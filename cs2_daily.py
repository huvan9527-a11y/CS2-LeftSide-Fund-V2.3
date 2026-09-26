# @title 点击运行 CS2 V2.6（完成后自动下载报告ZIP）
# -*- coding: utf-8 -*-
# CS2 LeftSide Fund V2.6 LIVE — 1–5天跌势减弱便携版
# 资产池: 代码内嵌原 hot_pool_v3.csv 的全部物品名称（643 个）
# 行情/K线: 从 SteamDT 实时 API 获取；不依赖原 CSV、JSON 文件
# 执行: python cs2_v26_portable.py
# 可选: STEAMDT_API_KEY 环境变量优先于代码内密钥，CS2_MODE=file 可切到文件行情模式
# 注意: 代码内密钥是明文，请勿公开分享；运行需要联网、可写目录和依赖包。
#
import importlib.util
import subprocess
import sys

for _pkg in ("pandas", "openpyxl", "requests", "tqdm"):
    if importlib.util.find_spec(_pkg) is None:
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", _pkg], check=False)

import time
import os
import re
import json
import glob
import math
from bisect import bisect_right
import requests
import pandas as pd
from tqdm import tqdm
from datetime import datetime, timedelta, timezone

try:
    from google.colab import files as colab_files
    IN_COLAB = True
except Exception:
    colab_files = None
    IN_COLAB = False

# ============================================================
# 0. CONFIG —— 所有参数集中在这里，调参只改这一块
# ============================================================
CONFIG = {
    # 平台字段 -> 标准名（统一大写后匹配，未知平台按原名小写保留）
    "PLATFORM_ALIAS": {
        "YOUPIN": "youpin", "YOUPIP": "youpin", "UUYP": "youpin",
        "悠悠有品": "youpin", "悠悠": "youpin",
        "STEAM": "steam", "STEAMDT": "steam",
        "BUFF": "buff", "BUFF165": "buff",
        "IGXE": "igxe", "C5": "c5", "C5GAME": "c5",
    },
    "EXCLUDE_PLATFORMS": {"steam"},                 # 彻底剔除，不解析不入表
    "SCORE_PLATFORM": "youpin",                     # 评分只基于悠悠
    "REFERENCE_PLATFORMS": ["buff", "c5"],          # 仅作参考列展示，不入分(实测c5买一价数据异常,勿信)

    # 价格单位: "yuan"=元 / "cents"=分 / "auto"=自动检测(交叉验证+中位数兜底)
    "PRICE_UNIT": "auto",

    # ---- V2.6：盘口50 + 品质5 + 位置10 + 下跌动能减弱20 + 低点10 + 卖压收敛5 = 100 ----
    # 盘口基于悠悠；刀/手套属性不再主导是否接近止跌。
    "DEPTH_STEPS": [(0.30, 20), (0.20, 17), (0.15, 14), (0.10, 11), (0.06, 7)],
    "DEPTH_BASE": 3,
    "DEPTH_MIN_BIDS": 5,
    "DEPTH_LOWBID_CAP": 8,
    "TIGHTNESS_STEPS": [(0.99, 15), (0.96, 12), (0.93, 9), (0.88, 6), (0.80, 3)],
    "TIGHTNESS_BASE": 1,
    "BID_PREMIUM_FLAG": 1.05,       # 超过5%可能是定向求购；不作为有效买一价
    # log10(悠悠挂卖数)分档: 900/500/250/100/30 —— 对数化压缩顶端,
    # 降低与深度分(buy/sell比)的相关性(原线性档位两者共变严重)
    "LIQ_STEPS": [(2.954, 15), (2.699, 12), (2.398, 9), (2.0, 7), (1.477, 4)],
    "LIQ_BASE": 2,
    "QUALITY_STAR": 5,
    "QUALITY_LEGEND": 5,
    "QUALITY_BASE": 2,

    # K线为小时数据：用前/近12小时、前/近24小时的跌幅比较，寻找1–5天左侧候选。
    "KLINE_TYPE": 1,
    "KLINE_MIN_POINTS": 20,
    "KLINE_STALE_HOURS": 12,       # 1–5天观察窗允许半天以内的K线，过期不作今日信号
    "KLINE_WINDOW_HOURS": 120,     # 回撤位置改用5日窗口(API实测返回90天小时线, 48h只覆盖2天)
    "KLINE_TREND_HOURS": 720,      # 30日趋势: 识别长期阴跌, 过滤下跌中继
    "LONG_DECLINE_30D": -0.30,     # 30日跌幅超过30%判为长期阴跌, 不出左侧候选
    "KLINE_MAX_GAP_HOURS": 3,      # 时间锚点缺口过大时不拿旧价代替
    "KLINE_MIN_CHANGES_24H": 2,    # 过少的价格更新不能等同于止跌
    "MIN_PREV_DROP_12H": 0.005,   # 前12h至少跌0.5%才比较减速
    "MIN_PREV_DROP_24H": 0.008,   # 前24h至少跌0.8%才比较减速
    "POSITION_STEPS": [(0.08, 10), (0.05, 8), (0.03, 6), (0.015, 4), (0.005, 2)],  # 作用于5日回撤dd120
    "POSITION_BASE": 0,
    "MIN_CANDIDATE_DRAWDOWN": 0.01,
    "MIN_CANDIDATE_MOMENTUM": 4,
    "MAX_CANDIDATE_REBOUND": 0.15,  # 距48h低点超过15%当作可能已错过左侧
    # 低点企稳联合判定: 未创新低 且 未大幅反弹 才计分(原两项独立相加, 逻辑不自洽)
    "REBOUND_FULL": 0.06,           # 反弹<=6%: 企稳分全额
    "REBOUND_HALF": 0.10,           # 反弹<=10%: 半额; 更远: 0分(左侧位置已过)

    # ---- SteamDT开放平台实时API ----
    "USE_LIVE_API": True,          # True且有key=实时拉取; 否则读本地文件
    "STEAMDT_API_KEY": "",  # 环境变量同名可覆盖，文件不要公开
    "API_BASE": "https://open.steamdt.com",
    "BATCH_SIZE": 100,             # 批量行情每批名字数(实测>100返回100002参数错误)
    "BATCH_SLEEP": 61,             # 批与批之间等待秒数(限频保护)
    "KLINE_SLEEP": 0.55,           # K线请求间隔秒(限频120/分)
    "API_MAX_RETRY": 4,

    # ---- 资金 ----
    "FUND_TIERS": [(85, 50000), (75, 25000), (60, 10000)],  # (分数下限, 金额)
    "TOTAL_BUDGET": 500000,          # 总预算
    "USE_BUDGET_CAP": True,          # True=按分数从高到低分配,总额封顶; False=只按单品定档

    # ---- 输出 ----
    "MD_TOP_N": 50,
    "SAVE_SNAPSHOT": True,
    "TZ_OFFSET_HOURS": 8,            # 东八区，避免Colab的UTC日期问题
}


def now_cn():
    return datetime.now(timezone(timedelta(hours=CONFIG["TZ_OFFSET_HOURS"])))


def fnum(x):
    """安全转float, 失败返回None"""
    try:
        if x is None:
            return None
        f = float(x)
        return f
    except Exception:
        return None


def fint(x):
    f = fnum(x)
    return int(f) if f is not None and f == f else 0  # NaN->0


def read_csv_any(path, **kw):
    """编码自适应: 先utf-8-sig(含BOM), 失败退gb18030(中文Windows导出)"""
    for enc in ("utf-8-sig", "gb18030"):
        try:
            return pd.read_csv(path, encoding=enc, **kw)
        except UnicodeDecodeError:
            continue
    raise Exception(f"无法识别文件编码(需UTF-8或GBK): {path}")


# ============================================================
# 1. 代码内嵌资产池（仅物品名称；原 CSV 的行情/K线较旧，不嵌入）
#    有效价格和 K 线均由 API 实时请求；不再要求外部资产池 CSV。
# ============================================================
EMBEDDED_POOL_NAMES = (
    'USP-S | Guardian (Factory New)',
    'Five-SeveN | Fowl Play (Factory New)',
    'USP-S | Caiman (Factory New)',
    'Five-SeveN | Kami (Factory New)',
    'Tec-9 | Blue Titanium (Factory New)',
    'Tec-9 | Bamboo Forest (Factory New)',
    'Glock-18 | Grinder (Factory New)',
    'Five-SeveN | Nightshade (Factory New)',
    'FAMAS | Neural Net (Factory New)',
    'Tec-9 | Ossified (Factory New)',
    'P250 | Splash (Factory New)',
    'AWP | Electric Hive (Factory New)',
    'M4A4 | Desert-Strike (Factory New)',
    'P250 | Cartel (Factory New)',
    'USP-S | Orion (Factory New)',
    'M4A4 | X-Ray (Factory New)',
    'Desert Eagle | Hypnotic (Factory New)',
    'Tec-9 | Avalanche (Factory New)',
    'Desert Eagle | Cobalt Disruption (Factory New)',
    'Tec-9 | Toxic (Factory New)',
    'USP-S | Blood Tiger (Factory New)',
    'MP9 | Rose Iron (Factory New)',
    'Glock-18 | Fade (Factory New)',
    'AWP | Lightning Strike (Factory New)',
    'Desert Eagle | Blaze (Factory New)',
    'MP9 | Hypnotic (Factory New)',
    'M4A4 | 龍王 (Dragon King) (Factory New)',
    'AWP | Graphite (Factory New)',
    'M4A1-S | Guardian (Factory New)',
    'MP9 | Ruby Poison Dart (Factory New)',
    'FAMAS | Survivor Z (Factory New)',
    'Glock-18 | Wraiths (Factory New)',
    'MP9 | Dart (Factory New)',
    'Glock-18 | Bunsen Burner (Factory New)',
    'P250 | Steel Disruption (Factory New)',
    'Galil AR | Kami (Factory New)',
    'Five-SeveN | Copper Galaxy (Factory New)',
    'Tec-9 | Titanium Bit (Factory New)',
    'P250 | Mint Kimono (Factory New)',
    'MP9 | Hot Rod (Factory New)',
    'Five-SeveN | Case Hardened (Factory New)',
    'FAMAS | Hexane (Factory New)',
    'Galil AR | Blue Titanium (Factory New)',
    'M4A1-S | Blood Tiger (Factory New)',
    'Five-SeveN | Anodized Gunmetal (Factory New)',
    'M4A1-S | Boreal Forest (Factory New)',
    'M4A4 | Evil Daimyo (Factory New)',
    'AK-47 | Blue Laminate (Factory New)',
    'FAMAS | Pulse (Factory New)',
    'AWP | Sun in Leo (Factory New)',
    'Tec-9 | Isaac (Factory New)',
    'AK-47 | Elite Build (Factory New)',
    'Desert Eagle | Naga (Factory New)',
    'FAMAS | Afterimage (Factory New)',
    'Glock-18 | Steel Disruption (Factory New)',
    'Desert Eagle | Conspiracy (Factory New)',
    'P250 | Wingshot (Factory New)',
    'Five-SeveN | Hot Shot (Factory New)',
    'USP-S | Stainless (Factory New)',
    'AK-47 | Emerald Pinstripe (Factory New)',
    'P250 | Undertow (Factory New)',
    'P250 | Muertos (Factory New)',
    'Desert Eagle | Sunset Storm 弐 (Factory New)',
    'Desert Eagle | Sunset Storm 壱 (Factory New)',
    'Desert Eagle | Golden Koi (Factory New)',
    'Desert Eagle | Pilot (Factory New)',
    'Desert Eagle | Crimson Web (Factory New)',
    'Desert Eagle | Midnight Storm (Factory New)',
    'Five-SeveN | Retrobution (Factory New)',
    'FAMAS | Djinn (Factory New)',
    'Galil AR | Rocket Pop (Factory New)',
    'M4A1-S | Nitro (Factory New)',
    'Glock-18 | Dragon Tattoo (Factory New)',
    'AWP | Hyper Beast (Factory New)',
    'AK-47 | Vulcan (Factory New)',
    'M4A1-S | Cyrex (Factory New)',
    'M4A4 | Griffin (Factory New)',
    'P250 | Hive (Factory New)',
    'P250 | Gunsmoke (Factory New)',
    'M4A1-S | Atomic Alloy (Factory New)',
    'AK-47 | Cartel (Factory New)',
    'Glock-18 | Water Elemental (Factory New)',
    'Galil AR | Orange DDPAT (Factory New)',
    'M4A1-S | Hyper Beast (Factory New)',
    'M4A1-S | Icarus Fell (Factory New)',
    'Galil AR | Hunting Blind (Factory New)',
    'MP9 | Deadly Poison (Factory New)',
    'M4A4 | Desert Storm (Factory New)',
    'AK-47 | Jaguar (Factory New)',
    'M4A1-S | Hot Rod (Factory New)',
    'M4A1-S | Knight (Factory New)',
    'AWP | BOOM (Factory New)',
    'AWP | Corticera (Factory New)',
    'AK-47 | Aquamarine Revenge (Factory New)',
    'AK-47 | Frontside Misty (Factory New)',
    'AK-47 | Jungle Spray (Factory New)',
    'M4A1-S | Basilisk (Factory New)',
    'P250 | Crimson Kimono (Factory New)',
    'USP-S | Para Green (Factory New)',
    'Tec-9 | Hades (Factory New)',
    'AK-47 | Case Hardened (Factory New)',
    'AWP | Dragon Lore (Factory New)',
    'Glock-18 | Twilight Galaxy (Factory New)',
    'FAMAS | Cyanospatter (Factory New)',
    'P250 | Contamination (Factory New)',
    'AWP | Safari Mesh (Factory New)',
    'FAMAS | Styx (Factory New)',
    'Galil AR | Aqua Terrace (Factory New)',
    'Galil AR | Cerberus (Factory New)',
    'Glock-18 | Night (Factory New)',
    'Galil AR | Urban Rubble (Factory New)',
    'MP9 | Green Plaid (Factory New)',
    'Desert Eagle | Heirloom (Factory New)',
    'M4A4 | Daybreak (Factory New)',
    'USP-S | Overgrowth (Factory New)',
    'Galil AR | Stone Cold (Factory New)',
    'Glock-18 | Brass (Factory New)',
    'USP-S | Road Rash (Factory New)',
    'AK-47 | Black Laminate (Factory New)',
    'USP-S | Serum (Factory New)',
    'AK-47 | Red Laminate (Factory New)',
    'M4A4 | Jungle Tiger (Factory New)',
    'Five-SeveN | Candy Apple (Factory New)',
    'Glock-18 | Blue Fissure (Factory New)',
    'Desert Eagle | Hand Cannon (Factory New)',
    'Tec-9 | Terrace (Factory New)',
    'USP-S | Royal Blue (Factory New)',
    'M4A4 | Bullet Rain (Factory New)',
    'P250 | Mehndi (Factory New)',
    'M4A4 | Faded Zebra (Factory New)',
    'M4A4 | Radiation Hazard (Factory New)',
    'AK-47 | Jet Set (Factory New)',
    'M4A1-S | Golden Coil (Factory New)',
    'AK-47 | Safari Mesh (Factory New)',
    'USP-S | Kill Confirmed (Factory New)',
    'Five-SeveN | Orange Peel (Factory New)',
    'USP-S | Forest Leaves (Factory New)',
    'AK-47 | Point Disarray (Factory New)',
    'MP9 | Dark Age (Factory New)',
    'Glock-18 | Reactor (Factory New)',
    "MP9 | Pandora's Box (Factory New)",
    'USP-S | Business Class (Factory New)',
    'M4A4 | Howl (Factory New)',
    'M4A4 | Urban DDPAT (Factory New)',
    'M4A4 | Royal Paladin (Factory New)',
    'AK-47 | Hydroponic (Factory New)',
    'Desert Eagle | Urban DDPAT (Factory New)',
    'M4A4 | Poseidon (Factory New)',
    'Five-SeveN | Jungle (Factory New)',
    'M4A4 | Zirka (Factory New)',
    'Desert Eagle | Mudder (Factory New)',
    'P250 | Facets (Factory New)',
    'P250 | Bone Mask (Factory New)',
    'AK-47 | Fire Serpent (Factory New)',
    'Galil AR | Shattered (Factory New)',
    'MP9 | Setting Sun (Factory New)',
    'AWP | Pink DDPAT (Factory New)',
    'M4A1-S | Master Piece (Factory New)',
    'AK-47 | Predator (Factory New)',
    'Five-SeveN | Neon Kimono (Factory New)',
    'Five-SeveN | Triumvirate (Factory New)',
    'M4A4 | The Battlestar (Factory New)',
    'MP9 | Storm (Factory New)',
    'USP-S | Lead Conduit (Factory New)',
    'Tec-9 | Jambiya (Factory New)',
    'Galil AR | Winter Forest (Factory New)',
    'P250 | Whiteout (Factory New)',
    'AK-47 | Fuel Injector (Factory New)',
    'FAMAS | Valence (Factory New)',
    'Desert Eagle | Kumicho Dragon (Factory New)',
    'Glock-18 | Royal Legion (Factory New)',
    'AWP | Elite Build (Factory New)',
    'Desert Eagle | Night (Factory New)',
    'AK-47 | Wasteland Rebel (Factory New)',
    'M4A4 | Tornado (Factory New)',
    'AK-47 | First Class (Factory New)',
    'Five-SeveN | Nitro (Factory New)',
    'Galil AR | Firefight (Factory New)',
    'Glock-18 | Sand Dune (Factory New)',
    'Tec-9 | Re-Entry (Factory New)',
    'AWP | Medusa (Factory New)',
    'AWP | Phobos (Factory New)',
    "M4A1-S | Chantico's Fire (Factory New)",
    'Tec-9 | Brass (Factory New)',
    'Five-SeveN | Contractor (Factory New)',
    'M4A4 | Modern Hunter (Factory New)',
    'M4A4 | Desolate Space (Factory New)',
    'Glock-18 | Wasteland Rebel (Factory New)',
    'M4A1-S | Mecha Industries (Factory New)',
    'P250 | Nuclear Threat (Factory New)',
    'Tec-9 | Nuclear Threat (Factory New)',
    'Tec-9 | Fuel Injector (Factory New)',
    'Five-SeveN | Scumbria (Factory New)',
    'AK-47 | Neon Revolution (Factory New)',
    'P250 | Iron Clad (Factory New)',
    'Desert Eagle | Directive (Factory New)',
    'Glock-18 | Weasel (Factory New)',
    'Glock-18 | Groundwater (Factory New)',
    'MP9 | Airlock (Factory New)',
    'MP9 | Dry Season (Factory New)',
    'FAMAS | Roll Cage (Factory New)',
    'P250 | Modern Hunter (Factory New)',
    'FAMAS | Spitfire (Factory New)',
    'MP9 | Sand Scale (Factory New)',
    'USP-S | Cyrex (Factory New)',
    'Galil AR | Black Sand (Factory New)',
    'MP9 | Bulldozer (Factory New)',
    '★ Hand Wraps | Slaughter (Field-Tested)',
    'M4A1-S | Flashback (Factory New)',
    'M4A4 | Buzz Kill (Factory New)',
    'FAMAS | Mecha Industries (Factory New)',
    '★ Bloodhound Gloves | Snakebite (Field-Tested)',
    '★ Bloodhound Gloves | Charred (Field-Tested)',
    'Glock-18 | Ironwork (Factory New)',
    '★ Moto Gloves | Spearmint (Minimal Wear)',
    '★ Driver Gloves | Convoy (Minimal Wear)',
    '★ Sport Gloves | Arid (Field-Tested)',
    '★ Bloodhound Gloves | Guerrilla (Field-Tested)',
    '★ Moto Gloves | Cool Mint (Field-Tested)',
    '★ Driver Gloves | Diamondback (Field-Tested)',
    '★ Driver Gloves | Lunar Weave (Field-Tested)',
    '★ Sport Gloves | Hedge Maze (Minimal Wear)',
    '★ Specialist Gloves | Crimson Kimono (Field-Tested)',
    '★ Bloodhound Gloves | Bronzed (Field-Tested)',
    '★ Hand Wraps | Badlands (Field-Tested)',
    '★ Driver Gloves | Crimson Weave (Minimal Wear)',
    '★ Moto Gloves | Boom! (Field-Tested)',
    'AWP | Snake Camo (Factory New)',
    '★ Moto Gloves | Eclipse (Field-Tested)',
    '★ Specialist Gloves | Forest DDPAT (Field-Tested)',
    '★ Bloodhound Gloves | Charred (Minimal Wear)',
    '★ Bloodhound Gloves | Guerrilla (Minimal Wear)',
    '★ Moto Gloves | Spearmint (Field-Tested)',
    '★ Hand Wraps | Badlands (Minimal Wear)',
    '★ Hand Wraps | Spruce DDPAT (Minimal Wear)',
    '★ Moto Gloves | Cool Mint (Minimal Wear)',
    '★ Bloodhound Gloves | Snakebite (Minimal Wear)',
    '★ Hand Wraps | Spruce DDPAT (Field-Tested)',
    "★ Sport Gloves | Pandora's Box (Field-Tested)",
    "★ Sport Gloves | Pandora's Box (Minimal Wear)",
    '★ Driver Gloves | Lunar Weave (Minimal Wear)',
    '★ Driver Gloves | Crimson Weave (Field-Tested)',
    '★ Specialist Gloves | Foundation (Field-Tested)',
    '★ Driver Gloves | Diamondback (Minimal Wear)',
    '★ Hand Wraps | Leather (Field-Tested)',
    '★ Hand Wraps | Slaughter (Minimal Wear)',
    '★ Specialist Gloves | Foundation (Minimal Wear)',
    '★ Sport Gloves | Superconductor (Minimal Wear)',
    '★ Specialist Gloves | Emerald Web (Minimal Wear)',
    '★ Sport Gloves | Hedge Maze (Field-Tested)',
    '★ Specialist Gloves | Crimson Kimono (Minimal Wear)',
    '★ Driver Gloves | Convoy (Field-Tested)',
    '★ Specialist Gloves | Forest DDPAT (Minimal Wear)',
    '★ Hand Wraps | Leather (Minimal Wear)',
    '★ Specialist Gloves | Emerald Web (Field-Tested)',
    'FAMAS | Contrast Spray (Factory New)',
    '★ Moto Gloves | Boom! (Minimal Wear)',
    '★ Sport Gloves | Superconductor (Field-Tested)',
    '★ Moto Gloves | Eclipse (Minimal Wear)',
    '★ Bloodhound Gloves | Bronzed (Minimal Wear)',
    'AWP | Fever Dream (Factory New)',
    'USP-S | Neo-Noir (Factory New)',
    'AK-47 | Bloodsport (Factory New)',
    'Galil AR | Crimson Tsunami (Factory New)',
    '★ Sport Gloves | Arid (Minimal Wear)',
    'M4A1-S | Decimator (Factory New)',
    'FAMAS | Macabre (Factory New)',
    'Galil AR | Sugar Rush (Factory New)',
    'AK-47 | Orbit Mk01 (Factory New)',
    'Tec-9 | Cut Out (Factory New)',
    'P250 | Red Rock (Factory New)',
    'AWP | Oni Taiji (Factory New)',
    'M4A4 | Hellfire (Factory New)',
    'M4A1-S | Briefing (Factory New)',
    'Five-SeveN | Hyper Beast (Factory New)',
    'USP-S | Blueprint (Factory New)',
    'Tec-9 | Cracked Opal (Factory New)',
    'AK-47 | The Empress (Factory New)',
    'Glock-18 | Off World (Factory New)',
    'M4A1-S | Leaded Glass (Factory New)',
    'MP9 | Goo (Factory New)',
    'P250 | See Ya Later (Factory New)',
    '★ Sport Gloves | Vice (Field-Tested)',
    'AWP | Mortis (Factory New)',
    '★ Sport Gloves | Bronze Morph (Field-Tested)',
    '★ Moto Gloves | Polygon (Field-Tested)',
    '★ Moto Gloves | Polygon (Minimal Wear)',
    '★ Specialist Gloves | Mogul (Field-Tested)',
    '★ Hydra Gloves | Case Hardened (Field-Tested)',
    'M4A4 | Neo-Noir (Factory New)',
    '★ Specialist Gloves | Mogul (Minimal Wear)',
    '★ Hydra Gloves | Rattler (Minimal Wear)',
    '★ Hydra Gloves | Mangrove (Field-Tested)',
    '★ Hand Wraps | Overprint (Field-Tested)',
    'USP-S | Cortex (Factory New)',
    '★ Driver Gloves | Imperial Plaid (Minimal Wear)',
    '★ Hydra Gloves | Emerald (Field-Tested)',
    '★ Hand Wraps | Duct Tape (Minimal Wear)',
    '★ Hand Wraps | Overprint (Minimal Wear)',
    '★ Driver Gloves | King Snake (Field-Tested)',
    '★ Hand Wraps | Arboreal (Field-Tested)',
    '★ Hydra Gloves | Rattler (Field-Tested)',
    '★ Hand Wraps | Cobalt Skulls (Minimal Wear)',
    '★ Moto Gloves | Transport (Field-Tested)',
    '★ Driver Gloves | Overtake (Field-Tested)',
    '★ Moto Gloves | POW! (Minimal Wear)',
    '★ Sport Gloves | Omega (Field-Tested)',
    '★ Moto Gloves | Turtle (Field-Tested)',
    '★ Hand Wraps | Duct Tape (Field-Tested)',
    '★ Moto Gloves | POW! (Field-Tested)',
    '★ Specialist Gloves | Crimson Web (Field-Tested)',
    '★ Sport Gloves | Amphibious (Field-Tested)',
    '★ Driver Gloves | Racing Green (Field-Tested)',
    '★ Specialist Gloves | Buckshot (Field-Tested)',
    '★ Hand Wraps | Arboreal (Minimal Wear)',
    '★ Driver Gloves | Racing Green (Minimal Wear)',
    '★ Specialist Gloves | Buckshot (Minimal Wear)',
    '★ Hand Wraps | Cobalt Skulls (Field-Tested)',
    '★ Moto Gloves | Transport (Minimal Wear)',
    '★ Hydra Gloves | Emerald (Minimal Wear)',
    '★ Specialist Gloves | Fade (Field-Tested)',
    '★ Sport Gloves | Bronze Morph (Minimal Wear)',
    '★ Driver Gloves | Imperial Plaid (Field-Tested)',
    '★ Hydra Gloves | Mangrove (Minimal Wear)',
    '★ Moto Gloves | Turtle (Minimal Wear)',
    '★ Sport Gloves | Amphibious (Minimal Wear)',
    '★ Driver Gloves | Overtake (Minimal Wear)',
    '★ Sport Gloves | Omega (Minimal Wear)',
    '★ Driver Gloves | King Snake (Minimal Wear)',
    '★ Sport Gloves | Vice (Minimal Wear)',
    '★ Hydra Gloves | Case Hardened (Minimal Wear)',
    '★ Specialist Gloves | Fade (Minimal Wear)',
    '★ Specialist Gloves | Crimson Web (Minimal Wear)',
    'AWP | PAW (Factory New)',
    'M4A1-S | Nightmare (Factory New)',
    'Desert Eagle | Code Red (Factory New)',
    'AK-47 | Neon Rider (Factory New)',
    'FAMAS | Eye of Athena (Factory New)',
    'Tec-9 | Snek-9 (Factory New)',
    'Glock-18 | Warhawk (Factory New)',
    'Tec-9 | Remote Control (Factory New)',
    'AWP | Acheron (Factory New)',
    'AK-47 | Safety Net (Factory New)',
    'Glock-18 | Nuclear Garden (Factory New)',
    'M4A1-S | Control Panel (Factory New)',
    'P250 | Vino Primo (Factory New)',
    'AWP | Neo-Noir (Factory New)',
    'Desert Eagle | Mecha Industries (Factory New)',
    'M4A4 | Magnesium (Factory New)',
    'AK-47 | Asiimov (Factory New)',
    'Tec-9 | Bamboozle (Factory New)',
    'Desert Eagle | Light Rail (Factory New)',
    'Five-SeveN | Angry Mob (Factory New)',
    'AWP | Atheris (Factory New)',
    'M4A4 | The Emperor (Factory New)',
    'AWP | Wildfire (Factory New)',
    'FAMAS | Commemoration (Factory New)',
    'FAMAS | Decommissioned (Factory New)',
    'Tec-9 | Flash Out (Factory New)',
    'P250 | Inferno (Factory New)',
    'Five-SeveN | Buddy (Factory New)',
    'Glock-18 | Sacrifice (Factory New)',
    'MP9 | Hydra (Factory New)',
    'USP-S | Pathfinder (Factory New)',
    'FAMAS | Night Borre (Factory New)',
    'AK-47 | Rat Rod (Factory New)',
    'Tec-9 | Orange Murano (Factory New)',
    'AWP | Containment Breach (Factory New)',
    'Tec-9 | Rust Leaf (Factory New)',
    'Tec-9 | Decimator (Factory New)',
    'M4A1-S | Moss Quartz (Factory New)',
    'M4A4 | Dark Blossom (Factory New)',
    'AK-47 | Baroque Purple (Factory New)',
    'P250 | Dark Filigree (Factory New)',
    'Desert Eagle | Emerald Jörmungandr (Factory New)',
    'Glock-18 | Synth Leaf (Factory New)',
    'FAMAS | Sundown (Factory New)',
    'MP9 | Stained Glass (Factory New)',
    'AWP | Gungnir (Factory New)',
    'Five-SeveN | Crimson Blossom (Factory New)',
    'Galil AR | Tornado (Factory New)',
    'MP9 | Wild Lily (Factory New)',
    'AWP | The Prince (Factory New)',
    'AK-47 | Wild Lotus (Factory New)',
    'AWP | Capillary (Factory New)',
    'AK-47 | Phantom Disruptor (Factory New)',
    'M4A1-S | Player Two (Factory New)',
    'Glock-18 | Bullet Queen (Factory New)',
    'Tec-9 | Brother (Factory New)',
    'M4A4 | Tooth Fairy (Factory New)',
    'Glock-18 | Vogue (Factory New)',
    'Desert Eagle | Printstream (Factory New)',
    'AK-47 | Legion of Anubis (Factory New)',
    'M4A4 | Global Offensive (Factory New)',
    'Desert Eagle | The Bronze (Factory New)',
    'Tec-9 | Phoenix Chalk (Factory New)',
    'Five-SeveN | Berries And Cherries (Factory New)',
    'P250 | Forest Night (Factory New)',
    'FAMAS | Prime Conspiracy (Factory New)',
    'USP-S | Monster Mashup (Factory New)',
    'Glock-18 | Franklin (Factory New)',
    'M4A1-S | Blue Phosphor (Factory New)',
    'AWP | Silk Tiger (Factory New)',
    'P250 | Contaminant (Factory New)',
    'M4A1-S | Printstream (Factory New)',
    'AWP | Fade (Factory New)',
    '★ Hand Wraps | Desert Shamagh (Minimal Wear)',
    '★ Driver Gloves | Snow Leopard (Minimal Wear)',
    '★ Moto Gloves | 3rd Commando Company (Field-Tested)',
    '★ Hand Wraps | Constrictor (Field-Tested)',
    'Desert Eagle | Night Heist (Factory New)',
    'M4A4 | Cyber Security (Factory New)',
    'Galil AR | Vandal (Factory New)',
    '★ Specialist Gloves | Marble Fade (Field-Tested)',
    '★ Broken Fang Gloves | Yellow-banded (Minimal Wear)',
    '★ Broken Fang Gloves | Needle Point (Field-Tested)',
    '★ Specialist Gloves | Field Agent (Field-Tested)',
    '★ Hand Wraps | Desert Shamagh (Field-Tested)',
    '★ Sport Gloves | Big Game (Field-Tested)',
    '★ Broken Fang Gloves | Unhinged (Field-Tested)',
    '★ Driver Gloves | Rezan the Red (Field-Tested)',
    'Glock-18 | Neo-Noir (Factory New)',
    'Five-SeveN | Fairy Tale (Factory New)',
    'USP-S | Target Acquired (Factory New)',
    'Galil AR | Dusk Ruins (Factory New)',
    '★ Moto Gloves | Finish Line (Field-Tested)',
    'Galil AR | Phoenix Blacklight (Factory New)',
    '★ Specialist Gloves | Tiger Strike (Field-Tested)',
    '★ Hand Wraps | Giraffe (Field-Tested)',
    '★ Sport Gloves | Slingshot (Minimal Wear)',
    'USP-S | Ancient Visions (Factory New)',
    '★ Sport Gloves | Nocts (Field-Tested)',
    '★ Driver Gloves | Queen Jaguar (Field-Tested)',
    '★ Driver Gloves | Snow Leopard (Field-Tested)',
    '★ Specialist Gloves | Lt. Commander (Field-Tested)',
    '★ Hand Wraps | Constrictor (Minimal Wear)',
    '★ Broken Fang Gloves | Unhinged (Minimal Wear)',
    '★ Sport Gloves | Scarlet Shamagh (Field-Tested)',
    '★ Hand Wraps | Giraffe (Minimal Wear)',
    '★ Specialist Gloves | Marble Fade (Minimal Wear)',
    '★ Sport Gloves | Big Game (Minimal Wear)',
    '★ Hand Wraps | CAUTION! (Field-Tested)',
    '★ Hand Wraps | CAUTION! (Minimal Wear)',
    '★ Specialist Gloves | Lt. Commander (Minimal Wear)',
    'Tec-9 | Blast From the Past (Factory New)',
    '★ Moto Gloves | Smoke Out (Minimal Wear)',
    'AK-47 | Panthera onca (Factory New)',
    '★ Driver Gloves | Black Tie (Field-Tested)',
    '★ Moto Gloves | Blood Pressure (Minimal Wear)',
    '★ Broken Fang Gloves | Jade (Minimal Wear)',
    '★ Moto Gloves | Finish Line (Minimal Wear)',
    '★ Sport Gloves | Scarlet Shamagh (Minimal Wear)',
    '★ Broken Fang Gloves | Yellow-banded (Field-Tested)',
    'P250 | Bengal Tiger (Factory New)',
    '★ Specialist Gloves | Field Agent (Minimal Wear)',
    '★ Driver Gloves | Black Tie (Minimal Wear)',
    '★ Moto Gloves | Blood Pressure (Field-Tested)',
    '★ Specialist Gloves | Tiger Strike (Minimal Wear)',
    '★ Broken Fang Gloves | Jade (Field-Tested)',
    '★ Moto Gloves | Smoke Out (Field-Tested)',
    '★ Sport Gloves | Slingshot (Field-Tested)',
    '★ Sport Gloves | Nocts (Minimal Wear)',
    '★ Moto Gloves | 3rd Commando Company (Minimal Wear)',
    '★ Driver Gloves | Rezan the Red (Minimal Wear)',
    '★ Driver Gloves | Queen Jaguar (Minimal Wear)',
    'M4A1-S | Welcome to the Jungle (Factory New)',
    '★ Broken Fang Gloves | Needle Point (Minimal Wear)',
    'AK-47 | X-Ray (Factory New)',
    'AK-47 | Slate (Factory New)',
    'MP9 | Food Chain (Factory New)',
    'Desert Eagle | Trigger Discipline (Factory New)',
    'P250 | Cyber Shell (Factory New)',
    'Galil AR | Chromatic Aberration (Factory New)',
    'M4A4 | In Living Color (Factory New)',
    'USP-S | The Traitor (Factory New)',
    'MP9 | Mount Fuji (Factory New)',
    'Tec-9 | Safety Net (Factory New)',
    'Glock-18 | Red Tire (Factory New)',
    'Galil AR | Amber Fade (Factory New)',
    'P250 | Black & Tan (Factory New)',
    'Glock-18 | Gamma Doppler (Factory New)',
    'USP-S | Purple DDPAT (Factory New)',
    'Five-SeveN | Boost Protocol (Factory New)',
    'MP9 | Music Box (Factory New)',
    'USP-S | Black Lotus (Factory New)',
    'FAMAS | Meltdown (Factory New)',
    'M4A4 | Red DDPAT (Factory New)',
    'M4A1-S | Fizzy POP (Factory New)',
    'AWP | POP AWP (Factory New)',
    'FAMAS | Faulty Wiring (Factory New)',
    'Desert Eagle | Sputnik (Factory New)',
    'FAMAS | ZX Spectron (Factory New)',
    'AWP | Desert Hydra (Factory New)',
    'USP-S | Orange Anolis (Factory New)',
    'M4A4 | Spider Lily (Factory New)',
    'Glock-18 | Snack Attack (Factory New)',
    'P250 | Digital Architect (Factory New)',
    'AK-47 | Leet Museo (Factory New)',
    'Desert Eagle | Ocean Drive (Factory New)',
    'M4A4 | The Coalition (Factory New)',
    'USP-S | Whiteout (Factory New)',
    'Five-SeveN | Fall Hazard (Factory New)',
    'Desert Eagle | Fennec Fox (Factory New)',
    'AK-47 | Gold Arabesque (Factory New)',
    'AK-47 | Green Laminate (Factory New)',
    'Galil AR | CAUTION! (Factory New)',
    'Glock-18 | Pink DDPAT (Factory New)',
    'M4A1-S | Imminent Danger (Factory New)',
    'AK-47 | Nightwish (Factory New)',
    'FAMAS | Rapid Eye Movement (Factory New)',
    'MP9 | Starlight Protector (Factory New)',
    'P250 | Visions (Factory New)',
    'AWP | Chromatic Aberration (Factory New)',
    'AK-47 | Ice Coaled (Factory New)',
    'USP-S | Printstream (Factory New)',
    'M4A4 | Temukau (Factory New)',
    'AK-47 | Head Shot (Factory New)',
    'AWP | Duality (Factory New)',
    'AK-47 | Steel Delta (Factory New)',
    'USP-S | Desert Tactical (Factory New)',
    'AWP | Black Nile (Factory New)',
    "Tec-9 | Mummy's Rot (Factory New)",
    'M4A4 | Eye of Horus (Factory New)',
    "Glock-18 | Ramese's Reach (Factory New)",
    'FAMAS | Waters of Nephthys (Factory New)',
    "P250 | Apep's Curse (Factory New)",
    'USP-S | Jawbreaker (Factory New)',
    'M4A4 | Etch Lord (Factory New)',
    'M4A1-S | Black Lotus (Factory New)',
    'AK-47 | Inheritance (Factory New)',
    'AWP | Chrome Cannon (Factory New)',
    'MP9 | Arctic Tri-Tone (Factory New)',
    'M4A1-S | Fade (Factory New)',
    'Galil AR | Rainbow Spoon (Factory New)',
    'AK-47 | Crossfade (Factory New)',
    'Desert Eagle | Starcade (Factory New)',
    'Desert Eagle | Heat Treated (Factory New)',
    'Desert Eagle | Calligraffiti (Factory New)',
    'USP-S | 27 (Factory New)',
    'M4A4 | Turbine (Factory New)',
    'M4A4 | Polysoup (Factory New)',
    'Five-SeveN | Heat Treated (Factory New)',
    'Glock-18 | Gold Toof (Factory New)',
    'Glock-18 | AXIA (Factory New)',
    'AWP | Crakow! (Factory New)',
    'P250 | Epicenter (Factory New)',
    'AK-47 | B the Monster (Factory New)',
    'AK-47 | The Outsiders (Factory New)',
    'M4A1-S | Vaporwave (Factory New)',
    'AWP | CMYK (Factory New)',
    'Glock-18 | Green Line (Factory New)',
    'USP-S | Royal Guard (Factory New)',
    'AWP | Printstream (Factory New)',
    'Glock-18 | Shinobu (Factory New)',
    'M4A1-S | Glitched Paint (Factory New)',
    'M4A4 | Hellish (Factory New)',
    'AK-47 | Searing Rage (Factory New)',
    'USP-S | Bleeding Edge (Factory New)',
    'Desert Eagle | Mulberry (Factory New)',
    'M4A4 | Sheet Lightning (Factory New)',
    'MP9 | Latte Rush (Factory New)',
    'AK-47 | Midnight Laminate (Factory New)',
    'FAMAS | Bad Trip (Factory New)',
    'AK-47 | Nouveau Rouge (Factory New)',
    'AWP | Green Energy (Factory New)',
    'Tec-9 | Whiteout (Factory New)',
    'AWP | LongDog (Factory New)',
    'M4A1-S | Stratosphere (Factory New)',
    'M4A1-S | Solitude (Factory New)',
    'M4A1-S | Liquidation (Factory New)',
    'Glock-18 | Mirror Mosaic (Factory New)',
    'AWP | Ice Coaled (Factory New)',
    'AK-47 | The Oligarch (Factory New)',
    'M4A4 | Full Throttle (Factory New)',
    'AK-47 | Aphrodite (Factory New)',
    'USP-S | Sleeping Potion (Factory New)',
    'AK-47 | Breakthrough (Factory New)',
    'Glock-18 | Trace Lock (Factory New)',
    'M4A1-S | Party Animal (Factory New)',
    'AWP | The End (Factory New)',
    'AWP | Exothermic (Factory New)',
    'AK-47 | Crane Flight (Factory New)',
    '★ Specialist Gloves | Lime Polycam (Minimal Wear)',
    '★ Driver Gloves | Wave Chaser (Field-Tested)',
    '★ Sport Gloves | Red Racer (Field-Tested)',
    '★ Specialist Gloves | Big Swell (Field-Tested)',
    '★ Sport Gloves | Ultra Violent (Minimal Wear)',
    '★ Specialist Gloves | Cloud Chaser (Minimal Wear)',
    '★ Specialist Gloves | Chocolate Chesterfield (Field-Tested)',
    'M4A1-S | Electrum (Factory New)',
    '★ Specialist Gloves | Lime Polycam (Field-Tested)',
    '★ Driver Gloves | Brocade Flowers (Field-Tested)',
    '★ Specialist Gloves | Big Swell (Minimal Wear)',
    'P250 | Kintsugi (Factory New)',
    '★ Sport Gloves | Occult (Minimal Wear)',
    "AWP | Queen's Gambit (Factory New)",
    '★ Sport Gloves | Blaze (Field-Tested)',
    '★ Specialist Gloves | Cloud Chaser (Field-Tested)',
    '★ Driver Gloves | Seigaiha (Field-Tested)',
    '★ Specialist Gloves | Blackbook (Field-Tested)',
    '★ Specialist Gloves | Pillow Punchers (Field-Tested)',
    '★ Sport Gloves | Occult (Field-Tested)',
    '★ Driver Gloves | Plum Quill (Minimal Wear)',
    '★ Sport Gloves | Frosty (Minimal Wear)',
    '★ Sport Gloves | Frosty (Field-Tested)',
    '★ Driver Gloves | Hand Sweaters (Field-Tested)',
    '★ Specialist Gloves | Sunburst (Field-Tested)',
    '★ Driver Gloves | Garden (Field-Tested)',
    '★ Specialist Gloves | Chocolate Chesterfield (Minimal Wear)',
    '★ Sport Gloves | Violet Beadwork (Minimal Wear)',
    '★ Sport Gloves | Ultra Violent (Field-Tested)',
    '★ Driver Gloves | Seigaiha (Minimal Wear)',
    '★ Sport Gloves | Creme Pinstripe (Field-Tested)',
    '★ Sport Gloves | Red Racer (Minimal Wear)',
    'Glock-18 | Fully Tuned (Factory New)',
    '★ Driver Gloves | Brocade Crane (Minimal Wear)',
    '★ Sport Gloves | Violet Beadwork (Field-Tested)',
    '★ Driver Gloves | Brocade Crane (Field-Tested)',
    '★ Driver Gloves | Brocade Flowers (Minimal Wear)',
    '★ Driver Gloves | Wave Chaser (Minimal Wear)',
    '★ Specialist Gloves | Pillow Punchers (Minimal Wear)',
    '★ Driver Gloves | Dragon Fists (Minimal Wear)',
    '★ Sport Gloves | Blaze (Minimal Wear)',
    '★ Driver Gloves | Garden (Minimal Wear)',
    '★ Driver Gloves | Dragon Fists (Field-Tested)',
    '★ Specialist Gloves | Blackbook (Minimal Wear)',
    '★ Driver Gloves | Hand Sweaters (Minimal Wear)',
    '★ Driver Gloves | Plum Quill (Field-Tested)',
    '★ Specialist Gloves | Sunburst (Minimal Wear)',
    '★ Sport Gloves | Creme Pinstripe (Minimal Wear)',
    'Tec-9 | Sultan (Factory New)',
    'M4A4 | Dark Operative (Factory New)',
    'P250 | Lotus Imprint (Factory New)',
    'AK-47 | Consequence of the Jinn (Factory New)',
    'AWP | Black Box (Factory New)',
    'Desert Eagle | Eastern Enigma (Factory New)',
    'M4A4 | Falak (Factory New)',
    'Glock-18 | Ifrit Lattice (Factory New)',
    'M4A1-S | Fatal Glitch (Factory New)',
    'USP-S | Spiral Glitch (Factory New)',
    'AWP | Sovereign Flame (Factory New)',
    'Glock-18 | Ghost Protocol (Factory New)',
    'AK-47 | AUTOEXEC (Factory New)',
)
print("环境:", "Google Colab" if IN_COLAB else "本地模式")

# 可选文件模式: 只有明确设置 CS2_MODE=file 时才读取本地行情文件；
# 实时模式可以不提供任何文件。
steam_file = youpin_file = None
for path in sorted(glob.glob("steamdt*.json") + glob.glob("steamdt*.csv") +
                   glob.glob("uploads/steamdt*.json") + glob.glob("uploads/steamdt*.csv")):
    if not steam_file:
        steam_file = path
for path in sorted(glob.glob("youpin*.json") + glob.glob("youpin*.csv") +
                   glob.glob("uploads/youpin*.json") + glob.glob("uploads/youpin*.csv")):
    if not youpin_file:
        youpin_file = path

API_KEY = (os.environ.get("STEAMDT_API_KEY") or CONFIG.get("STEAMDT_API_KEY") or "").strip()
file_mode = os.environ.get("CS2_MODE", "").lower() in ("file", "offline")
USE_LIVE = bool(CONFIG["USE_LIVE_API"] and API_KEY and not file_mode)
if not USE_LIVE and not file_mode:
    raise RuntimeError("请设置 STEAMDT_API_KEY；离线运行须设置 CS2_MODE=file")
if not USE_LIVE and not steam_file:
    raise RuntimeError("文件模式需要本地 steamdt*.json/csv；实时模式需要有效的 SteamDT API Key")

csv_file = f"代码内置资产池({len(EMBEDDED_POOL_NAMES)}件)"
if USE_LIVE:
    print(f"模式: 实时API(open.steamdt.com) | 资产池={csv_file}")
else:
    print(f"模式: 本地文件 | 资产池={csv_file} | SteamDT={steam_file} | YOUPIN={youpin_file or '(未提供)'}")

# ============================================================
# 2. 资产池（白名单）
# ============================================================
pool = pd.DataFrame({"marketHashName": EMBEDDED_POOL_NAMES})
pool_names = set(EMBEDDED_POOL_NAMES)
assert len(pool_names) == len(EMBEDDED_POOL_NAMES), "内嵌物品名称存在重复"
print("CSV池:", len(pool_names))

# ============================================================
# 3. 解析JSON（递归找所有含 marketHashName 的节点, 平台显式匹配）
# ============================================================
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def walk_items(obj, out):
    if isinstance(obj, dict):
        if "marketHashName" in obj:
            out.append(obj)
        for v in obj.values():
            walk_items(v, out)
    elif isinstance(obj, list):
        for v in obj:
            walk_items(v, out)


def extract_platform_rows(item):
    """优先 response.data, 否则找item下第一个'元素为含platform键的dict'的列表"""
    resp = item.get("response")
    if isinstance(resp, dict) and isinstance(resp.get("data"), list):
        return resp["data"]
    for v in item.values():
        if isinstance(v, list) and v and isinstance(v[0], dict) and "platform" in v[0]:
            return v
    return []


def norm_platform(p):
    if p is None:
        return None
    key = str(p).strip().upper()
    return CONFIG["PLATFORM_ALIAS"].get(key, key.lower())


def load_market_items(path):
    """steamdt物品列表: 支持原始JSON / 转换后的CSV(marketHashName,raw两列)"""
    if str(path).lower().endswith(".csv"):
        cdf = read_csv_any(path)
        if "marketHashName" not in cdf.columns or "raw" not in cdf.columns:
            raise Exception("CSV行情文件需包含 marketHashName 和 raw 两列")
        items = []
        for _, r in cdf.iterrows():
            try:
                resp = json.loads(r["raw"])
            except Exception:
                continue
            items.append({"marketHashName": r["marketHashName"], "response": resp})
        return items
    data = load_json(path)
    items = []
    walk_items(data, items)
    return items


def parse_file(path, source_label):
    """返回 records: [{name, platform, sell, buy, sell_count, buy_count, source}]"""
    items = load_market_items(path)
    records = []
    for it in items:
        name = it.get("marketHashName")
        if not name:
            continue
        for row in extract_platform_rows(it):
            if not isinstance(row, dict):
                continue
            plat = norm_platform(row.get("platform"))
            if plat is None:
                continue
            sell = fnum(row.get("sellPrice"))
            buy = fnum(row.get("biddingPrice"))
            sc = fint(row.get("sellCount"))
            bc = fint(row.get("biddingCount"))
            if not (sell or buy or sc or bc):
                continue  # 该平台无有效行情(全零占位行)
            records.append({
                "name": name,
                "platform": plat,
                "sell": sell,
                "buy": buy,
                "sell_count": sc,
                "buy_count": bc,
                "source": source_label,
            })
    print(f"{source_label}: 解析到 {len(items)} 个物品, {len(records)} 条有效平台行情")
    return records

# ============================================================
# 3.5 SteamDT开放平台实时拉取(USE_LIVE=True时使用)
#     文档: https://doc.steamdt.com  鉴权: Bearer {API_KEY}
#     限频: price/batch 1次/分 | item/kline 120次/分
# ============================================================
def api_headers():
    return {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}


_HTTP = None


def http_session():
    """复用TCP/TLS连接: 643次K线请求每次省一次握手(约100-300ms)"""
    global _HTTP
    if _HTTP is None:
        _HTTP = requests.Session()
        _HTTP.headers.update(api_headers())
    return _HTTP


AUTH_CODES = (4001, 4002, 2007)  # 2007=请先申请ApiKey(实测无效key返回此码)


def is_auth_error(code, msg):
    return code in AUTH_CODES or "key" in str(msg).lower()


def api_call(url, payload=None, timeout=30):
    """统一POST调用; 鉴权错误立即终止, 限频等61s重试, 其他错误等10s重试"""
    last_err = None
    for attempt in range(1, CONFIG["API_MAX_RETRY"] + 1):
        try:
            resp = requests.post(url, headers=api_headers(), json=payload, timeout=timeout)
            data = resp.json()
            if data.get("success"):
                return data
            code, msg = data.get("errorCode"), str(data.get("errorMsg", ""))
            if is_auth_error(code, msg):
                raise Exception(f"API_KEY无效或未授权(errorCode={code}: {msg}), 请到SteamDT「个人中心-API管理」检查")
            if resp.status_code == 429 or "频" in msg or "rate" in msg.lower():
                last_err, wait = f"触发限频({code}: {msg})", CONFIG["BATCH_SLEEP"]
            else:
                last_err, wait = f"errorCode={code} {msg}", 10
        except requests.RequestException as e:
            last_err, wait = f"网络异常: {e}", 10
        except json.JSONDecodeError as e:
            last_err, wait = f"响应非JSON: {e}", 10
        if attempt < CONFIG["API_MAX_RETRY"]:
            print(f"  [重试{attempt}/{CONFIG['API_MAX_RETRY']}] {last_err} -> 等待{wait}s", flush=True)
            time.sleep(wait)
    raise Exception(f"API调用失败: {last_err}")


def fetch_live_prices(names):
    """批量拉实时行情 -> 存档为与steamdt_raw.json相同结构并返回路径
    限频1批/分; 单批过大(100002)自动折半; 同批连续失败>=5次则终止"""
    url = f"{CONFIG['API_BASE']}/open/cs2/v1/price/batch"
    bs = CONFIG["BATCH_SIZE"]
    pending = [{"names": names[i:i + bs], "tries": 0}
               for i in range(0, len(names), bs)][::-1]
    est = len(pending) - 1
    print(f"[实时行情] {len(names)}个物品 分{len(pending)}批 (限频1批/分, 预计≈{est}分钟左右)", flush=True)
    out, n_call = [], 0
    while pending:
        item = pending.pop()
        chunk = item["names"]
        if n_call > 0:
            time.sleep(CONFIG["BATCH_SLEEP"])
        n_call += 1
        try:
            resp = http_session().post(url, json={"marketHashNames": chunk}, timeout=60)
            data = resp.json()
        except (requests.RequestException, json.JSONDecodeError, ValueError) as e:
            print(f"  批#{n_call}: 网络/解析异常({e}) 将重试", flush=True)
            _requeue(item, pending)
            continue
        if data.get("success"):
            rows = data.get("data") or []
            out += [{"marketHashName": it.get("marketHashName"),
                     "response": {"success": True, "data": it.get("dataList") or []}}
                    for it in rows]
            print(f"  批#{n_call}: +{len(rows)}个 (待处理{len(pending)}批)", flush=True)
            continue
        code, msg = data.get("errorCode"), str(data.get("errorMsg", ""))
        if is_auth_error(code, msg):
            print(f"\n[错误] API_KEY无效或未授权(errorCode={code}: {msg})")
            print("       请到 SteamDT「个人中心-API管理」申请/检查key后重试")
            sys.exit(1)
        if code == 100002 and len(chunk) > 10:  # 单批过大 -> 折半重试
            mid = len(chunk) // 2
            pending.append({"names": chunk[:mid], "tries": item["tries"]})
            pending.append({"names": chunk[mid:], "tries": item["tries"]})
            print(f"  批#{n_call}: {len(chunk)}个被拒(参数错误) -> 拆两批({mid}+{len(chunk) - mid})", flush=True)
            continue
        print(f"  批#{n_call}: errorCode={code} {msg} 将重试", flush=True)
        _requeue(item, pending)
    path = f"steamdt_live_{now_cn():%Y-%m-%d}.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    print(f"[实时行情] 共{len(out)}个, 已存档 {path}", flush=True)
    return path


def _requeue(item, pending):
    item["tries"] += 1
    if item["tries"] >= 5:
        print(f"\n[错误] 同一批连续失败{item['tries']}次, 终止运行(批次含{len(item['names'])}个物品)")
        sys.exit(1)
    pending.append(item)


def fetch_live_klines(names):
    """逐个拉K线；每25个存盘，设CS2_RESUME_KLINE=1可续跑当天未完成任务。"""
    url = f"{CONFIG['API_BASE']}/open/cs2/item/v1/kline"
    path = f"kline_live_{now_cn():%Y-%m-%d}.json"
    out, fails = {}, 0
    if os.environ.get("CS2_RESUME_KLINE", "") == "1" and os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                cached = json.load(handle)
            if isinstance(cached, dict):
                out = {n: v for n, v in cached.items() if n in names and isinstance(v, list)}
                print(f"[K线续跑] 已复用当天缓存{len(out)}个；如需最新K线请不要设置CS2_RESUME_KLINE")
        except (OSError, ValueError):
            pass

    def checkpoint():
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(out, handle, ensure_ascii=False)
        os.replace(tmp, path)

    pending = [n for n in names if n not in out]
    print(f"[K线拉取] 待处理{len(pending)}个, 间隔{CONFIG['KLINE_SLEEP']}s；网络耗时另计")
    for index, n in enumerate(tqdm(pending, desc="K线", unit="个"), 1):
        got = False
        for attempt in range(2):
            try:
                data = http_session().post(
                    url, json={"marketHashName": n, "type": CONFIG["KLINE_TYPE"]},
                    timeout=30).json()
                if data.get("success") and isinstance(data.get("data"), list):
                    out[n] = data["data"]
                    got = True
                    break
                if is_auth_error(data.get("errorCode"), data.get("errorMsg")):
                    checkpoint()
                    print(f"\n[错误] API_KEY无效(errorCode={data.get('errorCode')}: {data.get('errorMsg')}), 终止K线拉取")
                    sys.exit(1)
            except (requests.RequestException, json.JSONDecodeError, ValueError):
                pass
            if attempt == 0:
                time.sleep(5)
        if not got:
            fails += 1
        if index % 25 == 0:
            checkpoint()
        time.sleep(CONFIG["KLINE_SLEEP"])
    checkpoint()
    print(f"[K线拉取] 成功{len(out)} 失败{fails}，已存档{path}")
    return out


if USE_LIVE:
    steam_src = fetch_live_prices(sorted(pool_names))
    recs_steamdt = parse_file(steam_src, "steamdt-live")
else:
    steam_src = steam_file
    recs_steamdt = parse_file(steam_src, "steamdt")

if youpin_file:
    print("检测到独立YOUPIN文件, 其YOUPIN行情将覆盖steamdt内同名数据")
    recs_youpin = parse_file(youpin_file, "youpin")
else:
    print("未提供youpin_result.json, 直接使用steamdt内的YOUPIN行情")
    recs_youpin = []

# ============================================================
# 4. 价格单位归一（auto: 两文件悠悠价交叉验证优先, 中位数兜底）
# ============================================================
def price_values(records, plat=None):
    vals = []
    for r in records:
        if plat and r["platform"] != plat:
            continue
        for k in ("sell", "buy"):
            v = r[k]
            if v is not None and v == v and v > 0:
                vals.append(v)
    return sorted(vals)


def median(vals):
    return vals[len(vals) // 2] if vals else None


unit_mode = str(CONFIG["PRICE_UNIT"]).lower()
div_steamdt = div_youpin = 1

if unit_mode == "cents":
    div_steamdt = div_youpin = 100
elif unit_mode == "auto":
    # --- 4a. 交叉验证: 同一物品在两个文件里的悠悠价对比 ---
    y_by_name_a, y_by_name_b = {}, {}
    for r in recs_steamdt:
        if r["platform"] == "youpin" and r["sell"] and r["sell"] > 0:
            y_by_name_a.setdefault(r["name"], []).append(r["sell"])
    for r in recs_youpin:
        if r["platform"] == "youpin" and r["sell"] and r["sell"] > 0:
            y_by_name_b.setdefault(r["name"], []).append(r["sell"])
    common = sorted(set(y_by_name_a) & set(y_by_name_b))
    ratios = []
    for n in common:
        for pa in y_by_name_a[n]:
            for pb in y_by_name_b[n]:
                if pa > 0 and pb > 0:
                    ratios.append(pb / pa)
    ratios.sort()
    if len(ratios) >= 2 and ratios[len(ratios) // 2] >= 90:
        div_youpin = 100  # youpin文件是"分", steamdt是"元"
        basis = f"交叉验证 ratio≈{ratios[len(ratios)//2]:.0f} ({len(common)}个重合物品)"
    elif len(ratios) >= 2 and ratios[len(ratios) // 2] <= 1 / 90:
        div_steamdt = 100
        basis = f"交叉验证 ratio≈{ratios[len(ratios)//2]:.4f} ({len(common)}个重合物品)"
    else:
        # --- 4b. 兜底: 中位数>1000判为"分" ---
        basis = "中位数兜底(重合样本不足)"
        m1, m2 = median(price_values(recs_steamdt)), median(price_values(recs_youpin))
        div_steamdt = 100 if (m1 and m1 > 1000) else 1
        div_youpin = 100 if (m2 and m2 > 1000) else 1
    print(f"[单位检测] 模式=auto | {basis} | steamdt÷{div_steamdt} | youpin÷{div_youpin}")
else:
    print(f"[单位检测] 模式={unit_mode} | 全部÷{div_steamdt}")

for r in recs_steamdt:
    if div_steamdt != 1:
        if r["sell"] is not None:
            r["sell"] /= div_steamdt
        if r["buy"] is not None:
            r["buy"] /= div_steamdt
for r in recs_youpin:
    if div_youpin != 1:
        if r["sell"] is not None:
            r["sell"] /= div_youpin
        if r["buy"] is not None:
            r["buy"] /= div_youpin

# ============================================================
# 5. 合并（白名单过滤; steam已剔除; youpin文件优先覆盖同平台数据）
# ============================================================
item_map = {}  # name -> {platform: {...}}

for rec in (recs_steamdt + recs_youpin):
    if rec["platform"] in CONFIG["EXCLUDE_PLATFORMS"]:
        continue
    name = rec["name"]
    if name not in pool_names:
        continue
    slot = item_map.setdefault(name, {})
    slot[rec["platform"]] = {  # 后处理的youpin文件覆盖steamdt里的同平台行
        "sell": rec["sell"], "buy": rec["buy"],
        "sell_count": rec["sell_count"], "buy_count": rec["buy_count"],
        "source": rec["source"],
    }

# ---- K线/回撤指标: live拉取优先, 池CSV的kline列兜底 ----
kline_raw = {}
if "kline" in pool.columns:
    for _, r in pool.iterrows():
        v = r.get("kline")
        if isinstance(v, str) and v.startswith("["):
            try:
                kline_raw[r["marketHashName"]] = json.loads(v)
            except Exception:
                pass
    print(f"池CSV自带K线: {len(kline_raw)}个")
if USE_LIVE:
    kline_raw.update(fetch_live_klines(sorted(pool_names)))
else:
    offline_kline = os.environ.get("CS2_KLINE_FILE", "").strip()
    if offline_kline:
        kline_file_data = load_json(offline_kline)
        if not isinstance(kline_file_data, dict):
            raise ValueError("CS2_KLINE_FILE应是{物品名: K线数组}格式的JSON")
        kline_raw.update({n: v for n, v in kline_file_data.items()
                          if n in pool_names and isinstance(v, list)})
        print(f"本地K线文件: {offline_kline} | 已读取{len(kline_raw)}个物品")
print(f"K线数据合计: {len(kline_raw)}个")


def kline_features(raw):
    """从小时K线提取跌速变化与数据质量；按真实时间定位窗口，而非假定每根恰好1小时。"""
    result = {"valid": False, "age_hours": None, "points": 0, "changes24": 0,
              "r12_prev": None, "r12_now": None, "r24_prev": None, "r24_now": None,
              "dd48": None, "rebound48": None, "new_low_depth": None,
              "down_prev24": None, "down_now24": None,
              "dd120": None, "r120": None, "trend_720": None}
    if not isinstance(raw, list):
        return result
    by_ts = {}
    for bar in raw:
        if not isinstance(bar, (list, tuple)) or len(bar) < 5:
            continue
        try:
            ts = float(bar[0])
            close = float(bar[2])
            if ts > 1e12:  # 同时兼容毫秒时间戳
                ts /= 1000
            if not math.isfinite(ts) or not math.isfinite(close) or ts <= 0 or close <= 0:
                continue
            by_ts[int(ts)] = close
        except (TypeError, ValueError, OverflowError):
            continue
    points = sorted(by_ts.items())
    result["points"] = len(points)
    if len(points) < CONFIG["KLINE_MIN_POINTS"]:
        return result
    times = [t for t, _ in points]
    end = times[-1]
    price = by_ts[end]
    result["age_hours"] = round((now_cn().timestamp() - end) / 3600, 2)

    def close_hours_ago(hours):
        target = end - hours * 3600
        ix = bisect_right(times, target) - 1
        if ix < 0 or target - times[ix] > CONFIG["KLINE_MAX_GAP_HOURS"] * 3600:
            return None
        return points[ix][1]

    c12, c24, c48 = (close_hours_ago(h) for h in (12, 24, 48))
    if c24 is None:  # 不足24小时的窗口无从判断此前跌势
        return result
    result["valid"] = True
    result["r12_now"] = price / c12 - 1 if c12 else None
    result["r12_prev"] = c12 / c24 - 1 if c12 else None
    result["r24_now"] = price / c24 - 1
    result["r24_prev"] = c24 / c48 - 1 if c48 else None

    # 5日窗口: 回撤位置按120h回撤(策略周期1-5天, 48h只覆盖2天)
    win_h = CONFIG["KLINE_WINDOW_HOURS"] * 3600
    last_win = [(t, c) for t, c in points if t >= end - win_h]
    if last_win:
        result["dd120"] = max(0.0, 1 - price / max(c for _, c in last_win))
    c120 = close_hours_ago(CONFIG["KLINE_WINDOW_HOURS"])
    if c120:
        result["r120"] = price / c120 - 1
    # 30日趋势: 识别长期阴跌, 避免把下跌中继误判为左侧
    c720 = close_hours_ago(CONFIG["KLINE_TREND_HOURS"])
    if c720:
        result["trend_720"] = price / c720 - 1

    last48 = [(t, c) for t, c in points if t >= end - 48 * 3600]
    if last48:
        result["dd48"] = max(0.0, 1 - price / max(c for _, c in last48))
        result["rebound48"] = max(0.0, price / min(c for _, c in last48) - 1)
    recent12 = [c for t, c in points if t >= end - 12 * 3600]
    prev12 = [c for t, c in points if end - 24 * 3600 <= t < end - 12 * 3600]
    if recent12 and prev12:
        result["new_low_depth"] = 1 - min(recent12) / min(prev12)

    last24 = [(t, c) for t, c in points if t >= end - 24 * 3600]
    result["changes24"] = sum(abs(b / a - 1) >= 0.0005
                               for (_, a), (_, b) in zip(last24, last24[1:]))
    previous_down = []
    current_down = []
    for (t1, c1), (t2, c2) in zip(points, points[1:]):
        if t2 <= end - 48 * 3600 or t2 > end:
            continue
        hourly_drop = max(0.0, 1 - c2 / c1) / max(1.0, (t2 - t1) / 3600)
        if t2 <= end - 24 * 3600:
            previous_down.append(hourly_drop)
        else:
            current_down.append(hourly_drop)
    if len(previous_down) >= 8 and len(current_down) >= 8:
        result["down_prev24"] = sum(previous_down) / len(previous_down)
        result["down_now24"] = sum(current_down) / len(current_down)
    return result


kline_metrics = {name: kline_features(k) for name, k in kline_raw.items()}

matched = set(item_map)
missing = sorted(pool_names - matched)
print(f"匹配成功: {len(matched)} / 池内 {len(pool_names)} | 未匹配: {len(missing)}")

rows = []
for name, plats in item_map.items():
    yp = plats.get(CONFIG["SCORE_PLATFORM"], {})
    ys, yb = yp.get("sell"), yp.get("buy")
    ysc, ybc = yp.get("sell_count", 0), yp.get("buy_count", 0)
    depth_ratio = (ybc / ysc) if ysc else None
    tight_ratio = (yb / ys) if (ys and yb) else None
    row = {
        "marketHashName": name,
        "youpin_sell": ys,
        "youpin_buy": yb,
        "youpin_sell_count": ysc,
        "youpin_buy_count": ybc,
        "挂买/挂卖": round(depth_ratio, 4) if depth_ratio is not None else None,
        "买一/卖一": round(tight_ratio, 4) if tight_ratio is not None else None,
        "数据来源": yp.get("source", ""),
    }
    for p in CONFIG["REFERENCE_PLATFORMS"]:  # 参考列,不入分
        pr = plats.get(p, {})
        row[f"{p}_sell"] = pr.get("sell")
        row[f"{p}_sell_count"] = pr.get("sell_count", 0)
    notes = []
    if not yp:
        notes.append("无悠悠数据")
    else:
        if not ys:
            notes.append("缺卖一价")
        if not yb:
            notes.append("缺买一价")
        if tight_ratio and tight_ratio > CONFIG["BID_PREMIUM_FLAG"]:
            notes.append("溢价求购:买一高于卖一,多为定向收特定图案,买一价不可作退出价")
    row["备注"] = ";".join(notes)
    rows.append(row)

df = pd.DataFrame(rows)

if df.empty:
    raise RuntimeError("没有匹配到可评分的物品，请检查实时接口返回的平台字段与物品名称")

# ============================================================
# 6. V2.6：评分 + 左侧信号分级（不是未来收益预测）
# ============================================================
def stepped(value, steps, base):
    if value is None or (isinstance(value, float) and not math.isfinite(value)):
        return 0
    for threshold, points in steps:
        if value >= threshold:
            return points
    return base


def depth_score(row):
    sell_count, buy_count = row["youpin_sell_count"], row["youpin_buy_count"]
    if not sell_count or sell_count <= 0:
        return 0
    score = stepped(buy_count / sell_count, CONFIG["DEPTH_STEPS"], CONFIG["DEPTH_BASE"])
    return min(score, CONFIG["DEPTH_LOWBID_CAP"]) if buy_count < CONFIG["DEPTH_MIN_BIDS"] else score


def tightness_score(row):
    sell, buy = row["youpin_sell"], row["youpin_buy"]
    if not sell or not buy or sell <= 0 or buy <= 0:
        return 0
    ratio = buy / sell
    if ratio > CONFIG["BID_PREMIUM_FLAG"]:
        return 0  # 特殊求购可能只收特定图案，不视作可执行的退出报价
    return stepped(min(ratio, 1.0), CONFIG["TIGHTNESS_STEPS"], CONFIG["TIGHTNESS_BASE"])


def liquidity_score(row):
    count = row["youpin_sell_count"]
    if not count or count <= 0:
        return 0
    return stepped(math.log10(count), CONFIG["LIQ_STEPS"], CONFIG["LIQ_BASE"])


def quality_score(name):
    if "★" in name:
        return CONFIG["QUALITY_STAR"]
    if re.search(r"\bHowl\b", name) or "Dragon Lore" in name:
        return CONFIG["QUALITY_LEGEND"]
    return CONFIG["QUALITY_BASE"]


def deceleration_score(previous, recent, minimum_previous_drop, points=10):
    """跌幅从负变小才加分；前期没跌或最近跌得更快均给0分。"""
    if previous is None or recent is None or previous > -minimum_previous_drop:
        return 0.0
    improvement = max(0.0, min(1.0, (recent - previous) / abs(previous)))
    # 只是从暴跌变成较轻的暴跌，给部分分而不等同于止跌
    if recent < -0.04:
        improvement *= 0.4
    elif recent < -0.02:
        improvement *= 0.7
    return points * improvement


def technical_scores(features):
    """位置10、减速20、低点10、下行跌幅收敛5；过期和过少更新不算企稳。"""
    score = {"回撤位置分": 0, "跌速减弱分": 0, "低点企稳分": 0, "下行跌幅收敛分": 0}
    if not features["valid"] or features["age_hours"] is None:
        return score
    if features["age_hours"] > CONFIG["KLINE_STALE_HOURS"] or features["age_hours"] < -2:
        return score
    # 回撤位置: 优先5日窗口dd120, 数据不足退回dd48
    dd_for_position = features["dd120"] if features["dd120"] is not None else features["dd48"]
    score["回撤位置分"] = stepped(dd_for_position, CONFIG["POSITION_STEPS"], CONFIG["POSITION_BASE"])
    if features["changes24"] < CONFIG["KLINE_MIN_CHANGES_24H"]:
        return score  # 长时间无价格变化不当作下跌动能消失
    early = deceleration_score(features["r12_prev"], features["r12_now"],
                               CONFIG["MIN_PREV_DROP_12H"])
    medium = deceleration_score(features["r24_prev"], features["r24_now"],
                                CONFIG["MIN_PREV_DROP_24H"])
    score["跌速减弱分"] = round(early + medium)
    # 低点企稳改为联合判定: 未创新低 且 未大幅反弹 才计分
    # (原为两项独立相加: "创新低很深+反弹0%"也能拿分, 逻辑不自洽)
    new_low = features["new_low_depth"]
    new_low_score = 0
    if new_low is not None:
        if new_low <= 0.005:
            new_low_score = 6
        elif new_low <= 0.015:
            new_low_score = 4
        elif new_low <= 0.03:
            new_low_score = 2
    rebound = features["rebound48"]
    if rebound is None:
        rebound_factor = 1.0          # 无反弹数据不惩罚
    elif rebound <= CONFIG["REBOUND_FULL"]:
        rebound_factor = 1.0
    elif rebound <= CONFIG["REBOUND_HALF"]:
        rebound_factor = 0.5
    else:
        rebound_factor = 0.0          # 已反弹较远, 左侧位置基本错过
    score["低点企稳分"] = round(new_low_score * rebound_factor)
    down_prior, down_now = features["down_prev24"], features["down_now24"]
    if down_prior is not None and down_now is not None and down_prior > 0.0002:
        if down_now <= down_prior * 0.7:
            score["下行跌幅收敛分"] = 5
        elif down_now <= down_prior:
            score["下行跌幅收敛分"] = 3
    return score


def signal_label(row, features, scores):
    """观察分级: 含长期阴跌/反弹较远等排除项; 没有把'不创新低'设为入选门槛。"""
    if not row["youpin_sell"] or not row["youpin_buy"] or not row["youpin_sell_count"]:
        return "悠悠盘口不足"
    # 异常溢价求购单不剔除技术候选；只单独标注成交风险，资金暂缓。
    if not features["valid"]:
        return "K线不足"
    if features["age_hours"] is None or features["age_hours"] > CONFIG["KLINE_STALE_HOURS"] or features["age_hours"] < -2:
        return "K线过期"
    if features["changes24"] < CONFIG["KLINE_MIN_CHANGES_24H"]:
        return "价格更新稀疏"
    trend = features.get("trend_720")
    if trend is not None and trend <= CONFIG["LONG_DECLINE_30D"]:
        return "长期阴跌"  # 30日阴跌品, 短期减速多为下跌中继, 不出候选
    if features["rebound48"] is not None and features["rebound48"] > CONFIG["MAX_CANDIDATE_REBOUND"]:
        return "反弹较远"
    if features["dd48"] is None or features["dd48"] < CONFIG["MIN_CANDIDATE_DRAWDOWN"]:
        return "回撤不明显"
    mom = scores["跌速减弱分"]
    if mom >= 12 and scores["低点企稳分"] >= 5:
        return "减速明显"
    if mom >= CONFIG["MIN_CANDIDATE_MOMENTUM"]:
        return "左侧观察"
    accelerating_12h = (features["r12_prev"] is not None and features["r12_now"] is not None
                        and features["r12_now"] < features["r12_prev"] - 0.005)
    accelerating_24h = (features["r24_prev"] is not None and features["r24_now"] is not None
                        and features["r24_now"] < features["r24_prev"] - 0.005)
    if accelerating_12h or accelerating_24h:
        return "可能加速下跌"
    return "暂未减速"


df["盘口深度分"] = df.apply(depth_score, axis=1)          # /20
df["盘口紧密度分"] = df.apply(tightness_score, axis=1)    # /15
df["流动性分"] = df.apply(liquidity_score, axis=1)        # /15
df["品质分"] = df["marketHashName"].apply(quality_score)  # /5

score_map, label_map = {}, {}
for _, item in df.iterrows():
    name = item["marketHashName"]
    features = kline_metrics.get(name) or kline_features(None)
    scores = technical_scores(features)
    score_map[name] = scores
    label_map[name] = signal_label(item, features, scores)
for column in ("回撤位置分", "跌速减弱分", "低点企稳分", "下行跌幅收敛分"):
    df[column] = df["marketHashName"].map(lambda n: score_map[n][column])
for column, feature in (("前12h涨跌", "r12_prev"), ("近12h涨跌", "r12_now"),
                        ("前24h涨跌", "r24_prev"), ("近24h涨跌", "r24_now"),
                        ("48h回撤", "dd48"), ("距48h低点", "rebound48"),
                        ("5日回撤", "dd120"), ("5日涨跌", "r120"), ("30日涨跌", "trend_720"),
                        ("K线时效(小时)", "age_hours"), ("24h变价次数", "changes24")):
    df[column] = df["marketHashName"].map(lambda n: (kline_metrics.get(n) or {}).get(feature))
df["信号状态"] = df["marketHashName"].map(label_map)
df["报价风险"] = df.apply(
    lambda r: "买一异常溢价(不作为可成交退出价)"
    if r["youpin_sell"] and r["youpin_buy"]
    and r["youpin_buy"] > r["youpin_sell"] * CONFIG["BID_PREMIUM_FLAG"] else "",
    axis=1)
score_columns = ("盘口深度分", "盘口紧密度分", "流动性分", "品质分",
                 "回撤位置分", "跌速减弱分", "低点企稳分", "下行跌幅收敛分")
df["V2.6总分"] = df[list(score_columns)].sum(axis=1)
assert df["V2.6总分"].between(0, 100).all(), "评分超出0–100，请检查分项配置"

# 优先展示有跌速减弱证据的候选，同行内按分数/盘口排序。
df["_signal_priority"] = df["信号状态"].map({"减速明显": 2, "左侧观察": 1}).fillna(0)
df = df.sort_values(["_signal_priority", "V2.6总分", "挂买/挂卖", "买一/卖一"],
                    ascending=False, na_position="last").drop(columns="_signal_priority").reset_index(drop=True)
df.insert(0, "排名", df.index + 1)
print("信号状态统计:", df["信号状态"].value_counts().to_dict())

# ============================================================
# 7. 资金档位 + 总预算封顶分配
# ============================================================
def fund_from_score(score):
    for th, amount in sorted(CONFIG["FUND_TIERS"], key=lambda x: -x[0]):
        if score >= th:
            return amount
    return 0


df["建议资金"] = df.apply(
    lambda row: fund_from_score(row["V2.6总分"])
    if row["信号状态"] in ("减速明显", "左侧观察")
    and not row["报价风险"] else 0,
    axis=1)

if CONFIG["USE_BUDGET_CAP"]:
    remaining = CONFIG["TOTAL_BUDGET"]
    alloc = []
    for v in df["建议资金"]:
        a = min(v, remaining)
        alloc.append(a)
        remaining -= a
    df["实际获批"] = alloc
    print(f"总预算 {CONFIG['TOTAL_BUDGET']:,} | 已分配 {sum(alloc):,} | 剩余 {remaining:,}")
else:
    df["实际获批"] = df["建议资金"]

# ============================================================
# 8. 输出: Excel(美化) + Markdown + 每日快照
# ============================================================
today = now_cn().strftime("%Y-%m-%d")
xlsx = f"CS2_V26_Daily_Report_{today}.xlsx"
md = f"CS2_V26_Daily_Report_{today}.md"
snap = f"snapshot_CS2_V26_{today}.csv"

price_cols = [c for c in df.columns if c.endswith("_sell") or c.endswith("_buy")]
for c in price_cols:
    df[c] = pd.to_numeric(df[c], errors="coerce").round(2)

missing_df = pd.DataFrame({"marketHashName": missing})

with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
    df.to_excel(w, index=False, sheet_name="选品报告")
    missing_df.to_excel(w, index=False, sheet_name="未匹配清单")
    ws = w.sheets["选品报告"]

    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "C2"

    for i, col in enumerate(df.columns, 1):
        width = max(len(str(col)) * 1.6,
                    df[col].astype(str).str.len().max() if len(df) else 10, 10)
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 42)

    score_col_idx = list(df.columns).index("V2.6总分") + 1
    letter = get_column_letter(score_col_idx)
    ws.conditional_formatting.add(
        f"{letter}2:{letter}{len(df) + 1}",
        ColorScaleRule(start_type="min", start_color="F8696B",
                       mid_type="percentile", mid_value=50, mid_color="FFEB84",
                       end_type="max", end_color="63BE7B"))

with open(md, "w", encoding="utf-8") as f:
    f.write(f"# CS2 LeftSide V2.6 Daily（1–5天跌速减弱观察版）\n\n")
    f.write(f"- 日期: {today}\n")
    f.write(f"- 池内: {len(pool_names)} | 匹配: {len(matched)} | 未匹配: {len(missing)}\n")
    f.write(f"- 单位: steamdt÷{div_steamdt} / youpin÷{div_youpin}\n")
    f.write(f"- 预算: {CONFIG['TOTAL_BUDGET']:,} | 已分配: {int(df['实际获批'].sum()):,}\n\n")
    f.write("- 评分: 盘口50 + 品质5 + 5日回撤位置10 + 跌速减弱20 + 低点企稳10 + 下行跌幅收敛5\n")
    f.write("- 本版改动: 回撤位置改120h窗口; 低点企稳改联合判定; 流动性log化降相关; 新增30日长期阴跌排除\n")
    f.write("- 信号仅为1–5天左侧候选；异常溢价求购不剔除候选，但不模拟拨款。\n")
    f.write("- 分数尚未回测，资金为模型模拟分配，非交易建议；下行收敛只代表价格变化，无成交量证据。\n")
    f.write(f"- 状态统计: {df['信号状态'].value_counts().to_dict()}\n\n")
    f.write(df.head(CONFIG["MD_TOP_N"]).to_string())

if CONFIG["SAVE_SNAPSHOT"]:
    df.to_csv(snap, index=False, encoding="utf-8-sig")  # 积累历史,以后算真实回撤

print("=" * 24)
print("V2.6 完成（1–5天左侧观察版" + ("·API拉取" if USE_LIVE else "·本地文件") + ")")
print("=" * 24)
pd.set_option("display.width", 200)
pd.set_option("display.max_columns", None)
pd.set_option("display.unicode.east_asian_width", True)
print(df.head(20))

if missing:
    print(f"\n[提示] {len(missing)} 个池内物品未匹配到行情, 见Excel『未匹配清单』: ")
    print("  " + ", ".join(missing[:10]) + ("..." if len(missing) > 10 else ""))

for f_out in [xlsx, md] + ([snap] if CONFIG["SAVE_SNAPSHOT"] else []):
    print("已保存:", f_out)

# Colab浏览器有时会阻止连续下载：将全部结果打包为一次下载。
if IN_COLAB:
    from zipfile import ZipFile, ZIP_DEFLATED
    archive = f"CS2_V26_Daily_Report_{today}.zip"
    with ZipFile(archive, "w", compression=ZIP_DEFLATED) as bundle:
        for report_path in [xlsx, md] + ([snap] if CONFIG["SAVE_SNAPSHOT"] else []):
            bundle.write(report_path)
    print("已打包下载:", archive)
    colab_files.download(archive)

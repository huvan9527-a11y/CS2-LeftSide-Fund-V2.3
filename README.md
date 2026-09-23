# CS2 V2.5 + V2.6 每日双报告

`cs2_daily.py` 每天只抓取一次 643 个物品的价格和 K 线，然后用原 V2.5、V2.6 的各自参数分别生成 XLSX、Markdown 和 CSV。GitHub Actions 每天北京时间 20:00 尝试运行，并提交 `reports/` 中两版报告及当天原始快照；实际触发时间可能晚于整点。

## 加入仓库

将 `cs2_daily.py` 放在仓库根目录，将 `.github/workflows/cs2-daily.yml` 按原目录结构复制到仓库。在仓库 Settings → Secrets and variables → Actions 新增 `STEAMDT_API_KEY`。运行仓库 Actions 中的 `CS2 daily V2.5 and V2.6` → Run workflow 进行首次验证。

仓库的 Actions 需要启用；若仓库策略未允许工作流提交代码，需在 Settings → Actions → General 开启工作流的 Read and write permissions。SteamDT Key 若限定固定出口 IP，GitHub 托管 Runner 的动态 IP 可能无法通过白名单，请改用受信任的固定 IP 自托管 Runner。

## Colab 手动运行

上传脚本，在环境变量中设置 `STEAMDT_API_KEY`，执行 `!python cs2_daily.py`。如果需要浏览器自动下载，请使用同目录 `cs2_daily_colab.ipynb`，它只运行同一个脚本并下载一个双报告 ZIP。

原始 V2.5、V2.6 notebook 中包含明文 Key。不要上传它们；若此前已公开上传过 Key，应在 SteamDT 后台更换。工作流只从 GitHub Secret 读取 Key。

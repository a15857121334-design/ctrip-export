# 携程订单导出到 Excel 本地工具

这个工具在本项目目录里更新 `ctrip.xlsx`，不会上传登录态、订单表、备份或日志。

当前主流程优先使用携程后台的 **导出订单** 功能：脚本进入 `bst.ctrip.com` 订单管理页，按日期筛选，点击导出订单，读取携程导出的 `.xls`，再按你的 Excel 字段规则追加到 `ctrip.xlsx`。如果以后需要，也可以把 `order_source` 改成 `dom`，退回页面列表读取方式。

## 文件说明

- `main.py`：读取登录态，导出订单，去重、过滤、排序并写入 Excel。
- `save_login_state.py`：打开浏览器让你手动登录，然后保存本地登录态 `storage_state.json`。
- `config.yaml`：后台地址、导出按钮、字段映射、状态过滤、模板路径、补跑和网络重试设置。
- `requirements.txt`：Python 依赖。
- `setup_startup_catchup.ps1`：安装 Windows 登录后中午漏跑补跑任务。
- `unregister_startup_catchup.ps1`：取消 Windows 登录后中午漏跑补跑任务。
- `update_state.json`：本地自动更新状态，不进 Git。
- `ctrip.xlsx`：本地订单表，不进 Git。
- `ctrip_backup/`：按天备份目录，不进 Git。

## 安装依赖

在本项目目录运行：

```powershell
cd "C:\Users\15857\Documents\New project"
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## 保存登录态

运行：

```powershell
python save_login_state.py
```

浏览器打开后手动登录携程后台。脚本只保存浏览器登录态到：

```text
C:\Users\15857\Documents\New project\storage_state.json
```

不会保存账号密码。

## 检查配置

```powershell
python main.py --check-config
```

这一步只检查路径、依赖和配置，不抓取、不写 Excel。

## 预览订单

```powershell
python main.py --recent-days 7 --dry-run
```

dry-run 会导出并读取携程订单，打印新增数据预览，但不会写入 `ctrip.xlsx`。

## 正式更新

每天自动任务使用的命令是：

```powershell
python main.py --recent-days 7 --update-template --daily-backup --yes --run-slot noon
python main.py --recent-days 7 --update-template --daily-backup --yes --run-slot evening
```

手动执行时如果不加 `--yes`，脚本会在写入前再次确认。

## 开机后补跑漏跑任务

如果 12:00 或 18:00 自动更新没成功，Windows 登录后补跑任务会等待网络稳定，然后检查 `update_state.json`：

- 当天 12:00 后、18:00 前：如果今天中午没成功，自动补跑中午任务。
- 当天 18:00 后：如果今天晚间没成功，自动补跑晚间任务。
- 第二天 12:00 前：如果昨天晚间没成功，自动补跑昨天晚间任务。
- 对应时段已经成功：不再补跑，避免重复运行。

安装登录后补跑任务：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\setup_startup_catchup.ps1
```

取消登录后补跑任务：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\unregister_startup_catchup.ps1
```

补跑日志会写到：

```text
C:\Users\15857\Documents\New project\logs\startup_catchup.log
```

## 数据规则

- 只保留：`待确认、已确认、已完成、已归档`。
- 排除：`已取消、全部退订`。
- 优先用导出文件里的 `订单号` 对本次导出结果去重。
- 再用 `ctrip.xlsx` 已有行做兜底去重，避免重复录入旧订单。
- 写入后按下单日期排序。
- `日期` 来自 `下单时间`，`团期` 来自 `出行日期`。
- `卖价/优惠后卖价` 来自 `合同价`，`结算/加返后结算` 来自 `结算价`。
- `利润` 列继续使用公式：`=IF(I{row}="","",I{row}-K{row})`。

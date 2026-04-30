# 携程订单抓取到 Excel 本地工具

这个工具会读取桌面上的 `订单.xlsx` 作为只读模板，从你有权限访问的携程后台订单页抓取订单，并生成桌面文件 `ctrip.xlsx`。原始 `订单.xlsx` 不会被修改。

## 文件说明

- `config.yaml`：后台 URL、页面选择器、字段映射、模板路径和输出路径。
- `save_login_state.py`：打开浏览器让你手动登录，然后保存本地登录态 `storage_state.json`。
- `main.py`：读取登录态，按日期范围抓取订单，生成 Excel。
- `requirements.txt`：需要安装的 Python 依赖。

## 1. 安装依赖

在本项目目录运行：

```powershell
cd "C:\Users\15857\Documents\New project"
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## 2. 修改配置

打开 `config.yaml`，至少修改：

- `order_page_url`：改成真实的携程后台订单管理页地址。
- `selectors.table_rows`：订单列表每一行的 CSS 选择器。
- 日期筛选选择器：`selectors.start_date_input`、`selectors.end_date_input`、`selectors.search_button`。
- 字段映射：如果页面字段名和 Excel 表头不同，修改每个字段的 `header` 或 `selector`。

如果无法确认页面字段，优先使用 `selector`。例如：

```yaml
field_mapping:
  日期:
    selector: ".order-date"
    type: "date"
    default: ""
```

如果页面是标准表格，也可以配置 `selectors.table_headers` 和每个字段的 `header`，脚本会按表头列序读取。

## 3. 保存登录态

运行：

```powershell
python save_login_state.py
```

脚本会打开浏览器。你手动登录携程后台后，回到终端按提示输入 `确认`，才会写入：

```text
C:\Users\15857\Documents\New project\storage_state.json
```

脚本不会保存账号密码。

## 4. 检查配置

运行：

```powershell
python main.py --check-config
```

它只检查配置和本地路径，不抓取、不生成 Excel。

## 5. 抓取预演

先运行 dry-run：

```powershell
python main.py --start-date 2026-01-01 --end-date 2026-01-31 --dry-run
```

dry-run 会进入页面抓取并打印前几行预览，但不会生成 `ctrip.xlsx`。

## 6. 正式生成 Excel

确认 dry-run 结果正常后运行：

```powershell
python main.py --start-date 2026-01-01 --end-date 2026-01-31
```

保存前脚本会显示：

- 只读模板路径：`C:\Users\15857\Desktop\订单.xlsx`
- 输出文件路径：`C:\Users\15857\Desktop\ctrip.xlsx`
- 将写入的订单行数
- 是否会覆盖已有 `ctrip.xlsx`

只有你输入 `确认` 后，才会写入 Excel。

## Excel 输出规则

- 原模板 `工作表1` 会保留在新文件里，作为样例。
- 新增工作表 `ctrip` 写入抓取订单。
- 字段顺序固定为：`日期、操作人、来源、目的地、供应商、团期、人数、卖价、优惠后卖价、结算、加返后结算、利润、备注`。
- `利润` 列自动写入公式：`=IF(I{row}="","",I{row}-K{row})`。
- `日期` 和 `团期` 会尽量写成真正日期值；人数和金额会尽量转为数字。
- 原始模板文件始终只读，不会被写回。

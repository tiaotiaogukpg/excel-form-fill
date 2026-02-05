# excel-form-fill 各脚本用法说明

## 一、整体流程

```
Excel 文件 → excel_reader 解析 → 记录列表 → fill_form / run_single → browser-use Agent 在网页中填写
```

---

## 二、各文件作用与用法

### 1. `create_sample_excel.py` — 生成示例 Excel

**作用**：生成用于测试的示例 Excel（单表、双列两种格式），方便本地调试读表与填表流程。

**用法**（无参数，直接运行）：

```bash
python create_sample_excel.py
```

**输出**：

- `sample_data.xlsx`：单表示例（姓名、学号、成绩、备注）
- `sample_data_double_column.xlsx`：双列示例（同一行左右两栏各一条记录）

---

### 2. `excel_reader.py` — 从 Excel 读成记录列表

**作用**：把任意 Excel 表解析成「表头 → 行数据」的字典列表，供 `fill_form` 或其它脚本使用。支持单表、双列布局、语义识别表头（如「姓名」「学生姓名」）。

**作为库使用**（在其它脚本里 `import`）：

```python
from pathlib import Path
from excel_reader import read_excel_to_records, records_to_task_text

# 读取 Excel → 记录列表
records = read_excel_to_records(
    Path("sample_data.xlsx"),
    sheet=None,           # 第一个工作表；也可传工作表名或索引
    double_column=None,   # None=自动检测双列，True=强制双列，False=单表
)
# records = [{"姓名": "张三", "学号": "2024001", "成绩": "85", ...}, ...]

# 转成给 Agent 看的纯文本
text = records_to_task_text(records)
# "第1条：姓名=张三，学号=2024001，成绩=85，..."
```

**命令行**：`excel_reader.py` 没有内置 CLI，一般通过 `fill_form.py` 间接使用。

---

### 3. `fill_form.py` — 按 Excel 批量填表（整表/分页）

**作用**：用 `excel_reader` 读出记录，构造自然语言任务，交给 browser-use Agent 在浏览器里打开目标 URL，逐条（或分页）填写。适合「整张 Excel 往一个录入系统里填」的场景。

**前置**：`.env` 里配置 `DEEPSEEK_API_KEY` 或 `BROWSER_USE_API_KEY`（见 `.env.example`）。

**用法**：

```bash
# 最简：指定 Excel 和录入页 URL
python fill_form.py --excel sample_data.xlsx --url "http://localhost:5173"

# 常用参数
python fill_form.py -e sample_data.xlsx -u "http://localhost:5173" \
  --sheet 0 \              # 工作表名或索引，默认第一个
  --page-size 10 \        # 每 10 条一批，每批调一次 Agent（省 token）
  --max-rows 20 \         # 只填前 20 行
  --max-steps 50 \        # 每批 Agent 最大步数
  --timeout 900 \         # 总超时秒数（默认 15 分钟）
  --instructions "先选班级再填" \   # 补充说明给 Agent
  --double-column         # 强制双列解析
  # --no-double-column    # 禁用双列，按单表解析
  # --no-import-first     # 目标页无文件导入时关闭「导入优先」
  # --selector-map map.json   # 字段→CSS 选择器 JSON，给 React/Vue 表单用
```

**分页逻辑**：设 `--page-size N` 时，会按每 N 条一批调用 Agent，每批任务里会写明「本批第 X 页/共 Y 页」「填完本页后点一次提交本页」「有下一页请翻页」。

**导入优先（browser-use 优化，通用）**：为避免 Agent 在「逐条添加」与「文件导入」之间反复尝试，脚本会：  
1）把当前 Excel 的**绝对路径**加入 Agent 的 `available_file_paths`；  
2）在任务中用**通用表述**要求：若页面有文件上传（`input type=file`）或「导入/上传/选择文件/选择 Excel」等入口，**优先**用 **upload_file(index=…, path=…)** 选择本 Excel，再点「导入」「上传」「确定」等；若无文件上传，再寻找「添加」「新增」「导入」等入口。  
这样适配**任意带文件导入的录入系统**（成绩系统、OA、CRM、后台等），不限于本项目的成绩演示页。  
**导入**：任务中会明确「每步只做一次」；选择文件后点击一次「导入」（id=importBtn）即完成导入，不要反复点击文件选择或导入按钮、避免 Agent 反复尝试。  
**提交确认**：本页填完后只需点击一次「提交修改」「提交本页」「提交全部」或「保存本页」；若出现二次确认弹窗，只点击一次「确定」或「确认」即可。  
**提交成功确认**：点击提交后，**先检查页面相关标记**（如状态列「已提交」「提交于 xxx」、页面「提交成功」「保存成功」提示等）以确认；若**发现任一标记**则视为成功，不要再次点击提交或重新填写，直接点「下一页」。若**未发现任何上述标记**，也**自动进行下一步**（直接点「下一页」或结束本页），不要反复点击提交或重新填写。  
**导入后逐页校准**：导入后按页处理：当前页填写/核对成绩 → 点一次「提交本页」→ 检查页面标记；有则视为成功并点「下一页」，无则也自动点「下一页」。每页只提交一次。  
若目标系统**没有** Excel/文件导入功能，可加 **`--no-import-first`**，关闭「导入优先」说明，只引导 Agent 寻找「添加/新增/导入」等入口逐条或批量填表。  
**表格已有行时**：若目标页（如校园时空大数据等成绩管理页）表格中**已有**班级、姓名、课程、平时成绩、考试成绩等列和行，任务中会明确「无需先找添加学生或导入；直接在对应行填入本批数据，最后点提交修改/提交本页」。Agent 会优先在已有行中按行顺序填成绩并提交，不会浪费时间找「添加学生」或「导入」入口。

**成绩比例**：若目标页面（如成绩录入系统）有「成绩计算比例」文字和「设置比例」按钮，Excel 中若有「比例/设置/权重」工作表或表头，脚本会读出比例并写入任务；否则可手动补充：  
`--instructions "若页面有「成绩计算比例」或「设置比例」按钮，请先点击「设置比例」确认或设置为平时 40%、考试 60%（或与 Excel 一致）后再填写成绩。"`

---

### 4. `run_single.py` — 单条成绩录入（含登录账号）

**作用**：只录**一名学生**的成绩（平时+考试），并支持指定登录账号。内部会构造「打开 URL → 登录（若需）→ 搜索该生 → 填平时/考试成绩 → 提交」的任务，交给 browser-use 执行。适合验证流程或只录一个人的场景。

**用法**：

```bash
# 最简（账号可从 .env 读）
python run_single.py --url "http://localhost:5173" --name "张三" --usual 85 --exam 78

# 指定登录账号（命令行优先于 .env）
python run_single.py -u "http://localhost:5173" -n "张三" --usual 85 --exam 78 \
  --user admin --password admin123

# 班级、课程、权重（可与 .env 的 DEFAULT_CLASS / DEFAULT_COURSE / USUAL_WEIGHT / EXAM_WEIGHT 配合）
python run_single.py -u "http://localhost:5173" -n "张三" --usual 85 --exam 78 \
  --class "高一(1)班" --course "语文" --user admin -P admin123
```

**账号配置**：可在 `.env` 中设置 `GRADE_ENTRY_USER`、`GRADE_ENTRY_PASSWORD`（或 `LOGIN_USER`、`LOGIN_PASSWORD`），不传 `--user`/`--password` 时使用。

---

### 5. `tools_js_inject.py` — 自定义 Agent 工具（JS 注入）

**作用**：为 browser-use Agent 提供两个自定义 action：  
- `set_value(selector, value)`：用 CSS 选择器找到元素，写值并触发 `input`/`change`/`blur`  
- `click(selector)`：用 CSS 选择器点击元素  

用于 React/Vue 等动态表单，Agent 可优先用这些工具精确填值，而不是纯靠「点哪输哪」。

**用法**：不单独运行。`fill_form.py` 在启动时会尝试 `import tools_js_inject`，若成功则把其中的 `tools` 注入给 Agent。  
若你提供 `--selector-map map.json`（字段名 → CSS selector），任务里会带上这段映射说明，引导 Agent 使用 `set_value(selector, value)` 填写。

**示例 selector-map JSON**：

```json
{
  "姓名": "#name",
  "学号": "[data-field=studentId]",
  "成绩": ".score-input"
}
```

---

### 6. `main.py` — 项目入口提示

**作用**：目前仅打印一句提示，说明本项目的入口是 `fill_form.py`、`run_single.py` 等。可直接运行：

```bash
python main.py
```

---

## 三、典型使用顺序

1. **生成示例数据**：`python create_sample_excel.py`
2. **单条录入测试（含登录）**：`python run_single.py -u "http://localhost:5173" -n "张三" --usual 85 --exam 78 -U admin -P admin123`
3. **整表/分页批量填**：`python fill_form.py -e sample_data.xlsx -u "http://localhost:5173" --page-size 10`

Excel 结构若为「双列排版」（同一行左右两栏各一条），可不加参数（自动检测）或加 `--double-column`；若希望严格按单表解析，加 `--no-double-column`。

---

## 四、故障排除与已知问题

### 1. ScreenshotWatchdog 超时（15s）

若日志中出现 `ScreenshotWatchdog.on_ScreenshotEvent timed out after 15.0s` 或 `Clean screenshot timed out`，说明 browser-use 内部截图/事件处理超时（约 15 秒为库内默认）。可尝试：

- **减小每批条数**：`--page-size 5`（每批 5 条），减轻单次 DOM/截图负载；
- **增大总超时**：`--timeout 1200`（20 分钟），避免整轮因其它延迟被提前中止。

该超时由 browser-use/bubus 事件总线控制，本脚本暂无配置项；若页面 DOM 很大或动画多，可优先减小 `--page-size`。

### 2. Agent 填错行（按姓名匹配导致判失败）

任务文案已明确要求**按行顺序**：第 1 条数据填表格第 1 行、第 2 条填第 2 行……不按「姓名」搜索。若 Judge 仍判「填了错误学生」：

- 确认当前页是否已加载出表格行（例如先点「加载示例」、翻到对应页），再让 Agent 在本页按顺序填；
- 若页面必须「先搜姓名再填」，需在 `--instructions` 中补充说明，或改用支持按姓名定位的流程（如 `run_single` 单条录入）。

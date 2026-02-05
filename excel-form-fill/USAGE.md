# excel-form-fill 各程序用处和用法

## 一、整体流程（一句话）

**Excel 为唯一真源 → Python 解析 + 强校验 + 任务生成 → browser-use Agent 在网页中逐行比对并填写。**

---

## 二、各程序用处与用法

### 1. `main.py` — 项目入口

| 用处 | 说明 |
|------|------|
| **做什么** | 打印项目说明和常用命令，不执行填表。 |
| **何时用** | 第一次接触项目时，快速知道「该跑哪个脚本」。 |

**用法：**

```bash
python main.py
```

会输出：生成示例、填表命令等提示。

---

### 2. `create_sample_excel.py` — 生成示例 Excel

| 用处 | 说明 |
|------|------|
| **做什么** | 在项目目录下生成几份示例 Excel，用于测试读表与填表。 |
| **何时用** | 本地调试、验证 `excel_reader` 和 `fill_form` 是否正常。 |

**用法：** 无参数，直接运行。

```bash
python create_sample_excel.py
```

**生成文件：**

| 文件 | 内容 |
|------|------|
| `sample_data.xlsx` | 单表：姓名、学号、成绩、备注（仅一个「成绩」列） |
| `sample_data_double_column.xlsx` | 双列排版：同一行左右两栏各一条（姓名、学号、成绩） |
| `sample_data_usual_exam.xlsx` | **平时成绩 + 考试成绩两列**：姓名、学号、平时成绩、考试成绩、备注（供自动填表使用） |

**注意：** 当前填表流程要求 Excel 必须**同时**有「平时成绩」和「考试成绩」列，否则强校验会禁止填表。用于真实填表时，请用带这两列的表，或先用 `sample_data_usual_exam.xlsx` 测试。

---

### 3. `excel_reader.py` — Phase 1：Excel → 结构化数据（库，不单独运行）

| 用处 | 说明 |
|------|------|
| **做什么** | 把 Excel 按**语义**读成「记录列表 + 列识别元数据」；只认「平时成绩」「考试成绩」列，**不猜、不复制**。 |
| **何时用** | 被 `fill_form.py` 或其它脚本 `import` 使用，一般不直接命令行运行。 |

**原则：**

- 只从表头为「平时成绩」「平时」「考试成绩」「考试」「期末成绩」等**明确语义**的列取值。
- **不会**把单独一列「成绩」自动当成平时+考试（避免填错列）。

**在代码里怎么用：**

```python
from pathlib import Path
from excel_reader import read_excel_to_records, validate_records_for_fill, records_to_task_text

# 读取 → (记录列表, 元数据)
records, meta = read_excel_to_records(Path("成绩.xlsx"), sheet=0)

# 填表前强校验：未同时识别两列则禁止填表
ok, msg = validate_records_for_fill(records, meta)
if not ok:
    print(msg)  # 输出诊断信息，并应停止
else:
    text = records_to_task_text(records)  # 转成给 Agent 看的表格文本
```

**返回值：**

- `records`：每项为 `{"姓名": "...", "平时成绩": int或None, "考试成绩": int或None, ...}`。
- `meta`：`has_usual_column`、`has_exam_column`、`raw_columns`，供校验用。

---

### 4. `fill_form.py` — Phase 2：强校验 + 任务生成 + DeepSeek + browser-use 执行（主入口）

| 用处 | 说明 |
|------|------|
| **做什么** | 用 `excel_reader` 读 Excel → **强校验**（未同时识别「平时成绩」「考试成绩」列则**禁止填表**）→ 生成任务文本 → 使用 **DeepSeek** 模型 + **browser-use** Agent 在浏览器中打开录入页并按要求填写成绩。 |
| **何时用** | 每次要「按 Excel 往网页系统填成绩」时运行；可先 `--dry-run` 看任务文案，再去掉 `--dry-run` 真正跑 Agent。 |

**前置：配置 .env（项目根目录）**

在 `excel-form-fill` 目录下创建或编辑 `.env`，填写 DeepSeek API 信息（用于 browser-use 的 LLM）：

```env
# DeepSeek API（必填，用于 Agent 理解任务并操作浏览器）
DEEPSEEK_API_KEY=sk-你的密钥
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1

# 可选：录入系统登录账号（若页面需要登录，可后续在任务中说明或由 Agent 自行处理）
GRADE_ENTRY_USER=teacher1
GRADE_ENTRY_PASSWORD=你的密码
```

- **DEEPSEEK_API_KEY**：必填。未配置时，非 `--dry-run` 运行会提示并退出。
- **DEEPSEEK_BASE_URL**：可选，默认 `https://api.deepseek.com/v1`。
- 密钥获取： [DeepSeek 开放平台](https://platform.deepseek.com/) 注册并创建 API Key。

**用法：**

```bash
# 1）先 dry-run：只读 Excel + 校验 + 打印任务，不启动浏览器、不调 Agent
python fill_form.py -e sample_data_usual_exam.xlsx -u "http://你的录入页地址" --dry-run

# 2）真正成绩录入：使用 DeepSeek + browser-use 自动打开浏览器并填表
python fill_form.py -e sample_data_usual_exam.xlsx -u "http://localhost:5173"

# 常用参数
python fill_form.py -e 成绩.xlsx -u "http://localhost:5173" \
  --sheet 0           # 工作表名或索引，默认第一个
  --page-size 10       # 每 10 条一批（任务文案分页说明用）
  --max-rows 20        # 只处理前 20 行
  --max-steps 80       # Agent 最大步数（默认 80）
  --headless           # 无头模式（不显示浏览器窗口）
  --dry-run            # 仅校验+打印任务，不调 Agent
```

**行为说明：**

1. **强校验**：若 Excel 没有同时识别到「平时成绩」与「考试成绩」列，程序会**直接退出**并在 stderr 输出诊断，**不会**自动填表。
2. **任务格式**：通过校验后，生成「姓名 \| 平时成绩(目标值) \| 考试成绩(目标值)」表格 + 比对与提交规则，作为 Agent 的 `task` 文本。
3. **`--dry-run`**：只做读取、校验和打印任务，不启动浏览器、不调 LLM。建议先用此方式确认 Excel 与表头无误。
4. **非 dry-run**：使用 **DeepSeek**（从 `.env` 读 `DEEPSEEK_API_KEY` / `DEEPSEEK_BASE_URL`）作为 LLM，启动 **browser-use** 的 Agent 和本地浏览器，打开 `-u` 指定 URL，按任务文案在页面中逐行比对并填写平时成绩、考试成绩；默认有头模式（可看到浏览器窗口），加 `--headless` 则无头运行。

---

## 三、推荐使用顺序

1. **看入口说明**：`python main.py`
2. **配置 .env**：在项目根目录设置 `DEEPSEEK_API_KEY`（见上节）。
3. **生成示例**：`python create_sample_excel.py`
4. **试填表流程（不真填）**：`python fill_form.py -e sample_data_usual_exam.xlsx -u "http://你的录入页" --dry-run`
5. **真正成绩录入**：去掉 `--dry-run`，运行 `python fill_form.py -e sample_data_usual_exam.xlsx -u "http://录入页URL"`，使用 DeepSeek + browser-use 自动打开浏览器并填表。

---

## 四、设计要点（防踩雷）

- **Excel 是唯一正确来源**：成绩以 Excel 为准；网页只做「比对 + 按需修改」。
- **必须同时有「平时成绩」「考试成绩」两列**：缺一列或只认到一列 → 强校验禁止填表，并输出诊断。
- **不猜、不复制**：不会把单列「成绩」自动当成平时+考试，避免「平时=考试」类事故。
- **Agent 是执行器**：所有判断、校验在 Python 侧完成；给 Agent 的只是「已消化好」的表格和规则。

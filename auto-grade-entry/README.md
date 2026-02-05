## 自动录入成绩（Excel → 网页 → 自动化 → 分页提交）

### 目标
- **从 Excel 成绩单中自动提取学生成绩**
- **利用大模型 + 浏览器自动化，将成绩自动录入学校成绩系统**
- **提升录入效率，减少人工操作和 token 消耗**

本仓库包含一个本地“成绩录入系统”网页（模拟学校系统）与两套自动化脚本：
- **初版**：browser-use（LLM 驱动）跑通“单个学生录入 + 提交”
- **优化版**：整页批量输入 → **一次提交**，并处理分页循环（显著降 token）

---

### 目录结构
- `web/`：成绩录入网页（静态页面）
- `automation/`
  - `extract_excel.py`：从 Excel 导出 `grades.json`
  - `run_single_browser_use.py`：初版自动化（单人录入 + 提交）
  - `run_batch_playwright.py`：优化版自动化（整页批量 + 分页 + 一次提交）
  - `run_full_pipeline.py`：一键串联（Excel → JSON → 批量分页录入）

---

### 1) 运行成绩录入网页（本地）
#### 方式 A：带登录与数据隔离（推荐）

已加入**登录系统（老师/管理员）**与**按账号隔离的数据存储**，**必须**用内置服务器启动（会同时托管 `web/` 静态页面与 `/api` 接口）：

```bash
cd "auto-grade-entry/server"
npm install
npm run dev
```

浏览器打开：`http://localhost:5173`

**注意**：不要用其它方式打开页面（如 VS Code Live Server、`python -m http.server`、直接打开 `index.html` 等），否则 `/api` 请求会落到静态服务器，返回 HTML 而非 JSON，管理员将无法看到老师数据。

初始管理员账号：`admin / admin123`（可通过环境变量 `SEED_ADMIN_USER/SEED_ADMIN_PASS` 修改）

#### 方式 B：纯静态页面（不含登录）

在 `auto-grade-entry/web` 目录启动静态服务器（任选其一）：

**Python：**
```bash
cd "auto-grade-entry/web"
python -m http.server 5173
```

然后浏览器打开：`http://localhost:5173`

---

### 1.1) 直接在网页里导入 Excel（可选）
网页已支持 **上传 Excel（.xlsx/.xls）→ 自动填表 →（可选）自动逐页提交**。

**课程判断与「信息技术」误判**：  
Excel 中课程名按以下顺序解析：**标题区显式「课程：xxx」** > **主标题括号《科目》** > **同行课程格** > **关键字推断**。  
考查/考察科目表中，模板常用「课程：信息技术」表示**考查类型**（非课程名），真实课程多为同行的「语文」或主标题《语文》。因此「信息技术」「计算机」在推断与标题区中**不参与自动采纳**，仅当用户显式写「课程：信息技术」时才作为课程名。详见 `web/app.js` 中「课程判断逻辑」注释。

注意：网页端解析 Excel 使用了 `xlsx` 的 CDN，如果你的电脑不能访问 CDN：
- 可以先用 Python 脚本 `automation/extract_excel.py` 走“Excel→JSON→自动化”方案
- 或者我把 `xlsx` 库改成放在本地 `web/vendor/`（离线可用）

---

### 2) 从 Excel 导出 JSON
要求 Excel 至少包含这些列（列名可在脚本里调整）：
- `班级`、`姓名`、`课程`、`平时成绩`、`考试成绩`

**报表格式（课程成绩报告单）**：`extract_excel.py` 按「不信版面，只信语义」解析：
- 读取整个 Sheet 已使用区域，不假设列数；识别**所有**「学生姓名」列（左栏/右栏等）。
- **按行扫描**：同一行左栏、右栏若都有人名且至少有一项成绩，均归一为一条学生记录。
- 课程名整表只读一次（标题区或外部参数），不因新表头新建课程；成绩用「姓名列 + 相对偏移」取，不写死列号。
- 过滤：表头/空姓名、成绩全空的行跳过。

运行：
```bash
cd "auto-grade-entry"
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python automation\extract_excel.py --excel "你的成绩单.xlsx" --out "automation\grades.json"
```

---

### 3) 自动化（初版：先跑通流程）
> 需要安装 Playwright 浏览器内核。

```bash
cd "auto-grade-entry"
.\.venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install
copy .env.example .env
# 编辑 .env，填入 DEEPSEEK_API_KEY
python automation\run_single_browser_use.py --url "http://localhost:5173" --name "张三"
```

---

### 4) 自动化（优化版：整页批量 + 一次提交 + 分页）
```bash
cd "auto-grade-entry"
.\.venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install
python automation\run_batch_playwright.py --url "http://localhost:5173" --grades "automation\grades.json"
```

---

### 5) 一键闭环（推荐：先提取 Excel 再批量提交）
```bash
cd "auto-grade-entry"
.\.venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install
python automation\run_full_pipeline.py --excel "你的成绩单.xlsx" --url "http://localhost:5173"
```

### 页面交互说明（给自动化用）
- 点击成绩单元格会弹出输入框
- 回车保存、ESC 取消
- 最终成绩自动计算：`最终 = 平时 + 考试`
- 只要本页有修改，“提交本页”按钮出现；提交后消失
- 分页显示（默认每页 10 人）


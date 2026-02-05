const STORAGE_KEY = "auto-grade-entry:v1";
const DATA_CHANGED_KEY = "auto-grade-entry:dataChanged";

async function apiJson(path, opts) {
  const res = await fetch(path, { credentials: "include", ...opts });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || "请求失败");
  return data;
}

async function requireLoginOrRedirect() {
  try {
    const me = await apiJson("/api/me");
    return me.user;
  } catch {
    window.location.href = "/login.html";
    return null;
  }
}

const DEFAULT_CLASSES = [];
const DEFAULT_COURSES = [];
// 班级下拉中不展示的选项（可在此加入需要隐藏的班级名）
const CLASS_EXCLUDED_FROM_OPTIONS = [];
// “添加班级”选项的 value，选中时弹出输入新班级名
const ADD_CLASS_OPTION_VALUE = "__add_class__";
// “添加课程”选项的 value，选中时弹出输入新课程名
const ADD_COURSE_OPTION_VALUE = "__add_course__";
function clampInt(v, min, max) {
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  if (Number.isNaN(n)) return null;
  const r = Math.round(n);
  return Math.min(max, Math.max(min, r));
}

function loadStateLocal() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function saveStateLocal(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

async function loadStateRemote() {
  const data = await apiJson("/api/state");
  return data.state ?? null;
}

function showSyncError(msg) {
  const text = msg || "保存到服务器失败，请检查网络或重新登录。管理员将无法看到你的数据。";
  if (els.importHint) els.importHint.textContent = text;
}

function saveState(state) {
  // 本地一份（兜底/离线），服务端一份（按账号隔离）
  saveStateLocal(state);
  // 示例模式下不写入服务端，避免示例 38 人混入持久化数据
  if (state.isSample === true) return;
  apiJson("/api/state", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ state }),
  }).catch((err) => {
    showSyncError();
  });
}

// 供提交/导入等关键操作后调用：同步到服务器并返回是否成功（失败时已提示用户）
async function saveStateToServerAndConfirm(state) {
  saveStateLocal(state);
  if (state.isSample === true) return true;
  try {
    await apiJson("/api/state", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ state }),
    });
    return true;
  } catch (e) {
    showSyncError();
    return false;
  }
}

// 默认 46 开（平时 40%，考试 60%）；表格有权重时以表格读取比例优先
const DEFAULT_GRADE_FORMULA = {
  type: "weighted",
  terms: [
    { name: "平时成绩", weight: 0.4 },
    { name: "考试成绩", weight: 0.6 },
  ],
};

function getEffectiveGradeFormula() {
  return state.gradeFormula && state.gradeFormula.terms?.length >= 2
    ? state.gradeFormula
    : DEFAULT_GRADE_FORMULA;
}

function formatGradeFormulaHint(formula) {
  const f = formula || DEFAULT_GRADE_FORMULA;
  if (!f.terms?.length) return "成绩计算比例：平时 40%，考试 60%";
  if (f.type === "weighted" && f.terms.length >= 2) {
    const p0 = Math.round((f.terms[0].weight ?? 0.4) * 100);
    const p1 = Math.round((f.terms[1].weight ?? 0.6) * 100);
    return `成绩计算比例：${f.terms[0].name || "平时"} ${p0}%，${f.terms[1].name || "考试"} ${p1}%`;
  }
  return "成绩计算比例：平时 40%，考试 60%";
}

function buildInitialState() {
  return {
    selectedClass: "",
    selectedCourse: "",
    pageSize: 10,
    pageIndex: 1,
    search: "",
    rows: [],
    isSample: false,
    autoSubmitOnImport: true,
    // 删除数据时记录时间，用于显示「数据删除于 xxx」
    deletedAt: null,
    // 成绩计算方式：从 Excel 检索到的成绩列与权重，用于计算最终成绩
    gradeFormula: null,
  };
}

const els = {
  classSelect: document.getElementById("classSelect"),
  courseSelect: document.getElementById("courseSelect"),
  searchInput: document.getElementById("searchInput"),
  pageSizeSelect: document.getElementById("pageSizeSelect"),
  tbody: document.getElementById("tbody"),
  pager: document.getElementById("pager"),
  submitBtn: document.getElementById("submitBtn"),
  submitAllBtn: document.getElementById("submitAllBtn"),
  gradeFormulaHintPanel: document.getElementById("gradeFormulaHintPanel"),
  dirtyHint: document.getElementById("dirtyHint"),
  currentUser: document.getElementById("currentUser"),
  logoutBtn: document.getElementById("logoutBtn"),
  adminBtn: document.getElementById("adminBtn"),
  changePwdBtn: document.getElementById("changePwdBtn"),
  changePwdOverlay: document.getElementById("changePwdOverlay"),
  changePwdOld: document.getElementById("changePwdOld"),
  changePwdNew: document.getElementById("changePwdNew"),
  changePwdConfirm: document.getElementById("changePwdConfirm"),
  changePwdHint: document.getElementById("changePwdHint"),
  changePwdCancel: document.getElementById("changePwdCancel"),
  changePwdSubmit: document.getElementById("changePwdSubmit"),
  editorOverlay: document.getElementById("editorOverlay"),
  editorInput: document.getElementById("editorInput"),
  excelFile: document.getElementById("excelFile"),
  excelFileLabel: document.getElementById("excelFileLabel"),
  sheetSelect: document.getElementById("sheetSelect"),
  importBtn: document.getElementById("importBtn"),
  importAndSubmitAllBtn: document.getElementById("importAndSubmitAllBtn"),
  resetBtn: document.getElementById("resetBtn"),
  importHint: document.getElementById("importHint"),
  gradeFormulaHint: document.getElementById("gradeFormulaHint"),
  autoSubmitToggle: document.getElementById("autoSubmitToggle"),
  confirmOverlay: document.getElementById("confirmOverlay"),
  importPreviewBlock: document.getElementById("importPreviewBlock"),
  importPreviewSummary: document.getElementById("importPreviewSummary"),
  importPreviewTbody: document.getElementById("importPreviewTbody"),
  importPreviewSelectAll: document.getElementById("importPreviewSelectAll"),
  importPreviewCancel: document.getElementById("importPreviewCancel"),
  importPreviewHint: document.getElementById("importPreviewHint"),
  mainTableWrap: document.getElementById("mainTableWrap"),
  inlineEditOverlay: document.getElementById("inlineEditOverlay"),
  gradeInlineInput: document.getElementById("grade-inline-input"),
  mainFooterBar: document.getElementById("mainFooterBar"),
  confirmTitle: document.getElementById("confirmTitle"),
  confirmMessage: document.getElementById("confirmMessage"),
  confirmCancelBtn: document.getElementById("confirmCancelBtn"),
  confirmOkBtn: document.getElementById("confirmOkBtn"),
  addStudentBtn: document.getElementById("addStudentBtn"),
  addStudentOverlay: document.getElementById("addStudentOverlay"),
  addStudentClass: document.getElementById("addStudentClass"),
  addStudentName: document.getElementById("addStudentName"),
  addStudentCourse: document.getElementById("addStudentCourse"),
  addStudentUsual: document.getElementById("addStudentUsual"),
  addStudentExam: document.getElementById("addStudentExam"),
  addStudentFinal: document.getElementById("addStudentFinal"),
  addStudentHint: document.getElementById("addStudentHint"),
  addStudentCancel: document.getElementById("addStudentCancel"),
  addStudentSubmit: document.getElementById("addStudentSubmit"),
  addClassOverlay: document.getElementById("addClassOverlay"),
  addClassInput: document.getElementById("addClassInput"),
  addClassHint: document.getElementById("addClassHint"),
  addClassCancel: document.getElementById("addClassCancel"),
  addClassSubmit: document.getElementById("addClassSubmit"),
  addCourseOverlay: document.getElementById("addCourseOverlay"),
  addCourseInput: document.getElementById("addCourseInput"),
  addCourseHint: document.getElementById("addCourseHint"),
  addCourseCancel: document.getElementById("addCourseCancel"),
  addCourseSubmit: document.getElementById("addCourseSubmit"),
  gradeFormulaOverlay: document.getElementById("gradeFormulaOverlay"),
  gradeFormulaUsual: document.getElementById("gradeFormulaUsual"),
  gradeFormulaExam: document.getElementById("gradeFormulaExam"),
  gradeFormulaOverlayHint: document.getElementById("gradeFormulaOverlayHint"),
  gradeFormulaCancel: document.getElementById("gradeFormulaCancel"),
  gradeFormulaSubmit: document.getElementById("gradeFormulaSubmit"),
  setGradeFormulaBtn: document.getElementById("setGradeFormulaBtn"),
};

let state = loadStateLocal() ?? buildInitialState();
saveStateLocal(state);

// editor context
let editing = null; // { rowId, field, pageIndexAtOpen }

// 提交动画（仅 UI 层）：先清空再显示已提交数据
let uiSubmittingRowIds = new Set();
let uiSubmitTimer = null;
// 本会话是否有过修改或提交（重新登录后从服务器拉取的数据视为未修改，显示「未修改」）
// 删除/提交/导入等会让本地或服务器数据变化，刷新后通过 DATA_CHANGED_KEY 恢复，不显示「未修改」
let sessionHasLocalChanges = false;

function setDataChanged() {
  sessionHasLocalChanges = true;
  localStorage.setItem(DATA_CHANGED_KEY, "1");
}

// 通用确认弹窗（返回 Promise<boolean>，确定 true，取消/关闭 false）
let confirmResolve = null;
function showConfirm({ title = "确认", message = "", okText = "确定", cancelText = "取消" }) {
  return new Promise((resolve) => {
    confirmResolve = resolve;
    if (els.confirmTitle) els.confirmTitle.textContent = title;
    if (els.confirmMessage) els.confirmMessage.textContent = message;
    if (els.confirmOkBtn) els.confirmOkBtn.textContent = okText;
    if (els.confirmCancelBtn) els.confirmCancelBtn.textContent = cancelText;
    if (els.confirmOverlay) els.confirmOverlay.style.display = "grid";
  });
}
function closeConfirm(ok) {
  if (confirmResolve) confirmResolve(!!ok);
  confirmResolve = null;
  if (els.confirmOverlay) els.confirmOverlay.style.display = "none";
}

function startSubmitFlash(rowIds) {
  if (!rowIds?.length) return;
  uiSubmittingRowIds = new Set(rowIds);
  renderAll();
  if (uiSubmitTimer) clearTimeout(uiSubmitTimer);
  uiSubmitTimer = setTimeout(() => {
    uiSubmittingRowIds.clear();
    renderAll();
  }, 350);
}

function formatTime(ts) {
  if (!ts) return "";
  const d = new Date(ts);
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
function formatDateTime(ts) {
  if (!ts) return "";
  const d = new Date(ts);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function normalizeState() {
  state.rows = state.rows.map((r) => ({
    submittedAt: r.submittedAt ?? null,
    ...r,
  }));
  if (typeof state.isSample !== "boolean") state.isSample = false;
  if (typeof state.autoSubmitOnImport !== "boolean") state.autoSubmitOnImport = true;
  if (typeof state.deletedAt !== "number") state.deletedAt = null;
  if (state.rows.length > 0) state.deletedAt = null;
  if (state.gradeFormula === undefined) state.gradeFormula = null;
}

function ensureCourseOption(course) {
  const c = String(course ?? "").trim();
  if (!c) return;
  if (!DEFAULT_COURSES.includes(c)) {
    // const 数组仍可 push；用于让下拉能显示来自文件名/表头的科目
    DEFAULT_COURSES.push(c);
  }
}

function ensureClassOption(cls) {
  const c = String(cls ?? "").trim();
  if (!c) return;
  if (CLASS_EXCLUDED_FROM_OPTIONS.includes(c)) return;
  if (!DEFAULT_CLASSES.includes(c)) {
    // 识别到的班级若不在默认列表，加入下拉以便选中
    DEFAULT_CLASSES.push(c);
  }
}

// ========== 课程判断逻辑：问题起因与解析顺序 ==========
// 【问题起因】「信息技术」误判来源于三处：
// 1) 考查/考察科目表中，表格模板用「课程：信息技术」表示考查类型（非课程名），同一行或主标题才是真实课程（如《语文》）。
// 2) 标题或文件名中出现「信息技术」「计算机」时，被关键字/括号推断当成课程名。
// 3) 规范表「科目」列、多表合并按行数取最多课程时，信息技术行数多会覆盖真实课程。
// 【解析顺序】统一为：标题区显式「课程：xxx」 > 主标题括号《科目》 > 同行课程格 > 关键字推断；考查/考察表上「信息技术」「计算机」不参与推断、不占标题区课程名，仅当用户显式写「课程：信息技术」时才采纳。

/** 推断/标题区中一律不采纳为课程名的值（考查类型等）；仅当标题区显式「课程：信息技术」时可被 extractMetaFromTitleArea 采纳 */
const EXCLUDED_FROM_INFERRED_COURSE = ["信息技术", "计算机"];

function isExcludedInferredCourse(v) {
  const s = String(v ?? "").trim();
  return EXCLUDED_FROM_INFERRED_COURSE.includes(s);
}

/** 在标题区/考查表上下文中，该值是否可作为课程名采纳（排除考查类型误判） */
function acceptAsTitleCourse(v, onlyCourseLabel) {
  if (!v || !isLikelyCourseName(v)) return false;
  return !onlyCourseLabel || !isExcludedInferredCourse(v);
}

/** 科目推断（仅从文本关键字）：1) 括号《xxx》 2) 已知学科关键字（不含信息技术/计算机，避免误匹配） */
function inferCourseFromTexts(texts) {
  const candidates = (texts || [])
    .map((t) => String(t ?? "").trim())
    .filter((t) => t);

  // 1) 优先抓 “《科目》”“（科目）” 或 “(科目)”（与表格主标题一致）；括号内为「信息技术」「计算机」不采纳，避免考查科目表误判
  for (const t of candidates) {
    const m0 = t.match(/《([^》]{1,12})》/);
    if (m0?.[1]) {
      const v = m0[1].trim();
      if (v && !v.includes("学年") && !v.includes("学期") && isLikelyCourseName(v) && !isExcludedInferredCourse(v)) return v;
    }
    const m1 = t.match(/（([^）]{1,12})）/);
    if (m1?.[1]) {
      const v = m1[1].trim();
      if (v && !v.includes("学年") && !v.includes("学期") && isLikelyCourseName(v) && !isExcludedInferredCourse(v)) return v;
    }
    const m2 = t.match(/\(([^)]{1,12})\)/);
    if (m2?.[1]) {
      const v = m2[1].trim();
      if (v && !v.includes("学年") && !v.includes("学期") && isLikelyCourseName(v) && !isExcludedInferredCourse(v)) return v;
    }
  }

  // 2) 命中常见学科关键字（短科目名优先）；排除考查类型名，仅当标题区显式「课程：信息技术」时由 extractMetaFromTitleArea 提供
  const knownShort = ["语文", "数学", "英语", "物理", "化学", "生物", "政治", "历史", "地理", "体育", "音乐", "美术"];
  const known = [
    ...knownShort,
    ...DEFAULT_COURSES.filter((k) => !knownShort.includes(k) && !EXCLUDED_FROM_INFERRED_COURSE.includes(k)),
    "程序设计", "C语言", "Python", "Java", "数据结构",
  ].filter((k) => !EXCLUDED_FROM_INFERRED_COURSE.includes(k));
  for (const t of candidates) {
    for (const k of known) {
      if (t.includes(k)) return k;
    }
  }
  return "";
}

function inferClassFromTexts(texts) {
  const raw = (texts || []).map((t) => String(t ?? "").trim()).filter((t) => t);
  const candidates = [...raw];

  // 0) 若“班级”与“xxx班”分在两格，合并后再识别：相邻为「班级/班级:/班级：」+「xxx班」则取 xxx班
  for (let i = 0; i < raw.length - 1; i++) {
    const a = raw[i];
    const b = raw[i + 1];
    if (/^班级[：:]?\s*$/.test(a) && /^.+\s*班$/.test(b)) {
      const val = b.replace(/\s+$/, "");
      if (val.length >= 2 && val.length <= 30) return val;
    }
  }

  // 1) 优先匹配常见班级格式：高一(1)班、高二(2)班、高三(3)班
  for (const t of candidates) {
    const m1 = t.match(/([高初][一二三1-3])\s*[（(]\s*(\d+)\s*[）)]\s*班/);
    if (m1?.[1] && m1?.[2]) {
      const grade = m1[1].replace(/一/g, "1").replace(/二/g, "2").replace(/三/g, "3");
      return `${grade}(${m1[2]})班`;
    }
    const m2 = t.match(/([高初][一二三1-3])\s*(\d+)\s*班/);
    if (m2?.[1] && m2?.[2]) {
      const grade = m2[1].replace(/一/g, "1").replace(/二/g, "2").replace(/三/g, "3");
      return `${grade}(${m2[2]})班`;
    }
    const m2b = t.match(/([高初][一二三1-3])([一二三四五六七八九十\d]+)\s*班/);
    if (m2b?.[1] && m2b?.[2]) {
      const grade = m2b[1].replace(/一/g, "1").replace(/二/g, "2").replace(/三/g, "3");
      const cn = m2b[2];
      const num = /^\d+$/.test(cn) ? cn : { 一: "1", 二: "2", 三: "3", 四: "4", 五: "5", 六: "6", 七: "7", 八: "8", 九: "9", 十: "10" }[cn] || "1";
      return `${grade}(${num})班`;
    }
  }

  // 2) 匹配默认班级列表中的班级
  for (const t of candidates) {
    for (const cls of DEFAULT_CLASSES) {
      if (t.includes(cls)) return cls;
    }
  }

  // 3) 匹配“班级：xxx班”“班级: abc班”等（同一格内）
  for (const t of candidates) {
    const m3 = t.match(/(?:班级|行政班|所在班)[：:\s]*([^\s，。]+班)/);
    if (m3?.[1]) return m3[1].trim();
  }

  // 4) 若出现过“班级”标签，任选一个以“班”结尾的短串当作班级名（兼容“班级:”“abc班”分两格且顺序不紧邻）
  const hasClassLabel = raw.some((t) => /^班级[：:]?\s*$/.test(t) || t === "班级");
  if (hasClassLabel) {
    const xx = raw.find((t) => /^[^\s，。]{1,30}班$/.test(t));
    if (xx) return xx;
  }

  // 5) 匹配 2024级1班 等
  for (const t of candidates) {
    const m4 = t.match(/(\d{4})\s*级\s*(\d+)\s*班/);
    if (m4?.[1] && m4?.[2]) return `${m4[1]}级${m4[2]}班`;
  }

  return "";
}

function extractTitleTextsFromAoa(aoa, maxRows = 12, maxCols = 20) {
  const out = [];
  const rmax = Math.min(maxRows, aoa?.length ?? 0);
  for (let r = 0; r < rmax; r++) {
    const row = aoa[r] || [];
    for (let c = 0; c < Math.min(maxCols, row.length); c++) {
      const s = String(row[c] ?? "").trim();
      if (s) out.push(s);
    }
  }
  return out;
}

/** 判断是否为合理的科目名（排除文档标题类文字如「成绩报告单及教学质量分析表」） */
function isLikelyCourseName(v) {
  const s = String(v ?? "").trim();
  if (!s || s.length > 20) return false;
  if (/报告单|分析表|教学质量|成绩单|成绩表|统计表|汇总表/.test(s)) return false;
  return true;
}

/** 是否为「课程」或「科目」标签（含冒号、可能分多格） */
function isCourseOrSubjectLabel(s) {
  const t = String(s ?? "").trim();
  return /^(课程|科目)\s*[：:]?\s*$/.test(t) || t === "课程" || t === "科目";
}

/** 仅从标题区文本中取「《科目》」「（科目）」括号内的科目，用于覆盖误判；括号内为「信息技术」「计算机」不采纳 */
function extractCourseFromBrackets(texts) {
  const list = (texts || []).map((t) => String(t ?? "").trim()).filter((t) => t);
  for (const t of list) {
    const m0 = t.match(/《([^》]{1,12})》/);
    if (m0?.[1]) {
      const v = m0[1].trim();
      if (v && isLikelyCourseName(v) && !isExcludedInferredCourse(v)) return v;
    }
    const m1 = t.match(/（([^）]{1,12})）/);
    if (m1?.[1]) {
      const v = m1[1].trim();
      if (v && isLikelyCourseName(v) && !isExcludedInferredCourse(v)) return v;
    }
    const m2 = t.match(/\(([^)]{1,12})\)/);
    if (m2?.[1]) {
      const v = m2[1].trim();
      if (v && isLikelyCourseName(v) && !isExcludedInferredCourse(v)) return v;
    }
  }
  return "";
}

/** 工作表名是否为「考查科目」/「考察科目」或含该字样（此类表头里「科目」多指考查类型，不是课程名，只认「课程：xxx」） */
function isSheetNameAssessmentSubject(sheetName) {
  const s = String(sheetName || "").trim();
  if (!s) return false;
  return s === "考查科目" || s === "考察科目" || (s.includes("考查") && s.includes("科目")) || (s.includes("考察") && s.includes("科目"));
}

/** 从工作表标题区（前几行）显式读取「班级：xxx」「课程：xxx」/「科目：xxx」；分多格时整行扫。考查/考察科目表只认「课程：xxx」不认「科目：xxx」 */
function extractMetaFromTitleArea(aoa, maxRows = 12, sheetName = "") {
  const onlyCourseLabel = isSheetNameAssessmentSubject(sheetName);
  let titleClass = "";
  let titleCourse = "";
  const rmax = Math.min(maxRows, aoa?.length ?? 0);
  for (let r = 0; r < rmax; r++) {
    const row = aoa[r] || [];
    for (let c = 0; c < row.length; c++) {
      const s = String(row[c] ?? "").trim();
      if (!s) continue;
      if (s.includes("班级")) {
        for (const sep of ["：", ":", " "]) {
          if (s.includes(sep)) {
            const parts = s.split(sep, 2);
            if (parts[0].trim().includes("班级") || /^班级[：:]?\s*$/.test(parts[0].trim())) {
              const v = (parts[1] || "").trim();
              if (v && !titleClass) titleClass = v;
              break;
            }
          }
        }
        if (!titleClass && /^班级[：:]?\s*$/.test(s)) continue;
        if (!titleClass) {
          const v = s.replace(/^.*?班级[：:\s]*/, "").trim();
          if (v) titleClass = v;
        }
      }
      if (s.includes("课程") || s.includes("科目")) {
        for (const sep of ["：", ":", " "]) {
          if (s.includes(sep)) {
            const parts = s.split(sep, 2);
            const left = (parts[0] || "").trim();
            if (!/^(课程|科目)\s*[：:]?\s*$/.test(left)) continue;
            if (onlyCourseLabel && left.startsWith("科目")) continue;
            const v = (parts[1] || "").trim();
            if (v && !titleCourse && acceptAsTitleCourse(v, onlyCourseLabel)) titleCourse = v;
            break;
          }
        }
        if (!titleCourse && /^(课程|科目)[：:\s]/.test(s)) {
          if (!onlyCourseLabel || s.startsWith("课程")) {
            const v = s.replace(/^(课程|科目)[：:\s]*/, "").trim();
            if (acceptAsTitleCourse(v, onlyCourseLabel)) titleCourse = v;
          }
        }
        if (!titleCourse && isCourseOrSubjectLabel(s)) {
          if (onlyCourseLabel && s.startsWith("科目")) continue;
          const nextCell = String((row[c + 1] ?? "")).trim();
          const prevCell = c > 0 ? String((row[c - 1] ?? "")).trim() : "";
          if (acceptAsTitleCourse(nextCell, onlyCourseLabel)) titleCourse = nextCell;
          else if (acceptAsTitleCourse(prevCell, onlyCourseLabel)) titleCourse = prevCell;
          else {
            for (let j = 0; j < row.length; j++) {
              if (j === c) continue;
              const v = String(row[j] ?? "").trim();
              if (acceptAsTitleCourse(v, onlyCourseLabel)) {
                titleCourse = v;
                break;
              }
            }
          }
        }
      }
      if (!titleCourse && isLikelyCourseName(s) && s.length <= 8) {
        const prevCell = c > 0 ? String((row[c - 1] ?? "")).trim() : "";
        const prevIsCourseLabel =
          isCourseOrSubjectLabel(prevCell) ||
          (onlyCourseLabel && /课程\s*[：:]/.test(prevCell) && isExcludedInferredCourse((prevCell.split(/[：:]/)[1] || "").trim()));
        if (prevIsCourseLabel && (!onlyCourseLabel || prevCell.startsWith("课程")) && acceptAsTitleCourse(s, onlyCourseLabel))
          titleCourse = s;
      }
    }
    if (!titleCourse) {
      const hasLabel = row.some((cell) => isCourseOrSubjectLabel(String(cell ?? "").trim()));
      const hasCourseOnly = onlyCourseLabel && row.some((cell) => /^课程\s*[：:]?\s*$/.test(String(cell ?? "").trim()));
      if (hasLabel && (!onlyCourseLabel || hasCourseOnly)) {
        for (let j = 0; j < row.length; j++) {
          const v = String(row[j] ?? "").trim();
          if (acceptAsTitleCourse(v, onlyCourseLabel)) {
            titleCourse = v;
            break;
          }
        }
      }
    }
  }
  return { class: titleClass, course: titleCourse };
}

function optionize(selectEl, values) {
  selectEl.innerHTML = values
    .map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`)
    .join("");
}

/** 用于班级下拉：排除 abc班 等，并追加「添加班级」选项 */
function getClassOptionsForSelect() {
  const fromRows = (state.rows || [])
    .map((r) => (r.className != null ? String(r.className).trim() : ""))
    .filter((s) => s);
  const list = [...new Set([...DEFAULT_CLASSES, ...fromRows])].filter(
    (c) => c && !CLASS_EXCLUDED_FROM_OPTIONS.includes(c)
  );
  return list;
}

/** 填充班级下拉：选项列表 + 最后一项「添加班级」 */
function optionizeClassSelect(selectEl, classList) {
  const opts = classList.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
  selectEl.innerHTML = opts + `<option value="${escapeHtml(ADD_CLASS_OPTION_VALUE)}">添加班级</option>`;
}

/** 用于课程下拉：已有数据 + 默认课程，并追加「添加课程」选项 */
function getCourseOptionsForSelect() {
  const fromRows = (state.rows || [])
    .map((r) => (r.course != null ? String(r.course).trim() : ""))
    .filter((s) => s);
  return [...new Set([...DEFAULT_COURSES, ...fromRows])].filter((c) => c);
}

/** 填充课程下拉：选项列表 + 最后一项「添加课程」 */
function optionizeCourseSelect(selectEl, courseList) {
  const opts = courseList.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
  selectEl.innerHTML = opts + `<option value="${escapeHtml(ADD_COURSE_OPTION_VALUE)}">添加课程</option>`;
}

// ---------- 添加课程/学科弹窗 ----------
let addCourseOnConfirm = null;
let addCourseOnCancel = null;

function openAddCourseOverlay(onConfirm, onCancel) {
  addCourseOnConfirm = onConfirm;
  addCourseOnCancel = onCancel || null;
  if (els.addCourseInput) els.addCourseInput.value = "";
  if (els.addCourseHint) els.addCourseHint.textContent = "";
  if (els.addCourseOverlay) els.addCourseOverlay.style.display = "grid";
  setTimeout(() => els.addCourseInput?.focus(), 80);
}

function closeAddCourseOverlay(canceled = true) {
  if (canceled && typeof addCourseOnCancel === "function") addCourseOnCancel();
  addCourseOnConfirm = null;
  addCourseOnCancel = null;
  if (els.addCourseOverlay) els.addCourseOverlay.style.display = "none";
}

function submitAddCourse() {
  const name = (els.addCourseInput?.value ?? "").trim();
  if (els.addCourseHint) els.addCourseHint.textContent = "";
  if (!name) {
    if (els.addCourseHint) els.addCourseHint.textContent = "请填写课程名称。";
    return;
  }
  if (typeof addCourseOnConfirm === "function") addCourseOnConfirm(name);
  closeAddCourseOverlay(false);
}

// ---------- 自定义成绩比例弹窗 ----------
function openGradeFormulaOverlay() {
  const formula = getEffectiveGradeFormula();
  const p0 = Math.round((formula.terms?.[0]?.weight ?? 0.4) * 100);
  const p1 = Math.round((formula.terms?.[1]?.weight ?? 0.6) * 100);
  if (els.gradeFormulaUsual) els.gradeFormulaUsual.value = String(p0);
  if (els.gradeFormulaExam) els.gradeFormulaExam.value = String(p1);
  if (els.gradeFormulaOverlayHint) els.gradeFormulaOverlayHint.textContent = "";
  if (els.gradeFormulaOverlay) els.gradeFormulaOverlay.style.display = "grid";
  setTimeout(() => els.gradeFormulaUsual?.focus(), 80);
}

function closeGradeFormulaOverlay() {
  if (els.gradeFormulaOverlay) els.gradeFormulaOverlay.style.display = "none";
}

function submitGradeFormula() {
  const usualPct = clampInt(els.gradeFormulaUsual?.value, 0, 100);
  const examPct = clampInt(els.gradeFormulaExam?.value, 0, 100);
  if (els.gradeFormulaOverlayHint) els.gradeFormulaOverlayHint.textContent = "";
  if (usualPct == null || examPct == null) {
    if (els.gradeFormulaOverlayHint) els.gradeFormulaOverlayHint.textContent = "请填写平时与考试成绩的百分比（0–100）。";
    return;
  }
  const sum = (usualPct ?? 0) + (examPct ?? 0);
  if (Math.abs(sum - 100) > 1) {
    if (els.gradeFormulaOverlayHint) els.gradeFormulaOverlayHint.textContent = "平时% + 考试% 应为 100%。";
    return;
  }
  state.gradeFormula = {
    type: "weighted",
    terms: [
      { name: "平时成绩", weight: (usualPct ?? 40) / 100 },
      { name: "考试成绩", weight: (examPct ?? 60) / 100 },
    ],
  };
  setDataChanged();
  saveState(state);
  closeGradeFormulaOverlay();
  renderGradeFormulaHint();
  renderAll();
}

// ---------- 添加班级弹窗（统一样式，替代 window.prompt） ----------
let addClassOnConfirm = null;
let addClassOnCancel = null;

function openAddClassOverlay(onConfirm, onCancel) {
  addClassOnConfirm = onConfirm;
  addClassOnCancel = onCancel || null;
  if (els.addClassInput) els.addClassInput.value = "";
  if (els.addClassHint) els.addClassHint.textContent = "";
  if (els.addClassOverlay) els.addClassOverlay.style.display = "grid";
  setTimeout(() => els.addClassInput?.focus(), 80);
}

function closeAddClassOverlay(canceled = true) {
  if (canceled && typeof addClassOnCancel === "function") addClassOnCancel();
  addClassOnConfirm = null;
  addClassOnCancel = null;
  if (els.addClassOverlay) els.addClassOverlay.style.display = "none";
}

function submitAddClass() {
  const name = (els.addClassInput?.value ?? "").trim();
  if (els.addClassHint) els.addClassHint.textContent = "";
  if (!name) {
    if (els.addClassHint) els.addClassHint.textContent = "请填写班级名称。";
    return;
  }
  if (typeof addClassOnConfirm === "function") addClassOnConfirm(name);
  closeAddClassOverlay(false);
}

// ---------- 单个学生成绩录入 ----------
function openAddStudentOverlay() {
  const classes = getClassOptionsForSelect();
  const courses = getCourseOptionsForSelect();
  optionizeClassSelect(els.addStudentClass, classes.length ? classes : []);
  optionizeCourseSelect(els.addStudentCourse, courses.length ? courses : []);
  els.addStudentClass.value = (state.selectedClass && state.selectedClass !== ADD_CLASS_OPTION_VALUE && classes.includes(state.selectedClass)) ? state.selectedClass : (classes[0] ?? ADD_CLASS_OPTION_VALUE);
  els.addStudentCourse.value = (state.selectedCourse && state.selectedCourse !== ADD_COURSE_OPTION_VALUE && courses.includes(state.selectedCourse)) ? state.selectedCourse : (courses[0] ?? ADD_COURSE_OPTION_VALUE);
  els.addStudentName.value = "";
  els.addStudentUsual.value = "";
  els.addStudentExam.value = "";
  if (els.addStudentHint) els.addStudentHint.textContent = "";
  updateAddStudentFinal();
  if (els.addStudentOverlay) els.addStudentOverlay.style.display = "grid";
}

function closeAddStudentOverlay() {
  if (els.addStudentOverlay) els.addStudentOverlay.style.display = "none";
}

function updateAddStudentFinal() {
  const usual = clampInt(els.addStudentUsual?.value, 0, 100);
  const exam = clampInt(els.addStudentExam?.value, 0, 100);
  if (usual == null && exam == null) {
    if (els.addStudentFinal) els.addStudentFinal.textContent = "填写平时/考试后自动计算";
    return;
  }
  const formula = getEffectiveGradeFormula();
  const v0 = usual ?? 0;
  const v1 = exam ?? 0;
  const finalVal =
    formula.type === "weighted" && formula.terms?.length >= 2
      ? Math.round(v0 * (formula.terms[0].weight ?? 0.4) + v1 * (formula.terms[1].weight ?? 0.6))
      : v0 + v1;
  if (els.addStudentFinal) els.addStudentFinal.textContent = String(finalVal);
}

function submitAddStudent() {
  const className = (els.addStudentClass?.value ?? "").trim();
  const name = (els.addStudentName?.value ?? "").trim();
  const course = (els.addStudentCourse?.value ?? "").trim();
  const usual = clampInt(els.addStudentUsual?.value, 0, 100);
  const exam = clampInt(els.addStudentExam?.value, 0, 100);
  if (els.addStudentHint) els.addStudentHint.textContent = "";
  if (!name) {
    if (els.addStudentHint) els.addStudentHint.textContent = "请填写姓名。";
    return;
  }
  if (className === ADD_CLASS_OPTION_VALUE || !className.trim()) {
    if (els.addStudentHint) els.addStudentHint.textContent = "请选择班级，或选择「添加班级」后输入新班级名。";
    return;
  }
  const courseVal = (els.addStudentCourse?.value ?? "").trim();
  if (courseVal === ADD_COURSE_OPTION_VALUE || !courseVal) {
    if (els.addStudentHint) els.addStudentHint.textContent = "请选择课程，或选择「添加课程」后输入新课程名。";
    return;
  }
  const nextId = (state.rows.length + 1);
  const id = `S${String(nextId).padStart(3, "0")}`;
  const newRow = {
    id,
    className: className || state.selectedClass || "",
    name,
    course: courseVal || state.selectedCourse || "",
    usual: usual ?? null,
    exam: exam ?? null,
    submitted: false,
    submittedAt: null,
    dirty: true,
    lastUpdatedAt: Date.now(),
  };
  state.rows.push(newRow);
  ensureClassOption(className || state.selectedClass);
  ensureCourseOption(courseVal || state.selectedCourse);
  setDataChanged();
  saveState(state);
  closeAddStudentOverlay();
  state.selectedClass = newRow.className;
  state.selectedCourse = newRow.course;
  syncControlsFromState();
  state.pageIndex = 1;
  saveState(state);
  renderAll();
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normCol(s) {
  return String(s).trim().toLowerCase().replaceAll(" ", "");
}

function pickCol(columns, aliases) {
  const map = new Map(columns.map((c) => [normCol(c), c]));
  for (const a of aliases) {
    const hit = map.get(normCol(a));
    if (hit) return hit;
  }
  return null;
}

function toIntOrNull(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "string" && v.trim() === "") return null;
  const n = Number(v);
  if (Number.isNaN(n)) return null;
  const r = Math.round(n);
  return Math.min(100, Math.max(0, r));
}

/** 根据文件类型读取为工作簿（支持 .xlsx/.xls 与 .csv） */
async function getWorkbookFromFile(file) {
  if (!window.XLSX) throw new Error("未加载 XLSX 解析库（xlsx.full.min.js）。");
  const fileName = String(file?.name ?? "").toLowerCase();
  const isCsv = fileName.endsWith(".csv");
  if (isCsv) {
    const text = await file.text();
    return window.XLSX.read(text, { type: "string", raw: false });
  }
  const buf = await file.arrayBuffer();
  return window.XLSX.read(buf, { type: "array" });
}

/** 读取 Excel/CSV 工作簿的工作表列表及每个表首行标题（用于下拉展示） */
async function getWorkbookSheetInfo(file) {
  if (!window.XLSX) return { sheetNames: [], sheetTitles: [] };
  const wb = await getWorkbookFromFile(file);
  const names = wb.SheetNames || [];
  const sheetTitles = names.map((name) => {
    const ws = wb.Sheets[name];
    if (!ws) return { name, title: "" };
    const aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", range: 0 });
    const firstRow = aoa[0] || [];
    const title = firstRow
      .slice(0, 5)
      .map((c) => String(c ?? "").trim())
      .filter(Boolean)
      .join(" ")
      .slice(0, 30);
    return { name, title: title || name };
  });
  return { sheetNames: names, sheetTitles };
}

/** 从已打开的 workBook 解析单个工作表 */
function parseOneSheetFromWorkbook(wb, fileName, targetSheet) {
  const ws = wb.Sheets[targetSheet];
  if (!ws) throw new Error(`未找到工作表：${targetSheet}`);
  // 兼容两类 Excel：
  // - “规范表格”：第一行就是列名（姓名/平时成绩/考试成绩...）
  // - “报表格式”：有标题行、合并单元格，多行表头（比如“课程成绩报告单”）
  const json = window.XLSX.utils.sheet_to_json(ws, { defval: "" });
  const cols = json.length ? Object.keys(json[0]) : [];
  const colName = pickCol(cols, ["姓名", "name", "student", "学生姓名"]);
  const colClass = pickCol(cols, ["班级", "class", "classname", "行政班"]);
  const colCourse = pickCol(cols, ["课程", "course", "科目", "学科"]);
  const colUsual = pickCol(cols, ["平时成绩", "平时", "usual", "平时分", "过程性评价"]);
  const colExam = pickCol(cols, ["考试成绩", "考试", "exam", "期末", "期末成绩"]);

  // 从工作表标题区（前几行）显式读取「班级：xxx」「课程：xxx」/「科目：xxx」；再用标题区+文件名/工作表名推断
  const aoaForTitle = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const titleMeta = extractMetaFromTitleArea(aoaForTitle, 12, targetSheet);
  const titleTextsForClass = extractTitleTextsFromAoa(aoaForTitle, 12, 24);
  const inferredClassFromName =
    titleMeta.class || inferClassFromTexts([...titleTextsForClass, fileName, targetSheet]);
  // 课程解析顺序：标题区课程 > 考查表时优先括号《科目》 > 关键字推断；若推断结果为排除项则用括号覆盖
  let inferredFromTitleAndName = inferCourseFromTexts([...titleTextsForClass, fileName, targetSheet]);
  if (!titleMeta.course && isSheetNameAssessmentSubject(targetSheet)) {
    const fromBracket = extractCourseFromBrackets(titleTextsForClass);
    if (fromBracket) inferredFromTitleAndName = fromBracket;
  } else if (isExcludedInferredCourse(inferredFromTitleAndName) && !titleMeta.course) {
    const fromBracket = extractCourseFromBrackets(titleTextsForClass);
    if (fromBracket) inferredFromTitleAndName = fromBracket;
  }
  const inferredCourseFromTitle = titleMeta.course || inferredFromTitleAndName;

  if (colName) {
    const grades = [];
    // 考查/考察科目表：表格里的「科目」列多为考查类型（如信息技术），不作为课程，只用标题/括号推断的课程
    const useInferredCourseOnly = isSheetNameAssessmentSubject(targetSheet);
    for (const r of json) {
      const name = String(r[colName] ?? "").trim();
      if (!name) continue;
      const courseVal = useInferredCourseOnly ? "" : (colCourse ? String(r[colCourse] ?? "").trim() : "");
      const classVal = colClass ? String(r[colClass] ?? "").trim() : "";
      grades.push({
        name,
        class_name: classVal || inferredClassFromName,
        course: courseVal || inferredCourseFromTitle,
        usual: colUsual ? toIntOrNull(r[colUsual]) : null,
        exam: colExam ? toIntOrNull(r[colExam]) : null,
      });
    }
    const inferredCourse = uniqueNonEmpty(grades.map((g) => g.course)) || inferredCourseFromTitle;
    const inferredClass = uniqueNonEmpty(grades.map((g) => g.class_name)) || inferredClassFromName;
    if (inferredCourse) ensureCourseOption(inferredCourse);
    // 规范表无权重列时用默认 46 开（平时 40%，考试 60%）
    const gradeFormula = {
      type: "weighted",
      terms: [
        { name: colUsual || "平时成绩", weight: 0.4 },
        { name: colExam || "考试成绩", weight: 0.6 },
      ],
    };
    if (!colUsual && !colExam) gradeFormula.terms = [{ name: "平时成绩", weight: 0.4 }, { name: "考试成绩", weight: 0.6 }];
    return {
      sheet: targetSheet,
      mode: "table",
      detected_columns: { name: colName, class: colClass, course: colCourse, usual: colUsual, exam: colExam },
      inferred_course: inferredCourse || "",
      inferred_class: inferredClass || "",
      grades,
      gradeFormula,
    };
  }

  // 报表格式：用二维数组读取并自动定位“姓名/平时/考试”
  const aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const reportTitleMeta = extractMetaFromTitleArea(aoa, 12, targetSheet);
  const titleTexts = extractTitleTextsFromAoa(aoa, 12, 30);
  // 课程解析顺序：标题区课程 > 考查表时优先括号《科目》 > 关键字推断；若推断结果为排除项则用括号覆盖
  let reportInferredCourse = inferCourseFromTexts([...titleTexts, targetSheet, fileName]);
  if (!reportTitleMeta.course && isSheetNameAssessmentSubject(targetSheet)) {
    const fromBracket = extractCourseFromBrackets(titleTexts);
    if (fromBracket) reportInferredCourse = fromBracket;
  } else if (isExcludedInferredCourse(reportInferredCourse) && !reportTitleMeta.course) {
    const fromBracket = extractCourseFromBrackets(titleTexts);
    if (fromBracket) reportInferredCourse = fromBracket;
  }
  const inferredCourse = reportTitleMeta.course || reportInferredCourse || "";
  const inferredClass =
    reportTitleMeta.class ||
    inferClassFromTexts([...titleTexts, targetSheet, fileName]) ||
    inferredClassFromName;
  if (inferredCourse) ensureCourseOption(inferredCourse);

  const maxScanRows = Math.min(40, aoa.length);

  // 这类“报表”经常是左右两栏学生表：同一行里会出现多个“姓名”表头
  let nameHeaderRow = -1;
  let nameCols = [];
  for (let r = 0; r < maxScanRows; r++) {
    const row = aoa[r] || [];
    const cols = [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] ?? "");
      if (cell.includes("姓名")) cols.push(c);
    }
    if (cols.length) {
      nameHeaderRow = r;
      nameCols = cols;
      break;
    }
  }
  if (nameHeaderRow < 0 || !nameCols.length) {
    throw new Error("Excel 未找到“姓名”表头。若是报表格式，请确认文件不是图片/扫描件。");
  }

  // 分项表头通常在“姓名表头行”的下一行
  let scoreHeaderRow = nameHeaderRow + 1;
  if (scoreHeaderRow >= aoa.length) scoreHeaderRow = nameHeaderRow;
  // 如果下一行没有“平时”，再向下找
  const findRowContains = (kw, from, to) => {
    for (let r = from; r <= to; r++) {
      const row = aoa[r] || [];
      if (row.some((x) => String(x ?? "").includes(kw))) return r;
    }
    return -1;
  };
  if (!((aoa[scoreHeaderRow] || []).some((x) => String(x ?? "").includes("平时")))) {
    const rr = findRowContains("平时", nameHeaderRow, Math.min(nameHeaderRow + 5, aoa.length - 1));
    if (rr >= 0) scoreHeaderRow = rr;
  }

  const scoreRow = aoa[scoreHeaderRow] || [];
  const weightsRow = aoa[scoreHeaderRow + 1] || [];

  const findFirstCol = (predicate, startCol) => {
    for (let c = startCol; c < scoreRow.length; c++) {
      if (predicate(c)) return c;
    }
    return -1;
  };

  const dataStartRow = Math.min(aoa.length, scoreHeaderRow + 2);
  const grades = [];

  const toSeq = (v) => {
    if (v === null || v === undefined) return null;
    const s = String(v).trim();
    if (!s) return null;
    const n = Number(s);
    if (!Number.isFinite(n)) return null;
    if (!Number.isInteger(n)) return null;
    if (n <= 0) return null;
    return n;
  };
  const isLikelyName = (name) => {
    const s = String(name ?? "").trim();
    if (!s) return false;
    if (s.length > 10) return false;
    // 避免把说明段落当名字
    if (/[，。:：;；、\\n\\t]/.test(s)) return false;
    if (s.includes("说明") || s.includes("统计") || s.includes("成绩")) return false;
    return true;
  };

  const parseGroup = (nameCol) => {
    const seqCol = Math.max(0, nameCol - 1);
    const usualColIdx = findFirstCol(
      (c) => String(scoreRow[c] ?? "").includes("平时"),
      Math.max(0, nameCol + 1)
    );
    if (usualColIdx < 0) return null;

    const examColByWeight = (() => {
      for (let c = usualColIdx + 1; c < weightsRow.length; c++) {
        const v = Number(weightsRow[c]);
        if (!Number.isNaN(v) && Math.abs(v - 0.4) < 1e-6) return c;
      }
      return -1;
    })();
    const examColIdx =
      examColByWeight >= 0
        ? examColByWeight
        : findFirstCol(
            (c) => {
              const s = String(scoreRow[c] ?? "");
              return s.includes("考试") || s.includes("期末");
            },
            usualColIdx + 1
          );
    if (examColIdx < 0) return null;

    const wUsual = Number(weightsRow[usualColIdx]);
    const wExam = Number(weightsRow[examColIdx]);
    const hasWeights =
      !Number.isNaN(wUsual) &&
      !Number.isNaN(wExam) &&
      wUsual > 0 &&
      wExam > 0 &&
      wUsual <= 1 &&
      wExam <= 1 &&
      Math.abs(wUsual + wExam - 1) < 0.02;

    for (let r = dataStartRow; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const seq = toSeq(row[seqCol]);
      if (seq === null) continue;

      const name = String(row[nameCol] ?? "").trim();
      if (!isLikelyName(name)) continue;

      const rawUsual = Number(row[usualColIdx]);
      const rawExam = Number(row[examColIdx]);
      grades.push({
        name,
        class_name: inferredClass || "",
        course: inferredCourse,
        // 平时成绩、考试成绩一律存原数据（表格中显示原分数）；权重仅用于说明，不改写显示值
        usual: toIntOrNull(row[usualColIdx]),
        exam: toIntOrNull(row[examColIdx]),
      });
    }

    return {
      nameCol,
      seqCol,
      usualColIdx,
      examColIdx,
      weights: hasWeights ? { usual: wUsual, exam: wExam } : null,
    };
  };

  const groups = nameCols
    .map((c) => parseGroup(c))
    .filter(Boolean);

  if (!groups.length) {
    throw new Error("Excel 报表中未找到可用的“平时/考试”列组。");
  }

  const g0 = groups[0];
  // 表格有权重行则用表格比例（如 64 开），否则用默认 46 开（平时 40%，考试 60%）
  const gradeFormula = {
    type: "weighted",
    terms: [
      { name: String(scoreRow[g0.usualColIdx] ?? "平时").trim() || "平时", weight: g0.weights?.usual ?? 0.4 },
      { name: String(scoreRow[g0.examColIdx] ?? "考试").trim() || "考试", weight: g0.weights?.exam ?? 0.6 },
    ],
  };

  // 去重：同名取最后一次（避免重复页/重复表头）
  const mp = new Map();
  for (const g of grades) mp.set(g.name, g);
  const deduped = [...mp.values()];

  return {
    sheet: targetSheet,
    mode: "report",
    detected_columns: {
      name: `R${nameHeaderRow}C${nameCols.join(",")}`,
      score_row: `R${scoreHeaderRow}`,
    },
    groups,
    inferred_course: inferredCourse || "",
    inferred_class: inferredClass || "",
    grades: deduped,
    gradeFormula,
  };
}

async function parseExcel(file, sheetNameOrAll) {
  const fileName = String(file?.name ?? "");
  const wb = await getWorkbookFromFile(file);
  const sheetNames = wb.SheetNames || [];
  if (!sheetNames.length) throw new Error("表格未找到可用工作表。");

  const wantAll = sheetNameOrAll === "__all__";
  const targets = wantAll ? sheetNames : [sheetNames.includes(sheetNameOrAll) ? sheetNameOrAll : sheetNames[0]];

  const results = [];
  for (const targetSheet of targets) {
    try {
      results.push(parseOneSheetFromWorkbook(wb, fileName, targetSheet));
    } catch (e) {
      if (!wantAll) throw e;
      // 全选时某表解析失败则跳过
    }
  }

  if (results.length === 0) throw new Error("未解析到任何有效数据。");
  if (results.length === 1) return results[0];

  const allGrades = results.flatMap((r) => r.grades || []);
  const courseCounts = new Map();
  for (const g of allGrades) {
    const c = String(g.course ?? "").trim();
    if (c) courseCounts.set(c, (courseCounts.get(c) || 0) + 1);
  }
  let mainCourse = "";
  let maxCount = 0;
  for (const [c, n] of courseCounts) {
    if (n > maxCount) {
      maxCount = n;
      mainCourse = c;
    }
  }
  // 若按行数最多的课程是「信息技术」「计算机」（多为表格科目列误带出），优先用第一个表的推断课程，避免误判
  if (isExcludedInferredCourse(mainCourse) && results[0].inferred_course && !isExcludedInferredCourse(results[0].inferred_course)) {
    mainCourse = results[0].inferred_course;
  }
  const merged = {
    sheet: results.map((r) => r.sheet).join("、"),
    mode: results[0].mode,
    detected_columns: results[0].detected_columns,
    inferred_course: mainCourse || results[0].inferred_course || "",
    inferred_class: results[0].inferred_class || "",
    grades: allGrades,
  };
  if (results[0].weights) merged.weights = results[0].weights;
  if (results[0].gradeFormula) merged.gradeFormula = results[0].gradeFormula;
  return merged;
}

function uniqueNonEmpty(values) {
  const s = new Set(values.map((x) => String(x ?? "").trim()).filter((x) => x));
  return s.size === 1 ? [...s][0] : null;
}

function applyGradesToState(grades, parsed) {
  // 以表格读取比例优先；无表格或未解析出加权比例时用默认 46 开（平时 40%，考试 60%）
  const formula = parsed?.gradeFormula;
  state.gradeFormula =
    formula && formula.type === "weighted" && formula.terms?.length >= 2
      ? formula
      : { ...DEFAULT_GRADE_FORMULA, terms: DEFAULT_GRADE_FORMULA.terms.map((t) => ({ ...t })) };
  const courseUnique = uniqueNonEmpty(grades.map((g) => g.course));
  const classUnique = uniqueNonEmpty(grades.map((g) => g.class_name));
  // 只有当 Excel 中所有行的课程都相同时，才自动设置下拉框（但不强制统一所有行）
  // 若唯一课程为「信息技术」「计算机」且解析出更可信的 inferred_course（如语文），优先用解析出的课程，避免误判
  const effectiveCourse =
    courseUnique && !isExcludedInferredCourse(courseUnique)
      ? courseUnique
      : parsed?.inferred_course && !isExcludedInferredCourse(parsed.inferred_course)
        ? parsed.inferred_course
        : courseUnique;
  if (effectiveCourse) {
    ensureCourseOption(effectiveCourse);
    state.selectedCourse = effectiveCourse;
  }
  // 若 Excel 中所有行班级一致，自动加入下拉并选中
  if (classUnique) {
    ensureClassOption(classUnique);
    state.selectedClass = classUnique;
  }

  if (state.isSample) {
    state.rows = [];
    state.isSample = false;
  }
  state.deletedAt = null;
  setDataChanged();

  if (!state.rows.length) {
    const now = Date.now();
    state.rows = grades
      // 只要有姓名就导入；成绩为空的保留，避免“人员遗漏”
      .filter((g) => g?.name)
      .map((g, idx) => {
        const courseFromExcel = String(g.course ?? "").trim();
        const classFromExcel = String(g.class_name ?? "").trim();
        // 若 Excel 行为「信息技术」「计算机」且未显式课程，用解析出的课程填充，避免考查科目表误判
        const rowCourse =
          courseFromExcel && !isExcludedInferredCourse(courseFromExcel)
            ? courseFromExcel
            : state.selectedCourse;
        return {
          id: `S${String(idx + 1).padStart(3, "0")}`,
          // 优先用 Excel 中的班级值；没有才用当前选中的班级
          className: classFromExcel || state.selectedClass,
          name: String(g.name).trim(),
          // 优先用 Excel 中的课程值（排除误判的信息技术/计算机）；没有才用推断出的默认课程
          course: rowCourse,
          usual: g.usual ?? null,
          exam: g.exam ?? null,
          submitted: false,
          submittedAt: null,
          dirty: true,
          lastUpdatedAt: now,
        };
      });

    state.search = "";
    state.pageIndex = 1;
    saveState(state);
    renderAll();

    // 同步下拉选项与默认值
    syncControlsFromState();
    renderAll();
    return { filled: state.rows.length, missingInExcel: [], addedFromExcel: [] };
  }

  const byName = new Map(grades.map((g) => [g.name, g]));
  const missingInExcel = [];
  const addedFromExcel = [];

  let filled = 0;
  state.rows = state.rows.map((r) => {
    const g = byName.get(r.name);
    if (!g) {
      missingInExcel.push(r.name);
      return r;
    }
    const nextUsual = g.usual;
    const nextExam = g.exam;
    const courseFromExcel = String(g.course ?? "").trim();
    const classFromExcel = String(g.class_name ?? "").trim();
    if (nextUsual === null && nextExam === null && !courseFromExcel && !classFromExcel) return r;
    filled += 1;
    // 若 Excel 行为「信息技术」「计算机」且未显式课程，不覆盖为信息技术，保留解析出的课程
    const rowCourse =
      courseFromExcel && !isExcludedInferredCourse(courseFromExcel)
        ? courseFromExcel
        : state.selectedCourse || r.course;
    return {
      ...r,
      // 如果 Excel 中有班级值，更新行的班级（保留每行的独立班级）
      className: classFromExcel || r.className,
      // 如果 Excel 中有课程值，更新行的课程（排除误判的信息技术/计算机）
      course: rowCourse,
      usual: nextUsual,
      exam: nextExam,
      dirty: true,
      submitted: false,
      submittedAt: null,
      lastUpdatedAt: Date.now(),
    };
  });

  // Excel 里有、表格里没有的姓名：自动补到名单里（否则导入会“匹配不到”）
  const existingNames = new Set(state.rows.map((r) => r.name));
  const nextIdNum = () => state.rows.length + 1;
  for (const g of grades) {
    if (existingNames.has(g.name)) continue;
    const courseFromExcel = String(g.course ?? "").trim();
    const classFromExcel = String(g.class_name ?? "").trim();
    const idx = nextIdNum();
    state.rows.push({
      id: `S${String(idx).padStart(3, "0")}`,
      // 优先用 Excel 中的班级值；没有才用当前下拉框的值
      className: classFromExcel || state.selectedClass,
      name: g.name,
      // 优先用 Excel 中的课程值（排除误判的信息技术/计算机）；没有才用当前下拉框的值
      course:
        courseFromExcel && !isExcludedInferredCourse(courseFromExcel)
          ? courseFromExcel
          : state.selectedCourse,
      usual: g.usual ?? null,
      exam: g.exam ?? null,
      submitted: false,
      submittedAt: null,
      dirty: true,
      lastUpdatedAt: Date.now(),
    });
    existingNames.add(g.name);
    addedFromExcel.push(g.name);
    filled += 1;
  }

  // 如果 Excel 的班级一致，自动加入下拉并选中
  const cls = uniqueNonEmpty(grades.map((g) => g.class_name));
  if (cls) {
    ensureClassOption(cls);
    state.selectedClass = cls;
  }
  // 若本次导入的课程唯一，把该班级下所有行的课程统一为该课程，避免旧数据留下其他科目；用已选课程（已排除误判的信息技术）统一
  if (state.selectedCourse && classUnique) {
    state.rows = state.rows.map((r) =>
      r.className === classUnique ? { ...r, course: state.selectedCourse } : r
    );
  }

  state.search = "";
  state.pageIndex = 1;
  saveState(state);
  syncControlsFromState(); // 刷新班级/课程下拉，使识别到的班级出现在选项中并选中
  renderAll();

  return { filled, missingInExcel, addedFromExcel };
}

async function submitAllPages() {
  if (state.isSample === true) {
    if (els.importHint) els.importHint.textContent = "请先导入真实 Excel 再提交。";
    return;
  }
  // 逐页提交：只会提交 dirty 行
  const total = state.rows.length;
  const pageSize = state.pageSize;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  for (let p = 1; p <= totalPages; p++) {
    state.pageIndex = p;
    await submitCurrentPage();
  }
  state.pageIndex = 1;
  await saveStateToServerAndConfirm(state);
  renderAll();
}

function computeFinal(row) {
  const formula = getEffectiveGradeFormula();
  const v0 = row.usual ?? 0;
  const v1 = row.exam ?? 0;
  if (formula.type === "weighted" && formula.terms?.length >= 2) {
    const w0 = formula.terms[0].weight ?? 0.4;
    const w1 = formula.terms[1].weight ?? 0.6;
    return Math.round(v0 * w0 + v1 * w1);
  }
  return v0 + v1;
}

function setAllRowsClassAndCourse() {
  const cls = state.selectedClass;
  const course = state.selectedCourse;
  state.rows = state.rows.map((r) => ({
    ...r,
    className: cls,
    course,
  }));
}

function setAllRowsClass() {
  // 只更新班级，不更新课程（保留每行的独立课程）
  const cls = state.selectedClass;
  state.rows = state.rows.map((r) => ({
    ...r,
    className: cls,
  }));
}

function clearUnsubmittedScoresForCourse(course) {
  // 清空指定课程下所有未提交（dirty）行的成绩数据
  // 注意：已提交的数据（submitted: true）不会被清空
  state.rows = state.rows.map((r) => {
    // 只清空：课程匹配 && 未提交 && 有未提交标记
    if (r.course === course && r.dirty === true && r.submitted !== true) {
      return {
        ...r,
        usual: null,
        exam: null,
        dirty: false,
        lastUpdatedAt: Date.now(),
      };
    }
    return r;
  });
}

function clearUnsubmittedScoresForClass(className) {
  // 清空指定班级下所有未提交（dirty）行的成绩数据
  // 注意：已提交的数据（submitted: true）不会被清空
  state.rows = state.rows.map((r) => {
    // 只清空：班级匹配 && 未提交 && 有未提交标记
    if (r.className === className && r.dirty === true && r.submitted !== true) {
      return {
        ...r,
        usual: null,
        exam: null,
        dirty: false,
        lastUpdatedAt: Date.now(),
      };
    }
    return r;
  });
}

function getFilteredRows() {
  // 先按班级和课程过滤：只显示当前选中班级和课程的数据
  let rows = state.rows.filter((r) => 
    r.className === state.selectedClass && r.course === state.selectedCourse
  );
  
  // 再按姓名搜索过滤
  const kw = state.search.trim();
  if (kw) {
    rows = rows.filter((r) => r.name.includes(kw));
  }
  
  return rows;
}

function getPagedRows() {
  const rows = getFilteredRows();
  const total = rows.length;
  const pageSize = state.pageSize;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  const pageIndex = Math.min(totalPages, Math.max(1, state.pageIndex));
  const start = (pageIndex - 1) * pageSize;
  const end = start + pageSize;
  return {
    rows: rows.slice(start, end),
    total,
    totalPages,
    pageIndex,
  };
}

function anyDirtyOnCurrentPage() {
  const { rows } = getPagedRows();
  return rows.some((r) => r.dirty);
}

function anyDirtyAnywhere() {
  return state.rows.some((r) => r.dirty);
}

function renderDirtyHint() {
  const dirtyAny = state.rows.some((r) => r.dirty);
  const submittedAny = state.rows.some((r) => r.submitted);
  const hasRows = (state.rows || []).length > 0;
  const deletedAt = state.deletedAt;

  if (hasRows === false && deletedAt && sessionHasLocalChanges) {
    els.dirtyHint.className = "badge badge--danger";
    els.dirtyHint.textContent = `数据删除于 ${formatDateTime(deletedAt)}`;
  } else if (hasRows === false && !deletedAt) {
    els.dirtyHint.className = "badge badge--unmodified";
    els.dirtyHint.textContent = "未导入";
  } else if (dirtyAny && sessionHasLocalChanges) {
    // 仅在本会话有操作时显示「未提交」；刷新/重登后忽略历史 dirty，显示「未修改」
    els.dirtyHint.className = "badge badge--warn";
    els.dirtyHint.textContent = "未提交";
  } else if (submittedAny && sessionHasLocalChanges) {
    const submittedRows = state.rows.filter((r) => r.submitted && r.submittedAt);
    const latest = submittedRows.length ? Math.max(...submittedRows.map((r) => r.submittedAt)) : null;
    const timeStr = latest ? formatTime(latest) : "";
    els.dirtyHint.className = "badge badge--ok";
    els.dirtyHint.textContent = timeStr ? `已提交 ${timeStr}` : "已提交";
  } else if (!sessionHasLocalChanges) {
    // 只有重新登录后（本会话无本地操作）才显示「未修改」
    els.dirtyHint.className = "badge badge--unmodified";
    els.dirtyHint.textContent = "未修改";
  } else {
    // 本会话有过操作但当前无脏数据且未在本次会话提交：仍显示未提交
    els.dirtyHint.className = "badge badge--warn";
    els.dirtyHint.textContent = "未提交";
  }
}

function renderSubmitBtn() {
  const { rows } = getPagedRows();
  // 当页有数据时始终显示「提交本页」，便于自动化/Agent 通过 id=submitBtn 定位；无脏数据时点击为 no-op
  els.submitBtn.style.display = rows.length > 0 ? "inline-flex" : "none";
}

function renderSubmitAllBtn() {
  els.submitAllBtn.style.display = anyDirtyAnywhere() ? "inline-flex" : "none";
}

function renderTable() {
  const { rows, pageIndex } = getPagedRows();
  if (!rows.length) {
    els.tbody.innerHTML = `
      <tr>
        <td colspan="9">
          <span class="muted">当前无数据：请先导入 Excel。</span>
        </td>
      </tr>
    `;
    return;
  }

  els.tbody.innerHTML = rows
    .map((r, i) => {
      const idx = (pageIndex - 1) * state.pageSize + (i + 1);
      const isSubmitting = uiSubmittingRowIds.has(r.id);
      const final = isSubmitting ? "" : computeFinal(r);
      const timeText = r.submittedAt ? `提交于 ${formatTime(r.submittedAt)}` : "";
      const usualVal = r.usual != null ? escapeHtml(String(r.usual)) : "";
      const examVal = r.exam != null ? escapeHtml(String(r.exam)) : "";
      const usualCell = isSubmitting
        ? '<span class="muted">提交中…</span>'
        : `<input type="number" class="cellInput grade-inline-input" data-row-id="${escapeHtml(r.id)}" data-field="usual" value="${usualVal}" min="0" max="100" step="0.01" placeholder="0–100" aria-label="平时成绩" />`;
      const examCell = isSubmitting
        ? '<span class="muted">提交中…</span>'
        : `<input type="number" class="cellInput grade-inline-input" data-row-id="${escapeHtml(r.id)}" data-field="exam" value="${examVal}" min="0" max="100" step="0.01" placeholder="0–100" aria-label="考试成绩" />`;
      return `
        <tr data-row-id="${escapeHtml(r.id)}">
          <td><span class="muted">${idx}</span></td>
          <td>${escapeHtml(r.className)}</td>
          <td>${escapeHtml(r.name)}</td>
          <td>${escapeHtml(r.course)}</td>
          <td>${usualCell}</td>
          <td>${examCell}</td>
          <td><span class="cell">${final}</span></td>
          <td>
            ${
              isSubmitting
                ? `<span class="status status--submitting">提交中…</span>`
                : r.submitted
                ? `<div><span class="status status--submitted">已提交</span>${timeText ? `<div class="statusTime">${escapeHtml(timeText)}</div>` : ""}</div>`
                : `<span class="status status--draft">未提交</span>`
            }
          </td>
          <td>
            <div class="opCell">
              <button class="btn btn--mini ${r.dirty ? "btn--primary" : ""}" type="button" data-action="submit-row" aria-label="提交该行" ${r.dirty ? "" : "disabled"}>
                提交
              </button>
              <button class="btn btn--mini btn--danger-soft" type="button" data-action="delete-row">
                删除
              </button>
            </div>
          </td>
        </tr>
      `;
    })
    .join("");
}

async function submitSingleRow(rowId) {
  const idx = state.rows.findIndex((r) => r.id === rowId);
  if (idx < 0) return;
  const row = state.rows[idx];
  if (!row.dirty) return;
  // 只提交当前选中班级和课程的行
  if (row.className !== state.selectedClass || row.course !== state.selectedCourse) return;
  setDataChanged();
  const now = Date.now();
  state.rows[idx] = { ...row, dirty: false, submitted: true, submittedAt: now };
  if (state.isSample === true) {
    saveState(state);
    startSubmitFlash([rowId]);
    renderAll();
    return;
  }
  await saveStateToServerAndConfirm(state);
  startSubmitFlash([rowId]);
}

function deleteRow(rowId) {
  const idx = state.rows.findIndex((r) => r.id === rowId);
  if (idx < 0) return;
  const name = state.rows[idx]?.name ?? "";
  const ok = window.confirm(`确认删除该行？${name ? `\\n学生：${name}` : ""}`);
  if (!ok) return;
  state.rows = state.rows.filter((r) => r.id !== rowId);
  uiSubmittingRowIds.delete(rowId);
  setDataChanged();
  saveState(state);
  renderAll();
}

function renderPager() {
  const { total, totalPages, pageIndex } = getPagedRows();
  const makeBtn = (label, disabled, onClick, extraClass = "") => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `btn ${extraClass}`.trim();
    btn.textContent = label;
    btn.disabled = disabled;
    btn.addEventListener("click", onClick);
    return btn;
  };

  els.pager.innerHTML = "";
  els.pager.appendChild(
    makeBtn("上一页", pageIndex <= 1, () => {
      state.pageIndex = Math.max(1, state.pageIndex - 1);
      saveState(state);
      renderAll();
    })
  );

  // 简单页码（最多展示 7 个）
  const max = 7;
  let start = Math.max(1, pageIndex - Math.floor(max / 2));
  let end = Math.min(totalPages, start + max - 1);
  start = Math.max(1, end - max + 1);

  for (let p = start; p <= end; p++) {
    const active = p === pageIndex;
    els.pager.appendChild(
      makeBtn(
        String(p),
        false,
        () => {
          state.pageIndex = p;
          saveState(state);
          renderAll();
        },
        active ? "btn--primary" : ""
      )
    );
  }

  els.pager.appendChild(
    makeBtn("下一页", pageIndex >= totalPages, () => {
      state.pageIndex = Math.min(totalPages, state.pageIndex + 1);
      saveState(state);
      renderAll();
    })
  );

  const info = document.createElement("div");
  info.className = "muted";
  info.textContent = `共 ${total} 人 / ${totalPages} 页`;
  els.pager.appendChild(info);
}

function renderAll() {
  // 纠正 pageIndex
  const { totalPages, pageIndex } = getPagedRows();
  if (pageIndex !== state.pageIndex) {
    state.pageIndex = pageIndex;
    saveState(state);
  }

  renderTable();
  renderPager();
  renderSubmitBtn();
  renderSubmitAllBtn();
  renderDirtyHint();
  renderGradeFormulaHint();
}

function renderGradeFormulaHint() {
  const text = formatGradeFormulaHint(getEffectiveGradeFormula());
  if (els.gradeFormulaHint) els.gradeFormulaHint.textContent = text;
  if (els.gradeFormulaHintPanel) els.gradeFormulaHintPanel.textContent = text;
}

function openEditor(rowId, field) {
  const { pageIndex } = getPagedRows();
  editing = { rowId, field, pageIndexAtOpen: pageIndex };
  const row = state.rows.find((r) => r.id === rowId);
  const current = row?.[field] ?? "";
  els.editorInput.value = current;
  els.editorOverlay.style.display = "grid";
  setTimeout(() => els.editorInput.focus(), 0);
}

function closeEditor() {
  editing = null;
  els.editorOverlay.style.display = "none";
}

function commitInlineFromInput() {
  if (!editing?.inlineInput) return;
  const { rowId, field, pageIndexAtOpen, inlineInput } = editing;
  if (_inlineInputHandlers && inlineInput === els.gradeInlineInput) {
    inlineInput.removeEventListener("blur", _inlineInputHandlers.onBlur);
    inlineInput.removeEventListener("keydown", _inlineInputHandlers.onKeydown);
    _inlineInputHandlers = null;
  }
  if (els.inlineEditOverlay) els.inlineEditOverlay.style.display = "none";
  const rowIndex = state.rows.findIndex((r) => r.id === rowId);
  editing = null;
  if (rowIndex < 0) {
    renderAll();
    return;
  }
  const next = clampInt(inlineInput.value, 0, 100);
  const old = state.rows[rowIndex][field];
  if (next !== old) {
    setDataChanged();
    const updated = { ...state.rows[rowIndex] };
    updated[field] = next;
    updated.dirty = true;
    updated.submitted = false;
    updated.submittedAt = null;
    updated.lastUpdatedAt = Date.now();
    state.rows[rowIndex] = updated;
    saveState(state);
  }
  state.pageIndex = pageIndexAtOpen;
  saveState(state);
  renderAll();
}

function cancelInline() {
  if (!editing?.inlineInput) return;
  if (_inlineInputHandlers && editing.inlineInput === els.gradeInlineInput) {
    editing.inlineInput.removeEventListener("blur", _inlineInputHandlers.onBlur);
    editing.inlineInput.removeEventListener("keydown", _inlineInputHandlers.onKeydown);
    _inlineInputHandlers = null;
  }
  if (els.inlineEditOverlay) els.inlineEditOverlay.style.display = "none";
  editing = null;
  renderAll();
}

// 全局内联输入框的 blur/keydown 句柄，用于 commit/cancel 时移除
let _inlineInputHandlers = null;

function startInlineEdit(rowId, field) {
  if (editing?.inlineInput) commitInlineFromInput();
  const row = state.rows.find((r) => r.id === rowId);
  if (!row) return;
  const current = row[field] ?? "";
  const { pageIndex } = getPagedRows();
  const tr = Array.from(els.tbody.querySelectorAll("tr[data-row-id]")).find(
    (r) => r.getAttribute("data-row-id") === rowId
  );
  const cell = tr?.querySelector(`[data-field="${field}"]`);
  if (!cell || !els.gradeInlineInput || !els.inlineEditOverlay) return;

  const input = els.gradeInlineInput;
  const overlay = els.inlineEditOverlay;
  input.value = current;
  input.setAttribute("aria-label", field === "usual" ? "平时成绩" : "考试成绩");
  const rect = cell.getBoundingClientRect();
  overlay.style.left = rect.left + "px";
  overlay.style.top = rect.top + "px";
  overlay.style.width = Math.max(rect.width, 90) + "px";
  overlay.style.height = rect.height + "px";
  overlay.style.display = "block";
  input.focus();
  input.select();

  editing = { rowId, field, pageIndexAtOpen: pageIndex, inlineInput: input };

  function onBlur() {
    if (_inlineInputHandlers) {
      input.removeEventListener("blur", _inlineInputHandlers.onBlur);
      input.removeEventListener("keydown", _inlineInputHandlers.onKeydown);
      _inlineInputHandlers = null;
    }
    commitInlineFromInput();
  }
  function onKeydown(e) {
    if (e.key === "Enter") {
      e.preventDefault();
      if (_inlineInputHandlers) {
        input.removeEventListener("blur", _inlineInputHandlers.onBlur);
        input.removeEventListener("keydown", _inlineInputHandlers.onKeydown);
        _inlineInputHandlers = null;
      }
      commitInlineFromInput();
    } else if (e.key === "Escape") {
      e.preventDefault();
      if (_inlineInputHandlers) {
        input.removeEventListener("blur", _inlineInputHandlers.onBlur);
        input.removeEventListener("keydown", _inlineInputHandlers.onKeydown);
        _inlineInputHandlers = null;
      }
      cancelInline();
    }
  }
  _inlineInputHandlers = { onBlur, onKeydown };
  input.addEventListener("blur", onBlur);
  input.addEventListener("keydown", onKeydown);
}

function commitEditorValue() {
  if (!editing) return;
  if (editing.inlineInput) {
    commitInlineFromInput();
    return;
  }
  const { rowId, field, pageIndexAtOpen } = editing;
  // 编辑期间如果翻页/过滤变化，仍然允许提交（按 rowId 精准更新）
  const rowIndex = state.rows.findIndex((r) => r.id === rowId);
  if (rowIndex < 0) return closeEditor();

  const next = clampInt(els.editorInput.value, 0, 100);
  const old = state.rows[rowIndex][field];
  if (next !== old) {
    setDataChanged();
    const updated = { ...state.rows[rowIndex] };
    updated[field] = next;
    updated.dirty = true;
    updated.submitted = false;
    updated.submittedAt = null;
    updated.lastUpdatedAt = Date.now();
    state.rows[rowIndex] = updated;
    saveState(state);
  }

  // 尽量回到打开编辑时的页码（如果搜索变化导致页数不够，会被 renderAll 纠正）
  state.pageIndex = pageIndexAtOpen;
  saveState(state);
  closeEditor();
  renderAll();
}

async function submitCurrentPage() {
  const { rows } = getPagedRows();
  const ids = new Set(rows.map((r) => r.id));
  const currentClass = state.selectedClass;
  const currentCourse = state.selectedCourse;
  let changed = false;
  const now = Date.now();
  const submittedIds = [];
  state.rows = state.rows.map((r) => {
    if (!ids.has(r.id)) return r;
    if (!r.dirty) return r;
    // 只提交当前选中班级和课程的行
    if (r.className !== currentClass || r.course !== currentCourse) return r;
    changed = true;
    submittedIds.push(r.id);
    return { ...r, dirty: false, submitted: true, submittedAt: now };
  });
  if (changed) {
    setDataChanged();
    if (state.isSample === true) {
      saveState(state);
      renderAll();
    } else {
      await saveStateToServerAndConfirm(state);
    }
  }
  startSubmitFlash(submittedIds);
}

async function submitAllDirty() {
  if (!anyDirtyAnywhere()) return;
  setDataChanged();
  const currentClass = state.selectedClass;
  const currentCourse = state.selectedCourse;
  const now = Date.now();
  const currentPageDirtyIds = getPagedRows().rows
    .filter((r) => r.dirty && r.className === currentClass && r.course === currentCourse)
    .map((r) => r.id);
  state.rows = state.rows.map((r) => {
    if (!r.dirty) return r;
    if (r.className !== currentClass || r.course !== currentCourse) return r;
    return { ...r, dirty: false, submitted: true, submittedAt: now };
  });
  if (state.isSample === true) {
    saveState(state);
    renderAll();
  } else {
    await saveStateToServerAndConfirm(state);
  }
  startSubmitFlash(currentPageDirtyIds);
}

function wireEvents() {
  els.logoutBtn?.addEventListener("click", async () => {
    try {
      await apiJson("/api/auth/logout", { method: "POST" });
    } catch {}
    localStorage.removeItem(DATA_CHANGED_KEY);
    window.location.href = "/login.html";
  });

  els.adminBtn?.addEventListener("click", () => {
    window.location.href = "/admin.html";
  });
  // 修改密码
  els.changePwdBtn?.addEventListener("click", () => {
    if (!els.changePwdOverlay) return;
    els.changePwdOld.value = "";
    els.changePwdNew.value = "";
    els.changePwdConfirm.value = "";
    if (els.changePwdHint) els.changePwdHint.textContent = "";
    els.changePwdOverlay.style.display = "grid";
  });
  els.changePwdCancel?.addEventListener("click", () => {
    if (els.changePwdOverlay) els.changePwdOverlay.style.display = "none";
  });
  els.changePwdOverlay?.addEventListener("click", (e) => {
    if (e.target.classList.contains("overlay__backdrop")) els.changePwdOverlay.style.display = "none";
  });
  els.changePwdSubmit?.addEventListener("click", async () => {
    const oldPwd = els.changePwdOld?.value?.trim() ?? "";
    const newPwd = els.changePwdNew?.value?.trim() ?? "";
    const confirmPwd = els.changePwdConfirm?.value?.trim() ?? "";
    if (!els.changePwdHint) return;
    els.changePwdHint.textContent = "";
    if (!oldPwd || !newPwd) {
      els.changePwdHint.textContent = "请填写当前密码和新密码。";
      return;
    }
    if (newPwd !== confirmPwd) {
      els.changePwdHint.textContent = "两次输入的新密码不一致。";
      return;
    }
    if (newPwd.length < 1) {
      els.changePwdHint.textContent = "新密码不能为空。";
      return;
    }
    try {
      await apiJson("/api/auth/change-password", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ oldPassword: oldPwd, newPassword: newPwd }),
      });
      els.changePwdHint.textContent = "修改成功，请使用新密码登录。";
      els.changePwdHint.style.color = "var(--ok)";
      setTimeout(() => {
        els.changePwdOverlay.style.display = "none";
        window.location.href = "/login.html";
      }, 1200);
    } catch (err) {
      els.changePwdHint.textContent = err?.message || "修改失败";
      els.changePwdHint.style.color = "var(--danger)";
    }
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && els.changePwdOverlay?.style.display === "grid") {
      els.changePwdOverlay.style.display = "none";
    }
    if (e.key === "Escape" && els.addStudentOverlay?.style.display === "grid") {
      closeAddStudentOverlay();
    }
    if (e.key === "Escape" && els.addClassOverlay?.style.display === "grid") {
      closeAddClassOverlay();
    }
    if (e.key === "Escape" && els.addCourseOverlay?.style.display === "grid") {
      closeAddCourseOverlay();
    }
    if (e.key === "Escape" && els.gradeFormulaOverlay?.style.display === "grid") {
      closeGradeFormulaOverlay();
    }
  });
  // 单个学生成绩录入
  els.addStudentBtn?.addEventListener("click", () => openAddStudentOverlay());
  els.addStudentCancel?.addEventListener("click", () => closeAddStudentOverlay());
  els.addStudentOverlay?.addEventListener("click", (e) => {
    if (e.target.classList.contains("overlay__backdrop")) closeAddStudentOverlay();
  });
  els.addStudentUsual?.addEventListener("input", () => updateAddStudentFinal());
  els.addStudentExam?.addEventListener("input", () => updateAddStudentFinal());
  els.addStudentSubmit?.addEventListener("click", () => submitAddStudent());
  // 添加班级弹窗（统一样式）
  els.addClassCancel?.addEventListener("click", () => closeAddClassOverlay());
  els.addClassOverlay?.addEventListener("click", (e) => {
    if (e.target.classList.contains("overlay__backdrop")) closeAddClassOverlay();
  });
  els.addClassSubmit?.addEventListener("click", () => submitAddClass());
  els.addClassInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      submitAddClass();
    }
  });
  // 添加课程弹窗
  els.addCourseCancel?.addEventListener("click", () => closeAddCourseOverlay());
  els.addCourseOverlay?.addEventListener("click", (e) => {
    if (e.target.classList.contains("overlay__backdrop")) closeAddCourseOverlay();
  });
  els.addCourseSubmit?.addEventListener("click", () => submitAddCourse());
  els.addCourseInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      submitAddCourse();
    }
  });
  // 自定义成绩比例弹窗
  els.setGradeFormulaBtn?.addEventListener("click", () => openGradeFormulaOverlay());
  els.gradeFormulaCancel?.addEventListener("click", () => closeGradeFormulaOverlay());
  els.gradeFormulaOverlay?.addEventListener("click", (e) => {
    if (e.target.classList.contains("overlay__backdrop")) closeGradeFormulaOverlay();
  });
  els.gradeFormulaSubmit?.addEventListener("click", () => submitGradeFormula());
  els.gradeFormulaUsual?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); submitGradeFormula(); }
  });
  els.gradeFormulaExam?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); submitGradeFormula(); }
  });
  els.addStudentClass?.addEventListener("change", () => {
    if (els.addStudentClass?.value !== ADD_CLASS_OPTION_VALUE) return;
    const classes = getClassOptionsForSelect();
    const prevVal = state.selectedClass && classes.includes(state.selectedClass) ? state.selectedClass : (classes[0] || ADD_CLASS_OPTION_VALUE);
    openAddClassOverlay(
      (trimmed) => {
        ensureClassOption(trimmed);
        const classesNext = getClassOptionsForSelect();
        optionizeClassSelect(els.addStudentClass, classesNext.length ? classesNext : []);
        els.addStudentClass.value = trimmed;
      },
      () => {
        optionizeClassSelect(els.addStudentClass, classes.length ? classes : []);
        els.addStudentClass.value = prevVal;
      }
    );
  });
  els.addStudentCourse?.addEventListener("change", () => {
    if (els.addStudentCourse?.value !== ADD_COURSE_OPTION_VALUE) return;
    const courses = getCourseOptionsForSelect();
    const prevVal = (state.selectedCourse && courses.includes(state.selectedCourse)) ? state.selectedCourse : (courses[0] || ADD_COURSE_OPTION_VALUE);
    openAddCourseOverlay(
      (trimmed) => {
        ensureCourseOption(trimmed);
        const coursesNext = getCourseOptionsForSelect();
        optionizeCourseSelect(els.addStudentCourse, coursesNext.length ? coursesNext : []);
        els.addStudentCourse.value = trimmed;
      },
      () => {
        optionizeCourseSelect(els.addStudentCourse, courses.length ? courses : []);
        els.addStudentCourse.value = prevVal;
      }
    );
  });
  els.classSelect.addEventListener("change", () => {
    const newClass = els.classSelect.value;
    if (newClass === ADD_CLASS_OPTION_VALUE) {
      openAddClassOverlay(
        (trimmed) => {
          ensureClassOption(trimmed);
          state.selectedClass = trimmed;
          syncControlsFromState();
          state.pageIndex = 1;
          saveState(state);
          renderAll();
        },
        () => syncControlsFromState()
      );
      return;
    }
    const previousClass = state.selectedClass;
    if (previousClass !== newClass) clearUnsubmittedScoresForClass(previousClass);
    state.selectedClass = newClass;
    state.pageIndex = 1;
    saveState(state);
    renderAll();
  });
  els.courseSelect.addEventListener("change", () => {
    const newCourse = els.courseSelect.value;
    if (newCourse === ADD_COURSE_OPTION_VALUE) {
      openAddCourseOverlay(
        (trimmed) => {
          ensureCourseOption(trimmed);
          state.selectedCourse = trimmed;
          syncControlsFromState();
          state.pageIndex = 1;
          saveState(state);
          renderAll();
        },
        () => syncControlsFromState()
      );
      return;
    }
    const previousCourse = state.selectedCourse;
    if (previousCourse !== newCourse) clearUnsubmittedScoresForCourse(previousCourse);
    state.selectedCourse = newCourse;
    state.pageIndex = 1;
    saveState(state);
    renderAll();
  });
  els.searchInput.addEventListener("input", () => {
    state.search = els.searchInput.value;
    state.pageIndex = 1;
    saveState(state);
    renderAll();
  });
  els.pageSizeSelect.addEventListener("change", () => {
    state.pageSize = Number(els.pageSizeSelect.value);
    state.pageIndex = 1;
    saveState(state);
    renderAll();
  });

  // 导入后是否自动提交开关
  els.autoSubmitToggle?.addEventListener("change", () => {
    state.autoSubmitOnImport = !!els.autoSubmitToggle.checked;
    saveState(state);
  });

  // 行操作（提交/删除）
  els.tbody.addEventListener("click", (e) => {
    const btn = e.target.closest?.("button[data-action]");
    if (!btn) return;
    const tr = e.target.closest?.("tr[data-row-id]");
    if (!tr) return;
    const rowId = tr.getAttribute("data-row-id");
    if (!rowId) return;
    const action = btn.getAttribute("data-action");
    if (action === "submit-row") submitSingleRow(rowId);
    if (action === "delete-row") deleteRow(rowId);
  });

  function commitInCellScore(input) {
    if (!input) return;
    const rowId = input.getAttribute("data-row-id");
    const field = input.getAttribute("data-field");
    if (!rowId || !field || (field !== "usual" && field !== "exam")) return;
    const rowIndex = state.rows.findIndex((r) => r.id === rowId);
    if (rowIndex < 0) return;
    const next = clampInt(input.value, 0, 100);
    const old = state.rows[rowIndex][field];
    if (next !== old) {
      setDataChanged();
      const updated = { ...state.rows[rowIndex] };
      updated[field] = next;
      updated.dirty = true;
      updated.submitted = false;
      updated.submittedAt = null;
      updated.lastUpdatedAt = Date.now();
      state.rows[rowIndex] = updated;
      saveState(state);
      renderAll();
    }
  }
  // 表内常规 input：平时/考试成绩直接在框内显示与编辑，无覆盖层
  els.tbody.addEventListener("change", (e) => {
    const input = e.target.closest?.("input.cellInput[data-row-id][data-field]");
    if (!input) return;
    commitInCellScore(input);
  });
  els.tbody.addEventListener("keydown", (e) => {
    const input = e.target.closest?.("input.cellInput[data-row-id][data-field]");
    if (!input) return;
    if (e.key === "Enter") {
      e.preventDefault();
      commitInCellScore(input);
      input.blur();
    }
  });
  // 点击非 input 的单元格（如其他列）才打开覆盖层；成绩列已是表内 input，无需覆盖层
  els.tbody.addEventListener("click", (e) => {
    if (e.target.closest?.("button[data-action]")) return;
    if (e.target.classList?.contains("cellInput")) return;
    const cell = e.target.closest?.("[data-field]");
    if (!cell) return;
    const field = cell.getAttribute("data-field");
    if (field === "usual" || field === "exam") return;
    const tr = e.target.closest?.("tr[data-row-id]");
    if (!tr) return;
    const rowId = tr.getAttribute("data-row-id");
    if (!rowId || !field) return;
    startInlineEdit(rowId, field);
  });

  els.tbody.addEventListener("keydown", (e) => {
    if (e.target.classList?.contains("cellInput")) return;
    if (e.key !== "Enter") return;
    const cell = e.target.closest?.("[data-field]");
    if (!cell) return;
    const tr = e.target.closest?.("tr[data-row-id]");
    if (!tr) return;
    const rowId = tr.getAttribute("data-row-id");
    const field = cell.getAttribute("data-field");
    if (!rowId || !field) return;
    e.preventDefault();
    startInlineEdit(rowId, field);
  });

  // 通用确认弹窗
  els.confirmOkBtn?.addEventListener("click", () => closeConfirm(true));
  els.confirmCancelBtn?.addEventListener("click", () => closeConfirm(false));
  els.confirmOverlay?.addEventListener("click", (e) => {
    if (e.target.classList.contains("overlay__backdrop")) closeConfirm(false);
  });

  // overlay 交互
  els.editorOverlay.addEventListener("click", (e) => {
    const isBackdrop = e.target.classList.contains("overlay__backdrop");
    if (isBackdrop) closeEditor();
  });
  document.addEventListener("keydown", (e) => {
    if (els.editorOverlay.style.display === "none") return;
    if (e.key === "Escape") closeEditor();
    if (e.key === "Enter") commitEditorValue();
  });

  els.submitBtn.addEventListener("click", submitCurrentPage);
  els.submitAllBtn?.addEventListener("click", submitAllDirty);

  // Excel 导入：解析后直接应用全部数据，不显示预览
  const openImportPreview = async (submitAll = false) => {
    const file = els.excelFile?.files?.[0];
    if (!file) {
      els.importHint.textContent = "请先选择一个 .xlsx、.xls 或 .csv 文件。";
      return;
    }
    try {
      els.importHint.textContent = "正在解析表格…";
      const sheetValue = (els.sheetSelect?.value ?? "").trim();
      const parsed = await parseExcel(file, sheetValue || "__all__");
      const grades = (parsed.grades || []).filter((g) => g?.name);
      if (!grades.length) {
        els.importHint.textContent = "未解析到有效成绩数据，请检查文件与工作表。";
        return;
      }
      els.importHint.textContent = "";
      const { filled, addedFromExcel } = applyGradesToState(grades, parsed);
      if (submitAll) await submitAllPages();
      const extra = addedFromExcel.length
        ? `（已自动补充名单：${addedFromExcel.slice(0, 5).join("、")}${addedFromExcel.length > 5 ? "…" : ""}）`
        : "";
      const w = parsed?.weights ? `（按权重换算：平时×${parsed.weights.usual}，考试×${parsed.weights.exam}）` : "";
      const inferred = [];
      if (parsed?.inferred_course) inferred.push(`科目：${parsed.inferred_course}`);
      if (parsed?.inferred_class) inferred.push(`班级：${parsed.inferred_class}`);
      const inferredStr = inferred.length ? `，识别到 ${inferred.join("、")}` : "";
      els.importHint.textContent = `导入完成：工作表「${parsed?.sheet ?? ""}」，已填入 ${filled} 人 ${submitAll ? "并已提交全部" : ""}${inferredStr} ${w} ${extra}`;
    } catch (err) {
      els.importHint.textContent = `导入失败：${err?.message ?? String(err)}`;
    }
  };

  els.excelFile?.addEventListener("change", async () => {
    const file = els.excelFile?.files?.[0];
    if (els.excelFileLabel) {
      els.excelFileLabel.textContent = file ? file.name : "未选择文件";
    }
    els.excelFile?.parentElement?.classList.toggle("fileWrap--has-file", !!file);
    if (!els.sheetSelect) return;
    if (!file) {
      els.sheetSelect.innerHTML = '<option value="">请先选择文件</option>';
      els.sheetSelect.disabled = true;
      return;
    }
    try {
      els.sheetSelect.disabled = true;
      const { sheetTitles } = await getWorkbookSheetInfo(file);
      const options = [{ value: "__all__", text: "全选" }];
      sheetTitles.forEach((s, i) => {
        const label = s.title ? `第${i + 1}页 (${s.name}) - ${s.title}` : `第${i + 1}页 (${s.name})`;
        options.push({ value: s.name, text: label });
      });
      els.sheetSelect.innerHTML =
        '<option value="">请先选择文件</option>' +
        options.map((o) => `<option value="${escapeHtml(o.value)}">${escapeHtml(o.text)}</option>`).join("");
      els.sheetSelect.value = options[0].value;
      els.sheetSelect.disabled = false;
      // 选择文件后仅填充工作表下拉，不自动导入；用户点击「导入」后再解析并显示预览
    } catch (e) {
      els.sheetSelect.innerHTML = '<option value="">请先选择文件</option>';
      els.sheetSelect.disabled = false;
      if (els.importHint) els.importHint.textContent = "读取工作表列表失败：" + (e?.message || String(e));
    }
  });

  els.importBtn?.addEventListener("click", () => openImportPreview(!!state.autoSubmitOnImport));
  els.importAndSubmitAllBtn?.addEventListener("click", () => openImportPreview(true));

  els.resetBtn?.addEventListener("click", async () => {
    const ok = await showConfirm({
      title: "确认删除数据",
      message: "这会清空本页已导入/已修改的数据。",
      okText: "确定",
      cancelText: "取消",
    });
    if (!ok) return;
    const submitOk = await showConfirm({
      title: "是否同步到服务器？",
      message: "选择「确定」：同步到服务器，管理员端也会看到已清空。\n选择「取消」：仅本地清空，不同步到服务器。",
      okText: "确定",
      cancelText: "取消",
    });
    const emptyState = buildInitialState();
    emptyState.deletedAt = Date.now();
    state = emptyState;
    setDataChanged();
    saveStateLocal(emptyState);
    // 删除后立即更新「未修改」徽章，确保用户可见
    renderDirtyHint();
    if (submitOk) {
      const synced = await saveStateToServerAndConfirm(emptyState);
      if (synced) {
        localStorage.removeItem(STORAGE_KEY);
        window.location.reload();
        return;
      }
    }
    syncControlsFromState();
    renderAll();
    // 仅本地清空时再次刷新徽章和表格，避免弹窗关闭后未重绘
    requestAnimationFrame(() => {
      renderDirtyHint();
    });
  });
}

function syncControlsFromState() {
  // 班级：已有数据 + 默认列表，排除 abc班，并带「添加班级」选项
  const classes = getClassOptionsForSelect();
  const courses = getCourseOptionsForSelect();
  optionizeClassSelect(els.classSelect, classes);
  optionizeCourseSelect(els.courseSelect, courses);
  if (state.selectedClass === ADD_CLASS_OPTION_VALUE || CLASS_EXCLUDED_FROM_OPTIONS.includes(state.selectedClass) || (classes.length && !classes.includes(state.selectedClass))) {
    state.selectedClass = classes.length ? classes[0] : "";
  }
  if (!classes.length) state.selectedClass = "";
  if (state.selectedCourse === ADD_COURSE_OPTION_VALUE || (courses.length && !courses.includes(state.selectedCourse))) {
    state.selectedCourse = courses.length ? courses[0] : "";
  }
  if (!courses.length) state.selectedCourse = "";
  els.classSelect.value = state.selectedClass ?? "";
  els.courseSelect.value = state.selectedCourse ?? "";
  els.searchInput.value = state.search ?? "";
  els.pageSizeSelect.value = String(state.pageSize ?? 10);
  if (els.autoSubmitToggle) {
    els.autoSubmitToggle.checked = !!state.autoSubmitOnImport;
  }
}

async function boot() {
  const user = await requireLoginOrRedirect();
  if (!user) return;

  // 登录后立即显示成绩比例说明，便于用户和自动化脚本发现「设置比例」功能
  renderGradeFormulaHint();

  // 当前账号显示（便于与管理员侧 ID 对照）
  if (els.currentUser) {
    els.currentUser.textContent = `当前账号: ${user.username || ""} (ID: ${user.id || ""})`;
  }
  // 管理员入口
  if (els.adminBtn) {
    els.adminBtn.style.display = user.role === "admin" ? "inline-flex" : "none";
  }

  normalizeState();
  try {
    const remote = await loadStateRemote();
    if (remote) {
      state = remote;
    }
    // 刷新后：若本地曾有过删除/提交/导入等导致数据变化，恢复标记，不显示「未修改」
    sessionHasLocalChanges = false;
    if (localStorage.getItem(DATA_CHANGED_KEY)) sessionHasLocalChanges = true;
  } catch {}

  syncControlsFromState();
  // 不统一每行的班级/课程，保留每行独立值；表格按 getFilteredRows 只显示当前班级+课程
  // 若有本地数据，主动同步到服务器（失败时会提示），以便管理员能看见
  await saveStateToServerAndConfirm(state);
  wireEvents();
  renderAll();
}

boot();

// 给自动化脚本用的一些稳定钩子（避免依赖 UI 文案变化）
window.__AUTO_GRADE_ENTRY__ = {
  getState: () => JSON.parse(JSON.stringify(state)),
  getRowIdByName: (name) => {
    const r = state.rows.find((x) => x.name === name);
    return r?.id ?? null;
  },
  setRowScores: (rowId, usual, exam) => {
    const idx = state.rows.findIndex((r) => r.id === rowId);
    if (idx < 0) return false;
    const updated = { ...state.rows[idx] };
    const nextUsual = clampInt(usual, 0, 100);
    const nextExam = clampInt(exam, 0, 100);
    updated.usual = nextUsual;
    updated.exam = nextExam;
    updated.dirty = true;
    updated.submitted = false;
    updated.submittedAt = null;
    updated.lastUpdatedAt = Date.now();
    state.rows[idx] = updated;
    saveState(state);
    renderAll();
    return true;
  },
  setScoresByName: (name, usual, exam) => {
    const id = window.__AUTO_GRADE_ENTRY__.getRowIdByName(name);
    if (!id) return false;
    return window.__AUTO_GRADE_ENTRY__.setRowScores(id, usual, exam);
  },
  getVisibleRowIds: () => {
    const { rows } = getPagedRows();
    return rows.map((r) => r.id);
  },
  goToPage: (p) => {
    state.pageIndex = Number(p) || 1;
    saveState(state);
    renderAll();
  },
  submitPage: () => submitCurrentPage(),
  submitAll: () => submitAllDirty(),
  submitRow: (rowId) => submitSingleRow(rowId),
  deleteRow: (rowId) => deleteRow(rowId),
};


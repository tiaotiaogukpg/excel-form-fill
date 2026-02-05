import express from "express";
import cookieParser from "cookie-parser";
import path from "node:path";
import { fileURLToPath } from "node:url";
import bcrypt from "bcryptjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

import { ensureSeedAdmin, findUserByUsername, findUserById, signToken, requireAuth, requireRole, ROLES, createUser, updateUserPassword, deleteUser } from "./auth.js";
import { loadStateForUser, saveStateForUser, loadUsers, deleteStateForUser } from "./storage.js";

const app = express();
app.use(express.json({ limit: "2mb" }));
app.use(cookieParser());

ensureSeedAdmin();

app.post("/api/auth/login", (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: "缺少用户名或密码" });
  const u = findUserByUsername(String(username));
  if (!u) return res.status(401).json({ error: "用户名或密码错误" });
  const ok = bcrypt.compareSync(String(password), u.password_hash);
  if (!ok) return res.status(401).json({ error: "用户名或密码错误" });

  const token = signToken({ id: u.id, username: u.username, role: u.role });
  res.cookie("token", token, {
    httpOnly: true,
    sameSite: "lax",
  });
  res.json({ ok: true, user: { id: u.id, username: u.username, role: u.role } });
});

app.post("/api/auth/logout", (req, res) => {
  res.clearCookie("token");
  res.json({ ok: true });
});

// 自助注册为“老师”：如果账号不存在则创建；若已存在且密码正确则直接登录
app.post("/api/auth/register", async (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: "缺少用户名或密码" });
  try {
    const name = String(username);
    const pass = String(password);

    const existing = findUserByUsername(name);
    if (existing) {
      const ok = bcrypt.compareSync(pass, existing.password_hash);
      if (!ok) return res.status(401).json({ error: "用户名已存在且密码不匹配" });
      const token = signToken({ id: existing.id, username: existing.username, role: existing.role });
      res.cookie("token", token, {
        httpOnly: true,
        sameSite: "lax",
      });
      return res.json({ ok: true, user: { id: existing.id, username: existing.username, role: existing.role } });
    }

    const u = createUser({ username: name, password: pass, role: ROLES.teacher });
    const token = signToken({ id: u.id, username: u.username, role: u.role });
    res.cookie("token", token, {
      httpOnly: true,
      sameSite: "lax",
    });
    res.json({ ok: true, user: u });
  } catch (e) {
    res.status(400).json({ error: e?.message ?? String(e) });
  }
});

app.get("/api/me", requireAuth, (req, res) => {
  res.json({ user: req.user });
});

// 当前用户修改自己的密码
app.post("/api/auth/change-password", requireAuth, (req, res) => {
  const { oldPassword, newPassword } = req.body || {};
  if (!oldPassword || !newPassword) return res.status(400).json({ error: "请填写当前密码和新密码" });
  const u = findUserById(req.user.id);
  if (!u) return res.status(401).json({ error: "用户不存在" });
  if (!bcrypt.compareSync(String(oldPassword), u.password_hash)) {
    return res.status(401).json({ error: "当前密码错误" });
  }
  try {
    updateUserPassword(req.user.id, newPassword);
    res.json({ ok: true });
  } catch (e) {
    return res.status(400).json({ error: e?.message ?? String(e) });
  }
});

// 空 state 占位，仅用于管理员查看某账号时无文件情况，保证前端拿到统一结构
const EMPTY_STATE = { rows: [], selectedClass: "", selectedCourse: "", isSample: false };

// 账号隔离的状态读写（每个账号一份）；读写均用规范化 userId 保证一致
app.get("/api/state", requireAuth, (req, res) => {
  const uid = normalizeUserId(req.user.id);
  if (!uid) return res.status(400).json({ error: "无效的 userId" });
  const st = loadStateForUser(uid);
  res.json({ state: st });
});

app.put("/api/state", requireAuth, (req, res) => {
  const { state } = req.body || {};
  if (!state || typeof state !== "object") return res.status(400).json({ error: "state 必须是对象" });
  const uid = normalizeUserId(req.user.id);
  if (!uid) return res.status(400).json({ error: "无效的 userId" });
  saveStateForUser(uid, state);
  res.json({ ok: true });
});

// 管理员：创建老师账号
app.post("/api/admin/users", requireAuth, requireRole(ROLES.admin), (req, res) => {
  const { username, password, role } = req.body || {};
  const r = role === ROLES.admin ? ROLES.admin : ROLES.teacher;
  if (!username || !password) return res.status(400).json({ error: "缺少 username/password" });
  try {
    const u = createUser({ username: String(username), password: String(password), role: r });
    res.json({ ok: true, user: u });
  } catch (e) {
    res.status(400).json({ error: e?.message ?? String(e) });
  }
});

// 管理员：查看账号列表（不返回密码哈希）
app.get("/api/admin/users", requireAuth, requireRole(ROLES.admin), (req, res) => {
  const doc = loadUsers();
  const users = (doc.users || []).map((u) => ({
    id: u.id,
    username: u.username,
    role: u.role,
    created_at: u.created_at,
  }));
  res.json({ users });
});

// 管理员：强制修改某账号密码
app.post("/api/admin/users/:userId/password", requireAuth, requireRole(ROLES.admin), (req, res) => {
  const uid = normalizeUserId(req.params.userId);
  if (!uid) return res.status(400).json({ error: "缺少或无效的 userId" });
  const { newPassword } = req.body || {};
  if (!newPassword || String(newPassword).trim().length < 1) {
    return res.status(400).json({ error: "请填写新密码" });
  }
  try {
    updateUserPassword(uid, String(newPassword).trim());
    res.json({ ok: true });
  } catch (e) {
    return res.status(400).json({ error: e?.message ?? String(e) });
  }
});

// 管理员：删除账号（同时删除该账号的 state 数据）
app.delete("/api/admin/users/:userId", requireAuth, requireRole(ROLES.admin), (req, res) => {
  const uid = normalizeUserId(req.params.userId);
  if (!uid) return res.status(400).json({ error: "缺少或无效的 userId" });
  if (uid === req.user.id) return res.status(400).json({ error: "不能删除当前登录的账号" });
  try {
    deleteUser(uid);
    deleteStateForUser(uid);
    res.json({ ok: true });
  } catch (e) {
    return res.status(400).json({ error: e?.message ?? String(e) });
  }
});

// 规范化 userId，防止路径穿越
function normalizeUserId(userId) {
  const s = String(userId ?? "").trim();
  if (!s || s.includes("..") || s.includes("/") || s.includes("\\")) return null;
  return s;
}

// 管理员：查看指定账号的成绩录入状态（无文件时返回空 state，便于前端统一展示）
app.get("/api/admin/state/:userId", requireAuth, requireRole(ROLES.admin), (req, res) => {
  const uid = normalizeUserId(req.params.userId);
  if (!uid) return res.status(400).json({ error: "缺少或无效的 userId" });
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate");
  const st = loadStateForUser(uid);
  res.json({ state: st ?? EMPTY_STATE });
});

// 管理员：查看某账号 state 摘要（便于排查“管理员看不到数据”）
app.get("/api/admin/state-meta/:userId", requireAuth, requireRole(ROLES.admin), (req, res) => {
  const uid = normalizeUserId(req.params.userId);
  if (!uid) return res.status(400).json({ error: "缺少或无效的 userId" });
  const st = loadStateForUser(uid);
  const hasFile = st !== null;
  const rowCount = (st && Array.isArray(st.rows)) ? st.rows.length : 0;
  res.json({ userId: uid, hasFile, rowCount });
});

// 静态资源：web/（相对 server.js 所在目录的上一级）
const WEB_DIR = path.join(__dirname, "..", "web");
app.use(express.static(WEB_DIR));

// 任何其它路径：默认回到登录页/主页
app.get("*", (req, res) => {
  res.sendFile(path.join(WEB_DIR, "login.html"));
});

const port = Number(process.env.PORT || 5173);
app.listen(port, () => {
  // eslint-disable-next-line no-console
  console.log(`server listening on http://localhost:${port}`);
});


import crypto from "node:crypto";
import jwt from "jsonwebtoken";
import bcrypt from "bcryptjs";
import { loadUsers, saveUsers } from "./storage.js";

export const ROLES = /** @type {const} */ ({
  admin: "admin",
  teacher: "teacher",
});

export function newId(prefix = "U") {
  return `${prefix}_${crypto.randomBytes(10).toString("hex")}`;
}

export function getJwtSecret() {
  return process.env.JWT_SECRET || "dev-secret-change-me";
}

export function signToken(payload) {
  return jwt.sign(payload, getJwtSecret(), { expiresIn: "7d" });
}

export function verifyToken(token) {
  return jwt.verify(token, getJwtSecret());
}

export function getAuthUser(req) {
  const token = req.cookies?.token;
  if (!token) return null;
  try {
    const decoded = verifyToken(token);
    return decoded;
  } catch {
    return null;
  }
}

export function requireAuth(req, res, next) {
  const u = getAuthUser(req);
  if (!u) return res.status(401).json({ error: "UNAUTHENTICATED" });
  req.user = u;
  next();
}

export function requireRole(role) {
  return (req, res, next) => {
    const u = req.user;
    if (!u) return res.status(401).json({ error: "UNAUTHENTICATED" });
    if (u.role !== role) return res.status(403).json({ error: "FORBIDDEN" });
    next();
  };
}

export function ensureSeedAdmin() {
  const adminUser = process.env.SEED_ADMIN_USER || "admin";
  const adminPass = process.env.SEED_ADMIN_PASS || "admin123";

  const doc = loadUsers();
  const exists = doc.users.some((u) => u.username === adminUser);
  if (exists) return;

  const hash = bcrypt.hashSync(adminPass, 10);
  doc.users.push({
    id: newId("U"),
    username: adminUser,
    password_hash: hash,
    role: ROLES.admin,
    created_at: Date.now(),
  });
  saveUsers(doc);
  // eslint-disable-next-line no-console
  console.log(`[seed] created admin user: ${adminUser} / ${adminPass}`);
}

export function findUserByUsername(username) {
  const doc = loadUsers();
  return doc.users.find((u) => u.username === username) || null;
}

export function findUserById(id) {
  const doc = loadUsers();
  return doc.users.find((u) => u.id === id) || null;
}

export function updateUserPassword(userId, newPassword) {
  const doc = loadUsers();
  const u = doc.users.find((x) => x.id === userId);
  if (!u) throw new Error("用户不存在");
  u.password_hash = bcrypt.hashSync(String(newPassword), 10);
  saveUsers(doc);
}

export function deleteUser(userId) {
  const doc = loadUsers();
  const idx = doc.users.findIndex((u) => u.id === userId);
  if (idx < 0) throw new Error("用户不存在");
  const admins = doc.users.filter((u) => u.role === ROLES.admin);
  if (admins.length === 1 && doc.users[idx].role === ROLES.admin) {
    throw new Error("不能删除最后一个管理员");
  }
  doc.users.splice(idx, 1);
  saveUsers(doc);
}

export function createUser({ username, password, role }) {
  const doc = loadUsers();
  if (doc.users.some((u) => u.username === username)) {
    throw new Error("用户名已存在");
  }
  const hash = bcrypt.hashSync(password, 10);
  const u = {
    id: newId("U"),
    username,
    password_hash: hash,
    role,
    created_at: Date.now(),
  };
  doc.users.push(u);
  saveUsers(doc);
  return { id: u.id, username: u.username, role: u.role };
}


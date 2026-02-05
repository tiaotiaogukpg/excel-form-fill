import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = path.join(__dirname, "data");
const USERS_PATH = path.join(DATA_DIR, "users.json");
const STATES_DIR = path.join(DATA_DIR, "states");

function ensureDirs() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.mkdirSync(STATES_DIR, { recursive: true });
}

export function readJsonSafe(p, fallback) {
  try {
    const s = fs.readFileSync(p, "utf-8");
    return JSON.parse(s);
  } catch {
    return fallback;
  }
}

export function writeJsonAtomic(p, obj) {
  const tmp = `${p}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(obj, null, 2), "utf-8");
  fs.renameSync(tmp, p);
}

export function loadUsers() {
  ensureDirs();
  return readJsonSafe(USERS_PATH, { users: [] });
}

export function saveUsers(usersDoc) {
  ensureDirs();
  writeJsonAtomic(USERS_PATH, usersDoc);
}

export function userStatePath(userId) {
  ensureDirs();
  return path.join(STATES_DIR, `${userId}.json`);
}

export function loadStateForUser(userId) {
  const p = userStatePath(userId);
  return readJsonSafe(p, null);
}

export function saveStateForUser(userId, state) {
  const p = userStatePath(userId);
  writeJsonAtomic(p, state);
}

export function deleteStateForUser(userId) {
  const p = userStatePath(userId);
  try {
    fs.unlinkSync(p);
  } catch {
    // 文件不存在则忽略
  }
}


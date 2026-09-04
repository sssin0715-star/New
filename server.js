/**
 * 신규입사자 면담 예약 — Node.js 서버
 *
 * 외부 의존성이 없습니다. Node 18 이상이면 `node server.js` 로 바로 뜹니다.
 * 예약 데이터는 data/store.json 한 파일에 저장하고, 쓰기는 임시 파일에 쓴 뒤
 * 원자적으로 교체합니다.
 *
 * 환경 변수
 *   PORT        포트 (기본 3000)
 *   ADMIN_CODE  관리자 코드 (없으면 첫 실행 때 만들어 store.json 에 저장)
 *   DATA_DIR    저장 폴더 (기본 ./data) — 호스팅의 영구 디스크를 여기로 잡으세요
 */

import { createServer } from "node:http";
import { readFile, writeFile, rename, mkdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import { randomUUID, randomBytes } from "node:crypto";
import { join, dirname, extname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));

const PORT = Number(process.env.PORT) || 3000;
const DATA_DIR = process.env.DATA_DIR || join(__dirname, "data");
const STORE_PATH = join(DATA_DIR, "store.json");
const PUBLIC_DIR = join(__dirname, "public");

const DOW = ["일", "월", "화", "수", "목", "금", "토"];
const MAX_BODY = 16 * 1024;

/* ========================= 저장소 ========================= */

const DEFAULT_STORE = {
  config: {
    active: false,
    startDate: "",
    endDate: "",
    weekdays: [1, 2, 3, 4, 5],
    startTime: "10:00",
    endTime: "17:00",
    slotMinutes: 45,
    breakStart: "",
    breakEnd: "",
    durationLabel: "30~40분",
    locationLabel: "메일로 안내",
    blocked: [],
  },
  bookings: {}, // slotId -> { name, email, phone, bookedAt, cancelCode }
  adminCode: "",
};

let store = structuredClone(DEFAULT_STORE);

/** 쓰기를 한 줄로 세워 마지막 쓰기가 이전 쓰기를 덮어쓰지 않게 합니다. */
let writeChain = Promise.resolve();

async function loadStore() {
  await mkdir(DATA_DIR, { recursive: true });
  if (existsSync(STORE_PATH)) {
    try {
      const parsed = JSON.parse(await readFile(STORE_PATH, "utf8"));
      store = {
        config: { ...DEFAULT_STORE.config, ...(parsed.config || {}) },
        bookings: parsed.bookings || {},
        adminCode: parsed.adminCode || "",
      };
    } catch {
      console.error("store.json 을 읽지 못해 기본값으로 시작합니다.");
    }
  }
  if (!store.adminCode) {
    store.adminCode = process.env.ADMIN_CODE || randomBytes(4).toString("hex");
    await persist();
  }
  console.log(`관리자 코드: ${store.adminCode}`);
}

function persist() {
  writeChain = writeChain.then(async () => {
    const tmp = `${STORE_PATH}.${process.pid}.tmp`;
    await writeFile(tmp, JSON.stringify(store, null, 2), "utf8");
    await rename(tmp, STORE_PATH);
  });
  return writeChain;
}

/* 예약 확정은 "확인 후 기록"이 끊기면 안 되므로 한 번에 하나씩 처리합니다. */
let bookingGate = Promise.resolve();
function serialize(fn) {
  const run = bookingGate.then(fn, fn);
  bookingGate = run.then(
    () => undefined,
    () => undefined,
  );
  return run;
}

/* ========================= 날짜·슬롯 ========================= */

const pad2 = (n) => String(n).padStart(2, "0");

function todayStr(d = new Date()) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
function nowToken(d = new Date()) {
  return `${todayStr(d)}T${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}
function parseDate(dateStr) {
  const [y, m, d] = String(dateStr).split("-").map(Number);
  return new Date(y, m - 1, d);
}
function timeToMin(t) {
  const [h, m] = String(t).split(":").map(Number);
  return h * 60 + m;
}
const minToTime = (m) => `${pad2(Math.floor(m / 60))}:${pad2(m % 60)}`;

/** 설정이 만들어내는 모든 슬롯 id. 예약 요청 검증의 기준입니다. */
function allSlotIds(config) {
  const ids = new Set();
  if (!config.active || !config.startDate || !config.endDate) return ids;

  const times = [];
  const hasBreak = Boolean(config.breakStart && config.breakEnd);
  const bs = hasBreak ? timeToMin(config.breakStart) : 0;
  const be = hasBreak ? timeToMin(config.breakEnd) : 0;
  const endMin = timeToMin(config.endTime);
  for (let m = timeToMin(config.startTime); m + config.slotMinutes <= endMin; m += config.slotMinutes) {
    if (hasBreak && m < be && m + config.slotMinutes > bs) continue;
    times.push(minToTime(m));
  }

  const cursor = parseDate(config.startDate);
  const end = parseDate(config.endDate);
  for (let guard = 0; cursor <= end && guard < 400; guard += 1) {
    if (config.weekdays.includes(cursor.getDay())) {
      const ds = todayStr(cursor);
      for (const t of times) ids.add(`${ds}T${t}`);
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return ids;
}

function fmtWhen(slotId) {
  const [datePart, timePart] = String(slotId).split("T");
  const d = parseDate(datePart);
  return `${d.getMonth() + 1}월 ${d.getDate()}일 (${DOW[d.getDay()]}) ${timePart}`;
}

/* ========================= 검증 ========================= */

const isEmail = (s) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s));
const clean = (v, max) => String(v ?? "").trim().slice(0, max);

const TIME_RE = /^([01]\d|2[0-3]):[0-5]\d$/;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

function validateConfig(next) {
  if (!DATE_RE.test(next.startDate) || !DATE_RE.test(next.endDate)) return "날짜 형식을 확인해주세요.";
  if (next.startDate > next.endDate) return "시작일이 종료일보다 늦을 수 없어요.";
  if (!Array.isArray(next.weekdays) || next.weekdays.length === 0) return "요일을 하나 이상 선택해주세요.";
  if (next.weekdays.some((d) => !Number.isInteger(d) || d < 0 || d > 6)) return "요일 값이 올바르지 않아요.";
  if (!TIME_RE.test(next.startTime) || !TIME_RE.test(next.endTime)) return "시각 형식을 확인해주세요.";
  if (timeToMin(next.startTime) >= timeToMin(next.endTime)) return "시작 시각이 종료 시각보다 빨라야 해요.";
  if (![30, 45, 60].includes(next.slotMinutes)) return "간격은 30·45·60분 중에서 선택해주세요.";
  if (next.breakStart || next.breakEnd) {
    if (!TIME_RE.test(next.breakStart) || !TIME_RE.test(next.breakEnd)) return "제외 시간대 형식을 확인해주세요.";
    if (timeToMin(next.breakStart) >= timeToMin(next.breakEnd)) return "제외 시작이 종료보다 빨라야 해요.";
  }
  return null;
}

/* ========================= 요청 처리 ========================= */

/** 공개 상태 — 개인정보 없이 "찬 시간"만 내려갑니다. */
function publicState() {
  return {
    config: store.config,
    taken: Object.keys(store.bookings),
    serverNow: nowToken(),
  };
}

function adminState() {
  return {
    ok: true,
    config: store.config,
    bookings: Object.entries(store.bookings)
      .map(([id, b]) => ({ id, name: b.name, email: b.email, phone: b.phone, bookedAt: b.bookedAt }))
      .sort((a, b) => a.id.localeCompare(b.id)),
    serverNow: nowToken(),
  };
}

function requireAdmin(body) {
  return Boolean(store.adminCode) && clean(body.code, 64) === store.adminCode;
}

async function book(body) {
  const slotId = clean(body.slotId, 32);
  const name = clean(body.name, 20);
  const email = clean(body.email, 60);
  const phone = clean(body.phone, 20);

  if (!name) return { ok: false, message: "이름을 입력해주세요." };
  if (!isEmail(email)) return { ok: false, message: "이메일 주소를 확인해주세요." };
  if (phone.replace(/\D/g, "").length < 9) return { ok: false, message: "연락처를 확인해주세요." };

  return serialize(async () => {
    const { config, bookings } = store;
    if (!config.active) return { ok: false, message: "지금은 예약을 받지 않고 있어요." };
    if (!allSlotIds(config).has(slotId)) return { ok: false, message: "선택할 수 없는 시간이에요. 새로고침해주세요." };
    if (config.blocked.includes(slotId)) return { ok: false, message: "담당자가 마감한 시간이에요." };
    if (slotId < nowToken()) return { ok: false, message: "이미 지난 시간이에요." };
    if (bookings[slotId]) return { ok: false, message: "방금 다른 분이 예약했어요. 다른 시간을 선택해주세요." };

    const cancelCode = randomUUID();
    bookings[slotId] = { name, email, phone, cancelCode, bookedAt: new Date().toISOString() };
    await persist();
    console.log(`예약 확정: ${fmtWhen(slotId)} · ${name}`);
    return { ok: true, cancelCode };
  });
}

async function cancelMine(body) {
  const slotId = clean(body.slotId, 32);
  const cancelCode = clean(body.cancelCode, 64);
  return serialize(async () => {
    const found = store.bookings[slotId];
    if (!found || !cancelCode || found.cancelCode !== cancelCode) {
      return { ok: false, message: "예약을 찾지 못했어요. 새로고침해주세요." };
    }
    delete store.bookings[slotId];
    await persist();
    return { ok: true };
  });
}

async function handleApi(path, body) {
  switch (path) {
    case "/api/state":
      return publicState();

    case "/api/book":
      return book(body);

    case "/api/cancel":
      return cancelMine(body);

    case "/api/admin/state":
      return requireAdmin(body) ? adminState() : { ok: false, message: "코드가 맞지 않아요." };

    case "/api/admin/config": {
      if (!requireAdmin(body)) return { ok: false, message: "코드가 맞지 않아요." };
      const incoming = body.config || {};
      const next = {
        active: true,
        startDate: clean(incoming.startDate, 10),
        endDate: clean(incoming.endDate, 10),
        weekdays: Array.isArray(incoming.weekdays) ? incoming.weekdays.map(Number) : [],
        startTime: clean(incoming.startTime, 5),
        endTime: clean(incoming.endTime, 5),
        slotMinutes: Number(incoming.slotMinutes),
        breakStart: clean(incoming.breakStart, 5),
        breakEnd: clean(incoming.breakEnd, 5),
        durationLabel: clean(incoming.durationLabel, 20) || "30~40분",
        locationLabel: clean(incoming.locationLabel, 24) || "메일로 안내",
        blocked: store.config.blocked,
      };
      const problem = validateConfig(next);
      if (problem) return { ok: false, message: problem };
      store.config = next;
      await persist();
      return { ok: true, state: adminState() };
    }

    case "/api/admin/block": {
      if (!requireAdmin(body)) return { ok: false, message: "코드가 맞지 않아요." };
      const slotId = clean(body.slotId, 32);
      const blocked = new Set(store.config.blocked);
      if (body.block) blocked.add(slotId);
      else blocked.delete(slotId);
      store.config = { ...store.config, blocked: [...blocked] };
      await persist();
      return { ok: true, state: adminState() };
    }

    case "/api/admin/cancel": {
      if (!requireAdmin(body)) return { ok: false, message: "코드가 맞지 않아요." };
      const slotId = clean(body.slotId, 32);
      if (!store.bookings[slotId]) return { ok: false, message: "예약을 찾지 못했어요." };
      delete store.bookings[slotId];
      await persist();
      return { ok: true, state: adminState() };
    }

    default:
      return null;
  }
}

/* ========================= 간단한 호출 제한 ========================= */

const hits = new Map(); // ip -> { count, resetAt }
const WINDOW_MS = 60_000;
const MAX_HITS = 120;

function rateLimited(ip) {
  const now = Date.now();
  const entry = hits.get(ip);
  if (!entry || now > entry.resetAt) {
    hits.set(ip, { count: 1, resetAt: now + WINDOW_MS });
    return false;
  }
  entry.count += 1;
  return entry.count > MAX_HITS;
}

setInterval(() => {
  const now = Date.now();
  for (const [ip, entry] of hits) if (now > entry.resetAt) hits.delete(ip);
}, WINDOW_MS).unref();

/* ========================= 정적 파일 ========================= */

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".png": "image/png",
};

async function serveStatic(pathname, res) {
  const rel = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
  // 경로 탈출 차단
  const target = join(PUBLIC_DIR, rel);
  if (!target.startsWith(PUBLIC_DIR)) {
    res.writeHead(403).end("Forbidden");
    return;
  }
  try {
    const data = await readFile(target);
    res.writeHead(200, {
      "content-type": MIME[extname(target)] || "application/octet-stream",
      "cache-control": "no-cache",
    });
    res.end(data);
  } catch {
    res.writeHead(404, { "content-type": "text/plain; charset=utf-8" });
    res.end("페이지를 찾을 수 없어요.");
  }
}

/* ========================= 서버 ========================= */

function readBody(req) {
  return new Promise((resolve, reject) => {
    let raw = "";
    req.on("data", (chunk) => {
      raw += chunk;
      if (raw.length > MAX_BODY) {
        reject(new Error("body too large"));
        req.destroy();
      }
    });
    req.on("end", () => {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch {
        reject(new Error("bad json"));
      }
    });
    req.on("error", reject);
  });
}

const server = createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
  const ip = (req.headers["x-forwarded-for"]?.split(",")[0] || req.socket.remoteAddress || "?").trim();

  if (url.pathname.startsWith("/api/")) {
    if (rateLimited(ip)) {
      res.writeHead(429, { "content-type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: false, message: "요청이 너무 잦아요. 잠시 후 다시 시도해주세요." }));
      return;
    }

    let body = {};
    if (req.method === "POST") {
      try {
        body = await readBody(req);
      } catch {
        res.writeHead(400, { "content-type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ ok: false, message: "요청을 읽지 못했어요." }));
        return;
      }
    } else if (req.method !== "GET") {
      res.writeHead(405).end();
      return;
    }

    let result;
    try {
      result = await handleApi(url.pathname, body);
    } catch (err) {
      console.error(err);
      res.writeHead(500, { "content-type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: false, message: "서버에서 문제가 생겼어요." }));
      return;
    }

    if (result === null) {
      res.writeHead(404, { "content-type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: false, message: "없는 요청이에요." }));
      return;
    }

    res.writeHead(200, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
    res.end(JSON.stringify(result));
    return;
  }

  if (req.method !== "GET") {
    res.writeHead(405).end();
    return;
  }
  await serveStatic(url.pathname, res);
});

await loadStore();
server.listen(PORT, () => {
  console.log(`면담 예약 서버가 http://localhost:${PORT} 에서 실행 중입니다.`);
});

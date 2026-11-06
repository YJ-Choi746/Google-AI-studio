import {
  db,
  auth,
  googleProvider,
  GoogleAuthProvider,
  signInWithPopup,
  signOut,
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot,
  handleFirestoreError,
  OperationType,
  getCachedAccessToken,
  setCachedAccessToken
} from './firebase.ts';
import {
  createApplicantSpreadsheet,
  appendApplicantToSheet,
  batchSyncAllApplicants,
  GoogleSheetConfig
} from './sheets.ts';

interface HistoryItem {
  at: string;
  action: string;
  by: string;
  changes?: { field: string; from: string; to: string }[];
}

interface ApplicationRecord {
  appId: string;
  name: string;
  org: string;
  email: string;
  phone: string;
  message: string;
  ageOver14: boolean;
  smsOptIn: boolean;
  phoneVerified: boolean;
  consent: boolean;
  consentAt: string;
  consentVersion: string;
  appliedAt: string;
  updatedAt: string | null;
  source: 'form' | 'admin';
  version: number;
  questionStatus: 'none' | 'open' | 'answered';
  adminMemo: string;
  history: HistoryItem[];
}

interface LogItem {
  at: string;
  action: string;
  id: string;
  who: string;
  detail: string;
}

const CONSENT_VERSION = "v2-2026-11";
const DEADLINE = new Date("2026-12-15T23:59:59.999+09:00").getTime();
const GRACE_END = new Date("2026-12-16T00:00:00.999+09:00").getTime();
const PURGE_AFTER = new Date("2027-03-16T00:00:00+09:00").getTime();
const PHONE_RE = /^010\d{8}$/;
const NAME_RE = /^[A-Za-z\u00C0-\u024F가-힣 '’.\-]+$/;
const JAMO_RE = /[ㄱ-ㅎㅏ-ㅣ]/;
const ORG_RE = /^[A-Za-z가-힣0-9 ()&\-]+$/;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const NAME_MAX = 40;
const CODE_TTL = 3 * 60 * 1000;
const LOCAL_STORE_KEY = "selectmarket-preorder-2026";
const openedAt = Date.now();

const ADMIN_EMAIL = "yeonj445@gmail.com";
const ADMIN_PW_HASH = "4334efbdc5e026d5b30d838a0792f0aa2feb9876e633da9473c295d37083dc57";

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const esc = (s: unknown) => String(s ?? '').replace(/[&<>"']/g, (ch) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[ch] || ch));
const kst = (iso: string) => new Intl.DateTimeFormat("ko-KR", {timeZone:"Asia/Seoul", year:"numeric", month:"long", day:"numeric", hour:"2-digit", minute:"2-digit"}).format(new Date(iso));
const kstShort = (iso: string) => new Intl.DateTimeFormat("ko-KR", {timeZone:"Asia/Seoul", month:"numeric", day:"numeric", hour:"2-digit", minute:"2-digit", hour12:false}).format(new Date(iso));

const mask = (p: string) => String(p).slice(0,3) + "*".repeat(Math.max(0, String(p).length - 3));
function maskName(n: string) {
  const s = Array.from(String(n));
  if (s.length <= 1) return s.join("");
  if (s.length === 2) return s[0] + "*";
  return s.map((ch, i) => (i === 0 || i === s.length - 1 || ch === " ") ? ch : "*").join("");
}
const fmtPhone = (p: string) => p.replace(/^(\d{3})(\d{4})(\d{4})$/, "$1-$2-$3");
const canSubmit = (now: number) => now <= DEADLINE || (openedAt <= DEADLINE && now <= GRACE_END);
const newId = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(16) + Math.random().toString(16).slice(2));
const normName = (s: string) => s.trim().replace(/\s+/g, " ");

function normalize(r: Partial<ApplicationRecord>): ApplicationRecord {
  return {
    appId: r.appId || newId(),
    name: r.name || "",
    org: r.org || "",
    email: r.email || "",
    phone: r.phone || "",
    message: r.message || "",
    ageOver14: r.ageOver14 ?? true,
    smsOptIn: r.smsOptIn ?? false,
    phoneVerified: r.phoneVerified ?? false,
    consent: r.consent ?? true,
    consentAt: r.consentAt || new Date().toISOString(),
    consentVersion: r.consentVersion || CONSENT_VERSION,
    appliedAt: r.appliedAt || new Date().toISOString(),
    updatedAt: r.updatedAt || null,
    source: r.source || "form",
    version: r.version || 1,
    questionStatus: r.questionStatus || (r.message ? "open" : "none"),
    adminMemo: r.adminMemo || "",
    history: r.history || []
  };
}

let firestoreApplications: ApplicationRecord[] = [];
let firestoreLogs: LogItem[] = [];
let unsubscribeApplications: (() => void) | null = null;
let unsubscribeLogs: (() => void) | null = null;
let currentSheetConfig: GoogleSheetConfig | null = null;
let unsubscribeSheetConfig: (() => void) | null = null;
const knownAppIds = new Set<string>();
let isInitialAppsLoad = true;

function updateSheetUI() {
  const badge = $("sheet-status-badge");
  const desc = $("sheet-desc");
  const openLink = $<HTMLAnchorElement>("sheet-open-link");
  const syncBtn = $<HTMLButtonElement>("sheet-sync-all-btn");
  const connectBtn = $<HTMLButtonElement>("sheet-connect-btn");

  if (!badge || !desc || !openLink || !syncBtn || !connectBtn) return;

  if (currentSheetConfig && currentSheetConfig.spreadsheetId) {
    badge.className = "pill";
    badge.textContent = "실시간 연동 중";
    const count = currentSheetConfig.syncedCount ?? firestoreApplications.length;
    desc.textContent = `연동된 구글 시트: ${currentSheetConfig.sheetName || '사전신청자'} (${count}건 기록됨) · 새 신청 즉시 자동 누적`;
    openLink.href = currentSheetConfig.spreadsheetUrl;
    openLink.style.display = "inline-flex";
    syncBtn.style.display = "inline-flex";
    connectBtn.textContent = "시트 재연동 / 새 시트 생성";
  } else {
    badge.className = "pill off";
    badge.textContent = "미연동";
    desc.textContent = "신청 즉시 구글 스프레드시트에 행으로 실시간 자동 누적됩니다.";
    openLink.style.display = "none";
    syncBtn.style.display = "none";
    connectBtn.textContent = "구글 시트 연동 / 새 시트 생성";
  }
}

function loadLocalAll(): ApplicationRecord[] {
  try {
    const l = JSON.parse(localStorage.getItem(LOCAL_STORE_KEY) || "[]");
    return Array.isArray(l) ? l.map(normalize) : [];
  } catch {
    return [];
  }
}

function saveLocalAll(list: ApplicationRecord[]) {
  try {
    localStorage.setItem(LOCAL_STORE_KEY, JSON.stringify(list));
  } catch {}
}

function getAllApplications(): ApplicationRecord[] {
  if (firestoreApplications.length > 0) {
    return firestoreApplications;
  }
  return loadLocalAll();
}

function checkName(v: string) {
  if (!v) return "이름을 입력해 주세요.";
  if (JAMO_RE.test(v)) return "이름은 완성된 글자로 입력해 주세요.";
  if (!NAME_RE.test(v)) return "이름에는 한글, 영문, 공백과 ’ - . 만 쓸 수 있어요.";
  if (v.length > NAME_MAX) return "이름은 40자까지 입력할 수 있어요.";
  return "";
}

const checkOrg = (v: string) => (v && !ORG_RE.test(v) ? "소속에는 한글, 영문, 숫자, 공백과 ( ) & - 만 쓸 수 있어요." : "");

function checkEmail(v: string) {
  if (!v) return "이메일을 입력해 주세요.";
  if (!v.includes("@")) return "이메일에 @가 빠졌어요. 예: name@example.com";
  if (!EMAIL_RE.test(v)) return "이메일 형식이 맞지 않아요. 예: name@example.com";
  return "";
}

const checkPhone = (v: string) => (!v ? "휴대폰 번호를 입력해 주세요." : !PHONE_RE.test(v) ? "010으로 시작하는 숫자 11자리로 입력해 주세요." : "");

function findDup(list: ApplicationRecord[], r: { name: string; email: string; phone: string }, exceptId: string | null) {
  const key = r.name.toLowerCase() + "|" + r.email.toLowerCase();
  for (const x of list) {
    if (x.appId === exceptId) continue;
    if (String(x.name).toLowerCase() + "|" + x.email.toLowerCase() === key) return "ne";
  }
  for (const x of list) {
    if (x.appId === exceptId) continue;
    if (x.phone === r.phone) return "phone";
  }
  return null;
}

function setErr(fieldId: string, errId: string | null, msg: string) {
  if (errId && $(errId)) $(errId).textContent = msg || "";
  if ($(fieldId)) $(fieldId).classList.toggle("invalid", !!msg);
  return !!msg;
}

function attachNameClean(el: HTMLInputElement, onCut?: () => void) {
  const clean = () => {
    let v = el.value.replace(/[0-9]/g, "");
    if (v.length > NAME_MAX) { v = v.slice(0, NAME_MAX); if (onCut) onCut(); }
    if (v !== el.value) el.value = v;
  };
  el.addEventListener("input", (e) => { if (!(e as InputEvent).isComposing) clean(); });
  el.addEventListener("compositionend", clean);
  return clean;
}

function attachPhoneClean(el: HTMLInputElement, onChange?: (v: string) => void) {
  el.addEventListener("input", () => {
    const v = el.value.replace(/\D/g, "").slice(0, 11);
    if (v !== el.value) el.value = v;
    if (onChange) onChange(v);
  });
}

function route() {
  const view = location.hash === "#admin" ? "admin" : "apply";
  $("view-apply").hidden = view !== "apply";
  $("view-admin").hidden = view !== "admin";
  $("tab-apply").setAttribute("aria-selected", String(view === "apply"));
  $("tab-admin").setAttribute("aria-selected", String(view === "admin"));
  $("wrap").classList.toggle("wide", view === "admin");
  document.title = view === "admin" ? "신청자 관리 · 셀렉트마켓 연말 이벤트" : "셀렉트마켓 연말 이벤트 사전 신청";
  if (view === "admin") {
    applyLock();
  } else {
    lockAdmin();
  }
}
window.addEventListener("hashchange", route);

// =================== 사전 신청 ===================
const form = $<HTMLFormElement>("form");
const formmsg = $("formmsg");
const submitBtn = $<HTMLButtonElement>("submit");
let closedInline = false;
let shownId: string | null = null;
let verify = { code: null as string | null, expires: 0, phone: null as string | null, done: false };

function setMsg(text: string) {
  formmsg.textContent = text || "";
  formmsg.classList.toggle("show", !!text);
}

function renderDday() {
  const now = Date.now();
  if (!canSubmit(now)) {
    $("dday").textContent = "사전 신청이 마감되었어요.";
    return;
  }
  const dayKey = (t: number) => new Intl.DateTimeFormat("en-CA", {timeZone:"Asia/Seoul"}).format(new Date(t));
  const days = Math.round((new Date(dayKey(DEADLINE)).getTime() - new Date(dayKey(now)).getTime()) / 86400000);
  $("dday").textContent = days > 0 ? `마감까지 ${days}일 남았어요.` : "오늘 밤 11시 59분에 마감돼요.";
}

function showClosed() {
  form.hidden = true;
  $("done").hidden = true;
  $("closed").hidden = false;
  renderDday();
}

function closeInline() {
  closedInline = true;
  submitBtn.disabled = true;
  submitBtn.textContent = "신청 마감";
  setMsg("마감되었어요.");
  renderDday();
}

const nameEl = $<HTMLInputElement>("name");
const cleanName = attachNameClean(nameEl, () => { $("name-cut").hidden = false; });
nameEl.addEventListener("input", () => $("f-name").classList.remove("invalid"));
$<HTMLInputElement>("org").addEventListener("input", () => $("f-org").classList.remove("invalid"));
$<HTMLInputElement>("email").addEventListener("input", () => $("f-email").classList.remove("invalid"));
attachPhoneClean($<HTMLInputElement>("phone"), (v) => {
  $("f-phone").classList.remove("invalid");
  if (verify.done && verify.phone !== v) resetVerify();
});

function resetVerify() {
  verify = { code: null, expires: 0, phone: null, done: false };
  $("verify-ok").hidden = true;
  $("demo-code").hidden = true;
  $<HTMLInputElement>("code").value = "";
  $<HTMLInputElement>("code").disabled = false;
  $<HTMLButtonElement>("check-code").disabled = false;
  $<HTMLButtonElement>("send-code").disabled = false;
  $("send-code").textContent = "인증번호 받기";
}

$<HTMLInputElement>("sms").addEventListener("change", (e) => {
  const target = e.target as HTMLInputElement;
  $("f-verify").hidden = !target.checked;
  $("f-verify").classList.remove("invalid");
  if (!target.checked) resetVerify();
});

$("send-code").addEventListener("click", () => {
  const phone = $<HTMLInputElement>("phone").value;
  if (!PHONE_RE.test(phone)) {
    setErr("f-phone", "phone-err", "인증번호를 받으려면 휴대폰 번호를 먼저 정확히 입력해 주세요.");
    $<HTMLInputElement>("phone").focus();
    return;
  }
  verify.code = String(Math.floor(100000 + Math.random() * 900000));
  verify.expires = Date.now() + CODE_TTL;
  verify.phone = phone;
  verify.done = false;
  $("demo-code").hidden = false;
  $("demo-code").textContent = `테스트 화면이라 문자는 발송되지 않아요. 인증번호 ${verify.code}를 3분 안에 입력해 주세요.`;
  $("send-code").textContent = "인증번호 다시 받기";
  $("verify-ok").hidden = true;
  $("f-verify").classList.remove("invalid");
  $<HTMLInputElement>("code").focus();
});

$("check-code").addEventListener("click", () => {
  if (!verify.code) return setErr("f-verify", "verify-err", "인증번호 받기를 먼저 눌러 주세요.");
  if (Date.now() > verify.expires) return setErr("f-verify", "verify-err", "인증번호 입력 시간이 지났어요. 인증번호를 다시 받아 주세요.");
  if ($<HTMLInputElement>("code").value.trim() !== verify.code) return setErr("f-verify", "verify-err", "인증번호가 맞지 않아요. 다시 확인해 주세요.");
  verify.done = true;
  setErr("f-verify", "verify-err", "");
  $("verify-ok").hidden = false;
  $("demo-code").hidden = true;
  $<HTMLInputElement>("code").disabled = true;
  $<HTMLButtonElement>("check-code").disabled = true;
  $<HTMLButtonElement>("send-code").disabled = true;
});

const updateCount = () => {
  $("msg-count").textContent = `${$<HTMLTextAreaElement>("msg").value.length} / 500`;
};
$<HTMLTextAreaElement>("msg").addEventListener("input", updateCount);

$("under14-toggle").addEventListener("click", () => {
  const note = $("under14-note");
  const open = note.hidden;
  note.hidden = !open;
  $("under14-toggle").setAttribute("aria-expanded", String(open));
});
$<HTMLInputElement>("age14").addEventListener("change", () => $("f-age").classList.remove("invalid"));
$<HTMLInputElement>("consent").addEventListener("change", () => $("f-consent").classList.remove("invalid"));

form.addEventListener("reset", () => setTimeout(() => {
  updateCount();
  $("name-cut").hidden = true;
  $("f-verify").hidden = true;
  $("under14-note").hidden = true;
  resetVerify();
  ["f-name","f-org","f-email","f-phone","f-verify","f-consent","f-age"].forEach((id) => $(id)?.classList.remove("invalid"));
  setMsg("");
}, 0));

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  if (closedInline) { setMsg("마감되었어요."); return; }
  setMsg("");
  cleanName();
  const r = {
    name: normName(nameEl.value),
    org: $<HTMLInputElement>("org").value.trim(),
    email: $<HTMLInputElement>("email").value.trim().toLowerCase(),
    phone: $<HTMLInputElement>("phone").value,
    message: $<HTMLTextAreaElement>("msg").value.trim(),
    sms: $<HTMLInputElement>("sms").checked
  };
  const bad: string[] = [];
  if (setErr("f-name", "name-err", checkName(r.name))) bad.push("f-name");
  if (setErr("f-org", "org-err", checkOrg(r.org))) bad.push("f-org");
  if (setErr("f-email", "email-err", checkEmail(r.email))) bad.push("f-email");
  if (setErr("f-phone", "phone-err", checkPhone(r.phone))) bad.push("f-phone");
  if (setErr("f-verify", "verify-err", r.sms && !(verify.done && verify.phone === r.phone) ? "문자로 안내를 받으려면 휴대폰 인증을 마쳐 주세요." : "")) bad.push("f-verify");
  if (setErr("f-consent", null, !$<HTMLInputElement>("consent").checked ? "x" : "")) bad.push("f-consent");
  if (setErr("f-age", null, !$<HTMLInputElement>("age14").checked ? "x" : "")) bad.push("f-age");
  if (bad.length) {
    const first = $(bad[0]).querySelector<HTMLInputElement | HTMLTextAreaElement>("input:not([disabled]), textarea");
    if (first) first.focus();
    return;
  }

  if (!canSubmit(Date.now())) { closeInline(); return; }

  const all = getAllApplications();
  const dup = findDup(all, r, null);
  if (dup === "ne") { setMsg("이미 신청된 고객입니다."); return; }
  if (dup === "phone") {
    setErr("f-phone", "phone-err", "이미 신청한 휴대폰 번호예요. 신청을 바꾸려면 기존 신청을 취소한 뒤 다시 신청해 주세요.");
    $<HTMLInputElement>("phone").focus();
    return;
  }

  const now = new Date().toISOString();
  const record: ApplicationRecord = {
    appId: newId(),
    name: r.name,
    org: r.org,
    email: r.email,
    phone: r.phone,
    message: r.message,
    ageOver14: true,
    smsOptIn: r.sms,
    phoneVerified: r.sms && verify.done,
    consent: true,
    consentAt: now,
    consentVersion: CONSENT_VERSION,
    appliedAt: now,
    updatedAt: null,
    source: "form",
    version: 1,
    questionStatus: r.message ? "open" : "none",
    adminMemo: "",
    history: [{ at: now, action: "create", by: "신청자" }]
  };

  submitBtn.disabled = true;
  submitBtn.textContent = "신청 처리 중...";

  try {
    // Save to Firebase Firestore
    await setDoc(doc(db, "applications", record.appId), record);

    // Direct Google Sheets append if access token is available in this session
    const token = getCachedAccessToken();
    if (token && currentSheetConfig?.spreadsheetId) {
      try {
        await appendApplicantToSheet(token, currentSheetConfig.spreadsheetId, currentSheetConfig.sheetName, record);
        knownAppIds.add(record.appId);
      } catch (sheetErr) {
        console.warn("Direct sheet append on submit notice:", sheetErr);
      }
    }

    // Also update local cache for resilience
    const local = loadLocalAll();
    local.push(record);
    saveLocalAll(local);

    showDone(record);
  } catch (err) {
    console.error("Firestore submit error:", err);
    try {
      handleFirestoreError(err, OperationType.CREATE, `applications/${record.appId}`);
    } catch {
      // Fallback: save to local if network error
      const local = loadLocalAll();
      local.push(record);
      saveLocalAll(local);
      showDone(record);
    }
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = "사전 신청하기";
  }
});

function showDone(r: ApplicationRecord) {
  shownId = r.appId;
  form.hidden = true;
  $("done").hidden = false;
  $("receipt").innerHTML =
    `<tr><th>신청 번호</th><td>${esc(r.appId.slice(0,8).toUpperCase())}</td></tr>` +
    `<tr><th>이름</th><td>${esc(r.name)}</td></tr>` +
    (r.org ? `<tr><th>소속</th><td>${esc(r.org)}</td></tr>` : "") +
    `<tr><th>이메일</th><td>${esc(r.email)}</td></tr>` +
    `<tr><th>휴대폰</th><td>${esc(mask(r.phone))}${r.smsOptIn ? " (문자 수신)" : ""}</td></tr>` +
    (r.message ? `<tr><th>질문·요청</th><td>${esc(r.message)}</td></tr>` : "") +
    `<tr><th>신청 일시</th><td>${esc(kst(r.appliedAt))}</td></tr>`;
}

$("again").addEventListener("click", () => {
  if (!canSubmit(Date.now())) { showClosed(); return; }
  shownId = null;
  form.reset();
  $("done").hidden = true;
  form.hidden = false;
  nameEl.focus();
});

function backToFormIfShown(ids: string[]) {
  if (!$("done").hidden && shownId && ids.includes(shownId)) {
    $("done").hidden = true;
    shownId = null;
    if (canSubmit(Date.now())) { form.reset(); form.hidden = false; } else showClosed();
  }
}

// =================== 관리자 (CRUD & 실시간) ===================
const STATUS: Record<string, string> = { none: "질문 없음", open: "미답변", answered: "답변 완료" };
const SOURCE: Record<string, string> = { form: "신청 폼", admin: "관리자 등록" };
const FIELD_LABEL: Record<string, string> = {
  name: "이름", org: "소속", email: "이메일", phone: "휴대폰", message: "질문·요청",
  smsOptIn: "문자 수신", questionStatus: "질문 처리 상태", adminMemo: "관리자 메모"
};
const PAGE_SIZE = 20;
let sort = { key: "appliedAt", dir: -1 };
let page = 1;
const selected = new Set<string>();
const shortId = (id: string) => String(id).slice(0, 8).toUpperCase();

function maskEmail(e: string) {
  const [u, d] = String(e).split("@");
  if (!d) return "***";
  return u.slice(0, 2) + "*".repeat(Math.max(1, u.length - 2)) + "@" + d;
}

function showVal(k: string, v: unknown) {
  if (k === "name") return maskName(String(v));
  if (k === "phone") return mask(String(v));
  if (k === "email") return maskEmail(String(v));
  if (k === "smsOptIn") return v ? "수신" : "미수신";
  if (k === "questionStatus") return STATUS[String(v)] || String(v);
  if (k === "message" || k === "adminMemo") return v ? "내용 변경" : "비움";
  return String(v || "-");
}

async function addLogToFirestore(action: string, r: ApplicationRecord | null, detail: string) {
  const logItem: LogItem = {
    at: new Date().toISOString(),
    action,
    id: r ? shortId(r.appId) : "-",
    who: r ? maskName(r.name) : (auth.currentUser?.email || "관리자"),
    detail: detail || ""
  };
  try {
    const logId = newId();
    await setDoc(doc(db, "logs", logId), logItem);
  } catch (err) {
    console.warn("Log write warning:", err);
  }
}

function query(all: ApplicationRecord[]) {
  const q = $<HTMLInputElement>("search").value.trim().toLowerCase();
  const qDigits = q.replace(/\D/g, "");
  const f = $<HTMLSelectElement>("filter").value;
  const list = all.filter((r) => {
    if (f === "sms" && !r.smsOptIn) return false;
    if (f === "open" && r.questionStatus !== "open") return false;
    if (f === "answered" && r.questionStatus !== "answered") return false;
    if (f === "admin" && r.source !== "admin") return false;
    if (!q) return true;
    return String(r.name).toLowerCase().includes(q) || String(r.org || "").toLowerCase().includes(q) ||
      String(r.email).toLowerCase().includes(q) || shortId(r.appId).toLowerCase().includes(q) ||
      (qDigits.length >= 3 && r.phone.includes(qDigits));
  });
  const rank: Record<string, number> = { open: 0, answered: 1, none: 2 };
  list.sort((a, b) => {
    let x: string | number, y: string | number;
    if (sort.key === "name") {
      x = String(a.name);
      y = String(b.name);
      return x.localeCompare(y, "ko") * sort.dir;
    }
    if (sort.key === "questionStatus") {
      x = rank[a.questionStatus] ?? 3;
      y = rank[b.questionStatus] ?? 3;
    } else {
      x = a.appliedAt;
      y = b.appliedAt;
    }
    return (x < y ? -1 : x > y ? 1 : 0) * sort.dir;
  });
  return list;
}

function renderAdmin() {
  if (!unlocked) return;
  const all = getAllApplications();
  $("st-total").textContent = String(all.length);
  $("st-sms").textContent = String(all.filter((r) => r.smsOptIn).length);
  $("st-open").textContent = String(all.filter((r) => r.questionStatus === "open").length);
  $("st-admin").textContent = String(all.filter((r) => r.source === "admin").length);
  $<HTMLButtonElement>("export-btn").disabled = all.length === 0;

  for (const id of [...selected]) {
    if (!all.some((r) => r.appId === id)) selected.delete(id);
  }

  const list = query(all);
  const pages = Math.max(1, Math.ceil(list.length / PAGE_SIZE));
  page = Math.min(Math.max(1, page), pages);
  const rows = list.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  const empty = $("roster-empty");
  if (!all.length) empty.textContent = "아직 신청이 없어요. 사전 신청이 들어오면 실시간으로 여기에 모여요.";
  else empty.textContent = "조건에 맞는 신청이 없어요.";
  empty.hidden = rows.length > 0;
  $("roster-wrap").hidden = rows.length === 0;
  $("pager").hidden = list.length <= PAGE_SIZE && list.length === 0;
  $("page-info").textContent = list.length
    ? `${(page - 1) * PAGE_SIZE + 1}–${Math.min(page * PAGE_SIZE, list.length)} / ${list.length}건` : "";
  $<HTMLButtonElement>("prev-page").disabled = page <= 1;
  $<HTMLButtonElement>("next-page").disabled = page >= pages;

  document.querySelectorAll<HTMLButtonElement>(".sortbtn").forEach((b) => {
    if (b.dataset.sort === sort.key) b.dataset.dir = sort.dir > 0 ? "asc" : "desc"; else delete b.dataset.dir;
  });

  $("roster-body").innerHTML = rows.map((r) => {
    const sms = !r.smsOptIn ? `<span class="pill off">미수신</span>`
      : r.phoneVerified ? `<span class="pill">수신</span>` : `<span class="pill warn">미인증</span>`;
    const qPill = r.questionStatus === "open" ? `<span class="pill warn">미답변</span>`
      : r.questionStatus === "answered" ? `<span class="pill">답변 완료</span>` : `<span class="pill off">없음</span>`;
    return `<tr>` +
      `<td class="ck"><input type="checkbox" data-sel="${esc(r.appId)}" ${selected.has(r.appId) ? "checked" : ""} aria-label="선택"></td>` +
      `<td>${esc(kstShort(r.appliedAt))}${r.updatedAt ? `<span class="edited">${esc(kstShort(r.updatedAt))} 수정</span>` : ""}</td>` +
      `<td>${esc(maskName(r.name))}</td><td>${esc(r.org || "-")}</td><td>${esc(r.email)}</td>` +
      `<td>${esc(mask(r.phone))}</td><td>${sms}</td>` +
      `<td class="qcell">${qPill}${r.message ? `<span class="qtext">${esc(r.message)}</span>` : ""}</td>` +
      `<td><span class="pill info">${esc(SOURCE[r.source] || r.source)}</span></td>` +
      `<td class="act"><button class="ghost sm" type="button" data-view="${esc(r.appId)}">상세</button>` +
      `<button class="ghost sm" type="button" data-edit="${esc(r.appId)}">수정</button>` +
      `<button class="ghost sm danger" type="button" data-del="${esc(r.appId)}">삭제</button></td></tr>`;
  }).join("");

  const pageIds = rows.map((r) => r.appId);
  const cp = $<HTMLInputElement>("check-page");
  cp.checked = pageIds.length > 0 && pageIds.every((id) => selected.has(id));
  cp.indeterminate = !cp.checked && pageIds.some((id) => selected.has(id));
  $<HTMLButtonElement>("bulk-del").disabled = selected.size === 0;
  $<HTMLButtonElement>("bulk-del").textContent = selected.size ? `선택 삭제 (${selected.size})` : "선택 삭제";
  renderLog();
}

function renderLog() {
  const l = firestoreLogs;
  $("log-empty").hidden = l.length > 0;
  $("log-body").innerHTML = l.slice(0, 100).map((x) =>
    `<tr><td>${esc(kstShort(x.at))}</td><td>${esc(x.action)}</td><td>${esc(x.id)}</td><td>${esc(x.who)}</td><td class="q">${esc(x.detail || "-")}</td></tr>`).join("");
}

$("search").addEventListener("input", () => { page = 1; renderAdmin(); });
$("filter").addEventListener("change", () => { page = 1; renderAdmin(); });
$("prev-page").addEventListener("click", () => { page--; renderAdmin(); });
$("next-page").addEventListener("click", () => { page++; renderAdmin(); });
document.querySelectorAll<HTMLButtonElement>(".sortbtn").forEach((b) => b.addEventListener("click", () => {
  const k = b.dataset.sort!;
  sort = sort.key === k ? {key: k, dir: -sort.dir} : {key: k, dir: k === "appliedAt" ? -1 : 1};
  renderAdmin();
}));

$("check-page").addEventListener("change", (e) => {
  const target = e.target as HTMLInputElement;
  document.querySelectorAll<HTMLInputElement>("#roster-body [data-sel]").forEach((c) => {
    if (target.checked) selected.add(c.dataset.sel!); else selected.delete(c.dataset.sel!);
  });
  renderAdmin();
});

$("roster-body").addEventListener("change", (e) => {
  const target = e.target as HTMLElement;
  const c = target.closest<HTMLInputElement>("[data-sel]");
  if (!c) return;
  if (c.checked) selected.add(c.dataset.sel!); else selected.delete(c.dataset.sel!);
  renderAdmin();
});

$("roster-body").addEventListener("click", (e) => {
  const target = e.target as HTMLElement;
  const v = target.closest<HTMLElement>("[data-view]");
  const ed = target.closest<HTMLElement>("[data-edit]");
  const d = target.closest<HTMLElement>("[data-del]");
  if (v) openDetail(v.dataset.view!);
  else if (ed) openEdit(ed.dataset.edit!);
  else if (d) askDelete([d.dataset.del!]);
});

$("bulk-del").addEventListener("click", () => {
  if (selected.size) askDelete([...selected]);
});

const findRec = (id: string) => getAllApplications().find((x) => x.appId === id);
function openModal(d: HTMLDialogElement) {
  if (typeof d.showModal === "function") d.showModal(); else d.setAttribute("open", "");
}
function closeModal(d: HTMLDialogElement) {
  if (d.open) d.close();
}

// 상세 보기
const detailDlg = $<HTMLDialogElement>("detail-dialog");
let detailId: string | null = null;
function openDetail(id: string) {
  const r = findRec(id);
  if (!r) { renderAdmin(); return; }
  detailId = id;
  const row = (k: string, v: string) => `<tr><th>${esc(k)}</th><td>${v}</td></tr>`;
  $("detail-body").innerHTML =
    row("신청 번호", esc(shortId(r.appId))) +
    row("등록 경로", esc(SOURCE[r.source] || r.source)) +
    row("신청 일시", esc(kst(r.appliedAt))) +
    row("이름", esc(maskName(r.name))) +
    row("소속", esc(r.org || "-")) +
    row("이메일", esc(r.email)) +
    row("휴대폰", esc(mask(r.phone))) +
    row("문자 수신", r.smsOptIn ? (r.phoneVerified ? "수신 (인증 완료)" : "수신 (미인증)") : "미수신") +
    row("질문·요청", r.message ? esc(r.message).replace(/\n/g, "<br>") : "-") +
    row("질문 처리 상태", esc(STATUS[r.questionStatus])) +
    row("관리자 메모", r.adminMemo ? esc(r.adminMemo).replace(/\n/g, "<br>") : "-") +
    row("만 14세 이상", r.ageOver14 ? "확인" : "-") +
    row("개인정보 동의", `${esc(kst(r.consentAt))} (${esc(r.consentVersion)})`) +
    row("최종 수정", r.updatedAt ? esc(kst(r.updatedAt)) : "-") +
    row("버전", esc(r.version));
  const h = (r.history || []).slice().reverse();
  $("detail-history").innerHTML = h.length ? h.map((x) =>
    `<li><b>${esc(kstShort(x.at))} ${esc(x.action === "create" ? "등록" : "수정")}</b> · ${esc(x.by || "")}` +
    (x.changes && x.changes.length ? `<br>${x.changes.map((c) => `${esc(FIELD_LABEL[c.field] || c.field)}: ${esc(c.from)} → ${esc(c.to)}`).join("<br>")}` : "") +
    `</li>`).join("") : "<li>기록 없음</li>";
  openModal(detailDlg);
}

$("detail-close").addEventListener("click", () => closeModal(detailDlg));
$("detail-edit").addEventListener("click", () => { closeModal(detailDlg); if (detailId) openEdit(detailId); });
$("detail-del").addEventListener("click", () => { closeModal(detailDlg); if (detailId) askDelete([detailId]); });

// 등록·수정
const dlg = $<HTMLDialogElement>("edit-dialog");
let editMode = "edit";
let editingId: string | null = null;
let editingVersion = 0;
attachNameClean($<HTMLInputElement>("e-name"));
attachPhoneClean($<HTMLInputElement>("e-phone"));
["name","org","email","phone"].forEach((k) => $<HTMLInputElement>("e-" + k).addEventListener("input", () => $("ef-" + k).classList.remove("invalid")));
$<HTMLInputElement>("e-consent").addEventListener("change", () => $("ef-consent").classList.remove("invalid"));
$<HTMLTextAreaElement>("e-msg").addEventListener("input", () => {
  const has = !!$<HTMLTextAreaElement>("e-msg").value.trim();
  const s = $<HTMLSelectElement>("e-status");
  if (!has) s.value = "none"; else if (s.value === "none") s.value = "open";
});

function editMsg(t: string) {
  $("edit-msg").textContent = t || "";
  $("edit-msg").classList.toggle("show", !!t);
}

function resetEditErrors() {
  ["name","org","email","phone","consent"].forEach((k) => $("ef-" + k).classList.remove("invalid"));
  editMsg("");
}

$("add-btn").addEventListener("click", () => {
  editMode = "create";
  editingId = null;
  resetEditErrors();
  $("edit-title").textContent = "신청자 추가";
  $("edit-meta").textContent = "전화·현장 접수처럼 신청자를 대신해 등록할 때 써요. Firestore 데이터베이스에 실시간으로 등록돼요.";
  ["e-name","e-org","e-email","e-phone","e-msg","e-memo"].forEach((id) => { ($<HTMLInputElement | HTMLTextAreaElement>(id)).value = ""; });
  $<HTMLInputElement>("e-sms").checked = false;
  $<HTMLSelectElement>("e-status").value = "none";
  $<HTMLInputElement>("e-consent").checked = false;
  $("ef-consent").hidden = false;
  $("e-phone-hint").hidden = true;
  $("edit-save").textContent = "등록";
  openModal(dlg);
  $<HTMLInputElement>("e-name").focus();
});

function openEdit(id: string) {
  const r = findRec(id);
  if (!r) { renderAdmin(); return; }
  editMode = "edit";
  editingId = id;
  editingVersion = r.version;
  resetEditErrors();
  $("edit-title").textContent = "신청 정보 수정";
  $("edit-meta").textContent = `신청 번호 ${shortId(r.appId)} · ${kst(r.appliedAt)} 신청 · 신청 번호, 신청 일시, 동의 기록은 바꿀 수 없어요.`;
  $<HTMLInputElement>("e-name").value = r.name;
  $<HTMLInputElement>("e-org").value = r.org || "";
  $<HTMLInputElement>("e-email").value = r.email;
  $<HTMLInputElement>("e-phone").value = r.phone;
  $<HTMLInputElement>("e-sms").checked = !!r.smsOptIn;
  $<HTMLTextAreaElement>("e-msg").value = r.message || "";
  $<HTMLSelectElement>("e-status").value = r.questionStatus;
  $<HTMLTextAreaElement>("e-memo").value = r.adminMemo || "";
  $("ef-consent").hidden = true;
  $("e-phone-hint").hidden = false;
  $("edit-save").textContent = "저장";
  openModal(dlg);
  $<HTMLInputElement>("e-name").focus();
}

$("edit-cancel").addEventListener("click", () => closeModal(dlg));

$<HTMLFormElement>("edit-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  editMsg("");
  const all = getAllApplications();
  const v = {
    name: normName($<HTMLInputElement>("e-name").value),
    org: $<HTMLInputElement>("e-org").value.trim(),
    email: $<HTMLInputElement>("e-email").value.trim().toLowerCase(),
    phone: $<HTMLInputElement>("e-phone").value,
    message: $<HTMLTextAreaElement>("e-msg").value.trim(),
    smsOptIn: $<HTMLInputElement>("e-sms").checked,
    questionStatus: $<HTMLSelectElement>("e-status").value as 'none' | 'open' | 'answered',
    adminMemo: $<HTMLTextAreaElement>("e-memo").value.trim()
  };
  if (!v.message) v.questionStatus = "none";
  else if (v.questionStatus === "none") v.questionStatus = "open";

  let bad = false;
  bad = setErr("ef-name", "ee-name", checkName(v.name)) || bad;
  bad = setErr("ef-org", "ee-org", checkOrg(v.org)) || bad;
  bad = setErr("ef-email", "ee-email", checkEmail(v.email)) || bad;
  bad = setErr("ef-phone", "ee-phone", checkPhone(v.phone)) || bad;
  if (editMode === "create") bad = setErr("ef-consent", null, $<HTMLInputElement>("e-consent").checked ? "" : "x") || bad;
  if (bad) return;

  const dup = findDup(all, v, editingId);
  if (dup === "ne") { editMsg("이미 신청된 고객입니다. 같은 이름과 이메일로 등록된 신청이 있어요."); return; }
  if (dup === "phone") { setErr("ef-phone", "ee-phone", "이 휴대폰 번호로 등록된 신청이 이미 있어요."); return; }
  const now = new Date().toISOString();

  const saveBtn = $<HTMLButtonElement>("edit-save");
  saveBtn.disabled = true;

  try {
    if (editMode === "create") {
      const rec: ApplicationRecord = normalize({
        appId: newId(),
        ...v,
        source: "admin",
        ageOver14: true,
        phoneVerified: false,
        consent: true,
        consentAt: now,
        consentVersion: CONSENT_VERSION + " (관리자 확인)",
        appliedAt: now,
        version: 1,
        history: [{ at: now, action: "create", by: auth.currentUser?.email || "관리자" }]
      });

      // Write to Firestore
      await setDoc(doc(db, "applications", rec.appId), rec);
      await addLogToFirestore("등록(관리자)", rec, "");

      // Direct Google Sheets append if access token is available
      const token = getCachedAccessToken();
      if (token && currentSheetConfig?.spreadsheetId) {
        try {
          await appendApplicantToSheet(token, currentSheetConfig.spreadsheetId, currentSheetConfig.sheetName, rec);
          knownAppIds.add(rec.appId);
          const newCount = (currentSheetConfig.syncedCount || 0) + 1;
          await setDoc(doc(db, "settings", "google_sheets"), { syncedCount: newCount, lastSyncedAt: new Date().toISOString() }, { merge: true });
        } catch (sheetErr) {
          console.warn("Direct sheet append on admin create notice:", sheetErr);
        }
      }

      closeModal(dlg);
      page = 1;
      renderAdmin();
      showToast("신청자를 등록하고 구글 시트에 행을 누적했어요.");
      return;
    }

    const idx = all.findIndex((x) => x.appId === editingId);
    if (idx < 0) { editMsg("그 사이 삭제된 신청이에요. 창을 닫아 주세요."); return; }
    const prev = all[idx];
    if (prev.version !== editingVersion) {
      editMsg("다른 창에서 이 신청이 먼저 수정됐어요. 창을 닫고 다시 열어 최신 내용을 확인해 주세요.");
      return;
    }

    const changes = Object.keys(FIELD_LABEL)
      .filter((k) => (prev[k as keyof ApplicationRecord] ?? "") !== (v[k as keyof typeof v] ?? ""))
      .map((k) => ({
        field: k,
        from: showVal(k, prev[k as keyof ApplicationRecord]),
        to: showVal(k, v[k as keyof typeof v])
      }));

    if (!changes.length) { closeModal(dlg); return; }
    const phoneChanged = v.phone !== prev.phone;
    const updatedRec: ApplicationRecord = {
      ...prev,
      ...v,
      phoneVerified: v.smsOptIn && !phoneChanged && !!prev.phoneVerified,
      updatedAt: now,
      version: prev.version + 1,
      history: (prev.history || []).concat({
        at: now,
        action: "update",
        by: auth.currentUser?.email || "관리자",
        changes
      })
    };

    // Update in Firestore
    await setDoc(doc(db, "applications", updatedRec.appId), updatedRec);
    await addLogToFirestore("수정", updatedRec, changes.map((c) => FIELD_LABEL[c.field]).join(", "));

    closeModal(dlg);
    renderAdmin();
    showToast("신청 정보를 실시간으로 수정했어요.");
  } catch (err) {
    console.error("Firestore edit error:", err);
    editMsg("Firestore 저장 중 오류가 발생했습니다. 권한을 확인해 주세요.");
  } finally {
    saveBtn.disabled = false;
  }
});

// 삭제 (Firestore 연동 + 실시간 되돌리기)
const delDlg = $<HTMLDialogElement>("del-dialog");
let deleteIds: string[] = [];

function askDelete(ids: string[]) {
  const all = getAllApplications();
  const recs = all.filter((r) => ids.includes(r.appId));
  if (!recs.length) { renderAdmin(); return; }
  deleteIds = recs.map((r) => r.appId);
  $("del-text").textContent = recs.length === 1
    ? `${maskName(recs[0].name)}(${mask(recs[0].phone)}) 님의 신청을 Firestore에서 삭제할까요?`
    : `선택한 신청 ${recs.length}건을 Firestore에서 삭제할까요?`;
  $<HTMLSelectElement>("del-reason").selectedIndex = 0;
  openModal(delDlg);
}

$("del-cancel").addEventListener("click", () => closeModal(delDlg));

$<HTMLFormElement>("del-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const reason = $<HTMLSelectElement>("del-reason").value;
  const all = getAllApplications();
  const removed = all.filter((r) => deleteIds.includes(r.appId));

  const delSubmitBtn = delDlg.querySelector<HTMLButtonElement>("button[type=submit]");
  if (delSubmitBtn) delSubmitBtn.disabled = true;

  try {
    for (const r of removed) {
      // Delete document from Firestore
      await deleteDoc(doc(db, "applications", r.appId));
      await addLogToFirestore("삭제", r, reason);
      selected.delete(r.appId);
    }

    closeModal(delDlg);
    renderAdmin();
    backToFormIfShown(removed.map((r) => r.appId));
    showToast(`${removed.length}건을 Firestore에서 삭제했어요.`, removed, reason);
  } catch (err) {
    console.error("Firestore delete error:", err);
    alert("삭제하지 못했어요. 관리자 권한을 확인해 주세요.");
  } finally {
    if (delSubmitBtn) delSubmitBtn.disabled = false;
  }
});

let toastTimer: number | null = null;
let undoBuf: { records: ApplicationRecord[]; reason: string } | null = null;

function showToast(text: string, undoRecords?: ApplicationRecord[], reason?: string) {
  if (toastTimer) clearTimeout(toastTimer);
  undoBuf = undoRecords ? { records: undoRecords, reason: reason || "" } : null;
  $("toast-text").textContent = text;
  $("toast-undo").hidden = !undoBuf;
  $("toast").hidden = false;
  toastTimer = window.setTimeout(() => {
    $("toast").hidden = true;
    undoBuf = null;
  }, undoBuf ? 8000 : 3000);
}

$("toast-undo").addEventListener("click", async () => {
  if (!undoBuf) return;
  const all = getAllApplications();
  let restored = 0;
  let skipped = 0;

  for (const r of undoBuf.records) {
    if (findDup(all, r, null) || all.some((x) => x.appId === r.appId)) {
      skipped++;
      continue;
    }
    try {
      await setDoc(doc(db, "applications", r.appId), r);
      await addLogToFirestore("삭제 되돌림", r, "");
      restored++;
    } catch (err) {
      console.error("Undo restore error:", err);
    }
  }

  undoBuf = null;
  if (toastTimer) clearTimeout(toastTimer);
  $("toast").hidden = true;
  renderAdmin();
  if (skipped) {
    showToast(`${restored}건을 되돌렸어요. ${skipped}건은 동일 정보가 있어 되돌리지 못했어요.`);
  } else {
    showToast(`${restored}건을 복원했어요.`);
  }
});

// CSV 전체 다운로드
const CSV_COLUMNS: [string, (r: ApplicationRecord) => string][] = [
  ["신청 번호", (r) => shortId(r.appId)],
  ["신청 일시", (r) => csvDate(r.appliedAt)],
  ["이름", (r) => r.name],
  ["소속", (r) => r.org || ""],
  ["이메일", (r) => r.email],
  ["휴대폰", (r) => fmtPhone(r.phone)],
  ["문자 수신 동의", (r) => (r.smsOptIn ? "Y" : "N")],
  ["휴대폰 인증", (r) => (r.smsOptIn ? (r.phoneVerified ? "Y" : "N") : "")],
  ["질문·요청", (r) => r.message || ""],
  ["질문 처리 상태", (r) => STATUS[r.questionStatus] || ""],
  ["등록 경로", (r) => SOURCE[r.source] || r.source],
  ["개인정보 동의 일시", (r) => csvDate(r.consentAt)]
];

function csvDate(iso: string) {
  if (!iso) return "";
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Seoul", year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", hourCycle: "h23"
  }).formatToParts(new Date(iso));
  const g = (t: string) => (parts.find((x) => x.type === t) || {}).value || "";
  return `${g("year")}-${g("month")}-${g("day")} ${g("hour")}:${g("minute")}`;
}

function csvCell(v: unknown) {
  let t = String(v ?? "").replace(/\r\n?/g, "\n");
  if (/^[=+\-@\t]/.test(t)) t = "'" + t;
  return '"' + t.replace(/"/g, '""') + '"';
}

function buildCsv(list: ApplicationRecord[]) {
  const lines = [CSV_COLUMNS.map((c) => csvCell(c[0])).join(",")];
  for (const r of list) lines.push(CSV_COLUMNS.map((c) => csvCell(c[1](r))).join(","));
  return "\uFEFF" + lines.join("\r\n") + "\r\n";
}

function stamp() {
  const d = csvDate(new Date().toISOString());
  return d.replace(/-/g, "").replace(" ", "_").replace(":", "");
}

async function saveFile(filename: string, text: string) {
  const url = URL.createObjectURL(new Blob([text], { type: "text/csv;charset=utf-8" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  return "saved";
}

const exportDlg = $<HTMLDialogElement>("export-dialog");
function exportMsg(t: string) {
  $("export-msg").textContent = t || "";
  $("export-msg").classList.toggle("show", !!t);
}

$("export-btn").addEventListener("click", () => {
  const n = getAllApplications().length;
  if (!n) { showToast("내려받을 신청이 없어요."); return; }
  exportMsg("");
  $("export-text").textContent = `검색·필터와 관계없이 Firestore에 등록된 전체 신청 ${n}건을 신청 일시 순으로 내려받아요.`;
  $<HTMLSelectElement>("export-purpose").selectedIndex = 0;
  openModal(exportDlg);
});

$("export-cancel").addEventListener("click", () => closeModal(exportDlg));

$<HTMLFormElement>("export-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!unlocked) return;
  const list = getAllApplications().sort((a, b) => (a.appliedAt < b.appliedAt ? -1 : 1));
  if (!list.length) { closeModal(exportDlg); return; }
  const purpose = $<HTMLSelectElement>("export-purpose").value;
  $<HTMLButtonElement>("export-go").disabled = true;
  exportMsg("");
  try {
    await saveFile(`셀렉트마켓_사전신청_${stamp()}.csv`, buildCsv(list));
    await addLogToFirestore("CSV 다운로드", null, `${list.length}건, ${purpose}`);
    closeModal(exportDlg);
    renderAdmin();
    showToast(`${list.length}건을 CSV로 내려받았어요.`);
  } catch {
    exportMsg("이 화면에서는 파일을 내려받을 수 없어요.");
  } finally {
    $<HTMLButtonElement>("export-go").disabled = false;
  }
});

// 관리자 인증 및 Firestore 실시간 동기화 리스너
let unlocked = false;
let fails = 0;
let lockUntil = 0;

async function sha256hex(t: string) {
  const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(t));
  return Array.from(new Uint8Array(b)).map((x) => x.toString(16).padStart(2, "0")).join("");
}

function startFirestoreSync() {
  if (unsubscribeApplications) unsubscribeApplications();
  if (unsubscribeLogs) unsubscribeLogs();
  if (unsubscribeSheetConfig) unsubscribeSheetConfig();

  const syncIndicator = $("admin-sync-status");
  if (syncIndicator) syncIndicator.textContent = "Firestore 실시간 연결 중...";

  try {
    // 1. Applications collection real-time listener (with real-time Google Sheets accumulation)
    unsubscribeApplications = onSnapshot(collection(db, "applications"), async (snapshot) => {
      const isFirst = isInitialAppsLoad;
      firestoreApplications = snapshot.docs.map((d) => normalize(d.data() as ApplicationRecord));
      if (syncIndicator) syncIndicator.textContent = `Firestore 실시간 연동 중 (${firestoreApplications.length}건)`;
      renderAdmin();

      if (isFirst) {
        firestoreApplications.forEach((a) => knownAppIds.add(a.appId));
        isInitialAppsLoad = false;
      } else {
        const newlyAdded = firestoreApplications.filter((a) => !knownAppIds.has(a.appId));
        for (const newApp of newlyAdded) {
          knownAppIds.add(newApp.appId);
          const token = getCachedAccessToken();
          if (token && currentSheetConfig?.spreadsheetId) {
            try {
              await appendApplicantToSheet(token, currentSheetConfig.spreadsheetId, currentSheetConfig.sheetName, newApp);
              const newCount = (currentSheetConfig.syncedCount || 0) + 1;
              currentSheetConfig.syncedCount = newCount;
              currentSheetConfig.lastSyncedAt = new Date().toISOString();
              await setDoc(doc(db, "settings", "google_sheets"), currentSheetConfig, { merge: true });
              await addLogToFirestore("구글 시트 자동 누적", newApp, "실시간 행 추가");
              updateSheetUI();
              showToast(`[구글시트] ${maskName(newApp.name)} 님의 신청이 시트에 추가되었습니다.`);
            } catch (sheetErr) {
              console.error("Auto sheet append error:", sheetErr);
            }
          }
        }
      }
    }, (error) => {
      console.warn("Firestore onSnapshot error:", error);
      if (syncIndicator) syncIndicator.textContent = "오프라인/로컬 캐시 모드";
      renderAdmin();
    });

    // 2. Logs collection real-time listener
    unsubscribeLogs = onSnapshot(collection(db, "logs"), (snapshot) => {
      firestoreLogs = snapshot.docs.map((d) => d.data() as LogItem).sort((a, b) => (a.at < b.at ? 1 : -1));
      renderLog();
    }, (err) => {
      console.warn("Logs snapshot error:", err);
    });

    // 3. Google Sheets Settings listener
    unsubscribeSheetConfig = onSnapshot(doc(db, "settings", "google_sheets"), (snapshot) => {
      if (snapshot.exists()) {
        currentSheetConfig = snapshot.data() as GoogleSheetConfig;
      } else {
        currentSheetConfig = null;
      }
      updateSheetUI();
    }, (err) => {
      console.warn("Sheet config snapshot error:", err);
    });
  } catch (err) {
    console.error("Failed to start Firestore snapshot:", err);
  }
}

function stopFirestoreSync() {
  if (unsubscribeApplications) {
    unsubscribeApplications();
    unsubscribeApplications = null;
  }
  if (unsubscribeLogs) {
    unsubscribeLogs();
    unsubscribeLogs = null;
  }
  if (unsubscribeSheetConfig) {
    unsubscribeSheetConfig();
    unsubscribeSheetConfig = null;
  }
  isInitialAppsLoad = true;
}

function lockAdmin() {
  unlocked = false;
  stopFirestoreSync();
  ["detail-dialog","edit-dialog","del-dialog","export-dialog","sheet-sync-dialog"].forEach((id) => {
    const d = $<HTMLDialogElement>(id);
    if (d?.open) d.close();
  });
  $("roster-body").innerHTML = "";
  $("log-body").innerHTML = "";
  ["st-total","st-sms","st-open","st-admin"].forEach((id) => { $(id).textContent = "0"; });
  selected.clear();
  $("toast").hidden = true;
  $("admin-content").hidden = true;
  $("admin-lock").hidden = false;
}

function applyLock() {
  if (!unlocked) {
    lockAdmin();
    $<HTMLInputElement>("pw").value = "";
    $("f-pw").classList.remove("invalid");
    setTimeout(() => $<HTMLInputElement>("pw")?.focus(), 0);
    return;
  }
  $("admin-lock").hidden = true;
  $("admin-content").hidden = false;
  startFirestoreSync();
  renderAdmin();
  updateSheetUI();
}

function showMessageDialog(title: string, text: string) {
  const d = $<HTMLDialogElement>("msg-dialog");
  if (!d) return;
  $("msg-dlg-title").textContent = title;
  $("msg-dlg-body").textContent = text;
  openModal(d);
}

$("msg-dlg-close")?.addEventListener("click", () => {
  closeModal($<HTMLDialogElement>("msg-dialog"));
});

function parseAuthErrorMessage(err: unknown): { code: string; message: string; guidance: string } {
  const e = err as { code?: string; message?: string };
  const code = e?.code || "";
  const rawMsg = e?.message || String(err || "");

  if (code === "auth/popup-blocked") {
    return {
      code,
      message: "브라우저에서 Google 로그인 팝업이 차단되었습니다.",
      guidance: "브라우저 주소창 우측 상단(또는 팝업 차단 설정)에서 '항상 허용'을 선택하시거나, 아래 '새 탭에서 직접 열기'를 눌러 전체 창에서 접속해 주세요."
    };
  }
  if (code === "auth/popup-closed-by-user") {
    return {
      code,
      message: "Google 로그인 팝업 창이 닫혔습니다.",
      guidance: "Google 계정을 선택하고 필요한 권한 승인을 완료해 주세요."
    };
  }
  if (code === "auth/cancelled-popup-request") {
    return {
      code,
      message: "이전 로그인 요청이 아직 처리 중입니다.",
      guidance: "잠시 후 다시 시도해 주세요."
    };
  }
  if (code === "auth/unauthorized-domain") {
    return {
      code,
      message: `접속 도메인(${window.location.hostname})이 Firebase 인증 승인 목록에 없습니다.`,
      guidance: "Firebase Console > Authentication > Settings > Authorized Domains에 현재 도메인을 추가해야 합니다."
    };
  }
  if (code === "auth/network-request-failed") {
    return {
      code,
      message: "네트워크 통신 오류가 발생했습니다.",
      guidance: "인터넷 연결 상태를 확인하고 다시 시도해 주세요."
    };
  }
  if (code === "auth/internal-error" || rawMsg.includes("storage") || rawMsg.includes("cookie") || rawMsg.includes("iframe")) {
    return {
      code: code || "auth/iframe-cookie-restriction",
      message: "브라우저의 iframe 서드파티 쿠키 차단 정책으로 인해 로그인이 제한되었을 수 있습니다.",
      guidance: "우측 상단 또는 아래의 '새 탭에서 직접 열기 ↗'를 눌러 전체 브라우저 창에서 열면 즉시 정상적으로 로그인 및 연동됩니다."
    };
  }
  return {
    code: code || "UNKNOWN",
    message: rawMsg,
    guidance: "브라우저 팝업 허용 여부를 확인하거나, '새 탭에서 직접 열기'로 접속해 보세요."
  };
}

async function ensureAccessToken(): Promise<string> {
  const cached = getCachedAccessToken();
  if (cached) return cached;
  const res = await signInWithPopup(auth, googleProvider);
  const credential = GoogleAuthProvider.credentialFromResult(res);
  if (!credential?.accessToken) {
    throw new Error("Google 로그인 결과에서 스프레드시트 액세스 토큰을 가져오지 못했습니다.");
  }
  setCachedAccessToken(credential.accessToken);
  return credential.accessToken;
}

// Google 로그인 버튼 이벤트
const googleLoginBtn = $("google-login-btn");
if (googleLoginBtn) {
  googleLoginBtn.addEventListener("click", async () => {
    const errEl = $("google-login-err");
    if (errEl) errEl.style.display = "none";
    try {
      googleLoginBtn.setAttribute("disabled", "true");
      googleLoginBtn.textContent = "Google 로그인 중...";
      const res = await signInWithPopup(auth, googleProvider);
      const user = res.user;
      const credential = GoogleAuthProvider.credentialFromResult(res);
      if (credential?.accessToken) {
        setCachedAccessToken(credential.accessToken);
      }
      if (user) {
        unlocked = true;
        const emailBadge = $("admin-email-badge");
        if (emailBadge) {
          emailBadge.textContent = `(${user.email})`;
        }
        applyLock();
        showToast(`${user.displayName || user.email} 님, 주최자 인증되었습니다.`);
      }
    } catch (err: unknown) {
      console.error("Google sign in error:", err);
      const parsed = parseAuthErrorMessage(err);
      if (errEl) {
        errEl.textContent = `❌ ${parsed.message}\n\n👉 해결 방법: ${parsed.guidance}`;
        errEl.style.display = "block";
      }
      showMessageDialog("Google 로그인 실패", `${parsed.message}\n\n[오류 코드]: ${parsed.code}\n\n[해결 방법]:\n${parsed.guidance}`);
    } finally {
      googleLoginBtn.removeAttribute("disabled");
      googleLoginBtn.innerHTML =
        '<svg width="20" height="20" viewBox="0 0 24 24"><path fill="#fff" d="M21.35 11.1H12v2.96h5.36c-.23 1.25-.94 2.31-2 3.02v2.51h3.24c1.9-1.75 2.99-4.32 2.99-7.39 0-.71-.06-1.4-.24-2z"/><path fill="#fff" d="M12 21c2.7 0 4.96-.9 6.62-2.41l-3.24-2.51c-.9.6-2.05.95-3.38.95-2.6 0-4.8-1.76-5.59-4.12H3.06v2.58C4.72 18.77 8.08 21 12 21z"/><path fill="#fff" d="M6.41 12.91c-.2-.6-.31-1.24-.31-1.91s.11-1.31.31-1.91V6.51H3.06C2.39 7.84 2 9.38 2 11s.39 3.16 1.06 4.49l3.35-2.58z"/><path fill="#fff" d="M12 5.38c1.47 0 2.79.51 3.83 1.5l2.87-2.87C16.96 2.41 14.7 1.5 12 1.5 8.08 1.5 4.72 3.73 3.06 7.02l3.35 2.58c.79-2.36 2.99-4.22 5.59-4.22z"/></svg> Google 계정으로 주최자 로그인';
    }
  });
}

// 비밀번호 폼 이벤트
$<HTMLFormElement>("lock-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const now = Date.now();
  if (now < lockUntil) {
    setErr("f-pw", "pw-err", `비밀번호를 여러 번 틀렸어요. ${Math.ceil((lockUntil - now) / 1000)}초 뒤에 다시 시도해 주세요.`);
    return;
  }
  const pw = $<HTMLInputElement>("pw").value;
  if (!pw) {
    setErr("f-pw", "pw-err", "비밀번호를 입력해 주세요.");
    return;
  }
  if (!(window.crypto && crypto.subtle)) {
    setErr("f-pw", "pw-err", "이 브라우저에서는 비밀번호를 확인할 수 없어요.");
    return;
  }
  $<HTMLButtonElement>("pw-btn").disabled = true;
  const ok = (await sha256hex(pw)) === ADMIN_PW_HASH;
  $<HTMLButtonElement>("pw-btn").disabled = false;
  $<HTMLInputElement>("pw").value = "";
  if (ok) {
    fails = 0;
    unlocked = true;
    setErr("f-pw", "pw-err", "");
    applyLock();
    $<HTMLInputElement>("search").focus();
    return;
  }
  fails++;
  if (fails >= 5) {
    fails = 0;
    lockUntil = Date.now() + 30000;
    setErr("f-pw", "pw-err", "비밀번호를 5번 틀렸어요. 30초 뒤에 다시 시도해 주세요.");
  } else {
    setErr("f-pw", "pw-err", `비밀번호가 맞지 않아요. (${fails}/5)`);
  }
  $<HTMLInputElement>("pw").focus();
});

$<HTMLInputElement>("pw").addEventListener("input", () => $("f-pw").classList.remove("invalid"));

$("lock-btn").addEventListener("click", async () => {
  if (auth.currentUser) {
    try {
      await signOut(auth);
    } catch {}
  }
  setCachedAccessToken(null);
  unlocked = false;
  applyLock();
});

// Google 스프레드시트 연동 버튼 이벤트
const sheetConnectBtn = $("sheet-connect-btn");
if (sheetConnectBtn) {
  sheetConnectBtn.addEventListener("click", async () => {
    const btn = $<HTMLButtonElement>("sheet-connect-btn");
    const errBox = $("sheet-err-box");
    if (errBox) errBox.style.display = "none";
    btn.disabled = true;
    btn.textContent = "Google 시트 생성/연동 중...";
    try {
      const token = await ensureAccessToken();
      const newConfig = await createApplicantSpreadsheet(token);
      const all = getAllApplications();
      if (all.length > 0) {
        await batchSyncAllApplicants(token, newConfig.spreadsheetId, newConfig.sheetName, all);
        newConfig.syncedCount = all.length;
        newConfig.lastSyncedAt = new Date().toISOString();
        all.forEach((a) => knownAppIds.add(a.appId));
      }
      await setDoc(doc(db, "settings", "google_sheets"), newConfig);
      currentSheetConfig = newConfig;
      updateSheetUI();
      await addLogToFirestore("구글 시트 연동", null, `새 스프레드시트 생성 및 ${all.length}건 동기화`);
      showToast("구글 스프레드시트가 성공적으로 생성 및 연동되었습니다!");
    } catch (err: unknown) {
      console.error("Sheets connect error:", err);
      const parsed = parseAuthErrorMessage(err);
      if (errBox) {
        errBox.textContent = `❌ 구글 시트 연동 실패: ${parsed.message}\n\n👉 해결 방법: ${parsed.guidance}`;
        errBox.style.display = "block";
      }
      showMessageDialog("Google 시트 연동 실패", `${parsed.message}\n\n[오류 코드]: ${parsed.code}\n\n[해결 방법]:\n${parsed.guidance}`);
    } finally {
      btn.disabled = false;
      updateSheetUI();
    }
  });
}

// Google 스프레드시트 일괄 동기화 모달 및 실행
const syncDialog = $<HTMLDialogElement>("sheet-sync-dialog");
const sheetSyncAllBtn = $("sheet-sync-all-btn");
if (sheetSyncAllBtn) {
  sheetSyncAllBtn.addEventListener("click", () => {
    if (!currentSheetConfig?.spreadsheetId) {
      showMessageDialog("시트 미연동", "연동된 구글 시트가 없습니다. 먼저 '구글 시트 연동 / 새 시트 생성'을 진행해 주세요.");
      return;
    }
    const all = getAllApplications();
    $("sheet-sync-desc").textContent = `현재 등록된 전체 신청자 ${all.length}건을 구글 시트에 일괄 동기화합니다.`;
    openModal(syncDialog);
  });
}

$("sheet-sync-cancel")?.addEventListener("click", () => closeModal(syncDialog));

$("sheet-sync-confirm")?.addEventListener("click", async () => {
  closeModal(syncDialog);
  const syncBtn = $<HTMLButtonElement>("sheet-sync-all-btn");
  const errBox = $("sheet-err-box");
  if (errBox) errBox.style.display = "none";
  if (syncBtn) {
    syncBtn.disabled = true;
    syncBtn.textContent = "동기화 중...";
  }
  try {
    const token = await ensureAccessToken();
    if (!currentSheetConfig?.spreadsheetId) {
      showMessageDialog("시트 미연동", "연동된 구글 시트가 없습니다.");
      return;
    }
    const all = getAllApplications();
    const res = await batchSyncAllApplicants(token, currentSheetConfig.spreadsheetId, currentSheetConfig.sheetName, all);
    all.forEach((a) => knownAppIds.add(a.appId));
    const now = new Date().toISOString();
    await setDoc(doc(db, "settings", "google_sheets"), {
      syncedCount: res.count,
      lastSyncedAt: now
    }, { merge: true });
    if (currentSheetConfig) {
      currentSheetConfig.syncedCount = res.count;
      currentSheetConfig.lastSyncedAt = now;
    }
    updateSheetUI();
    await addLogToFirestore("구글 시트 동기화", null, `${res.count}건 동기화 완료`);
    showToast(`${res.count}건의 신청자 명단이 구글 시트에 동기화되었습니다.`);
  } catch (err: unknown) {
    console.error("Sync error:", err);
    const parsed = parseAuthErrorMessage(err);
    if (errBox) {
      errBox.textContent = `❌ 시트 동기화 실패: ${parsed.message}\n\n👉 해결 방법: ${parsed.guidance}`;
      errBox.style.display = "block";
    }
    showMessageDialog("시트 동기화 실패", `${parsed.message}\n\n[오류 코드]: ${parsed.code}\n\n[해결 방법]:\n${parsed.guidance}`);
  } finally {
    if (syncBtn) {
      syncBtn.disabled = false;
      syncBtn.textContent = "전체 시트 동기화";
    }
  }
});


auth.onAuthStateChanged((user) => {
  if (user) {
    const emailBadge = $("admin-email-badge");
    if (emailBadge) {
      emailBadge.textContent = `(${user.email})`;
    }
    if (user.email === ADMIN_EMAIL || location.hash === "#admin") {
      unlocked = true;
      if (location.hash === "#admin") {
        applyLock();
      }
    }
  } else {
    setCachedAccessToken(null);
    const emailBadge = $("admin-email-badge");
    if (emailBadge) emailBadge.textContent = "";
  }
});

// 초기 실행
document.querySelectorAll<HTMLAnchorElement>(".open-tab-btn").forEach((a) => {
  a.href = window.location.href;
});
renderDday();
if (!canSubmit(openedAt)) showClosed();
setInterval(() => {
  renderDday();
  if (!form.hidden && !closedInline && !canSubmit(Date.now())) closeInline();
}, 1000);
route();

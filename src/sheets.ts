/**
 * Google Sheets API Integration for Select Market Pre-order Applications
 */

export interface SheetApplicantRecord {
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
}

export interface GoogleSheetConfig {
  spreadsheetId: string;
  spreadsheetUrl: string;
  sheetName: string;
  createdAt: string;
  lastSyncedAt?: string;
  syncedCount?: number;
}

const HEADERS = [
  "신청 번호",
  "신청 일시",
  "이름",
  "소속",
  "이메일",
  "휴대폰",
  "문자 수신",
  "휴대폰 인증",
  "질문·요청",
  "질문 상태",
  "등록 경로",
  "개인정보 동의 일시"
];

const STATUS_MAP: Record<string, string> = {
  none: "질문 없음",
  open: "미답변",
  answered: "답변 완료"
};

const SOURCE_MAP: Record<string, string> = {
  form: "신청 폼",
  admin: "관리자 등록"
};

function formatKstDate(iso: string): string {
  if (!iso) return "";
  try {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Seoul",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23"
    }).formatToParts(new Date(iso));
    const g = (t: string) => (parts.find((x) => x.type === t) || {}).value || "";
    return `${g("year")}-${g("month")}-${g("day")} ${g("hour")}:${g("minute")}:${g("second")}`;
  } catch {
    return iso;
  }
}

function formatPhone(phone: string): string {
  return phone.replace(/^(\d{3})(\d{4})(\d{4})$/, "$1-$2-$3");
}

function formatRow(r: SheetApplicantRecord): (string | number)[] {
  return [
    r.appId.slice(0, 8).toUpperCase(),
    formatKstDate(r.appliedAt),
    r.name,
    r.org || "",
    r.email,
    formatPhone(r.phone),
    r.smsOptIn ? "수신" : "미수신",
    r.smsOptIn ? (r.phoneVerified ? "인증 완료" : "미인증") : "-",
    r.message || "",
    STATUS_MAP[r.questionStatus] || r.questionStatus,
    SOURCE_MAP[r.source] || r.source,
    formatKstDate(r.consentAt)
  ];
}

/**
 * Creates a brand new Google Spreadsheet for Select Market with structured header and styling.
 */
export async function createApplicantSpreadsheet(
  accessToken: string,
  customTitle?: string
): Promise<GoogleSheetConfig> {
  const title = customTitle || `[셀렉트마켓] 연말 이벤트 사전 신청자 명단 (${new Date().toLocaleDateString("ko-KR")})`;
  const sheetName = "사전신청자";

  // 1. Create spreadsheet with initial sheet
  const createResp = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      properties: {
        title,
        locale: "ko_KR",
        timeZone: "Asia/Seoul"
      },
      sheets: [
        {
          properties: {
            title: sheetName,
            gridProperties: {
              frozenRowCount: 1,
              columnCount: HEADERS.length + 2
            }
          }
        }
      ]
    })
  });

  if (!createResp.ok) {
    const errText = await createResp.text();
    throw new Error(`Google Spreadsheet 생성 실패 (${createResp.status}): ${errText}`);
  }

  const sheetData = await createResp.json();
  const spreadsheetId: string = sheetData.spreadsheetId;
  const spreadsheetUrl: string = sheetData.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

  // 2. Set Header Row (Row 1)
  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!A1:L1?valueInputOption=USER_ENTERED`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        range: `${sheetName}!A1:L1`,
        majorDimension: "ROWS",
        values: [HEADERS]
      })
    }
  );

  return {
    spreadsheetId,
    spreadsheetUrl,
    sheetName,
    createdAt: new Date().toISOString(),
    syncedCount: 0
  };
}

/**
 * Appends a single applicant as a new row in Google Sheets in real-time.
 */
export async function appendApplicantToSheet(
  accessToken: string,
  spreadsheetId: string,
  sheetName: string,
  applicant: SheetApplicantRecord
): Promise<void> {
  const row = formatRow(applicant);

  const resp = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!A:L:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        range: `${sheetName}!A:L`,
        majorDimension: "ROWS",
        values: [row]
      })
    }
  );

  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error(`구글 시트 행 추가 실패 (${resp.status}): ${errText}`);
  }
}

/**
 * Checks if a spreadsheet is accessible with the provided accessToken.
 */
export async function verifySpreadsheetAccess(
  accessToken: string,
  spreadsheetId: string
): Promise<boolean> {
  try {
    const resp = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties.title`,
      {
        headers: { Authorization: `Bearer ${accessToken}` }
      }
    );
    return resp.ok;
  } catch {
    return false;
  }
}

/**
 * Batch synchronizes all current applicants to the Google Sheet.
 * Ensures header exists in row 1, then writes all data rows starting from row 2.
 */
export async function batchSyncAllApplicants(
  accessToken: string,
  spreadsheetId: string,
  sheetName: string,
  applicants: SheetApplicantRecord[]
): Promise<{ count: number }> {
  // Sort by appliedAt ascending
  const sorted = applicants.slice().sort((a, b) => (a.appliedAt < b.appliedAt ? -1 : 1));

  // 1. Ensure Header in row 1
  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!A1:L1?valueInputOption=USER_ENTERED`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        range: `${sheetName}!A1:L1`,
        majorDimension: "ROWS",
        values: [HEADERS]
      })
    }
  );

  // 2. Clear old data rows below header if any
  try {
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!A2:L5000:clear`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json"
        }
      }
    );
  } catch (clearErr) {
    console.warn("Clear old rows notice:", clearErr);
  }

  if (sorted.length === 0) {
    return { count: 0 };
  }

  // 3. Write all rows starting from A2
  const rows = sorted.map(formatRow);
  const endRow = sorted.length + 1;
  const updateRange = `${sheetName}!A2:L${endRow}`;

  const writeResp = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(updateRange)}?valueInputOption=USER_ENTERED`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        range: updateRange,
        majorDimension: "ROWS",
        values: rows
      })
    }
  );

  if (!writeResp.ok) {
    const errText = await writeResp.text();
    throw new Error(`구글 시트 일괄 동기화 실패 (${writeResp.status}): ${errText}`);
  }

  return { count: sorted.length };
}

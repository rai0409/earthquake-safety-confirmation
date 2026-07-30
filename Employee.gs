/**
 * Employee.gs
 * 社員情報の取得・検証
 */

/**
 * @typedef {Object} Employee
 * @property {string} employeeId
 * @property {string} name
 * @property {string} email
 * @property {boolean} active
 * @property {string} group
 */

// employeesシートの列インデックス（0始まり）
const EMP_COL = {
  EMPLOYEE_ID: 0,
  NAME: 1,
  EMAIL: 2,
  ACTIVE: 3,
  GROUP: 4,
};

/**
 * アクティブな社員一覧を取得
 * @returns {{ employees: Employee[], invalidCount: number }}
 */
function getActiveEmployees() {
  const sheet = getSheet(SHEET.EMPLOYEES);
  if (!sheet) throw new Error('employeesシートが見つかりません');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { employees: [], invalidCount: 0 };

  const employees = [];
  let invalidCount = 0;

  // ヘッダー行をスキップ（i=1から）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // 空行を無視
    const rowStr = row.map(c => String(c).trim()).join('');
    if (!rowStr) continue;

    const validation = validateEmployee(row);
    if (!validation.valid) {
      Logger.log(`社員検証エラー 行${i + 1}: ${validation.reason}`);
      invalidCount++;
      continue;
    }

    const active = parseBoolean(row[EMP_COL.ACTIVE]);
    if (!active) continue; // activeがFALSEの社員はスキップ

    employees.push({
      employeeId: String(row[EMP_COL.EMPLOYEE_ID]).trim(),
      name: String(row[EMP_COL.NAME]).trim(),
      email: normalizeEmail(String(row[EMP_COL.EMAIL])),
      active: true,
      group: String(row[EMP_COL.GROUP] || '').trim(),
    });
  }

  return { employees, invalidCount };
}

/**
 * 社員行データを検証
 * @param {Array} row - スプレッドシートの1行データ
 * @returns {{ valid: boolean, reason: string }}
 */
function validateEmployee(row) {
  if (!row || row.length < 3) {
    return { valid: false, reason: '列が不足しています' };
  }

  const employeeId = String(row[EMP_COL.EMPLOYEE_ID] || '').trim();
  const name = String(row[EMP_COL.NAME] || '').trim();
  const email = String(row[EMP_COL.EMAIL] || '').trim();

  if (!employeeId) {
    return { valid: false, reason: 'employee_idが空です' };
  }
  if (!name) {
    return { valid: false, reason: `employee_id "${employeeId}": nameが空です` };
  }
  if (!email) {
    return { valid: false, reason: `employee_id "${employeeId}": emailが空です` };
  }
  if (!isValidEmail(email)) {
    return { valid: false, reason: `employee_id "${employeeId}": メール形式が不正 "${email}"` };
  }

  return { valid: true, reason: '' };
}

/**
 * employee_idの重複チェック
 * @returns {{ duplicates: string[] }}
 */
function checkDuplicateEmployeeIds() {
  const sheet = getSheet(SHEET.EMPLOYEES);
  if (!sheet) return { duplicates: [] };

  const data = sheet.getDataRange().getValues();
  const ids = {};
  const duplicates = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowStr = row.map(c => String(c).trim()).join('');
    if (!rowStr) continue;

    const id = String(row[EMP_COL.EMPLOYEE_ID]).trim();
    if (!id) continue;

    if (ids[id]) {
      if (!duplicates.includes(id)) duplicates.push(id);
    } else {
      ids[id] = true;
    }
  }

  return { duplicates };
}

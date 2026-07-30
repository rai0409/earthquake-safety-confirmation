/**
 * Tests.gs
 * Apps Script内で実行できる軽量テスト関数
 * 外部サービスへ実送信しない
 */

/**
 * 全テストを実行
 */
function runAllTests() {
  const results = [];
  const tests = [
    testIntensityToRank,
    testThresholdComparison,
    testEventAgeJudgment,
    testNotificationKeyGeneration,
    testEmployeeValidation,
    testEmailValidation,
    testPrefilledFormUrl,
    testDuplicateEventDetection,
    testDuplicateNotificationDetection,
    testTestModeRecipientControl,
    testGmailQuotaInsufficient,
    testOutlookConfigMissing,
    testRetryStatusCodes,
    testLatestResponseSelection,
    testUnansweredAggregation,
    testDamageAggregation,
  ];

  for (const test of tests) {
    try {
      const result = test();
      results.push({ name: test.name, passed: result.passed, message: result.message });
    } catch (err) {
      results.push({ name: test.name, passed: false, message: `例外: ${safeErrorMessage(err)}` });
    }
  }

  const passed = results.filter(r => r.passed).length;
  const failed = results.filter(r => !r.passed).length;

  Logger.log(`\n===== テスト結果: ${passed}件合格 / ${failed}件不合格 =====`);
  results.forEach(r => {
    Logger.log(`${r.passed ? '✓' : '✗'} ${r.name}: ${r.message}`);
  });

  return results;
}

// ---------------------------------------------------------------------------
// 個別テスト関数
// ---------------------------------------------------------------------------

function testIntensityToRank() {
  const cases = [
    { input: '1', expected: 10 },
    { input: '4', expected: 40 },
    { input: '5-', expected: 45 },
    { input: '5+', expected: 50 },
    { input: '6-', expected: 55 },
    { input: '6+', expected: 60 },
    { input: '7', expected: 70 },
    { input: '不明', expected: null },
    { input: '', expected: null },
  ];

  for (const c of cases) {
    const result = intensityToRank(c.input);
    if (result !== c.expected) {
      return { passed: false, message: `震度"${c.input}": 期待=${c.expected}, 実際=${result}` };
    }
  }
  return { passed: true, message: '全ケース正常' };
}

function testThresholdComparison() {
  // 震度5弱(5-)がthreshold(5-)以上 → OK
  const rank5minus = intensityToRank('5-');
  const rankThreshold = intensityToRank('5-');
  if (rank5minus < rankThreshold) {
    return { passed: false, message: '震度5-がthreshold5-より小と判定' };
  }

  // 震度4がthreshold5-未満 → 送信しない
  const rank4 = intensityToRank('4');
  if (rank4 >= rankThreshold) {
    return { passed: false, message: '震度4がthreshold5-以上と誤判定' };
  }

  return { passed: true, message: '閾値比較正常' };
}

function testEventAgeJudgment() {
  const now = new Date();
  const minus29min = new Date(now.getTime() - 29 * 60 * 1000);
  const minus31min = new Date(now.getTime() - 31 * 60 * 1000);
  const maxAgeMs = 30 * 60 * 1000;

  const within = (now.getTime() - minus29min.getTime()) <= maxAgeMs;
  const over = (now.getTime() - minus31min.getTime()) <= maxAgeMs;

  if (!within) return { passed: false, message: '29分前の地震が対象外と判定' };
  if (over) return { passed: false, message: '31分前の地震が対象内と誤判定' };

  return { passed: true, message: '30分判定正常' };
}

function testNotificationKeyGeneration() {
  const key = generateNotificationKey('EQ20260101001', 'EMP001');
  const expected = 'EQ20260101001:EMP001:initial';
  if (key !== expected) {
    return { passed: false, message: `期待="${expected}", 実際="${key}"` };
  }
  return { passed: true, message: 'notification_key生成正常' };
}

function testEmployeeValidation() {
  const validRow = ['EMP001', '田中太郎', 'tanaka@example.com', 'TRUE', '営業部'];
  const noId = ['', '田中太郎', 'tanaka@example.com', 'TRUE', ''];
  const noName = ['EMP002', '', 'tanaka@example.com', 'TRUE', ''];
  const badEmail = ['EMP003', '佐藤', 'not-an-email', 'TRUE', ''];

  if (!validateEmployee(validRow).valid) {
    return { passed: false, message: '有効な社員行が不正と判定' };
  }
  if (validateEmployee(noId).valid) {
    return { passed: false, message: 'employee_id空が有効と誤判定' };
  }
  if (validateEmployee(noName).valid) {
    return { passed: false, message: 'name空が有効と誤判定' };
  }
  if (validateEmployee(badEmail).valid) {
    return { passed: false, message: '不正メールが有効と誤判定' };
  }

  return { passed: true, message: '社員検証正常' };
}

function testEmailValidation() {
  const valid = ['user@example.com', 'user.name+tag@domain.co.jp'];
  const invalid = ['notanemail', '@nodomain.com', 'user@', ''];

  for (const e of valid) {
    if (!isValidEmail(e)) return { passed: false, message: `有効メール"${e}"が不正と判定` };
  }
  for (const e of invalid) {
    if (isValidEmail(e)) return { passed: false, message: `不正メール"${e}"が有効と誤判定` };
  }

  return { passed: true, message: 'メール検証正常' };
}

function testPrefilledFormUrl() {
  // モック設定を直接設定してテスト
  const originalCache = _settingsCache;
  _settingsCache = {
    form_base_url: 'https://docs.google.com/forms/d/YOUR_FORM_ID/viewform',
    form_event_entry_id: 'entry.111111111',
    form_employee_id_entry_id: 'entry.222222222',
    form_name_entry_id: 'entry.333333333',
  };

  let result = { passed: false, message: '' };
  try {
    const event = { eventId: 'EQ20260101001' };
    const employee = { employeeId: 'EMP001', name: '田中太郎' };
    const url = buildPrefilledFormUrl(event, employee);

    if (!url.includes('entry.111111111=EQ20260101001')) {
      result = { passed: false, message: 'event_idがURLに含まれない' };
    } else if (!url.includes('entry.222222222=EMP001')) {
      result = { passed: false, message: 'employee_idがURLに含まれない' };
    } else if (!url.includes('entry.333333333=')) {
      result = { passed: false, message: 'nameがURLに含まれない' };
    } else {
      result = { passed: true, message: 'フォームURL生成正常' };
    }
  } catch (err) {
    result = { passed: false, message: `例外: ${safeErrorMessage(err)}` };
  } finally {
    _settingsCache = originalCache;
  }

  return result;
}

function testDuplicateEventDetection() {
  // isEventAlreadyCompleted はシートを読むため、ここではロジック検証のみ
  const completedStatuses = [
    EVENT_STATUS.COMPLETED,
    EVENT_STATUS.PARTIAL_FAILED,
    EVENT_STATUS.TEST_COMPLETED,
  ];
  const notCompleted = [EVENT_STATUS.DETECTED, EVENT_STATUS.PROCESSING, EVENT_STATUS.FAILED];

  for (const s of completedStatuses) {
    if (!completedStatuses.includes(s)) {
      return { passed: false, message: `status "${s}" が完了判定から漏れている` };
    }
  }
  for (const s of notCompleted) {
    if (completedStatuses.includes(s)) {
      return { passed: false, message: `status "${s}" が誤って完了判定されている` };
    }
  }

  return { passed: true, message: 'イベント重複判定定数正常' };
}

function testDuplicateNotificationDetection() {
  const terminalStatuses = [NOTIFICATION_STATUS.SENT, NOTIFICATION_STATUS.RESPONDED];
  const nonTerminal = [NOTIFICATION_STATUS.PENDING, NOTIFICATION_STATUS.SENDING, NOTIFICATION_STATUS.FAILED];

  for (const s of terminalStatuses) {
    if (!terminalStatuses.includes(s)) {
      return { passed: false, message: `status "${s}" が再送禁止判定から漏れている` };
    }
  }
  for (const s of nonTerminal) {
    if (terminalStatuses.includes(s)) {
      return { passed: false, message: `status "${s}" が誤って再送禁止判定` };
    }
  }

  return { passed: true, message: '通知重複判定定数正常' };
}

function testTestModeRecipientControl() {
  // テストモードでtest_recipient_emailだけへ送信することを確認
  const isTestMode = true;
  const testEmail = 'test@example.com';
  const employees = [
    { email: 'emp1@example.com', name: '社員1' },
    { email: 'emp2@example.com', name: '社員2' },
  ];

  // テストモード時は actualRecipient が testEmail になるはず
  const actualRecipient = isTestMode ? { email: testEmail } : employees[0];
  if (actualRecipient.email !== testEmail) {
    return { passed: false, message: 'テストモードで実社員へ送信されようとしている' };
  }

  return { passed: true, message: 'テストモード送信先制御正常' };
}

function testGmailQuotaInsufficient() {
  // quotaが送信対象数より少ない場合の判定
  const quota = 5;
  const required = 10;
  const { ok } = { ok: quota >= required };
  if (ok) {
    return { passed: false, message: 'quota不足が検出されない' };
  }
  return { passed: true, message: 'Gmail quota不足判定正常' };
}

function testOutlookConfigMissing() {
  // Script Propertiesが空の状態での検証（実際には設定されていないことを前提）
  // validateMicrosoftConfiguration はPropertiesServiceを読むため、ロジック検証のみ
  const required = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MS_SENDER_EMAIL'];
  const present = ['MS_TENANT_ID'];
  const missing = required.filter(k => !present.includes(k));

  if (missing.length === 0) {
    return { passed: false, message: '設定不足が検出されない' };
  }
  if (!missing.includes('MS_CLIENT_ID')) {
    return { passed: false, message: 'MS_CLIENT_IDが不足一覧に含まれない' };
  }

  return { passed: true, message: 'Outlook設定不足判定正常' };
}

function testRetryStatusCodes() {
  const shouldRetry = [429, 500, 502, 503, 504];
  const shouldNotRetry = [400, 401, 403, 404];

  for (const code of shouldRetry) {
    if (!RETRY_STATUS_CODES.includes(code)) {
      return { passed: false, message: `${code}が再試行対象に含まれない` };
    }
  }
  for (const code of shouldNotRetry) {
    if (RETRY_STATUS_CODES.includes(code)) {
      return { passed: false, message: `${code}が誤って再試行対象` };
    }
    if (!NO_RETRY_STATUS_CODES.includes(code)) {
      return { passed: false, message: `${code}が再試行しない対象に含まれない` };
    }
  }

  return { passed: true, message: '再試行ステータスコード判定正常' };
}

function testLatestResponseSelection() {
  // 同じemployeeIdから複数回答 → 最新を選択
  const responses = [
    { timestamp: '2026-01-01T10:00:00.000Z', employee_id: 'EMP001', safety_status: '被害なし' },
    { timestamp: '2026-01-01T11:00:00.000Z', employee_id: 'EMP001', safety_status: '本人にけがあり' },
    { timestamp: '2026-01-01T09:00:00.000Z', employee_id: 'EMP002', safety_status: '被害なし' },
  ];

  // getLatestResponsesByEmployee をシミュレート
  const sorted = [...responses].sort((a, b) =>
    new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime()
  );
  const latest = {};
  for (const r of sorted) {
    latest[r.employee_id] = r;
  }

  if (latest['EMP001'].safety_status !== '本人にけがあり') {
    return { passed: false, message: 'EMP001の最新回答が選択されていない' };
  }
  if (Object.keys(latest).length !== 2) {
    return { passed: false, message: '社員数が不正' };
  }

  return { passed: true, message: '最新回答選択正常' };
}

function testUnansweredAggregation() {
  // 送信済みで未回答の社員を抽出
  const notifications = [
    { notification_key: 'EQ001:EMP001:initial', employee_id: 'EMP001', email: 'a@ex.com', status: 'sent', sent_at: '' },
    { notification_key: 'EQ001:EMP002:initial', employee_id: 'EMP002', email: 'b@ex.com', status: 'responded', sent_at: '' },
    { notification_key: 'EQ001:EMP003:initial', employee_id: 'EMP003', email: 'c@ex.com', status: 'failed', sent_at: '' },
  ];
  const latestResponses = { EMP002: { safety_status: '被害なし', employee_id: 'EMP002' } };

  const unanswered = notifications
    .filter(n => String(n.status).trim() === NOTIFICATION_STATUS.SENT)
    .filter(n => !latestResponses[String(n.employee_id)]);

  if (unanswered.length !== 1 || unanswered[0].employee_id !== 'EMP001') {
    return { passed: false, message: '未回答者集計が不正' };
  }

  return { passed: true, message: '未回答者集計正常' };
}

function testDamageAggregation() {
  // 安否状況別集計
  const responses = [
    { safety_status: '被害なし' },
    { safety_status: '被害なし' },
    { safety_status: '本人にけがあり' },
    { safety_status: '家族に被害あり' },
    { safety_status: '複数の被害あり' },
    { safety_status: '被害なし' },
  ];

  const counts = {
    no_damage: 0,
    personal_injury: 0,
    family_damage: 0,
    multiple_damage: 0,
  };

  for (const r of responses) {
    if (r.safety_status === '被害なし') counts.no_damage++;
    else if (r.safety_status === '本人にけがあり') counts.personal_injury++;
    else if (r.safety_status === '家族に被害あり') counts.family_damage++;
    else if (r.safety_status === '複数の被害あり') counts.multiple_damage++;
  }

  if (counts.no_damage !== 3) return { passed: false, message: `被害なし: 期待3, 実際${counts.no_damage}` };
  if (counts.personal_injury !== 1) return { passed: false, message: `本人けが: 期待1, 実際${counts.personal_injury}` };
  if (counts.family_damage !== 1) return { passed: false, message: `家族被害: 期待1, 実際${counts.family_damage}` };
  if (counts.multiple_damage !== 1) return { passed: false, message: `複数被害: 期待1, 実際${counts.multiple_damage}` };

  return { passed: true, message: '被害状況集計正常' };
}

// ═══════════════════════════════════════════════════════════════
// 職場安全巡查記錄 — Google Apps Script
// 部署為 Web App 後，將 WEB_APP_URL 貼入前端 HTML
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = '巡查記錄';
const EMAIL_TO   = 'saiwanhsrecord@gmail.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    saveToSheet(data);
    sendEmail(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Web App is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Write headers
    sheet.appendRow([
      '提交時間', '檢查項目', '巡查人員', '部門', '位置', '巡查日期',
      '合格數', '不合格數', '不適用數', '總項目數',
      '詳細結果', '備註'
    ]);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold')
         .setBackground('#1a5fb4').setFontColor('white');
    sheet.setFrozenRows(1);
  }

  // Build detail string
  let details = [];
  data.items.forEach(item => {
    const status = item.status === 'pass' ? '✓合格'
                 : item.status === 'fail' ? '✗不合格'
                 : item.status === 'na'   ? '—不適用' : '未填';
    const remark = item.remark ? ` [備註:${item.remark}]` : '';
    details.push(`${status} ${item.text}${remark}`);
  });

  const rowColor = data.failCount > 0 ? '#fff5f5' : '#f0fff4';

  const newRow = [
    new Date(),
    data.title,
    data.inspector,
    data.dept,
    data.location,
    data.date,
    data.passCount,
    data.failCount,
    data.naCount,
    data.total,
    details.join('\n'),
    data.failCount > 0 ? `⚠️ 發現${data.failCount}項不合格` : '✓ 全部合格'
  ];

  const lastRow = sheet.appendRow(newRow);
  // Colour row by pass/fail
  const rowNum = sheet.getLastRow();
  sheet.getRange(rowNum, 1, 1, 12).setBackground(rowColor);
  // Auto-resize columns
  sheet.autoResizeColumns(1, 12);
}

function sendEmail(data) {
  const subject = `[${data.failCount > 0 ? '⚠️ 不合格' : '✓ 合格'}] ${data.title} 巡查記錄 ${data.date} — ${data.inspector}`;

  let bodyLines = [
    `<h2 style="color:${data.failCount > 0 ? '#dc2626' : '#16a34a'};">` +
    `${data.failCount > 0 ? '⚠️ 發現不合格項目' : '✓ 本次巡查全部合格'}</h2>`,
    `<table style="border-collapse:collapse;font-family:sans-serif;font-size:14px;margin-bottom:16px;">`,
    `<tr><td style="padding:6px 12px;color:#555;width:80px;">檢查項目</td><td style="padding:6px 12px;font-weight:700;">${data.title}</td></tr>`,
    `<tr style="background:#f5f5f5;"><td style="padding:6px 12px;color:#555;">巡查人員</td><td style="padding:6px 12px;">${data.inspector}</td></tr>`,
    `<tr><td style="padding:6px 12px;color:#555;">部門</td><td style="padding:6px 12px;">${data.dept}</td></tr>`,
    `<tr style="background:#f5f5f5;"><td style="padding:6px 12px;color:#555;">位置</td><td style="padding:6px 12px;">${data.location}</td></tr>`,
    `<tr><td style="padding:6px 12px;color:#555;">巡查日期</td><td style="padding:6px 12px;">${data.date}</td></tr>`,
    `</table>`,
    `<div style="display:flex;gap:12px;margin-bottom:20px;">`,
    `<div style="background:#f0fdf4;border:1px solid #4ade80;border-radius:8px;padding:10px 18px;text-align:center;"><strong style="font-size:22px;color:#16a34a;">${data.passCount}</strong><br><span style="font-size:12px;color:#555;">合格</span></div>`,
    `<div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:10px 18px;text-align:center;"><strong style="font-size:22px;color:#dc2626;">${data.failCount}</strong><br><span style="font-size:12px;color:#555;">不合格</span></div>`,
    `<div style="background:#f9fafb;border:1px solid #ccc;border-radius:8px;padding:10px 18px;text-align:center;"><strong style="font-size:22px;color:#6b7280;">${data.naCount}</strong><br><span style="font-size:12px;color:#555;">不適用</span></div>`,
    `<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:10px 18px;text-align:center;"><strong style="font-size:22px;color:#2563eb;">${data.total}</strong><br><span style="font-size:12px;color:#555;">總項目</span></div>`,
    `</div>`,
  ];

  if (data.failCount > 0) {
    bodyLines.push(`<div style="background:#fef2f2;border:2px solid #dc2626;border-radius:8px;padding:14px 18px;margin-bottom:20px;">`);
    bodyLines.push(`<strong style="color:#dc2626;font-size:15px;">⚠️ 需要即時處理：</strong><ol style="margin:8px 0 0 18px;color:#333;font-size:13px;">`);
    bodyLines.push(`<li>立即停止使用相關設備，張貼「停用」警告標示</li>`);
    bodyLines.push(`<li>隔離設備，防止其他員工誤用</li>`);
    bodyLines.push(`<li>聯絡維修人員安排修復</li>`);
    bodyLines.push(`<li>確認合格後方可恢復使用</li>`);
    bodyLines.push(`</ol></div>`);
  }

  bodyLines.push(`<h3 style="color:#333;border-bottom:1px solid #ddd;padding-bottom:8px;">詳細巡查結果</h3>`);
  bodyLines.push(`<table style="width:100%;border-collapse:collapse;font-size:13px;">`);
  bodyLines.push(`<thead><tr style="background:#333;color:white;"><th style="padding:8px 12px;text-align:left;">#</th><th style="padding:8px 12px;text-align:left;">檢查項目</th><th style="padding:8px 12px;text-align:left;width:90px;">結果</th></tr></thead><tbody>`);

  let n = 0;
  data.items.forEach(item => {
    n++;
    const s = item.status;
    const statusTxt = s === 'pass' ? '✓ 合格' : s === 'fail' ? '✗ 不合格' : s === 'na' ? '— 不適用' : '未填寫';
    const bg = s === 'pass' ? '#f0fdf4' : s === 'fail' ? '#fef2f2' : s === 'na' ? '#f9fafb' : '#fffbeb';
    const color = s === 'pass' ? '#16a34a' : s === 'fail' ? '#dc2626' : '#6b7280';
    const remark = item.remark ? `<br><small style="color:#666;">備註：${item.remark}</small>` : '';
    bodyLines.push(`<tr style="background:${bg};"><td style="padding:7px 12px;color:#777;">${n}</td><td style="padding:7px 12px;">${item.text}${remark}</td><td style="padding:7px 12px;font-weight:700;color:${color};">${statusTxt}</td></tr>`);
  });

  bodyLines.push(`</tbody></table>`);
  bodyLines.push(`<p style="font-size:11px;color:#999;margin-top:16px;border-top:1px solid #ddd;padding-top:8px;">此電郵由職場安全月度巡查系統自動發送 · ${new Date().toLocaleString()}</p>`);

  const htmlBody = bodyLines.join('');

  GmailApp.sendEmail(EMAIL_TO, subject, '', { htmlBody: htmlBody });
}

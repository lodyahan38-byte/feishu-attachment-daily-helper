const FEISHU_API = 'https://open.feishu.cn/open-apis';

const CONFIG = {
  sourceWikiToken: process.env.SOURCE_WIKI_TOKEN,
  sourceSheetId: process.env.SOURCE_SHEET_ID,
  targetWikiToken: process.env.TARGET_WIKI_TOKEN,
  targetSheetId: process.env.TARGET_SHEET_ID,
  snapshotSheetId: process.env.SNAPSHOT_SHEET_ID,
  reportGroupName: process.env.REPORT_GROUP_NAME,
  idColumn: process.env.ID_COLUMN || 'I',
  attachmentColumn: process.env.ATTACHMENT_COLUMN || 'AJ',
  demandTimeColumn: process.env.DEMAND_TIME_COLUMN || 'A',
  locationColumn: process.env.LOCATION_COLUMN || 'C',
  brandColumn: process.env.BRAND_COLUMN || 'K',
  seriesColumn: process.env.SERIES_COLUMN || 'L',
  modelColumn: process.env.MODEL_COLUMN || 'M'
};

function required(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing environment variable: ${name}`);
  return value;
}

function colToIndex(col) {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}

function cellText(value) {
  if (value == null) return '';
  if (typeof value === 'string' || typeof value === 'number') return String(value).trim();
  if (Array.isArray(value)) return value.map(cellText).filter(Boolean).join(' ');
  if (typeof value === 'object') {
    return String(value.text || value.name || value.file_name || value.filename || value.value || '').trim();
  }
  return String(value).trim();
}

function normalizeAttachment(value) {
  if (!value) return { name: '', token: '', raw: value };
  const list = Array.isArray(value) ? value : [value];
  const items = list
    .map((item) => {
      if (item == null) return null;
      if (typeof item === 'string') return { name: item.trim(), token: item.trim() };
      if (typeof item === 'object') {
        const name = item.text || item.name || item.file_name || item.filename || item.value || '';
        const token = item.fileToken || item.file_token || item.token || item.id || name;
        return { name: String(name).trim(), token: String(token || '').trim() };
      }
      return { name: String(item).trim(), token: String(item).trim() };
    })
    .filter((item) => item && (item.name || item.token));

  return {
    name: items.map((item) => item.name).filter(Boolean).join(' | '),
    token: items.map((item) => item.token || item.name).filter(Boolean).join('|'),
    raw: value
  };
}

function compareAttachment(oldItem, currentItem) {
  const oldToken = oldItem?.token || '';
  const newToken = currentItem?.token || '';
  const oldName = oldItem?.name || '';
  const newName = currentItem?.name || '';

  if (!oldToken && newToken) return '新增成片';
  if (oldToken && !newToken) return '异常清空';
  if (oldToken && newToken && oldToken !== newToken) return '修改成片';
  if (oldName && newName && oldName !== newName) return '修改成片';
  return '无变化';
}

function stripZip(name) {
  return String(name || '').replace(/\.zip$/i, '').trim();
}

function normalizeForCompare(text) {
  return String(text || '')
    .replace(/\.zip$/i, '')
    .replace(/\s+/g, '')
    .replace(/[＿_—–]/g, '-')
    .trim()
    .toLowerCase();
}

function validateNaming({ attachmentName, carId, brand, series, model, location, demandTime }) {
  const baseName = stripZip(attachmentName);
  const expected = `成片-${carId}-${brand} ${series} ${model}-${location}-${demandTime}.zip`;
  const problems = [];

  if (!attachmentName) problems.push('附件名为空');
  if (!/^成片-/.test(baseName)) problems.push('缺少固定前缀：成片-');
  if (!/\.zip$/i.test(attachmentName)) problems.push('文件后缀不是 .zip');

  const normalizedName = normalizeForCompare(attachmentName);
  const checks = [
    ['车源商品ID', carId],
    ['品牌名称', brand],
    ['车系名称', series],
    ['车型名称', model],
    ['所属地', location],
    ['需求时间', demandTime]
  ];

  for (const [label, value] of checks) {
    const normalizedValue = normalizeForCompare(value);
    if (normalizedValue && !normalizedName.includes(normalizedValue)) problems.push(`${label}不匹配或缺失：${value}`);
  }

  return {
    expected,
    result: problems.length ? '不通过' : '通过',
    problems
  };
}

async function feishuFetch(path, options = {}) {
  const token = await getTenantAccessToken();
  const res = await fetch(`${FEISHU_API}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json; charset=utf-8',
      ...(options.headers || {})
    }
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`Feishu API failed: ${path} ${JSON.stringify(data)}`);
  }
  return data;
}

let cachedToken = null;
let cachedTokenExpireAt = 0;
async function getTenantAccessToken() {
  if (cachedToken && Date.now() < cachedTokenExpireAt) return cachedToken;
  const res = await fetch(`${FEISHU_API}/auth/v3/tenant_access_token/internal`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify({
      app_id: required('FEISHU_APP_ID'),
      app_secret: required('FEISHU_APP_SECRET')
    })
  });
  const data = await res.json();
  if (data.code !== 0) throw new Error(`Failed to get tenant_access_token: ${JSON.stringify(data)}`);
  cachedToken = data.tenant_access_token;
  cachedTokenExpireAt = Date.now() + Math.max((data.expire || 7200) - 300, 60) * 1000;
  return cachedToken;
}

async function resolveSpreadsheetToken(wikiToken) {
  if (!wikiToken) return '';
  try {
    const data = await feishuFetch(`/wiki/v2/spaces/get_node?token=${encodeURIComponent(wikiToken)}`);
    return data.data?.node?.obj_token || data.data?.node?.origin_node_token || wikiToken;
  } catch (err) {
    console.warn('resolve wiki token failed, fallback to original token:', err.message);
    return wikiToken;
  }
}

async function readRange(spreadsheetToken, range) {
  const encoded = encodeURIComponent(range);
  const data = await feishuFetch(`/sheets/v2/spreadsheets/${spreadsheetToken}/values/${encoded}`, { method: 'GET' });
  return data.data?.valueRange?.values || [];
}

async function writeRange(spreadsheetToken, range, values) {
  return feishuFetch(`/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    body: JSON.stringify({
      valueRange: { range, values }
    })
  });
}

async function readCurrentAttachmentState(spreadsheetToken) {
  const values = await readRange(spreadsheetToken, `${CONFIG.sourceSheetId}!A2:${CONFIG.attachmentColumn}1000`);
  const idIdx = colToIndex(CONFIG.idColumn);
  const attachmentIdx = colToIndex(CONFIG.attachmentColumn);

  return values
    .map((row, idx) => {
      const rowNumber = idx + 2;
      const carId = cellText(row[idIdx]);
      const attachment = normalizeAttachment(row[attachmentIdx]);
      if (!carId) return null;
      return { carId, rowNumber, attachmentName: attachment.name, token: attachment.token };
    })
    .filter(Boolean);
}

async function readSnapshot(spreadsheetToken) {
  const values = await readRange(spreadsheetToken, `${CONFIG.snapshotSheetId}!A2:E1000`);
  const map = new Map();
  for (const row of values) {
    const carId = cellText(row[0]);
    if (!carId) continue;
    map.set(carId, {
      carId,
      rowNumber: Number(cellText(row[1])) || null,
      name: cellText(row[2]),
      token: cellText(row[3]),
      snapshotTime: cellText(row[4])
    });
  }
  return map;
}

async function writeSnapshot(spreadsheetToken, currentState) {
  const now = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai', hour12: false });
  const rows = [
    ['车源商品ID', '行号', '附件名', 'file_token', '快照时间'],
    ...currentState.map((item) => [item.carId, item.rowNumber, item.attachmentName || '', item.token || '', now])
  ];
  await writeRange(spreadsheetToken, `${CONFIG.snapshotSheetId}!A1:E${rows.length}`, rows);
}

async function readNamingFields(spreadsheetToken, rowNumber) {
  const row = (await readRange(spreadsheetToken, `${CONFIG.sourceSheetId}!A${rowNumber}:M${rowNumber}`))[0] || [];
  return {
    demandTime: cellText(row[colToIndex(CONFIG.demandTimeColumn)]),
    location: cellText(row[colToIndex(CONFIG.locationColumn)]),
    carId: cellText(row[colToIndex(CONFIG.idColumn)]),
    brand: cellText(row[colToIndex(CONFIG.brandColumn)]),
    series: cellText(row[colToIndex(CONFIG.seriesColumn)]),
    model: cellText(row[colToIndex(CONFIG.modelColumn)])
  };
}

async function findReportChatId() {
  const groupName = required('REPORT_GROUP_NAME');
  let pageToken = '';
  for (let i = 0; i < 5; i++) {
    const query = new URLSearchParams({ page_size: '100' });
    if (pageToken) query.set('page_token', pageToken);
    const data = await feishuFetch(`/im/v1/chats?${query.toString()}`, { method: 'GET' });
    const groups = data.data?.items || [];
    const found = groups.find((g) => g.name === groupName);
    if (found?.chat_id) return found.chat_id;
    pageToken = data.data?.page_token || '';
    if (!data.data?.has_more) break;
  }
  throw new Error(`未找到日报群：${groupName}。请确认应用机器人已加入该群，并已开通获取群列表权限。`);
}

async function sendTextMessage(text) {
  const chatId = process.env.REPORT_CHAT_ID || (await findReportChatId());
  const content = JSON.stringify({ text });
  return feishuFetch('/im/v1/messages?receive_id_type=chat_id', {
    method: 'POST',
    body: JSON.stringify({ receive_id: chatId, msg_type: 'text', content })
  });
}

function buildReport(changes) {
  const date = new Date().toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });
  const added = changes.filter((c) => c.type === '新增成片');
  const modified = changes.filter((c) => c.type === '修改成片');
  const cleared = changes.filter((c) => c.type === '异常清空');

  const lines = [];
  lines.push(`【外包成片附件更新日报｜${date}】`);
  lines.push('');
  lines.push('检测范围：外包对接表 AJ「输出成片（zip）」');
  lines.push(`本次检测到变化：${changes.length} 条`);
  lines.push(`新增成片：${added.length} 条；修改成片：${modified.length} 条；异常清空：${cleared.length} 条`);
  lines.push('');

  function section(title, items) {
    lines.push(`${title}：${items.length} 条`);
    if (!items.length) {
      lines.push('- 无');
      lines.push('');
      return;
    }
    items.slice(0, 30).forEach((item, index) => {
      lines.push(`${index + 1}. 行号：${item.rowNumber}`);
      lines.push(`   车源商品ID：${item.carId}`);
      if (item.vehicleText) lines.push(`   车辆：${item.vehicleText}`);
      if (item.location || item.demandTime) lines.push(`   所属地 / 需求时间：${item.location || '-'} / ${item.demandTime || '-'}`);
      if (item.oldName) lines.push(`   昨日附件：${item.oldName || '空'}`);
      if (item.newName) lines.push(`   今日附件：${item.newName || '空'}`);
      if (item.naming) {
        lines.push(`   命名校验：${item.naming.result}`);
        if (item.naming.problems.length) lines.push(`   问题：${item.naming.problems.join('；')}`);
      }
    });
    if (items.length > 30) lines.push(`   ……剩余 ${items.length - 30} 条未展示`);
    lines.push('');
  }

  section('一、新增成片', added);
  section('二、修改成片', modified);
  section('三、异常清空', cleared);

  lines.push('请审核以上成片视频。审核合格后，在源表标记「合格」，后续可同步到对客表。');
  return lines.join('\n');
}

async function runDailyCheck({ dryRun = false } = {}) {
  const sourceSpreadsheetToken = await resolveSpreadsheetToken(required('SOURCE_WIKI_TOKEN'));
  const currentState = await readCurrentAttachmentState(sourceSpreadsheetToken);
  const snapshot = await readSnapshot(sourceSpreadsheetToken);

  const changes = [];
  for (const item of currentState) {
    const old = snapshot.get(item.carId) || { name: '', token: '' };
    const type = compareAttachment(old, { name: item.attachmentName, token: item.token });
    if (type === '无变化') continue;

    const change = {
      type,
      rowNumber: item.rowNumber,
      carId: item.carId,
      oldName: old.name,
      newName: item.attachmentName
    };

    if (type !== '异常清空') {
      const namingFields = await readNamingFields(sourceSpreadsheetToken, item.rowNumber);
      change.vehicleText = [namingFields.brand, namingFields.series, namingFields.model].filter(Boolean).join(' ');
      change.location = namingFields.location;
      change.demandTime = namingFields.demandTime;
      change.naming = validateNaming({
        attachmentName: item.attachmentName,
        carId: namingFields.carId,
        brand: namingFields.brand,
        series: namingFields.series,
        model: namingFields.model,
        location: namingFields.location,
        demandTime: namingFields.demandTime
      });
    }
    changes.push(change);
  }

  const report = buildReport(changes);
  if (!dryRun) {
    await sendTextMessage(report);
    await writeSnapshot(sourceSpreadsheetToken, currentState);
  }
  return { changes, report, currentCount: currentState.length, dryRun };
}

export default async function handler(req, res) {
  try {
    const dryRun = req.query?.dryRun === '1' || req.query?.dryRun === 'true';
    const result = await runDailyCheck({ dryRun });
    res.status(200).json({ ok: true, ...result });
  } catch (error) {
    console.error(error);
    res.status(500).json({ ok: false, error: error.message });
  }
}

if (process.env.RUN_LOCAL === '1') {
  runDailyCheck({ dryRun: process.env.DRY_RUN === '1' })
    .then((result) => console.log(JSON.stringify(result, null, 2)))
    .catch((err) => {
      console.error(err);
      process.exit(1);
    });
}

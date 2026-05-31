const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const querystring = require('querystring');
const os = require('os');
const { execFileSync } = require('child_process');

const PORT = 3000;
const BRIDGE_HOST = '100.89.188.128';
const BRIDGE_PORT = 18800;
const BRIDGE_TOKEN = 'jzkj2026';

const ALIPAY_APP_ID = '2018123062720430';
const ALIPAY_PUBLIC_KEY = `-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAz64ANCgXnb8DVo3SS8T
797PA8bcRKuY/Yl2I5datO9tYe63Y8fmNJ4cWF6ik8TJCokQ1rE2YeAdzIXTsMR
3XyXsxRhbgrwW+WCMBZcBG3h0qwGls7J9uqE67oL7NjOxKhC3W7eVEtlY8AHnB
8U/rkNoRaORGnK8vaGsvvvXCeXdbX8l+gkmR1BoBa/NcApKYr37KO4TIPIF1/rp
PtkcejyQAmgT0zueceuAvsCMkW1upyn/MSLpjCbJMrIAFhFRn8ofhUe0lYZX5n3X
weorUPDESO/zuS/OdxApp0KK57KuvmMOa7x7Tqbzg+XHTAQCYo4Sh3dqynGuqYr
Xj+OxsQIDAQAB
-----END PUBLIC KEY-----`;

// ── 套餐：按次数，不按天 ──
const PLAN_MAP = {
  '19.00': { name: '体验包', quota: 500  },
  '49.00': { name: '标准包', quota: 2000 },
};

const WHITELIST_FILE = path.join(__dirname, 'whitelist.json');
const ORDERS_FILE    = path.join(__dirname, 'orders.json');
const QR_DIR         = path.join(__dirname, 'qrcodes');
const CASE_CONTEXTS_FILE = path.join(__dirname, 'case_contexts.json');
const CUSTOMER_MEMORY_FILE = path.join(__dirname, 'customer_memories.json');
const AUDIT_TASK_LOG_FILE = '/var/log/n1-audit-task.log';
const LOCAL_UPLOAD_ROOT = '/tmp/n1_uploads';
const CASE_CONTEXTS = readJSON(CASE_CONTEXTS_FILE);
const CUSTOMER_MEMORIES = readJSON(CUSTOMER_MEMORY_FILE);
let ACTIVE_CASE_ID = '';
const SMB_CONFIG_FILE = '/opt/ai001/smb_config.json';
const SUPPORTED_AUDIT_EXTENSIONS = new Set([
  '.pdf', '.xlsx', '.csv', '.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.zip',
]);
const IGNORED_RESULT_DIR_PATTERNS = [
  /^AI审计结果(?:_|$)/,
  /^AI票据结果(?:_|$)/,
  /^AI审计失败(?:_|$)/,
  /^AI审计结果$/,
  /_整理后(?:_|$)/,
  /^整理后(?:_|$)/,
];

// ──────────────────────────────────────────
// 工具函数
// ──────────────────────────────────────────
function serveFile(res, filePath, contentType) {
  fs.readFile(filePath, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
}

function readJSON(file) {
  try {
    if (!fs.existsSync(file)) return {};
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch { return {}; }
}

function writeJSON(file, data) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf8');
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function isSupportedAuditFile(fileName) {
  return SUPPORTED_AUDIT_EXTENSIONS.has(path.extname(fileName || '').toLowerCase());
}

function shouldIgnoreAuditRelativePath(relativePath) {
  const parts = String(relativePath || '').split(/[\\/]+/).filter(Boolean);
  return parts.some((part) => IGNORED_RESULT_DIR_PATTERNS.some((pattern) => pattern.test(part)));
}

function listMountedAuditFiles(rootDir) {
  const result = [];

  function walk(currentDir) {
    const names = fs.readdirSync(currentDir).sort((a, b) => a.localeCompare(b, 'zh-CN'));
    names.forEach((name) => {
      const absolutePath = path.join(currentDir, name);
      let stat;
      try {
        stat = fs.statSync(absolutePath);
      } catch {
        return;
      }

      if (stat.isDirectory()) {
        const relativeDir = path.relative(rootDir, absolutePath).replaceAll(path.sep, '/');
        if (shouldIgnoreAuditRelativePath(relativeDir)) {
          return;
        }
        walk(absolutePath);
        return;
      }

      if (!stat.isFile() || !isSupportedAuditFile(name)) {
        return;
      }

      const relativeFile = path.relative(rootDir, absolutePath).replaceAll(path.sep, '/');
      if (shouldIgnoreAuditRelativePath(relativeFile)) {
        return;
      }

      result.push({
        name: relativeFile,
        size: stat.size,
        isDir: false,
      });
    });
  }

  walk(rootDir);
  return result;
}

function readTailLines(filePath, limit = 200) {
  try {
    if (!fs.existsSync(filePath)) return [];
    const raw = fs.readFileSync(filePath, 'utf8');
    return raw.split('\n').filter(Boolean).slice(-limit).reverse();
  } catch {
    return [];
  }
}

function createCaseId() {
  return 'case_' + Date.now();
}

function slugifyName(value, fallback = '审计案件') {
  const text = String(value || '').trim();
  const cleaned = text
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return cleaned || fallback;
}

function normalizeCustomerKey(value) {
  return slugifyName(value, 'default_customer').toLowerCase();
}

function summarizeCustomerMemory(ctx) {
  if (!ctx) return null;
  return {
    customerKey: ctx.customerKey || '',
    customerName: ctx.customerName || '',
    caseName: ctx.caseName || '',
    accountNo: ctx.accountNo || '',
    cardNo: ctx.cardNo || '',
    bankName: ctx.bankName || '',
    reportBaseName: ctx.reportBaseName || '',
    preferredStrategy: ctx.strategy || '本地标准',
    lastAnalysisCondition: ctx.analysisCondition || '',
    lastResultStatus: ctx.resultStatus || '',
    lastResultDir: ctx.resultDir || '',
    lastResultFiles: Array.isArray(ctx.resultFiles) ? ctx.resultFiles : [],
    updatedAt: new Date().toISOString(),
  };
}

function saveCustomerMemory(ctx) {
  const key = ctx && ctx.customerKey ? ctx.customerKey : '';
  if (!key) return;
  CUSTOMER_MEMORIES[key] = {
    ...(CUSTOMER_MEMORIES[key] || {}),
    ...summarizeCustomerMemory(ctx),
  };
  writeJSON(CUSTOMER_MEMORY_FILE, CUSTOMER_MEMORIES);
}

function normalizeCaseContext(raw) {
  const base = raw && typeof raw === 'object' ? raw : {};
  const customerName = base.customerName || base.caseName || '';
  const reportBaseName = slugifyName(base.reportBaseName || customerName || base.caseName || '审计案件');
  const customerKey = normalizeCustomerKey(base.customerKey || customerName || base.caseName || reportBaseName);
  return {
    caseName: base.caseName || '未命名案件',
    customerName,
    customerKey,
    accountNo: base.accountNo || '',
    cardNo: base.cardNo || '',
    bankName: base.bankName || '',
    reportBaseName,
    primaryReportFile: base.primaryReportFile || '',
    resultStatus: base.resultStatus || '等待生成',
    selectedCount: Number(base.selectedCount || 0),
    strategy: base.strategy || '本地标准',
    selectedFiles: Array.isArray(base.selectedFiles) ? base.selectedFiles : [],
    taskType: base.taskType || 'case',
    analysisCondition: base.analysisCondition || '',
    lastTaskId: base.lastTaskId || '',
    resultDir: base.resultDir || '',
    sharePath: base.sharePath || '',
    resultFiles: Array.isArray(base.resultFiles) ? base.resultFiles : [],
    updatedAt: new Date().toISOString(),
  };
}

function saveCaseContext(caseId, raw) {
  const id = caseId || createCaseId();
  const prev = CASE_CONTEXTS[id] || {};
  CASE_CONTEXTS[id] = { id, ...prev, ...normalizeCaseContext({ ...prev, ...raw }) };
  ACTIVE_CASE_ID = id;
  writeJSON(CASE_CONTEXTS_FILE, CASE_CONTEXTS);
  saveCustomerMemory(CASE_CONTEXTS[id]);
  return CASE_CONTEXTS[id];
}

function getCaseContext(caseId) {
  if (caseId && CASE_CONTEXTS[caseId]) return CASE_CONTEXTS[caseId];
  if (ACTIVE_CASE_ID && CASE_CONTEXTS[ACTIVE_CASE_ID]) return CASE_CONTEXTS[ACTIVE_CASE_ID];
  return null;
}

function formatCaseContextPrompt(ctx) {
  if (!ctx) return '';
  const files = ctx.selectedFiles && ctx.selectedFiles.length ? ctx.selectedFiles.join(', ') : '未记录';
  const resultFiles = ctx.resultFiles && ctx.resultFiles.length ? ctx.resultFiles.join(', ') : '未记录';
  const memory = ctx.customerKey ? CUSTOMER_MEMORIES[ctx.customerKey] : null;
  const memoryBlock = memory ? [
    '【该客户本地记忆】',
    `客户名称: ${memory.customerName || '未记录'}`,
    `常用策略: ${memory.preferredStrategy || '未记录'}`,
    `最近要求: ${memory.lastAnalysisCondition || '未记录'}`,
    `最近结果状态: ${memory.lastResultStatus || '未记录'}`,
    `最近结果文件: ${memory.lastResultFiles && memory.lastResultFiles.length ? memory.lastResultFiles.join(', ') : '未记录'}`,
  ].join('\n') : '';
  return [
    '【当前案件上下文】',
    `案件名称: ${ctx.caseName || '未命名案件'}`,
    `客户名称: ${ctx.customerName || '未记录'}`,
    `客户标识: ${ctx.customerKey || '未记录'}`,
    `交易账号: ${ctx.accountNo || '未记录'}`,
    `交易卡号: ${ctx.cardNo || '未记录'}`,
    `开户行: ${ctx.bankName || '未记录'}`,
    `任务类型: ${ctx.taskType || 'case'}`,
    `结果状态: ${ctx.resultStatus || '等待生成'}`,
    `本次材料数: ${ctx.selectedCount || 0}`,
    `处理策略: ${ctx.strategy || '本地标准'}`,
    `本次文件: ${files}`,
    `案件要求: ${ctx.analysisCondition || '未填写'}`,
    `最近任务ID: ${ctx.lastTaskId || '未生成'}`,
    `结果目录: ${ctx.resultDir || '未生成'}`,
    `结果文件: ${resultFiles}`,
    memoryBlock,
    '请基于以上上下文回答，并默认延续当前案件，不要把用户当作新开话题。',
  ].filter(Boolean).join('\n');
}

function readSmbConfig() {
  try {
    if (!fs.existsSync(SMB_CONFIG_FILE)) return null;
    return JSON.parse(fs.readFileSync(SMB_CONFIG_FILE, 'utf8'));
  } catch {
    return null;
  }
}

function getCurrentMountedFiles() {
  const cfg = readSmbConfig();
  if (!cfg || !cfg.mountPoint) {
    return { ok: false, error: '未找到挂载配置' };
  }
  if (!fs.existsSync(cfg.mountPoint)) {
    return { ok: false, error: '挂载点不存在' };
  }
  try {
    const files = listMountedAuditFiles(cfg.mountPoint);
    const archiveInsights = files
      .filter((item) => /\.(zip|rar|7z|tar|gz)$/i.test(item.name || ''))
      .map((item) => {
        const absolutePath = path.join(cfg.mountPoint, item.name);
        const insight = inspectArchiveForPrep(absolutePath);
        return {
          path: item.name,
          ...insight,
        };
      });
    return { ok: true, config: cfg, files, archiveInsights };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function detectBankNameByText(text) {
  const source = String(text || '');
  if (/广发|cgb/i.test(source)) return '广发';
  if (/建设|建行|ccb/i.test(source)) return '建设';
  if (/民生|cmbc/i.test(source)) return '民生';
  if (/招商|招行|cmb/i.test(source)) return '招商';
  if (/农业|农行|abc/i.test(source)) return '农业';
  if (/工商|工行|icbc/i.test(source)) return '工商';
  if (/中国银行|中行|boc/i.test(source)) return '中国银行';
  if (/交通银行|交行|bocom/i.test(source)) return '交通';
  if (/邮储|邮政储蓄|psbc/i.test(source)) return '邮储';
  if (/平安|ping\s*an/i.test(source)) return '平安';
  if (/上海银行/i.test(source)) return '上海银行';
  return '';
}

function inferBankHints(files) {
  const hints = new Map();
  (files || []).forEach((item) => {
    const relativeName = String(item && item.name || '');
    const bankName = detectBankNameByText(relativeName);
    if (!bankName) return;
    const base = path.basename(relativeName).toLowerCase();
    if (!base) return;
    if (!hints.has(base)) hints.set(base, new Map());
    const bankMap = hints.get(base);
    bankMap.set(bankName, (bankMap.get(bankName) || 0) + 1);
  });
  return hints;
}

function pickHintedBank(baseName, hintMap) {
  const bankMap = hintMap.get(String(baseName || '').toLowerCase());
  if (!bankMap || !bankMap.size) return '';
  return Array.from(bankMap.entries()).sort((a, b) => b[1] - a[1])[0][0] || '';
}

function detectBankNameFromPath(name, hintMap) {
  const text = String(name || '');
  const direct = detectBankNameByText(text);
  if (direct) return direct;
  const hinted = pickHintedBank(path.basename(text), hintMap || new Map());
  if (hinted) return hinted;
  if (/信用卡开户|信用卡流水|开户资料|交易流水/.test(text)) {
    const fallback = pickHintedBank(path.basename(text), hintMap || new Map());
    if (fallback) return fallback;
  }
  return '未识别银行';
}

function bankFolderName(bankName) {
  if (bankName === '未识别银行') return '99_未识别材料';
  return `01_${bankName}`;
}

function materialSubfolder(fileName) {
  const text = String(fileName || '');
  const ext = path.extname(text).toLowerCase();
  if (/\.(zip|rar|7z|tar|gz)$/i.test(text)) return '98_压缩包待处理';
  if (ext === '.csv' || ext === '.xlsx' || ext === '.xls') return '01_表格材料';
  if (ext === '.pdf') return '02_PDF材料';
  if (['.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp'].includes(ext)) return '03_图片材料';
  return '09_其他材料';
}

function makeTimestampStamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function collectSupportedFiles(rootDir, currentDir = rootDir) {
  const result = [];
  const names = fs.readdirSync(currentDir).sort((a, b) => a.localeCompare(b, 'zh-CN'));
  names.forEach((name) => {
    const absolutePath = path.join(currentDir, name);
    let stat;
    try {
      stat = fs.statSync(absolutePath);
    } catch {
      return;
    }
    if (stat.isDirectory()) {
      result.push(...collectSupportedFiles(rootDir, absolutePath));
      return;
    }
    if (!stat.isFile() || !isSupportedAuditFile(name)) {
      return;
    }
    result.push({
      name: path.relative(rootDir, absolutePath).replaceAll(path.sep, '/'),
      size: stat.size,
      absolutePath,
    });
  });
  return result;
}

function extractZipForPrep(zipPath) {
  const extractRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'n1_prep_zip_'));
  try {
    execFileSync('/usr/bin/unzip', ['-o', '-qq', zipPath, '-d', extractRoot], { stdio: 'pipe' });
    const files = collectSupportedFiles(extractRoot);
    return { ok: true, extractRoot, files };
  } catch (e) {
    try { fs.rmSync(extractRoot, { recursive: true, force: true }); } catch {}
    return { ok: false, error: e.message || 'unzip_failed' };
  }
}

function classifyArchiveFailure(message) {
  const text = String(message || '');
  if (/password|unable to get password|encrypted/i.test(text)) {
    return {
      code: 'password_required',
      label: '压缩包带密码',
      hint: '该压缩包需要密码，当前不会自动解压，请先手动解压后再放入共享目录。',
    };
  }
  if (/End-of-central-directory|cannot find zipfile directory|not a zipfile|unsupported/i.test(text)) {
    return {
      code: 'invalid_archive',
      label: '压缩包格式异常',
      hint: '该压缩包可能损坏或格式不标准，建议重新获取原始压缩包或手动解压后再投入。',
    };
  }
  return {
    code: 'extract_failed',
    label: '压缩包解压失败',
    hint: '该压缩包未能自动解压，请人工检查后再处理。',
  };
}

function inspectArchiveForPrep(archivePath) {
  const name = path.basename(String(archivePath || ''));
  if (!/\.zip$/i.test(name)) {
    return {
      archive: name,
      code: 'unsupported_archive',
      label: '暂不支持自动探测',
      hint: '当前扫描阶段只对 zip 做快速探测，其他压缩格式先按人工处理看待。',
    };
  }
  try {
    execFileSync('/usr/bin/unzip', ['-t', '-qq', archivePath], { stdio: 'pipe' });
    return {
      archive: name,
      code: 'auto_extract_ready',
      label: '可自动解压',
      hint: '该压缩包扫描判断可自动解压，生成整理目录时会优先尝试自动展开。',
    };
  } catch (e) {
    return {
      archive: name,
      ...classifyArchiveFailure(e.message || 'unzip_failed'),
    };
  }
}

function organizeMountedMaterials(sourceSubDir) {
  const mounted = getCurrentMountedFiles();
  if (!mounted.ok) {
    return mounted;
  }
  const cfg = mounted.config;
  const rootDir = cfg.mountPoint;
  const subDir = (sourceSubDir || '').trim().replace(/^\/+|\/+$/g, '');
  let files;
  if (subDir) {
    const scanDir = path.join(rootDir, subDir);
    if (!fs.existsSync(scanDir)) {
      return { ok: false, error: `案件文件夹不存在: ${subDir}` };
    }
    files = listMountedAuditFiles(scanDir).map(f => ({ ...f, name: subDir + '/' + f.name }));
  } else {
    files = mounted.files || [];
  }
  const shareRootName = String(cfg.share || '当前材料').split(/[\\/]+/).filter(Boolean).pop() || '当前材料';
  const targetDirName = `${shareRootName}_整理后_${makeTimestampStamp()}`;
  const targetDir = path.join(rootDir, targetDirName);
  fs.mkdirSync(targetDir, { recursive: true });

  const copied = [];
  const skippedDuplicates = [];
  const extractedArchives = [];
  const archiveFailures = [];
  const seen = new Set();
  const hintMap = inferBankHints(files);

  function copyMaterial(sourcePath, relativeName, size, hintSourceMap) {
    const dedupeKey = `${path.basename(relativeName).toLowerCase()}__${size || 0}`;
    if (seen.has(dedupeKey)) {
      skippedDuplicates.push(relativeName);
      return;
    }
    seen.add(dedupeKey);

    const bankName = detectBankNameFromPath(relativeName, hintSourceMap || hintMap);
    const bankDir = bankFolderName(bankName);
    const subDir = materialSubfolder(relativeName);
    const targetFolder = path.join(targetDir, bankDir, subDir);
    fs.mkdirSync(targetFolder, { recursive: true });

    let targetName = path.basename(relativeName);
    let targetPath = path.join(targetFolder, targetName);
    let idx = 2;
    while (fs.existsSync(targetPath)) {
      const ext = path.extname(targetName);
      const stem = path.basename(targetName, ext);
      targetName = `${stem}_${idx}${ext}`;
      targetPath = path.join(targetFolder, targetName);
      idx += 1;
    }
    fs.copyFileSync(sourcePath, targetPath);
    copied.push({
      source: relativeName,
      bankName,
      target: path.relative(rootDir, targetPath).replaceAll(path.sep, '/'),
    });
  }

  files.forEach((item) => {
    const relativeName = String(item.name || '');
    const sourcePath = path.join(rootDir, relativeName);
    if (/\.zip$/i.test(relativeName)) {
      const extracted = extractZipForPrep(sourcePath);
      if (extracted.ok) {
        const nestedHintMap = inferBankHints(extracted.files.map((file) => ({ name: relativeName + '/' + file.name })));
        extracted.files.forEach((file) => {
          copyMaterial(
            file.absolutePath,
            `${relativeName}#解压/${file.name}`,
            file.size,
            nestedHintMap.size ? nestedHintMap : hintMap
          );
        });
        extractedArchives.push({
          archive: relativeName,
          extractedCount: extracted.files.length,
        });
        try { fs.rmSync(extracted.extractRoot, { recursive: true, force: true }); } catch {}
        return;
      }
      archiveFailures.push({
        archive: relativeName,
        error: extracted.error || 'unzip_failed',
        ...classifyArchiveFailure(extracted.error || 'unzip_failed'),
      });
    }
    copyMaterial(sourcePath, relativeName, item.size, hintMap);
  });

  if (!copied.length && archiveFailures.length) {
    archiveFailures.forEach((item) => {
      const sourcePath = path.join(rootDir, item.archive);
      const stat = fs.statSync(sourcePath);
      copyMaterial(sourcePath, item.archive, stat.size, hintMap);
    });
  }

  const readmePath = path.join(targetDir, '00_说明区');
  fs.mkdirSync(readmePath, { recursive: true });
  fs.writeFileSync(
    path.join(readmePath, '整理说明.txt'),
    [
      '本目录由材料整理自动生成。',
      '原始材料未被覆盖，请继续保留原始目录。',
      '建议先处理 01_表格材料 与少量关键 PDF。',
      extractedArchives.length
        ? `本次已自动解压压缩包 ${extractedArchives.length} 份。`
        : '本次未成功自动解压任何压缩包。',
      archiveFailures.length
        ? `以下压缩包未能自动解压，已保留待人工处理：${archiveFailures.map((item) => item.archive).join('；')}`
        : '无需要人工处理的压缩包。',
      skippedDuplicates.length ? `已跳过疑似重复材料 ${skippedDuplicates.length} 份。` : '未发现需要跳过的重复材料。',
    ].join('\n'),
    'utf8'
  );

  return {
    ok: true,
    config: cfg,
    targetDir,
    targetDirName,
    sharePath: toSharePath(targetDir),
    copiedCount: copied.length,
    skippedDuplicates,
    extractedArchives,
    archiveFailures,
    copied,
  };
}

function toSharePath(targetPath) {
  const cfg = readSmbConfig();
  if (!cfg || !cfg.mountPoint || !cfg.host || !cfg.share || !targetPath) return '';
  if (!targetPath.startsWith(cfg.mountPoint)) return '';
  const suffix = targetPath.slice(cfg.mountPoint.length).replace(/^\/+/, '').replaceAll('/', '\\');
  return `\\\\${cfg.host}\\${cfg.share}${suffix ? '\\' + suffix : ''}`;
}

function ensureVersionDir(baseDir, versionLabel) {
  const parent = path.dirname(baseDir);
  const baseName = path.basename(baseDir);
  let candidate = path.join(parent, `${baseName}_${versionLabel}`);
  let idx = 2;
  while (fs.existsSync(candidate)) {
    candidate = path.join(parent, `${baseName}_${versionLabel}_${idx}`);
    idx += 1;
  }
  fs.mkdirSync(candidate, { recursive: true });
  return candidate;
}

function findPrimaryReportFile(ctx) {
  if (!ctx || !ctx.resultDir) return null;
  const candidates = [];
  if (ctx.primaryReportFile) candidates.push(ctx.primaryReportFile);
  if (Array.isArray(ctx.resultFiles)) candidates.push(...ctx.resultFiles);
  const deduped = Array.from(new Set(candidates.filter(Boolean)));
  for (const name of deduped) {
    if (!/\.md$/i.test(name)) continue;
    const target = path.join(ctx.resultDir, name);
    if (fs.existsSync(target)) return { fileName: name, fullPath: target };
  }
  try {
    const files = fs.readdirSync(ctx.resultDir);
    const md = files.find((name) => /审计报告.*\.md$/i.test(name)) || files.find((name) => /\.md$/i.test(name));
    if (md) return { fileName: md, fullPath: path.join(ctx.resultDir, md) };
  } catch {}
  return null;
}

function summarizeResultStatus(taskType) {
  if (taskType === 'invoice') return '已生成票据台账';
  if (taskType === 'contract') return '已生成合同解析结果';
  if (taskType === 'custom') return '已生成自定义结果';
  return '已生成附表与报告';
}

function decodeBase64Payload(value) {
  const text = String(value || '');
  const idx = text.indexOf('base64,');
  const payload = idx >= 0 ? text.slice(idx + 7) : text;
  return Buffer.from(payload, 'base64');
}

// ──────────────────────────────────────────
// 白名单：按次数管理
// ──────────────────────────────────────────
function activateUser(outTradeNo, totalAmount, buyerId, passbackParams) {
  // 优先从 passback_params 拿 planId，降级用金额匹配
  let plan = PLAN_MAP[totalAmount];
  try {
    const pb = JSON.parse(decodeURIComponent(passbackParams || '{}'));
    if (pb.quota) plan = { name: pb.plan_id, quota: pb.quota };
  } catch {}

  if (!plan) {
    console.error('[激活失败] 未知金额:', totalAmount);
    return;
  }

  const wl = readJSON(WHITELIST_FILE);

  if (wl[buyerId]) {
    // 续费 → 叠加额度
    wl[buyerId].quota += plan.quota;
    wl[buyerId].total += plan.quota;
    wl[buyerId].orders = wl[buyerId].orders || [];
    wl[buyerId].orders.push(outTradeNo);
    wl[buyerId].last_renewed = new Date().toISOString();
    console.log(`[续费] ${buyerId} +${plan.quota}次，当前剩余: ${wl[buyerId].quota}`);
  } else {
    // 新用户
    wl[buyerId] = {
      quota: plan.quota,
      total: plan.quota,
      plan: plan.name,
      created_at: new Date().toISOString(),
      orders: [outTradeNo],
    };
    console.log(`[新开通] ${buyerId} ${plan.name} ${plan.quota}次`);
  }

  writeJSON(WHITELIST_FILE, wl);

  // 记录订单
  const orders = readJSON(ORDERS_FILE);
  orders[outTradeNo] = {
    buyerId, amount: totalAmount, plan: plan.name,
    quota: plan.quota, created_at: new Date().toISOString(),
  };
  writeJSON(ORDERS_FILE, orders);
}

function consumeQuota(userId) {
  const wl = readJSON(WHITELIST_FILE);
  const user = wl[userId];
  if (!user) return { allowed: false, reason: 'not_found' };
  if (user.quota <= 0) return { allowed: false, reason: 'exhausted', total: user.total };
  user.quota -= 1;
  writeJSON(WHITELIST_FILE, wl);
  return { allowed: true, remaining: user.quota };
}

function getQuotaInfo(userId) {
  const wl = readJSON(WHITELIST_FILE);
  return wl[userId] || null;
}

// ──────────────────────────────────────────
// 支付宝验签
// ──────────────────────────────────────────
function verifyAlipaySign(params) {
  const sign = params.sign;
  if (!sign) return false;
  const sortedKeys = Object.keys(params)
    .filter(k => k !== 'sign' && k !== 'sign_type')
    .sort();
  const signStr = sortedKeys.map(k => `${k}=${params[k]}`).join('&');
  try {
    const verify = crypto.createVerify('RSA-SHA256');
    verify.update(signStr, 'utf8');
    return verify.verify(ALIPAY_PUBLIC_KEY, sign, 'base64');
  } catch (e) {
    console.error('[验签异常]', e.message);
    return false;
  }
}

// ──────────────────────────────────────────
// 转发到 Mac Mini bridge
// ──────────────────────────────────────────
function callBridge(apiPath, payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const req = http.request({
      hostname: BRIDGE_HOST, port: BRIDGE_PORT,
      path: apiPath, method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${BRIDGE_TOKEN}`,
      },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ──────────────────────────────────────────
// Telegram Bot 主动推送
// ──────────────────────────────────────────
const TG_TOKEN = process.env.TG_BOT_TOKEN || '';

function tgSendMessage(chatId, text) {
  if (!TG_TOKEN) return;
  const body = JSON.stringify({
    chat_id: chatId,
    text,
    parse_mode: 'Markdown',
  });
  const req = https.request({
    hostname: 'api.telegram.org',
    path: `/bot${TG_TOKEN}/sendMessage`,
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  });
  req.write(body);
  req.end();
}

function tgSendPhoto(chatId, photoBuffer, caption) {
  // 用 sendPhoto multipart 发图
  if (!TG_TOKEN || !photoBuffer) {
    tgSendMessage(chatId, caption);
    return;
  }
  const boundary = '----TGBoundary' + Date.now();
  const CRLF = '\r\n';
  const head =
    `--${boundary}${CRLF}` +
    `Content-Disposition: form-data; name="chat_id"${CRLF}${CRLF}${chatId}${CRLF}` +
    `--${boundary}${CRLF}` +
    `Content-Disposition: form-data; name="caption"${CRLF}${CRLF}${caption}${CRLF}` +
    `--${boundary}${CRLF}` +
    `Content-Disposition: form-data; name="photo"; filename="qr.png"${CRLF}` +
    `Content-Type: image/png${CRLF}${CRLF}`;
  const tail = `${CRLF}--${boundary}--${CRLF}`;
  const headBuf = Buffer.from(head);
  const tailBuf = Buffer.from(tail);
  const fullBuf = Buffer.concat([headBuf, photoBuffer, tailBuf]);

  const req = https.request({
    hostname: 'api.telegram.org',
    path: `/bot${TG_TOKEN}/sendPhoto`,
    method: 'POST',
    headers: {
      'Content-Type': `multipart/form-data; boundary=${boundary}`,
      'Content-Length': fullBuf.length,
    },
  });
  req.write(fullBuf);
  req.end();
}

const https = require('https');

// ──────────────────────────────────────────
// HTTP 服务
// ──────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.writeHead(200); res.end(); return; }

  const url = new URL(req.url, `http://localhost:${PORT}`);

  // ── 静态页面 ──
  if (url.pathname === '/' || url.pathname === '/index.html') {
    serveFile(res, path.join(__dirname, "index.html"), 'text/html; charset=utf-8');
    return;
  }
  if (url.pathname === '/chat') {
    serveFile(res, path.join(__dirname, 'chat.html'), 'text/html; charset=utf-8');
    return;
  }
  if (url.pathname === '/pay' || url.pathname === '/pay.html') {
    serveFile(res, path.join(__dirname, 'pay.html'), 'text/html; charset=utf-8');
    return;
  }

  // ── 二维码图片静态服务 ──
  if (url.pathname.startsWith('/qrcodes/')) {
    const fname = path.basename(url.pathname);
    serveFile(res, path.join(QR_DIR, fname), 'image/png');
    return;
  }

  // ── 对话接口（加额度检查）──
  if (url.pathname === '/api/chat' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { message, device_id } = JSON.parse(body);

        // /buy 和 /quota 指令直接走 bridge buy 接口
        if (message && (message.startsWith('/buy') || message.startsWith('/quota'))) {
          const result = await callBridge('/api/buy', { message, userId: device_id });
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ reply: result.caption || result.reply || '请查看套餐', qr: result.qrBase64 || null }));
          return;
        }

        // 额度检查
        const check = consumeQuota(device_id);
        if (!check.allowed) {
          const reason = check.reason === 'not_found'
            ? '您还未开通服务，发送 /buy 查看套餐。'
            : `您的额度已用完（共${check.total}次）。\n发送 /buy 续费，用完即止，随时续。`;
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ reply: reason }));
          return;
        }

        console.log(`[${device_id}] 剩余${check.remaining}次 | ${message}`);
        const result = await callBridge('/api/chat', { message, agentId: 'shangye' });
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ reply: result.reply || '无回复' }));

      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ reply: '抱歉，服务暂时不可用。' }));
      }
    });
    return;
  }

  // ── 支付宝回调 ──
  if (url.pathname === '/api/pay/callback' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => { body += chunk.toString(); });
    req.on('end', () => {
      try {
        const params = querystring.parse(body);
        console.log('[支付回调]', params.out_trade_no, params.trade_status, params.total_amount);

        if (!verifyAlipaySign(params)) {
          console.error('[验签失败]');
          res.writeHead(400); res.end('验签失败'); return;
        }

        if (params.trade_status === 'TRADE_SUCCESS' || params.trade_status === 'TRADE_FINISHED') {
          activateUser(
            params.out_trade_no,
            params.total_amount,
            params.buyer_id,
            params.passback_params,
          );

          // 如果能从订单号拿到 Telegram chatId，主动推送激活消息
          // 订单号格式：JZKJ_plan_b_chatId_timestamp
          const parts = (params.out_trade_no || '').split('_');
          const chatId = parts[2];
          if (chatId) {
            const info = getQuotaInfo(params.buyer_id);
            const notifyText = '✅ *到账成功！*\n\n已为您激活 *' + (info ? info.quota : '') + '次* 对话额度。\n用完随时发 /buy 续费，额度自动叠加。';
            const notifyBody = JSON.stringify({token:'jzkj2026', chat_id: chatId, text: notifyText});
            const nReq = require('http').request({
              hostname: '192.168.2.118',
              port: 18800,
              path: '/api/notify',
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer jzkj2026',
                'Content-Length': Buffer.byteLength(notifyBody)
              }
            }, r => r.resume());
            nReq.on('error', e => console.error('[notify error]', e.message));
            nReq.write(notifyBody); nReq.end();
          }
        }

        // 管理员通知
        const adminText = '💰 *新订单到账！*\n\n套餐：' + (PLAN_MAP[params.total_amount] ? PLAN_MAP[params.total_amount].name + ' ¥' + params.total_amount : '¥' + params.total_amount) + '\n买家：' + params.buyer_id + '\n订单号：' + params.out_trade_no;
        const adminBody = JSON.stringify({token:'jzkj2026', chat_id:'7674359237', text: adminText});
        const aReq = require('http').request({
          hostname: '192.168.2.118', port: 18800, path: '/api/notify', method: 'POST',
          headers: {'Content-Type':'application/json','Authorization':'Bearer jzkj2026','Content-Length':Buffer.byteLength(adminBody)}
        }, r => r.resume());
        aReq.on('error', e => console.error('[admin notify error]', e.message));
        aReq.write(adminBody); aReq.end();

        res.writeHead(200, { 'Content-Type': 'text/plain' });
        res.end('success');
      } catch (e) {
        console.error('[回调异常]', e.message);
        res.writeHead(500); res.end('error');
      }
    });
    return;
  }

  // ── 额度查询接口 ──
  if (url.pathname === '/api/quota' && req.method === 'GET') {
    const userId = url.searchParams.get('userId') || '';
    const info = userId ? getQuotaInfo(userId) : null;
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(info || { error: 'not_found' }));
    return;
  }


  // ── 面板页面 ──
  if (url.pathname === '/panel') {
    serveFile(res, path.join(__dirname, 'n1_panel.html'), 'text/html; charset=utf-8');
    return;
  }

  // ── 审计日志接口 ──
  if (url.pathname === '/api/audit-log' && req.method === 'GET') {
    try {
      const limit = Number(url.searchParams.get('limit') || 200);
      const lines = readTailLines(AUDIT_TASK_LOG_FILE, Number.isFinite(limit) ? limit : 200);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true, file: AUDIT_TASK_LOG_FILE, lines }));
    } catch(e) {
      res.writeHead(500);
      res.end(JSON.stringify({ ok: false, file: AUDIT_TASK_LOG_FILE, lines: [], error: e.message }));
    }
    return;
  }

  // ── 审计任务提交 ──
  if (url.pathname === '/api/task-start' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { filePath, fileName, fileSize, task, condition, outputDir, selectedFiles, caseId, caseContext, customerName: bodyCustomerName, relatedSubjects: bodyRelatedSubjects, auditMode } = JSON.parse(body);
        const callerIp = (req.socket.remoteAddress || '').replace(/^::ffff:/, '') || '';
        fs.appendFileSync('/var/log/n1-audit-task.log', `[${new Date().toLocaleString('zh-CN', { hour12: false })}] API task-start file=${fileName || ''} task=${task || ''} path=${filePath || ''}
`);
        const savedCase = saveCaseContext(caseId, {
          ...(caseContext || {}),
          customerName: bodyCustomerName || (caseContext && caseContext.customerName) || '',
          relatedSubjects: bodyRelatedSubjects || (caseContext && caseContext.relatedSubjects) || '',
          taskType: task || 'case',
          analysisCondition: condition || '',
          selectedFiles: Array.isArray(selectedFiles) ? selectedFiles : [],
          selectedCount: Array.isArray(selectedFiles) ? selectedFiles.length : 0,
          resultStatus: '处理中',
        });
        const taskManager = require('./task-manager');
        const taskId = taskManager.createTask({
          localPath: filePath,
          name: fileName,
          size: fileSize,
          taskType: savedCase.taskType || task || 'case',
          condition,
          outputDir,
          selectedFiles,
          caseId: savedCase.id,
          caseName: savedCase.caseName,
          customerName: savedCase.customerName,
          customerKey: savedCase.customerKey,
          accountNo: savedCase.accountNo,
          cardNo: savedCase.cardNo,
          bankName: savedCase.bankName,
          reportBaseName: savedCase.reportBaseName,
          auditMode: auditMode || 'hybrid',
          callerIp,
        });
        saveCaseContext(savedCase.id, { lastTaskId: taskId });
        taskManager.runTask(taskId);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, taskId, caseId: savedCase.id }));
      } catch(e) {
        res.writeHead(500); res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
    return;
  }

  if (url.pathname === '/api/task-upload-start' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { files, task, condition, caseId, caseContext, auditMode } = JSON.parse(body || '{}');
        const uploadFiles = Array.isArray(files) ? files : [];
        if (!uploadFiles.length) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: '未收到本机上传文件' }));
          return;
        }
        ensureDir(LOCAL_UPLOAD_ROOT);
        const savedCase = saveCaseContext(caseId, {
          ...(caseContext || {}),
          taskType: task || 'case',
          analysisCondition: condition || '',
          selectedFiles: uploadFiles.map((item) => item.name).filter(Boolean),
          selectedCount: uploadFiles.length,
          resultStatus: '处理中',
        });
        const uploadDir = path.join(LOCAL_UPLOAD_ROOT, `${Date.now()}_${savedCase.id}`);
        ensureDir(uploadDir);
        uploadFiles.forEach((item, index) => {
          const safeName = slugifyName(item.name || `upload_${index + 1}`, `upload_${index + 1}`);
          const target = path.join(uploadDir, safeName);
          fs.writeFileSync(target, decodeBase64Payload(item.contentBase64 || ''));
        });
        fs.appendFileSync(AUDIT_TASK_LOG_FILE, `[${new Date().toLocaleString('zh-CN', { hour12: false })}] API task-upload-start task=${task || ''} files=${uploadFiles.length} dir=${uploadDir}\n`);
        const taskManager = require('./task-manager');
        const taskId = taskManager.createTask({
          localPath: uploadDir,
          localUploadDir: uploadDir,
          name: '本机上传文件',
          size: uploadFiles.reduce((sum, item) => sum + Number(item.size || 0), 0),
          taskType: savedCase.taskType || task || 'case',
          auditMode: auditMode || 'hybrid',
          condition,
          outputDir: uploadDir,
          selectedFiles: uploadFiles.map((item) => item.name).filter(Boolean),
          caseId: savedCase.id,
          caseName: savedCase.caseName,
          customerName: savedCase.customerName,
          customerKey: savedCase.customerKey,
          accountNo: savedCase.accountNo,
          cardNo: savedCase.cardNo,
          bankName: savedCase.bankName,
          reportBaseName: savedCase.reportBaseName,
        });
        saveCaseContext(savedCase.id, { lastTaskId: taskId });
        taskManager.runTask(taskId);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, taskId, caseId: savedCase.id }));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
    return;
  }

  // ── 审计任务状态 ──
  if (url.pathname === '/api/task-status' && req.method === 'GET') {
    const id = url.searchParams.get('id') || '';
    const taskManager = require('./task-manager');
    const t = taskManager.getTask(id);
    if (t && t.fileInfo && t.fileInfo.caseId && t.status === 'done') {
      const resultFiles = t.result && Array.isArray(t.result.files) ? t.result.files : [];
      const taskType = (t.result && t.result.taskType) || t.fileInfo.taskType || 'case';
      const primaryReport = resultFiles.find((name) => /审计报告.*\.md$/i.test(name)) || resultFiles.find((name) => /\.md$/i.test(name)) || '';
      saveCaseContext(t.fileInfo.caseId, {
        resultStatus: summarizeResultStatus(taskType),
        lastTaskId: id,
        resultDir: t.result && t.result.outDir ? t.result.outDir : '',
        sharePath: t.result && t.result.sharePath ? t.result.sharePath : '',
        resultFiles,
        primaryReportFile: primaryReport,
      });
    }
    if (t && t.fileInfo && t.fileInfo.caseId && t.status === 'failed') {
      saveCaseContext(t.fileInfo.caseId, {
        resultStatus: '任务失败',
        lastTaskId: id,
      });
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(t || { error: 'not_found' }));
    return;
  }

  // ── 案件上下文接口 ──
  if (url.pathname === '/api/case-context' && req.method === 'GET') {
    const caseId = url.searchParams.get('caseId') || '';
    const ctx = getCaseContext(caseId);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, caseContext: ctx, activeCaseId: ACTIVE_CASE_ID }));
    return;
  }

  if (url.pathname === '/api/case-context' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { caseId, caseContext } = JSON.parse(body || '{}');
        const saved = saveCaseContext(caseId, caseContext || {});
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, caseId: saved.id, caseContext: saved }));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
    return;
  }

  // ── 案件助手接口 ──
  if (url.pathname === '/api/chat-proxy' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { message, caseId, mode } = JSON.parse(body || '{}');
        const ctx = getCaseContext(caseId);
        const enriched = `${formatCaseContextPrompt(ctx)}\n\n【会话模式】${mode || 'assistant'}\n【用户消息】${message || ''}`;
        const result = await callBridge('/api/chat', { message: enriched, agentId: 'shangye' });
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, reply: result.reply || '无回复', caseContext: ctx || null }));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, reply: '抱歉，案件助手暂时不可用。', error: e.message }));
      }
    });
    return;
  }

  if (url.pathname === '/api/refine-writeback' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { caseId, instruction } = JSON.parse(body || '{}');
        const ctx = getCaseContext(caseId);
        if (!ctx || !ctx.resultDir) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: '当前案件还没有可回写的结果目录' }));
          return;
        }

        const primaryReport = findPrimaryReportFile(ctx);
        if (!primaryReport) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: '当前结果目录中未找到审计报告.md' }));
          return;
        }

        const sourceContent = fs.readFileSync(primaryReport.fullPath, 'utf8');
        const bridgePayload = {
          message: [
            formatCaseContextPrompt(ctx),
            '【任务】请基于原始 Markdown 审计报告和用户微调要求，输出一份可直接保存的新 Markdown 审计报告。',
            '【要求】直接输出完整 Markdown 正文，不要加解释，不要加代码块围栏。',
            `【用户微调要求】${instruction || ''}`,
            '【原始报告】',
            sourceContent,
          ].join('\n\n'),
          agentId: 'shangye',
        };
        const result = await callBridge('/api/chat', bridgePayload);
        const refined = (result && result.reply ? String(result.reply) : '').trim();
        if (!refined) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: '微调结果为空' }));
          return;
        }

        const versionDir = ensureVersionDir(ctx.resultDir, 'V2');
        const carryFiles = (ctx.resultFiles || []).filter(Boolean);
        carryFiles.forEach((name) => {
          const src = path.join(ctx.resultDir, name);
          const dst = path.join(versionDir, name);
          if (fs.existsSync(src) && !fs.existsSync(dst)) {
            fs.copyFileSync(src, dst);
          }
        });

        const fileName = primaryReport.fileName;
        const targetPath = path.join(versionDir, fileName);
        fs.writeFileSync(targetPath, refined + '\n', 'utf8');

        const updatedFiles = Array.from(new Set([...(ctx.resultFiles || []), fileName]));
        const saved = saveCaseContext(caseId, {
          resultStatus: '已生成微调版本',
          resultDir: versionDir,
          sharePath: toSharePath(versionDir),
          resultFiles: updatedFiles,
          primaryReportFile: fileName,
        });

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          ok: true,
          message: '微调结果已写回共享目录',
          fileName,
          versionDirName: path.basename(versionDir),
          resultDir: versionDir,
          sharePath: toSharePath(versionDir),
          filePath: targetPath,
          fileSharePath: toSharePath(targetPath),
          caseContext: saved,
        }));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
    return;
  }

  // ── SMB挂载接口 ──
  if (url.pathname === '/api/smb-mount' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { host, share, user, pass } = JSON.parse(body);
        const mountPoint = '/mnt/smb_client';
        if (!fs.existsSync(mountPoint)) fs.mkdirSync(mountPoint, { recursive: true });
        const { execSync } = require('child_process');
        try { execSync(`umount ${mountPoint} 2>/dev/null || true`); } catch(e) {}
        try {
          execSync(`mount -t cifs //${host}/${share} ${mountPoint} -o username=${user||'guest'},password=${pass||''},iocharset=utf8,vers=3.0 2>&1`);
        } catch(me) {
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: me.message }));
          return;
        }
        // 挂载成功后保存配置，供 run_audit.sh 动态读取
        fs.writeFileSync('/opt/ai001/smb_config.json', JSON.stringify({ host, share, mountPoint, user: user||'guest' }, null, 2));
        const files = listMountedAuditFiles(mountPoint);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, files }));
      } catch(e) {
        res.writeHead(500); res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
    return;
  }

  if (url.pathname === '/api/smb-config' && req.method === 'GET') {
    try {
      const cfg = readSmbConfig();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        ok: true,
        config: cfg || null,
      }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  if (url.pathname === '/api/smb-files' && req.method === 'GET') {
    try {
      const mounted = getCurrentMountedFiles();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(mounted));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  if (url.pathname === '/api/material-prep-organize' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        let sourceSubDir = '';
        try { sourceSubDir = (JSON.parse(body).sourceSubDir || '').trim(); } catch(e) {}
        const result = organizeMountedMaterials(sourceSubDir);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
    return;
  }

// ── 批量审计：整理后目录一键提交所有银行任务 ──
  if (url.pathname === '/api/batch-task-from-organized' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { organizedDirName, customerName } = JSON.parse(body);
        const cfg = readSmbConfig();
        if (!cfg || !cfg.mountPoint) { res.writeHead(400); res.end(JSON.stringify({ ok: false, error: '未找到挂载配置' })); return; }
        const organizedDir = path.join(cfg.mountPoint, organizedDirName);
        if (!fs.existsSync(organizedDir)) { res.writeHead(400); res.end(JSON.stringify({ ok: false, error: '整理目录不存在: ' + organizedDirName })); return; }

        const bankFolders = fs.readdirSync(organizedDir).filter(e =>
          /^01_/.test(e) && fs.statSync(path.join(organizedDir, e)).isDirectory()
        );
        if (!bankFolders.length) { res.writeHead(400); res.end(JSON.stringify({ ok: false, error: '未找到银行子目录（01_XXX）' })); return; }

        const taskManager = require('./task-manager');
        const tasks = [];

        for (const bankFolder of bankFolders) {
          const bankName = bankFolder.replace(/^01_/, '');
          const bankDir = path.join(organizedDir, bankFolder);
          const selectedFiles = [];
          for (const subDir of ['01_表格材料', '02_PDF材料']) {
            const subPath = path.join(bankDir, subDir);
            if (!fs.existsSync(subPath)) continue;
            fs.readdirSync(subPath).forEach(f => {
              if (/^\.|~$/.test(f)) return;
              selectedFiles.push(path.join(organizedDirName, bankFolder, subDir, f));
            });
          }
          if (!selectedFiles.length) continue;

          const taskCustomerName = customerName ? customerName + '_' + bankName : bankName;
          const savedCase = saveCaseContext(null, {
            customerName: taskCustomerName,
            taskType: 'case',
            selectedFiles,
            selectedCount: selectedFiles.length,
            resultStatus: '处理中',
          });
          const callerIpBatch = (req.socket.remoteAddress || '').replace(/^::ffff:/, '') || '';
          const taskId = taskManager.createTask({
            localPath: cfg.mountPoint,
            name: 'SMB共享目录',
            taskType: 'case',
            selectedFiles,
            caseId: savedCase.id,
            caseName: savedCase.caseName,
            customerName: taskCustomerName,
            customerKey: taskCustomerName,
            bankName,
            reportBaseName: taskCustomerName,
            auditMode: 'hybrid',
            callerIp: callerIpBatch,
          });
          saveCaseContext(savedCase.id, { lastTaskId: taskId });
          taskManager.runTask(taskId);
          tasks.push({ bank: bankName, taskId, fileCount: selectedFiles.length, caseId: savedCase.id });
        }

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true, tasks }));
      } catch(e) {
        res.writeHead(500); res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
    return;
  }

    res.writeHead(404); res.end('Not found');
});

server.listen(PORT, () => {
  console.log(`✅ N1 Server :${PORT} → Bridge ${BRIDGE_HOST}:${BRIDGE_PORT}`);
  console.log(`   额度模式：按次扣减，用完续费`);
});

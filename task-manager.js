const fs = require('fs');
const { execSync } = require('child_process');
const http = require('http');

const TASKS = {};

function createTask(fileInfo) {
  const id = 'task_' + Date.now();
  TASKS[id] = { id, status: 'pending', fileInfo, createdAt: new Date().toISOString(), result: null, error: null };
  return id;
}

function getTask(id) {
  return TASKS[id] || null;
}

function runTask(id) {
  const task = TASKS[id];
  if (!task) return;
  task.status = 'running';
  task.startedAt = new Date().toISOString();

  // 第一步：SCP文件到Mac Mini
  try {
    execSync(`scp -o StrictHostKeyChecking=no ${task.fileInfo.localPath} lin@192.168.2.111:/tmp/n1_inbound/${task.fileInfo.name}`);
  } catch(e) {
    task.status = 'failed';
    task.error = 'SCP失败: ' + e.message;
    return;
  }

  // 第二步：调墨影处理，指定输出目录
  const remotePath = '/tmp/n1_inbound/' + task.fileInfo.name;
  const resultDir = '/tmp/n1_results/' + id;
  const message = `按 PLAYBOOK_司法审计.md 做。文件：${remotePath} 暂无嫌疑人，先出通用版。结果文件全部输出到目录：${resultDir}`;

  const body = JSON.stringify({ message, agentId: 'shangye' });

  const req = http.request({
    hostname: '192.168.2.111',
    port: 18800,
    path: '/api/chat',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer jzkj2026',
      'Content-Length': Buffer.byteLength(body)
    },
    timeout: 600000
  }, res => {
    let data = '';
    res.on('data', chunk => data += chunk);
    res.on('end', () => {
      try {
        // 不等bridge返回，直接轮询结果目录
        pollResults(id, resultDir, 0);
      } catch(e) {
        task.status = 'failed';
        task.error = e.message;
      }
    });
  });

  req.on('error', () => {
    // bridge超时没关系，继续轮询结果目录
    pollResults(id, resultDir, 0);
  });

  req.on('timeout', () => {
    req.destroy();
    pollResults(id, resultDir, 0);
  });

  req.write(body);
  req.end();

  // 启动后立即开始轮询（不等bridge）
  setTimeout(() => pollResults(id, resultDir, 0), 30000);
}

function pollResults(id, resultDir, attempts) {
  const task = TASKS[id];
  if (!task || task.status === 'done' || task.status === 'failed') return;
  if (attempts > 20) {
    task.status = 'failed';
    task.error = '超时：墨影10分钟内未输出结果';
    return;
  }

  // 检查Mac Mini上结果目录是否有文件
  try {
    const output = execSync(`ssh -o StrictHostKeyChecking=no lin@192.168.2.111 "ls ${resultDir}/*.xlsx ${resultDir}/*.docx 2>/dev/null | head -5"`).toString().trim();
    if (output) {
      // 有结果文件，回传到N1
      const localResultDir = '/tmp/n1_results/' + id;
      if (!fs.existsSync(localResultDir)) fs.mkdirSync(localResultDir, {recursive:true});
      execSync(`scp -o StrictHostKeyChecking=no -r lin@192.168.2.111:${resultDir}/ /tmp/n1_results/`);

      // 复制到SMB挂载目录
      const smbResultDir = '/mnt/smb_n1/AI审计结果';
      try {
        if (!fs.existsSync(smbResultDir)) fs.mkdirSync(smbResultDir, {recursive:true});
        execSync(`cp -r ${localResultDir}/* ${smbResultDir}/`);
      } catch(e) {}

      const files = output.split('\n').filter(Boolean).map(f => f.split('/').pop());
      task.status = 'done';
      task.result = { files, resultDir: localResultDir, message: '审计完成，结果已回传到您的共享目录' };
      task.finishedAt = new Date().toISOString();

      // 写审计日志
      const audit = require('./audit');
      audit.writeAudit({
        operation: 'AUDIT_DONE',
        fileSize: task.fileInfo.size,
        modelName: 'audit-local-ocr',
        uploaded: false,
        durationMs: Date.now() - new Date(task.startedAt).getTime()
      });
      return;
    }
  } catch(e) {}

  // 还没有结果，30秒后再查
  setTimeout(() => pollResults(id, resultDir, attempts + 1), 30000);
}

module.exports = { createTask, getTask, runTask };

# audit-bank-ocr · 司法银行流水审计 Skill

> **最后更新：** 2026-05-29（首版，基于周贤案完整流程沉淀）  
> **适用系统：** N1 盒子（192.168.2.109）+ Mac Mini M4 算力节点

---

## 一、Skill 能做什么

接收法院司法查询的银行材料（xlsx / csv / PDF），自动：
1. OCR 解析流水 → 结构化附表 xlsx
2. 生成审计报告（.docx + .md）
3. 结果写回 SMB 共享目录
4. 多银行汇总 → 总台账底稿 + 汇总分析报告

---

## 二、输入材料准备

### 2.1 支持的文件类型

| 类型 | 说明 | 质量 |
|------|------|------|
| `.xlsx` / `.csv` | 银行原始导出结构化流水 | ★★★ 最佳 |
| `.pdf`（文本层） | 内嵌文字的 PDF，直接提取 | ★★ 良好 |
| `.pdf`（扫描件） | 图像 PDF，走 OCR+Vision | ★ 取决于清晰度 |

### 2.2 关键规则：每银行单独跑

**不要把多个银行的文件混在一个目录里触发单次任务。**  
系统优先级：发现 xlsx/csv → 跳过 PDF OCR。混合输入会导致 PDF 被忽略。

正确做法：
```
广发 → 单独任务（只放广发文件）
建设 → 单独任务（只放建设文件）
民生 → 单独任务（只放民生文件）
最后 → 手动汇总
```

### 2.3 文件放置位置

```bash
# 文件放到 SMB 根目录，用 selectedFiles 精确指定
/mnt/smb_client/文件名.xlsx
/mnt/smb_client/文件名.pdf
```

---

## 三、触发审计任务

### 3.1 标准触发（xlsx/csv 或简单 PDF）

```bash
curl -s -X POST http://192.168.2.109:3000/api/task-start \
  -H "Content-Type: application/json" \
  -d '{
    "task": "case",
    "customerName": "被审计人姓名",
    "selectedFiles": [
      "文件1.xlsx",
      "文件2.pdf"
    ]
  }'
```

返回：`{"ok":true,"taskId":"task_xxx","caseId":"case_xxx"}`

### 3.2 强制 vision 模式（复杂 PDF 格式）

部分银行 PDF 有文本层但格式特殊（如民生银行多列排版），文本解析器识别率极低。
需在 `run_audit_local_ocr.py` 直接调用：

```bash
python3 /Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/scripts/run_audit_local_ocr.py \
  --input-dir /tmp/task_xxx/源数据 \
  --template /Volumes/AI_Agent/AI_Controlled_Zone/audit_base/空白审计模板.xlsx \
  --output-xlsx /tmp/task_xxx/output.xlsx \
  --artifacts-dir /tmp/task_xxx/审计中间结果 \
  --customer-name "姓名" \
  --mode vision
```

OCR 完成后，把 output.xlsx 重命名含 `审计附表` 上传 SMB，再触发报告任务：

```bash
# 上传 xlsx（用中转方式避免中文路径问题）
sshpass -p '139319' scp output.xlsx root@192.168.2.109:/tmp/temp.xlsx
sshpass -p '139319' ssh root@192.168.2.109 "cp /tmp/temp.xlsx '/mnt/smb_client/XXX_审计附表.xlsx'"

# 触发报告生成
curl -s -X POST http://192.168.2.109:3000/api/task-start \
  -d '{"task":"case","customerName":"姓名","selectedFiles":["XXX_审计附表.xlsx"]}'
```

---

## 四、监控任务进度

```bash
# 实时日志
sshpass -p '139319' ssh root@192.168.2.109 "tail -f /var/log/n1-audit-task.log"

# 本地 OCR 进度（Mac 上）
ls ~/tmp/audit-local-ocr/*/vision/ | wc -l  # 已处理页数

# 关键完成标志
grep "结果已写回" /var/log/n1-audit-task.log
```

---

## 五、结果汇总（多银行）

系统暂无 API 接口，手动命令行完成：

```bash
# 1. 建汇总目录
OUTDIR="/mnt/smb_client/AI案件汇总_${姓名}_$(date +%Y%m%d_%H%M%S)"
sshpass -p '139319' ssh root@192.168.2.109 "
  mkdir -p '$OUTDIR/01_银行A' '$OUTDIR/02_银行B' '$OUTDIR/04_总台账底稿'
  cp /mnt/smb_client/AI审计结果_XXX1/* '$OUTDIR/01_银行A/'
  cp /mnt/smb_client/AI审计结果_XXX2/* '$OUTDIR/02_银行B/'
"

# 2. 合并总台账（Mac 本地 Python）
# 参考 /tmp/总台账底稿_*.csv 生成逻辑

# 3. 生成汇总分析报告
# 参考 00_汇总分析.md/docx 模板
```

---

## 六、结果目录结构

```
AI案件汇总_姓名_时间戳/
├── 00_汇总分析.docx   ← 三账户总览+风险红旗（参照国际审计标准）
├── 00_汇总分析.md
├── 01_广发/
│   ├── xxx_审计附表.xlsx
│   ├── xxx_审计报告.docx
│   └── xxx_审计报告.md
├── 02_建设/...
├── 03_民生/...
└── 04_总台账底稿/
    └── 总台账底稿.csv   ← 所有银行交易合并，首列为银行来源
```

---

## 七、常见问题

| 现象 | 原因 | 解决 |
|------|------|------|
| SOURCES 只有 xlsx，PDF 被忽略 | 目录里有结构化文件，PDF 被跳过 | 各银行单独任务 |
| `no transactions parsed` | PDF 格式特殊，文本解析器不认识 | 用 `--mode vision` 直接调 OCR 脚本 |
| vision 全返回 0 条 | `extract_json_object` 未定义 | 已修复（见 BUGS_FIXED.md #3） |
| `Permission denied (publickey,password)` | SSH key 未加载 | `sshpass -p '139319'` 前缀 |
| 文件被锁定无法覆盖 | Windows Excel 开着该文件 | 关闭 Excel 再操作 |
| `/tmp/` 下 tesseract 报找不到文件 | macOS TCC 沙盒限制 | 已修复（见 BUGS_FIXED.md #2） |

---

## 八、关键路径速查

| 资源 | 路径 |
|------|------|
| N1 面板（局域网） | http://192.168.2.109:3000/panel |
| N1 面板（公网） | https://n2.jzkjbj.cn/panel |
| 审计日志 | N1: /var/log/n1-audit-task.log |
| OCR 脚本 | Mac: /Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/scripts/run_audit_local_ocr.py |
| 审计脚本 | N1: /opt/ai001/run_audit.rewired.sh |
| 结果目录 | N1: /mnt/smb_client/AI审计结果_* |
| 临时目录 | Mac: ~/tmp/audit-local-ocr/ |
| Vision 配置 | Mac: ~/.openclaw/openclaw.json → models.providers.aliyun |

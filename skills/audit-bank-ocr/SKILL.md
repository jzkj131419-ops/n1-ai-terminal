# audit-bank-ocr · 司法银行流水审计 Skill

> **版本：** v2 · 2026-05-29  
> **主 PLAYBOOK：** `/Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/PLAYBOOK_司法审计.md`  
> **适用：** 非吸 / 职侵 / 诈骗 类经济案件银行流水审计

---

## 0. 任务识别

满足以下任一特征即按本手册执行：
- 用户说"银行流水审计"、"司法审计流水"、"职务侵占分析"、"非吸案分析"
- 用户提供 PDF/xlsx 路径，文件名含数字流水号
- 用户提到"嫌疑人 XXX"、"分析与 XXX 有关的资金"

**不适用：** 个人记账、对账核账、税务申报。

---

## 一、客户主操作流程

> 正常情况下客户自己在 N1 面板完成，全程不需要命令行。

### 第1步：上传 / 挂载材料

将银行材料（xlsx / csv / PDF / ZIP）放入 SMB 共享目录，支持子文件夹。  
挂载路径示例：`\\192.168.2.109\smb_share\` 或通过面板"挂载"入口填写路径。

### 第2步：材料整理与去重

打开面板左侧"**材料整理**"入口，点击"扫描"：
- 自动识别重复文件（内容级去重）
- ZIP 自动解压（最多2层），生成 `整理后_时间戳` 目录
- 超过2层 / 带密码 / 异常包 → 标记提示，需人工处理
- 确认无误后点"整理"，将结果同步到"案件助手"

> 若同批材料包含**明显不相关的主体或跨案件混入**，在此阶段手动拆分，不要混入同一案件任务。

### 第3步：生成案件结果

1. 面板左侧选"**案件助手**"
2. 确认挂载路径，选中整理好的材料文件
3. 填写**被审计人姓名**，点击"开始审计"

系统自动完成：OCR 解析 → 生成审计附表 xlsx → 生成报告 → 写回 SMB。  
结果写回到 `AI审计结果_时间戳/`，包含：
- `xxx_审计附表.xlsx`（主交付物，含14+张分析 sheet）
- `xxx_审计报告.docx`（14节 Word 正式报告）
- `xxx_审计报告.md`（同步 Markdown）

### 第4步：验收结果

**客户版（日常验收）：**
- 结果目录已正常生成（无 `AI审计失败_*`）
- 打开 `xxx_审计报告.docx`，确认被审计人姓名、账户期间、收支合计无明显异常
- 附表 sheet `1交易记录汇总` 有数据，条数不为零

**内部版（精确核验，交付前必做）：**
对照 **PDF 第一页头部** 的三个合计数：

| 验收项 | 方法 |
|--------|------|
| 条数对上 | 解析条数 = PDF 头部"交易明细合计" |
| 净额对上 | 解析(收入-支出) = PDF头部(收入合计-支出合计) |
| 差额定位 | 有差额必须定位到具体冲正/红冲交易 |

三条全过 → 数据可信，报告可发。有差额参见"排障 §T3"。

---

## 一A、可选增强

### 增强1：嫌疑人专题分析

获得嫌疑人姓名后补做，聚焦5大红旗：
1. 单笔精确 ¥50,000（规避监管卡限）
2. 单日多笔向同一对手（化整为零）
3. 摘要无业务实质（全是"转账/汇入"）
4. 极度单向性（流出>>流入）
5. 即收即转（公司大额收款当日转嫌疑人）

> 详见 PLAYBOOK §6 + 本地 LLM 批量判定 §6.5。

### 增强2：多银行汇总

多个银行分别出审计结果后，合并到一个交付目录：
```
AI案件汇总_姓名_时间戳/
├── 00_汇总分析.docx   ← 多账户总览+风险红旗
├── 01_广发/
├── 02_建设/
├── 03_民生/
└── 04_总台账底稿/总台账底稿.csv
```
> 当前面板"案件汇总"按钮暂无后端，需命令行操作，详见"排障 §T7"。

---

## 二、交付物清单

每单完整交付：

```
AI案件汇总_姓名_时间戳/
├── 00_汇总分析.docx/.md         ← 多银行总览+风险红旗（参照国际审计标准）
├── 01_银行A/
│   ├── xxx_审计附表.xlsx         ← 主交付物，含1-17张sheet
│   ├── xxx_审计报告.docx         ← 14节正式报告
│   └── xxx_审计报告.md
├── 02_银行B/...
└── 04_总台账底稿/
    └── 总台账底稿.csv            ← 所有银行合并，首列为银行来源
```

如有嫌疑人专题：额外附 `xxx_嫌疑人_专题分析.docx/.md`。

---

## 三、特殊情况 / 内部排障

> 以下内容是**技术人员排障用**，不属于客户操作流程。

### T1：某银行 PDF 被系统忽略（SOURCES 只有 xlsx）

**原因：** 同一目录里有 xlsx/csv 文件时，系统优先走结构化解析，跳过 PDF OCR。

**正确做法：** 每个银行单独发起一次任务，不要把多家银行的文件放在同一个目录触发。

```bash
# 广发（xlsx）单独任务
curl -s -X POST http://192.168.2.109:3000/api/task-start \
  -H "Content-Type: application/json" \
  -d '{"task":"case","customerName":"周贤","selectedFiles":["55776信用卡流水.xlsx","55776信用卡开户.xlsx"]}'

# 建设（PDF）单独任务
curl -s -X POST http://192.168.2.109:3000/api/task-start \
  -d '{"task":"case","customerName":"周贤_建设","selectedFiles":["建设PDF1.pdf","建设PDF2.pdf"]}'
```

### T2：OCR 完成但解析出极少交易（如 PDF 有100页却只出2条）

**原因：** 该银行 PDF 有内嵌文本层，但格式为多列交叉排版（如民生银行司法查询），文本解析器不识别。

**处理：** 绕过 N1 任务系统，直接在 Mac 上调 OCR 脚本强制 vision 模式：

```bash
RESULT_DIR="/tmp/vision_$(date +%s)"
mkdir -p "$RESULT_DIR/源数据" "$RESULT_DIR/审计中间结果"
cp /tmp/n1_results/task_xxx/审计源数据/文件.pdf "$RESULT_DIR/源数据/"

python3 /Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/scripts/run_audit_local_ocr.py \
  --input-dir "$RESULT_DIR/源数据" \
  --template "/Volumes/AI_Agent/AI_Controlled_Zone/audit_base/空白审计模板.xlsx" \
  --output-xlsx "$RESULT_DIR/审计中间结果/output.xlsx" \
  --artifacts-dir "$RESULT_DIR/审计中间结果" \
  --customer-name "姓名" \
  --mode vision

# 完成后上传 xlsx 并重触发报告
sshpass -p '139319' scp "$RESULT_DIR/审计中间结果/output.xlsx" root@192.168.2.109:/tmp/temp.xlsx
sshpass -p '139319' ssh root@192.168.2.109 "cp /tmp/temp.xlsx '/mnt/smb_client/XXX_审计附表.xlsx'"
curl -s -X POST http://192.168.2.109:3000/api/task-start \
  -d '{"task":"case","customerName":"姓名","selectedFiles":["XXX_审计附表.xlsx"]}'
```

### T3：验收有差额（净额对不上）

**两种情况：**
1. **净额对上但收入支出各高一个相同数** → 冲正交易问题（正常）
   - 搜摘要含"冲正/红冲/撤销"的交易，金额等于差额即可确认
   - 在备注列打 `[冲正]` 标签便于筛除
2. **净额对不上** → 解析有错，回查列位置/页边界

### T4：Vision 全返回 0 条

逐一排查：

```bash
# 1. 看是 JSON 还是 error 文件
ls ~/tmp/audit-local-ocr/*/vision/*.error.txt 2>/dev/null

# 2. 常见错误及对应修复
cat ~/tmp/audit-local-ocr/*/vision/*.error.txt
# "No module named 'openai'" → /opt/homebrew/bin/pip3 install openai --break-system-packages
# "extract_json_object" NameError → 确认 BUGS_FIXED.md #3 已修复
```

### T5：文件上传到 SMB 遇到中文路径问题

```bash
# 用中转方式
sshpass -p '139319' scp 文件.pdf root@192.168.2.109:/tmp/temp.pdf
sshpass -p '139319' ssh root@192.168.2.109 "cp /tmp/temp.pdf '/mnt/smb_client/正式文件名.pdf'"
```

### T6：Excel 文件被锁（覆盖时报 Permission Denied）

Windows Excel 打开着该文件。关闭 Excel 再操作。

### T7：手动触发多银行汇总（案件汇总面板按钮暂无后端）

```bash
OUTDIR="/mnt/smb_client/AI案件汇总_${姓名}_$(date +%Y%m%d_%H%M%S)"
sshpass -p '139319' ssh root@192.168.2.109 "
  mkdir -p '$OUTDIR/01_广发' '$OUTDIR/02_建设' '$OUTDIR/03_民生' '$OUTDIR/04_总台账底稿'
  cp /mnt/smb_client/AI审计结果_XXX1/* '$OUTDIR/01_广发/'
  cp /mnt/smb_client/AI审计结果_XXX2/* '$OUTDIR/02_建设/'
  cp /mnt/smb_client/AI审计结果_XXX3/* '$OUTDIR/03_民生/'
"
# 然后本地 Python 合并 xlsx → 总台账底稿.csv
# 生成 00_汇总分析.md + .docx
```

---

## 四、已知限制（当前版本）

| 限制 | 影响 | 计划修复 |
|------|------|---------|
| 民生银行多列格式自动降级未实现 | 文本层PDF只出2条，需手动触发vision | Bug #7，待排期 |
| `task-start` API 不支持 `auditMode` 参数 | 无法通过面板指定 vision 模式 | Bug #6，待排期 |
| 案件汇总面板按钮无后端 | 多银行汇总只能命令行操作 | 待实装 `/api/case-summary-*` |
| 民生司法查询数据质量 | 日期/对手字段有误差，需人工复核关键交易 | 格式解析器优化 |

> **注：** 当前系统默认按案件视角处理材料；若同批材料存在明显无关主体或跨案件混入，建议先在材料整理阶段拆分，避免结果混淆。

---

## 五、关键路径速查

| 资源 | 路径 |
|------|------|
| **主 PLAYBOOK** | `/Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/PLAYBOOK_司法审计.md` |
| **银行格式手册** | `BANK_FORMATS.md`（同目录）|
| **Bug 记录** | `BUGS_FIXED.md`（同目录）|
| N1 面板（局域网） | `http://192.168.2.109:3000/panel` |
| N1 面板（公网） | `https://n2.jzkjbj.cn/panel` |
| N1 SSH | `sshpass -p '139319' ssh -o StrictHostKeyChecking=no root@192.168.2.109` |
| 审计日志 | N1: `/var/log/n1-audit-task.log` |
| OCR 脚本 | `/Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/scripts/run_audit_local_ocr.py` |
| 批量分类脚本 | `.../scripts/batch_classify_audit.py` |
| 报告模板 | `.../templates/generate_audit_report.py` + `generate_suspect_analysis.py` |
| Excel 空模板 | `/Volumes/AI_Agent/AI_Controlled_Zone/audit_base/空白审计模板.xlsx` |
| OCR 临时目录 | `~/tmp/audit-local-ocr/`（不能用 /tmp，macOS TCC 限制）|
| Vision 配置 | `~/.openclaw/openclaw.json → models.providers.aliyun` |

---

## 六、历史案例索引

| 案件 | 银行 | 条数 | 关键特征 |
|------|------|------|---------|
| 沙暴文化（2026-05） | 农商+农行+建行 | 7003+2818+1827 | 嫌疑人周贤，净转出¥1761万，本地LLM首次实战 |
| 周贤（2026-05-29） | 广发+建设+民生 | 14789+142+467 | 三银行司法查询，修复7个OCR系统bug |

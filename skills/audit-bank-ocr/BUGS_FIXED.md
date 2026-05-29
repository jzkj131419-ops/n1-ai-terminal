# BUGS_FIXED · OCR 审计系统已修复缺陷记录

> **作用：** 防止回归，理解根因，未来升级时优先回归测试这些场景。  
> **最后更新：** 2026-05-29

---

## Bug #1 · tesseract `-c` 参数顺序错误

**发现时间：** 2026-05-29  
**影响：** 所有 PDF 本地 OCR 失败，报 `read_params_file: Can't open -c`

**根因：**  
`run_audit_local_ocr.py` 中 tesseract 调用顺序错误，`-c` 放在了输出格式 `tsv` 之后：
```python
# 错误
["tesseract", img, "stdout", "-l", langs, "--psm", psm, "tsv", "-c", "preserve_interword_spaces=1"]

# 正确（-c 必须在 tsv 之前）
["tesseract", img, "stdout", "-l", langs, "--psm", psm, "-c", "preserve_interword_spaces=1", "tsv"]
```

tesseract 把 `tsv` 之后的所有参数都当成配置文件名来读，导致 `Can't open -c`。

**修复文件：**  
`/Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/scripts/run_audit_local_ocr.py` line ~600

**回归测试：** OCR 任意 PDF，确认 `*_ocr_lines.txt` 正常生成，无 `read_params_file` 报错。

---

## Bug #2 · macOS TCC 沙盒阻止 tesseract 读取 `/tmp/` 文件

**发现时间：** 2026-05-29  
**影响：** 所有 OCR 图像处理失败，报 `image file not found: /tmp/...`（文件存在但读不到）

**根因：**  
macOS TCC（隐私保护）限制部分应用（tesseract）读取 `/tmp/claude-xxx/` 下的文件。  
原代码：
```python
with tempfile.TemporaryDirectory(prefix="audit-local-ocr-") as temp_dir:
    # 默认在 /tmp/ 下创建，tesseract 读不到
```

**修复：**  
把临时目录改到 `~/tmp/` 下：
```python
_tess_tmp = Path.home() / "tmp" / "audit-local-ocr"
_tess_tmp.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory(prefix="audit-local-ocr-", dir=_tess_tmp) as temp_dir:
```

**修复文件：** `run_audit_local_ocr.py` line ~1295

**回归测试：** 确认 `~/tmp/audit-local-ocr/` 目录有生成，OCR 正常产出结果。

---

## Bug #3 · `extract_json_object` 函数缺失导致 Vision 结果全部被吞

**发现时间：** 2026-05-29  
**影响：** Vision API 调用成功，模型返回了正确的交易 JSON，但全部被解析为 0 条

**根因：**  
`_call_vision_single()` 调用了 `extract_json_object(content)` 但该函数从未定义，也没有 import。  
Python 的 `except Exception: data = {}` 默默吞掉了 `NameError`，返回空 dict，导致所有 Vision 结果归零。

```python
try:
    data = extract_json_object(content)  # NameError 在这里，被吞掉了
except Exception:
    data = {}                            # 永远走这里
return data.get("transactions") or []   # 永远返回 []
```

**修复：** 在 `_call_vision_single` 上方添加函数定义：
```python
def extract_json_object(text: str) -> dict:
    start = text.find("{")
    end = text.rfind("}") + 1
    if start == -1 or end == 0:
        raise ValueError("no JSON object found")
    return json.loads(text[start:end])
```

**修复文件：** `run_audit_local_ocr.py` line ~741

**回归测试：** 跑一个 PDF vision 任务，确认 `vision/*.json` 的 `transactions_top/bot` > 0。

---

## Bug #4 · Vision 图像压缩上限 1800px 导致密集表格识别率为 0

**发现时间：** 2026-05-29  
**影响：** 建设银行 PDF 第一页 Vision 返回 0 条，手动直接发原图则正常

**根因：**  
`_call_vision_single()` 将图像压缩到最大 1800px，再竖向对半切割：
- 建设银行 PDF 原图：3434×2480px（横向）
- 压缩后：1800×1303px
- 对半切后每半：1800×651px → 密集表格在这个尺寸下 qwen-vl-max 无法识别

**修复：** 将压缩上限从 1800 改为 4000：
```python
# 修复前
if max(img.size) > 1800:
    img.thumbnail((1800, 1800))

# 修复后
if max(img.size) > 4000:
    img.thumbnail((4000, 4000))
```

**修复文件：** `run_audit_local_ocr.py` line ~745

**回归测试：** 建设银行 PDF 第一页 Vision 返回条数 > 10。

---

## Bug #5 · `openai` 模块未安装在 homebrew python3

**发现时间：** 2026-05-29  
**影响：** Vision 调用报 `No module named 'openai'`，所有 Vision 失败

**根因：**  
`run_audit.rewired.sh` 使用 `/opt/homebrew/bin/python3`（MAC_PATH 中的 Python），  
但 `openai` 包只安装在 miniforge conda 环境中。两套 Python 互不共享包。

**修复：**
```bash
/opt/homebrew/bin/pip3 install openai --break-system-packages
```

**回归测试：** `/opt/homebrew/bin/python3 -c "import openai; print(openai.__version__)"` 无报错。

**注意：** macOS 系统升级或 Homebrew 大版本更新后可能丢失，需重新安装。

---

## Bug #6 · `auditMode` 参数未在 `task-start` API 端点解析

**发现时间：** 2026-05-29  
**影响：** 传入 `auditMode: "vision"` 无效，任务始终以 hybrid 模式运行

**根因：**  
`/api/task-start` 的 body 解构中没有 `auditMode` 字段（只有 `task-upload-start` 有）：
```javascript
// task-start（缺少 auditMode）
const { filePath, fileName, fileSize, task, condition, outputDir, 
        selectedFiles, caseId, caseContext, customerName, relatedSubjects } = JSON.parse(body);

// task-upload-start（有 auditMode）
const { files, task, condition, caseId, caseContext, auditMode } = JSON.parse(body);
```

**当前状态：** 未修复（绕过方案：直接调 OCR 脚本，见 SKILL.md §3.2）  
**建议修复：** 在 `server.js` task-start 解构中加入 `auditMode`，并传给 `taskManager.createTask`。

---

## Bug #7 · 民生银行文本层 PDF 解析仅得 2 条交易

**发现时间：** 2026-05-29  
**影响：** 16635 行文本，应有 4000+ 条交易，实际只解析出 2 条

**根因：**  
民生银行司法查询 PDF 的文本层格式为"多列交叉排版"：每条交易的字段分散在 3-5 行，且列顺序不符合标准流水格式。文本解析器基于标准单行/双行格式设计，无法识别该排版。

系统检测到 PDF 有文本层后，走文本提取路径，不会自动降级到图像渲染+Vision。

**当前状态：** 未修复（绕过方案：强制 `--mode vision`）  
**建议修复：**
1. 文本提取后检查"每页平均提取交易数"，若远低于预期页数×5，自动降级到渲染+Vision
2. 或针对民生银行特殊格式开发专属文本解析器

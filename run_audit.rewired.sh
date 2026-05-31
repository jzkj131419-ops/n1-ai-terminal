#!/bin/bash
set -u

TASK_ID="${1:-}"
RESULT_DIR_ARG="${2:-}"
SELECTED_FILES_JSON="${3:-}"
SCRIPT_DIR="/Users/lin/Documents/Codex/2026-05-25/base-lin-mac-mini-memory-sed"
# Mac SSH 地址自动识别：
# - 调用方是 Mac（callerIp 非 SMB 宿主）→ 直接 SSH 回调用方
# - 调用方是 Windows/其他（callerIp == SMB 宿主或为空）→ 用 macSshHost 配置
_SMB_HOST=$(python3 -c "import json; c=json.load(open('/opt/ai001/smb_config.json')); print(c.get('host',''))" 2>/dev/null || true)
_CFG_MAC=$(python3 -c "import json; c=json.load(open('/opt/ai001/smb_config.json')); print(c.get('macSshHost',''))" 2>/dev/null || true)
# 回环地址（反代场景）和 SMB 宿主都走 macSshHost 配置
_IS_LOOPBACK=false
case "${N1_CALLER_IP:-}" in 127.0.0.1|::1|localhost|"") _IS_LOOPBACK=true ;; esac
if [ "${_IS_LOOPBACK}" = "false" ] && [ -n "${N1_CALLER_IP:-}" ] && [ "${N1_CALLER_IP}" != "${_SMB_HOST}" ]; then
  MAC="lin@${N1_CALLER_IP}"
elif [ -n "$_CFG_MAC" ]; then
  MAC="$_CFG_MAC"
else
  MAC="lin@100.89.188.128"
fi
MAC_PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
OCR_SCRIPT="/Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/scripts/run_audit_local_ocr.py"
REPORT_SCRIPT="$SCRIPT_DIR/generate_audit_report.dynamic.py"
BUNDLE_SCRIPT="/Users/lin/Documents/Codex/2026-05-25/base-lin-mac-mini-memory-sed/generate_intake_bundle.py"
CSV_AUDIT_SCRIPT="$SCRIPT_DIR/generate_audit_from_bank_csv.py"
BASE="/Volumes/AI_Agent/AI_Controlled_Zone/audit_base"
TEMPLATE="$BASE/空白审计模板.xlsx"
LOG="/var/log/n1-audit-task.log"
CONFIG="/opt/ai001/smb_config.json"
META_DIR="/tmp/n1_task_meta"
META_FILE="$META_DIR/$TASK_ID.json"
MATERIALS_TSV="$META_DIR/${TASK_ID}_materials.tsv"
TASK_TYPE="${N1_TASK_TYPE:-case}"
ANALYSIS_CONDITION="${N1_ANALYSIS_CONDITION:-}"
LOCAL_UPLOAD_DIR="${N1_LOCAL_UPLOAD_DIR:-}"

CASE_NAME="${N1_CASE_NAME:-审计案件}"
CUSTOMER_NAME="${N1_CUSTOMER_NAME:-$CASE_NAME}"
CARD_NO="${N1_CARD_NO:-}"
BANK_NAME="${N1_BANK_NAME:-}"
ACCOUNT_NO="${N1_ACCOUNT_NO:-}"
OCR_MODE="${N1_AUDIT_MODE:-hybrid}"
echo "[audit] OCR_MODE=$OCR_MODE" >> "$LOG"
REPORT_MODE="${N1_REPORT_MODE:-single_account}"
ACCOUNT_SCOPE="${N1_ACCOUNT_SCOPE:-}"
BANK_SCOPE="${N1_BANK_SCOPE:-}"
REPORT_BASENAME="${N1_REPORT_BASENAME:-$CUSTOMER_NAME}"
REPORT_BASENAME=$(printf '%s' "$REPORT_BASENAME" | tr '/\\:*?"<>|' '_' | tr -d '\r' | sed 's/[[:space:]]//g')
[ -z "$REPORT_BASENAME" ] && REPORT_BASENAME="审计案件"

is_placeholder_name() {
  case "$1" in
    ""|"审计案件"|"当前任务"|"单案件处理"|"未填写客户名称"|"未命名客户")
      return 0
      ;;
    *)
      return 1
      ;;
  esac
}

sanitize_report_basename() {
  local raw="$1"
  local cleaned
  cleaned=$(printf '%s' "$raw" | tr '/\\:*?"<>|' '_' | tr -d '\r' | sed 's/[[:space:]]//g')
  [ -z "$cleaned" ] && cleaned="审计案件"
  printf '%s' "$cleaned"
}

infer_bank_from_name() {
  python3 - <<'PY' "$1"
import re
import sys

name = sys.argv[1]
m = re.search(r'([一-龥A-Za-z0-9]{2,20}银行)', name)
print(m.group(1) if m else "")
PY
}

sanitize_field() {
  printf '%s' "$1" | tr '\t\r\n' '   '
}

append_material_record() {
  local status="$1"
  local category="$2"
  local original_path="$3"
  local detail="${4:-}"
  printf '%s\t%s\t%s\t%s\n' \
    "$(sanitize_field "$status")" \
    "$(sanitize_field "$category")" \
    "$(sanitize_field "$original_path")" \
    "$(sanitize_field "$detail")" >> "$MATERIALS_TSV"
}

write_meta() {
  python3 - <<'PY' "$META_FILE" "$1" "$2" "$3" "$4" "$5" "$6"
import json
import sys

meta_file, outdir, copied, task_type, status, reason, files_json = sys.argv[1:8]
payload = {
    "outDir": outdir,
    "files": json.loads(files_json),
    "copiedCount": int(copied or 0),
    "taskType": task_type,
    "status": status,
}
if reason:
    payload["failureReason"] = reason
with open(meta_file, "w", encoding="utf-8") as fh:
    json.dump(payload, fh, ensure_ascii=False)
PY
}

create_failure_artifacts() {
  local reason="$1"
  local failure_dir="$SMB_MOUNT/AI审计失败_$(date +%Y%m%d_%H%M%S)"
  local failure_md="失败说明.md"
  local failure_txt="未处理原因.txt"
  local failure_xlsx="材料识别清单.xlsx"
  mkdir -p "$failure_dir"

  python3 - <<'PY' "$failure_dir/$failure_md" "$failure_dir/$failure_txt" "$failure_dir/$failure_xlsx" "$MATERIALS_TSV" "$reason" "$CUSTOMER_NAME" "$TASK_TYPE"
from pathlib import Path
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font
import sys

md_path = Path(sys.argv[1])
txt_path = Path(sys.argv[2])
xlsx_path = Path(sys.argv[3])
tsv_path = Path(sys.argv[4])
reason = sys.argv[5]
customer = sys.argv[6]
task_type = sys.argv[7]
rows = []
if tsv_path.exists():
    rows = [line.rstrip("\n").split("\t") for line in tsv_path.read_text(encoding="utf-8", errors="ignore").splitlines()[1:] if line.strip()]
accepted = sum(1 for row in rows if row and row[0] == "accepted")
skipped = sum(1 for row in rows if row and row[0] == "skipped")
failed = sum(1 for row in rows if row and row[0] == "failed")
md_path.write_text(
    "\n".join([
        "# 失败说明",
        "",
        f"- 客户名称：{customer or '未命名客户'}",
        f"- 任务类型：{task_type}",
        f"- 失败原因：{reason}",
        f"- 生成日期：{date.today().isoformat()}",
        "",
        "## 材料处理概况",
        f"- 已接纳材料：{accepted}",
        f"- 已跳过材料：{skipped}",
        f"- 处理失败材料：{failed}",
        "",
        "## 建议",
        "- 优先提供标准银行 CSV、银行原始导出 Excel 或完整银行流水 PDF。",
        "- 若材料来自压缩包，请先确认压缩包未加密、未分卷，或先在本地解压后再上传。",
        "- 若目录中混有历史结果文件，请清理后重试。",
    ]),
    encoding="utf-8",
)
txt_path.write_text(reason + "\n", encoding="utf-8")

wb = Workbook()
ws = wb.active
ws.title = "材料识别清单"
headers = ["状态", "分类", "原始路径", "说明"]
for idx, header in enumerate(headers, start=1):
    cell = ws.cell(1, idx, header)
    cell.font = Font(bold=True)
for row_idx, row in enumerate(rows, start=2):
    values = (row + ["", "", "", ""])[:4]
    for col_idx, value in enumerate(values, start=1):
        ws.cell(row_idx, col_idx, value)
for col in ("A", "B", "C", "D"):
    ws.column_dimensions[col].width = 28 if col != "C" else 80
wb.save(xlsx_path)
PY

  write_meta "$failure_dir" "$COPIED" "$TASK_TYPE" "failed" "$reason" "[\"$failure_md\",\"$failure_txt\",\"$failure_xlsx\"]"
  echo "[$(date)] 失败说明已写回 $failure_dir" >> "$LOG"
}

fail_task() {
  local reason="$1"
  echo "[$(date)] 错误：$reason" >> "$LOG"
  create_failure_artifacts "$reason"
  exit 1
}

is_supported_input() {
  local lower
  lower=$(printf '%s' "$1" | tr '[:upper:]' '[:lower:]')
  case "$lower" in
    *.pdf|*.xlsx|*.csv|*.jpg|*.jpeg|*.png|*.tif|*.tiff|*.bmp)
      return 0
      ;;
    *)
      return 1
      ;;
  esac
}

is_supported_archive() {
  local lower
  lower=$(printf '%s' "$1" | tr '[:upper:]' '[:lower:]')
  case "$lower" in
    *.zip)
      return 0
      ;;
    *)
      return 1
      ;;
  esac
}

make_stage_target() {
  python3 - <<'PY' "$INPUT_STAGE_DIR" "$1"
from pathlib import Path
import os
import re
import sys

stage_dir = Path(sys.argv[1])
raw_name = (sys.argv[2] or "input").replace("\\", "/").strip("/")
safe_name = raw_name.replace("/", "__")
safe_name = re.sub(r'[<>:"|?*]+', "_", safe_name)
safe_name = re.sub(r"\s+", "_", safe_name).strip("._")
if not safe_name:
    safe_name = "input"
base, ext = os.path.splitext(safe_name)
candidate = stage_dir / safe_name
index = 2
while candidate.exists():
    candidate = stage_dir / f"{base}_{index}{ext}"
    index += 1
print(candidate)
PY
}

stage_supported_file() {
  local source_file="$1"
  local label="$2"
  local stage_target

  if ! is_supported_input "$source_file"; then
    return 1
  fi

  stage_target=$(make_stage_target "$label")
  cp "$source_file" "$stage_target"
  STAGED_COUNT=$((STAGED_COUNT + 1))
  append_material_record "accepted" "input" "$label" "$(basename "$stage_target")"
  return 0
}

stage_zip_archive() {
  local archive_file="$1"
  local label="$2"
  local extract_dir
  local extracted_any=0

  extract_dir=$(mktemp -d "/tmp/n1_zip_${TASK_ID}_XXXXXX")
  local unzip_output
  if ! unzip_output=$(python3 - <<'PY' "$archive_file" "$extract_dir" 2>&1
from pathlib import Path
from zipfile import ZipFile
import sys

archive = Path(sys.argv[1])
target = Path(sys.argv[2])
with ZipFile(archive) as zf:
    zf.extractall(target)
PY
); then
    echo "[$(date)] 警告：压缩包解压失败：$archive_file" >> "$LOG"
    append_material_record "failed" "archive" "$label" "$unzip_output"
    rm -rf "$extract_dir"
    return 1
  fi

  while IFS= read -r -d '' extracted_file; do
    local relative_inside
    relative_inside=$(python3 - <<'PY' "$extract_dir" "$extracted_file"
from pathlib import Path
import sys

root = Path(sys.argv[1])
current = Path(sys.argv[2])
print(current.relative_to(root).as_posix())
PY
)
    if stage_supported_file "$extracted_file" "${label%.*}/$relative_inside"; then
      extracted_any=1
    fi
  done < <(find "$extract_dir" -type f -print0)

  rm -rf "$extract_dir"

  if [ "$extracted_any" -eq 0 ]; then
    echo "[$(date)] 警告：压缩包内未找到可处理文件：$archive_file" >> "$LOG"
    append_material_record "skipped" "archive" "$label" "压缩包内未找到可处理文件"
  else
    append_material_record "accepted" "archive" "$label" "已解压并纳入处理"
  fi
}

stage_input_candidate() {
  local source_file="$1"
  local label="${2:-$(basename "$source_file")}"

  if [ ! -f "$source_file" ]; then
    echo "[$(date)] 警告：输入文件不存在：$source_file" >> "$LOG"
    append_material_record "failed" "missing" "$label" "输入文件不存在"
    return 1
  fi

  if is_supported_input "$source_file"; then
    if ! stage_supported_file "$source_file" "$label"; then
      echo "[$(date)] 警告：输入文件整理失败：$source_file" >> "$LOG"
      return 1
    fi
    return 0
  fi

  if is_supported_archive "$source_file"; then
    stage_zip_archive "$source_file" "$label"
    return 0
  fi

  echo "[$(date)] 跳过不支持的文件：$source_file" >> "$LOG"
  append_material_record "skipped" "unsupported" "$label" "文件类型不支持"
  return 0
}

XLSX_NAME="${REPORT_BASENAME}_审计附表.xlsx"
DOCX_NAME="${REPORT_BASENAME}_审计报告.docx"
MD_NAME="${REPORT_BASENAME}_审计报告.md"
INITIAL_XLSX_NAME="$XLSX_NAME"
INITIAL_DOCX_NAME="$DOCX_NAME"
INITIAL_MD_NAME="$MD_NAME"
BUNDLE_MD_NAME="${REPORT_BASENAME}_票据材料清单.md"
BUNDLE_XLSX_NAME="${REPORT_BASENAME}_票据台账.xlsx"

echo "[$(date)] TASK=$TASK_ID 开始" >> "$LOG"
mkdir -p "$META_DIR"
printf '状态\t分类\t原始路径\t说明\n' > "$MATERIALS_TSV"

if [ -z "$TASK_ID" ]; then
  echo "[$(date)] 错误：缺少 TASK_ID" >> "$LOG"
  exit 1
fi

# 读挂载点
if [ ! -f "$CONFIG" ]; then
  echo "[$(date)] 错误：未找到 smb_config.json" >> "$LOG"
  exit 1
fi
SMB_MOUNT=$(python3 -c "import json; print(json.load(open('$CONFIG'))['mountPoint'])" 2>/dev/null || echo "/mnt/smb_client")
echo "[$(date)] 挂载点: $SMB_MOUNT" >> "$LOG"

# Mac 侧临时目录
RESULT_DIR="${RESULT_DIR_ARG:-/tmp/n1_results/$TASK_ID}"
SOURCE_DIR="$RESULT_DIR/审计源数据"
ARTIFACTS_DIR="$RESULT_DIR/审计中间结果"
INPUT_STAGE_DIR="/tmp/n1_input_stage/$TASK_ID"
ssh -o StrictHostKeyChecking=no "$MAC" "mkdir -p '$SOURCE_DIR' '$ARTIFACTS_DIR'" >> "$LOG" 2>&1
rm -rf "$INPUT_STAGE_DIR"
mkdir -p "$INPUT_STAGE_DIR"
trap 'rm -rf "$INPUT_STAGE_DIR"' EXIT

# 把本次选中的输入文件整理到 N1 暂存目录：
# 1. 递归支持 SMB 子文件夹
# 2. 自动解压 zip
# 3. 统一平铺后再送到 Mac，尽量不改后续主链
STAGED_COUNT=0
if [ -n "$LOCAL_UPLOAD_DIR" ] && [ -d "$LOCAL_UPLOAD_DIR" ]; then
  while IFS= read -r -d '' FILE; do
    REL_NAME=$(python3 - <<'PY' "$LOCAL_UPLOAD_DIR" "$FILE"
from pathlib import Path
import sys

root = Path(sys.argv[1])
current = Path(sys.argv[2])
print(current.relative_to(root).as_posix())
PY
)
    stage_input_candidate "$FILE" "$REL_NAME"
  done < <(
    find "$LOCAL_UPLOAD_DIR" -type f -print0
  )
elif [ -n "$SELECTED_FILES_JSON" ] && [ "$SELECTED_FILES_JSON" != "[]" ]; then
  while IFS= read -r FILE_NAME; do
    [ -z "$FILE_NAME" ] && continue
    FILE_PATH="$SMB_MOUNT/$FILE_NAME"
    stage_input_candidate "$FILE_PATH" "$FILE_NAME"
  done < <(python3 -c 'import json,sys; [print(x) for x in json.loads(sys.argv[1])] ' "$SELECTED_FILES_JSON")
else
  while IFS= read -r -d '' FILE; do
    REL_NAME=$(python3 - <<'PY' "$SMB_MOUNT" "$FILE"
from pathlib import Path
import sys

root = Path(sys.argv[1])
current = Path(sys.argv[2])
print(current.relative_to(root).as_posix())
PY
)
    stage_input_candidate "$FILE" "$REL_NAME"
  done < <(
    find "$SMB_MOUNT" -type f -print0
  )
fi
echo "[$(date)] 已整理输入文件 $STAGED_COUNT 个" >> "$LOG"

COPIED=0
while IFS= read -r -d '' FILE; do
  scp -o StrictHostKeyChecking=no "$FILE" "$MAC:$SOURCE_DIR/" >> "$LOG" 2>&1
  if [ $? -eq 0 ]; then
    COPIED=$((COPIED + 1))
  else
    echo "[$(date)] 警告：输入文件传输失败：$FILE" >> "$LOG"
  fi
done < <(find "$INPUT_STAGE_DIR" -maxdepth 1 -type f -print0)
echo "[$(date)] 已传输输入文件 $COPIED 个" >> "$LOG"

if [ "$COPIED" -eq 0 ]; then
  fail_task "客户共享目录未找到可处理输入文件（含子文件夹与 zip）"
fi

if [ "$TASK_TYPE" = "invoice" ]; then
  DETECTED_TASK_TYPE=$(ssh -o StrictHostKeyChecking=no "$MAC" \
    "PATH='$MAC_PATH' python3 '$BUNDLE_SCRIPT' \
      --input-dir '$SOURCE_DIR' \
      --output-dir '$RESULT_DIR' \
      --task-type '$TASK_TYPE' \
      --case-name '$CASE_NAME' \
      --customer-name '$CUSTOMER_NAME' \
      --analysis-condition '$ANALYSIS_CONDITION' \
      --report-base-name '$REPORT_BASENAME' \
      --ocr-mode '$OCR_MODE' \
      --detect-only" 2>>"$LOG" | tail -n 1 | tr -d '\r')
  if [ "$DETECTED_TASK_TYPE" = "case" ]; then
    echo "[$(date)] 智能分流：检测到材料更像审计流水，改走审计案件主链" >> "$LOG"
    TASK_TYPE="case"
  else
    echo "[$(date)] 智能分流：保留票据台账链（detect=${DETECTED_TASK_TYPE:-unknown}）" >> "$LOG"
  fi
fi

if [ "$TASK_TYPE" = "invoice" ]; then
  echo "[$(date)] 进入票据台账独立流程" >> "$LOG"
  ssh -o StrictHostKeyChecking=no "$MAC" \
    "PATH='$MAC_PATH' python3 '$BUNDLE_SCRIPT' \
      --input-dir '$SOURCE_DIR' \
      --output-dir '$RESULT_DIR' \
      --task-type 'invoice' \
      --case-name '$CASE_NAME' \
      --customer-name '$CUSTOMER_NAME' \
      --analysis-condition '$ANALYSIS_CONDITION' \
      --report-base-name '$REPORT_BASENAME' \
      --ocr-mode '$OCR_MODE'" >> "$LOG" 2>&1
  CODE=$?
  echo "[$(date)] 票据台账脚本退出码 $CODE" >> "$LOG"
  [ "$CODE" -ne 0 ] && fail_task "票据台账脚本执行失败"

  OUTDIR="$SMB_MOUNT/AI票据结果_$(date +%Y%m%d_%H%M%S)"
  mkdir -p "$OUTDIR"
  scp -o StrictHostKeyChecking=no \
    "$MAC:$RESULT_DIR/$BUNDLE_XLSX_NAME" \
    "$MAC:$RESULT_DIR/$BUNDLE_MD_NAME" \
    "$OUTDIR/" >> "$LOG" 2>&1
  CODE=$?
  echo "[$(date)] 票据结果写回退出码 $CODE" >> "$LOG"
  [ "$CODE" -ne 0 ] && fail_task "票据结果写回共享目录失败"

  write_meta "$OUTDIR" "$COPIED" "invoice" "done" "" "[\"$BUNDLE_XLSX_NAME\",\"$BUNDLE_MD_NAME\"]"

  echo "[$(date)] 完成，票据结果已写回 $OUTDIR" >> "$LOG"
  exit 0
fi

# 优先走结构化材料（客户附表 / 银行 CSV / 原始银行 XLSX），只有缺少结构化流水时才回退 OCR。
ssh -o StrictHostKeyChecking=no "$MAC" "
  set -e
  export PATH='$MAC_PATH'
  PROVIDED_XLSX=\$(find '$SOURCE_DIR' -maxdepth 1 -type f -name '*审计附表*.xlsx' | head -n 1)
  STRUCTURED_FILE=\$(find '$SOURCE_DIR' -maxdepth 1 -type f \\( -name '*.csv' -o -name '*.xlsx' \\) ! -name '~$*' | head -n 1)
  if [ -n \"\$PROVIDED_XLSX\" ]; then
    cp \"\$PROVIDED_XLSX\" '$RESULT_DIR/$XLSX_NAME'
    echo '使用客户提供的审计附表'
  elif [ -n \"\$STRUCTURED_FILE\" ]; then
    python3 '$CSV_AUDIT_SCRIPT' \
      --input-dir '$SOURCE_DIR' \
      --template '$TEMPLATE' \
      --output-xlsx '$RESULT_DIR/$XLSX_NAME' \
      --customer-name '$CUSTOMER_NAME' \
      --card-no '$CARD_NO' \
      --account-no '$ACCOUNT_NO'
  else
    python3 '$OCR_SCRIPT' \
      --input-dir '$SOURCE_DIR' \
      --template '$TEMPLATE' \
      --output-xlsx '$RESULT_DIR/$XLSX_NAME' \
      --artifacts-dir '$ARTIFACTS_DIR' \
      --customer-name '$CUSTOMER_NAME' \
      --card-no '$CARD_NO' \
      --account-no '$ACCOUNT_NO' \
      --mode '$OCR_MODE'
  fi
" >> "$LOG" 2>&1
CODE=$?
echo "[$(date)] 附表准备退出码 $CODE" >> "$LOG"
[ "$CODE" -ne 0 ] && fail_task "审计附表准备失败，请检查输入材料格式或压缩包是否可正常解压"

PRIMARY_CUSTOMER=$(python3 - <<'PY' "$LOG"
from pathlib import Path
import re
import sys

text = Path(sys.argv[1]).read_text(encoding='utf-8', errors='ignore')
matches = re.findall(r'PRIMARY_CUSTOMER::([^\r\n]+)', text)
print(matches[-1].strip() if matches else "")
PY
)

PRIMARY_ACCOUNT=$(python3 - <<'PY' "$LOG"
from pathlib import Path
import re
import sys

text = Path(sys.argv[1]).read_text(encoding='utf-8', errors='ignore')
matches = re.findall(r'PRIMARY_ACCOUNT::([^\r\n]+)', text)
print(matches[-1].strip() if matches else "")
PY
)

REPORT_MODE_FROM_LOG=$(python3 - <<'PY' "$LOG"
from pathlib import Path
import re
import sys

text = Path(sys.argv[1]).read_text(encoding='utf-8', errors='ignore')
matches = re.findall(r'REPORT_MODE::([^\r\n]+)', text)
print(matches[-1].strip() if matches else "")
PY
)

SUBJECT_SCOPE_FROM_LOG=$(python3 - <<'PY' "$LOG"
from pathlib import Path
import re
import sys

text = Path(sys.argv[1]).read_text(encoding='utf-8', errors='ignore')
matches = re.findall(r'CUSTOMERS::([^\r\n]+)', text)
print(matches[-1].strip() if matches else "")
PY
)

ACCOUNT_SCOPE_FROM_LOG=$(python3 - <<'PY' "$LOG"
from pathlib import Path
import re
import sys

text = Path(sys.argv[1]).read_text(encoding='utf-8', errors='ignore')
matches = re.findall(r'ACCOUNTS::([^\r\n]+)', text)
print(matches[-1].strip() if matches else "")
PY
)

BANK_SCOPE_FROM_LOG=$(python3 - <<'PY' "$LOG"
from pathlib import Path
import re
import sys

text = Path(sys.argv[1]).read_text(encoding='utf-8', errors='ignore')
matches = re.findall(r'BANKS::([^\r\n]+)', text)
print(matches[-1].strip() if matches else "")
PY
)

if is_placeholder_name "$CUSTOMER_NAME" && ! is_placeholder_name "$PRIMARY_CUSTOMER"; then
  CUSTOMER_NAME="$PRIMARY_CUSTOMER"
fi

if [ -n "$PRIMARY_ACCOUNT" ]; then
  ACCOUNT_NO="$PRIMARY_ACCOUNT"
fi

if [ -n "$REPORT_MODE_FROM_LOG" ]; then
  REPORT_MODE="$REPORT_MODE_FROM_LOG"
fi

if [ -n "$ACCOUNT_SCOPE_FROM_LOG" ]; then
  ACCOUNT_SCOPE="$ACCOUNT_SCOPE_FROM_LOG"
fi

if [ -n "$BANK_SCOPE_FROM_LOG" ]; then
  BANK_SCOPE="$BANK_SCOPE_FROM_LOG"
fi

if [ "$REPORT_MODE" = "multi_account" ]; then
  REPORT_BASENAME=$(sanitize_report_basename "${CUSTOMER_NAME}_多账户汇总")
else
  REPORT_BASENAME=$(sanitize_report_basename "$CUSTOMER_NAME")
fi
XLSX_NAME="${REPORT_BASENAME}_审计附表.xlsx"
DOCX_NAME="${REPORT_BASENAME}_审计报告.docx"
MD_NAME="${REPORT_BASENAME}_审计报告.md"

if [ "$XLSX_NAME" != "$INITIAL_XLSX_NAME" ]; then
  ssh -o StrictHostKeyChecking=no "$MAC" "
    if [ -f '$RESULT_DIR/$INITIAL_XLSX_NAME' ]; then
      mv '$RESULT_DIR/$INITIAL_XLSX_NAME' '$RESULT_DIR/$XLSX_NAME'
    fi
  " >> "$LOG" 2>&1
fi

if [ -z "$BANK_NAME" ]; then
  BANK_NAME=$(python3 - <<'PY' "$LOG"
from pathlib import Path
import re
import sys

log_path = Path(sys.argv[1])
text = log_path.read_text(encoding='utf-8', errors='ignore')
matches = re.findall(r'BANK::([^\r\n]+)', text)
print(matches[-1].strip() if matches else "")
PY
)
fi

if [ -z "$BANK_NAME" ]; then
  CSV_NAME=$(python3 - <<'PY' "$SOURCE_DIR"
from pathlib import Path
import sys
src = Path(sys.argv[1])
csvs = sorted(src.glob("*.csv"))
print(csvs[0].name if csvs else "")
PY
)
  if [ -n "$CSV_NAME" ]; then
    INFERRED_BANK=$(infer_bank_from_name "$CSV_NAME")
    [ -n "$INFERRED_BANK" ] && BANK_NAME="$INFERRED_BANK"
  fi
fi

# 生成报告
ssh -o StrictHostKeyChecking=no "$MAC" \
  "PATH='$MAC_PATH' \
N1_OUTPUT_DIR='$RESULT_DIR' \
N1_AUDIT_XLSX='$RESULT_DIR/$XLSX_NAME' \
N1_AUDIT_DOCX='$RESULT_DIR/$DOCX_NAME' \
N1_AUDIT_MD='$RESULT_DIR/$MD_NAME' \
N1_CUSTOMER_NAME='$CUSTOMER_NAME' \
N1_ACCOUNT_NO='$ACCOUNT_NO' \
N1_BANK_NAME='$BANK_NAME' \
N1_REPORT_MODE='$REPORT_MODE' \
N1_SUBJECT_SCOPE='$SUBJECT_SCOPE_FROM_LOG' \
N1_ACCOUNT_SCOPE='$ACCOUNT_SCOPE' \
N1_BANK_SCOPE='$BANK_SCOPE' \
python3 '$REPORT_SCRIPT'" >> "$LOG" 2>&1
CODE=$?
echo "[$(date)] 报告脚本退出码 $CODE" >> "$LOG"
[ "$CODE" -ne 0 ] && fail_task "审计报告生成失败，通常是未识别到有效交易或材料结构不完整"

# 结果写回客户共享文件夹
OUTDIR="$SMB_MOUNT/AI审计结果_$(date +%Y%m%d_%H%M%S)"
mkdir -p "$OUTDIR"
scp -o StrictHostKeyChecking=no \
  "$MAC:$RESULT_DIR/$DOCX_NAME" \
  "$MAC:$RESULT_DIR/$MD_NAME" \
  "$MAC:$RESULT_DIR/$XLSX_NAME" \
  "$OUTDIR/" >> "$LOG" 2>&1
CODE=$?
echo "[$(date)] 结果写回退出码 $CODE" >> "$LOG"
[ "$CODE" -ne 0 ] && fail_task "审计结果写回共享目录失败"

write_meta "$OUTDIR" "$COPIED" "case" "done" "" "[\"$DOCX_NAME\",\"$MD_NAME\",\"$XLSX_NAME\"]"

echo "[$(date)] 完成，结果已写回 $OUTDIR" >> "$LOG"

# 银行流水漏录核查与补录流程手册

> 沉淀自：2026-06-01 周贤案民生银行漏录184条补录实操  
> 适用：任何已完成 OCR 解析的审计附表，发现 PDF 条数 > 附表条数 时触发

---

## 一、触发条件

**触发规则：PDF 解析条数 vs 附表 `1交易记录汇总` 非空行数，差值 > 0 即触发。**

```python
# 快速检测命令
python3 - <<'EOF'
import openpyxl, subprocess, re

xlsx = "/path/to/周贤_民生_审计附表.xlsx"
wb = openpyxl.load_workbook(xlsx)
ws = wb["1交易记录汇总"]
attached_count = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r[0])

# PDF 条数从 run_audit_local_ocr.py 输出中取（或手动查 PDF 封面合计行数）
pdf_count = 4712  # 填入 PDF 封面标注的交易条数

diff = pdf_count - attached_count
print(f"附表: {attached_count} 条 | PDF: {pdf_count} 条 | 差: {diff} 条")
if diff > 0:
    print(">>> 触发漏录核查流程")
else:
    print(">>> 条数对上，无需核查")
EOF
```

---

## 二、生成漏录核查清单

### 2.1 脚本：`generate_supplement_checklist.py`

```python
#!/usr/bin/env python3
"""
生成银行流水漏录核查清单
用法：
  python3 generate_supplement_checklist.py \
    --xlsx  <审计附表.xlsx> \
    --pdf-csv <PDF全量解析CSV> \
    --out   <输出核查清单.xlsx>

PDF全量解析CSV 格式（run_audit_local_ocr.py 的中间产物，或另外单跑提取）：
  日期,方向,金额,余额,摘要,对手方,页码
"""
import argparse, csv
from collections import defaultdict
from datetime import datetime, date
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill

RED_FILL   = PatternFill("solid", fgColor="FF9999")   # 完全漏录行（差额>0）
AMBER_FILL = PatternFill("solid", fgColor="FFD580")   # 部分差异行
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")   # 无差行（留作参考）

def to_date(v):
    if isinstance(v, (datetime, date)):
        return v.date() if isinstance(v, datetime) else v
    return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()

def run(xlsx_path, pdf_csv_path, out_path):
    # ── 读附表 ──────────────────────────────────
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb["1交易记录汇总"]
    # 按 (日期, 方向) 聚合附表金额和笔数
    attached = defaultdict(lambda: {"n": 0, "total": 0.0})
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        d = to_date(r[0])
        direction = "收入" if r[5] == "C" else "支出"
        amt = float(r[6]) if r[6] else 0.0
        key = (d, direction)
        attached[key]["n"] += 1
        attached[key]["total"] += amt

    # ── 读 PDF CSV ──────────────────────────────
    pdf_rows = []
    with open(pdf_csv_path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            pdf_rows.append(row)

    # ── 找差异行 ────────────────────────────────
    discrepancy_dates = set()
    for row in pdf_rows:
        d = datetime.strptime(row["日期"][:10], "%Y-%m-%d").date()
        direction = row["方向"]
        amt = float(row["金额"])
        key = (d, direction)
        at = attached.get(key, {"n": 0, "total": 0.0})
        diff = amt - at["total"]  # 正数=PDF多；负数=附表已多
        if abs(diff) > 0.01:
            discrepancy_dates.add((d, direction))

    # ── 生成清单（仅差异日期的PDF行）───────────
    checklist_rows = []
    for row in pdf_rows:
        d = datetime.strptime(row["日期"][:10], "%Y-%m-%d").date()
        direction = row["方向"]
        if (d, direction) not in discrepancy_dates:
            continue
        amt = float(row["金额"])
        bal = float(row["余额"]) if row.get("余额") else 0.0
        key = (d, direction)
        at = attached.get(key, {"n": 0, "total": 0.0})
        diff = amt - at["total"]  # 单行差额（简化：以本行金额 vs 附表合计）
        checklist_rows.append({
            "日期": d,
            "方向": direction,
            "PDF金额": amt,
            "PDF余额": bal,
            "PDF摘要": row.get("摘要", ""),
            "PDF对手方": row.get("对手方", ""),
            "PDF页码": row.get("页码", ""),
            "附表现有同日同方向笔数": at["n"],
            "附表同日同方向合计": at["total"],
            "差额": round(amt - at["total"], 2),
            "核查结论（待填）": None,
        })

    # ── 写 Excel ────────────────────────────────
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "漏录核查清单"

    headers = list(checklist_rows[0].keys()) if checklist_rows else []
    out_ws.append(headers)

    for item in checklist_rows:
        row_vals = [item[h] for h in headers]
        out_ws.append(row_vals)
        ri = out_ws.max_row
        diff_val = item["差额"]
        fill = RED_FILL if diff_val > 0 else (AMBER_FILL if diff_val < 0 else GREEN_FILL)
        for ci in range(1, len(headers) + 1):
            out_ws.cell(ri, ci).fill = fill

    # 说明 sheet
    info_ws = out_wb.create_sheet("说明")
    info_ws.append(["颜色", "含义", "建议操作"])
    info_ws.append(["红色", "差额>0，PDF多于附表，可能漏录", "建议填写：确认补录"])
    info_ws.append(["橙色", "差额<0，附表多于PDF", "核实后决定；通常填：不补（附表已有）"])
    info_ws.append(["绿色", "差额=0，但同日存在其他差异", "留存参考"])
    info_ws.append([])
    info_ws.append(["K列填写规则：确认补录 | 不补（已存在） | 不补（金额有误） | 不补（其他原因）"])

    out_wb.save(out_path)
    total = len(checklist_rows)
    red   = sum(1 for r in checklist_rows if r["差额"] > 0)
    amber = sum(1 for r in checklist_rows if r["差额"] < 0)
    print(f"核查清单已生成: {out_path}")
    print(f"  总行数: {total}（红{red} 橙{amber} 绿{total-red-amber}）")

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx",    required=True)
    ap.add_argument("--pdf-csv", required=True)
    ap.add_argument("--out",     required=True)
    args = ap.parse_args()
    run(args.xlsx, args.pdf_csv, args.out)
```

**调用示例（周贤民生案）：**
```bash
python3 generate_supplement_checklist.py \
  --xlsx    "/Volumes/AI审计/.../03_民生/周贤_民生_审计附表.xlsx" \
  --pdf-csv "/tmp/民生_pdf_全量.csv" \
  --out     "/Volumes/AI审计/.../03_民生/民生_漏录核查清单.xlsx"
```

### 2.2 PDF全量CSV 从哪来

`run_audit_local_ocr.py` 在 `--artifacts-dir` 下会输出 `all_transactions.csv`（每行一条交易），格式为 `日期,方向,金额,余额,摘要,对手方,页码`。直接用作 `--pdf-csv` 参数。

---

## 三、核查结论填写规则

发给客户的 Excel K 列，填写以下四种之一：

| 结论 | 含义 | 适用场景 |
|------|------|---------|
| `确认补录` | 确认该条PDF交易确实漏录，需写入附表 | 红色行（差额>0），经核实确实缺失 |
| `不补（已存在）` | 附表中已有对应条目，只是日期/金额略有偏差 | 橙色行，附表实际已覆盖 |
| `不补（金额有误）` | PDF金额与实际不符，以附表为准 | 任何行，PDF数字疑有OCR误读 |
| `不补（其他原因）` | 如冲正配对、内部转账等 | 与其他交易形成对偶抵消的条目 |

**快速决策口诀：**
- 红色行 → 默认 `确认补录`，有疑问再改
- 橙色行 → 默认 `不补（已存在）`，除非附表那笔金额明显不对

---

## 四、批量写入附表的脚本

```python
#!/usr/bin/env python3
"""
批量补录：将核查清单中"确认补录"行写入审计附表
用法：
  python3 supplement_to_xlsx.py \
    --checklist <核查清单.xlsx> \
    --xlsx      <审计附表.xlsx> \
    --ledger    <总台账底稿.csv> \
    --bank      <银行名称，如"民生银行"> \
    --card      <卡号> \
    --account   <账号> \
    --customer  <客户名>
"""
import argparse, csv, shutil
from datetime import datetime, date
from decimal import Decimal
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill

ORANGE = PatternFill("solid", fgColor="FFD580")

def to_date(v):
    if isinstance(v, (datetime, date)):
        return v.date() if isinstance(v, datetime) else v
    return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()

def run(checklist_path, xlsx_path, ledger_path, bank, card, account, customer, confirm_all=False):
    # ── 读核查清单 ──────────────────────────────
    ck_wb = openpyxl.load_workbook(checklist_path)
    ck_ws = ck_wb.active
    new_rows = []
    for r in ck_ws.iter_rows(min_row=2, values_only=True):
        dt, direction, amt, bal, summary, cp, _, _, _, _, conclusion = r[:11]
        if dt is None:
            continue
        # 筛选：confirm_all=True 时全部补录；否则只补 K列="确认补录" 的行
        if not confirm_all and (conclusion or "").strip() != "确认补录":
            continue

        tx_date = to_date(dt)
        d_flag  = "D" if direction == "支出" else "C"
        income  = float(amt) if direction == "收入" else 0.0
        expense = float(amt) if direction == "支出" else 0.0

        new_rows.append((
            tx_date, "00:00:00", customer, card, account,
            d_flag, float(amt), income, expense, float(bal or 0),
            "", cp or "", "", summary or "", "", "[漏录补录]",
        ))

    if not new_rows:
        print("没有需要补录的行（核查清单K列均未填写"确认补录"，或使用 --confirm-all）")
        return

    print(f"待补录: {len(new_rows)} 行")

    # ── 附表备份 + 追加 ──────────────────────────
    backup = str(xlsx_path).replace(".xlsx", "_pre_buhu.xlsx")
    shutil.copy2(xlsx_path, backup)
    print(f"附表备份: {backup}")

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb["1交易记录汇总"]

    old_data = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
    all_data = old_data + [tuple(r) for r in new_rows]

    def sort_key(r):
        d = r[0]
        if isinstance(d, datetime): d = d.date()
        return (d, str(r[1]) if r[1] else "00:00:00")

    all_data.sort(key=sort_key)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row: cell.value = None

    for ri, row in enumerate(all_data, 2):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(ri, ci, val)
            if row[15] == "[漏录补录]" and ci in (11, 12):
                cell.fill = ORANGE

    wb.save(xlsx_path)
    print(f"附表已保存，合计 {len(all_data)} 行")

    # ── 总台账 CSV ──────────────────────────────
    if ledger_path:
        shutil.copy2(ledger_path, str(ledger_path).replace(".csv", "_pre_buhu.csv"))
        with open(ledger_path, encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            header = next(reader)
            existing = list(reader)

        added = [[bank, r[0].strftime("%Y-%m-%d"), r[1], r[2], r[3], r[4],
                  r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13], r[14], r[15]]
                 for r in new_rows]
        all_csv = sorted(existing + added, key=lambda r: (r[0], r[1], r[2] or ""))

        with open(ledger_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerows(all_csv)
        print(f"总台账已更新: {len(all_csv)} 条（新增 {len(added)} 条）")

    # ── 补录后合计校验 ──────────────────────────
    wb2 = openpyxl.load_workbook(xlsx_path)
    ws2 = wb2["1交易记录汇总"]
    ti = te = Decimal("0")
    cnt = 0
    for r in ws2.iter_rows(min_row=2, values_only=True):
        if not r[0]: continue
        if r[7]: ti += Decimal(str(r[7]))
        if r[8]: te += Decimal(str(r[8]))
        cnt += 1
    print(f"\n=== 补录后校验 ===")
    print(f"总条数: {cnt}")
    print(f"收入: ¥{ti:,.2f}  支出: ¥{te:,.2f}  净额: ¥{ti-te:,.2f}")

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--checklist", required=True)
    ap.add_argument("--xlsx",      required=True)
    ap.add_argument("--ledger")
    ap.add_argument("--bank",      required=True)
    ap.add_argument("--card",      required=True)
    ap.add_argument("--account",   required=True)
    ap.add_argument("--customer",  required=True)
    ap.add_argument("--confirm-all", action="store_true",
                    help="忽略K列，全部补录（用于人工已电话确认全部补录的场景）")
    args = ap.parse_args()
    run(args.checklist, Path(args.xlsx), Path(args.ledger) if args.ledger else None,
        args.bank, args.card, args.account, args.customer, args.confirm_all)
```

**调用示例（周贤民生案 — 用户口头确认全部184条）：**
```bash
python3 supplement_to_xlsx.py \
  --checklist "/Volumes/AI审计/.../03_民生/民生_漏录核查清单.xlsx" \
  --xlsx      "/Volumes/AI审计/.../03_民生/周贤_民生_审计附表.xlsx" \
  --ledger    "/Volumes/AI审计/.../04_总台账底稿/总台账底稿.csv" \
  --bank      "民生银行" \
  --card      "6216910103670714" \
  --account   "50000000000414464745" \
  --customer  "周贤" \
  --confirm-all
```

> **注意**：`--confirm-all` 仅在用户已明确口头/书面确认"全部补录"时使用。  
> 正常流程应由客户填写K列后再跑，不加此参数。

---

## 五、补录后重新生成报告

### 5.1 重跑审计报告

```bash
# 复制模板，改三个常量
cp /Volumes/AI_Agent/openclaw/workspace/skills/audit-local-ocr/templates/generate_audit_report.py /tmp/gen_report.py

# 编辑头部：
#   XLSX → 附表路径
#   DOCX / MD → 输出路径
#   CUSTOMER / ACCOUNT / BANK → 对应信息
vim /tmp/gen_report.py

python3 /tmp/gen_report.py
```

**注意**：模板第 276 行日期解析需兼容带时间的字符串：
```python
# 原始（会报错）：
date.fromisoformat(str(a["last"]))
# 修正：
date.fromisoformat(str(a["last"])[:10])
```
（此 Bug 已在周贤案修复，后续使用直接改模板本体。）

### 5.2 更新 00_汇总分析

运行完报告后，手动更新 `00_汇总分析.md` 中的以下字段：

| 字段 | 来源 |
|------|------|
| 民生笔数 | 报告第一节或附表行数 |
| 民生收入/支出/净额 | 报告第二节 |
| 三账户合计 | 民生新值 + 广发/建设原值 |
| 本版说明 | 注明"补录N条" |
| 报告日期 | 今日日期 |

---

## 六、完整流程时序

```
1. PDF解析完成，附表生成
        ↓
2. 条数核查：pdf_count vs attached_count
   差值 > 0 → 进入核查流程
        ↓
3. 运行 generate_supplement_checklist.py
   → 输出 漏录核查清单.xlsx（红/橙/绿色）
        ↓
4. 发给客户/审计员填写 K 列
   （确认补录 / 不补+原因）
        ↓
5. 运行 supplement_to_xlsx.py
   → 附表条数增加，总台账同步更新
   → 打印新收支合计，与PDF标注比对
        ↓
6. 重跑 generate_audit_report.py
   → 附表分析 sheet 全部刷新
   → 生成新 docx + md 报告
        ↓
7. 手动更新 00_汇总分析.md 数字
        ↓
8. 三铁律验收：条数对 / 净额对 / 冲正定位 ✓
```

---

## 七、本次周贤案实操数据（参考）

| 指标 | 数值 |
|------|------|
| 银行 | 中国民生银行 |
| 账号 | 50000000000414464745 |
| OCR原始附表条数 | 4,660 |
| PDF条数 | 4,712 |
| 核查清单行数 | 184（差额>0红色60行，差额<0橙色104行，无差绿色20行） |
| 用户确认结论 | 全部184条确认补录 |
| 补录后附表条数 | 4,844 |
| 补录新增收入 | ¥1,924,080.08 |
| 补录新增支出 | ¥1,329,569.34 |
| 补录后民生净额 | ¥-780,864.10 |
| 三账户补录后净额 | ¥-794,153.71 |
| 附表备份 | 周贤_民生_审计附表_pre_buhu.xlsx |
| 总台账备份 | 总台账底稿_pre_buhu.csv |

---

*本手册由 Claude Code 根据周贤案实操沉淀 · 2026-06-01*

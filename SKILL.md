---
name: generate-lbank-kol-promo-v1
description: Use when generating local v1 LBank KOL or affiliate promo copy from an Excel/CSV table and pasted or file-based activity text, with BD-confirmed rows, per-language copy matching, invite-code registration links, event icode links, and user-selected local output folders.
---

# Generate LBank KOL Promo V1

## Overview

Use this skill to turn LBank KOL/affiliate tables plus activity copy into local promo output files. V1 is local-file only: it does not read/write Feishu, send Telegram messages, edit images, or translate copy. AI may be used to inspect inputs, clarify ambiguous business rules, and explain anomalies, but the final workbook generation is deterministic script output.

## Inputs

Supported inputs:

- Preferred template workbook with sheets `KOL状态 最新状态`, `本期活动内容`, and `宣发及回链内容`
- Legacy `.xlsx` or `.csv` table with the fixed fields in `references/field_schema.md`
- For legacy mode only: activity copy pasted in chat, one `.txt` file, or a folder of language `.txt` files

The input table may already contain empty output columns: `生成宣发文案`, `生成状态`, and `跳过原因`. The script ignores old values in those columns and writes a new output workbook.

For one unmarked activity text, apply the same copy to every generated row. For pasted or file-based multilingual copy, mark each block with a line like `[繁中]` or `[en]`; the marker should match the table's `语言` value after whitespace/case normalization. For a folder, name each `.txt` file after the `语言` value; add `default.txt` as fallback if useful.

## Workflow

1. Confirm the user has provided or identified the input table.
2. Confirm the output folder before running. If the user has not provided one, ask for it explicitly; do not assume a machine-specific path.
3. For legacy mode only, confirm the activity copy source.
4. If field behavior is unclear, read `references/field_schema.md`.
5. Run the bundled generator with `--output-root`.

Preferred three-sheet template mode:

```bash
python /path/to/generate-lbank-kol-promo-v1/scripts/generate_lbank_kol_promo.py \
  --input-table /path/to/AI_宣发模板.xlsx \
  --output-root /path/to/output_folder
```

Legacy single-table mode:

```bash
python /path/to/generate-lbank-kol-promo-v1/scripts/generate_lbank_kol_promo.py \
  --input-table /path/to/kol_table.xlsx \
  --activity-text /path/to/activity.txt \
  --output-root /path/to/output_folder
```

If the user pasted activity copy directly in chat, either pass it with `--activity-copy` or save it to a temporary `.txt` first. Use `--activity-text /path/to/activity_text_folder` when there are per-language `.txt` files.

6. Report the created output folder and generation/skipped counts.
7. Never mark rows sent and never push messages in v1.

## Generation Rules

- In preferred template mode, only generate rows where `投放复核状态` is `通过`.
- In preferred template mode, each non-empty row in `本期活动内容` is an independent activity. One activity produces one output promo sheet; two activities produce two sheets; three activities produce three sheets.
- For a single activity, write `宣发及回链内容`. For multiple activities, write `宣发及回链内容_活动1`, `宣发及回链内容_活动2`, etc.
- Each activity sheet contains one row per invite code/KOL record. Do not merge multiple activities into one long promo message or one mixed tracking sheet.
- In preferred template mode, the `语言` column on KOL rows is retained and can affect generated step labels, but activity rows without a `语言` column are treated as common activity copy for all KOL languages. The script does not translate copy.
- In legacy mode, only generate rows where `BD确认状态` is confirmed and `本活动是否已发送` is not already sent.
- Step 1 registration link is always `https://{使用域名}/ref/{代理邀请码}`.
- Step 2 event links always carry `icode={代理邀请码}`:
  - Replace an existing `icode=...`.
  - Append `?icode=...` or `&icode=...` when missing.
- Remove all whitespace inside `代理邀请码` before generating links; for example, `AB C003` becomes `ABC003`.
- In preferred template mode, images embedded on the `配图` cell in `本期活动内容` are saved once under `assets/` per activity. The matching `活动宣发素材图` cells contain the relative asset path, avoiding repeated image embedding and oversized Excel files.
- Replace code placeholders such as `xxxx`, `{invite_code}`, and `{邀请码}`.
- Replace registration link placeholders such as `{registration_link}`, `{ref_link}`, and `{注册链接}`.
- Replace sample invite codes found in `/ref/...` and `icode=...` examples throughout the copy.
- Use `lbank.com` when `使用域名` is blank.

## Outputs

The script creates the next local run folder under the selected output folder. If `--output-root` is omitted, direct script execution defaults to `./output`, but skill usage should still confirm the user's intended output folder first.

```text
<output-folder>/test1
<output-folder>/test2
...
```

Preferred three-sheet template mode writes the main workbook at the run-folder root and puts check/debug files under `debug/`:

```text
testN/
  filled_template.xlsx
  assets/
    activity_1_image_1.png
  debug/
    input_original.xlsx
    generated_messages.xlsx
    generated_messages.csv
    generated_messages.md
    skipped_rows.xlsx
    skipped_rows.csv
```

Legacy single-table mode has no filled template, so `generated_messages.xlsx` is the root main workbook and the remaining check/debug files are under `debug/`.

## Verification

Before calling the skill complete after editing it, run:

```bash
python -m unittest discover -s /path/to/generate-lbank-kol-promo-v1/scripts -p "test_*.py"
python /path/to/.codex/skills/.system/skill-creator/scripts/quick_validate.py /path/to/generate-lbank-kol-promo-v1
```

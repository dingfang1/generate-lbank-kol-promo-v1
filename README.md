# Generate LBank KOL Promo V1

Local Codex skill for generating LBank KOL / affiliate promo output files from an Excel template.

## What It Does

- Reads the three-sheet Excel template:
  - `KOL状态 最新状态`
  - `本期活动内容`
  - `宣发及回链内容`
- Generates per-KOL promo copy.
- Treats each non-empty row in `本期活动内容` as an independent activity and writes one promo sheet per activity.
- Builds registration links from each KOL's domain and invite code.
- Rewrites LBank event links with `icode=<invite_code>`.
- Saves embedded activity images once under `assets/` and writes relative paths into `活动宣发素材图`.
- Writes local output files only. It does not send Telegram messages, write Feishu, or mark rows as sent.

## Install As A Codex Skill

Clone this repository into your Codex skills folder:

```powershell
git clone <repo-url> "$env:USERPROFILE\.codex\skills\generate-lbank-kol-promo-v1"
```

Then restart Codex so it can scan the skill.

## Python Dependencies

```powershell
python -m pip install -r requirements.txt
```

## Usage

Choose both paths before running:

- `--input-table`: your source Excel or CSV file
- `--output-root`: the folder where run folders such as `test1`, `test2`, etc. should be created

Three-sheet template mode:

```powershell
python "$env:USERPROFILE\.codex\skills\generate-lbank-kol-promo-v1\scripts\generate_lbank_kol_promo.py" `
  --input-table "D:\path\to\input.xlsx" `
  --output-root "D:\path\to\output"
```

The script creates the next available output folder under your selected output folder:

```text
<output-folder>\test1
<output-folder>\test2
...
```

In three-sheet template mode, each run folder keeps the main workbook at the root and puts check files under `debug/`:

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

In legacy single-table mode, `generated_messages.xlsx` is the root main workbook and the remaining check files are under `debug/`.

## Notes

- Direct script execution defaults to `.\output` only when `--output-root` is omitted. For normal use, pass `--output-root` explicitly so the output goes where you expect.
- Only rows with `投放复核状态 = 通过` are generated in three-sheet template mode.
- If `本期活动内容` has multiple activity rows, the output workbook uses sheets such as `宣发及回链内容_活动1` and `宣发及回链内容_活动2`.
- Blank invite codes are skipped.

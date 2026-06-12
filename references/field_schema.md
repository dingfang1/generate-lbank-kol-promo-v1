# Field Schema

## Preferred Three-Sheet Template

Use this structure when possible.

### `KOL状态 最新状态`

| Field | Purpose |
| --- | --- |
| `代理UID` | Affiliate/KOL identifier for reporting. |
| `代理邀请码` | Invite suffix used in registration and event links. |
| `语言` | Language version used to match `本期活动内容`. |
| `地区` | Region metadata retained in output. |
| `使用域名` | Domain for Step 1 registration link. Defaults to `lbank.com` if blank. |
| `所属BD` | Retained in output for BD handoff. |
| `BD申请状态` | Retained source status; not a generation gate. |
| `投放复核状态` | Generation gate. Only `通过` rows are generated. |

### `本期活动内容`

Each non-empty activity row is treated as one independent activity. In preferred template mode, one activity row produces one output promo sheet; multiple activity rows produce `宣发及回链内容_活动1`, `宣发及回链内容_活动2`, etc. Do not combine multiple activities into one output sheet.

| Field | Purpose |
| --- | --- |
| `语言` | Optional language metadata. If blank or absent, the activity row is common copy for all KOL languages. |
| `标题` | Prepended to generated promo copy. |
| `正文` | Main activity copy. Event links inside it are rewritten with `icode`. |
| `链接` | Canonical Step 2 activity link. It is rewritten with `icode`. |
| `配图` | Text values are copied to `活动宣发素材图`. Embedded images anchored on this cell are saved once under `assets/`, and the matching output cells contain the relative asset path. |

### `宣发及回链内容`

The generator writes a new `filled_template.xlsx` copy and fills this sheet. It does not modify the input workbook. If there are multiple activity rows, this sheet is replaced by one sheet per activity, named `宣发及回链内容_活动1`, `宣发及回链内容_活动2`, etc.

Generated fields include `注册链接`, `活动链接`, `活动宣发内容`, `活动宣发素材图`, `生成状态`, `跳过原因`, and `生成时间`. `回链`, `截图`, and payment fields remain blank for later operations.

## Legacy Required Input Fields

| Field | Purpose |
| --- | --- |
| `代理UID` | Affiliate/KOL identifier for reporting. |
| `代理邀请码` | Invite suffix used in registration and event links. |
| `语言` | Language version. Used to choose matching `.txt` when activity copy is provided as a folder. |
| `地区` | Region metadata retained in output. |
| `使用域名` | Domain for Step 1 registration link. Defaults to `lbank.com` if blank. |
| `BD确认状态` | Generation gate. Only confirmed rows are generated. |
| `本活动是否已发送` | Safety gate. Rows already sent for this activity are skipped. This does not mean whether the row should be sent now. |
| `TG群名` | Retained in Markdown and CSV output. |
| `TG链接` | Retained in Markdown and CSV output. |

## Optional Empty Input Fields

The input workbook may include these fields as empty columns. They are filled in the generated output workbook, not in the original file.

| Field | Purpose |
| --- | --- |
| `生成宣发文案` | Final generated promo copy. |
| `生成状态` | `已生成` or `已跳过`. |
| `跳过原因` | Reason a row was skipped. |

## Status Rules

Confirmed values include `已确认`, `确认`, `是`, `yes`, `true`, `1`, `ok`, and `confirmed`.

`本活动是否已发送` sent values include `已发送`, `已推送`, `已发`, `发送`, `是`, `yes`, `true`, `1`, and `sent`.

Rows are skipped when:

- `代理邀请码` is blank
- `BD确认状态` is not confirmed
- `本活动是否已发送` is already sent
- per-language activity copy is missing when a language folder is used

## Link Rules

Registration link:

```text
https://{使用域名}/ref/{代理邀请码}
```

Before link generation, all whitespace inside `代理邀请码` is removed. For example, `AB C003` becomes `ABC003`.

Event links containing `/event-new/` must include:

```text
icode={代理邀请码}
```

If an event URL already has `icode`, replace its value. If not, append `icode` while preserving existing query parameters.

## Activity Copy Rules

In preferred template mode, each `本期活动内容` row is an independent activity. Activity rows without `语言` are applied to all KOL languages. The KOL `语言` value is retained and may affect generated step labels, but the script does not translate activity copy.

Unmarked legacy copy is applied to all generated rows.

For multilingual pasted copy or one multilingual `.txt`, use simple language markers:

```text
[繁中]
繁中文案...

[en]
English copy...
```

The marker should match the table's `语言` value. If a row's language has no matching block and no `default` block exists, the row is skipped with `缺少对应语言活动文案`.

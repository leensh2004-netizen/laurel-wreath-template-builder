# Laurel Wreath Template Builder V11 - Movement Tables

This version keeps all V9 fixes and adds a reviewed workflow for movement tables.

## Main features

- Fixed Arabic template included inside `templates/`.
- Employee fills company info, partners, policies, sections, and custom sections.
- Trial balance Excel upload supports `.xls` and `.xlsx`.
- Current year normally uses Excel column `النهائي`.
- Previous/comparative year uses the numeric column immediately before `المجموعة`.
- Financial note rows are extracted by exact normalized `المجموعة` matching.
- User can manually edit extracted values before generating.
- V11 adds fixed-asset movement table suggestions for note `(8) الممتلكات والمعدات`.

## Important V11 note

The fixed-asset movement table should be reviewed before applying. It is disabled by default.
Turn on:

`Also update 2024 fixed-asset movement tables`

only after checking the suggested values.

The movement logic is:

- Opening cost = previous/comparative cost balance
- Additions = current cost minus previous cost, if positive
- Closing cost = current cost
- Opening accumulated depreciation = absolute previous accumulated depreciation
- Depreciation for the year = absolute current accumulated depreciation minus absolute previous accumulated depreciation, if positive
- Closing accumulated depreciation = absolute current accumulated depreciation
- Net book value = current cost minus accumulated depreciation

V11 leaves unclear categories unchanged instead of guessing.

## Run

Double click `run_app.bat`, or run:

```cmd
python -m streamlit run app.py
```


V11 update:
- Removed sample placeholder text from the website input fields for a cleaner manager-facing review version.
- Kept the fixed template, Excel mapping, editable policy boxes, and movement-table workflow.

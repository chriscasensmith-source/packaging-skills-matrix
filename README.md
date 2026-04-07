# Packaging Skills Matrix Workbook

This project builds a practical Excel workbook for packaging training leaders to manage:
- Core skill proficiency for PLT/MT II role family
- Line qualification status across 21 packaging lines
- Critical gap visibility
- Leadership dashboard coverage and risk views

## Files
- `build_workbook.py` - Python generator script (openpyxl only, no macros/VBA)
- `packaging_skills_matrix.xlsx` - Generated workbook
- `README.md` - Usage and maintenance guide

## How to run
1. Install dependency:
   ```bash
   pip install openpyxl
   ```
2. Build workbook:
   ```bash
   python build_workbook.py
   ```
3. Open `packaging_skills_matrix.xlsx` in Excel.

## Workbook tabs (in order)
1. `Instructions`
2. `Employee_List`
3. `Skill_Definitions`
4. `Core_Skill_Assessments`
5. `Line_Qualifications`
6. `Dashboard`
7. `Lists` (hidden by default)

## Where to update data
### Update employees
- Go to `Employee_List` table.
- Add/edit rows in the roster.
- Keep `Employee_ID` unique.
- Use `Active_Status` for active/inactive management.

### Update line names
- Unhide `Lists` tab.
- Edit `line_name` values.
- These names feed dashboard labels and admin lists.

### Update skill definitions
- Go to `Skill_Definitions`.
- Maintain one row per competency.
- `Required_Level_For_Role` defaults to 3 and can be adjusted by competency.

### Update dropdown values
- Unhide `Lists` tab and edit:
  - skill levels
  - line qualification levels
  - shifts
  - cells
  - role family
  - active status
  - low coverage threshold

## Dashboard logic
- **Line Coverage Summary:** counts Level 3+ and Level 4 by line; flags `LOW` when Level 3+ is below threshold.
- **Critical Skill Gaps by Employee:** critical gaps + total below-required counts per employee.
- **Critical Skill Gaps by Competency:** shows which critical skills have most gaps.
- **Cross Training Opportunity:** highlights employees with `Core_Critical_Ready=Yes` and limited line depth.
- **Trainer Bench Strength:** compares Level 4 line coverage by employee.

## Assumptions used
- PLT and MT II are combined into one role family value: `PLT_MT2`.
- Example employee roster uses placeholders (12 rows) for visibility and testing.
- No confirmed real line names were found in this repository; default placeholders `Line_01` to `Line_21` were used.
- Low coverage threshold default is `3` (editable in `Lists`).

## Suggested next customizations
1. Replace placeholder employee names with current roster.
2. Replace line placeholders with actual line names used on site.
3. Assign primary trainers and team leads based on your org.
4. Preload current known qualification levels to get immediate dashboard value.
5. Review `Critical_Flag` and required levels with operations and quality leadership.


## Validation checklist
After running the generator, verify:
1. `packaging_skills_matrix.xlsx` exists in the project root.
2. Re-open it with openpyxl:
   ```bash
   python -c "from openpyxl import load_workbook; load_workbook('packaging_skills_matrix.xlsx'); print('Workbook opened successfully')"
   ```
3. Open the workbook in Excel and confirm there are no repair warnings.

## Troubleshooting
- If `ModuleNotFoundError: No module named openpyxl` appears, install dependency first with `pip install openpyxl` in an environment with package access.
- If package installation is blocked by network policy, run the script on a workstation or CI environment that has openpyxl preinstalled.

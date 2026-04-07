from __future__ import annotations

from datetime import date
from pathlib import Path
import re
from typing import List, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

OUTPUT_FILE = "packaging_skills_matrix.xlsx"

EMPLOYEES = [
    ("E1001", "Avery Carter", "PLT_MT2", "A", "Cell 1", "Jordan Lee", "M. Patel", "2022-03-14", "Active"),
    ("E1002", "Blake Nguyen", "PLT_MT2", "A", "Cell 2", "Jordan Lee", "S. Morris", "2021-10-03", "Active"),
    ("E1003", "Casey Ramirez", "PLT_MT2", "B", "Cell 1", "Taylor Kim", "M. Patel", "2023-01-09", "Active"),
    ("E1004", "Dakota Singh", "PLT_MT2", "B", "Cell 3", "Taylor Kim", "A. Brooks", "2020-06-21", "Active"),
    ("E1005", "Emerson Wright", "PLT_MT2", "C", "Cell 2", "Robin Hall", "S. Morris", "2024-02-12", "Active"),
    ("E1006", "Finley Torres", "PLT_MT2", "C", "Cell 3", "Robin Hall", "A. Brooks", "2022-11-18", "Active"),
    ("E1007", "Gray Allen", "PLT_MT2", "D", "Cell 1", "Jordan Lee", "M. Patel", "2019-09-27", "Active"),
    ("E1008", "Harper Diaz", "PLT_MT2", "D", "Cell 2", "Taylor Kim", "S. Morris", "2021-04-30", "Active"),
    ("E1009", "Indigo Foster", "PLT_MT2", "A", "Cell 3", "Robin Hall", "A. Brooks", "2020-12-07", "Active"),
    ("E1010", "Jules Bennett", "PLT_MT2", "B", "Cell 2", "Jordan Lee", "M. Patel", "2023-07-19", "Active"),
    ("E1011", "Kai Morgan", "PLT_MT2", "C", "Cell 1", "Taylor Kim", "S. Morris", "2024-05-13", "Active"),
    ("E1012", "Logan Price", "PLT_MT2", "D", "Cell 3", "Robin Hall", "A. Brooks", "2022-08-05", "Inactive"),
]

SKILLS = [
    ("Safety and Compliance", "C01", "Applies LOTO correctly for the task being performed", "Performs LOTO without missing steps and verifies zero energy before work begins", "Yes", 3, ""),
    ("Safety and Compliance", "C02", "Identifies emergency stops and safety devices on the assigned line", "Can locate and explain the purpose of major safety devices without prompting", "Yes", 3, ""),
    ("Safety and Compliance", "C03", "Uses proper PPE and follows GMP and GDP during active work", "Uses correct PPE and maintains compliant practices during normal work without reminders", "Yes", 3, ""),
    ("Line and Equipment Understanding", "C04", "Explains line flow from start to finish", "Can walk the line and explain the purpose of each major section in the process", "No", 3, ""),
    ("Line and Equipment Understanding", "C05", "Identifies major components and their function", "Correctly identifies major machine components and explains their basic function", "No", 3, ""),
    ("Line and Equipment Understanding", "C06", "Navigates HMI screens needed for operation", "Uses the HMI to monitor status and perform normal operating actions without help", "No", 3, ""),
    ("Startup and Pre Run Checks", "C07", "Performs pre run checks completely", "Completes all required pre run checks without skipping steps", "Yes", 3, ""),
    ("Startup and Pre Run Checks", "C08", "Verifies correct components against the work order", "Confirms components match the work order and catches mismatches before startup", "Yes", 3, ""),
    ("Startup and Pre Run Checks", "C09", "Confirms line readiness before startup", "Verifies line setup, materials, and readiness before startup without prompting", "Yes", 3, ""),
    ("Running the Line", "C10", "Maintains steady line operation", "Runs the line with normal stability and avoids unnecessary stops caused by operator error", "No", 3, ""),
    ("Running the Line", "C11", "Makes basic adjustments to maintain flow", "Performs normal operating adjustments to keep the line flowing without needing frequent help", "No", 3, ""),
    ("Running the Line", "C12", "Communicates issues before they escalate", "Recognizes developing issues and communicates early with the right level of detail", "No", 3, ""),
    ("In Process Quality", "C13", "Verifies coding against the work order", "Confirms coding matches the work order before release and recognizes obvious coding errors", "Yes", 3, ""),
    ("In Process Quality", "C14", "Performs FAI verification correctly", "Completes required FAI verification steps accurately without missing key checks", "Yes", 3, ""),
    ("In Process Quality", "C15", "Identifies common defects and responds appropriately", "Recognizes common product or package defects and takes the correct next step", "Yes", 3, ""),
    ("Troubleshooting and Mechanical Skill", "C16", "Clears jams safely", "Clears common jams safely without causing additional issues", "No", 3, ""),
    ("Troubleshooting and Mechanical Skill", "C17", "Performs minor mechanical adjustments", "Completes routine minor adjustments correctly and knows when the issue is beyond their level", "No", 3, ""),
    ("Troubleshooting and Mechanical Skill", "C18", "Identifies likely root cause of common stops", "Can explain the likely cause of a common stop and take the correct first response", "No", 3, ""),
    ("Changeover and Setup", "C19", "Performs routine format change steps", "Completes routine changeover tasks in the correct sequence with minimal correction", "Yes", 3, ""),
    ("Changeover and Setup", "C20", "Verifies correct parts and settings after changeover", "Confirms parts and settings are correct before startup after a changeover", "Yes", 3, ""),
    ("Documentation and Systems", "C21", "Completes work orders accurately", "Enters required work order information accurately and completely", "No", 3, ""),
    ("Documentation and Systems", "C22", "Uses Informance correctly including downtime tracking", "Uses correct downtime codes and enters data accurately without rework", "No", 3, ""),
    ("Documentation and Systems", "C23", "Updates whiteboards and batch information accurately", "Keeps shift and production information accurate and current", "No", 3, ""),
    ("Sanitation and Shutdown", "C24", "Performs proper shutdown", "Shuts down the line in the correct sequence and leaves equipment in the proper state", "No", 3, ""),
    ("Sanitation and Shutdown", "C25", "Leaves the line clean and ready for next shift or sanitation", "Leaves the area and equipment in acceptable condition for the next step in the process", "No", 3, ""),
    ("Escalation and Ownership", "C26", "Knows when and how to escalate", "Escalates at the right time with clear, useful information", "No", 3, ""),
    ("Escalation and Ownership", "C27", "Gives clear shift handoff and ownership communication", "Provides a complete and accurate handoff without leaving important issues unclear", "No", 3, ""),
]


def find_line_names(repo_root: Path, max_lines: int = 21) -> Tuple[List[str], str]:
    candidates: set[str] = set()
    line_pattern = re.compile(r"\b(Line[\s_-]?\d{1,2}|L\d{1,2})\b", re.IGNORECASE)
    for path in repo_root.glob("**/*"):
        if path.is_dir():
            continue
        if path.suffix.lower() not in {".md", ".txt", ".csv", ".json", ".yml", ".yaml"}:
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        for match in line_pattern.findall(text):
            normalized = match.replace(" ", "_").replace("-", "_")
            if normalized.lower().startswith("l") and not normalized.lower().startswith("line"):
                digits = re.sub(r"\D", "", normalized)
                normalized = f"Line_{int(digits):02d}" if digits else ""
            if normalized:
                candidates.add(normalized.title())
    cleaned = sorted([c for c in candidates if c.lower().startswith("line_")], key=lambda x: int(re.sub(r"\D", "", x) or 999))
    if cleaned:
        return cleaned[:max_lines] + [f"Line_{i:02d}" for i in range(len(cleaned) + 1, max_lines + 1)], "Detected line-like names from repository text files"
    return [f"Line_{i:02d}" for i in range(1, max_lines + 1)], "No line names found; used placeholders Line_01 to Line_21"


def style_header(ws, row=1):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9D9D9")
    for cell in ws[row]:
        if cell.value is None:
            continue
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def add_table(ws, name: str, start_cell: str, end_cell: str):
    table = Table(displayName=name, ref=f"{start_cell}:{end_cell}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)


def build_workbook():
    repo_root = Path(__file__).resolve().parent
    line_names, line_name_note = find_line_names(repo_root)

    wb = Workbook()
    ws_instructions = wb.active
    ws_instructions.title = "Instructions"
    ws_employee = wb.create_sheet("Employee_List")
    ws_skills = wb.create_sheet("Skill_Definitions")
    ws_assess = wb.create_sheet("Core_Skill_Assessments")
    ws_lines = wb.create_sheet("Line_Qualifications")
    ws_dash = wb.create_sheet("Dashboard")
    ws_lists = wb.create_sheet("Lists")

    # Instructions
    ws_instructions["A1"] = "Packaging Skills Matrix & Line Qualification Workbook"
    ws_instructions["A1"].font = Font(size=14, bold=True)
    instruction_text = [
        ("Workbook purpose", "Track packaging core skills, line qualifications, and staffing risk for PLT/MT II."),
        ("Tab guide", "Employee_List = master roster; Skill_Definitions = one row per competency; Core_Skill_Assessments = ratings per person/skill; Line_Qualifications = ratings per person/line; Dashboard = risk and coverage; Lists = admin values for dropdowns/thresholds."),
        ("Core skill levels", "1 Awareness, 2 Assisted, 3 Independent (target), 4 Trainer."),
        ("Line qualification levels", "1 Not trained, 2 Can run with help, 3 Can run independently, 4 Can train others."),
        ("Update employees", "Add or edit rows in Employee_List table. Keep Employee_ID unique. Active_Status controls who appears in assessment and line views when rebuilt."),
        ("Update line names", "Edit line names in Lists tab (line_names list). Re-run script to refresh formulas and dashboard labels."),
        ("Update skill definitions", "Edit rows in Skill_Definitions table. Required level defaults to 3; set different value only where needed."),
        ("Critical gap logic", "Critical_Gap = Yes when Critical_Flag = Yes and Current_Level is below Required_Level_For_Role."),
        ("Line readiness logic", "Core_Critical_Ready = Yes only when employee has zero critical gaps in Core_Skill_Assessments."),
        ("Dashboard flags", "Red/amber flags indicate low line coverage or critical risk. Low coverage threshold is editable in Lists tab."),
    ]
    r = 3
    for title, body in instruction_text:
        ws_instructions[f"A{r}"] = title
        ws_instructions[f"A{r}"].font = Font(bold=True)
        ws_instructions[f"B{r}"] = body
        ws_instructions[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    ws_instructions.column_dimensions["A"].width = 28
    ws_instructions.column_dimensions["B"].width = 120
    ws_instructions.freeze_panes = "A3"

    # Lists
    ws_lists.append(["List_Name", "Value", "Label"])
    rows = [
        ("core_skill_level", 1, "Awareness"),
        ("core_skill_level", 2, "Assisted"),
        ("core_skill_level", 3, "Independent"),
        ("core_skill_level", 4, "Trainer"),
        ("line_level", 1, "Not trained"),
        ("line_level", 2, "Can run with help"),
        ("line_level", 3, "Can run independently"),
        ("line_level", 4, "Can train others"),
        ("role_family", "PLT_MT2", "PLT + MT II"),
        ("shift", "A", "Shift A"),
        ("shift", "B", "Shift B"),
        ("shift", "C", "Shift C"),
        ("shift", "D", "Shift D"),
        ("cell", "Cell 1", "Cell 1"),
        ("cell", "Cell 2", "Cell 2"),
        ("cell", "Cell 3", "Cell 3"),
        ("active_status", "Active", "Active"),
        ("active_status", "Inactive", "Inactive"),
        ("admin", "Low_Coverage_Threshold", "3"),
    ]
    for line in line_names:
        rows.append(("line_name", line, line))
    for row in rows:
        ws_lists.append(row)
    style_header(ws_lists)
    add_table(ws_lists, "ListsTable", "A1", f"C{ws_lists.max_row}")
    ws_lists.column_dimensions["A"].width = 22
    ws_lists.column_dimensions["B"].width = 28
    ws_lists.column_dimensions["C"].width = 36

    # Named ranges for validation lists
    wb.defined_names.add(DefinedName(name="RoleFamilyList", attr_text="Lists!$B$10:$B$10"))
    wb.defined_names.add(DefinedName(name="ShiftList", attr_text="Lists!$B$11:$B$14"))
    wb.defined_names.add(DefinedName(name="CellList", attr_text="Lists!$B$15:$B$17"))
    wb.defined_names.add(DefinedName(name="ActiveStatusList", attr_text="Lists!$B$18:$B$19"))
    wb.defined_names.add(DefinedName(name="CoreLevelList", attr_text="Lists!$B$2:$B$5"))
    wb.defined_names.add(DefinedName(name="LineLevelList", attr_text="Lists!$B$6:$B$9"))
    wb.defined_names.add(DefinedName(name="LowCoverageThreshold", attr_text="Lists!$C$20"))
    wb.defined_names.add(DefinedName(name="LineNamesList", attr_text=f"Lists!$B$21:$B${20 + len(line_names)}"))

    # Employee list
    employee_headers = ["Employee_ID", "Employee_Name", "Role_Family", "Shift", "Cell", "Team_Lead", "Primary_Trainer", "Hire_Date", "Active_Status"]
    ws_employee.append(employee_headers)
    for e in EMPLOYEES:
        ws_employee.append(list(e))
    style_header(ws_employee)
    add_table(ws_employee, "EmployeeTable", "A1", f"I{ws_employee.max_row}")
    ws_employee.freeze_panes = "A2"
    for c, w in {"A": 14, "B": 22, "C": 14, "D": 8, "E": 10, "F": 16, "G": 16, "H": 12, "I": 12}.items():
        ws_employee.column_dimensions[c].width = w
    for col, rng in [("C", "RoleFamilyList"), ("D", "ShiftList"), ("E", "CellList"), ("I", "ActiveStatusList")]:
        dv = DataValidation(type="list", formula1=f"={rng}", allow_blank=False)
        ws_employee.add_data_validation(dv)
        dv.add(f"{col}2:{col}500")

    # Skill definitions
    skills_headers = ["Category", "Competency_ID", "Competency", "Level_3_Standard", "Critical_Flag", "Required_Level_For_Role", "Notes"]
    ws_skills.append(skills_headers)
    for skill in SKILLS:
        ws_skills.append(list(skill))
    style_header(ws_skills)
    add_table(ws_skills, "SkillTable", "A1", f"G{ws_skills.max_row}")
    ws_skills.freeze_panes = "A2"
    for col, w in {"A": 30, "B": 14, "C": 50, "D": 70, "E": 12, "F": 22, "G": 24}.items():
        ws_skills.column_dimensions[col].width = w
    for row in ws_skills.iter_rows(min_row=2, max_row=ws_skills.max_row, min_col=1, max_col=7):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Core skill assessments
    assess_headers = [
        "Employee_ID", "Employee_Name", "Role_Family", "Shift", "Cell", "Competency_ID", "Category", "Competency",
        "Level_3_Standard", "Critical_Flag", "Required_Level_For_Role", "Current_Level", "Meets_Required", "Critical_Gap",
        "Last_Assessed_By", "Last_Assessed_Date", "Comments"
    ]
    ws_assess.append(assess_headers)
    active_employees = [e for e in EMPLOYEES if e[8] == "Active"]
    for emp in active_employees:
        for skill in SKILLS:
            ws_assess.append([
                emp[0], emp[1], emp[2], emp[3], emp[4], skill[1], skill[0], skill[2], skill[3], skill[4], skill[5], 2,
                "", "", "", "", ""
            ])
    for r in range(2, ws_assess.max_row + 1):
        ws_assess[f"M{r}"] = f'=IF(L{r}>=K{r},"Yes","No")'
        ws_assess[f"N{r}"] = f'=IF(AND(J{r}="Yes",L{r}<K{r}),"Yes","No")'
    style_header(ws_assess)
    add_table(ws_assess, "AssessmentTable", "A1", f"Q{ws_assess.max_row}")
    ws_assess.freeze_panes = "A2"
    widths = [12, 20, 12, 8, 8, 14, 28, 42, 60, 10, 16, 12, 12, 12, 16, 14, 24]
    for i, w in enumerate(widths, 1):
        ws_assess.column_dimensions[get_column_letter(i)].width = w
    for row in ws_assess.iter_rows(min_row=2, max_row=ws_assess.max_row, min_col=7, max_col=17):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    dv_core = DataValidation(type="list", formula1="=CoreLevelList", allow_blank=False)
    ws_assess.add_data_validation(dv_core)
    dv_core.add(f"L2:L{ws_assess.max_row}")
    # CF
    ws_assess.conditional_formatting.add(
        f"L2:L{ws_assess.max_row}",
        FormulaRule(formula=["$L2<$K2"], fill=PatternFill("solid", fgColor="FCE4D6"))
    )
    ws_assess.conditional_formatting.add(
        f"A2:Q{ws_assess.max_row}",
        FormulaRule(formula=["$N2=\"Yes\""], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True))
    )

    # Line qualifications
    line_headers = ["Employee_ID", "Employee_Name", "Role_Family", "Shift", "Cell", "Core_Critical_Ready"] + [f"Line_{i:02d}" for i in range(1, 22)] + ["Total_Lines_Level_3_Plus", "Total_Lines_Level_4", "Notes"]
    ws_lines.append(line_headers)
    for emp in active_employees:
        ws_lines.append([emp[0], emp[1], emp[2], emp[3], emp[4], ""] + [1] * 21 + [0, 0, ""])
    for r in range(2, ws_lines.max_row + 1):
        ws_lines[f"F{r}"] = f'=IF(COUNTIFS(Core_Skill_Assessments!$A:$A,A{r},Core_Skill_Assessments!$N:$N,"Yes")=0,"Yes","No")'
        ws_lines[f"AB{r}"] = f"=COUNTIF(G{r}:AA{r},\">=3\")"
        ws_lines[f"AC{r}"] = f"=COUNTIF(G{r}:AA{r},4)"
    style_header(ws_lines)
    add_table(ws_lines, "LineQualTable", "A1", f"AD{ws_lines.max_row}")
    ws_lines.freeze_panes = "A2"
    ws_lines.page_setup.orientation = "landscape"
    ws_lines.page_setup.fitToWidth = 1
    ws_lines.page_setup.fitToHeight = 0
    for i in range(1, 31):
        ws_lines.column_dimensions[get_column_letter(i)].width = 11 if 7 <= i <= 27 else 14
    ws_lines.column_dimensions["B"].width = 20
    dv_line = DataValidation(type="list", formula1="=LineLevelList", allow_blank=False)
    ws_lines.add_data_validation(dv_line)
    dv_line.add(f"G2:AA{ws_lines.max_row}")
    ws_lines.conditional_formatting.add(f"G2:AA{ws_lines.max_row}", CellIsRule(operator="equal", formula=["1"], fill=PatternFill("solid", fgColor="F8CBAD")))
    ws_lines.conditional_formatting.add(f"G2:AA{ws_lines.max_row}", CellIsRule(operator="equal", formula=["2"], fill=PatternFill("solid", fgColor="FCE4D6")))
    ws_lines.conditional_formatting.add(f"G2:AA{ws_lines.max_row}", CellIsRule(operator="equal", formula=["3"], fill=PatternFill("solid", fgColor="E2F0D9")))
    ws_lines.conditional_formatting.add(f"G2:AA{ws_lines.max_row}", CellIsRule(operator="equal", formula=["4"], fill=PatternFill("solid", fgColor="C6E0B4")))
    ws_lines.conditional_formatting.add(
        f"G2:AA{ws_lines.max_row}",
        FormulaRule(formula=["=AND($F2=\"No\",G2>=3)"], fill=PatternFill("solid", fgColor="FFD966"), font=Font(bold=True, color="7F6000"))
    )

    # Dashboard
    ws_dash["A1"] = "Packaging Skills Matrix Dashboard"
    ws_dash["A1"].font = Font(size=14, bold=True)
    ws_dash["A2"] = f"Generated: {date.today().isoformat()}"

    # Section A
    ws_dash["A4"] = "Section A: Line Coverage Summary"
    ws_dash["A4"].font = Font(bold=True)
    ws_dash.append(["Line", "Employees Level 3+", "Employees Level 4", "Low Coverage Flag"])
    header_row = ws_dash.max_row
    for idx in range(1, 22):
        row = ws_dash.max_row + 1
        line_col = get_column_letter(6 + idx)
        ws_dash[f"A{row}"] = f"=Lists!B{20 + idx}"
        ws_dash[f"B{row}"] = f"=COUNTIF(Line_Qualifications!{line_col}:{line_col},\">=3\")"
        ws_dash[f"C{row}"] = f"=COUNTIF(Line_Qualifications!{line_col}:{line_col},4)"
        ws_dash[f"D{row}"] = f"=IF(B{row}<LowCoverageThreshold,\"LOW\",\"OK\")"

    # Section B
    start_b = ws_dash.max_row + 2
    ws_dash[f"A{start_b}"] = "Section B: Critical Skill Gaps by Employee"
    ws_dash[f"A{start_b}"].font = Font(bold=True)
    ws_dash[f"A{start_b+1}"] = "Employee_Name"
    ws_dash[f"B{start_b+1}"] = "Critical_Gaps"
    ws_dash[f"C{start_b+1}"] = "All_Skills_Below_Required"
    for i, emp in enumerate(active_employees, start=start_b + 2):
        ws_dash[f"A{i}"] = emp[1]
        ws_dash[f"B{i}"] = f'=COUNTIFS(Core_Skill_Assessments!$B:$B,A{i},Core_Skill_Assessments!$N:$N,"Yes")'
        ws_dash[f"C{i}"] = f'=COUNTIFS(Core_Skill_Assessments!$B:$B,A{i},Core_Skill_Assessments!$M:$M,"No")'

    # Section C
    start_c = ws_dash.max_row + 2
    ws_dash[f"A{start_c}"] = "Section C: Critical Skill Gaps by Competency"
    ws_dash[f"A{start_c}"].font = Font(bold=True)
    ws_dash[f"A{start_c+1}"] = "Competency_ID"
    ws_dash[f"B{start_c+1}"] = "Competency"
    ws_dash[f"C{start_c+1}"] = "Employees_Below_Required"
    crit_skills = [s for s in SKILLS if s[4] == "Yes"]
    for i, skill in enumerate(crit_skills, start=start_c + 2):
        ws_dash[f"A{i}"] = skill[1]
        ws_dash[f"B{i}"] = skill[2]
        ws_dash[f"C{i}"] = f'=COUNTIFS(Core_Skill_Assessments!$F:$F,A{i},Core_Skill_Assessments!$N:$N,"Yes")'

    # Section D
    start_d = ws_dash.max_row + 2
    ws_dash[f"A{start_d}"] = "Section D: Cross Training Opportunity View"
    ws_dash[f"A{start_d}"].font = Font(bold=True)
    ws_dash[f"A{start_d+1}"] = "Employee_Name"
    ws_dash[f"B{start_d+1}"] = "Core_Critical_Ready"
    ws_dash[f"C{start_d+1}"] = "Lines_Level_3_Plus"
    row_d = start_d + 2
    for idx in range(2, ws_lines.max_row + 1):
        ws_dash[f"A{row_d}"] = f"=Line_Qualifications!B{idx}"
        ws_dash[f"B{row_d}"] = f"=Line_Qualifications!F{idx}"
        ws_dash[f"C{row_d}"] = f"=Line_Qualifications!AB{idx}"
        row_d += 1

    # Section E
    start_e = ws_dash.max_row + 2
    ws_dash[f"A{start_e}"] = "Section E: Trainer Bench Strength"
    ws_dash[f"A{start_e}"].font = Font(bold=True)
    ws_dash[f"A{start_e+1}"] = "Employee_Name"
    ws_dash[f"B{start_e+1}"] = "Level_4_Line_Qualifications"
    for i, idx in enumerate(range(2, ws_lines.max_row + 1), start=start_e + 2):
        ws_dash[f"A{i}"] = f"=Line_Qualifications!B{idx}"
        ws_dash[f"B{i}"] = f"=Line_Qualifications!AC{idx}"

    style_header(ws_dash, row=header_row)
    style_header(ws_dash, row=start_b + 1)
    style_header(ws_dash, row=start_c + 1)
    style_header(ws_dash, row=start_d + 1)
    style_header(ws_dash, row=start_e + 1)
    ws_dash.freeze_panes = "A6"
    ws_dash.column_dimensions["A"].width = 36
    ws_dash.column_dimensions["B"].width = 24
    ws_dash.column_dimensions["C"].width = 24
    ws_dash.column_dimensions["D"].width = 18
    ws_dash.conditional_formatting.add(f"D{header_row+1}:D{header_row+21}", FormulaRule(formula=[f'$D{header_row+1}="LOW"'], fill=PatternFill("solid", fgColor="FFC7CE")))

    # common alignment
    for ws in [ws_employee, ws_skills, ws_assess, ws_lines, ws_dash]:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for c in row:
                if c.alignment is None:
                    c.alignment = Alignment(vertical="top")

    ws_lists.sheet_state = "hidden"

    wb.save(repo_root / OUTPUT_FILE)

    # Validation reload
    loaded = load_workbook(repo_root / OUTPUT_FILE)
    tab_names = loaded.sheetnames

    print("Completion summary")
    print(f"- Files created: build_workbook.py, {OUTPUT_FILE}")
    print(f"- Tabs created: {', '.join(tab_names)}")
    print(f"- Assumptions/fallbacks: Role_Family unified as PLT_MT2; seeded with 12 placeholder employees; {line_name_note}.")
    print(f"- Line names source: {'Detected from repository' if 'Detected' in line_name_note else 'Placeholder names used'}")


if __name__ == "__main__":
    build_workbook()

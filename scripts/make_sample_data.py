from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from roster_updater.constants import BASELINE_HEADERS, CORE_HEADERS


def row(
    number,
    display_name,
    logon,
    email,
    title,
    department,
    capability,
    manager,
    manager_email,
    erp_category="",
):
    first, *last = display_name.split()
    return {
        "S.No.": number,
        "Display Name": display_name,
        "First Name": first,
        "Last Name": " ".join(last),
        "Logon Name": logon,
        "Email Address": email,
        "Title": title,
        "Department": department,
        "Capability": capability,
        "Manager": manager,
        "Manager Email Address": manager_email,
        "Office": "Sample Office",
        "City": "Sample City",
        "Employee Type": "Employee",
        "ERP Category": erp_category,
    }


def write(path: Path, sheet_name: str, headers, rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(headers)
    for item in rows:
        worksheet.append([item.get(header, "") for header in headers])
    worksheet.freeze_panes = "A2"
    workbook.save(path)


def main():
    output = Path(__file__).resolve().parents[1] / "sample_data"
    output.mkdir(exist_ok=True)
    baseline = [
        row(1, "Alex Manager", "amanager", "alex@example.test", "Manager", "Operations", "Enable", "Erin Executive", "erin@example.test", "Supply Chain"),
        row(2, "Blake Builder", "bbuilder", "blake@example.test", "Technician", "Production", "Enable", "Alex Manager", "alex@example.test"),
        row(3, "Casey Departing", "cdeparting", "casey@example.test", "Designer", "Design", "Make", "Alex Manager", "alex@example.test"),
        row(4, "Erin Executive", "eexecutive", "erin@example.test", "Executive", "Leadership", "Enable", "Erin Executive", "erin@example.test"),
    ]
    monthly = [
        row(1, "Alex Manager", "amanager", "alex@example.test", "Senior Manager", "Operations", "", "Erin Executive", "erin@example.test"),
        row(2, "Blake Builder", "bbuilder", "blake@example.test", "Technician", "Production", "", "Alex Manager", "alex@example.test"),
        row(3, "Dana New", "dnew", "dana@example.test", "Coordinator", "Operations", "", "Alex Manager", "alex@example.test"),
        row(4, "Evan New", "enew", "evan@example.test", "Specialist", "Operations", "", "Inactive Manager", "inactive@example.test"),
        row(5, "Erin Executive", "eexecutive", "erin@example.test", "Executive", "Leadership", "", "Erin Executive", "erin@example.test"),
    ]
    write(output / "sample_baseline.xlsx", "All_Users_Active_Info (1)", BASELINE_HEADERS, baseline)
    write(output / "sample_monthly_pull.xlsx", "All Users", CORE_HEADERS, monthly)
    print(f"Created sample files in {output}")


if __name__ == "__main__":
    main()



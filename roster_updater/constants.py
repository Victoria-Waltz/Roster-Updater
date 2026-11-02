"""Roster column names and accepted header aliases."""
from __future__ import annotations




CORE_HEADERS = [
    "S.No.",
    "Display Name",
    "First Name",
    "Last Name",
    "Logon Name",
    "Email Address",
    "Title",
    "Department",
    "Capability",
    "Manager",
    "Manager Email Address",
    "Office",
    "City",
    "Employee Type",
]


BASELINE_HEADERS = CORE_HEADERS + ["ERP Category"]


KEY_HEADER = "Logon Name"


CAPABILITY_HEADER = "Capability"


ERP_HEADER = "ERP Category"


COMPARE_FIELDS = [
    "Display Name",
    "First Name",
    "Last Name",
    "Email Address",
    "Title",
    "Department",
    "Capability",
    "Manager",
    "Manager Email Address",
    "Office",
    "City",
    "Employee Type",
]


HEADER_ALIASES = {
    "sno": "S.No.",
    "serialnumber": "S.No.",
    "displayname": "Display Name",
    "firstname": "First Name",
    "lastname": "Last Name",
    "logonname": "Logon Name",
    "loginname": "Logon Name",
    "emailaddress": "Email Address",
    "email": "Email Address",
    "title": "Title",
    "department": "Department",
    "capability": "Capability",
    "competency": "Capability",
    "manager": "Manager",
    "manageremailaddress": "Manager Email Address",
    "manageremail": "Manager Email Address",
    "office": "Office",
    "city": "City",
    "employeetype": "Employee Type",
    "erpcategory": "ERP Category",
}

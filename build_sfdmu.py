"""Build SFDMU CSVs + export.json for the Atlas Manufacturing dev seed from the review workbook.

Field API names, record types and picklist values were checked against the dev sandbox
(CQSE-Golden-Dev, CQ 16.0) on 2026-11-09.
"""
import csv, json, os, shutil, zlib
from openpyxl import load_workbook

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Atlas_Manufacturing_Dev_Seed_Data.xlsx")
OUT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(OUT, "data")
CQ = "compliancequest__"
DUE_DATE = "2027-06-30"   # fixed future date: CQ rejects past due dates
# The user running the load. grant-cq-access.ps1 writes data/User.csv with this made-up Id and that org's
# username, so assignees resolve to the load user in any org.
LOAD_USER = "S99000000000000001"

wb = load_workbook(SRC)


def rows(tab):
    ws = wb[tab]
    hdr = [c.value for c in ws[3]]
    out = []
    for r in ws.iter_rows(min_row=4, values_only=True):
        if all(v is None for v in r):
            continue
        out.append({h: ("" if v is None else v) for h, v in zip(hdr, r)})
    return out


objects = []   # (sobject, externalId, operation, [fields for query], csv rows as dicts)


def add(sobject, external_id, columns, data, query_fields=None, operation="Upsert"):
    objects.append(dict(sobject=sobject, externalId=external_id, operation=operation,
                        columns=columns, rows=data, query_fields=query_fields))


def tf(v):
    return "TRUE" if str(v).strip().lower() in ("yes", "true", "1") else "FALSE"


site_division = {"Houston Assembly Plant": "Consumer Drones", "Piscataway Components Plant": "Commercial Drones"}

# ---------- Org structure ----------
add("compliancequest__SQX_Division__c", "Name", ["Name"],
    [{"Name": r["Name"]} for r in rows("Division")])
add("compliancequest__SQX_Business_Unit__c", "Name", ["Name"],
    [{"Name": r["Name"]} for r in rows("Business Unit")])
add("compliancequest__SQX_Site__c", "Name", ["Name", f"{CQ}SQX_Business_Unit__r.Name"],
    [{"Name": r["Name"], f"{CQ}SQX_Business_Unit__r.Name": "Atlas Manufacturing"} for r in rows("Site")],
    query_fields=["Name", f"{CQ}SQX_Business_Unit__c"])
add("compliancequest__SQX_Department__c", "Name", ["Name"],
    [{"Name": r["Name"]} for r in rows("Department")])
add("compliancequest__SQX_Standard_Service__c", "Name", ["Name", f"{CQ}Description__c"],
    [{"Name": r["Name"], f"{CQ}Description__c": r["Description"]} for r in rows("Process")])

# ---------- Accounts / contacts / suppliers / personnel ----------
acc = []
for r in rows("Accounts"):
    city, _, country = r["City"].partition(", ")
    acc.append({"Name": r["Account Name"],
                "RecordType.DeveloperName": "CQ_Supplier" if r["Type"] == "Supplier" else "Default",  # Default: the seed's own (force-app), deployed by grant-cq-access.ps1
                "BillingCity": city, "BillingState": country if len(country) == 2 and country not in ("CN", "KR", "TW") else "",
                "BillingCountry": {"CN": "China", "KR": "South Korea", "TW": "Taiwan"}.get(country, "United States"),
                "Description": r["Supplies / Buys"]})
add("Account", "Name", ["Name", "RecordType.DeveloperName", "BillingCity", "BillingState", "BillingCountry", "Description"],
    acc, query_fields=["Name", "RecordTypeId", "BillingCity", "BillingState", "BillingCountry", "Description"])

add("Contact", "Email", ["FirstName", "LastName", "Email", "Title", "Account.Name"],
    [{"FirstName": r["First Name"], "LastName": r["Last Name"], "Email": r["Email"], "Title": r["Title"],
      "Account.Name": r["Account"]} for r in rows("Contacts")],
    query_fields=["FirstName", "LastName", "Email", "Title", "AccountId"])

sup_cols = ["Name", f"{CQ}Supplier_Number__c", f"{CQ}SQX_Account__r.Name", f"{CQ}Type__c", f"{CQ}Supplier_Type__c",
            f"{CQ}Qualification_Status__c", f"{CQ}Supplier_QMS_Status__c", f"{CQ}Supplier_Supply_Chain_Status__c",
            f"{CQ}Status__c", f"{CQ}Active__c"]
add("compliancequest__SQX_Supplier__c", f"{CQ}Supplier_Number__c", sup_cols,
    [{"Name": r["Supplier Name"], f"{CQ}Supplier_Number__c": r["Supplier Name"],
      f"{CQ}SQX_Account__r.Name": r["Account"], f"{CQ}Type__c": r["Type"],
      f"{CQ}Supplier_Type__c": r["Supplier Type"], f"{CQ}Qualification_Status__c": r["Qualification Status"],
      f"{CQ}Supplier_QMS_Status__c": r["Supplier QMS Status"],
      f"{CQ}Supplier_Supply_Chain_Status__c": r["Supply Chain Status"],
      f"{CQ}Status__c": r["Status"], f"{CQ}Active__c": "Yes"} for r in rows("Suppliers")],
    query_fields=[c.replace("__r.Name", "__c") for c in sup_cols])

# Supplier Contacts: CQ checks that an NC's contact is linked to the NC's supplier through this object
SUPR = f"{CQ}SQX_Supplier__r.{CQ}Supplier_Number__c"
add("compliancequest__SQX_Supplier_Contact__c", f"{SUPR};{CQ}SQX_Contact__r.Email",
    [SUPR, f"{CQ}SQX_Contact__r.Email"],
    [{SUPR: r["Supplier"], f"{CQ}SQX_Contact__r.Email": r["Contact (Email)"]} for r in rows("Supplier Contacts")],
    query_fields=[f"{CQ}SQX_Supplier__c", f"{CQ}SQX_Contact__c"])

job = {"Quality Engineer": "Quality Engineer", "Assembly Lead": "Assembly Supervisor",
       "Receiving Inspector": "Quality Inspector", "Customer Service Rep": "Customer Service Representative",
       "Quality Manager": "Quality Manager"}
site_of_dept = {r["Name"]: r["Site"] for r in rows("Department")}
pers_cols = [f"{CQ}Identification_Number__c", f"{CQ}Full_Name__c", f"{CQ}Job_Title__c", f"{CQ}Personnel_Status__c",
             f"{CQ}Active__c", f"{CQ}Employment_Type__c", f"{CQ}Email_Address__c", f"{CQ}SQX_Department__r.Name",
             f"{CQ}SQX_Site__r.Name", f"{CQ}SQX_Business_Unit__r.Name"]
pers = []
for r in rows("Personnel"):
    site = site_of_dept[r["Department"]]
    pers.append({f"{CQ}Identification_Number__c": r["Identification Number"], f"{CQ}Full_Name__c": r["Full Name"],
                 f"{CQ}Job_Title__c": job[r["Job Title"]], f"{CQ}Personnel_Status__c": "Active",
                 f"{CQ}Active__c": "TRUE", f"{CQ}Employment_Type__c": "Full Time",
                 f"{CQ}Email_Address__c": r["Full Name"].lower().replace(" ", ".") + "@atlas-mfg.example.com",
                 f"{CQ}SQX_Department__r.Name": r["Department"], f"{CQ}SQX_Site__r.Name": site,
                 f"{CQ}SQX_Business_Unit__r.Name": "Atlas Manufacturing"})
add("compliancequest__SQX_Personnel__c", f"{CQ}Identification_Number__c", pers_cols, pers,
    query_fields=[c.replace("__r.Name", "__c") for c in pers_cols])

# ---------- Parts ----------
add("compliancequest__SQX_Part_Family__c", "Name", ["Name"], [{"Name": r["Part Family Name"]} for r in rows("Part Families")])

part_cols = [f"{CQ}Part_Number__c", "Name", "cqext__Title__c", f"{CQ}Part_Family__r.Name", f"{CQ}Active__c",
             f"{CQ}Part_Risk_Level__c", f"{CQ}Part_Type__c", f"{CQ}Part_Description__c"]
part_title = {}
part_family = {}
parts = []
for r in rows("Parts"):
    part_title[r["Part Number"]] = r["Title"]
    part_family[r["Part Number"]] = r["Part Family"]
    parts.append({f"{CQ}Part_Number__c": r["Part Number"], "Name": r["Title"], "cqext__Title__c": r["Title"],
                  f"{CQ}Part_Family__r.Name": r["Part Family"], f"{CQ}Active__c": tf(r["Active"]),
                  f"{CQ}Part_Risk_Level__c": r["Risk Level"], f"{CQ}Part_Type__c": "Sub Assembly",
                  f"{CQ}Part_Description__c": f"{r['Title']} supplied by {r['Supplier']}"})
add("compliancequest__SQX_Part__c", f"{CQ}Part_Number__c", part_cols, parts,
    query_fields=[c.replace("__r.Name", "__c") for c in part_cols])

add("compliancequest__SQX_Lot_Info__c", "Name",
    ["Name", f"{CQ}Part_Number__c", f"{CQ}Manufacturing_Date__c", f"{CQ}Expiration_Date__c"],
    [{"Name": r["Lot Number"], f"{CQ}Part_Number__c": r["Part"], f"{CQ}Manufacturing_Date__c": r["Manufactured Date"],
      f"{CQ}Expiration_Date__c": r["Expiration Date"]} for r in rows("Lot Info")])

# ---------- Codes ----------
code_name = {}
codes = []
for r in rows("Hierarchical Codes"):
    code_name[r["Code (Seed Key)"]] = r["Name"]
    if r["Used On"] == "Nonconformance":
        ctype = "Process Failure Code" if r["Code (Seed Key)"] == "NC-FW" else "Product Failure Code"
        rt = "Non_Complaint_Related"
    elif r["Used On"] == "Complaint":
        ctype, rt = "Complaint Code", "Complaint_Related"
    else:
        ctype, rt = "Complaint Conclusion", "Complaint_Related"
    codes.append({"Name": r["Name"], "RecordType.DeveloperName": rt, f"{CQ}Type__c": ctype,
                  f"{CQ}Defect_Category__c": r["Category"], f"{CQ}Level__c": r["Level"],
                  f"{CQ}Active__c": tf(r["Active"]), f"{CQ}Description__c": f"Atlas code {r['Code (Seed Key)']}"})
add("compliancequest__SQX_Defect_Code__c", "Name",
    ["Name", "RecordType.DeveloperName", f"{CQ}Type__c", f"{CQ}Defect_Category__c", f"{CQ}Level__c", f"{CQ}Active__c",
     f"{CQ}Description__c"], codes,
    query_fields=["Name", "RecordTypeId", f"{CQ}Type__c", f"{CQ}Defect_Category__c", f"{CQ}Level__c", f"{CQ}Active__c",
                  f"{CQ}Description__c"])

# ---------- PHR configuration ----------
add("compliancequest__SQX_Result_Type__c", "Name", ["Name"], [{"Name": "Yes/No"}])
rtv_cols = ["Name", f"{CQ}SQX_Result_Type__r.Name", "cqext__Response_Label__c", "cqext__Response_Value__c",
            f"{CQ}Color__c", "cqext__Language__c"]
add("compliancequest__SQX_Result_Type_Value__c", f"{CQ}SQX_Result_Type__r.Name;Name", rtv_cols,
    [{"Name": r["Value"], f"{CQ}SQX_Result_Type__r.Name": r["Result Type (Seed Key)"],
      "cqext__Response_Label__c": r["Response Label"], "cqext__Response_Value__c": r["Response Value"],
      f"{CQ}Color__c": r["Color"], "cqext__Language__c": "English"} for r in rows("Result Type")],
    query_fields=[c.replace("__r.Name", "__c") for c in rtv_cols])

task_cols = ["Name", "RecordType.DeveloperName", f"{CQ}Task_Type__c", f"{CQ}Record_Type__c", f"{CQ}Active__c",
             f"{CQ}Allowed_Days__c", f"{CQ}Description__c"]
add("compliancequest__SQX_Task__c", "Name", task_cols,
    [{"Name": r["Task Name (Seed Key)"], "RecordType.DeveloperName": "Complaint_Management_Step",
      f"{CQ}Task_Type__c": r["Task Type"], f"{CQ}Record_Type__c": r["Record Type (picklist)"],
      f"{CQ}Active__c": tf(r["Active"]), f"{CQ}Allowed_Days__c": r["Allowed Days"],
      f"{CQ}Description__c": r["Description"]} for r in rows("CQ Tasks")],
    query_fields=["Name", "RecordTypeId"] + task_cols[2:])

q_cols = ["Name", f"{CQ}SQX_Task__r.Name", "RecordType.DeveloperName", f"{CQ}QuestionText__c", f"{CQ}Answer_Type__c",
          f"{CQ}SQX_Result_Type__r.Name", f"{CQ}Picklist_Category__c", f"{CQ}Allow_Comment__c"]
add("compliancequest__SQX_Task_Question__c", f"{CQ}SQX_Task__r.Name;Name", q_cols,
    [{"Name": r["Question No. (Seed Key)"], f"{CQ}SQX_Task__r.Name": r["Task"],
      "RecordType.DeveloperName": "Simple_Question", f"{CQ}QuestionText__c": r["Question Text"],
      f"{CQ}Answer_Type__c": r["Answer Type"], f"{CQ}SQX_Result_Type__r.Name": r["Result Type"],
      f"{CQ}Picklist_Category__c": r["Picklist Category"], f"{CQ}Allow_Comment__c": tf(r["Allow Comment"])}
     for r in rows("PHR Questions")],
    query_fields=["Name", f"{CQ}SQX_Task__c", "RecordTypeId", f"{CQ}QuestionText__c", f"{CQ}Answer_Type__c",
                  f"{CQ}SQX_Result_Type__c", f"{CQ}Picklist_Category__c", f"{CQ}Allow_Comment__c"])

# ---------- Nonconformances ----------
nc_cols = [f"{CQ}NC_Title__c", f"{CQ}Type__c", f"{CQ}Type_Of_Issue__c", f"{CQ}Is_Part_Impacted__c",
           f"{CQ}SQX_Part__r.{CQ}Part_Number__c", f"{CQ}SQX_Service__r.Name", f"{CQ}SQX_Supplier__r.{CQ}Supplier_Number__c",
           f"{CQ}SQX_Account__r.Name", f"{CQ}SQX_External_Contact__r.Email", f"{CQ}SQX_Site__r.Name",
           f"{CQ}SQX_Division__r.Name", f"{CQ}SQX_Business_Unit__r.Name", f"{CQ}SQX_Department__r.Name",
           f"{CQ}Occurrence_Date__c", f"{CQ}Aware_Date__c", f"{CQ}Description__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
contact_email = {f"{r['First Name']} {r['Last Name']}": r["Email"] for r in rows("Contacts")}
ncs = []
for r in rows("Nonconformances"):
    ncs.append({f"{CQ}NC_Title__c": r["Title (Seed Key)"], f"{CQ}Type__c": r["NC Type"],
                f"{CQ}Type_Of_Issue__c": r["Issue Type"], f"{CQ}Is_Part_Impacted__c": r["Is Part Impacted?"],
                f"{CQ}SQX_Part__r.{CQ}Part_Number__c": r["Part"], f"{CQ}SQX_Service__r.Name": r["Process"],
                f"{CQ}SQX_Supplier__r.{CQ}Supplier_Number__c": r["Supplier"], f"{CQ}SQX_Account__r.Name": r["Account"],
                f"{CQ}SQX_External_Contact__r.Email": contact_email.get(r["Contact"], ""),
                f"{CQ}SQX_Site__r.Name": r["Site"], f"{CQ}SQX_Division__r.Name": site_division[r["Site"]],
                f"{CQ}SQX_Business_Unit__r.Name": "Atlas Manufacturing", f"{CQ}SQX_Department__r.Name": r["Department"],
                f"{CQ}Occurrence_Date__c": r["Occurrence Date"], f"{CQ}Aware_Date__c": r["Aware Date"],
                f"{CQ}Description__c": r["Description"], f"{CQ}Status__c": "Draft", f"{CQ}Record_Stage__c": "Draft"})
nc_q = [f"{CQ}NC_Title__c", f"{CQ}Type__c", f"{CQ}Type_Of_Issue__c", f"{CQ}Is_Part_Impacted__c", f"{CQ}SQX_Part__c",
        f"{CQ}SQX_Service__c", f"{CQ}SQX_Supplier__c", f"{CQ}SQX_Account__c", f"{CQ}SQX_External_Contact__c",
        f"{CQ}SQX_Site__c", f"{CQ}SQX_Division__c", f"{CQ}SQX_Business_Unit__c", f"{CQ}SQX_Department__c",
        f"{CQ}Occurrence_Date__c", f"{CQ}Aware_Date__c", f"{CQ}Description__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add("compliancequest__SQX_Nonconformance__c", f"{CQ}NC_Title__c", nc_cols, ncs, query_fields=nc_q)

NCR = f"{CQ}SQX_Nonconformance__r.{CQ}NC_Title__c"
ip_cols = [NCR, f"{CQ}SQX_Impacted_Part__r.{CQ}Part_Number__c", f"{CQ}Lot_Number__c", f"{CQ}Lot_Quantity__c"]
add("compliancequest__SQX_NC_Impacted_Product__c", f"{NCR};{CQ}Lot_Number__c", ip_cols,
    [{NCR: r["Nonconformance (Title)"], f"{CQ}SQX_Impacted_Part__r.{CQ}Part_Number__c": r["Impacted Part"],
      f"{CQ}Lot_Number__c": r["Lot Number"], f"{CQ}Lot_Quantity__c": r["Quantity"]} for r in rows("NC Impacted Products")],
    query_fields=[f"{CQ}SQX_Nonconformance__c", f"{CQ}SQX_Impacted_Part__c", f"{CQ}Lot_Number__c", f"{CQ}Lot_Quantity__c"])

nc_part = {r["Title (Seed Key)"]: r["Part"] for r in rows("Nonconformances")}
nc_lot = {r["Nonconformance (Title)"]: r["Lot Number"] for r in rows("NC Impacted Products")}
d_cols = [NCR, f"{CQ}SQX_Defect_Code__r.Name", f"{CQ}SQX_Part__r.{CQ}Part_Number__c", f"{CQ}Lot_Ser_Number__c",
          f"{CQ}Defective_Quantity__c", f"{CQ}Number_of_defects__c"]
add("compliancequest__SQX_NC_Defect__c", f"{NCR};{CQ}SQX_Defect_Code__r.Name", d_cols,
    [{NCR: r["Nonconformance (Title)"], f"{CQ}SQX_Defect_Code__r.Name": code_name[r["Failure Code"]],
      f"{CQ}SQX_Part__r.{CQ}Part_Number__c": nc_part[r["Nonconformance (Title)"]],
      f"{CQ}Lot_Ser_Number__c": nc_lot.get(r["Nonconformance (Title)"], ""),
      f"{CQ}Defective_Quantity__c": r["Defective Quantity"], f"{CQ}Number_of_defects__c": r["Number of Defects"]}
     for r in rows("NC Defects")],
    query_fields=[f"{CQ}SQX_Nonconformance__c", f"{CQ}SQX_Defect_Code__c", f"{CQ}SQX_Part__c", f"{CQ}Lot_Ser_Number__c",
                  f"{CQ}Defective_Quantity__c", f"{CQ}Number_of_defects__c"])

c_cols = [f"{CQ}Title__c", NCR, f"{CQ}Description__c", f"{CQ}Due_Date__c", f"{CQ}Status__c", f"{CQ}SQX_User__c"]
add("compliancequest__SQX_Containment__c", f"{CQ}Title__c", c_cols,
    [{f"{CQ}Title__c": r["Containment Action"][:255], NCR: r["Nonconformance (Title)"],
      f"{CQ}Description__c": r["Containment Action"], f"{CQ}Due_Date__c": DUE_DATE, f"{CQ}Status__c": "Draft",
      f"{CQ}SQX_User__c": LOAD_USER}   # Assignee: CQ's trigger copies it to the owner, which can't be blank
     for r in rows("Containments")],
    query_fields=[f"{CQ}Title__c", f"{CQ}SQX_Nonconformance__c", f"{CQ}Description__c", f"{CQ}Due_Date__c", f"{CQ}Status__c",
                  f"{CQ}SQX_User__c"])

# ---------- Complaints ----------
CR = f"{CQ}SQX_Complaint__r.{CQ}Complaint_Title__c"
cmp_cols = [f"{CQ}Complaint_Title__c", f"{CQ}SQX_Account__r.Name", f"{CQ}SQX_External_Contact__r.Email",
            f"{CQ}SQX_Part__r.{CQ}Part_Number__c", f"{CQ}SQX_Part_Family__r.Name", f"{CQ}SQX_Lot_Info__r.Name",
            f"{CQ}Lot_Number__c", f"{CQ}SQX_Complaint_Code__r.Name", f"{CQ}Occurrence_Date__c", f"{CQ}Aware_Date__c",
            f"{CQ}Reported_Date__c", f"{CQ}Complaint_Quantity__c", f"{CQ}Country_of_Origin__c",
            f"{CQ}Complaint_Record_Type__c", f"{CQ}SQX_Site__r.Name", f"{CQ}SQX_Division__r.Name",
            f"{CQ}SQX_Business_Unit__r.Name", f"{CQ}Description_As_Reported__c", f"{CQ}Description__c",
            f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
cmps = []
cmp_part = {}
for r in rows("Complaints"):
    cmp_part[r["Title (Seed Key)"]] = (r["Part"], r["Lot Number"])
    cmps.append({f"{CQ}Complaint_Title__c": r["Title (Seed Key)"], f"{CQ}SQX_Account__r.Name": r["Account"],
                 f"{CQ}SQX_External_Contact__r.Email": contact_email[r["Contact"]],
                 f"{CQ}SQX_Part__r.{CQ}Part_Number__c": r["Part"], f"{CQ}SQX_Part_Family__r.Name": part_family[r["Part"]],
                 f"{CQ}SQX_Lot_Info__r.Name": r["Lot Number"], f"{CQ}Lot_Number__c": r["Lot Number"],
                 f"{CQ}SQX_Complaint_Code__r.Name": code_name[r["Complaint Code"]],
                 f"{CQ}Occurrence_Date__c": r["Occurrence Date"], f"{CQ}Aware_Date__c": r["Aware Date"],
                 f"{CQ}Reported_Date__c": r["Reported Date"], f"{CQ}Complaint_Quantity__c": r["Complaint Quantity"],
                 f"{CQ}Country_of_Origin__c": "US", f"{CQ}Complaint_Record_Type__c": "Product Quality",
                 f"{CQ}SQX_Site__r.Name": r["Site"], f"{CQ}SQX_Division__r.Name": site_division[r["Site"]],
                 f"{CQ}SQX_Business_Unit__r.Name": "Atlas Manufacturing",
                 f"{CQ}Description_As_Reported__c": r["Description As Reported"], f"{CQ}Description__c": r["Description"],
                 f"{CQ}Status__c": "Draft", f"{CQ}Record_Stage__c": "Draft"})
cmp_q = [f"{CQ}Complaint_Title__c", f"{CQ}SQX_Account__c", f"{CQ}SQX_External_Contact__c", f"{CQ}SQX_Part__c",
         f"{CQ}SQX_Part_Family__c", f"{CQ}SQX_Lot_Info__c", f"{CQ}Lot_Number__c", f"{CQ}SQX_Complaint_Code__c",
         f"{CQ}Occurrence_Date__c", f"{CQ}Aware_Date__c", f"{CQ}Reported_Date__c", f"{CQ}Complaint_Quantity__c",
         f"{CQ}Country_of_Origin__c", f"{CQ}Complaint_Record_Type__c", f"{CQ}SQX_Site__c", f"{CQ}SQX_Division__c",
         f"{CQ}SQX_Business_Unit__c", f"{CQ}Description_As_Reported__c", f"{CQ}Description__c", f"{CQ}Status__c",
         f"{CQ}Record_Stage__c"]
add("compliancequest__SQX_Complaint__c", f"{CQ}Complaint_Title__c", cmp_cols, cmps, query_fields=cmp_q)

cd_cols = [CR, f"{CQ}SQX_Defect_Code__r.Name", f"{CQ}SQX_Part__r.{CQ}Part_Number__c", f"{CQ}Lot_Ser_Number__c",
           f"{CQ}Defective_Quantity__c"]
add("compliancequest__SQX_Complaint_Defect__c", f"{CR};{CQ}SQX_Defect_Code__r.Name", cd_cols,
    [{CR: r["Complaint (Title)"], f"{CQ}SQX_Defect_Code__r.Name": code_name[r["Complaint Code"]],
      f"{CQ}SQX_Part__r.{CQ}Part_Number__c": cmp_part[r["Complaint (Title)"]][0],
      f"{CQ}Lot_Ser_Number__c": cmp_part[r["Complaint (Title)"]][1],
      f"{CQ}Defective_Quantity__c": r["Defective Quantity"]} for r in rows("Complaint Defects")],
    query_fields=[f"{CQ}SQX_Complaint__c", f"{CQ}SQX_Defect_Code__c", f"{CQ}SQX_Part__c", f"{CQ}Lot_Ser_Number__c",
                  f"{CQ}Defective_Quantity__c"])

# ---------- Audits ----------
DOCR = f"{CQ}SQX_Controlled_Document__r.{CQ}Document_Number__c"
CRITR = f"{CQ}SQX_Audit_Criteria_Document__r.{CQ}Document_Number__c"
PROGR = f"{CQ}SQX_Audit_Program__r.{CQ}Title__c"
AUDR = f"{CQ}SQX_Audit__r.{CQ}Title__c"

# Controlled Documents: the audit criteria (Audit_Criteria record type, Draft: requirements can only be added to a
# Draft document; release them by hand afterwards) and the Atlas SOPs, work instructions and inspection criteria
# from the Controlled Documents tab (released ones load as Current). One SFDMU object for both, keyed on number.
# Files are not loaded here: link-document-files.ps1 uploads files/<number>.pdf and sets Primary Content.
RT_NAME = {"Controlled Document": "Controlled_Document", "Inspection Criteria": "Inspection_Criteria",
           "Audit Criteria": "Audit_Criteria"}
ac_cols = [f"{CQ}Document_Number__c", f"{CQ}Title__c", "RecordType.DeveloperName", f"{CQ}Document_Status__c",
           f"{CQ}Revision__c", f"{CQ}Document_Category__c", f"{CQ}SQX_Business_Unit__r.Name", f"{CQ}Description__c",
           f"{CQ}SQX_Site__r.Name", f"{CQ}SQX_Department__r.Name", f"{CQ}SQX_Part__r.{CQ}Part_Number__c",
           f"{CQ}Effective_Date__c", f"{CQ}Review_Interval__c", f"{CQ}Duration__c"]
docs = [{f"{CQ}Document_Number__c": r["Document Number (Seed Key)"], f"{CQ}Title__c": r["Title"],
         "RecordType.DeveloperName": "Audit_Criteria", f"{CQ}Document_Status__c": r["Document Status"],
         f"{CQ}Revision__c": r["Revision"], f"{CQ}Document_Category__c": r["Document Category"],
         f"{CQ}SQX_Business_Unit__r.Name": r["Business Unit"], f"{CQ}Description__c": r["Description"]}
        for r in rows("Audit Criteria")]
for r in rows("Controlled Documents"):
    eff = r["Effective Date"]
    docs.append({f"{CQ}Document_Number__c": r["Document Number (Seed Key)"], f"{CQ}Title__c": r["Title"],
                 "RecordType.DeveloperName": RT_NAME[r["Record Type"]], f"{CQ}Document_Status__c": r["Document Status"],
                 f"{CQ}Revision__c": r["Revision"], f"{CQ}Document_Category__c": r["Document Category"],
                 f"{CQ}SQX_Business_Unit__r.Name": r["Business Unit"], f"{CQ}Description__c": r["Description"],
                 f"{CQ}SQX_Site__r.Name": r["Site"], f"{CQ}SQX_Department__r.Name": r["Department"],
                 f"{CQ}SQX_Part__r.{CQ}Part_Number__c": r["Product (Part Number)"],
                 f"{CQ}Effective_Date__c": str(eff)[:10] if eff else "",
                 f"{CQ}Review_Interval__c": r["Review Interval (days)"], f"{CQ}Duration__c": r["Days to Complete Training"]})
add("compliancequest__SQX_Controlled_Document__c", f"{CQ}Document_Number__c", ac_cols, docs,
    query_fields=[f"{CQ}Document_Number__c", f"{CQ}Title__c", "RecordTypeId", f"{CQ}Document_Status__c",
                  f"{CQ}Revision__c", f"{CQ}Document_Category__c", f"{CQ}SQX_Business_Unit__c", f"{CQ}Description__c",
                  f"{CQ}SQX_Site__c", f"{CQ}SQX_Department__c", f"{CQ}SQX_Part__c",
                  f"{CQ}Effective_Date__c", f"{CQ}Review_Interval__c", f"{CQ}Duration__c"])

req_cols = [DOCR, f"{CQ}Section__c", f"{CQ}Topic__c", f"{CQ}Objective__c", f"{CQ}Result_Type__c"]
add("compliancequest__SQX_Doc_Criterion_Requirement__c", f"{DOCR};{CQ}Topic__c", req_cols,
    [{DOCR: r["Audit Criteria (Document Number)"], f"{CQ}Section__c": r["Section"], f"{CQ}Topic__c": r["Topic"],
      f"{CQ}Objective__c": r["Objective"], f"{CQ}Result_Type__c": r["Result Type"]}
     for r in rows("Criterion Requirements")],
    query_fields=[f"{CQ}SQX_Controlled_Document__c"] + req_cols[1:])

prog_cols = [f"{CQ}Title__c", f"{CQ}Audit_Program_Year__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c", f"{CQ}Objective__c"]
add("compliancequest__SQX_Audit_Program__c", f"{CQ}Title__c", prog_cols,
    [{f"{CQ}Title__c": r["Title (Seed Key)"], f"{CQ}Audit_Program_Year__c": str(r["Audit Program Year"]),
      f"{CQ}Status__c": r["Status"], f"{CQ}Record_Stage__c": r["Stage"], f"{CQ}Objective__c": r["Objective"]}
     for r in rows("Audit Program")])

add("compliancequest__SQX_Audit_Program_Scope__c", f"{PROGR};{CRITR}", [PROGR, CRITR],
    [{PROGR: r["Audit Program"], CRITR: r["Audit Criteria (Document Number)"]} for r in rows("Audit Program Scope")],
    query_fields=[f"{CQ}SQX_Audit_Program__c", f"{CQ}SQX_Audit_Criteria_Document__c"])

# No record type on audits: "Standard_Audit" in the dev sandbox is org-made, not CQ's, so audits take the org default.
# Auto_Set_Title is off so CQ keeps our title (it is the external ID).
aud_cols = [f"{CQ}Title__c", f"{CQ}Auto_Set_Title__c", f"{CQ}Audit_Type__c", f"{CQ}Audit_Category__c",
            f"{CQ}Audit_Format__c", f"{CQ}Status__c", f"{CQ}Stage__c", PROGR, CRITR, f"{CQ}Audit_Year__c",
            f"{CQ}SQX_Business_Unit__r.Name", f"{CQ}SQX_Division__r.Name", f"{CQ}SQX_Site__r.Name",
            f"{CQ}SQX_Department__r.Name", f"{CQ}Account__r.Name", f"{CQ}SQX_Supplier__r.{CQ}Supplier_Number__c",
            f"{CQ}Start_Date__c", f"{CQ}End_Date__c", f"{CQ}Scope_Description__c"]
add("compliancequest__SQX_Audit__c", f"{CQ}Title__c", aud_cols,
    [{f"{CQ}Title__c": r["Title (Seed Key)"], f"{CQ}Auto_Set_Title__c": "FALSE", f"{CQ}Audit_Type__c": r["Audit Type"],
      f"{CQ}Audit_Category__c": r["Audit Category"], f"{CQ}Audit_Format__c": r["Audit Format"],
      f"{CQ}Status__c": r["Status"], f"{CQ}Stage__c": r["Stage"], PROGR: r["Audit Program"],
      CRITR: r["Audit Criteria (Document Number)"], f"{CQ}Audit_Year__c": str(r["Plan Year"]),
      f"{CQ}SQX_Business_Unit__r.Name": r["Business Unit"], f"{CQ}SQX_Division__r.Name": r["Division"],
      f"{CQ}SQX_Site__r.Name": r["Site"], f"{CQ}SQX_Department__r.Name": r["Department"],
      f"{CQ}Account__r.Name": r["Account"], f"{CQ}SQX_Supplier__r.{CQ}Supplier_Number__c": r["Supplier"],
      f"{CQ}Start_Date__c": str(r["Start Date"])[:10], f"{CQ}End_Date__c": str(r["End Date"])[:10],
      f"{CQ}Scope_Description__c": r["Scope/Description"]} for r in rows("Audits")],
    query_fields=[c.replace("__r.Name", "__c").replace(f"__r.{CQ}Supplier_Number__c", "__c")
                  .replace(f"__r.{CQ}Title__c", "__c").replace(f"__r.{CQ}Document_Number__c", "__c") for c in aud_cols])

# ---------- Equipment ----------
# Equipment number (Name) is the key. Calibration status is carried on the equipment (Last/Next Calibration
# Date). No calibration records are loaded: they get added in CQ. Procedures point at Controlled Documents.
EQR = f"{CQ}SQX_Equipment__r.Name"
eq_cols = ["Name", f"{CQ}Equipment_Description__c", f"{CQ}Equipment_Type__c", f"{CQ}Equipment_Status__c",
           f"{CQ}Manufacturer__c", f"{CQ}Model_Number__c", f"{CQ}Serial_Number__c", f"{CQ}Equipment_Location__c",
           f"{CQ}SQX_Site__r.Name", f"{CQ}SQX_Department__r.Name", f"{CQ}SQX_Business_Unit__r.Name",
           f"{CQ}Requires_Periodic_Calibration__c", f"{CQ}Calibration_Interval__c", f"{CQ}Calibration_Source__c",
           f"{CQ}SQX_Calibration_Supplier__r.Name", f"{CQ}Last_Calibration_Date__c", f"{CQ}Next_Calibration_Date__c",
           f"{CQ}SQX_Applicable_Procedure_Rev__r.{CQ}Document_Number__c"]


def d10(v):
    return str(v)[:10] if v else ""


add("compliancequest__SQX_Equipment__c", "Name", eq_cols,
    [{"Name": r["Equipment Number (Seed Key)"], f"{CQ}Equipment_Description__c": r["Description"],
      f"{CQ}Equipment_Type__c": r["Equipment Type"], f"{CQ}Equipment_Status__c": r["Equipment Status"],
      f"{CQ}Manufacturer__c": r["Manufacturer"], f"{CQ}Model_Number__c": r["Model Number"],
      f"{CQ}Serial_Number__c": r["Serial Number"], f"{CQ}Equipment_Location__c": r["Location"],
      f"{CQ}SQX_Site__r.Name": r["Site"], f"{CQ}SQX_Department__r.Name": r["Department"],
      f"{CQ}SQX_Business_Unit__r.Name": "Atlas Manufacturing",
      f"{CQ}Requires_Periodic_Calibration__c": tf(r["Requires Periodic Calibration"]),
      f"{CQ}Calibration_Interval__c": r["Calibration Interval (days)"], f"{CQ}Calibration_Source__c": r["Calibration Source"],
      f"{CQ}SQX_Calibration_Supplier__r.Name": r["Calibration Lab (Account)"],
      f"{CQ}Last_Calibration_Date__c": d10(r["Last Calibration Date"]), f"{CQ}Next_Calibration_Date__c": d10(r["Next Calibration Date"]),
      f"{CQ}SQX_Applicable_Procedure_Rev__r.{CQ}Document_Number__c": r["Applicable Procedure (Document Number)"]}
     for r in rows("Equipment")],
    query_fields=["Name", f"{CQ}Equipment_Description__c", f"{CQ}Equipment_Type__c", f"{CQ}Equipment_Status__c",
                  f"{CQ}Manufacturer__c", f"{CQ}Model_Number__c", f"{CQ}Serial_Number__c", f"{CQ}Equipment_Location__c",
                  f"{CQ}SQX_Site__c", f"{CQ}SQX_Department__c", f"{CQ}SQX_Business_Unit__c",
                  f"{CQ}Requires_Periodic_Calibration__c", f"{CQ}Calibration_Interval__c", f"{CQ}Calibration_Source__c",
                  f"{CQ}SQX_Calibration_Supplier__c", f"{CQ}Last_Calibration_Date__c", f"{CQ}Next_Calibration_Date__c",
                  f"{CQ}SQX_Applicable_Procedure_Rev__c"])

sch_cols = [EQR, f"{CQ}Uniqueness_Constraint__c", f"{CQ}Active__c", f"{CQ}Recurring_Event__c", f"{CQ}Recurring_Interval__c",
            f"{CQ}Recurring_Unit__c", f"{CQ}Schedule_Basis__c", f"{CQ}Last_Performed_Date__c", f"{CQ}Next_Due_Date__c",
            f"{CQ}SQX_Applicable_Procedure__r.{CQ}Document_Number__c", f"{CQ}Comments__c"]
add("compliancequest__SQX_Equipment_Event_Schedule__c", f"{EQR};{CQ}Uniqueness_Constraint__c", sch_cols,
    [{EQR: r["Equipment (Seed Key)"], f"{CQ}Uniqueness_Constraint__c": r["Event Name"], f"{CQ}Active__c": tf(r["Active"]),
      f"{CQ}Recurring_Event__c": "TRUE", f"{CQ}Recurring_Interval__c": r["Recurring Interval"],
      f"{CQ}Recurring_Unit__c": r["Recurring Unit"], f"{CQ}Schedule_Basis__c": r["Schedule Basis"],
      f"{CQ}Last_Performed_Date__c": d10(r["Last Performed Date"]), f"{CQ}Next_Due_Date__c": d10(r["Next Due Date"]),
      f"{CQ}SQX_Applicable_Procedure__r.{CQ}Document_Number__c": r["Applicable Procedure (Document Number)"],
      f"{CQ}Comments__c": r["Comments"]}
     for r in rows("Equipment Schedules")],
    query_fields=[f"{CQ}SQX_Equipment__c", f"{CQ}Uniqueness_Constraint__c", f"{CQ}Active__c", f"{CQ}Recurring_Event__c",
                  f"{CQ}Recurring_Interval__c", f"{CQ}Recurring_Unit__c", f"{CQ}Schedule_Basis__c",
                  f"{CQ}Last_Performed_Date__c", f"{CQ}Next_Due_Date__c", f"{CQ}SQX_Applicable_Procedure__c", f"{CQ}Comments__c"])

# ---------- Inspections ----------
# Specifications sit under the Inspection Criteria documents (key: Specification Number). CQ only accepts them on a
# Draft criteria document, so the IC-REC documents load as Draft. Inspections (key: Lot Number) load as Draft with
# no criteria and no detail lines: criteria get added in CQ when an inspection is worked (CQ refuses detail lines on a
# closed inspection, so none are loaded closed). Create NC If Inspection Fails is off; failed lots point at their NCs.
from datetime import datetime
from zoneinfo import ZoneInfo


def qf(cols):
    # query field for each column: a lookup's __r.<key> column queries the lookup's Id field
    return ["RecordTypeId" if c.startswith("RecordType.") else c.split(".")[0][:-3] + "__c" if "." in c else c for c in cols]


def central_dt(v):
    # "2026-10-21 09:00" Central time -> UTC date-time
    if not v:
        return ""
    t = datetime.strptime(str(v)[:16], "%Y-%m-%d %H:%M").replace(tzinfo=ZoneInfo("America/Chicago"))
    return t.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%S.000Z")


def num(v):
    return "" if v in ("", None) else v


CRIT = f"{CQ}SQX_Inspection_Criteria__r.{CQ}Document_Number__c"
spec_cols = [f"{CQ}Specification_Number__c", CRIT, f"{CQ}Characteristics__c", f"{CQ}Specification__c",
             f"{CQ}Lower_Spec_Limit__c", f"{CQ}Upper_Spec_Limit__c", f"{CQ}UOM__c", f"{CQ}Impact__c",
             f"{CQ}Inspection_Method__c", EQR]
spec_by_num = {}
specs = []
for r in rows("Specifications"):
    s = {f"{CQ}Specification_Number__c": r["Specification Number (Seed Key)"],
         CRIT: r["Inspection Criteria (Document Number)"], f"{CQ}Characteristics__c": r["Characteristic"],
         f"{CQ}Specification__c": r["Specification"], f"{CQ}Lower_Spec_Limit__c": num(r["Lower Spec Limit"]),
         f"{CQ}Upper_Spec_Limit__c": num(r["Upper Spec Limit"]), f"{CQ}UOM__c": r["UOM"], f"{CQ}Impact__c": r["Impact"],
         f"{CQ}Inspection_Method__c": r["Inspection Method"], EQR: r["Equipment"]}
    specs.append(s)
    spec_by_num[s[f"{CQ}Specification_Number__c"]] = s
add("compliancequest__SQX_Specification__c", f"{CQ}Specification_Number__c", spec_cols, specs, query_fields=qf(spec_cols))

account_of_supplier = {r["Supplier Name"]: r["Account"] for r in rows("Suppliers")}
ins_cols = [f"{CQ}Lot_Number__c", f"{CQ}SQX_Part__r.{CQ}Part_Number__c", f"{CQ}SQX_Supplier__r.{CQ}Supplier_Number__c",
            f"{CQ}SQX_Supplier_Number__r.Name", f"{CQ}Inspection_Source__c", f"{CQ}Inspection_Type__c",
            f"{CQ}SQX_Department__r.Name", f"{CQ}Lot_Quantity__c", f"{CQ}Sample_Size__c", f"{CQ}Status__c",
            f"{CQ}Initiated_Date_Time__c", f"{CQ}SQX_NC_Number__r.{CQ}NC_Title__c",
            f"{CQ}Create_NC_If_Inspection_Fails__c", f"{CQ}Comment__c"]
add("compliancequest__SQX_Inspection__c", f"{CQ}Lot_Number__c", ins_cols,
    [{f"{CQ}Lot_Number__c": r["Lot Number (Seed Key)"], f"{CQ}SQX_Part__r.{CQ}Part_Number__c": r["Part"],
      f"{CQ}SQX_Supplier__r.{CQ}Supplier_Number__c": r["Supplier"],
      f"{CQ}SQX_Supplier_Number__r.Name": account_of_supplier[r["Supplier"]],
      f"{CQ}Inspection_Source__c": r["Inspection Source"], f"{CQ}Inspection_Type__c": r["Inspection Type"],
      f"{CQ}SQX_Department__r.Name": r["Department"], f"{CQ}Lot_Quantity__c": r["Lot Quantity"],
      f"{CQ}Sample_Size__c": r["Sample Size"], f"{CQ}Status__c": r["Status"],
      f"{CQ}Initiated_Date_Time__c": central_dt(r["Initiated"]),
      f"{CQ}SQX_NC_Number__r.{CQ}NC_Title__c": r["Nonconformance (Title)"],
      f"{CQ}Create_NC_If_Inspection_Fails__c": "FALSE", f"{CQ}Comment__c": r["Comment"]}
     for r in rows("Inspections")],
    query_fields=qf(ins_cols))

# ---------- Supplier records, CAPA, changes, training, risks, meetings, projects ----------
# Every workflow record loads at its starting point (Draft). CQ names the supplier records itself unless
# "Prevent Name Update" is set, so it is set to keep the seed's names (the Seed Key). Training requirements and
# personnel job functions load inactive (CQ's default): activating them in CQ creates the document training.
SUPN = f"{CQ}SQX_Supplier__r.{CQ}Supplier_Number__c"
PARTN = f"{CQ}SQX_Part__r.{CQ}Part_Number__c"
DRAFT = {f"{CQ}Status__c": "Draft", f"{CQ}Record_Stage__c": "Draft"}


def add_q(sobject, key, cols, data, operation="Upsert"):
    add(sobject, key, cols, data, query_fields=qf(cols), operation=operation)


add_q("compliancequest__SQX_Supplier_Part__c", f"{CQ}Account__r.Name;{CQ}Part__r.{CQ}Part_Number__c",
      [f"{CQ}Account__r.Name", f"{CQ}Part__r.{CQ}Part_Number__c", SUPN, f"{CQ}Supplier_Reference_Part_Number__c",
       f"{CQ}Active__c", f"{CQ}Approved__c"],
      [{f"{CQ}Account__r.Name": r["Supplier Account"], f"{CQ}Part__r.{CQ}Part_Number__c": r["Part"],
        SUPN: r["Supplier Account"], f"{CQ}Supplier_Reference_Part_Number__c": r["Supplier Reference Part Number"],
        f"{CQ}Active__c": tf(r["Active"]), f"{CQ}Approved__c": tf(r["Approved"])} for r in rows("Supplier Parts")])

sd_cols = ["Name", f"{CQ}Prevent_Name_Update_Flag__c", SUPN, f"{CQ}SQX_Account__r.Name", PARTN, f"{CQ}Deviation_Category__c",
           f"{CQ}Affected_Lots__c", f"{CQ}Affected_Quantities__c", f"{CQ}Effective_Date__c", f"{CQ}Expiration_Date__c",
           f"{CQ}Reason_for_Deviation__c", f"{CQ}Scope__c", f"{CQ}Account_Type__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_Supplier_Deviation__c", "Name", sd_cols,
      [{"Name": r["Name (Seed Key)"], f"{CQ}Prevent_Name_Update_Flag__c": "TRUE", SUPN: r["Supplier"],
        f"{CQ}SQX_Account__r.Name": account_of_supplier[r["Supplier"]], PARTN: r["Part"],
        f"{CQ}Deviation_Category__c": r["Deviation Category"], f"{CQ}Affected_Lots__c": r["Affected Lots"],
        f"{CQ}Affected_Quantities__c": r["Affected Quantity"], f"{CQ}Effective_Date__c": d10(r["Effective Date"]),
        f"{CQ}Expiration_Date__c": d10(r["Expiration Date"]), f"{CQ}Reason_for_Deviation__c": r["Reason for Deviation"],
        f"{CQ}Scope__c": r["Scope"], f"{CQ}Account_Type__c": "CQ_Supplier", **DRAFT} for r in rows("Supplier Deviations")])

se_cols = ["Name", f"{CQ}Prevent_Name_Update__c", SUPN, f"{CQ}SQX_Account__r.Name", PARTN, f"{CQ}Subject__c",
           f"{CQ}Escalation_Level__c", f"{CQ}Target_Escalation_Level__c", f"{CQ}Rationale_for_Escalation__c",
           f"{CQ}Account_Type__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_Supplier_Escalation__c", "Name", se_cols,
      [{"Name": r["Name (Seed Key)"], f"{CQ}Prevent_Name_Update__c": "TRUE", SUPN: r["Supplier"],
        f"{CQ}SQX_Account__r.Name": account_of_supplier[r["Supplier"]], PARTN: r["Part"], f"{CQ}Subject__c": r["Subject"],
        f"{CQ}Escalation_Level__c": str(r["Escalation Level"]), f"{CQ}Target_Escalation_Level__c": str(r["Target Escalation Level"]),
        f"{CQ}Rationale_for_Escalation__c": r["Rationale for Escalation"], f"{CQ}Account_Type__c": "CQ_Supplier", **DRAFT}
       for r in rows("Supplier Escalations")])

sq_cols = ["Name", f"{CQ}Prevent_Name_Update_Flag__c", f"{CQ}Company_Name__c", f"{CQ}Supplier_Name__c", f"{CQ}First_Name__c",
           f"{CQ}Last_Name__c", f"{CQ}Job_Title__c", f"{CQ}Location__c", PARTN, f"{CQ}Justification_for_SI__c",
           f"{CQ}Account_Type__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_New_Supplier_Introduction__c", "Name", sq_cols,
      [{"Name": r["Name (Seed Key)"], f"{CQ}Prevent_Name_Update_Flag__c": "TRUE", f"{CQ}Company_Name__c": r["Company Name"],
        f"{CQ}Supplier_Name__c": r["Company Name"], f"{CQ}First_Name__c": r["Contact First Name"],
        f"{CQ}Last_Name__c": r["Contact Last Name"], f"{CQ}Job_Title__c": r["Contact Job Title"],
        f"{CQ}Location__c": r["Location"], PARTN: r["Part"], f"{CQ}Justification_for_SI__c": r["Justification"],
        f"{CQ}Account_Type__c": "CQ_Supplier", **DRAFT} for r in rows("Supplier Qualifications")])

si_cols = ["Name", f"{CQ}Prevent_Name_Update__c", SUPN, f"{CQ}SQX_Account__r.Name", PARTN, f"{CQ}Interaction_Reason__c",
           f"{CQ}Renew_Due_Date__c", f"{CQ}Scope__c", f"{CQ}Account_Type__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_Supplier_Interaction__c", "Name", si_cols,
      [{"Name": r["Name (Seed Key)"], f"{CQ}Prevent_Name_Update__c": "TRUE", SUPN: r["Supplier"],
        f"{CQ}SQX_Account__r.Name": account_of_supplier[r["Supplier"]], PARTN: r["Part"],
        f"{CQ}Interaction_Reason__c": r["Interaction Reason"], f"{CQ}Renew_Due_Date__c": d10(r["Renew Due Date"]),
        f"{CQ}Scope__c": r["Scope"], f"{CQ}Account_Type__c": "CQ_Supplier", **DRAFT} for r in rows("Supplier Interactions")])

capa_cols = [f"{CQ}Title__c", "RecordType.DeveloperName", f"{CQ}CAPA_Type__c", SUPN, PARTN, f"{CQ}SQX_Site__r.Name",
             f"{CQ}SQX_Division__r.Name", f"{CQ}SQX_Business_Unit__r.Name", f"{CQ}SQX_Department__r.Name", f"{CQ}Severity__c",
             f"{CQ}Probability__c", f"{CQ}Target_Due_Date__c", f"{CQ}Description__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_CAPA__c", f"{CQ}Title__c", capa_cols,
      [{f"{CQ}Title__c": r["Title (Seed Key)"], "RecordType.DeveloperName": "CAPA", f"{CQ}CAPA_Type__c": r["CAPA Type"],
        SUPN: r["Supplier"], PARTN: r["Part"], f"{CQ}SQX_Site__r.Name": r["Site"],
        f"{CQ}SQX_Division__r.Name": site_division[r["Site"]], f"{CQ}SQX_Business_Unit__r.Name": "Atlas Manufacturing",
        f"{CQ}SQX_Department__r.Name": r["Department"], f"{CQ}Severity__c": r["Severity"],
        f"{CQ}Probability__c": r["Probability"], f"{CQ}Target_Due_Date__c": d10(r["Target Due Date"]),
        f"{CQ}Description__c": r["Description"], **DRAFT} for r in rows("CAPAs")])

cr_cols = [f"{CQ}Title__c", f"{CQ}Change_Category__c", f"{CQ}Priority__c", f"{CQ}SQX_Department__r.Name", PARTN,
           f"{CQ}SQX_Process__r.Name", f"{CQ}Description__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_Change_Request__c", f"{CQ}Title__c", cr_cols,
      [{f"{CQ}Title__c": r["Title (Seed Key)"], f"{CQ}Change_Category__c": r["Change Category"], f"{CQ}Priority__c": r["Priority"],
        f"{CQ}SQX_Department__r.Name": r["Department"], PARTN: r["Part"], f"{CQ}SQX_Process__r.Name": r["Process"],
        f"{CQ}Description__c": r["Description"], **DRAFT} for r in rows("Change Requests")])

add_q("compliancequest__SQX_Job_Function__c", "Name", ["Name", f"{CQ}Description__c", f"{CQ}Active__c"],
      [{"Name": r["Name (Seed Key)"], f"{CQ}Description__c": r["Description"], f"{CQ}Active__c": "TRUE"}
       for r in rows("Job Functions")])

REQDOC = f"{CQ}SQX_Controlled_Document__r.{CQ}Document_Number__c"
JFN = f"{CQ}SQX_Job_Function__r.Name"
add_q("compliancequest__SQX_Requirement__c", f"{REQDOC};{JFN}",
      [REQDOC, JFN, f"{CQ}Level_of_Competency__c", f"{CQ}Active__c", f"{CQ}Optional__c"],
      [{REQDOC: r["Controlled Document (Number)"], JFN: r["Job Function"], f"{CQ}Level_of_Competency__c": r["Level of Competency"],
        f"{CQ}Active__c": "FALSE", f"{CQ}Optional__c": "FALSE"} for r in rows("Training Requirements")])

PERSN = f"{CQ}SQX_Personnel__r.{CQ}Identification_Number__c"
add_q("compliancequest__SQX_Personnel_Job_Function__c", f"{PERSN};{JFN}", [PERSN, JFN, f"{CQ}Active__c"],
      [{PERSN: r["Personnel (Identification Number)"], JFN: r["Job Function"], f"{CQ}Active__c": "FALSE"}
       for r in rows("Personnel Job Functions")])

risk_cols = [f"{CQ}Risk_Title__c", f"{CQ}Risk_Category__c", f"{CQ}Reported_Priority__c", f"{CQ}Identified_Date__c", PARTN,
             f"{CQ}SQX_Process__r.Name", f"{CQ}SQX_Department__r.Name", SUPN, f"{CQ}Risk_Description__c",
             f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_Risk__c", f"{CQ}Risk_Title__c", risk_cols,
      [{f"{CQ}Risk_Title__c": r["Risk Title (Seed Key)"], f"{CQ}Risk_Category__c": r["Risk Category"],
        f"{CQ}Reported_Priority__c": r["Reported Priority"], f"{CQ}Identified_Date__c": d10(r["Identified Date"]),
        PARTN: r["Part"], f"{CQ}SQX_Process__r.Name": r["Process"], f"{CQ}SQX_Department__r.Name": r["Department"],
        SUPN: r["Supplier"], f"{CQ}Risk_Description__c": r["Risk Description"], **DRAFT} for r in rows("Risks")])

mtg_cols = [f"{CQ}Title__c", "RecordType.DeveloperName", f"{CQ}Meeting_Date__c", f"{CQ}Objective__c", f"{CQ}Agenda__c",
            f"{CQ}Status__c", f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_Meeting__c", f"{CQ}Title__c", mtg_cols,
      [{f"{CQ}Title__c": r["Title (Seed Key)"], "RecordType.DeveloperName": "Management_Review",
        f"{CQ}Meeting_Date__c": d10(r["Meeting Date"]), f"{CQ}Objective__c": r["Objective"], f"{CQ}Agenda__c": r["Agenda"],
        **DRAFT} for r in rows("Meetings")])

prj_cols = [f"{CQ}Title__c", "RecordType.DeveloperName", "cqext__Project_Category__c", f"{CQ}SQX_Department__r.Name",
            f"{CQ}Project_Start_Date__c", f"{CQ}Project_End_Date__c", f"{CQ}Description__c", f"{CQ}Status__c",
            f"{CQ}Record_Stage__c"]
add_q("compliancequest__SQX_Project__c", f"{CQ}Title__c", prj_cols,
      [{f"{CQ}Title__c": r["Title (Seed Key)"], "RecordType.DeveloperName": "Default",
        "cqext__Project_Category__c": r["Project Category"], f"{CQ}SQX_Department__r.Name": r["Department"],
        f"{CQ}Project_Start_Date__c": d10(r["Project Start Date"]), f"{CQ}Project_End_Date__c": d10(r["Project End Date"]),
        f"{CQ}Description__c": r["Description"], **DRAFT} for r in rows("Projects")])

# ---------- Source Ids ----------
# SFDMU did not resolve the "<Lookup>__r.<ExternalId>" columns on their own (every lookup and record type
# came through blank), so the CSVs are written the way SFDMU itself exports them: an Id column with a
# made-up source Id per record, and each lookup Id column filled with the parent's made-up Id.
# The __r columns stay alongside. Source Ids never reach the org; SFDMU maps them to the new records.
PARENT = {  # relationship name (without namespace) -> parent object
    "SQX_Business_Unit__r": "SQX_Business_Unit__c", "SQX_Site__r": "SQX_Site__c",
    "SQX_Department__r": "SQX_Department__c", "SQX_Division__r": "SQX_Division__c",
    "SQX_Account__r": "Account", "Account": "Account", "SQX_External_Contact__r": "Contact", "SQX_Contact__r": "Contact",
    "SQX_Part__r": "SQX_Part__c", "SQX_Impacted_Part__r": "SQX_Part__c",
    "Part_Family__r": "SQX_Part_Family__c", "SQX_Part_Family__r": "SQX_Part_Family__c",
    "SQX_Service__r": "SQX_Standard_Service__c", "SQX_Supplier__r": "SQX_Supplier__c",
    "SQX_Lot_Info__r": "SQX_Lot_Info__c", "SQX_Defect_Code__r": "SQX_Defect_Code__c",
    "SQX_Complaint_Code__r": "SQX_Defect_Code__c", "SQX_Result_Type__r": "SQX_Result_Type__c",
    "SQX_Task__r": "SQX_Task__c", "SQX_Nonconformance__r": "SQX_Nonconformance__c",
    "SQX_Complaint__r": "SQX_Complaint__c",
    "SQX_Controlled_Document__r": "SQX_Controlled_Document__c", "SQX_Audit_Criteria_Document__r": "SQX_Controlled_Document__c",
    "SQX_Audit_Program__r": "SQX_Audit_Program__c", "SQX_Audit__r": "SQX_Audit__c", "Account__r": "Account",
    "SQX_Equipment__r": "SQX_Equipment__c", "SQX_Calibration_Supplier__r": "Account",
    "SQX_Applicable_Procedure_Rev__r": "SQX_Controlled_Document__c", "SQX_Applicable_Procedure__r": "SQX_Controlled_Document__c",
    "SQX_Inspection_Criteria__r": "SQX_Controlled_Document__c", "SQX_Supplier_Number__r": "Account",
    "SQX_NC_Number__r": "SQX_Nonconformance__c", "Part__r": "SQX_Part__c",
    "SQX_Job_Function__r": "SQX_Job_Function__c", "SQX_Personnel__r": "SQX_Personnel__c",
    "SQX_Process__r": "SQX_Standard_Service__c",
}


def full(name):
    return name if name in ("Account", "Contact") else CQ + name


def id_field(rel):
    if rel == "Account":
        return "AccountId"
    if rel == "RecordType":
        return "RecordTypeId"
    return rel[:-3] + "__c"


by_obj = {o["sobject"]: o for o in objects}
for n, o in enumerate(objects, start=1):
    for i, r in enumerate(o["rows"], start=1):
        r["Id"] = f"S{n:02d}{i:015d}"          # 18 characters, unique across all objects
problems = []
record_types = {}   # made-up Id -> (DeveloperName, SobjectType)
for o in objects:
    cols = ["Id"]
    for c in o["columns"]:
        if "." in c:
            rel, ext = c.split(".", 1)
            idc = id_field(rel)
            cols.append(idc)
            for r in o["rows"]:
                v = r.get(c, "")
                if v in ("", None):
                    r[idc] = ""
                elif rel == "RecordType":
                    r[idc] = "RT" + format(zlib.crc32(f"{o['sobject']}:{v}".encode()), "016d")
                    record_types[r[idc]] = (v, o["sobject"])
                else:
                    parent = by_obj[full(PARENT[rel.replace(CQ, "")])]
                    match = [p for p in parent["rows"] if str(p.get(ext, "")) == str(v)]
                    if len(match) != 1:
                        problems.append(f"{o['sobject']}.{c} = {v!r}: {len(match)} parent matches")
                        r[idc] = ""
                    else:
                        r[idc] = match[0]["Id"]
        cols.append(c)
    o["query_fields"] = o["query_fields"] or list(o["columns"])
    o["columns"] = cols
if problems:
    raise SystemExit("Lookup problems:\n  " + "\n  ".join(problems))

# ---------- Write ----------
if os.path.exists(DATA):
    shutil.rmtree(DATA)
os.makedirs(DATA)
# Record types: SFDMU matches these to the target org by DeveloperName + object. No record type name repeats
# within an object across namespaces (checked in the dev sandbox), so the namespace isn't part of the key;
# that keeps the seed's own Account "Default" (no namespace) matchable.
with open(os.path.join(DATA, "RecordType.csv"), "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f, quoting=csv.QUOTE_ALL)
    w.writerow(["Id", "DeveloperName", "SobjectType"])
    for rid, (dev, sobj) in sorted(record_types.items(), key=lambda x: (x[1][1], x[1][0])):
        w.writerow([rid, dev, sobj])
export = {"objects": [{"query": "SELECT Id, DeveloperName, SobjectType FROM RecordType",
                       "operation": "Readonly", "externalId": "DeveloperName;SobjectType"},
                      # data/User.csv is written per org by grant-cq-access.ps1 (the load user only).
                      {"query": "SELECT Id, Username FROM User", "operation": "Readonly", "externalId": "Username"}]}
for o in objects:
    with open(os.path.join(DATA, o["sobject"] + ".csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=o["columns"], quoting=csv.QUOTE_ALL)
        w.writeheader()
        for r in o["rows"]:
            w.writerow({k: r.get(k, "") for k in o["columns"]})
    qf = o["query_fields"] or o["columns"]
    export["objects"].append({"query": f"SELECT Id, {', '.join(qf)} FROM {o['sobject']}",
                              "operation": o["operation"], "externalId": o["externalId"]})
with open(os.path.join(DATA, "export.json"), "w") as f:
    json.dump(export, f, indent=2)
print(f"{len(objects)} objects, {sum(len(o['rows']) for o in objects)} records")
for o in objects:
    print(f"  {o['sobject']}: {len(o['rows'])}")

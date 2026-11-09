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
                "RecordType.DeveloperName": "CQ_Supplier" if r["Type"] == "Supplier" else "",  # customers: org default ("Default" is not a CQ record type)
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

c_cols = [f"{CQ}Title__c", NCR, f"{CQ}Description__c", f"{CQ}Due_Date__c", f"{CQ}Status__c"]
add("compliancequest__SQX_Containment__c", f"{CQ}Title__c", c_cols,
    [{f"{CQ}Title__c": r["Containment Action"][:255], NCR: r["Nonconformance (Title)"],
      f"{CQ}Description__c": r["Containment Action"], f"{CQ}Due_Date__c": DUE_DATE, f"{CQ}Status__c": "Draft"}
     for r in rows("Containments")],
    query_fields=[f"{CQ}Title__c", f"{CQ}SQX_Nonconformance__c", f"{CQ}Description__c", f"{CQ}Due_Date__c", f"{CQ}Status__c"])

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

inv_cols = [CR, "RecordType.DeveloperName", f"{CQ}SQX_Part__r.{CQ}Part_Number__c", f"{CQ}SQX_Part_Family__r.Name",
            f"{CQ}Lot_Number__c", f"{CQ}Type_of_Investigation__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c",
            f"{CQ}Investigation_Summary__c"]
invs = []
for r in rows("Complaint Investigations"):
    p, lot = cmp_part[r["Complaint (Title)"]]
    invs.append({CR: r["Complaint (Title)"], "RecordType.DeveloperName": "Complaint_Investigation",
                 f"{CQ}SQX_Part__r.{CQ}Part_Number__c": p, f"{CQ}SQX_Part_Family__r.Name": part_family[p],
                 f"{CQ}Lot_Number__c": lot, f"{CQ}Type_of_Investigation__c": "Complaint Investigation",
                 f"{CQ}Status__c": "Draft", f"{CQ}Record_Stage__c": "Draft",
                 f"{CQ}Investigation_Summary__c": r["Investigation Summary"]})
add("compliancequest__SQX_Investigation__c", CR, inv_cols, invs,
    query_fields=[f"{CQ}SQX_Complaint__c", "RecordTypeId", f"{CQ}SQX_Part__c", f"{CQ}SQX_Part_Family__c",
                  f"{CQ}Lot_Number__c", f"{CQ}Type_of_Investigation__c", f"{CQ}Status__c", f"{CQ}Record_Stage__c",
                  f"{CQ}Investigation_Summary__c"])

phr_cols = [CR, f"{CQ}SQX_Task__r.Name", f"{CQ}Status__c", f"{CQ}Product_History_Review_Summary__c"]
add("compliancequest__SQX_Product_History_Review__c", CR, phr_cols,
    [{CR: r["Complaint (Title)"], f"{CQ}SQX_Task__r.Name": "Atlas Drone Lot History Review", f"{CQ}Status__c": "Draft",
      f"{CQ}Product_History_Review_Summary__c": r["Summary"]} for r in rows("Product History Reviews")],
    query_fields=[f"{CQ}SQX_Complaint__c", f"{CQ}SQX_Task__c", f"{CQ}Status__c", f"{CQ}Product_History_Review_Summary__c"])

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
# Record types: SFDMU matches these to the target org by DeveloperName + namespace + object.
# All the record types the seed uses are CQ's own (compliancequest namespace).
with open(os.path.join(DATA, "RecordType.csv"), "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f, quoting=csv.QUOTE_ALL)
    w.writerow(["Id", "DeveloperName", "NamespacePrefix", "SobjectType"])
    for rid, (dev, sobj) in sorted(record_types.items(), key=lambda x: (x[1][1], x[1][0])):
        w.writerow([rid, dev, "compliancequest", sobj])
export = {"objects": [{"query": "SELECT Id, DeveloperName, NamespacePrefix, SobjectType FROM RecordType",
                       "operation": "Readonly", "externalId": "DeveloperName;NamespacePrefix;SobjectType"}]}
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

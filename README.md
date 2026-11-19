# Atlas Manufacturing – dev seed data

Fictional, non-regulated drone maker. ~5 records per object, all records with a status load as **Draft**.
Built from `Atlas_Manufacturing_Dev_Seed_Data.xlsx`. Field names checked against the dev sandbox (CQ 16.0) on 2026-11-09.

## Load into a scratch org

Run from this folder (the one containing `data` and `config`):

```
sf org create scratch --definition-file config\cq16-snapshot-def.json --alias atlas-seed --duration-days 7 --target-dev-hub golden --no-namespace --wait 30
powershell -ExecutionPolicy Bypass -File .\grant-cq-access.ps1 -Org atlas-seed
sf sfdmu run --sourceusername csvfile --targetusername atlas-seed --path data --noprompt
```

`grant-cq-access.ps1` sets up the load user (the scratch org's admin) in one step:

1. Deploys the seed's own Account record type, `Default`, for the customer accounts. The snapshot has no
   customer record type of its own; the dev sandbox's is also called `Default`.
2. Builds a permission set, `Atlas_Seed_Record_Types`, that makes every active CQ record type in that org
   (plus Account `Default`) visible, and deploys it. Without this, records that name a record type fail with
   "invalid record type".
3. Assigns that permission set plus **every** CQ permission set and permission set group
   (`compliancequest`, `cqext`, `cqcollab`) to the load user. Anything already assigned is skipped, and any
   that Salesforce refuses (for example, one needing a licence the user lacks) is listed rather than stopping the run.
4. Turns on **Salesforce CRM Content User** for the load user and creates the document libraries that CQ's
   **CQ Default Library** settings name (Controlled Document Draft/Release, Template Document Draft/Release), with the
   load user as Library Administrator. Libraries that already exist are left alone. This replaces setting them up by hand.
5. Writes `data\User.csv` with the load user's username. Assignees in the seed (for now, containment
   assignees: CQ copies the assignee to the owner, which can't be blank) point at this user, so they work in
   any org. Run the script again after rebuilding the CSVs, since the rebuild clears the `data` folder.

Everything includes `CQ_Migration_User`, which CQ's validation rules check for and skip. It's labelled CAUTION
in CQ, so the script stops on any org that doesn't look like a scratch org unless you add `-NotScratchOk`.
In a real sandbox, grant record type access and permissions the normal way instead.

If your PowerShell policy blocks the script, run its steps by hand: deploy `force-app` after generating the
permission set, then `sf apex run --file scripts\assign-cq-access.apex --target-org atlas-seed`.

Results and errors from the load land in `data\target` and `data\logs`.

Then attach the document files (see Controlled Documents below):

```
powershell -ExecutionPolicy Bypass -File .\link-document-files.ps1 -Org atlas-seed -DryRun
powershell -ExecutionPolicy Bypass -File .\link-document-files.ps1 -Org atlas-seed
```

## What loads (38 objects, 185 records)

Divisions, Business Unit, Sites, Departments, Processes, Accounts, Contacts, Suppliers, Supplier Contacts, Personnel,
Part Families, Parts, Lot Info, Hierarchical Codes, Yes/No Result Type, CQ Tasks (PHR task + 5 Questions, Investigation task),
5 Nonconformances (+ Impacted Products, Defects, Containments), 5 Complaints (+ Defects),
2 Complaint Investigations, 2 Product History Reviews,
15 Controlled Documents (SOPs, work instructions, inspection criteria), 2 Audit Criteria documents (+ 8 Requirements), 1 Audit Program (+ 2 Scopes), 3 Audits, 3 Audit Findings,
6 Equipment (+ 5 Calibrations, 2 Schedules), 10 Specifications on the 3 Inspection Criteria, 4 Inspections (+ 14 Inspection Details).

## Choices made when converting from the spreadsheet

- **Users:** assignees that CQ requires (containments) are set to the load user through `data\User.csv`.
  Other owners and "Initiated By" default to the user running the load. Personnel records are not linked to users.
  On inspections the load user is also completer and closure reviewer; CQ's migration permission allows that for the load.
- **Inspections don't raise NCs:** "Create NC If Inspection Fails" is off; the failed inspections point at the existing NCs.
- **Associated Items are not loaded.** Complaints carry Part, Part Family and Lot on the header;
  CQ creates the primary associated item itself as the complaint moves forward.
- **Containment due dates are 2027-06-30** (CQ rejects past due dates).
- **NC failure codes:** hardware codes are "Product Failure Code", the firmware code is "Process Failure Code",
  to satisfy CQ's issue-type rule.
- **Supplier accounts** use the `CQ_Supplier` record type (what CQ's supplier rules expect); customers use `Default`.
- **Business Unit → Site:** both sites sit under the Atlas Manufacturing BU. NCs and complaints carry
  Division by site (Houston = Consumer Drones, Piscataway = Commercial Drones).
- **PHRs** link to the complaint and the Atlas PHR task, not to the investigation.
  CQ builds PHR checklists when a PHR is created from the complaint's PHR step; a loaded PHR may not get them.

## Controlled Documents

- 15 Atlas documents from the **Controlled Documents** tab: 8 SOPs, 4 work instructions and 3 inspection criteria
  (`Inspection_Criteria` record type), plus the 2 audit criteria documents. Released ones load as **Current**;
  WI-TST-040 is a new document still in **Draft**. Revisions are left to change orders (CQ's Revise action).
- **Files** are one-page Word documents in `files\`, one per document, named by document number. Word, because a
  controlled document's Primary Content is its editable source (a PDF would be the Secondary rendition).
  `build_document_files.py` holds the text and runs `scripts\make-docx.js` (docx-js keeps each file about 12 KB, so
  the link script can send it in one Apex run).
- **`link-document-files.ps1`** uploads each file and links it the way CQ does. A loaded document has no library, so
  its Draft and Release Library fields are filled from CQ's default libraries. The file goes into the document's
  library (Release for Current, Draft for Draft) with its Controlled Document field set, it's linked to the document,
  and the document's Primary Content is set to it. The file is titled with the document's title, as CQ's own upload does. If no library is found, or it refuses the upload, the file is
  published to the document itself and the run says so. Documents that already have their file only get the library
  fields, the library link and the title, if those are missing, so it's safe to rerun. `-DryRun` previews without changing anything.

## Audits

- **Audit Criteria documents load as Draft.** CQ only lets requirements be added to a Draft document.
  Release them by hand before moving audits forward.
- **Audit checklists are not loaded.** CQ copies the criteria's requirements into the checklist as the audit moves on.
- **Audits use the org's default record type.** "Standard Audit" in the dev sandbox was created in that org, not by CQ.
- **Lead Auditor and Auditee Contact are blank** (users differ per org); fill them in when you work the audit.
- Record types: `grant-cq-access.ps1` makes every CQ record type visible, including Audit_Criteria and Audit_Finding.

## Rebuild

The CSVs and `export.json` are generated from the spreadsheet by `build_sfdmu.py`.
Edit the spreadsheet, rerun the script, reload. Upserts match on the external IDs in `export.json`,
so reloading updates records rather than duplicating them.

The `Id` columns in the CSVs are made-up source Ids (`S01…`, `RT…`) that tie child rows to parent rows.
They never reach the org: SFDMU swaps them for the real Ids of the records it creates.

## Keeping your local copy up to date

First time, from `C:\Users\LouSanatore\Projects` (rename or remove any old `atlas-seed` folder first):

```
git clone https://github.com/lsanatore77/atlas-seed
```

After that, to pick up changes, from inside `atlas-seed`:

```
git pull
```

# Atlas Manufacturing – dev seed data

Fictional, non-regulated drone maker. ~5 records per object, all records with a status load as **Draft**.
Built from `Atlas_Manufacturing_Dev_Seed_Data.xlsx`. Field names checked against the dev sandbox (CQ 16.0) on 2026-11-09.

## Load into a scratch org

Run from this folder (the one containing `data` and `config`):

```
sf org create scratch --definition-file config\cq16-snapshot-def.json --alias atlas-seed --duration-days 7 --target-dev-hub golden --no-namespace --wait 30
sf org assign permset --name compliancequest__CQ_Migration_User --target-org atlas-seed
sf sfdmu run --sourceusername csvfile --targetusername atlas-seed --path data --noprompt
```

`CQ_Migration_User` gives the load user CQ's data-migration permission, which CQ's own validation rules
check for and skip. It's labelled CAUTION in CQ: fine in a scratch org, remove it afterwards in any real org.

Results and errors land in `data\target` and `data\logs`.

## What loads (26 objects, 109 records)

Divisions, Business Unit, Sites, Departments, Processes, Accounts, Contacts, Suppliers, Supplier Contacts, Personnel,
Part Families, Parts, Lot Info, Hierarchical Codes, Yes/No Result Type, CQ Tasks (PHR task + 5 Questions, Investigation task),
5 Nonconformances (+ Impacted Products, Defects, Containments), 5 Complaints (+ Defects),
2 Complaint Investigations, 2 Product History Reviews.

## Choices made when converting from the spreadsheet

- **Users are not set.** Owners, assignees and "Initiated By" default to the user running the load.
  Personnel records are not linked to users.
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

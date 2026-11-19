"""Build the Atlas controlled document files (one short Word document per document) into files/.

Word, because a controlled document's Primary Content is its editable source; a PDF would be the Secondary rendition.

The document list matches the Controlled Documents tab of the review workbook (number, title,
revision, status, effective date). link-document-files.ps1 uploads them after the seed load and
sets each document's Primary Content. Atlas Manufacturing is fictional; this is test data.
"""
import os
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "files")

# ---------- content ----------
# Each document: list of sections. A section is (heading, [paragraphs]) or (heading, {"table": [...rows]})
# or (heading, {"steps": [...]}) for numbered steps.
SOP_RESP = "The Quality Manager owns this procedure. Department leads make sure their staff are trained on it."
C = {
"SOP-QA-001": [
    ("Purpose", ["Defines how Atlas Manufacturing creates, reviews, approves, releases, revises and retires controlled documents, so that people always work to the current approved version."]),
    ("Scope", ["All procedures, work instructions, inspection criteria, forms and specifications used at the Houston Assembly Plant and the Piscataway Components Plant."]),
    ("Responsibilities", ["Document owners write and maintain their documents. Approvers named in the approval matrix review content. Quality releases documents and assigns training. " + SOP_RESP]),
    ("Procedure", {"steps": [
        "Request a new document or revision in CQ. Give the reason for the change.",
        "Draft the document from the approved template. Use the next number in the series (SOP-, WI-, IC-).",
        "Route for approval. Approvers check accuracy, impact on other documents and training needs.",
        "On approval, Quality releases the document. The previous revision becomes obsolete on the same date.",
        "CQ assigns document training to the affected job functions. Staff must complete it before working to the new revision.",
        "Review each document at least every 365 days. Record the review even when nothing changes.",
    ]}),
    ("Records", ["Approval history, training completion and review dates are kept in CQ for the life of the document plus three years."]),
],
"SOP-QA-002": [
    ("Purpose", ["Defines how nonconforming material is identified, segregated and dispositioned so it cannot be used or shipped by mistake."]),
    ("Scope", ["Purchased parts, work in progress and finished drones at both Atlas sites, including material returned by customers."]),
    ("Responsibilities", ["Anyone who finds a suspected nonconformance raises it. Quality Engineering leads the investigation. The Material Review Board (Quality, Engineering, Production) decides the disposition."]),
    ("Procedure", {"steps": [
        "Tag the material with a red HOLD label and move it to the MRB cage. Record the nonconformance in CQ the same shift.",
        "Record part, lot, quantity, where it was found and a description with photos.",
        "Contain: check stock, work in progress and recent shipments from the same lot.",
        "MRB dispositions within five working days: use as is, rework, return to supplier or scrap. Use-as-is needs an engineering justification.",
        "Open a CAPA when the issue repeats, affects safety, or comes from a systemic cause (see SOP-QA-003).",
        "Close the nonconformance when the disposition is complete and the records are attached.",
    ]}),
    ("Records", ["Nonconformance records, MRB decisions and scrap or return paperwork are kept in CQ."]),
],
"SOP-QA-003": [
    ("Purpose", ["Defines when Atlas opens a corrective or preventive action, how the root cause is found, and how effectiveness is checked."]),
    ("Scope", ["Repeat or significant nonconformances, customer complaints, audit findings, supplier escalations and trends from inspection data."]),
    ("Responsibilities", ["The CAPA owner leads the investigation and action plan. The Quality Manager approves the plan and the effectiveness check. " + "Action owners complete their tasks by the due date."]),
    ("Procedure", {"steps": [
        "Open a CAPA in CQ and link the source records (nonconformance, complaint, finding).",
        "Contain the problem if it is not already contained.",
        "Find the root cause using 5 Whys or a fishbone. Test the cause against the evidence before acting on it.",
        "Plan corrective actions and, where the same cause could affect other products, preventive actions.",
        "Implement the actions. Changes to documents or parts go through change control (SOP-QA-007).",
        "Check effectiveness after 90 days of production, or three lots, whichever is longer. Reopen if the problem returns.",
    ]}),
    ("Records", ["CAPA records, investigation evidence and effectiveness results are kept in CQ."]),
],
"SOP-QA-004": [
    ("Purpose", ["Defines how purchased parts are sampled, inspected and released at the Piscataway receiving dock."]),
    ("Scope", ["All production parts from approved suppliers. Consumables and tools are out of scope."]),
    ("Responsibilities", ["Receiving Inspection performs inspections to the part's inspection criteria (IC- documents). Quality Engineering owns the criteria and sampling plans."]),
    ("Procedure", {"steps": [
        "Check the delivery against the purchase order: part number, quantity, lot and certificate of conformance.",
        "Find the inspection criteria for the part number. If none exist, hold the lot and tell Quality Engineering.",
        "Sample the lot to the plan in the criteria (ANSI/ASQ Z1.4, general level II, unless the criteria say otherwise).",
        "Inspect and record every result in CQ. Use only calibrated equipment.",
        "Accept: label the lot RELEASED and move it to stock. Reject: label it HOLD and raise a nonconformance (SOP-QA-002).",
        "Lots from a supplier on escalation get tightened inspection until Quality removes it.",
    ]}),
    ("Records", ["Inspection results, certificates of conformance and release decisions are kept in CQ against the lot."]),
],
"SOP-QA-005": [
    ("Purpose", ["Defines calibration and preventive maintenance for production and test equipment, and what to do when equipment is found out of tolerance."]),
    ("Scope", ["Torque drivers, motor test stands, battery cyclers, X-ray, scales and any gauge used to accept product."]),
    ("Responsibilities", ["Quality schedules calibration and keeps equipment records. Users check the calibration label before use and report damage."]),
    ("Procedure", {"steps": [
        "Give each item an asset number and record it in CQ with its calibration interval (default 12 months; torque drivers 6 months).",
        "Calibrate against standards traceable to national standards. Keep the certificate.",
        "Label the item with its calibration due date. Do not use equipment past its due date.",
        "If an item is found out of tolerance, quarantine it and assess product made since the last good calibration. Raise a nonconformance if product may be affected.",
        "Carry out preventive maintenance to the manufacturer's schedule and record it.",
    ]}),
    ("Records", ["Equipment records, calibration certificates and maintenance logs are kept in CQ."]),
],
"SOP-QA-006": [
    ("Purpose", ["Defines how Atlas approves new suppliers, monitors their performance, handles deviations and escalates when a supplier underperforms."]),
    ("Scope", ["Suppliers of production parts and components, including motors, battery cells, PCBs, composites and optics."]),
    ("Responsibilities", ["Supplier Quality runs qualification and monitoring. Purchasing buys only from approved suppliers. The Quality Manager approves escalations."]),
    ("Procedure", {"steps": [
        "Qualify a new supplier with a questionnaire, a quality agreement and, for high-risk parts, an on-site audit.",
        "Approve the supplier in CQ for specific parts. Purchasing checks approval before ordering.",
        "Review supplier performance quarterly: lot acceptance rate, on-time delivery and response to nonconformances.",
        "Approve deviations (temporary acceptance of a known difference) only with an engineering justification and an end date.",
        "Escalate a supplier with two rejected lots in a quarter or a safety-related failure. Escalation means tightened inspection and a supplier corrective action request.",
    ]}),
    ("Records", ["Qualification files, scorecards, deviations and escalations are kept in CQ against the supplier."]),
],
"SOP-QA-007": [
    ("Purpose", ["Defines how changes to parts, processes and documents are requested, assessed, approved and implemented."]),
    ("Scope", ["Design changes, process changes, supplier or material changes, and changes to controlled documents."]),
    ("Responsibilities", ["Anyone may request a change. Engineering assesses impact. The change board (Engineering, Quality, Production) approves change orders."]),
    ("Procedure", {"steps": [
        "Raise a change request in CQ describing the change and why it is needed.",
        "Assess the impact: affected parts, documents, inspection criteria, stock, suppliers and training.",
        "Approve or reject the request. Approved requests become a change order with tasks and owners.",
        "Implement the tasks: revise documents, update parts, notify suppliers, use up or scrap old stock.",
        "Verify the change before closing the change order. Confirm training on revised documents is complete.",
    ]}),
    ("Records", ["Change requests, change orders, impact assessments and verification evidence are kept in CQ."]),
],
"SOP-HR-001": [
    ("Purpose", ["Defines how Atlas makes sure people are trained and qualified for their job before they work unsupervised."]),
    ("Scope", ["All employees and contractors at both sites whose work affects product quality."]),
    ("Responsibilities", ["Managers define training needs for each job function. Quality assigns document training. Qualified trainers deliver on-the-job training."]),
    ("Procedure", {"steps": [
        "Map each job function to the documents and courses it needs in CQ.",
        "New starters complete their job function's training plan before working unsupervised.",
        "When a document is released or revised, CQ assigns document training to the affected job functions. Complete it within the days set on the document.",
        "Practical skills (soldering, torque, flight test) need a trainer sign-off after supervised work.",
        "Managers review overdue training weekly.",
    ]}),
    ("Records", ["Training records, trainer qualifications and sign-offs are kept in CQ."]),
],
"WI-ASM-010": [
    ("Purpose", ["Mount 2306 brushless motors (ATL-PR-2306) to AX-200 carbon arms with the correct thread locker and torque."]),
    ("Equipment", ["Calibrated torque driver set to 0.6 N·m, M3 hex bit, medium-strength thread locker, motor screws M3 x 6 mm."]),
    ("Steps", {"steps": [
        "Check the torque driver's calibration label. Do not use it if it is past due.",
        "Check the motor spins freely by hand with no grinding or noise. Set aside any noisy motor and raise a nonconformance.",
        "Apply one drop of thread locker to each of the four screws.",
        "Fit the motor with the leads facing the arm root. Torque the screws in a cross pattern to 0.6 N·m.",
        "Mark each screw head with a torque stripe.",
        "Record the motor lot number on the build traveler.",
    ]}),
    ("Quality checks", ["All four torque stripes present. Motor spins freely after mounting. Lot recorded."]),
],
"WI-ASM-020": [
    ("Purpose", ["Assemble 4S 5000 mAh LiPo packs (ATL-PW-4S5000) from matched cells and check them for swelling before release."]),
    ("Equipment", ["Cell tester (voltage and internal resistance), spot welder, pack compression jig, feeler gauge set."]),
    ("Steps", {"steps": [
        "Measure each cell's open-circuit voltage and internal resistance. Record both.",
        "Group cells so that voltage is within 0.01 V and internal resistance within 2 mΩ across the pack.",
        "Weld the series connections and the balance leads. Check every weld with a pull test on the first pack of each shift.",
        "Wrap and label the pack with the lot number.",
        "Measure pack thickness in the jig. Reject any pack more than 0.5 mm over nominal: possible swelling.",
        "Store finished packs at storage charge (3.80 V per cell).",
    ]}),
    ("Quality checks", ["Cell matching record complete. Thickness within limit. Lot label applied."]),
],
"WI-ASM-030": [
    ("Purpose", ["Load and verify flight controller firmware on F7 boards (ATL-EL-F7FC)."]),
    ("Equipment", ["Firmware station PC, USB-C cable, released firmware image from the engineering share."]),
    ("Steps", {"steps": [
        "Check the firmware version on the build traveler matches the released image on the station.",
        "Connect the board in bootloader mode. Flash the image.",
        "Read back the checksum and compare it with the published checksum. A mismatch is a reject: reflash once, then raise a nonconformance.",
        "Power-cycle the board and confirm the version string on the configurator screen.",
        "Record the board serial number, firmware version and checksum.",
    ]}),
    ("Quality checks", ["Checksum matches. Version string confirmed after power-cycle. Record complete."]),
],
"WI-TST-040": [
    ("Purpose", ["Final flight test of each finished drone before packaging. DRAFT: content still being written."]),
    ("Equipment", ["Tethered test cage, ground station, calibrated tachometer."]),
    ("Steps", {"steps": [
        "Tethered hover for 60 seconds. Check for vibration and drift.",
        "Check motor balance: each motor's RPM at hover within 5% of the average.",
        "Failsafe test: switch off the transmitter and confirm the drone lands or returns as configured.",
        "[To be completed: camera and gimbal check, battery telemetry check, pass/fail limits.]",
    ]}),
    ("Quality checks", ["[To be completed.]"]),
],
"IC-REC-2306": [
    ("Purpose", ["Incoming inspection of 2306 brushless motors (ATL-PR-2306) from Voltcore Motors."]),
    ("Sampling", ["ANSI/ASQ Z1.4 general level II, AQL 0.65 for functional checks. Tightened inspection while the supplier is on escalation."]),
    ("Checks", {"table": [
        ["Characteristic", "Method", "Requirement"],
        ["Bearing noise", "Spin by hand and listen", "No grinding or clicking"],
        ["Shaft end play", "Dial indicator", "≤ 0.05 mm"],
        ["KV rating", "Motor test stand", "2450 KV ± 5%"],
        ["Winding resistance", "Milliohm meter", "Within 5% across all phases"],
        ["Leads and solder joints", "Visual", "No damaged insulation or cold joints"],
    ]}),
],
"IC-REC-4S5000": [
    ("Purpose", ["Incoming inspection of 4S 5000 mAh LiPo battery packs (ATL-PW-4S5000) and cells from Helion Cell Co."]),
    ("Sampling", ["ANSI/ASQ Z1.4 general level II, AQL 0.4. Any swollen pack rejects the lot."]),
    ("Checks", {"table": [
        ["Characteristic", "Method", "Requirement"],
        ["Open-circuit voltage", "Voltmeter", "3.75 to 3.85 V per cell"],
        ["Cell balance", "Cell checker", "Spread ≤ 0.01 V"],
        ["Swelling", "Thickness jig", "≤ nominal + 0.5 mm"],
        ["Connector and leads", "Visual", "No damage; correct polarity"],
        ["Label and lot", "Visual", "Lot matches certificate of conformance"],
    ]}),
],
"IC-REC-F7": [
    ("Purpose", ["Incoming inspection of F7 flight controller boards (ATL-EL-F7FC) from Brightline PCB."]),
    ("Sampling", ["Visual and power-on test: ANSI/ASQ Z1.4 general level II, AQL 0.65. X-ray: 5 boards per lot, 100% while the cold-solder CAPA is open."]),
    ("Checks", {"table": [
        ["Characteristic", "Method", "Requirement"],
        ["Solder joints", "Visual, 10x", "No cold, bridged or missing joints"],
        ["BGA solder joints", "X-ray", "Voids ≤ 25% of ball area"],
        ["Power-on", "Bench supply 5 V", "Current 120 to 180 mA; status LED on"],
        ["Connectors", "Visual", "Seated, no bent pins"],
        ["Label and lot", "Visual", "Lot matches certificate of conformance"],
    ]}),
],
}

# ---------- build ----------
# The Word files are written by scripts/make-docx.js (docx-js keeps them small, about 10 KB, so
# link-document-files.ps1 can send each one in a single Apex run). This script supplies the text.
import json, subprocess, tempfile


def spec(meta, sections):
    num, title, rtype, status, rev, eff, fname = meta
    out = []
    for i, (heading, body) in enumerate(sections):
        s = {"id": i, "heading": heading}
        if isinstance(body, list):
            s["paragraphs"] = body
        elif "steps" in body:
            s["steps"] = body["steps"]
        else:
            s["table"] = body["table"]
        out.append(s)
    return {"number": num, "title": title, "recordType": rtype, "status": status, "revision": rev,
            "effective": eff, "file": fname, "sections": out}


def main():
    wb = load_workbook(os.path.join(HERE, "Atlas_Manufacturing_Dev_Seed_Data.xlsx"))
    ws = wb["Controlled Documents"]
    hdr = [c.value for c in ws[3]]
    os.makedirs(OUT, exist_ok=True)
    specs = []
    for r in ws.iter_rows(min_row=4, values_only=True):
        if all(v is None for v in r):
            continue
        d = dict(zip(hdr, r))
        num = d["Document Number (Seed Key)"]
        if num not in C:
            raise SystemExit(f"No content written for {num}")
        eff = str(d["Effective Date"] or "")[:10]
        specs.append(spec((num, d["Title"], d["Record Type"], d["Document Status"], d["Revision"], eff,
                           d["File"] or f"{num}.docx"), C[num]))
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as f:
        json.dump(specs, f)
    subprocess.run(["node", os.path.join(HERE, "scripts", "make-docx.js"), f.name, OUT], check=True)
    os.remove(f.name)
    # Drop the zip's folder entries: Word has refused files that have them when opened through online editors.
    import zipfile
    for sp in specs:
        fp = os.path.join(OUT, sp["file"])
        with zipfile.ZipFile(fp) as z:
            items = [(i, z.read(i.filename)) for i in z.infolist() if not i.filename.endswith("/")]
        with zipfile.ZipFile(fp, "w", zipfile.ZIP_DEFLATED) as z:
            for i, data in items:
                z.writestr(i, data)


if __name__ == "__main__":
    main()

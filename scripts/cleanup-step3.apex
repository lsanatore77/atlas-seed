// One-time cleanup for a scratch org loaded before the inspection redesign (2026-11-19).
// Deletes the four seed inspections (with their detail lines), the three IC-REC Inspection Criteria documents that
// loaded as Current (plus each document's Primary Content file) and the Closed calibration on ATL-EQ-001, so the
// next load recreates them at their starting point: documents Draft with their specifications, inspections Draft,
// the calibration Open. cleanup-step3.ps1 fills in __DRY__.
Boolean dryRun = __DRY__;
Set<String> lots = new Set<String>{ 'LOT-PR-2410-07', 'LOT-EL-2409-03', 'LOT-PW-2411-04', 'LOT-PR-2411-02' };
Set<String> docNumbers = new Set<String>{ 'IC-REC-2306', 'IC-REC-4S5000', 'IC-REC-F7' };
List<String> out = new List<String>();

List<compliancequest__SQX_Inspection__c> ins = [
    SELECT Id, Name, compliancequest__Lot_Number__c, compliancequest__Status__c
    FROM compliancequest__SQX_Inspection__c WHERE compliancequest__Lot_Number__c IN :lots];
List<compliancequest__SQX_Inspection_Detail__c> details = [
    SELECT Id FROM compliancequest__SQX_Inspection_Detail__c WHERE compliancequest__SQX_Inspection__c IN :ins];
List<compliancequest__SQX_Controlled_Document__c> docs = [
    SELECT Id, compliancequest__Document_Number__c, compliancequest__Document_Status__c, compliancequest__Content_Reference__c
    FROM compliancequest__SQX_Controlled_Document__c WHERE compliancequest__Document_Number__c IN :docNumbers];
List<compliancequest__SQX_Calibration__c> cals = [
    SELECT Id, compliancequest__Calibration_Date__c FROM compliancequest__SQX_Calibration__c
    WHERE compliancequest__SQX_Equipment__r.Name = 'ATL-EQ-001' AND compliancequest__Calibration_Status__c = 'Close'];
Integer specCount = [SELECT COUNT() FROM compliancequest__SQX_Specification__c WHERE compliancequest__SQX_Inspection_Criteria__c IN :docs];

for (compliancequest__SQX_Inspection__c i : ins) {
    out.add((dryRun ? 'WOULD DELETE' : 'DELETING') + ' inspection ' + i.Name + ' (' + i.compliancequest__Lot_Number__c + ', ' + i.compliancequest__Status__c + ')');
}
out.add((dryRun ? 'WOULD DELETE ' : 'DELETING ') + details.size() + ' inspection detail lines');
for (compliancequest__SQX_Controlled_Document__c d : docs) {
    out.add((dryRun ? 'WOULD DELETE' : 'DELETING') + ' document ' + d.compliancequest__Document_Number__c + ' (' + d.compliancequest__Document_Status__c + ')' +
            (String.isBlank(d.compliancequest__Content_Reference__c) ? '' : ' and its file'));
}
for (compliancequest__SQX_Calibration__c c : cals) {
    out.add((dryRun ? 'WOULD DELETE' : 'DELETING') + ' closed calibration on ATL-EQ-001 (' + String.valueOf(c.compliancequest__Calibration_Date__c) + ')');
}
if (specCount > 0) { out.add('(' + specCount + ' specifications go with the documents)'); }

if (!dryRun) {
    Integer failed = 0;
    for (Database.DeleteResult r : Database.delete(details, false)) {
        if (!r.isSuccess()) { failed++; out.add('  detail NOT deleted: ' + r.getErrors()[0].getMessage()); }
    }
    for (Database.DeleteResult r : Database.delete(ins, false)) {
        if (!r.isSuccess()) { failed++; out.add('  inspection NOT deleted: ' + r.getErrors()[0].getMessage()); }
    }
    for (Database.DeleteResult r : Database.delete(cals, false)) {
        if (!r.isSuccess()) { failed++; out.add('  calibration NOT deleted: ' + r.getErrors()[0].getMessage()); }
    }
    List<Id> files = new List<Id>();
    Database.DeleteResult[] dr = Database.delete(docs, false);
    for (Integer k = 0; k < dr.size(); k++) {
        if (dr[k].isSuccess()) {
            String ref = docs[k].compliancequest__Content_Reference__c;
            try { if (String.isNotBlank(ref)) { files.add(Id.valueOf(ref)); } } catch (Exception e) { }
        } else {
            failed++; out.add('  ' + docs[k].compliancequest__Document_Number__c + ' NOT deleted: ' + dr[k].getErrors()[0].getMessage());
        }
    }
    List<ContentDocument> cds = [SELECT Id FROM ContentDocument WHERE Id IN :files];
    for (Database.DeleteResult r : Database.delete(cds, false)) {
        if (!r.isSuccess()) { failed++; out.add('  file NOT deleted: ' + r.getErrors()[0].getMessage()); }
    }
    out.add(failed == 0 ? 'CLEANUP DONE. Run the seed load and link-document-files.ps1 next.' : 'CLEANUP FINISHED WITH ' + failed + ' PROBLEMS (see above).');
}
for (String line : out) { System.debug(line); }

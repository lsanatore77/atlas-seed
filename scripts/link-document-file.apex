// Uploads one document file and links it as that document's Primary Content.
// link-document-files.ps1 fills in the placeholders for each file in files\ and runs this once per file.
// The file goes into the document's library (Release for released documents, Draft for drafts) and is
// linked to the document; if the library refuses it, the file is published to the document itself.
String docNumber = '__NUM__';
String fileName = '__FILE__';
Boolean dryRun = __DRY__;
String fileBase64 = '__B64__';

String result;
try {
    List<compliancequest__SQX_Controlled_Document__c> docs = [
        SELECT Id, compliancequest__Title__c, compliancequest__Document_Status__c, compliancequest__Content_Reference__c,
               compliancequest__Draft_Vault__c, compliancequest__Release_Vault__c
        FROM compliancequest__SQX_Controlled_Document__c
        WHERE compliancequest__Document_Number__c = :docNumber];
    if (docs.size() != 1) {
        result = 'NO DOCUMENT: ' + docs.size() + ' documents numbered ' + docNumber;
    } else {
        compliancequest__SQX_Controlled_Document__c d = docs[0];
        Boolean hasFile = false;
        if (String.isNotBlank(d.compliancequest__Content_Reference__c)) {
            try {
                Id current = Id.valueOf(d.compliancequest__Content_Reference__c);
                hasFile = [SELECT COUNT() FROM ContentDocument WHERE Id = :current] > 0;
            } catch (Exception e) { hasFile = false; }
        }
        if (hasFile) {
            result = 'ALREADY LINKED: ' + docNumber;
        } else if (dryRun) {
            result = 'WOULD LINK: ' + docNumber + ' (' + d.compliancequest__Document_Status__c + ') <- ' + fileName;
        } else {
            String libText = d.compliancequest__Document_Status__c == 'Draft'
                ? d.compliancequest__Draft_Vault__c : d.compliancequest__Release_Vault__c;
            Id lib = null;
            try { if (String.isNotBlank(libText)) { lib = Id.valueOf(libText); } } catch (Exception e) { lib = null; }
            Blob body = EncodingUtil.base64Decode(fileBase64);
            String title = (docNumber + ' ' + d.compliancequest__Title__c).left(255);
            String placedIn = 'library';
            ContentVersion v = new ContentVersion(Title = title, PathOnClient = fileName, VersionData = body,
                compliancequest__Controlled_Document__c = d.Id, FirstPublishLocationId = (lib != null ? lib : d.Id));
            Database.SaveResult sr = Database.insert(v, false);
            if (lib == null) { placedIn = 'document (no library set)'; }
            if (!sr.isSuccess() && lib != null) {
                String libError = sr.getErrors()[0].getMessage();
                v = new ContentVersion(Title = title, PathOnClient = fileName, VersionData = body,
                    compliancequest__Controlled_Document__c = d.Id, FirstPublishLocationId = d.Id);
                sr = Database.insert(v, false);
                lib = null;
                placedIn = 'document (library refused: ' + libError + ')';
            }
            if (!sr.isSuccess()) {
                result = 'UPLOAD FAILED: ' + docNumber + ': ' + sr.getErrors()[0].getMessage();
            } else {
                Id contentDocId = [SELECT ContentDocumentId FROM ContentVersion WHERE Id = :v.Id].ContentDocumentId;
                String linkNote = '';
                if (lib != null) {
                    Database.SaveResult lr = Database.insert(new ContentDocumentLink(ContentDocumentId = contentDocId,
                        LinkedEntityId = d.Id, ShareType = 'V', Visibility = 'AllUsers'), false);
                    if (!lr.isSuccess()) { linkNote = '; link to document failed: ' + lr.getErrors()[0].getMessage(); }
                }
                d.compliancequest__Content_Reference__c = contentDocId;
                Database.SaveResult ur = Database.update(d, false);
                result = ur.isSuccess()
                    ? 'LINKED: ' + docNumber + ' <- ' + fileName + ' (file in ' + placedIn + linkNote + ')'
                    : 'UPLOADED BUT PRIMARY CONTENT NOT SET: ' + docNumber + ': ' + ur.getErrors()[0].getMessage();
            }
        }
    }
} catch (Exception e) {
    result = 'ERROR: ' + docNumber + ': ' + e.getMessage();
}
System.debug(result);

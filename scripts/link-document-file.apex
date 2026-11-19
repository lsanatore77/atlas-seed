// Uploads one document file, links it as that document's Primary Content and files it in CQ's library.
// link-document-files.ps1 fills in the placeholders for each file in files\ and runs this once per file.
// - Library fields: a loaded document has no Draft/Release Library, so they are filled from the libraries named
//   in CQ's Default Library settings (created by grant-cq-access.ps1).
// - The file goes into the Release library for released documents and the Draft library for drafts, and is
//   linked to the document. If the library is missing or refuses it, the file is published to the document.
// - A document that already has its file only gets the library fields and the library link, if missing.
String docNumber = '__NUM__';
String fileName = '__FILE__';
Boolean dryRun = __DRY__;
String fileBase64 = '__B64__';

String result;
try {
    List<compliancequest__SQX_Controlled_Document__c> docs = [
        SELECT Id, compliancequest__Title__c, compliancequest__Document_Status__c, compliancequest__Content_Reference__c,
               compliancequest__Draft_Vault__c, compliancequest__Release_Vault__c, RecordType.DeveloperName
        FROM compliancequest__SQX_Controlled_Document__c
        WHERE compliancequest__Document_Number__c = :docNumber];
    if (docs.size() != 1) {
        result = 'NO DOCUMENT: ' + docs.size() + ' documents numbered ' + docNumber;
    } else {
        compliancequest__SQX_Controlled_Document__c d = docs[0];

        // Libraries: the document's own, or CQ's defaults for this kind of document
        String settingName = d.RecordType.DeveloperName == 'Template_Document'
            ? 'CQ_Template_Document_Default_Library' : 'CQ_Controlled_Document_Default_Library';
        Map<String, Id> libByName = new Map<String, Id>();
        List<compliancequest__CQ_Default_Library__mdt> setting = [
            SELECT compliancequest__Draft_Library__c, compliancequest__Release_Library__c
            FROM compliancequest__CQ_Default_Library__mdt WHERE DeveloperName = :settingName LIMIT 1];
        String draftName = setting.isEmpty() ? null : setting[0].compliancequest__Draft_Library__c;
        String releaseName = setting.isEmpty() ? null : setting[0].compliancequest__Release_Library__c;
        for (ContentWorkspace w : [SELECT Id, Name FROM ContentWorkspace WHERE Name = :draftName OR Name = :releaseName]) {
            libByName.put(w.Name, w.Id);
        }
        Boolean fieldsChanged = false;
        if (String.isBlank(d.compliancequest__Draft_Vault__c) && libByName.containsKey(draftName)) {
            d.compliancequest__Draft_Vault__c = libByName.get(draftName); fieldsChanged = true;
        }
        if (String.isBlank(d.compliancequest__Release_Vault__c) && libByName.containsKey(releaseName)) {
            d.compliancequest__Release_Vault__c = libByName.get(releaseName); fieldsChanged = true;
        }
        String libText = d.compliancequest__Document_Status__c == 'Draft'
            ? d.compliancequest__Draft_Vault__c : d.compliancequest__Release_Vault__c;
        Id lib = null;
        try { if (String.isNotBlank(libText)) { lib = Id.valueOf(libText); } } catch (Exception e) { lib = null; }

        // Does the document already have its file?
        Id currentFile = null;
        if (String.isNotBlank(d.compliancequest__Content_Reference__c)) {
            try {
                Id ref = Id.valueOf(d.compliancequest__Content_Reference__c);
                if ([SELECT COUNT() FROM ContentDocument WHERE Id = :ref] > 0) { currentFile = ref; }
            } catch (Exception e) { currentFile = null; }
        }

        if (currentFile != null) {
            Boolean inLibrary = lib == null || [SELECT COUNT() FROM ContentDocumentLink
                                                WHERE ContentDocumentId = :currentFile AND LinkedEntityId = :lib] > 0;
            if (!fieldsChanged && inLibrary) {
                result = 'ALREADY LINKED: ' + docNumber;
            } else if (dryRun) {
                result = 'WOULD FILE IN LIBRARY: ' + docNumber + (lib == null ? ' (no library found)' : '');
            } else {
                String note = '';
                if (!inLibrary) {
                    Database.SaveResult lr = Database.insert(new ContentDocumentLink(ContentDocumentId = currentFile,
                        LinkedEntityId = lib, ShareType = 'I', Visibility = 'AllUsers'), false);
                    note = lr.isSuccess() ? 'file added to library' : 'file NOT added to library: ' + lr.getErrors()[0].getMessage();
                }
                if (fieldsChanged) {
                    Database.SaveResult ur = Database.update(d, false);
                    note += (note == '' ? '' : '; ') + (ur.isSuccess() ? 'library fields set' : 'library fields NOT set: ' + ur.getErrors()[0].getMessage());
                }
                result = 'UPDATED: ' + docNumber + ' (' + note + ')';
            }
        } else if (dryRun) {
            result = 'WOULD LINK: ' + docNumber + ' (' + d.compliancequest__Document_Status__c + ') <- ' + fileName +
                     (lib == null ? ' (no library found: file goes on the document)' : '');
        } else {
            Blob body = EncodingUtil.base64Decode(fileBase64);
            String title = (docNumber + ' ' + d.compliancequest__Title__c).left(255);
            String placedIn = lib == null ? 'document (no library found)' : 'library';
            ContentVersion v = new ContentVersion(Title = title, PathOnClient = fileName, VersionData = body,
                compliancequest__Controlled_Document__c = d.Id, FirstPublishLocationId = (lib != null ? lib : d.Id));
            Database.SaveResult sr = Database.insert(v, false);
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

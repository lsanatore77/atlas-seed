// Creates the libraries CQ's Default Library settings (CQ_Default_Library__mdt) name, if they don't exist yet,
// and makes the running user a Library Administrator of each. Run by grant-cq-access.ps1.
Id me = UserInfo.getUserId();
Set<String> names = new Set<String>();
for (compliancequest__CQ_Default_Library__mdt m : [
        SELECT compliancequest__Draft_Library__c, compliancequest__Release_Library__c, compliancequest__Distribution_Library__c
        FROM compliancequest__CQ_Default_Library__mdt]) {
    for (String n : new List<String>{ m.compliancequest__Draft_Library__c, m.compliancequest__Release_Library__c,
                                      m.compliancequest__Distribution_Library__c }) {
        if (String.isNotBlank(n)) { names.add(n); }
    }
}
Map<String, Id> libs = new Map<String, Id>();
for (ContentWorkspace w : [SELECT Id, Name FROM ContentWorkspace WHERE Name IN :names]) { libs.put(w.Name, w.Id); }

List<ContentWorkspace> toCreate = new List<ContentWorkspace>();
for (String n : names) {
    if (!libs.containsKey(n)) { toCreate.add(new ContentWorkspace(Name = n)); }
}
Integer created = 0;
List<Database.SaveResult> rs = Database.insert(toCreate, false);
for (Integer i = 0; i < rs.size(); i++) {
    if (rs[i].isSuccess()) { created++; libs.put(toCreate[i].Name, rs[i].getId()); }
    else { System.debug('LIBRARY NOT CREATED ' + toCreate[i].Name + ': ' + rs[i].getErrors()[0].getMessage()); }
}

// Library Administrator membership for the load user
Set<Id> memberOf = new Set<Id>();
for (ContentWorkspaceMember wm : [SELECT ContentWorkspaceId FROM ContentWorkspaceMember WHERE MemberId = :me AND ContentWorkspaceId IN :libs.values()]) {
    memberOf.add(wm.ContentWorkspaceId);
}
List<ContentWorkspacePermission> admin = [SELECT Id FROM ContentWorkspacePermission WHERE Type = 'Admin' LIMIT 1];
List<ContentWorkspaceMember> joins = new List<ContentWorkspaceMember>();
if (!admin.isEmpty()) {
    for (Id libId : libs.values()) {
        if (!memberOf.contains(libId)) {
            joins.add(new ContentWorkspaceMember(ContentWorkspaceId = libId, MemberId = me, ContentWorkspacePermissionId = admin[0].Id));
        }
    }
}
Integer joined = 0;
List<Database.SaveResult> js = Database.insert(joins, false);
for (Integer i = 0; i < js.size(); i++) {
    if (js[i].isSuccess()) { joined++; }
    else { System.debug('NOT ADDED TO LIBRARY ' + joins[i].ContentWorkspaceId + ': ' + js[i].getErrors()[0].getMessage()); }
}
System.debug('LIBRARIES: ' + names.size() + ' named in CQ settings, ' + created + ' created, ' + (names.size() - created) +
             ' already there; load user added to ' + joined);

// Gives the running user every ComplianceQuest permission set and permission set group,
// plus Atlas_Seed_Record_Types. Ones already held are skipped. Failures are listed, not fatal.
Set<String> cqNamespaces = new Set<String>{ 'compliancequest', 'cqext', 'cqcollab' };
Id me = UserInfo.getUserId();

// Libraries (CQ's document content) need the Salesforce CRM Content User setting on the user.
User loadUser = [SELECT Id, UserPermissionsSFContentUser FROM User WHERE Id = :me];
if (!loadUser.UserPermissionsSFContentUser) {
    loadUser.UserPermissionsSFContentUser = true;
    Database.SaveResult cu = Database.update(loadUser, false);
    System.debug(cu.isSuccess() ? 'Salesforce CRM Content User turned on'
                                : 'Salesforce CRM Content User NOT turned on: ' + cu.getErrors()[0].getMessage());
}

Set<Id> held = new Set<Id>();
for (PermissionSetAssignment a : [SELECT PermissionSetId, PermissionSetGroupId FROM PermissionSetAssignment WHERE AssigneeId = :me]) {
    held.add(a.PermissionSetId);
    if (a.PermissionSetGroupId != null) { held.add(a.PermissionSetGroupId); }
}

List<PermissionSetAssignment> toAdd = new List<PermissionSetAssignment>();
List<String> labels = new List<String>();
Integer skipped = 0;
for (PermissionSet p : [
    SELECT Id, Name, NamespacePrefix FROM PermissionSet
    WHERE IsOwnedByProfile = false AND Type = 'Regular'
      AND (NamespacePrefix IN :cqNamespaces OR (NamespacePrefix = null AND Name = 'Atlas_Seed_Record_Types'))
]) {
    if (held.contains(p.Id)) { skipped++; continue; }
    toAdd.add(new PermissionSetAssignment(AssigneeId = me, PermissionSetId = p.Id));
    labels.add((p.NamespacePrefix == null ? '' : p.NamespacePrefix + '__') + p.Name);
}
for (PermissionSetGroup g : [SELECT Id, DeveloperName, NamespacePrefix FROM PermissionSetGroup WHERE NamespacePrefix IN :cqNamespaces]) {
    if (held.contains(g.Id)) { skipped++; continue; }
    toAdd.add(new PermissionSetAssignment(AssigneeId = me, PermissionSetGroupId = g.Id));
    labels.add(g.NamespacePrefix + '__' + g.DeveloperName + ' (group)');
}

Integer ok = 0;
List<Database.SaveResult> results = Database.insert(toAdd, false);
for (Integer i = 0; i < results.size(); i++) {
    if (results[i].isSuccess()) { ok++; }
    else { System.debug('NOT ASSIGNED ' + labels[i] + ': ' + results[i].getErrors()[0].getMessage()); }
}
System.debug('ASSIGNED ' + ok + ' of ' + toAdd.size() + ' new; ' + skipped + ' already held');

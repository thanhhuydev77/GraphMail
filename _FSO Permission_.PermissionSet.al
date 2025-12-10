permissionset 70000 "FSO Permission"
{
    Assignable = true;
    Permissions = table "FSO Email - Graph API Account"=X,
        tabledata "FSO Email - Graph API Account"=RIMD,
        codeunit "FSO Email - Graph API Helper"=X,
        codeunit "FSO Graph User Connector"=X,
        page "FSO Graph User Email Account"=X;
}

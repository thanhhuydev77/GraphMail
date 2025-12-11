permissionset 70000 " Permission"
{
    Assignable = true;
    Permissions = table "Email - Graph API Account" = X,
        tabledata "Email - Graph API Account" = RIMD,
        codeunit "Email - Graph API Helper" = X,
        codeunit "Graph User Connector" = X,
        page " Graph UserEmail Account" = X;
}

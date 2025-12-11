namespace GraphMail;
permissionset 70000 "ALE Permission"
{
    Assignable = true;
    Permissions = table "ALE Email - Graph API Account" = X,
        tabledata "ALE Email - Graph API Account" = RIMD,
        codeunit "ALE Email - Graph API Helper" = X,
        codeunit "ALE Graph User Connector" = X,
        page "ALE Graph User Email Account" = X;
}

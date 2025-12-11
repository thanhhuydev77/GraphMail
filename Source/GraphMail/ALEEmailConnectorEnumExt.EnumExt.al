namespace GraphMail;

using System.Email;
enumextension 70000 "ALE Email Connector EnumExt" extends "Email Connector"
{
    value(70000; "Graph User")
    {
        Caption = 'Shared Email';
        Implementation = "Email Connector" = "ALE Graph User Connector", "Default Email Rate Limit" = "ALE Graph User Connector";
    }
}

enumextension 70000 "FSO Email Connector" extends "Email Connector"
{
    value(70000; "Graph User")
    {
    Caption = 'Graph User';
    Implementation = "Email Connector"="FSO Graph User Connector", "Default Email Rate Limit"="FSO Graph User Connector";
    }
}

enumextension 70000 "Email Connector" extends "Email Connector"
{
    value(70000; "Graph User")
    {
        Caption = 'Graph User';
        Implementation = "Email Connector" = "Graph User Connector", "Default Email Rate Limit" = "Graph User Connector";
    }
}

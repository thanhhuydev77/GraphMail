page 70000 " Graph UserEmail Account"
{
    Caption = 'Graph UserEmail Account';
    InsertAllowed = false;
    PageType = NavigatePage;
    Permissions = tabledata "Email - Graph API Account" = rimd;
    SourceTable = "Email - Graph API Account";
    SourceTableTemporary = true;

    layout
    {
        area(Content)
        {
            group(New)
            {
                InstructionalText = 'Everyone will sendEmail messages from thisEmail account.';

                field("Email Address"; Rec."Email Address")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of theEmail Address field.';
                }
                field(Name; Rec.Name)
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Account Name field.';
                }
                field("Tenant ID"; Rec."Tenant ID")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the value of the Tenant ID field.';
                }
                field("Client ID"; Rec."Client ID")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the value of the Client ID field.';
                }
                field("Client Secrect"; Rec."Client Secrect")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the value of the Client Secrect field.';
                }
            }
        }
    }
    actions
    {
        area(Processing)
        {
            action("Next")
            {
                ApplicationArea = All;
                Caption = 'Next';
                Image = NextRecord;
                InFooterBar = true;
                ToolTip = 'Executes the Next action.';

                trigger OnAction()
                var
                    EmailGraphAPIAccount: Record "Email - Graph API Account";
                begin
                    EmailGraphAPIAccount.Init();
                    EmailGraphAPIAccount := Rec;
                    EmailGraphAPIAccount."Graph APIEmail Connector" := Enum::"Email Connector"::"Graph User";
                    AccountAdded := EmailGraphAPIAccount.Insert();
                    CurrPage.Close();
                end;
            }
        }
    }
    trigger OnOpenPage()
    begin
        NewMode := Rec."Email Address" = '';
        if NewMode then begin
            Rec.Init();
            Rec.ID := CreateGuid();
            Rec.Insert();
        end
    end;

    procedure GetAccount(var Account: Record "Email Account"): Boolean
    begin
        if AccountAdded then begin
            Account."Email Address" := Rec."Email Address";
            Account.Name := Rec.Name;
            Account.Connector := Enum::"Email Connector"::"Graph User";
            exit(true);
        end;
        exit(false);
    end;

    var
        AccountAdded: Boolean;
        NewMode: Boolean;
}

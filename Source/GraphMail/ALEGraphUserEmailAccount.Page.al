namespace GraphMail;

using System.Email;
using System.Azure.Identity;
page 70000 "ALE Graph User Email Account"
{
    Caption = 'Graph User Email Account';
    InsertAllowed = false;
    PageType = NavigatePage;
    Permissions = tabledata "ALE Email - Graph API Account" = rimd;
    SourceTable = "ALE Email - Graph API Account";
    SourceTableTemporary = true;

    layout
    {
        area(Content)
        {
            group(New)
            {
                InstructionalText = 'Everyone will send email messages from this email address.';

                field("Email Address"; Rec."Email Address")
                {
                    ApplicationArea = all;
                    ToolTip = 'Specifies the value of the Email Address field.';
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
                field("Client Secrect"; Rec."Client Secret")
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
                    EmailGraphAPIAccount: Record "ALE Email - Graph API Account";
                begin
                    EmailGraphAPIAccount.Init();
                    EmailGraphAPIAccount := Rec;
                    EmailGraphAPIAccount."Graph API Email Connector" := Enum::"Email Connector"::"Graph User";
                    AccountAdded := EmailGraphAPIAccount.Insert();
                    CurrPage.Close();
                end;
            }
        }
    }
    trigger OnOpenPage()
    var
        AzureADTenant: Codeunit "Azure AD Tenant";
    begin
        NewMode := Rec."Email Address" = '';
        if NewMode then begin
            Rec.Init();
            Rec.ID := CreateGuid();
            Rec."Tenant ID" := AzureADTenant.GetAadTenantId();
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

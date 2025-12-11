namespace GraphMail;

using System.Email;
using System.Text;
codeunit 70000 "ALE Graph User Connector" implements "Email Connector v4", "Default Email Rate Limit"
{
    Access = Internal;
    Permissions = tabledata "ALE Email - Graph API Account" = rimd;

    var
        EmailGraphAPIHelper: Codeunit "ALE Email - Graph API Helper";
        ConnectorDescriptionTxt: Label 'Users send emails from one shared email.';

    procedure Reply(var EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    begin
    end;

    procedure MarkAsRead(AccountId: Guid; ExternalMessageId: Text)
    begin
    end;

    procedure Send(EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    begin
        EmailGraphAPIHelper.Send(EmailMessage, AccountId);
    end;

    procedure RegisterAccount(var EmailAccount: Record "Email Account"): Boolean
    var
        CurrentUserEmailAccount: Page "ALE Graph User Email Account";
    begin
        CurrentUserEmailAccount.RunModal();
        exit(CurrentUserEmailAccount.GetAccount(EmailAccount));
    end;

    procedure ShowAccountInformation(AccountId: Guid);
    var
        EmailOutlookAccount: Record "ALE Email - Graph API Account";
    begin
        EmailOutlookAccount.SetRange("Graph API Email Connector", Enum::"Email Connector"::"Graph User");
        if EmailOutlookAccount.FindFirst() then Page.RunModal(Page::"ALE Graph User Email Account", EmailOutlookAccount);
    end;

    procedure GetAccounts(var EmailAccount: Record "Email Account")
    begin
        EmailGraphAPIHelper.GetAccounts(Enum::"Email Connector"::"Graph User", EmailAccount);
    end;

    procedure DeleteAccount(AccountId: Guid): Boolean
    begin
        exit(EmailGraphAPIHelper.DeleteAccount(AccountId));
    end;

    procedure GetDescription(): Text[250]
    begin
        exit(ConnectorDescriptionTxt);
    end;

    procedure GetLogoAsBase64(): Text
    var
        Base64Convert: codeunit "Base64 Convert";
        InS: InStream;
        ImageNameTok: label 'GraphAPIlogo.png';
    begin
        NavApp.GetResource(ImageNameTok, InS);
        exit(Base64Convert.ToBase64(InS));
    end;

    procedure GetDefaultEmailRateLimit(): Integer
    begin
        exit(EmailGraphAPIHelper.DefaultEmailRateLimit());
    end;

    procedure RetrieveEmails(AccountId: Guid; var EmailInbox: Record "Email Inbox"; var Filters: Record "Email Retrieval Filters" temporary)
    begin
    end;

    procedure GetEmailFolders(AccountId: Guid; var EmailFolders: Record "Email Folders" temporary)
    begin
    end;
}

namespace GraphMail;

using System.Email;
using System.RestClient;
using System.Text;
using System.Utilities;
using System.DataAdministration;
codeunit 70001 "ALE Email - Graph API Helper"
{
    Permissions = tabledata "Email Inbox" = ri,
        tabledata "ALE Email - Graph API Account" = rimd;

    var
        AccountNotFoundErr: Label 'We could not find the account. Typically, this is because the account has been deleted.';
        EmailBodyTooLargeErr: Label 'The Email is too large to send. The size limit is 4 MB, not including attachments.', Locked = true;
        gGraphURLTok: Label 'https://graph.microsoft.com/v1.0/users/', Locked = true;

    procedure GetAccounts(Connector: Enum "Email Connector"; var Accounts: Record "Email Account")
    var
        EmailGraphAPIAccount: Record "ALE Email - Graph API Account";
    begin
        EmailGraphAPIAccount.SetRange("Graph API Email Connector", Connector);
        if EmailGraphAPIAccount.FindSet() then
            repeat
                Accounts."Account Id" := EmailGraphAPIAccount.Id;
                Accounts."Email Address" := EmailGraphAPIAccount."Email Address";
                Accounts.Name := EmailGraphAPIAccount.Name;
                Accounts.Connector := Connector;
                Accounts.Insert();
            until EmailGraphAPIAccount.Next() = 0;
    end;

    procedure DeleteAccount(AccountId: Guid): Boolean
    var
        OutlookAccount: Record "ALE Email - Graph API Account";
    begin
        if OutlookAccount.Get(AccountId) then
            if OutlookAccount.WritePermission() then exit(OutlookAccount.Delete());
        exit(false);
    end;

    procedure EmailMessageToJson(EmailMessage: Codeunit "Email Message"; Account: Record "ALE Email - Graph API Account"): JsonObject
    var
        EmailAddressJson: JsonObject;
        EmailMessageJson: JsonObject;
        FromJson: JsonObject;
    begin
        EmailAddressJson.Add('address', Account."Email Address");
        EmailAddressJson.Add('name', Account."Name");
        FromJson.Add('emailAddress', EmailAddressJson);
        EmailMessageJson.Add('from', FromJson);
        exit(EmailMessageToJson(EmailMessage, EmailMessageJson))
    end;

    procedure AddEmailAttachments(EmailMessage: Codeunit "Email Message"; var MessageJson: JsonObject)
    var
        AttachmentsArray: JsonArray;
        AttachmentItemJson: JsonObject;
        AttachmentJson: JsonObject;
    begin
        if not EmailMessage.Attachments_First() then exit;
        repeat
            Clear(AttachmentJson);
            Clear(AttachmentItemJson);
            AttachmentJson.Add('name', EmailMessage.Attachments_GetName());
            AttachmentJson.Add('contentType', EmailMessage.Attachments_GetContentType());
            AttachmentJson.Add('isInline', EmailMessage.Attachments_IsInline());
            if EmailMessage.Attachments_GetLength() <= MaximumAttachmentSizeInBytes() then begin
                AttachmentJson.Add('@odata.type', '#microsoft.graph.fileAttachment');
                AttachmentJson.Add('contentBytes', EmailMessage.Attachments_GetContentBase64());
                AttachmentsArray.Add(AttachmentJson);
            end
            else begin
                AttachmentJson.Add('attachmentType', 'file');
                AttachmentJson.Add('size', EmailMessage.Attachments_GetLength());
                AttachmentJson.Add('contentBytes', EmailMessage.Attachments_GetContentBase64());
                AttachmentItemJson.Add('AttachmentItem', AttachmentJson);
                AttachmentsArray.Add(AttachmentItemJson);
            end;
        until EmailMessage.Attachments_Next() = 0;
        MessageJson.Add('attachments', AttachmentsArray);
    end;

    local procedure EmailMessageToJson(EmailMessage: Codeunit "Email Message"; EmailMessageJson: JsonObject): JsonObject
    var
        EmailBody: JsonObject;
        MessageJson: JsonObject;
        MessageText: Text;
    begin
        if EmailMessage.IsBodyHTMLFormatted() then
            EmailBody.Add('contentType', 'HTML')
        else
            EmailBody.Add('contentType', 'text');
        EmailBody.Add('content', EmailMessage.GetBody());
        EmailMessageJson.Add('subject', EmailMessage.GetSubject());
        EmailMessageJson.Add('body', EmailBody);
        EmailMessageJson.Add('toRecipients', GetEmailRecipients(EmailMessage, Enum::"Email Recipient Type"::"To"));
        EmailMessageJson.Add('ccRecipients', GetEmailRecipients(EmailMessage, Enum::"Email Recipient Type"::Cc));
        EmailMessageJson.Add('bccRecipients', GetEmailRecipients(EmailMessage, Enum::"Email Recipient Type"::Bcc));
        // If message json > max request size, then error as theEmail body is too large.
        EmailMessageJson.WriteTo(MessageText);
        if StrLen(MessageText) > MaximumRequestSizeInBytes() then Error(EmailBodyTooLargeErr);
        AddEmailAttachments(EmailMessage, EmailMessageJson);
        // If message json <= max request size, wrap it in message object to send in a single request.
        EmailMessageJson.WriteTo(MessageText);
        if StrLen(MessageText) > MaximumRequestSizeInBytes() then
            MessageJson := EmailMessageJson
        else begin
            MessageJson.Add('message', EmailMessageJson);
            MessageJson.Add('saveToSentItems', true);
        end;
        exit(MessageJson);
    end;

    local procedure GetEmailRecipients(EmailMessage: Codeunit "Email Message"; EmailRecipientType: enum "Email Recipient Type"): JsonArray
    var
        RecipientsJson: JsonArray;
        Address: JsonObject;
        EmailAddress: JsonObject;
        Recipients: List of [Text];
        Value: Text;
    begin
        EmailMessage.GetRecipients(EmailRecipientType, Recipients);
        foreach value in Recipients do begin
            clear(Address);
            clear(EmailAddress);
            Address.Add('address', value);
            EmailAddress.Add('emailAddress', Address);
            RecipientsJson.Add(EmailAddress);
        end;
        exit(RecipientsJson);
    end;

    procedure Send(EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    var
        EmailGraphAPIAccount: Record "ALE Email - Graph API Account";
        OauthHeader: Codeunit HttpAuthOAuthClientCredentials;
    begin
        if not EmailGraphAPIAccount.Get(AccountId) then Error(AccountNotFoundErr);
        // OauthHeader := GetAccessTokenByEmailAccount(EmailGraphAPIAccount);
        // SendEmail(OauthHeader, EmailMessageToJson(EmailMessage, EmailGraphAPIAccount), EmailGraphAPIAccount."Email Address");
        OauthHeader := GetOAuth2HeaderByEmail(EmailGraphAPIAccount);
        SendEmail(OauthHeader, EmailMessageToJson(EmailMessage, EmailGraphAPIAccount), EmailGraphAPIAccount."Email Address");
    end;

    procedure GetOAuth2HeaderByEmail(EmailGraphAPIAccount: Record "ALE Email - Graph API Account") HttpAuthOAuthClientCredentials: Codeunit HttpAuthOAuthClientCredentials
    var
        AuthorizeBaseUrlLbl: Label 'https://login.microsoftonline.com/';
        ScopeLbl: Label 'https://graph.microsoft.com/.default';
        ListScope: List of [Text];
        TenantID: Text;
    begin
        TenantID := RemoveParenthesis(EmailGraphAPIAccount."Tenant ID".ToText());
        ListScope.Add(ScopeLbl);
        HttpAuthOAuthClientCredentials.Initialize(AuthorizeBaseUrlLbl + TenantID, EmailGraphAPIAccount."Client ID", EmailGraphAPIAccount."Client Secret", ListScope);
    end;

    procedure RemoveParenthesis(pInputTxt: Text): Text
    begin
        pInputTxt := DELCHR(pInputTxt, '=', '{');
        pInputTxt := DELCHR(pInputTxt, '=', '}');
        exit(pInputTxt);
    end;

    procedure SendEmail(OauthHeader: Codeunit HttpAuthOAuthClientCredentials; MessageJson: JsonObject; FromEmail: Text[250])
    var
        Attachments: JsonArray;
        Attachment: JsonToken;
        JToken: JsonToken;
        MessageId: Text;
    begin
        if MessageJson.Contains('message') then
            SendMailSingleRequest(OauthHeader, MessageJson, FromEmail)
        else begin
            MessageJson.Get('attachments', JToken);
            Attachments := JToken.AsArray();
            MessageJson.Remove('attachments');
            MessageId := CreateDraftMail(OauthHeader, MessageJson, FromEmail);
            foreach Attachment in Attachments do
                if Attachment.AsObject().Contains('AttachmentItem') then
                    UploadAttachment(OauthHeader, FromEmail, Attachment.AsObject(), MessageId)
                else
                    PostAttachment(OauthHeader, FromEmail, Attachment.AsObject(), MessageId);
            SendDraftMail(OauthHeader, MessageId, FromEmail);
        end;
    end;

    local procedure CreateDraftMail(OauthHeader: Codeunit HttpAuthOAuthClientCredentials; MessageJson: JsonObject; FromEmail: Text): Text
    var
        RestClient: Codeunit "Rest Client";
        ResponseJson: JsonObject;
        JToken: JsonToken;
        MessageId: Text;
        MessageJsonText: Text;
        RequestUri: Text;
    begin
        MessageJson.WriteTo(MessageJsonText);
        RequestUri := gGraphURLTok + FromEmail + '/sendMail';

        RestClient.Initialize(OauthHeader);
        JToken := RestClient.PostAsJson(RequestUri, MessageJson);
        ResponseJson.Get('id', JToken);
        MessageId := JToken.AsValue().AsText();
        exit(MessageId);
    end;

    local procedure PostAttachment(OauthHeader: Codeunit HttpAuthOAuthClientCredentials; EmailAddress: Text[250]; AttachmentJson: JsonObject; MessageId: Text)
    var
        RestClient: Codeunit "Rest Client";
        PostAttachmentUriTxt: Label '%1/messages/%2/attachments', Locked = true;
        RequestUri: Text;
    begin
        RequestUri := gGraphURLTok + StrSubstNo(PostAttachmentUriTxt, EmailAddress, MessageId);
        RestClient.Initialize(OauthHeader);
        RestClient.PostAsJson(RequestUri, AttachmentJson);
    end;

    local procedure SendDraftMail(OauthHeader: Codeunit HttpAuthOAuthClientCredentials; MessageId: Text; FromEmail: Text): Text
    var
        RestClient: Codeunit "Rest Client";
        MailHttpContent: Codeunit "Http Content";
        RequestUri: Text;
    begin
        RestClient.Initialize(OauthHeader);
        RequestUri := gGraphURLTok + FromEmail + '/messages/' + MessageId + '/send';
        RestClient.Post(RequestUri, MailHttpContent);
    end;

    local procedure SendMailSingleRequest(OauthHeader: Codeunit HttpAuthOAuthClientCredentials; MessageJson: JsonObject; FromEmail: Text)
    var
        RestClient: Codeunit "Rest Client";
        MessageJsonText: Text;
        RequestUri: Text;
    begin
        MessageJson.WriteTo(MessageJsonText);
        RequestUri := gGraphURLTok + FromEmail + '/sendMail';
        RestClient.Initialize(OauthHeader);
        RestClient.PostAsJson(RequestUri, MessageJson);
    end;

    local procedure UploadAttachment(OauthHeader: Codeunit HttpAuthOAuthClientCredentials; EmailAddress: Text[250]; AttachmentJson: JsonObject; MessageId: Text)
    var
        RestClient: Codeunit "Rest Client";
        Base64Convert: Codeunit "Base64 Convert";
        AttachmentTempBlob: Codeunit "Temp Blob";
        jtoken: JsonToken;
        AttachmentInStream: Instream;
        FromByte, Range, ToByte, TotalBytes : Integer;

        UploadAttachmentMeUriTxt: Label '/messages/%1/attachments/createUploadSession', Locked = true;
        AttachmentOutStream: OutStream;
        AttachmentContentInBase64, RequestUri, UploadUrl : Text;
    begin
        RequestUri := gGraphURLTok + EmailAddress + StrSubstNo(UploadAttachmentMeUriTxt, MessageId);
        RestClient.Initialize(OauthHeader);
        AttachmentContentInBase64 := GetAttachmentContent(AttachmentJson);

        jtoken := RestClient.PostAsJson(RequestUri, AttachmentJson);
        UploadUrl := GetUploadUrl(jtoken);
        FromByte := 0;
        TotalBytes := GetAttachmentSize(AttachmentJson);
        Range := MaximumAttachmentSizeInBytes();
        AttachmentTempBlob.CreateOutStream(AttachmentOutStream);
        Base64Convert.FromBase64(AttachmentContentInBase64, AttachmentOutStream);
        AttachmentTempBlob.CreateInStream(AttachmentInStream);
        while FromByte < TotalBytes do begin
            ToByte := FromByte + Range - 1;
            if ToByte >= TotalBytes then begin
                ToByte := TotalBytes - 1;
                Range := ToByte - FromByte + 1;
            end;
            UploadAttachmentRange(UploadUrl, AttachmentInStream, FromByte, ToByte, TotalBytes, Range);
            FromByte := ToByte + 1;
        end;
    end;

    local procedure GetAttachmentSize(AttachmentJson: JsonObject): Integer
    var
        JToken: JsonToken;
    begin
        AttachmentJson.Get('AttachmentItem', JToken);
        JToken.AsObject().Get('size', JToken);
        exit(JToken.AsValue().AsInteger());
    end;

    local procedure GetUploadUrl(AttachmentHttpResponseMessage: JsonToken): Text
    var
        ResponseJson: JsonObject;
        JToken: JsonToken;
    begin
        ResponseJson := AttachmentHttpResponseMessage.AsObject();
        ResponseJson.Get('uploadUrl', JToken);
        exit(JToken.AsValue().AsText());
    end;

    local procedure GetAttachmentContent(var AttachmentJson: JsonObject): Text
    var
        JToken: JsonToken;
        JTokenContent: JsonToken;
        AttachmentContentInBase64: Text;
    begin
        AttachmentJson.Get('AttachmentItem', JToken);
        JToken.AsObject().Get('contentBytes', JTokenContent);
        AttachmentContentInBase64 := JTokenContent.AsValue().AsText();
        Jtoken.AsObject().Remove('contentBytes');
        exit(AttachmentContentInBase64);
    end;

    local procedure UploadAttachmentRange(UploadUrl: Text; AttachmentInStream: InStream; FromByte: Integer; ToByte: Integer; TotalBytes: Integer; Range: Integer)
    var
        AttachmentRangeTempBlob: Codeunit "Temp Blob";
        AttachmentHttpClient: HttpClient;
        AttachmentHttpContent: HttpContent;
        AttachmentHttpContentHeaders: HttpHeaders;
        AttachmentHttpRequestHeaders: HttpHeaders;
        AttachmentHttpRequestMessage: HttpRequestMessage;
        AttachmentHttpResponseMessage: HttpResponseMessage;
        AttachmentRangeInStream: Instream;
        ContentLength: Integer;
        ContentRangeLbl: Label 'bytes %1-%2/%3', Comment = '%1 - From byte, %2 - To byte, %3 - Total bytes', Locked = true;
        SendEmailErr: Label 'Could not send the Email message. Try again later.';
        AttachmentOutStream: OutStream;
        HttpErrorMessage: Text;
    begin
        AttachmentRangeTempBlob.CreateOutStream(AttachmentOutStream);
        CopyStream(AttachmentOutStream, AttachmentInStream, Range); // copy range of bytes to upload
        AttachmentRangeTempBlob.CreateInStream(AttachmentRangeInStream);
        AttachmentHttpRequestMessage.Method('PUT');
        AttachmentHttpRequestMessage.SetRequestUri(UploadUrl);
        AttachmentHttpContent.WriteFrom(AttachmentRangeInStream);
        AttachmentHttpContent.GetHeaders(AttachmentHttpContentHeaders);
        AttachmentHttpContentHeaders.Clear();
        ContentLength := ToByte - FromByte + 1;
        AttachmentHttpContentHeaders.Add('Content-Type', 'application/octet-stream');
        AttachmentHttpContentHeaders.Add('Content-Length', Format(ContentLength));
        AttachmentHttpContentHeaders.Add('Content-Range', StrSubstNo(ContentRangeLbl, FromByte, ToByte, TotalBytes));
        AttachmentHttpRequestMessage.Content := AttachmentHttpContent;
        AttachmentHttpRequestMessage.GetHeaders(AttachmentHttpRequestHeaders);
        AttachmentHttpRequestHeaders.Clear();
        AttachmentHttpRequestHeaders.Add('Keep-alive', 'true');
        if not AttachmentHttpClient.Send(AttachmentHttpRequestMessage, AttachmentHttpResponseMessage) then Error(SendEmailErr);
        if AttachmentHttpResponseMessage.HttpStatusCode <> 200 then begin
            if AttachmentHttpResponseMessage.HttpStatusCode = 201 then exit;
            HttpErrorMessage := GetHttpErrorMessageAsText(AttachmentHttpResponseMessage);
            Error(HttpErrorMessage);
        end;
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Environment Cleanup", 'OnClearCompanyConfig', '', false, false)]
    local procedure ClearCompanyConfigGeneral(CompanyName: Text; SourceEnv: Enum "Environment Type"; DestinationEnv: Enum "Environment Type")
    var
        EmailGraphAPIAccount: Record "ALE Email - Graph API Account";
    begin
        EmailGraphAPIAccount.DeleteAll();
    end;

    local procedure MaximumRequestSizeInBytes(): Integer
    begin
        exit(4194304); // 4 mb
    end;

    local procedure MaximumAttachmentSizeInBytes(): Integer
    begin
        exit(3145728); // 3 mb
    end;

    procedure DefaultEmailRateLimit(): Integer
    begin
        exit(30);
    end;

    local procedure GetHttpErrorMessageAsText(MailHttpResponseMessage: HttpResponseMessage): Text
    var
        SendEmailErr: Label 'Could not send the Email message. Try again later.';
        ErrorMessage: Text;
    begin
        if not TryGetErrorMessage(MailHttpResponseMessage, ErrorMessage) then ErrorMessage := SendEmailErr;
        exit(ErrorMessage);
    end;

    [TryFunction]
    local procedure TryGetErrorMessage(MailHttpResponseMessage: HttpResponseMessage; var ErrorMessage: Text);
    var
        ResponseJson: JsonObject;
        JToken: JsonToken;
        ResponseJsonText: Text;
    begin
        MailHttpResponseMessage.Content.ReadAs(ResponseJsonText);
        ResponseJson.ReadFrom(ResponseJsonText);
        ResponseJson.Get('error', JToken);
        JToken.AsObject().Get('message', JToken);
        ErrorMessage := JToken.AsValue().AsText();
    end;
}

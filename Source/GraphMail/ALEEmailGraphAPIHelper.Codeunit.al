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

    procedure EmailMessageToReplyJson(EmailMessage: Codeunit "Email Message"; ReplyBody: Text; AsHtml: Boolean): JsonObject
    var
        Recipients: JsonArray;
        EmailBody: JsonObject;
        EmailMessageJson: JsonObject;
        MessageText: Text;
    begin
        EmailBody.add('content', ReplyBody);
        if AsHtml then
            EmailBody.add('contentType', 'html')
        else
            EmailBody.add('contentType', 'text');
        EmailMessageJson.Add('body', EmailBody);
        Recipients := GetEmailRecipients(EmailMessage, Enum::"Email Recipient Type"::"To");
        if Recipients.Count > 0 then EmailMessageJson.Add('toRecipients', Recipients);
        Recipients := GetEmailRecipients(EmailMessage, Enum::"Email Recipient Type"::Cc);
        if Recipients.Count > 0 then EmailMessageJson.Add('ccRecipients', Recipients);
        Recipients := GetEmailRecipients(EmailMessage, Enum::"Email Recipient Type"::Bcc);
        if Recipients.Count > 0 then EmailMessageJson.Add('bccRecipients', Recipients);
        // If message json > max request size, then error as theEmail body is too large.
        EmailMessageJson.WriteTo(MessageText);
        if StrLen(MessageText) > MaximumRequestSizeInBytes() then Error(EmailBodyTooLargeErr);
        AddEmailAttachments(EmailMessage, EmailMessageJson);
        exit(EmailMessageJson);
    end;

    procedure EmailMessageToJson(EmailMessage: Codeunit "Email Message"): JsonObject
    var
        EmailMessageJson: JsonObject;
    begin
        exit(EmailMessageToJson(EmailMessage, EmailMessageJson));
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
        OauthHeader: Dictionary of [Text, SecretText];
    begin
        if not EmailGraphAPIAccount.Get(AccountId) then Error(AccountNotFoundErr);
        OauthHeader := GetAccessTokenByEmailAccount(EmailGraphAPIAccount);
        SendEmail(OauthHeader, EmailMessageToJson(EmailMessage, EmailGraphAPIAccount), EmailGraphAPIAccount."Email Address");
    end;

    procedure GetAccessTokenByEmailAccount(EmailGraphAPIAccount: Record "ALE Email - Graph API Account") OauthHeader: Dictionary of [Text, SecretText]
    var
        HttpAuthOAuthClientCredentials: Codeunit HttpAuthOAuthClientCredentials;
        AuthorizeBaseUrlLbl: Label 'https://login.microsoftonline.com/';
        ScopeLbl: Label 'https://graph.microsoft.com/.default';
        ListScope: List of [Text];
        TenantID: Text;
    begin
        TenantID := RemoveParenthesis(EmailGraphAPIAccount."Tenant ID".ToText());
        ListScope.Add(ScopeLbl);
        HttpAuthOAuthClientCredentials.Initialize(AuthorizeBaseUrlLbl + TenantID, EmailGraphAPIAccount."Client ID", EmailGraphAPIAccount."Client Secrect", ListScope);
        OauthHeader := HttpAuthOAuthClientCredentials.GetAuthorizationHeaders();
    end;

    procedure RemoveParenthesis(pInputTxt: Text): Text
    begin
        pInputTxt := DELCHR(pInputTxt, '=', '{');
        pInputTxt := DELCHR(pInputTxt, '=', '}');
        exit(pInputTxt);
    end;

    procedure SendEmail(OauthHeader: Dictionary of [Text, SecretText]; MessageJson: JsonObject; FromEmail: Text[250])
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

    local procedure CreateDraftMail(OauthHeader: Dictionary of [Text, SecretText]; MessageJson: JsonObject; FromEmail: Text): Text
    var
        MailHttpClient: HttpClient;
        MailHttpContent: HttpContent;
        MailContentHeaders: HttpHeaders;
        MailRequestHeaders: HttpHeaders;
        MailHttpRequestMessage: HttpRequestMessage;
        MailHttpResponseMessage: HttpResponseMessage;
        ResponseJson: JsonObject;
        JToken: JsonToken;
        EnvironmentBlocksErr: Label 'The request to sendEmail has been blocked. To resolve the problem, enable outgoing HTTP requests for theEmail - Outlook REST API app on the Extension Management page.';
        GraphURLTxt: Label 'https://graph.microsoft.com/v1.0/users/';
        HttpErrorMessage: Text;
        MessageId: Text;
        MessageJsonText: Text;
        RequestUri: Text;
        ResponseJsonText: Text;
    begin
        MessageJson.WriteTo(MessageJsonText);
        RequestUri := GraphURLTxt + FromEmail + '/sendMail';
        MailHttpRequestMessage.Method('POST');
        MailHttpRequestMessage.SetRequestUri(RequestUri);
        MailHttpRequestMessage.GetHeaders(MailRequestHeaders);
        MailRequestHeaders.Add('Authorization', OauthHeader.Get('Authorization'));
        MailHttpContent.WriteFrom(MessageJsonText);
        MailHttpContent.GetHeaders(MailContentHeaders);
        MailContentHeaders.Clear();
        MailContentHeaders.Add('Content-Type', 'application/json');
        MailHttpRequestMessage.Content := MailHttpContent;
        if not MailHttpClient.Send(MailHttpRequestMessage, MailHttpResponseMessage) then
            if MailHttpResponseMessage.IsBlockedByEnvironment() then Error(EnvironmentBlocksErr);
        if MailHttpResponseMessage.HttpStatusCode <> 201 then begin
            HttpErrorMessage := GetHttpErrorMessageAsText(MailHttpResponseMessage);
            Error(HttpErrorMessage);
        end
        else begin
            MailHttpResponseMessage.Content.ReadAs(ResponseJsonText);
            ResponseJson.ReadFrom(ResponseJsonText);
            ResponseJson.Get('id', JToken);
            MessageId := JToken.AsValue().AsText();
        end;
        exit(MessageId);
    end;

    local procedure PostAttachment(OauthHeader: Dictionary of [Text, SecretText]; EmailAddress: Text[250]; AttachmentJson: JsonObject; MessageId: Text)
    var
        AttachmentHttpClient: HttpClient;
        AttachmentHttpContent: HttpContent;
        AttachmentContentHeaders: HttpHeaders;
        AttachmentRequestHeaders: HttpHeaders;
        AttachmentHttpRequestMessage: HttpRequestMessage;
        AttachmentHttpResponseMessage: HttpResponseMessage;
        GraphURLTxt: Label 'https://graph.microsoft.com';
        PostAttachmentUriTxt: Label '/v1.0/users/%1/messages/%2/attachments', Locked = true;
        SendEmailErr: Label 'Could not send theEmail message. Try again later.';
        AttachmentRequestJsonText: Text;
        HttpErrorMessage: Text;
        RequestUri: Text;
    begin
        RequestUri := GraphURLTxt + StrSubstNo(PostAttachmentUriTxt, EmailAddress, MessageId);
        AttachmentHttpRequestMessage.Method('POST');
        AttachmentHttpRequestMessage.SetRequestUri(RequestUri);
        AttachmentHttpRequestMessage.GetHeaders(AttachmentRequestHeaders);
        AttachmentRequestHeaders.Add('Authorization', OauthHeader.Get('Authorization'));
        AttachmentJson.WriteTo(AttachmentRequestJsonText);
        AttachmentHttpContent.WriteFrom(AttachmentRequestJsonText);
        AttachmentHttpContent.GetHeaders(AttachmentContentHeaders);
        AttachmentContentHeaders.Clear();
        AttachmentContentHeaders.Add('Content-Type', 'application/json');
        AttachmentHttpRequestMessage.Content := AttachmentHttpContent;
        if not AttachmentHttpClient.Send(AttachmentHttpRequestMessage, AttachmentHttpResponseMessage) then Error(SendEmailErr);
        if AttachmentHttpResponseMessage.HttpStatusCode <> 201 then begin
            HttpErrorMessage := GetHttpErrorMessageAsText(AttachmentHttpResponseMessage);
            Error(HttpErrorMessage);
        end;
    end;

    local procedure SendDraftMail(OauthHeader: Dictionary of [Text, SecretText]; MessageId: Text; FromEmail: Text): Text
    var
        MailHttpClient: HttpClient;
        MailHttpContent: HttpContent;
        MailContentHeaders: HttpHeaders;
        MailRequestHeaders: HttpHeaders;
        MailHttpRequestMessage: HttpRequestMessage;
        MailHttpResponseMessage: HttpResponseMessage;
        GraphURLTxt: Label 'https://graph.microsoft.com';
        SendEmailErr: Label 'Could not send theEmail message. Try again later.';
        HttpErrorMessage: Text;
        RequestUri: Text;
    begin
        RequestUri := GraphURLTxt + '/v1.0/users/' + FromEmail + '/messages/' + MessageId + '/send';
        MailHttpRequestMessage.Method('POST');
        MailHttpRequestMessage.SetRequestUri(RequestUri);
        MailHttpRequestMessage.GetHeaders(MailRequestHeaders);
        MailRequestHeaders.Add('Authorization', OauthHeader.Get('Authorization'));
        MailHttpContent.GetHeaders(MailContentHeaders);
        MailContentHeaders.Clear();
        MailContentHeaders.Add('Content-Length', '0');
        if not MailHttpClient.Send(MailHttpRequestMessage, MailHttpResponseMessage) then Error(SendEmailErr);
        if MailHttpResponseMessage.HttpStatusCode <> 202 then begin
            HttpErrorMessage := GetHttpErrorMessageAsText(MailHttpResponseMessage);
            Error(HttpErrorMessage);
        end;
    end;

    local procedure SendMailSingleRequest(OauthHeader: Dictionary of [Text, SecretText]; MessageJson: JsonObject; FromEmail: Text)
    var
        MailHttpClient: HttpClient;
        MailHttpContent: HttpContent;
        MailContentHeaders: HttpHeaders;
        MailRequestHeaders: HttpHeaders;
        MailHttpRequestMessage: HttpRequestMessage;
        MailHttpResponseMessage: HttpResponseMessage;
        EnvironmentBlocksErr: Label 'The request to sendEmail has been blocked. To resolve the problem, enable outgoing HTTP requests for theEmail - Outlook REST API app on the Extension Management page.';
        GraphURLTxt: Label 'https://graph.microsoft.com/v1.0/users/';
        MessageJsonText: Text;
        RequestUri: Text;
    begin
        MessageJson.WriteTo(MessageJsonText);
        RequestUri := GraphURLTxt + FromEmail + '/sendMail';
        MailHttpRequestMessage.Method('POST');
        MailHttpRequestMessage.SetRequestUri(RequestUri);
        MailHttpRequestMessage.GetHeaders(MailRequestHeaders);
        MailRequestHeaders.Add('Authorization', OauthHeader.Get('Authorization'));
        MailHttpContent.WriteFrom(MessageJsonText);
        MailHttpContent.GetHeaders(MailContentHeaders);
        MailContentHeaders.Clear();
        MailContentHeaders.Add('Content-Type', 'application/json');
        MailHttpRequestMessage.Content := MailHttpContent;
        if not MailHttpClient.Send(MailHttpRequestMessage, MailHttpResponseMessage) then;
        if MailHttpResponseMessage.IsBlockedByEnvironment() then Error(EnvironmentBlocksErr)
    end;

    local procedure UploadAttachment(OauthHeader: Dictionary of [Text, SecretText]; EmailAddress: Text[250]; AttachmentJson: JsonObject; MessageId: Text)
    var
        Base64Convert: Codeunit "Base64 Convert";
        AttachmentTempBlob: Codeunit "Temp Blob";
        AttachmentHttpClient: HttpClient;
        AttachmentHttpContent: HttpContent;
        AttachmentContentHeaders: HttpHeaders;
        AttachmentRequestHeaders: HttpHeaders;
        AttachmentHttpRequestMessage: HttpRequestMessage;
        AttachmentHttpResponseMessage: HttpResponseMessage;
        AttachmentInStream: Instream;
        FromByte, Range, ToByte, TotalBytes : Integer;
        GraphURLTxt: Label 'https://graph.microsoft.com/v1.0/users/';
        SendEmailErr: Label 'Could not send theEmail message. Try again later.';
        UploadAttachmentMeUriTxt: Label '/messages/%1/attachments/createUploadSession', Locked = true;
        AttachmentOutStream: OutStream;
        AttachmentContentInBase64, HttpErrorMessage, RequestJsonText, RequestUri, UploadUrl : Text;
    begin
        RequestUri := GraphURLTxt + EmailAddress + StrSubstNo(UploadAttachmentMeUriTxt, MessageId);
        AttachmentContentInBase64 := GetAttachmentContent(AttachmentJson);
        AttachmentJson.WriteTo(RequestJsonText);
        AttachmentHttpRequestMessage.Method('POST');
        AttachmentHttpRequestMessage.SetRequestUri(RequestUri);
        AttachmentHttpRequestMessage.GetHeaders(AttachmentRequestHeaders);
        AttachmentRequestHeaders.Add('Authorization', OauthHeader.Get('Authorization'));
        AttachmentHttpContent.WriteFrom(RequestJsonText);
        AttachmentHttpContent.GetHeaders(AttachmentContentHeaders);
        AttachmentContentHeaders.Clear();
        AttachmentContentHeaders.Add('Content-Type', 'application/json');
        AttachmentHttpRequestMessage.Content := AttachmentHttpContent;
        if not AttachmentHttpClient.Send(AttachmentHttpRequestMessage, AttachmentHttpResponseMessage) then Error(SendEmailErr);
        if AttachmentHttpResponseMessage.HttpStatusCode <> 201 then begin
            HttpErrorMessage := GetHttpErrorMessageAsText(AttachmentHttpResponseMessage);
            Error(HttpErrorMessage);
        end
        else
            UploadUrl := GetUploadUrl(AttachmentHttpResponseMessage);
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

    local procedure GetUploadUrl(AttachmentHttpResponseMessage: HttpResponseMessage): Text
    var
        ResponseJson: JsonObject;
        JToken: JsonToken;
        ResponseJsonText: Text;
    begin
        AttachmentHttpResponseMessage.Content.ReadAs(ResponseJsonText);
        ResponseJson.ReadFrom(ResponseJsonText);
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
        SendEmailErr: Label 'Could not send theEmail message. Try again later.';
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

    procedure MarkEmailAsRead(AccountId: Guid; ExternalMessageId: Text)
    var
    begin
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
        SendEmailErr: Label 'Could not send theEmail message. Try again later.';
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

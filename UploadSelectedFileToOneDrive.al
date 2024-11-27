pageextension 50100 DocumentAttachmentDetailsExt extends "Document Attachment Details"
{
    actions
    {
        addafter(UploadFile)
        {
            action(UploadFileToOneDrive)
            {
                Caption = 'Upload File to OneDrive';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = GetActionMessages;

                trigger OnAction()
                var
                    OneDriveHandler: Codeunit OneDriveHandler;
                    DocAttach: Record "Document Attachment";
                begin
                    DocAttach.Reset();
                    CurrPage.SetSelectionFilter(DocAttach);
                    if DocAttach.FindFirst() then
                        OneDriveHandler.UploadFilesToOneDrive(DocAttach);
                end;
            }
        }
    }
}

codeunit 50120 OneDriveHandler
{
    procedure UploadFilesToOneDrive(DocAttach: Record "Document Attachment")
    var
        HttpClient: HttpClient;
        HttpRequestMessage: HttpRequestMessage;
        HttpResponseMessage: HttpResponseMessage;
        Headers: HttpHeaders;
        ContentHeader: HttpHeaders;
        RequestContent: HttpContent;
        JsonResponse: JsonObject;
        AuthToken: SecretText;
        OneDriveFileUrl: Text;
        ResponseText: Text;
        OutStream: OutStream;
        FileContent: InStream;
        TempBlob: Codeunit "Temp Blob";
        FileName: Text;
        TenantMedia: Record "Tenant Media";
        MimeType: Text;
    begin
        // Get OAuth token
        AuthToken := GetOAuthToken();

        if AuthToken.IsEmpty() then
            Error('Failed to obtain access token.');

        if DocAttach."Document Reference ID".HasValue then begin
            TempBlob.CreateOutStream(OutStream);
            DocAttach."Document Reference ID".ExportStream(OutStream);
            TempBlob.CreateInStream(FileContent);
            FileName := DocAttach."File Name" + '.' + DocAttach."File Extension";
            if TenantMedia.Get(DocAttach."Document Reference ID".MediaId) then
                MimeType := TenantMedia."Mime Type";
        end;

        // Define the OneDrive folder URL

        // delegated permissions
        //OneDriveFileUrl := 'https://graph.microsoft.com/v1.0/me/drive/root/children';

        // application permissions (replace with the actual user principal name)
        OneDriveFileUrl := 'https://graph.microsoft.com/v1.0/users/Admin@2qcj3x.onmicrosoft.com/drive/root:/OneDriveAPITest/' + FileName + ':/content';
        // Initialize the HTTP request
        HttpRequestMessage.SetRequestUri(OneDriveFileUrl);
        HttpRequestMessage.Method := 'PUT';
        HttpRequestMessage.GetHeaders(Headers);
        Headers.Add('Authorization', SecretStrSubstNo('Bearer %1', AuthToken));
        RequestContent.GetHeaders(ContentHeader);
        ContentHeader.Clear();
        ContentHeader.Add('Content-Type', MimeType);
        HttpRequestMessage.Content.WriteFrom(FileContent);

        // Send the HTTP request
        if HttpClient.Send(HttpRequestMessage, HttpResponseMessage) then begin
            // Log the status code for debugging
            //Message('HTTP Status Code: %1', HttpResponseMessage.HttpStatusCode());

            if HttpResponseMessage.IsSuccessStatusCode() then begin
                HttpResponseMessage.Content.ReadAs(ResponseText);
                JsonResponse.ReadFrom(ResponseText);
                Message(ResponseText);
            end else begin
                //Report errors!
                HttpResponseMessage.Content.ReadAs(ResponseText);
                Error('Failed to upload files to OneDrive: %1 %2', HttpResponseMessage.HttpStatusCode(), ResponseText);
            end;
        end else
            Error('Failed to send HTTP request to OneDrive');
    end;

    procedure GetOAuthToken() AuthToken: SecretText
    var
        ClientID: Text;
        ClientSecret: Text;
        TenantID: Text;
        AccessTokenURL: Text;
        OAuth2: Codeunit OAuth2;
        Scopes: List of [Text];
    begin
        ClientID := 'b4fe1687-f1ab-4bfa-b494-0e2236ed50bd';
        ClientSecret := 'huL8Q~edsQZ4pwyxka3f7.WUkoKNcPuqlOXv0bww';
        TenantID := '7e47da45-7f7d-448a-bd3d-1f4aa2ec8f62';
        AccessTokenURL := 'https://login.microsoftonline.com/' + TenantID + '/oauth2/v2.0/token';
        Scopes.Add('https://graph.microsoft.com/.default');
        if not OAuth2.AcquireTokenWithClientCredentials(ClientID, ClientSecret, AccessTokenURL, '', Scopes, AuthToken) then
            Error('Failed to get access token from response\%1', GetLastErrorText());
    end;
}

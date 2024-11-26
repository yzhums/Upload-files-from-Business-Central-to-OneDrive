pageextension 50100 CustomerListExt extends "Customer List"
{
    actions
    {
        addafter("Sent Emails")
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
                begin
                    OneDriveHandler.UploadFilesToOneDrive();
                end;
            }
        }
    }
}

codeunit 50120 OneDriveHandler
{
    procedure UploadFilesToOneDrive()
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
        DocAttach: Record "Document Attachment";
        TempBlob: Codeunit "Temp Blob";
    begin
        // Get OAuth token
        AuthToken := GetOAuthToken();

        if AuthToken.IsEmpty() then
            Error('Failed to obtain access token.');

        if DocAttach.Get(18, 10000, Enum::"Attachment Document Type"::"Quote", 0, 1) then
            if DocAttach."Document Reference ID".HasValue then begin
                TempBlob.CreateOutStream(OutStream);
                DocAttach."Document Reference ID".ExportStream(OutStream);
                TempBlob.CreateInStream(FileContent);
            end;

        // Define the OneDrive folder URL

        // delegated permissions
        //OneDriveFileUrl := 'https://graph.microsoft.com/v1.0/me/drive/root/children';

        // application permissions (replace with the actual user principal name)
        OneDriveFileUrl := 'https://graph.microsoft.com/v1.0/users/Admin@2qcj3x.onmicrosoft.com/drive/root:/OneDriveAPITest/OneDriveUploadTest.csv:/content';
        // Initialize the HTTP request
        HttpRequestMessage.SetRequestUri(OneDriveFileUrl);
        HttpRequestMessage.Method := 'PUT';
        HttpRequestMessage.GetHeaders(Headers);
        Headers.Add('Authorization', SecretStrSubstNo('Bearer %1', AuthToken));
        RequestContent.GetHeaders(ContentHeader);
        ContentHeader.Clear();
        ContentHeader.Add('Content-Type', 'text/csv');
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

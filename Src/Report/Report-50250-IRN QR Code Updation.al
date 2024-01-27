report 50250 "IRN QR Code Updation"
{
    Permissions = TableData 112 = rm,
                  TableData 114 = rm,
                  TableData 124 = rm,
                  TableData 5744 = rm,
                  TableData "GST Ledger Entry" = rm;
    ProcessingOnly = true;
    UseRequestPage = false;
    UseSystemPrinter = false;

    dataset
    {
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(IRN; IRN)
                {
                    ApplicationArea = all;
                    Visible = false;
                }
            }
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPostReport()
    begin

        //CUEinvoice.SetLocation(LocRec.Code);
        //CUEinvoice.GetToken(LocRec."E-Inv UserName",LocRec."E-Inv Password",LocRec."E-Inv ForceRefreshAccessToken",ClientId,AuthToken,Sek,TokenExpiry,decryptedSek
        //,LocRec."E-Inv ClientId",LocRec."E-Inv Client_Secret");

        //DJ Govt API Comment
        /*
        LocRec.GET(LocCode);
        GetIRNDetails(LocRec."E-Inv Token",LocRec."E-Inv Sek",IRN,LocRec."E-Inv ClientId",LocRec."E-Inv Client_Secret",LocRec."GST Registration No.",LocRec."E-Inv UserName");
        MESSAGE('IRN Details Updated Successfully');
        */
        //DJ Govt API Comment

        //DJ GSP API
        LocRec.GET(LocCode);
        // CUEinvoice.SetLocation(LocRec.Code);
        //  CUEinvoice.GetToken(LocRec."E-Inv ClientId", LocRec."E-Inv Client_Secret", AuthToken, TokenExpiry);
        GetIRNDetails(AuthToken, Sek, IRN, LocRec."E-Inv ClientId", LocRec."E-Inv Client_Secret", LocRec."GST Registration No.", LocRec."E-Inv UserName", LocRec."E-Inv Password");
        MESSAGE('IRN Details Updated Successfully');
        //DJ GSP API

    end;

    var
        LocRec: Record 14;
        // CUEinvoice: Codeunit 50011;
        ClientId: Text;
        AuthToken: Text;
        Sek: Text;
        Client_ID: Text;
        Client_Secret: Text;
        TokenExpiry: Text;
        decryptedSek: Text;
        GetIrnDetailsText1: Label 'https://api.einvoice1.gst.gov.in/eicore/v1.03/Invoice/irn/';
        GetIrnDetailsText: Label 'https://gsp.adaequare.com/enriched/ei/api/invoice/irn?irn=';
        DocumentNo: Code[20];
        LocCode: Code[20];
        IRN: Text[250];

    //  //[Scope('Internal')]
    procedure GetIRNDetails_Govt(AuthToken: Text; Sek: Text; Irn: Text; Client_ID: Text; Client_Secret: Text; Gstin: Text; user_name: Text)
    var
        //  HttpClient: DotNet HttpClient;
        // URI: DotNet Uri;
        // ReqHdr: DotNet HttpRequestHeaders;
        // HttpStringContent: DotNet StringContent;
        txtJsonResult: Text;
        //HttpResponseMessage: DotNet HttpResponseMessage;
        //JObject: DotNet JObject;
        // Encoding: DotNet Encoding;
        txtJsonResponse: Text;
        txtResult: Text;
        txtJsonRequest: Text;
        FileName: Text;
        TestFile: File;
        FileMgt: Codeunit 419;
        MyOutStream: OutStream;
        //byteAppKey: DotNet Byte;
        ErrorDetails: Text;
        ErrorMessage: Text;
        ErrorDetailsPos1: Integer;
        ErrorDetailsPos2: Integer;
        // dtJSONConvertor: DotNet JsonConvert;
        txtData: Text;
        // ConvertCode: DotNet Convert;
        SignedQRCode: Text;
        SignedInvoice: Text;
        GSTLedgerEntry: Record "GST Ledger Entry";
        // cuEInvoiceAPI: Codeunit 50011;
        //  TempBlob: Record 99008535;
        SalesInvHeader: Record 112;
        // [RunOnClient]
        // GetAppkey: DotNet eInv;
        SaleCrMemo: Record 114;
        PurchCrMemo: Record 124;
    begin
        // HttpClient := HttpClient.HttpClient;

        // URI := URI.Uri(GetIrnDetailsText + Irn);
        // HttpClient.BaseAddress(URI);
        // HttpClient.DefaultRequestHeaders.Add('Client_ID', Client_ID);
        //  HttpClient.DefaultRequestHeaders.Add('Client_Secret', Client_Secret);
        //  HttpClient.DefaultRequestHeaders.Add('Gstin', Gstin);
        // HttpClient.DefaultRequestHeaders.Add('user_name', user_name);
        // HttpClient.DefaultRequestHeaders.Add('AuthToken', AuthToken);
        // HttpStringContent := HttpStringContent.StringContent(txtJsonRequest, Encoding.UTF8, 'application/json');
        //  HttpResponseMessage := HttpClient.GetAsync(URI).Result;
        //  txtJsonResponse := HttpResponseMessage.Content.ReadAsStringAsync().Result;

        //  JObject := JObject.JObject;
        //  JObject := JObject.Parse(txtJsonResponse);
        //  JObject := JObject.GetValue('Status');
        // IF JObject.ToString = '0' THEN BEGIN
        //     JObject := JObject.JObject;
        //     JObject := JObject.Parse(txtJsonResponse);
        //     JObject := JObject.GetValue('ErrorDetails');
        //     ErrorDetails := JObject.ToString;
        //     ErrorDetailsPos1 := STRPOS(ErrorDetails, '{');
        //     ErrorDetailsPos2 := STRPOS(ErrorDetails, '}');
        //     ErrorDetails := COPYSTR(ErrorDetails, ErrorDetailsPos1 + 1, ErrorDetailsPos2 - 1);
        //     ErrorDetailsPos2 := STRPOS(ErrorDetails, '}');
        //     ErrorDetails := '{' + COPYSTR(ErrorDetails, 1, ErrorDetailsPos2 - 1) + '}';

        //     JObject := JObject.JObject;
        //     JObject := JObject.Parse(ErrorDetails);
        //     JObject := JObject.GetValue('ErrorMessage');
        //     ErrorMessage := JObject.ToString;
        //     ERROR(ErrorMessage)
        // END ELSE BEGIN
        //     JObject := JObject.JObject;
        //     JObject := JObject.Parse(txtJsonResponse);
        //     JObject := JObject.GetValue('Data');
        //     txtResult := JObject.ToString;

        //     txtResult := GetAppkey.DecryptBySymmetricKey(txtResult, ConvertCode.FromBase64String(Sek));
        //     txtResult := GetAppkey.Base64Decode(txtResult);

        //     JObject := JObject.JObject;
        //     JObject := JObject.Parse(txtResult);
        //     JObject := JObject.GetValue('SignedQRCode');
        //     SignedQRCode := JObject.ToString;
        // END;

        IF SignedQRCode <> '' THEN BEGIN
            //    cuEInvoiceAPI.CreateQRCode(SignedQRCode, TempBlob);
            IF SalesInvHeader.GET(DocumentNo) THEN BEGIN
                //  SalesInvHeader."QR code" := TempBlob.Blob;
                ;
                SalesInvHeader.MODIFY;
            END;
            IF SaleCrMemo.GET(DocumentNo) THEN BEGIN
                //SaleCrMemo."QR code" := TempBlob.Blob;
                SaleCrMemo.MODIFY;
            END;
            IF PurchCrMemo.GET(DocumentNo) THEN BEGIN
                //PurchCrMemo."QR Code" := TempBlob.Blob;
                PurchCrMemo.MODIFY;
            END;
            GSTLedgerEntry.RESET;
            GSTLedgerEntry.SETRANGE("Document No.", DocumentNo);
            IF GSTLedgerEntry.FINDSET THEN BEGIN
                REPEAT
                    // GSTLedgerEntry."Scan QrCode E-Invoice" := cuEInvoiceAPI.GetQr(SignedQRCode);
                    //  GSTLedgerEntry."QR Code" := TempBlob.Blob;
                    GSTLedgerEntry."E-Inv Irn" := Irn;
                    GSTLedgerEntry.MODIFY;
                UNTIL GSTLedgerEntry.NEXT = 0;
            END;
        END;
    end;

    //  //[Scope('Internal')]
    procedure SetDocument(Loc: Code[20]; Docno: Code[20]; IRNVar: Text)
    begin
        LocCode := Loc;
        DocumentNo := Docno;
        IRN := IRNVar;
    end;

    // //[Scope('Internal')]
    procedure GetIRNDetails(AuthToken: Text; Sek: Text; Irn: Text; Client_ID: Text; Client_Secret: Text; Gstin: Text; user_name: Text; Password: Text)
    var
        //  HttpClient: DotNet HttpClient;
        // URI: DotNet Uri;
        // ReqHdr: DotNet HttpRequestHeaders;
        //HttpStringContent: DotNet StringContent;
        txtJsonResult: Text;
        // HttpResponseMessage: DotNet HttpResponseMessage;
        // JObject: DotNet JObject;
        // Encoding: DotNet Encoding;
        txtJsonResponse: Text;
        txtResult: Text;
        txtJsonRequest: Text;
        FileName: Text;
        TestFile: File;
        //FileMgt: Codeunit 419;
        MyOutStream: OutStream;
        // byteAppKey: DotNet Byte;
        ErrorDetails: Text;
        ErrorMessage: Text;
        ErrorDetailsPos1: Integer;
        ErrorDetailsPos2: Integer;
        //dtJSONConvertor: DotNet JsonConvert;
        txtData: Text;
        // ConvertCode: DotNet Convert;
        SignedQRCode: Text;
        SignedInvoice: Text;
        GSTLedgerEntry: Record "GST Ledger Entry";
        //cuEInvoiceAPI: Codeunit 50011;
        //TempBlob: Record 99008535;
        SalesInvHeader: Record 112;
        SaleCrMemo: Record 114;
        PurchCrmemo: Record 124;
    begin

        // HttpClient := HttpClient.HttpClient;

        //   URI := URI.Uri(GetIrnDetailsText + Irn);
        //  HttpClient.BaseAddress(URI);
        //  HttpClient.DefaultRequestHeaders.Add('user_name', user_name);
        //  HttpClient.DefaultRequestHeaders.Add('password', Password);
        // HttpClient.DefaultRequestHeaders.Add('gstin', Gstin);
        // HttpClient.DefaultRequestHeaders.Add('requestid', FORMAT(TODAY) + FORMAT(TIME));
        // HttpClient.DefaultRequestHeaders.Add('Authorization', 'Bearer' + ' ' + AuthToken);

        // HttpStringContent := HttpStringContent.StringContent(txtJsonRequest, Encoding.UTF8, 'application/json');
        // HttpResponseMessage := HttpClient.GetAsync(URI).Result;
        // txtJsonResponse := HttpResponseMessage.Content.ReadAsStringAsync().Result;

        // JObject := JObject.JObject;
        // JObject := JObject.Parse(txtJsonResponse);
        // JObject := JObject.GetValue('success');
        // IF UPPERCASE(JObject.ToString) = 'FALSE' THEN BEGIN
        //     JObject := JObject.JObject;
        //     //JObject := JObject.Parse(ErrorDetails); //DJ 23/02/2021
        //     JObject := JObject.Parse(txtJsonResponse);
        //     JObject := JObject.GetValue('message');
        //     ErrorMessage := JObject.ToString;

        //     ERROR(ErrorMessage);
        // END ELSE BEGIN
        //     JObject := JObject.JObject;
        //     JObject := JObject.Parse(txtJsonResponse);
        //     JObject := JObject.GetValue('result');
        //     txtResult := JObject.ToString;

        //     JObject := JObject.JObject;
        //     JObject := JObject.Parse(txtResult);
        //     JObject := JObject.GetValue('SignedQRCode');
        //     SignedQRCode := JObject.ToString;
        // END;

        IF SignedQRCode <> '' THEN BEGIN
            //  cuEInvoiceAPI.CreateQRCode(SignedQRCode, TempBlob);
            IF SalesInvHeader.GET(DocumentNo) THEN BEGIN
                // SalesInvHeader."QR code" := TempBlob.Blob;
                ;
                SalesInvHeader.MODIFY;
            END;
            IF SaleCrMemo.GET(DocumentNo) THEN BEGIN
                // SaleCrMemo."QR code" := TempBlob.Blob;
                SaleCrMemo.MODIFY;
            END;
            IF PurchCrmemo.GET(DocumentNo) THEN BEGIN
                // PurchCrmemo."QR Code" := TempBlob.Blob;
                PurchCrmemo.MODIFY;
            END;
            GSTLedgerEntry.RESET;
            //GSTLedgerEntry.SETRANGE("Document No.", DocumentNo);
            IF GSTLedgerEntry.FINDSET THEN BEGIN
                REPEAT
                    // GSTLedgerEntry."Scan QrCode E-Invoice" := cuEInvoiceAPI.GetQr(SignedQRCode);
                    // GSTLedgerEntry."QR Code" := TempBlob.Blob;
                    GSTLedgerEntry."E-Inv Irn" := Irn;
                    GSTLedgerEntry.MODIFY;
                UNTIL GSTLedgerEntry.NEXT = 0;
            END;
        END;
    end;
}


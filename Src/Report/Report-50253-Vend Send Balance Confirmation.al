report 50253 "Vend Send Balance Confirmation"
{
    ProcessingOnly = true;

    dataset
    {
        dataitem(Vendor; Vendor)
        {
            CalcFields = "Balance (LCY)";
            DataItemTableView = SORTING("No.")
                                ORDER(Ascending)
                                WHERE(Blocked = FILTER(<> All),
                                      "E-Mail" = FILTER(<> ''),
                                      "Exclude From Bal Confir Mail" = FILTER(false),
                                      "Vendor Posting Group" = FILTER('EMPLOYEE | BITUMEN | EXPENSE | "EXP-FO" | FUELOIL | IMPREST | PKGMAT | RAWMAT | TRANSPORT'),
                                      "Gen. Bus. Posting Group" = FILTER('DOMESTIC'));
            RequestFilterFields = "No.";

            trigger OnAfterGetRecord()
            begin
                //SMTPMailSetup.GET;
                CompanyInformation.GET;
                CLEAR(VendorBalanceConfirmation);
                CLEAR(FileName);
                CLEAR(PathName);
                CLEAR(Path);
                CLEAR(VendPostGroup); //Fahim 09-08-2022
                recVendor.RESET;
                recVendor.SETFILTER("No.", Vendor."No.");
                IF recVendor.FINDFIRST THEN BEGIN
                    //MESSAGE('Test');
                    // Fahim 09-08-2022
                    VendPostGroup := Vendor."Vendor Posting Group";
                    // End Fahim 09-08-2022
                    Clear(TempBlob);
                    TempBlob.CreateOutStream(OutStr);
                    VendorBalanceConfirmation.USEREQUESTPAGE(FALSE);
                    VendorBalanceConfirmation.SETTABLEVIEW(recVendor);
                    VendorBalanceConfirmation.CalculateAsOnDate(AsOnDate);
                    PathName := CONVERTSTR(recVendor."No.", '/', '-') + '.PDF';
                    // Path := TEMPORARYPATH + PathName;
                    FileName := Path;
                    // IF VendorBalanceConfirmation.SaveAsPdf(FileName) THEN BEGIN
                    if VendorBalanceConfirmation.SaveAs('', ReportFormat::Pdf, OutStr) then begin
                        //SMTPMail.CreateMessage('',SMTPMailSetup."User ID",Vendor."E-Mail",'Send Balance Confirmation','',TRUE);
                        // SMTPMail.CreateMessage('', SMTPMailSetup."User ID", Vendor."E-Mail", 'Vendor Balance Confirmation-' + VendPostGroup, '', TRUE); // Fahim 09-08-2022
                        Emailmsg.Create(Vendor."E-Mail", 'Vendor Balance Confirmation-' + VendPostGroup, '', TRUE);

                        EmailMsg.AppendToBody('<Br>');
                        EmailMsg.AppendToBody('Respected Sir/Madam,');
                        EmailMsg.AppendToBody('<Br>');
                        EmailMsg.AppendToBody('<Br>');
                        //EmailMsg.AppendToBody('In connection with the audit of our accounts, we request you to confirm the balance as on date as per attached report.');
                        EmailMsg.AppendToBody('In connection with the audit of our accounts, we request you to confirm the balance as on date as mentioned in attachment.'); // Fahim 09-08-2022
                        EmailMsg.AppendToBody('<Br>');
                        EmailMsg.AppendToBody('<Br>');
                        EmailMsg.AppendToBody('Thanks & Regards');
                        EmailMsg.AppendToBody('<Br>');
                        EmailMsg.AppendToBody(CompanyInformation.Name);
                        EmailMsg.AppendToBody('<Br>');
                        EmailMsg.AppendToBody('<Br>');
                        EmailMsg.AppendToBody('<HR>');
                        EmailMsg.AppendToBody('This is a system generated mail. please do not reply to this email id');
                        // SMTPMail.AddAttachment(FileName, '');
                        TempBlob.CreateInStream(InStr);
                        EmailMsg.AddAttachment(FileName, '.pdf', InStr);

                        //Send email Vendor Posting Group wise 27-07-2022 Fahim
                        IF VendPostGroup = 'EMPLOYEE' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'employee.gppl@gpglobal.com');
                        IF VendPostGroup = 'BITUMEN' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'bitumen.gppl@gpglobal.com');
                        IF VendPostGroup = 'EXPENSE' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'expense.gppl@gpglobal.com');
                        IF VendPostGroup = 'EXP-FO' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'expfo.gppl@gpglobal.com');
                        IF VendPostGroup = 'FUELOIL' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'fueloil.gppl@gpglobal.com');
                        IF VendPostGroup = 'IMPREST' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'imprest.gppl@gpglobal.com');
                        IF VendPostGroup = 'PKGMAT' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'pkgmat.gppl@gpglobal.com');
                        IF VendPostGroup = 'RAWMAT' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'rawmat.gppl@gpglobal.com');
                        IF VendPostGroup = 'TRANSPORT' THEN
                            EmailMsg.AddRecipient(RecipientType::"Cc", 'tpt.gppl@gpglobal.com');
                        //End Send email Vendor Posting Group wise 27-07-2022 Fahim

                        EmailObj.Send(Emailmsg, Enum::"Email Scenario"::Default);
                        //SMTPMail.Send;
                    END;
                END;
            end;

            trigger OnPostDataItem()
            begin
                MESSAGE('Mail Send Successfully');
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("As On Date"; AsOnDate)
                {
                    ApplicationArea = all;
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

    var
        //  SMTPMail: Codeunit 400;
        // SMTPMailSetup: Record 409;
        TempBlob: Codeunit "Temp Blob";
        OutStr: OutStream;
        Instr: InStream;
        EmailMsg: Codeunit "Email Message";
        EmailObj: Codeunit Email;
        RecipientType: Enum "Email Recipient Type";
        CompanyInformation: Record 79;
        VendorBalanceConfirmation: Report 50127;
        FileName: Text;
        recVendor: Record 23;
        Path: Text;
        PathName: Text;
        EndingDate: Date;
        Path2: Label '''E:\Avinash\Vendor Balance Report\VendorBal.pdf''';
        AsOnDate: Date;
        xSubject: Text;
        VendPostGroup: Text;
}


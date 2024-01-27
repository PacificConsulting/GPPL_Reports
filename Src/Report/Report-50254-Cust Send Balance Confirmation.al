report 50254 "Cust Send Balance Confirmation"
{
    ProcessingOnly = true;

    dataset
    {
        dataitem(Customer; Customer)
        {
            CalcFields = "Balance (LCY)";
            DataItemTableView = SORTING("No.")
                                ORDER(Ascending)
                                WHERE(Blocked = FILTER(<> All),
                                      "Exclude From Bal Confir Mail" = FILTER(false),
                                      "E-Mail" = FILTER(<> ''),
                                      "Customer Posting Group" = FILTER('AUTOMOTIVE | BITUMEN | BOILS | CEPSA | CHEMICAL | "HIGH SEAS" | FO | INDUSTRIAL | REPSOL | RPO | TOILS'),
                                      "Balance (LCY)" = FILTER(> 0));
            RequestFilterFields = "No.";

            trigger OnAfterGetRecord()
            begin
                //SMTPMailSetup.GET;
                CompanyInformation.GET;
                CLEAR(CustomerBalanceConfirmation);
                CLEAR(FileName);
                CLEAR(PathName);
                CLEAR(Path);
                recCustomer.RESET;
                recCustomer.SETFILTER("No.", Customer."No.");
                IF recCustomer.FINDFIRST THEN BEGIN
                    CLEAR(SPEmail);
                    CLEAR(HODEmail);
                    CLEAR(HOD);
                    CLEAR(Zone);
                    //Sales Person email id-Fahim 26-07-2022
                    recSalesPerson.RESET;
                    recSalesPerson.SETRANGE(recSalesPerson.Code, Customer."Salesperson Code");
                    IF recSalesPerson.FINDFIRST THEN BEGIN
                        SPEmail := recSalesPerson."E-Mail";
                        HOD := recSalesPerson."Region Head Code";
                        Zone := recSalesPerson."Region Code"
                    END;
                    //EndSales Person email id-Fahim 26-07-2022
                    //HOD email id-Fahim 26-07-2022
                    recSalesPerson2.RESET;
                    recSalesPerson2.SETRANGE(recSalesPerson2.Code, HOD);
                    IF recSalesPerson2.FINDFIRST THEN BEGIN
                        HODEmail := recSalesPerson2."E-Mail";
                    END;
                    //End HOD email id-Fahim 26-07-2022
                    Clear(TempBlob);
                    TempBlob.CreateOutStream(OutStr);

                    CustomerBalanceConfirmation.USEREQUESTPAGE(FALSE);
                    CustomerBalanceConfirmation.SETTABLEVIEW(recCustomer);
                    CustomerBalanceConfirmation.CalculateAsOnDate(AsOnDate);


                    PathName := CONVERTSTR(recCustomer."No.", '/', '-') + '.PDF';
                    //Path := TEMPORARYPATH + PathName;
                    FileName := PathName;
                    //IF CustomerBalanceConfirmation.SAVEASPDF(FileName) THEN BEGIN
                    if CustomerBalanceConfirmation.SaveAs('', ReportFormat::Pdf, OutStr) then begin
                        //SMTPMail.CreateMessage('', SMTPMailSetup."User ID", recCustomer."E-Mail", 'Customer Balance Confirmation', '', TRUE);
                        EmailMsg.Create(recCustomer."E-Mail", 'Customer Balance Confirmation', '', true);
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
                        TempBlob.CreateInStream(InStr);
                        EmailMsg.AddAttachment(FileName, '.pdf', InStr);
                        //SMTPMail.AddAttachment(FileName, '');
                        //Send email Zone wise 27-07-2022 Fahim
                        IF Zone = 'NORTH' THEN
                            IF HOD <> '' THEN BEGIN
                                //MESSAGE('%1', HOD);
                                //SMTPMail.AddCC(SPEmail + ';' + HODEmail + ';' + 'north.accounts@gpglobal.com');

                                EmailMsg.AddRecipient(RecipientType::"Cc", 'north.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", HODEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                            END ELSE BEGIN
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'north.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                //SMTPMail.AddCC(SPEmail + ';' + 'north.accounts@gpglobal.com');
                            END;
                        IF Zone = 'EAST' THEN
                            IF HOD <> '' THEN BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + HODEmail + ';' + 'east.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", HODEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'east.accounts@gpglobal.com');
                            END ELSE BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + 'east.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'east.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                            END;
                        IF Zone = 'WEST' THEN
                            IF HOD <> '' THEN BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + HODEmail + ';' + 'west.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", HODEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'west.accounts@gpglobal.com');
                            END ELSE BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + 'west.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'west.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                            END;
                        IF Zone = 'SOUTH' THEN
                            IF HOD <> '' THEN BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + HODEmail + ';' + 'south.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", HODEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'south.accounts@gpglobal.com');
                            END ELSE BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + 'south.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'south.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                            END;
                        IF Zone = 'WEST 1' THEN
                            IF HOD <> '' THEN BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + HODEmail + ';' + 'west.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", HODEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'west.accounts@gpglobal.com');
                            END ELSE BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + 'west.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'west.accounts@gpglobal.com');
                            END;
                        IF Zone = 'WEST 2' THEN
                            IF HOD <> '' THEN BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + HODEmail + ';' + 'west.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", HODEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'west.accounts@gpglobal.com');
                            END ELSE BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + 'west.accounts@gpglobal.com');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'west.accounts@gpglobal.com');
                            END;
                        IF Zone = '' THEN
                            IF HOD <> '' THEN BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + HODEmail + ';' + 'admin@pngco.in');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", HODEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'admin@pngco.in');
                            END ELSE BEGIN
                                //SMTPMail.AddCC(SPEmail + ';' + 'admin@pngco.in');
                                EmailMsg.AddRecipient(RecipientType::"Cc", SPEmail);
                                EmailMsg.AddRecipient(RecipientType::"Cc", 'admin@pngco.in');
                            END;
                        //End Send email Zone wise 27-07-2022 Fahim
                        //SMTPMail.Send;
                        EmailObj.Send(EmailMsg, Enum::"Email Scenario"::Default);
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
        //SMTPMail: Codeunit 400;
        //SMTPMailSetup: Record 409;
        CompanyInformation: Record 79;
        CustomerBalanceConfirmation: Report 50128;
        TempBlob: Codeunit "Temp Blob";
        OutStr: OutStream;
        Instr: InStream;
        EmailMsg: Codeunit "Email Message";
        EmailObj: Codeunit Email;
        RecipientType: Enum "Email Recipient Type";
        FileName: Text;
        recCustomer: Record 18;
        Path: Text;
        PathName: Text;
        EndingDate: Date;
        Path2: Label '''E:\Avinash\Vendor Balance Report\VendorBal.pdf''';
        AsOnDate: Date;
        SPEmail: Text;
        HODEmail: Text;
        HOD: Text;
        Zone: Text;
        recSalesPerson: Record 13;
        recSalesPerson2: Record 13;
        ZoneEmail: Text;
}


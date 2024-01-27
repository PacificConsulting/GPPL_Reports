report 50143 "Transport Details Cover Note"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 16Mar2018   RB-N         Freight Type
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TransportDetailsCoverNote.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Transport Details"; "Transport Details")
        {
            DataItemTableView = SORTING("TPT Invoice No.")
                                ORDER(Ascending)
                                WHERE(Open = FILTER(true),
                                      "Cancelled Invoice" = FILTER(false));
            RequestFilterFields = "Entry Date", "From Location Code";

            trigger OnAfterGetRecord()
            begin

                TCOUNT += 1;//21Apr2017

                //>>1

                recGateEntryLine.RESET;
                recGateEntryLine.SETRANGE(recGateEntryLine."Source No.", "Transport Details"."Invoice No.");
                IF recGateEntryLine.FINDFIRST THEN
                    IF recGateEntry.GET(recGateEntryLine."Entry Type", recGateEntryLine."Gate Entry No.") THEN;

                "vTPInvoiceNo." := "Transport Details"."TPT Invoice No.";
                //<<1


                //>>21Apr2017_DataGrouping
                TempTransport.SETCURRENTKEY("TPT Invoice No.");
                TempTransport.SETRANGE("TPT Invoice No.", "TPT Invoice No.");
                IF TempTransport.FINDLAST THEN BEGIN
                    ICOUNT := TempTransport.COUNT;
                END;

                GBillAmt += "Transport Details"."Bill Amount";
                GPassAmt += "Transport Details"."Passed Amount";
                GChargeAmt += "Transport Details"."Other Charges";
                GTotAmt += "Transport Details"."Total Amount";
                GTptBillAmt += "Transport Details"."Local TPT Bill Amount";
                GTptPassAmt += "Transport Details"."Local TPT Passed Amount";
                GChargeAmtL += "Transport Details"."Other Charges-Local";
                GTotAmtL += "Transport Details"."Total Amount-Local";

                //<<21Apr2017_DataGrouping


                //>>21Apr2017 DataHeaderGrouping
                IF NOT Summary AND (TCOUNT = 1) THEN
                //IF NOT Summary AND (TCOUNT = ICOUNT) THEN
                BEGIN

                    EnterCell(NNN, 1, 'TPT Invoice No.', TRUE, FALSE, '', 1);
                    EnterCell(NNN, 2, "Transport Details"."TPT Invoice No.", TRUE, FALSE, '', 1);

                    NNN += 1;
                END;

                //<<21Apr2017 DataHeaderGrouping

                //>>2
                /*
                //Transport Details, GroupHeader (1) - OnPreSection()
                
                Sr:=0;
                IF Summary=FALSE THEN BEGIN
                  Row+=1;
                  XLSHEET.Range('A'+FORMAT(Row)).Value :='TPT Invoice No.';
                  XLSHEET.Range('A'+FORMAT(Row)).Font.Bold :=TRUE;
                  XLSHEET.Range('B'+FORMAT(Row)).Value :=FORMAT("Transport Details"."TPT Invoice No.");
                  XLSHEET.Range('B'+FORMAT(Row)).Font.Bold :=TRUE;
                END;
                *///AutomationCommented
                  //<<2

                //>>3

                //Transport Details, Body (2) - OnPreSection()
                //IF NOT Summary THEN
                //IF "vTPInvoiceNo."="Transport Details"."TPT Invoice No." THEN
                //Sr+=1;

                //CreateBody;
                //<<3

                //>>4
                //Transport Details, GroupFooter (3) - OnPreSection()
                //RSPL-CAS-05774-C5J1Z0-START
                /*
                IF Summary=FALSE THEN BEGIN
                Row+=1;
                XLSHEET.Range('A'+FORMAT(Row)).Value :='Group Total';
                XLSHEET.Range('A'+FORMAT(Row)).Font.Bold :=TRUE;
                XLSHEET.Range('C'+FORMAT(Row)).Value :='';
                XLSHEET.Range('G'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Bill Amount");
                XLSHEET.Range('G'+FORMAT(Row)).Font.Bold :=TRUE;
                XLSHEET.Range('H'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Passed Amount");
                XLSHEET.Range('H'+FORMAT(Row)).Font.Bold :=TRUE;
                XLSHEET.Range('I'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Other Charges");
                XLSHEET.Range('I'+FORMAT(Row)).Font.Bold :=TRUE;
                XLSHEET.Range('J'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Total Amount");
                XLSHEET.Range('J'+FORMAT(Row)).Font.Bold :=TRUE;
                //RSPL-Sourav010415
                XLSHEET.Range('O'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Local TPT Bill Amount");
                XLSHEET.Range('O'+FORMAT(Row)).Font.Bold :=TRUE;
                XLSHEET.Range('P'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Local TPT Passed Amount");
                XLSHEET.Range('P'+FORMAT(Row)).Font.Bold :=TRUE;
                XLSHEET.Range('Q'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Other Charges-Local");
                XLSHEET.Range('Q'+FORMAT(Row)).Font.Bold :=TRUE;
                XLSHEET.Range('R'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Total Amount-Local");
                XLSHEET.Range('R'+FORMAT(Row)).Font.Bold :=TRUE;
                //RSPL-Sourav010415
                XLAPP.Visible(TRUE);
                END;
                
                Sr+=1;
                IF Summary THEN BEGIN
                Row+=1;
                XLSHEET.Range('A'+FORMAT(Row)).Value :=FORMAT(Sr);
                XLSHEET.Range('B'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Shipping Agent Name");
                XLSHEET.Range('C'+FORMAT(Row)).Value :=FORMAT("Transport Details"."TPT Invoice No.");
                XLSHEET.Range('D'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Bill Amount");
                XLSHEET.Range('E'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Passed Amount");
                XLSHEET.Range('F'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Total Amount");
                
                //RSPL-Sourav010415
                XLSHEET.Range('G'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Local Shipment Agent Name");
                XLSHEET.Range('H'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Local TPT Invoice No.");
                XLSHEET.Range('I'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Local TPT Bill Amount");
                XLSHEET.Range('J'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Local TPT Passed Amount");
                XLSHEET.Range('K'+FORMAT(Row)).Value :=FORMAT("Transport Details"."Total Amount-Local");
                //RSPL-Sourav010415
                XLAPP.Visible(TRUE);
                END;
                //RSPL-CAS-05774-C5J1Z0-END
                
                //RSPL-CAS-05774-C5J1Z0-END
                *///AutomationCommented
                  //<<4


                //>>21Apr2017 DataBody
                IF NOT Summary THEN BEGIN
                    Sr += 1;//
                    EnterCell(NNN, 1, FORMAT(Sr), FALSE, FALSE, '', 0);
                    EnterCell(NNN, 2, "Transport Details"."Shipping Agent Name", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 3, "Transport Details"."TPT Invoice No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 4, FORMAT("Transport Details"."TPT Invoice Date"), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 5, "Transport Details"."LR No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 6, "Transport Details"."Invoice No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 7, FORMAT("Transport Details"."Bill Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 8, FORMAT("Transport Details"."Passed Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 9, FORMAT("Transport Details"."Other Charges"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 10, FORMAT("Transport Details"."Total Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 11, "Transport Details"."Local Shipment Agent Name", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 12, "Transport Details"."Local TPT Invoice No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 13, FORMAT("Transport Details"."Local TPT Invoice Date"), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 14, "Transport Details"."Local LR No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 15, FORMAT("Transport Details"."Local TPT Bill Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 16, FORMAT("Transport Details"."Local TPT Passed Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 17, FORMAT("Transport Details"."Other Charges-Local"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 18, FORMAT("Transport Details".Deductions), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 19, FORMAT("Transport Details"."Freight Type"), FALSE, FALSE, '', 1);//16Mar2018
                    EnterCell(NNN, 20, "Transport Details".Reason, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 21, FORMAT("Transport Details"."Total Amount-Local"), FALSE, FALSE, '#,###0.00', 0);

                    NNN += 1;

                END;
                //<<21Apr2017 DataBody


                //>>21Apr2017 DataGroupBody
                IF TCOUNT = ICOUNT THEN BEGIN

                    IF NOT Summary THEN BEGIN

                        EnterCell(NNN, 1, 'Group Total', TRUE, FALSE, '', 1);
                        EnterCell(NNN, 7, FORMAT(GBillAmt), TRUE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 8, FORMAT(GPassAmt), TRUE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 9, FORMAT(GChargeAmt), TRUE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 10, FORMAT(GTotAmt), TRUE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 15, FORMAT(GTptBillAmt), TRUE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 16, FORMAT(GTptPassAmt), TRUE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 17, FORMAT(GChargeAmtL), TRUE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 21, FORMAT(GTotAmtL), TRUE, FALSE, '#,###0.00', 0);

                        NNN += 1;

                        Sr := 0;
                    END;

                    IF Summary THEN BEGIN
                        Sr += 1;
                        EnterCell(NNN, 1, FORMAT(Sr), FALSE, FALSE, '', 0);
                        EnterCell(NNN, 2, "Transport Details"."Shipping Agent Name", FALSE, FALSE, '', 1);
                        EnterCell(NNN, 3, "Transport Details"."TPT Invoice No.", FALSE, FALSE, '', 1);
                        EnterCell(NNN, 4, FORMAT(GBillAmt), FALSE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 5, FORMAT(GPassAmt), FALSE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 6, FORMAT(GTotAmt), FALSE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 7, "Transport Details"."Local Shipment Agent Name", FALSE, FALSE, '', 1);
                        EnterCell(NNN, 8, "Transport Details"."Local TPT Invoice No.", FALSE, FALSE, '', 1);
                        EnterCell(NNN, 9, FORMAT(GTptBillAmt), FALSE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 10, FORMAT(GTptPassAmt), FALSE, FALSE, '#,###0.00', 0);
                        EnterCell(NNN, 11, FORMAT(GTotAmtL), FALSE, FALSE, '#,###0.00', 0);

                        NNN += 1;
                    END;



                    TCOUNT := 0;
                    GBillAmt := 0;
                    GPassAmt := 0;
                    GChargeAmt := 0;
                    GTotAmt := 0;
                    GTptBillAmt := 0;
                    GTptPassAmt := 0;
                    GChargeAmtL := 0;
                    GTotAmtL := 0;
                END;

                //<<21Apr2017 DataGroupBody

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                recCompInfo.GET;
                //recUserSetup.GET(USERID);

                //CreateXLSHEET;
                //CreateHeader;
                Row := 5;
                //<<1



                //>>21Apr2017 ExcelHeader
                IF NOT Summary THEN BEGIN
                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '', 1);
                    EnterCell(1, 8, 'Date', TRUE, FALSE, '', 1);
                    EnterCell(1, 9, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

                    EnterCell(2, 1, 'Subject : Approved freight Bill', TRUE, FALSE, '', 1);
                    EnterCell(2, 8, 'USER ID', TRUE, FALSE, '', 1);
                    EnterCell(2, 9, USERID, TRUE, FALSE, '', 1);

                    EnterCell(3, 1, 'Please find here-under the following Transporters Invoices, sending thru todays courier '
                    + FORMAT("Transport Details".GETFILTER("Entry Date")), FALSE, FALSE, '', 1);

                    EnterCell(4, 1, 'as per detailed below:-', FALSE, FALSE, '', 1);

                    EnterCell(5, 1, 'No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 2, 'Transporter Name', TRUE, TRUE, '', 1);
                    EnterCell(5, 3, 'Transporter Bill', TRUE, TRUE, '', 1);
                    EnterCell(5, 4, 'Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 5, 'LR. No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 6, 'Document No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 7, 'Bill Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 8, 'Approved Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 9, 'Other Charges Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 10, 'Total Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 11, 'Local Transporter Name', TRUE, TRUE, '', 1);
                    EnterCell(5, 12, 'Local Transporter Bill', TRUE, TRUE, '', 1);
                    EnterCell(5, 13, 'Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 14, 'Local LR. No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 15, 'Local Bill Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 16, 'Local Appvd Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 17, 'Local Other Charges Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 18, 'Deductions', TRUE, TRUE, '', 1);
                    EnterCell(5, 19, 'Freight Type', TRUE, TRUE, '', 1);//16Mar2018
                    EnterCell(5, 20, 'Reason Code', TRUE, TRUE, '', 1);
                    EnterCell(5, 21, 'Local Total Amount', TRUE, TRUE, '', 1);

                    NNN := 6;
                END;


                IF Summary THEN BEGIN
                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '', 1);
                    EnterCell(1, 8, 'Date', TRUE, FALSE, '', 1);
                    EnterCell(1, 9, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

                    EnterCell(2, 1, 'Subject : Approved freight Bill', TRUE, FALSE, '', 1);
                    EnterCell(2, 8, 'USER ID', TRUE, FALSE, '', 1);
                    EnterCell(2, 9, USERID, TRUE, FALSE, '', 1);

                    EnterCell(3, 1, 'Please find here-under the following Transporters Invoices, sending thru todays courier '
                    + FORMAT("Transport Details".GETFILTER("Entry Date")), FALSE, FALSE, '', 1);

                    EnterCell(4, 1, 'as per detailed below:-', FALSE, FALSE, '', 1);

                    EnterCell(5, 1, 'No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 2, 'Transporter Name', TRUE, TRUE, '', 1);
                    EnterCell(5, 3, 'Transporter Bill', TRUE, TRUE, '', 1);
                    EnterCell(5, 4, 'Bill Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 5, 'Approved Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 6, 'Total Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 7, 'Local Transporter Name', TRUE, TRUE, '', 1);
                    EnterCell(5, 8, 'Local Transporter Bill', TRUE, TRUE, '', 1);
                    EnterCell(5, 9, 'Local Bill Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 10, 'Local Approved Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 11, 'Local Total Amount', TRUE, TRUE, '', 1);

                    NNN := 6;
                END;

                //<<21Apr2017 Excel Header

                //>>20Apr2017 CopyFilter
                TempTransport.COPYFILTERS("Transport Details");

                TCOUNT := 0;
                //<<20Apr2017 CopyFilter
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(Summary; Summary)
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

    trigger OnPostReport()
    begin

        //>>21Apr2017
        /* ExcelBuf.CreateBook('', 'DATA');
        ExcelBuf.CreateBookAndOpenExcel('', 'DATA', '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook('DATA');
        ExcelBuf.WriteSheet('DATA', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('DATA', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<21Apr2017
    end;

    var
        Row: Integer;
        Sr: Integer;
        recGateEntryLine: Record "Gate Entry Line";
        recGateEntry: Record "Posted Gate Entry Header";
        recCompInfo: Record 79;
        recUserSetup: Record 91;
        Summary: Boolean;
        "vTPInvoiceNo.": Code[30];
        "---20Apr2017": Integer;
        ExcelBuf: Record 370 temporary;
        NNN: Integer;
        TempTransport: Record 50020;
        TCOUNT: Integer;
        ICOUNT: Integer;
        GBillAmt: Decimal;
        GPassAmt: Decimal;
        GChargeAmt: Decimal;
        GTotAmt: Decimal;
        GTptBillAmt: Decimal;
        GTptPassAmt: Decimal;
        GChargeAmtL: Decimal;
        GTotAmtL: Decimal;

    // //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin


        ExcelBuf.INIT;
        ExcelBuf.VALIDATE("Row No.", Rowno);
        ExcelBuf.VALIDATE("Column No.", columnno);
        ExcelBuf."Cell Value as Text" := Cellvalue;
        ExcelBuf.Formula := '';
        ExcelBuf.Bold := Bold;
        ExcelBuf.Underline := Underline;
        ExcelBuf.NumberFormat := NoFormat;
        ExcelBuf."Cell Type" := CType;
        ExcelBuf.INSERT;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


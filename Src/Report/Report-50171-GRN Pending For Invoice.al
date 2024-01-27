report 50171 "GRN Pending For Invoice"
{
    DefaultLayout = RDLC;

    RDLCLayout = 'Src/Report Layout/GRNPendingForInvoice.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Purchase Header"; "Purchase Header")
        {
            RequestFilterFields = "Posting Date", "Vendor Posting Group", "Location Code";
            dataitem("Purchase Line"; "Purchase Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Posting Date");
                dataitem("Purch. Rcpt. Line"; "Purch. Rcpt. Line")
                {
                    DataItemLink = "Order No." = FIELD("Document No."),
                                   "Order Line No." = FIELD("Line No.");
                    DataItemTableView = WHERE(Correction = FILTER(false));

                    trigger OnAfterGetRecord()
                    begin

                        //>>09Mar2017 LocName1

                        CLEAR(LocName1);
                        recLOC.RESET;
                        recLOC.SETRANGE(Code, "Purch. Rcpt. Line"."Location Code");
                        IF recLOC.FINDFIRST THEN BEGIN
                            LocName1 := recLOC.Name;
                        END;
                        //<< LocName1

                        PurchRcptheader.RESET;
                        IF PurchRcptheader.GET("Purch. Rcpt. Line"."Document No.") THEN BEGIN
                            IF PurchRcptheader."Closed GRN" = TRUE THEN
                                CurrReport.SKIP;
                        END;

                        InvoiceDetails.RESET;
                        //InvoiceDetails.SETRANGE("Receipt Document No.", "Purch. Rcpt. Line"."Document No.");
                        //InvoiceDetails.SETRANGE("Receipt Document Line No.", "Purch. Rcpt. Line"."Line No.");
                        IF InvoiceDetails.FIND('-') THEN BEGIN
                            InvoiceNo := InvoiceDetails."Document No.";
                            Invoiceqty := InvoiceDetails.Quantity;

                            InvoiceHeader.RESET;
                            InvoiceHeader.SETRANGE("No.", InvoiceDetails."Document No.");
                            IF InvoiceHeader.FIND('-') THEN;
                            Invoicedate := InvoiceHeader."Document Date";
                        END
                        ELSE BEGIN
                            InvoiceNo := '';
                            Invoiceqty := 0;
                            Invoicedate := 0D;

                        END;

                        MakeExcelDataBody;
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>09Mar2017 LocName1

                    CLEAR(LocName);
                    recLOC.RESET;
                    recLOC.SETRANGE(Code, "Purchase Line"."Location Code");
                    IF recLOC.FINDFIRST THEN BEGIN
                        LocName := recLOC.Name;
                    END;
                    //<< LocName1
                end;
            }
        }
    }

    requestpage
    {

        layout
        {
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
        CreateExcelbook;
    end;

    trigger OnPreReport()
    begin

        DateFilter := 'GRN Pending For Invoice as on ' + "Purchase Header".GETFILTER("Purchase Header"."Posting Date");

        MakeExcelInfo;
    end;

    var
        InvoiceDetails: Record 123;
        InvoiceHeader: Record 122;
        InvoiceNo: Text[60];
        Invoiceqty: Decimal;
        Invoicedate: Date;
        ExcelBuf: Record 370 temporary;
        Printtoexcel: Boolean;
        PurchRcptheader: Record 120;
        DateFilter: Text[100];
        Text16500: Label 'As per Details';
        Text000: Label 'Data';
        Text001: Label 'GRN Pending For Invoice***';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "----09Mar2017": Integer;
        LocCode: Code[100];
        recLOC: Record 14;
        LocName: Text[250];
        LocName1: Text[100];

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin

        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('GRN Pending For Invoice*', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('50171', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        /*  ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', USERID);
         //ExcelBuf.CreateBookAndOpenExcel('',Text000,Text001,COMPANYNAME,USERID);

         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text000);
        ExcelBuf.WriteSheet(Text000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>RSPL-
        ExcelBuf.NewRow;
        MakeHederColmn(COMPANYNAME, '', TODAY);
        //<<
        //EBT091---Start
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(DateFilter, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        //EBT091---End

        //ExcelBuf.NewRow;
        MakeHederColmn('', USERID, 0D);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('PO DETAILS', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);   //EBT091
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 01Dec2020
        ExcelBuf.AddColumn('GRN DETAILS', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('INVOICE DETAILS', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Date);


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('PO No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017
        //ExcelBuf.AddColumn('Location Code',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Buy from Vendor No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Buy from Vendor Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 01Dec2020
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO Rate', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRN No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty Recd', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017
        //ExcelBuf.AddColumn('Location Code',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Voucher No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Voucher Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Inv No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Inv Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Inv Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Inv Amt', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    //   //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        IF ("Purch. Rcpt. Line"."Quantity Invoiced" + "Purch. Rcpt. Line"."Qty. Rcd. Not Invoiced" <> 0) AND (InvoiceNo = '') THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn("Purchase Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn("Purchase Line"."Order Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
            ExcelBuf.AddColumn(LocName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//09Mar2017
                                                                                                        //ExcelBuf.AddColumn("Purchase Line"."Location Code",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn("Purchase Line"."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn("Purchase Header"."Buy-from Vendor Name", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 01Dec2020
            ExcelBuf.AddColumn("Purchase Line".Type, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn("Purchase Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn("Purchase Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn("Purchase Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("Purchase Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn("Purchase Line"."Unit Cost", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("Purch. Rcpt. Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn("Purch. Rcpt. Line"."Quantity Invoiced" +
            "Purch. Rcpt. Line"."Qty. Rcd. Not Invoiced", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("Purch. Rcpt. Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(LocName1, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//09Mar2017
                                                                                                         //ExcelBuf.AddColumn("Purch. Rcpt. Line"."Location Code",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(InvoiceNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(Invoicedate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(InvoiceHeader."Vendor Invoice No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(InvoiceHeader."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
            ExcelBuf.AddColumn(Invoiceqty, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(InvoiceDetails."Amount To Vendor", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
        END;
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelGroupdata()
    begin
    end;

    local procedure MakeHederColmn(CompName: Code[250]; vUserId: Code[80]; vDate: Date)
    begin
        ExcelBuf.AddColumn(CompName, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);   //EBT091
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 01Dec2020
        IF vUserId = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(FORMAT(vDate, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END ELSE
            ExcelBuf.AddColumn(vUserId, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
    end;
}


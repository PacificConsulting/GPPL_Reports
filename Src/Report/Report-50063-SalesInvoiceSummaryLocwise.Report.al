report 50063 "Sales Invoice Summary Loc wise"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesInvoiceSummaryLocwise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("Location Code");
            RequestFilterFields = "Posting Date", "Location Code";
            column(COMPANYNAME; COMPANYNAME)
            {
            }
            column(today; FORMAT(TODAY, 0, 4))
            {
            }
            column(USERID; USERID)
            {
            }
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No."),
                               "Location Code" = FIELD("Location Code");
                DataItemTableView = SORTING("Location Code", "No.", "Sell-to Customer No.")
                                    WHERE(Type = FILTER(Item | 'Charge (Item)'),
                                          Quantity = FILTER(<> 0));
                column(LocationCode_SalesInvoiceLine; "Sales Invoice Line"."Location Code")
                {
                }
                column(decQty; decQty)
                {
                }
                column(BillAmount; BillAmount)
                {
                }
                column(NetAmt; NetAmt)
                {
                }
                column(ExcDuty; ExcDuty)
                {
                }
                column(TaxAmt; TaxAmt)
                {
                }
                column(InvDiscountAmount_SalesInvoiceLine; "Sales Invoice Line"."Inv. Discount Amount")
                {
                }
                column(TradeDisc; TradeDisc)
                {
                }
                column(TradeSubsidy; TradeSubsidy)
                {
                }
                column(Foc; Foc)
                {
                }
                column(Charges; Charges)
                {
                }
                column(EntryTax; EntryTax)
                {
                }
                column(gross; gross)
                {
                }

                trigger OnAfterGetRecord()
                begin



                    vSalesLine.RESET;
                    vSalesLine.SETRANGE("Location Code", "Sales Invoice Header"."Location Code");
                    vSalesLine.SETRANGE("Posting Date", Mindate, MaxDate);
                    vSalesLine.SETFILTER(Type, '%1|%2', "Sales Invoice Line".Type::Item, "Sales Invoice Line".Type::"Charge (Item)");
                    vSalesLine.SETFILTER(Quantity, '<>%1', 0);
                    IF vSalesLine.FINDSET THEN
                        vCount := vSalesLine.COUNT;

                    IF recitem.GET("Sales Invoice Line"."No.") THEN;

                    recItemUOM.RESET;
                    recItemUOM.SETRANGE(recItemUOM."Item No.", recitem."No.");
                    recItemUOM.SETRANGE(recItemUOM.Code, recitem."Sales Unit of Measure");
                    IF recItemUOM.FINDFIRST THEN
                        IF "Sales Invoice Line"."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                            SalesPrice.RESET;
                            SalesPrice.SETRANGE(SalesPrice."Item No.", "Sales Invoice Line"."No.");
                            IF SalesPrice.FINDLAST THEN BEGIN
                                UPricePerLT := SalesPrice."Basic Price";
                                UPrice := SalesPrice."Basic Price" * "Sales Invoice Line"."Qty. per Unit of Measure";
                            END;
                        END
                        ELSE BEGIN
                            UPricePerLT := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            UPrice := "Sales Invoice Line"."Unit Price";
                        END;

                    IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                        BillAmount := BillAmount + 0;// "Sales Invoice Line"."Amount To Customer";
                    END
                    ELSE BEGIN
                        BillAmount := BillAmount + 0/* "Sales Invoice Line"."Amount To Customer"*// "Sales Invoice Header"."Currency Factor";
                    END;


                    IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                        NetAmt := NetAmt + "Sales Invoice Line".Amount;
                    END
                    ELSE BEGIN
                        NetAmt := NetAmt + UPrice * "Sales Invoice Line".Quantity / "Sales Invoice Header"."Currency Factor";
                    END;

                    ExcDuty := ExcDuty + 0;//"Sales Invoice Line"."Excise Amount";
                    TaxAmt := TaxAmt + 0;// "Sales Invoice Line"."Tax Amount";

                    IF "Sales Invoice Line".Type <> "Sales Invoice Line".Type::"Charge (Item)" THEN
                        decQty := decQty + "Sales Invoice Line"."Quantity (Base)";

                    gross := NetAmt + BillAmount + decQty + ExcDuty + TaxAmt + TradeDisc + TradeSubsidy;

                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Group", '%1|%2', 'FOC', 'FOC DIS');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT */
                    // Foc := Foc + 0;//PostedStrOrderRec.Amount;
                    //  UNTIL PostedStrOrderRec.NEXT = 0;

                    /* PostedStrOrderRec.RESET;
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Group", 'STATE DISC');
                    IF PostedStrOrderRec.FINDSET THEN
                        REPEAT */
                    StateDisc := StateDisc + 0;// PostedStrOrderRec.Amount;
                                               // UNTIL PostedStrOrderRec.NEXT = 0;


                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Group", 'ENTRYTAX');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT */
                    EntryTax := EntryTax + 0;//PostedStrOrderRec.Amount;
                                             // UNTIL PostedStrOrderRec.NEXT = 0;
                                             /* 
                                                                 PostedStrOrderRec.RESET;
                                                                 PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                                                                 PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                                                                 PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                                                                 PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Group", '%1|%2', 'TRANSSUBS', 'TPTSUBSIDY');
                                                                 IF PostedStrOrderRec.FINDSET THEN
                                                                     REPEAT */
                    TradeSubsidy := TradeSubsidy + 0;// PostedStrOrderRec.Amount;
                                                     // UNTIL PostedStrOrderRec.NEXT = 0;

                    /*  recDetailedTaxEntry.RESET;
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Invoice Line"."Document No.");
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."No.", "Sales Invoice Line"."No.");
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Invoice Line"."Tax Area Code");
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document Line No.", "Sales Invoice Line"."Line No.");
                     recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Invoice Line"."Tax Area Code");
                     IF recDetailedTaxEntry.FINDSET THEN
                         REPEAT */
                    CessAddduty := CessAddduty + 0;// recDetailedTaxEntry."Tax Amount";
                                                   //   UNTIL recDetailedTaxEntry.NEXT = 0;

                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Group", 'DISCOUNT');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT */
                    InvDiscAmt := InvDiscAmt + 0;//PostedStrOrderRec.Amount;
                                                 // UNTIL PostedStrOrderRec.NEXT = 0;

                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Group", 'CASHDISC');
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Calculation Value", '<%1', 0);
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT */
                    CashDiscAmt := CashDiscAmt + 0;//PostedStrOrderRec.Amount;
                                                   //  UNTIL PostedStrOrderRec.NEXT = 0;

                    /* PostedStrOrderRec.RESET;
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Group", 'CASHDISC');
                    PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Calculation Value", '>%1', 0);
                    IF PostedStrOrderRec.FINDSET THEN
                        REPEAT */
                    SurChargeInterest := SurChargeInterest + 0;//PostedStrOrderRec.Amount;
                                                               //  UNTIL PostedStrOrderRec.NEXT = 0;

                    /* PostedStrOrderRec.RESET;
                    PostedStrOrderRec.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                    PostedStrOrderRec.SETRANGE("Invoice No.", "Document No.");
                    PostedStrOrderRec.SETRANGE("Item No.", "No.");
                    PostedStrOrderRec.SETRANGE("Line No.", "Line No.");
                    PostedStrOrderRec.SETRANGE("Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                    PostedStrOrderRec.SETRANGE("Tax/Charge Group", 'CCFC');
                    IF PostedStrOrderRec.FINDSET THEN
                        REPEAT */
                    SurChargeInterest += 0;// PostedStrOrderRec.Amount;
                                           //  UNTIL PostedStrOrderRec.NEXT = 0;

                    //CLEAR(FreightAmt);
                    recGLentry.RESET;
                    recGLentry.SETFILTER(recGLentry."G/L Account No.", '%1|%2', '75001010', '75001025');
                    recGLentry.SETRANGE(recGLentry."Document No.", "Sales Invoice Header"."No.");
                    IF recGLentry.FINDSET THEN
                        REPEAT
                            FreightAmt += recGLentry.Amount;
                        UNTIL recGLentry.NEXT = 0;

                    /* PostedStrOrderRec.RESET;
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                    PostedStrOrderRec.SETFILTER("Tax/Charge Group", '%1|%2|%3|%4|%5|%6|%7', 'ADD-DIS', 'AVP-SP-SAN',
                                               'MONTH-DIS', 'NEW-PRI-SP', 'SPOT-DISC', 'VAT-DIF-DI', 'T2 DISC');
                    IF PostedStrOrderRec.FINDSET THEN
                        REPEAT */
                    Charges := Charges + 0;//PostedStrOrderRec.Amount;
                                           // UNTIL PostedStrOrderRec.NEXT = 0;

                    LocRec.RESET;
                    LocRec.SETRANGE(LocRec.Code, "Sales Invoice Line"."Location Code");
                    IF LocRec.FINDFIRST THEN
                        LocName := LocRec.Name;
                    LocCode := LocRec."Global Dimension 2 Code";

                    i += 1;
                    IF vCount = i THEN
                        IF Printtoexcel THEN BEGIN
                            MakeExcelGroupFooter;
                            i := 0;
                            decQty := 0;
                            BillAmount := 0;
                            NetAmt := 0;
                            ExcDuty := 0;
                            TaxAmt := 0;
                            InvDiscAmt := 0;
                            TradeDisc := 0;
                            TradeSubsidy := 0;
                            CashDiscAmt := 0;
                            EntryTax := 0;
                            Charges := 0;
                            //  Foc := 0;
                            CessAddduty := 0;
                            FreightAmt := 0;
                            StateDisc := 0;
                            SurChargeInterest := 0;
                        END;
                end;

                trigger OnPreDataItem()
                begin
                    "Sales Invoice Line".SETRANGE("Sales Invoice Line"."Document No.", "Sales Invoice Header"."No.");
                end;
            }

            trigger OnAfterGetRecord()
            begin
                Mindate := "Sales Invoice Header".GETRANGEMIN("Posting Date");
                MaxDate := "Sales Invoice Header".GETRANGEMAX("Posting Date");
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print to Excel"; Printtoexcel)
                {
                    Caption = 'Print to Excel';
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
        //IF Printtoexcel THEN
        // Makeexcelfooter;


        IF Printtoexcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin
        //>>RSPL-Rahul Code Commented for table not found
        /*
        //EBT STIVAN ---(21082012)--- To Filter Report as per the Report View Role ------START
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (MemberOf.FINDFIRST) THEN
        ERROR('You do not have permission to run the report');
        //EBT STIVAN ---(21082012)--- To Filter Report as per the Report View Role --------END
        */
        //<<RSPl-Rahul
        CLEAR(ExcelBuf);
        IF Printtoexcel THEN
            MakeExcelInfo;

    end;

    var
        //  PostedStrOrderRec: Record "13798";
        Foc: Decimal;
        EntryTax: Decimal;
        Charges: Decimal;
        AMt: Decimal;
        decQty: Decimal;
        BillAmount: Decimal;
        NetAmt: Decimal;
        ExcDuty: Decimal;
        TaxAmt: Decimal;
        InvDiscAmt: Decimal;
        TradeDisc: Decimal;
        TradeSubsidy: Decimal;
        ExcelBuf: Record 370 temporary;
        Printtoexcel: Boolean;
        gross: Decimal;
        Datefilter: Text[50];
        LocRec: Record 14;
        LocName: Text[50];
        LocCode: Code[10];
        CashDiscAmt: Decimal;
        recitem: Record 27;
        CessAddduty: Decimal;
        SIH: Record 112;
        recSalesCrMemoHeader: Record 114;
        SIL: Record 113;
        recItem1: Record 27;
        recItemUOM: Record 5404;
        FreightAmt: Decimal;
        FreightAmount: Decimal;
        UPricePerLT: Decimal;
        UPrice: Decimal;
        SalesPrice: Record 7002;
        recGLentry: Record 17;
        StateDisc: Decimal;
        SurChargeInterest: Decimal;
        //        recDetailedTaxEntry: Record "16522";
        Text16500: Label 'As per Details';
        Text000: Label 'Data';
        Text001: Label 'Sales Invoice Summary Location Wise ';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "--": Integer;
        vCount: Integer;
        i: Integer;
        vSalesLine: Record 113;
        Mindate: Date;
        MaxDate: Date;

    //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin

        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn(FORMAT(Text0004),FALSE,'',TRUE,FALSE,FALSE,'',1);
        //ExcelBuf.AddColumn(COMPANYNAME,FALSE,'',FALSE,FALSE,FALSE,'',1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(FORMAT(Text0008), FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(TODAY, FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(FORMAT(Text0006), FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('Sales Invoice Summary Location Wise ', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        //ExcelBuf.NewRow;
        ExcelBuf.AddColumn(FORMAT(Text0007), FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);
        //ExcelBuf.NewRow;

        //ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // [Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook(Text000);
        //ExcelBuf.CreateBook('Demo','Demo2');
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.CreateBookAndOpenExcel(Text000,Text001,'',COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //// ExcelBuf.CreateBookAndOpenExcel('', Text001, '', COMPANYNAME, USERID);
        //ERROR('');

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    //  [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('SALES SUMMARY LOCATION-WISE', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Header".GETFILTER("Posting Date"), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Quantity in Ltrs/Kg.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Gross Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Net Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Excise Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Invoice Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Cash Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Freight Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Transport Subsidy', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('FOC', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Entry Tax', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Cess/Add.Tax', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('State Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Surcharge Interest', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.NewRow;
    end;

    // [Scope('Internal')]
    procedure MakeExcelGroupFooter()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(LocRec.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(LocCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(decQty, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(BillAmount, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(NetAmt, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ExcDuty, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TaxAmt + CessAddduty, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(InvDiscAmt, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(CashDiscAmt, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ABS(FreightAmt), FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TradeSubsidy, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(Foc, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(Charges, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(EntryTax, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(CessAddduty, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(StateDisc, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(SurChargeInterest, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
    end;

    //  [Scope('Internal')]
    procedure Makeexcelfooter()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(decQty, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(BillAmount, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(NetAmt, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(ExcDuty, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(TaxAmt + CessAddduty, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(InvDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(CashDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(ABS(FreightAmt), FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(TradeSubsidy, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(Foc, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(Charges, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(EntryTax, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(CessAddduty, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(StateDisc, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(SurChargeInterest, FALSE, '', TRUE, FALSE, TRUE, '', 1);
    end;
}


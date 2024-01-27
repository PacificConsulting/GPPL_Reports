report 50081 "Customer - Ledger"
{
    // //03May2017 R_104--Customer - Detail Trial Bal. ---> R_50081--Customer Ledger
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CustomerLedger.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Customer - Ledger';

    dataset
    {
        dataitem(Customer; Customer)
        {
            DataItemTableView = SORTING("No.")
                                WHERE(Blocked = FILTER(<> All));
            PrintOnlyIfDetail = true;
            RequestFilterFields = "No.", "Date Filter", "Responsibility Center", "Salesperson Code";
            column(TodayFormatted; FORMAT(TODAY, 0, 4))
            {
            }
            column(PeriodCustDatetFilter; STRSUBSTNO(Text000, CustDateFilter))
            {
            }
            column(CompanyName; COMPANYNAME)
            {
            }
            column(PrintAmountsInLCY; PrintAmountsInLCY)
            {
            }
            column(ExcludeBalanceOnly; ExcludeBalanceOnly)
            {
            }
            column(CustFilterCaption; TABLECAPTION + ': ' + CustFilter)
            {
            }
            column(CustFilter; CustFilter)
            {
            }
            column(AmountCaption; AmountCaption)
            {
            }
            column(RemainingAmtCaption; RemainingAmtCaption)
            {
            }
            column(No_Cust; "No.")
            {
            }
            column(Name_Cust; Name)
            {
            }
            column(PhoneNo_Cust; "Phone No.")
            {
                IncludeCaption = true;
            }
            column(PageGroupNo; PageGroupNo)
            {
            }
            column(StartBalanceLCY; StartBalanceLCY)
            {
                AutoFormatType = 1;
            }
            column(StartBalAdjLCY; StartBalAdjLCY)
            {
                AutoFormatType = 1;
            }
            column(CustBalanceLCY; CustBalanceLCY)
            {
                AutoFormatType = 1;
            }
            column(CustLedgerEntryAmtLCY; "Cust. Ledger Entry"."Amount (LCY)" + Correction + ApplicationRounding)
            {
                AutoFormatType = 1;
            }
            column(StartBalanceLCYAdjLCY; StartBalanceLCY + StartBalAdjLCY)
            {
                AutoFormatType = 1;
            }
            column(StrtBalLCYCustLedgEntryAmt; StartBalanceLCY + "Cust. Ledger Entry"."Amount (LCY)" + Correction + ApplicationRounding)
            {
                AutoFormatType = 1;
            }
            column(CustDetailTrialBalCaption; CustDetailTrialBalCaptionLbl)
            {
            }
            column(PageNoCaption; PageNoCaptionLbl)
            {
            }
            column(AllAmtsLCYCaption; AllAmtsLCYCaptionLbl)
            {
            }
            column(RepInclCustsBalCptn; RepInclCustsBalCptnLbl)
            {
            }
            column(PostingDateCaption; PostingDateCaptionLbl)
            {
            }
            column(DueDateCaption; DueDateCaptionLbl)
            {
            }
            column(BalanceLCYCaption; BalanceLCYCaptionLbl)
            {
            }
            column(AdjOpeningBalCaption; AdjOpeningBalCaptionLbl)
            {
            }
            column(BeforePeriodCaption; BeforePeriodCaptionLbl)
            {
            }
            column(TotalCaption; TotalCaptionLbl)
            {
            }
            column(OpeningBalCaption; OpeningBalCaptionLbl)
            {
            }
            column(FirstDate; FirstDate)
            {
            }
            column(OpenBalance; OpenBalance)
            {
            }
            column(OpeningBalance1; OpeningBalance1)
            {
            }
            dataitem("Cust. Ledger Entry"; "Cust. Ledger Entry")
            {
                DataItemLink = "Customer No." = FIELD("No."),
                               "Posting Date" = FIELD("Date Filter"),
                               "Global Dimension 2 Code" = FIELD("Global Dimension 2 Filter"),
                               "Global Dimension 1 Code" = FIELD("Global Dimension 1 Filter"),
                               "Date Filter" = FIELD("Date Filter");
                DataItemTableView = SORTING("Customer No.", "Posting Date");
                column(PostDate_CustLedgEntry; FORMAT("Posting Date"))
                {
                }
                column(DocType_CustLedgEntry; "Document Type")
                {
                    IncludeCaption = true;
                }
                column(DocNo_CustLedgEntry; "Document No.")
                {
                    IncludeCaption = true;
                }
                column(Desc_CustLedgEntry; Description)
                {
                    IncludeCaption = true;
                }
                column(CustAmount; CustAmount)
                {
                    AutoFormatExpression = CustCurrencyCode;
                    AutoFormatType = 1;
                }
                column(CustRemainAmount; CustRemainAmount)
                {
                    AutoFormatExpression = CustCurrencyCode;
                    AutoFormatType = 1;
                }
                column(CustEntryDueDate; FORMAT(CustEntryDueDate))
                {
                }
                column(EntryNo_CustLedgEntry; "Entry No.")
                {
                    IncludeCaption = true;
                }
                column(CustCurrencyCode; CustCurrencyCode)
                {
                }
                column(CustBalanceLCY1; CustBalanceLCY)
                {
                    AutoFormatType = 1;
                }
                column(Cheque; Cheque)
                {
                }
                column(ChequeDatails; ChequeDatails)
                {
                }
                column(Narration; Narration)
                {
                }
                column(SourceCode; "Source Code")
                {
                }
                column(DebitAmtLCY; "Debit Amount (LCY)")
                {
                }
                column(CreditAmtLCY; "Credit Amount (LCY)")
                {
                }
                dataitem("Detailed Cust. Ledg. Entry"; "Detailed Cust. Ledg. Entry")
                {
                    DataItemLink = "Cust. Ledger Entry No." = FIELD("Entry No.");
                    DataItemTableView = SORTING("Cust. Ledger Entry No.", "Entry Type", "Posting Date")
                                        WHERE("Entry Type" = CONST("Correction of Remaining Amount"));
                    column(EntryType_DtldCustLedgEntry; FORMAT("Entry Type"))
                    {
                    }
                    column(Correction; Correction)
                    {
                        AutoFormatType = 1;
                    }
                    column(CustBalanceLCY2; CustBalanceLCY)
                    {
                        AutoFormatType = 1;
                    }
                    column(ApplicationRounding; ApplicationRounding)
                    {
                        AutoFormatType = 1;
                    }

                    trigger OnAfterGetRecord()
                    begin
                        //>>NAV2009
                        /*
                        CASE "Entry Type" OF
                          "Entry Type"::"Appln. Rounding":
                            ApplicationRounding := ApplicationRounding + "Amount (LCY)";
                          "Entry Type"::"Correction of Remaining Amount":
                            Correction := Correction + "Amount (LCY)";
                        END;
                        */
                        //<<NAV2009 Commented

                        CustBalanceLCY := CustBalanceLCY + "Amount (LCY)";

                    end;

                    trigger OnPreDataItem()
                    begin
                        SETFILTER("Posting Date", CustDateFilter);
                        Correction := 0;
                        ApplicationRounding := 0;
                    end;
                }
                dataitem("Detailed Cust. Ledg. Entry2"; "Detailed Cust. Ledg. Entry")
                {
                    DataItemLink = "Cust. Ledger Entry No." = FIELD("Entry No.");
                    DataItemTableView = SORTING("Cust. Ledger Entry No.", "Entry Type", "Posting Date")
                                        WHERE("Entry Type" = CONST("Appln. Rounding"));

                    trigger OnAfterGetRecord()
                    begin

                        //>>1 NAV2009

                        ApplicationRounding := ApplicationRounding + "Amount (LCY)";
                        CustBalanceLCY := CustBalanceLCY + "Amount (LCY)";
                        IF CONFIRM('%1 %2', TRUE, Text001, ApplicationRounding) THEN;
                        //<<1 NAV2009
                    end;

                    trigger OnPreDataItem()
                    begin

                        //>>1 NAV2009
                        SETFILTER("Posting Date", CustDateFilter);
                        //<<1 NAV2009
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>NAV2009

                    CLEAR(Narration);
                    CLEAR(ChequeDatails);
                    CLEAR(DocnoCancInv);
                    //<<NAV2009

                    CALCFIELDS(Amount, "Remaining Amount", "Amount (LCY)", "Remaining Amt. (LCY)");

                    CustLedgEntryExists := TRUE;
                    IF PrintAmountsInLCY THEN BEGIN
                        CustAmount := "Amount (LCY)";
                        CustRemainAmount := "Remaining Amt. (LCY)";
                        CustCurrencyCode := '';
                    END ELSE BEGIN
                        CustAmount := Amount;
                        CustRemainAmount := "Remaining Amount";
                        CustCurrencyCode := "Currency Code";
                    END;
                    CustBalanceLCY := CustBalanceLCY + "Amount (LCY)";


                    IF ("Document Type" = "Document Type"::Payment) OR ("Document Type" = "Document Type"::Refund) THEN
                        CustEntryDueDate := 0D
                    ELSE
                        CustEntryDueDate := "Due Date";

                    //>>NAV2009

                    DebitAmt := "Debit Amount (LCY)";
                    CreditAmt := "Credit Amount (LCY)";

                    TCreditAmt += CreditAmt;//04May2017
                    TDebitAmt += DebitAmt;//04May2017

                    PostdNarration.RESET;
                    PostdNarration.SETRANGE(PostdNarration."Document No.", "Cust. Ledger Entry"."Document No.");
                    PostdNarration.SETRANGE("Transaction No.", "Cust. Ledger Entry"."Transaction No.");
                    IF PostdNarration.FINDSET THEN
                        REPEAT
                            Narration := Narration + ' ' + PostdNarration.Narration;
                        UNTIL PostdNarration.NEXT = 0;

                    recSalesCommentLine.RESET;
                    recSalesCommentLine.SETRANGE(recSalesCommentLine."Document Type", recSalesCommentLine."Document Type"::"Posted Credit Memo");
                    recSalesCommentLine.SETRANGE(recSalesCommentLine."No.", "Cust. Ledger Entry"."Document No.");
                    IF recSalesCommentLine.FINDFIRST THEN
                        REPEAT
                            Narration := Narration + ' ' + recSalesCommentLine.Comment;
                        UNTIL recSalesCommentLine.NEXT = 0;

                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.", "Cust. Ledger Entry"."Document No.");
                    recSIH.SETFILTER(recSIH."Order No.", '%1', '');
                    IF recSIH.FINDFIRST THEN BEGIN
                        recSalesCommentLine.RESET;
                        recSalesCommentLine.SETRANGE(recSalesCommentLine."Document Type", recSalesCommentLine."Document Type"::"Posted Invoice");
                        recSalesCommentLine.SETRANGE(recSalesCommentLine."No.", "Cust. Ledger Entry"."Document No.");
                        IF recSalesCommentLine.FINDFIRST THEN
                            REPEAT
                                Narration := Narration + ' ' + recSalesCommentLine.Comment;
                            UNTIL recSalesCommentLine.NEXT = 0;
                    END;

                    BankAccLedEntry.RESET;
                    BankAccLedEntry.SETRANGE(BankAccLedEntry."Document No.", "Cust. Ledger Entry"."Document No.");
                    IF BankAccLedEntry.FINDFIRST THEN BEGIN
                        Cheque := 'Cheque No:-';
                        ChequeDatails := BankAccLedEntry."Cheque No.";
                    END ELSE BEGIN
                        Cheque := '';
                        ChequeDatails := '';
                    END;
                    //<<NAV2009

                    //>>NAV2009
                    //Cust. Ledger Entry, Body (2) - OnPreSection()
                    IF PostdNarration.GET("Entry No.") THEN;

                    salescreditmemo.RESET;
                    salescreditmemo.SETRANGE(salescreditmemo."No.", "Document No.");


                    IF salescreditmemo.FIND('-') THEN
                        IF salescreditmemo."Applies-to Doc. No." <> '' THEN
                            //DocnoCancInv:=COPYSTR(salescreditmemo."Applies-to Doc. No.",STRLEN(salescreditmemo."Applies-to Doc. No.")-3,4)
                            DocnoCancInv := COPYSTR(salescreditmemo."Applies-to Doc. No.", STRLEN(salescreditmemo."Applies-to Doc. No.") - 2, 2)
                        ELSE
                            DocnoCancInv := '';



                    IF PrintToExcel THEN
                        MakeExcelDataBody;
                    //<<NAV2009
                end;

                trigger OnPreDataItem()
                begin
                    CustLedgEntryExists := FALSE;

                    //CurrReport.CREATETOTALS(CustAmount,"Amount (LCY)",DebitAmt,CreditAmt);//NAV2009 //111
                    CurrReport.CREATETOTALS(CustAmount, "Amount (LCY)");
                end;
            }
            dataitem(Integer; Integer)
            {
                DataItemTableView = SORTING(Number)
                                    WHERE(Number = CONST(1));
                column(Name1_Cust; Customer.Name)
                {
                }
                column(CustBalanceLCY4; CustBalanceLCY)
                {
                    AutoFormatType = 1;
                }
                column(StartBalanceLCY2; StartBalanceLCY)
                {
                }
                column(StartBalAdjLCY2; StartBalAdjLCY)
                {
                }
                column(CustBalStBalStBalAdjLCY; CustBalanceLCY - StartBalanceLCY - StartBalAdjLCY)
                {
                    AutoFormatType = 1;
                }

                trigger OnAfterGetRecord()
                begin
                    IF NOT CustLedgEntryExists AND ((StartBalanceLCY = 0) OR ExcludeBalanceOnly) THEN BEGIN
                        StartBalanceLCY := 0;
                        CurrReport.SKIP;
                    END;


                    //>>NAV2009
                    //Integer, Body (1) - OnPreSection()
                    IF CustBalanceLCY > 0 THEN BEGIN
                        ClosingBalance2 := 0;
                        ClosingBalance1 := CustBalanceLCY;
                    END ELSE
                        IF CustBalanceLCY < 0 THEN BEGIN
                            ClosingBalance1 := 0;
                            ClosingBalance2 := CustBalanceLCY;
                        END;


                    //CurrReport.SHOWOUTPUT(NOT PrintAmountsInLCY);

                    IF PrintToExcel THEN
                        MakeExcelDataFooter;

                    //<<NAV2009
                end;
            }

            trigger OnAfterGetRecord()
            begin
                IF PrintOnlyOnePerPage THEN
                    PageGroupNo := PageGroupNo + 1;

                OpenBalance := 0; //NAV2009
                OpeningBalance1 := 0;//NAV2009
                StartBalanceLCY := 0;
                StartBalAdjLCY := 0;
                IF CustDateFilter <> '' THEN BEGIN
                    IF GETRANGEMIN("Date Filter") <> 0D THEN BEGIN
                        SETRANGE("Date Filter", 0D, GETRANGEMIN("Date Filter") - 1);
                        CALCFIELDS("Net Change (LCY)");
                        StartBalanceLCY := "Net Change (LCY)";
                    END;
                    SETFILTER("Date Filter", CustDateFilter);
                    CALCFIELDS("Net Change (LCY)");
                    StartBalAdjLCY := "Net Change (LCY)";
                    CustLedgEntry.SETCURRENTKEY("Customer No.", "Posting Date");
                    CustLedgEntry.SETRANGE("Customer No.", "No.");
                    CustLedgEntry.SETFILTER("Posting Date", CustDateFilter);
                    IF CustLedgEntry.FIND('-') THEN
                        REPEAT
                            CustLedgEntry.SETFILTER("Date Filter", CustDateFilter);
                            CustLedgEntry.CALCFIELDS("Amount (LCY)");
                            StartBalAdjLCY := StartBalAdjLCY - CustLedgEntry."Amount (LCY)";
                            "Detailed Cust. Ledg. Entry".SETCURRENTKEY("Cust. Ledger Entry No.", "Entry Type", "Posting Date");
                            "Detailed Cust. Ledg. Entry".SETRANGE("Cust. Ledger Entry No.", CustLedgEntry."Entry No.");
                            "Detailed Cust. Ledg. Entry".SETFILTER("Entry Type", '%1|%2',
                              "Detailed Cust. Ledg. Entry"."Entry Type"::"Correction of Remaining Amount",
                              "Detailed Cust. Ledg. Entry"."Entry Type"::"Appln. Rounding");
                            "Detailed Cust. Ledg. Entry".SETFILTER("Posting Date", CustDateFilter);
                            IF "Detailed Cust. Ledg. Entry".FIND('-') THEN
                                REPEAT
                                    StartBalAdjLCY := StartBalAdjLCY - "Detailed Cust. Ledg. Entry"."Amount (LCY)";
                                UNTIL "Detailed Cust. Ledg. Entry".NEXT = 0;
                            "Detailed Cust. Ledg. Entry".RESET;
                        UNTIL CustLedgEntry.NEXT = 0;
                END;
                //   CurrReport.PRINTONLYIFDETAIL := ExcludeBalanceOnly OR (StartBalanceLCY = 0);
                CustBalanceLCY := StartBalanceLCY + StartBalAdjLCY;

                //>>NAV2009

                IF CustDateFilter = '' THEN
                    ERROR('Please give the Date Filter');

                FirstDate := COPYSTR(FORMAT(CustDateFilter), 1, 8);
                //<<NAV2009


                //>>NAV2009
                //Customer, Body (6) - OnPreSection()
                // Changed Customer.full name to customer.name   EBT MILAN 130114
                DetailsCust := 'Customer No/Name  : ' + Customer."No." + ' ; ' + Customer.Name;
                IF StartBalanceLCY > 0 THEN BEGIN
                    OpeningBalance1 := 0;
                    OpenBalance := StartBalanceLCY;
                END ELSE
                    IF StartBalanceLCY < 0 THEN BEGIN
                        OpenBalance := 0;
                        OpeningBalance1 := StartBalanceLCY;
                    END;

                FirstDate := COPYSTR(FORMAT(CustDateFilter), 1, 8);

                IF PrintToExcel THEN
                    MakeCustHeader;
                //<<NAV2009
            end;

            trigger OnPreDataItem()
            begin
                PageGroupNo := 1;
                CurrReport.NEWPAGEPERRECORD := PrintOnlyOnePerPage;
                CurrReport.CREATETOTALS("Cust. Ledger Entry"."Amount (LCY)", StartBalanceLCY, StartBalAdjLCY, Correction, ApplicationRounding);

                //>>NAV2009
                /*
                Memberof.RESET;
                Memberof.SETRANGE(Memberof."User ID",USERID);
                Memberof.SETRANGE(Memberof."Role ID",'MARKETING');
                IF Memberof.FINDFIRST THEN
                BEGIN
                  Customer.SETFILTER(Customer."Responsibility Center",CustResCenterFilter,CustResCenter);
                  Customer.SETRANGE(Customer."Salesperson Code",USERID);
                END;
                *///Commented NAV2009
                  //<<NAV2009

                //>>NAV2009 ExcelHeader
                //Customer, Header (1) - OnPreSection()
                DateDetails := STRSUBSTNO(Text000, CustDateFilter);

                IF PrintToExcel AND Flag THEN BEGIN
                    MakeExcelHeader;
                    Flag := FALSE;
                END;

                //<<NAV2009 ExcelHeader

            end;
        }
    }

    requestpage
    {
        SaveValues = true;

        layout
        {
            area(content)
            {
                group(Options)
                {
                    Caption = 'Options';
                    field(ShowAmountsInLCY; PrintAmountsInLCY)
                    {
                        Caption = 'Show Amounts in LCY';
                    }
                    field(NewPageperCustomer; PrintOnlyOnePerPage)
                    {
                        Caption = 'New Page per Customer';
                    }
                    field(ExcludeCustHaveaBalanceOnly; ExcludeBalanceOnly)
                    {
                        Caption = 'Exclude Customers That Have a Balance Only';
                        MultiLine = true;
                    }
                    field("Print to Excel"; PrintToExcel)
                    {
                    }
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

        //>>1

        IF PrintToExcel THEN
            CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin
        //CustFilter := Customer.GETFILTERS;
        //CustDateFilter := Customer.GETFILTER("Date Filter");

        //>>NAV2009
        /*
        CLEAR(CustResCenter);
        recCust.RESET;
        recCust.SETRANGE(recCust."No.",Customer.GETFILTER(Customer."No."));
        IF recCust.FINDFIRST THEN
        BEGIN
         CustResCenter := recCust."Responsibility Center";
        
         Memberof.RESET;
         Memberof.SETRANGE(Memberof."User ID",USERID);
         Memberof.SETRANGE(Memberof."Role ID",'MARKETING');
         IF Memberof.FINDFIRST THEN
         BEGIN
          CSOmapping.RESET;
          CSOmapping.SETRANGE(CSOmapping."User Id",USERID);
          CSOmapping.SETRANGE(CSOmapping.Type,CSOmapping.Type::"Responsibility Center");
          CSOmapping.SETFILTER(CSOmapping.Value,CustResCenter);
          IF NOT CSOmapping.FINDFIRST THEN
            ERROR('You are not allowed to run the report other than your Responsibility Center')
         END;
        
         user.GET(USERID);
         IF (user."User ID" = 'CF24') OR (user."User ID" = 'CF09') OR (user."User ID" = 'CF11') OR (user."User ID" = 'CF12') OR
         (user."User ID" = 'CF16') OR (user."User ID" = 'CF22') OR (user."User ID" = 'CF29') OR (user."User ID" = 'CF31') OR
         (user."User ID" = 'CF32') OR (user."User ID" = 'CF34') OR (user."User ID" = 'CF35') OR (user."User ID" = 'CF36') OR
         (user."User ID" = 'CF40') OR (user."User ID" = 'CF41') OR (user."User ID" = 'CF42') OR (user."User ID" = 'CF43') OR
         (user."User ID" = 'CF44') OR (user."User ID" = 'CF46') OR (user."User ID" = 'CF47') OR (user."User ID" = 'CF48') OR
         (user."User ID" = 'CF49') OR (user."User ID" = 'CF50') OR (user."User ID" = 'CF51') OR (user."User ID" = 'CF52') OR
         (user."User ID" = 'CF53') OR (user."User ID" = 'CF55') THEN
         ERROR('You are not allowed to run this Report');
        END;
        
        */ //Commented 03May2017


        Flag := TRUE;
        CustFilter := Customer.GETFILTERS;
        CustDateFilter := Customer.GETFILTER("Date Filter");

        //<<NAV2009

        WITH "Cust. Ledger Entry" DO
            IF PrintAmountsInLCY THEN BEGIN
                AmountCaption := FIELDCAPTION("Amount (LCY)");
                RemainingAmtCaption := FIELDCAPTION("Remaining Amt. (LCY)");
            END ELSE BEGIN
                AmountCaption := FIELDCAPTION(Amount);
                RemainingAmtCaption := FIELDCAPTION("Remaining Amount");
            END;


        IF PrintToExcel THEN
            MakeExcelInfo;

        //<<

    end;

    var
        Text000: Label 'Period: %1';
        CustLedgEntry: Record 21;
        PrintAmountsInLCY: Boolean;
        PrintOnlyOnePerPage: Boolean;
        ExcludeBalanceOnly: Boolean;
        CustFilter: Text;
        CustDateFilter: Text[30];
        AmountCaption: Text[80];
        RemainingAmtCaption: Text[30];
        CustAmount: Decimal;
        CustRemainAmount: Decimal;
        CustBalanceLCY: Decimal;
        CustCurrencyCode: Code[10];
        CustEntryDueDate: Date;
        StartBalanceLCY: Decimal;
        StartBalAdjLCY: Decimal;
        Correction: Decimal;
        ApplicationRounding: Decimal;
        CustLedgEntryExists: Boolean;
        PageGroupNo: Integer;
        CustDetailTrialBalCaptionLbl: Label 'Customer - Detail Trial Bal.';
        PageNoCaptionLbl: Label 'Page';
        AllAmtsLCYCaptionLbl: Label 'All amounts are in LCY';
        RepInclCustsBalCptnLbl: Label 'This report also includes customers that only have balances.';
        PostingDateCaptionLbl: Label 'Posting Date';
        DueDateCaptionLbl: Label 'Due Date';
        BalanceLCYCaptionLbl: Label 'Balance (LCY)';
        AdjOpeningBalCaptionLbl: Label 'Adj. of Opening Balance';
        BeforePeriodCaptionLbl: Label 'Total (LCY) Before Period';
        TotalCaptionLbl: Label 'Total (LCY)';
        OpeningBalCaptionLbl: Label 'Total Adj. of Opening Balance';
        "---03May2017-----NAV2009": Integer;
        Flag: Boolean;
        Text001: Label 'Appln Rounding:';
        Text0000: Label 'Data';
        Text0001: Label 'Customer - Ledger';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        FirstDate: Text[30];
        DebitAmt: Decimal;
        CreditAmt: Decimal;
        Narration: Text[1024];
        ChequeDatails: Text[30];
        DocnoCancInv: Text[60];
        PostdNarration: Record "Posted Narration";
        recSalesCommentLine: Record 44;
        recSIH: Record 112;
        BankAccLedEntry: Record 271;
        Cheque: Text[30];
        salescreditmemo: Record 114;
        OpenBalance: Decimal;
        OpeningBalance1: Decimal;
        DetailsCust: Text[100];
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        DateDetails: Text[50];
        Narration1: Text[250];
        ClosingBalance1: Decimal;
        ClosingBalance2: Decimal;
        TCreditAmt: Decimal;
        TDebitAmt: Decimal;

    // [Scope('Internal')]
    procedure InitializeRequest(ShowAmountInLCY: Boolean; SetPrintOnlyOnePerPage: Boolean; SetExcludeBalanceOnly: Boolean)
    begin
        PrintOnlyOnePerPage := SetPrintOnlyOnePerPage;
        PrintAmountsInLCY := ShowAmountInLCY;
        ExcludeBalanceOnly := SetExcludeBalanceOnly;
    end;

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//04May2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Customer - Ledger', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Customer - Ledger", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
    end;

    //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text0000,Text0001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>04May2017
        /*  ExcelBuf.CreateBook('', Text0001);
         ExcelBuf.CreateBookAndOpenExcel('', Text0001, '', '', USERID);
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<04May2017
    end;

    // [Scope('Internal')]
    procedure MakeExcelHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2 //04May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3 //04May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4 //04May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5 //04May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6 //04May2017
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7 //04May2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Customer Ledger', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2 //04May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3 //04May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4 //04May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5 //04May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6 //04May2017
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7 //04May2017
        ExcelBuf.NewRow; //04May2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(DateDetails, FALSE, '', FALSE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Posting Date', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Type', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Document No', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Narration', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Debit(Rs.)', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Credit(Rs.)', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Balance(Rs.)', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//7
    end;

    // [Scope('Internal')]
    procedure MakeCustHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(DetailsCust, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Opening Balance As On:', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn(FirstDate, FALSE, '', FALSE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//4
        IF OpenBalance <> 0 THEN
            ExcelBuf.AddColumn(ABS(OpenBalance), FALSE, '', FALSE, FALSE, TRUE, '#,###0.00', 0)//5
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//5

        IF OpeningBalance1 <> 0 THEN
            ExcelBuf.AddColumn(ABS(OpeningBalance1), FALSE, '', FALSE, FALSE, TRUE, '#,###0.00', 0)//6
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//6
    end;

    //  [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Cust. Ledger Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn("Cust. Ledger Entry"."Source Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Cust. Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        Narration1 := COPYSTR(Narration, 1, 250);
        ExcelBuf.AddColumn(Narration1, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        Narration1 := ''; //04May2017
        IF "Cust. Ledger Entry"."Debit Amount (LCY)" <> 0 THEN
            ExcelBuf.AddColumn("Cust. Ledger Entry"."Debit Amount (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0)//5
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5

        IF "Cust. Ledger Entry"."Credit Amount (LCY)" <> 0 THEN
            ExcelBuf.AddColumn("Cust. Ledger Entry"."Credit Amount (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0)//6
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6

        ExcelBuf.AddColumn(CustBalanceLCY, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//7
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn(Cheque, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(ChequeDatails, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4

        ExcelBuf.AddColumn((TDebitAmt + ABS(OpenBalance)), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//5 //04May2017
        TDebitAmt := 0;//04May2017
        //ExcelBuf.AddColumn((DebitAmt+ABS(OpenBalance)),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//5

        ExcelBuf.AddColumn((TCreditAmt + ABS(OpeningBalance1)), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//6 //04May2017
        TCreditAmt := 0;//04May2017

        //ExcelBuf.AddColumn((CreditAmt+ABS(OpeningBalance1)),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//6
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Closing Balance', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        IF ClosingBalance1 <> 0 THEN
            ExcelBuf.AddColumn(ABS(ClosingBalance1), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0)//5
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5

        IF ClosingBalance2 <> 0 THEN
            ExcelBuf.AddColumn(ABS(ClosingBalance2), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0)//6
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
    end;
}


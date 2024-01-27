report 50103 "Customer  Aging Detailed"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 15Apr2017   RB-N         [duedate]-->[CLEEndDueDate]
    // 17Apr2017   RB-N         [OriginalAmtCptn]-->Resp. Center,SalespersonCode RESPONSIBLITYCENTRE__GROUPING
    // 18Nov2017   RB-N         DueDate Calculation for AgingBy::"Due Date"
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CustomerAgingDetailed.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Aged Accounts Receivable';

    dataset
    {
        dataitem(Customer; Customer)
        {
            RequestFilterFields = "No.", "Responsibility Center", "Salesperson Code", Type;
            column(TodayFormatted; FORMAT(TODAY, 0, 4))
            {
            }
            column(CompanyName; COMPANYNAME)
            {
            }
            column(FormatEndingDate; STRSUBSTNO(Text006, FORMAT(EndingDate, 0, 4)))
            {
            }
            column(PostingDate; STRSUBSTNO(Text007, SELECTSTR(AgingBy + 1, Text009)))
            {
            }
            column(PrintAmountInLCY; PrintAmountInLCY)
            {
            }
            column(TableCaptnCustFilter; TABLECAPTION + ': ' + CustFilter)
            {
            }
            column(CustFilter; CustFilter)
            {
            }
            column(AgingByDueDate; AgingBy = AgingBy::"Due Date")
            {
            }
            column(AgedbyDocumnetDate; STRSUBSTNO(Text004, SELECTSTR(AgingBy + 1, Text009)))
            {
            }
            column(HeaderText5; HeaderText[5])
            {
            }
            column(HeaderText4; HeaderText[4])
            {
            }
            column(HeaderText3; HeaderText[3])
            {
            }
            column(HeaderText2; HeaderText[2])
            {
            }
            column(HeaderText1; HeaderText[1])
            {
            }
            column(PrintDetails; PrintDetails)
            {
            }
            column(GrandTotalCLE5RemAmt; GrandTotalCustLedgEntry[5]."Remaining Amt. (LCY)")
            {
                AutoFormatType = 1;
            }
            column(GrandTotalCLE4RemAmt; GrandTotalCustLedgEntry[4]."Remaining Amt. (LCY)")
            {
                AutoFormatType = 1;
            }
            column(GrandTotalCLE3RemAmt; GrandTotalCustLedgEntry[3]."Remaining Amt. (LCY)")
            {
                AutoFormatType = 1;
            }
            column(GrandTotalCLE2RemAmt; GrandTotalCustLedgEntry[2]."Remaining Amt. (LCY)")
            {
                AutoFormatType = 1;
            }
            column(GrandTotalCLE1RemAmt; GrandTotalCustLedgEntry[1]."Remaining Amt. (LCY)")
            {
                AutoFormatType = 1;
            }
            column(GrandTotalCLEAmtLCY; GrandTotalCustLedgEntry[1]."Amount (LCY)")
            {
                AutoFormatType = 1;
            }
            column(GrandTotalCLE1CustRemAmtLCY; Pct(GrandTotalCustLedgEntry[1]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
            {
            }
            column(GrandTotalCLE2CustRemAmtLCY; Pct(GrandTotalCustLedgEntry[2]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
            {
            }
            column(GrandTotalCLE3CustRemAmtLCY; Pct(GrandTotalCustLedgEntry[3]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
            {
            }
            column(GrandTotalCLE4CustRemAmtLCY; Pct(GrandTotalCustLedgEntry[4]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
            {
            }
            column(GrandTotalCLE5CustRemAmtLCY; Pct(GrandTotalCustLedgEntry[5]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
            {
            }
            column(AgedAccReceivableCptn; AgedAccReceivableCptnLbl)
            {
            }
            column(CurrReportPageNoCptn; CurrReportPageNoCptnLbl)
            {
            }
            column(AllAmtinLCYCptn; AllAmtinLCYCptnLbl)
            {
            }
            column(AgedOverdueAmtCptn; AgedOverdueAmtCptnLbl)
            {
            }
            column(CLEEndDateAmtLCYCptn; CLEEndDateAmtLCYCptnLbl)
            {
            }
            column(CLEEndDateDueDateCptn; CLEEndDateDueDateCptnLbl)
            {
            }
            column(CLEEndDateDocNoCptn; CLEEndDateDocNoCptnLbl)
            {
            }
            column(CLEEndDatePstngDateCptn; CLEEndDatePstngDateCptnLbl)
            {
            }
            column(CLEEndDateDocTypeCptn; CLEEndDateDocTypeCptnLbl)
            {
            }
            column(OriginalAmtCptn; OriginalAmtCptnLbl)
            {
            }
            column(TotalLCYCptn; TotalLCYCptnLbl)
            {
            }
            column(NewPagePercustomer; NewPagePercustomer)
            {
            }
            column(PageGroupNo; PageGroupNo)
            {
            }
            column(ResponsibilityCenter_Customer; Customer."Responsibility Center")
            {
            }
            column(CustCreditLCY; Customer."Credit Limit (LCY)")
            {
            }
            dataitem("Cust. Ledger Entry"; "Cust. Ledger Entry")
            {
                DataItemLink = "Customer No." = FIELD("No.");
                DataItemTableView = SORTING("Customer No.", "Posting Date", "Currency Code");

                trigger OnAfterGetRecord()
                var
                    CustLedgEntry: Record 21;
                begin
                    //>>1
                    CustLedgEntry.SETCURRENTKEY("Closed by Entry No.");
                    CustLedgEntry.SETRANGE("Closed by Entry No.", "Entry No.");
                    CustLedgEntry.SETRANGE("Posting Date", 0D, EndingDate);
                    IF CustLedgEntry.FINDSET(FALSE, FALSE) THEN
                        REPEAT
                            InsertTemp(CustLedgEntry);
                        UNTIL CustLedgEntry.NEXT = 0;

                    IF "Closed by Entry No." <> 0 THEN BEGIN
                        CustLedgEntry.SETRANGE("Closed by Entry No.", "Closed by Entry No.");
                        IF CustLedgEntry.FINDSET(FALSE, FALSE) THEN
                            REPEAT
                                InsertTemp(CustLedgEntry);
                            UNTIL CustLedgEntry.NEXT = 0;
                    END;

                    CustLedgEntry.RESET;
                    CustLedgEntry.SETRANGE("Entry No.", "Closed by Entry No.");
                    CustLedgEntry.SETRANGE("Posting Date", 0D, EndingDate);
                    IF CustLedgEntry.FINDSET(FALSE, FALSE) THEN
                        REPEAT
                            InsertTemp(CustLedgEntry);
                        UNTIL CustLedgEntry.NEXT = 0;
                    //<<1

                    CurrReport.SKIP;
                end;

                trigger OnPreDataItem()
                begin
                    SETRANGE("Posting Date", EndingDate + 1, 99991231D/*31129999D*/);
                end;
            }
            dataitem(OpenCustLedgEntry; "Cust. Ledger Entry")
            {
                DataItemLink = "Customer No." = FIELD("No.");
                DataItemTableView = SORTING("Customer No.", Open, Positive, "Due Date", "Currency Code");

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    CASE AgingBy OF
                        /*AgingBy::"Due Date":
                          IF "Due Date" > EndingDate THEN
                            CurrReport.SKIP; */
                        AgingBy::"Document Date":
                            IF "Document Date" > EndingDate THEN
                                CurrReport.SKIP;
                        AgingBy::"Posting Date":
                            BEGIN
                                CALCFIELDS("Remaining Amt. (LCY)");
                                IF "Remaining Amt. (LCY)" = 0 THEN
                                    CurrReport.SKIP;
                            END;
                    END;

                    InsertTemp(OpenCustLedgEntry);
                    //<<1
                    /*
                    IF AgingBy = AgingBy::"Posting Date" THEN BEGIN
                      CALCFIELDS("Remaining Amt. (LCY)");
                      IF "Remaining Amt. (LCY)" = 0 THEN
                        CurrReport.SKIP;
                    END;
                    
                    InsertTemp(OpenCustLedgEntry);
                    CurrReport.SKIP;
                    */

                end;

                trigger OnPreDataItem()
                begin
                    IF AgingBy = AgingBy::"Posting Date" THEN BEGIN
                        SETRANGE("Posting Date", 0D, EndingDate);
                        SETRANGE("Date Filter", 0D, EndingDate);
                    END;
                end;
            }
            dataitem(CurrencyLoop; Integer)
            {
                DataItemTableView = SORTING(Number)
                                    WHERE(Number = FILTER(1 ..));
                PrintOnlyIfDetail = true;
                dataitem(TempCustLedgEntryLoop; Integer)
                {
                    DataItemTableView = SORTING(Number)
                                        WHERE(Number = FILTER(1 ..));
                    column(Name1_Cust; Customer.Name)
                    {
                        IncludeCaption = true;
                    }
                    column(No_Cust; Customer."No.")
                    {
                        IncludeCaption = true;
                    }
                    column(CLEEndDateRemAmtLCY; CustLedgEntryEndingDate."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(AgedCLE1RemAmtLCY; AgedCustLedgEntry[1]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(AgedCLE2RemAmtLCY; AgedCustLedgEntry[2]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(AgedCLE3RemAmtLCY; AgedCustLedgEntry[3]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(AgedCLE4RemAmtLCY; AgedCustLedgEntry[4]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(AgedCLE5RemAmtLCY; AgedCustLedgEntry[5]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(CLEEndDateAmtLCY; CustLedgEntryEndingDate."Amount (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(CLEEndDueDate; FORMAT(CustLedgEntryEndingDate."Due Date"))
                    {
                    }
                    column(duedate; DueDate)
                    {
                    }
                    column(CLEEndDateDocNo; CustLedgEntryEndingDate."Document No.")
                    {
                    }
                    column(CLEDocType; FORMAT(CustLedgEntryEndingDate."Document Type"))
                    {
                    }
                    column(CLEPostingDate; FORMAT(CustLedgEntryEndingDate."Posting Date"))
                    {
                    }
                    column(AgedCLE5TempRemAmt; AgedCustLedgEntry[5]."Remaining Amount")
                    {
                        AutoFormatExpression = CurrencyCode;
                        AutoFormatType = 1;
                    }
                    column(AgedCLE4TempRemAmt; AgedCustLedgEntry[4]."Remaining Amount")
                    {
                        AutoFormatExpression = CurrencyCode;
                        AutoFormatType = 1;
                    }
                    column(AgedCLE3TempRemAmt; AgedCustLedgEntry[3]."Remaining Amount")
                    {
                        AutoFormatExpression = CurrencyCode;
                        AutoFormatType = 1;
                    }
                    column(AgedCLE2TempRemAmt; AgedCustLedgEntry[2]."Remaining Amount")
                    {
                        AutoFormatExpression = CurrencyCode;
                        AutoFormatType = 1;
                    }
                    column(AgedCLE1TempRemAmt; AgedCustLedgEntry[1]."Remaining Amount")
                    {
                        AutoFormatExpression = CurrencyCode;
                        AutoFormatType = 1;
                    }
                    column(RemAmt_CLEEndDate; CustLedgEntryEndingDate."Remaining Amount")
                    {
                        AutoFormatExpression = CurrencyCode;
                        AutoFormatType = 1;
                    }
                    column(CLEEndDate; CustLedgEntryEndingDate.Amount)
                    {
                        AutoFormatExpression = CurrencyCode;
                        AutoFormatType = 1;
                    }
                    column(Name_Cust; STRSUBSTNO(Text005, Customer.Name))
                    {
                    }
                    column(TotalCLE1AmtLCY; TotalCustLedgEntry[1]."Amount (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE1RemAmtLCY; TotalCustLedgEntry[1]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE2RemAmtLCY; TotalCustLedgEntry[2]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE3RemAmtLCY; TotalCustLedgEntry[3]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE4RemAmtLCY; TotalCustLedgEntry[4]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE5RemAmtLCY; TotalCustLedgEntry[5]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(CurrrencyCode; CurrencyCode)
                    {
                        AutoFormatExpression = CurrencyCode;
                        AutoFormatType = 1;
                    }
                    column(TotalCLE5RemAmt; TotalCustLedgEntry[5]."Remaining Amount")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE4RemAmt; TotalCustLedgEntry[4]."Remaining Amount")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE3RemAmt; TotalCustLedgEntry[3]."Remaining Amount")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE2RemAmt; TotalCustLedgEntry[2]."Remaining Amount")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE1RemAmt; TotalCustLedgEntry[1]."Remaining Amount")
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCLE1Amt; TotalCustLedgEntry[1].Amount)
                    {
                        AutoFormatType = 1;
                    }
                    column(TotalCheck; CustFilterCheck)
                    {
                    }
                    column(GrandTotalCLE1AmtLCY; GrandTotalCustLedgEntry[1]."Amount (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(GrandTotalCLE5PctRemAmtLCY; Pct(GrandTotalCustLedgEntry[5]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
                    {
                    }
                    column(GrandTotalCLE3PctRemAmtLCY; Pct(GrandTotalCustLedgEntry[3]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
                    {
                    }
                    column(GrandTotalCLE2PctRemAmtLCY; Pct(GrandTotalCustLedgEntry[2]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
                    {
                    }
                    column(GrandTotalCLE1PctRemAmtLCY; Pct(GrandTotalCustLedgEntry[1]."Remaining Amt. (LCY)", GrandTotalCustLedgEntry[1]."Amount (LCY)"))
                    {
                    }
                    column(GrandTotalCLE5RemAmtLCY; GrandTotalCustLedgEntry[5]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(GrandTotalCLE4RemAmtLCY; GrandTotalCustLedgEntry[4]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(GrandTotalCLE3RemAmtLCY; GrandTotalCustLedgEntry[3]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(GrandTotalCLE2RemAmtLCY; GrandTotalCustLedgEntry[2]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(GrandTotalCLE1RemAmtLCY; GrandTotalCustLedgEntry[1]."Remaining Amt. (LCY)")
                    {
                        AutoFormatType = 1;
                    }
                    column(Region; Region)
                    {
                    }
                    column(salespersonvalue; salespersonvalue)
                    {
                    }
                    column(resposiblityvalue; resposiblityvalue)
                    {
                    }
                    column(creditvalue; creditvalue)
                    {
                    }
                    column(DivisionCode; DivisionCode)
                    {
                    }

                    trigger OnAfterGetRecord()
                    var
                        PeriodIndex: Integer;
                    begin
                        IF Number = 1 THEN BEGIN
                            IF NOT TempCustLedgEntry.FINDSET(FALSE, FALSE) THEN
                                CurrReport.BREAK;
                        END ELSE
                            IF TempCustLedgEntry.NEXT = 0 THEN
                                CurrReport.BREAK;


                        // EBT MILAN 100114 To Upnd (S) Behind Supp. Invoice.
                        CLEAR(SuppText);
                        suup.RESET;
                        suup.SETRANGE(suup."No.", CustLedgEntryEndingDate."Document No.");
                        suup.SETRANGE(suup."Supplimentary Invoice", TRUE);
                        IF suup.FINDFIRST THEN
                            SuppText := '(S)'
                        ELSE
                            SuppText := '';
                        //

                        CustLedgEntryEndingDate := TempCustLedgEntry;
                        DetailedCustomerLedgerEntry.SETRANGE("Cust. Ledger Entry No.", CustLedgEntryEndingDate."Entry No.");
                        IF DetailedCustomerLedgerEntry.FINDSET(FALSE, FALSE) THEN
                            REPEAT
                                IF (DetailedCustomerLedgerEntry."Entry Type" =
                                    DetailedCustomerLedgerEntry."Entry Type"::"Initial Entry") AND
                                   (CustLedgEntryEndingDate."Posting Date" > EndingDate) AND
                                   (AgingBy <> AgingBy::"Posting Date")
                                THEN BEGIN
                                    IF CustLedgEntryEndingDate."Document Date" <= EndingDate THEN
                                        DetailedCustomerLedgerEntry."Posting Date" :=
                                          CustLedgEntryEndingDate."Document Date"
                                    ELSE
                                        IF (CustLedgEntryEndingDate."Due Date" <= EndingDate) AND
                                           (AgingBy = AgingBy::"Due Date")
                                        THEN
                                            DetailedCustomerLedgerEntry."Posting Date" :=
                                              CustLedgEntryEndingDate."Due Date"
                                END;

                                IF (DetailedCustomerLedgerEntry."Posting Date" <= EndingDate) OR
                                   (TempCustLedgEntry.Open AND
                                    (AgingBy = AgingBy::"Due Date") AND
                                    (CustLedgEntryEndingDate."Due Date" > EndingDate) AND
                                    (CustLedgEntryEndingDate."Posting Date" <= EndingDate))
                                THEN BEGIN
                                    IF DetailedCustomerLedgerEntry."Entry Type" IN
                                       [DetailedCustomerLedgerEntry."Entry Type"::"Initial Entry",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Unrealized Loss",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Unrealized Gain",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Realized Loss",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Realized Gain",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Discount",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Discount (VAT Excl.)",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Discount (VAT Adjustment)",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Tolerance",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Discount Tolerance",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Tolerance (VAT Excl.)",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Tolerance (VAT Adjustment)",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Discount Tolerance (VAT Excl.)",
                                        DetailedCustomerLedgerEntry."Entry Type"::"Payment Discount Tolerance (VAT Adjustment)"]
                                    THEN BEGIN
                                        CustLedgEntryEndingDate.Amount := CustLedgEntryEndingDate.Amount + DetailedCustomerLedgerEntry.Amount;
                                        CustLedgEntryEndingDate."Amount (LCY)" :=
                                          CustLedgEntryEndingDate."Amount (LCY)" + DetailedCustomerLedgerEntry."Amount (LCY)";
                                    END;
                                    IF DetailedCustomerLedgerEntry."Posting Date" <= EndingDate THEN BEGIN
                                        CustLedgEntryEndingDate."Remaining Amount" :=
                                          CustLedgEntryEndingDate."Remaining Amount" + DetailedCustomerLedgerEntry.Amount;
                                        CustLedgEntryEndingDate."Remaining Amt. (LCY)" :=
                                          CustLedgEntryEndingDate."Remaining Amt. (LCY)" + DetailedCustomerLedgerEntry."Amount (LCY)";
                                    END;
                                END;
                            UNTIL DetailedCustomerLedgerEntry.NEXT = 0;

                        IF CustLedgEntryEndingDate."Remaining Amount" = 0 THEN
                            CurrReport.SKIP;

                        //>>28july2017
                        //DueDate:=CALCDATE(calcform,CustLedgEntryEndingDate."Posting Date");
                        //DueDate := CALCDATE(Customer."Approved Payment Days",CustLedgEntryEndingDate."Due Date");//28July2017
                        //DueDate := CALCDATE('0D',CustLedgEntryEndingDate."Due Date");//05Aug2017

                        //<<28July2017
                        /*
                        //>>18Nov2017 DueDate Calculation for AgingBy::"Due Date"
                        //IF AgingBy = AgingBy::"Due Date" THEN
                        //BEGIN
                          CLEAR(PayCode18);
                          PayTerm18.RESET;
                          PayTerm18.SETRANGE(Code,Customer."Payment Terms Code");
                          IF PayTerm18.FINDFIRST THEN
                          BEGIN
                            PayCode18 := PayTerm18."Due Date Calculation";
                          END;
                        
                          DueDate := CALCDATE(PayCode18,CustLedgEntryEndingDate."Posting Date");
                        
                        //END;
                        //>>18Nov2017 DueDate Calculation for AgingBy::"Due Date"
                        */

                        DueDate := CustLedgEntryEndingDate."Due Date";//29Aug2018

                        CASE AgingBy OF
                            AgingBy::"Due Date":
                                //PeriodIndex := GetPeriodIndex(CustLedgEntryEndingDate."Due Date");
                                PeriodIndex := GetPeriodIndex(DueDate);//28July2017
                            AgingBy::"Posting Date":
                                PeriodIndex := GetPeriodIndex(CustLedgEntryEndingDate."Posting Date");
                            AgingBy::"Document Date":
                                BEGIN
                                    IF CustLedgEntryEndingDate."Document Date" > EndingDate THEN BEGIN
                                        CustLedgEntryEndingDate."Remaining Amount" := 0;
                                        CustLedgEntryEndingDate."Remaining Amt. (LCY)" := 0;
                                        CustLedgEntryEndingDate."Document Date" := CustLedgEntryEndingDate."Posting Date";
                                    END;
                                    PeriodIndex := GetPeriodIndex(CustLedgEntryEndingDate."Document Date");
                                END;
                        END;
                        CLEAR(AgedCustLedgEntry);
                        AgedCustLedgEntry[PeriodIndex]."Remaining Amount" := CustLedgEntryEndingDate."Remaining Amount";
                        AgedCustLedgEntry[PeriodIndex]."Remaining Amt. (LCY)" := CustLedgEntryEndingDate."Remaining Amt. (LCY)";
                        TotalCustLedgEntry[PeriodIndex]."Remaining Amount" += CustLedgEntryEndingDate."Remaining Amount";
                        TotalCustLedgEntry[PeriodIndex]."Remaining Amt. (LCY)" += CustLedgEntryEndingDate."Remaining Amt. (LCY)";
                        GrandTotalCustLedgEntry[PeriodIndex]."Remaining Amt. (LCY)" += CustLedgEntryEndingDate."Remaining Amt. (LCY)";
                        TotalCustLedgEntry[1].Amount += CustLedgEntryEndingDate."Remaining Amount";
                        TotalCustLedgEntry[1]."Amount (LCY)" += CustLedgEntryEndingDate."Remaining Amt. (LCY)";
                        GrandTotalCustLedgEntry[1]."Amount (LCY)" += CustLedgEntryEndingDate."Remaining Amt. (LCY)";


                        //>>2

                        IF PrintDetails = TRUE THEN BEGIN
                            RecDimSet.RESET;//10Apr2017
                                            //DimensionValue.RESET;
                            IF CustLedgEntryEndingDate."Document Type" = CustLedgEntryEndingDate."Document Type"::Invoice THEN BEGIN
                                SalesInvoiceHeader.RESET;
                                IF SalesInvoiceHeader.GET(CustLedgEntryEndingDate."Document No.") THEN;
                                SalesShipmentHeader.RESET;
                                SalesShipmentHeader.SETRANGE(SalesShipmentHeader."Order No.", SalesInvoiceHeader."Order No.");
                                IF SalesShipmentHeader.FINDFIRST THEN BEGIN

                                    RecDimSet.SETRANGE("Dimension Set ID", SalesShipmentHeader."Dimension Set ID");
                                    RecDimSet.SETFILTER("Dimension Code", 'DIVISION');
                                    RecDimSet.SETRANGE("Dimension Value Code", SalesShipmentHeader."Shortcut Dimension 1 Code");
                                    IF RecDimSet.FINDFIRST THEN BEGIN

                                    END;
                                    //DimensionValue.GET('Division',SalesShipmentHeader."Shortcut Dimension 1 Code");
                                END;
                            END;
                        END;


                        //DimensionValue.RESET;
                        RecDimSet.RESET;
                        RecDimSet.SETRANGE("Dimension Set ID", TempCustLedgEntry."Dimension Set ID");
                        RecDimSet.SETRANGE("Dimension Value Code", TempCustLedgEntry."Global Dimension 1 Code");
                        //DimensionValue.SETRANGE(DimensionValue.Code,TempCustLedgEntry."Global Dimension 1 Code");
                        IF RecDimSet.FINDFIRST THEN BEGIN

                            RecDimSet.CALCFIELDS("Dimension Value Name");
                            DivisionCode := RecDimSet."Dimension Value Name";
                        END ELSE
                            DivisionCode := '';

                        /*
                        IF DimensionValue.FINDFIRST THEN
                          DivisionCode := DimensionValue.Name
                         ELSE
                          DivisionCode :='';
                        */


                        //<<2

                        //>>RSPl/Migration/Rahul

                        RespTotalCustLedgEntry[PeriodIndex]."Remaining Amt. (LCY)" += CustLedgEntryEndingDate."Remaining Amt. (LCY)";
                        RespTotalCustLedgEntry[PeriodIndex]."Remaining Amount" += CustLedgEntryEndingDate."Remaining Amount";
                        RespTotalCustLedgEntry[1]."Amount (LCY)" += CustLedgEntryEndingDate."Remaining Amt. (LCY)";
                        RespTotalCustLedgEntry[1].Amount += CustLedgEntryEndingDate."Remaining Amount";

                        CLEAR(ZoneCode);//RSPLSUM 16Apr2020
                        "Sales/purchaser".RESET;
                        "Sales/purchaser".SETRANGE("Sales/purchaser".Code, Customer."Salesperson Code");
                        IF "Sales/purchaser".FINDFIRST THEN BEGIN
                            salespersonvalue := "Sales/purchaser".Name;
                            ZoneCode := "Sales/purchaser"."Zone Code";//RSPLSUM 16Apr2020
                        END;//RSPLSUM 16Apr2020

                        resposiblityctr.RESET;
                        resposiblityctr.SETRANGE(resposiblityctr.Code, Customer."Responsibility Center");
                        IF resposiblityctr.FINDFIRST THEN
                            resposiblityvalue := resposiblityctr.Name;

                        IF PrintToExcel AND PrintDetails THEN
                            MakeExcelDataBody;
                        //<<

                    end;

                    trigger OnPostDataItem()
                    begin
                        IF NOT PrintAmountInLCY THEN
                            UpdateCurrencyTotals;
                        //>>RSPl
                        IF PrintToExcel AND PrintDetails AND (TotalCustLedgEntry[1].Amount <> 0) THEN
                            MakeExcelDataFooter
                        ELSE
                            IF PrintToExcel AND (NOT PrintDetails) AND (TotalCustLedgEntry[1].Amount <> 0) THEN
                                MakeExcelDataFooterforDetail;
                    end;

                    trigger OnPreDataItem()
                    begin
                        IF NOT PrintAmountInLCY THEN BEGIN
                            IF (TempCurrency.Code = '') OR (TempCurrency.Code = GLSetup."LCY Code") THEN
                                TempCustLedgEntry.SETFILTER("Currency Code", '%1|%2', GLSetup."LCY Code", '')
                            ELSE
                                TempCustLedgEntry.SETRANGE("Currency Code", TempCurrency.Code);
                        END;

                        //>>1 10Apr2017
                        paymnettermsrec.RESET;
                        paymnettermsrec.SETRANGE(paymnettermsrec.Code, Customer."Payment Terms Code");
                        IF paymnettermsrec.FINDFIRST THEN BEGIN
                            creditvalue := paymnettermsrec.Description;
                            calcform := FORMAT(paymnettermsrec."Due Date Calculation");
                        END;

                        //<<1

                        PageGroupNo := NextPageGroupNo;
                        IF NewPagePercustomer AND (NumberOfCurrencies > 0) THEN
                            NextPageGroupNo := PageGroupNo + 1;
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    CLEAR(TotalCustLedgEntry);

                    IF Number = 1 THEN BEGIN
                        IF NOT TempCurrency.FINDSET(FALSE, FALSE) THEN
                            CurrReport.BREAK;
                    END ELSE
                        IF TempCurrency.NEXT = 0 THEN
                            CurrReport.BREAK;

                    IF TempCurrency.Code <> '' THEN
                        CurrencyCode := TempCurrency.Code
                    ELSE
                        CurrencyCode := GLSetup."LCY Code";

                    NumberOfCurrencies := NumberOfCurrencies + 1;
                end;

                trigger OnPreDataItem()
                begin
                    NumberOfCurrencies := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin
                IF NewPagePercustomer THEN
                    PageGroupNo += 1;

                TempCurrency.RESET;
                TempCurrency.DELETEALL;
                TempCustLedgEntry.RESET;
                TempCustLedgEntry.DELETEALL;



                CLEAR(Region);
                CLEAR(RegionCode);
                CLEAR(RegionCodeSP);//RSPLSUM 22Jun2020
                recCust.RESET;
                recCust.SETRANGE(recCust."No.", "No.");
                IF recCust.FINDFIRST THEN BEGIN
                    recState.RESET;
                    recState.SETRANGE(recState.Code, recCust."State Code");
                    IF recState.FINDFIRST THEN BEGIN
                        Region := 'Region' + ': ' + '';//recState."Region Code";
                        RegionCode := '';//recState."Region Code";
                    END;

                    RecSalesperPuchser.RESET;//RSPLSUM 22Jun2020
                    IF RecSalesperPuchser.GET(recCust."Salesperson Code") THEN//RSPLSUM 22Jun2020
                        RegionCodeSP := RecSalesperPuchser."Region Code";//RSPLSUM 22Jun2020

                END;

                IF recPaymentTerms.GET(Customer."Payment Terms Code") THEN;   //RSPL-CAS-04861-D0D6V7

                //sn-Begin
                dFCRate := 0;
                recCustLed.RESET;
                recCustLed.SETRANGE(recCustLed."Customer No.", Customer."No.");
                IF recCustLed.FINDFIRST THEN;
                recCustLed.CALCFIELDS(recCustLed.Amount, recCustLed."Amount (LCY)");
            end;

            trigger OnPostDataItem()
            begin

                IF PrintToExcel AND NOT PrintDetails THEN
                    Makeexcelgroupdetail;
            end;

            trigger OnPreDataItem()
            begin

                Customer.SETFILTER(Customer."Responsibility Center", CustResCenterFilter, CustResCenter);
                //Customer.SETRANGE(Customer."Salesperson Code",USERID);
            end;
        }
        dataitem(CurrencyTotals; Integer)
        {
            DataItemTableView = SORTING(Number)
                                WHERE(Number = FILTER(1 ..));
            column(CurrNo; Number = 1)
            {
            }
            column(TempCurrCode; TempCurrency2.Code)
            {
                AutoFormatExpression = CurrencyCode;
                AutoFormatType = 1;
            }
            column(AgedCLE6RemAmt; AgedCustLedgEntry[6]."Remaining Amount")
            {
                AutoFormatExpression = CurrencyCode;
                AutoFormatType = 1;
            }
            column(AgedCLE1RemAmt; AgedCustLedgEntry[1]."Remaining Amount")
            {
                AutoFormatExpression = CurrencyCode;
                AutoFormatType = 1;
            }
            column(AgedCLE2RemAmt; AgedCustLedgEntry[2]."Remaining Amount")
            {
                AutoFormatExpression = CurrencyCode;
                AutoFormatType = 1;
            }
            column(AgedCLE3RemAmt; AgedCustLedgEntry[3]."Remaining Amount")
            {
                AutoFormatExpression = CurrencyCode;
                AutoFormatType = 1;
            }
            column(AgedCLE4RemAmt; AgedCustLedgEntry[4]."Remaining Amount")
            {
                AutoFormatExpression = CurrencyCode;
                AutoFormatType = 1;
            }
            column(AgedCLE5RemAmt; AgedCustLedgEntry[5]."Remaining Amount")
            {
                AutoFormatExpression = CurrencyCode;
                AutoFormatType = 1;
            }
            column(CurrSpecificationCptn; CurrSpecificationCptnLbl)
            {
            }

            trigger OnAfterGetRecord()
            begin
                IF Number = 1 THEN BEGIN
                    IF NOT TempCurrency2.FINDSET(FALSE, FALSE) THEN
                        CurrReport.BREAK;
                END ELSE
                    IF TempCurrency2.NEXT = 0 THEN
                        CurrReport.BREAK;

                CLEAR(AgedCustLedgEntry);
                TempCurrencyAmount.SETRANGE("Currency Code", TempCurrency2.Code);
                IF TempCurrencyAmount.FINDSET(FALSE, FALSE) THEN
                    REPEAT
                        IF TempCurrencyAmount.Date <> 99991231D /*31129999D*/ THEN
                            AgedCustLedgEntry[GetPeriodIndex(TempCurrencyAmount.Date)]."Remaining Amount" :=
                              TempCurrencyAmount.Amount
                        ELSE
                            AgedCustLedgEntry[6]."Remaining Amount" := TempCurrencyAmount.Amount;
                    UNTIL TempCurrencyAmount.NEXT = 0;
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
                    field(AgedAsOf; EndingDate)
                    {
                        Caption = 'Aged As Of';
                    }
                    field(Agingby; AgingBy)
                    {
                        Caption = 'Aging by';
                        OptionCaption = 'Due Date,Posting Date,Document Date';
                    }
                    field(PeriodLength; PeriodLength)
                    {
                        Caption = 'Period Length';
                    }
                    field(AmountsinLCY; PrintAmountInLCY)
                    {
                        Caption = 'Print Amounts in LCY';
                    }
                    field(PrintDetails; PrintDetails)
                    {
                        Caption = 'Print Details';
                    }
                    field(HeadingType; HeadingType)
                    {
                        Caption = 'Heading Type';
                        OptionCaption = 'Date Interval,Number of Days';
                    }
                    field(perCustomer; NewPagePercustomer)
                    {
                        Caption = 'New Page per Customer';
                    }
                    field(PrintToExcel; PrintToExcel)
                    {
                        Caption = 'Print to Excel';
                    }
                }
            }
        }

        actions
        {
        }

        trigger OnOpenPage()
        begin
            IF EndingDate = 0D THEN
                EndingDate := WORKDATE;
        end;
    }

    labels
    {
        BalanceCaption = 'Balance';
    }

    trigger OnPostReport()
    begin

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin
        CustFilter := Customer.GETFILTERS;

        CLEAR(CustResCenterFilter);
        CustResCenterFilter := Customer.GETFILTER(Customer."Responsibility Center");
        SalespersonFilter := Customer.GETFILTER(Customer."Salesperson Code");
        GLSetup.GET;

        CLEAR(CustResCenter);
        recCust.RESET;
        recCust.SETRANGE(recCust."No.", Customer.GETFILTER(Customer."No."));
        IF recCust.FINDFIRST THEN
            CustResCenter := recCust."Responsibility Center";

        CalcDates;
        CreateHeadings;
        //>>RSPL/Migration/Rahul
        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<
        PageGroupNo := 1;
        NextPageGroupNo := 1;
        CustFilterCheck := (CustFilter <> 'No.');
    end;

    var
        GLSetup: Record 98;
        TempCustLedgEntry: Record 21 temporary;
        CustLedgEntryEndingDate: Record 21;
        TotalCustLedgEntry: array[5] of Record 21;
        GrandTotalCustLedgEntry: array[5] of Record 21;
        AgedCustLedgEntry: array[6] of Record 21;
        TempCurrency: Record 4 temporary;
        TempCurrency2: Record 4 temporary;
        TempCurrencyAmount: Record 264 temporary;
        DetailedCustomerLedgerEntry: Record 379;
        CustFilter: Text;
        PrintAmountInLCY: Boolean;
        EndingDate: Date;
        AgingBy: Option "Due Date","Posting Date","Document Date";
        PeriodLength: DateFormula;
        PrintDetails: Boolean;
        HeadingType: Option "Date Interval","Number of Days";
        NewPagePercustomer: Boolean;
        PeriodStartDate: array[5] of Date;
        PeriodEndDate: array[5] of Date;
        HeaderText: array[5] of Text[30];
        Text000: Label 'Not Due';
        Text001: Label 'Before';
        CurrencyCode: Code[10];
        Text002: Label 'days';
        Text003: Label 'More than';
        Text004: Label 'Aged by %1';
        Text005: Label 'Total for %1';
        Text006: Label 'Aged as of %1';
        Text007: Label 'Aged by %1';
        NumberOfCurrencies: Integer;
        Text009: Label 'Due Date,Posting Date,Document Date';
        Text010: Label 'The Date Formula %1 cannot be used. Try to restate it. E.g. 1M+CM instead of CM+1M.';
        PageGroupNo: Integer;
        NextPageGroupNo: Integer;
        CustFilterCheck: Boolean;
        Text032: Label '-%1', Comment = 'Negating the period length: %1 is the period length';
        AgedAccReceivableCptnLbl: Label 'Aged Accounts Receivable';
        CurrReportPageNoCptnLbl: Label 'Page';
        AllAmtinLCYCptnLbl: Label 'All Amounts in LCY';
        AgedOverdueAmtCptnLbl: Label 'Aged Overdue Amounts';
        CLEEndDateAmtLCYCptnLbl: Label 'Original Amount ';
        CLEEndDateDueDateCptnLbl: Label 'Due Date';
        CLEEndDateDocNoCptnLbl: Label 'Document No.';
        CLEEndDatePstngDateCptnLbl: Label 'Posting Date';
        CLEEndDateDocTypeCptnLbl: Label 'Document Type';
        OriginalAmtCptnLbl: Label 'Currency Code';
        TotalLCYCptnLbl: Label 'Total (LCY)';
        CurrSpecificationCptnLbl: Label 'Currency Specification';
        "---RSPL-Rahul----": Integer;
        CustSalesPersonFilter: Text[250];
        CustResCenterFilter: Text[250];
        CustResCenter: Code[20];
        SalespersonFilter: Code[30];
        recCust: Record 18;
        Region: Text[30];
        recState: Record State;
        RegionCode: Text[30];
        ExcelBuf: Record 370 temporary;
        Text011: Label 'Data';
        Text012: Label 'Customer - Aging Detailed';
        Text013: Label 'Company Name';
        Text014: Label 'Report No.';
        Text015: Label 'Report Name';
        Text016: Label 'User ID';
        Text017: Label 'Date';
        Text018: Label 'Customer Filters';
        Text019: Label 'Cust. Ledger Entry Filters';
        Text008: Label 'All Amounts in LCY';
        resposiblityvalue: Text[50];
        salespersonvalue: Text[50];
        DivisionCode: Text[30];
        DueDate: Date;
        RespTotalCustLedgEntry: array[5] of Record 21;
        PrintToExcel: Boolean;
        dFCRate2: Decimal;
        recCustLed: Record 21;
        recPaymentTerms: Record 3;
        dFCRate: Decimal;
        "Sales/purchaser": Record 13;
        resposiblityctr: Record 5714;
        "---10APr2017": Integer;
        SuppText: Text[30];
        suup: Record 112;
        RecDimSet: Record 480;
        SalesInvoiceHeader: Record 112;
        SalesShipmentHeader: Record 110;
        paymnettermsrec: Record 3;
        creditvalue: Text[50];
        calcform: Text[30];
        "--------18Nov2017": Integer;
        PayTerm18: Record 3;
        PayCode18: DateFormula;
        ZoneCode: Code[10];
        RecSalesperPuchser: Record 13;
        RegionCodeSP: Code[10];

    local procedure CalcDates()
    var
        i: Integer;
        PeriodLength2: DateFormula;
    begin
        EVALUATE(PeriodLength2, STRSUBSTNO(Text032, PeriodLength));
        IF AgingBy = AgingBy::"Due Date" THEN BEGIN
            PeriodEndDate[1] := 99991231D;//31129999D;
            PeriodStartDate[1] := EndingDate + 1;
        END ELSE BEGIN
            PeriodEndDate[1] := EndingDate;
            PeriodStartDate[1] := CALCDATE(PeriodLength2, EndingDate + 1);
        END;
        FOR i := 2 TO ARRAYLEN(PeriodEndDate) DO BEGIN
            PeriodEndDate[i] := PeriodStartDate[i - 1] - 1;
            PeriodStartDate[i] := CALCDATE(PeriodLength2, PeriodEndDate[i] + 1);
        END;
        PeriodStartDate[i] := 0D;

        FOR i := 1 TO ARRAYLEN(PeriodEndDate) DO
            IF PeriodEndDate[i] < PeriodStartDate[i] THEN
                ERROR(Text010, PeriodLength);
    end;

    local procedure CreateHeadings()
    var
        i: Integer;
    begin
        IF AgingBy = AgingBy::"Due Date" THEN BEGIN
            HeaderText[1] := Text000;
            i := 2;
        END ELSE
            i := 1;
        WHILE i < ARRAYLEN(PeriodEndDate) DO BEGIN
            IF HeadingType = HeadingType::"Date Interval" THEN
                HeaderText[i] := STRSUBSTNO('%1\..%2', PeriodStartDate[i], PeriodEndDate[i])
            ELSE
                HeaderText[i] :=
                  STRSUBSTNO('%1 - %2 %3', EndingDate - PeriodEndDate[i] + 1, EndingDate - PeriodStartDate[i] + 1, Text002);
            i := i + 1;
        END;
        IF HeadingType = HeadingType::"Date Interval" THEN
            HeaderText[i] := STRSUBSTNO('%1 %2', Text001, PeriodStartDate[i - 1])
        ELSE
            HeaderText[i] := STRSUBSTNO('%1 \%2 %3', Text003, EndingDate - PeriodStartDate[i - 1] + 1, Text002);
    end;

    local procedure InsertTemp(var CustLedgEntry: Record 21)
    var
        Currency: Record 4;
    begin
        WITH TempCustLedgEntry DO BEGIN
            IF GET(CustLedgEntry."Entry No.") THEN
                EXIT;
            TempCustLedgEntry := CustLedgEntry;
            INSERT;
            IF PrintAmountInLCY THEN BEGIN
                CLEAR(TempCurrency);
                TempCurrency."Amount Rounding Precision" := GLSetup."Amount Rounding Precision";
                IF TempCurrency.INSERT THEN;
                EXIT;
            END;
            IF TempCurrency.GET("Currency Code") THEN
                EXIT;
            IF TempCurrency.GET('') AND ("Currency Code" = GLSetup."LCY Code") THEN
                EXIT;
            IF TempCurrency.GET(GLSetup."LCY Code") AND ("Currency Code" = '') THEN
                EXIT;
            IF "Currency Code" <> '' THEN
                Currency.GET("Currency Code")
            ELSE BEGIN
                CLEAR(Currency);
                Currency."Amount Rounding Precision" := GLSetup."Amount Rounding Precision";
            END;
            TempCurrency := Currency;
            TempCurrency.INSERT;
        END;
    end;

    local procedure GetPeriodIndex(Date: Date): Integer
    var
        i: Integer;
    begin
        FOR i := 1 TO ARRAYLEN(PeriodEndDate) DO
            IF Date IN [PeriodStartDate[i] .. PeriodEndDate[i]] THEN
                EXIT(i);
    end;

    local procedure Pct(a: Decimal; b: Decimal): Text[30]
    begin
        IF b <> 0 THEN
            EXIT(FORMAT(ROUND(100 * a / b, 0.1), 0, '<Sign><Integer><Decimals,2>') + '%');
    end;

    local procedure UpdateCurrencyTotals()
    var
        i: Integer;
    begin
        TempCurrency2.Code := CurrencyCode;
        IF TempCurrency2.INSERT THEN;
        WITH TempCurrencyAmount DO BEGIN
            FOR i := 1 TO ARRAYLEN(TotalCustLedgEntry) DO BEGIN
                "Currency Code" := CurrencyCode;
                Date := PeriodStartDate[i];
                IF FIND THEN BEGIN
                    Amount := Amount + TotalCustLedgEntry[i]."Remaining Amount";
                    MODIFY;
                END ELSE BEGIN
                    "Currency Code" := CurrencyCode;
                    Date := PeriodStartDate[i];
                    Amount := TotalCustLedgEntry[i]."Remaining Amount";
                    INSERT;
                END;
            END;
            "Currency Code" := CurrencyCode;
            Date := 99991231D;//31129999D;
            IF FIND THEN BEGIN
                Amount := Amount + TotalCustLedgEntry[1].Amount;
                MODIFY;
            END ELSE BEGIN
                "Currency Code" := CurrencyCode;
                Date := 99991231D;//31129999D; 
                Amount := TotalCustLedgEntry[1].Amount;
                INSERT;
            END;
        END;
    end;

    // //[Scope('Internal')]
    procedure InitializeRequest(NewEndingDate: Date; NewAgingBy: Option; NewPeriodLength: DateFormula; NewPrintAmountInLCY: Boolean; NewPrintDetails: Boolean; NewHeadingType: Option; NewPagePercust: Boolean)
    begin
        EndingDate := NewEndingDate;
        AgingBy := NewAgingBy;
        PeriodLength := NewPeriodLength;
        PrintAmountInLCY := NewPrintAmountInLCY;
        PrintDetails := NewPrintDetails;
        HeadingType := NewHeadingType;
        NewPagePercustomer := NewPagePercust;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//10Apr2017
        ExcelBuf.AddInfoColumn(FORMAT(Text013), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text015), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(FORMAT(Text012), FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text014), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"Customer  Aging Detailed", FALSE, FALSE, FALSE, FALSE, '', 0);//10Apr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text016), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text017), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//10Apr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text018), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(Customer.GETFILTERS, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text019), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn("Cust. Ledger Entry".GETFILTERS, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF PrintAmountInLCY THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddInfoColumn(PrintAmountInLCY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(COPYSTR(Text004, 1, 7)), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(FORMAT(SELECTSTR(AgingBy + 1, Text009)), FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    local procedure MakeExcelDataHeader()
    begin
        //>>**Header
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 03Jun2020
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        //>>**Header2
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Customer Aging Detailed', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 03Jun2020
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        //<<

        //>>10Apr2017
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('As of : ' + FORMAT(EndingDate), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 03Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        //<<10Apr2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Customer.FIELDCAPTION("No."), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Customer.FIELDCAPTION(Name), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Responsibility center', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Salesperson Name', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Region', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Salesperson Region Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn('Zone', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn(TempCustLedgEntry.FIELDCAPTION("Posting Date"), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TempCustLedgEntry.FIELDCAPTION("Document Type"), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TempCustLedgEntry.FIELDCAPTION("Document No."), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TempCustLedgEntry.FIELDCAPTION("Due Date"), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Customer Posting Group', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FC Booked Rate', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Credit Limit', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 03Jun2020
        ExcelBuf.AddColumn('Credit Period of Customer', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);

        IF NOT PrintAmountInLCY THEN BEGIN
            ExcelBuf.AddColumn(TempCustLedgEntry.FIELDCAPTION("Original Amount"), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(TempCustLedgEntry.FIELDCAPTION("Remaining Amount"), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        END ELSE BEGIN
            ExcelBuf.AddColumn(TempCustLedgEntry.FIELDCAPTION("Original Amt. (LCY)"), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(TempCustLedgEntry.FIELDCAPTION("Remaining Amt. (LCY)"), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        END;
        ExcelBuf.AddColumn(HeaderText[1], FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(HeaderText[2], FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(HeaderText[3], FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(HeaderText[4], FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(HeaderText[5], FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Customer."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Customer.Name, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(resposiblityvalue, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(salespersonvalue, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(DivisionCode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RegionCode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RegionCodeSP, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn(ZoneCode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn(TempCustLedgEntry."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//10Apr2017
        ExcelBuf.AddColumn(FORMAT(TempCustLedgEntry."Document Type"), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TempCustLedgEntry."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(DueDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//10Apr2017
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);

        IF CustLedgEntryEndingDate.Amount > 0 THEN//RSPLSUM
            dFCRate := CustLedgEntryEndingDate."Amount (LCY)" / CustLedgEntryEndingDate.Amount;//RSPLSUM

        //SN-BEGIN
        ExcelBuf.AddColumn(TempCustLedgEntry."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11 //10APr2017
                                                                                                              //RSPLSUM 03Jun2020--IF recCustLed."Customer Posting Group"='FOREIGN' THEN BEGIN
        IF (TempCustLedgEntry."Customer Posting Group" = 'FOREIGN') OR (TempCustLedgEntry."Customer Posting Group" = 'FO') THEN BEGIN//RSPLSUM 03Jun2020
            ExcelBuf.AddColumn(CustLedgEntryEndingDate."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12 //10APr2017
            ExcelBuf.AddColumn(dFCRate, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//13 //10APr2017
            ExcelBuf.AddColumn(CustLedgEntryEndingDate.Amount, FALSE, '', FALSE, FALSE, FALSE, '', 0);//14 //10APr2017
        END ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
        END;

        ExcelBuf.AddColumn(Customer."Credit Limit (LCY)", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15 //RSPLSUM 03Jun2020
        ExcelBuf.AddColumn(recPaymentTerms.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//15 //RSPL-CAS-04861-D0D6V7
        //SN-END

        IF PrintAmountInLCY THEN BEGIN
            ExcelBuf.AddColumn(CustLedgEntryEndingDate."Amount (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(CustLedgEntryEndingDate."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[1]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[2]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[3]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[4]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[5]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END ELSE BEGIN
            ExcelBuf.AddColumn(CustLedgEntryEndingDate.Amount, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(CustLedgEntryEndingDate."Remaining Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[1]."Remaining Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[2]."Remaining Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[3]."Remaining Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[4]."Remaining Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(AgedCustLedgEntry[5]."Remaining Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END;
    end;

    //  //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text011,Text012,COMPANYNAME,USERID);
        /* ExcelBuf.CreateBookAndOpenExcel('', Text012, '', '', '');
        ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook(Text012);
        ExcelBuf.WriteSheet(Text012, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        ERROR('');
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//18 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//19 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//20 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//21 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//22 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//23//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//23//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//24//RSPLSUM 03Jun2020

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//5 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//9 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//10 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//11 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//12 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//13 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//14//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//14//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//15//RSPLSUM 03Jun2020
        IF PrintAmountInLCY THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//16
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//17
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//18
            ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Amount (LCY)", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//19
            ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//20
            ExcelBuf.AddColumn(TotalCustLedgEntry[2]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//21
            ExcelBuf.AddColumn(TotalCustLedgEntry[3]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//22
            ExcelBuf.AddColumn(TotalCustLedgEntry[4]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//23
            ExcelBuf.AddColumn(TotalCustLedgEntry[5]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//24
        END ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//16
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//17
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//18
            ExcelBuf.AddColumn(TotalCustLedgEntry[1].Amount, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//19
            ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Remaining Amount", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//20
            ExcelBuf.AddColumn(TotalCustLedgEntry[2]."Remaining Amount", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//21
            ExcelBuf.AddColumn(TotalCustLedgEntry[3]."Remaining Amount", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//22
            ExcelBuf.AddColumn(TotalCustLedgEntry[4]."Remaining Amount", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//23
            ExcelBuf.AddColumn(TotalCustLedgEntry[5]."Remaining Amount", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//24
        END;

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//18 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//19 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//20 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//21 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//22 //10Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//23//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//23//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//24//RSPLSUM 03Jun2020
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataFooterforDetail()
    begin
        /*
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Customer."No.",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Customer.Name,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(resposiblityvalue,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(salespersonvalue,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RegionCode,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        IF PrintAmountInLCY THEN BEGIN
           ExcelBuf.AddColumn(DueDate,FALSE,'',FALSE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
           ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text); //RSPL
            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        
           ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
           ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Amount (LCY)",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Remaining Amt. (LCY)",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[2]."Remaining Amt. (LCY)",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[3]."Remaining Amt. (LCY)",FALSE,'',FALSE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[4]."Remaining Amt. (LCY)",FALSE,'',FALSE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[5]."Remaining Amt. (LCY)",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        END ELSE BEGIN
            ExcelBuf.AddColumn(DueDate,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text); //RSPl
        
           ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
           ExcelBuf.AddColumn(TotalCustLedgEntry[1].Amount,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Remaining Amount",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[2]."Remaining Amount",FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[3]."Remaining Amount",FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[4]."Remaining Amount",FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
           ExcelBuf.AddColumn(TotalCustLedgEntry[5]."Remaining Amount",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        END;
        */

        IF CustLedgEntryEndingDate.Amount > 0 THEN
            dFCRate2 := CustLedgEntryEndingDate."Amount (LCY)" / CustLedgEntryEndingDate.Amount;

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Customer."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Customer.Name, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(resposiblityvalue, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(salespersonvalue, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RegionCode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RegionCodeSP, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn(ZoneCode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        IF PrintAmountInLCY THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Date);//10 DueDate
                                                                                                    //ExcelBuf.AddColumn(recCustLed."Customer Posting Group",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(Customer."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11 //07Aug2017
                                                                                                         //RSPLSUM 03Jun2020--IF recCustLed."Customer Posting Group"='FOREIGN' THEN BEGIN
            IF (CustLedgEntryEndingDate."Customer Posting Group" = 'FOREIGN') OR (CustLedgEntryEndingDate."Customer Posting Group" = 'FO') THEN BEGIN//RSPLSUM 03Jun2020
                ExcelBuf.AddColumn(CustLedgEntryEndingDate."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn(dFCRate2, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                ExcelBuf.AddColumn(CustLedgEntryEndingDate.Amount, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
            END ELSE BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            END;

            ExcelBuf.AddColumn(Customer."Credit Limit (LCY)", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);  //RSPLSUM 03Jun2020
            ExcelBuf.AddColumn(recPaymentTerms.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);  //RSPL-CAS-04861-D0D6V7
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            //ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Amount (LCY)",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Amount (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//17 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//18 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[2]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//19 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[3]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//20 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[4]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//21 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[5]."Remaining Amt. (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22 07Aug2017
        END ELSE BEGIN
            //ExcelBuf.AddColumn(DueDate,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//07Aug2017 //10 DueDate
                                                                                                   //ExcelBuf.AddColumn(recCustLed."Customer Posting Group",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(Customer."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11 //07Aug2017
                                                                                                         //RSPLSUM 03Jun2020--IF recCustLed."Customer Posting Group"='FOREIGN' THEN
            IF (CustLedgEntryEndingDate."Customer Posting Group" = 'FOREIGN') OR (CustLedgEntryEndingDate."Customer Posting Group" = 'FO') THEN //RSPLSUM 03Jun2020
            BEGIN
                ExcelBuf.AddColumn(CustLedgEntryEndingDate."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn(dFCRate2, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                ExcelBuf.AddColumn(CustLedgEntryEndingDate.Amount, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
            END ELSE BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            END;

            ExcelBuf.AddColumn(Customer."Credit Limit (LCY)", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);  //RSPLSUM 03Jun2020
            ExcelBuf.AddColumn(recPaymentTerms.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);  //RSPL-CAS-04861-D0D6V7
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(TotalCustLedgEntry[1].Amount, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//17 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[1]."Remaining Amount", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//18 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[2]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//19 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[3]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//20 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[4]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//21 07Aug2017
            ExcelBuf.AddColumn(TotalCustLedgEntry[5]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//22 07Aug2017
        END;

    end;

    //  //[Scope('Internal')]
    procedure Makeexcelgroupdetail()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        //ExcelBuf.AddColumn(resposiblityvalue,FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//10  //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//11  //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//12  //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//13  //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//14  //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//15//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//15//RSPLSUM 22Jun2020
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//15//RSPLSUM 03Jun2020

        IF PrintAmountInLCY THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//16
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//17
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[1]."Amount (LCY)", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//18
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[1]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//19
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[2]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//20
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[3]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//21
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[4]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//22
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[5]."Remaining Amt. (LCY)", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//23
        END ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[1].Amount, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[1]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[2]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[3]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[4]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(RespTotalCustLedgEntry[5]."Remaining Amount", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END;
        ExcelBuf.NewRow;
    end;
}


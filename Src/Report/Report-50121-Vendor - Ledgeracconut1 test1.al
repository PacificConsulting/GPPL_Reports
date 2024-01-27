report 50121 "Vendor - Ledgeracconut1 test1"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 27Aug2018   RB-N         Displaying Opening Balance
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/VendorLedgeracconut1test1.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Vendor - Detail Trial Balance';
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem(Vendor; Vendor)
        {
            DataItemTableView = SORTING("No.");
            PrintOnlyIfDetail = true;
            RequestFilterFields = "No.", "Search Name", "Vendor Posting Group", "Date Filter";
            column(VendDatetFilterPeriod; STRSUBSTNO(Text000, VendDateFilter))
            {
            }
            column(CompanyName; COMPANYNAME)
            {
            }
            column(VendorTblCapVendFltr; TABLECAPTION + ': ' + VendFilter)
            {
            }
            column(VendFilter; VendFilter)
            {
            }
            column(PageGroupNo; PageGroupNo)
            {
            }
            column(PrintAmountsInLCY; PrintAmountsInLCY)
            {
            }
            column(ExcludeBalanceOnly; ExcludeBalanceOnly)
            {
            }
            column(PrintOnlyOnePerPage; PrintOnlyOnePerPage)
            {
            }
            column(AmountCaption; AmountCaption)
            {
                AutoFormatExpression = "Currency Code";
                AutoFormatType = 1;
            }
            column(RemainingAmtCaption; RemainingAmtCaption)
            {
                AutoFormatExpression = "Currency Code";
                AutoFormatType = 1;
            }
            column(No_Vendor; "No.")
            {
            }
            column(Name_Vendor; Name)
            {
            }
            column(PhoneNo_Vendor; "Phone No.")
            {
                IncludeCaption = true;
            }
            column(StartBalanceLCY; StartBalanceLCY)
            {
                AutoFormatType = 1;
            }
            column(StartBalAdjLCY; StartBalAdjLCY)
            {
                AutoFormatType = 1;
            }
            column(VendBalanceLCY; VendBalanceLCY)
            {
                AutoFormatType = 1;
            }
            column(StrtBalLCYStartBalAdjLCY; StartBalanceLCY + StartBalAdjLCY)
            {
                AutoFormatType = 1;
            }
            column(VendDetailTrialBalCap; VendDetailTrialBalCapLbl)
            {
            }
            column(PageCaption; PageCaptionLbl)
            {
            }
            column(AllamountsareinLCYCaption; AllamountsareinLCYCaptionLbl)
            {
            }
            column(ReportIncludesvendorshavebalanceCaption; ReportIncludesvendorshavebalanceCaptionLbl)
            {
            }
            column(PostingDateCaption; PostingDateCaptionLbl)
            {
            }
            column(BalanceLCYCaption; BalanceLCYCaptionLbl)
            {
            }
            column(DueDateCaption; DueDateCaptionLbl)
            {
            }
            column(AdjofOpeningBalanceCaption; AdjofOpeningBalanceCaptionLbl)
            {
            }
            column(TotalLCYCaption; TotalLCYCaptionLbl)
            {
            }
            column(TotalAdjofOpenBalCaption; TotalAdjofOpenBalCaptionLbl)
            {
            }
            column(TotalLCYBeforePeriodCaption; TotalLCYBeforePeriodCaptionLbl)
            {
            }
            column(OpeningCrBal; ABS(OpenBalance))
            {
            }
            column(OpeningDrBal; ABS(OpeningBalance1))
            {
            }
            dataitem("Vendor Ledger Entry"; "Vendor Ledger Entry")
            {
                DataItemLink = "Vendor No." = FIELD("No."),
                               "Posting Date" = FIELD("Date Filter"),
                               "Global Dimension 1 Code" = FIELD("Global Dimension 1 Filter"),
                               "Global Dimension 2 Code" = FIELD("Global Dimension 2 Filter"),
                               "Date Filter" = FIELD("Date Filter");
                DataItemTableView = SORTING("Vendor No.", "Posting Date");
                column(PostingDate_VendLedgEntry; FORMAT("Posting Date"))
                {
                }
                column(DocType_VendLedgEntry; "Document Type")
                {
                    IncludeCaption = true;
                }
                column(DocNo_VendLedgerEntry; "Document No.")
                {
                    IncludeCaption = true;
                }
                column(Desc_VendLedgerEntry; Description)
                {
                    IncludeCaption = true;
                }
                column(VendAmount; VendAmount)
                {
                    AutoFormatExpression = VendCurrencyCode;
                    AutoFormatType = 1;
                }
                column(VendBalLCY2; VendBalanceLCY)
                {
                    AutoFormatType = 1;
                }
                column(VendRemainAmount; VendRemainAmount)
                {
                    AutoFormatExpression = VendCurrencyCode;
                    AutoFormatType = 1;
                }
                column(VendEntryDueDate; FORMAT(VendEntryDueDate))
                {
                }
                column(EntryNo_VendorLedgerEntry; "Entry No.")
                {
                    IncludeCaption = true;
                }
                column(VendCurrencyCodes; VendCurrencyCode)
                {
                }
                column(Narration; Narration)
                {
                }
                column(ChequeDatails; ChequeDatails)
                {
                }
                column(Cheque; Cheque)
                {
                }
                column(DebitAmountLCY_VendorLedgerEntry; "Vendor Ledger Entry"."Debit Amount (LCY)")
                {
                }
                column(CreditAmountLCY_VendorLedgerEntry; "Vendor Ledger Entry"."Credit Amount (LCY)")
                {
                }
                column(Balance1; Balance1)
                {
                }
                column(vServiceTaxAmt; vServiceTaxAmt)
                {
                }
                column(TotalTDS; -"Vendor Ledger Entry"."Total TDS Including SHE CESS")
                {
                }
                column(SourceCode_VendorLedgerEntry; "Vendor Ledger Entry"."Source Code")
                {
                }
                column(PurchaseLCY_VendorLedgerEntry; ABS("Vendor Ledger Entry"."Purchase (LCY)"))
                {
                }
                column(OpenBalance; OpenBalance)
                {
                }
                column(OpeningBalance1; OpeningBalance1)
                {
                }
                column(DebitAmtLCY_InitialEntry; RecDetVenLedEnt."Debit Amount (LCY)")
                {
                }
                column(CreditAmtLCY_InitialEntry; RecDetVenLedEnt."Credit Amount (LCY)")
                {
                }
                column(GainLossDebitAmt; GainLossDebit)
                {
                }
                column(GainLossCreditAmt; GainLossCredit)
                {
                }
                column(GainLossDebitNew; GainLossDebitNew)
                {
                }
                column(GainLossCreditNew; GainLossCreditNew)
                {
                }
                column(GainLossBal; GainLossBal)
                {
                }
                column(GainLossDebitNewLastEntry; GainLossDebitNewLastEntry)
                {
                }
                column(GainLossCreditNewLastEntry; GainLossCreditNewLastEntry)
                {
                }
                column(LEntry; LEntry)
                {
                }
                column(LastEntryGainLossBal; LastEntryGainLossBal)
                {
                }
                dataitem("Detailed Vendor Ledg. Entry"; "Detailed Vendor Ledg. Entry")
                {
                    DataItemLink = "Vendor Ledger Entry No." = FIELD("Entry No.");
                    DataItemTableView = SORTING("Vendor Ledger Entry No.", "Entry Type", "Posting Date")
                                        WHERE("Entry Type" = CONST("Correction of Remaining Amount"));
                    column(EntryTyp_DetVendLedgEntry; "Entry Type")
                    {
                    }
                    column(Correction; Correction)
                    {
                        AutoFormatType = 1;
                    }

                    trigger OnAfterGetRecord()
                    begin
                        Correction := Correction + "Amount (LCY)";
                        VendBalanceLCY := VendBalanceLCY + "Amount (LCY)";
                    end;

                    trigger OnPostDataItem()
                    begin
                        SumCorrections := SumCorrections + Correction;
                    end;

                    trigger OnPreDataItem()
                    begin
                        SETFILTER("Posting Date", VendDateFilter);
                        Correction := 0;
                    end;
                }
                dataitem("Detailed Vendor Ledg. Entry2"; "Detailed Vendor Ledg. Entry")
                {
                    DataItemLink = "Vendor Ledger Entry No." = FIELD("Entry No.");
                    DataItemTableView = SORTING("Vendor Ledger Entry No.", "Entry Type", "Posting Date")
                                        WHERE("Entry Type" = CONST("Appln. Rounding"));
                    column(Entry_DetVendLedgEntry2; "Entry Type")
                    {
                    }
                    column(VendBalanceLCY3; VendBalanceLCY)
                    {
                        AutoFormatType = 1;
                    }
                    column(ApplicationRounding; ApplicationRounding)
                    {
                        AutoFormatType = 1;
                    }

                    trigger OnAfterGetRecord()
                    begin
                        ApplicationRounding := ApplicationRounding + "Amount (LCY)";
                        VendBalanceLCY := VendBalanceLCY + "Amount (LCY)";
                    end;

                    trigger OnPreDataItem()
                    begin
                        SETFILTER("Posting Date", VendDateFilter);
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    CLEAR(Narration);
                    CLEAR(ChequeDatails);
                    CALCFIELDS(Amount, "Remaining Amount", "Amount (LCY)", "Remaining Amt. (LCY)");

                    VendLedgEntryExists := TRUE;
                    IF PrintAmountsInLCY THEN BEGIN
                        VendAmount := "Amount (LCY)";
                        VendRemainAmount := "Remaining Amt. (LCY)";
                        VendCurrencyCode := '';
                    END ELSE BEGIN
                        VendAmount := Amount;
                        VendRemainAmount := "Remaining Amount";
                        VendCurrencyCode := "Currency Code";
                    END;
                    VendBalanceLCY := VendBalanceLCY + "Amount (LCY)";
                    IF ("Document Type" = "Document Type"::Payment) OR ("Document Type" = "Document Type"::Refund) THEN
                        VendEntryDueDate := 0D
                    ELSE
                        VendEntryDueDate := "Due Date";

                    //>>Body

                    //CustBalanceLCY := CustBalanceLCY + "Amount (LCY)";


                    PostdNarration.RESET;
                    PostdNarration.SETRANGE(PostdNarration."Document No.", "Vendor Ledger Entry"."Document No.");

                    IF PostdNarration.FINDSET THEN
                        REPEAT
                            Narration := Narration + ' ' + PostdNarration.Narration;
                        UNTIL PostdNarration.NEXT = 0;

                    IF "Vendor Ledger Entry"."Source Code" = 'PURCHASES' THEN BEGIN
                        PurchCommentLine.RESET;
                        PurchCommentLine.SETRANGE(PurchCommentLine."No.", "Vendor Ledger Entry"."Document No.");
                        IF PurchCommentLine.FINDSET THEN
                            REPEAT
                                Narration := Narration + ' ' + PurchCommentLine.Comment;
                            UNTIL PurchCommentLine.NEXT = 0;
                    END;


                    IF "Vendor Ledger Entry"."Source Code" = 'PURCHASES' THEN BEGIN
                        PurchInvH.RESET;
                        PurchInvH.SETRANGE(PurchInvH."Vendor Invoice No.", "Vendor Ledger Entry"."External Document No.");
                        IF PurchInvH.FINDFIRST THEN;
                        Cheque := 'Inv. No.';
                        ChequeDatails := "Vendor Ledger Entry"."External Document No." + '-' + FORMAT("Vendor Ledger Entry"."Document Date");
                    END
                    ELSE BEGIN
                        BankAccLedEntry.RESET;
                        BankAccLedEntry.SETRANGE(BankAccLedEntry."Document No.", "Vendor Ledger Entry"."Document No.");
                        IF BankAccLedEntry.FINDFIRST THEN BEGIN
                            Cheque := 'Chq. No:-';
                            ChequeDatails := BankAccLedEntry."Cheque No.";
                        END
                        ELSE BEGIN
                            Cheque := '';
                            ChequeDatails := '';
                        END;
                    END;
                    CLEAR(vServiceTaxAmt);
                    /* recSTE.RESET;
                    recSTE.SETRANGE(recSTE."Document No.", "Vendor Ledger Entry"."Document No.");
                    IF recSTE.FINDFIRST THEN
                        vServiceTaxAmt := recSTE."Service Tax Amount" + recSTE."eCess Amount" + recSTE."SHE Cess Amount"; */
                    //<<

                    //RSPLSUM 19Nov19>>
                    RecDetVenLedEnt.RESET;
                    RecDetVenLedEnt.SETCURRENTKEY("Ledger Entry Amount", "Vendor Ledger Entry No.", "Entry Type", "Posting Date");
                    RecDetVenLedEnt.SETRANGE("Ledger Entry Amount", TRUE);
                    RecDetVenLedEnt.SETRANGE("Vendor Ledger Entry No.", "Entry No.");
                    RecDetVenLedEnt.SETRANGE("Entry Type", RecDetVenLedEnt."Entry Type"::"Initial Entry");
                    RecDetVenLedEnt.SETFILTER("Posting Date", VendDateFilter);
                    IF RecDetVenLedEnt.FINDFIRST THEN;

                    CLEAR(GainLossDebit);
                    CLEAR(GainLossCredit);
                    RecDVLE.RESET;
                    RecDVLE.SETCURRENTKEY("Ledger Entry Amount", "Vendor Ledger Entry No.", "Entry Type", "Posting Date");
                    RecDVLE.SETRANGE("Ledger Entry Amount", TRUE);
                    RecDVLE.SETRANGE("Vendor Ledger Entry No.", "Entry No.");
                    RecDVLE.SETFILTER("Entry Type", '%1|%2|%3|%4', RecDVLE."Entry Type"::"Unrealized Gain", RecDVLE."Entry Type"::"Unrealized Loss",
                                        RecDVLE."Entry Type"::"Realized Gain", RecDVLE."Entry Type"::"Realized Loss");
                    RecDVLE.SETFILTER("Posting Date", VendDateFilter);
                    IF RecDVLE.FINDSET THEN
                        REPEAT
                            GainLossDebit += RecDVLE."Debit Amount (LCY)";
                            GainLossCredit += RecDVLE."Credit Amount (LCY)"
                        UNTIL RecDVLE.NEXT = 0;


                    CLEAR(GainLossCreditNew);
                    CLEAR(GainLossDebitNew);
                    RecDVLENew.RESET;
                    RecDVLENew.SETCURRENTKEY("Ledger Entry Amount", "Vendor Ledger Entry No.", "Entry Type", "Posting Date");
                    RecDVLENew.SETRANGE("Vendor No.", "Vendor Ledger Entry"."Vendor No.");
                    RecDVLENew.SETRANGE("Ledger Entry Amount", TRUE);
                    //RecDVLENew.SETRANGE("Vendor Ledger Entry No.","Entry No.");
                    RecDVLENew.SETFILTER("Entry Type", '%1|%2|%3|%4', RecDVLE."Entry Type"::"Unrealized Gain", RecDVLE."Entry Type"::"Unrealized Loss",
                                        RecDVLE."Entry Type"::"Realized Gain", RecDVLE."Entry Type"::"Realized Loss");
                    IF Fentry THEN BEGIN
                        RecDVLENew.SETRANGE("Posting Date", Vendor.GETRANGEMIN("Date Filter"), "Vendor Ledger Entry"."Posting Date" - 1);
                        Fentry := FALSE;
                    END ELSE
                        RecDVLENew.SETRANGE("Posting Date", TempPostingDate, "Vendor Ledger Entry"."Posting Date" - 1);
                    IF RecDVLENew.FINDSET THEN
                        REPEAT
                            GainLossDebitNew += RecDVLENew."Debit Amount (LCY)";
                            GainLossCreditNew += RecDVLENew."Credit Amount (LCY)";
                        UNTIL RecDVLENew.NEXT = 0;
                    //RSPLSUM 19Nov19<<

                    IF (LastEntryNo = "Vendor Ledger Entry"."Entry No.") AND (LEntry = TRUE) THEN BEGIN
                        CLEAR(GainLossDebitNewLastEntry);
                        CLEAR(GainLossCreditNewLastEntry);
                        RecDVLENew.RESET;
                        RecDVLENew.SETCURRENTKEY("Ledger Entry Amount", "Vendor Ledger Entry No.", "Entry Type", "Posting Date");
                        RecDVLENew.SETRANGE("Vendor No.", "Vendor Ledger Entry"."Vendor No.");
                        RecDVLENew.SETRANGE("Ledger Entry Amount", TRUE);
                        //RecDVLENew.SETRANGE("Vendor Ledger Entry No.","Entry No.");
                        RecDVLENew.SETFILTER("Entry Type", '%1|%2|%3|%4', RecDVLE."Entry Type"::"Unrealized Gain", RecDVLE."Entry Type"::"Unrealized Loss",
                                            RecDVLE."Entry Type"::"Realized Gain", RecDVLE."Entry Type"::"Realized Loss");
                        RecDVLENew.SETRANGE("Posting Date", LastRecordPostingDate, Vendor.GETRANGEMAX("Date Filter"));
                        IF RecDVLENew.FINDSET THEN
                            REPEAT
                                GainLossDebitNewLastEntry += RecDVLENew."Debit Amount (LCY)";
                                GainLossCreditNewLastEntry += RecDVLENew."Credit Amount (LCY)";
                            UNTIL RecDVLENew.NEXT = 0;

                    END;

                    //Code have been changed not to show adjustment entry and recalculate the balance which is not derived from adj. entry
                    "Vendor Ledger Entry".CALCFIELDS("Amount (LCY)");
                    "Vendor Ledger Entry".CALCFIELDS("Debit Amount (LCY)");
                    "Vendor Ledger Entry".CALCFIELDS("Credit Amount (LCY)");

                    NewCustBalanceLCY := (StartBalanceLCY) + "Vendor Ledger Entry"."Amount (LCY)";
                    IF PostdNarration.GET("Entry No.") THEN;
                    /*
                    IF VendNew THEN
                    BEGIN
                       Balance1 := NewCustBalanceLCY;
                       //Balance1 := Balance1+"Vendor Ledger Entry"."Debit Amount (LCY)"-"Vendor Ledger Entry"."Credit Amount (LCY)";
                       VendNew := FALSE;
                    END
                    ELSE
                       Balance1 := Balance1+RecDetVenLedEnt."Debit Amount (LCY)"-RecDetVenLedEnt."Credit Amount (LCY)";
                    */
                    IF VendNew THEN BEGIN
                        //Balance1 := NewCustBalanceLCY;
                        GainLossBal := DcOpenBal + GainLossDebitNew - GainLossCreditNew;
                        //Balance1 := DcOpenBal+RecDetVenLedEnt."Debit Amount (LCY)"-RecDetVenLedEnt."Credit Amount (LCY)"+GainLossDebit-GainLossCredit;
                        Balance1 := GainLossBal + RecDetVenLedEnt."Debit Amount (LCY)" - RecDetVenLedEnt."Credit Amount (LCY)";
                        VendNew := FALSE;
                    END
                    ELSE BEGIN
                        GainLossBal := Balance1 + GainLossDebitNew - GainLossCreditNew;
                        Balance1 := GainLossBal + RecDetVenLedEnt."Debit Amount (LCY)" - RecDetVenLedEnt."Credit Amount (LCY)";
                    END;

                    IF (LastEntryNo = "Vendor Ledger Entry"."Entry No.") AND (LEntry = TRUE) THEN BEGIN
                        LastEntryGainLossBal := Balance1 + GainLossDebitNewLastEntry - GainLossCreditNewLastEntry;
                        LEntry := FALSE;
                    END;

                    ClosingBalanceNew := LastEntryGainLossBal;

                    TempPostingDate := "Vendor Ledger Entry"."Posting Date";

                    IF PrintToExcel THEN
                        MakeExcelDataBody;

                end;

                trigger OnPreDataItem()
                begin
                    VendLedgEntryExists := FALSE;
                    CurrReport.CREATETOTALS(VendAmount, "Amount (LCY)");
                end;
            }
            dataitem(Integer; Integer)
            {
                DataItemTableView = SORTING(Number)
                                    WHERE(Number = CONST(1));
                column(VendBalanceLCY4; VendBalanceLCY)
                {
                    AutoFormatType = 1;
                }
                column(StartBalAdjLCY1; StartBalAdjLCY)
                {
                }
                column(StartBalanceLCY1; StartBalanceLCY)
                {
                }
                column(VendBalStrtBalStrtBalAdj; VendBalanceLCY - StartBalanceLCY)
                {
                    AutoFormatType = 1;
                }
                column(ClosingBalance1; ABS(ClosingBalance1))
                {
                }
                column(ClosingBalance2; ABS(ClosingBalance2))
                {
                }

                trigger OnAfterGetRecord()
                begin
                    IF NOT VendLedgEntryExists AND ((StartBalanceLCY = 0) OR ExcludeBalanceOnly) THEN BEGIN
                        StartBalanceLCY := 0;
                        CurrReport.SKIP;
                    END;


                    /*
                    IF VendBalanceLCY  > 0 THEN BEGIN
                       ClosingBalance2 := 0;
                       ClosingBalance1 := VendBalanceLCY ;
                    END ELSE
                    IF VendBalanceLCY  < 0 THEN BEGIN
                       ClosingBalance1 := 0;
                       ClosingBalance2 := VendBalanceLCY ;
                    END;
                    */

                    IF VendBalanceLCY > 0 THEN BEGIN
                        ClosingBalance2 := 0;
                        ClosingBalance1 := ClosingBalanceNew;
                    END ELSE
                        IF VendBalanceLCY < 0 THEN BEGIN
                            ClosingBalance1 := 0;
                            ClosingBalance2 := ClosingBalanceNew;
                        END;

                    IF PrintToExcel THEN
                        MakeExcelDataFooter;

                end;
            }

            trigger OnAfterGetRecord()
            begin
                IF PrintOnlyOnePerPage THEN
                    PageGroupNo := PageGroupNo + 1;

                StartBalanceLCY := 0;
                StartBalAdjLCY := 0;
                IF VendDateFilter <> '' THEN BEGIN
                    IF GETRANGEMIN("Date Filter") <> 0D THEN BEGIN
                        SETRANGE("Date Filter", 0D, GETRANGEMIN("Date Filter") - 1);
                        CALCFIELDS("Net Change (LCY)");
                        StartBalanceLCY := -"Net Change (LCY)";
                    END;
                    SETFILTER("Date Filter", VendDateFilter);
                    CALCFIELDS("Net Change (LCY)");
                    StartBalAdjLCY := -"Net Change (LCY)";
                    VendorLedgerEntry.SETCURRENTKEY("Vendor No.", "Posting Date");
                    VendorLedgerEntry.SETRANGE("Vendor No.", "No.");
                    VendorLedgerEntry.SETFILTER("Posting Date", VendDateFilter);
                    IF VendorLedgerEntry.FIND('-') THEN
                        REPEAT
                            VendorLedgerEntry.SETFILTER("Date Filter", VendDateFilter);
                            VendorLedgerEntry.CALCFIELDS("Amount (LCY)");
                            StartBalAdjLCY := StartBalAdjLCY - VendorLedgerEntry."Amount (LCY)";
                            "Detailed Vendor Ledg. Entry".SETCURRENTKEY("Vendor Ledger Entry No.", "Entry Type", "Posting Date");
                            "Detailed Vendor Ledg. Entry".SETRANGE("Vendor Ledger Entry No.", VendorLedgerEntry."Entry No.");
                            "Detailed Vendor Ledg. Entry".SETFILTER("Entry Type", '%1|%2',
                              "Detailed Vendor Ledg. Entry"."Entry Type"::"Correction of Remaining Amount",
                              "Detailed Vendor Ledg. Entry"."Entry Type"::"Appln. Rounding");
                            "Detailed Vendor Ledg. Entry".SETFILTER("Posting Date", VendDateFilter);
                            IF "Detailed Vendor Ledg. Entry".FIND('-') THEN
                                REPEAT
                                    StartBalAdjLCY := StartBalAdjLCY - "Detailed Vendor Ledg. Entry"."Amount (LCY)";
                                UNTIL "Detailed Vendor Ledg. Entry".NEXT = 0;
                            "Detailed Vendor Ledg. Entry".RESET;
                        UNTIL VendorLedgerEntry.NEXT = 0;
                END;
                //CurrReport.PRINTONLYIFDETAIL := ExcludeBalanceOnly OR (StartBalanceLCY = 0);
                VendBalanceLCY := StartBalanceLCY + StartBalAdjLCY;
                //>>Migration
                VendorDetails := 'Vendor No/Name  : ' + Vendor."No." + ' ; ' + Vendor."Full Name";

                IF StartBalanceLCY > 0 THEN BEGIN
                    OpeningBalance1 := 0;
                    OpenBalance := StartBalanceLCY;
                END ELSE
                    IF StartBalanceLCY < 0 THEN BEGIN
                        OpenBalance := 0;
                        OpeningBalance1 := StartBalanceLCY;
                    END;

                //RSPLSUM 19Nov19>>
                /*CLEAR(DcOpenBal);
                RecVenLedEnt.RESET;
                RecVenLedEnt.SETCURRENTKEY("Vendor No.","Posting Date","Global Dimension 1 Code","Global Dimension 2 Code");
                RecVenLedEnt.SETRANGE("Vendor No.",Vendor."No.");
                IF VendDateFilter <> '' THEN
                  RecVenLedEnt.SETRANGE("Posting Date",0D,Vendor.GETRANGEMIN("Date Filter") - 1);
                IF VendDim1Filter <> '' THEN
                  RecVenLedEnt.SETFILTER("Global Dimension 1 Code",VendDim1Filter);
                IF VendDim2Filter <> '' THEN
                  RecVenLedEnt.SETFILTER("Global Dimension 2 Code",VendDim2Filter);
                IF RecVenLedEnt.FINDSET THEN
                  REPEAT
                    RecVenLedEnt.CALCFIELDS("Amount (LCY)");
                    DcOpenBal += RecVenLedEnt."Amount (LCY)";
                  UNTIL RecVenLedEnt.NEXT = 0;
                */
                RecDVLENEew.RESET;
                RecDVLENEew.SETCURRENTKEY("Ledger Entry Amount", "Vendor Ledger Entry No.", "Entry Type", "Posting Date");
                RecDVLENEew.SETRANGE("Ledger Entry Amount", TRUE);
                RecDVLENEew.SETRANGE("Vendor No.", Vendor."No.");
                IF VendDateFilter <> '' THEN
                    RecDVLENEew.SETRANGE("Posting Date", 0D, Vendor.GETRANGEMIN("Date Filter") - 1);
                IF VendDim1Filter <> '' THEN
                    RecDVLENEew.SETFILTER("Initial Entry Global Dim. 1", VendDim1Filter);
                IF VendDim2Filter <> '' THEN
                    RecDVLENEew.SETFILTER("Initial Entry Global Dim. 2", VendDim2Filter);
                RecDVLENEew.CALCSUMS("Amount (LCY)");

                DcOpenBal := RecDVLENEew."Amount (LCY)";

                IF DcOpenBal > 0 THEN BEGIN
                    OpeningBalance1 := 0;
                    OpenBalance := DcOpenBal;
                END ELSE
                    IF DcOpenBal < 0 THEN BEGIN
                        OpenBalance := 0;
                        OpeningBalance1 := DcOpenBal;
                    END;
                //RSPLSUM 19Nov19<<

                RecVLENew.RESET;
                RecVLENew.SETCURRENTKEY("Vendor No.", "Posting Date", "Global Dimension 1 Code", "Global Dimension 2 Code");
                RecVLENew.SETRANGE("Vendor No.", Vendor."No.");
                IF VendDateFilter <> '' THEN
                    RecVLENew.SETFILTER("Posting Date", VendDateFilter);
                IF VendDim1Filter <> '' THEN
                    RecVLENew.SETFILTER("Global Dimension 1 Code", VendDim1Filter);
                IF VendDim2Filter <> '' THEN
                    RecVLENew.SETFILTER("Global Dimension 2 Code", VendDim2Filter);
                IF RecVLENew.FINDLAST THEN BEGIN
                    LastRecordPostingDate := RecVLENew."Posting Date";
                    LastEntryNo := RecVLENew."Entry No.";
                    LEntry := TRUE;
                END;

                TempPostingDate := 0D;

                FirstDate := COPYSTR(FORMAT(CustDateFilter), 1, 8);

                IF PrintToExcel THEN
                    MakeCustHeader;
                VendNew := TRUE;               //For calculating the balance only first time from opening balance calculation
                Fentry := TRUE;
                //<<

            end;

            trigger OnPreDataItem()
            begin
                PageGroupNo := 1;
                SumCorrections := 0;

                CurrReport.NEWPAGEPERRECORD := PrintOnlyOnePerPage;
                CurrReport.CREATETOTALS("Vendor Ledger Entry"."Amount (LCY)", StartBalanceLCY, StartBalAdjLCY, Correction, ApplicationRounding);
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
                        ApplicationArea = all;
                        Caption = 'Show Amounts in LCY';
                    }
                    field(PrintOnlyOnePerPage; PrintOnlyOnePerPage)
                    {
                        ApplicationArea = all;
                        Caption = 'New Page per Vendor';
                    }
                    field(ExcludeBalanceOnly; ExcludeBalanceOnly)
                    {
                        ApplicationArea = all;
                        Caption = 'Exclude Vendors That Have A Balance Only';
                        MultiLine = true;
                    }
                    field(PrintToExcel; PrintToExcel)
                    {
                        ApplicationArea = all;
                        Caption = 'Print To Excel';
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

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin
        VendFilter := Vendor.GETFILTERS;
        VendDateFilter := Vendor.GETFILTER("Date Filter");
        VendDim1Filter := Vendor.GETFILTER("Global Dimension 1 Filter");//RSPLSUM 19Nov2019
        VendDim2Filter := Vendor.GETFILTER("Global Dimension 2 Filter");//RSPLSUM 19Nov2019

        WITH "Vendor Ledger Entry" DO
            IF PrintAmountsInLCY THEN BEGIN
                AmountCaption := FIELDCAPTION("Amount (LCY)");
                RemainingAmtCaption := FIELDCAPTION("Remaining Amt. (LCY)");
            END ELSE BEGIN
                AmountCaption := FIELDCAPTION(Amount);
                RemainingAmtCaption := FIELDCAPTION("Remaining Amount");
            END;


        IF PrintToExcel THEN
            MakeExcelInfo;
    end;

    var
        Text000: Label 'Period: %1';
        VendorLedgerEntry: Record 25;
        VendFilter: Text;
        VendDateFilter: Text[30];
        VendAmount: Decimal;
        VendRemainAmount: Decimal;
        VendBalanceLCY: Decimal;
        VendEntryDueDate: Date;
        StartBalanceLCY: Decimal;
        StartBalAdjLCY: Decimal;
        Correction: Decimal;
        ApplicationRounding: Decimal;
        ExcludeBalanceOnly: Boolean;
        PrintAmountsInLCY: Boolean;
        PrintOnlyOnePerPage: Boolean;
        VendLedgEntryExists: Boolean;
        AmountCaption: Text[30];
        RemainingAmtCaption: Text[30];
        VendCurrencyCode: Code[10];
        PageGroupNo: Integer;
        SumCorrections: Decimal;
        VendDetailTrialBalCapLbl: Label 'Vendor - Detail Trial Balance';
        PageCaptionLbl: Label 'Page';
        AllamountsareinLCYCaptionLbl: Label 'All amounts are in LCY.';
        ReportIncludesvendorshavebalanceCaptionLbl: Label 'This report also includes vendors that only have balances.';
        PostingDateCaptionLbl: Label 'Posting Date';
        BalanceLCYCaptionLbl: Label 'Balance (LCY)';
        DueDateCaptionLbl: Label 'Due Date';
        AdjofOpeningBalanceCaptionLbl: Label 'Adj. of Opening Balance';
        TotalLCYCaptionLbl: Label 'Total (LCY)';
        TotalAdjofOpenBalCaptionLbl: Label 'Total Adj. of Opening Balance';
        TotalLCYBeforePeriodCaptionLbl: Label 'Total (LCY) Before Period';
        "---Roboso--": Integer;
        PrintToExcel: Boolean;
        VendorDetails: Text[100];
        NewCustBalanceLCY: Decimal;
        VendNew: Boolean;
        Balance1: Decimal;
        PurchCommentLine: Record 43;
        PurchInvH: Record 122;
        DocDate: Date;
        //recSTE: Record 16473;
        vServiceTaxAmt: Decimal;
        vTdsAmt: Text[50];
        ExcelBuf: Record 370 temporary;
        Text001: Label 'Appln Rounding:';
        Text0000: Label 'Data';
        Text0001: Label 'Vendor - Ledger';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        FirstDate: Text[30];
        ClosingBalance1: Decimal;
        ClosingBalance2: Decimal;
        BankAccLedEntry: Record 271;
        ChequeDatails: Text[50];
        Cheque: Text[30];
        OpenBalance: Decimal;
        OpeningBalance1: Decimal;
        Narration: Text[1024];
        PostdNarration: Record "Posted Narration";
        DebitAmt: Decimal;
        CreditAmt: Decimal;
        CompanyInfo: Record 79;
        CustDateFilter: Text[30];
        RecDetVenLedEnt: Record 380;
        DcOpenBal: Decimal;
        RecVenLedEnt: Record 25;
        VendDim1Filter: Text;
        VendDim2Filter: Text;
        ClosingBalanceNew: Decimal;
        RecDVLE: Record 380;
        GainLossDebit: Decimal;
        GainLossCredit: Decimal;
        RecDVLENew: Record 380;
        GainLossDebitNew: Decimal;
        GainLossCreditNew: Decimal;
        TempPostingDate: Date;
        Fentry: Boolean;
        GainLossBal: Decimal;
        RecVLENew: Record 25;
        LastRecordPostingDate: Date;
        LEntry: Boolean;
        GainLossDebitNewLastEntry: Decimal;
        GainLossCreditNewLastEntry: Decimal;
        LastEntryNo: Integer;
        LastEntryGainLossBal: Decimal;
        RecDVLENEew: Record 380;

    // //[Scope('Internal')]
    procedure InitializeRequest(NewPrintAmountsInLCY: Boolean; NewPrintOnlyOnePerPage: Boolean; NewExcludeBalanceOnly: Boolean)
    begin
        PrintAmountsInLCY := NewPrintAmountsInLCY;
        PrintOnlyOnePerPage := NewPrintOnlyOnePerPage;
        ExcludeBalanceOnly := NewExcludeBalanceOnly;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        /*
        ExcelBuf.AddInfoColumn(FORMAT(Text0004),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('Vendor - Ledger',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::50129,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        */
        MakeExcelHeader;

    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text0000,Text0001,COMPANYNAME,USERID);
        /*  ExcelBuf.CreateBookAndOpenExcel('', 'Vendor Leger Account', '', '', '');
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook('Vendor Leger Account');
        ExcelBuf.WriteSheet('Vendor Leger Account', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Vendor Leger Account', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // //[Scope('Internal')]
    procedure MakeExcelHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Vendor Ledger', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text); //RSPL-CAS-04786-L4N9T8
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Narration', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);  //RSPL-CAS-04786-L4N9T8
        ExcelBuf.AddColumn('Debit(Rs.)', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Credit(Rs.)', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Balance(Rs.)', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure MakeCustHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(VendorDetails, FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPL-CAS-04786-L4N9T8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPL-CAS-04786-L4N9T8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPL-CAS-04786-L4N9T8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPL-CAS-04786-L4N9T8


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Opening Balance As On:', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FirstDate, FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        IF OpenBalance <> 0 THEN
            ExcelBuf.AddColumn(ABS(OpenBalance), FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text)
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        IF OpeningBalance1 <> 0 THEN
            ExcelBuf.AddColumn(ABS(OpeningBalance1), FALSE, '', FALSE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number)
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text); //RSPL-CAS-04786-L4N9T8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text); //RSPL-CAS-04786-L4N9T8
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        "Vendor Ledger Entry".CALCFIELDS("Debit Amount (LCY)");
        "Vendor Ledger Entry".CALCFIELDS("Credit Amount (LCY)");
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Vendor Ledger Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn("Vendor Ledger Entry"."Source Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Vendor Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(COPYSTR(Narration, 1, 248), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Vendor Ledger Entry"."Purchase (LCY)" <> 0 THEN
            ExcelBuf.AddColumn(ABS("Vendor Ledger Entry"."Purchase (LCY)"), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number) //RSPL-CAS-04786-L4N9T8
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //RSPL-CAS-04786-L4N9T8

        IF "Vendor Ledger Entry"."Debit Amount (LCY)" <> 0 THEN
            ExcelBuf.AddColumn("Vendor Ledger Entry"."Debit Amount (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number)
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Vendor Ledger Entry"."Credit Amount (LCY)" <> 0 THEN
            ExcelBuf.AddColumn("Vendor Ledger Entry"."Credit Amount (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number)
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Balance1, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        //RSPL-CAS-04786-L4N9T8
        IF vServiceTaxAmt <> 0 THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('Service Tax  Amount', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            IF vServiceTaxAmt <> 0 THEN
                ExcelBuf.AddColumn(vServiceTaxAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number)
            ELSE
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;
        IF "Vendor Ledger Entry"."Total TDS Including SHE CESS" <> 0 THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('TDS Amount', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            IF "Vendor Ledger Entry"."Total TDS Including SHE CESS" <> 0 THEN
                ExcelBuf.AddColumn(-("Vendor Ledger Entry"."Total TDS Including SHE CESS"), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number)
            ELSE
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.NewRow;
        END;

        //RSPL-CAS-04786-L4N9T8
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Cheque, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ChequeDatails, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(COPYSTR(Narration, 249, 500), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        //RSPL-CAS-04786-L4N9T8
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        "Vendor Ledger Entry".CALCFIELDS("Debit Amount (LCY)");
        "Vendor Ledger Entry".CALCFIELDS("Credit Amount (LCY)");
        DebitAmt += "Vendor Ledger Entry"."Debit Amount (LCY)";
        CreditAmt += "Vendor Ledger Entry"."Credit Amount (LCY)";
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);  //RSPL-CAS-04786-L4N9T8
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);  //RSPL-CAS-04786-L4N9T8
        //ExcelBuf.AddColumn(FORMAT(DebitAmt+ABS(OpenBalance))+'**Test1',FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn((DebitAmt + ABS(OpenBalance)), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn((CreditAmt + ABS(OpeningBalance1)), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        /*
        "Detailed Vendor Ledg. Entry".CALCSUMS("Debit Amount (LCY)");
        "Detailed Vendor Ledg. Entry".CALCSUMS("Credit Amount (LCY)");
        ExcelBuf.AddColumn(("Detailed Vendor Ledg. Entry"."Credit Amount (LCY)"+ABS(OpenBalance)),FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(("Detailed Vendor Ledg. Entry"."Debit Amount (LCY)"+ABS(OpeningBalance1)),FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        */
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);  //RSPL-CAS-04786-L4N9T8
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Closing Balance', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text); //RSPL-CAS-04786-L4N9T8
        IF ClosingBalance1 <> 0 THEN
            ExcelBuf.AddColumn(ABS(ClosingBalance1), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number)
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        IF ClosingBalance2 <> 0 THEN
            ExcelBuf.AddColumn(ABS(ClosingBalance2), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number)
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);  //RSPL-CAS-04786-L4N9T8

    end;
}


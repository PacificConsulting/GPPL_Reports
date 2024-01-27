report 50029 "Ledger New 2017--1"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 28Apr2017   RB-N          Changed AccountName[=iif(Fields!ControlAccountName.Value="",Fields!AccountName.Value,Fields!ControlAccountName.Value)]
    //                           --> Max(AccountName2)--for displaying as per NAV2009
    //                           Changed SourceType[=Fields!SourceDesc.Value] ---> "G/L Entry"."Global Dimension 2 Code" as per NAV2009
    // 
    // 29Apr2017   RB-N          AccountName29
    // 02May2017   RB-N          FIRST(DrCrTextBalance2)-->LAST(DrCrTextBalance3)
    // 25May2017   RB-N          GLNo--OLDVALUE(10)--->NEWVALUE(100)
    // 11July2017  RB-N          Report#16563 --saved as R#50029--Ledger New Report
    // 13Nov2018   RB-N          Displaying Division Code & Divsion Name
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/LedgerNew20171.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    Caption = 'Ledger New 2017';

    dataset
    {
        dataitem("G/L Account"; "G/L Account")
        {
            DataItemTableView = SORTING("No.")
                                ORDER(Ascending)
                                WHERE("Account Type" = FILTER(Posting));
            RequestFilterFields = "No.", "Date Filter", "Global Dimension 1 Filter", "Global Dimension 2 Filter";
            column(TodayFormatted; FORMAT(TODAY, 0, 4))
            {
            }
            column(CompInfoName; CompInfo.Name)
            {
            }
            column(LedgerName; Name + '  ' + 'Ledger')
            {
            }
            column(GetFilters; GETFILTERS)
            {
            }
            column(LocationFilter; LocationFilter)
            {
            }
            column(OpeningBalDesc; 'Opening Balance As On' + ' ' + FORMAT(GETRANGEMIN("Date Filter")))
            {
            }
            column(OpeningDRBal; OpeningDRBal)
            {
            }
            column(OpeningCRBal; OpeningCRBal)
            {
            }
            column(DRCRBal; ABS(OpeningDRBal - OpeningCRBal))
            {
            }
            column(DrCrTextBalance; DrCrTextBalance)
            {
            }
            column(OpeningCRBalGLEntryCreditAmt; OpeningCRBal + "G/L Entry"."Credit Amount")
            {
            }
            column(OpeningDRBalGLEntryDebitAmt; OpeningDRBal + "G/L Entry"."Debit Amount")
            {
            }
            column(SumOpeningDRCRBalTransDRCR; ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits))
            {
            }
            column(DrCrTextBalance2; DrCrTextBalance2)
            {
            }
            column(TotalDebitAmount; TotalDebitAmount)
            {
            }
            column(TotalCreditAmount; TotalCreditAmount)
            {
            }
            column(TransDebits; TransDebits)
            {
            }
            column(TransCredits; TransCredits)
            {
            }
            column(No_GLAccount; "No.")
            {
            }
            column(DateFilter_GLAccount; "Date Filter")
            {
            }
            column(GlobalDim1Filter_GLAccount; "Global Dimension 1 Filter")
            {
            }
            column(GlobalDim2Filter_GLAccount; "Global Dimension 2 Filter")
            {
            }
            column(PageNoCaption; PageCaptionLbl)
            {
            }
            column(PostingDateCaption; PostingDateCaptionLbl)
            {
            }
            column(DocumentNoCaption; DocumentNoCaptionLbl)
            {
            }
            column(Debit_AmountCaption; DebitAmountCaptionLbl)
            {
            }
            column(CreditAmountCaption; CreditAmountCaptionLbl)
            {
            }
            column(AccountNameCaption; AccountNameCaptionLbl)
            {
            }
            column(VoucherTypeCaption; VoucherTypeCaptionLbl)
            {
            }
            column(LocationCodeCaption; LocationCodeCaptionLbl)
            {
            }
            column(BalanceCaption; BalanceCaptionLbl)
            {
            }
            column(Closing_BalanceCaption; Closing_BalanceCaptionLbl)
            {
            }
            column(DrCrOpening13; DrCrOpening13)
            {
            }
            dataitem("G/L Entry"; "G/L Entry")
            {
                DataItemLink = "G/L Account No." = FIELD("No."),
                               "Posting Date" = FIELD("Date Filter"),
                               "Global Dimension 1 Code" = FIELD("Global Dimension 1 Filter"),
                               "Global Dimension 2 Code" = FIELD("Global Dimension 2 Filter");
                DataItemTableView = SORTING("Posting Date", "Document No.")
                                    ORDER(Ascending)
                                    WHERE(Reversed = FILTER(false));
                RequestFilterFields = "Document No.";
                column(ControlAccountName; ControlAccountName)
                {
                }
                column(PostingDate_GLEntry; FORMAT("Posting Date"))
                {
                }
                column(DocumentNo_GLEntry; "Document No.")
                {
                }
                column(AccountName; AccountName)
                {
                }
                column(AccountName29; AccountName29)
                {
                }
                column(DebitAmount_GLEntry; IntDrAmt)
                {
                }
                column(CreditAmoutnt_GLEntry; IntCrAmt)
                {
                }
                column(SourceDesc; SourceDesc)
                {
                }
                column(LocationCode_GLEntry; '"G/L Entry"."Location Code"')
                {
                }
                column(SumOpeningDRCRBalTransDRCR2; ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits))
                {
                }
                column(SumOpeningDRCRBalTransDRCR4; ABS(OpeningDRBal - OpeningCRBal))
                {
                }
                column(DrCrTextBalance3; DrCrTextBalance)
                {
                }
                column(OneEntryRecord; OneEntryRecord)
                {
                }
                column(EntryNo_GLEntry; "Entry No.")
                {
                }
                column(DimCode; "G/L Entry"."Global Dimension 2 Code")
                {
                }
                column(GLE_DebitAmt; "G/L Entry"."Debit Amount")
                {
                }
                column(GLE_CreditAmt; "G/L Entry"."Credit Amount")
                {
                }
                column(DebitAmt12; DebitAmt12)
                {
                }
                column(CreditAmt12; CreditAmt12)
                {
                }
                column(LineNarVis; LineNarVis)
                {
                }
                column(VouNarNis; VouNarNis)
                {
                }
                column(SalesComm; SalesComm)
                {
                }
                column(PurchComm; PurchComm)
                {
                }
                dataitem(GLDoc; "G/L Entry")
                {
                    DataItemLink = "Document No." = FIELD("Document No."),
                                   "G/L Account No." = FIELD("G/L Account No.");
                    DataItemTableView = SORTING("Posting Date", "Document No.")
                                        ORDER(Ascending)
                                        WHERE(Reversed = FILTER(false));
                    column(DrCrTextBalance3_13; DrCrTextBalance)
                    {
                    }
                    column(SumOpeningDRCRBalTransDRCR2_13; ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits))
                    {
                    }
                    column(SumOpeningDRCRBalTransDRCR4_13; ABS(OpeningDRBal - OpeningCRBal))
                    {
                    }
                    column(LocCode13; '"Location Code"')
                    {
                    }
                    column(DimCode13; "Global Dimension 2 Code")
                    {
                    }
                    column(GLE_DebitAmt13; "Debit Amount")
                    {
                    }
                    column(GLE_CreditAmt13; "Credit Amount")
                    {
                    }
                    column(DebitAmt13; DebitAmt12)
                    {
                    }
                    column(CreditAmt13; CreditAmt12)
                    {
                    }
                    column(EntryNo_GL13; "Entry No.")
                    {
                    }
                    column(PostingDate_GL13; FORMAT("Posting Date"))
                    {
                    }
                    column(DocumentNo_GL13; "Document No.")
                    {
                    }
                    column(AccountName13; AccountName)
                    {
                    }
                    column(AccountName29_13; AccountName2913)
                    {
                    }
                    column(PrintDetail_13; PrintDetail)
                    {
                    }
                    column(DetailVisible13; DetailVisible13)
                    {
                    }
                    column(TCSBaseAmount; DcTCSBaseAmt)
                    {
                    }
                    column(TDSBaseAmount; DcTDSBaseAmt)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        //MESSAGE('%1 no. of Records \\ %2 Doc No  \\ Entry No %3',COUNT,"Document No.","Entry No.");

                        //>>13July2017
                        GLAcc.GET("G/L Account No.");
                        ControlAccount := FindControlAccount("Source Type", "Entry No.", "Source No.", "G/L Account No.");
                        IF ControlAccount THEN
                            ControlAccountName := Daybook.FindGLAccName("Source Type", "Entry No.", "Source No.", "G/L Account No.");

                        IF Amount > 0 THEN
                            TransDebits := TransDebits + Amount;
                        IF Amount < 0 THEN
                            TransCredits := TransCredits - Amount;

                        SourceDesc := '';
                        IF "Source Code" <> '' THEN BEGIN
                            SourceCode.GET("Source Code");
                            SourceDesc := SourceCode.Description;
                        END;

                        //>>29Apr2017
                        AccountName2913 := '';
                        IF PrintDetail THEN BEGIN
                            IF "Source Type" = "Source Type"::Vendor THEN
                            //IF VendLedgEntry.GET("Entry No.") THEN
                            BEGIN
                                Vend.GET("Source No.");
                                AccountName2913 := Vend.Name;
                            END;

                            IF "Source Type" = "Source Type"::Customer THEN
                            //IF CustLedgEntry.GET("Entry No.") THEN
                            BEGIN
                                Cust.GET("Source No.");
                                AccountName2913 := Cust.Name;
                                TANNO := Cust.TANNO;
                            END;

                            IF "Source Type" = "Source Type"::"Bank Account" THEN
                            //IF BankLedgEntry.GET("Entry No.") THEN
                            BEGIN
                                Bank.GET("Source No.");
                                AccountName2913 := Bank.Name;
                            END;

                            IF "Source Type" = "Source Type"::"Fixed Asset" THEN BEGIN
                                FALedgEntry.RESET;
                                FALedgEntry.SETCURRENTKEY("G/L Entry No.");
                                FALedgEntry.SETRANGE("G/L Entry No.", "Entry No.");
                                IF FALedgEntry.FINDFIRST THEN BEGIN
                                    FA.GET("Source No.");
                                    AccountName2913 := FA.Description;
                                END;
                            END;

                            IF "Source Type" = "Source Type"::" " THEN BEGIN
                                GLEntry29.RESET;
                                GLEntry29.SETCURRENTKEY("Document No.", "Posting Date", Amount);//"Entry No."
                                GLEntry29.SETRANGE("Document No.", "Document No.");
                                GLEntry29.SETFILTER("Source Type", '<>%1', GLEntry29."Source Type"::" ");
                                IF GLEntry29.FINDFIRST THEN
                                //IF GLEntry29."Source Type" <> GLEntry29."Source Type" :: " "  THEN
                                BEGIN

                                    IF GLEntry29."Source Type" = GLEntry29."Source Type"::Vendor THEN BEGIN
                                        Vend.GET(GLEntry29."Source No.");
                                        AccountName2913 := Vend.Name;

                                    END;

                                    IF GLEntry29."Source Type" = GLEntry29."Source Type"::Customer THEN BEGIN

                                        Cust.GET(GLEntry29."Source No.");
                                        AccountName2913 := Cust.Name;
                                        TANNO := Cust.TANNO;
                                    END;

                                    //GLEntry29.SETFILTER("Source Type",
                                END ELSE BEGIN
                                    GLAccount.GET("G/L Account No.");
                                    AccountName2913 := GLAccount.Name;
                                END;
                            END;

                        END;


                        IF NOT PrintDetail THEN
                            AccountName2913 := Text16500;

                        //<<29Apr2017
                        AccountName := '';
                        IF OneEntryRecord AND (ControlAccountName <> '') THEN BEGIN
                            //AccountName := Daybook.FindGLAccName(GLEntry."Source Type",GLEntry."Entry No.",GLEntry."Source No.",GLEntry."G/L Account No.");
                            //Commented asper NAV2009 28Apr2017
                            DrCrTextBalance := '';
                            IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                                DrCrTextBalance := 'Dr';
                            IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                                DrCrTextBalance := 'Cr';
                        END ELSE
                            IF OneEntryRecord AND (ControlAccountName = '') THEN BEGIN
                                AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.");
                                DrCrTextBalance := '';
                                IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                                    DrCrTextBalance := 'Dr';
                                IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                                    DrCrTextBalance := 'Cr';
                            END;

                        IF GLAccountNo <> "G/L Account"."No." THEN
                            GLAccountNo := "G/L Account"."No.";

                        IF GLAccountNo = "G/L Account"."No." THEN BEGIN
                            DrCrTextBalance2 := '';
                            IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                                DrCrTextBalance2 := 'Dr';
                            IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                                DrCrTextBalance2 := 'Cr';
                        END;

                        TotalDebitAmount += "Debit Amount";
                        TotalCreditAmount += "Credit Amount";

                        //RSPLSUM 15Nov2019>>
                        CLEAR(DcTDSBaseAmt);
                        //IF "Vendor Ledger Entry"."Total TDS Including SHE CESS" <> 0 THEN BEGIN
                        // RecTDSEntry.RESET;
                        // //RecTDSEntry.SETRANGE("Document Type","Document Type");
                        // RecTDSEntry.SETRANGE("Document No.", "Document No.");
                        // //RecTDSEntry.SETRANGE("Party Code","Vendor Ledger Entry"."Vendor No.");
                        // IF RecTDSEntry.FINDFIRST THEN
                        //     REPEAT
                        //         DcTDSBaseAmt := RecTDSEntry."TDS Base Amount";
                        //     UNTIL RecTDSEntry.NEXT = 0;
                        //END;
                        //RSPLSUM 14Nov2019<<

                        //Nb Start
                        CLEAR(DcTCSBaseAmt);
                        // RecTcsEntry.RESET;
                        // RecTcsEntry.SETRANGE("Document No.", "Document No.");
                        // IF RecTcsEntry.FINDFIRST THEN
                        //     REPEAT
                        //         DcTCSBaseAmt := RecTcsEntry."TCS Base Amount";
                        //     UNTIL RecTcsEntry.NEXT = 0;
                        //Nb End



                        IF PrintDetail THEN
                            DetailVisible13 := TRUE;

                        //<<13July2017

                        //>>10Aug2017 DocumentBody
                        IF PrintToExcel AND PrintDetail THEN
                            MakeExcelDataBody4;
                        //<<10Aug2017 DocumentBody

                        //>>10Aug2017 GCrAmt
                        GDrAmt += "Debit Amount";
                        GCrAmt += "Credit Amount";
                        //<<10Aug2017
                    end;

                    trigger OnPostDataItem()
                    begin

                        //>>10Aug2017 DocumentFooter
                        IF PrintToExcel THEN
                            MakeExcelDataGroupFooter;
                        //<<10Aug2017 DocumentFooter
                    end;

                    trigger OnPreDataItem()
                    begin

                        SETRANGE(GLDoc."Posting Date", SDATE26, EDATE26);//13Jun2017

                        //>>06Oct2017
                        IF Dim1Filter <> '' THEN
                            SETRANGE(GLDoc."Global Dimension 1 Code", Dim1Filter);


                        IF Dim2Filter <> '' THEN
                            SETRANGE(GLDoc."Global Dimension 2 Code", Dim2Filter);
                        //>>06Oct2017
                    end;
                }
                dataitem(Integer; Integer)
                {
                    DataItemTableView = SORTING(Number);
                    column(ControlAccountName1; ControlAccountName)
                    {
                    }
                    column(PostingDate_GLEntry2; FORMAT(GLEntry."Posting Date"))
                    {
                    }
                    column(GLEntryDocumentNo; GLEntry."Document No.")
                    {
                    }
                    column(AccountName2; AccountName)
                    {
                    }
                    column(GLEntryDebitAmount; "G/L Entry"."Debit Amount")
                    {
                    }
                    column(GLEntryCreditAmount; "G/L Entry"."Credit Amount")
                    {
                    }
                    column(DetailAmt; ABS(DetailAmt))
                    {
                    }
                    column(SourceDesc2; SourceDesc)
                    {
                    }
                    column(DrCrText; DrCrText)
                    {
                    }
                    column(LocationCode_GLEntry2; '"G/L Entry"."Location Code"')
                    {
                    }
                    column(SumOpeningDRCRBalTransDRCR3; ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits))
                    {
                    }
                    column(DrCrTextBalance4; DrCrTextBalance)
                    {
                    }
                    column(FirstRecord; FirstRecord)
                    {
                    }
                    column(Amount_GLEntry2; ABS(GLEntry.Amount))
                    {
                    }
                    column(PrintDetail; PrintDetail)
                    {
                    }
                    column(DocNo_GLEntry2; GLEntry2."Document No.")
                    {
                    }
                    column(vTotCount; vTotCount)
                    {
                    }
                    column(vCount2; vCount2)
                    {
                    }
                    column(IntCrAmt; IntCrAmt)
                    {
                    }
                    column(IntDrAmt; IntDrAmt)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        DrCrText := '';
                        IF Number > 1 THEN BEGIN
                            FirstRecord := FALSE;
                            GLEntry.NEXT;
                        END;
                        IF FirstRecord AND (ControlAccountName <> '') THEN BEGIN
                            DetailAmt := 0;
                            IF PrintDetail THEN
                                DetailAmt := GLEntry.Amount;   //RSPL-Rahul
                            IF DetailAmt > 0 THEN
                                DrCrText := 'Dr';
                            IF DetailAmt < 0 THEN
                                DrCrText := 'Cr';

                            IF NOT PrintDetail THEN
                                AccountName := Text16500
                            ELSE
                                AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.")
                            ;

                            DrCrTextBalance := '';
                            IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                                DrCrTextBalance := 'Dr';
                            IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                                DrCrTextBalance := 'Cr';
                        END ELSE
                            IF FirstRecord AND (ControlAccountName = '') THEN BEGIN
                                DetailAmt := 0;
                                IF PrintDetail THEN
                                    DetailAmt := GLEntry.Amount;

                                IF DetailAmt > 0 THEN
                                    DrCrText := 'Dr';
                                IF DetailAmt < 0 THEN
                                    DrCrText := 'Cr';

                                IF NOT PrintDetail THEN
                                    AccountName := Text16500
                                ELSE
                                    AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.")
                                ;

                                DrCrTextBalance := '';
                                IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                                    DrCrTextBalance := 'Dr';
                                IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                                    DrCrTextBalance := 'Cr';
                            END;

                        IF PrintDetail AND (NOT FirstRecord) THEN BEGIN
                            IF GLEntry.Amount > 0 THEN
                                DrCrText := 'Dr';
                            IF GLEntry.Amount < 0 THEN
                                DrCrText := 'Cr';
                            AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.");
                        END;
                    end;

                    trigger OnPreDataItem()
                    begin
                        SETRANGE(Number, 1, GLEntry.COUNT);
                        FirstRecord := TRUE;

                        IF GLEntry.COUNT = 1 THEN
                            CurrReport.BREAK;
                        //>>Robosoft/Rahul**Code added to group document no
                        vTotCount := GLEntry.COUNT;
                        IF GLEnryDocNo <> GLEntry."Document No." THEN
                            vCount2 := 0;
                        GLEnryDocNo := GLEntry."Document No.";
                        vCount2 += 1;
                        /*
                        GLEntry.CALCSUMS("Credit Amount");
                        GLEntry.CALCSUMS("Debit Amount");
                        IntCrAmt := GLEntry."Credit Amount";
                        
                        IntDrAmt := GLEntry."Debit Amount";
                        */
                        //<<

                    end;
                }
                dataitem("Posted Narration"; "Posted Narration")
                {
                    DataItemLink = "Entry No." = FIELD("Entry No.");
                    DataItemTableView = SORTING("Entry No.", "Transaction No.", "Line No.")
                                        ORDER(Ascending);
                    column(Narration_PostedNarration; Narration)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin
                        //>>12July2017
                        //CLEAR(LineNarVis);
                        //LineNarVis := TRUE;

                        //>>12July2017
                    end;

                    trigger OnPreDataItem()
                    begin
                        IF NOT PrintLineNarration THEN
                            CurrReport.BREAK;
                    end;
                }
                dataitem(PostedNarration1; "Posted Narration")
                {
                    DataItemLink = "Transaction No." = FIELD("Transaction No.");
                    DataItemTableView = SORTING("Entry No.", "Transaction No.", "Line No.")
                                        WHERE("Entry No." = FILTER(0));
                    column(Narration_PostedNarration1; Narration)
                    {
                    }
                    column(EntryNo_PostedNarration1; PostedNarration1."Entry No.")
                    {
                    }
                    column(TransactionNo_PostedNarration1; PostedNarration1."Transaction No.")
                    {
                    }
                    column(LineNo_PostedNarration1; PostedNarration1."Line No.")
                    {
                    }
                    column(PostingDate_PostedNarration1; PostedNarration1."Posting Date")
                    {
                    }
                    column(DocumentType_PostedNarration1; PostedNarration1."Document Type")
                    {
                    }
                    column(DocumentNo_PostedNarration1; PostedNarration1."Document No.")
                    {
                    }
                    column(VoucherNarration; Narration)
                    {
                    }
                    column(NarationVoucherNo; NarationVoucherNo)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        //>>12July2017
                        NarationVoucherNo += 1;

                        //CLEAR(VouNarNis);
                        //VouNarNis := TRUE;

                        //IF VouNarNis THEN
                        //MESSAGE(' %1 Voucher No \\ Narration %2',PostedNarration1."Document No.",PostedNarration1.Narration);

                        //>>12July2017
                    end;

                    trigger OnPreDataItem()
                    begin
                        IF NOT PrintVchNarration THEN
                            CurrReport.BREAK;

                        GLEntry2.RESET;
                        GLEntry2.SETCURRENTKEY("Posting Date", "Source Code", "Transaction No.");
                        GLEntry2.SETRANGE("Posting Date", "G/L Entry"."Posting Date");
                        GLEntry2.SETRANGE("Source Code", "G/L Entry"."Source Code");
                        GLEntry2.SETRANGE("Transaction No.", "G/L Entry"."Transaction No.");
                        GLEntry2.FINDLAST;
                        IF NOT (GLEntry2."Entry No." = "G/L Entry"."Entry No.") THEN
                            //CurrReport.BREAK;  //Migration


                            NarationVoucherNo := 0; //12July2017
                    end;
                }
                dataitem("Purch. Comment Line"; "Purch. Comment Line")
                {
                    DataItemLink = "No." = FIELD("Document No.");
                    DataItemTableView = SORTING("Document Type", "No.", "Line No.");
                    column(No_PurchCommentLine; "Purch. Comment Line"."No.")
                    {
                    }
                    column(LineNo_PurchCommentLine; "Purch. Comment Line"."Line No.")
                    {
                    }
                    column(Comment_PurchCommentLine; "Purch. Comment Line".Comment)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        IF NOT ("G/L Entry"."Source Code" = 'PURCHASES') THEN
                            CurrReport.SKIP;
                    end;
                }
                dataitem("Sales Comment Line"; "Sales Comment Line")
                {
                    DataItemLink = "No." = FIELD("Document No.");
                    DataItemTableView = SORTING("Document Type", "No.", "Document Line No.", "Line No.")
                                        WHERE("Document Type" = FILTER("Posted Credit Memo" | "Posted Invoice"));
                    column(No_SalesCommentLine; "Sales Comment Line"."No.")
                    {
                    }
                    column(LineNo_SalesCommentLine; "Sales Comment Line"."Line No.")
                    {
                    }
                    column(Comment_SalesCommentLine; "Sales Comment Line".Comment)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        IF NOT ("G/L Entry"."Source Code" = 'SALES') THEN
                            CurrReport.SKIP;
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    CLEAR(TotalNarration);//02May2017

                    //>>11July2017 Skip ExchangeRate Entry

                    //IF "Source Code" = 'EXCHRATADJ' THEN
                    //CurrReport.SKIP;

                    //<<11July2017 Skip ExchangeRate Entry

                    GLEntry.SETRANGE("Transaction No.", "Transaction No.");
                    GLEntry.SETFILTER("Entry No.", '<>%1', "Entry No.");
                    IF GLEntry.FINDFIRST THEN;

                    DrCrText := '';
                    ControlAccount := FALSE;
                    OneEntryRecord := TRUE;
                    IF GLEntry.COUNT > 1 THEN
                        OneEntryRecord := FALSE;

                    GLAcc.GET("G/L Account No.");
                    ControlAccount := FindControlAccount("Source Type", "Entry No.", "Source No.", "G/L Account No.");
                    IF ControlAccount THEN
                        ControlAccountName := Daybook.FindGLAccName("Source Type", "Entry No.", "Source No.", "G/L Account No.");
                    /*
                    IF Amount>0 THEN
                     TransDebits := TransDebits + Amount;
                    IF Amount<0 THEN
                     TransCredits := TransCredits - Amount;
                    */
                    SourceDesc := '';
                    IF "Source Code" <> '' THEN BEGIN
                        SourceCode.GET("Source Code");
                        SourceDesc := SourceCode.Description;
                    END;

                    //>>29Apr2017
                    AccountName29 := '';
                    //IF PrintDetail THEN
                    //BEGIN
                    IF "Source Type" = "Source Type"::Vendor THEN
                    //IF VendLedgEntry.GET("Entry No.") THEN
                    BEGIN
                        Vend.GET("Source No.");
                        AccountName29 := Vend.Name;
                    END;

                    IF "Source Type" = "Source Type"::Customer THEN
                    //IF CustLedgEntry.GET("Entry No.") THEN
                    BEGIN
                        Cust.GET("Source No.");
                        AccountName29 := Cust.Name;
                        TANNO := Cust.TANNO;
                    END;

                    IF "Source Type" = "Source Type"::"Bank Account" THEN
                    //IF BankLedgEntry.GET("Entry No.") THEN
                    BEGIN
                        Bank.GET("Source No.");
                        AccountName29 := Bank.Name;
                    END;

                    IF "Source Type" = "Source Type"::"Fixed Asset" THEN BEGIN
                        FALedgEntry.RESET;
                        FALedgEntry.SETCURRENTKEY("G/L Entry No.");
                        FALedgEntry.SETRANGE("G/L Entry No.", "Entry No.");
                        IF FALedgEntry.FINDFIRST THEN BEGIN
                            FA.GET("Source No.");
                            AccountName29 := FA.Description;
                        END;
                    END;

                    IF "Source Type" = "Source Type"::" " THEN BEGIN
                        GLEntry29.RESET;
                        GLEntry29.SETCURRENTKEY("Document No.", "Posting Date", Amount);//"Entry No."
                        GLEntry29.SETRANGE("Document No.", "Document No.");
                        GLEntry29.SETFILTER("Source Type", '<>%1', GLEntry29."Source Type"::" ");
                        IF GLEntry29.FINDFIRST THEN
                        //IF GLEntry29."Source Type" <> GLEntry29."Source Type" :: " "  THEN
                        BEGIN

                            IF GLEntry29."Source Type" = GLEntry29."Source Type"::Vendor THEN BEGIN
                                Vend.GET(GLEntry29."Source No.");
                                AccountName29 := Vend.Name;

                            END;

                            IF GLEntry29."Source Type" = GLEntry29."Source Type"::Customer THEN BEGIN

                                Cust.GET(GLEntry29."Source No.");
                                AccountName29 := Cust.Name;
                                TANNO := Cust.TANNO;
                            END;

                            //GLEntry29.SETFILTER("Source Type",
                        END ELSE BEGIN
                            GLAccount.GET("G/L Account No.");
                            AccountName29 := GLAccount.Name;
                        END;
                    END;

                    //END;

                    /*
                    IF NOT PrintDetail THEN
                        AccountName29:=Text16500;
                    *///Commented on 13July2017

                    //<<29Apr2017
                    AccountName := '';
                    IF OneEntryRecord AND (ControlAccountName <> '') THEN BEGIN
                        //AccountName := Daybook.FindGLAccName(GLEntry."Source Type",GLEntry."Entry No.",GLEntry."Source No.",GLEntry."G/L Account No.");
                        //Commented asper NAV2009 28Apr2017
                        DrCrTextBalance := '';
                        IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                            DrCrTextBalance := 'Dr';
                        IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                            DrCrTextBalance := 'Cr';
                    END ELSE
                        IF OneEntryRecord AND (ControlAccountName = '') THEN BEGIN
                            AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.");
                            DrCrTextBalance := '';
                            IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                                DrCrTextBalance := 'Dr';
                            IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                                DrCrTextBalance := 'Cr';
                        END;

                    IF GLAccountNo <> "G/L Account"."No." THEN
                        GLAccountNo := "G/L Account"."No.";

                    IF GLAccountNo = "G/L Account"."No." THEN BEGIN
                        DrCrTextBalance2 := '';
                        IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                            DrCrTextBalance2 := 'Dr';
                        IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                            DrCrTextBalance2 := 'Cr';
                    END;
                    /*
                    TotalDebitAmount += "G/L Entry"."Debit Amount";
                    TotalCreditAmount += "G/L Entry"."Credit Amount";
                    */

                    //>>13July2017 Testing
                    DocInserted := FALSE;
                    IF DocNo13 <> "Document No." THEN BEGIN
                        DocNo13 := "Document No.";
                        //MESSAGE('Doc no %1 ',DocNo13);


                        TempGL.SETRANGE("Cell Value as Text", "Document No.");
                        IF NOT TempGL.FINDFIRST THEN BEGIN
                            i13 += 1;
                            TempGL.INIT;
                            TempGL."Row No." := i13;
                            TempGL."Column No." := 1;
                            TempGL."Cell Value as Text" := "Document No.";
                            TempGL.INSERT;
                            //MESSAGE('%1 Doc No Inserted ',TempGL."Document No.");
                            DocInserted := TRUE;

                        END;

                        IF NOT DocInserted THEN BEGIN
                            TempGL.SETRANGE("Cell Value as Text", "Document No.");
                            IF TempGL.FINDFIRST THEN BEGIN

                                //MESSAGE('%1 Doc No Skip  ',DocNo13);
                                CurrReport.SKIP;
                            END;
                        END;

                    END;

                    IF NOT DocInserted THEN
                        CurrReport.SKIP;

                    //>>13July2017 Testing

                    //>>
                    IF PrintVchNarration THEN BEGIN
                        PrintNarration;
                        PrintComment;
                    END;

                    //>>Grp
                    CLEAR(IntCrAmt);
                    CLEAR(IntDrAmt);
                    vGLEntry.RESET;
                    vGLEntry.SETCURRENTKEY("G/L Account No.", "Posting Date");
                    vGLEntry.SETRANGE("Document No.", "Document No.");
                    //vGLEntry.SETFILTER("Location Code", LocationCode);
                    vGLEntry.SETRANGE(Reversed, FALSE);
                    vGLEntry.SETRANGE("Posting Date", "G/L Account".GETRANGEMIN("Date Filter"), "G/L Account".GETRANGEMAX("Date Filter"));
                    vGLEntry.SETRANGE("G/L Account No.", "G/L Account No.");
                    IF vGLEntry.FINDFIRST THEN BEGIN //13July2017
                                                     //IF vGLEntry.FINDSET THEN BEGIN
                        vCount := vGLEntry.COUNT;
                        vGLEntry.CALCSUMS("Credit Amount");
                        vGLEntry.CALCSUMS("Debit Amount");
                        vGLEntry.CALCSUMS(vGLEntry.Amount);
                        //IF vGLEntry.Amount <0 THEN
                        IntCrAmt := vGLEntry."Credit Amount";
                        //ELSE
                        IntDrAmt := vGLEntry."Debit Amount";
                        //MESSAGE(FORMAT(vGLEntry."Credit Amount"));
                    END;

                    /*
                   //>>02May2017 GCrAmt
                   GDrAmt += "G/L Entry"."Debit Amount";
                   GCrAmt += "G/L Entry"."Credit Amount";
                   //<<02May2017
                   */

                    vLoop += 1;
                    IF PrintToExcel THEN
                        IF vLoop = vCount THEN BEGIN
                            vLoop := 0;
                            // IntCrAmt :=0;
                            // IntDrAmt :=0;
                            //MakeExcelDataGroupFooter;//07Sep2017

                        END;

                    //GLEnryDocNo := "G/L Entry"."Document No.";
                    //<<

                    //MESSAGE('Entry No. %1 \\ Amount No. %2 \\ Date %3 \\ Docno %4 \\ Count %5',"Entry No.",Amount,"Posting Date","Document No.",COUNT);

                    DebitAmt12 += "G/L Entry"."Debit Amount"; //12July2017
                    CreditAmt12 += "G/L Entry"."Credit Amount";//12July2017

                    //>>12July2017 CommentVisiblity

                    CLEAR(LineNarVis);
                    CLEAR(VouNarNis);
                    CLEAR(SalesComm);
                    CLEAR(PurchComm);

                    PosNarration.RESET;
                    PosNarration.SETRANGE("Entry No.", "Entry No.");
                    IF PosNarration.FINDFIRST THEN
                        LineNarVis := TRUE;

                    PosNarration.RESET;
                    PosNarration.SETRANGE("Transaction No.", "Transaction No.");
                    PosNarration.SETRANGE("Entry No.", 0);
                    IF PosNarration.FINDFIRST THEN
                        VouNarNis := TRUE;

                    SCom12.RESET;
                    SCom12.SETRANGE("No.", "G/L Entry"."Document No.");
                    IF SCom12.FINDFIRST THEN
                        SalesComm := TRUE;

                    PCom12.RESET;
                    PCom12.SETRANGE("No.", GLEntry."Document No.");
                    IF PCom12.FINDFIRST THEN
                        PurchComm := TRUE;

                    //<<12July2017 CommentVisiblity

                end;

                trigger OnPostDataItem()
                begin
                    AccountChanged := TRUE;

                    IF PrintToExcel THEN
                        MakeExcelDataFooter;
                end;

                trigger OnPreDataItem()
                begin
                    // IF LocationCode <> '' THEN
                    //     SETFILTER("Location Code", LocationCode);
                    GLEntry.RESET;
                    GLEntry.SETCURRENTKEY("Transaction No.");
                    TotalDebitAmount := 0;
                    TotalCreditAmount := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin
                IF AccountNo <> "No." THEN BEGIN
                    AccountNo := "No.";
                    OpeningDRBal := 0;
                    OpeningCRBal := 0;

                    GLEntry2.RESET;
                    GLEntry2.SETCURRENTKEY("G/L Account No.", "Business Unit Code", "Global Dimension 1 Code",
                    "Global Dimension 2 Code", "Close Income Statement Dim. ID", "Posting Date");//, "Location Code");
                    GLEntry2.SETRANGE("G/L Account No.", "No.");
                    GLEntry2.SETFILTER("Posting Date", '%1..%2', 0D, CLOSINGDATE(GETRANGEMIN("Date Filter") - 1));
                    IF "Global Dimension 1 Filter" <> '' THEN
                        GLEntry2.SETFILTER("Global Dimension 1 Code", "Global Dimension 1 Filter");
                    IF "Global Dimension 2 Filter" <> '' THEN
                        GLEntry2.SETFILTER("Global Dimension 2 Code", "Global Dimension 2 Filter");
                    // IF LocationCode <> '' THEN
                    //     GLEntry2.SETFILTER("Location Code", LocationCode);

                    GLEntry2.CALCSUMS(Amount);
                    IF GLEntry2.Amount > 0 THEN
                        OpeningDRBal := GLEntry2.Amount;
                    IF GLEntry2.Amount < 0 THEN
                        OpeningCRBal := -GLEntry2.Amount;

                    DrCrTextBalance := '';
                    IF OpeningDRBal - OpeningCRBal > 0 THEN
                        DrCrOpening13 := 'Dr';//13Jul2017
                                              //DrCrTextBalance := 'Dr';
                    IF OpeningDRBal - OpeningCRBal < 0 THEN
                        DrCrOpening13 := 'Cr';//13Jul2017
                    //DrCrTextBalance := 'Cr';
                END;

                //>>26May2017 Skipping GLAccount if no Entry & If No Opening Balance Amount
                GLE26.RESET;
                GLE26.SETCURRENTKEY("G/L Account No.", "Posting Date");
                GLE26.SETRANGE("G/L Account No.", "No.");
                GLE26.SETRANGE("Posting Date", SDATE26, EDATE26);

                IF "Global Dimension 1 Filter" <> '' THEN
                    GLE26.SETFILTER("Global Dimension 1 Code", "Global Dimension 1 Filter");

                IF "Global Dimension 2 Filter" <> '' THEN
                    GLE26.SETFILTER("Global Dimension 2 Code", "Global Dimension 2 Filter");

                // IF LocationCode <> '' THEN
                //     GLE26.SETFILTER("Location Code", LocationCode);

                IF (OpeningCRBal = 0) AND (OpeningDRBal = 0) AND (GLE26.COUNT = 0) THEN
                    //IF (COUNT = 0) AND (OpeningCRBal = 0) AND (OpeningDRBal = 0) THEN
                    CurrReport.SKIP;

                //>>26May2017 Skipping GLAccount if no Entry & If No Opening Balance Amount

                //>>
                IF PrintToExcel THEN
                    MakeExcelDataGroupHeader;
                //<<
            end;

            trigger OnPreDataItem()
            begin
                CurrReport.CREATETOTALS(TransDebits, TransCredits, "G/L Entry"."Debit Amount", "G/L Entry"."Credit Amount");
                IF LocationCode <> '' THEN
                    LocationFilter := 'Location Code: ' + LocationCode;
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
                    field(PrintDetail; PrintDetail)
                    {
                        Caption = 'Print Detail';
                    }
                    field(PrintLineNarration; PrintLineNarration)
                    {
                        Caption = 'Print Line Narration';
                    }
                    field(PrintVchNarration; PrintVchNarration)
                    {
                        Caption = 'Print Voucher Narration';
                    }
                    field(LocationCode; LocationCode)
                    {
                        Caption = 'Location Code';
                        TableRelation = Location;
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
    }

    labels
    {
    }

    trigger OnPostReport()
    begin

        TempGL.DELETEALL;//13Jun2017

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin
        //CompInfo.GET;

        TempGL.DELETEALL;//13Jun2017

        //>>26May2017

        SDATE26 := "G/L Account".GETRANGEMIN("G/L Account"."Date Filter");
        EDATE26 := "G/L Account".GETRANGEMAX("G/L Account"."Date Filter");
        //<<26May2017

        //>>06Oct2017
        Dim1Filter := "G/L Account".GETFILTER("G/L Account"."Global Dimension 1 Filter");
        Dim2Filter := "G/L Account".GETFILTER("G/L Account"."Global Dimension 2 Filter");

        //<<06Oct2017

        CompInfo.GET;
        GLNo := "G/L Account".GETFILTER("G/L Account"."No.");
        datefilt := FORMAT("G/L Account".GETFILTER("G/L Account"."Date Filter"));
        divifilt := "G/L Account".GETFILTER("G/L Account"."Global Dimension 1 Filter");
        Respfilt := "G/L Account".GETFILTER("G/L Account"."Global Dimension 2 Filter");

        GLNAME := '';
        "G/L Account".RESET;
        "G/L Account".SETRANGE("G/L Account"."No.", GLNo);
        IF "G/L Account".FINDFIRST THEN BEGIN
            GLNAME := "G/L Account".Name;
        END;
        IF PrintToExcel THEN
            MakeExcelInfo;
    end;

    var
        CompInfo: Record 79;
        GLAcc: Record 15;
        GLEntry: Record 17;
        GLEntry2: Record 17;
        Daybook: Report "Day Book";
        SourceCode: Record 230;
        OpeningDRBal: Decimal;
        OpeningCRBal: Decimal;
        TransDebits: Decimal;
        TransCredits: Decimal;
        OneEntryRecord: Boolean;
        FirstRecord: Boolean;
        PrintDetail: Boolean;
        PrintLineNarration: Boolean;
        PrintVchNarration: Boolean;
        DetailAmt: Decimal;
        AccountName: Text[100];
        ControlAccountName: Text[100];
        ControlAccount: Boolean;
        SourceDesc: Text[50];
        DrCrText: Text[2];
        DrCrTextBalance: Text[2];
        LocationCode: Code[20];
        LocationFilter: Text[100];
        Text16500: Label 'As per Details';
        AccountChanged: Boolean;
        AccountNo: Code[20];
        DrCrTextBalance2: Text[2];
        GLAccountNo: Code[20];
        TotalDebitAmount: Decimal;
        TotalCreditAmount: Decimal;
        PageCaptionLbl: Label 'Page';
        PostingDateCaptionLbl: Label 'Posting Date';
        DocumentNoCaptionLbl: Label 'Document No.';
        DebitAmountCaptionLbl: Label 'Debit Amount';
        CreditAmountCaptionLbl: Label 'Credit Amount';
        AccountNameCaptionLbl: Label 'Account Name';
        VoucherTypeCaptionLbl: Label 'Voucher Type';
        LocationCodeCaptionLbl: Label 'Location Code';
        BalanceCaptionLbl: Label 'Balance';
        Closing_BalanceCaptionLbl: Label 'Closing Balance';
        "---RSPL--Excel---": Integer;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        Text0000: Label 'Data';
        Text0001: Label 'Ledger';
        Text0002: Label 'Company Name';
        Text0003: Label 'Report No.';
        Text0004: Label 'Report Name';
        Text0005: Label 'USER Id';
        Text0006: Label 'Date';
        GLNAME: Text[60];
        GLNo: Code[100];
        datefilt: Text[30];
        divifilt: Code[10];
        Respfilt: Code[500];
        OpeningBalance: Text[50];
        VendInvNo: Code[20];
        PurchInvHeader: Record 122;
        TotalNarration: Text[1024];
        recPostedNarration: Record "Posted Narration";
        TotalComment: Text[500];
        recPurchCommentLine: Record 43;
        recSalesCommentLine: Record 44;
        recLoc: Record 14;
        recResCentre: Record 5714;
        recSIH: Record 112;
        i: Integer;
        GLEnryDocNo: Code[30];
        vGLEntry: Record 17;
        vCount: Integer;
        vLoop: Integer;
        vTotCount: Integer;
        vCount2: Integer;
        IntCrAmt: Decimal;
        IntDrAmt: Decimal;
        "----29Apr2017": Integer;
        AccountName29: Text[100];
        VendLedgEntry: Record 25;
        Vend: Record 23;
        CustLedgEntry: Record 21;
        Cust: Record 18;
        BankLedgEntry: Record 271;
        Bank: Record 270;
        FALedgEntry: Record 5601;
        FA: Record 5600;
        GLAccount: Record 15;
        GLEntry29: Record 17;
        "--02May2017": Integer;
        GCrAmt: Decimal;
        GDrAmt: Decimal;
        "---26May2017": Integer;
        GLE26: Record 17;
        SDATE26: Date;
        EDATE26: Date;
        "--12July2017": Integer;
        LineNarVis: Boolean;
        VouNarNis: Boolean;
        NarationLineNo: Integer;
        NarationVoucherNo: Integer;
        SalesComm: Boolean;
        PurchComm: Boolean;
        SCom12: Record 44;
        PCom12: Record 43;
        PosNarration: Record "Posted Narration";
        DebitAmt12: Decimal;
        CreditAmt12: Decimal;
        "--13July2017": Integer;
        DocNo13: Code[20];
        i13: Integer;
        DocNo13seq: array[99999] of Code[20];
        TempGL: Record 370 temporary;
        DocInserted: Boolean;
        DrCrOpening13: Text[2];
        DetailVisible13: Boolean;
        AccountName2913: Text[100];
        "-------06Oct2017": Integer;
        Dim1Filter: Text[1024];
        Dim2Filter: Text[1024];
        DimVal13: Record 349;
        DcTDSBaseAmt: Decimal;
        //RecTDSEntry: Record 13729;
        //RecTcsEntry: Record 16514;
        DcTCSBaseAmt: Decimal;
        TANNO: Text[20];

    //  //[Scope('Internal')]
    procedure FindControlAccount("Source Type": Option " ",Customer,Vendor,"Bank Account","Fixed Asset"; "Entry No.": Integer; "Source No.": Code[20]; "G/L Account No.": Code[20]): Boolean
    var
        VendLedgEntry: Record 25;
        CustLedgEntry: Record 21;
        BankLedgEntry: Record 271;
        FALedgEntry: Record 5601;
        SubLedgerFound: Boolean;
    begin
        IF "Source Type" = "Source Type"::Vendor THEN BEGIN
            IF VendLedgEntry.GET("Entry No.") THEN
                SubLedgerFound := TRUE;
        END;
        IF "Source Type" = "Source Type"::Customer THEN BEGIN
            IF CustLedgEntry.GET("Entry No.") THEN
                SubLedgerFound := TRUE;
        END;
        IF "Source Type" = "Source Type"::"Bank Account" THEN
            IF BankLedgEntry.GET("Entry No.") THEN BEGIN
                SubLedgerFound := TRUE;
            END;
        IF "Source Type" = "Source Type"::"Fixed Asset" THEN BEGIN
            FALedgEntry.RESET;
            FALedgEntry.SETCURRENTKEY("G/L Entry No.");
            FALedgEntry.SETRANGE("G/L Entry No.", "Entry No.");
            IF FALedgEntry.FINDFIRST THEN
                SubLedgerFound := TRUE;
        END;
        EXIT(SubLedgerFound);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text0002), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Ledger New ', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0003), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(50029, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin

        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text0000,Text0001,COMPANYNAME,USERID);
        // ExcelBuf.CreateBookAndOpenExcel('', 'Ledger Details', '', '', '');
        // ExcelBuf.GiveUserControl;
        // ERROR('');

        ExcelBuf.CreateNewBook('Ledger Details');
        ExcelBuf.WriteSheet('Ledger Details', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Ledger Details', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Account No. :' + '  ' + GLNo + ' - ' + GLNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date :' + '  ' + datefilt, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Division :' + '  ' + divifilt, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Resp.Center :' + '  ' + Respfilt, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(LocationCode, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 15Nov2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 15Nov2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);// Fahim 23-Mar-2022 for TAN No.
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Account Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Dr/Cr', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Voucher Type', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Division Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('Division Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('Res. Centre Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Res. Centre Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Debit Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Credit Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Balance', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Narration', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('TDS Base Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 15Nov2019
        ExcelBuf.AddColumn('TCS Base Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//NB
        ExcelBuf.AddColumn('TAN No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//Fahim 23-Mar-22
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 15Nov2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 15Nov2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 23-Mar-22
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataGroupHeader()
    begin
        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);
        ExcelBuf.AddColumn("G/L Account"."No.", FALSE, '', TRUE, FALSE, TRUE, '', 1);//16Apr2019
        ExcelBuf.AddColumn('Opening Balance', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(DrCrOpening13, FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Nov2018
        ExcelBuf.AddColumn(OpeningDRBal, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(OpeningCRBal, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 15Nov2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLnb
    end;

    ////[Scope('Internal')]
    procedure MakeExcelDataBody1()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("G/L Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(VendInvNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(AccountName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(SourceDesc, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Entry"."Debit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("G/L Entry"."Credit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalNarration, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10Aug2017
        //ExcelBuf.AddColumn(TotalNarration[2],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[3],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[4],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[5],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody2()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("G/L Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(VendInvNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(AccountName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(SourceDesc, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Entry"."Debit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("G/L Entry"."Credit Amount", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalNarration, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10Aug2017
        //ExcelBuf.AddColumn(TotalNarration[2],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[3],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[4],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[5],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody3()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(AccountName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn((GLEntry.Amount), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(DrCrText, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataBody4()
    begin

        /*
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("G/L Entry"."Posting Date",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn("G/L Entry"."Document No.",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn(VendInvNo,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn(ControlAccountName,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//7
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//8
        ExcelBuf.AddColumn(SourceDesc,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//9
        ExcelBuf.AddColumn("G/L Entry"."Debit Amount",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);//10
        ExcelBuf.AddColumn("G/L Entry"."Credit Amount",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);//11
        ExcelBuf.AddColumn(ABS(OpeningDRBal-OpeningCRBal+TransDebits-TransCredits),FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//12
        ExcelBuf.AddColumn(TotalNarration,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//10Aug2017 //13
        //ExcelBuf.AddColumn(TotalNarration[2],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[3],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[4],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalNarration[5],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(AccountName,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn((GLEntry.Amount),FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(DrCrText,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        */

        //>>10Aug2017 DataBody
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(GLDoc."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(GLDoc."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3 VendorInvoiceNo
        ExcelBuf.AddColumn(AccountName2913, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn(GLDoc.Amount, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//5 Amount
        ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6 Dr/Cr
        ExcelBuf.AddColumn(SourceDesc, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7 //VoucherType
        ExcelBuf.AddColumn('GLDoc."Location Code"', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8 //LocationCode

        recLoc.RESET;
        //recLoc.SETRANGE(recLoc.Code, GLDoc."Location Code");
        IF recLoc.FINDFIRST THEN BEGIN
            ExcelBuf.AddColumn(recLoc.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//9 LocationName
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//9 LocationName

        //>>13Nov2018
        ExcelBuf.AddColumn(GLDoc."Global Dimension 1 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        DimVal13.RESET;
        DimVal13.SETRANGE("Dimension Code", 'DIVISION');
        DimVal13.SETRANGE(Code, GLDoc."Global Dimension 1 Code");
        IF DimVal13.FINDLAST THEN
            ExcelBuf.AddColumn(DimVal13.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1)
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
        //<<13Nov2018

        ExcelBuf.AddColumn(GLDoc."Global Dimension 2 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 Responsibility Center

        recResCentre.RESET;
        recResCentre.SETRANGE(recResCentre.Code, GLDoc."Global Dimension 2 Code");
        IF recResCentre.FINDFIRST THEN BEGIN
            ExcelBuf.AddColumn(recResCentre.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11 Responsibility CenterName
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//11 Responsibility CenterName

        ExcelBuf.AddColumn(GLDoc."Debit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//12 Debit Amount
        ExcelBuf.AddColumn(GLDoc."Credit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//13 Credit Amount
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//14 Balance
        ExcelBuf.AddColumn(COPYSTR(TotalNarration, 1, 250), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//15
        ExcelBuf.AddColumn(DcTDSBaseAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//16 //RSPLSUM 15Nov2019
        ExcelBuf.AddColumn(DcTCSBaseAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//16 //rspl Nb
        ExcelBuf.AddColumn(TANNO, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);// Fahim 23-Mar-22
        //<<10Aug2017 DataBody

    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody5()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("G/L Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ControlAccountName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(SourceDesc, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Entry"."Debit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("G/L Entry"."Credit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
    end;

    ////[Scope('Internal')]
    procedure MakeExcelDataGroupFooter()
    begin
        /*//RSPLSUM
        IF PrintDetail THEN
        BEGIN
        
          ExcelBuf.NewRow;
          ExcelBuf.AddColumn("G/L Entry"."Posting Date",FALSE,'',TRUE,FALSE,FALSE,'',2);//1
          ExcelBuf.AddColumn("G/L Entry"."Document No.",FALSE,'',TRUE,FALSE,FALSE,'',1);//2
          ExcelBuf.AddColumn(VendInvNo,FALSE,'',TRUE,FALSE,FALSE,'',1);//3
          ExcelBuf.AddColumn(AccountName29,FALSE,'',TRUE,FALSE,FALSE,'',1);//17May2017//4
          //ExcelBuf.AddColumn(AccountName,FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//5
          ExcelBuf.AddColumn(DrCrTextBalance,FALSE,'',TRUE,FALSE,FALSE,'',1);//6
          ExcelBuf.AddColumn(SourceDesc,FALSE,'',TRUE,FALSE,FALSE,'',1);//7
          ExcelBuf.AddColumn("G/L Entry"."Location Code",FALSE,'',TRUE,FALSE,FALSE,'',1);//8
          recLoc.RESET;
          recLoc.SETRANGE(recLoc.Code,"G/L Entry"."Location Code");
          IF recLoc.FINDLAST THEN
          BEGIN
          ExcelBuf.AddColumn(recLoc.Name,FALSE,'',TRUE,FALSE,FALSE,'',1);//9
          END ELSE
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//9
        
          //>>13Nov2018
          ExcelBuf.AddColumn(GLDoc."Global Dimension 1 Code",FALSE,'',TRUE,FALSE,FALSE,'',1);
          DimVal13.RESET;
          DimVal13.SETRANGE("Dimension Code",'DIVISION');
          DimVal13.SETRANGE(Code,GLDoc."Global Dimension 1 Code");
          IF DimVal13.FINDLAST THEN
            ExcelBuf.AddColumn(DimVal13.Name,FALSE,'',TRUE,FALSE,FALSE,'',1)
          ELSE
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);
          //<<13Nov2018
        
          ExcelBuf.AddColumn("G/L Entry"."Global Dimension 2 Code",FALSE,'',TRUE,FALSE,FALSE,'',1);//10
        
          recResCentre.RESET;
          recResCentre.SETRANGE(recResCentre.Code,"G/L Entry"."Global Dimension 2 Code");
          IF recResCentre.FINDLAST THEN
          BEGIN
          ExcelBuf.AddColumn(recResCentre.Name,FALSE,'',TRUE,FALSE,FALSE,'',1);//11
          END ELSE
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//11
        
        
          ExcelBuf.AddColumn(GDrAmt,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//10Aug2017 //12
          GDrAmt := 0;//10Aug2017
          //ExcelBuf.AddColumn("G/L Entry"."Debit Amount",FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        
          ExcelBuf.AddColumn(GCrAmt,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//10Aug2017 //13
          GCrAmt := 0;//10Aug2017
          //ExcelBuf.AddColumn("G/L Entry"."Credit Amount",FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
          ExcelBuf.AddColumn(ABS(OpeningDRBal-OpeningCRBal+TransDebits-TransCredits),FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//14
          //ExcelBuf.AddColumn(TotalNarration,FALSE,'',TRUE,FALSE,FALSE,'',1);//10Aug2017 //15
          ExcelBuf.AddColumn(COPYSTR(TotalNarration,1,250),FALSE,'',TRUE,FALSE,FALSE,'',1);//RSPL-PA
          ExcelBuf.AddColumn(COPYSTR(TotalNarration,251,250),FALSE,'',TRUE,FALSE,FALSE,'',1);//RSPL-PA
          //ExcelBuf.AddColumn(TotalNarration[2],FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          //ExcelBuf.AddColumn(TotalNarration[3],FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          //ExcelBuf.AddColumn(TotalNarration[4],FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
          //ExcelBuf.AddColumn(TotalNarration[5],FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        END;
        *///RSPLSUM
        IF NOT PrintDetail THEN BEGIN

            ExcelBuf.NewRow;
            ExcelBuf.AddColumn("G/L Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
            ExcelBuf.AddColumn("G/L Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
            ExcelBuf.AddColumn(VendInvNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
            ExcelBuf.AddColumn(AccountName29, FALSE, '', FALSE, FALSE, FALSE, '', 1);//17May2017//4
                                                                                     //ExcelBuf.AddColumn(AccountName,FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
            ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
            ExcelBuf.AddColumn(SourceDesc, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
            ExcelBuf.AddColumn('"G/L Entry"."Location Code"', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
            recLoc.RESET;
            //recLoc.SETRANGE(recLoc.Code, "G/L Entry"."Location Code");
            IF recLoc.FINDLAST THEN BEGIN
                ExcelBuf.AddColumn(recLoc.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
            END ELSE
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//9

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//RSPLSUM
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//RSPLSUM
                                                                          //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//RSPLSUM
                                                                          //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//RSPLSUM
                                                                          /*//RSPLSUM
                                                                          //>>13Nov2018
                                                                          ExcelBuf.AddColumn(GLDoc."Global Dimension 1 Code",FALSE,'',FALSE,FALSE,FALSE,'',1);
                                                                          DimVal13.RESET;
                                                                          DimVal13.SETRANGE("Dimension Code",'DIVISION');
                                                                          DimVal13.SETRANGE(Code,GLDoc."Global Dimension 1 Code");
                                                                          IF DimVal13.FINDLAST THEN
                                                                            ExcelBuf.AddColumn(DimVal13.Name,FALSE,'',FALSE,FALSE,FALSE,'',1)
                                                                          ELSE
                                                                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);
                                                                          //<<13Nov2018
                                                                          *///RSPLSUM
            ExcelBuf.AddColumn("G/L Entry"."Global Dimension 2 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
            recResCentre.RESET;
            recResCentre.SETRANGE(recResCentre.Code, "G/L Entry"."Global Dimension 2 Code");
            IF recResCentre.FINDLAST THEN BEGIN
                ExcelBuf.AddColumn(recResCentre.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
            END ELSE
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//11

            ExcelBuf.AddColumn(GDrAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//10Aug2017 //12
            GDrAmt := 0;//10Aug2017

            ExcelBuf.AddColumn(GCrAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//10Aug2017 //13
            GCrAmt := 0;//10Aug2017

            ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//14
                                                                                                                                               //ExcelBuf.AddColumn(TotalNarration,FALSE,'',FALSE,FALSE,FALSE,'',1);//10Aug2017 //15 //RSPL-PA
            ExcelBuf.AddColumn(COPYSTR(TotalNarration, 1, 250), FALSE, '', FALSE, FALSE, FALSE, '', 1);//10Aug2017 //15 //RSPL-PA
            ExcelBuf.AddColumn(COPYSTR(TotalNarration, 251, 250), FALSE, '', FALSE, FALSE, FALSE, '', 1);//10Aug2017 //15 //RSPL-PA
                                                                                                         //ExcelBuf.AddColumn(TotalNarration[2],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
                                                                                                         //ExcelBuf.AddColumn(TotalNarration[3],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
                                                                                                         //ExcelBuf.AddColumn(TotalNarration[4],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
                                                                                                         //ExcelBuf.AddColumn(TotalNarration[5],FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        END;

    end;

    //// //[Scope('Internal')]
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 15Nov2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLNB
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 23-Mar-22
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Closing Balance', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Nov2018
        "G/L Entry".CALCSUMS("Debit Amount");
        "G/L Entry".CALCSUMS("Credit Amount");
        ExcelBuf.AddColumn(OpeningDRBal + "G/L Entry"."Debit Amount", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(OpeningCRBal + "G/L Entry"."Credit Amount", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 15Nov2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLnb
    end;

    // //[Scope('Internal')]
    procedure PrintNarration()
    var
        PosNar: Record "Posted Narration";
    begin
        i := 0;

        //CLEAR(TotalNarration); //10Aug2017
        recPostedNarration.RESET;
        recPostedNarration.SETRANGE(recPostedNarration."Document No.", "G/L Entry"."Document No.");
        recPostedNarration.SETRANGE(recPostedNarration."Transaction No.", "G/L Entry"."Transaction No.");
        IF recPostedNarration.FINDFIRST THEN
            REPEAT
                //i +=1;
                TotalNarration := TotalNarration + recPostedNarration.Narration; //10Aug2017
                                                                                 //TotalNarration[i] :=   recPostedNarration.Narration;

            UNTIL recPostedNarration.NEXT = 0;
    end;

    ////[Scope('Internal')]
    procedure PrintComment()
    var
        PurchCommentLine: Record "43";
    begin
        //CLEAR(TotalNarration); //10Aug2017
        IF "G/L Entry"."Source Code" = 'PURCHASES' THEN BEGIN
            recPurchCommentLine.RESET;
            recPurchCommentLine.SETRANGE(recPurchCommentLine."No.", "G/L Entry"."Document No.");
            IF recPurchCommentLine.FINDFIRST THEN
                TotalNarration := TotalNarration + recPurchCommentLine.Comment;//10Aug2017
                                                                               //TotalNarration[5] :=   recPurchCommentLine.Comment;
        END;

        IF "G/L Entry"."Source Code" = 'SALES' THEN BEGIN
            recSalesCommentLine.RESET;
            recSalesCommentLine.SETFILTER(recSalesCommentLine."Document Type", '%1|%2',
                                          recSalesCommentLine."Document Type"::"Posted Credit Memo",
                                          recSalesCommentLine."Document Type"::"Posted Invoice");
            recSalesCommentLine.SETRANGE(recSalesCommentLine."No.", "G/L Entry"."Document No.");
            IF recSalesCommentLine.FINDFIRST THEN BEGIN

                IF recSalesCommentLine."Document Type" = recSalesCommentLine."Document Type"::"Posted Invoice" THEN BEGIN
                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.", "G/L Entry"."Document No.");
                    recSIH.SETRANGE(recSIH."Order No.", '');
                    recSIH.SETRANGE(recSIH."Supplimentary Invoice", FALSE);
                    IF recSIH.FINDFIRST THEN BEGIN
                        TotalNarration := TotalNarration + recSalesCommentLine.Comment;//10Aug2017
                                                                                       //TotalNarration[1] :=  recSalesCommentLine.Comment;
                    END;
                END;


                IF recSalesCommentLine."Document Type" = recSalesCommentLine."Document Type"::"Posted Credit Memo" THEN BEGIN
                    TotalNarration := TotalNarration + recSalesCommentLine.Comment;//10Aug2017
                                                                                   //TotalNarration[2] := recSalesCommentLine.Comment;
                END;

            END;
        END;
    end;
}


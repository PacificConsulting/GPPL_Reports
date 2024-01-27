report 50187 "G/L Ledger dimension wise"
{
    // =iif(Fields!FirstRecord.Value and Fields!ControlAccountName.Value="",false,true)----1
    // =iif(Fields!PrintDetail.Value and not(Fields!FirstRecord.Value),false,true)----2 (Modified 2Condition) to display the FirstLine PrintDetail
    // VendInvNo--->OldValue20-->ChangeValue30-->15May2017
    // DimCode & DimensionValue Added-->15May2017 RB-N.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/GLLedgerdimensionwise.rdl';

    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'G/L Ledger dimension wise';

    dataset
    {
        dataitem("G/L Account"; "G/L Account")
        {
            DataItemTableView = SORTING("No.")
                                ORDER(Ascending)
                                WHERE("Account Type" = FILTER(Posting));
            RequestFilterFields = "No.", "Date Filter", "Global Dimension 1 Filter", "Global Dimension 2 Filter";
            column(AccNameNo; 'Account :' + '  ' + "G/L Account".Name + '               ' + 'Dimension :' + varrecdimension.Name + '  ' + '  Value :' + varrecDimensionvalue.Name)
            {
            }
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
                column(DebitAmount_GLEntry; "Debit Amount")
                {
                }
                column(CreditAmoutnt_GLEntry; "Credit Amount")
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
                column(DrCrTextBalance3; DrCrTextBalance)
                {
                }
                column(OneEntryRecord; OneEntryRecord)
                {
                }
                column(EntryNo_GLEntry; "Entry No.")
                {
                }
                column(VendInvNo; VendInvNo)
                {
                }
                column(RespCenter; "Global Dimension 2 Code")
                {
                }
                column(Dimcode15; Dimcode15)
                {
                }
                column(DimName15; DimName15)
                {
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
                                DetailAmt := GLEntry.Amount;
                            IF DetailAmt > 0 THEN
                                DrCrText := 'Dr';
                            IF DetailAmt < 0 THEN
                                DrCrText := 'Cr';

                            IF NOT PrintDetail THEN
                                AccountName := Text16500;
                            /*   ELSE
                                  AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.")
                              ; */

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
                                    AccountName := Text16500;
                                /* ELSE
                                    AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.")
                                ; */

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
                            // AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.");
                        END;


                        IF (FirstRecord) AND (PrintToExcel) AND (ControlAccountName = '') THEN
                            MakeExcelDataBody;

                        IF (FirstRecord) AND (PrintToExcel) AND (ControlAccountName <> '') THEN
                            MakeExcelDataBody1;

                        IF (FirstRecord) AND (PrintToExcel) AND (ControlAccountName <> '') THEN
                            MakeExcelDetailBody1;

                        IF PrintDetail AND PrintToExcel THEN
                            MakeExcelDetailBody;
                    end;

                    trigger OnPreDataItem()
                    begin
                        SETRANGE(Number, 1, GLEntry.COUNT);
                        FirstRecord := TRUE;

                        IF GLEntry.COUNT = 1 THEN;
                        //CurrReport.BREAK; //20Apr2017 Commented asPerNAV2009
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

                        IF (PrintToExcel) AND (Narration <> '') THEN BEGIN
                            ExcelBuf.NewRow;
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
                            ExcelBuf.AddColumn(Narration, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
                            ExcelBuf.NewRow;
                        END;
                    end;

                    trigger OnPreDataItem()
                    begin
                        IF NOT PrintLineNarration THEN
                            CurrReport.BREAK;
                    end;
                }
                dataitem(PostedNarration1; "Posted Narration")
                {
                    DataItemLink = "Document No." = FIELD("Document No."),
                                   "Transaction No." = FIELD("Transaction No.");
                    DataItemTableView = SORTING("Entry No.", "Transaction No.", "Line No.")
                                        WHERE("Entry No." = FILTER(0));
                    column(Narration_PostedNarration1; Narration)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        IF (PrintToExcel) AND (Narration <> '') THEN BEGIN
                            /*
                            ExcelBuf.NewRow;
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//1
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//3
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//4
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//5
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//6
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//7
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//8
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//9
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//10
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//11
                            ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//12
                            */
                            TxtNarration += ', ' + PostedNarration1.Narration;

                        END;

                    end;

                    trigger OnPostDataItem()
                    begin
                        TxtNarration := COPYSTR(TxtNarration, 2, STRLEN(TxtNarration));
                        ExcelBuf.AddColumn(TxtNarration, FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
                    end;

                    trigger OnPreDataItem()
                    begin
                        CLEAR(TxtNarration);
                        IF NOT PrintVchNarration THEN
                            CurrReport.BREAK;

                        GLEntry2.RESET;
                        GLEntry2.SETCURRENTKEY("Posting Date", "Source Code", "Transaction No.");
                        GLEntry2.SETRANGE("Posting Date", "G/L Entry"."Posting Date");
                        GLEntry2.SETRANGE("Source Code", "G/L Entry"."Source Code");
                        GLEntry2.SETRANGE("Transaction No.", "G/L Entry"."Transaction No.");
                        GLEntry2.FINDLAST;
                        IF NOT (GLEntry2."Entry No." = "G/L Entry"."Entry No.") THEN;
                        //CurrReport.BREAK; //20Apr2017 Commented asPerNAV2009
                    end;
                }
                dataitem("Purch. Comment Line"; "Purch. Comment Line")
                {
                    DataItemLink = "No." = FIELD("Document No.");
                    DataItemTableView = SORTING("Document Type", "No.", "Line No.");
                    column(PurchComm; "Purch. Comment Line".Comment)
                    {
                    }
                    column(PurDocNo; "Purch. Comment Line"."No.")
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        //>>1NAV2009

                        IF NOT ("G/L Entry"."Source Code" = 'PURCHASES') THEN
                            CurrReport.SKIP;

                        //<<1NAV2009


                        IF (PrintToExcel) AND ("Purch. Comment Line".Comment <> '') THEN BEGIN
                            /*
                              ExcelBuf.NewRow;
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//1
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//3
                              ExcelBuf.AddColumn("Purch. Comment Line".Comment,FALSE,'',FALSE,FALSE,FALSE,'',1);//4
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//5
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//6
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//7
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//8
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//9
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//10
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//11
                              ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//12
                            */
                            TxtComment += ', ' + "Purch. Comment Line".Comment;
                        END;

                    end;

                    trigger OnPostDataItem()
                    begin
                        TxtComment := COPYSTR(TxtComment, 2, STRLEN(TxtComment));
                        ExcelBuf.AddColumn(TxtComment, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
                    end;

                    trigger OnPreDataItem()
                    begin
                        CLEAR(TxtComment);
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1NAV2009

                    //EBT STIVAN ---(030111)--------To Capture Staff Code and Name-------------------------------------START
                    //++NAV2016Dimension
                    StaffCode := '';
                    StaffName := '';

                    recDimSet.RESET;
                    recDimSet.SETRANGE("Dimension Set ID", "G/L Entry"."Dimension Set ID");
                    recDimSet.SETRANGE("Dimension Code", 'Staff');
                    IF recDimSet.FINDFIRST THEN BEGIN
                        StaffCode := recDimSet."Dimension Value Code";
                        recDimSet.CALCFIELDS("Dimension Value Name");

                        StaffName := recDimSet."Dimension Value Name"
                    END;

                    /*
                    DocumentDimension.RESET;
                    DocumentDimension.SETRANGE(DocumentDimension."Table ID",17);
                    DocumentDimension.SETRANGE(DocumentDimension."Entry No.","G/L Entry"."Entry No.");
                    DocumentDimension.SETRANGE(DocumentDimension."Dimension Code",'Staff');
                    IF DocumentDimension.FINDFIRST THEN
                       StaffCode := DocumentDimension."Dimension Value Code";
                    REPEAT
                    StaffName := '';
                    recDimensionvalue.RESET;
                    recDimensionvalue.SETRANGE(recDimensionvalue."Dimension Code",'Staff');
                    recDimensionvalue.SETRANGE(recDimensionvalue.Code,StaffCode);
                    IF recDimensionvalue.FINDFIRST THEN
                     StaffName := recDimensionvalue.Name;
                    UNTIL DocumentDimension.NEXT = 0;
                    //EBT STIVAN ---(030111)--------To Capture Staff Code and Name---------------------------------------END
                    */
                    //--NAV2016Dimension


                    //RSPL
                    //++NAV2016Dimension
                    VehicleCode := '';
                    VehicleName := '';

                    recDimSet.RESET;
                    recDimSet.SETRANGE("Dimension Set ID", "G/L Entry"."Dimension Set ID");
                    recDimSet.SETRANGE("Dimension Code", 'VEHICLE');
                    IF recDimSet.FINDFIRST THEN BEGIN
                        VehicleCode := recDimSet."Dimension Value Code";
                        recDimSet.CALCFIELDS("Dimension Value Name");

                        VehicleName := recDimSet."Dimension Value Name"
                    END;

                    //--NAV2016Dimension
                    /*
                    DocumentDimension.RESET;
                    DocumentDimension.SETRANGE(DocumentDimension."Table ID",17);
                    DocumentDimension.SETRANGE(DocumentDimension."Entry No.","G/L Entry"."Entry No.");
                    DocumentDimension.SETRANGE(DocumentDimension."Dimension Code",'VEHICLE');
                    IF DocumentDimension.FINDFIRST THEN
                       VehicleCode := DocumentDimension."Dimension Value Code";
                    REPEAT
                    StaffName := '';
                    recDimensionvalue.RESET;
                    recDimensionvalue.SETRANGE(recDimensionvalue."Dimension Code",'VEHICLE');
                    recDimensionvalue.SETRANGE(recDimensionvalue.Code,StaffCode);
                    IF recDimensionvalue.FINDFIRST THEN
                     VehicleName := recDimensionvalue.Name;
                    UNTIL DocumentDimension.NEXT = 0;
                    */


                    //++NAV2016Dimension
                    PhoneCode := '';
                    PhoneName := '';

                    recDimSet.RESET;
                    recDimSet.SETRANGE("Dimension Set ID", "G/L Entry"."Dimension Set ID");
                    recDimSet.SETRANGE("Dimension Code", 'PHONE');
                    IF recDimSet.FINDFIRST THEN BEGIN
                        PhoneCode := recDimSet."Dimension Value Code";
                        recDimSet.CALCFIELDS("Dimension Value Name");

                        PhoneName := recDimSet."Dimension Value Name"
                    END;

                    //--NAV2016Dimension

                    /*
                    PhoneCode := '';
                    PhoneName:='';
                    DocumentDimension.RESET;
                    DocumentDimension.SETRANGE(DocumentDimension."Table ID",17);
                    DocumentDimension.SETRANGE(DocumentDimension."Entry No.","G/L Entry"."Entry No.");
                    DocumentDimension.SETRANGE(DocumentDimension."Dimension Code",'PHONE');
                    IF DocumentDimension.FINDFIRST THEN
                       PhoneCode := DocumentDimension."Dimension Value Code";
                    REPEAT
                    StaffName := '';
                    recDimensionvalue.RESET;
                    recDimensionvalue.SETRANGE(recDimensionvalue."Dimension Code",'PHONE');
                    recDimensionvalue.SETRANGE(recDimensionvalue.Code,StaffCode);
                    IF recDimensionvalue.FINDFIRST THEN
                     PhoneName := recDimensionvalue.Name;
                    UNTIL DocumentDimension.NEXT = 0;
                    */
                    //RSPL
                    //ebt===================================

                    //++NAV2016 DimsetFilter
                    recDimSet2.RESET;
                    recDimSet2.SETRANGE("Dimension Set ID", "G/L Entry"."Dimension Set ID");
                    IF vardimension <> '' THEN
                        recDimSet2.SETRANGE("Dimension Code", vardimension);

                    IF vardimensionvalue <> '' THEN
                        recDimSet2.SETRANGE("Dimension Value Code", vardimensionvalue);

                    IF (NOT recDimSet2.FIND('-')) THEN
                        CurrReport.SKIP;

                    //--NAV2016 DimsetFilter
                    /*
                    ledgerdim.RESET;
                    ledgerdim.SETRANGE(ledgerdim."Entry No.","G/L Entry"."Entry No.");
                    ledgerdim.SETRANGE(ledgerdim."Table ID",17);
                    IF vardimension<>'' THEN
                      ledgerdim.SETRANGE("Dimension Code",vardimension);
                    IF vardimensionvalue <> '' THEN
                      ledgerdim.SETRANGE("Dimension Value Code",vardimensionvalue);
                    
                    IF  (NOT ledgerdim.FIND('-')) THEN
                       CurrReport.SKIP;
                    */

                    //ebt===================================

                    //>>NAV2009 VendorInvoiceNo
                    VendInvNo := ''; //19Apr2017
                    IF "G/L Entry"."Source Code" = 'PURCHASES' THEN BEGIN
                        PurchInvHeader.RESET;
                        PurchInvHeader.SETRANGE(PurchInvHeader."No.", "G/L Entry"."Document No.");
                        IF PurchInvHeader.FINDFIRST THEN BEGIN
                            VendInvNo := PurchInvHeader."Vendor Invoice No.";
                            //MESSAGE('Vendor Inv No. %1 \ Doc No %2',VendInvNo,"G/L Entry"."Document No.");
                        END;
                    END;
                    //<<NAV2009 VendorInvoiceNo

                    //<<1NAV2009

                    //>>15May2017
                    CLEAR(Dimcode15);
                    CLEAR(DimName15);
                    IF vardimension <> '' THEN BEGIN
                        recDimSet15.RESET;
                        recDimSet15.SETRANGE("Dimension Set ID", "G/L Entry"."Dimension Set ID");
                        recDimSet15.SETRANGE("Dimension Code", vardimension);
                        IF recDimSet15.FINDFIRST THEN BEGIN
                            Dimcode15 := recDimSet15."Dimension Value Code";
                            recDimSet15.CALCFIELDS("Dimension Value Name");
                            DimName15 := recDimSet15."Dimension Value Name";
                        END;
                    END;

                    IF vardimension = 'STAFF' THEN BEGIN
                        recDimSet15.RESET;
                        recDimSet15.SETRANGE("Dimension Set ID", "G/L Entry"."Dimension Set ID");
                        recDimSet15.SETRANGE("Dimension Code", vardimension);
                        IF recDimSet15.FINDFIRST THEN BEGIN
                            IF recDimSet15.COUNT = 1 THEN BEGIN
                                Dimcode15 := recDimSet15."Dimension Value Code";
                                recDimSet15.CALCFIELDS("Dimension Value Name");
                                DimName15 := recDimSet15."Dimension Value Name";
                            END;
                        END;
                    END;
                    //<<15May2017


                    ControlAccountName := ''; //20Apr2017 NAV2009
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
                    IF ControlAccount THEN;
                    // ControlAccountName := Daybook.FindGLAccName("Source Type", "Entry No.", "Source No.", "G/L Account No.");

                    IF Amount > 0 THEN
                        TransDebits := TransDebits + Amount;
                    IF Amount < 0 THEN
                        TransCredits := TransCredits - Amount;

                    SourceDesc := '';
                    IF "Source Code" <> '' THEN BEGIN
                        SourceCode.GET("Source Code");
                        SourceDesc := SourceCode.Description;
                    END;

                    AccountName := '';
                    IF OneEntryRecord AND (ControlAccountName <> '') THEN BEGIN
                        //  AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.");
                        DrCrTextBalance := '';
                        IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits > 0 THEN
                            DrCrTextBalance := 'Dr';
                        IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits < 0 THEN
                            DrCrTextBalance := 'Cr';
                    END ELSE
                        IF OneEntryRecord AND (ControlAccountName = '') THEN BEGIN
                            // AccountName := Daybook.FindGLAccName(GLEntry."Source Type", GLEntry."Entry No.", GLEntry."Source No.", GLEntry."G/L Account No.");
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

                        IF OpeningDRBal - OpeningCRBal + TransDebits - TransCredits = 0 THEN //24Apr2016
                            DrCrTextBalance2 := '';//24Apr2016
                    END;

                    TotalDebitAmount += "G/L Entry"."Debit Amount";
                    TotalCreditAmount += "G/L Entry"."Credit Amount";


                    IF (OneEntryRecord) AND (ControlAccountName <> '') AND (PrintToExcel) THEN
                        MakeExcelDataBody2;

                    IF (OneEntryRecord) AND (ControlAccountName <> '') AND (PrintToExcel) THEN
                        MakeExcelDetailBody2;

                end;

                trigger OnPostDataItem()
                begin
                    AccountChanged := TRUE;
                end;

                trigger OnPreDataItem()
                begin
                    IF LocationCode <> '' THEN
                        //SETFILTER("Location Code", LocationCode);
                    GLEntry.RESET;
                    GLEntry.SETCURRENTKEY("Transaction No.");
                    TotalDebitAmount := 0;
                    TotalCreditAmount := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1NAV2009

                IF ("G/L Account"."No." = '72010010') OR ("G/L Account"."No." = '72010020') OR ("G/L Account"."No." = '72010030')
                OR ("G/L Account"."No." = '72010040') OR ("G/L Account"."No." = '72010050') OR ("G/L Account"."No." = '72010060')
                OR ("G/L Account"."No." = '72010070') OR ("G/L Account"."No." = '72010080') THEN
                    CurrReport.SKIP;

                varrecdimension.RESET;
                varrecdimension.SETRANGE(varrecdimension.Code, vardimension);
                IF varrecdimension.FIND('-') THEN;

                varrecDimensionvalue.RESET;
                varrecDimensionvalue.SETRANGE(varrecDimensionvalue."Dimension Code", vardimension);
                varrecDimensionvalue.SETRANGE(varrecDimensionvalue.Code, vardimensionvalue);
                IF varrecDimensionvalue.FIND('-') THEN;
                //>>1NAV2009


                IF AccountNo <> "No." THEN BEGIN
                    AccountNo := "No.";
                    OpeningDRBal := 0;
                    OpeningCRBal := 0;

                    GLEntry2.RESET;
                    GLEntry2.SETCURRENTKEY("G/L Account No.", "Business Unit Code", "Global Dimension 1 Code",
                    "Global Dimension 2 Code", "Close Income Statement Dim. ID", "Posting Date"/*, "Location Code"*/);
                    GLEntry2.SETRANGE("G/L Account No.", "No.");
                    GLEntry2.SETFILTER("Posting Date", '%1..%2', 0D, CLOSINGDATE(GETRANGEMIN("Date Filter") - 1));
                    IF "Global Dimension 1 Filter" <> '' THEN
                        GLEntry2.SETFILTER("Global Dimension 1 Code", "Global Dimension 1 Filter");
                    IF "Global Dimension 2 Filter" <> '' THEN
                        GLEntry2.SETFILTER("Global Dimension 2 Code", "Global Dimension 2 Filter");
                    IF LocationCode <> '' THEN;
                    // GLEntry2.SETFILTER("Location Code", LocationCode);

                    //GLEntry2.CALCSUMS(Amount); //Commented Asper NAV2009

                    IF GLEntry2.FIND('-') THEN
                        REPEAT
                            recDimSet3.RESET;
                            recDimSet3.SETRANGE("Dimension Set ID", GLEntry2."Dimension Set ID");
                            IF vardimension <> '' THEN
                                recDimSet3.SETRANGE("Dimension Code", vardimension);
                            IF vardimensionvalue <> '' THEN
                                recDimSet3.SETRANGE("Dimension Value Code", vardimensionvalue);

                            IF recDimSet3.FIND('-') THEN
                                varopening += GLEntry2.Amount;
                        /*
                        //>>OldCode Commented
                        ledgerdim.RESET;
                        ledgerdim.SETRANGE(ledgerdim."Entry No.",GLEntry2."Entry No.");
                        ledgerdim.SETRANGE(ledgerdim."Table ID",17);
                        IF vardimension<>'' THEN
                          ledgerdim.SETRANGE("Dimension Code",vardimension);
                        IF vardimensionvalue <> '' THEN
                          ledgerdim.SETRANGE("Dimension Value Code",vardimensionvalue);
                        IF  ledgerdim.FIND('-') THEN
                          varopening:=varopening+GLEntry2.Amount;
                        //<<OldCode Commented
                        */
                        UNTIL GLEntry2.NEXT = 0;


                    IF varopening > 0 THEN
                        OpeningDRBal := varopening;
                    IF varopening < 0 THEN
                        OpeningCRBal := -varopening;

                    /*
                    IF GLEntry2.Amount > 0 THEN
                      OpeningDRBal := GLEntry2.Amount;
                    IF GLEntry2.Amount < 0 THEN
                      OpeningCRBal := -GLEntry2.Amount;

                    */
                    DrCrTextBalance := '';
                    IF OpeningDRBal - OpeningCRBal > 0 THEN
                        DrCrTextBalance := 'Dr';
                    IF OpeningDRBal - OpeningCRBal < 0 THEN
                        DrCrTextBalance := 'Cr';
                    IF (OpeningCRBal = 0) AND (OpeningDRBal = 0) THEN //24Apr2017
                        DrCrTextBalance := '';//24Apr2017
                END;


                IF PrintToExcel THEN BEGIN
                    ExcelBuf.NewRow;
                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
                    ExcelBuf.AddColumn('Opening Balance As On' + ' ' + FORMAT("G/L Account".GETRANGEMIN("Date Filter")), FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
                    ExcelBuf.AddColumn(OpeningDRBal, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);//8
                    ExcelBuf.AddColumn(OpeningCRBal, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);//9
                    ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal), FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);//10
                    ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', TRUE, FALSE, FALSE, '', 0);//11
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
                    ExcelBuf.NewRow;
                END;

            end;

            trigger OnPostDataItem()
            begin

                IF PrintToExcel THEN BEGIN
                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
                    ExcelBuf.AddColumn(Closing_BalanceCaptionLbl, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
                    ExcelBuf.AddColumn(OpeningDRBal + TotalDebitAmount, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);//8
                    ExcelBuf.AddColumn(OpeningCRBal + TotalCreditAmount, FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);//9
                    ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', TRUE, FALSE, FALSE, '#,##0.00', 0);//10
                    ExcelBuf.AddColumn(DrCrTextBalance2, FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
                    ExcelBuf.NewRow;
                END;
            end;

            trigger OnPreDataItem()
            begin
                CurrReport.CREATETOTALS(TransDebits, TransCredits, "G/L Entry"."Debit Amount", "G/L Entry"."Credit Amount");
                IF LocationCode <> '' THEN
                    LocationFilter := 'Location Code: ' + LocationCode;

                IF PrintToExcel THEN
                    MakeExcelDataHeader;
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
                        ApplicationArea = all;
                        Caption = 'Print Detail';
                    }
                    field(PrintLineNarration; PrintLineNarration)
                    {
                        ApplicationArea = all;
                        Caption = 'Print Line Narration';
                    }
                    field(PrintVchNarration; PrintVchNarration)
                    {
                        ApplicationArea = all;
                        Caption = 'Print Voucher Narration';
                    }
                    field(LocationCode; LocationCode)
                    {
                        ApplicationArea = all;
                        Caption = 'Location Code';
                        TableRelation = Location;
                    }
                    field(Dimension; vardimension)
                    {
                        ApplicationArea = all;
                        TableRelation = Dimension;
                    }
                    field("Dimension Value"; vardimensionvalue)
                    {
                        ApplicationArea = all;
                        TableRelation = "Dimension Value".Code;

                        trigger OnLookup(var Text: Text): Boolean
                        begin

                            //>>1
                            IF vardimension = '' THEN
                                ERROR('Please select dimension');

                            recDimensionvalue.RESET;
                            recDimensionvalue.SETRANGE("Dimension Code", vardimension);
                            IF recDimensionvalue.FIND('-') THEN
                                IF PAGE.RUNMODAL(0, recDimensionvalue) = ACTION::LookupOK THEN
                                    vardimensionvalue := recDimensionvalue.Code;
                            //IF FORM.RUNMODAL(0,recDimensionvalue)=ACTION::LookupOK THEN
                            vardimensionvalue := recDimensionvalue.Code;

                            //<<1
                        end;
                    }
                    field(PrintToExcel; PrintToExcel)
                    {
                        ApplicationArea = all;
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
        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin
        CompInfo.GET;


        //>>1NAV2009

        /*
        //EBT STIVAN ---(12122012)--- To Filter Report as per the Report View Role ------START
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (MemberOf.FINDFIRST) THEN
        ERROR('You do not have permission to run the report');
        //EBT STIVAN ---(12122012)--- To Filter Report as per the Report View Role --------END
        *///Commented 19Apr2017


        GLNo := "G/L Account".GETFILTER("G/L Account"."No.");
        datefilter := FORMAT("G/L Account".GETFILTER("G/L Account"."Date Filter"));
        DIVifilter := "G/L Account".GETFILTER("G/L Account"."Global Dimension 1 Filter");
        "RESP.filter" := "G/L Account".GETFILTER("G/L Account"."Global Dimension 2 Filter");

        GLNAME := '';
        "G/L Account".RESET;
        "G/L Account".SETRANGE("G/L Account"."No.", GLNo);
        IF "G/L Account".FINDFIRST THEN BEGIN
            GLNAME := "G/L Account".Name;
        END;

        varrecdimension.RESET;
        varrecdimension.SETRANGE(varrecdimension.Code, vardimension);
        IF varrecdimension.FIND('-') THEN;

        varrecDimensionvalue.RESET;
        varrecDimensionvalue.SETRANGE(varrecDimensionvalue.Code, vardimensionvalue);
        IF varrecDimensionvalue.FIND('-') THEN;
        //<<1NAV2009

        IF PrintToExcel THEN
            MakeExcelInfo;

    end;

    var
        CompInfo: Record 79;
        GLAcc: Record 15;
        GLEntry: Record 17;
        GLEntry2: Record 17;
        // Daybook: Report "Duty Code No.";
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
        "----19Apr2017--NAV2009Mig----": Integer;
        recDimSet: Record 480;
        recDimSet2: Record 480;
        recDimSet3: Record 480;
        vardimension: Code[20];
        vardimensionvalue: Code[120];
        GLNo: Code[10];
        GLNAME: Text[50];
        datefilter: Text[50];
        DIVifilter: Text[50];
        "RESP.filter": Text[50];
        GLentryNo: Integer;
        VehicleCode: Code[20];
        VehicleName: Text[50];
        PhoneCode: Code[20];
        PhoneName: Text[50];
        GLACCNAME: Text[50];
        GLACCCODE: Code[20];
        recGLacc: Record 15;
        varrecdimension: Record 348;
        varrecDimensionvalue: Record 349;
        StaffCode: Code[10];
        StaffName: Text[50];
        PurchInvHeader: Record 122;
        VendInvNo: Code[35];
        recDimensionvalue: Record 349;
        varopening: Decimal;
        "----15May2017": Integer;
        recDimSet15: Record 480;
        Dimcode15: Code[20];
        DimName15: Text[80];
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        Text0000: Label 'Data';
        Text0001: Label 'G/L Ledger dimension wise';
        Text0002: Label 'Company Name';
        Text0003: Label 'Report No.';
        Text0004: Label 'Report Name';
        Text0005: Label 'USER Id';
        Text0006: Label 'Date';
        TxtNarration: Text;
        TxtComment: Text;

    // //[Scope('Internal')]
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
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0002), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('G/L Ledger dimension wise', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0003), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"G/L Ledger dimension wise", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;

        IF PrintToExcel THEN BEGIN

            //>>23Feb2017
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
            ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
            ExcelBuf.NewRow;

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
            ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
            ExcelBuf.NewRow;
            //<<

        END;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin

        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text0000,Text0001,COMPANYNAME,USERID);
        /*  ExcelBuf.CreateBookAndOpenExcel('', Text0000, '', '', '');
         ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook(Text0000);
        ExcelBuf.WriteSheet(Text0000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text0000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(FORMAT(GLEntry."Posting Date"), FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(GLEntry."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(VendInvNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(AccountName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ABS(DetailAmt), FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//5
        ExcelBuf.AddColumn(DrCrText, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn("G/L Entry"."Global Dimension 2 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn("G/L Entry"."Debit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//8
        ExcelBuf.AddColumn("G/L Entry"."Credit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//9
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//10
        ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn(Dimcode15 + ' ' + DimName15, FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody1()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(FORMAT(GLEntry."Posting Date"), FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(GLEntry."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(VendInvNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(ControlAccountName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn("G/L Entry"."Global Dimension 2 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn("G/L Entry"."Debit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//8
        ExcelBuf.AddColumn("G/L Entry"."Credit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//9
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//10
        ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody2()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(FORMAT("G/L Entry"."Posting Date"), FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn("G/L Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(VendInvNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        IF ControlAccountName = '' THEN
            ExcelBuf.AddColumn(AccountName, FALSE, '', FALSE, FALSE, FALSE, '', 1)//4
        ELSE
            ExcelBuf.AddColumn(ControlAccountName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//5
        ExcelBuf.AddColumn(DrCrText, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn("G/L Entry"."Global Dimension 2 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn("G/L Entry"."Debit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//8
        ExcelBuf.AddColumn("G/L Entry"."Credit Amount", FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//9
        ExcelBuf.AddColumn(ABS(OpeningDRBal - OpeningCRBal + TransDebits - TransCredits), FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//10
        ExcelBuf.AddColumn(DrCrTextBalance, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn(Dimcode15 + ' ' + DimName15, FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('G/L Ledger DimensionWise GP', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('For The Period', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('FROM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn(FORMAT("G/L Account".GETRANGEMIN("G/L Account"."Date Filter")), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('TO', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn(FORMAT("G/L Account".GETRANGEMAX("G/L Account"."Date Filter")), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Vendor Inv. No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Account Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Resp. Center', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Debit Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Credit Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Balance', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Dimension', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Narration', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Purchase Comment', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDetailBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(AccountName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ABS(GLEntry.Amount), FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//5
        ExcelBuf.AddColumn(DrCrText, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDetailBody1()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(AccountName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ABS(DetailAmt), FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//5
        ExcelBuf.AddColumn(DrCrText, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDetailBody2()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(AccountName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
    end;
}


report 50068 "Bank Payment Posted Voucher"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 21Sep2017   RB-N         Amount Paid Value retireved from Closed By Amount(LCY)
    // 16Mar2018   RB-N         Skipping UnApplied Entries from Detail Vendor Ledgers
    // 14Sep2018   RB-N         Issue for displaying CreditMemo Amount
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BankPaymentPostedVoucher.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Posted Voucher';

    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("Document No.", "Posting Date", Amount)
                                ORDER(Descending)
                                WHERE("G/L Account No." = FILTER(<> 12912050 | <> 12912060));
            RequestFilterFields = "Posting Date", "Document No.";
            column(VoucherSourceDesc; SourceDesc + ' Voucher')
            {
            }
            column(DocumentNo_GLEntry; "Document No.")
            {
            }
            column(PostingDateFormatted; 'Date: ' + FORMAT("Posting Date"))
            {
            }
            column(CompanyInformationAddress; CompanyInformation.Address + ' ' + CompanyInformation."Address 2" + '  ' + CompanyInformation.City)
            {
            }
            column(CompanyInformationName; CompanyInformation.Name)
            {
            }
            column(CreditAmount_GLEntry; "Credit Amount")
            {
            }
            column(DebitAmount_GLEntry; "Debit Amount")
            {
            }
            column(DrText; DrText)
            {
            }
            column(GLAccName; GLAccName)
            {
            }
            column(CrText; CrText)
            {
            }
            column(DebitAmountTotal; DebitAmountTotal)
            {
            }
            column(CreditAmountTotal; CreditAmountTotal)
            {
            }
            column(ChequeDetail; 'Cheque No: ' + ChequeNo + '  Dated: ' + FORMAT(ChequeDate))
            {
            }
            column(ChequeNo; ChequeNo)
            {
            }
            column(ChequeDate; ChequeDate)
            {
            }
            column(RsNumberText1NumberText2; 'Rs. ' + NumberText[1] + ' ' + NumberText[2])
            {
            }
            column(EntryNo_GLEntry; "Entry No.")
            {
            }
            column(PostingDate_GLEntry; "Posting Date")
            {
            }
            column(TransactionNo_GLEntry; "Transaction No.")
            {
            }
            column(VoucherNoCaption; VoucherNoCaptionLbl)
            {
            }
            column(CreditAmountCaption; CreditAmountCaptionLbl)
            {
            }
            column(DebitAmountCaption; DebitAmountCaptionLbl)
            {
            }
            column(ParticularsCaption; ParticularsCaptionLbl)
            {
            }
            column(AmountInWordsCaption; AmountInWordsCaptionLbl)
            {
            }
            column(PreparedByCaption; PreparedByCaptionLbl)
            {
            }
            column(CheckedByCaption; CheckedByCaptionLbl)
            {
            }
            column(ApprovedByCaption; ApprovedByCaptionLbl)
            {
            }
            column(Name_UserName; vUser."Full Name")
            {
            }
            column(DimensionValue; DimensionValue)
            {
            }
            column(StaffDimension; StaffDimension)
            {
            }
            column(Resp; RespCtr.Code)
            {
            }
            column(CheckPrintName_GLEntry; "G/L Entry"."Check Print Name")
            {
            }
            column(SatffCode_Footer; StaffCode + ' - ' + StaffName)
            {
            }
            column(DepositEMDcode; DepositEMDcode + ' - ' + DepositEMDName)
            {
            }
            column(DepositDealerCode; DepositDealerCode + ' - ' + DepositDealerName)
            {
            }
            column(DepositOthersCode; DepositOthersCode + ' - ' + DepositOthersName)
            {
            }
            column(DepositRentCode; DepositRentCode + ' - ' + DepositRentName)
            {
            }
            column(PhoneNo; PhoneNo + ' - ' + PhoneNoDesc)
            {
            }
            column(VehicleCode; VehicleCode + ' - ' + VehicaleName)
            {
            }
            column(SatffCode_Footer1; StaffCode + '' + StaffName)
            {
            }
            column(DepositEMDcode1; DepositEMDcode + '' + DepositEMDName)
            {
            }
            column(DepositDealerCode1; DepositDealerCode + ' ' + DepositDealerName)
            {
            }
            column(DepositOthersCode1; DepositOthersCode + ' ' + DepositOthersName)
            {
            }
            column(DepositRentCode1; DepositRentCode + '' + DepositRentName)
            {
            }
            column(PhoneNo1; PhoneNo + '' + PhoneNoDesc)
            {
            }
            column(VehicleCode1; VehicleCode + ' ' + VehicaleName)
            {
            }
            column(UserName; UserName)
            {
            }
            dataitem(LineNarration; "Posted Narration")
            {
                DataItemLink = "Transaction No." = FIELD("Transaction No."),
                               "Entry No." = FIELD("Entry No.");
                DataItemTableView = SORTING("Entry No.", "Transaction No.", "Line No.");
                column(Narration_LineNarration; Narration)
                {
                }
                column(PrintLineNarration; PrintLineNarration)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    IF PrintLineNarration THEN BEGIN
                        PageLoop := PageLoop - 1;
                        LinesPrinted := LinesPrinted + 1;
                    END;
                end;
            }
            dataitem(DataItem5444; Integer)
            {
                DataItemTableView = SORTING(Number);
                column(IntegerOccurcesCaption; IntegerOccurcesCaptionLbl)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    PageLoop := PageLoop - 1;
                end;

                trigger OnPreDataItem()
                begin
                    GLEntry.SETCURRENTKEY("Document No.", "Posting Date", Amount);
                    GLEntry.ASCENDING(FALSE);
                    GLEntry.SETRANGE("Posting Date", "G/L Entry"."Posting Date");
                    GLEntry.SETRANGE("Document No.", "G/L Entry"."Document No.");
                    GLEntry.FINDLAST;
                    IF NOT (GLEntry."Entry No." = "G/L Entry"."Entry No.") THEN
                        CurrReport.BREAK;

                    SETRANGE(Number, 1, PageLoop)
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
                column(NarrationCaption; NarrationCaptionLbl)
                {
                }

                trigger OnPreDataItem()
                begin
                    GLEntry.SETCURRENTKEY("Document No.", "Posting Date", Amount);
                    GLEntry.ASCENDING(FALSE);
                    GLEntry.SETRANGE("Posting Date", "G/L Entry"."Posting Date");
                    GLEntry.SETRANGE("Document No.", "G/L Entry"."Document No.");
                    GLEntry.FINDLAST;
                    IF NOT (GLEntry."Entry No." = "G/L Entry"."Entry No.") THEN
                        CurrReport.BREAK;
                end;
            }
            dataitem("Dimension Set Entry"; "Dimension Set Entry")
            {
                DataItemLink = "Dimension Set ID" = FIELD("Dimension Set ID");
                DataItemTableView = WHERE("Dimension Code" = FILTER('STAFF"|"DEPOSIT-EMD"|"DEPOSIT-DEALER"|"DEPOSIT-OTHERS"|DEPOSIT-RENT|PHONE|VEHICLE'));
                column(DimensionValueCode_DimensionSetEntry; "Dimension Set Entry"."Dimension Value Code")
                {
                }
                column(DimensionValueName_DimensionSetEntry; "Dimension Set Entry"."Dimension Value Name")
                {
                }
            }
            dataitem("Vendor Ledger Entry"; "Vendor Ledger Entry")
            {
                DataItemLink = "Applied Bank Payment Doc .No" = FIELD("Document No."),
                               "Vendor No." = FIELD("Bal. Account No.");
                DataItemTableView = SORTING("Document No.", "Document Type", "Vendor No.");

                trigger OnAfterGetRecord()
                begin
                    "Vendor Ledger Entry".CALCFIELDS("Vendor Ledger Entry".Amount);

                    //>>16Mar2018
                    DVLE.RESET;
                    DVLE.SETRANGE("Document No.", "Applied Bank Payment Doc .No");
                    DVLE.SETRANGE("Vendor Ledger Entry No.", "Entry No.");
                    DVLE.SETRANGE(Unapplied, FALSE);
                    IF NOT DVLE.FINDFIRST THEN
                        CurrReport.SKIP;
                    //<<16Mar2018

                    Buffer.INIT;
                    Z := Z + 1;
                    Buffer."Sr. No." := Z;
                    Buffer."Document No." := "Vendor Ledger Entry"."Document No.";
                    Buffer."Invoice No." := "Vendor Ledger Entry"."External Document No.";
                    Buffer.Date := "Vendor Ledger Entry"."Document Date";
                    //Buffer.Amount := ABS("Vendor Ledger Entry".Amount) + "Vendor Ledger Entry"."Total TDS Including SHE CESS";
                    Buffer."TDS Amount" := ABS("Vendor Ledger Entry"."Total TDS Including SHE CESS");
                    //Buffer."Net Amount" := ABS("Vendor Ledger Entry"."Closed by Amount (LCY)"); //RB-N 21Sep2017
                    //Buffer."Net Amount" := ABS("Vendor Ledger Entry"."Applied Amount of Bank Payment");\
                    //>>14Sep2018
                    Buffer.Amount := -1 * ("Vendor Ledger Entry".Amount) + "Vendor Ledger Entry"."Total TDS Including SHE CESS";
                    Buffer."Net Amount" := -1 * "Vendor Ledger Entry"."Closed by Amount (LCY)";
                    //<<14Sep2018
                    Buffer.INSERT;
                end;
            }
            dataitem("Report Buffer Table1"; "Report Buffer Table1")
            {
                DataItemTableView = SORTING("Sr. No.");
                column(SrNo; "Report Buffer Table1"."Sr. No.")
                {
                }
                column(DocumentNo; "Report Buffer Table1"."Document No.")
                {
                }
                column(NetAmount; "Report Buffer Table1"."Net Amount")
                {
                }
                column(InvoiceNo; "Report Buffer Table1"."Invoice No.")
                {
                }
                column(InvDate; "Report Buffer Table1".Date)
                {
                }
                column(TDSAmount; "Report Buffer Table1"."TDS Amount")
                {
                }
                column(Amount; "Report Buffer Table1".Amount)
                {
                }
                column(AppliestoAmt; ABS("Vendor Ledger Entry"."Applied Amount of Bank Payment"))
                {
                }

                trigger OnAfterGetRecord()
                begin

                    IF SourceDesc = 'CONTRA' THEN
                        CurrReport.SKIP;
                    IF "Report Buffer Table1"."Document No." = '' THEN
                        CurrReport.SKIP;

                    TotalAmount := TotalAmount + Buffer.Amount;
                    TotalTDSAmount := TotalTDSAmount + Buffer."TDS Amount";
                    TotalNetAmount := TotalNetAmount + Buffer."Net Amount";
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //>>RSPL/Rahul/PostMigration***Code added to block interim entries
                IF ("G/L Entry"."G/L Account No." = '12912050') OR ("G/L Entry"."G/L Account No." = '12912060') THEN
                    //("G/L Entry"."G/L Account No." = '75001090') OR ("G/L Entry"."G/L Account No." = '75001090')THEN    //To be view in future
                    CurrReport.SKIP;
                //<<
                GLAccName := FindGLAccName("Source Type", "Entry No.", "Source No.", "G/L Account No.");

                //IF vUserSetup.GET(USERID) THEN
                IF vUserSetup.GET("G/L Entry"."User ID") THEN //RB-N 21Sep2017
                    UserName := vUserSetup.Name;

                RespCtr.RESET;
                IF "G/L Entry"."Global Dimension 2 Code" <> '' THEN
                    RespCtr.GET("G/L Entry"."Global Dimension 2 Code");

                IF Amount < 0 THEN BEGIN
                    CrText := 'To';
                    DrText := '';
                END ELSE BEGIN
                    CrText := '';
                    DrText := 'Dr';
                END;

                SourceDesc := '';
                VendorInvoNo := '';
                IF "Source Code" <> '' THEN BEGIN
                    SourceCode.GET("Source Code");
                    SourceDesc := SourceCode.Description;
                END;
                //>>RSPL/Migration

                IF "G/L Entry"."Source Code" = 'PURCHASES' THEN BEGIN
                    IF VendorSource.GET("G/L Entry"."Source No.") THEN
                        SourceDesc := VendorSource."Vendor Posting Group";
                    IF PostedPInv.GET("G/L Entry"."Document No.") THEN
                        VendorInvoNo := PostedPInv."Vendor Invoice No." + ',     Date :- ' + FORMAT(PostedPInv."Document Date")
                    ELSE
                        VendorInvoNo := '';
                END;//<<
                PageLoop := PageLoop - 1;
                LinesPrinted := LinesPrinted + 1;

                ChequeNo := '';
                ChequeDate := 0D;
                IF ("Source No." <> '') AND ("Source Type" = "Source Type"::"Bank Account") THEN BEGIN
                    IF BankAccLedgEntry.GET("Entry No.") THEN BEGIN
                        ChequeNo := BankAccLedgEntry."Cheque No.";
                        ChequeDate := BankAccLedgEntry."Cheque Date";
                    END;
                END;

                IF (ChequeNo <> '') AND (ChequeDate <> 0D) THEN BEGIN
                    PageLoop := PageLoop - 1;
                    LinesPrinted := LinesPrinted + 1;
                END;
                IF PostingDate <> "Posting Date" THEN BEGIN
                    PostingDate := "Posting Date";
                    TotalDebitAmt := 0;
                END;
                IF DocumentNo <> "Document No." THEN BEGIN
                    DocumentNo := "Document No.";
                    TotalDebitAmt := 0;
                END;

                IF PostingDate = "Posting Date" THEN BEGIN
                    InitTextVariable;
                    TotalDebitAmt += "Debit Amount";
                    FormatNoText(NumberText, ABS(TotalDebitAmt), '');
                    PageLoop := NUMLines;
                    LinesPrinted := 0;
                END;
                IF (PrePostingDate <> "Posting Date") OR (PreDocumentNo <> "Document No.") THEN BEGIN
                    DebitAmountTotal := 0;
                    CreditAmountTotal := 0;
                    PrePostingDate := "Posting Date";
                    PreDocumentNo := "Document No.";
                END;

                DebitAmountTotal := DebitAmountTotal + "Debit Amount";
                CreditAmountTotal := CreditAmountTotal + "Credit Amount";


                //EBT STIVAN ---(18062012)--- To Capture Description as Narration in case of PLA Entries -------START
                CLEAR(PLADescription);
                /* recPLAEntry.RESET;
                recPLAEntry.SETRANGE(recPLAEntry."Document No.", "G/L Entry"."Document No.");
                IF recPLAEntry.FINDFIRST THEN BEGIN */
                PLADescription := 'Narration: ' + '';//recPLAEntry.Description;
                                                     //  END;
                                                     //EBT STIVAN ---(18062012)--- To Capture Description as Narration in case of PLA Entries ---------END

                //>>Robosoft\rahul

                DimensionValue := '';
                StaffDimension := '';
                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'DIVISION');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DimensionValue := DimSetEntry."Dimension Value Code";
                END;
                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'STAFF');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    Dimen.RESET;
                    Dimen.SETRANGE("Dimension Code", 'STAFF');
                    Dimen.SETRANGE(Code, DimSetEntry."Dimension Value Code");
                    IF Dimen.FINDFIRST THEN
                        StaffDimension := DimSetEntry."Dimension Value Code";
                END;
                //<<


                //>>RSPL/Migration/Rahul**Code added to capture dimensions

                CLEAR(StaffCode);
                CLEAR(StaffName);
                CLEAR(DepositEMDcode);
                CLEAR(DepositEMDName);
                CLEAR(DepositDealerCode);
                CLEAR(DepositDealerName);
                CLEAR(DepositOthersCode);
                CLEAR(DepositOthersName);
                CLEAR(DepositRentCode);
                CLEAR(DepositRentName);
                CLEAR(PhoneNo);
                CLEAR(PhoneNoDesc);
                CLEAR(VehicleCode);
                CLEAR(VehicaleName);

                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'STAFF');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    StaffCode := DimSetEntry."Dimension Value Code";
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    StaffName := DimSetEntry."Dimension Value Name";
                END;

                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'DEPOSIT-EMD');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DepositEMDcode := DimSetEntry."Dimension Value Code";
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    DepositEMDName := DimSetEntry."Dimension Value Name";
                END;

                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'DEPOSIT-DEALER');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DepositDealerCode := DimSetEntry."Dimension Value Code";
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    DepositDealerName := DimSetEntry."Dimension Value Name";
                END;

                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'DEPOSIT-OTHERS');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DepositOthersCode := DimSetEntry."Dimension Value Code";
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    DepositOthersName := DimSetEntry."Dimension Value Name";
                END;

                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'DEPOSIT-RENT');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DepositRentCode := DimSetEntry."Dimension Value Code";
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    DepositRentName := DimSetEntry."Dimension Value Name";
                END;

                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'PHONE');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    PhoneNo := DimSetEntry."Dimension Value Code";
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    PhoneNoDesc := DimSetEntry."Dimension Value Name";
                END;

                DimSetEntry.RESET;
                DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                DimSetEntry.SETRANGE("Dimension Code", 'VEHICLE');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    VehicleCode := DimSetEntry."Dimension Value Code";
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    VehicaleName := DimSetEntry."Dimension Value Name";
                END;

                //<<
            end;

            trigger OnPreDataItem()
            begin
                NUMLines := 13;
                PageLoop := NUMLines;
                LinesPrinted := 0;
                DebitAmountTotal := 0;
                CreditAmountTotal := 0;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                group(Options)
                {
                    Caption = 'Options';
                    field(PrintLineNarration; PrintLineNarration)
                    {
                        Caption = 'PrintLineNarration';
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

    trigger OnPreReport()
    begin
        CompanyInformation.GET;

        VendorInvoNo := '';
        Buffer.DELETEALL;
        Z := 0;
    end;

    var
        CompanyInformation: Record 79;
        SourceCode: Record 230;
        GLEntry: Record 17;
        BankAccLedgEntry: Record 271;
        GLAccName: Text[80];
        SourceDesc: Text[80];
        CrText: Text[2];
        DrText: Text[2];
        NumberText: array[2] of Text[80];
        PageLoop: Integer;
        LinesPrinted: Integer;
        NUMLines: Integer;
        ChequeNo: Code[80];
        ChequeDate: Date;
        Text16526: Label 'ZERO';
        Text16527: Label 'HUNDRED';
        Text16528: Label 'AND';
        Text16529: Label '%1 results in a written number that is too long.';
        Text16532: Label 'ONE';
        Text16533: Label 'TWO';
        Text16534: Label 'THREE';
        Text16535: Label 'FOUR';
        Text16536: Label 'FIVE';
        Text16537: Label 'SIX';
        Text16538: Label 'SEVEN';
        Text16539: Label 'EIGHT';
        Text16540: Label 'NINE';
        Text16541: Label 'TEN';
        Text16542: Label 'ELEVEN';
        Text16543: Label 'TWELVE';
        Text16544: Label 'THIRTEEN';
        Text16545: Label 'FOURTEEN';
        Text16546: Label 'FIFTEEN';
        Text16547: Label 'SIXTEEN';
        Text16548: Label 'SEVENTEEN';
        Text16549: Label 'EIGHTEEN';
        Text16550: Label 'NINETEEN';
        Text16551: Label 'TWENTY';
        Text16552: Label 'THIRTY';
        Text16553: Label 'FORTY';
        Text16554: Label 'FIFTY';
        Text16555: Label 'SIXTY';
        Text16556: Label 'SEVENTY';
        Text16557: Label 'EIGHTY';
        Text16558: Label 'NINETY';
        Text16559: Label 'THOUSAND';
        Text16560: Label 'MILLION';
        Text16561: Label 'BILLION';
        Text16562: Label 'LAKH';
        Text16563: Label 'CRORE';
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        PrintLineNarration: Boolean;
        PostingDate: Date;
        TotalDebitAmt: Decimal;
        DocumentNo: Code[20];
        DebitAmountTotal: Decimal;
        CreditAmountTotal: Decimal;
        PrePostingDate: Date;
        PreDocumentNo: Code[80];
        VoucherNoCaptionLbl: Label 'Voucher No. :';
        CreditAmountCaptionLbl: Label 'Credit Amount';
        DebitAmountCaptionLbl: Label 'Debit Amount';
        ParticularsCaptionLbl: Label 'Particulars';
        AmountInWordsCaptionLbl: Label 'Amount (in words):';
        PreparedByCaptionLbl: Label 'Prepared by:';
        CheckedByCaptionLbl: Label 'Checked by:';
        ApprovedByCaptionLbl: Label 'Approved by:';
        IntegerOccurcesCaptionLbl: Label 'IntegerOccurces';
        NarrationCaptionLbl: Label 'Narration :';
        "---": Integer;
        TotalAmount: Decimal;
        TotalTDSAmount: Decimal;
        TotalNetAmount: Decimal;
        Buffer: Record 50019;
        v: Integer;
        vUser: Record 2000000120;
        VendorSource: Record 23;
        PostedPInv: Record 122;
        VendorInvoNo: Code[80];
        PLADescription: Text[100];
        //recPLAEntry: Record 13723;
        DimSetEntry: Record 480;
        DimensionValue: Text[80];
        StaffDimension: Text[80];
        Dimen: Record 349;
        RespCtr: Record 5714;
        StaffCode: Code[10];
        StaffName: Text[80];
        RecDimensionValue: Record 349;
        DepositEMDcode: Code[10];
        DepositEMDName: Text[80];
        RecDimensionValue1: Record 349;
        DepositDealerCode: Code[10];
        DepositDealerName: Text[80];
        PhoneNo: Code[20];
        PhoneNoDesc: Text[80];
        DepositOthersCode: Code[10];
        DepositOthersName: Text[80];
        DepositRentCode: Code[10];
        DepositRentName: Text[80];
        VehicleCode: Code[20];
        VehicaleName: Text[80];
        UserName: Text;
        vUserSetup: Record 91;
        Z: Integer;
        "-----16Mar2018": Integer;
        DVLE: Record 380;

    //[Scope('Internal')]
    procedure FindGLAccName("Source Type": Option " ",Customer,Vendor,"Bank Account","Fixed Asset"; "Entry No.": Integer; "Source No.": Code[20]; "G/L Account No.": Code[20]): Text[80]
    var
        AccName: Text[80];
        VendLedgerEntry: Record 25;
        Vend: Record 23;
        CustLedgerEntry: Record 21;
        Cust: Record 18;
        BankLedgerEntry: Record 271;
        Bank: Record 270;
        FALedgerEntry: Record 5601;
        FA: Record 5600;
        GLAccount: Record 15;
    begin
        IF "Source Type" = "Source Type"::Vendor THEN
            IF VendLedgerEntry.GET("Entry No.") THEN BEGIN
                Vend.GET("Source No.");
                //AccName := Vend.Name;
                AccName := Vend."No." + ' - ' + Vend.Name;//RSPLSUM 20Feb2020
            END ELSE BEGIN
                GLAccount.GET("G/L Account No.");
                AccName := GLAccount.Name;
            END
        ELSE
            IF "Source Type" = "Source Type"::Customer THEN
                IF CustLedgerEntry.GET("Entry No.") THEN BEGIN
                    Cust.GET("Source No.");
                    AccName := Cust.Name;
                END ELSE BEGIN
                    GLAccount.GET("G/L Account No.");
                    AccName := GLAccount.Name;
                END
            ELSE
                IF "Source Type" = "Source Type"::"Bank Account" THEN
                    IF BankLedgerEntry.GET("Entry No.") THEN BEGIN
                        Bank.GET("Source No.");
                        AccName := Bank.Name;
                    END ELSE BEGIN
                        GLAccount.GET("G/L Account No.");
                        AccName := GLAccount.Name;
                    END
                ELSE BEGIN
                    GLAccount.GET("G/L Account No.");
                    AccName := GLAccount.Name;
                END;

        IF "Source Type" = "Source Type"::" " THEN BEGIN
            GLAccount.GET("G/L Account No.");
            AccName := GLAccount.Name;
        END;

        EXIT(AccName);
    end;

    //[Scope('Internal')]
    procedure FormatNoText(var NoText: array[2] of Text[80]; No: Decimal; CurrencyCode: Code[10])
    var
        PrintExponent: Boolean;
        Ones: Integer;
        Tens: Integer;
        Hundreds: Integer;
        Exponent: Integer;
        NoTextIndex: Integer;
        Currency: Record 4;
        TensDec: Integer;
        OnesDec: Integer;
    begin
        CLEAR(NoText);
        NoTextIndex := 1;
        NoText[1] := '';

        IF No < 1 THEN
            AddToNoText(NoText, NoTextIndex, PrintExponent, Text16526)
        ELSE BEGIN
            FOR Exponent := 4 DOWNTO 1 DO BEGIN
                PrintExponent := FALSE;
                IF No > 99999 THEN BEGIN
                    Ones := No DIV (POWER(100, Exponent - 1) * 10);
                    Hundreds := 0;
                END ELSE BEGIN
                    Ones := No DIV POWER(1000, Exponent - 1);
                    Hundreds := Ones DIV 100;
                END;
                Tens := (Ones MOD 100) DIV 10;
                Ones := Ones MOD 10;
                IF Hundreds > 0 THEN BEGIN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Hundreds]);
                    AddToNoText(NoText, NoTextIndex, PrintExponent, Text16527);
                END;
                IF Tens >= 2 THEN BEGIN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, TensText[Tens]);
                    IF Ones > 0 THEN
                        AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Ones]);
                END ELSE
                    IF (Tens * 10 + Ones) > 0 THEN
                        AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Tens * 10 + Ones]);
                IF PrintExponent AND (Exponent > 1) THEN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, ExponentText[Exponent]);
                IF No > 99999 THEN
                    No := No - (Hundreds * 100 + Tens * 10 + Ones) * POWER(100, Exponent - 1) * 10
                ELSE
                    No := No - (Hundreds * 100 + Tens * 10 + Ones) * POWER(1000, Exponent - 1);
            END;
        END;

        IF CurrencyCode <> '' THEN BEGIN
            Currency.GET(CurrencyCode);
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + '');//Currency."Currency Numeric Description");
        END ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, 'RUPEES');

        AddToNoText(NoText, NoTextIndex, PrintExponent, Text16528);

        TensDec := ((No * 100) MOD 100) DIV 10;
        OnesDec := (No * 100) MOD 10;
        IF TensDec >= 2 THEN BEGIN
            AddToNoText(NoText, NoTextIndex, PrintExponent, TensText[TensDec]);
            IF OnesDec > 0 THEN
                AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[OnesDec]);
        END ELSE
            IF (TensDec * 10 + OnesDec) > 0 THEN
                AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[TensDec * 10 + OnesDec])
            ELSE
                AddToNoText(NoText, NoTextIndex, PrintExponent, Text16526);
        IF (CurrencyCode <> '') THEN
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + '')//Currency."Currency Decimal Description" + ' ONLY')
        ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' PAISA ONLY');
    end;

    local procedure AddToNoText(var NoText: array[2] of Text[80]; var NoTextIndex: Integer; var PrintExponent: Boolean; AddText: Text[30])
    begin
        PrintExponent := TRUE;

        WHILE STRLEN(NoText[NoTextIndex] + ' ' + AddText) > MAXSTRLEN(NoText[1]) DO BEGIN
            NoTextIndex := NoTextIndex + 1;
            IF NoTextIndex > ARRAYLEN(NoText) THEN
                ERROR(Text16529, AddText);
        END;

        NoText[NoTextIndex] := DELCHR(NoText[NoTextIndex] + ' ' + AddText, '<');
    end;

    // [Scope('Internal')]
    procedure InitTextVariable()
    begin
        OnesText[1] := Text16532;
        OnesText[2] := Text16533;
        OnesText[3] := Text16534;
        OnesText[4] := Text16535;
        OnesText[5] := Text16536;
        OnesText[6] := Text16537;
        OnesText[7] := Text16538;
        OnesText[8] := Text16539;
        OnesText[9] := Text16540;
        OnesText[10] := Text16541;
        OnesText[11] := Text16542;
        OnesText[12] := Text16543;
        OnesText[13] := Text16544;
        OnesText[14] := Text16545;
        OnesText[15] := Text16546;
        OnesText[16] := Text16547;
        OnesText[17] := Text16548;
        OnesText[18] := Text16549;
        OnesText[19] := Text16550;

        TensText[1] := '';
        TensText[2] := Text16551;
        TensText[3] := Text16552;
        TensText[4] := Text16553;
        TensText[5] := Text16554;
        TensText[6] := Text16555;
        TensText[7] := Text16556;
        TensText[8] := Text16557;
        TensText[9] := Text16558;

        ExponentText[1] := '';
        ExponentText[2] := Text16559;
        ExponentText[3] := Text16562;
        ExponentText[4] := Text16563;
    end;
}


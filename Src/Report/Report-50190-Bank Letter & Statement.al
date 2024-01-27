report 50190 "Bank Letter & Statement"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BankLetterStatement.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Bank Account Ledger Entry"; "Bank Account Ledger Entry")
        {
            RequestFilterFields = "Document No.";
            column(FullName_recBank; recBank."Full Name")
            {
            }
            column(DocumentNo_BankAccountLedgerEntry; "Bank Account Ledger Entry"."Document No.")
            {
            }
            column(PostingDate_BankAccountLedgerEntry; "Bank Account Ledger Entry"."Posting Date")
            {
            }
            column(Addr_recBank; recBank.Address)
            {
            }
            column(Addr2_recBank; recBank."Address 2")
            {
            }
            column(City_recBank; recBank.City + ' - ' + recBank."Post Code")
            {
            }
            column(AccNo_recBank; recBank."Bank Account No.")
            {
            }

            trigger OnAfterGetRecord()
            begin
                recBank.GET("Bank Account Ledger Entry"."Bank Account No.");
            end;

            trigger OnPostDataItem()
            begin
                //CurrReport.BREAK;
            end;
        }
        dataitem("Bank Account Ledger Entry1"; "Bank Account Ledger Entry")
        {
            DataItemTableView = SORTING("Document No.", "Posting Date");
            RequestFilterFields = "Document No.";
            dataitem("G/L Entry"; "G/L Entry")
            {
                DataItemLink = "Document No." = FIELD("Document No.");
                DataItemTableView = SORTING("Posting Date", "Document No.")
                                    WHERE("System-Created Entry" = FILTER(false),
                                          "Document No." = FILTER(<> ''));
                column(DocumentNo_GLEntry; "G/L Entry"."Document No.")
                {
                }
                column(StaffName; UPPERCASE(StaffName))
                {
                }
                column(BankAccountNo; "BankAccountNo.")
                {
                }
                column(Amount_GLEntry; "G/L Entry".Amount)
                {
                }
                column(TotalAmountinWords; ('Rs.' + TotalAmountinWords[1]))
                {
                }
                column(PostingDate_GLEntry; "G/L Entry"."Posting Date")
                {
                }
                column(GTCode; GTCode)
                {
                }
                column(BankName; BankName)
                {
                }
                column(IFSCCode; IFSCCode)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //StaffName := '';
                    "BankAccountNo." := '';
                    CLEAR(BankName);//RSPLSUM 15Nov2019
                    CLEAR(IFSCCode);//RSPLSUM 15Nov2019
                                    //recLedEntryDim.RESET;
                                    //recLedEntryDim.SETRANGE(recLedEntryDim."Table ID",17);
                                    //recLedEntryDim.SETRANGE(recLedEntryDim."Entry No.","G/L Entry"."Entry No.");
                                    //recLedEntryDim.SETRANGE(recLedEntryDim."Dimension Code",'STAFF');
                                    //IF recLedEntryDim.FINDFIRST THEN/
                                    //BEGIN
                                    /*
                                     recDimValue.RESET;
                                     recDimValue.SETRANGE(recDimValue."Dimension Code",'STAFF');
                                     //recDimValue.SETRANGE(recDimValue.Code,recLedEntryDim."Dimension Value Code");
                                     IF recDimValue.FINDFIRST THEN
                                     BEGIN
                                      StaffName := recDimValue.Name;
                                     END;
                                    */

                    CLEAR(GTCode);//ss
                    DimSetEntry.RESET;
                    DimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                    DimSetEntry.SETRANGE("Dimension Code", 'STAFF');
                    IF DimSetEntry.FINDFIRST THEN BEGIN
                        DimSetEntry.CALCFIELDS("Dimension Value Name");
                        StaffName := DimSetEntry."Dimension Value Name";
                        GTCode := DimSetEntry."Dimension Value Code";//ss
                    END;

                    recEmployee.RESET;
                    //recEmployee.SETRANGE(recEmployee."No.",recLedEntryDim."Dimension Value Code");
                    recEmployee.SETRANGE("No.", DimSetEntry."Dimension Value Code");
                    IF recEmployee.FINDFIRST THEN BEGIN
                        "BankAccountNo." := recEmployee."Bank Account No.";
                        BankName := recEmployee."Bank Name";//RSPLSUM 15Nov2019
                        IFSCCode := recEmployee."IFSC Code";//RSPLSUM 15Nov2019
                    END;
                    //END;
                    TotalAmount += "G/L Entry".Amount;
                    SrNo := SrNo + 1;

                    IF PrintToExcel THEN
                        MakeEcelDataBody;

                end;

                trigger OnPostDataItem()
                begin

                    IF PrintToExcel THEN
                        MakeEcelDataFooter;

                    FormatNoText(TotalAmountinWords, ROUND((Amount), 0.01), '');
                end;
            }

            trigger OnAfterGetRecord()
            begin
                i += 1;
                IF i = 1 THEN BEGIN
                    IF PrintToExcel THEN
                        MakeExcelDataHeader1;
                END
            end;

            trigger OnPreDataItem()
            begin
                "Bank Account Ledger Entry1".SETRANGE("Bank Account Ledger Entry1"."Document No.", "Bank Account Ledger Entry"."Document No.");
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(PrintToExcel; PrintToExcel)
                {
                    ApplicationArea = all;
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

    trigger OnInitReport()
    begin

        InitTextVariable;

        SrNo := 0;
    end;

    trigger OnPostReport()
    begin

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin
        i := 0;
    end;

    var
        Reference: Code[30];
        SrNo: Integer;
        recDimValue: Record 349;
        StaffName: Text[50];
        recGLentry: Record "G/L Entry";
        recEmployee: Record Employee;
        "BankAccountNo.": Code[15];
        DocumentNoFilter: Code[20];
        recBank: Record "Bank Account";
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        TotalAmountinWords: array[2] of Text[250];
        ExcelBuf: Record 370 temporary;
        CompInfo: Record 79;
        PrintToExcel: Boolean;
        Text026: Label 'Zero';
        Text027: Label 'Hundred';
        Text028: Label '&';
        Text029: Label '%1 results in a written number that is too long.';
        Text032: Label 'One';
        Text033: Label 'Two';
        Text034: Label 'Three';
        Text035: Label 'Four';
        Text036: Label 'Five';
        Text037: Label 'Six';
        Text038: Label 'Seven';
        Text039: Label 'Eight';
        Text040: Label 'Nine';
        Text041: Label 'Ten';
        Text042: Label 'Eleven';
        Text043: Label 'Twelve';
        Text044: Label 'Thirteen';
        Text045: Label 'Fourteen';
        Text046: Label 'Fifteen';
        Text047: Label 'Sixteen';
        Text048: Label 'Seventeen';
        Text049: Label 'Eighteen';
        Text050: Label 'Nineteen';
        Text051: Label 'Twenty';
        Text052: Label 'Thirty';
        Text053: Label 'Forty';
        Text054: Label 'Fifty';
        Text055: Label 'Sixty';
        Text056: Label 'Seventy';
        Text057: Label 'Eighty';
        Text058: Label 'Ninety';
        Text059: Label 'Thousand';
        Text060: Label 'Million';
        Text061: Label 'Billion';
        Text1280000: Label 'Lakh';
        Text1280001: Label 'Crore';
        Text0000: Label 'Bank Letter & Statement';
        Text0001: Label 'Bank Letter & Statement';
        "--": Integer;
        DimSetEntry: Record 480;
        GTCode: Code[20];
        BankName: Text[50];
        IFSCCode: Code[20];
        TotalAmount: Decimal;
        i: Integer;
        Text063: Label 'Dear Sir/Madam,';
        Text064: Label 'Thanking You';
        Text065: Label 'Yours Very truly,';
        Text066: Label 'For GP PETROLEUMS LIMITED';
        Text067: Label 'Authorised Signatory';
        Text068: Label 'Kindly  credit  our  staff  account  as  per  the  detailed  statement  below towards  reimbursement  of their  expenses  and  debit  our  account  no.';

    // //[Scope('Internal')]
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
            AddToNoText(NoText, NoTextIndex, PrintExponent, Text026)
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
                    AddToNoText(NoText, NoTextIndex, PrintExponent, Text027);
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
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + '' /*Currency."Currency Numeric Description"*/);
        END ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, 'Rupees');

        AddToNoText(NoText, NoTextIndex, PrintExponent, Text028);

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
                AddToNoText(NoText, NoTextIndex, PrintExponent, Text026);
        IF (CurrencyCode <> '') THEN
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + /*Currency."Currency Decimal Description"*/'' + ' ONLY')
        ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, 'Paisa Only');
    end;

    local procedure AddToNoText(var NoText: array[2] of Text[80]; var NoTextIndex: Integer; var PrintExponent: Boolean; AddText: Text[30])
    begin
        PrintExponent := TRUE;

        WHILE STRLEN(NoText[NoTextIndex] + ' ' + AddText) > MAXSTRLEN(NoText[1]) DO BEGIN
            NoTextIndex := NoTextIndex + 1;
            IF NoTextIndex > ARRAYLEN(NoText) THEN
                ERROR(Text029, AddText);
        END;

        NoText[NoTextIndex] := DELCHR(NoText[NoTextIndex] + ' ' + AddText, '<');
    end;

    // //[Scope('Internal')]
    procedure InitTextVariable()
    begin
        OnesText[1] := Text032;
        OnesText[2] := Text033;
        OnesText[3] := Text034;
        OnesText[4] := Text035;
        OnesText[5] := Text036;
        OnesText[6] := Text037;
        OnesText[7] := Text038;
        OnesText[8] := Text039;
        OnesText[9] := Text040;
        OnesText[10] := Text041;
        OnesText[11] := Text042;
        OnesText[12] := Text043;
        OnesText[13] := Text044;
        OnesText[14] := Text045;
        OnesText[15] := Text046;
        OnesText[16] := Text047;
        OnesText[17] := Text048;
        OnesText[18] := Text049;
        OnesText[19] := Text050;

        TensText[1] := '';
        TensText[2] := Text051;
        TensText[3] := Text052;
        TensText[4] := Text053;
        TensText[5] := Text054;
        TensText[6] := Text055;
        TensText[7] := Text056;
        TensText[8] := Text057;
        TensText[9] := Text058;

        ExponentText[1] := '';
        ExponentText[2] := Text059;
        ExponentText[3] := Text1280000;
        ExponentText[4] := Text1280001;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;vb
        // ExcelBuf.CreateSheet(Text0000,CompInfo.Name,Text0001,USERID);
        /*  ExcelBuf.CreateBookAndOpenExcel('', Text0000, '', '', '');
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text0000);
        ExcelBuf.WriteSheet(Text0000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text0000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader1()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Ref. No. :', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Bank Account Ledger Entry1"."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date :', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT("Bank Account Ledger Entry1"."Posting Date", 10, '<Month,2>/<Day,2>/<Year4>'), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn(recBank.Name,FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(recBank."Full Name", FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 06Dec19
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(recBank.Address, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(recBank."Address 2", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(recBank.City + '-' + recBank."Post Code", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text063, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        //--RSPLSUM 06Dec19--ExcelBuf.AddColumn(Text068+' '+recBank."No.",FALSE,'',FALSE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Text068 + ' ' + recBank."Bank Account No.", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 06Dec19
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GT Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 06Dec19
        ExcelBuf.AddColumn('Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Bank Account No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Bank Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('IFSC Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure MakeEcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(SrNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(GTCode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 06Dec19
        ExcelBuf.AddColumn(UPPERCASE(StaffName), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("BankAccountNo.", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Entry".Amount, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(BankName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(IFSCCode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure MakeEcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPLSUM 06Dec19
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TotalAmount, FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text064, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text065, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text066, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text067, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
    end;
}


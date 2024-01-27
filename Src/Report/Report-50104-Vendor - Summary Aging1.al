report 50104 "Vendor - Summary Aging1"
{
    // EBT STIVAN  Dated-22/10/11 Done Excel Coding
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/VendorSummaryAging1.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Vendor; Vendor)
        {
            DataItemTableView = SORTING("No.")
                                WHERE(Blocked = FILTER(' '));
            RequestFilterFields = "No.", "Search Name", "Vendor Posting Group", "Currency Filter";

            trigger OnAfterGetRecord()
            begin

                paymnettermsrec.SETRANGE(paymnettermsrec.Code, Vendor."Payment Terms Code");
                IF paymnettermsrec.FINDFIRST THEN
                    PeriodStartDate[1] := 20010101D;//010101D;
                DateExp := ('-' + FORMAT(paymnettermsrec."Due Date Calculation"));
                testdate := CALCDATE(DateExp, EndingDate);

                VendorLedgerEntry.RESET;
                VendorLedgerEntry.SETCURRENTKEY("Vendor No.", "Posting Date");
                VendorLedgerEntry.SETRANGE("Vendor No.", "No.");
                VendorLedgerEntry.SETRANGE("Posting Date", 0D, EndingDate);
                IF VendorLedgerEntry.FIND('-') THEN
                    REPEAT
                        DtldVendorLedgerEntry.RESET;
                        DtldVendorLedgerEntry.SETCURRENTKEY("Vendor Ledger Entry No.", "Posting Date");
                        DtldVendorLedgerEntry.SETRANGE("Vendor Ledger Entry No.", VendorLedgerEntry."Entry No.");

                        IF DtldVendorLedgerEntry.FIND('-') THEN
                            REPEAT
                                IF DtldVendorLedgerEntry."Posting Date" <= EndingDate THEN BEGIN
                                    IF (testdate - VendorLedgerEntry."Posting Date") <= 0 THEN
                                        VendBalanceDue[5] := VendBalanceDue[5] + DtldVendorLedgerEntry."Amount (LCY)"
                                    ELSE
                                        IF ((testdate - VendorLedgerEntry."Posting Date") >= 1) AND ((testdate - VendorLedgerEntry."Posting Date") <= 30) THEN
                                            VendBalanceDue[4] := VendBalanceDue[4] + DtldVendorLedgerEntry."Amount (LCY)"
                                        ELSE
                                            IF ((testdate - VendorLedgerEntry."Posting Date") >= 31) AND ((testdate - VendorLedgerEntry."Posting Date") <= 60) THEN
                                                VendBalanceDue[3] := VendBalanceDue[3] + DtldVendorLedgerEntry."Amount (LCY)"
                                            ELSE
                                                IF ((testdate - VendorLedgerEntry."Posting Date") >= 61) AND ((testdate - VendorLedgerEntry."Posting Date") <= 90) THEN
                                                    VendBalanceDue[2] := VendBalanceDue[2] + DtldVendorLedgerEntry."Amount (LCY)"
                                                ELSE
                                                    IF ((testdate - VendorLedgerEntry."Posting Date") >= 91) THEN
                                                        VendBalanceDue[1] := VendBalanceDue[1] + DtldVendorLedgerEntry."Amount (LCY)";
                                END;
                            UNTIL DtldVendorLedgerEntry.NEXT = 0;
                    UNTIL VendorLedgerEntry.NEXT = 0;


                LineTotalVendAmountDue := VendBalanceDue[1] + VendBalanceDue[2] +
                               VendBalanceDue[3] + VendBalanceDue[4] + VendBalanceDue[5];

                IF LineTotalVendAmountDue = 0 THEN
                    CurrReport.SKIP;


                IF PrintToExcel THEN
                    MakeExcelDataBody;
            end;

            trigger OnPostDataItem()
            begin

                IF PrintToExcel THEN
                    MakeExcelDataFooter;
            end;

            trigger OnPreDataItem()
            begin
                LineTotalVendAmountDue := 0;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(PeriodStartDate; PeriodStartDate[2])
                {
                    ApplicationArea = ALL;
                    Caption = 'Starting Date';
                }
                field(EndingDate; EndingDate)
                {
                    ApplicationArea = ALL;
                    Caption = 'Ending Date';
                }
                field(PeriodLength; PeriodLength)
                {
                    ApplicationArea = ALL;
                    Caption = 'Period Length';
                }
                field(PrintAmountsInLCY; PrintAmountsInLCY)
                {
                    ApplicationArea = ALL;
                    Caption = 'Show Amounts in LCY';
                }
                field(HeadingType; HeadingType)
                {
                    ApplicationArea = ALL;
                    Caption = 'Heading Type';
                }
                field(PrintToExcel; PrintToExcel)
                {
                    ApplicationArea = ALL;
                    Caption = 'Print To Excel';
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
        PeriodStartDate[5] := EndingDate;
        PeriodStartDate[6] := 99981231D;//31129998D;
        FOR i := 4 DOWNTO 2 DO
            PeriodStartDate[i]
            := CALCDATE('<-30D>', PeriodStartDate[i + 1]);

        CalcDates;
        CreateHeadings;

        IF PrintToExcel THEN
            MakeExcelInfo;
    end;

    var
        Currency: Record 4;
        Currency2: Record 4 temporary;
        PrintAmountsInLCY: Boolean;
        VendFilter: Text[250];
        PeriodStartDate: array[6] of Date;
        LineTotalVendAmountDue: Decimal;
        TotalVendAmtDueLCY: Decimal;
        VendBalanceDue: array[5] of Decimal;
        VendBalanceDueLCY: array[5] of Decimal;
        PeriodLength: DateFormula;
        PrintLine: Boolean;
        i: Integer;
        EndingDate: Date;
        AgingBy: Option "Due Date","Posting Date","Document Date";
        HeadingType: Option "Date Interval","Number of Days";
        PeriodEndDate: array[5] of Date;
        HeaderText: array[5] of Text[40];
        CurrencyCode: Code[10];
        NumberOfCurrencies: Integer;
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        DateExp: Text[50];
        paymnettermsrec: Record 3;
        testdate: Date;
        VendorLedgerEntry: Record 25;
        DtldVendorLedgerEntry: Record 380;
        Text000: Label 'Not Due';
        Text001: Label 'Vendor - Summary Aging';
        Text002: Label 'days';
        Text003: Label 'More than';
        Text004: Label 'Aged by %1';
        Text005: Label 'Total for %1';
        Text006: Label 'Aged as of %1';
        Text007: Label 'Aged by %1';
        Text008: Label 'All Amounts in LCY';
        Text009: Label 'Due Date,Posting Date,Document Date';
        Text010: Label 'The Date Formula %1 cannot be used. Try to restate it. E.g. 1M+CM instead of CM+1M.';
        Text011: Label 'Company Name';
        Text012: Label 'Report Name';
        Text013: Label 'Report No.';
        Text014: Label 'USER ID';
        Text015: Label 'Date';
        "--10Apr2017": Integer;
        TVendBalanceDue: array[5] of Decimal;
        TLineTotalVendAmountDue: Decimal;

    local procedure CalcDates()
    var
        i: Integer;
        PeriodLength2: DateFormula;
    begin
        EVALUATE(PeriodLength2, '-' + FORMAT(PeriodLength));
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

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text011), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ;
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text012), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('Vendor - Summary Aging', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ;
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text013), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"Vendor - Summary Aging1", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ;
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text014), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ;
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text015), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ;
        ExcelBuf.ClearNewRow;

        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        /*  ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', '');
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
        //>>10APr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text001, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8


        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter : ' + FORMAT(PeriodStartDate[2]) + '..' + FORMAT(EndingDate), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8


        //<<10APr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('NAME', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn('NOT DUE', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn('0-30 DAYS', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn('31-60 DAYS', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn('61-90 DAYS', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn('OVER 90 DAYS', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//7
        ExcelBuf.AddColumn('BALANCE', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//8
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Vendor."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn(Vendor.Name, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn(VendBalanceDue[1], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//3
        TVendBalanceDue[1] += VendBalanceDue[1];//10Apr2017
        VendBalanceDue[1] := 0;

        ExcelBuf.AddColumn(VendBalanceDue[2], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//4
        TVendBalanceDue[2] += VendBalanceDue[2];//10Apr2017
        VendBalanceDue[2] := 0;

        ExcelBuf.AddColumn(VendBalanceDue[3], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//5
        TVendBalanceDue[3] += VendBalanceDue[3];//10Apr2017
        VendBalanceDue[3] := 0;

        ExcelBuf.AddColumn(VendBalanceDue[4], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//6
        TVendBalanceDue[4] += VendBalanceDue[4];//10Apr2017
        VendBalanceDue[4] := 0;

        ExcelBuf.AddColumn(VendBalanceDue[5], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//7
        TVendBalanceDue[5] += VendBalanceDue[5];//10Apr2017
        VendBalanceDue[5] := 0;

        ExcelBuf.AddColumn(LineTotalVendAmountDue, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//8
        TLineTotalVendAmountDue += LineTotalVendAmountDue;//10Apr2017
        LineTotalVendAmountDue := 0;
    end;

    //  //[Scope('Internal')]
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
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//2
        //ExcelBuf.AddColumn(VendBalanceDue[1],FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//3 //10Apr2017
        ExcelBuf.AddColumn(TVendBalanceDue[1], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//3 //10Apr2017
        ExcelBuf.AddColumn(TVendBalanceDue[2], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//4 //10Apr2017

        ExcelBuf.AddColumn(TVendBalanceDue[3], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//5 //10Apr2017
        ExcelBuf.AddColumn(TVendBalanceDue[4], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//6 //10Apr2017
        ExcelBuf.AddColumn(TVendBalanceDue[5], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//7 //10Apr2017
        ExcelBuf.AddColumn(TLineTotalVendAmountDue, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//8 //10Apr2017
    end;
}


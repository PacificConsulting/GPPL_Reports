report 50101 "Customer - Aging  Summary"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CustomerAgingSummary.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Customer; Customer)
        {
            DataItemTableView = SORTING("Global Dimension 2 Code");
            RequestFilterFields = "No.", "Customer Posting Group", "Payment Terms Code", "Global Dimension 2 Code", "Salesperson Code";
            column(CompName; companyinfo.Name)
            {
            }
            column(DateFilter; STRSUBSTNO(Text001, FORMAT(StartDate)) + '....' + FORMAT(EndDate))
            {
            }
            column(HDate; FORMAT(TODAY, 0, 4))
            {
            }
            column(CustGlobalCode; "Global Dimension 2 Code")
            {
            }
            column(CustNo; Customer."No.")
            {
            }
            column(CustName; Customer.Name)
            {
            }
            column(resposiblityvalue; resposiblityvalue)
            {
            }
            column(Region; Region)
            {
            }
            column(salespersonvalue; salespersonvalue)
            {
            }
            column(creditvalue; creditvalue)
            {
            }
            column(totalcustbal; totalcustbal)
            {
            }
            column(CustBal5; CustBalanceDueLCY[5])
            {
            }
            column(CustBal4; CustBalanceDueLCY[4])
            {
            }
            column(CustBal3; CustBalanceDueLCY[3])
            {
            }
            column(CustBal2; CustBalanceDueLCY[2])
            {
            }
            column(CustBal1; CustBalanceDueLCY[1])
            {
            }

            trigger OnAfterGetRecord()
            begin

                TCOUNT += 1;//03Apr2017

                //>>DataGrouping 03Apr2017
                CUST03.SETCURRENTKEY("Global Dimension 2 Code");
                CUST03.SETRANGE("Global Dimension 2 Code", "Global Dimension 2 Code");
                IF CUST03.FINDLAST THEN BEGIN
                    SCOUNT := CUST03.COUNT;
                    //MESSAGE('No. of TCOUNT %1 \\ SCOUNT %2 \\ GGG %3',TCOUNT,SCOUNT,"Global Dimension 2 Code");
                END;
                //<<DataGrouping 03Apr2017


                //>>1

                //usersetup.GET(USERID);

                // to get dimension name
                dimvalue.SETRANGE(dimvalue.Code, Customer."Global Dimension 2 Code");
                IF dimvalue.FINDFIRST THEN
                    resposiblityvalue := dimvalue.Name;

                dimvalue.SETRANGE(dimvalue.Code, Customer."Global Dimension 2 Code");
                IF dimvalue.FINDFIRST THEN
                    resposiblityvalue := dimvalue.Name;

                // to get salesperson name
                CLEAR(ZoneCode);//RSPLSUM 16Apr2020
                "Sales/purchaser".SETRANGE("Sales/purchaser".Code, Customer."Salesperson Code");
                IF "Sales/purchaser".FINDFIRST THEN BEGIN
                    salespersonvalue := "Sales/purchaser".Name;
                    ZoneCode := "Sales/purchaser"."Zone Code";//RSPLSUM 16Apr2020
                END;//RSPLSUM 16Apr2020

                //to get payment terms name
                paymnettermsrec.SETRANGE(paymnettermsrec.Code, Customer."Payment Terms Code");
                IF paymnettermsrec.FINDFIRST THEN
                    creditvalue := paymnettermsrec.Description;

                PeriodStartDate[1] := 20010101D/*010101D*/;
                DateExp := ('-' + FORMAT(paymnettermsrec."Due Date Calculation"));
                testdate := CALCDATE(DateExp, EndDate);

                CLEAR(Region);
                CLEAR(RegionCode);
                recCust.RESET;
                recCust.SETRANGE(recCust."No.", "No.");
                IF recCust.FINDFIRST THEN BEGIN
                    recState.RESET;
                    recState.SETRANGE(recState.Code, recCust."State Code");
                    IF recState.FINDFIRST THEN BEGIN
                        Region := 'Region' + ': ' + '';//recState."Region Code";
                        RegionCode := '';// recState."Region Code";
                    END;
                END;

                //set filter on posting date and then loop through detailed cust. ledger entry table based on cust. ledger table
                //then split the balance based on the aging
                CLEAR(CustBalanceDueLCY[5]);//03Apr2017
                CLEAR(CustBalanceDueLCY[4]);//03Apr2017
                CLEAR(CustBalanceDueLCY[3]);//03Apr2017
                CLEAR(CustBalanceDueLCY[2]);//03Apr2017
                CLEAR(CustBalanceDueLCY[1]);//03Apr2017
                CLEAR(totalcustbal);//03Apr2017

                CustLedgEntry.RESET;
                CustLedgEntry.SETCURRENTKEY("Customer No.", "Posting Date");
                CustLedgEntry.SETRANGE("Customer No.", "No.");
                CustLedgEntry.SETRANGE("Posting Date", 0D, EndDate);
                IF CustLedgEntry.FIND('-') THEN
                    REPEAT
                        DtldCustLedgEntry.RESET;
                        DtldCustLedgEntry.SETCURRENTKEY("Cust. Ledger Entry No.", "Posting Date");
                        DtldCustLedgEntry.SETRANGE("Cust. Ledger Entry No.", CustLedgEntry."Entry No.");

                        IF DtldCustLedgEntry.FIND('-') THEN
                            REPEAT
                                IF DtldCustLedgEntry."Posting Date" <= EndDate THEN BEGIN
                                    IF (testdate - CustLedgEntry."Posting Date") <= 0 THEN
                                        CustBalanceDueLCY[5] := CustBalanceDueLCY[5] + DtldCustLedgEntry."Amount (LCY)"
                                    ELSE
                                        IF ((testdate - CustLedgEntry."Posting Date") >= 1) AND ((testdate - CustLedgEntry."Posting Date") <= 30) THEN
                                            CustBalanceDueLCY[4] := CustBalanceDueLCY[4] + DtldCustLedgEntry."Amount (LCY)"
                                        ELSE
                                            IF ((testdate - CustLedgEntry."Posting Date") >= 31) AND ((testdate - CustLedgEntry."Posting Date") <= 60) THEN
                                                CustBalanceDueLCY[3] := CustBalanceDueLCY[3] + DtldCustLedgEntry."Amount (LCY)"
                                            ELSE
                                                IF ((testdate - CustLedgEntry."Posting Date") >= 61) AND ((testdate - CustLedgEntry."Posting Date") <= 90) THEN
                                                    CustBalanceDueLCY[2] := CustBalanceDueLCY[2] + DtldCustLedgEntry."Amount (LCY)"
                                                ELSE
                                                    IF ((testdate - CustLedgEntry."Posting Date") >= 91) THEN
                                                        CustBalanceDueLCY[1] := CustBalanceDueLCY[1] + DtldCustLedgEntry."Amount (LCY)";
                                END;
                            UNTIL DtldCustLedgEntry.NEXT = 0;
                    UNTIL CustLedgEntry.NEXT = 0;

                totalcustbal := CustBalanceDueLCY[1] + CustBalanceDueLCY[2] +
                               CustBalanceDueLCY[3] + CustBalanceDueLCY[4] + CustBalanceDueLCY[5];


                CLEAR(NAH);
                IF totalcustbal = 0 THEN
                    NAH := TRUE;//03Apr2017
                                //CurrReport.SKIP;
                                //<<1


                //>>2

                //Customer, Body (6) - OnPreSection()
                IF PrintToExcel THEN
                    MakeExcelDataBody;
                //<<2


                //>>03Apr2017
                Gtotalcustbal += totalcustbal;
                Ttotalcustbal += totalcustbal;

                //<<03Apr2017

                //>>3

                //Customer, GroupFooter (7) - OnPreSection()
                //>>03Apr2017
                CLEAR(NAH1);
                IF Gtotalcustbal = 0 THEN
                    NAH1 := TRUE;

                //<<03Apr2017

                IF (TCOUNT = SCOUNT) AND PrintToExcel THEN BEGIN

                    MakeExcelGroupFooter;
                    TCOUNT := 0;
                END;
                //<<3
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                //Customer, Footer (9) - OnPreSection()
                IF PrintToExcel THEN
                    MakeExcelDataFooter;
                //<<1
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                //CurrReport.CREATETOTALS(CustBalanceDueLCY,totalcustbal);
                companyinfo.GET;
                totalcustbal := 0;
                //<<1

                CUST03.COPYFILTERS(Customer);//03Apr2017

                TCOUNT := 0;//03Apr2017
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; StartDate)
                {
                    ApplicationArea = all;
                }
                field("End Date"; EndDate)
                {
                    ApplicationArea = all;
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = all;
                }
            }
        }

        actions
        {
        }

        trigger OnOpenPage()
        begin
            //>>1
            IF StartDate = 0D THEN
                StartDate := WORKDATE;

            AccountingPeriod.SETFILTER("Starting Date", '<=%1', WORKDATE);
            AccountingPeriod.SETRANGE("New Fiscal Year", TRUE);
            IF AccountingPeriod.FIND('+') THEN
                //PeriodStartingDate := AccountingPeriod."Starting Date";
                StartDate := AccountingPeriod."Starting Date";
            //<<1
        end;
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

        //>>1
        /*
        user.GET(USERID);
        IF (user."User ID" = 'CF24') OR (user."User ID" = 'CF09') OR (user."User ID" = 'CF11') OR (user."User ID" = 'CF12') OR
        (user."User ID" = 'CF16') OR (user."User ID" = 'CF22') OR (user."User ID" = 'CF29') OR (user."User ID" = 'CF31') OR
        (user."User ID" = 'CF32') OR (user."User ID" = 'CF34') OR (user."User ID" = 'CF35') OR (user."User ID" = 'CF36') OR
        (user."User ID" = 'CF40') OR (user."User ID" = 'CF41') OR (user."User ID" = 'CF42') OR (user."User ID" = 'CF43') OR
        (user."User ID" = 'CF44') OR (user."User ID" = 'CF46') OR (user."User ID" = 'CF47') OR (user."User ID" = 'CF48') OR
        (user."User ID" = 'CF49') OR (user."User ID" = 'CF50') OR (user."User ID" = 'CF51') OR (user."User ID" = 'CF52') OR
        (user."User ID" = 'CF53') OR (user."User ID" = 'CF55') THEN
        ERROR('You are not allowed to run this Report');
        *///Commented UserTable

        CustFilter := Customer.GETFILTERS;

        PeriodStartDate[5] := EndDate;
        PeriodStartDate[6] := 99981231D;//31129998D;
        FOR i := 4 DOWNTO 2 DO
            PeriodStartDate[i]
            := CALCDATE('<-30D>', PeriodStartDate[i + 1]);


        IF PrintToExcel THEN
            MakeExcelInfo;

        //<<1

    end;

    var
        DtldCustLedgEntry: Record 379;
        StartDate: Date;
        CustFilter: Text[250];
        PeriodStartDate: array[6] of Date;
        CustBalanceDueLCY: array[5] of Decimal;
        PrintCust: Boolean;
        i: Integer;
        companyinfo: Record 79;
        totalcustbal: Decimal;
        "Sales/purchaser": Record 13;
        resposiblityctr: Record 5714;
        resposiblityvalue: Text[50];
        salespersonvalue: Text[50];
        creditvalue: Text[50];
        paymnettermsrec: Record 3;
        dimvalue: Record 349;
        AccountingPeriod: Record 50;
        EndDate: Date;
        DateExp: Text[40];
        CustLedgEntry: Record 21;
        TempCustLedgEntry: Record 21 temporary;
        tempDtldCustLedgEntry: Record 379 temporary;
        test: Text[30];
        testdate: Date;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        usersetup: Record 91;
        user: Record 2000000120;
        recCust: Record 18;
        recState: Record State;
        Region: Text[30];
        RegionCode: Text[30];
        Text001: Label 'As of %1';
        Text000: Label 'Data';
        Text0001: Label 'Customer - Aging  Summary';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "---03Apr2017": Integer;
        CUST03: Record 18;
        TCOUNT: Integer;
        SCOUNT: Integer;
        Gtotalcustbal: Decimal;
        Ttotalcustbal: Decimal;
        GCustBalanceDueLCY: array[5] of Decimal;
        TCustBalanceDueLCY: array[5] of Decimal;
        NAH: Boolean;
        NAH1: Boolean;
        ZoneCode: Code[10];

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Customer - Aging  Summary', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Customer - Aging  Summary", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>3Apr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//12

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Customer - Aging Summary', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.NewRow;

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        //ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter : ' + FORMAT(StartDate) + '..' + FORMAT(EndDate), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        //>>3Apr2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12//RSPLSUM 16Apr2020

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('NAME', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('SALESPERSON CODE', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Region', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Zone', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('CREDIT PERIOD', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('OUTSTANDING', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('NOT DUE', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('0-30 DAYS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('31-60 DAYS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('61-90 DAYS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('OVER 90 DAYS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        IF NAH = FALSE THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn(Customer."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
            ExcelBuf.AddColumn(Customer.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
            ExcelBuf.AddColumn(salespersonvalue, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
            ExcelBuf.AddColumn(RegionCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
            ExcelBuf.AddColumn(ZoneCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5//RSPLSUM 16Apr2020
            ExcelBuf.AddColumn(creditvalue, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
            ExcelBuf.AddColumn(totalcustbal, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//7


            ExcelBuf.AddColumn(CustBalanceDueLCY[5], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//8
                                                                                                     //>>03Apr2017
            GCustBalanceDueLCY[5] += CustBalanceDueLCY[5];
            TCustBalanceDueLCY[5] += CustBalanceDueLCY[5];
            //>>03Apr2017

            ExcelBuf.AddColumn(CustBalanceDueLCY[4], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//9
                                                                                                     //>>03Apr2017
            GCustBalanceDueLCY[4] += CustBalanceDueLCY[4];
            TCustBalanceDueLCY[4] += CustBalanceDueLCY[4];
            //>>03Apr2017

            ExcelBuf.AddColumn(CustBalanceDueLCY[3], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//10
                                                                                                     //>>03Apr2017
            GCustBalanceDueLCY[3] += CustBalanceDueLCY[3];
            TCustBalanceDueLCY[3] += CustBalanceDueLCY[3];
            //>>03Apr2017

            ExcelBuf.AddColumn(CustBalanceDueLCY[2], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//11
                                                                                                     //>>03Apr2017
            GCustBalanceDueLCY[2] += CustBalanceDueLCY[2];
            TCustBalanceDueLCY[2] += CustBalanceDueLCY[2];
            //>>03Apr2017

            ExcelBuf.AddColumn(CustBalanceDueLCY[1], FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//12
                                                                                                     //>>03Apr2017
            GCustBalanceDueLCY[1] += CustBalanceDueLCY[1];
            TCustBalanceDueLCY[1] += CustBalanceDueLCY[1];
            //>>03Apr2017

            //ExcelBuf.AddColumn(Customer."Global Dimension 2 Code"+'--'+resposiblityvalue,FALSE,'',TRUE,TRUE,TRUE,'@',1);//test
        END;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>03Apr2017 RB-N
        /* ExcelBuf.CreateBook('', Text0001);
        ExcelBuf.CreateBookAndOpenExcel('', Text0001, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 16Apr2020
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6//RSPLSUM 16Apr2020
        //ExcelBuf.AddColumn(totalcustbal,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//6
        //ExcelBuf.AddColumn(CustBalanceDueLCY[5],FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//7
        //ExcelBuf.AddColumn(CustBalanceDueLCY[4],FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//8
        //ExcelBuf.AddColumn(CustBalanceDueLCY[3],FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//9
        //ExcelBuf.AddColumn(CustBalanceDueLCY[2],FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//10
        //ExcelBuf.AddColumn(CustBalanceDueLCY[1],FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//11

        //>>GrandTotal 03Apr2017
        ExcelBuf.AddColumn(Ttotalcustbal, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//7
        ExcelBuf.AddColumn(TCustBalanceDueLCY[5], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//8
        ExcelBuf.AddColumn(TCustBalanceDueLCY[4], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//9
        ExcelBuf.AddColumn(TCustBalanceDueLCY[3], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//10
        ExcelBuf.AddColumn(TCustBalanceDueLCY[2], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//11
        ExcelBuf.AddColumn(TCustBalanceDueLCY[1], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//12
        //<<GrandTotal 03Apr2017
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelGroupFooter()
    begin
        IF NAH1 = FALSE THEN BEGIN

            ExcelBuf.NewRow;
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 16Apr2020
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
            ExcelBuf.AddColumn('TOTAL FOR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
            ExcelBuf.AddColumn(resposiblityvalue, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6//RSPLSUM 16Apr2020
                                                                         //ExcelBuf.AddColumn(totalcustbal,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//6
            ExcelBuf.AddColumn(Gtotalcustbal, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//7
            Gtotalcustbal := 0;//03Apr2017

            ExcelBuf.AddColumn(GCustBalanceDueLCY[5], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//8
            GCustBalanceDueLCY[5] := 0;//03Apr2017

            ExcelBuf.AddColumn(GCustBalanceDueLCY[4], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//9
            GCustBalanceDueLCY[4] := 0;//03Apr2017

            ExcelBuf.AddColumn(GCustBalanceDueLCY[3], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//10
            GCustBalanceDueLCY[3] := 0;//03Apr2017

            ExcelBuf.AddColumn(GCustBalanceDueLCY[2], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//11
            GCustBalanceDueLCY[2] := 0;//03Apr2017

            ExcelBuf.AddColumn(GCustBalanceDueLCY[1], FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//12
            GCustBalanceDueLCY[1] := 0;//03Apr2017

        END;
    end;
}


report 50153 "Overdue Reminder"
{
    // 01Aug2017 ::RB-N, DueDate Calculation
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/OverdueReminder.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Customer; Customer)
        {
            DataItemTableView = SORTING("No.")
                                WHERE(Blocked = FILTER(<> All));
            RequestFilterFields = "No.", Type;
            column(ComName; CompnayInfo.Name)
            {
            }
            column(No_Customer; Customer."No.")
            {
            }
            column(Name_Customer; Customer.Name)
            {
            }
            column(Address_Customer; Customer.Address)
            {
            }
            column(Address2_Customer; Customer."Address 2")
            {
            }
            column(City_Customer; Customer.City)
            {
            }
            column(PostCode_Customer; Customer."Post Code")
            {
            }
            column(vAsonDate; vAsonDate)
            {
            }
            dataitem("Cust. Ledger Entry"; "Cust. Ledger Entry")
            {
                DataItemLink = "Customer No." = FIELD("No.");
                DataItemTableView = SORTING("Customer No.", "Posting Date")
                                    ORDER(Ascending)
                                    WHERE("Document Type" = FILTER(Invoice),
                                          "Remaining Amt. (LCY)" = FILTER(<> 0));
                column(EntryNo; "Entry No.")
                {
                }
                column(DueDate01; DueDate01)
                {
                }
                column(RemainingAmount_CustLedgerEntry; "Cust. Ledger Entry"."Remaining Amount")
                {
                }
                column(AmountLCY_CustLedgerEntry; "Cust. Ledger Entry"."Amount (LCY)")
                {
                }
                column(DueDate_CustLedgerEntry; "Cust. Ledger Entry"."Due Date")
                {
                }
                column(PostingDate_CustLedgerEntry; "Cust. Ledger Entry"."Posting Date")
                {
                }
                column(DocumentNo_CustLedgerEntry; "Cust. Ledger Entry"."Document No.")
                {
                }
                column(vBillalreadydue; vBillalreadydue)
                {
                }
                column(vBillsbecomingdue; vBillsbecomingdue)
                {
                }
                column(vBillbecomenextcycle; vBillbecomenextcycle)
                {
                }
                column(vAsMonthDescription; 'Bills becoming due in ' + vAsMonthDescription)
                {
                }
                column(vAsYear; ' Total Remmitance for ' + vAsMonthDescription + ' ' + FORMAT(vAsYear))
                {
                }
                column(PgNo; PgN)
                {
                }
                column(vBillalreadydue11; "Cust. Ledger Entry"."Remaining Amt. (LCY)")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    "Cust. Ledger Entry".CALCFIELDS("Amount (LCY)");
                    "Cust. Ledger Entry".CALCFIELDS("Remaining Amt. (LCY)");

                    //>>01Aug2017 DueDate
                    CLEAR(DueDate01);
                    //DueDate01 := CALCDATE(Customer."Approved Payment Days","Due Date");
                    DueDate01 := CALCDATE('0D', "Due Date");//05Aug2017

                    //<<01Aug2017 DueDate


                    CLEAR(vBillalreadydue);//01Aug2017
                    //IF "Cust. Ledger Entry"."Due Date" <= vLastEnd THEN
                    //IF "Cust. Ledger Entry"."Due Date" <vAsonDate THEN
                    IF DueDate01 < vAsonDate THEN //01Aug2017
                        vBillalreadydue := "Remaining Amt. (LCY)";

                    CLEAR(vBillsbecomingdue);//01Aug2017
                    //IF ("Cust. Ledger Entry"."Due Date" >= vAsonDate) AND ("Cust. Ledger Entry"."Due Date" <= vEndDateofason)  THEN
                    IF (DueDate01 >= vAsonDate) AND (DueDate01 <= vEndDateofason) THEN //01Aug2017
                        vBillsbecomingdue := "Remaining Amt. (LCY)";

                    CLEAR(vBillbecomenextcycle);//01Aug2017
                    //IF ("Cust. Ledger Entry"."Due Date" > vEndDateofason) THEN
                    IF (DueDate01 > vEndDateofason) THEN
                        vBillbecomenextcycle := "Remaining Amt. (LCY)";

                    PgN := CurrReport.PAGENO;
                end;

                trigger OnPreDataItem()
                begin

                    LastFieldNo := FIELDNO("Entry No.");
                    //CurrReport.CREATETOTALS(vBillalreadydue,vBillsbecomingdue,vBillbecomenextcycle);
                    vBillalreadydue := 0;
                    vBillsbecomingdue := 0;
                    vBillbecomenextcycle := 0;
                    //"Cust. Ledger Entry".SETRANGE("Cust. Ledger Entry"."Due Date",vMontSrat,vMontEnd);
                    CLEAR(vBillsbecomingdue);
                    CLEAR(vBillbecomenextcycle);
                    CLEAR(vBillalreadydue);
                end;
            }

            trigger OnAfterGetRecord()
            begin

                Customer.CALCFIELDS(Balance);
                IF Balance = 0 THEN
                    CurrReport.SKIP;

                IF recCountry.GET(Customer."Country/Region Code") THEN;
                IF recState.GET(Customer."State Code") THEN;
            end;

            trigger OnPreDataItem()
            begin
                CompnayInfo.GET;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(vAsonDate; vAsonDate)
                {
                    ApplicationArea = all;
                    Caption = 'As on Date';
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

        //vMonth:=DATE2DMY(WORKDATE,2);
        //vYear:=DATE2DMY(WORKDATE,3);

        //07-01-14
        IF vAsonDate = 0D THEN
            ERROR('Please Mention As on Date');

        vAsDate := DATE2DMY(vAsonDate, 1);
        vAsMonth := DATE2DMY(vAsonDate, 2);
        vAsYear := DATE2DMY(vAsonDate, 3);


        IF (vAsMonth = 1) OR (vAsMonth = 3) OR (vAsMonth = 5) OR (vAsMonth = 7)
              OR (vAsMonth = 8) OR (vAsMonth = 10) OR (vAsMonth = 12) THEN
            vDays := 31
        ELSE
            IF (vAsMonth = 4) OR (vAsMonth = 6) OR (vAsMonth = 9) OR (vAsMonth = 11) THEN
                vDays := 30
            ELSE
                IF (vAsMonth = 2) THEN
                    vDays := 28;

        vEndDateofason := DMY2DATE(vDays, vAsMonth, vAsYear);
        //MESSAGE('%1 EndDate',vEndDateofason);

        IF vAsMonth = 1 THEN
            vAsMonthDescription := 'January'
        ELSE
            IF vAsMonth = 2 THEN
                vAsMonthDescription := 'February'
            ELSE
                IF vAsMonth = 3 THEN
                    vAsMonthDescription := 'March'
                ELSE
                    IF vAsMonth = 4 THEN
                        vAsMonthDescription := 'April'
                    ELSE
                        IF vAsMonth = 5 THEN
                            vAsMonthDescription := 'May'
                        ELSE
                            IF vAsMonth = 6 THEN
                                vAsMonthDescription := 'June'
                            ELSE
                                IF vAsMonth = 7 THEN
                                    vAsMonthDescription := 'July'
                                ELSE
                                    IF vAsMonth = 8 THEN
                                        vAsMonthDescription := 'August'
                                    ELSE
                                        IF vAsMonth = 9 THEN
                                            vAsMonthDescription := 'September'
                                        ELSE
                                            IF vAsMonth = 10 THEN
                                                vAsMonthDescription := 'October'
                                            ELSE
                                                IF vAsMonth = 11 THEN
                                                    vAsMonthDescription := 'November'
                                                ELSE
                                                    IF vAsMonth = 12 THEN
                                                        vAsMonthDescription := 'December';

        //07-01-14
        /*
        IF vStartMonth=0 THEN
            vMonth:=1 ;
        IF vStartMonth=1 THEN
            vMonth:=2;
        IF vStartMonth=2 THEN
            vMonth:=3;
        IF vStartMonth=3 THEN
            vMonth:=4 ;
        IF vStartMonth=4 THEN
            vMonth:=5;
        IF vStartMonth=5 THEN
            vMonth:=6 ;
        IF vStartMonth=6 THEN
            vMonth:=7 ;
        IF vStartMonth=7 THEN
            vMonth:=8;
        IF vStartMonth=8 THEN
            vMonth:=9;
        IF vStartMonth=9 THEN
            vMonth:=10 ;
        IF vStartMonth=10 THEN
            vMonth:=11;
        IF vStartMonth=11 THEN
            vMonth:=12;
        
        
        vStratDate:=DMY2DATE(1,vMonth,vStartYear);
        vLastEnd:=(vStratDate-1);
        
        
        IF (vStartMonth=0) OR (vStartMonth=2) OR (vStartMonth=4) OR (vStartMonth=6)
              OR (vStartMonth=8) OR (vStartMonth=10) OR (vStartMonth=12) THEN
           vDay:=31
          ELSE
           IF (vStartMonth=3) OR (vStartMonth=5) OR (vStartMonth=8) OR (vStartMonth=10) THEN
               vDay:=30
           ELSE
            IF (vStartMonth=1) THEN
                vDay:=28;
        
        vMontSrat1:=(vStratDate-1);
        vMontSrat2:=DATE2DMY(vMontSrat1,2);
        vMontSrat3:=DATE2DMY(vMontSrat1,3);
        vMontSrat:=DMY2DATE(1,vMontSrat2,vMontSrat3);
        
        IF vStartMonth=11 THEN
          vEndDate:=DMY2DATE(31,1,vStartYear+1)
        ELSE IF vStartMonth=0 THEN
           vEndDate:=DMY2DATE(28,vMonth+1,vStartYear)
        ELSE IF vStartMonth=1 THEN
           vEndDate:=DMY2DATE(31,vMonth+1,vStartYear)
        ELSE IF vStartMonth=2 THEN
           vEndDate:=DMY2DATE(30,vMonth+1,vStartYear)
        ELSE IF vStartMonth=3 THEN
           vEndDate:=DMY2DATE(31,vMonth+1,vStartYear)
        ELSE IF vStartMonth=4 THEN
           vEndDate:=DMY2DATE(30,vMonth+1,vStartYear)
        ELSE IF vStartMonth=5 THEN
           vEndDate:=DMY2DATE(31,vMonth+1,vStartYear)
        ELSE IF vStartMonth=6 THEN
           vEndDate:=DMY2DATE(31,vMonth+1,vStartYear)
        ELSE IF vStartMonth=7 THEN
           vEndDate:=DMY2DATE(30,vMonth+1,vStartYear)
        ELSE IF vStartMonth=8 THEN
           vEndDate:=DMY2DATE(31,vMonth+1,vStartYear)
        ELSE IF vStartMonth=9 THEN
           vEndDate:=DMY2DATE(30,vMonth+1,vStartYear)
        ELSE IF vStartMonth=10 THEN
           vEndDate:=DMY2DATE(31,vMonth+1,vStartYear);
        
        vMontEnd:=vEndDate;
        
        vEndDate1:=DATE2DMY(vEndDate,2);
        vEndDate2:=DATE2DMY(vEndDate,3);
        vEndDate3:=DMY2DATE(1,vEndDate1,vEndDate2);
        
        
        IF vStartMonth=0  THEN
           vMonthDes:='Jan'
          ELSE IF vStartMonth=1  THEN
           vMonthDes:='Feb'
          ELSE IF vStartMonth=2  THEN
           vMonthDes:='March'
          ELSE IF vStartMonth=3  THEN
           vMonthDes:='April'
          ELSE IF vStartMonth=4  THEN
           vMonthDes:='May'
          ELSE IF vStartMonth=5  THEN
           vMonthDes:='June'
          ELSE IF vStartMonth=6  THEN
           vMonthDes:='July'
          ELSE IF vStartMonth=7  THEN
           vMonthDes:='Aug'
          ELSE IF vStartMonth=8  THEN
            vMonthDes:='Sept'
          ELSE IF vStartMonth=9  THEN
           vMonthDes:='Oct'
          ELSE IF vStartMonth=10  THEN
           vMonthDes:='Nov'
          ELSE IF vStartMonth=11  THEN
           vMonthDes:='Dec';
        
        
        IF vStartMonth=0  THEN
           vMonthDescription:='January'
          ELSE IF vStartMonth=1  THEN
           vMonthDescription:='February'
          ELSE IF vStartMonth=2  THEN
           vMonthDescription:='March'
          ELSE IF vStartMonth=3  THEN
           vMonthDescription:='April'
          ELSE IF vStartMonth=4  THEN
           vMonthDescription:='May'
          ELSE IF vStartMonth=5  THEN
           vMonthDescription:='June'
          ELSE IF vStartMonth=6  THEN
           vMonthDescription:='July'
          ELSE IF vStartMonth=7  THEN
           vMonthDescription:='August'
          ELSE IF vStartMonth=8  THEN
            vMonthDescription:='September'
          ELSE IF vStartMonth=9  THEN
           vMonthDescription:='October'
          ELSE IF vStartMonth=10  THEN
           vMonthDescription:='November'
          ELSE IF vStartMonth=11  THEN
           vMonthDescription:='December';
        */

    end;

    var
        CompnayInfo: Record 79;
        pgno: Integer;
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        vBillalreadydue: Decimal;
        vBillsbecomingdue: Decimal;
        vBillbecomenextcycle: Decimal;
        vDay: Integer;
        vMonth: Integer;
        vYear: Integer;
        vDate: Integer;
        vStratDate: Date;
        vEndDate: Date;
        vLastEnd: Date;
        vMonthDes: Text[50];
        vStartMonth: Option Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec;
        vStartYear: Integer;
        vMontSrat: Date;
        vMontSrat1: Date;
        vMontSrat2: Integer;
        vMontSrat3: Integer;
        vMontEnd: Date;
        vEndDate1: Integer;
        vEndDate2: Integer;
        vEndDate3: Date;
        recCountry: Record 9;
        recState: Record State;
        vCustomerNo: Code[20];
        vMonthDescription: Code[20];
        vAsonDate: Date;
        vAsDate: Integer;
        vAsMonth: Integer;
        vAsMonthDescription: Text[30];
        vAsYear: Integer;
        vDays: Integer;
        vEndDateofason: Date;
        "---Robosoft---": Integer;
        PgN: Integer;
        "---01Aug2017": Integer;
        DueDate01: Date;

    //   //[Scope('Internal')]
    procedure Month()
    begin
    end;
}


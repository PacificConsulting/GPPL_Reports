report 50210 "Foreign Currency Balance - New"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 30Jun2018   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ForeignCurrencyBalanceNew.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem(Customer; Customer)
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Currency Code" = FILTER(<> ''));
            column(No_Customer; Customer."No.")
            {
            }
            column(Name_Customer; Customer.Name)
            {
            }
            column(CurrencyCode_Customer; Customer."Currency Code")
            {
            }
            column(totalcustbal; totalcustbal)
            {
            }
            column(EDate; EDate)
            {
            }
            column(CustBalUSD; CustBalUSD)
            {
            }
            column(CustBalOth; CustBalOth)
            {
            }
            column(EYear2; EYear2)
            {
            }
            column(EYear3; EYear3)
            {
            }
            column(EYear4; EYear4)
            {
            }
            column(EYear5; EYear5)
            {
            }
            column(CustBalUSD2; CustBalUSD2)
            {
            }
            column(CustBalUSD3; CustBalUSD3)
            {
            }
            column(CustBalUSD4; CustBalUSD4)
            {
            }
            column(CustBalUSD5; CustBalUSD5)
            {
            }
            column(CustBalOth2; CustBalOth2)
            {
            }
            column(CustBalOth3; CustBalOth3)
            {
            }
            column(CustBalOth4; CustBalOth4)
            {
            }
            column(CustBalOth5; CustBalOth5)
            {
            }

            trigger OnAfterGetRecord()
            begin

                CLEAR(ActualDate);
                PayTerm.RESET;
                IF PayTerm.GET("Payment Terms Code") THEN;

                DateExp := ('-' + FORMAT(PayTerm."Due Date Calculation"));
                ActualDate := CALCDATE(PayTerm."Due Date Calculation", EDate);

                CLEAR(totalcustbal);
                CLEAR(CustBalance);

                CLE.RESET;
                CLE.SETCURRENTKEY("Customer No.", "Posting Date");
                CLE.SETRANGE("Customer No.", "No.");
                CLE.SETRANGE("Posting Date", 0D, EDate);
                CLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                IF CLE.FINDFIRST THEN
                    REPEAT

                        DCLE.RESET;
                        DCLE.SETCURRENTKEY("Cust. Ledger Entry No.", "Posting Date");
                        DCLE.SETRANGE("Cust. Ledger Entry No.", CLE."Entry No.");
                        IF DCLE.FINDFIRST THEN
                            REPEAT
                                IF DCLE."Posting Date" <= EDate THEN BEGIN
                                    IF (ActualDate - CLE."Posting Date") <= 0 THEN
                                        CustBalance[5] += DCLE.Amount
                                    ELSE
                                        IF ((ActualDate - CLE."Posting Date") >= 1) AND ((ActualDate - CLE."Posting Date") <= 30) THEN
                                            CustBalance[4] += DCLE.Amount
                                        ELSE
                                            IF ((ActualDate - CLE."Posting Date") >= 31) AND ((ActualDate - CLE."Posting Date") <= 60) THEN
                                                CustBalance[3] += DCLE.Amount
                                            ELSE
                                                IF ((ActualDate - CLE."Posting Date") >= 61) AND ((ActualDate - CLE."Posting Date") <= 90) THEN
                                                    CustBalance[2] += DCLE.Amount
                                                ELSE
                                                    IF ((ActualDate - CLE."Posting Date") >= 91) THEN
                                                        CustBalance[1] += DCLE.Amount;
                                END;
                            UNTIL DCLE.NEXT = 0;
                    UNTIL CLE.NEXT = 0;

                totalcustbal := CustBalance[1] + CustBalance[2] + CustBalance[3] + CustBalance[4] + CustBalance[5];

                CLEAR(CustBalOth);
                CLEAR(CustBalUSD);

                IF "Currency Code" = 'USD' THEN
                    CustBalUSD := totalcustbal
                ELSE
                    CustBalOth := totalcustbal;

                //2nd Year
                CLEAR(totalcustbal2);
                CLEAR(CustBalOth2);
                CLEAR(CustBalUSD2);
                CLEAR(CustBalance);
                IF EYear2 <> 0D THEN BEGIN

                    ActualDate := CALCDATE(PayTerm."Due Date Calculation", EYear2);
                    CLEAR(CustBalance);

                    CLE.RESET;
                    CLE.SETCURRENTKEY("Customer No.", "Posting Date");
                    CLE.SETRANGE("Customer No.", "No.");
                    CLE.SETRANGE("Posting Date", 0D, EYear2);
                    CLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                    IF CLE.FINDFIRST THEN
                        REPEAT

                            DCLE.RESET;
                            DCLE.SETCURRENTKEY("Cust. Ledger Entry No.", "Posting Date");
                            DCLE.SETRANGE("Cust. Ledger Entry No.", CLE."Entry No.");
                            IF DCLE.FINDFIRST THEN
                                REPEAT
                                    IF DCLE."Posting Date" <= EYear2 THEN BEGIN
                                        IF (ActualDate - CLE."Posting Date") <= 0 THEN
                                            CustBalance[5] += DCLE.Amount
                                        ELSE
                                            IF ((ActualDate - CLE."Posting Date") >= 1) AND ((ActualDate - CLE."Posting Date") <= 30) THEN
                                                CustBalance[4] += DCLE.Amount
                                            ELSE
                                                IF ((ActualDate - CLE."Posting Date") >= 31) AND ((ActualDate - CLE."Posting Date") <= 60) THEN
                                                    CustBalance[3] += DCLE.Amount
                                                ELSE
                                                    IF ((ActualDate - CLE."Posting Date") >= 61) AND ((ActualDate - CLE."Posting Date") <= 90) THEN
                                                        CustBalance[2] += DCLE.Amount
                                                    ELSE
                                                        IF ((ActualDate - CLE."Posting Date") >= 91) THEN
                                                            CustBalance[1] += DCLE.Amount;
                                    END;
                                UNTIL DCLE.NEXT = 0;
                        UNTIL CLE.NEXT = 0;

                    totalcustbal2 := CustBalance[1] + CustBalance[2] + CustBalance[3] + CustBalance[4] + CustBalance[5];

                    IF "Currency Code" = 'USD' THEN
                        CustBalUSD2 := totalcustbal2
                    ELSE
                        CustBalOth2 := totalcustbal2;
                END;

                //3rd Year
                CLEAR(totalcustbal3);
                CLEAR(CustBalOth3);
                CLEAR(CustBalUSD3);
                CLEAR(CustBalance);
                IF EYear3 <> 0D THEN BEGIN

                    ActualDate := CALCDATE(PayTerm."Due Date Calculation", EYear3);
                    CLEAR(CustBalance);

                    CLE.RESET;
                    CLE.SETCURRENTKEY("Customer No.", "Posting Date");
                    CLE.SETRANGE("Customer No.", "No.");
                    CLE.SETRANGE("Posting Date", 0D, EYear3);
                    CLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                    IF CLE.FINDFIRST THEN
                        REPEAT

                            DCLE.RESET;
                            DCLE.SETCURRENTKEY("Cust. Ledger Entry No.", "Posting Date");
                            DCLE.SETRANGE("Cust. Ledger Entry No.", CLE."Entry No.");
                            IF DCLE.FINDFIRST THEN
                                REPEAT
                                    IF DCLE."Posting Date" <= EYear3 THEN BEGIN
                                        IF (ActualDate - CLE."Posting Date") <= 0 THEN
                                            CustBalance[5] += DCLE.Amount
                                        ELSE
                                            IF ((ActualDate - CLE."Posting Date") >= 1) AND ((ActualDate - CLE."Posting Date") <= 30) THEN
                                                CustBalance[4] += DCLE.Amount
                                            ELSE
                                                IF ((ActualDate - CLE."Posting Date") >= 31) AND ((ActualDate - CLE."Posting Date") <= 60) THEN
                                                    CustBalance[3] += DCLE.Amount
                                                ELSE
                                                    IF ((ActualDate - CLE."Posting Date") >= 61) AND ((ActualDate - CLE."Posting Date") <= 90) THEN
                                                        CustBalance[2] += DCLE.Amount
                                                    ELSE
                                                        IF ((ActualDate - CLE."Posting Date") >= 91) THEN
                                                            CustBalance[1] += DCLE.Amount;
                                    END;
                                UNTIL DCLE.NEXT = 0;
                        UNTIL CLE.NEXT = 0;

                    totalcustbal3 := CustBalance[1] + CustBalance[2] + CustBalance[3] + CustBalance[4] + CustBalance[5];

                    IF "Currency Code" = 'USD' THEN
                        CustBalUSD3 := totalcustbal3
                    ELSE
                        CustBalOth3 := totalcustbal3;
                END;

                //4th Year
                CLEAR(totalcustbal4);
                CLEAR(CustBalOth4);
                CLEAR(CustBalUSD4);
                CLEAR(CustBalance);
                IF EYear4 <> 0D THEN BEGIN

                    ActualDate := CALCDATE(PayTerm."Due Date Calculation", EYear4);
                    CLEAR(CustBalance);

                    CLE.RESET;
                    CLE.SETCURRENTKEY("Customer No.", "Posting Date");
                    CLE.SETRANGE("Customer No.", "No.");
                    CLE.SETRANGE("Posting Date", 0D, EYear4);
                    CLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                    IF CLE.FINDFIRST THEN
                        REPEAT

                            DCLE.RESET;
                            DCLE.SETCURRENTKEY("Cust. Ledger Entry No.", "Posting Date");
                            DCLE.SETRANGE("Cust. Ledger Entry No.", CLE."Entry No.");
                            IF DCLE.FINDFIRST THEN
                                REPEAT
                                    IF DCLE."Posting Date" <= EYear4 THEN BEGIN
                                        IF (ActualDate - CLE."Posting Date") <= 0 THEN
                                            CustBalance[5] += DCLE.Amount
                                        ELSE
                                            IF ((ActualDate - CLE."Posting Date") >= 1) AND ((ActualDate - CLE."Posting Date") <= 30) THEN
                                                CustBalance[4] += DCLE.Amount
                                            ELSE
                                                IF ((ActualDate - CLE."Posting Date") >= 31) AND ((ActualDate - CLE."Posting Date") <= 60) THEN
                                                    CustBalance[3] += DCLE.Amount
                                                ELSE
                                                    IF ((ActualDate - CLE."Posting Date") >= 61) AND ((ActualDate - CLE."Posting Date") <= 90) THEN
                                                        CustBalance[2] += DCLE.Amount
                                                    ELSE
                                                        IF ((ActualDate - CLE."Posting Date") >= 91) THEN
                                                            CustBalance[1] += DCLE.Amount;
                                    END;
                                UNTIL DCLE.NEXT = 0;
                        UNTIL CLE.NEXT = 0;

                    totalcustbal4 := CustBalance[1] + CustBalance[2] + CustBalance[3] + CustBalance[4] + CustBalance[5];

                    IF "Currency Code" = 'USD' THEN
                        CustBalUSD4 := totalcustbal4
                    ELSE
                        CustBalOth4 := totalcustbal4;
                END;


                //5th Year
                CLEAR(totalcustbal5);
                CLEAR(CustBalOth5);
                CLEAR(CustBalUSD5);
                CLEAR(CustBalance);
                IF EYear5 <> 0D THEN BEGIN

                    ActualDate := CALCDATE(PayTerm."Due Date Calculation", EYear5);
                    CLEAR(CustBalance);

                    CLE.RESET;
                    CLE.SETCURRENTKEY("Customer No.", "Posting Date");
                    CLE.SETRANGE("Customer No.", "No.");
                    CLE.SETRANGE("Posting Date", 0D, EYear5);
                    CLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                    IF CLE.FINDFIRST THEN
                        REPEAT

                            DCLE.RESET;
                            DCLE.SETCURRENTKEY("Cust. Ledger Entry No.", "Posting Date");
                            DCLE.SETRANGE("Cust. Ledger Entry No.", CLE."Entry No.");
                            IF DCLE.FINDFIRST THEN
                                REPEAT
                                    IF DCLE."Posting Date" <= EYear5 THEN BEGIN
                                        IF (ActualDate - CLE."Posting Date") <= 0 THEN
                                            CustBalance[5] += DCLE.Amount
                                        ELSE
                                            IF ((ActualDate - CLE."Posting Date") >= 1) AND ((ActualDate - CLE."Posting Date") <= 30) THEN
                                                CustBalance[4] += DCLE.Amount
                                            ELSE
                                                IF ((ActualDate - CLE."Posting Date") >= 31) AND ((ActualDate - CLE."Posting Date") <= 60) THEN
                                                    CustBalance[3] += DCLE.Amount
                                                ELSE
                                                    IF ((ActualDate - CLE."Posting Date") >= 61) AND ((ActualDate - CLE."Posting Date") <= 90) THEN
                                                        CustBalance[2] += DCLE.Amount
                                                    ELSE
                                                        IF ((ActualDate - CLE."Posting Date") >= 91) THEN
                                                            CustBalance[1] += DCLE.Amount;
                                    END;
                                UNTIL DCLE.NEXT = 0;
                        UNTIL CLE.NEXT = 0;

                    totalcustbal5 := CustBalance[1] + CustBalance[2] + CustBalance[3] + CustBalance[4] + CustBalance[5];

                    IF "Currency Code" = 'USD' THEN
                        CustBalUSD5 := totalcustbal5
                    ELSE
                        CustBalOth5 := totalcustbal5;
                END;

                IF (totalcustbal = 0) AND (totalcustbal2 = 0) AND (totalcustbal3 = 0) AND (totalcustbal4 = 0) AND (totalcustbal5 = 0) THEN
                    CurrReport.SKIP;
            end;
        }
        dataitem(Vendor; Vendor)
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Currency Code" = FILTER(<> ''));
            column(No_Vendor; Vendor."No.")
            {
            }
            column(Name_Vendor; Vendor.Name)
            {
            }
            column(CurrencyCode_Vendor; Vendor."Currency Code")
            {
            }
            column(TotalVenBal; TotalVenBal)
            {
            }
            column(VenBalUSD; VenBalUSD)
            {
            }
            column(VenBalOth; VenBalOth)
            {
            }
            column(VenBalUSD2; VenBalUSD2)
            {
            }
            column(VenBalUSD3; VenBalUSD3)
            {
            }
            column(VenBalUSD4; VenBalUSD4)
            {
            }
            column(VenBalUSD5; VenBalUSD5)
            {
            }
            column(VenBalOth2; VenBalOth2)
            {
            }
            column(VenBalOth3; VenBalOth3)
            {
            }
            column(VenBalOth4; VenBalOth4)
            {
            }
            column(VenBalOth5; VenBalOth5)
            {
            }

            trigger OnAfterGetRecord()
            begin

                CLEAR(ActualDate1);
                PayTerm.RESET;
                IF PayTerm.GET("Payment Terms Code") THEN;

                DateExp1 := ('-' + FORMAT(PayTerm."Due Date Calculation"));
                ActualDate1 := CALCDATE(PayTerm."Due Date Calculation", EDate);

                CLEAR(TotalVenBal);
                CLEAR(VenBalance);

                VLE.RESET;
                VLE.SETCURRENTKEY("Vendor No.", "Posting Date");
                VLE.SETRANGE("Vendor No.", "No.");
                VLE.SETRANGE("Posting Date", 0D, EDate);
                VLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                IF VLE.FINDFIRST THEN
                    REPEAT

                        DVLE.RESET;
                        DVLE.SETCURRENTKEY("Vendor Ledger Entry No.", "Posting Date");
                        DVLE.SETRANGE("Vendor Ledger Entry No.", VLE."Entry No.");
                        IF DVLE.FINDFIRST THEN
                            REPEAT
                                IF DVLE."Posting Date" <= EDate THEN BEGIN
                                    IF (ActualDate1 - VLE."Posting Date") <= 0 THEN
                                        VenBalance[5] += DVLE.Amount
                                    ELSE
                                        IF ((ActualDate1 - VLE."Posting Date") >= 1) AND ((ActualDate1 - VLE."Posting Date") <= 30) THEN
                                            VenBalance[4] += DVLE.Amount
                                        ELSE
                                            IF ((ActualDate1 - VLE."Posting Date") >= 31) AND ((ActualDate1 - VLE."Posting Date") <= 60) THEN
                                                VenBalance[3] += DVLE.Amount
                                            ELSE
                                                IF ((ActualDate1 - VLE."Posting Date") >= 61) AND ((ActualDate1 - VLE."Posting Date") <= 90) THEN
                                                    VenBalance[2] += DVLE.Amount
                                                ELSE
                                                    IF ((ActualDate1 - VLE."Posting Date") >= 91) THEN
                                                        VenBalance[1] += DVLE.Amount;
                                END;
                            UNTIL DVLE.NEXT = 0;
                    UNTIL VLE.NEXT = 0;

                TotalVenBal := VenBalance[1] + VenBalance[2] + VenBalance[3] + VenBalance[4] + VenBalance[5];

                CLEAR(VenBalUSD);
                CLEAR(VenBalOth);

                IF "Currency Code" = 'USD' THEN
                    VenBalUSD := TotalVenBal
                ELSE
                    VenBalOth := TotalVenBal;


                //2nd Year
                CLEAR(TotalVenBal2);
                CLEAR(VenBalUSD2);
                CLEAR(VenBalOth2);
                CLEAR(VenBalance);
                IF EYear2 <> 0D THEN BEGIN
                    ActualDate1 := CALCDATE(PayTerm."Due Date Calculation", EYear2);

                    VLE.RESET;
                    VLE.SETCURRENTKEY("Vendor No.", "Posting Date");
                    VLE.SETRANGE("Vendor No.", "No.");
                    VLE.SETRANGE("Posting Date", 0D, EYear2);
                    VLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                    IF VLE.FINDFIRST THEN
                        REPEAT

                            DVLE.RESET;
                            DVLE.SETCURRENTKEY("Vendor Ledger Entry No.", "Posting Date");
                            DVLE.SETRANGE("Vendor Ledger Entry No.", VLE."Entry No.");
                            IF DVLE.FINDFIRST THEN
                                REPEAT
                                    IF DVLE."Posting Date" <= EYear2 THEN BEGIN
                                        IF (ActualDate1 - VLE."Posting Date") <= 0 THEN
                                            VenBalance[5] += DVLE.Amount
                                        ELSE
                                            IF ((ActualDate1 - VLE."Posting Date") >= 1) AND ((ActualDate1 - VLE."Posting Date") <= 30) THEN
                                                VenBalance[4] += DVLE.Amount
                                            ELSE
                                                IF ((ActualDate1 - VLE."Posting Date") >= 31) AND ((ActualDate1 - VLE."Posting Date") <= 60) THEN
                                                    VenBalance[3] += DVLE.Amount
                                                ELSE
                                                    IF ((ActualDate1 - VLE."Posting Date") >= 61) AND ((ActualDate1 - VLE."Posting Date") <= 90) THEN
                                                        VenBalance[2] += DVLE.Amount
                                                    ELSE
                                                        IF ((ActualDate1 - VLE."Posting Date") >= 91) THEN
                                                            VenBalance[1] += DVLE.Amount;
                                    END;
                                UNTIL DVLE.NEXT = 0;
                        UNTIL VLE.NEXT = 0;

                    TotalVenBal2 := VenBalance[1] + VenBalance[2] + VenBalance[3] + VenBalance[4] + VenBalance[5];

                    IF "Currency Code" = 'USD' THEN
                        VenBalUSD2 := TotalVenBal2
                    ELSE
                        VenBalOth2 := TotalVenBal2;
                END;

                //3rd Year
                CLEAR(TotalVenBal3);
                CLEAR(VenBalUSD3);
                CLEAR(VenBalOth3);
                CLEAR(VenBalance);
                IF EYear3 <> 0D THEN BEGIN
                    ActualDate1 := CALCDATE(PayTerm."Due Date Calculation", EYear3);

                    VLE.RESET;
                    VLE.SETCURRENTKEY("Vendor No.", "Posting Date");
                    VLE.SETRANGE("Vendor No.", "No.");
                    VLE.SETRANGE("Posting Date", 0D, EYear3);
                    VLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                    IF VLE.FINDFIRST THEN
                        REPEAT

                            DVLE.RESET;
                            DVLE.SETCURRENTKEY("Vendor Ledger Entry No.", "Posting Date");
                            DVLE.SETRANGE("Vendor Ledger Entry No.", VLE."Entry No.");
                            IF DVLE.FINDFIRST THEN
                                REPEAT
                                    IF DVLE."Posting Date" <= EYear3 THEN BEGIN
                                        IF (ActualDate1 - VLE."Posting Date") <= 0 THEN
                                            VenBalance[5] += DVLE.Amount
                                        ELSE
                                            IF ((ActualDate1 - VLE."Posting Date") >= 1) AND ((ActualDate1 - VLE."Posting Date") <= 30) THEN
                                                VenBalance[4] += DVLE.Amount
                                            ELSE
                                                IF ((ActualDate1 - VLE."Posting Date") >= 31) AND ((ActualDate1 - VLE."Posting Date") <= 60) THEN
                                                    VenBalance[3] += DVLE.Amount
                                                ELSE
                                                    IF ((ActualDate1 - VLE."Posting Date") >= 61) AND ((ActualDate1 - VLE."Posting Date") <= 90) THEN
                                                        VenBalance[2] += DVLE.Amount
                                                    ELSE
                                                        IF ((ActualDate1 - VLE."Posting Date") >= 91) THEN
                                                            VenBalance[1] += DVLE.Amount;
                                    END;
                                UNTIL DVLE.NEXT = 0;
                        UNTIL VLE.NEXT = 0;

                    TotalVenBal3 := VenBalance[1] + VenBalance[2] + VenBalance[3] + VenBalance[4] + VenBalance[5];

                    IF "Currency Code" = 'USD' THEN
                        VenBalUSD3 := TotalVenBal3
                    ELSE
                        VenBalOth3 := TotalVenBal3;
                END;

                //4th Year
                CLEAR(TotalVenBal4);
                CLEAR(VenBalUSD4);
                CLEAR(VenBalOth4);
                CLEAR(VenBalance);
                IF EYear4 <> 0D THEN BEGIN
                    ActualDate1 := CALCDATE(PayTerm."Due Date Calculation", EYear4);

                    VLE.RESET;
                    VLE.SETCURRENTKEY("Vendor No.", "Posting Date");
                    VLE.SETRANGE("Vendor No.", "No.");
                    VLE.SETRANGE("Posting Date", 0D, EYear4);
                    VLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                    IF VLE.FINDFIRST THEN
                        REPEAT

                            DVLE.RESET;
                            DVLE.SETCURRENTKEY("Vendor Ledger Entry No.", "Posting Date");
                            DVLE.SETRANGE("Vendor Ledger Entry No.", VLE."Entry No.");
                            IF DVLE.FINDFIRST THEN
                                REPEAT
                                    IF DVLE."Posting Date" <= EYear4 THEN BEGIN
                                        IF (ActualDate1 - VLE."Posting Date") <= 0 THEN
                                            VenBalance[5] += DVLE.Amount
                                        ELSE
                                            IF ((ActualDate1 - VLE."Posting Date") >= 1) AND ((ActualDate1 - VLE."Posting Date") <= 30) THEN
                                                VenBalance[4] += DVLE.Amount
                                            ELSE
                                                IF ((ActualDate1 - VLE."Posting Date") >= 31) AND ((ActualDate1 - VLE."Posting Date") <= 60) THEN
                                                    VenBalance[3] += DVLE.Amount
                                                ELSE
                                                    IF ((ActualDate1 - VLE."Posting Date") >= 61) AND ((ActualDate1 - VLE."Posting Date") <= 90) THEN
                                                        VenBalance[2] += DVLE.Amount
                                                    ELSE
                                                        IF ((ActualDate1 - VLE."Posting Date") >= 91) THEN
                                                            VenBalance[1] += DVLE.Amount;
                                    END;
                                UNTIL DVLE.NEXT = 0;
                        UNTIL VLE.NEXT = 0;

                    TotalVenBal4 := VenBalance[1] + VenBalance[2] + VenBalance[3] + VenBalance[4] + VenBalance[5];

                    IF "Currency Code" = 'USD' THEN
                        VenBalUSD4 := TotalVenBal4
                    ELSE
                        VenBalOth4 := TotalVenBal4;
                END;

                //5th Year
                CLEAR(TotalVenBal5);
                CLEAR(VenBalUSD5);
                CLEAR(VenBalOth5);
                CLEAR(VenBalance);
                IF EYear5 <> 0D THEN BEGIN
                    ActualDate1 := CALCDATE(PayTerm."Due Date Calculation", EYear5);

                    VLE.RESET;
                    VLE.SETCURRENTKEY("Vendor No.", "Posting Date");
                    VLE.SETRANGE("Vendor No.", "No.");
                    VLE.SETRANGE("Posting Date", 0D, EYear5);
                    VLE.SETFILTER("Currency Code", '<>%1', '');//RB-N 05Jul2018
                    IF VLE.FINDFIRST THEN
                        REPEAT

                            DVLE.RESET;
                            DVLE.SETCURRENTKEY("Vendor Ledger Entry No.", "Posting Date");
                            DVLE.SETRANGE("Vendor Ledger Entry No.", VLE."Entry No.");
                            IF DVLE.FINDFIRST THEN
                                REPEAT
                                    IF DVLE."Posting Date" <= EYear5 THEN BEGIN
                                        IF (ActualDate1 - VLE."Posting Date") <= 0 THEN
                                            VenBalance[5] += DVLE.Amount
                                        ELSE
                                            IF ((ActualDate1 - VLE."Posting Date") >= 1) AND ((ActualDate1 - VLE."Posting Date") <= 30) THEN
                                                VenBalance[4] += DVLE.Amount
                                            ELSE
                                                IF ((ActualDate1 - VLE."Posting Date") >= 31) AND ((ActualDate1 - VLE."Posting Date") <= 60) THEN
                                                    VenBalance[3] += DVLE.Amount
                                                ELSE
                                                    IF ((ActualDate1 - VLE."Posting Date") >= 61) AND ((ActualDate1 - VLE."Posting Date") <= 90) THEN
                                                        VenBalance[2] += DVLE.Amount
                                                    ELSE
                                                        IF ((ActualDate1 - VLE."Posting Date") >= 91) THEN
                                                            VenBalance[1] += DVLE.Amount;
                                    END;
                                UNTIL DVLE.NEXT = 0;
                        UNTIL VLE.NEXT = 0;

                    TotalVenBal5 := VenBalance[1] + VenBalance[2] + VenBalance[3] + VenBalance[4] + VenBalance[5];

                    IF "Currency Code" = 'USD' THEN
                        VenBalUSD5 := TotalVenBal5
                    ELSE
                        VenBalOth5 := TotalVenBal5;
                END;

                IF (TotalVenBal = 0) AND (TotalVenBal2 = 0) AND (TotalVenBal3 = 0) AND (TotalVenBal4 = 0) AND (TotalVenBal5 = 0) THEN
                    CurrReport.SKIP;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Year"; SYear)
                {

                    trigger OnValidate()
                    begin
                        IF (SYear < 2009) OR (SYear > 2050) THEN
                            ERROR('You have entered %1 Year, which is not allowed for this report', SYear);

                        YearDiff := EYear - SYear;

                        IF YearDiff > 5 THEN
                            ERROR('This report is allowed for upto 5 Years Periods');

                        SDate := DMY2DATE(1, 4, SYear - 1);
                        EDate := DMY2DATE(31, 3, EYear);

                        IF YearDiff = 2 THEN BEGIN
                            EYear2 := DMY2DATE(31, 3, EYear - 1);
                            EYear3 := DMY2DATE(31, 3, EYear - 2);
                        END;

                        IF YearDiff = 3 THEN BEGIN
                            EYear2 := DMY2DATE(31, 3, EYear - 1);
                            EYear3 := DMY2DATE(31, 3, EYear - 2);
                            EYear4 := DMY2DATE(31, 3, EYear - 3);
                        END;

                        IF YearDiff = 4 THEN BEGIN
                            EYear2 := DMY2DATE(31, 3, EYear - 1);
                            EYear3 := DMY2DATE(31, 3, EYear - 2);
                            EYear4 := DMY2DATE(31, 3, EYear - 3);
                            EYear5 := DMY2DATE(31, 3, EYear - 4);
                        END;

                        IF YearDiff = 5 THEN BEGIN
                            EYear2 := DMY2DATE(31, 3, EYear - 1);
                            EYear3 := DMY2DATE(31, 3, EYear - 2);
                            EYear4 := DMY2DATE(31, 3, EYear - 3);
                            EYear5 := DMY2DATE(31, 3, EYear - 4);
                        END;
                    end;
                }
                field("End Year"; EYear)
                {

                    trigger OnValidate()
                    begin
                        IF (EYear < SYear) THEN
                            ERROR('End Year must be greater than Start Year');

                        IF (EYear < 2009) OR (EYear > 2050) THEN
                            ERROR('You have entered %1 Year, which is not allowed for this report', EYear);

                        YearDiff := EYear - SYear;

                        IF YearDiff > 5 THEN
                            ERROR('This report is allowed for upto 5 Years Periods');

                        EDate := DMY2DATE(31, 3, EYear);

                        IF YearDiff = 2 THEN BEGIN
                            EYear2 := DMY2DATE(31, 3, EYear - 1);
                            EYear3 := DMY2DATE(31, 3, EYear - 2);
                        END;

                        IF YearDiff = 3 THEN BEGIN
                            EYear2 := DMY2DATE(31, 3, EYear - 1);
                            EYear3 := DMY2DATE(31, 3, EYear - 2);
                            EYear4 := DMY2DATE(31, 3, EYear - 3);
                        END;

                        IF YearDiff = 4 THEN BEGIN
                            EYear2 := DMY2DATE(31, 3, EYear - 1);
                            EYear3 := DMY2DATE(31, 3, EYear - 2);
                            EYear4 := DMY2DATE(31, 3, EYear - 3);
                            EYear5 := DMY2DATE(31, 3, EYear - 4);
                        END;

                        IF YearDiff = 5 THEN BEGIN
                            EYear2 := DMY2DATE(31, 3, EYear - 1);
                            EYear3 := DMY2DATE(31, 3, EYear - 2);
                            EYear4 := DMY2DATE(31, 3, EYear - 3);
                            EYear5 := DMY2DATE(31, 3, EYear - 4);
                        END;
                    end;
                }
            }
        }

        actions
        {
        }

        trigger OnOpenPage()
        begin

            EYear := DATE2DMY(TODAY, 3);
            SYear := EYear - 1;

            SDate := DMY2DATE(1, 4, EYear - 1);
            EDate := DMY2DATE(31, 3, EYear);

            YearDiff := EYear - SYear;
        end;
    }

    labels
    {
    }

    trigger OnPreReport()
    begin

        IF SYear = 0 THEN
            ERROR('Enter The Year');
    end;

    var
        YearDiff: Integer;
        SDate: Date;
        EDate: Date;
        SYear: Integer;
        EYear: Integer;
        EYear2: Date;
        EYear3: Date;
        EYear4: Date;
        EYear5: Date;
        PayTerm: Record 3;
        DateExp: Text;
        ActualDate: Date;
        totalcustbal: Decimal;
        totalcustbal2: Decimal;
        totalcustbal3: Decimal;
        totalcustbal4: Decimal;
        totalcustbal5: Decimal;
        CustBalUSD: Decimal;
        CustBalOth: Decimal;
        CustBalUSD2: Decimal;
        CustBalOth2: Decimal;
        CustBalUSD3: Decimal;
        CustBalOth3: Decimal;
        CustBalUSD4: Decimal;
        CustBalOth4: Decimal;
        CustBalUSD5: Decimal;
        CustBalOth5: Decimal;
        CLE: Record 21;
        DCLE: Record 379;
        CustBalance: array[5] of Decimal;
        DateExp1: Text;
        ActualDate1: Date;
        TotalVenBal: Decimal;
        TotalVenBal2: Decimal;
        TotalVenBal3: Decimal;
        TotalVenBal4: Decimal;
        TotalVenBal5: Decimal;
        VLE: Record 25;
        DVLE: Record 380;
        VenBalance: array[5] of Decimal;
        VenBalUSD: Decimal;
        VenBalOth: Decimal;
        VenBalUSD2: Decimal;
        VenBalOth2: Decimal;
        VenBalUSD3: Decimal;
        VenBalOth3: Decimal;
        VenBalUSD4: Decimal;
        VenBalOth4: Decimal;
        VenBalUSD5: Decimal;
        VenBalOth5: Decimal;
}


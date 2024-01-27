report 50243 "Customer Summary Aging GP"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CustomerSummaryAgingGP.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Customer; Customer)
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(AgingBy; 'Aging By: ' + FORMAT(AgingBy))
            {
            }
            column(CountryName; CountryName)
            {
            }
            column(ResponsibilityCenter_Customer; Customer."Responsibility Center")
            {
            }
            column(AvgSales; AvgSales)
            {
            }
            column(AvgSalesPeriod; AvgSalesPeriod)
            {
            }
            column(CompanyInfo_Name; CompanyInfo.Name)
            {
            }
            column(ColumnHead_1_; ColumnHead[1])
            {
            }
            column(ColumnHead_2_; ColumnHead[2])
            {
            }
            column(ColumnHead_3_; ColumnHead[3])
            {
            }
            column(ColumnHead_4_; ColumnHead[4])
            {
            }
            column(ColumnHead_5_; ColumnHead[5])
            {
            }
            column(ColumnHead_6_; ColumnHead[6])
            {
            }
            column(ColumnHead_7_; ColumnHead[7])
            {
            }
            column(FORMAT_TODAY_0_4_; FORMAT(TODAY, 0, 4))
            {
            }
            column(USERID; USERID)
            {
            }
            column(EndingDate; FORMAT(EndingDate, 0, 4))
            {
            }
            column(Customer_TABLECAPTION_____CustFilter; Customer.GETFILTERS)
            {
            }
            column(CompanyInfo_Picture; CompanyInfo.Picture)
            {
            }
            column(CompanyRegNo; CompanyRegNo)
            {
            }
            column(Customer_Customer__No__; Customer."No.")
            {
            }
            column(Customer_Customer_Name; Customer.Name)
            {
            }
            column(Customer_Customer__Salesperson_Code_; SalesPersonDesc)
            {
            }
            column(Customer_Customer__Global_Dimension_1_Code_; Customer."Global Dimension 1 Code")
            {
            }
            column(Insurance_Limit___Management_Limit___Temporary_Limit_; InsLimit + MgmtLimit + TempLimit)
            {
            }
            column(Customer_Customer__Payment_Terms_Code_; PaymentTermsDesc)
            {
            }
            column(Customer_Balance; Customer."Balance (LCY)")
            {
            }
            column(CustBalanceDueLCY_1_; CustBalanceDueLCY[1])
            {
            }
            column(CustBalanceDueLCY_2_; CustBalanceDueLCY[2])
            {
            }
            column(CustBalanceDueLCY_3_; CustBalanceDueLCY[3])
            {
            }
            column(CustBalanceDueLCY_4_; CustBalanceDueLCY[4])
            {
            }
            column(CustBalanceDueLCY_5_; CustBalanceDueLCY[5])
            {
            }
            column(CustBalanceDueLCY_6_; CustBalanceDueLCY[6])
            {
            }
            column(CustBalanceDueLCY_7_; CustBalanceDueLCY[7])
            {
            }
            column(Customer_Balance_Control1000000023; (CustBalanceDueLCY[1] + CustBalanceDueLCY[2] + CustBalanceDueLCY[3] + CustBalanceDueLCY[4] + CustBalanceDueLCY[5] + CustBalanceDueLCY[6] + CustBalanceDueLCY[7]))
            {
            }
            column(Customer_Customer__Insurance_Limit_; InsLimit)
            {
            }
            column(Customer_Customer__Management_Limit_; MgmtLimit)
            {
            }
            column(Customer_Customer__Temporary_Limit_; TempLimit)
            {
            }
            column(PDCAmt; PDCAmt)
            {
            }
            column(Million; Million)
            {
            }
            column(LCAmt; LCAmt)
            {
            }
            column(PayRpt; PayRpt)
            {
            }
            column(UnappliedPayRpt; UnappliedPayRpt)
            {
            }
            column(OutStanding_SalesORDER_Amt; vOutStAmt)
            {
            }
            column(AverageDays; ROUND(vAvgDays, 1))
            {
                DecimalPlaces = 0 : 0;
            }
            column(OrderShipnotInvoiced; DecAmountSameCurrency)
            {
            }
            column(No_Caption; No_CaptionLbl)
            {
            }
            column(NameCaption; NameCaptionLbl)
            {
            }
            column(Sales_PersonCaption; Sales_PersonCaptionLbl)
            {
            }
            column(DepartmentCaption; DepartmentCaptionLbl)
            {
            }
            column(Credit_LimitCaption; Credit_LimitCaptionLbl)
            {
            }
            column(Payment_Terms_CodeCaption; Payment_Terms_CodeCaptionLbl)
            {
            }
            column(Outstanding_BalancyCaption; Outstanding_BalancyCaptionLbl)
            {
            }
            column(CurrReport_PAGENOCaption; CurrReport_PAGENOCaptionLbl)
            {
            }
            dataitem(Integer; Integer)
            {
                DataItemTableView = SORTING(Number);

                trigger OnAfterGetRecord()
                begin
                    IF Number = 1 THEN
                        Currency2.FIND('-')
                    ELSE
                        IF Currency2.NEXT = 0 THEN
                            CurrReport.BREAK;
                    Currency2.CALCFIELDS("Cust. Ledg. Entries in Filter");
                    IF NOT Currency2."Cust. Ledg. Entries in Filter" THEN
                        CurrReport.SKIP;

                    PrintLine := FALSE;
                    LineTotalCustBalance := 0;

                    IF SalesPerson <> '' THEN
                        Customer.SETRANGE(Customer."Salesperson Code", SalesPerson);
                    IF Customer.FIND THEN BEGIN
                        FOR j := 1 TO 7 DO BEGIN
                            CustLedgerEntry.RESET;
                            CustLedgerEntry.SETCURRENTKEY("Customer No.", "Posting Date");
                            CustLedgerEntry.SETRANGE("Customer No.", Customer."No.");
                            CustLedgerEntry.SETFILTER("Remaining Amt. (LCY)", '<>%1', 0);
                            CustLedgerEntry.SETFILTER("Date Filter", '..%1', EndingDate);
                            IF AgingBy = AgingBy::"Due Date" THEN BEGIN
                                IF j = 1 THEN
                                    CustLedgerEntry.SETRANGE("Due Date", PeriodEndDate[j + 1], PeriodEndDate[j])
                                ELSE
                                    CustLedgerEntry.SETRANGE("Due Date", (PeriodEndDate[j + 1]), PeriodEndDate[j] - 1);
                            END;
                            IF AgingBy = AgingBy::"Posting Date" THEN BEGIN
                                IF j = 1 THEN
                                    CustLedgerEntry.SETRANGE("Posting Date", (PeriodEndDate[j + 1]), PeriodEndDate[j])
                                ELSE
                                    CustLedgerEntry.SETRANGE("Posting Date", (PeriodEndDate[j + 1]), PeriodEndDate[j] - 1);
                            END;
                            CustLedgerEntry.SETRANGE("Currency Code", Currency2.Code);

                            CustLedgerEntry.CALCFIELDS(CustLedgerEntry."Remaining Amt. (LCY)");
                            CustBalanceDue[j] := CustBalanceDue[j] + CustLedgerEntry."Remaining Amt. (LCY)";

                            //Customer as Vendor
                            IF CustVend THEN BEGIN
                                VendLedgerEntry.SETCURRENTKEY("Vendor No.", "Posting Date");
                                VendLedgerEntry.SETRANGE("Vendor No.", Customer."Vendor No.");
                                CustLedgerEntry.SETFILTER("Remaining Amt. (LCY)", '<>%1', 0);
                                CustLedgerEntry.SETFILTER("Date Filter", '..%1', EndingDate);
                                IF AgingBy = AgingBy::"Due Date" THEN BEGIN
                                    IF j = 1 THEN
                                        VendLedgerEntry.SETRANGE("Due Date", PeriodEndDate[j + 1], PeriodEndDate[j])
                                    ELSE
                                        VendLedgerEntry.SETRANGE("Due Date", (PeriodEndDate[j + 1]), PeriodEndDate[j] - 1);
                                END;
                                IF AgingBy = AgingBy::"Posting Date" THEN BEGIN
                                    IF j = 1 THEN
                                        VendLedgerEntry.SETRANGE("Posting Date", (PeriodEndDate[j + 1]), PeriodEndDate[j])
                                    ELSE
                                        VendLedgerEntry.SETRANGE("Posting Date", (PeriodEndDate[j + 1]), PeriodEndDate[j] - 1);
                                END;
                                VendLedgerEntry.SETRANGE(VendLedgerEntry.Open, TRUE);
                                VendLedgerEntry.CALCFIELDS(VendLedgerEntry."Remaining Amt. (LCY)");
                                VendBalanceDue[j] := VendBalanceDue[j] + CustBalanceDue[j] + VendLedgerEntry."Remaining Amt. (LCY)";
                            END;

                            IF ShowMillion THEN BEGIN
                                CustBalanceDue[j] := (CustBalanceDue[j] + VendBalanceDue[j]) / 100000;
                                Balance := (Balance + CustBalanceDueLCY[j]) / 100000;
                            END ELSE BEGIN
                                CustBalanceDue[j] := CustBalanceDue[j] + VendBalanceDue[j];
                                Balance := Balance + CustBalanceDueLCY[j];
                            END;

                            IF CustBalanceDue[j] <> 0 THEN
                                PrintLine := TRUE;
                            LineTotalCustBalance := LineTotalCustBalance + CustBalanceDue[j];
                        END;
                    END;
                end;

                trigger OnPreDataItem()
                begin
                    IF PrintAmountsInLCY OR NOT PrintLine THEN
                        CurrReport.BREAK;
                    Currency2.RESET;
                    Currency2.SETRANGE("Customer Filter", Customer."No.");
                    Customer.COPYFILTER("Currency Filter", Currency2.Code);
                    IF (Customer.GETFILTER("Global Dimension 1 Filter") <> '') OR
                       (Customer.GETFILTER("Global Dimension 2 Filter") <> '')
                    THEN BEGIN
                        Customer.COPYFILTER("Global Dimension 1 Filter", Currency2."Global Dimension 1 Filter");
                        Customer.COPYFILTER("Global Dimension 2 Filter", Currency2."Global Dimension 2 Filter");
                    END;
                end;
            }

            trigger OnAfterGetRecord()
            begin
                GLSetup.GET;
                CLEAR(Balance);
                CLEAR(CustBalanceDue);
                CLEAR(CustBalanceDueLCY);
                CLEAR(VendBalanceDueLCY);
                CLEAR(VendBalanceDue);
                CLEAR(SalesPersonDesc);
                CLEAR(PDCAmt);
                CLEAR(PayRpt);
                CLEAR(LCAmt);
                CLEAR(UnappliedPayRpt);
                PrintLine := FALSE;
                LineTotalCustBalance := 0;
                //Get Currency Value
                CurrExchRate.RESET;
                CurrExchRate.SETCURRENTKEY("Currency Code", "Starting Date");
                //CurrExchRate.SETFILTER(CurrExchRate."Currency Code",'<>%1',GLSetup."LCY Code");
                //CurrExchRate.SETFILTER(CurrExchRate."Currency Code","Currency Code");
                CurrExchRate.SETFILTER(CurrExchRate."Currency Code", 'USD');//15Mar2019
                IF CurrExchRate.FINDLAST THEN;


                Customer.COPYFILTER("Currency Filter", CustLedgerEntry."Currency Code");
                //Payment Unapplied Entries As of period -
                CustLedgerEntry.RESET;
                CustLedgerEntry.SETCURRENTKEY("Customer No.", "Posting Date");
                CustLedgerEntry.SETRANGE("Customer No.", "No.");
                CustLedgerEntry.SETFILTER("Document Type", '%1|%2', CustLedgerEntry."Document Type"::Payment,
                                                      CustLedgerEntry."Document Type"::" ");
                CustLedgerEntry.SETFILTER("Posting Date", '..%1', CALCDATE('-1D', EndingDate));
                CustLedgerEntry.SETRANGE(Open, TRUE);
                //
                CustLedgerEntry.SETFILTER("Source Code", '<>%1', 'JOURNALV');
                //
                IF CustLedgerEntry.FINDSET THEN BEGIN
                    REPEAT
                        //CustLedgerEntry.CALCFIELDS(CustLedgerEntry."Credit Amount (LCY)");
                        CustLedgerEntry.CALCFIELDS("Remaining Amt. (LCY)");
                        IF ShowMillion THEN BEGIN
                            IF AmountUSD THEN
                                // UnappliedPayRpt := (UnappliedPayRpt+CustLedgerEntry."Remaining Amt. (LCY)"/1000000) *
                                //                                            (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount")
                                //>>05Mar2019
                                IF NOT AfricaCountry THEN
                                    UnappliedPayRpt := (UnappliedPayRpt + CustLedgerEntry."Remaining Amt. (LCY)" / 1000000) *
                                                                             (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount")
                                ELSE
                                    UnappliedPayRpt := (UnappliedPayRpt + CustLedgerEntry."Remaining Amt. (LCY)" / 1000000) / (1 / CustLedgerEntry."Original Currency Factor")
                            //<<05Mar2019
                            ELSE
                                UnappliedPayRpt := UnappliedPayRpt + CustLedgerEntry."Remaining Amt. (LCY)" / 1000000;
                        END ELSE BEGIN
                            IF AmountUSD THEN
                                // UnappliedPayRpt := (UnappliedPayRpt+CustLedgerEntry."Remaining Amt. (LCY)") *
                                //                                             (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount")
                                //>>05Mar2019
                                IF NOT AfricaCountry THEN
                                    UnappliedPayRpt := (UnappliedPayRpt + CustLedgerEntry."Remaining Amt. (LCY)") *
                                                                             (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount")
                                ELSE
                                    UnappliedPayRpt := (UnappliedPayRpt + CustLedgerEntry."Remaining Amt. (LCY)") / (1 / CustLedgerEntry."Original Currency Factor")
                            //<<05Mar2019
                            ELSE
                                UnappliedPayRpt := UnappliedPayRpt + CustLedgerEntry."Remaining Amt. (LCY)";
                        END;
                    UNTIL CustLedgerEntry.NEXT = 0;
                END;

                //Payment Unapplied Entries As of period +

                //>>28Feb2019
                AvgSales := 0;
                CLE.RESET;
                CLE.SETCURRENTKEY("Customer No.", "Posting Date", "Currency Code");
                CLE.SETRANGE("Customer No.", "No.");
                CLE.SETRANGE("Document Type", CLE."Document Type"::Invoice);
                CLE.SETRANGE("Posting Date", SDate, EndingDate);
                IF CLE.FINDSET THEN
                    REPEAT
                        CLE.CALCFIELDS("Amount (LCY)");
                        IF ShowMillion THEN BEGIN
                            IF AmountUSD THEN
                                AvgSales := (AvgSales + CLE."Amount (LCY)" / 1000000) *
                                                                            (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount")
                            ELSE
                                AvgSales := (AvgSales + CLE."Amount (LCY)" / 1000000);
                        END ELSE BEGIN
                            IF AmountUSD THEN
                                AvgSales := (AvgSales + CLE."Amount (LCY)") *
                                                                            (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount")
                            ELSE
                                AvgSales := (AvgSales + CLE."Amount (LCY)");
                        END;
                    UNTIL CLE.NEXT = 0;


                IF AvgSales <> 0 THEN BEGIN
                    IF AvgOption = AvgOption::"180 Days" THEN
                        AvgSales := AvgSales / 6;
                    IF AvgOption = AvgOption::"365 Days" THEN
                        AvgSales := AvgSales / 12;
                END;

                //<<28Feb2019
                /*
                //PDC value taking beyond the given period
                CustLedgerEntry.RESET;
                CustLedgerEntry.SETCURRENTKEY("Customer No.","Posting Date");
                CustLedgerEntry.SETRANGE("Customer No.","No.");
                CustLedgerEntry.SETRANGE(CustLedgerEntry.PDC,TRUE);
                CustLedgerEntry.SETFILTER("Posting Date",'>%1',EndingDate);
                IF CustLedgerEntry.FINDSET THEN BEGIN
                  REPEAT
                   CustLedgerEntry.CALCFIELDS(CustLedgerEntry."Credit Amount (LCY)");
                   IF ShowMillion THEN BEGIN
                    IF AmountUSD THEN
                      PDCAmt:= (PDCAmt+CustLedgerEntry."Credit Amount (LCY)"/1000000) *
                                                (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount")
                    ELSE
                      PDCAmt:= PDCAmt+CustLedgerEntry."Credit Amount (LCY)"/1000000;
                   END ELSE BEGIN
                    IF AmountUSD THEN
                      PDCAmt:= PDCAmt+CustLedgerEntry."Credit Amount (LCY)" *
                                                (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount")
                    ELSE
                      PDCAmt:= PDCAmt+CustLedgerEntry."Credit Amount (LCY)";
                   END;
                  UNTIL CustLedgerEntry.NEXT = 0;
                END;
                */
                /*
                //LC value taking beyond the given period
                FacilityUtility.RESET;
                FacilityUtility.SETRANGE("Party Code","No.");
                FacilityUtility.SETRANGE("Facility Type",'LC');
                FacilityUtility.SETRANGE("Transaction Type",FacilityUtility."Transaction Type"::Sale);
                FacilityUtility.SETRANGE(Status,FacilityUtility.Status::Released);
                FacilityUtility.SETFILTER("Maturity Date",'>%1',EndingDate);
                IF FacilityUtility.FINDSET THEN REPEAT
                   IF FacilityUtility."Source Currency Code" <> GLSetup."LCY Code" THEN BEGIN
                     CurrExchRate1.RESET;
                     CurrExchRate1.SETFILTER("Currency Code",FacilityUtility."Source Currency Code");
                     IF CurrExchRate1.FINDLAST THEN;
                
                     IF ShowMillion THEN BEGIN
                       IF AmountUSD THEN
                         LCAmt += FacilityUtility."Facility Value" / 1000000
                       ELSE
                         LCAmt += (FacilityUtility."Facility Value" / 1000000) /
                                        (CurrExchRate1."Exchange Rate Amount" / CurrExchRate1."Relational Exch. Rate Amount")
                     END ELSE BEGIN
                       IF AmountUSD THEN
                         LCAmt += FacilityUtility."Facility Value"
                       ELSE
                         LCAmt += FacilityUtility."Facility Value" /
                                  (CurrExchRate1."Exchange Rate Amount" / CurrExchRate1."Relational Exch. Rate Amount")
                     END;
                   END
                   ELSE BEGIN
                     IF ShowMillion THEN BEGIN
                       IF AmountUSD THEN
                         LCAmt += (FacilityUtility."Facility Value" / 1000000) * CurrExchRate."Relational Exch. Rate Amount"
                       ELSE
                         LCAmt += FacilityUtility."Facility Value" / 1000000
                     END ELSE BEGIN
                       IF AmountUSD THEN
                         LCAmt += FacilityUtility."Facility Value" * CurrExchRate."Relational Exch. Rate Amount"
                       ELSE
                         LCAmt += FacilityUtility."Facility Value"
                     END;
                  END;
                UNTIL FacilityUtility.NEXT = 0;
                */

                //>>28Feb2019
                CountryName := '';
                RecCountry.RESET;
                IF RecCountry.GET(Customer."Country/Region Code") THEN
                    CountryName := RecCountry.Name;
                //<<28Feb2019

                //Payment value taking beyond the given period
                CustLedgerEntry.RESET;
                CustLedgerEntry.SETCURRENTKEY("Customer No.", "Posting Date");
                CustLedgerEntry.SETRANGE("Customer No.", "No.");
                CustLedgerEntry.SETFILTER("Document Type", '%1|%2', CustLedgerEntry."Document Type"::Payment,
                                                      CustLedgerEntry."Document Type"::" ");
                //CustLedgerEntry.SETRANGE(CustLedgerEntry.PDC,FALSE);
                CustLedgerEntry.SETFILTER("Posting Date", '>%1', EndingDate);
                IF CustLedgerEntry.FINDSET THEN BEGIN
                    REPEAT
                        CustLedgerEntry.CALCFIELDS(CustLedgerEntry."Credit Amount (LCY)");
                        IF ShowMillion THEN BEGIN
                            IF AmountUSD THEN
                                PayRpt := (PayRpt + CustLedgerEntry."Credit Amount (LCY)" / 1000000) *
                                                            (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount")
                            ELSE
                                PayRpt := PayRpt + CustLedgerEntry."Credit Amount (LCY)" / 1000000
                        END ELSE BEGIN
                            IF AmountUSD THEN
                                PayRpt := (PayRpt + CustLedgerEntry."Credit Amount (LCY)") *
                                                              (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount")
                            ELSE
                                PayRpt := PayRpt + CustLedgerEntry."Credit Amount (LCY)";
                        END;
                    UNTIL CustLedgerEntry.NEXT = 0;
                END;


                FOR j := 1 TO 7 DO BEGIN
                    CustLedgerEntry.RESET;
                    CustLedgerEntry.SETCURRENTKEY("Customer No.", "Posting Date");
                    CustLedgerEntry.SETRANGE("Customer No.", "No.");
                    CustLedgerEntry.SETFILTER("Remaining Amt. (LCY)", '<>%1', 0);
                    CustLedgerEntry.SETFILTER("Date Filter", '..%1', EndingDate);
                    IF AgingBy = AgingBy::"Due Date" THEN BEGIN
                        IF j = 1 THEN
                            CustLedgerEntry.SETRANGE("Due Date", PeriodEndDate[j + 1], PeriodEndDate[j])
                        ELSE
                            CustLedgerEntry.SETRANGE("Due Date", (PeriodEndDate[j + 1]), PeriodEndDate[j] - 1);
                    END;
                    IF AgingBy = AgingBy::"Posting Date" THEN BEGIN
                        IF j = 1 THEN
                            CustLedgerEntry.SETRANGE("Posting Date", (PeriodEndDate[j + 1]), PeriodEndDate[j])
                        ELSE
                            CustLedgerEntry.SETRANGE("Posting Date", (PeriodEndDate[j + 1]), PeriodEndDate[j] - 1);
                    END;
                    IF CustLedgerEntry.FINDSET THEN BEGIN
                        REPEAT
                            CustLedgerEntry.CALCFIELDS(CustLedgerEntry."Remaining Amt. (LCY)", "Remaining Amount");
                            //CustBalanceDueLCY[j] := CustBalanceDueLCY[j]+CustLedgerEntry."Remaining Amt. (LCY)";
                            //>>05Mar2019
                            IF CustLedgerEntry."Remaining Amount" = CustLedgerEntry."Remaining Amt. (LCY)" THEN BEGIN
                                NewExRate := 0;
                                CER.RESET;
                                CER.SETCURRENTKEY("Currency Code", "Starting Date");
                                CER.SETRANGE("Starting Date", 0D, CustLedgerEntry."Posting Date");
                                CER.SETFILTER("Currency Code", 'USD');
                                IF CER.FINDLAST THEN
                                    NewExRate := CER."Exchange Rate Amount" / CER."Relational Exch. Rate Amount"
                                ELSE
                                    NewExRate := 1;
                            END ELSE
                                NewExRate := 1;
                            IF AmountUSD THEN BEGIN
                                IF NOT AfricaCountry THEN
                                    CustBalanceDueLCY[j] := CustBalanceDueLCY[j] + CustLedgerEntry."Remaining Amt. (LCY)"
                                ELSE
                                    CustBalanceDueLCY[j] := CustBalanceDueLCY[j] + CustLedgerEntry."Remaining Amount" * NewExRate;
                            END ELSE
                                CustBalanceDueLCY[j] := CustBalanceDueLCY[j] + CustLedgerEntry."Remaining Amt. (LCY)";
                        //<<05Mar2019
                        UNTIL CustLedgerEntry.NEXT = 0;
                    END;

                    //Customer as Vendor
                    IF CustVend THEN BEGIN
                        VendLedgerEntry.SETCURRENTKEY("Vendor No.", "Posting Date");
                        VendLedgerEntry.SETRANGE("Vendor No.", "Vendor No.");
                        VendLedgerEntry.SETFILTER("Remaining Amt. (LCY)", '<>%1', 0);
                        VendLedgerEntry.SETFILTER("Date Filter", '..%1', EndingDate);
                        IF AgingBy = AgingBy::"Due Date" THEN BEGIN
                            IF j = 1 THEN
                                VendLedgerEntry.SETRANGE("Due Date", PeriodEndDate[j + 1], PeriodEndDate[j])
                            ELSE
                                VendLedgerEntry.SETRANGE("Due Date", (PeriodEndDate[j + 1]), PeriodEndDate[j] - 1);
                        END;
                        IF AgingBy = AgingBy::"Posting Date" THEN BEGIN
                            IF j = 1 THEN
                                VendLedgerEntry.SETRANGE("Posting Date", (PeriodEndDate[j + 1]), PeriodEndDate[j])
                            ELSE
                                VendLedgerEntry.SETRANGE("Posting Date", (PeriodEndDate[j + 1]), PeriodEndDate[j] - 1);
                        END;

                        IF VendLedgerEntry.FINDSET THEN
                            REPEAT
                                VendLedgerEntry.CALCFIELDS(VendLedgerEntry."Remaining Amt. (LCY)");
                                VendBalanceDueLCY[j] := VendBalanceDueLCY[j] + VendLedgerEntry."Remaining Amt. (LCY)";
                            UNTIL VendLedgerEntry.NEXT = 0;
                    END;

                    IF ShowMillion THEN BEGIN
                        IF AmountUSD THEN BEGIN
                            CustBalanceDue[j] := ((CustBalanceDue[j] + VendBalanceDue[j]) / 1000000) *
                                                                        (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                            CustBalanceDueLCY[j] := ((CustBalanceDueLCY[j] + VendBalanceDueLCY[j]) / 1000000) *
                                                                        (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                            Balance := ((Balance + CustBalanceDueLCY[j]) / 1000000) *
                                                                        (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                        END ELSE BEGIN
                            CustBalanceDue[j] := (CustBalanceDue[j] + VendBalanceDue[j]) / 1000000;
                            CustBalanceDueLCY[j] := (CustBalanceDueLCY[j] + VendBalanceDueLCY[j]) / 1000000;
                            Balance := (Balance + CustBalanceDueLCY[j]) / 1000000;
                        END
                    END ELSE BEGIN
                        IF AmountUSD THEN BEGIN
                            //    CustBalanceDue[j]    := (CustBalanceDue[j]+VendBalanceDue[j]) *
                            //                                             (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount");
                            //    CustBalanceDueLCY[j] := (CustBalanceDueLCY[j]+VendBalanceDueLCY[j]) *
                            //                                              (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount");
                            //    Balance := (Balance+CustBalanceDueLCY[j]) *
                            //                                              (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount");
                            //>>05Mar2019
                            IF NOT AfricaCountry THEN BEGIN
                                CustBalanceDue[j] := (CustBalanceDue[j] + VendBalanceDue[j]) *
                                                                       (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                                CustBalanceDueLCY[j] := (CustBalanceDueLCY[j] + VendBalanceDueLCY[j]) *
                                                                        (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                                Balance := (Balance + CustBalanceDueLCY[j]) *
                                                                          (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                            END ELSE BEGIN
                                CustBalanceDue[j] := (CustBalanceDue[j] + VendBalanceDue[j]);
                                CustBalanceDueLCY[j] := (CustBalanceDueLCY[j] + VendBalanceDueLCY[j]);
                                Balance := (Balance + CustBalanceDueLCY[j]);
                            END;
                            //<<05Mar2019
                        END;
                    END;

                    IF CustBalanceDue[j] <> 0 THEN
                        PrintLine := TRUE;
                    IF ShowMillion THEN BEGIN
                        IF AmountUSD THEN BEGIN
                            LineTotalCustBalance := ((LineTotalCustBalance + CustBalanceDueLCY[j]) / 1000000) *
                                                                        (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                            TotalCustBalanceLCY := ((TotalCustBalanceLCY + CustBalanceDueLCY[j]) / 1000000) *
                                                                        (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                        END ELSE BEGIN
                            LineTotalCustBalance := (LineTotalCustBalance + CustBalanceDueLCY[j]) / 1000000;
                            TotalCustBalanceLCY := (TotalCustBalanceLCY + CustBalanceDueLCY[j]) / 1000000;
                        END;
                    END ELSE BEGIN
                        IF AmountUSD THEN BEGIN
                            //LineTotalCustBalance:= (LineTotalCustBalance + CustBalanceDueLCY[j]) *
                            //                                            (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount");
                            //TotalCustBalanceLCY := (TotalCustBalanceLCY + CustBalanceDueLCY[j]) *
                            //                                            (CurrExchRate."Exchange Rate Amount"/CurrExchRate."Relational Exch. Rate Amount");
                            //>>05Mar2019
                            IF NOT AfricaCountry THEN BEGIN
                                LineTotalCustBalance := (LineTotalCustBalance + CustBalanceDueLCY[j]) *
                                                                          (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                                TotalCustBalanceLCY := (TotalCustBalanceLCY + CustBalanceDueLCY[j]) *
                                                                          (CurrExchRate."Exchange Rate Amount" / CurrExchRate."Relational Exch. Rate Amount");
                            END ELSE BEGIN
                                LineTotalCustBalance := (LineTotalCustBalance + CustBalanceDueLCY[j]);
                                TotalCustBalanceLCY := (TotalCustBalanceLCY + CustBalanceDueLCY[j]);
                            END;
                            //<<05Mar2019
                        END ELSE BEGIN
                            LineTotalCustBalance := LineTotalCustBalance + CustBalanceDueLCY[j];
                            TotalCustBalanceLCY := TotalCustBalanceLCY + CustBalanceDueLCY[j];
                        END;
                    END;
                END;

                GLSetup.GET;
                InsLimit := 0;
                MgmtLimit := 0;
                TempLimit := 0;
                IF ("Currency Code" = '') OR (GLSetup."LCY Code" = "Currency Code") THEN BEGIN
                    Million := 'All Figures are in million and in' + ' ' + GLSetup."LCY Code";
                    IF AfricaCountry THEN
                        IF AmountUSD THEN
                            Million := 'All Figures are in million and in USD';
                    IF ShowMillion AND ("Credit Limit Approval Status" = "Credit Limit Approval Status"::Approved) THEN BEGIN
                        IF "Insurance Limit Exp Date" >= WORKDATE THEN
                            InsLimit := (Customer."Insurance Limit" * "Insurance Percentage" / 100) / 1000000;
                        IF "Management Limit Exp Date" >= WORKDATE THEN
                            MgmtLimit := (Customer."Management Limit" * "Management Percentage" / 100) / 1000000;
                        IF "Temporary Limit Exp Date" >= WORKDATE THEN
                            TempLimit := (Customer."Temporary Limit" * "Temporary Percentage" / 100) / 1000000;
                    END
                    ELSE
                        IF "Credit Limit Approval Status" = "Credit Limit Approval Status"::Approved THEN BEGIN
                            IF "Insurance Limit Exp Date" >= WORKDATE THEN
                                InsLimit := Customer."Insurance Limit" * "Insurance Percentage" / 100;
                            IF "Management Limit Exp Date" >= WORKDATE THEN
                                MgmtLimit := Customer."Management Limit" * "Management Percentage" / 100;
                            IF "Temporary Limit Exp Date" >= WORKDATE THEN
                                TempLimit := (Customer."Temporary Limit" * "Temporary Percentage" / 100);
                        END;
                END
                ELSE BEGIN
                    IF ShowMillion AND ("Credit Limit Approval Status" = "Credit Limit Approval Status"::Approved) THEN BEGIN
                        IF NOT AmountUSD THEN BEGIN
                            Million := 'All Figures are in million and in' + ' ' + GLSetup."LCY Code";
                            IF "Insurance Limit Exp Date" >= WORKDATE THEN
                                InsLimit := ((Customer."Insurance Limit" * "Insurance Percentage" / 100) / 1000000) * CurrExchRate."Relational Exch. Rate Amount";
                            IF "Management Limit Exp Date" >= WORKDATE THEN
                                MgmtLimit := ((Customer."Management Limit" * "Management Percentage" / 100) / 1000000) *
                                   CurrExchRate."Relational Exch. Rate Amount";
                            IF "Temporary Limit Exp Date" >= WORKDATE THEN
                                TempLimit := ((Customer."Temporary Limit" * "Temporary Percentage" / 100) / 1000000) * CurrExchRate."Relational Exch. Rate Amount"
                        END ELSE BEGIN
                            Million := 'All Figures are in million and in USD';
                            IF "Insurance Limit Exp Date" >= WORKDATE THEN
                                InsLimit := (Customer."Insurance Limit" * "Insurance Percentage" / 100) / 1000000;
                            IF "Management Limit Exp Date" >= WORKDATE THEN
                                MgmtLimit := (Customer."Management Limit" * "Management Percentage" / 100) / 1000000;
                            IF "Temporary Limit Exp Date" >= WORKDATE THEN
                                TempLimit := (Customer."Temporary Limit" * "Temporary Percentage" / 100) / 1000000;
                        END;
                    END ELSE
                        IF ("Credit Limit Approval Status" = "Credit Limit Approval Status"::Approved) THEN BEGIN
                            IF NOT AmountUSD THEN BEGIN
                                Million := 'All Figures are in' + ' ' + GLSetup."LCY Code";
                                IF "Insurance Limit Exp Date" >= WORKDATE THEN
                                    InsLimit := (Customer."Insurance Limit" * "Insurance Percentage" / 100) * CurrExchRate."Relational Exch. Rate Amount";
                                IF "Management Limit Exp Date" >= WORKDATE THEN
                                    MgmtLimit := (Customer."Management Limit" * "Management Percentage" / 100) * CurrExchRate."Relational Exch. Rate Amount";
                                IF "Temporary Limit Exp Date" >= WORKDATE THEN
                                    TempLimit := (Customer."Temporary Limit" * "Temporary Percentage" / 100) * CurrExchRate."Relational Exch. Rate Amount";
                            END ELSE BEGIN
                                Million := 'All Figures are in USD';
                                IF "Insurance Limit Exp Date" >= WORKDATE THEN
                                    InsLimit := Customer."Insurance Limit" * "Insurance Percentage" / 100;
                                IF "Management Limit Exp Date" >= WORKDATE THEN
                                    MgmtLimit := Customer."Management Limit" * "Management Percentage" / 100;
                                IF "Temporary Limit Exp Date" >= WORKDATE THEN
                                    TempLimit := Customer."Temporary Limit" * "Temporary Percentage" / 100;
                            END;
                        END;
                END;


                IF "Payment Terms Code" <> '' THEN
                    IF PaymentTerms.GET("Payment Terms Code") THEN BEGIN
                        PaymentTerms.TranslateDescription(PaymentTerms, "Language Code");
                        PaymentTermsDesc := PaymentTerms.Description;
                    END;

                IF "Salesperson Code" <> '' THEN
                    IF SalesPersonTable.GET("Salesperson Code") THEN BEGIN
                        SalesPersonDesc := SalesPersonTable.Name;
                    END;

                //RSPL008 >>>>
                vOutStAmt := 0;
                SalesLine.RESET;
                SalesLine.CHANGECOMPANY(COMPANYNAME);
                SalesLine.SETCURRENTKEY("Document Type", "Bill-to Customer No.", Status, "Currency Code", "Credit Checking Not Required", Type);
                SalesLine.SETFILTER("Document Type", '%1', SalesLine."Document Type"::Order);
                SalesLine.SETFILTER("Bill-to Customer No.", "No.");
                SalesLine.SETFILTER(Status, '%1', SalesLine.Status::Released);
                SalesLine.SETFILTER("Credit Checking Not Required", '%1', FALSE);
                SalesLine.SETFILTER(Type, '%1', SalesLine.Type::Item);
                IF SalesLine.FINDSET THEN BEGIN
                    SalesLine.CALCSUMS(SalesLine."Outstanding Amount", SalesLine."Outstanding Amount (LCY)");
                    IF SalesLine."Outstanding Amount (LCY)" <> 0 THEN
                        vOutStAmt := vOutStAmt + SalesLine."Outstanding Amount (LCY)"
                    ELSE
                        vOutStAmt := vOutStAmt + SalesLine."Outstanding Amount";
                END;
                /*
                recSH.RESET;
                recSH.SETCURRENTKEY("Document Type",Status,"Credit Limit Approval Status");
                recSH.SETRANGE("Document Type",recSH."Document Type"::Order);
                recSH.SETRANGE(Status,recSH.Status::Released);
                recSH.SETRANGE("Credit Limit Approval Status",recSH."Credit Limit Approval Status"::Released);
                recSH.SETRANGE("Sell-to Customer No.","No.");
                IF recSH.FINDFIRST THEN REPEAT
                  recSH.CALCFIELDS("Amount to Customer");
                 IF recSH."Currency Factor" <> 0 THEN
                    vOutStAmt := vOutStAmt+ recSH."Amount to Customer"/recSH."Currency Factor"
                   ELSE
                    vOutStAmt := vOutStAmt+ recSH."Amount to Customer";
                 UNTIL recSH.NEXT=0;
                */
                vDaysDiff := 0;
                vPntEntrCnt := 0;
                vAvgDays := 0;
                recCustLdEntryInv.RESET;
                recCustLdEntryInv.SETCURRENTKEY("Customer No.", "Posting Date", "Currency Code");
                recCustLdEntryInv.SETRANGE("Customer No.", "No.");
                recCustLdEntryInv.SETRANGE("Document Type", recCustLdEntryInv."Document Type"::Invoice);
                recCustLdEntryInv.SETRANGE("Source Code", 'SALES');
                recCustLdEntryInv.SETRANGE(Open, FALSE);
                IF recCustLdEntryInv.FINDFIRST THEN
                    REPEAT
                        IF recCustLdEntryInv."Closed by Entry No." <> 0 THEN BEGIN
                            recCustLdEntryPay.RESET;
                            recCustLdEntryPay.SETCURRENTKEY("Entry No.");
                            recCustLdEntryPay.SETRANGE("Entry No.", recCustLdEntryInv."Closed by Entry No.");
                            IF recCustLdEntryPay.FINDFIRST THEN BEGIN
                                vPntEntrCnt += 1;
                                vDaysDiff := vDaysDiff + (recCustLdEntryPay."Posting Date" - recCustLdEntryInv."Posting Date");
                            END;
                        END;
                    UNTIL recCustLdEntryInv.NEXT = 0;
                IF vPntEntrCnt <> 0 THEN
                    vAvgDays := vDaysDiff / vPntEntrCnt;
                //RSPL008 <<<<

                //GULF/001 >>>> Order Ship not Invoiced Value Capture
                DecAmountSameCurrency := 0;
                SalesLine.RESET;
                SalesLine.CHANGECOMPANY(COMPANYNAME);
                SalesLine.SETCURRENTKEY("Document Type", "Bill-to Customer No.", Status, "Currency Code", "Credit Checking Not Required", Type);
                SalesLine.SETFILTER("Document Type", '%1', SalesLine."Document Type"::Order);
                SalesLine.SETFILTER("Bill-to Customer No.", "No.");
                SalesLine.SETFILTER(Status, '%1', SalesLine.Status::Released);
                SalesLine.SETFILTER("Credit Checking Not Required", '%1', FALSE);
                SalesLine.SETFILTER(Type, '%1', SalesLine.Type::Item);
                IF SalesLine.FINDSET THEN BEGIN
                    SalesLine.CALCSUMS(SalesLine."Shipped Not Invoiced", SalesLine."Shipped Not Invoiced (LCY)", SalesLine."Outstanding Amount (LCY)");
                    IF SalesLine."Shipped Not Invoiced (LCY)" <> 0 THEN
                        DecAmountSameCurrency := DecAmountSameCurrency + SalesLine."Shipped Not Invoiced (LCY)"
                    ELSE
                        DecAmountSameCurrency := DecAmountSameCurrency + SalesLine."Shipped Not Invoiced"
                END;
                //GULF/001 <<<< Order Ship not Invoiced Value Capture

            end;

            trigger OnPreDataItem()
            begin
                CompanyInfo.GET;
                //>>05Mar2019
                AfricaCountry := FALSE;
                IF (CompanyInfo."Country/Region Code" = 'KE') OR (CompanyInfo."Country/Region Code" = 'UGANDA')
                OR (CompanyInfo."Country/Region Code" = 'TZ') OR (CompanyInfo."Country/Region Code" = 'NG') THEN
                    AfricaCountry := TRUE;
                //<<05Mar2019
                CompanyInfo.CALCFIELDS(CompanyInfo.Picture);
                ColumnHead[1] := FORMAT(PeriodEndDate[1] - PeriodEndDate[1])
                                 + ' - '
                                 + FORMAT(PeriodEndDate[1] - PeriodEndDate[2])
                                 + Text012;

                ColumnHead[2] := FORMAT(PeriodEndDate[1] - PeriodEndDate[2] + 1)
                                 + ' - '
                                 + FORMAT(PeriodEndDate[1] - PeriodEndDate[3])
                                 + Text012;

                ColumnHead[3] := FORMAT(PeriodEndDate[1] - PeriodEndDate[3] + 1)
                                 + ' - '
                                 + FORMAT(PeriodEndDate[1] - PeriodEndDate[4])
                                 + Text012;
                ColumnHead[4] := FORMAT(PeriodEndDate[1] - PeriodEndDate[4] + 1)
                                 + ' - '
                                 + FORMAT(PeriodEndDate[1] - PeriodEndDate[5])
                                 + Text012;

                ColumnHead[5] := FORMAT(PeriodEndDate[1] - PeriodEndDate[5] + 1)
                                 + ' - '
                                 + FORMAT(PeriodEndDate[1] - PeriodEndDate[6])
                                 + Text012;

                ColumnHead[6] := FORMAT(PeriodEndDate[1] - PeriodEndDate[6] + 1)
                                 + ' - '
                                 + FORMAT(PeriodEndDate[1] - PeriodEndDate[7])
                                 + Text012;

                ColumnHead[7] := Text020
                                 + ' '
                                 + FORMAT(PeriodEndDate[1] - PeriodEndDate[7])
                                 + Text012;

                // CurrReport.CREATETOTALS(CustBalanceDue, CustBalanceDueLCY, TotalCustBalanceLCY);
                Currency2.Code := '';
                Currency2.INSERT;
                IF Currency.FIND('-') THEN
                    REPEAT
                        Currency2 := Currency;
                        Currency2.INSERT;
                    UNTIL Currency.NEXT = 0;
                //INA-GP-001++
                Bool := TRUE;
                CompanyRegNo := CompanyInfo."Registration No.";

                IF COMPANYNAME = 'Gulf Petrochem SG' THEN
                    Bool := FALSE;
                //INA-GP-001--
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(EndingDate; EndingDate)
                {
                    ApplicationArea = all;
                    Caption = 'Aged As Of';
                }
                field(AgingBy; AgingBy)
                {
                    ApplicationArea = all;
                    Caption = 'Aging by';
                }
                field(ShowMillion; ShowMillion)
                {
                    ApplicationArea = all;
                    Caption = 'Show amount in millions';
                }
                field(AmountUSD; AmountUSD)
                {
                    ApplicationArea = all;
                    Caption = 'Show amount in USD';
                }
                field("Average Sales"; AvgOption)
                {
                    ApplicationArea = all;
                    Caption = 'Average Sales';
                }
                group("Period Length")
                {

                    Caption = 'Period Length';
                    field(Slab1PeriodCalculation; Slab1PeriodCalculation)
                    {
                        ApplicationArea = all;
                        Style = Favorable;
                        StyleExpr = TRUE;
                    }
                    field(Slab2PeriodCalculation; Slab2PeriodCalculation)
                    {
                        ApplicationArea = all;
                        Style = Favorable;
                        StyleExpr = TRUE;
                    }
                    field(Slab3PeriodCalculation; Slab3PeriodCalculation)
                    {
                        ApplicationArea = all;
                        Style = Favorable;
                        StyleExpr = TRUE;
                    }
                    field(Slab4PeriodCalculation; Slab4PeriodCalculation)
                    {
                        ApplicationArea = all;
                        Style = Favorable;
                        StyleExpr = TRUE;
                    }
                    field(Slab5PeriodCalculation; Slab5PeriodCalculation)
                    {
                        ApplicationArea = all;
                        Style = Favorable;
                        StyleExpr = TRUE;
                    }
                    field(Slab6PeriodCalculation; Slab6PeriodCalculation)
                    {
                        ApplicationArea = all;
                        Style = Favorable;
                        StyleExpr = TRUE;
                    }
                    field("Customer As Vendor"; CustVend)
                    {
                        ApplicationArea = all;
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

            IF PeriodEndDate[1] = 0D THEN BEGIN
                PeriodEndDate[1] := WORKDATE;
                Slab1PeriodCalculation := Text002;
                Slab2PeriodCalculation := Text003;
                Slab3PeriodCalculation := Text004;
                Slab4PeriodCalculation := Text005;
                Slab5PeriodCalculation := Text006;
                Slab6PeriodCalculation := Text007;
            END;
        end;
    }

    labels
    {
    }

    trigger OnPreReport()
    begin
        CustFilter := Customer.GETFILTERS;
        PeriodEndDate[1] := EndingDate;
        IF AgingBy = AgingBy::"Due Date" THEN BEGIN
            PeriodEndDate[2] := PeriodEndDate[1];
            FOR j := 2 TO 7 DO
                IF j = 2 THEN BEGIN
                    PeriodCalculation := Slab1PeriodCalculation;
                    PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                END ELSE
                    IF j = 3 THEN BEGIN
                        PeriodCalculation := Slab2PeriodCalculation;
                        PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                    END ELSE
                        IF j = 4 THEN BEGIN
                            PeriodCalculation := Slab3PeriodCalculation;
                            PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                        END ELSE
                            IF j = 5 THEN BEGIN
                                PeriodCalculation := Slab4PeriodCalculation;
                                PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                            END ELSE
                                IF j = 6 THEN BEGIN
                                    PeriodCalculation := Slab5PeriodCalculation;
                                    PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                                END ELSE
                                    IF j = 7 THEN BEGIN
                                        PeriodCalculation := Slab6PeriodCalculation;
                                        PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                                    END;
        END ELSE BEGIN
            FOR j := 2 TO 7 DO
                IF j = 2 THEN BEGIN
                    PeriodCalculation := Slab1PeriodCalculation;
                    PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                END ELSE
                    IF j = 3 THEN BEGIN
                        PeriodCalculation := Slab2PeriodCalculation;
                        PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                    END ELSE
                        IF j = 4 THEN BEGIN
                            PeriodCalculation := Slab3PeriodCalculation;
                            PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                        END ELSE
                            IF j = 5 THEN BEGIN
                                PeriodCalculation := Slab4PeriodCalculation;
                                PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                            END ELSE
                                IF j = 6 THEN BEGIN
                                    PeriodCalculation := Slab5PeriodCalculation;
                                    PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                                END ELSE
                                    IF j = 7 THEN BEGIN
                                        PeriodCalculation := Slab6PeriodCalculation;
                                        PeriodEndDate[j] := CALCDATE('-(' + PeriodCalculation + ')', PeriodEndDate[1]);
                                    END;
        END;
        PeriodEndDate[8] := 0D;

        //>>28Feb2019
        CLEAR(SDate);
        CLEAR(AvgSalesPeriod);
        IF AvgOption = AvgOption::"180 Days" THEN BEGIN
            SDate := CALCDATE('-180D', EndingDate);
            AvgSalesPeriod := '6 Months';
        END;
        IF AvgOption = AvgOption::"365 Days" THEN BEGIN
            SDate := CALCDATE('-365D', EndingDate);
            AvgSalesPeriod := '12 Months';
        END;
        //<<28Feb2019
    end;

    var
        Currency: Record 4;
        Currency2: Record 4 temporary;
        CustFilter: Text[250];
        PrintAmountsInLCY: Boolean;
        PeriodLength: DateFormula;
        PeriodStartDate: array[7] of Date;
        CustBalanceDue: array[7] of Decimal;
        CustBalanceDueLCY: array[7] of Decimal;
        LineTotalCustBalance: Decimal;
        TotalCustBalanceLCY: Decimal;
        PrintLine: Boolean;
        j: Integer;
        PaymentTermsDesc: Text[60];
        SalesPersonDesc: Text[50];
        DateFilter: Option BillDate,DueDate;
        SalesPerson: Code[10];
        SalesPersonTable: Record 13;
        test: Text[50];
        GroupBySalesPerson: Boolean;
        PrintAmountInLCY: Boolean;
        PaymentTerms: Record 3;
        EndingDate: Date;
        AgingBy: Option "Posting Date","Due Date";
        PeriodEndDate: array[8] of Date;
        HeaderText: array[7] of Text[30];
        Balance: Decimal;
        DateDiff: Integer;
        CompanyInfo: Record 79;
        Slab1PeriodCalculation: Code[10];
        Slab2PeriodCalculation: Code[10];
        Slab3PeriodCalculation: Code[10];
        Slab4PeriodCalculation: Code[10];
        Slab5PeriodCalculation: Code[10];
        PeriodCalculation: Code[10];
        CompanyInformation: Record 79;
        TempCustLedgEntry: Record 21 temporary;
        ColumnHead: array[7] of Text[20];
        ColumnHeadHead: Text[59];
        CustVend: Boolean;
        VendBalanceDue: array[8] of Decimal;
        VendBalanceDueLCY: array[8] of Decimal;
        CustLedgerEntry: Record 21;
        VendLedgerEntry: Record 25;
        PDCAmt: Decimal;
        ShowMillion: Boolean;
        InsLimit: Decimal;
        MgmtLimit: Decimal;
        TempLimit: Decimal;
        Million: Text[50];
        SaleInvhdr: Record 112;
        FacilityUtility: Record 50003;
        LCAmt: Decimal;
        PayRpt: Decimal;
        UnappliedPayRpt: Decimal;
        AmountUSD: Boolean;
        GLSetup: Record 98;
        CurrExchRate: Record 330;
        CompanyRegNo: Text[30];
        Bool: Boolean;
        CurrExchRate1: Record 330;
        "-------RSPL----": Integer;
        recSH: Record 36;
        vOutStAmt: Decimal;
        recCustLdEntryInv: Record 21;
        recCustLdEntryPay: Record 21;
        vDaysDiff: Integer;
        vPntEntrCnt: Integer;
        vAvgDays: Decimal;
        vDays: Integer;
        SalesLine: Record 37;
        DecAmountSameCurrency: Decimal;
        "--------28Feb2019": Integer;
        CountryName: Text;
        RecCountry: Record 9;
        AvgOption: Option "180 Days","365 Days";
        AvgSales: Decimal;
        AvgSalesPeriod: Text;
        CLE: Record 21;
        SDate: Date;
        AfricaCountry: Boolean;
        CER: Record 330;
        NewExRate: Decimal;
        Text001: Label 'Days';
        Text002: Label '30D';
        Text003: Label '60D';
        Text004: Label '90D';
        Text005: Label '120D';
        Text006: Label '180D';
        Text007: Label '365D';
        Text011: Label 'Up To';
        Text012: Label 'Days';
        Text013: Label 'Aged Overdue Amounts';
        Text016: Label 'Aged Customer Balances';
        Text020: Label 'Over';
        No_CaptionLbl: Label 'No.';
        NameCaptionLbl: Label 'Name';
        Sales_PersonCaptionLbl: Label 'Sales Person';
        DepartmentCaptionLbl: Label 'Department';
        Credit_LimitCaptionLbl: Label 'Credit Limit';
        Payment_Terms_CodeCaptionLbl: Label 'Payment Terms Code';
        Outstanding_BalancyCaptionLbl: Label 'Outstanding Balancy';
        CurrReport_PAGENOCaptionLbl: Label 'Label1000000022';
        Slab6PeriodCalculation: Code[10];
}


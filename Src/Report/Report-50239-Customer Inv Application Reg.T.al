report 50239 "Customer Inv Application Reg.T"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 19Nov2018   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CustomerInvApplicationRegT.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Customer Invoice Application Register';

    dataset
    {
        dataitem(CLE; "Cust. Ledger Entry")
        {
            DataItemTableView = SORTING("Document No.", "Document Type", "Customer No.")
                                WHERE("Document Type" = CONST(Invoice));
            RequestFilterFields = "Entry No.";
            column(EntryNo_CLE; CLE."Entry No.")
            {
            }
            column(CustomerNo_CLE; CLE."Customer No.")
            {
            }
            column(PostingDate_CLE; CLE."Posting Date")
            {
            }
            column(DocumentType_CLE; CLE."Document Type")
            {
            }
            column(DocumentNo_CLE; CLE."Document No.")
            {
            }
            column(CustName; CustName)
            {
            }
            column(Dim1; CLE."Global Dimension 1 Code")
            {
            }
            column(Dim2; CLE."Global Dimension 2 Code")
            {
            }
            column(DimText; DimText)
            {
            }
            column(AmountLCY_CLE; CLE."Amount (LCY)")
            {
            }
            column(SalesLCY_CLE; CLE."Sales (LCY)")
            {
            }
            column(ProfitLCY_CLE; CLE."Profit (LCY)")
            {
            }
            column(OriginalAmtLCY_CLE; CLE."Original Amt. (LCY)")
            {
            }
            column(Cdper; CdPer)
            {
            }
            column(CdAmt; CdAmt)
            {
            }
            column(TxtFilter; TxtFilter)
            {
            }
            column(PayDocNo; PayDocNo)
            {
            }
            column(PayDocDate; PayDocDate)
            {
            }
            column(BankName; BankName)
            {
            }
            column(ChqNo; ChqNo)
            {
            }
            column(ChqDate; ChqDate)
            {
            }
            column(ClosedbyAmountLCY_CLE; CLE."Closed by Amount (LCY)")
            {
            }
            column(TCSAmount; TCSAmount)
            {
            }
            dataitem(APPCLE2; "Cust. Ledger Entry")
            {
                DataItemLink = "Closed by Entry No." = FIELD("Entry No.");
                DataItemTableView = SORTING("Entry No.");
                column(EntryNo_APPCLE; "Entry No.")
                {
                }
                column(CustomerNo_APPCLE; "Customer No.")
                {
                }
                column(PostingDate_APPCLE; FORMAT("Posting Date"))
                {
                }
                column(DocumentType_APPCLE; "Document Type")
                {
                }
                column(DocumentNo_APPCLE; "Document No.")
                {
                }
                column(Description_APPCLE; Description)
                {
                }
                column(CurrencyCode_APPCLE; "Currency Code")
                {
                }
                column(Amount_APPCLE; Amount)
                {
                }
                column(RemainingAmount_APPCLE; "Remaining Amount")
                {
                }
                column(PayDocNo2; PayDocNo2)
                {
                }
                column(PayDocDate2; PayDocDate2)
                {
                }
                column(BankName2; BankName2)
                {
                }
                column(ChqNo2; ChqNo2)
                {
                }
                column(ChqDate2; FORMAT(ChqDate2))
                {
                }
                column(ClosebyAmountLCY_APPCLE2; ABS("Closed by Amount (LCY)"))
                {
                }

                trigger OnAfterGetRecord()
                begin


                    PayDocNo2 := '';
                    CLEAR(PayDocDate2);
                    CLEAR(BankName2);
                    CLEAR(ChqNo2);
                    CLEAR(ChqDate2);

                    PayDocNo2 := "Document No.";
                    PayDocDate2 := "Posting Date";

                    BLE2.RESET;
                    BLE2.SETCURRENTKEY("Document No.");
                    BLE2.SETRANGE("Document No.", PayDocNo2);
                    IF BLE2.FINDFIRST THEN BEGIN
                        ChqNo2 := BLE2."Cheque No.";
                        ChqDate2 := BLE2."Cheque Date";

                        BankAcc2.RESET;
                        IF BankAcc2.GET(BLE2."Bank Account No.") THEN
                            BankName2 := BankAcc2.Name;
                    END;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                CustName := '';
                Cust.RESET;
                IF Cust.GET("Customer No.") THEN
                    CustName := Cust.Name;

                DimText := '';
                Dim.RESET;
                IF Dim.GET('DIVISION', CLE."Global Dimension 1 Code") THEN
                    DimText := Dim.Name;


                CdPer := 0;
                CdAmt := 0;
                /*  PstrLine.RESET;
                 PstrLine.SETCURRENTKEY("Invoice No.");
                 PstrLine.SETRANGE("Invoice No.", "Document No.");
                 PstrLine.SETRANGE("Tax/Charge Group", 'CASHDISC');
                 IF PstrLine.FINDFIRST THEN BEGIN
                     CdPer := ABS(PstrLine."Calculation Value");
                     PstrLine.CALCSUMS("Amount (LCY)");
                     CdAmt := PstrLine."Amount (LCY)";
                 END; */

                PayDocNo := '';
                CLEAR(PayDocDate);
                CLEAR(BankName);
                CLEAR(ChqNo);
                CLEAR(ChqDate);
                IF CLE."Closed by Entry No." <> 0 THEN BEGIN
                    AppCLE.RESET;
                    IF AppCLE.GET(CLE."Closed by Entry No.") THEN BEGIN
                        PayDocNo := AppCLE."Document No.";
                        //MESSAGE(AppCLE."Document No.");
                        PayDocDate := AppCLE."Posting Date";

                        BLE.RESET;
                        BLE.SETCURRENTKEY("Document No.");
                        BLE.SETRANGE("Document No.", PayDocNo);
                        IF BLE.FINDFIRST THEN BEGIN
                            ChqNo := BLE."Cheque No.";
                            ChqDate := BLE."Cheque Date";

                            BankAcc.RESET;
                            IF BankAcc.GET(BLE."Bank Account No.") THEN
                                BankName := BankAcc.Name;
                        END;
                    END;
                END;

                //RSPLSUM-TCS 06Nov2020>>
                CLEAR(TCSAmount);
                RecSIL.RESET;
                RecSIL.SETCURRENTKEY("Document No.", Type, Quantity);
                RecSIL.SETRANGE("Document No.", CLE."Document No.");
                RecSIL.SETFILTER(Quantity, '<>%1', 0);
                IF RecSIL.FINDSET THEN
                    REPEAT
                        TCSAmount += 0;// RecSIL."TDS/TCS Amount";
                    UNTIL RecSIL.NEXT = 0;
                //RSPLSUM-TCS 06Nov2020<<
            end;

            trigger OnPreDataItem()
            begin

                SETRANGE("Posting Date", SDate, EDate);

                IF LocFilter <> '' THEN
                    SETFILTER("Location Code", LocFilter);

                IF CusFilter <> '' THEN
                    SETFILTER("Customer No.", CusFilter);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; SDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin

                        EDate := SDate;
                    end;
                }
                field("End Date"; EDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin

                        IF EDate < SDate THEN
                            ERROR('EndDate must be greater than StartDate');
                    end;
                }
                field("Location Filter"; LocFilter)
                {
                    ApplicationArea = all;
                    TableRelation = Location;
                }
                field("Customer Filter"; CusFilter)
                {
                    ApplicationArea = all;
                    TableRelation = Customer;
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

        IF SDate = 0D THEN
            ERROR('Please Specify Start Date');

        IF EDate = 0D THEN
            ERROR('Please Specify End Date');

        TxtFilter := '';
        TxtFilter := 'Date Filter : ' + FORMAT(SDate) + '..' + FORMAT(EDate);

        IF LocFilter <> '' THEN
            TxtFilter += ', Location : ' + LocFilter;

        IF CusFilter <> '' THEN
            TxtFilter += ', Customer : ' + CusFilter;
    end;

    var
        SDate: Date;
        EDate: Date;
        LocFilter: Code[100];
        CusFilter: Code[100];
        Cust: Record 18;
        CustName: Text;
        Dim: Record 349;
        DimText: Text;
        // PstrLine: Record 13798;
        CdPer: Decimal;
        CdAmt: Decimal;
        TxtFilter: Text;
        AppCLE: Record 21;
        BLE: Record 271;
        PayDocNo: Code[20];
        PayDocDate: Date;
        BankName: Text;
        ChqNo: Code[20];
        ChqDate: Date;
        BankAcc: Record 270;
        "-----------------------RSPL-----------------": Integer;
        AppCLEPay: Record 21;
        BLE2: Record 271;
        PayDocNo2: Code[20];
        PayDocDate2: Date;
        BankName2: Text;
        ChqNo2: Code[20];
        ChqDate2: Date;
        BankAcc2: Record 270;
        RecSIL: Record 113;
        TCSAmount: Decimal;
}


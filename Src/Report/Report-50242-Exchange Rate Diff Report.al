report 50242 "Exchange Rate Diff Report"
{
    // //RSPLSUM 05Jan21-- added Unrealized Gain/Loss in filter of Detailed Cust Ledger Entry and Detailed Vendor Ledger Entry
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ExchangeRateDiffReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Detailed Cust. Ledg. Entry"; "Detailed Cust. Ledg. Entry")
        {
            DataItemTableView = ORDER(Ascending)
                                WHERE("Currency Code" = FILTER(<> ''),
                                      "Entry Type" = FILTER(Application | "Unrealized Gain" | "Unrealized Loss"));
            RequestFilterFields = "Posting Date", "Document No.";
            column(EntryUnrealizedGainLoss; EntryUnrealizedGainLoss)
            {
            }
            column(CustomerDetail; CustomerDetail)
            {
            }
            column(CurranceFactorDetail; CurranceFactorDetail)
            {
            }
            column(AmountLCY_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Amount (LCY)")
            {
            }
            column(EntryNo_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Entry No.")
            {
            }
            column(CustLedgerEntryNo_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Cust. Ledger Entry No.")
            {
            }
            column(PostingDate_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Posting Date")
            {
            }
            column(DocumentNo_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Document No.")
            {
            }
            dataitem("Cust. Ledger Entry"; "Cust. Ledger Entry")
            {
                DataItemLink = "Entry No." = FIELD("Cust. Ledger Entry No.");
                DataItemTableView = WHERE("Document Type" = FILTER(Invoice));
                column(CustomerDetail1; CustomerDetail)
                {
                }
                column(DocumentType_CustLedgerEntry; "Cust. Ledger Entry"."Document Type")
                {
                }
                column(CustomerName; CustomerName)
                {
                }
                column(Exchangefedai; Exchangefedai)
                {
                }
                column(CurranceFactor; CurranceFactor)
                {
                }
                column(PostingDate_CustLedgerEntry; "Cust. Ledger Entry"."Posting Date")
                {
                }
                column(DocumentNo_CustLedgerEntry; "Cust. Ledger Entry"."Document No.")
                {
                }
                column(Amount_CustLedgerEntry; AmtInForeignCurrency)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    CLEAR(CurranceFactor);
                    IF Customer_rec.GET("Cust. Ledger Entry"."Customer No.") THEN
                        CustomerName := Customer_rec.Name;

                    CurranceFactor := ROUND((1 / "Cust. Ledger Entry"."Original Currency Factor"), 0.01);

                    DetailedCustLedgEntry.RESET;
                    DetailedCustLedgEntry.SETCURRENTKEY("Cust. Ledger Entry No.", "Customer No.");
                    DetailedCustLedgEntry.SETRANGE("Cust. Ledger Entry No.", "Entry No.");
                    DetailedCustLedgEntry.SETRANGE("Customer No.", "Customer No.");
                    //DetailedCustLedgEntry.SETFILTER("Posting Date",'%1<',FromDate);
                    IF DetailedCustLedgEntry.FINDFIRST THEN BEGIN
                        REPEAT
                            IF DetailedCustLedgEntry."Source Code" = 'EXCHRATADJ' THEN BEGIN
                                //IF CurrencyExchangeRate.GET(DetailedCustLedgEntry."Currency Code",DetailedCustLedgEntry."Posting Date") THEN BEGIN
                                CurrencyExchangeRate.RESET;
                                CurrencyExchangeRate.SETRANGE("Currency Code", DetailedCustLedgEntry."Currency Code");
                                CurrencyExchangeRate.SETRANGE("Starting Date", DetailedCustLedgEntry."Posting Date");
                                IF CurrencyExchangeRate.FINDFIRST THEN BEGIN
                                    IF DetailedCustLedgEntry."Posting Date" < FromDate THEN
                                        Exchangefedai := CurrencyExchangeRate."Relational Exch. Rate Amount";
                                END;
                            END;
                        UNTIL DetailedCustLedgEntry.NEXT = 0;
                    END;

                    //RSPLSUM 30Dec2020>>
                    CLEAR(ApplicationTillFromDate);
                    RecDCLE.RESET;
                    RecDCLE.SETCURRENTKEY("Cust. Ledger Entry No.", "Entry Type", "Posting Date", "Customer No.");
                    RecDCLE.SETRANGE("Cust. Ledger Entry No.", "Entry No.");
                    RecDCLE.SETRANGE("Entry Type", RecDCLE."Entry Type"::Application);
                    RecDCLE.SETFILTER("Posting Date", '<%1', FromDate);
                    RecDCLE.SETRANGE("Customer No.", "Customer No.");
                    IF RecDCLE.FINDSET THEN
                        REPEAT
                            ApplicationTillFromDate += ABS(RecDCLE.Amount);
                        UNTIL RecDCLE.NEXT = 0;

                    CLEAR(AmtInForeignCurrency);
                    "Cust. Ledger Entry".CALCFIELDS("Cust. Ledger Entry".Amount);
                    AmtInForeignCurrency := "Cust. Ledger Entry".Amount - ApplicationTillFromDate;
                    //RSPLSUM 30Dec2020<<
                end;

                trigger OnPreDataItem()
                begin

                    CLEAR(Exchangefedai);
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //RSPLSUM 05Jan21>>
                IF ("Detailed Cust. Ledg. Entry"."Entry Type" = "Detailed Cust. Ledg. Entry"."Entry Type"::Application) AND
                    (("Detailed Cust. Ledg. Entry"."Posting Date" < FromDate) OR ("Detailed Cust. Ledg. Entry"."Posting Date" > ToDate)) THEN
                    CurrReport.SKIP;

                IF (("Detailed Cust. Ledg. Entry"."Entry Type" = "Detailed Cust. Ledg. Entry"."Entry Type"::"Unrealized Gain") OR
                    ("Detailed Cust. Ledg. Entry"."Entry Type" = "Detailed Cust. Ledg. Entry"."Entry Type"::"Unrealized Loss")) AND
                      ("Detailed Cust. Ledg. Entry"."Posting Date" <> FromDate - 1) THEN
                    CurrReport.SKIP;

                IF (("Detailed Cust. Ledg. Entry"."Entry Type" = "Detailed Cust. Ledg. Entry"."Entry Type"::"Unrealized Gain") OR
                    ("Detailed Cust. Ledg. Entry"."Entry Type" = "Detailed Cust. Ledg. Entry"."Entry Type"::"Unrealized Loss")) THEN BEGIN
                    RecDetCustLedEnt.RESET;
                    RecDetCustLedEnt.SETRANGE("Cust. Ledger Entry No.", "Cust. Ledger Entry No.");
                    RecDetCustLedEnt.SETRANGE("Customer No.", "Customer No.");
                    RecDetCustLedEnt.SETRANGE("Entry Type", RecDetCustLedEnt."Entry Type"::Application);
                    RecDetCustLedEnt.SETRANGE("Posting Date", FromDate, ToDate);
                    IF RecDetCustLedEnt.FINDFIRST THEN
                        CurrReport.SKIP;
                END;

                EntryUnrealizedGainLoss := FALSE;
                IF (("Detailed Cust. Ledg. Entry"."Entry Type" = "Detailed Cust. Ledg. Entry"."Entry Type"::"Unrealized Gain") OR
                    ("Detailed Cust. Ledg. Entry"."Entry Type" = "Detailed Cust. Ledg. Entry"."Entry Type"::"Unrealized Loss")) THEN
                    EntryUnrealizedGainLoss := TRUE;
                //RSPLSUM 05Jan21<<

                //DJ05012021
                IF CustomerLedgerEntryNo = "Cust. Ledger Entry No." THEN
                    CurrReport.SKIP;
                CustomerLedgerEntryNo := "Cust. Ledger Entry No.";
                //DJ05012021
                CLEAR(CurranceFactorDetail);
                CustLedgerEntry_rec.RESET;
                CustLedgerEntry_rec.SETRANGE("Document No.", "Document No.");
                IF CustLedgerEntry_rec.FINDFIRST THEN
                    CurranceFactorDetail := ROUND((1 / CustLedgerEntry_rec."Original Currency Factor"), 0.01);
            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("Cust. Ledger Entry No."); //DJ05012021
                //RSPLSUM 05Jan21--SETRANGE("Detailed Cust. Ledg. Entry"."Posting Date",FromDate,ToDate);
                SETRANGE("Detailed Cust. Ledg. Entry"."Posting Date", FromDate - 1, ToDate);//RSPLSUM 05Jan21
            end;
        }
        dataitem("Detailed Vendor Ledg. Entry"; "Detailed Vendor Ledg. Entry")
        {
            DataItemTableView = WHERE("Currency Code" = FILTER(<> ''),
                                      "Entry Type" = FILTER(Application | "Unrealized Gain" | "Unrealized Loss"));
            RequestFilterFields = "Document No.", "Posting Date";
            column(EntryUnrealizedGainLossVend; EntryUnrealizedGainLoss)
            {
            }
            column(VendorDetail; VendorDetail)
            {
            }
            column(CurranceFactorForVendorDetail; CurranceFactorForVendorDetail)
            {
            }
            column(EntryNo_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Entry No.")
            {
            }
            column(VendorLedgerEntryNo_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Vendor Ledger Entry No.")
            {
            }
            column(PostingDate_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Posting Date")
            {
            }
            column(DocumentNo_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Document No.")
            {
            }
            column(AmountLCY_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Amount (LCY)")
            {
            }
            dataitem("Vendor Ledger Entry"; "Vendor Ledger Entry")
            {
                DataItemLink = "Entry No." = FIELD("Vendor Ledger Entry No.");
                DataItemTableView = WHERE("Document Type" = FILTER(Invoice));
                column(VendorDetail1; VendorDetail)
                {
                }
                column(DocumentType_VendorLedgerEntry; "Vendor Ledger Entry"."Document Type")
                {
                }
                column(CurranceFactorForVendor; CurranceFactorForVendor)
                {
                }
                column(ExchangefedaiVendor; ExchangefedaiVendor)
                {
                }
                column(VendorName; VendorName)
                {
                }
                column(PostingDate_VendorLedgerEntry; "Vendor Ledger Entry"."Posting Date")
                {
                }
                column(DocumentNo_VendorLedgerEntry; "Vendor Ledger Entry"."Document No.")
                {
                }
                column(Amount_VendorLedgerEntry; AmtInForeignCurrencyVendor)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    CLEAR(CurranceFactor);

                    IF Vendor_Rec.GET("Vendor Ledger Entry"."Vendor No.") THEN
                        VendorName := Vendor_Rec.Name;


                    CurranceFactorForVendor := ROUND((1 / "Vendor Ledger Entry"."Original Currency Factor"), 0.01);

                    DetailedVendorLedgEntryRec.RESET;
                    DetailedVendorLedgEntryRec.SETCURRENTKEY("Vendor Ledger Entry No.", "Vendor No.");
                    DetailedVendorLedgEntryRec.SETRANGE("Vendor Ledger Entry No.", "Entry No.");
                    DetailedVendorLedgEntryRec.SETRANGE("Vendor No.", "Vendor No.");
                    IF DetailedVendorLedgEntryRec.FINDFIRST THEN BEGIN
                        REPEAT
                            IF DetailedVendorLedgEntryRec."Source Code" = 'EXCHRATADJ' THEN BEGIN
                                //IF CurrencyExchangeRate.GET(DetailedVendorLedgEntryRec."Currency Code",DetailedVendorLedgEntryRec."Posting Date") THEN
                                CurrencyExchangeRate.RESET;
                                CurrencyExchangeRate.SETRANGE("Currency Code", DetailedVendorLedgEntryRec."Currency Code");
                                CurrencyExchangeRate.SETRANGE("Starting Date", DetailedVendorLedgEntryRec."Posting Date");
                                IF CurrencyExchangeRate.FINDFIRST THEN BEGIN
                                    IF DetailedVendorLedgEntryRec."Posting Date" < FromDate THEN
                                        ExchangefedaiVendor := CurrencyExchangeRate."Relational Exch. Rate Amount";
                                END;
                            END;
                        UNTIL DetailedVendorLedgEntryRec.NEXT = 0;
                    END;

                    //RSPLSUM 30Dec2020>>
                    CLEAR(ApplicationTillFromDateVendor);
                    RecDVLE.RESET;
                    RecDVLE.SETCURRENTKEY("Vendor Ledger Entry No.", "Entry Type", "Posting Date", "Vendor No.");
                    RecDVLE.SETRANGE("Vendor Ledger Entry No.", "Entry No.");
                    RecDVLE.SETRANGE("Entry Type", RecDVLE."Entry Type"::Application);
                    RecDVLE.SETFILTER("Posting Date", '<%1', FromDate);
                    RecDVLE.SETRANGE("Vendor No.", "Vendor No.");
                    IF RecDVLE.FINDSET THEN
                        REPEAT
                            ApplicationTillFromDateVendor += ABS(RecDVLE.Amount);
                        UNTIL RecDVLE.NEXT = 0;

                    CLEAR(AmtInForeignCurrencyVendor);
                    "Vendor Ledger Entry".CALCFIELDS("Vendor Ledger Entry".Amount);
                    AmtInForeignCurrencyVendor := "Vendor Ledger Entry".Amount - ApplicationTillFromDateVendor;
                    //RSPLSUM 30Dec2020<<
                end;

                trigger OnPreDataItem()
                begin
                    CLEAR(ExchangefedaiVendor);
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //RSPLSUM 05Jan21>>
                IF ("Detailed Vendor Ledg. Entry"."Entry Type" = "Detailed Vendor Ledg. Entry"."Entry Type"::Application) AND
                    (("Detailed Vendor Ledg. Entry"."Posting Date" < FromDate) OR ("Detailed Vendor Ledg. Entry"."Posting Date" > ToDate)) THEN
                    CurrReport.SKIP;

                IF (("Detailed Vendor Ledg. Entry"."Entry Type" = "Detailed Vendor Ledg. Entry"."Entry Type"::"Unrealized Gain") OR
                    ("Detailed Vendor Ledg. Entry"."Entry Type" = "Detailed Vendor Ledg. Entry"."Entry Type"::"Unrealized Loss")) AND
                      ("Detailed Vendor Ledg. Entry"."Posting Date" <> FromDate - 1) THEN
                    CurrReport.SKIP;

                IF (("Detailed Vendor Ledg. Entry"."Entry Type" = "Detailed Vendor Ledg. Entry"."Entry Type"::"Unrealized Gain") OR
                    ("Detailed Vendor Ledg. Entry"."Entry Type" = "Detailed Vendor Ledg. Entry"."Entry Type"::"Unrealized Loss")) THEN BEGIN
                    RecDetVendLedEnt.RESET;
                    RecDetVendLedEnt.SETRANGE("Vendor Ledger Entry No.", "Vendor Ledger Entry No.");
                    RecDetVendLedEnt.SETRANGE("Vendor No.", "Vendor No.");
                    RecDetVendLedEnt.SETRANGE("Entry Type", RecDetVendLedEnt."Entry Type"::Application);
                    RecDetVendLedEnt.SETRANGE("Posting Date", FromDate, ToDate);
                    IF RecDetVendLedEnt.FINDFIRST THEN
                        CurrReport.SKIP;
                END;

                EntryUnrealizedGainLoss := FALSE;
                IF (("Detailed Vendor Ledg. Entry"."Entry Type" = "Detailed Vendor Ledg. Entry"."Entry Type"::"Unrealized Gain") OR
                    ("Detailed Vendor Ledg. Entry"."Entry Type" = "Detailed Vendor Ledg. Entry"."Entry Type"::"Unrealized Loss")) THEN
                    EntryUnrealizedGainLoss := TRUE;
                //RSPLSUM 05Jan21<<

                //DJ05012021
                IF VendorLedgerEntryNo = "Vendor Ledger Entry No." THEN
                    CurrReport.SKIP;
                VendorLedgerEntryNo := "Vendor Ledger Entry No.";
                //DJ05012021
                CLEAR(CurranceFactorForVendorDetail);
                VendorLedgerEntry_Rec.RESET;
                VendorLedgerEntry_Rec.SETRANGE("Document No.", "Document No.");
                IF VendorLedgerEntry_Rec.FINDFIRST THEN
                    CurranceFactorForVendorDetail := ROUND((1 / VendorLedgerEntry_Rec."Original Currency Factor"), 0.01);
            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("Vendor Ledger Entry No.");//DJ05012021
                //RSPLSUM 05Jan21--SETRANGE("Detailed Vendor Ledg. Entry"."Posting Date",FromDate,ToDate);
                SETRANGE("Detailed Vendor Ledg. Entry"."Posting Date", FromDate - 1, ToDate);//RSPLSUM 05Jan21
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("From Date"; FromDate)
                {
                }
                field("To Date"; ToDate)
                {
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

    var
        Customer_rec: Record 18;
        CustomerName: Text;
        CurranceFactor: Decimal;
        CurranceFactorDetail: Decimal;
        CustLedgerEntry_rec: Record 21;
        CurranceFactorForVendor: Decimal;
        VendorName: Text;
        Vendor_Rec: Record 23;
        VendorLedgerEntry_Rec: Record 25;
        CurranceFactorForVendorDetail: Decimal;
        DetailedCustLedgEntry: Record 379;
        CurrencyExchangeRate: Record 330;
        Exchangefedai: Decimal;
        ExchangefedaiVendor: Decimal;
        DetailedVendorLedgEntryRec: Record 380;
        AmountExchange: Decimal;
        FromDate: Date;
        ToDate: Date;
        VendorDetail: Boolean;
        CustomerDetail: Boolean;
        RecDCLE: Record 379;
        ApplicationTillFromDate: Decimal;
        AmtInForeignCurrency: Decimal;
        RecDVLE: Record 380;
        ApplicationTillFromDateVendor: Decimal;
        AmtInForeignCurrencyVendor: Decimal;
        VendorLedgerEntryNo: Integer;
        CustomerLedgerEntryNo: Integer;
        RecDetCustLedEnt: Record 379;
        EntryUnrealizedGainLoss: Boolean;
        RecDetVendLedEnt: Record 380;
}


report 50218 "Fedai Register"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 22Oct2018   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/FedaiRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Fedai Register';

    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("Entry No.")
                                WHERE("System-Created Entry" = CONST(true),
                                      "Source Code" = CONST('EXCHRATADJ'));
            RequestFilterFields = "Posting Date", "Document No.", "G/L Account No.";
            dataitem("Detailed Cust. Ledg. Entry"; "Detailed Cust. Ledg. Entry")
            {
                DataItemLink = "Document No." = FIELD("Document No.");
                DataItemTableView = SORTING("Document No.", "Document Type", "Posting Date");
                column(PostingDate_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Posting Date")
                {
                }
                column(DocumentType_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Document Type")
                {
                }
                column(DocumentNo_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Document No.")
                {
                }
                column(AmountLCY_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Amount (LCY)")
                {
                }
                column(CustomerNo_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Customer No.")
                {
                }
                column(CurrencyCode_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Currency Code")
                {
                }
                column(EntryNo_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Entry No.")
                {
                }
                column(CustLedgerEntryNo_DetailedCustLedgEntry; "Detailed Cust. Ledg. Entry"."Cust. Ledger Entry No.")
                {
                }
                column(CustName; CustName)
                {
                }
                column(CustDocNo; CustDocNo)
                {
                }
                column(CustExDocNo; CustExDocNo)
                {
                }
                column(CustExRate; CustExRate)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    CustName := '';
                    CustDocNo := '';
                    CustExDocNo := '';
                    CustExRate := 0;

                    Cust.RESET;
                    IF Cust.GET("Detailed Cust. Ledg. Entry"."Customer No.") THEN
                        CustName := Cust.Name;

                    CLE.RESET;
                    IF CLE.GET("Detailed Cust. Ledg. Entry"."Cust. Ledger Entry No.") THEN BEGIN
                        CustDocNo := CLE."Document No.";
                        CustExDocNo := CLE."External Document No.";
                        IF CLE."Original Currency Factor" <> 0 THEN
                            CustExRate := 1 / CLE."Original Currency Factor"
                        ELSE
                            CustExRate := 1;
                    END;

                    CurrExRate.RESET;
                    CurrExRate.SETRANGE("Currency Code", "Currency Code");
                    CurrExRate.SETRANGE("Starting Date", 0D, "G/L Entry"."Posting Date");
                    IF CurrExRate.FINDLAST THEN
                        CustExRate := CurrExRate."Relational Adjmt Exch Rate Amt";
                end;
            }
            dataitem("Detailed Vendor Ledg. Entry"; "Detailed Vendor Ledg. Entry")
            {
                DataItemLink = "Document No." = FIELD("Document No.");
                DataItemTableView = SORTING("Document No.", "Document Type", "Posting Date");
                column(EntryNo_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Entry No.")
                {
                }
                column(VendorLedgerEntryNo_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Vendor Ledger Entry No.")
                {
                }
                column(PostingDate_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Posting Date")
                {
                }
                column(DocumentType_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Document Type")
                {
                }
                column(DocumentNo_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Document No.")
                {
                }
                column(AmountLCY_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Amount (LCY)")
                {
                }
                column(VendorNo_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Vendor No.")
                {
                }
                column(CurrencyCode_DetailedVendorLedgEntry; "Detailed Vendor Ledg. Entry"."Currency Code")
                {
                }
                column(VenName; VenName)
                {
                }
                column(VenDocNo; VenDocNo)
                {
                }
                column(VenExDocNo; VenExDocNo)
                {
                }
                column(VenExRate; VenExRate)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    VenName := '';
                    VenDocNo := '';
                    VenExDocNo := '';
                    VenExRate := 0;

                    Ven.RESET;
                    IF Ven.GET("Detailed Vendor Ledg. Entry"."Vendor No.") THEN
                        VenName := Ven.Name;


                    VLE.RESET;
                    IF VLE.GET("Detailed Vendor Ledg. Entry"."Vendor Ledger Entry No.") THEN BEGIN
                        VenDocNo := VLE."Document No.";
                        VenExDocNo := VLE."External Document No.";
                        IF VLE."Original Currency Factor" <> 0 THEN
                            VenExRate := 1 / VLE."Original Currency Factor"
                        ELSE
                            VenExRate := 1;
                    END;

                    CurrExRate.RESET;
                    CurrExRate.SETRANGE("Currency Code", "Currency Code");
                    CurrExRate.SETRANGE("Starting Date", 0D, "G/L Entry"."Posting Date");
                    IF CurrExRate.FINDLAST THEN
                        CustExRate := CurrExRate."Relational Adjmt Exch Rate Amt";
                end;
            }

            trigger OnAfterGetRecord()
            begin

                IF TempDocNo <> "G/L Entry"."Document No." THEN BEGIN
                    TempDocNo := "G/L Entry"."Document No.";
                END ELSE
                    CurrReport.SKIP;
            end;
        }
    }

    requestpage
    {

        layout
        {
        }

        actions
        {
        }
    }

    labels
    {
    }

    var
        Cust: Record 18;
        CustName: Text;
        CustDocNo: Code[20];
        CustExDocNo: Code[50];
        Ven: Record 23;
        VenName: Text;
        VenDocNo: Code[20];
        VenExDocNo: Code[35];
        CLE: Record 21;
        VLE: Record 25;
        TempDocNo: Code[20];
        CustExRate: Decimal;
        VenExRate: Decimal;
        CurrExRate: Record 330;
}


report 50088 "Advance Payment Ledger"
{
    // 11Aug2017 :: RB-N, New Report Advence Payment Ledger
    DefaultLayout = RDLC;
    RDLCLayout = 'src/GPPL/Report Layout/PlantInvILN.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    //RDLCLayout = 'src/GPPL/Report Layout/PlantInvILN.rdl';


    dataset
    {
        dataitem("Bank Account Ledger Entry"; "Bank Account Ledger Entry")
        {
            DataItemTableView = SORTING("Posting Date")
                                WHERE("Bal. Account Type" = CONST(Vendor));
            RequestFilterFields = "Posting Date";
            column(CreditAmount_BankAccountLedgerEntry; "Bank Account Ledger Entry"."Credit Amount")
            {
            }
            column(EntryNo_BankAccountLedgerEntry; "Bank Account Ledger Entry"."Entry No.")
            {
            }
            column(BankAccountNo_BankAccountLedgerEntry; "Bank Account Ledger Entry"."Bank Account No.")
            {
            }
            column(PostingDate_BankAccountLedgerEntry; "Bank Account Ledger Entry"."Posting Date")
            {
            }
            column(DocumentNo_BankAccountLedgerEntry; "Bank Account Ledger Entry"."Document No.")
            {
            }
            column(Description_BankAccountLedgerEntry; "Bank Account Ledger Entry".Description)
            {
            }
            column(Amount_BankAccountLedgerEntry; "Bank Account Ledger Entry".Amount)
            {
            }
            column(GSTNo; GSTNo)
            {
            }
            column(BalAccountNo_BankAccountLedgerEntry; "Bank Account Ledger Entry"."Bal. Account No.")
            {
            }
            column(VendorName; Ven.Name)
            {
            }

            trigger OnAfterGetRecord()
            begin

                GSTNo := '';
                IF Ven.GET("Bank Account Ledger Entry"."Bal. Account No.") THEN
                    GSTNo := Ven."GST Registration No.";
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
        GSTNo: Code[15];
        Ven: Record 23;
}


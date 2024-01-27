report 50338 "Print CLGL Entries"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PrintCLGLEntries.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Cust. Ledger Entry"; "Cust. Ledger Entry")
        {
            RequestFilterFields = "Posting Date";
            column(EntryNo_CustLedgerEntry; "Cust. Ledger Entry"."Entry No.")
            {
            }
            column(AmountLCY_CustLedgerEntry; "Cust. Ledger Entry"."Amount (LCY)")
            {
            }
            dataitem("G/L Entry"; "G/L Entry")
            {
                DataItemLink = "Entry No." = FIELD("Entry No.");
                column(GLAccountNo_GLEntry; "G/L Entry"."G/L Account No.")
                {
                }
                column(Amount_GLEntry; "G/L Entry".Amount)
                {
                }
            }
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
}


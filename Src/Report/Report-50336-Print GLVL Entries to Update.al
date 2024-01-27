report 50336 "Print GLVL Entries to Update"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PrintGLVLEntriestoUpdate.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("G/L - Item Ledger Relation"; "G/L - Item Ledger Relation")
        {
            RequestFilterFields = "G/L Register No.";
            dataitem("G/L Entry"; "G/L Entry")
            {
                DataItemLink = "Entry No." = FIELD("G/L Entry No.");
                DataItemTableView = WHERE("Posting Date" = FILTER('<= 31-03-20'));
                column(EntryNo_GLEntry; "G/L Entry"."Entry No.")
                {
                }
                column(Amount_GLEntry; "G/L Entry".Amount)
                {
                }
                column(DebitAmount_GLEntry; "G/L Entry"."Debit Amount")
                {
                }
                column(CreditAmount_GLEntry; "G/L Entry"."Credit Amount")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    IF "G/L Entry"."Posting Date" > 20200331D/*310320D*/ THEN
                        CurrReport.SKIP;
                end;
            }
            dataitem("Value Entry"; "Value Entry")
            {
                DataItemLink = "Entry No." = FIELD("Value Entry No.");
                DataItemTableView = WHERE("Posting Date" = FILTER('<= 31-03-20'));
                column(EntryNo_ValueEntry; "Value Entry"."Entry No.")
                {
                }
                column(CostperUnit_ValueEntry; "Value Entry"."Cost per Unit")
                {
                }
                column(CostAmountActual_ValueEntry; "Value Entry"."Cost Amount (Actual)")
                {
                }
                column(CostPostedtoGL_ValueEntry; "Value Entry"."Cost Posted to G/L")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    IF "Value Entry"."Posting Date" > 20200331D/*310320D*/ THEN
                        CurrReport.SKIP;
                end;
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

    trigger OnPreReport()
    begin
        GlRegNo := "G/L - Item Ledger Relation".GETFILTER("G/L - Item Ledger Relation"."G/L Register No.");
        IF GlRegNo = '' THEN
            ERROR('Please enter G/L Reg. No.');
    end;

    var
        GlRegNo: Text;
}


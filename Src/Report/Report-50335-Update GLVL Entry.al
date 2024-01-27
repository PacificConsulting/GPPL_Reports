report 50335 "Update GLVL Entry"
{
    DefaultLayout = RDLC;

    RDLCLayout = 'Src/Report Layout/UpdateGLVLEntry.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Permissions = TableData 17 = rm,
                  TableData 5802 = rm;

    dataset
    {
        dataitem("G/L - Item Ledger Relation"; "G/L - Item Ledger Relation")
        {
            RequestFilterFields = "G/L Register No.";
            dataitem("G/L Entry"; "G/L Entry")
            {
                DataItemLink = "Entry No." = FIELD("G/L Entry No.");
                DataItemTableView = WHERE("Posting Date" = FILTER('<= 31-03-20'));

                trigger OnAfterGetRecord()
                begin
                    IF "G/L Entry"."Posting Date" > 20200331D/*310320D*/ THEN
                        CurrReport.SKIP;

                    "G/L Entry"."Debit Amount" := 0;
                    "G/L Entry"."Credit Amount" := 0;
                    "G/L Entry".Amount := 0;
                    "G/L Entry".MODIFY;
                end;
            }
            dataitem("Value Entry"; "Value Entry")
            {
                DataItemLink = "Entry No." = FIELD("Value Entry No.");
                DataItemTableView = WHERE("Posting Date" = FILTER('<= 31-03-20'));

                trigger OnAfterGetRecord()
                begin
                    IF "Value Entry"."Posting Date" > 20200331D/*310320D*/ THEN
                        CurrReport.SKIP;

                    "Value Entry"."Cost per Unit" := 0;
                    "Value Entry"."Cost Amount (Actual)" := 0;
                    "Value Entry"."Cost Posted to G/L" := 0;
                    "Value Entry".MODIFY;
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

    trigger OnPostReport()
    begin
        MESSAGE('Done');
    end;

    trigger OnPreReport()
    begin
        GlRegNo := "G/L - Item Ledger Relation".GETFILTER("G/L - Item Ledger Relation"."G/L Register No.");
        IF GlRegNo = '' THEN
            ERROR('Please enter G/L Reg. No.');
    end;

    var
        GlRegNo: Text;
}


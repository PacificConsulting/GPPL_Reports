report 50337 "Update GL Entry"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/UpdateGLEntry.rdl';
    Permissions = TableData 17 = rm;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = WHERE("Posting Date" = FILTER('<= 31-03-20'));
            RequestFilterFields = "Entry No.";

            trigger OnAfterGetRecord()
            begin
                IF "G/L Entry"."Posting Date" > 20200331D/*310320D*/ THEN
                    CurrReport.SKIP;

                IF ("G/L Entry"."Entry No." >= "G/L Entry".GETRANGEMIN("G/L Entry"."Entry No.")) AND
                    ("G/L Entry"."Entry No." <= "G/L Entry".GETRANGEMAX("G/L Entry"."Entry No.")) THEN BEGIN
                    "G/L Entry"."Debit Amount" := 0;
                    "G/L Entry"."Credit Amount" := 0;
                    "G/L Entry".Amount := 0;
                    "G/L Entry".MODIFY;
                END;
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

    trigger OnPostReport()
    begin
        MESSAGE('Done');
    end;

    trigger OnPreReport()
    begin
        EntryNoFilterStart := FORMAT("G/L Entry".GETRANGEMIN("G/L Entry"."Entry No."));
        EntryNoFilterEnd := FORMAT("G/L Entry".GETRANGEMAX("G/L Entry"."Entry No."));

        IF (EntryNoFilterStart = '') OR (EntryNoFilterEnd = '') THEN
            ERROR('Please enter Entry No. in range');
    end;

    var
        EntryNoFilterStart: Text;
        EntryNoFilterEnd: Text;
}


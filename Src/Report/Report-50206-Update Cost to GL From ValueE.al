report 50206 "Update Cost to GL From ValueE"
{
    Permissions = TableData 5802 = rim;
    ProcessingOnly = true;

    dataset
    {
        dataitem("Value Entry"; "Value Entry")
        {
            DataItemTableView = SORTING("Entry No.")
                                ORDER(Ascending)
                                WHERE("Entry Type" = CONST(Revaluation),
                                      "Source Code" = CONST('TRANSFER'));
            RequestFilterFields = "Item No.", "Inventory Posting Group";

            trigger OnAfterGetRecord()
            begin
                "Value Entry"."OLD Cost Posted to G/L" := "Value Entry"."Cost Posted to G/L";
                "Value Entry"."Cost Posted to G/L" := "Value Entry"."Cost Amount (Actual)";
                MODIFY;
                // MESSAGE('DONE');
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
        MESSAGE('Updated');
    end;

    trigger OnPreReport()
    begin
        IF "Value Entry".GETFILTER("Value Entry"."Inventory Posting Group") = '' THEN
            ERROR('Please Select Inventory Posting Group');

        IF "Value Entry".GETFILTER("Value Entry"."Posting Date") = '' THEN
            ERROR('Please Enter Posting Date');
    end;
}


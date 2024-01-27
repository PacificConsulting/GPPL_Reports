pageextension 80023 PostedInwardGateEntryExt extends "Posted Inward Gate Entry"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        // Add changes to page actions here
        addlast(Reporting)
        {
            action("Inward Vehicle Posted")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                var
                    PstGateEntry: Record "Posted Gate Entry Header";
                BEGIN
                    PstGateEntry.RESET;
                    PstGateEntry.SETRANGE(PstGateEntry."Entry Type", rec."Entry Type");
                    PstGateEntry.SETRANGE(PstGateEntry."No.", rec."No.");
                    REPORT.RUN(50056, TRUE, FALSE, PstGateEntry);
                END;
            }

        }
    }

    var
        myInt: Integer;


}
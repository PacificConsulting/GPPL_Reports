pageextension 80025 PostedOutwardGateEntryExt extends "Posted Outward Gate Entry"
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
            action("Outward Vehicle Posted")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    PstGateEntry: Record "Posted Gate Entry Header";
                BEGIN
                    PstGateEntry.RESET;
                    PstGateEntry.SETRANGE(PstGateEntry."Entry Type", rec."Entry Type");
                    PstGateEntry.SETRANGE(PstGateEntry."No.", rec."No.");
                    REPORT.RUN(50058, TRUE, FALSE, PstGateEntry);
                END;
            }
        }
    }

    var
        myInt: Integer;


}

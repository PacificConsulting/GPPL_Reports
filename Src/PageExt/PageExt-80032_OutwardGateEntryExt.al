pageextension 80032 OutwardGateEntryExt extends "Outward Gate Entry"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        // Add changes to page actions here
        addlast(reporting)
        {
            action("&Print")
            {

                Promoted = true;
                Image = Print;
                PromotedCategory = Process;
                ApplicationArea = ALL;
                trigger OnAction()
                var

                    GateENtry: Record "Gate Entry Header";
                BEGIN
                    GateENtry.RESET;
                    GateENtry := Rec;
                    GateENtry.SETRECFILTER;
                    REPORT.RUN(50149, TRUE, FALSE, GateENtry);
                END;
            }
        }
    }

    var
        myInt: Integer;
}
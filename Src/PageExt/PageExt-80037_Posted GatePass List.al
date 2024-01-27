pageextension 80037 PostedGatePassList extends "Posted GatePass List"
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
            action("Posted GatePass2")
            {
                ApplicationArea = all;
                trigger OnAction()
                BEGIN
                    CurrPage.SETSELECTIONFILTER(Rec);
                    REPORT.RUNMODAL(50219, TRUE, FALSE, Rec)
                END;
            }
        }
    }

    var
        myInt: Integer;
}
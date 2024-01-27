pageextension 80040 OutwardGateEntryLogisticsExt extends "Outward Gate Entry-Logistics"
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
            action("Vehicle Packing List2")
            {
                Promoted = true;
                Visible = TRUE;
                PromotedIsBig = true;
                Image = Report;
                ApplicationArea = all;
                trigger OnAction()
                var
                    GateEntryHead: Record "Gate Entry Header";
                BEGIN
                    GateEntryHead.RESET;
                    GateEntryHead.SETRANGE(GateEntryHead."Entry Type", rec."Entry Type");
                    GateEntryHead.SETRANGE(GateEntryHead."No.", rec."No.");
                    REPORT.RUN(50249, TRUE, FALSE, GateEntryHead);
                END;
            }

        }
    }

    var
        myInt: Integer;
}
pageextension 80042 MonthlyAttendanceExt extends "Leave Master"
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
            action("Create &Leaves2")
            {
                Promoted = true;
                PromotedCategory = Process;
                ApplicationArea = all;
                trigger OnAction()
                var
                    LeaveMaster: Record "Leave Master";

                BEGIN
                    // CLEAR(GenerateLeave);
                    IF LeaveMaster.FIND('-') THEN
                        REPEAT
                            LeaveMaster.TESTFIELD("No. of Leaves in Year");
                            LeaveMaster.TESTFIELD("Crediting Interval");
                            LeaveMaster.TESTFIELD("Applicable Date");
                            IF LeaveMaster."Carry Forward" THEN
                                LeaveMaster.TESTFIELD(LeaveMaster."Max.Leaves to Carry Forward");
                            IF LeaveMaster."Applicable During Probation" THEN
                                LeaveMaster.TESTFIELD(LeaveMaster."No.of Leaves During Probation");
                        UNTIL LeaveMaster.NEXT = 0;
                    // GenerateLeave.RUNMODAL;
                END;
            }
        }
    }

    var
        myInt: Integer;
}
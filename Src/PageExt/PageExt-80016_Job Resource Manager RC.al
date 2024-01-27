pageextension 80016 "JobResourceManagerRCEtrpt" extends "Job Resource Manager RC"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addafter("Navi&gate")
        {
            group("Purchase Reports")
            {


                action("Item Specification-RM")
                {
                    ApplicationArea = All;

                    RunObject = report 50028;
                }
                action("QC Test Report for Plant")
                {
                    ApplicationArea = All;

                    RunObject = report 50090;
                }
                action("Item Specification")
                {
                    ApplicationArea = All;

                    RunObject = report 50168;
                }
            }
        }
    }

    var
        myInt: Integer;
}
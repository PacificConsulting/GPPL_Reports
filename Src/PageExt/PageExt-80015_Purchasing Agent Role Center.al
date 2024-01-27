pageextension 80015 "Purchasgent RoleCenterEtrpt" extends "Purchasing Agent Role Center"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addafter("Posted Documents")
        {
            group("Purchase Reports")
            {

                action("Vendor Item Analysis")
                {
                    ApplicationArea = All;

                    RunObject = report 50023;
                }
                action("Bin-wise Raw Material - Plant")
                {
                    ApplicationArea = All;

                    RunObject = report 50024;
                }
                action("Gate Entry Vehicle Status I/W")
                {
                    ApplicationArea = All;

                    RunObject = report 50043;
                }
                action("BRT Register")
                {
                    ApplicationArea = All;

                    RunObject = report 50086;
                }
                action("Purchase Register GST")
                {
                    ApplicationArea = All;

                    RunObject = report 50093;
                }
                action("RM Consumption Detailed")
                {
                    ApplicationArea = All;

                    RunObject = report 50096;
                }
                action("RM Consumption Summary")
                {
                    ApplicationArea = All;

                    RunObject = report 50097;
                }
                action("Vendor Ageing Summary")
                {
                    ApplicationArea = All;

                    RunObject = report 50104;
                }
                action("Vendor Balance Confirmation")
                {
                    ApplicationArea = All;

                    RunObject = report 50127;
                }
                action("Vendor - Ledger Acconut")
                {
                    ApplicationArea = All;

                    RunObject = report 50129;
                }
                action("Purchase Register")
                {
                    ApplicationArea = All;

                    RunObject = report 50151;
                }
                action("Raw Material - Ageing")
                {
                    ApplicationArea = All;

                    RunObject = report 50162;
                }
                action("GRN Register")
                {
                    ApplicationArea = All;

                    RunObject = report 50164;
                }
                action("GRN Pending for Invoice")
                {
                    ApplicationArea = All;

                    RunObject = report 50171;
                }
                action("PO Statement")
                {
                    ApplicationArea = All;

                    RunObject = report 50196;
                }





            }
        }
    }

    var
        myInt: Integer;


}
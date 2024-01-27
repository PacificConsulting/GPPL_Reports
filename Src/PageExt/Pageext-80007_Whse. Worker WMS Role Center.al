pageextension 80007 "Whse.WorkWMSRoleCenterExtrpt" extends "Whse. Worker WMS Role Center"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addafter("Reference Data")
        {
            group("Inventory Reports")
            {
                action("In-Transit Item Valuation")
                {
                    ApplicationArea = All;
                    RunObject = report 50015;
                }
                action("Daily Production Memo(Excel)")
                {
                    ApplicationArea = All;
                    RunObject = report 50030;
                }
                action("Item by location (RM);")
                {
                    ApplicationArea = All;
                    RunObject = report 50032;
                }
                action("Bin-wise FG Material - Plant")
                {
                    ApplicationArea = all;
                    RunObject = report 50038;
                }
                action("RM Movement (Locwise)")
                {
                    ApplicationArea = all;
                    RunObject = report 50051;
                }
                action("FG Movement (Locwise)")
                {
                    ApplicationArea = all;
                    RunObject = report 50052;
                }
                action("Gate Entry Report I/O")
                {
                    ApplicationArea = all;
                    RunObject = report 50065;
                }
                action("BRT Register - GST")
                {
                    ApplicationArea = all;
                    RunObject = report 50086;
                }
                action("Sales Register - GST")
                {
                    ApplicationArea = all;
                    RunObject = report 50092;
                }
                action("Consumption Report Detailed")
                {
                    ApplicationArea = all;
                    RunObject = report 50096;
                }
                action("=Ledger Entries - RM")
                {
                    ApplicationArea = all;
                    RunObject = report 50098;
                }
                action("Ledger Entries - FG")
                {
                    ApplicationArea = all;
                    RunObject = report 50099;
                }

                action("POP Sales Register")
                {
                    ApplicationArea = all;
                    RunObject = report 50144;
                }
                action("Purchase Register")
                {
                    ApplicationArea = all;
                    RunObject = report 50151;
                }
                action("Finished Goods - Ageing")
                {
                    ApplicationArea = all;
                    RunObject = report 50160;
                }
                action("Raw Material - Ageing")
                {
                    ApplicationArea = all;
                    RunObject = report 50162;
                }
                action("GRN Register")
                {
                    ApplicationArea = all;
                    RunObject = report 50164;
                }
                action("Dimension wise POP Item")
                {
                    ApplicationArea = all;
                    RunObject = report 50174;
                }
                action("Pending Indent/CSO Itemwise")
                {
                    ApplicationArea = all;
                    RunObject = report 50185;
                }
                action("Purchase Order Statement")
                {
                    ApplicationArea = all;
                    RunObject = report 50196;
                }
                action("RM Valuation")
                {
                    ApplicationArea = all;
                    RunObject = report 50216;
                }
                action("Binwise Consumption Report")
                {
                    ApplicationArea = all;
                    RunObject = report 50222;
                }



            }

        }
    }

    var
        myInt: Integer;



}
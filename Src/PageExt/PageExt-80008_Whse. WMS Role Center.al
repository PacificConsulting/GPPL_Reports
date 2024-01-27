pageextension 80008 "Whse. WMS Role CenterExtrpt" extends "Whse. WMS Role Center"
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
                action("Location Wise Finished Goods")
                {
                    ApplicationArea = All;
                    RunObject = report 50016;
                }
                action("Bin-wise Raw Material - Plant")
                {
                    ApplicationArea = All;
                    RunObject = report 50024;
                }
                action("Detailed Sales Register")
                {
                    ApplicationArea = All;
                    RunObject = report 50027;
                }
                action("Daily Production Register")
                {
                    ApplicationArea = All;
                    RunObject = report 50030;
                }
                action("Pop Item by Location")
                {
                    ApplicationArea = All;
                    RunObject = report 50035;
                }
                action("Finished Goods Inventory Bin-Wise")
                {
                    ApplicationArea = All;
                    RunObject = report 50038;
                }
                action("Pending CSO Datewise")
                {
                    ApplicationArea = all;
                    RunObject = report 50040;
                }
                action("FG Movement (Locwise)")
                {
                    ApplicationArea = all;
                    RunObject = report 50052;
                }
                action("BRT Register GST")
                {
                    ApplicationArea = all;
                    RunObject = report 50086;
                }
                action("Sales Register GST")
                {
                    ApplicationArea = all;
                    RunObject = report 50092;
                }
                action("Purchase Register GST")
                {
                    ApplicationArea = all;
                    RunObject = report 50093;
                }
                action("Yield Report")
                {
                    ApplicationArea = all;
                    RunObject = report 50115;
                }
                action("FG Stock Analysis MonthWise")
                {
                    ApplicationArea = all;
                    RunObject = report 50133;
                }
                action("Purchase Register")
                {
                    ApplicationArea = all;
                    RunObject = report 50151;
                }
                action("GRA Register")
                {
                    ApplicationArea = all;
                    RunObject = report 50164;
                }
                action("Pending Indent/CSO Itemwise")
                {
                    ApplicationArea = all;
                    RunObject = report 50185;
                }
                action("Summary Sales/Transfer")
                {
                    ApplicationArea = all;
                    RunObject = report 50186;
                }
                action("Pending Indent  CSO Summary")
                {
                    ApplicationArea = all;
                    RunObject = report 50186;
                }
            }

        }
    }

    var
        myInt: Integer;

}
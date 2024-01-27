pageextension 80011 "ShopSupervisorRoleCenterExtrpt" extends "Shop Supervisor Role Center"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {

        addlast(reporting)
        {
            action("Location Wise Finished Goods")
            {
                ApplicationArea = All;
                RunObject = report 50016;
            }
            action("Binwise Raw Material Plant")
            {
                ApplicationArea = All;
                RunObject = report 50024;
            }
            action("Daily Production Memo")
            {
                ApplicationArea = All;
                RunObject = report 50030;
            }
            action("Binwise FG Material Plant")
            {
                ApplicationArea = All;
                RunObject = report 50038;
            }
            action("RM Movement Locationwise")
            {
                ApplicationArea = All;
                RunObject = report 50051;
            }
            action("FG Movement (Locwise)")
            {
                ApplicationArea = All;
                RunObject = report 50052;
            }
            action("QC Test Report for Plant")
            {
                ApplicationArea = All;
                RunObject = report 50090;
            }
            action("Consumption Report Detailed")
            {
                ApplicationArea = All;
                RunObject = report 50096;
            }
            action("Consumption Report Summary")
            {
                ApplicationArea = All;
                RunObject = report 50097;
            }
            action("Ledger Entries - RM")
            {
                ApplicationArea = All;
                RunObject = report 50098;
            }
            action("Ledger Entries - FG")
            {
                ApplicationArea = All;
                RunObject = report 50099;
            }
            action("Item Wise Sales Report")
            {
                ApplicationArea = All;
                RunObject = report 50113;
            }
            action("Yield Report")
            {
                ApplicationArea = All;
                RunObject = report 50115;
            }
            action("BOM Explosion Prod. Planning")
            {
                ApplicationArea = All;
                RunObject = report 50152;
            }
            action("Finished Goods - Ageing-N")
            {
                ApplicationArea = All;
                RunObject = report 50160;
            }
            action("Raw Material - Ageing")
            {
                ApplicationArea = All;
                RunObject = report 50162;
            }
            action("Pending Indent/CSO Itemwise")
            {
                ApplicationArea = All;
                RunObject = report 50185;
            }
            action("Production Order Comparision")
            {
                ApplicationArea = All;
                RunObject = report 50215;
            }


        }
    }

    var
        myInt: Integer;




}
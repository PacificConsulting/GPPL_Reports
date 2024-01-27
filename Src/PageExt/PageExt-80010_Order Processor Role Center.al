pageextension 80010 "OrdProcessorRole CenterExtrpt" extends "Order Processor Role Center"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {

        addlast(Action62)
        {
            action("Location Wise Finished Goods")
            {
                ApplicationArea = All;
                RunObject = report 50016;
            }
            action("Detailed Sales Register")
            {
                ApplicationArea = All;
                RunObject = report 50027;
            }
            action("G/L Ledger")
            {
                ApplicationArea = All;
                RunObject = report 50029;
            }
            action("Pop Item by Location")
            {
                ApplicationArea = All;
                RunObject = report 50035;
            }
            action("Pending CSO Order Date Wise")
            {
                ApplicationArea = All;
                RunObject = report 50040;
            }
            action("FG Movement Locationwise")
            {
                ApplicationArea = All;
                RunObject = report 50052;
            }
            action("Collection Report")
            {
                ApplicationArea = All;
                RunObject = report 50057;
            }
            action("Customer - Ledger")
            {
                ApplicationArea = All;
                RunObject = report 50081;
            }
            action("Automotive Discount Register")
            {
                ApplicationArea = All;
                RunObject = report 50085;
            }
            action("BRT Register GST")
            {
                ApplicationArea = All;
                RunObject = report 50086;
            }
            action("Sales Register GST")
            {
                ApplicationArea = All;
                RunObject = report 50092;
            }
            action("Customer Ageing Summary")
            {
                ApplicationArea = All;
                RunObject = report 50101;
            }
            action("Customer Ageing Detailed")
            {
                ApplicationArea = All;
                RunObject = report 50103;
            }
            action("Customer Balance Confirmation")
            {
                ApplicationArea = All;
                RunObject = report 50128;
            }
            action("Vendor Ledger")
            {
                ApplicationArea = All;
                RunObject = report 50129;
                Promoted = true;
                PromotedIsBig = true;
                Image = VendorLedger;
                PromotedCategory = Report;
            }
            action("Quarterly E-Return Depot")
            {
                ApplicationArea = All;
                RunObject = report 50131;
                Image = Report;
            }
            action("POP Sales Register")
            {
                ApplicationArea = All;
                RunObject = report 50144;
                Image = Report;
            }
            action("Overdue Reminder")
            {
                ApplicationArea = All;
                RunObject = report 50153;
                Image = Report;
            }
            action("Finished Goods - Ageing-N")
            {
                ApplicationArea = All;
                RunObject = report 50160;
                Image = Report;
            }
            action("CSO Analysis")
            {
                ApplicationArea = All;
                RunObject = report 50167;
                Image = Report;
            }
            action("Transport Details Report")
            {
                ApplicationArea = All;
                RunObject = report 50175;
                Image = Radio;
            }
            action("Pending Indent / CSO Item Wise Depo")
            {
                ApplicationArea = All;
                RunObject = report 50177;
                Image = Radio;
            }
            action("Pending Indent/CSO Item Wise")
            {
                ApplicationArea = All;
                RunObject = report 50185;
                Image = Report;
            }
            action("PO Statement")
            {
                ApplicationArea = All;
                RunObject = report 50196;
                Image = Report;
            }
            action("Customer Inv Application Reg.")
            {
                ApplicationArea = All;
                RunObject = report 50220;
                Image = Report;
            }
            action("Customer Summary Aging GP")
            {
                ApplicationArea = All;
                RunObject = report 50243;
                Image = Report;
            }
        }
    }

    var
        myInt: Integer;




}
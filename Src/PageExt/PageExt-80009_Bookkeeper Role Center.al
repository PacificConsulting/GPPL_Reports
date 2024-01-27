pageextension 80009 "Bookkeeper Role CenterExtrpt" extends "Bookkeeper Role Center"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addafter("&Trial Balance")
        {
            group("Inventory Reports")
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
                action("Item by Location - FG")
                {
                    ApplicationArea = All;
                    RunObject = report 50031;
                }
                action("Item by Location - RM;")
                {
                    ApplicationArea = All;
                    RunObject = report 50032;
                }
                action("Pop Item by Location")
                {
                    ApplicationArea = All;
                    RunObject = report 50035;
                }
                action("Sales Debit Memo - VAT")
                {
                    ApplicationArea = All;
                    RunObject = report 50044;
                }
                action("Collection Report")
                {
                    ApplicationArea = All;
                    RunObject = report 50057;
                }
                action("Sales Cr. Memo Summary Locationwise")
                {
                    ApplicationArea = All;
                    RunObject = report 50062;
                }
                action("Sales Invoice Summary Locationwise")
                {
                    ApplicationArea = All;
                    RunObject = report 50063;
                }
                action("Posted Bank Payment Voucher")
                {
                    ApplicationArea = All;
                    RunObject = report 50068;
                }
                action("Payment Voucher GST")
                {
                    ApplicationArea = All;
                    RunObject = report 50080;
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
                action("Service Tax Register")
                {
                    ApplicationArea = All;
                    RunObject = report 50089;
                }
                action("Sales Register GST")
                {
                    ApplicationArea = All;
                    RunObject = report 50092;
                }
                action("Purchase Register GST")
                {
                    ApplicationArea = All;
                    RunObject = report 50093;
                }
                action("Customer Aging Summary")
                {
                    ApplicationArea = All;
                    RunObject = report 50101;
                    Image = Report;
                }
                action("Excise Duty Differences")
                {
                    ApplicationArea = All;
                    RunObject = report 50102;
                    Image = Report;
                }
                action("Customer Aging Detail")
                {
                    ApplicationArea = All;
                    RunObject = report 50103;
                    Image = Report;
                }
                action("Vendor Aging Summary")
                {
                    ApplicationArea = All;
                    RunObject = report 50104;
                    Image = Report;
                }
                action("Summary of Sales/Transfer")
                {
                    ApplicationArea = All;
                    RunObject = report 50105;
                    Image = Report;
                }
                action("Sales Invoice Summary")
                {
                    ApplicationArea = All;
                    RunObject = report 50108;
                    Image = Report;
                }
                action("Sales Cancel Invoice - Summary")
                {
                    ApplicationArea = All;
                    RunObject = report 50109;
                    Image = Report;
                }
                action("Sales Return - Summary")
                {
                    ApplicationArea = All;
                    RunObject = report 50110;
                    Image = Report;
                }
                action("Sales Credit Memo - Summary")
                {
                    ApplicationArea = All;
                    RunObject = report 50111;
                    Image = Report;
                }
                action("Sales Debit Memo - Summary")
                {
                    ApplicationArea = All;
                    RunObject = report 50112;
                    Image = Report;
                }
                action("Bank Reconciliation Statement")
                {
                    ApplicationArea = All;
                    RunObject = report 50114;
                    Image = Report;
                }
                action("GST Differences")
                {
                    ApplicationArea = All;
                    RunObject = report 50122;
                    Image = Report;
                }
                action("Vendor Balance Confirmation")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50127;
                }
                // action("Page General Journal Batches")
                // {
                //     CaptionML = ENU = 'Vendor Balance Confirmation';
                //     RunObject = Report 50127;
                //     Image = report;
                // }
                action("Customer Balance Confirmation")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50128;
                }
                action("=Vendor - Ledger Acconut")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50129;
                }
                action("Qtrly E-Return Report")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50131;
                }
                action("Purchase Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50151;
                }
                action("Overdue Reminder")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50153;
                }
                action("Finished Goods Ageing")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50160;
                }
                action("Raw Material Aging")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50162;
                }
                action("GRN Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50164;
                }
                action("GRN Pending For Invoice")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50171;
                }
                action("Transport Detail Report")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50175;
                }
                action("Cancell Invoice / SRO Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50178;
                }
                action("Vendor Payment Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50183;
                }
                action("Customer Payment Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50184;
                }
                action("Ledger Dimension Wise")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50187;
                }
                action("Credit Memo Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50188;
                }
                action("Debit Memo Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50189;
                }
                action("Bank Letter & Statement")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50190;
                }
                action("Purchase Order Statement")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50196;
                }
                action("Sales Register Costing")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50202;
                }
                action("JV Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50203;
                }
                action("RM Valulation Report")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50216;
                }
                action("Vendor Invoice Application Register")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50221;
                }
                action("Bank Payment Report")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50230;
                }
                action("Exchange Difference Report")
                {
                    ApplicationArea = all;
                    Image = Report;
                    RunObject = Report 50242;
                }






            }

        }
    }

    var
        myInt: Integer;


}
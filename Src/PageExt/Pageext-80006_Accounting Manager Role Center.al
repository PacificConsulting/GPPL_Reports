pageextension 80006 "AccoManager Role CenterExtrpt" extends "Accounting Manager Role Center"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addafter("Cost Accounting")
        {
            group("Finance Reports")
            {
                action("In-Transit Item Valuation")
                {
                    ApplicationArea = All;
                    RunObject = report 50015;
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
                action("G/L Ledger")
                {
                    ApplicationArea = All;
                    RunObject = report 50029;
                }
                action("Daily Production Memo (Excel)")
                {
                    ApplicationArea = All;
                    RunObject = report 50030;
                }
                action("[Item By Location FG ]")
                {
                    ApplicationArea = All;
                    RunObject = report 50031;
                }
                action("Item By Location RM")
                {
                    ApplicationArea = All;
                    RunObject = report 50032;
                }
                action("Pop Item by Location")
                {
                    ApplicationArea = All;
                    RunObject = report 50035;
                }
                action("Bin-wise FG Material - Plant")
                {
                    ApplicationArea = All;
                    RunObject = report 50038;
                }
                action("RM Movement (Locwise)")
                {
                    ApplicationArea = All;
                    RunObject = report 50051;
                }
                action("FG Movement (Locwise)")
                {
                    ApplicationArea = All;
                    RunObject = report 50052;
                }
                action("Collection Report")
                {
                    ApplicationArea = All;
                    RunObject = report 50057;
                }
                action("Sales Register Costing")
                {
                    ApplicationArea = All;
                    RunObject = report 50069;
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
                action("BRT Register GST")
                {
                    ApplicationArea = All;
                    RunObject = report 50086;
                }
                action("Advance Payment Ledger")
                {
                    ApplicationArea = All;
                    RunObject = report 50088;
                }
                action("FG Valutn (After Revaluation)")
                {
                    ApplicationArea = All;
                    RunObject = report 50091;
                }

                action("FG Valuation Report")
                {
                    ApplicationArea = All;
                    RunObject = report 50091;
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
                    Image = Report;
                }
                action("Consumption Report Detailed")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50096;
                }

                action("Customer Aging Summary")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50101;
                }
                action("Excise Duty Differences")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50102;
                }
                action("Customer Aging Detail;")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50103;
                }
                action("Vendor Aging Summary")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50104;
                }
                action("Sales Summary")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50108;
                }
                action("Sales Credit Memo - Summary")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50111;
                }
                action("Sales Debit Memo - Summary")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50112;
                }
                action("Bank Reconciliation Statement")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50114;
                }
                action("Yeild Report")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50115;
                }
                action("GST Differences")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50122;
                }
                action("Staff Wise Expense")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50124;
                }
                action("Vendor Balance Confirmation")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50127;
                }
                action("Customer Balance Confirmation")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50128;
                }
                action("Vendor - Ledger Acconut")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50129;
                }
                action("POP Sales Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50144;
                }
                action("Consilated RG23D Report")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50147;
                }
                action("Purchase Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50151;
                }
                action("Overdue Reminder")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50153;
                }
                action("Trial Balance - GP")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50156;
                }
                action("FG Aging")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50160;
                }
                action("Raw Material Aging")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50162;
                }
                action("Item Sales Price History")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50163;
                }
                action("GRN Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50164;
                }
                action("CSO Analysis")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50167;
                }
                action("Statutory Form Remider")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50169;
                }
                action("Transport Detail Report")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50175;
                }
                action("Vendor Payment Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50183;
                }
                action("Customer Payment Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50184;
                }
                action("Ledger Dimension Wise")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50187;
                }
                action("Credit Memo Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50188;
                }
                action("Debit Memo Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50189;
                }
                action("Purchase Order Statement")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50196;
                }
                action("Sales Register Costing - New")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50202;
                }
                action("JV Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50203;
                }
                action("Foreign Currency Balance - New")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50210;
                }
                action("Inventory Valuation - New")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50212;
                }
                action("Production Order Comparision")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50215;
                }
                action("RM Valuatoin Report")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50216;
                }
                action("RM Valuation Report")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50216;
                }
                action("Fedai Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50218;
                }
                action("Customer Invoice Application Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50220;
                }
                action("Vendor Invoice Application Register")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50221;
                }
                action("Export G/L Entry")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50226;
                }
                action("Bank Payment Report")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50230;
                }
                action("RM Valuation Report - Testing")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50237;
                }
                action("Exchange Rate Diff Report")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50242;
                }
                action("Customer Summary Aging GP")
                {
                    ApplicationArea = All;
                    Image = Report;
                    RunObject = report 50243;
                }













            }
        }
    }

    var
        myInt: Integer;


}
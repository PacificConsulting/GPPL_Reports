pageextension 80021 PurchaseOrderExt extends "Purchase Order"
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
            action("Purchase Order - &Import")
            {
                ApplicationArea = all;
                trigger OnAction()
                BEGIN
                    CurrPage.SETSELECTIONFILTER(Rec);
                    REPORT.RUNMODAL(50054, TRUE, FALSE, Rec)
                END;
            }
        }
    }

    var
        myInt: Integer;
}
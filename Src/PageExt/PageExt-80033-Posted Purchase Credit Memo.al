pageextension 80033 ClosedBlanketPurchaseOrderExt extends "Posted Purchase Credit Memo"
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
            action("Purchase Return Invoice GST")
            {
                ApplicationArea = all;
                trigger OnAction()
                VAR
                    PH: Record 38;
                    Ord05: Record 224;
                    Ven05: Record 23;
                BEGIN
                    CurrPage.SETSELECTIONFILTER(Rec);
                    REPORT.RUNMODAL(50173, TRUE, FALSE, Rec)
                END;
            }
            action("Purchase &Return Invoice")
            {
                ApplicationArea = all;
                trigger OnAction()
                BEGIN
                    CurrPage.SETSELECTIONFILTER(Rec);
                    REPORT.RUNMODAL(50180, TRUE, FALSE, Rec)
                END;
            }
        }
    }

    var
        myInt: Integer;
}
pageextension 80034 PostedPurchaseCreditMemosExt extends "Posted Purchase Credit Memos"
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
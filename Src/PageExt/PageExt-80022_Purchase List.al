pageextension 80022 PurchaseListExt extends "Purchase List"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        // Add changes to page actions here
        addlast(reporting)
        {
            action("Purchase Order - &Import1")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    myInt: Integer;
                BEGIN
                    CurrPage.SETSELECTIONFILTER(Rec);
                    REPORT.RUNMODAL(50054, TRUE, FALSE, Rec)
                END;
            }
            action("Purchase Order - R&egular Detailed")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    RecPurchheader: Record "Purchase Header";
                BEGIN
                    RecPurchheader.RESET;
                    RecPurchheader := Rec;
                    RecPurchheader.SETRECFILTER;
                    REPORT.RUNMODAL(50087, TRUE, FALSE, RecPurchheader);
                END;
            }
        }
    }

    var
        myInt: Integer;
}
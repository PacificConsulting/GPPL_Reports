pageextension 80043 PostedReturnShipmentExt extends "Posted Return Shipment"
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

            action("&Print2")
            {
                Ellipsis = true;
                Promoted = true;
                Image = Print;
                PromotedCategory = Process;
                ApplicationArea = all;
                trigger OnAction()
                var
                    ReturnShptHeader: Record "Return Shipment Header";
                BEGIN
                    //CurrPage.SETSELECTIONFILTER(ReturnShptHeader);
                    //ReturnShptHeader.PrintRecords(TRUE);


                    ReturnShptHeader.RESET;
                    ReturnShptHeader := Rec;
                    ReturnShptHeader.SETRECFILTER;
                    REPORT.RUN(50179, TRUE, FALSE, ReturnShptHeader);
                END;
            }
        }
    }

    var
        myInt: Integer;
}
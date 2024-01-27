pageextension 80027 PostedPurchaseInvoicesExt extends "Posted Purchase Invoices"
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
            action("Self Invoice GST")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()

                VAR
                    PIH: Record 122;
                BEGIN
                    //28July2017
                    PIH.RESET;
                    PIH.SETRANGE("No.", rec."No.");
                    IF PIH.FINDFIRST THEN
                        REPORT.RUNMODAL(50084, TRUE, TRUE, PIH);
                    //28July2017
                END;
            }
            // action("&Print")
            // {
            //     ApplicationArea = all;
            //     Image = Print;
            //     trigger OnAction()

            //     VAR
            //         PIH: Record 122;
            //     BEGIN
            //         //28July2017
            //         PIH.RESET;
            //         PIH.SETRANGE("No.", rec."No.");
            //         IF PIH.FINDFIRST THEN
            //             REPORT.RUNMODAL(50084, TRUE, TRUE, PIH);
            //         //28July2017
            //     END;
            // }
        }
    }

    var
        myInt: Integer;
}
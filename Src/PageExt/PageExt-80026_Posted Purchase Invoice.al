pageextension 80026 PostedPurchaseInvoiceExt extends "Posted Purchase Invoice"
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
            action("&Print")
            {
                ApplicationArea = all;
                Image = Print;
                Promoted = true;
                PromotedCategory = Process;
                trigger OnAction()
                VAR
                    GLEntry: Record "G/L Entry";
                BEGIN
                    //Commeted by EBT STIVAN ---(20/04/2012)--- To allow Print of Posted Purchase Invoice ---START
                    //              {
                    //              CurrPage.SETSELECTIONFILTER(PurchInvHeader);
                    // PurchInvHeader.PrintRecords(TRUE);
                    //              }
                    //Commented by EBT STIVAN ---(20/04/2012)--- To allow Print of Posted Purchase Invoice -----END

                    //EBT STIVAN ---(20/04/2012)--- To allow Print of Posted Purchase Invoice ---START
                    GLEntry.SETCURRENTKEY("Document No.", "Posting Date");
                    GLEntry.SETRANGE("Document No.", rec."No.");
                    GLEntry.SETRANGE("Posting Date", rec."Posting Date");
                    IF GLEntry.FINDFIRST THEN
                        //REPORT.RUNMODAL(REPORT::"Posted Voucher",TRUE,TRUE,GLEntry);
                        REPORT.RUNMODAL(50207, TRUE, TRUE, GLEntry);//08Jul2019
                                                                    //EBT STIVAN ---(20/04/2012)--- To allow Print of Posted Purchase Invoice -----END
                END;
            }
        }
    }

    var
        myInt: Integer;
}
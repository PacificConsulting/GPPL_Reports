pageextension 80039 EWAYBILLPAGE extends "E-WAY BILL PAGE"
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
            action("E-WAY BILL REPORT TRANSFER C2")
            {
                ApplicationArea = all;
                trigger OnAction()
                VAR
                    ReportEwayBillTransferC: Report 50236;
                    RecTSH: Record 5744;
                BEGIN
                    RecTSH.RESET;
                    RecTSH.SETRANGE("No.", rec."Document No.");
                    IF RecTSH.FINDFIRST THEN BEGIN
                        CLEAR(ReportEwayBillTransferC);
                        ReportEwayBillTransferC.SetParams(rec."EWB No.", rec.Cancelled);
                        ReportEwayBillTransferC.SETTABLEVIEW(RecTSH);
                        ReportEwayBillTransferC.USEREQUESTPAGE(TRUE);
                        ReportEwayBillTransferC.RUNMODAL;
                    END ELSE
                        ERROR('Invalid Document to run the report');
                END;
            }

        }
    }

    var
        myInt: Integer;
}
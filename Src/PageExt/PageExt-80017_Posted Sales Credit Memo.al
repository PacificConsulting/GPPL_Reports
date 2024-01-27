pageextension 80017 PostedSalesCreditMemoExt extends "Posted Sales Credit Memo"
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
            action("Print SRO - &Industrial")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    SalesCreditMemoHeader: Record "Sales Cr.Memo Header";
                BEGIN
                    SalesCreditMemoHeader.RESET;
                    SalesCreditMemoHeader := Rec;
                    SalesCreditMemoHeader.SETRECFILTER;
                    REPORT.RUN(50033, TRUE, FALSE, SalesCreditMemoHeader);
                END;

            }
            action("Print SRO - &Automotive")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    SalesCreditMemoHeader: Record "Sales Cr.Memo Header";
                BEGIN
                    SalesCreditMemoHeader.RESET;
                    SalesCreditMemoHeader := Rec;
                    SalesCreditMemoHeader.SETRECFILTER;
                    REPORT.RUN(50034, TRUE, FALSE, SalesCreditMemoHeader);
                END;

            }
            action("Print Sturcture Discount")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    SalesCreditMemoHeader: Record "Sales Cr.Memo Header";
                BEGIN
                    SalesCreditMemoHeader.RESET;
                    SalesCreditMemoHeader := Rec;
                    SalesCreditMemoHeader.SETRECFILTER;
                    REPORT.RUN(50035, TRUE, FALSE, SalesCreditMemoHeader);
                END;
            }
            action("Print SRO - &Export")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    SalesCreditMemoHeader: Record "Sales Cr.Memo Header";
                BEGIN
                    SalesCreditMemoHeader.RESET;
                    SalesCreditMemoHeader := Rec;
                    SalesCreditMemoHeader.SETRECFILTER;
                    REPORT.RUN(50036, TRUE, FALSE, SalesCreditMemoHeader);
                END;
            }
            action("Print &Credit Memo")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    SalesCreditMemoHeader: Record "Sales Cr.Memo Header";
                BEGIN
                    SalesCreditMemoHeader.RESET;
                    SalesCreditMemoHeader := Rec;
                    SalesCreditMemoHeader.SETRECFILTER;
                    REPORT.RUN(50037, TRUE, FALSE, SalesCreditMemoHeader);
                END;
            }
            action("Credit Note GST")
            {
                ApplicationArea = all;
                trigger OnAction()
                VAR
                    RecSalesCrMemoLine: Record 115;
                    SalesCreditMemoHeader: Record "Sales Cr.Memo Header";
                BEGIN
                    //RSPLSUM 13Jan21


                    IF REC."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(IRN);
                        RecSalesCrMemoLine.RESET;
                        RecSalesCrMemoLine.SETRANGE("Document No.", REC."No.");
                        RecSalesCrMemoLine.SETFILTER(Quantity, '%1', 0);
                        //  RecSalesCrMemoLine.SETFILTER("Total GST Amount", '%1', 0); tempory comment 1-jan-23 pcpl-065
                        IF RecSalesCrMemoLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    SalesCreditMemoHeader.RESET;
                    SalesCreditMemoHeader := Rec;
                    SalesCreditMemoHeader.SETRECFILTER;
                    REPORT.RUN(50074, TRUE, FALSE, SalesCreditMemoHeader);
                END;
            }
        }
    }

    var
        myInt: Integer;
}
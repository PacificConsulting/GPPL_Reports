pageextension 80018 PostedSalesCreditMemosExt extends "Posted Sales Credit Memos"
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
                var
                    SalesCreditMemoHeader: Record "Sales Cr.Memo Header";
                BEGIN
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
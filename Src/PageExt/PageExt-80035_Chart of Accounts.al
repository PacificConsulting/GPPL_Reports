pageextension 80035 ChartofAccountsExtREP extends "Chart of Accounts"
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
            action("&Print")
            {
                ApplicationArea = all;
                trigger OnAction()
                VAR

                BEGIN
                    CurrPage.SETSELECTIONFILTER(Rec);
                    REPORT.RUNMODAL(50173, TRUE, FALSE, Rec)
                END;
            }
            action("Trail Balance GST2")
            {
                RunObject = Report 50182;
                Image = Report;
                ApplicationArea = all;

            }
        }
    }

    var
        myInt: Integer;
}
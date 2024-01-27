pageextension 80004 SalesOrderCardREpCstm extends "Sales Order"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addlast(Reporting)
        {
            action("Sales Order Industrial - GST")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    SH: Record "Sales Header";
                begin
                    SH.RESET;
                    SH := Rec;
                    SH.SETRECFILTER;
                    REPORT.RUN(50013, TRUE, FALSE, SH);
                end;
            }
            action("Sales Order Automotive - GST")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    SH: Record "Sales Header";
                begin
                    SH.RESET;
                    SH := Rec;
                    SH.SETRECFILTER;
                    REPORT.RUN(50014, TRUE, FALSE, SH);
                end;
            }
            action("Sales Order - & Structure Details")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    SH: Record "Sales Header";
                BEGIN
                    SH.RESET;
                    SH := Rec;
                    SH.SETRECFILTER;
                    REPORT.RUN(50140, TRUE, FALSE, SH);
                END;
            }
            action("Sales Order Export/RPO - GST")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    SH: Record "Sales Header";
                BEGIN
                    SH.RESET;
                    SH := Rec;
                    SH.SETRECFILTER;
                    REPORT.RUN(50155, TRUE, FALSE, SH);
                END;
            }
            action("Sales Confirmation")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    SH: Record "Sales Header";
                BEGIN
                    SH.RESET;
                    SH := Rec;
                    SH.SETRECFILTER;
                    REPORT.RUN(50231, TRUE, FALSE, SH);
                END;
            }
        }
    }

    var
        myInt: Integer;
}
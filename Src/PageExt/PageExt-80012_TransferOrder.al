pageextension 80012 TransferOrderExtRptCstm extends "Transfer Order"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addlast(reporting)
        {
            action("Structure Details")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    TrasnHdr: Record "Transfer Header";
                begin
                    TrasnHdr.RESET;
                    TrasnHdr := Rec;
                    TrasnHdr.SETRECFILTER;
                    REPORT.RUN(50018, TRUE, FALSE, TrasnHdr);
                end;
            }
            action("GST Calculation")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    TrasnHdr: Record "Transfer Header";
                begin
                    TrasnHdr.RESET;
                    TrasnHdr := Rec;
                    TrasnHdr.SETRECFILTER;
                    REPORT.RUN(50018, TRUE, FALSE, TrasnHdr);
                end;
            }
        }
    }

    var
        myInt: Integer;
}
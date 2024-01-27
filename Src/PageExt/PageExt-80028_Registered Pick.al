pageextension 80028 RegisteredPickExt extends "Registered Pick"
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
            action(Print)
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                var
                    regiwarehouseactihe: Record "Registered Whse. Activity Hdr.";
                BEGIN
                    regiwarehouseactihe.RESET;
                    regiwarehouseactihe := Rec;
                    regiwarehouseactihe.SETRECFILTER;
                    REPORT.RUN(50086, TRUE, FALSE, regiwarehouseactihe);
                END;
            }
        }
    }

    var
        myInt: Integer;

}
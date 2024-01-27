pageextension 80029 AdministratorRoleCenterExt extends "Administrator Role Center"
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
            action("Sales Register GST")
            {
                ApplicationArea = all;
                Image = Report;
                RunObject = report 50092;

            }

        }
    }

    var
        myInt: Integer;
}
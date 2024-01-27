pageextension 80030 ReleasedProductionOrdersExt extends "Released Production Orders"
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
            action("Job Card")
            {
                ApplicationArea = all;
                Image = Report;
                RunObject = Report 50100;
            }
        }
    }

    var
        myInt: Integer;
}
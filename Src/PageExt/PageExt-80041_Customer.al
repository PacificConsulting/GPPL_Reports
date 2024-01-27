pageextension 80041 CustomerListExt extends "Customer List"
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
            action("Send Balance Confirmation")
            {
                ApplicationArea = all;
                RunObject = Report 50254;
                Promoted = true;
                PromotedIsBig = true;
                Image = MailAttachment;
                PromotedCategory = Process;
            }
        }
    }

    var
        myInt: Integer;
}
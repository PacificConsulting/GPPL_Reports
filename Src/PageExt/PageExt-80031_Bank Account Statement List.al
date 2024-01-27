pageextension 80031 BankAccountStatementListExt extends "Bank Account Statement List"
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

            action("Bank Reco Statement")
            {
                ApplicationArea = all;
                RunObject = Report 50114;
                Promoted = true;
                PromotedIsBig = true;
                Image = Report;
                PromotedCategory = Report;

            }
        }

    }

    var
        myInt: Integer;
}
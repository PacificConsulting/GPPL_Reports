pageextension 80036 VendorLedgerEntriesExt extends "Vendor Ledger Entries"
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
            action("Print Voucher GP2")
            {
                ApplicationArea = all;
                Image = PrintVoucher;
                Ellipsis = true;

                trigger OnAction()
                VAR
                    GLEntry: Record 17;
                BEGIN
                    GLEntry.SETCURRENTKEY("Document No.", "Posting Date");
                    GLEntry.SETRANGE("Document No.", rec."Document No.");
                    GLEntry.SETRANGE("Posting Date", rec."Posting Date");
                    IF GLEntry.FIND('-') THEN
                        REPORT.RUNMODAL(50207, TRUE, TRUE, GLEntry);
                END;
            }
        }
    }

    var
        myInt: Integer;
}
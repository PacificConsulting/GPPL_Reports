pageextension 60002 Vendor_List_Ext extends "Vendor List"
{
    layout
    {
        addafter(Name)
        {
            field("Creation Date"; rec."Creation Date")
            {
                ApplicationArea = all;
            }
            field("E-Mail"; Rec."E-Mail")
            {
                ApplicationArea = all;
            }
            field("Shipping Agent"; rec."Shipping Agent")
            {
                ApplicationArea = all;
            }
            field("P.A.N. No."; Rec."P.A.N. No.")
            {
                ApplicationArea = all;
            }
            field("State Code"; Rec."State Code")
            {
                ApplicationArea = ALL;
            }
            // field("Service Tax Registration No."; rec."Service Tax Registration No.")
            // {
            //     ApplicationArea = all;
            // }
            // field("Service Entity Type"; rec."Service Entity Type")
            // {
            //     ApplicationArea = all;
            // }
            field("GST Vendor Type"; rec."GST Vendor Type")
            {
                ApplicationArea = all;
            }
            field(Address; rec.Address)
            {
                ApplicationArea = all;
            }
            field("MSME Status"; rec."MSME Status")
            {
                ApplicationArea = all;
            }
            /* field("Balance (LCY)"; Rec."Balance (LCY)")
            {
                ApplicationArea = all;
            } */
            field("Net Change"; Rec."Net Change")
            {
                ApplicationArea = all;
            }
            field("Address 2"; Rec."Address 2")
            {
                ApplicationArea = all;
            }
            field(City; Rec.City)
            {
                ApplicationArea = all;
            }

        }
        addafter("Location Code")
        {
            /*  field("Responsibility Center"; Rec."Responsibility Center")
             {
                 ApplicationArea = all;
             } */
            // field("T.I.N. No."; rec."T.I.N. No.")
            // {
            //     ApplicationArea = all;
            // }
            field(Balance; Rec.Balance)
            {
                ApplicationArea = all;
            }
            field("Debit Amount (LCY)"; Rec."Debit Amount (LCY)")
            {
                ApplicationArea = all;
            }
            field("Credit Amount (LCY)"; Rec."Credit Amount (LCY)")
            {
                ApplicationArea = all;
            }
        }
        addafter("Post Code")
        {
            field("Debit Amount"; Rec."Debit Amount")
            {
                ApplicationArea = all;
            }
            field("Credit Amount"; Rec."Credit Amount")
            {
                ApplicationArea = all;
            }
            field("Shipping Agent Code"; Rec."Shipping Agent Code")
            {
                ApplicationArea = all;
            }
        }
        addafter("Base Calendar Code")
        {
            field("GST Registration No."; Rec."GST Registration No.")
            {
                ApplicationArea = all;
            }
            field("IRN Applicable"; rec."IRN Applicable")
            {
                ApplicationArea = all;
            }
            field("Linking of Aadhaar with PAN"; rec."Linking of Aadhaar with PAN")
            {
                ApplicationArea = all;
            }
            field("Exclude From Bal Confir Mail"; rec."Exclude From Bal Confir Mail")
            {
                ApplicationArea = all;
            }
            field("ITR filled for last 02 years"; rec."ITR filled for last 02 years")
            {
                ApplicationArea = all;
            }
        }
        modify("Post Code")
        {
            Enabled = false;
        }
        // Add changes to page layout here
    }

    actions
    {
        addafter("Vendor - Trial Balance")
        {
            action("Send Balance Confirmation")
            {
                RunObject = Report 50253;
                Promoted = true;
                PromotedIsBig = true;
                Image = MailAttachment;
                PromotedCategory = Process;
                ApplicationArea = all;

            }

        }

        // Add changes to page actions here
    }
    trigger OnOpenPage()
    begin
        //SETCURRENTKEY(Name);             //EBT STIVAN (09062012) --- To sort as per Name wise
        rec.SETFILTER(Blocked, '<>%1', 2);     //EBT STIVAN (09062012) --- To Filter Blocked as <>ALL
    end;

    var
        myInt: Integer;
}
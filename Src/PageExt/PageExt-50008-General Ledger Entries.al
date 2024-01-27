pageextension 50008 GeneralLedgerEntriesExt extends "General Ledger Entries"
{
    layout
    {
        // Add changes to page layout here
        addafter(Control1)
        {
            field("Prior-Year Entry"; rec."Prior-Year Entry")
            {
                ApplicationArea = all;
            }
            field("Close Income Statement Dim. ID"; rec."Close Income Statement Dim. ID")
            {
                ApplicationArea = all;
            }
            field("System-Created Entry"; rec."System-Created Entry")
            {
                ApplicationArea = all;
            }
            // field("External Document No."; rec."External Document No.")
            // {
            //     ApplicationArea = all;
            // }
            field("Exp/Purchase Invoice Doc. No."; rec."Exp/Purchase Invoice Doc. No.")
            {
                ApplicationArea = all;
            }
            // field("Source Type"; rec."Source Type")
            // {
            //     ApplicationArea = all;
            // }
            field("Transaction No."; rec."Transaction No.")
            {
                ApplicationArea = all;
            }
            // field("Source No."; rec."Source No.")
            // {
            //     ApplicationArea = all;
            // }
            // field("Location Code"; rec."Location Code")
            // {
            //     ApplicationArea = all;
            // }
            // field("Dimension Set ID"; rec."Dimension Set ID")
            // {
            //     ApplicationArea = all;
            // }
            // field("Debit Amount"; rec."Debit Amount")
            // {
            //     ApplicationArea = all;
            // }
            // field("Credit Amount"; rec."Credit Amount")
            // {
            //     ApplicationArea = all;
            // }
            // field(Amount; rec.Amount)
            // {
            //     ApplicationArea = all;
            //     DecimalPlaces = 5 : 5;
            // }
            // field("Bal. Account Type"; rec."Bal. Account Type")
            // {
            //     ApplicationArea = all;
            // }
            // field("Bal. Account No."; rec."Bal. Account No.")
            // {
            //     ApplicationArea = all;
            // }
            // field("Job No."; rec."Job No.")
            // {
            //     ApplicationArea = all;
            //     Visible = FALSE;
            // }
            // field("IC Partner Code"; REC."IC Partner Code")
            // {
            //     ApplicationArea = ALL;
            //     Visible = FALSE;
            // }
            // field("Gen. Posting Type"; REC."Gen. Posting Type")
            // {
            //     ApplicationArea = ALL;
            //     Visible = false;
            // }
            // field("Gen. Bus. Posting Group"; REC."Gen. Bus. Posting Group")
            // {
            //     ApplicationArea = ALL;
            //     Visible = false;
            // }
            // field("Gen. Prod. Posting Group"; REC."Gen. Prod. Posting Group")
            // {
            //     ApplicationArea = ALL;
            //     Visible = false;

            // }
            // field(Quantity; REC.Quantity)
            // {
            //     ApplicationArea = ALL;
            //     Visible = false;
            // }
            // field("Additional-Currency Amount"; REC."Additional-Currency Amount")
            // {
            //     ApplicationArea = ALL;
            //     Visible = false;
            // }

            // field("VAT Amount"; REC."VAT Amount")
            // {
            //     ApplicationArea = ALL;
            //     Visible = false;
            // }
            field(Narration; Narration)
            {
                ApplicationArea = ALL;
                //CaptionML ='Narration';
            }
            field(GLAccIncomeBalance; GLAccIncomeBalance)
            {
                ApplicationArea = ALL;
                //CaptionML ="G/L Type";
            }
            field("Creation Date"; REC."Creation Date")
            {
                ApplicationArea = ALL;
            }
            // field("Blanket Oreder No."; REC."Blanket Oreder No.")
            // {
            //     ApplicationArea = ALL;
            // }
        }
    }


    actions
    {
        // Add changes to page actions here
        // modify("Print Voucher")
        // {
        //     //Visible = false;
        // }
        addafter("&Navigate")
        {
            action("Export G/L Entry to Excel")
            {
                ApplicationArea = all;
                RunObject = Report 50226;
                Image = Excel;
            }

            action("Print Voucher2")
            {
                ApplicationArea = all;
                Image = PrintVoucher;
                Caption = 'Print Voucher';
                Visible = false;
                trigger OnAction()
                VAR
                    GLEntry: Record 17;
                BEGIN
                    GLEntry.SETCURRENTKEY("Document No.", "Posting Date");
                    GLEntry.SETRANGE("Document No.", REC."Document No.");
                    GLEntry.SETRANGE("Posting Date", REC."Posting Date");
                    IF GLEntry.FINDFIRST THEN begin

                    end;
                    //REPORT.RUNMODAL(REPORT::"Posted Voucher",TRUE,TRUE,GLEntry);
                    //08Jul2019
                    // IF GLEntry."Source Code" = 'JOURNALV' THEN
                    // REPORT.RUNMODAL(REPORT::"Posted Voucher" TRUE, TRUE, GLEntry)
                    //ELSE
                    // REPORT.RUNMODAL(50207, TRUE, TRUE, GLEntry);
                    //
                END;
            }


        }
    }
    trigger OnAfterGetRecord()
    var
        myInt: Integer;
    BEGIN
        //RSPLSUM 06Dec19>>
        CLEAR(Narration);
        RecPostedNarration.RESET;
        RecPostedNarration.SETCURRENTKEY("Entry No.", "Transaction No.");
        RecPostedNarration.SETRANGE("Entry No.", 0);
        RecPostedNarration.SETRANGE("Transaction No.", rec."Transaction No.");
        IF RecPostedNarration.FINDSET THEN
            REPEAT
                Narration += RecPostedNarration.Narration;
            UNTIL RecPostedNarration.NEXT = 0;

        //RSPLSUM 09Dec19>>
        RecPurchComtLine.RESET;
        RecPurchComtLine.SETCURRENTKEY("Document Type", "No.");
        RecPurchComtLine.SETFILTER("Document Type", '%1|%2', RecPurchComtLine."Document Type"::"Posted Invoice",
                                    RecPurchComtLine."Document Type"::"Posted Credit Memo");
        RecPurchComtLine.SETRANGE("No.", rec."Document No.");
        IF RecPurchComtLine.FINDSET THEN
            REPEAT
                Narration += RecPurchComtLine.Comment;
            UNTIL RecPurchComtLine.NEXT = 0;
        //RSPLSUM 09Dec19<<

        //RSPLSUM 13Dec19>>
        RecSalesComtLine.RESET;
        RecSalesComtLine.SETCURRENTKEY("Document Type", "No.");
        RecSalesComtLine.SETFILTER("Document Type", '%1|%2', RecSalesComtLine."Document Type"::"Posted Credit Memo",
                                  RecSalesComtLine."Document Type"::"Posted Invoice");
        RecSalesComtLine.SETRANGE("No.", rec."Document No.");
        IF RecSalesComtLine.FINDSET THEN
            REPEAT
                Narration += RecSalesComtLine.Comment;
            UNTIL RecSalesComtLine.NEXT = 0;

        //RSPLSUM 09Dec19>>
        RecInventoryComtline.RESET;
        RecInventoryComtline.SETCURRENTKEY("Document Type", "No.");
        RecInventoryComtline.SETRANGE("Document Type", RecInventoryComtline."Document Type"::"Posted Transfer Receipt");
        RecInventoryComtline.SETRANGE("No.", rec."Document No.");
        IF RecInventoryComtline.FINDSET THEN
            REPEAT
                Narration += RecInventoryComtline.Comment;
            UNTIL RecInventoryComtline.NEXT = 0;
        //RSPLSUM 09Dec19<<

        //RSPLSUM 21Jan20>>
        IF Narration = '' THEN BEGIN
            IF (rec.Reversed = TRUE) AND (rec."Source Code" = 'REVERSAL') THEN BEGIN
                RecGLE.RESET;
                IF RecGLE.GET(rec."Reversed Entry No.") THEN BEGIN
                    RecPostedNarration.RESET;
                    RecPostedNarration.SETCURRENTKEY("Entry No.", "Transaction No.");
                    RecPostedNarration.SETRANGE("Entry No.", 0);
                    RecPostedNarration.SETRANGE("Transaction No.", RecGLE."Transaction No.");
                    IF RecPostedNarration.FINDSET THEN
                        REPEAT
                            Narration += RecPostedNarration.Narration;
                        UNTIL RecPostedNarration.NEXT = 0;
                END;
            END;
        END;
        //RSPLSUM 21Jan20<<

        //Fahim 29-09-2021
        CLEAR(BLPUORNo);
        IF rec."Source Type" = rec."Source Type"::Vendor THEN BEGIN
            PuInvHeader.RESET;
            IF PuInvHeader.GET(rec."Document No.") THEN
                BLPUORNo := PuInvHeader."Blanket Order No.";
        END;

        //RSPLSUM 13Dec19<<

        CLEAR(GLAccIncomeBalance);
        RecGLAcc.RESET;
        IF RecGLAcc.GET(rec."G/L Account No.") THEN
            GLAccIncomeBalance := FORMAT(RecGLAcc."Income/Balance");
        //RSPLSUM 06Dec19<<
    END;

    var
        RecPostedNarration: Record "Posted Narration";
        Narration: Text;
        GLAccIncomeBalance: Text;
        RecGLAcc: Record 15;
        RecPurchComtLine: Record 43;
        RecSalesComtLine: Record 44;
        RecInventoryComtline: Record 5748;
        RecGLE: Record 17;
        BLPUORNo: Code[50];
        PuInvHeader: Record 122;
}
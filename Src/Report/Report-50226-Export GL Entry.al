report 50226 "Export G/L Entry"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ExportGLEntry.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            RequestFilterFields = "Posting Date", "G/L Account No.";
            column(PostingDate_GLEntry; "G/L Entry"."Posting Date")
            {
            }
            column(GLAccountNo_GLEntry; "G/L Entry"."G/L Account No.")
            {
            }
            column(GLAccountName_GLEntry; "G/L Entry"."G/L Account Name")
            {
            }
            column(Description_GLEntry; "G/L Entry".Description)
            {
            }
            column(DocumentNo_GLEntry; "G/L Entry"."Document No.")
            {
            }
            column(ExternalDocumentNo_GLEntry; "G/L Entry"."External Document No.")
            {
            }
            column(DebitAmount_GLEntry; "G/L Entry"."Debit Amount")
            {
            }
            column(CreditAmount_GLEntry; "G/L Entry"."Credit Amount")
            {
            }
            column(Amount_GLEntry; "G/L Entry".Amount)
            {
            }
            column(SourceCode_GLEntry; "G/L Entry"."Source Code")
            {
            }
            column(GlobalDimension1Code_GLEntry; "G/L Entry"."Global Dimension 1 Code")
            {
            }
            column(GlobalDimension2Code_GLEntry; "G/L Entry"."Global Dimension 2 Code")
            {
            }
            column(SourceType_GLEntry; "G/L Entry"."Source Type")
            {
            }
            column(SourceNo_GLEntry; "G/L Entry"."Source No.")
            {
            }
            column(LocationCode_GLEntry; '"G/L Entry"."Location Code"')
            {
            }
            column(BalAccountType_GLEntry; "G/L Entry"."Bal. Account Type")
            {
            }
            column(BalAccountNo_GLEntry; "G/L Entry"."Bal. Account No.")
            {
            }
            column(UserID_GLEntry; "G/L Entry"."User ID")
            {
            }
            column(ExpPurchaseInvoiceDocNo_GLEntry; "G/L Entry"."Exp/Purchase Invoice Doc. No.")
            {
            }
            column(Income_Balance; GLAccIncomeBalance)
            {
            }
            column(Narration; Narration)
            {
            }
            column(SourceName; SourceName)
            {
            }
            column(BLPUORNo; BLPUORNo)
            {
            }

            trigger OnAfterGetRecord()
            begin
                CLEAR(GLAccIncomeBalance);
                RecGLAcc.RESET;
                IF RecGLAcc.GET("G/L Account No.") THEN
                    GLAccIncomeBalance := FORMAT(RecGLAcc."Income/Balance");

                //RSPLSUM 06Dec19>>
                CLEAR(Narration);
                RecPostedNarration.RESET;
                RecPostedNarration.SETCURRENTKEY("Entry No.", "Transaction No.");
                RecPostedNarration.SETRANGE("Entry No.", 0);
                RecPostedNarration.SETRANGE("Transaction No.", "Transaction No.");
                IF RecPostedNarration.FINDSET THEN
                    REPEAT
                        Narration += RecPostedNarration.Narration;
                    UNTIL RecPostedNarration.NEXT = 0;



                //RSPLSUM 09Dec19>>
                IF Narration = '' THEN BEGIN
                    RecPurchComtLine.RESET;
                    RecPurchComtLine.SETCURRENTKEY("Document Type", "No.");
                    RecPurchComtLine.SETFILTER("Document Type", '%1|%2', RecPurchComtLine."Document Type"::"Posted Invoice",
                                                RecPurchComtLine."Document Type"::"Posted Credit Memo");
                    RecPurchComtLine.SETRANGE("No.", "Document No.");
                    IF RecPurchComtLine.FINDSET THEN
                        REPEAT
                            Narration += RecPurchComtLine.Comment;
                        UNTIL RecPurchComtLine.NEXT = 0;
                END;
                //RSPLSUM 09Dec19<<

                //RSPLSUM 13Dec19>>
                IF Narration = '' THEN BEGIN
                    RecSalesComtLine.RESET;
                    RecSalesComtLine.SETCURRENTKEY("Document Type", "No.");
                    RecSalesComtLine.SETFILTER("Document Type", '%1|%2', RecSalesComtLine."Document Type"::"Posted Credit Memo",
                                              RecSalesComtLine."Document Type"::"Posted Invoice");
                    RecSalesComtLine.SETRANGE("No.", "Document No.");
                    IF RecSalesComtLine.FINDSET THEN
                        REPEAT
                            Narration += RecSalesComtLine.Comment;
                        UNTIL RecSalesComtLine.NEXT = 0;
                END;
                //RSPLSUM 09Dec19>>
                IF Narration = '' THEN BEGIN
                    RecInventoryComtline.RESET;
                    RecInventoryComtline.SETCURRENTKEY("Document Type", "No.");
                    RecInventoryComtline.SETRANGE("Document Type", RecInventoryComtline."Document Type"::"Posted Transfer Receipt");
                    RecInventoryComtline.SETRANGE("No.", "Document No.");
                    IF RecInventoryComtline.FINDSET THEN
                        REPEAT
                            Narration += RecInventoryComtline.Comment;
                        UNTIL RecInventoryComtline.NEXT = 0;
                END;
                //RSPLSUM 09Dec19<<

                //RSPLSUM 21Jan20>>
                IF Narration = '' THEN BEGIN
                    IF (Reversed = TRUE) AND ("Source Code" = 'REVERSAL') THEN BEGIN
                        RecGLE.RESET;
                        IF RecGLE.GET("Reversed Entry No.") THEN BEGIN
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

                CLEAR(SourceName);
                IF "G/L Entry"."Source Type" = "G/L Entry"."Source Type"::"Bank Account" THEN BEGIN
                    RecBankAcc.RESET;
                    IF RecBankAcc.GET("G/L Entry"."Source No.") THEN
                        SourceName := RecBankAcc.Name;
                END;

                IF "G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Customer THEN BEGIN
                    RecCust.RESET;
                    IF RecCust.GET("G/L Entry"."Source No.") THEN
                        SourceName := RecCust.Name;
                END;

                IF "G/L Entry"."Source Type" = "G/L Entry"."Source Type"::"Fixed Asset" THEN BEGIN
                    RecFA.RESET;
                    IF RecFA.GET("G/L Entry"."Source No.") THEN
                        SourceName := RecFA.Description;
                END;

                IF "G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Vendor THEN BEGIN
                    RecVend.RESET;
                    IF RecVend.GET("G/L Entry"."Source No.") THEN
                        SourceName := RecVend.Name;
                END;

                //Fahim 29-09-2021
                CLEAR(BLPUORNo);
                IF "G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Vendor THEN BEGIN
                    PuInvHeader.RESET;
                    IF PuInvHeader.GET("G/L Entry"."Document No.") THEN
                        BLPUORNo := PuInvHeader."Blanket Order No.";
                END;
            end;
        }
    }

    requestpage
    {

        layout
        {
        }

        actions
        {
        }
    }

    labels
    {
    }

    var
        GLAccIncomeBalance: Text;
        RecGLAcc: Record 15;
        Narration: Text;
        RecPostedNarration: Record "Posted Narration";
        RecPurchComtLine: Record 43;
        RecSalesComtLine: Record 44;
        RecInventoryComtline: Record 5748;
        RecGLE: Record 17;
        RecBankAcc: Record 270;
        RecVend: Record 23;
        RecCust: Record 18;
        RecFA: Record 5600;
        SourceName: Text;
        PuInvHeader: Record 122;
        BLPUORNo: Text[20];
}


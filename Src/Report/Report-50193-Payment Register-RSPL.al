report 50193 "Payment Register-RSPL"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PaymentRegisterRSPL.rdl';
    Caption = 'Payment Register-RSPL';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Detailed GST Ledger Entry"; "Detailed GST Ledger Entry")
        {
            DataItemTableView = SORTING("Entry No.")
                                WHERE("Document Type" = CONST(Payment),
                                      "GST on Advance Payment" = CONST(true));
            RequestFilterFields = "Posting Date";
            column(EntryNo; "Detailed GST Ledger Entry"."Entry No.")
            {
            }
            column(EntryType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Entry Type")
            {
            }
            column(TransactionType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Transaction Type")
            {
            }
            column(DocumentType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Document Type")
            {
            }
            column(DocumentNo; "Detailed GST Ledger Entry"."Document No.")
            {
            }
            column(PostingDate_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Posting Date")
            {
            }
            column(Type_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry".Type)
            {
            }
            column(ItmNo; "Detailed GST Ledger Entry"."No.")
            {
            }
            column(ProductType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Product Type")
            {
            }
            column(SourceType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Source Type")
            {
            }
            column(SourceNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Source No.")
            {
            }
            column(HSNSACCode_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."HSN/SAC Code")
            {
            }
            column(GSTComponentCode_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Component Code")
            {
            }
            column(GSTGroupCode_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Group Code")
            {
            }
            column(GSTJurisdictionType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Jurisdiction Type")
            {
            }
            column(GSTBaseAmount_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Base Amount")
            {
            }
            column(GST_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST %")
            {
            }
            column(GSTAmount_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Amount")
            {
            }
            column(ExternalDocumentNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."External Document No.")
            {
            }
            column(AmountLoadedonItem_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Amount Loaded on Item")
            {
            }
            column(Quantity_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry".Quantity)
            {
            }
            column(GSTWithoutPaymentofDuty_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Without Payment of Duty")
            {
            }
            column(GLAccountNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."G/L Account No.")
            {
            }
            column(ReversedbyEntryNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Reversed by Entry No.")
            {
            }
            column(Reversed_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry".Reversed)
            {
            }
            column(UserID_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."User ID"')
            {
            }
            column(Positive_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry".Positive')
            {
            }
            column(DocumentLineNo; "Detailed GST Ledger Entry"."Document Line No.")
            {
            }
            column(ItemChargeEntry_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Item Charge Entry")
            {
            }
            column(ReverseCharge_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Reverse Charge")
            {
            }
            column(GSTonAdvancePayment_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST on Advance Payment")
            {
            }
            column(NatureofSupply_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Nature of Supply"')
            {
            }
            column(PaymentDocumentNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Payment Document No.")
            {
            }
            column(GSTExemptedGoods_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Exempted Goods")
            {
            }
            column(LocationStateCode_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Location State Code"')
            {
            }
            column(BuyerSellerStateCode_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Buyer/Seller State Code"')
            {
            }
            column(ShippingAddressStateCode_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Shipping Address State Code"')
            {
            }
            column(LocationRegNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Location  Reg. No.")
            {
            }
            column(BuyerSellerRegNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Buyer/Seller Reg. No.")
            {
            }
            column(GSTGroupType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Group Type")
            {
            }
            column(GSTCredit_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Credit")
            {
            }
            column(ReversalEntry_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Reversal Entry")
            {
            }
            column(TransactionNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Transaction No.")
            {
            }
            column(CurrencyCode_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Currency Code")
            {
            }
            column(CurrencyFactor_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Currency Factor")
            {
            }
            column(ApplicationDocType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Application Doc. Type")
            {
            }
            column(ApplicationDocNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Application Doc. No")
            {
            }
            column(OriginalDocType_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Original Doc. Type"')
            {
            }
            column(OriginalDocNo_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Original Doc. No."')
            {
            }
            column(AppliedFromEntryNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Applied From Entry No.")
            {
            }
            column(ReversedEntryNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Reversed Entry No.")
            {
            }
            column(RemainingClosed_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Remaining Closed")
            {
            }
            column(GSTRoundingPrecision_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Rounding Precision")
            {
            }
            column(GSTRoundingType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Rounding Type")
            {
            }
            column(LocationCode_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Location Code")
            {
            }
            column(GSTCustomerType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Customer Type")
            {
            }
            column(GSTVendorType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."GST Vendor Type")
            {
            }
            column(CLEVLEEntryNo_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."CLE/VLE Entry No."')
            {
            }
            column(BillOfExportNo_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Bill Of Export No."')
            {
            }
            column(BillOfExportDate_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Bill Of Export Date"')
            {
            }
            column(eCommMerchantId_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."e-Comm. Merchant Id"')
            {
            }
            column(eCommOperatorGSTRegNo_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."e-Comm. Operator GST Reg. No."')
            {
            }
            column(InvoiceType_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Invoice Type"')
            {
            }
            column(OriginalInvoiceNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Original Invoice No.")
            {
            }
            column(OriginalInvoiceDate_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Original Invoice Date"')
            {
            }
            column(ReconciliationMonth_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Reconciliation Month")
            {
            }
            column(ReconciliationYear_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Reconciliation Year")
            {
            }
            column(Reconciled_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry".Reconciled)
            {
            }
            column(CreditAvailed_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Credit Availed")
            {
            }
            column(Paid_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry".Paid)
            {
            }
            column(AmounttoCustomerVendor_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Amount to Customer/Vendor"')
            {
            }
            column(CreditAdjustmentType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Credit Adjustment Type")
            {
            }
            column(AdvPmtAdjustment_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Adv. Pmt. Adjustment"')
            {
            }
            column(OriginalAdvPmtDocNo_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Original Adv. Pmt Doc. No."')
            {
            }
            column(OriginalAdvPmtDocDate_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Original Adv. Pmt Doc. Date"')
            {
            }
            column(PaymentDocumentDate; '"Detailed GST Ledger Entry"."Payment Document Date"')
            {
            }
            column(Cess_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry".Cess')
            {
            }
            column(UnApplied_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry".UnApplied)
            {
            }
            column(ItemLedgerEntryNo_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Item Ledger Entry No."')
            {
            }
            column(CreditReversal_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Credit Reversal"')
            {
            }
            column(GSTPlaceofSupply; "Detailed GST Ledger Entry"."GST Place of Supply")
            {
            }
            column(ItemChargeAssgnLineNo_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Item Charge Assgn. Line No."')
            {
            }
            column(PaymentType_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Payment Type")
            {
            }
            column(Distributed_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry".Distributed)
            {
            }
            column(DistributedReversed_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Distributed Reversed")
            {
            }
            column(InputServiceDistribution_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Input Service Distribution")
            {
            }
            column(Opening_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry".Opening)
            {
            }
            column(RemainingAmountClosed_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Remaining Amount Closed"')
            {
            }
            column(RemainingBaseAmount_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Remaining Base Amount")
            {
            }
            column(RemainingGSTAmount_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Remaining GST Amount")
            {
            }
            column(GenBusPostingGroup_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Gen. Bus. Posting Group"')
            {
            }
            column(GenProdPostingGroup_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Gen. Prod. Posting Group"')
            {
            }
            column(ReasonCode_DetailedGSTLedgerEntry; '"Detailed GST Ledger Entry"."Reason Code"')
            {
            }
            column(DistDocumentNo_DetailedGSTLedgerEntry; "Detailed GST Ledger Entry"."Dist. Document No.")
            {
            }
            column(CGST17; CGST17)
            {
            }
            column(SGST17; SGST17)
            {
            }
            column(IGST17; IGST17)
            {
            }
            column(UTGST17; UTGST17)
            {
            }
            column(CESS17; CESS17)
            {
            }
            column(CompName; CompInfo.Name)
            {
            }
            column(SNo; SNo)
            {
            }
            column(CompAdd; CompInfo.Address + ' ' + CompInfo."Address 2" + ' ' + CompInfo.City + ' - ' + CompInfo."Post Code")
            {
            }
            column(GSTPer; GSTPer)
            {
            }
            column(ItmDesc; ItmDesc)
            {
            }
            column(LocationToName; Loc11.Name)
            {
            }
            column(LocationToAdd; Loc11.Address + ' ' + Loc11."Address 2")
            {
            }
            column(Name17; Name17)
            {
            }
            column(Address17; Address17)
            {
            }
            column(PlaceofSupply; PlaceofSupply)
            {
            }

            trigger OnAfterGetRecord()
            begin

                TTT += 1;
                //>>Document No.

                DGST18.RESET;
                DGST18.SETRANGE("Document No.", "Document No.");
                DGST18.SETRANGE("Document Line No.", "Document Line No.");
                IF DGST18.FINDFIRST THEN BEGIN
                    III := DGST18.COUNT;
                END;

                IF TTT = III THEN BEGIN
                    SNo += 1;
                    TTT := 0;

                END;
                //>>Document No.


                //>>ItmDescription
                CLEAR(ItmDesc);
                IF Type = Type::Item THEN BEGIN
                    Itm17.RESET;
                    IF Itm17.GET("No.") THEN
                        ItmDesc := Itm17.Description;

                END;

                IF Type = Type::"G/L Account" THEN BEGIN
                    GL17.RESET;
                    IF GL17.GET("No.") THEN
                        ItmDesc := GL17.Name;

                END;
                //>>ItmDescription

                //>>Customer-Vendor Details
                CLEAR(Name17);
                CLEAR(Address17);
                CLEAR(PlaceofSupply);

                IF "Source Type" = "Source Type"::Customer THEN BEGIN
                    Cust17.RESET;
                    IF Cust17.GET("Source No.") THEN BEGIN
                        Name17 := Cust17.Name;
                        Address17 := Cust17.Address + ' ' + Cust17."Address 2" + ' ' + Cust17.City + ' - ' + Cust17."Post Code";

                        State18.RESET;
                        IF State18.GET(Cust17."State Code") THEN
                            PlaceofSupply := State18.Description;
                    END;
                END;

                IF "Source Type" = "Source Type"::Vendor THEN BEGIN
                    Vend17.RESET;
                    IF Vend17.GET("Source No.") THEN BEGIN
                        Name17 := Vend17.Name;
                        Address17 := Vend17.Address + ' ' + Vend17."Address 2" + ' ' + Vend17.City + ' - ' + Vend17."Post Code";

                        State18.RESET;
                        /* IF State18.GET("Location State Code") THEN
                            PlaceofSupply := State18.Description; */

                    END;
                END;

                //>>Customer-Vendor Details

                //>>DetailedGST
                CLEAR(GSTPer);
                CLEAR(CGST17);
                CLEAR(SGST17);
                CLEAR(IGST17);
                CLEAR(UTGST17);
                CLEAR(CESS17);

                DGST.RESET;
                DGST.SETRANGE("Document No.", "Document No.");
                DGST.SETRANGE("Document Line No.", "Document Line No.");
                IF DGST.FINDSET THEN
                    REPEAT


                        IF DGST."GST Component Code" = 'CGST' THEN
                            CGST17 := DGST."GST Amount";

                        IF DGST."GST Component Code" = 'SGST' THEN
                            SGST17 := DGST."GST Amount";

                        IF DGST."GST Component Code" = 'IGST' THEN
                            IGST17 := DGST."GST Amount";

                        IF DGST."GST Component Code" = 'UTGST' THEN
                            UTGST17 := DGST."GST Amount";

                        IF DGST."GST Component Code" = 'CESS' THEN
                            CESS17 := DGST."GST Amount";



                        GSTPer += DGST."GST %";

                    UNTIL DGST.NEXT = 0;
                //>>DetailedGST
            end;

            trigger OnPreDataItem()
            begin

                //SNo := 0;
                TTT := 0;
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

    trigger OnPreReport()
    begin

        CompInfo.GET;
    end;

    var
        CGST17: Decimal;
        SGST17: Decimal;
        IGST17: Decimal;
        UTGST17: Decimal;
        CESS17: Decimal;
        CompInfo: Record 79;
        Loc17: Record 14;
        Loc11: Record 14;
        SNo: Integer;
        TSH: Record 5744;
        DGST: Record "Detailed GST Ledger Entry";
        GSTPer: Decimal;
        Itm17: Record 27;
        GL17: Record 15;
        ItmDesc: Text[50];
        Cust17: Record 18;
        Vend17: Record 23;
        Name17: Text[50];
        Address17: Text[150];
        State18: Record State;
        Loc18: Record 14;
        PlaceofSupply: Text[50];
        TempDocNo: Code[20];
        TempLineNo: Integer;
        DGST18: Record "Detailed GST Ledger Entry";
        III: Integer;
        TTT: Integer;
}


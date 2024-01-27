report 50216 "RM Valuation Report --- 1"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 31Jan2018   RB-N         Document Type--SalesReturnShipment & Displaying Posted SalesCreditMemo Details
    // 16Jul2018   RB-N         Skipping Location with Include in Valuation--False
    // 04Sep2018   RB-N         Allowing only RawMaterial ItemCode for REPSOL--CAT15
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/RMValuationReport1.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            DataItemTableView = SORTING("Item Category Code", "Item No.", "Lot No.")
                                WHERE("Item Category Code" = FILTER('CAT01' | 'CAT05' | 'CAT06' | 'CAT04' | 'CAT07' | 'CAT08' | 'CAT09' | 'CAT10' | 'CAT16' | 'CAT15' | 'CAT19' | 'CAT21' | 'CAT17' | 'CAT14'),
                                      "Document Type" = FILTER("Transfer Shipment" | "Transfer Receipt" | "Purchase Receipt" | "Sales Return Receipt" | "Sales Credit Memo" | ' '));
            RequestFilterFields = "Location Code", "Item Category Code";

            trigger OnAfterGetRecord()
            begin

                //>>04Sep2018
                RepItmNo := '';
                IF "Item Category Code" = 'CAT15' THEN BEGIN
                    RepItmNo := COPYSTR("Item No.", 1, 4);
                    IF (RepItmNo = 'FGRP') OR (RepItmNo = 'FGRE') THEN
                        CurrReport.SKIP;
                END;
                //<<04Sep2018

                //>>05Sep2018
                IF "Document Type" = "Document Type"::"Transfer Shipment" THEN
                    IF "Location Code" <> 'IN-TRANS' THEN
                        CurrReport.SKIP;

                //<<05Sep2018

                //>>09Mar2017 LocName1
                CLEAR(LocName1);
                recLOC.RESET;
                recLOC.SETRANGE(Code, "Item Ledger Entry"."Location Code");
                IF recLOC.FINDFIRST THEN BEGIN
                    LocName1 := recLOC.Name;
                END;
                //<< LocName1

                //>>RB-N 16Jul2018
                recLOC.RESET;
                recLOC.SETRANGE(Code, "Location Code");
                recLOC.SETRANGE("Include in Valuation Report", FALSE);
                IF recLOC.FINDFIRST THEN
                    CurrReport.SKIP;
                //<<RB-N 16Jul2018

                //>>RB-N 05Sep2018 Skipping Closed Location
                recLOC.RESET;
                recLOC.SETRANGE(Code, "Location Code");
                recLOC.SETRANGE(Closed, TRUE);
                IF recLOC.FINDFIRST THEN
                    CurrReport.SKIP;
                //<<RB-N 05Sep2018 Skipping Closed Location

                vLandedCostAmt := 0;
                PurcRcptLine.RESET;
                PurcRcptLine.SETRANGE(PurcRcptLine."Document No.", "Item Ledger Entry"."Document No.");
                PurcRcptLine.SETRANGE(PurcRcptLine."No.", "Item Ledger Entry"."Item No.");
                IF PurcRcptLine.FIND('-') THEN BEGIN
                    recPurchInvLine.RESET;
                    // recPurchInvLine.SETRANGE(recPurchInvLine."Receipt Document No.", PurcRcptLine."Document No.");
                    recPurchInvLine.SETRANGE(recPurchInvLine."No.", "Item Ledger Entry"."Item No.");
                    IF recPurchInvLine.FINDSET THEN
                        REPEAT
                            vLandedCostAmt += recPurchInvLine."Landed Cost";
                        UNTIL recPurchInvLine.NEXT = 0;
                END;

                itemappnotfind := FALSE;
                IMPORTLOCAL := '';
                //CLEAR(GRNTRans);
                "OPENING/POSITIVE" := '';
                posadj := FALSE;
                outptbool := FALSE;
                IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Positive Adjmt." THEN BEGIN
                    "OPENING/POSITIVE" := 'OPENING/POSITIVE';
                END;

                IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Consumption THEN
                    IF NOT "Item Ledger Entry".Positive THEN
                        CurrReport.SKIP;

                IF "Item Ledger Entry"."Location Code" = 'INVADJ0001' THEN
                    CurrReport.SKIP;

                ItemRec.GET("Item Ledger Entry"."Item No.");
                VendorName := '';
                //IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Purchase THEN
                IF VendorRec.GET("Item Ledger Entry"."Source No.") THEN
                    VendorName := VendorRec.Name;

                RemQty := 0;
                ItemApplication.RESET;
                ItemApplication.SETRANGE(ItemApplication."Inbound Item Entry No.", "Item Ledger Entry"."Entry No.");
                ItemApplication.SETRANGE(ItemApplication."Posting Date", 0D, PostingDate);
                IF ItemApplication.FINDSET THEN
                    REPEAT
                        RemQty += ItemApplication.Quantity;
                    UNTIL ItemApplication.NEXT = 0;

                //>>05May2019
                CLEAR(NegativeCorrection);
                IF "Entry Type" = "Entry Type"::"Negative Adjmt." THEN
                    IF "Applies-to Entry" <> 0 THEN BEGIN
                        CLEAR(ILE0505);
                        IF ILE0505.GET("Applies-to Entry") THEN
                            IF "Posting Date" < ILE0505."Posting Date" THEN BEGIN
                                NegativeCorrection := TRUE;
                                RemQty := Quantity;
                            END;
                    END;
                //<<05May2019
                //>>17May2019
                IF NegativeCorrection THEN BEGIN
                    IF "Posting Date" < PostingDate THEN
                        NegativeCorrection := FALSE;
                END;
                //<<17May2019

                //>>05Sep2018
                IF NOT NegativeCorrection THEN //05May2019
                    IF RemQty < 0 THEN
                        CurrReport.SKIP;
                //<<05Sep2018

                IF RemQty = 0 THEN
                    IF NOT Showall THEN
                        CurrReport.SKIP;

                RateofItem := 0;
                TotalCost := 0;
                TotalQty := 0;
                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");//07May2018
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Charge No.", '');
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        TotalCost += ValueEntry."Cost Amount (Expected)" + ValueEntry."Cost Amount (Actual)";
                        TotalQty += ValueEntry."Item Ledger Entry Quantity";
                    UNTIL ValueEntry.NEXT = 0;
                RateofItem := TotalCost / TotalQty;

                ExpCost := 0;
                ActCost := 0;
                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");//07May2018
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '');
                ValueEntry.SETFILTER(ValueEntry."Item Ledger Entry Quantity", '<>%1', 0);
                ValueEntry.SETRANGE("Document Type", ValueEntry."Document Type"::"Purchase Invoice");
                //ValueEntry.SETRANGE("Document Type",ValueEntry."Document Type"::"Purchase Credit Memo");//28Jun2018 //  rspl ad
                IF ValueEntry.FINDFIRST THEN BEGIN
                    ExpCost += ValueEntry."Cost per Unit";

                END;
                //rspl ad

                CLEAR(PCM);
                CLEAR(PCI);
                CLEAR(Totalv);
                ValueEntry.RESET;
                ValueEntry.SETRANGE(ValueEntry."Document No.", "Document No.");
                IF ValueEntry.FINDFIRST THEN BEGIN
                    newValueEntry.RESET();
                    newValueEntry.SETRANGE(newValueEntry."Item Ledger Entry No.", ValueEntry."Item Ledger Entry No.");
                    newValueEntry.SETRANGE(newValueEntry."Document Type", newValueEntry."Document Type"::"Purchase Credit Memo");
                    IF newValueEntry.FINDFIRST THEN
                        REPEAT
                            PCM += newValueEntry."Cost per Unit";
                        UNTIL newValueEntry.NEXT = 0;
                    // MESSAGE('%1',ABS(PCM));
                    //  newValueEntry.RESET();
                    //  newValueEntry.SETRANGE(newValueEntry."Item Ledger Entry No.",ValueEntry."Item Ledger Entry No.");
                    // newValueEntry.SETRANGE(newValueEntry."Document Type",newValueEntry."Document Type"::"Purchase Invoice");
                    // IF newValueEntry.FINDFIRST THEN
                    // REPEAT
                    //PCI+= newValueEntry."Cost per Unit";
                    //UNTIL newValueEntry.NEXT=0;
                    // Totalv:=  ABS(PCM)- ABS(ExpCost);
                END;
                //rspl ad



                //>>06Sep2018
                ExpCost := 0;

                CALCFIELDS("Cost Amount (Expected)", "Cost Amount (Actual)");
                ExpCost := ("Cost Amount (Expected)" + "Cost Amount (Actual)") / Quantity;
                //<<06Sep2018



                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");//07May2018
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '');
                ValueEntry.SETFILTER(ValueEntry."Invoiced Quantity", '<>%1', 0);
                ValueEntry.SETRANGE("Document Type", ValueEntry."Document Type"::"Purchase Invoice");//28Jun2018
                IF ValueEntry.FINDFIRST THEN BEGIN
                    ActCost := ValueEntry."Cost per Unit";
                END;

                TRRNo := '';
                TRRDate := 0D;
                WHBENo := '';//31Jan2018
                DebondBENo := '';//31Jan2018
                IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer THEN BEGIN
                    TRRNo := "Item Ledger Entry"."Document No.";
                    TRRDate := "Item Ledger Entry"."Posting Date";
                    TransferRecipt.RESET;
                    //TransferRecipt.GET("Item Ledger Entry"."Document No.");
                    IF TransferRecipt.GET("Item Ledger Entry"."Document No.") THEN //07May2018
                    BEGIN
                        WHBENo := TransferRecipt."WH Bill Entry No.";
                        DebondBENo := TransferRecipt."Debond Bill Entry No.";
                    END;
                END;

                BRTno := '';
                BRTdate := 0D;
                recTSH.RESET;
                recTSH.SETCURRENTKEY("Transfer Order No.");//07May2018
                recTSH.SETRANGE(recTSH."Transfer Order No.", "Item Ledger Entry"."Order No.");
                IF recTSH.FINDFIRST THEN BEGIN
                    BRTno := recTSH."No.";
                    BRTdate := recTSH."Posting Date"
                END;

                //>>10Aug2018
                VendorCode := '';
                VendName := '';
                VendorInvoiceno := '';
                GRNTRans := '';
                vend := FALSE;
                posadj := FALSE;
                opbal := '';
                opbalDate := 0D;
                //<<10Aug2018

                IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer THEN BEGIN
                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");
                    recILE.SETRANGE(recILE."Document No.", BRTno);
                    recILE.SETRANGE(recILE."Item No.", "Item Ledger Entry"."Item No.");
                    recILE.SETFILTER(recILE."Location Code", '<>%1', 'IN-TRANS');
                    IF recILE.FINDFIRST THEN BEGIN
                        recItemApplEntry.RESET;
                        recItemApplEntry.SETRANGE(recItemApplEntry."Outbound Item Entry No.", recILE."Entry No.");
                        //recItemApplEntry.SETFILTER(recItemApplEntry.Quantity,'<=%1',recILE.Quantity);
                        recItemApplEntry.SETFILTER(recItemApplEntry.Quantity, '<=%1', "Item Ledger Entry".Quantity);
                        IF recItemApplEntry.FINDFIRST THEN BEGIN
                            recItemLedEntry.RESET;
                            recItemLedEntry.SETRANGE(recItemLedEntry."Entry No.", recItemApplEntry."Inbound Item Entry No.");
                            IF recItemLedEntry.FINDFIRST THEN BEGIN
                                IF recItemLedEntry."Document Type" = recItemLedEntry."Document Type"::"Purchase Receipt" THEN BEGIN
                                    recRPH.RESET;
                                    recRPH.SETRANGE(recRPH."No.", recItemLedEntry."Document No.");
                                    IF recRPH.FINDFIRST THEN BEGIN
                                        VendorCode := '';
                                        VendName := '';
                                        VendorInvoiceno := '';
                                        VendorCode := recRPH."Buy-from Vendor No.";
                                        VendName := recRPH."Buy-from Vendor Name";
                                        VendorInvoiceno := recItemLedEntry."External Document No.";
                                        GRNTRans := recRPH."No.";
                                        vend := TRUE;
                                    END;
                                END;

                                IF recItemLedEntry."Entry Type" = recItemLedEntry."Entry Type"::"Positive Adjmt." THEN BEGIN
                                    VendorCode := '';
                                    VendName := '';
                                    VendorInvoiceno := '';
                                    GRNTRans := '';
                                    vend := TRUE;
                                    posadj := TRUE;
                                    opbal := recItemLedEntry."Document No.";
                                    opbalDate := recItemLedEntry."Posting Date";
                                END;

                                IF recItemLedEntry."Entry Type" = recItemLedEntry."Entry Type"::Output THEN BEGIN
                                    VendorCode := '';
                                    VendName := '';
                                    VendorInvoiceno := '';
                                    GRNTRans := '';
                                    outptbool := TRUE;
                                    outpt := recItemLedEntry."Document No.";
                                    outptDate := recItemLedEntry."Posting Date";
                                END ELSE BEGIN
                                    TRansrec.RESET;
                                    TRansrec.SETRANGE("No.", recItemLedEntry."Document No.");
                                    IF TRansrec.FINDFIRST THEN BEGIN
                                        recILEnew.RESET;
                                        recILEnew.SETCURRENTKEY("Item No.");
                                        recILEnew.SETRANGE(recILEnew."Order No.", TRansrec."Transfer Order No.");
                                        recILEnew.SETRANGE(recILEnew."Item No.", "Item Ledger Entry"."Item No.");
                                        recILEnew.SETFILTER(recILEnew."Location Code", '<>%1', 'IN-TRANS');
                                        IF recILEnew.FINDFIRST THEN BEGIN
                                            recItemApplicEntry1.RESET;
                                            recItemApplicEntry1.SETRANGE(recItemApplicEntry1."Outbound Item Entry No.", recILEnew."Entry No.");
                                            recItemApplicEntry1.SETFILTER(recItemApplicEntry1.Quantity, '<=%1', "Item Ledger Entry".Quantity);
                                            IF recItemApplicEntry1.FINDFIRST THEN BEGIN
                                                recItemledgentry.RESET;
                                                recItemledgentry.SETRANGE(recItemledgentry."Entry No.", recItemApplicEntry1."Inbound Item Entry No.");
                                                IF recItemledgentry.FINDFIRST THEN BEGIN
                                                    IF recItemledgentry."Document Type" = recItemledgentry."Document Type"::"Purchase Receipt" THEN BEGIN
                                                        purchrecpthdr.RESET;
                                                        purchrecpthdr.SETRANGE(purchrecpthdr."No.", recItemledgentry."Document No.");
                                                        IF purchrecpthdr.FINDFIRST THEN BEGIN
                                                            VendorCode := '';
                                                            VendName := '';
                                                            VendorInvoiceno := '';
                                                            VendorCode := purchrecpthdr."Buy-from Vendor No.";
                                                            VendName := purchrecpthdr."Buy-from Vendor Name";
                                                            VendorInvoiceno := recItemledgentry."External Document No.";
                                                            GRNTRans := purchrecpthdr."No.";
                                                            vend := TRUE;
                                                        END;
                                                    END;
                                                    IF recItemledgentry."Entry Type" = recItemledgentry."Entry Type"::"Positive Adjmt." THEN BEGIN
                                                        VendorCode := '';
                                                        VendName := '';
                                                        VendorInvoiceno := '';
                                                        GRNTRans := '';
                                                        vend := TRUE;
                                                        opbal := recItemledgentry."Document No.";
                                                        opbalDate := recItemledgentry."Posting Date";
                                                        posadj := TRUE;
                                                    END;

                                                    IF recItemledgentry."Entry Type" = recItemledgentry."Entry Type"::Output THEN BEGIN
                                                        VendorCode := '';
                                                        VendName := '';
                                                        VendorInvoiceno := '';
                                                        GRNTRans := '';
                                                        outpt := recItemledgentry."Document No.";
                                                        outptDate := recItemledgentry."Posting Date";
                                                        outptbool := TRUE;
                                                    END ELSE BEGIN
                                                        Transsrechd.RESET;
                                                        Transsrechd.SETRANGE(Transsrechd."No.", recItemledgentry."Document No.");
                                                        IF Transsrechd.FINDFIRST THEN BEGIN
                                                            ILErec.RESET;
                                                            ILErec.SETRANGE(ILErec."Order No.", Transsrechd."Transfer Order No.");
                                                            ILErec.SETRANGE(ILErec."Item No.", "Item Ledger Entry"."Item No.");
                                                            ILErec.SETFILTER(ILErec."Location Code", '<>%1', 'IN-TRANS');
                                                            IF ILErec.FINDFIRST THEN BEGIN
                                                                IAErec.RESET;
                                                                IAErec.SETRANGE(IAErec."Outbound Item Entry No.", ILErec."Entry No.");
                                                                IAErec.SETFILTER(IAErec.Quantity, '<=%1', "Item Ledger Entry".Quantity);
                                                                IF IAErec.FINDFIRST THEN BEGIN
                                                                    ILErec1.RESET;
                                                                    ILErec1.SETRANGE(ILErec1."Entry No.", IAErec."Inbound Item Entry No.");
                                                                    IF ILErec1.FINDFIRST THEN BEGIN
                                                                        IF ILErec1."Document Type" = ILErec1."Document Type"::"Purchase Receipt" THEN BEGIN
                                                                            Purchrcpthhd.RESET;
                                                                            Purchrcpthhd.SETRANGE(Purchrcpthhd."No.", ILErec1."Document No.");
                                                                            IF Purchrcpthhd.FINDFIRST THEN BEGIN
                                                                                VendorCode := '';
                                                                                VendName := '';
                                                                                VendorInvoiceno := '';
                                                                                VendorCode := Purchrcpthhd."Buy-from Vendor No.";
                                                                                VendName := Purchrcpthhd."Buy-from Vendor Name";
                                                                                VendorInvoiceno := ILErec1."External Document No.";
                                                                                GRNTRans := Purchrcpthhd."No.";
                                                                                vend := TRUE;
                                                                            END;
                                                                        END;

                                                                        IF ILErec1."Entry Type" = ILErec1."Entry Type"::"Positive Adjmt." THEN BEGIN
                                                                            VendorCode := '';
                                                                            VendName := '';
                                                                            VendorInvoiceno := '';
                                                                            GRNTRans := '';
                                                                            opbal := ILErec1."Document No.";
                                                                            opbalDate := ILErec1."Posting Date";
                                                                            vend := TRUE;
                                                                            posadj := TRUE;
                                                                        END ELSE BEGIN
                                                                            TRSHD.RESET;
                                                                            TRSHD.SETRANGE(TRSHD."No.", ILErec1."Document No.");
                                                                            IF TRSHD.FINDFIRST THEN BEGIN
                                                                                ILErec2.RESET;
                                                                                ILErec2.SETRANGE(ILErec2."Order No.", TRSHD."Transfer Order No.");
                                                                                ILErec2.SETRANGE(ILErec2."Item No.", "Item Ledger Entry"."Item No.");
                                                                                ILErec2.SETFILTER(ILErec2."Location Code", '<>%1', 'IN-TRANS');
                                                                                IF ILErec2.FINDFIRST THEN BEGIN
                                                                                    ITMAPPENT.RESET;
                                                                                    ITMAPPENT.SETRANGE(ITMAPPENT."Outbound Item Entry No.", ILErec2."Entry No.");
                                                                                    ITMAPPENT.SETFILTER(ITMAPPENT.Quantity, '<=%1', "Item Ledger Entry".Quantity);
                                                                                    IF ITMAPPENT.FINDFIRST THEN BEGIN
                                                                                        ILErec3.RESET;
                                                                                        ILErec3.SETRANGE(ILErec3."Entry No.", ITMAPPENT."Inbound Item Entry No.");
                                                                                        IF ILErec3.FINDFIRST THEN BEGIN
                                                                                            IF ILErec3."Document Type" = ILErec3."Document Type"::"Purchase Receipt" THEN BEGIN
                                                                                                PURCHRECHDR.RESET;
                                                                                                PURCHRECHDR.SETRANGE(PURCHRECHDR."No.", ILErec3."Document No.");
                                                                                                IF PURCHRECHDR.FINDFIRST THEN BEGIN
                                                                                                    VendorCode := '';
                                                                                                    VendName := '';
                                                                                                    VendorInvoiceno := '';
                                                                                                    VendorCode := PURCHRECHDR."Buy-from Vendor No.";
                                                                                                    VendName := PURCHRECHDR."Buy-from Vendor Name";
                                                                                                    VendorInvoiceno := ILErec3."External Document No.";
                                                                                                    GRNTRans := PURCHRECHDR."No.";
                                                                                                    vend := TRUE;
                                                                                                END;
                                                                                            END;

                                                                                            IF ILErec3."Entry Type" = ILErec3."Entry Type"::"Positive Adjmt." THEN BEGIN
                                                                                                VendorCode := '';
                                                                                                VendName := '';
                                                                                                VendorInvoiceno := '';
                                                                                                GRNTRans := '';
                                                                                                opbal := ILErec3."Document No.";
                                                                                                opbalDate := ILErec3."Posting Date";
                                                                                                vend := TRUE;
                                                                                                posadj := TRUE;
                                                                                            END ELSE BEGIN
                                                                                                TRRECHD.RESET;
                                                                                                TRRECHD.SETRANGE(TRRECHD."No.", ILErec3."Document No.");
                                                                                                IF TRRECHD.FINDFIRST THEN BEGIN
                                                                                                    ILErec4.RESET;
                                                                                                    ILErec4.SETRANGE(ILErec4."Order No.", TRRECHD."Transfer Order No.");
                                                                                                    ILErec4.SETRANGE(ILErec4."Item No.", "Item Ledger Entry"."Item No.");
                                                                                                    ILErec4.SETRANGE(ILErec4."Location Code", '<>%1', 'IN-TRANS');
                                                                                                    IF ILErec4.FINDFIRST THEN BEGIN
                                                                                                        ITMAPPPLENTRY.RESET;
                                                                                                        ITMAPPPLENTRY.SETRANGE(ITMAPPPLENTRY."Outbound Item Entry No.", ILErec4."Entry No.");
                                                                                                        ITMAPPPLENTRY.SETFILTER(ITMAPPPLENTRY.Quantity, '<=%1', "Item Ledger Entry".Quantity);
                                                                                                        IF ITMAPPPLENTRY.FINDFIRST THEN BEGIN
                                                                                                            ILErec5.RESET;
                                                                                                            ILErec5.SETRANGE(ILErec5."Entry No.", ITMAPPPLENTRY."Inbound Item Entry No.");
                                                                                                            IF ILErec5.FINDFIRST THEN BEGIN
                                                                                                                IF ILErec5."Document Type" = ILErec5."Document Type"::"Purchase Receipt" THEN BEGIN
                                                                                                                    Purchrecheaderr.RESET;
                                                                                                                    Purchrecheaderr.SETRANGE(Purchrecheaderr."No.", ILErec5."Document No.");
                                                                                                                    IF Purchrecheaderr.FINDFIRST THEN BEGIN
                                                                                                                        VendorCode := '';
                                                                                                                        VendName := '';
                                                                                                                        VendorInvoiceno := '';
                                                                                                                        VendorCode := Purchrecheaderr."Buy-from Vendor No.";
                                                                                                                        VendName := Purchrecheaderr."Buy-from Vendor Name";
                                                                                                                        VendorInvoiceno := ILErec5."External Document No.";
                                                                                                                        GRNTRans := Purchrecheaderr."No.";
                                                                                                                        vend := TRUE;
                                                                                                                    END;
                                                                                                                END;

                                                                                                                IF ILErec5."Entry Type" = ILErec5."Entry Type"::"Positive Adjmt." THEN BEGIN
                                                                                                                    VendorCode := '';
                                                                                                                    VendName := '';
                                                                                                                    VendorInvoiceno := '';
                                                                                                                    GRNTRans := '';
                                                                                                                    opbal := ILErec5."Document No.";
                                                                                                                    opbalDate := ILErec5."Posting Date";
                                                                                                                    vend := TRUE;
                                                                                                                    posadj := TRUE;
                                                                                                                END;
                                                                                                            END;
                                                                                                        END;
                                                                                                    END;
                                                                                                END;
                                                                                            END;
                                                                                        END;
                                                                                    END;
                                                                                END;
                                                                            END;
                                                                        END;
                                                                    END;
                                                                END;
                                                            END;
                                                        END;
                                                    END;
                                                END;
                                            END;
                                        END;
                                    END;
                                END;
                            END;
                        END ELSE
                            itemappnotfind := TRUE;  // ADDED on 200114 Item application entry not finded in transfer case EBT MILAN

                    END;
                END;


                //>>05Sep2018 GRN No. for the Transfer
                TempILEEntryNo05 := 0;
                GRNFound05 := FALSE;
                GRNEntryNo := 0;

                IF ("Entry Type" = "Entry Type"::Transfer) AND (GRNTRans = '') THEN BEGIN
                    ItmApp05.RESET;
                    ItmApp05.SETCURRENTKEY("Inbound Item Entry No.");
                    ItmApp05.SETRANGE("Inbound Item Entry No.", "Entry No.");
                    ItmApp05.SETFILTER(Quantity, '>0');
                    IF ItmApp05.FINDFIRST THEN BEGIN
                        IF ItmApp05."Transferred-from Entry No." <> 0 THEN BEGIN
                            TempILEEntryNo05 := ItmApp05."Transferred-from Entry No.";

                            ItmApp05_01.RESET;
                            ItmApp05_01.SETCURRENTKEY("Inbound Item Entry No.");
                            ItmApp05_01.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                            ItmApp05_01.SETFILTER(Quantity, '>0');
                            IF ItmApp05_01.FINDFIRST THEN
                                IF ItmApp05_01."Transferred-from Entry No." <> 0 THEN BEGIN
                                    TempILEEntryNo05 := ItmApp05."Transferred-from Entry No.";
                                END ELSE BEGIN
                                    GRNFound05 := TRUE;
                                    GRNEntryNo := ItmApp05_01."Inbound Item Entry No.";
                                END;

                            //2ndLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_02.RESET;
                                ItmApp05_02.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_02.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_02.SETFILTER(Quantity, '>0');
                                IF ItmApp05_02.FINDFIRST THEN
                                    IF ItmApp05_02."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_02."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_02."Inbound Item Entry No.";
                                    END;
                            END;

                            //3rdLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_03.RESET;
                                ItmApp05_03.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_03.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_03.SETFILTER(Quantity, '>0');
                                IF ItmApp05_03.FINDFIRST THEN
                                    IF ItmApp05_03."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_03."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_03."Inbound Item Entry No.";
                                    END;
                            END;

                            //4thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_04.RESET;
                                ItmApp05_04.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_04.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_04.SETFILTER(Quantity, '>0');
                                IF ItmApp05_04.FINDFIRST THEN
                                    IF ItmApp05_04."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_04."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_04."Inbound Item Entry No.";
                                    END;
                            END;

                            //5thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_05.RESET;
                                ItmApp05_05.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_05.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_05.SETFILTER(Quantity, '>0');
                                IF ItmApp05_05.FINDFIRST THEN
                                    IF ItmApp05_05."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_05."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_05."Inbound Item Entry No.";
                                    END;
                            END;

                            //6thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_01.RESET;
                                ItmApp05_01.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_01.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_01.SETFILTER(Quantity, '>0');
                                IF ItmApp05_01.FINDFIRST THEN
                                    IF ItmApp05_01."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_01."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_01."Inbound Item Entry No.";
                                    END;
                            END;

                            //7thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_02.RESET;
                                ItmApp05_02.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_02.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_02.SETFILTER(Quantity, '>0');
                                IF ItmApp05_02.FINDFIRST THEN
                                    IF ItmApp05_02."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_02."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_02."Inbound Item Entry No.";
                                    END;
                            END;

                            //8thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_03.RESET;
                                ItmApp05_03.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_03.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_03.SETFILTER(Quantity, '>0');
                                IF ItmApp05_03.FINDFIRST THEN
                                    IF ItmApp05_03."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_03."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_03."Inbound Item Entry No.";
                                    END;
                            END;

                            //9thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_04.RESET;
                                ItmApp05_04.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_04.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_04.SETFILTER(Quantity, '>0');
                                IF ItmApp05_04.FINDFIRST THEN
                                    IF ItmApp05_04."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_04."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_04."Inbound Item Entry No.";
                                    END;
                            END;

                            //10thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_05.RESET;
                                ItmApp05_05.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_05.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_05.SETFILTER(Quantity, '>0');
                                IF ItmApp05_05.FINDFIRST THEN
                                    IF ItmApp05_05."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_05."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_05."Inbound Item Entry No.";
                                    END;
                            END;

                            //11thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_01.RESET;
                                ItmApp05_01.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_01.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_01.SETFILTER(Quantity, '>0');
                                IF ItmApp05_01.FINDFIRST THEN
                                    IF ItmApp05_01."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_01."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_01."Inbound Item Entry No.";
                                    END;
                            END;

                            //12thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_02.RESET;
                                ItmApp05_02.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_02.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_02.SETFILTER(Quantity, '>0');
                                IF ItmApp05_02.FINDFIRST THEN
                                    IF ItmApp05_02."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_02."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_02."Inbound Item Entry No.";
                                    END;
                            END;

                            //13thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_03.RESET;
                                ItmApp05_03.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_03.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_03.SETFILTER(Quantity, '>0');
                                IF ItmApp05_03.FINDFIRST THEN
                                    IF ItmApp05_03."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_03."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_03."Inbound Item Entry No.";
                                    END;
                            END;

                            //14thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_04.RESET;
                                ItmApp05_04.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_04.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_04.SETFILTER(Quantity, '>0');
                                IF ItmApp05_04.FINDFIRST THEN
                                    IF ItmApp05_04."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_04."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_04."Inbound Item Entry No.";
                                    END;
                            END;

                            //15thLevel
                            IF NOT GRNFound05 THEN BEGIN
                                ItmApp05_05.RESET;
                                ItmApp05_05.SETCURRENTKEY("Inbound Item Entry No.");
                                ItmApp05_05.SETRANGE("Inbound Item Entry No.", TempILEEntryNo05);
                                ItmApp05_05.SETFILTER(Quantity, '>0');
                                IF ItmApp05_05.FINDFIRST THEN
                                    IF ItmApp05_05."Transferred-from Entry No." <> 0 THEN BEGIN
                                        TempILEEntryNo05 := ItmApp05_05."Transferred-from Entry No.";
                                    END ELSE BEGIN
                                        GRNFound05 := TRUE;
                                        GRNEntryNo := ItmApp05_05."Inbound Item Entry No.";
                                    END;
                            END;

                        END;
                    END;
                END;

                IF GRNFound05 AND (GRNEntryNo <> 0) THEN BEGIN
                    ILE05.RESET;
                    IF ILE05.GET(GRNEntryNo) THEN
                        IF ILE05."Document Type" = ILE05."Document Type"::"Purchase Receipt" THEN BEGIN
                            PRH05.RESET;
                            IF PRH05.GET(ILE05."Document No.") THEN BEGIN
                                VendorCode := '';
                                VendName := '';
                                VendorCode := PRH05."Buy-from Vendor No.";
                                VendName := PRH05."Buy-from Vendor Name";
                                VendorInvoiceno := ILE05."External Document No.";
                                GRNTRans := PRH05."No.";
                                vend := TRUE;
                            END;
                        END;

                END;
                //<<05Sep2018 GRN No. for the Transfer

                ExciseAmount := 0;
                IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Purchase THEN BEGIN
                    recPHL.RESET;
                    recPHL.SETRANGE(recPHL."Document No.", "Item Ledger Entry"."Document No.");
                    recPHL.SETRANGE(recPHL."No.", "Item Ledger Entry"."Item No.");
                    recPHL.SETRANGE(recPHL."Line No.", "Item Ledger Entry"."Document Line No.");
                    IF recPHL.FINDFIRST THEN BEGIN
                        "Excise%" := 0;
                        /* recExcisePostingSetup.RESET;
                        recExcisePostingSetup.SETRANGE(recExcisePostingSetup."Excise Bus. Posting Group", recPHL."Excise Bus. Posting Group");
                        recExcisePostingSetup.SETRANGE(recExcisePostingSetup."Excise Prod. Posting Group", recPHL."Excise Prod. Posting Group");
                        IF recExcisePostingSetup.FINDFIRST THEN BEGIN
                            "Excise%" := recExcisePostingSetup."BED %" + (recExcisePostingSetup."BED %" * recExcisePostingSetup."eCess %" / 100) +
                                         (recExcisePostingSetup."BED %" * recExcisePostingSetup."SHE Cess %" / 100);
                        END; */
                        //  ExciseAmount :=  (RemQty * ROUND((ExpCost),0.01)) * "Excise%"/100;
                        //   ExciseAmount := (accessvalue * RemQty)
                    END;
                END;

                ShortageCharge := 0;
                TransportCharge := 0;
                CustomCharge := 0;
                AnyOtherCharge := 0;
                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '%1|%2|%3', 'AD 01', 'BO 01', 'CH 01');
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        ShortageCharge += ValueEntry."Cost per Unit";
                    UNTIL ValueEntry.NEXT = 0;

                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '%1|%2|%3|%4|%5|%6', 'FRTADD', 'FRTBASEOIL', 'FRTBASEOILIMP',
                                                          'FRTCHEM', 'FRTCONSU', 'FRTPKGMAT');
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        TransportCharge += ValueEntry."Cost per Unit";
                    UNTIL ValueEntry.NEXT = 0;


                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '%1|%2|%3', 'CUSTOMADDITIVES', 'CUSTOMBASEOILS', 'CUSTOMCHEMICALS');
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        CustomCharge += ValueEntry."Cost per Unit";
                    UNTIL ValueEntry.NEXT = 0;

                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '%1|%2|%3|%4|%5|%6|%7', 'AD 03', 'BO 03', 'CH 03', 'PM 01', 'PM 03',
                                                                         'ROUNDING', 'SKMKT FEES');
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        AnyOtherCharge += ValueEntry."Cost per Unit";
                    UNTIL ValueEntry.NEXT = 0;

                // EBT MILAN (18122013)---------------ADDED NEW CALCULATION ON DISC.------------------------------------------------START
                CLEAR(DiscountCost);
                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '%1|%2|%3|%4|%5|%6', 'AD 02', 'BO 02', 'CH 02', 'PKGMAT 01', 'PM 02', 'PURDISCONT');//RSPLSID19Dec2022
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        //MESSAGE('%1',"Item Ledger Entry"."Entry No.");
                        DiscountCost += ValueEntry."Cost per Unit";
                    //MESSAGE('%1',DiscountCost); // Fahim
                    UNTIL ValueEntry.NEXT = 0;

                //CLEAR(TotalCharges);
                //TotalCharges := ShortageCharge+TransportCharge+DiscountCost+CustomCharge+AnyOtherCharge;

                CLEAR(Pruchinvno);
                IF ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Purchase) AND
                        ("Item Ledger Entry"."Entry Type" <> "Item Ledger Entry"."Entry Type"::"Positive Adjmt.") THEN BEGIN
                    Purchinvline.RESET;
                    // Purchinvline.SETRANGE(Purchinvline."Receipt Document No.", "Item Ledger Entry"."Document No.");
                    IF Purchinvline.FINDFIRST THEN BEGIN
                        Pruchinvno := Purchinvline."Document No.";
                        PruchinvnoDate := Purchinvline."Posting Date";
                    END
                    ELSE
                        Pruchinvno := 'Pending';
                    /*
                      CLEAR(IMPORTLOCAL);
                      Purchrecpthead.RESET;
                      Purchrecpthead.SETRANGE(Purchrecpthead."No.","Item Ledger Entry"."Document No.");
                      Purchrecpthead.SETFILTER(Purchrecpthead."Location Code",'BOND0001');
                      IF Purchrecpthead.FINDFIRST THEN
                      IMPORTLOCAL := 'IMPORT'
                      ELSE
                      IMPORTLOCAL := 'LOCAL';
                    */
                END;



                CLEAR(Purchinvtrans);
                CLEAR(vendcodtran);
                CLEAR(PurchinvExtrans);
                CLEAR(Purchinvtrans);
                vLandedCostAmt := 0;
                //IF GRNTRans <> '' THEN
                //BEGIN
                Purchinvline.RESET;
                // Purchinvline.SETRANGE(Purchinvline."Receipt Document No.", GRNTRans);
                Purchinvline.SETRANGE(Purchinvline.Type, Purchinvline.Type::Item);
                Purchinvline.SETRANGE(Purchinvline."No.", "Item Ledger Entry"."Item No."); // ADDED MILAN 06012014
                IF Purchinvline.FINDFIRST THEN BEGIN

                    recPurchInvLine.RESET;
                    //recPurchInvLine.SETRANGE(recPurchInvLine."Receipt Document No.", GRNTRans);
                    recPurchInvLine.SETRANGE(recPurchInvLine."No.", "Item Ledger Entry"."Item No.");
                    IF recPurchInvLine.FINDSET THEN
                        REPEAT
                            vLandedCostAmt += recPurchInvLine."Landed Cost";
                        UNTIL recPurchInvLine.NEXT = 0;
                    //mayuri


                    Purchinvtrans := Purchinvline."Document No.";
                    PurchinvtransDate := Purchinvline."Posting Date";     // EBT MILAN Document Posing Date 260214
                                                                          //    vendcodtran := Purchinvline."Buy-from Vendor No.";
                                                                          //    vendtrans.GET(Purchinvline."Buy-from Vendor No.");
                    CLEAR(Vendnametran);
                    vendtrans.RESET;
                    vendtrans.SETRANGE(vendtrans."No.", Purchinvln."Buy-from Vendor No.");

                    IF vendtrans.FINDFIRST THEN BEGIN
                        Vendnametran := vendtrans.Name;
                        vendcodtran := vendtrans."No.";
                    END;
                    Purchinvhd.RESET;
                    Purchinvhd.SETRANGE(Purchinvhd."No.", Purchinvline."Document No.");
                    IF Purchinvhd.FINDFIRST THEN
                        PurchinvExtrans := Purchinvhd."Vendor Invoice No.";
                END
                ELSE
                    Purchinvtrans := 'Pending';

                // IF posadj=TRUE THEN
                // Purchinvtrans := opbal;

                //opbal := '';

                /*
                 CLEAR(IMPORTLOCAL);
                 Purchrecpthead.RESET;
                 Purchrecpthead.SETRANGE(Purchrecpthead."No.",GRNTRans);
                 Purchrecpthead.SETFILTER(Purchrecpthead."Location Code",'BOND0001');
                 IF Purchrecpthead.FINDFIRST THEN
                 IMPORTLOCAL := 'IMPORT'
                 ELSE
                 IMPORTLOCAL := 'LOCAL';
                 */

                //END;

                //MESSAGE('%1',GRNTRans);
                CLEAR(finalinvno);
                IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer THEN BEGIN
                    IF GRNTRans <> '' THEN
                        finalinvno := Purchinvtrans
                    ELSE
                        finalinvno := VendorInvoiceno;
                END ELSE BEGIN
                    IF ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Output) OR
                             ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Positive Adjmt.") THEN
                        finalinvno := "Item Ledger Entry"."Document No."
                    ELSE
                        finalinvno := Pruchinvno;
                END;

                //MESSAGE('%1',PRUCHINVNO);
                CLEAR(costcat5);
                CLEAR(accessvalue);
                Purchinvln.RESET;
                Purchinvln.SETRANGE(Purchinvln."Document No.", finalinvno);
                Purchinvln.SETRANGE(Purchinvln."No.", "Item Ledger Entry"."Item No.");
                IF Purchinvln.FINDFIRST THEN BEGIN
                    accessvalue := 0;// Purchinvln."Assessable Value";
                    ADEAmt := 0;//Purchinvln."ADE Amount";
                    costcat5 := Purchinvln."Unit Cost (LCY)";
                    //MESSAGE('%1....%2',accessvalue,costcat5);
                END;

                // EBT MILAN ADDED NOT TO OVERWRIGHT EXPCOST IN CASE OF OPENING BALANCE
                IF "Item Ledger Entry"."Entry Type" <> "Item Ledger Entry"."Entry Type"::"Positive Adjmt." THEN BEGIN
                    Valueentryrec.RESET;
                    Valueentryrec.SETRANGE(Valueentryrec."Document No.", finalinvno);
                    Valueentryrec.SETRANGE(Valueentryrec."Item No.", "Item Ledger Entry"."Item No."); // ADDED Expcost was wrong in case of two item in
                                                                                                      // one invoice.  20-01-2014
                                                                                                      //Valueentryrec.SETFILTER(Valueentryrec."Item Charge No.",'');//RSPLSID
                    ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '%1|%2', '', 'BO 04');
                    IF Valueentryrec.FINDFIRST THEN BEGIN
                        ExpCost := Valueentryrec."Cost per Unit";
                    END;
                END;


                CLEAR(bondedrate);
                Purchinvlinee.RESET;
                Purchinvlinee.SETRANGE(Purchinvlinee."Document No.", finalinvno);
                Purchinvlinee.SETRANGE(Purchinvlinee."No.", "Item Ledger Entry"."Item No.");
                IF Purchinvlinee.FINDFIRST THEN BEGIN
                    bondedrate := Purchinvlinee."Bonded Rate";
                    Exbondedrate := Purchinvlinee."Exbond Rate";
                END;
                /*//RSPLSUM 07Jul2020>>
                CLEAR(ShortageCharge);
                Valueentryrec.RESET;
                Valueentryrec.SETRANGE(Valueentryrec."Document No.",finalinvno);
                Valueentryrec.SETFILTER(Valueentryrec."Item Charge No.",'%1|%2|%3','AD 01','BO 01','CH 01');
                IF Valueentryrec.FINDFIRST THEN
                BEGIN
                 ShortageCharge := Valueentryrec."Cost per Unit";
                END;
                *///RSPLSUM 07Jul2020<<

                //>>11Dec2018

                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
                ValueEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
                ValueEntry.SETRANGE("Entry Type", ValueEntry."Entry Type"::Revaluation);
                ValueEntry.SETRANGE("Item Ledger Entry Type", ValueEntry."Item Ledger Entry Type"::Transfer);
                ValueEntry.SETRANGE("Journal Batch Name", '');
                IF ValueEntry.FINDFIRST THEN BEGIN
                    ValueEntry.CALCSUMS("Cost per Unit");
                    CustomCharge += ValueEntry."Cost per Unit";
                END;
                //Code Commented 04Jul2019
                AddRevCost := 0;
                ValueEntry.RESET;
                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
                ValueEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
                ValueEntry.SETRANGE("Entry Type", ValueEntry."Entry Type"::Revaluation);
                ValueEntry.SETRANGE("Item Ledger Entry Type", ValueEntry."Item Ledger Entry Type"::Purchase);
                IF ValueEntry.FINDFIRST THEN BEGIN
                    ValueEntry.CALCSUMS("Cost per Unit");
                    AddRevCost := ValueEntry."Cost per Unit";
                END;
                //<<11Dec2018


                CLEAR(TotalCharges);
                TotalCharges := ShortageCharge + TransportCharge + DiscountCost + CustomCharge + AnyOtherCharge;


                IF ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Output) OR
                          ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Positive Adjmt.") OR
                          ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Purchase) THEN BEGIN
                    purchrcpthed.RESET;
                    purchrcpthed.SETRANGE(purchrcpthed."No.", "Item Ledger Entry"."Document No.");
                    IF purchrcpthed.FINDFIRST THEN BEGIN
                        IF purchrcpthed."Location Code" = 'BOND0001' THEN
                            IMPORTLOCAL := 'IMPORT'
                        ELSE
                            IMPORTLOCAL := 'LOCAL';
                    END;
                END;

                IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer THEN BEGIN
                    purchrcpthed.RESET;
                    purchrcpthed.SETRANGE(purchrcpthed."No.", GRNTRans);
                    IF purchrcpthed.FINDFIRST THEN BEGIN
                        IF purchrcpthed."Location Code" = 'BOND0001' THEN
                            IMPORTLOCAL := 'IMPORT'
                        ELSE
                            IMPORTLOCAL := 'LOCAL';
                    END;
                END;

                /*
                
                // EBT MILAN 070114 To show invoice rate of positive adj. at vasai and trans. to daman--------------------START
                IF posadj = TRUE THEN
                BEGIN
                 ValueEntryreccc.RESET;
                 ValueEntryreccc.SETRANGE(ValueEntryreccc."Document No.",opbal);
                 ValueEntryreccc.SETRANGE(ValueEntryreccc."Item No.","Item Ledger Entry"."Item No.");
                 IF ValueEntryreccc.FINDFIRST THEN
                 BEGIN
                  ExpCost := ValueEntryreccc."Cost per Unit";
                 END;
                END;
                // EBT MILAN 070114 To show invoice rate of positive adj. at vasai and trans. to daman----------------------END
                *///Code Commented 11Jun2019

                // EBT MILAN (18122013)---------------ADDED NEW CALCULATION ON DISC.--------------------------------------------------END


                //RSPL007

                /*
                PurcRcptLine.RESET;
                PurcRcptLine.SETRANGE(PurcRcptLine."Document No.","Document No.");
                PurcRcptLine.SETRANGE(PurcRcptLine."No.","Item No.");
                IF PurcRcptLine.FINDSET THEN
                REPEAT
                  vLandedCostAmt+=PurcRcptLine."Landed Cost";
                UNTIL PurcRcptLine.NEXT=0; */

                /*
                IF "Document No." = 'GRN/15/1415/12/0283' THEN BEGIN
                MESSAGE("Document No.");
                MESSAGE("Item No.");
                END;
                */

                //RSPL007

                MakeExcelDataBody;

            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("Item Category Code", "Item No.", "Lot No.");//28Feb2017

                SETRANGE("Posting Date", 0D, PostingDate);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(PostingDate; PostingDate)
                {
                    ApplicationArea = all;
                    Caption = 'As On';
                }
                field(Showall; Showall)
                {
                    ApplicationArea = all;
                    Caption = 'Show Zero Rem. Qty';
                }
            }
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPostReport()
    begin
        CreateExcelbook;
    end;

    trigger OnPreReport()
    begin

        //>>09Mar2017 LocName
        LocCode := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code");
        recLOC.RESET;
        recLOC.SETRANGE(Code, LocCode);
        IF recLOC.FINDFIRST THEN BEGIN
            LocName := recLOC.Name;
        END;
        //<< LocName

        ExcelBuf.RESET;
        ExcelBuf.DELETEALL;

        MakeExcelInfo;
    end;

    var
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        ItemRec: Record 27;
        VendorRec: Record 23;
        ValueEntry: Record 5802;
        ChargeofItem: Decimal;
        ItemApplication: Record 339;
        RemQty: Decimal;
        PostingDate: Date;
        Showall: Boolean;
        VendorName: Text[60];
        ExpCost: Decimal;
        ActCost: Decimal;
        DiscountCost: Decimal;
        TRRNo: Code[20];
        TRRDate: Date;
        TransferRecipt: Record 5746;
        WHBENo: Code[20];
        DebondBENo: Code[20];
        ShortageCharge: Decimal;
        TransportCharge: Decimal;
        CustomCharge: Decimal;
        AnyOtherCharge: Decimal;
        RateofItem: Decimal;
        TotalCost: Decimal;
        TotalQty: Decimal;
        recTSH: Record 5744;
        BRTno: Code[20];
        BRTdate: Date;
        //recExcisePostingSetup: Record 13711;
        recTRL: Record 5747;
        ExciseAmount: Decimal;
        "Excise%": Decimal;
        recPHL: Record 121;
        recILE: Record 32;
        recTSL: Record 5745;
        recItemApplEntry: Record 339;
        recItemLedEntry: Record 32;
        recRPH: Record 120;
        VendorCode: Code[20];
        VendName: Text[50];
        VendorInvoiceno: Code[50];
        "OPENING/POSITIVE": Code[20];
        TotalCharges: Decimal;
        Pruchinvno: Code[20];
        Purchinvline: Record 123;
        GRNTRans: Code[20];
        Purchinvtrans: Code[20];
        vendcodtran: Code[20];
        Vendnametran: Text[100];
        vendinvno: Code[20];
        vendtrans: Record 23;
        Purchinvhd: Record 122;
        finalinvno: Code[20];
        Purchinvln: Record 123;
        accessvalue: Decimal;
        IMPORTLOCAL: Text[30];
        Purchrecpthead: Record 120;
        PurchinvExtrans: Code[50];
        vend: Boolean;
        purchrcpthed: Record 120;
        ADEAmt: Decimal;
        recILEnew: Record 32;
        TRansrec: Record 5746;
        recItemApplicEntry1: Record 339;
        recItemledgentry: Record 32;
        purchrecpthdr: Record 120;
        costcat5: Decimal;
        Transsrechd: Record 5746;
        ILErec: Record 32;
        IAErec: Record 339;
        ILErec1: Record 32;
        Purchrcpthhd: Record 120;
        posadj: Boolean;
        opbal: Code[20];
        ILErec2: Record 32;
        TRSHD: Record 5746;
        ITMAPPENT: Record 339;
        ILErec3: Record 32;
        PURCHRECHDR: Record 120;
        TRRECHD: Record 5746;
        ILErec4: Record 32;
        ITMAPPPLENTRY: Record 339;
        ILErec5: Record 32;
        Purchrecheaderr: Record 120;
        Valueentryrec: Record 5802;
        outpt: Code[20];
        outptbool: Boolean;
        Purchinvlinee: Record 123;
        bondedrate: Decimal;
        Exbondedrate: Decimal;
        ValueEntryreccc: Record 5802;
        itemappnotfind: Boolean;
        Itemapplientry: Record 339;
        PurchinvtransDate: Date;
        opbalDate: Date;
        outptDate: Date;
        PruchinvnoDate: Date;
        vLandedCostAmt: Decimal;
        PurcRcptLine: Record 121;
        recPurchInvLine: Record 123;
        "----09Mar2017": Integer;
        LocCode: Code[100];
        recLOC: Record 14;
        LocName: Text[250];
        LocName1: Text[100];
        "------31Jan2018": Integer;
        RRH31: Record 6660;
        SCMH31: Record 114;
        RepItmNo: Code[20];
        "---------05Sep2018": Integer;
        GRNFound05: Boolean;
        PRH05: Record 120;
        ILE05: Record 32;
        ItmApp05: Record 339;
        ItmApp05_01: Record 339;
        ItmApp05_02: Record 339;
        ItmApp05_03: Record 339;
        ItmApp05_04: Record 339;
        ItmApp05_05: Record 339;
        TempILEEntryNo05: Integer;
        GRNEntryNo: Integer;
        AddRevCost: Decimal;
        "------06May2019": Integer;
        NegativeCorrection: Boolean;
        ILE0505: Record 32;
        PCM: Decimal;
        Totalv: Decimal;
        ILE: Record 32;
        ILEEntryno: Integer;
        newValueEntry: Record 5802;
        PCI: Decimal;
        VendCellPrint: Boolean;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        MakeExcelDataHeader;
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin

        //>>23Feb2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
        //ExcelBuf.NewRow;
        //ExcelBuf.NewRow;

        //<<

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Raw Material Valuation report', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Inventory As On', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(PostingDate, FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(LocName, FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017
        //ExcelBuf.AddColumn("Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code"),FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn('IMPORT/LOCAL', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//5 //09Mar2017
        //ExcelBuf.AddColumn('Location Code',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn('Original Qty in Kg / Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//7
        ExcelBuf.AddColumn('Remaining Qty in Kg / Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//8
        ExcelBuf.AddColumn('Vendor Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//9
        ExcelBuf.AddColumn('Vendor', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//10
        ExcelBuf.AddColumn('Vendor Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//11
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//12
        ExcelBuf.AddColumn('Document Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//13
        ExcelBuf.AddColumn('Receipt Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//14
        ExcelBuf.AddColumn('Invoice Rate', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//15

        ExcelBuf.AddColumn('Shortage', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//16
        ExcelBuf.AddColumn('Transport', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//17
        ExcelBuf.AddColumn('Discount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//18
        ExcelBuf.AddColumn('Custom Duty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//19
        ExcelBuf.AddColumn('Any Other Charge', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//20
        ExcelBuf.AddColumn('Total Charge', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//21

        //ExcelBuf.AddColumn('Charges',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Total Rate / Ltr. Kg', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//22
        ExcelBuf.AddColumn('Inventory Value', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//23
        ExcelBuf.AddColumn('Ass. value', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//24
        ExcelBuf.AddColumn('Bonded Rate', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//25
        ExcelBuf.AddColumn('ExBond Rate', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//26
        ExcelBuf.AddColumn('Proportionate Excise Duty on Invoice', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//27
        ExcelBuf.AddColumn('TRR No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//28
        ExcelBuf.AddColumn('TRR Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//29
        ExcelBuf.AddColumn('BRT No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//30
        ExcelBuf.AddColumn('BRT Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//31
        ExcelBuf.AddColumn('Mater WH BE No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//32
        ExcelBuf.AddColumn('Mater WH BE Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//33
        ExcelBuf.AddColumn('Dbond BE No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//34
        ExcelBuf.AddColumn('Dbond BE Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//35
        //ExcelBuf.AddColumn('Ass. value',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Remarks', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//36
        //RSPL007
        ExcelBuf.AddColumn('Landed Cost', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//37
        //RSPL007
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Item Ledger Entry"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn(ItemRec.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn("Item Ledger Entry"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn(IMPORTLOCAL, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn(LocName1, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//5 //09Mar2017
        //ExcelBuf.AddColumn("Item Ledger Entry"."Location Code",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn("Item Ledger Entry"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn("Item Ledger Entry".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', ExcelBuf."Cell Type"::Number);//7
        ExcelBuf.AddColumn(RemQty, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', ExcelBuf."Cell Type"::Number);//8

        VendCellPrint := FALSE;//RSPLSUM 07Feb2020
        IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer THEN BEGIN
            IF (GRNTRans <> '') AND (vend = FALSE) THEN BEGIN
                ExcelBuf.AddColumn(vendcodtran, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//9
                ExcelBuf.AddColumn(Vendnametran, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
                ExcelBuf.AddColumn(PurchinvExtrans, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//11
                ExcelBuf.AddColumn(Purchinvtrans, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//12
                ExcelBuf.AddColumn(PurchinvtransDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);  //13
                ExcelBuf.AddColumn(GRNTRans, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
                VendCellPrint := TRUE;//RSPLSUM 07Feb2020
            END;
            IF (GRNTRans <> '') AND (vend = TRUE) THEN BEGIN
                ExcelBuf.AddColumn(VendorCode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//9
                ExcelBuf.AddColumn(VendName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
                ExcelBuf.AddColumn(VendorInvoiceno, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//11
                ExcelBuf.AddColumn(Purchinvtrans, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//12
                ExcelBuf.AddColumn(PurchinvtransDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);  //13
                ExcelBuf.AddColumn(GRNTRans, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
                VendCellPrint := TRUE;//RSPLSUM 07Feb2020
            END;
            IF posadj = TRUE THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//9
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//11
                ExcelBuf.AddColumn(opbal, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//12
                ExcelBuf.AddColumn(opbalDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//13  // EBT MILAN 260214
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
                "OPENING/POSITIVE" := 'OPENING/POSITIVE';
                VendCellPrint := TRUE;//RSPLSUM 07Feb2020
            END;
            IF outptbool = TRUE THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//9
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//11
                ExcelBuf.AddColumn(outpt, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//12
                ExcelBuf.AddColumn(outptDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//13  // EBT MILAN 260214
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
                VendCellPrint := TRUE;//RSPLSUM 07Feb2020
            END;

            IF itemappnotfind = TRUE THEN  // ADDED on 200114 Item application entry not finded in transfer case EBT MILAN
            BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//9
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//11
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//12
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//13  // EBT MILAN 260214
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
                VendCellPrint := TRUE;//RSPLSUM 07Feb2020
            END;

            itemappnotfind := FALSE;

            //RSPLSUM 07Feb2020>>
            IF NOT VendCellPrint THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//9
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//11
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//12
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//13  // EBT MILAN 260214
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
            END;
            //RSPLSUM 07Feb2020<<
        END ELSE BEGIN
            ExcelBuf.AddColumn("Item Ledger Entry"."Source No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//9
            ExcelBuf.AddColumn(VendorName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
            ExcelBuf.AddColumn("Item Ledger Entry"."External Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//11
            IF ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Output) OR
               ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Positive Adjmt.") THEN BEGIN
                ExcelBuf.AddColumn("Item Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//12
                ExcelBuf.AddColumn("Item Ledger Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//13
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
            END ELSE BEGIN

                //IF "Item Ledger Entry"."Document No."= 'GRN/26/1415/01/0012' THEN
                //ERROR('FOUND');
                vLandedCostAmt := 0;
                PurcRcptLine.RESET;
                PurcRcptLine.SETRANGE(PurcRcptLine."Document No.", "Item Ledger Entry"."Document No.");
                PurcRcptLine.SETRANGE(PurcRcptLine."No.", "Item Ledger Entry"."Item No.");
                IF PurcRcptLine.FIND('-') THEN BEGIN
                    recPurchInvLine.RESET;
                    // recPurchInvLine.SETRANGE(recPurchInvLine."Receipt Document No.", PurcRcptLine."Document No.");
                    recPurchInvLine.SETRANGE(recPurchInvLine."No.", "Item Ledger Entry"."Item No.");
                    IF recPurchInvLine.FINDSET THEN
                        REPEAT
                            vLandedCostAmt += recPurchInvLine."Landed Cost";
                        UNTIL recPurchInvLine.NEXT = 0;
                END;

                //>>RB-N 31Jan2018
                IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Sales Return Receipt" THEN BEGIN
                    CLEAR(Pruchinvno);
                    CLEAR(PruchinvnoDate);
                    RRH31.RESET;
                    RRH31.SETRANGE("No.", "Item Ledger Entry"."Document No.");
                    IF RRH31.FINDFIRST THEN BEGIN
                        SCMH31.RESET;
                        SCMH31.SETCURRENTKEY("Return Order No.");
                        SCMH31.SETRANGE("Return Order No.", RRH31."Return Order No.");
                        IF SCMH31.FINDFIRST THEN BEGIN
                            Pruchinvno := SCMH31."No.";
                            PruchinvnoDate := SCMH31."Posting Date";
                        END;
                    END;
                END;
                //<<RB-N 31Jan2018

                ExcelBuf.AddColumn(Pruchinvno, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //12
                ExcelBuf.AddColumn(PruchinvnoDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//13  // EBT MILAN 260214
                ExcelBuf.AddColumn("Item Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
            END;
        END;

        //>>RB-N 12Jul2018
        IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Sales Return Receipt" THEN BEGIN
            CLEAR(ExpCost);
            ValueEntry.RESET;
            ValueEntry.SETCURRENTKEY("Document No.");
            ValueEntry.SETRANGE("Document No.", Pruchinvno);
            ValueEntry.SETRANGE("Document Line No.", "Item Ledger Entry"."Document Line No.");
            ValueEntry.SETRANGE("Item No.", "Item Ledger Entry"."Item No.");
            ValueEntry.SETRANGE("Document Type", ValueEntry."Document Type"::"Sales Credit Memo");
            ValueEntry.SETRANGE(Adjustment, FALSE);
            IF ValueEntry.FINDFIRST THEN BEGIN
                ExpCost := ValueEntry."Cost per Unit";
            END;

        END;
        //>>RB-N 12Jul2018

        //>>11Jun2019
        IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Positive Adjmt." THEN BEGIN
            CLEAR(ExpCost);
            ValueEntry.RESET;
            ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
            ValueEntry.SETRANGE("Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
            IF ValueEntry.FINDFIRST THEN BEGIN
                ValueEntry.CALCSUMS("Cost per Unit");
                ExpCost := ValueEntry."Cost per Unit";
            END;
        END;
        //<<11Jun2019

        //>>11Dec2018
        IF AddRevCost <> 0 THEN
            ExpCost += AddRevCost;
        //<<11Dec2018

        IF ("Item Ledger Entry"."Item Category Code" = 'CAT10') OR ("Item Ledger Entry"."Item Category Code" = 'CAT09') THEN
            ExcelBuf.AddColumn(ROUND((RateofItem), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number)//15
        ELSE
            ExcelBuf.AddColumn(ROUND((ExpCost - ABS(PCM)), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//15//RSPL AD

        ExcelBuf.AddColumn(ShortageCharge, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//16
        ExcelBuf.AddColumn(TransportCharge, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//17
        ExcelBuf.AddColumn(DiscountCost, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//18
        ExcelBuf.AddColumn(CustomCharge, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//19
        ExcelBuf.AddColumn(AnyOtherCharge, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//20
        ExcelBuf.AddColumn(TotalCharges, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//21

        //ExcelBuf.AddColumn(ROUND(DiscountCost,0.01),FALSE,'',FALSE,FALSE,FALSE,'');
        IF ("Item Ledger Entry"."Item Category Code" = 'CAT10') OR ("Item Ledger Entry"."Item Category Code" = 'CAT09') THEN BEGIN
            ExcelBuf.AddColumn((TotalCharges + RateofItem), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//22
            ExcelBuf.AddColumn((TotalCharges + RateofItem) * RemQty, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//23
        END ELSE BEGIN
            ExcelBuf.AddColumn(TotalCharges + ExpCost, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//22
            ExcelBuf.AddColumn((TotalCharges + ExpCost) * RemQty, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//23
        END;

        ExcelBuf.AddColumn((accessvalue * RemQty), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//24
        ExcelBuf.AddColumn(bondedrate, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//25
        ExcelBuf.AddColumn(Exbondedrate, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//26
        //ExcelBuf.AddColumn(ExciseAmount,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(((accessvalue * RemQty) * "Excise%" / 100), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//27
        ExcelBuf.AddColumn(TRRNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//28
        ExcelBuf.AddColumn(TRRDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//29
        ExcelBuf.AddColumn(BRTno, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//30
        ExcelBuf.AddColumn(BRTdate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//31
        ExcelBuf.AddColumn(WHBENo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//32
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//33
        ExcelBuf.AddColumn(DebondBENo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//34
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//35
        //>>05May2019
        IF NegativeCorrection THEN
            "OPENING/POSITIVE" := 'Negative Correction';
        //<<05May2019
        ExcelBuf.AddColumn("OPENING/POSITIVE", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//36

        ExcelBuf.AddColumn(vLandedCostAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//37
        IMPORTLOCAL := '';
        itemappnotfind := FALSE;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet('Data','Raw Material Inventory Valuation Report',COMPANYNAME,USERID);

        //ExcelBuf.CreateBookAndOpenExcel('','Raw Material Inventory Valuation Report','','','');
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>28Feb2017 RB-N
        /* ExcelBuf.CreateBook('', 'Raw Material Inventory Valuation Report');
        ExcelBuf.CreateBookAndOpenExcel('', 'Raw Material Inventory Valuation Report', '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook('Raw Material Inventory Valuation Report');
        ExcelBuf.WriteSheet('Raw Material Inventory Valuation Report', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Raw Material Inventory Valuation Report', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<
    end;
}


report 50093 "Purchase Register GST"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 11Sep2017   RB-N         Filtering of Report on the basis of Location GSTNo.
    // 20Sep2017   RB-N         PrintSummary Option & GST %
    // 15Nov2017   RB-N         Receiving State, Dealer Category
    // 07Jul2018   RB-N         GST Group Code, GST Category & -ve Value for Credit Memo.
    // 19Oct2018   RB-N         PO Details for ChargeItemLine & Bill of Entry Details
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseRegisterGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Purchase Register GST';

    dataset
    {
        dataitem("Purch. Inv. Header"; "Purch. Inv. Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "Posting Date", "Location Code", "Vendor Posting Group", "Gen. Bus. Posting Group";
            dataitem("Purch. Inv. Line"; "Purch. Inv. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                RequestFilterFields = "Gen. Prod. Posting Group", Type, "No.", "Item Category Code";
                DataItemTableView = SORTING("Document No.", "Line No.")
                                     ORDER(Ascending)
                                     WHERE(Quantity = FILTER(<> 0),
                                           "No." = FILTER(<> 74012210));

                trigger OnAfterGetRecord()
                begin
                    //IF  (PrinttoExcel) AND  (PurchaseInvoice) THEN //05May2017
                    //BEGIN  //05May2017

                    //>>
                    CLEAR(MRNo);
                    PRL.RESET;
                    //  PRL.SETRANGE("Document No.", "Purch. Inv. Line"."Receipt Document No.");
                    //  PRL.SETRANGE("Line No.", "Purch. Inv. Line"."Receipt Document Line No.");
                    IF PRL.FINDFIRST THEN BEGIN
                        MRNo := PRL."MR No";
                    END;
                    //<<
                    CLEAR(BEDpercent);
                    /* recExciseEntry.RESET;
                    recExciseEntry.SETRANGE(recExciseEntry."Document No.", "Purch. Inv. Line"."Document No.");
                    IF recExciseEntry.FINDFIRST THEN BEGIN
                        BEDpercent := recExciseEntry."BED %";
                    END; */

                    CLEAR(recItemCategory);
                    recItemCategory.SETRANGE(recItemCategory.Code, "Purch. Inv. Line"."Item Category Code");
                    IF recItemCategory.FINDFIRST THEN;
                    //RSPL007
                    vLandedCost += "Landed Cost";
                    //RSPL007
                    CLEAR(POno);
                    CLEAR(POdate);
                    CLEAR(BlanketOrderNo);
                    CLEAR(BlanketOrderDate);
                    CLEAR(GRNdate);
                    recPRH.RESET;
                    // recPRH.SETRANGE(recPRH."No.", "Purch. Inv. Line"."Receipt Document No.");
                    IF recPRH.FINDFIRST THEN BEGIN
                        POno := recPRH."Order No.";
                        GRNdate := recPRH."Posting Date";
                        BlanketOrderNo := recPRH."Blanket Order No.";
                        BlanketOrderDate := recPRH."Order Date";
                    END;

                    recPH.RESET;
                    recPH.SETRANGE(recPH."Document Type", recPH."Document Type"::Order);
                    recPH.SETRANGE(recPH."No.", POno);
                    IF recPH.FINDFIRST THEN BEGIN
                        POdate := recPH."Posting Date";
                    END;


                    IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)" THEN BEGIN
                        /* PostedStrOrdrLineDetails.RESET;
                        PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Purch. Inv. Header"."No.");
                        PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Item No.", "Purch. Inv. Line"."No.");
                        PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                        PostedStrOrdrLineDetails.SETFILTER(PostedStrOrdrLineDetails."Tax/Charge Group", '%1|%2', 'EXCISE', 'SHORTAGE');
                        IF PostedStrOrdrLineDetails.FINDFIRST THEN
                            REPEAT
                                AddExcise += AddExcise + PostedStrOrdrLineDetails."Calculation Value";
                            UNTIL PostedStrOrdrLineDetails.NEXT = 0; */
                        //>>19Oct2018
                        CLEAR(ILEDocNo);
                        VE19.RESET;
                        VE19.SETCURRENTKEY("Document No.");
                        VE19.SETRANGE("Document No.", "Document No.");
                        VE19.SETRANGE("Item Charge No.", "No.");
                        IF VE19.FINDFIRST THEN BEGIN
                            ILE19.RESET;
                            IF ILE19.GET(VE19."Item Ledger Entry No.") THEN
                                ILEDocNo := ILE19."Document No.";
                        END;

                        recPRH.RESET;
                        recPRH.SETRANGE(recPRH."No.", ILEDocNo);
                        IF recPRH.FINDFIRST THEN BEGIN
                            POno := recPRH."Order No.";
                            GRNdate := recPRH."Posting Date";
                            BlanketOrderNo := recPRH."Blanket Order No.";
                            BlanketOrderDate := recPRH."Order Date";
                        END;

                        recPH.RESET;
                        recPH.SETRANGE(recPH."Document Type", recPH."Document Type"::Order);
                        recPH.SETRANGE(recPH."No.", POno);
                        IF recPH.FINDFIRST THEN BEGIN
                            POdate := recPH."Posting Date";
                        END;
                        //<<19Oct2018
                    END;

                    CLEAR(DisChgs);
                    /*  recPostedStrOrdDetails.RESET;
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Inv. Header"."No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Inv. Line"."Line No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'DISCOUNT');//03Aug2017
                                                                                                            //recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code",'DISCOUNT');
                     IF recPostedStrOrdDetails.FINDFIRST THEN
                         REPEAT
                             DisChgs := DisChgs + recPostedStrOrdDetails.Amount
                         UNTIL recPostedStrOrdDetails.NEXT = 0; */

                    //>>03Aug2017 FreightAmt
                    CLEAR(freightamt);
                    /* recPostedStrOrdDetails.RESET;
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Inv. Header"."No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Inv. Line"."Line No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'FREIGHT');
                    IF recPostedStrOrdDetails.FINDFIRST THEN
                        REPEAT
                            freightamt := freightamt + recPostedStrOrdDetails.Amount
                        UNTIL recPostedStrOrdDetails.NEXT = 0; */
                    //>>03Aug2017 FreightAmt

                    //DJ 17042021
                    CLEAR(TCSPCharges);
                    /*  recPostedStrOrdDetails.RESET;
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Inv. Header"."No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Inv. Line"."Line No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'TCSP');//03Aug2017
                     IF recPostedStrOrdDetails.FINDFIRST THEN
                         REPEAT
                             TCSPCharges := TCSPCharges + recPostedStrOrdDetails."Amount (LCY)";
                         UNTIL recPostedStrOrdDetails.NEXT = 0; */
                    LineTCSPCharges += TCSPCharges;
                    //DJ 17042021

                    IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)" THEN
                        ExcessShortQty := ExcessShortQty + "Purch. Inv. Line".Quantity;//20Sep2017
                                                                                       //ExcessShortQty += ExcessShortQty + "Purch. Inv. Line".Quantity;



                    //>>RSPl-Rahul**Migration**test3
                    vLineAmt += "Purch. Inv. Line"."Line Amount";
                    vExciseBaseAmt += 0;//"Purch. Inv. Line"."Excise Base Amount";
                    vBEDAmt += 0;// "Purch. Inv. Line"."BED Amount";
                    vEcessAmt += 0;//"Purch. Inv. Line"."eCess Amount";
                    vSheCessAmt += 0;//"Purch. Inv. Line"."SHE Cess Amount";
                    vADEAmt += 0;// "Purch. Inv. Line"."ADE Amount";
                    //vTaxBaseAmt+="Purch. Inv. Line"."Tax Base Amount";
                    //>>09Mar2017
                    IF ("Purch. Inv. Line".Type = "Purch. Inv. Line".Type::Item) THEN
                        vTaxBaseAmt += 0;//"Purch. Inv. Line"."Tax Base Amount";
                    //<<

                    vTaxAmt += 0;//"Purch. Inv. Line"."Tax Amount";
                    vTDSAmt += 0;// "Purch. Inv. Line"."TDS Amount";
                    //<<





                    CLEAR(TaxPer);
                    CLEAR(TaxGroupCode);
                    /*  RecDTE.RESET;
                     RecDTE.SETRANGE(RecDTE."Transaction Type", RecDTE."Transaction Type"::Purchase);
                     RecDTE.SETRANGE(RecDTE."Document Type", RecDTE."Document Type"::Invoice);
                     RecDTE.SETRANGE("Document No.", "Document No.");
                     //RecDTE.SETRANGE(Type,Type);  RSPL
                     //RecDTE.SETRANGE("No.","No.");RSPL
                     RecDTE.SETRANGE("Document Line No.", "Line No.");
                     IF RecDTE.FINDFIRST THEN BEGIN
                         IF RecTG.GET(RecDTE."Tax Group Code") THEN;
                         IF RecDTE."Form Code" <> '' THEN
                             TaxGroupCode := RecTG.Description + 'WITH FORM ' + RecDTE."Form Code"
                         ELSE
                             TaxGroupCode := RecTG.Description;
                         TaxPer := FORMAT(RecDTE."Tax %") + '%';

                     END; */
                    // IF (Vendor."Vendor Posting Group" = 'FOREIGN') AND ("Purch. Inv. Header"."Form Code" = '') THEN BEGIN
                    TaxGroupCode := 'NoTax Ag.Export';
                    // END;


                    //RSPL-Dhananjay START

                    CLEAR(Vtext);
                    IF RecLocation.GET("Purch. Inv. Line"."Location Code") THEN BEGIN
                        LocationState := RecLocation."State Code"
                    END ELSE
                        IF rECResposiblityCenter.GET("Purch. Inv. Line"."Responsibility Center") THEN
                            LocationState := rECResposiblityCenter.State;

                    IF RecVendor.GET("Purch. Inv. Line"."Buy-from Vendor No.") THEN BEGIN
                        IF RecVendor."Gen. Bus. Posting Group" = 'DOMESTIC' THEN BEGIN
                            recPurchInv.RESET;
                            recPurchInv.SETRANGE(recPurchInv."No.", "Purch. Inv. Line"."Document No.");
                            IF recPurchInv.FINDSET THEN
                                IF recPurchInv."Order Address Code" <> '' THEN BEGIN
                                    orderadd.RESET;
                                    orderadd.SETRANGE("Vendor No.", "Buy-from Vendor No.");//24May2017
                                    orderadd.SETRANGE(orderadd.Code, recPurchInv."Order Address Code");
                                    IF orderadd.FINDFIRST THEN //24May2017
                                                               //IF orderadd.FINDSET THEN
                                        IF orderadd.State = LocationState THEN
                                            Vtext := 'INTRA STATE' //03Aug2017
                                                                   //Vtext := 'LOCAL'
                                        ELSE
                                            Vtext := 'INTER STATE';
                                END ELSE
                                    IF recPurchInv."Order Address Code" = '' THEN BEGIN
                                        IF RecVendor."State Code" = LocationState THEN
                                            Vtext := 'INTRA STATE' //03Aug2017
                                                                   //Vtext := 'LOCAL'
                                        ELSE
                                            Vtext := 'INTER STATE';
                                    END;
                        END;
                    END ELSE
                        IF RecVendor."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            Vtext := 'IMPORT';

                    //>>02Aug2017 GST
                    CLEAR(vCGST);
                    CLEAR(vSGST);
                    CLEAR(vIGST);
                    CLEAR(vUTGST);
                    CLEAR(vCGSTRate);
                    CLEAR(vSGSTRate);
                    CLEAR(vIGSTRate);
                    CLEAR(vUTGSTRate);
                    CLEAR(GSTPer);//RB-N 20Sep2017

                    DetailGSTEntry.RESET;
                    DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                    DetailGSTEntry.SETRANGE("Entry Type", DetailGSTEntry."Entry Type"::"Initial Entry");//29Mar2019
                    //DetailGSTEntry.SETRANGE(Type,Type); //05Aug2017
                    //DetailGSTEntry.SETRANGE(DetailGSTEntry."No.","No."); //05Aug2017
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT
                            GSTPer += DetailGSTEntry."GST %";//RB-N 20Sep2017
                            IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                vCGST += DetailGSTEntry."GST Amount";
                                vCGSTRate := DetailGSTEntry."GST %";
                            END;

                            IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                vSGST += DetailGSTEntry."GST Amount";
                                vSGSTRate := DetailGSTEntry."GST %";

                            END;

                            IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                vUTGST += DetailGSTEntry."GST Amount";
                                vUTGSTRate := DetailGSTEntry."GST %";

                            END;


                            IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                vIGST += DetailGSTEntry."GST Amount";
                                vIGSTRate := DetailGSTEntry."GST %";
                            END;

                        UNTIL DetailGSTEntry.NEXT = 0;

                    TCGST += vCGST;
                    TTCGST += vCGST;

                    TSGST += vSGST;
                    TTSGST += vSGST;

                    TIGST += vIGST;
                    TTIGST += vIGST;

                    TUTGST += vUTGST;
                    TTUTGST += vUTGST;
                    //<<02Aug2017 GST

                    //RSPLAM29072021
                    ExchangeRate := 0;
                    IF "Purch. Inv. Header"."Currency Code" <> '' THEN
                        ExchangeRate := 1 / "Purch. Inv. Header"."Currency Factor"
                    ELSE
                        ExchangeRate := 1;
                    //RSPLAM29072021

                    IF PrinttoExcel AND Detail AND PurchaseInvoice THEN BEGIN
                        MakePurchInvDataBody;
                    END;
                    //END; //05May2017

                    //>>20Sep2017
                    IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::Item THEN BEGIN
                        IF "Purch. Inv. Line"."Item Category Code" = 'CAT22' THEN
                            TotalQtyofITEM += "Purch. Inv. Line"."Quantity (Base)"//09Mar2017
                        ELSE
                            TotalQtyofITEM += "Purch. Inv. Line".Quantity;//09Mar2017
                    END;

                    IF "Purch. Inv. Header"."Currency Factor" = 0 THEN BEGIN
                        GSTBaseAmt += 0;// "Purch. Inv. Line"."GST Base Amount";//02Aug2017
                        TGSTBaseAmt += 0;//"Purch. Inv. Line"."GST Base Amount";//02Aug2017
                        GTaxBaseAmt += 0;//"Purch. Inv. Line"."Tax Base Amount";
                        TGTaxBaseAmt += 0;// "Purch. Inv. Line"."Tax Base Amount";

                    END;


                    IF "Purch. Inv. Header"."Currency Factor" <> 0 THEN BEGIN
                        GSTBaseAmt += /*"Purch. Inv. Line"."GST Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor";//02Aug2017
                        TGSTBaseAmt += /*"Purch. Inv. Line"."GST Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor";//02Aug2017
                        GTaxBaseAmt += /*"Purch. Inv. Line"."Tax Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor";
                        TGTaxBaseAmt += /*"Purch. Inv. Line"."Tax Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor";
                    END;
                    //<<20Sep2017

                    //>>09Mar2017
                    IF "Purch. Inv. Header"."Currency Code" = '' THEN BEGIN
                        TotalAmount1 += "Purch. Inv. Line"."Line Amount";
                        TotalBillAmtInclusivEcx += BillAmtInclusivEcx;
                        //TotalGrossValue +=  "Purch. Inv. Line"."Amount To Vendor";
                        TotalNetValue += 0;// "Purch. Inv. Line".Amount;
                        TotalBedAmt += 0;//"Purch. Inv. Line"."BED Amount";
                        TotaleCessAmt += 0;// "Purch. Inv. Line"."eCess Amount";
                        TotalSHECess += 0;// "Purch. Inv. Line"."SHE Cess Amount";
                        TotalTaxBaseAmt += "Purch. Inv. Line".Amount + 0;//"Purch. Inv. Line"."Charges To Vendor";
                        //TotalTaxAmt     +=  "Tax Amount";//Sourav  110215
                        //totadditionalduty +=  "Purch. Inv. Line"."ADE Amount";
                        totalfreight += freightamt;
                        totDisChgs += DisChgs;
                        TotalTCSPCharges += TCSPCharges;//DJ 170421

                        TotalDeliveryChrgs += DeliveryChrgs;
                        TotalTDSAmount += 0;// "Purch. Inv. Line"."TDS Amount";
                    END
                    ELSE BEGIN
                        TotalAmount1 += "Purch. Inv. Line"."Line Amount" / "Purch. Inv. Header"."Currency Factor";
                        TotalBillAmtInclusivEcx += BillAmtInclusivEcx / "Purch. Inv. Header"."Currency Factor";
                        TotalGrossValue += "Purch. Inv. Line"."Amount To Vendor" / "Purch. Inv. Header"."Currency Factor";
                        TotalNetValue += "Purch. Inv. Line".Amount / "Purch. Inv. Header"."Currency Factor";
                        TotalBedAmt += /*"Purch. Inv. Line"."BED Amount"*/0 / "Purch. Inv. Header"."Currency Factor";
                        TotaleCessAmt += /*"Purch. Inv. Line"."eCess Amount"*/0 / "Purch. Inv. Header"."Currency Factor";
                        TotalSHECess += /*"Purch. Inv. Line"."SHE Cess Amount"*/0 / "Purch. Inv. Header"."Currency Factor";
                        TotalTaxBaseAmt += "Purch. Inv. Line".Amount +/*"Purch. Inv. Line"."Charges To Vendor"*/ 0 / "Purch. Inv. Header"."Currency Factor";
                        TotalTaxAmt += /*"Tax Amount"*/0 / "Purch. Inv. Header"."Currency Factor";
                        //totadditionalduty +=  "Purch. Inv. Line"."ADE Amount"/"Purch. Inv. Header"."Currency Factor";
                        totalfreight += freightamt / "Purch. Inv. Header"."Currency Factor";
                        totDisChgs += DisChgs / "Purch. Inv. Header"."Currency Factor";
                        TotalAddExcise += AddExcise / "Purch. Inv. Header"."Currency Factor";
                        TotalTaxableAmount += /*"Purch. Inv. Line"."Tax Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor";
                        TotalDeliveryChrgs += DeliveryChrgs / "Purch. Inv. Header"."Currency Factor";
                        TotalTDSAmount += /*"Purch. Inv. Line"."TDS Amount"*/ 0 / "Purch. Inv. Header"."Currency Factor";
                    END;



                    //<<

                    //**Migration

                    Location.SETRANGE(Location.Code, "Purch. Inv. Line"."Location Code");
                    IF Location.FINDFIRST THEN
                        "Rec Addss" := Location.Name + ',' + Location.City;
                end;

                trigger OnPostDataItem()
                begin


                    IF PINVLINE <> 0 THEN
                        //IF TotalQtyofITEM <> 0 THEN //02May2017
                        IF PrinttoExcel AND PurchaseInvoice THEN BEGIN
                            MakePurchInvTotalDataGrouping;
                        END;

                    /*
                    IF "Purch. Inv. Header"."Currency Code" = '' THEN
                    BEGIN
                    TotalAmount1 += "Purch. Inv. Line"."Line Amount";
                    TotalBillAmtInclusivEcx +=  BillAmtInclusivEcx;
                    //TotalGrossValue +=  "Purch. Inv. Line"."Amount To Vendor";
                    TotalNetValue +=  "Purch. Inv. Line".Amount;
                    TotalBedAmt   +=  "Purch. Inv. Line"."BED Amount";
                    TotaleCessAmt +=  "Purch. Inv. Line"."eCess Amount";
                    TotalSHECess  +=  "Purch. Inv. Line"."SHE Cess Amount";
                    TotalTaxBaseAmt +=  "Purch. Inv. Line".Amount+"Purch. Inv. Line"."Charges To Vendor";
                    //TotalTaxAmt     +=  "Tax Amount";//Sourav  110215
                    //totadditionalduty +=  "Purch. Inv. Line"."ADE Amount";
                    totalfreight      += freightamt;
                    totDisChgs  +=  DisChgs;
                    
                    
                    TotalDeliveryChrgs  +=  DeliveryChrgs;
                    TotalTDSAmount += "Purch. Inv. Line"."TDS Amount";
                    END
                    ELSE
                    BEGIN
                    TotalAmount1 += "Purch. Inv. Line"."Line Amount"/"Purch. Inv. Header"."Currency Factor";
                    TotalBillAmtInclusivEcx +=  BillAmtInclusivEcx/"Purch. Inv. Header"."Currency Factor";
                    TotalGrossValue +=  "Purch. Inv. Line"."Amount To Vendor"/"Purch. Inv. Header"."Currency Factor";
                    TotalNetValue +=  "Purch. Inv. Line".Amount/"Purch. Inv. Header"."Currency Factor";
                    TotalBedAmt   +=  "Purch. Inv. Line"."BED Amount"/"Purch. Inv. Header"."Currency Factor";
                    TotaleCessAmt +=  "Purch. Inv. Line"."eCess Amount"/"Purch. Inv. Header"."Currency Factor";
                    TotalSHECess  +=  "Purch. Inv. Line"."SHE Cess Amount"/"Purch. Inv. Header"."Currency Factor";
                    TotalTaxBaseAmt +=  "Purch. Inv. Line".Amount+"Purch. Inv. Line"."Charges To Vendor"/"Purch. Inv. Header"."Currency Factor";
                    TotalTaxAmt     +=  "Tax Amount"/"Purch. Inv. Header"."Currency Factor";
                    //totadditionalduty +=  "Purch. Inv. Line"."ADE Amount"/"Purch. Inv. Header"."Currency Factor";
                    totalfreight      += freightamt/"Purch. Inv. Header"."Currency Factor";
                    totDisChgs  +=  DisChgs/"Purch. Inv. Header"."Currency Factor";
                    TotalAddExcise  +=  AddExcise/"Purch. Inv. Header"."Currency Factor";
                    TotalTaxableAmount  += "Purch. Inv. Line"."Tax Base Amount"/"Purch. Inv. Header"."Currency Factor";
                    TotalDeliveryChrgs  +=  DeliveryChrgs/"Purch. Inv. Header"."Currency Factor";
                    TotalTDSAmount += "Purch. Inv. Line"."TDS Amount"/"Purch. Inv. Header"."Currency Factor";
                    END;
                    *///Commented 09Mar2017

                end;

                trigger OnPreDataItem()
                begin

                    //>>02May2017

                    //IF "Purch. Inv. Line"."Gen. Prod. Posting Group" <> '' THEN
                    //SETFILTER("Gen. Prod. Posting Group","Purch. Inv. Line"."Gen. Prod. Posting Group");

                    //<<02May2017

                    PINVLINE := COUNT;//05May2017
                    LineTCSPCharges := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>03Aug2017
                CLEAR(BillStateCode);
                CLEAR(BillStateName);
                CLEAR(ShipStateCode);
                CLEAR(ShipStateName);
                CLEAR(VenGSTNo);

                Ven03.RESET;
                IF Ven03.GET("Buy-from Vendor No.") THEN BEGIN
                    VenGSTNo := Ven03."GST Registration No.";
                    State03.RESET;
                    State03.SETRANGE(Code, Ven03."State Code");
                    IF State03.FINDFIRST THEN BEGIN
                        BillStateCode := State03."State Code (GST Reg. No.)";
                        BillStateName := State03.Description;
                    END;
                END;

                IF "Order Address Code" = '' THEN BEGIN
                    ShipStateCode := BillStateCode;
                    ShipStateName := BillStateName;
                END;

                CLEAR(ShipGSTNo); //11Aug2017
                IF "Order Address Code" <> '' THEN BEGIN
                    Ord03.RESET;
                    Ord03.SETRANGE("Vendor No.", "Buy-from Vendor No.");
                    Ord03.SETRANGE(Code, "Order Address Code");
                    IF Ord03.FINDFIRST THEN BEGIN
                        ShipGSTNo := Ord03."GST Registration No.";//11Aug2017
                                                                  //MESSAGE('%1 \ %2 \ %3 ',"Buy-from Vendor No.","Order Address Code",Ord03.State);
                        State03.RESET;
                        State03.SETRANGE(Code, Ord03.State);
                        IF State03.FINDFIRST THEN BEGIN
                            ShipStateCode := State03."State Code (GST Reg. No.)";
                            ShipStateName := State03.Description;

                        END;
                    END;
                END;

                CLEAR(LocStateCode);
                CLEAR(LocGSTNo);
                CLEAR(LocName15);
                CLEAR(LocState15);

                Loc03.RESET;//15Nov2017
                IF Loc03.GET("Location Code") THEN BEGIN
                    LocGSTNo := Loc03."GST Registration No.";
                    LocName15 := Loc03.Name;//15Nov2017

                    State03.RESET;
                    State03.SETRANGE(Code, Loc03."State Code");
                    IF State03.FINDFIRST THEN BEGIN
                        LocStateCode := State03."State Code (GST Reg. No.)";
                        LocState15 := State03.Description;//15Nov2017

                    END;
                END;

                //<<03Aug2017

                //>>RB-N 11Sep2017 GSTNo. Filter

                IF TempGSTNo <> '' THEN BEGIN
                    IF TempGSTNo <> LocGSTNo THEN
                        CurrReport.SKIP;

                END;
                //<<RB-N 11Sep2017 GSTNo. Filter


                AddTax := 0;
                additnlamt := 0;
                freightamt := 0;
                BillAmtInclusivEcx := 0;
                DeliveryChrgs := 0;
                Amount1 := 0;
                TaxAmt := 0;
                GrossAmt := 0;
                DisChgs := 0;
                vLandedCost := 0;//RSPL007
                BillNoLen := STRLEN("Purch. Inv. Header"."No.");
                "Bill No." := "Purch. Inv. Header"."No.";
                IF "Purch. Inv. Header"."Ship-to Code" <> '' THEN BEGIN
                    "Rec Addss" := Location.Code + Location.City;
                END ELSE BEGIN
                    "Rec Addss" := 'Same As Above';
                END;

                IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
                    OrdAddCity := '';
                    OrdAddTIN := '';
                    OrdAddCST := '';
                    recOrderAddress.RESET;
                    recOrderAddress.SETRANGE(recOrderAddress."Vendor No.", "Purch. Inv. Header"."Buy-from Vendor No.");
                    recOrderAddress.SETRANGE(recOrderAddress.Code, "Purch. Inv. Header"."Order Address Code");
                    IF recOrderAddress.FINDFIRST THEN BEGIN
                        OrdAddCity := recOrderAddress.City;
                        OrdAddTIN := recOrderAddress."T.I.N. No.";
                        OrdAddCST := recOrderAddress."C.S.T. No.";
                    END;
                END;

                /*  PostedStrDetail.RESET;
                 PostedStrDetail.SETRANGE(PostedStrDetail.Type, PostedStrDetail.Type::Purchase);
                 PostedStrDetail.SETRANGE(PostedStrDetail."Document Type", PostedStrDetail."Document Type"::Invoice);
                 PostedStrDetail.SETRANGE(PostedStrDetail."No.", "Purch. Inv. Header"."No.");
                 IF PostedStrDetail.FIND('-') THEN
                     REPEAT
                     BEGIN
                         IF (PostedStrDetail."Tax/Charge Group" = 'EXCISE') THEN BEGIN
                             exciseamt := PostedStrDetail."Calculation Value";
                         END;
                         IF (PostedStrDetail."Tax/Charge Group" = 'FREIGHT') THEN BEGIN
                             freightamt := PostedStrDetail."Calculation Value";
                         END;
                         IF (PostedStrDetail."Tax/Charge Group" = 'ADD TAX') THEN BEGIN
                             additnlamt := PostedStrDetail."Calculation Value";
                         END;
                         IF (PostedStrDetail."Tax/Charge Group" = 'DISCOUNT') THEN BEGIN
                             DisChgs := PostedStrDetail."Calculation Value";
                         END;

                     END;
                     UNTIL PostedStrDetail.NEXT = 0; */

                Vendor.RESET;
                IF Vendor.GET("Purch. Inv. Header"."Buy-from Vendor No.") THEN;
                PurchOrdeDate := "Purch. Inv. Header"."Posting Date";

                RespCtr.SETRANGE(RespCtr.Code, "Purch. Inv. Header"."Responsibility Center");
                IF RespCtr.FINDFIRST THEN;
                Vendor.RESET;
                IF Vendor.GET("Purch. Inv. Header"."Buy-from Vendor No.") THEN;

                CLEAR(InvTaxDescription);
                FormCode := '';
                /* DetTaxEntry.RESET;
                DetTaxEntry.SETRANGE(DetTaxEntry."Document No.", "Purch. Inv. Line"."No.");
                IF DetTaxEntry.FINDFIRST THEN BEGIN
                    IF DetTaxEntry."Form Code" <> '' THEN BEGIN
                        FormCode := 'with Form ' + DetTaxEntry."Form Code";
                    END
                    ELSE
                        FormCode := '';
                    InvTaxDescription := FORMAT(DetTaxEntry."Tax Type") + FORMAT(DetTaxEntry."Tax %") + '%' + FormCode;
                END ELSE BEGIN
                    InvTaxDescription := '';
                END;
                IF (Vendor."Vendor Posting Group" = 'FOREIGN') AND ("Purch. Inv. Header"."Form Code" = '') THEN BEGIN
                    InvTaxDescription := 'NoTax Ag.Export';
                END; */


                /* PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Purch. Inv. Header"."No.");
                PostedStrOrdrLineDetails.SETFILTER(PostedStrOrdrLineDetails."Tax/Charge Group", '%1', 'ADDL.DUTY');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                    REPEAT
                        AddTax := AddTax + PostedStrOrdrLineDetails.Amount;
                    UNTIL PostedStrOrdrLineDetails.NEXT = 0; */

                //RSPL-290115
                tComment := '';
                CommentLine.RESET;
                CommentLine.SETRANGE(CommentLine."Document Type", CommentLine."Document Type"::"Posted Invoice");
                CommentLine.SETRANGE(CommentLine."No.", "Purch. Inv. Header"."No.");
                CommentLine.SETRANGE(CommentLine."Document Line No.", 0);
                IF CommentLine.FINDFIRST THEN
                    tComment := CommentLine.Comment;
                //RSPL-290115
            end;

            trigger OnPostDataItem()
            begin
                IF PrinttoExcel AND PurchaseInvoice THEN BEGIN
                    MakePurchInvGrandTotal;
                END;
            end;

            trigger OnPreDataItem()
            begin

                CompInfo.GET;

                CurrReport.CREATETOTALS(additnlamt, TotalTaxBaseAmt, TaxableAmount, totDisChgs, exciseamt, AddTax, BillAmtInclusivEcx);
                CurrReport.CREATETOTALS(TAXAmountvalue);
            end;
        }
        dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Pre-Assigned No." = FILTER(<> ''));
            dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0),
                                          "No." = FILTER(<> 74012210));

                trigger OnAfterGetRecord()
                begin
                    IF PrinttoExcel AND PurchaseCreditMemo THEN BEGIN
                        /* ExciseSetup.RESET;
                        ExciseSetup.SETRANGE("Excise Bus. Posting Group", "Purch. Inv. Line"."Excise Bus. Posting Group");
                        ExciseSetup.SETRANGE("Excise Prod. Posting Group", "Excise Prod. Posting Group");
                        IF ExciseSetup.FINDLAST THEN; */

                        CLEAR(recItemCategory);
                        recItemCategory.SETRANGE(recItemCategory.Code, "Purch. Cr. Memo Line"."Item Category Code");
                        IF recItemCategory.FINDFIRST THEN;


                        IF "Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::"Charge (Item)" THEN BEGIN
                            /* PostedStrOrdrLineDetails.RESET;
                            PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Purch. Cr. Memo Hdr."."No.");
                            PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Item No.", "Purch. Cr. Memo Line"."No.");
                            PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                            PostedStrOrdrLineDetails.SETFILTER(PostedStrOrdrLineDetails."Tax/Charge Group", '%1', 'EXCISE');
                            IF PostedStrOrdrLineDetails.FINDFIRST THEN
                                REPEAT
                                    AddExcise1 := AddExcise1 + PostedStrOrdrLineDetails."Calculation Value";
                                UNTIL PostedStrOrdrLineDetails.NEXT = 0; */
                        END;

                        CLEAR(DisChgs);
                        /*  recPostedStrOrdDetails.RESET;
                         recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Inv. Header"."No.");
                         recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Inv. Line"."Line No.");
                         recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                         recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'DISCOUNT');//03Aug2017
                         IF recPostedStrOrdDetails.FINDFIRST THEN
                             REPEAT
                                 DisChgs := DisChgs + recPostedStrOrdDetails.Amount
                             UNTIL recPostedStrOrdDetails.NEXT = 0; */

                        //>>03Aug2017 FreightAmt
                        CLEAR(freightamt);
                        /* recPostedStrOrdDetails.RESET;
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Inv. Header"."No.");
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Inv. Line"."Line No.");
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'FREIGHT');
                        IF recPostedStrOrdDetails.FINDFIRST THEN
                            REPEAT
                                freightamt := freightamt + recPostedStrOrdDetails.Amount
                            UNTIL recPostedStrOrdDetails.NEXT = 0; */
                        //>>03Aug2017 FreightAmt

                        //DJ 17042021
                        CLEAR(TCSPCharges);
                        /* recPostedStrOrdDetails.RESET;
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Cr. Memo Line"."No.");
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Cr. Memo Line"."Line No.");
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'TCSP');//03Aug2017
                        IF recPostedStrOrdDetails.FINDFIRST THEN
                            REPEAT
                                TCSPCharges := TCSPCharges + recPostedStrOrdDetails."Amount (LCY)";
                            UNTIL recPostedStrOrdDetails.NEXT = 0; */
                        LineTCSPCharges += TCSPCharges;
                        //DJ 17042021

                        //CLEAR(TotalQtyofITEM1);
                        IF "Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::Item THEN BEGIN
                            IF "Purch. Cr. Memo Line"."Item Category Code" = 'CAT22' THEN
                                TotalQtyofITEM1 := TotalQtyofITEM1 + "Purch. Cr. Memo Line"."Quantity (Base)"//20Sep2017
                            ELSE
                                TotalQtyofITEM1 := TotalQtyofITEM1 + "Purch. Cr. Memo Line".Quantity;//20Sep2017
                                                                                                     //TotalQtyofITEM1 += TotalQtyofITEM1 + "Purch. Cr. Memo Line".Quantity;
                        END;

                        //CLEAR(ExcessShortQty1);
                        IF "Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::"Charge (Item)" THEN
                            ExcessShortQty := ExcessShortQty + "Purch. Inv. Line".Quantity;

                        TotalQty1 += TotalQtyofITEM1;
                        TotalExcessShortQty1 += ExcessShortQty;


                        CLEAR(TaxPer);
                        CLEAR(TaxGroupCode);
                        /* RecDTE.RESET;
                        RecDTE.SETRANGE(RecDTE."Transaction Type", RecDTE."Transaction Type"::Purchase);
                        RecDTE.SETRANGE(RecDTE."Document Type", RecDTE."Document Type"::"Credit Memo");
                        RecDTE.SETRANGE("Document No.", "Document No.");
                        RecDTE.SETRANGE(Type, Type);
                        RecDTE.SETRANGE("No.", "No.");
                        RecDTE.SETRANGE("Document Line No.", "Line No.");
                        IF RecDTE.FINDFIRST THEN BEGIN
                            IF RecTG.GET(RecDTE."Tax Group Code") THEN;
                            IF RecDTE."Form Code" <> '' THEN
                                TaxGroupCode := RecTG.Description + 'WITH FORM ' + RecDTE."Form Code"
                            ELSE 
                                TaxGroupCode := RecTG.Description;
                            TaxPer := FORMAT(RecDTE."Tax %") + '%';

                        END
                        ELSE BEGIN
                            TaxGroupCode := 'NoTax Ag.Form' + "Purch. Cr. Memo Hdr."."Form Code";
                        END;
                        IF (Vendor."Vendor Posting Group" = 'FOREIGN') AND ("Purch. Inv. Header"."Form Code" = '') THEN BEGIN
                            TaxGroupCode := 'NoTax Ag.Export';
                        END; */

                        //RSPL-Dhananjay START
                        CLEAR(Vtext);
                        IF RecLocation.GET("Purch. Cr. Memo Line"."Location Code") THEN BEGIN
                            LocationState := RecLocation."State Code"
                        END ELSE
                            IF rECResposiblityCenter.GET("Purch. Cr. Memo Line"."Responsibility Center") THEN
                                LocationState := rECResposiblityCenter.State;

                        IF RecVendor.GET("Purch. Cr. Memo Line"."Buy-from Vendor No.") THEN BEGIN
                            IF RecVendor."Gen. Bus. Posting Group" = 'DOMESTIC' THEN BEGIN
                                RecPurchCreLine.RESET;
                                RecPurchCreLine.SETRANGE(RecPurchCreLine."No.", "Purch. Cr. Memo Line"."Document No.");
                                IF RecPurchCreLine.FINDSET THEN
                                    IF RecPurchCreLine."Order Address Code" <> '' THEN BEGIN
                                        orderadd.RESET;
                                        orderadd.SETRANGE("Vendor No.", "Buy-from Vendor No.");//24May2017
                                        orderadd.SETRANGE(orderadd.Code, RecPurchCreLine."Order Address Code");
                                        IF orderadd.FINDSET THEN
                                            IF orderadd.State = LocationState THEN
                                                Vtext := 'INTRA STATE' //03Aug2017
                                                                       //Vtext := 'LOCAL'
                                            ELSE
                                                Vtext := 'INTER STATE';
                                    END ELSE
                                        IF RecPurchCreLine."Order Address Code" = '' THEN BEGIN
                                            IF RecVendor."State Code" = LocationState THEN
                                                Vtext := 'INTRA STATE' //03Aug2017
                                                                       //Vtext := 'LOCAL'
                                            ELSE
                                                Vtext := 'INTER STATE';
                                        END
                            END
                        END
                        ELSE
                            IF RecVendor."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                                Vtext := 'IMPORT';

                        //>>02Aug2017 GST
                        CLEAR(vCGST);
                        CLEAR(vSGST);
                        CLEAR(vIGST);
                        CLEAR(vUTGST);
                        CLEAR(vCGSTRate);
                        CLEAR(vSGSTRate);
                        CLEAR(vIGSTRate);
                        CLEAR(vUTGSTRate);

                        DetailGSTEntry.RESET;
                        DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                        DetailGSTEntry.SETRANGE("Entry Type", DetailGSTEntry."Entry Type"::"Initial Entry");//29Mar2019
                                                                                                            //DetailGSTEntry.SETRANGE(Type,Type);//05Aug2017
                                                                                                            //DetailGSTEntry.SETRANGE(DetailGSTEntry."No.","No.");//05Aug2017
                        IF DetailGSTEntry.FINDSET THEN
                            REPEAT

                                IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                    vCGST += ABS(DetailGSTEntry."GST Amount");
                                    vCGSTRate := DetailGSTEntry."GST %";
                                END;

                                IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                    vSGST += ABS(DetailGSTEntry."GST Amount");
                                    vSGSTRate := DetailGSTEntry."GST %";

                                END;

                                IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                    vUTGST += ABS(DetailGSTEntry."GST Amount");
                                    vUTGSTRate := DetailGSTEntry."GST %";

                                END;


                                IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                    vIGST += ABS(DetailGSTEntry."GST Amount");
                                    vIGSTRate := DetailGSTEntry."GST %";
                                    //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");
                                END;

                            UNTIL DetailGSTEntry.NEXT = 0;

                        TCGST += vCGST;
                        TTCGST += vCGST;

                        TSGST += vSGST;
                        TTSGST += vSGST;

                        TIGST += vIGST;
                        TTIGST += vIGST;

                        TUTGST += vUTGST;
                        TTUTGST += vUTGST;
                        //<<02Aug2017 GST
                        ExchangeRateCrMemo := 0;
                        IF "Purch. Cr. Memo Hdr."."Currency Code" <> '' THEN
                            ExchangeRateCrMemo := 1 / "Purch. Cr. Memo Hdr."."Currency Factor"
                        ELSE
                            ExchangeRateCrMemo := 1;
                        //RSPLAM29072021  --

                        //RSPLAM29072021  ++

                        IF PrinttoExcel AND Detail AND PurchaseCreditMemo THEN BEGIN
                            MakePurchCrMemoDataBody;
                        END;

                        //>>20Sep2017
                        IF "Purch. Cr. Memo Hdr."."Currency Factor" = 0 THEN BEGIN
                            GSTBaseAmt += 0;// "Purch. Cr. Memo Line"."GST Base Amount";//02Aug2017
                            TGSTBaseAmt += 0;// "Purch. Cr. Memo Line"."GST Base Amount";//02Aug2017
                            GTaxBaseAmt += 0;//"Purch. Cr. Memo Line"."Tax Base Amount";
                            TGTaxBaseAmt += 0;// "Purch. Cr. Memo Line"."Tax Base Amount";
                        END;

                        IF "Purch. Cr. Memo Hdr."."Currency Factor" <> 0 THEN BEGIN

                            GSTBaseAmt += 0/*"Purch. Cr. Memo Line"."GST Base Amount"*/ / "Purch. Cr. Memo Hdr."."Currency Factor";//02Aug2017
                            TGSTBaseAmt += 0/* "Purch. Cr. Memo Line"."GST Base Amount" *// "Purch. Cr. Memo Hdr."."Currency Factor";//02Aug2017

                            GTaxBaseAmt += 0/*"Purch. Cr. Memo Line"."Tax Base Amount"*/ / "Purch. Cr. Memo Hdr."."Currency Factor";
                            TGTaxBaseAmt += 0/* "Purch. Cr. Memo Line"."Tax Base Amount"*/ / "Purch. Cr. Memo Hdr."."Currency Factor";
                        END;

                        //<<20Sep2017


                        //RSPL-Dhananjay END
                        //>>TEST10
                        vCrLineAmt += "Purch. Cr. Memo Line"."Line Amount";
                        vCrExciseBase += 0;// "Purch. Cr. Memo Line"."Excise Base Amount";
                        vCrBEDAmt += 0;// "Purch. Cr. Memo Line"."BED Amount";
                        vCrEcessAmt += 0;// "Purch. Cr. Memo Line"."eCess Amount";
                        vCrSheCessAmt += 0;// "Purch. Cr. Memo Line"."SHE Cess Amount";
                        vCrAEDAmt += 0;// "Purch. Cr. Memo Line"."ADE Amount";
                        VCrTaxAmt += 0;// "Purch. Cr. Memo Line"."Tax Amount";
                        vCrTaxBaseAmt += 0;// "Purch. Cr. Memo Line"."Tax Base Amount";
                        //<<

                    END;
                end;

                trigger OnPostDataItem()
                begin
                    //**Migration

                    IF CPINVLINE <> 0 THEN
                        //IF TotalQtyofITEM1 <> 0 THEN //02May2017
                        IF PrinttoExcel AND PurchaseCreditMemo THEN BEGIN
                            MakePurchCrMemoTotDataGrouping;
                        END;

                    IF "Purch. Cr. Memo Hdr."."Currency Code" = '' THEN BEGIN
                        //TotalAmount2  +=  "Purch. Cr. Memo Line"."Line Amount";
                        TotalBillAmtInclusivEcx1 += BillAmtInclusivEcx1;
                        TotalGrossValue1 += 0;// "Purch. Cr. Memo Line"."Amount To Vendor";
                        TotalNetValue1 += 0;//"Purch. Cr. Memo Line".Amount;
                        TotalExciseBaseAmt1 += 0;// "Purch. Cr. Memo Line"."Excise Base Amount";
                        TotalBedAmt1 += 0;// "Purch. Cr. Memo Line"."BED Amount";
                        TotaleCessAmt1 += 0;//"Purch. Cr. Memo Line"."eCess Amount";
                        TotalSHECess1 += 0;// "Purch. Cr. Memo Line"."SHE Cess Amount";
                        TotalTaxBaseAmt1 += "Purch. Cr. Memo Line".Amount + 0;// "Purch. Cr. Memo Line"."Charges To Vendor";
                        TotalTaxAmt1 += 0;//"Purch. Cr. Memo Line"."Tax Amount";
                        totadditionalduty1 += additnlamt1;
                        totalfreight1 += freightamt1;
                        totDisChgs1 += DisChgs1;
                        TotalAddExcise1 += AddExcise1;
                        //TotalTaxableAmount1  +=  "Purch. Cr. Memo Line"."Tax Base Amount";
                        TotalDeliveryChrgs1 += DeliveryChrgs1;
                    END
                    ELSE BEGIN
                        //TotalAmount2  +=  "Purch. Cr. Memo Line"."Line Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalBillAmtInclusivEcx1 += BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalGrossValue1 += 0;//"Purch. Cr. Memo Line"."Amount To Vendor" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalNetValue1 += 0;// "Purch. Cr. Memo Line".Amount / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalExciseBaseAmt1 += 0;// "Purch. Cr. Memo Line"."Excise Base Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalBedAmt1 += 0;// "Purch. Cr. Memo Line"."BED Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotaleCessAmt1 += 0;//"Purch. Cr. Memo Line"."eCess Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalSHECess1 += 0;//"Purch. Cr. Memo Line"."SHE Cess Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalTaxBaseAmt1 += "Purch. Cr. Memo Line".Amount + 0;//"Purch. Cr. Memo Line"."Charges To Vendor"
                                                                              // / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalTaxAmt1 += 0;// "Purch. Cr. Memo Line"."Tax Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        totadditionalduty1 += additnlamt1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        totalfreight1 += freightamt1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        totDisChgs1 += DisChgs1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalAddExcise1 += AddExcise1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        //TotalTaxableAmount1  +=  "Purch. Cr. Memo Line"."Tax Base Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalDeliveryChrgs1 += DeliveryChrgs1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                    END;
                    TotalTCSPCharges += TCSPCharges; //DJ 170421
                end;

                trigger OnPreDataItem()
                begin
                    /*
                    IF PrinttoExcel AND PurchaseReturn THEN
                    BEGIN
                     MakePurchReturnDataHeader
                    END;
                    */

                    //>>02Apr2018
                    IF "Purch. Inv. Line".GETFILTER(Type) <> '' THEN
                        SETFILTER(Type, "Purch. Inv. Line".GETFILTER(Type));

                    IF "Purch. Inv. Line".GETFILTER("No.") <> '' THEN
                        SETFILTER("No.", "Purch. Inv. Line".GETFILTER("No."));

                    IF "Purch. Inv. Line".GETFILTER("Item Category Code") <> '' THEN
                        SETFILTER("Item Category Code", "Purch. Inv. Line".GETFILTER("Item Category Code"));

                    IF "Purch. Inv. Line".GETFILTER("Gen. Prod. Posting Group") <> '' THEN
                        SETFILTER("Gen. Prod. Posting Group", "Purch. Inv. Line".GETFILTER("Gen. Prod. Posting Group"));
                    //<<02Apr2018

                    CPINVLINE := COUNT;//05May2017
                    LineTCSPCharges := 0;

                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>03Aug2017
                CLEAR(BillStateCode);
                CLEAR(BillStateName);
                CLEAR(ShipStateCode);
                CLEAR(ShipStateName);
                CLEAR(VenGSTNo);

                Ven03.RESET;
                IF Ven03.GET("Buy-from Vendor No.") THEN BEGIN
                    VenGSTNo := Ven03."GST Registration No.";
                    State03.RESET;
                    State03.SETRANGE(Code, Ven03."State Code");
                    IF State03.FINDFIRST THEN BEGIN
                        BillStateCode := State03."State Code (GST Reg. No.)";
                        BillStateName := State03.Description;
                    END;
                END;

                IF "Order Address Code" = '' THEN BEGIN
                    ShipStateCode := BillStateCode;
                    ShipStateName := BillStateName;
                END;


                CLEAR(ShipGSTNo); //11Aug2017
                IF "Order Address Code" <> '' THEN BEGIN
                    Ord03.RESET;
                    Ord03.SETRANGE("Vendor No.", "Buy-from Vendor No.");
                    Ord03.SETRANGE(Code, "Order Address Code");
                    IF Ord03.FINDFIRST THEN BEGIN
                        ShipGSTNo := Ord03."GST Registration No.";//11Aug2017
                                                                  //MESSAGE('%1 \ %2 \ %3 ',"Buy-from Vendor No.","Order Address Code",Ord03.State);
                        State03.RESET;
                        State03.SETRANGE(Code, Ord03.State);
                        IF State03.FINDFIRST THEN BEGIN
                            ShipStateCode := State03."State Code (GST Reg. No.)";
                            ShipStateName := State03.Description;

                        END;
                    END;
                END;

                CLEAR(LocStateCode);
                CLEAR(LocGSTNo);
                CLEAR(LocName15);
                CLEAR(LocState15);

                Loc03.RESET;//15Nov2017
                IF Loc03.GET("Location Code") THEN BEGIN
                    LocGSTNo := Loc03."GST Registration No.";
                    LocName15 := Loc03.Name;//15Nov2017

                    State03.RESET;
                    State03.SETRANGE(Code, Loc03."State Code");
                    IF State03.FINDFIRST THEN BEGIN
                        LocStateCode := State03."State Code (GST Reg. No.)";
                        LocState15 := State03.Description;//15Nov2017

                    END;
                END;

                //<<03Aug2017

                //>>RB-N 11Sep2017 GSTNo. Filter

                IF TempGSTNo <> '' THEN BEGIN
                    IF TempGSTNo <> LocGSTNo THEN
                        CurrReport.SKIP;

                END;
                //<<RB-N 11Sep2017 GSTNo. Filter

                "Purch. Cr. Memo Hdr.".CALCFIELDS("Purch. Cr. Memo Hdr."."E-Inv Irn");//RSPLSUM04Mar21

                BillNoLen1 := STRLEN("Purch. Cr. Memo Hdr."."No.");
                "Bill No.1" := "Purch. Cr. Memo Hdr."."No.";

                RespCtr1.SETRANGE(RespCtr1.Code, "Purch. Cr. Memo Hdr."."Responsibility Center");
                IF RespCtr1.FINDFIRST THEN;


                Vendor.RESET;
                IF Vendor.GET("Purch. Inv. Header"."Buy-from Vendor No.") THEN;

                IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
                    OrdAddCity := '';
                    OrdAddTIN := '';
                    OrdAddCST := '';
                    recOrderAddress.RESET;
                    recOrderAddress.SETRANGE(recOrderAddress."Vendor No.", "Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
                    recOrderAddress.SETRANGE(recOrderAddress.Code, "Purch. Cr. Memo Hdr."."Order Address Code");
                    IF recOrderAddress.FINDFIRST THEN BEGIN
                        OrdAddCity := recOrderAddress.City;
                        OrdAddTIN := recOrderAddress."T.I.N. No.";
                        OrdAddCST := recOrderAddress."C.S.T. No.";
                    END;
                END;

                /* DetTaxEntry.RESET;
                DetTaxEntry.SETRANGE(DetTaxEntry."Document No.", "Purch. Cr. Memo Hdr."."No.");
                IF DetTaxEntry.FINDFIRST THEN BEGIN
                    IF DetTaxEntry."Form Code" <> '' THEN BEGIN
                        FormCode1 := 'with Form ' + DetTaxEntry."Form Code";
                    END;
                    CrTaxDescription := FORMAT(DetTaxEntry."Tax Type") + FORMAT(DetTaxEntry."Tax %") + '%' + FormCode1;
                END ELSE BEGIN
                    CrTaxDescription := 'NoTax Ag.Form' + "Purch. Cr. Memo Hdr."."Form Code";
                END;
                IF (Vendor."Vendor Posting Group" = 'FOREIGN') AND ("Purch. Cr. Memo Hdr."."Form Code" = '') THEN BEGIN
                    CrTaxDescription := 'NoTax Ag.Export';
                END; */

                /*  PostedStrDetail.RESET;
                 PostedStrDetail.SETRANGE(PostedStrDetail.Type, PostedStrDetail.Type::Purchase);
                 PostedStrDetail.SETRANGE(PostedStrDetail."Document Type", PostedStrDetail."Document Type"::"Credit Memo");
                 PostedStrDetail.SETRANGE(PostedStrDetail."No.", "Purch. Cr. Memo Hdr."."No.");
                 IF PostedStrDetail.FIND('-') THEN
                     REPEAT
                     BEGIN
                         IF (PostedStrDetail."Tax/Charge Group" = 'EXCISE') THEN BEGIN
                             exciseamt1 := PostedStrDetail."Calculation Value";
                         END;
                         IF (PostedStrDetail."Tax/Charge Group" = 'FREIGHT') THEN BEGIN
                             freightamt1 := PostedStrDetail."Calculation Value";
                         END;
                         IF (PostedStrDetail."Tax/Charge Group" = 'ADD TAX') THEN BEGIN
                             additnlamt1 := PostedStrDetail."Calculation Value";
                         END;
                         IF (PostedStrDetail."Tax/Charge Group" = 'DISCOUNT') THEN BEGIN
                             DisChgs1 := PostedStrDetail."Calculation Value";
                         END;

                     END;
                     UNTIL PostedStrDetail.NEXT = 0; */
                DeliveryChrgs1 := exciseamt1 + freightamt1 + additnlamt1 + DisChgs1;

                PIH.RESET;
                PIH.SETRANGE(PIH."No.", "Purch. Cr. Memo Hdr."."Applies-to Doc. No.");
                IF PIH.FINDFIRST THEN;

                //RSPL-290115
                tCommentCrMemo := '';
                CommentLineCrMemo.RESET;
                CommentLineCrMemo.SETRANGE(CommentLineCrMemo."Document Type", CommentLineCrMemo."Document Type"::"Posted Credit Memo");
                CommentLineCrMemo.SETRANGE(CommentLineCrMemo."No.", "Purch. Cr. Memo Hdr."."No.");
                CommentLineCrMemo.SETRANGE(CommentLineCrMemo."Document Line No.", 0);
                IF CommentLineCrMemo.FINDFIRST THEN
                    tCommentCrMemo := CommentLineCrMemo.Comment;
                //RSPL-290115
            end;

            trigger OnPostDataItem()
            begin

                IF PrinttoExcel AND PurchaseCreditMemo THEN BEGIN
                    MakePurchCrMemoGrandTotal;
                END;

                TotalQty2 := ABS(TotalQty - TotalQty1);
                TotalNetValue2 := ABS(TotalNetValue - TotalNetValue1);
                TotalGrossValue2 := ABS(TotalGrossValue - TotalGrossValue1);
                TotalBedAmt2 := ABS(TotalBedAmt - TotalBedAmt1);
                TotaleCessAmt2 := ABS(TotaleCessAmt - TotaleCessAmt1);
                TotalSHECess2 := ABS(TotalSHECess - TotalSHECess1);
                TotalTaxBaseAmt2 := ABS(TotalTaxBaseAmt - TotalTaxBaseAmt1);
                TotalTaxAmt2 := ABS(TotalTaxAmt - TotalTaxAmt1);
                totalfreight2 := ABS(totalfreight - totalfreight1);
                totDisChgs2 := ABS(totDisChgs - totDisChgs1);
                totadditionalduty2 := ABS(totadditionalduty - totadditionalduty2);
                TotalAmount3 := ABS(TotalAmount1 - TotalAmount2);
                TotalAddExcise2 := ABS(TotalAddExcise - TotalAddExcise1);
                TotalTaxableAmount2 := ABS(TotalTaxableAmount - TotalTaxableAmount1);
                TotalDeliveryChrgs2 := ABS(TotalDeliveryChrgs - TotalDeliveryChrgs1);
                TotalBillAmtInclusivEcx2 := ABS(TotalBillAmtInclusivEcx - TotalBillAmtInclusivEcx1);
                TotalExciseBaseAmt2 := ABS(TotalExciseBaseAmt - TotalExciseBaseAmt1);
                TotalTDSAmount1 := ABS(TotalTDSAmount);
            end;

            trigger OnPreDataItem()
            begin

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Posting Date") <> '' THEN BEGIN
                    "Purch. Cr. Memo Hdr.".SETRANGE("Purch. Cr. Memo Hdr."."Posting Date",
                     "Purch. Inv. Header".GETRANGEMIN("Purch. Inv. Header"."Posting Date"), "Purch. Inv. Header".GETRANGEMAX("Purch. Inv. Header"."Posting Date"));
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Location Code") <> '' THEN BEGIN
                    SETFILTER("Location Code", "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Location Code"));
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Vendor Posting Group") <> '' THEN BEGIN
                    SETFILTER("Vendor Posting Group", "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Vendor Posting Group"));
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Gen. Bus. Posting Group") <> '' THEN BEGIN
                    SETFILTER("Gen. Bus. Posting Group", "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Gen. Bus. Posting Group"));
                END;

                IF PrinttoExcel AND PurchaseCreditMemo THEN BEGIN
                    MakePurchCrMemoDataHeader;
                END;

                GSTBaseAmt := 0;//02Aug2017
                TSGST := 0;//02Aug2017
                TCGST := 0;//02Aug2017
                TIGST := 0;//02Aug2017
                TUTGST := 0;//02Aug2017
            end;
        }
        dataitem("Purch. Cr. Memo Hdr.-ForReturn"; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Pre-Assigned No." = FILTER(''));
            dataitem("Purch. Cr. Memo Line-ForReturn"; "Purch. Cr. Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0),
                                          "No." = FILTER(<> 74012210));

                trigger OnAfterGetRecord()
                begin
                    RINVLINE := COUNT; //20Sep2019
                    CLEAR(recItemCategory);
                    recItemCategory.SETRANGE(recItemCategory.Code, "Purch. Cr. Memo Line-ForReturn"."Item Category Code");
                    IF recItemCategory.FINDFIRST THEN;




                    CLEAR(TaxPer);
                    CLEAR(TaxGroupCode);
                    /* RecDTE.RESET;
                    RecDTE.SETRANGE(RecDTE."Transaction Type", RecDTE."Transaction Type"::Purchase);
                    RecDTE.SETRANGE("Document No.", "Document No.");
                    RecDTE.SETRANGE(Type, Type);
                    RecDTE.SETRANGE("No.", "No.");
                    IF RecDTE.FINDFIRST THEN BEGIN
                        IF RecTG.GET(RecDTE."Tax Group Code") THEN;
                        TaxGroupCode := RecTG.Description;
                        TaxPer := FORMAT(RecDTE."Tax %") + '%'; */


                    // END;

                    CLEAR(DisChgs);
                    /* recPostedStrOrdDetails.RESET;
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Inv. Header"."No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Inv. Line"."Line No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'DISCOUNT');//03Aug2017
                                                                                                           //recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code",'DISCOUNT');
                    IF recPostedStrOrdDetails.FINDFIRST THEN
                        REPEAT
                            DisChgs := DisChgs + recPostedStrOrdDetails.Amount
                        UNTIL recPostedStrOrdDetails.NEXT = 0;

                    //>>03Aug2017 FreightAmt
                    CLEAR(freightamt);
                    recPostedStrOrdDetails.RESET;
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Inv. Header"."No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Inv. Line"."Line No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'FREIGHT');
                    IF recPostedStrOrdDetails.FINDFIRST THEN
                        REPEAT
                            freightamt := freightamt + recPostedStrOrdDetails.Amount
                        UNTIL recPostedStrOrdDetails.NEXT = 0; */
                    //>>03Aug2017 FreightAmt

                    //DJ 17042021
                    CLEAR(TCSPCharges);
                    /*  recPostedStrOrdDetails.RESET;
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Cr. Memo Line-ForReturn"."No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Cr. Memo Line-ForReturn"."Line No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'TCSP');//03Aug2017
                     IF recPostedStrOrdDetails.FINDFIRST THEN
                         REPEAT
                             TCSPCharges := TCSPCharges + recPostedStrOrdDetails."Amount (LCY)";
                         UNTIL recPostedStrOrdDetails.NEXT = 0; */
                    LineTCSPCharges += TCSPCharges;
                    //DJ 17042021

                    //RSPL-Dhananjay START
                    CLEAR(Vtext);
                    IF RecLocation.GET("Purch. Cr. Memo Line-ForReturn"."Location Code") THEN BEGIN
                        LocationState := RecLocation."State Code"
                    END ELSE
                        IF rECResposiblityCenter.GET("Purch. Cr. Memo Line-ForReturn"."Responsibility Center") THEN
                            LocationState := rECResposiblityCenter.State;

                    IF RecVendor.GET("Purch. Cr. Memo Line"."Buy-from Vendor No.") THEN BEGIN
                        IF RecVendor."Gen. Bus. Posting Group" = 'DOMESTIC' THEN BEGIN
                            IF RecVendor."State Code" = LocationState THEN
                                Vtext := 'INTRA STATE' //03Aug2017
                                                       //Vtext := 'LOCAL'
                            ELSE
                                Vtext := 'INTER STATE';
                        END
                    END ELSE
                        IF RecVendor."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            Vtext := 'IMPORT';
                    //RSPL-Dhananjay END


                    //>>02Aug2017 GST
                    CLEAR(vCGST);
                    CLEAR(vSGST);
                    CLEAR(vIGST);
                    CLEAR(vUTGST);
                    CLEAR(vCGSTRate);
                    CLEAR(vSGSTRate);
                    CLEAR(vIGSTRate);
                    CLEAR(vUTGSTRate);

                    DetailGSTEntry.RESET;
                    DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                    DetailGSTEntry.SETRANGE("Entry Type", DetailGSTEntry."Entry Type"::"Initial Entry");//29Mar2019
                    //DetailGSTEntry.SETRANGE(Type,Type);//05Aug2017
                    //DetailGSTEntry.SETRANGE(DetailGSTEntry."No.","No.");//05Aug2017
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT

                            IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                vCGST += ABS(DetailGSTEntry."GST Amount");
                                vCGSTRate := DetailGSTEntry."GST %";
                            END;

                            IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                vSGST += ABS(DetailGSTEntry."GST Amount");
                                vSGSTRate := DetailGSTEntry."GST %";

                            END;

                            IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                vUTGST += ABS(DetailGSTEntry."GST Amount");
                                vUTGSTRate := DetailGSTEntry."GST %";

                            END;


                            IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                vIGST += ABS(DetailGSTEntry."GST Amount");
                                vIGSTRate := DetailGSTEntry."GST %";
                                //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");
                            END;

                        UNTIL DetailGSTEntry.NEXT = 0;

                    TCGST += vCGST;
                    TTCGST += vCGST;

                    TSGST += vSGST;
                    TTSGST += vSGST;

                    TIGST += vIGST;
                    TTIGST += vIGST;

                    TUTGST += vUTGST;
                    TTUTGST += vUTGST;
                    //<<02Aug2017 GST


                    IF PrinttoExcel AND Detail AND PurchaseReturn THEN BEGIN
                        MakePurchReturnDataBody;
                    END;


                    //>>20Sep2017 GroupData

                    IF Type = Type::Item THEN BEGIN
                        IF "Purch. Cr. Memo Line-ForReturn"."Item Category Code" = 'CAT22' THEN BEGIN
                            rtQty += "Purch. Cr. Memo Line-ForReturn"."Quantity (Base)";//04Aug2017
                            vQtyGr += "Purch. Cr. Memo Line-ForReturn"."Quantity (Base)";//04Aug2017
                        END ELSE BEGIN
                            rtQty += "Purch. Cr. Memo Line-ForReturn".Quantity;//04Aug2017
                            vQtyGr += "Purch. Cr. Memo Line-ForReturn".Quantity;//04Aug2017
                        END;
                    END;

                    rtExciseQty += ExcessShortQty1;//04Aug2017

                    IF "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor" = 0 THEN BEGIN
                        rtAmt += "Purch. Cr. Memo Line-ForReturn"."Line Amount";//04Aug2017
                        rtExciseBaseAmt += 0;//"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount";//04Aug2017
                        GSTBaseAmt += 0;// "Purch. Cr. Memo Line-ForReturn"."GST Base Amount";//02Aug2017
                        TGSTBaseAmt += 0;//"Purch. Cr. Memo Line-ForReturn"."GST Base Amount";//02Aug2017
                        rtBEDAmt += 0;// "Purch. Cr. Memo Line-ForReturn"."BED Amount";//04Aug2017
                        rtAddDuty += 0;// "Purch. Cr. Memo Line-ForReturn"."ADE Amount";//04Aug2017
                        GTaxBaseAmt += 0;//"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount";
                        TGTaxBaseAmt += 0;// "Purch. Cr. Memo Line-ForReturn"."Tax Base Amount";
                        rtTaxAmt += 0;// "Purch. Cr. Memo Line-ForReturn"."Tax Amount";//04Aug2017
                        rtGrossAmt += 0;//"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor";//04Aug2017
                    END;

                    IF "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor" <> 0 THEN BEGIN
                        rtAmt += "Purch. Cr. Memo Line-ForReturn"."Line Amount" / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//04Aug2017
                        rtExciseBaseAmt += /*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//04Aug2017
                        GSTBaseAmt += /*"Purch. Cr. Memo Line-ForReturn"."GST Base Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//02Aug2017
                        TGSTBaseAmt += /*0"Purch. Cr. Memo Line-ForReturn"."GST Base Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//02Aug2017
                        rtBEDAmt += /*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//04aug2017
                        rtAddDuty += /*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//04aug2017
                        GTaxBaseAmt += /*"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount"*/ 0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
                        TGTaxBaseAmt += /*"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount"*/ 0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
                        rtTaxAmt += /*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"* / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//04Aug2017
                        rtGrossAmt += /*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//04Aug2017

                    END;
                    //<<20Sep2017 GroupData

                    //>>04Aug2017
                    vCrLineAmt += "Line Amount";
                    vCrExciseBase += 0;// "Excise Base Amount";
                    vCrBEDAmt += 0;// "BED Amount";
                    vCrAEDAmt += 0;//"ADE Amount";
                    VCrTaxAmt += 0;// "Tax Amount";
                    vCrTaxBaseAmt += 0;// "Tax Base Amount";
                    //<<04Aug2017
                end;

                trigger OnPostDataItem()
                begin
                    //IF vQtyGr <> 0 THEN //02May2017
                    IF RINVLINE <> 0 THEN //20Sep2017
                        IF PrinttoExcel AND PurchaseReturn THEN BEGIN
                            MakePurchReturnTotDataGrouping;
                        END;
                end;

                trigger OnPreDataItem()
                begin
                    //>>02Apr2018
                    IF "Purch. Inv. Line".GETFILTER(Type) <> '' THEN
                        SETFILTER(Type, "Purch. Inv. Line".GETFILTER(Type));

                    IF "Purch. Inv. Line".GETFILTER("No.") <> '' THEN
                        SETFILTER("No.", "Purch. Inv. Line".GETFILTER("No."));

                    IF "Purch. Inv. Line".GETFILTER("Item Category Code") <> '' THEN
                        SETFILTER("Item Category Code", "Purch. Inv. Line".GETFILTER("Item Category Code"));

                    IF "Purch. Inv. Line".GETFILTER("Gen. Prod. Posting Group") <> '' THEN
                        SETFILTER("Gen. Prod. Posting Group", "Purch. Inv. Line".GETFILTER("Gen. Prod. Posting Group"));
                    //<<02Apr2018
                    LineTCSPCharges := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin
                IF PrinttoExcel AND PurchaseReturn THEN BEGIN
                    IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
                        OrdAddCity := '';
                        OrdAddTIN := '';
                        OrdAddCST := '';
                        recOrderAddress.RESET;
                        recOrderAddress.SETRANGE(recOrderAddress."Vendor No.", "Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.");
                        recOrderAddress.SETRANGE(recOrderAddress.Code, "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code");
                        IF recOrderAddress.FINDFIRST THEN BEGIN
                            OrdAddCity := recOrderAddress.City;
                            OrdAddTIN := recOrderAddress."T.I.N. No.";
                            OrdAddCST := recOrderAddress."C.S.T. No.";
                        END;
                    END;


                    //RSPL-290115
                    tCommentCrMemo2 := '';
                    CommentLineCrMemo2.RESET;
                    CommentLineCrMemo2.SETRANGE(CommentLineCrMemo2."Document Type", CommentLineCrMemo2."Document Type"::"Posted Credit Memo");
                    CommentLineCrMemo2.SETRANGE(CommentLineCrMemo2."No.", "Purch. Cr. Memo Hdr.-ForReturn"."No.");
                    CommentLineCrMemo2.SETRANGE(CommentLineCrMemo2."Document Line No.", 0);
                    IF CommentLineCrMemo2.FINDFIRST THEN
                        tCommentCrMemo2 := CommentLineCrMemo2.Comment;
                    //RSPL-290115
                END;

                //>>03Aug2017
                CLEAR(BillStateCode);
                CLEAR(BillStateName);
                CLEAR(ShipStateCode);
                CLEAR(ShipStateName);
                CLEAR(VenGSTNo);

                Ven03.RESET;
                IF Ven03.GET("Buy-from Vendor No.") THEN BEGIN
                    VenGSTNo := Ven03."GST Registration No.";
                    State03.RESET;
                    State03.SETRANGE(Code, Ven03."State Code");
                    IF State03.FINDFIRST THEN BEGIN
                        BillStateCode := State03."State Code (GST Reg. No.)";
                        BillStateName := State03.Description;
                    END;
                END;

                IF "Order Address Code" = '' THEN BEGIN
                    ShipStateCode := BillStateCode;
                    ShipStateName := BillStateName;
                END;


                CLEAR(ShipGSTNo); //11Aug2017
                IF "Order Address Code" <> '' THEN BEGIN
                    Ord03.RESET;
                    Ord03.SETRANGE("Vendor No.", "Buy-from Vendor No.");
                    Ord03.SETRANGE(Code, "Order Address Code");
                    IF Ord03.FINDFIRST THEN BEGIN
                        ShipGSTNo := Ord03."GST Registration No.";//11Aug2017
                                                                  //MESSAGE('%1 \ %2 \ %3 ',"Buy-from Vendor No.","Order Address Code",Ord03.State);
                        State03.RESET;
                        State03.SETRANGE(Code, Ord03.State);
                        IF State03.FINDFIRST THEN BEGIN
                            ShipStateCode := State03."State Code (GST Reg. No.)";
                            ShipStateName := State03.Description;

                        END;
                    END;
                END;


                CLEAR(LocStateCode);
                CLEAR(LocGSTNo);
                CLEAR(LocName15);
                CLEAR(LocState15);

                Loc03.RESET;//15Nov2017
                IF Loc03.GET("Location Code") THEN BEGIN
                    LocGSTNo := Loc03."GST Registration No.";
                    LocName15 := Loc03.Name;//15Nov2017

                    State03.RESET;
                    State03.SETRANGE(Code, Loc03."State Code");
                    IF State03.FINDFIRST THEN BEGIN
                        LocStateCode := State03."State Code (GST Reg. No.)";
                        LocState15 := State03.Description;//15Nov2017

                    END;
                END;

                //<<03Aug2017

                //>>RB-N 11Sep2017 GSTNo. Filter

                IF TempGSTNo <> '' THEN BEGIN
                    IF TempGSTNo <> LocGSTNo THEN
                        CurrReport.SKIP;

                END;
                //<<RB-N 11Sep2017 GSTNo. Filter

                "Purch. Cr. Memo Hdr.-ForReturn".CALCFIELDS("Purch. Cr. Memo Hdr.-ForReturn"."E-Inv Irn");//RSPLSUM04Mar21
            end;

            trigger OnPostDataItem()
            begin
                IF PrinttoExcel AND PurchaseReturn THEN BEGIN
                    MakePurchReturnGrandTotal;
                END;
            end;

            trigger OnPreDataItem()
            begin

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Posting Date") <> '' THEN BEGIN
                    "Purch. Cr. Memo Hdr.-ForReturn".SETRANGE("Purch. Cr. Memo Hdr.-ForReturn"."Posting Date",
                    "Purch. Inv. Header".GETRANGEMIN("Purch. Inv. Header"."Posting Date"), "Purch. Inv. Header".GETRANGEMAX("Purch. Inv. Header"."Posting Date"));
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Location Code") <> '' THEN BEGIN
                    SETFILTER("Location Code", "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Location Code"));
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Vendor Posting Group") <> '' THEN BEGIN
                    SETFILTER("Vendor Posting Group", "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Vendor Posting Group"));
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Gen. Bus. Posting Group") <> '' THEN BEGIN
                    SETFILTER("Gen. Bus. Posting Group", "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Gen. Bus. Posting Group"));
                END;

                IF PrinttoExcel AND PurchaseReturn THEN BEGIN
                    MakePurchReturnDataHeader;
                    CLEAR(ExcessShortQty1);
                    CLEAR(BillAmtInclusivEcx1);
                    CLEAR(DisChgs);
                    CLEAR(freightamt);
                    //CLEAR(TotalQtyofITEM1);
                END;

                GSTBaseAmt := 0;//02Aug2017
                TSGST := 0;//02Aug2017
                TCGST := 0;//02Aug2017
                TIGST := 0;//02Aug2017
                TUTGST := 0;//02Aug2017
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("GST Register No."; TempGSTNo)
                {
                    TableRelation = "GST Registration Nos.";
                }
                field(PrinttoExcel; PrinttoExcel)
                {
                    Caption = 'Print to Excel';
                }
                field(Detail; Detail)
                {
                    Caption = 'Detail';
                }
                field("Print Summary"; PrintSummary)
                {
                    Visible = false;
                }
                field(PurchaseInvoice; PurchaseInvoice)
                {
                    Caption = 'Purchase Invoice';
                }
                field(PurchaseCreditMemo; PurchaseCreditMemo)
                {
                    Caption = 'Purchase Credit Memo';
                }
                field(PurchaseReturn; PurchaseReturn)
                {
                    Caption = 'Purchase Return';
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
        /*
        IF PrinttoExcel AND PurchaseInvoice AND PurchaseCreditMemo THEN
        BEGIN
         MakeExcelDataFooter;
        END;
        *///03Aug2017

        IF PrinttoExcel THEN
            CreateExcelbook;

    end;

    trigger OnPreReport()
    begin

        CompInfo.GET;
        DateFilter := "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Posting Date");
        PostingGroup := "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Vendor Posting Group");
        IF PostingGroup = 'RAW MAT' THEN
            datefiltertext := 'RAW MATERIAL REGISTER ' + ' For the Period of  ' + DateFilter
        ELSE
            datefiltertext := PostingGroup + ' REGISTER ' + ' For the Period of  ' + DateFilter;

        IF PrinttoExcel THEN
            MakeExcelInfo;
    end;

    var
        "Bill No.": Code[30];
        CSONo: Code[10];
        CompInfo: Record 79;
        Location: Record 14;
        TotalGrossValue: Decimal;
        TotalNetValue: Decimal;
        TotalBedAmt: Decimal;
        TotaleCessAmt: Decimal;
        TotalSHECess: Decimal;
        TotalQty: Decimal;
        PurchOrdeDate: Date;
        "Bill No.1": Code[20];
        RefInvNo: Code[20];
        TotalQty1: Decimal;
        TotalGrossValue1: Decimal;
        TotalNetValue1: Decimal;
        TotalExciseBaseAmt1: Decimal;
        TotalBedAmt1: Decimal;
        TotaleCessAmt1: Decimal;
        TotalSHECess1: Decimal;
        TotalQty2: Decimal;
        TotalGrossValue2: Decimal;
        TotalNetValue2: Decimal;
        TotalBedAmt2: Decimal;
        TotaleCessAmt2: Decimal;
        TotalSHECess2: Decimal;
        RespCtr: Record 5714;
        RespCtr1: Record 5714;
        Vendor: Record 23;
        PIH: Record 122;
        PIL: Record 123;
        Item: Record 27;
        Date: Date;
        DateFilter: Text[30];
        BillNoLen: Integer;
        BillNoLen1: Integer;
        UserSetup: Record 91;
        USERID: Text[30];
        Loc: Text[30];
        filtertext: Text[30];
        "Rec Addss": Text[100];
        PrinttoExcel: Boolean;
        datefiltertext: Text[250];
        RefInvNoLen: Integer;
        InvTaxDescription: Text[80];
        InvoiceStatus: Text[80];
        TotalTaxBaseAmt: Decimal;
        TotalTaxAmt: Decimal;
        CrTaxDescription: Text[80];
        TotalTaxBaseAmt1: Decimal;
        TotalTaxAmt1: Decimal;
        TotalTaxBaseAmt2: Decimal;
        TotalTaxAmt2: Decimal;
        // DetTaxEntry: Record "16522";
        InvTaxType: Text[30];
        CrTaxType: Text[30];
        FormCode: Text[30];
        FormCode1: Text[30];
        totalfreight: Decimal;
        totDisChgs: Decimal;
        totadditionalduty: Decimal;
        // PostedStrDetail: Record "13760";
        exciseamt: Decimal;
        freightamt: Decimal;
        additnlamt: Decimal;
        totalfreight1: Decimal;
        totDisChgs1: Decimal;
        totalfreight2: Decimal;
        totDisChgs2: Decimal;
        TaxableAmount: Decimal;
        PostingGroup: Text[30];
        // PostedStrOrdrLineDetails: Record "13798";
        AddTax: Decimal;
        AddExcise: Decimal;
        BillAmtInclusivEcx: Decimal;
        DeliveryChrgs: Decimal;
        Amount1: Decimal;
        TaxAmt: Decimal;
        GrossAmt: Decimal;
        PurchRcptHdr: Record 120;
        GRANo: Code[20];
        RecVendor: Record 23;
        VATTINNo: Code[20];
        CSTTINNo: Code[20];
        TotalAmount1: Decimal;
        TotalBillAmtInclusivEcx: Decimal;
        Amount2: Decimal;
        TaxAmt1: Decimal;
        GrossAmt1: Decimal;
        TotalAmount2: Decimal;
        TotalBillAmtInclusivEcx2: Decimal;
        BillAmtInclusivEcx1: Decimal;
        TaxableAmount1: Decimal;
        freightamt1: Decimal;
        DeliveryChrgs1: Decimal;
        exciseamt1: Decimal;
        additnlamt1: Decimal;
        TotalBillAmtInclusivEcx1: Decimal;
        totadditionalduty1: Decimal;
        DisChgs: Decimal;
        DisChgs1: Decimal;
        AddExcise1: Decimal;
        TotalAddExcise: Decimal;
        TotalAddExcise1: Decimal;
        TotalTaxableAmount: Decimal;
        TotalTaxableAmount1: Decimal;
        TotalDeliveryChrgs: Decimal;
        TotalDeliveryChrgs1: Decimal;
        TotalAmount3: Decimal;
        totadditionalduty2: Decimal;
        TotalAddExcise2: Decimal;
        TotalTaxableAmount2: Decimal;
        TotalDeliveryChrgs2: Decimal;
        // ExciseSetup: Record "13711";
        TAXAmountvalue: Decimal;
        ExcessShortQty: Decimal;
        TotalQtyofITEM: Decimal;
        // recPostedStrOrdDetails: Record "13798";
        // recExciseEntry: Record "13712";
        BEDpercent: Decimal;
        TotalTDSAmount: Decimal;
        TotalTDSAmount1: Decimal;
        ExcelBuf: Record 370 temporary;
        Detail: Boolean;
        PurchaseInvoice: Boolean;
        PurchaseCreditMemo: Boolean;
        TotalQtyofITEM1: Decimal;
        ExcessShortQty1: Decimal;
        TotalExciseBaseAmt: Decimal;
        TotalExciseBaseAmt2: Decimal;
        PurchaseReturn: Boolean;
        TotalExcessShortQty: Decimal;
        TotalExcessShortQty1: Decimal;
        recItemCategory: Record 5722;
        recPRH: Record 120;
        POno: Code[20];
        POdate: Date;
        BlanketOrderNo: Code[20];
        BlanketOrderDate: Date;
        GRNdate: Date;
        recPH: Record 38;
        ChargeItemTaxBaseAmt: Decimal;
        VendorPostingGroupFilter: Text[100];
        recOrderAddress: Record 224;
        OrdAddCity: Code[20];
        OrdAddTIN: Code[20];
        OrdAddCST: Code[20];
        recVend: Record 23;
        recOrderAddr: Record 224;
        recState: Record State;
        recState2: Record State;
        StateDesc: Text[30];
        // RecDTE: Record "16522";
        TaxPer: Text[100];
        TaxGroupCode: Text[100];
        RecTG: Record 321;
        vLandedCost: Decimal;
        "---Comments---": Integer;
        CommentLine: Record 43;
        tComment: Text[250];
        CommentLineCrMemo: Record 43;
        tCommentCrMemo: Text[250];
        CommentLineCrMemo2: Record 43;
        tCommentCrMemo2: Text[250];
        RecLocation: Record 14;
        LocationState: Code[10];
        Vtext: Text[30];
        rECResposiblityCenter: Record 5714;
        recPurchInv: Record 122;
        OrdeState: Code[10];
        orderadd: Record 224;
        RecPurchCreLine: Record 124;
        Text002: Label 'Data';
        Text003: Label 'Purchase Register';
        Text004: Label 'Company Name';
        Text005: Label 'Report No.';
        Text006: Label 'Report Name';
        Text007: Label 'User ID';
        Text008: Label 'Date';
        Text009: Label 'Location Filter';
        "--RSPL-Rahul": Integer;
        vLineAmt: Decimal;
        vExciseBaseAmt: Decimal;
        vBEDAmt: Decimal;
        vEcessAmt: Decimal;
        vSheCessAmt: Decimal;
        vADEAmt: Decimal;
        vTaxBaseAmt: Decimal;
        vTaxAmt: Decimal;
        vTDSAmt: Decimal;
        "--Cr": Integer;
        vCrLineAmt: Decimal;
        vCrBEDAmt: Decimal;
        vCrEcessAmtp: Decimal;
        vCrSheCessAmt: Decimal;
        vCrAEDAmt: Decimal;
        vCrExciseBase: Decimal;
        VCrTaxAmt: Decimal;
        vCrTaxBaseAmt: Decimal;
        "--Rt": Integer;
        rtQty: Decimal;
        rtExciseQty: Decimal;
        rtAmt: Decimal;
        rtBEDAmt: Decimal;
        rtExciseBaseAmt: Decimal;
        rtEcess: Decimal;
        rtSheCess: Decimal;
        rtAddDuty: Decimal;
        rtExciseDutyShortage: Decimal;
        rtBillAmt: Decimal;
        rtTaxableAmt: Decimal;
        rtTaxAmt: Decimal;
        rtGrossAmt: Decimal;
        vQtyGr: Decimal;
        vCrEcessAmt: Decimal;
        "-----05May2017": Integer;
        PINVLINE: Integer;
        CPINVLINE: Integer;
        RINVLINE: Integer;
        "----------------GST--02Aug2017": Integer;
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        vUTGST: Decimal;
        vUTGSTRate: Decimal;
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        TCGST: Decimal;
        TTCGST: Decimal;
        TSGST: Decimal;
        TTSGST: Decimal;
        TIGST: Decimal;
        TTIGST: Decimal;
        TUTGST: Decimal;
        TTUTGST: Decimal;
        GSTBaseAmt: Decimal;
        TGSTBaseAmt: Decimal;
        BillStateCode: Code[10];
        BillStateName: Text[50];
        ShipStateCode: Code[10];
        ShipStateName: Text[50];
        VenGSTNo: Code[15];
        State03: Record State;
        Ord03: Record 224;
        Ven03: Record 23;
        Loc03: Record 14;
        LocStateCode: Code[10];
        LocGSTNo: Code[15];
        GTaxBaseAmt: Decimal;
        TGTaxBaseAmt: Decimal;
        ShipGSTNo: Code[15];
        "------11Sep2017": Integer;
        TempGSTNo: Code[15];
        "----20Sep2017": Integer;
        PrintSummary: Boolean;
        GSTPer: Decimal;
        "----------15Nov2017": Integer;
        LocName15: Text;
        LocState15: Text;
        "----------07Jul2018": Integer;
        GSTGroup07: Record "GST Group";
        GSTCat07: Text;
        "--------19Oct2018": Integer;
        VE19: Record 5802;
        ILE19: Record 32;
        ILEDocNo: Code[20];
        MRNo: Code[30];
        PRL: Record 121;
        "------RK-----": Integer;
        PurRecLine_rec: Record 121;
        UintOfMeasure: Code[20];
        RecItemN: Record 27;
        UintOfMeasureCR: Code[20];
        UintOfMeasureCRRet: Code[20];
        TCSPCharges: Decimal;
        TotalTCSPCharges: Decimal;
        LineTCSPCharges: Decimal;
        ExchangeRate: Decimal;
        ExchangeRateCrMemo: Decimal;

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(FORMAT(Text003), FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(50093, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text007), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;

        IF PrinttoExcel AND PurchaseInvoice THEN BEGIN
            MakePurchInvDataHeader;
        END;
    end;

    // [Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text002,Text003,COMPANYNAME,USERID);
        /*  ExcelBuf.CreateBookAndOpenExcel('', 'Purchase Register GST', '', '', '');
         ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook('Purchase Register GST');
        ExcelBuf.WriteSheet('Purchase Register GST', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Purchase Register GST', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//41
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//42
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//51
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//52
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//53
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//54
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//59
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//60 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//61 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//62 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//63 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//64 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//65 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//66 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//67 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//68 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//69 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//70 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//71 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//72 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//73 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//74 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//75 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//76 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//77 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//78 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//79 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//80 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//84 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//85 03Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//88 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//89 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//90 DJ 17042021
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPLAM29072021
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//80

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Purchase Register', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//60 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//61 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//62 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//63 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//64 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//65 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//66 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//67 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//68 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//69 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//70 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//71 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//72 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//73 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//74 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//75 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//76 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//77 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//78 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//79 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//80 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//84 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//85 03Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//88 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//89 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//90 DJ 17042021
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPLAM29072021
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//80

        //>>09Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(DateFilter, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //<<
    end;

    // [Scope('Internal')]
    procedure MakePurchInvDataHeader()
    begin
        //ExcelBuf.NewRow;
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Invoice Details', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//77 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//78 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//85 03Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//88 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//89 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//90 DJ 17042021
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPLAM29072021
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//80
        //ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('Type of Material', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('Type of Supply', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('State Code of Shipping Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('Name of Shipping Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('State Code of Billing Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('Name of Billing Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('Vendor GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('Vendor Shipping GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn('Vendor Inv.No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('Vendor Inv. Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('Party Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('Party Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('Dealer Category', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14 15Nov2017
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('Product HSN Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('Item Discription', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('GST Group Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('GST Category', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('Payment Terms', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('Qty Received', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('Excise Short Qty', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('Unit Cost', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('INR Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('FC Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('Insurance', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('Frieght', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('Discount Charges', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('Excise Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('GST Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('BED %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('SGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('CGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('IGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('UTGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('Additional Duty Amout', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('SGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('CGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41
        ExcelBuf.AddColumn('IGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//42
        ExcelBuf.AddColumn('UTGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43
        ExcelBuf.AddColumn('Excise Duty On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn('GST On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//45
        ExcelBuf.AddColumn('Bill Amount Inclusive Excise', FALSE, '', TRUE, FALSE, TRUE, '', 1);//46
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//47
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '', 1);//48
        ExcelBuf.AddColumn('Rate of Tax', FALSE, '', TRUE, FALSE, TRUE, '', 1);//49
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('Other Adjustment', FALSE, '', TRUE, FALSE, TRUE, '', 1);//51
        ExcelBuf.AddColumn('TDS Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53
        ExcelBuf.AddColumn('E-Way Bill No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54
        ExcelBuf.AddColumn('E-Way Bill Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55
        ExcelBuf.AddColumn('Resp. Centre', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56
        ExcelBuf.AddColumn('Receiving Location /Place of Supply', FALSE, '', TRUE, FALSE, TRUE, '', 1);//57
        ExcelBuf.AddColumn('Receiving State', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58 15Nov2017
        ExcelBuf.AddColumn('State Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//59
        ExcelBuf.AddColumn('GPPL GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60
        ExcelBuf.AddColumn('PO No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61
        ExcelBuf.AddColumn('PO Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62
        ExcelBuf.AddColumn('Blanket Order No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63
        ExcelBuf.AddColumn('Blanket Order Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64
        ExcelBuf.AddColumn('GRN No. / TRR No', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65
        ExcelBuf.AddColumn('GRN Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66
        ExcelBuf.AddColumn('Category Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67
        ExcelBuf.AddColumn('Category Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68
        ExcelBuf.AddColumn('Vendor VAT TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69
        ExcelBuf.AddColumn('Vendor CST TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74
        ExcelBuf.AddColumn('State', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75
        ExcelBuf.AddColumn('Bonded Rate', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76
        ExcelBuf.AddColumn('ExBond Rate', FALSE, '', TRUE, FALSE, TRUE, '', 1);//77
        ExcelBuf.AddColumn('Landed Cost', FALSE, '', TRUE, FALSE, TRUE, '', 1);//78
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '', 1); //79
        ExcelBuf.AddColumn('BOE No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('BOE Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('BOE Value', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
        ExcelBuf.AddColumn('MR No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//85 03Apr2019
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '', 1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Requisitioner', FALSE, '', TRUE, FALSE, TRUE, '', 1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Department', FALSE, '', TRUE, FALSE, TRUE, '', 1);//88 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Deal Sheet No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//89 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('TCSP Charges', FALSE, '', TRUE, FALSE, TRUE, '', 1);//90 DJ 17042021
        ExcelBuf.AddColumn('Exchange Rate', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLAM29072021
    end;

    //  [Scope('Internal')]
    procedure MakePurchInvDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Inv. Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn("Purch. Inv. Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Vtext, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ShipStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(ShipStateName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(BillStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(BillStateName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(VenGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn(ShipGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Invoice No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Purch. Inv. Header"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//11
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Purch. Inv. Header"."GST Vendor Type", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14 15Nov2017 Dealer Category
        ExcelBuf.AddColumn("Purch. Inv. Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn("Purch. Inv. Line"."HSN/SAC Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn("Purch. Inv. Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn("Purch. Inv. Line"."GST Group Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);// 07Jul2018
        GSTCat07 := '';
        GSTGroup07.RESET;
        GSTGroup07.SETRANGE(Code, "Purch. Inv. Line"."GST Group Code");
        GSTGroup07.SETRANGE("Reverse Charge", TRUE);
        IF GSTGroup07.FINDFIRST THEN
            GSTCat07 := 'RCM'
        ELSE
            GSTCat07 := '';
        ExcelBuf.AddColumn(GSTCat07, FALSE, '', FALSE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn("Purch. Inv. Header"."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17

        //RSPL-RK-Start
        PurRecLine_rec.RESET;
        PurRecLine_rec.SETRANGE("Document No.", '' /*"Purch. Inv. Line"."Receipt Document No."*/);
        //PurRecLine_rec.SETRANGE("Line No.",'' /*"Purch. Inv. Line"."Receipt Document Line No."*/);
        IF PurRecLine_rec.FINDFIRST THEN BEGIN
            IF PurRecLine_rec."Item Category Code" = 'CAT22' THEN BEGIN
                RecItemN.RESET;
                IF RecItemN.GET(PurRecLine_rec."No.") THEN
                    UintOfMeasure := RecItemN."Base Unit of Measure";
            END ELSE
                UintOfMeasure := PurRecLine_rec."Unit of Measure Code";
        END;

        IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::Item THEN
            ExcelBuf.AddColumn(UintOfMeasure, FALSE, '', FALSE, FALSE, FALSE, '', 1)//18
        ELSE
            ExcelBuf.AddColumn("Purch. Inv. Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//18
        //RSPL-RK-End

        IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::Item THEN BEGIN
            IF "Purch. Inv. Line"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn("Purch. Inv. Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0)//19
            ELSE
                ExcelBuf.AddColumn("Purch. Inv. Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//19
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//19

        ExcelBuf.AddColumn(ExcessShortQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//20
        ExcelBuf.AddColumn("Purch. Inv. Header"."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        IF "Purch. Inv. Line"."Item Category Code" = 'CAT22' THEN
            ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)" / "Purch. Inv. Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//22
        ELSE
            ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22

        IF "Purch. Inv. Header"."Currency Code" = '' THEN BEGIN
            IF "Purch. Inv. Line"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)" / "Purch. Inv. Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//23
            ELSE
                ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//25
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn(freightamt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn(DisChgs, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn("Purch. Inv. Line"."Line Amount", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(/* "Purch. Inv. Line"."Excise Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."GST Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31


            ExcelBuf.AddColumn(BEDpercent, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %

            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."BED Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."ADE Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//38
            ExcelBuf.AddColumn(vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn(vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn(vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(AddExcise, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//44 GSTAmountShortage

            IF (/*"Purch. Inv. Line"."Excise Amount"*/0 <> 0) AND (/*"Purch. Inv. Line"."BED Amount"*/0 = 0) THEN BEGIN
                BillAmtInclusivEcx := "Purch. Inv. Line"."Line Amount" + 0/*"Purch. Inv. Line"."Excise Amount"*/ +
                                              AddExcise;
            END ELSE BEGIN
                BillAmtInclusivEcx := "Purch. Inv. Line"."Line Amount" + 0;// "Purch. Inv. Line"."BED Amount" +
                                                                           //"Purch. Inv. Line"."eCess Amount" + "Purch. Inv. Line"."SHE Cess Amount" +
                                                                           //"Purch. Inv. Line"."ADE Amount" + AddExcise;
            END;

            ExcelBuf.AddColumn(BillAmtInclusivEcx, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//45

            IF ("Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)") AND
               (("Purch. Inv. Line"."No." = 'AD 01') OR ("Purch. Inv. Line"."No." = 'BO 01') OR ("Purch. Inv. Line"."No." = 'CH 01'))
               AND (/*"Purch. Inv. Line"."Tax Amount"*/0 = 0) THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//46
                ChargeItemTaxBaseAmt += 0;// "Purch. Inv. Line"."Tax Base Amount";
            END
            ELSE BEGIN
                ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Tax Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//46

            END;
            ExcelBuf.AddColumn(FORMAT(TaxGroupCode), FALSE, '', FALSE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn(FORMAT(TaxPer), FALSE, '', FALSE, FALSE, FALSE, '', 1);//48

            TAXAmountvalue := 0;//"Purch. Inv. Line"."Tax Amount";
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Tax Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//49

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//50
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."TDS Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//51
            ExcelBuf.AddColumn("Purch. Inv. Line"."Amount To Vendor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//52

        END ELSE BEGIN
            //
            IF "Purch. Inv. Line"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)" / "Purch. Inv. Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//23
            ELSE
                ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
            ExcelBuf.AddColumn("Purch. Inv. Line"."Direct Unit Cost", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24
            ExcelBuf.AddColumn("Purch. Inv. Line"."Line Amount", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//25
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn(freightamt / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn(DisChgs / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn("Purch. Inv. Line"."Line Amount" / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Excise Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."GST Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31


            ExcelBuf.AddColumn(BEDpercent, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %

            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."BED Amount"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."ADE Amount"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//38
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(vUTGST / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(AddExcise / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//44 GSTAmountShortage

            IF (/*"Purch. Inv. Line"."Excise Amount"*/0 <> 0) AND (/*"Purch. Inv. Line"."BED Amount"*/0 = 0) THEN BEGIN
                BillAmtInclusivEcx := "Purch. Inv. Line"."Line Amount" + 0/*"Purch. Inv. Line"."Excise Amount"*/ +
                                              AddExcise;
            END ELSE BEGIN
                BillAmtInclusivEcx := "Purch. Inv. Line"."Line Amount" + 0;// "Purch. Inv. Line"."BED Amount" +
                                                                           //"Purch. Inv. Line"."eCess Amount" + "Purch. Inv. Line"."SHE Cess Amount" +
                                                                           //"Purch. Inv. Line"."ADE Amount" + AddExcise;
            END;

            ExcelBuf.AddColumn(BillAmtInclusivEcx / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//45

            //CLEAR(ChargeItemTaxBaseAmt);
            IF ("Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)") AND
               (("Purch. Inv. Line"."No." = 'AD 01') OR ("Purch. Inv. Line"."No." = 'BO 01') OR
               ("Purch. Inv. Line"."No." = 'CH 01')) AND (/*"Purch. Inv. Line"."Tax Amount"*/0 = 0) THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//46
                ChargeItemTaxBaseAmt += 0;// "Purch. Inv. Line"."Tax Base Amount";
            END
            ELSE BEGIN
                ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Tax Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//46

            END;
            ExcelBuf.AddColumn(FORMAT(TaxGroupCode), FALSE, '', FALSE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn(FORMAT(TaxPer), FALSE, '', FALSE, FALSE, FALSE, '', 1);//48
            TAXAmountvalue := 0;// "Purch. Inv. Line"."Tax Amount";

            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Tax Amount"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//49
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//50
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."TDS Amount"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//51
            ExcelBuf.AddColumn("Purch. Inv. Line"."Amount To Vendor" / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//52
        END;

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//53 E-Way Bill No.
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//54 E-Way Bill Date
        ExcelBuf.AddColumn(RespCtr.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LocName15, FALSE, '', FALSE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn(LocState15, FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 Receiving State 15Nov2017
        ExcelBuf.AddColumn(LocStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//57 GSTState Code
        ExcelBuf.AddColumn(LocGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 GSTRegNo
        ExcelBuf.AddColumn(POno, FALSE, '', FALSE, FALSE, FALSE, '', 1);//59
        ExcelBuf.AddColumn(POdate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//60
        ExcelBuf.AddColumn(BlanketOrderNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//61
        ExcelBuf.AddColumn(BlanketOrderDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//62
        IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)" THEN
            ExcelBuf.AddColumn(ILEDocNo, FALSE, '', FALSE, FALSE, FALSE, '', 1)//63
        ELSE
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Receipt Document No."*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//63
        ExcelBuf.AddColumn(GRNdate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//64
        ExcelBuf.AddColumn("Purch. Inv. Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//66
        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Inv. Header"."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';//RecVendor."T.I.N. No.";
            CSTTINNo := '';// RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn("Purch. Inv. Header"."Order Address Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//69
        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END;
        //sn-BEGIN
        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Inv. Header"."Buy-from Vendor No.", "Purch. Inv. Header"."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END
        ELSE
            IF recState.GET(/*"Purch. Inv. Header".State*/) THEN;

        ExcelBuf.AddColumn(recState.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//73
                                                                                        //sn-end
        ExcelBuf.AddColumn("Purch. Inv. Line"."Bonded Rate", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//74
        ExcelBuf.AddColumn("Purch. Inv. Line"."Exbond Rate", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//75
        ExcelBuf.AddColumn("Purch. Inv. Line"."Landed Cost", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//76
        ExcelBuf.AddColumn(tComment, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//77

        ExcelBuf.AddColumn("Purch. Inv. Header"."Bill of Entry No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn("Purch. Inv. Header"."Bill of Entry Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//83 19Oct2018
        ExcelBuf.AddColumn("Purch. Inv. Header"."Bill of Entry Value", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//84 19Oct2018
        ExcelBuf.AddColumn(MRNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//85 03Apr2019
        ExcelBuf.AddColumn("Purch. Inv. Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn("Purch. Inv. Header"."Purchaser Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn("Purch. Inv. Header"."Department Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//88 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn("Purch. Inv. Header"."Deal Sheet No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//89 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn(TCSPCharges, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//90 DJ 170421
        ExcelBuf.AddColumn(ExchangeRate, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//RSPLAM29072021
    end;

    // [Scope('Internal')]
    procedure MakePurchInvTotalDataGrouping()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Inv. Header"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn("Purch. Inv. Header"."No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Posting Group", FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Vtext, FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ShipStateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(ShipStateName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(BillStateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(BillStateName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(VenGSTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn(ShipGSTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Invoice No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Purch. Inv. Header"."Document Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//11
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Purch. Inv. Header"."GST Vendor Type", FALSE, '', TRUE, FALSE, FALSE, '', 1);//14 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn(TotalQtyofITEM, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//19
        ExcelBuf.AddColumn(ExcessShortQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//20
        ExcelBuf.AddColumn("Purch. Inv. Header"."Currency Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//23

        IF "Purch. Inv. Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(vLineAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(vExciseBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn(GSTBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//31 GSTBaseAmount
            GSTBaseAmt := 0;//02Aug2017

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//36 //UTGST %
            ExcelBuf.AddColumn(vBEDAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(vADEAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0); //38
            ExcelBuf.AddColumn(TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn(TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn(TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            TSGST := 0;//02Aug2017
            TCGST := 0;//02Aug2017
            TIGST := 0;//02Aug2017
            TUTGST := 0;//02Aug2017

            ExcelBuf.AddColumn(AddExcise, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//44 GSTShortageAmount

            IF (/*"Purch. Inv. Line"."Excise Amount"*/0 <> 0) AND (vBEDAmt = 0) THEN BEGIN
                BillAmtInclusivEcx := vLineAmt + /*"Purch. Inv. Line"."Excise Amount"*/0 + AddExcise;
            END ELSE BEGIN
                BillAmtInclusivEcx := vLineAmt + vBEDAmt + vEcessAmt + vSheCessAmt + AddExcise + AddExcise;
            END;

            ExcelBuf.AddColumn(BillAmtInclusivEcx, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//45

            ExcelBuf.AddColumn(GTaxBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//46
            GTaxBaseAmt := 0;//03Aug2017

            ExcelBuf.AddColumn(TaxGroupCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn(TaxPer, FALSE, '', TRUE, FALSE, FALSE, '', 1);//48
            ExcelBuf.AddColumn(vTaxAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0); //49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//50
            ExcelBuf.AddColumn(vTDSAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//51
            //"Purch. Inv. Header".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(/*"Purch. Inv. Header"."Amount to Vendor"*/0, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//52
        END ELSE BEGIN

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn(vLineAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//25
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(vLineAmt / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(vExciseBaseAmt / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//31 GSTBaseAmount
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//32
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %
            ExcelBuf.AddColumn(vBEDAmt / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(vADEAmt / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//38
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(AddExcise / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//44 GSTShortageAmount

            IF (/*"Purch. Inv. Line"."Excise Amount"*/0 <> 0) AND (vBEDAmt = 0) THEN BEGIN
                BillAmtInclusivEcx := vLineAmt + /*"Purch. Inv. Line"."Excise Amount"*/0 + AddExcise;
            END ELSE BEGIN
                BillAmtInclusivEcx := vLineAmt + vBEDAmt + vEcessAmt + vSheCessAmt + vADEAmt + AddExcise;
            END;

            ExcelBuf.AddColumn(BillAmtInclusivEcx / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//45
            ExcelBuf.AddColumn(GTaxBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//46
            GTaxBaseAmt := 0;//03Aug2017
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//48
            ExcelBuf.AddColumn(vTaxAmt / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//50
            ExcelBuf.AddColumn(vTDSAmt / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//51

            //"Purch. Inv. Header".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(/*"Purch. Inv. Header"."Amount to Vendor"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//52
        END;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//53 E-Way Bill No
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//54 E-Way Bill No
        ExcelBuf.AddColumn(RespCtr.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LocName15, FALSE, '', TRUE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn(LocState15, FALSE, '', TRUE, FALSE, FALSE, '', 1);//58 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//57 GSTState Code
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 GSTRegNo
        ExcelBuf.AddColumn(POno, FALSE, '', TRUE, FALSE, FALSE, '', 1);//59
        ExcelBuf.AddColumn(POdate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//60
        ExcelBuf.AddColumn(BlanketOrderNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//61
        ExcelBuf.AddColumn(BlanketOrderDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//62
        IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)" THEN
            ExcelBuf.AddColumn(ILEDocNo, FALSE, '', TRUE, FALSE, FALSE, '', 1)//63
        ELSE
            ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Receipt Document No."*/'', FALSE, '', TRUE, FALSE, FALSE, '', 1);//63
        ExcelBuf.AddColumn(GRNdate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//66

        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Inv. Header"."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';//RecVendor."T.I.N. No.";
            CSTTINNo := '';// RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn("Purch. Inv. Header"."Order Address Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//69
        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', TRUE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', TRUE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', TRUE, FALSE, FALSE, '', 1);//72
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END;

        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Inv. Header"."Buy-from Vendor No.", "Purch. Inv. Header"."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END
        ELSE
            IF recState.GET(/*"Purch. Inv. Header".State*/) THEN;
        ExcelBuf.AddColumn(recState.Description, FALSE, '', TRUE, FALSE, FALSE, '', 1);//73

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//74
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//75

        ExcelBuf.AddColumn("Purch. Inv. Line"."Landed Cost", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//76
        ExcelBuf.AddColumn(tComment, FALSE, '', TRUE, FALSE, FALSE, '', 1);//77

        ExcelBuf.AddColumn("Purch. Inv. Header"."Bill of Entry No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn("Purch. Inv. Header"."Bill of Entry Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//83 19Oct2018
        ExcelBuf.AddColumn("Purch. Inv. Header"."Bill of Entry Value", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//84 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//85 03Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//88 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//89 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn(LineTCSPCharges, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//90 DJ170421
        ExcelBuf.AddColumn(ExchangeRate, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//RSPLAM29072021

        //**Migration/Clear Variables test8
        TotalQty += TotalQtyofITEM;
        TotalExcessShortQty += ExcessShortQty;
        TotalExciseBaseAmt += vExciseBaseAmt;
        TotalAddExcise += AddExcise;
        TotalTaxableAmount += vTaxBaseAmt - ChargeItemTaxBaseAmt;
        TotalTaxAmt += vTaxAmt;//Sourav  110215
                               // "Purch. Inv. Header".CALCFIELDS("Amount to Vendor");
        TotalGrossValue += 0;// "Purch. Inv. Header"."Amount to Vendor";
        //vLandedCost += "Purch. Inv. Line"."Landed Cost";
        totadditionalduty += vADEAmt;
        //TotalQtyofITEM += "Purch. Inv. Line".Quantity;
        CLEAR(TotalQtyofITEM);
        CLEAR(ExcessShortQty);
        CLEAR(AddExcise);
        CLEAR(vLineAmt);
        CLEAR(vExciseBaseAmt);
        CLEAR(vBEDAmt);
        CLEAR(vEcessAmt);
        CLEAR(vSheCessAmt);
        CLEAR(vADEAmt);
        CLEAR(ChargeItemTaxBaseAmt);
        CLEAR(vTaxBaseAmt);
        CLEAR(vTaxAmt);
        CLEAR(vTDSAmt);
        CLEAR(vTaxBaseAmt);
    end;

    // [Scope('Internal')]
    procedure MakePurchInvGrandTotal()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '1', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '2', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '3', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '4', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '5', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '6', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '7', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '8', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '9', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '10', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '11', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '12', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '13', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '14', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '15', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '16', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '17', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '18', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '19', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '20', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '21', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '22', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '23', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '24', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '25', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '26', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '27', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '28', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '29', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '30', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '31', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '32', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '33', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '34', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //59
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //60
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //61
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //63
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //66
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //67
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //68
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //69
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //70
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //71
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //72
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //73
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //74
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //76
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //78 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //79 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //80 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //82 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//85 03Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//88 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//89 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//90 dj170421
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLAM29072021

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GRAND TOTAL(Invoice)', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '1', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '2', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '3', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '4', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '5', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '6', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '7', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '8', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '9', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '10', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //11 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //14 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); // 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); // 07Jul2018
        ExcelBuf.AddColumn(TotalQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0); //19
        ExcelBuf.AddColumn(TotalExcessShortQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0); //20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn(totalfreight, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//27 //Freight
        ExcelBuf.AddColumn(totDisChgs, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//28 //Discount
        totalfreight := 0; //03Aug2017
        totDisChgs := 0; //03Aug2017

        ExcelBuf.AddColumn(TotalAmount1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0); //29
        ExcelBuf.AddColumn(TotalExciseBaseAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//30
        ExcelBuf.AddColumn(TGSTBaseAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//31 GSTBaseAmt
        TGSTBaseAmt := 0;//02Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn(TotalBedAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//37
        ExcelBuf.AddColumn(totadditionalduty, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//38
        ExcelBuf.AddColumn(TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//39
        ExcelBuf.AddColumn(TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//40
        ExcelBuf.AddColumn(TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//41
        ExcelBuf.AddColumn(TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//42
        TTSGST := 0;//02Aug2017
        TTCGST := 0;//02Aug2017
        TTIGST := 0;//02Aug2017
        TTUTGST := 0;//02Aug2017

        ExcelBuf.AddColumn(TotalAddExcise, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//44 GSTShortageAmount
        ExcelBuf.AddColumn(TotalBillAmtInclusivEcx, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//45
        //ExcelBuf.AddColumn(TotalTaxableAmount-ChargeItemTaxBaseAmt -rtTaxableAmt,FALSE,'',TRUE,FALSE,TRUE,'#,#0.00',0);//46
        ExcelBuf.AddColumn(TGTaxBaseAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//46 //03Aug2017
        TGTaxBaseAmt := 0; //04Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //48
        ExcelBuf.AddColumn(TotalTaxAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn(TotalTDSAmount, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//51
        ExcelBuf.AddColumn(TotalGrossValue, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//52
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//59
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//85 03Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//88 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//89 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn(TotalTCSPCharges, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//DJ 170421
        TotalTCSPCharges := 0; //DJ 170421
    end;

    // [Scope('Internal')]
    procedure MakePurchCrMemoDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Credit Memo Details', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
        /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//85 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//88 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//85 RSPLSUM04Mar21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//86 DJ170421

        //ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('Type of Material', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('Type of Supply', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('State Code of Shipping Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('Name of Shipping Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('State Code of Billing Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('Name of Billing Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('Vendor GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('Vendor Shipping GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn('Vendor Inv.No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('Vendor Inv. Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('Party Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('Party Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('Dealers Category', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('Product HSN Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('Item Discription', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('GST Group Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('GST Category', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('Payment Terms', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('Qty Received', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('Excise Short Qty', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('Unit Cost', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('INR Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('FC Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('Insurance', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('Frieght', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('Discount Charges', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('Excise Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('GST Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('BED %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('SGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('CGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('IGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('UTGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('Additional Duty Amout', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('SGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('CGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('IGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41
        ExcelBuf.AddColumn('UTGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//42
        ExcelBuf.AddColumn('Excise Duty On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43
        ExcelBuf.AddColumn('GST On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn('Bill Amount Inclusive Excise', FALSE, '', TRUE, FALSE, TRUE, '', 1);//45
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//46
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '', 1);//47
        ExcelBuf.AddColumn('Rate of Tax', FALSE, '', TRUE, FALSE, TRUE, '', 1);//48
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//49
        ExcelBuf.AddColumn('Other Adjustment', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('TDS Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//51
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
        ExcelBuf.AddColumn('E-Way Bill No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53
        ExcelBuf.AddColumn('E-Way Bill Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54
        ExcelBuf.AddColumn('Resp. Centre', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55
        ExcelBuf.AddColumn('Receiving Location /Place of Supply', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56
        ExcelBuf.AddColumn('Receiving State', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58 15Nov2017
        ExcelBuf.AddColumn('State Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//57
        ExcelBuf.AddColumn('GPPL GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58
        ExcelBuf.AddColumn('PO No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//59
        ExcelBuf.AddColumn('PO Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60
        ExcelBuf.AddColumn('Blanket Order No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61
        ExcelBuf.AddColumn('Blanket Order Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62
        ExcelBuf.AddColumn('GRN No. / TRR No', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63
        ExcelBuf.AddColumn('GRN Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64
        ExcelBuf.AddColumn('Category Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65
        ExcelBuf.AddColumn('Category Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66
        ExcelBuf.AddColumn('Vendor VAT TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67
        ExcelBuf.AddColumn('Vendor CST TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72
        ExcelBuf.AddColumn('State', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73
        ExcelBuf.AddColumn('Bonded Rate', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74
        ExcelBuf.AddColumn('ExBond Rate', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75
        ExcelBuf.AddColumn('Landed Cost', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '', 1); //77
        ExcelBuf.AddColumn('BOE No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('BOE Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('BOE Value', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
        /*ExcelBuf.AddColumn('Item Category',FALSE,'',TRUE,FALSE,TRUE,'',1);//85 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Requisitioner',FALSE,'',TRUE,FALSE,TRUE,'',1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Department',FALSE,'',TRUE,FALSE,TRUE,'',1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Deal Sheet Number',FALSE,'',TRUE,FALSE,TRUE,'',1);//88 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn('IRN', FALSE, '', TRUE, FALSE, TRUE, '', 1);//85 RSPLSUM04Mar21
        ExcelBuf.AddColumn('TCSP Charges', FALSE, '', TRUE, FALSE, TRUE, '', 1);//86 DJ170421
        ExcelBuf.AddColumn('Exchange Rate', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLAM29072021

    end;

    // [Scope('Internal')]
    procedure MakePurchCrMemoDataBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Vendor Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Vtext, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ShipStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(ShipStateName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(BillStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(BillStateName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(VenGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn(ShipGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(PIH."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//11
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."GST Vendor Type", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14 15Nov2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."HSN/SAC Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn("Purch. Cr. Memo Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."GST Group Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);// 07Jul2018
        GSTCat07 := '';
        GSTGroup07.RESET;
        GSTGroup07.SETRANGE(Code, "Purch. Cr. Memo Line"."GST Group Code");
        GSTGroup07.SETRANGE("Reverse Charge", TRUE);
        IF GSTGroup07.FINDFIRST THEN
            GSTCat07 := 'RCM'
        ELSE
            GSTCat07 := '';
        ExcelBuf.AddColumn(GSTCat07, FALSE, '', FALSE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
        CLEAR(UintOfMeasureCR);
        IF "Purch. Cr. Memo Line"."Item Category Code" = 'CAT22' THEN BEGIN
            RecItemN.RESET;
            IF RecItemN.GET("Purch. Cr. Memo Line"."No.") THEN
                UintOfMeasureCR := RecItemN."Base Unit of Measure";
            ExcelBuf.AddColumn(UintOfMeasureCR, FALSE, '', FALSE, FALSE, FALSE, '', 1)//18
        END ELSE
            ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//18

        IF "Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::Item THEN BEGIN
            IF "Purch. Cr. Memo Line"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0)//19
            ELSE
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//19
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//19

        ExcelBuf.AddColumn(-1 * ExcessShortQty1, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//20
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        IF "Purch. Cr. Memo Line"."Item Category Code" = 'CAT22' THEN
            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Unit Cost (LCY)" / "Purch. Cr. Memo Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//22
        ELSE
            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22

        IF "Purch. Cr. Memo Hdr."."Currency Code" = '' THEN BEGIN
            IF "Purch. Cr. Memo Line"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Unit Cost (LCY)" / "Purch. Cr. Memo Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//23
            ELSE
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//25
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn(-1 * freightamt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn(-1 * DisChgs, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Line Amount", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line"."Excise Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line"."GST Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31


            ExcelBuf.AddColumn(/*ExciseSetup."BED %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %

            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line"."BED Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line"."ADE Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//38
            ExcelBuf.AddColumn(-1 * vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn(-1 * vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(-1 * vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn(-1 * vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(-1 * AddExcise1, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//44 GSTAmountShortage

            IF (/*"Purch. Cr. Memo Line"."Excise Amount"*/0 <> 0) AND (/*"Purch. Cr. Memo Line"."BED Amount"*/0 = 0) THEN BEGIN
                BillAmtInclusivEcx := "Purch. Cr. Memo Line"."Line Amount" + 0/*"Purch. Cr. Memo Line"."Excise Amount"*/ +
                                              AddExcise;
            END ELSE BEGIN
                BillAmtInclusivEcx := "Purch. Cr. Memo Line"."Line Amount" + 0;// "Purch. Cr. Memo Line"."BED Amount" +
                                                                               //"Purch. Cr. Memo Line"."eCess Amount" + "Purch. Cr. Memo Line"."SHE Cess Amount" +
                                                                               //"Purch. Cr. Memo Line"."ADE Amount" + AddExcise;
            END;

            ExcelBuf.AddColumn(-1 * BillAmtInclusivEcx, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//45

            IF ("Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::"Charge (Item)") AND
               (("Purch. Cr. Memo Line"."No." = 'AD 01') OR ("Purch. Cr. Memo Line"."No." = 'BO 01') OR ("Purch. Cr. Memo Line"."No." = 'CH 01'))
               AND (/*"Purch. Cr. Memo Line"."Tax Amount"*/0 = 0) THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//46
                ChargeItemTaxBaseAmt += 0;// "Purch. Cr. Memo Line"."Tax Base Amount";
            END
            ELSE BEGIN
                ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line"."Tax Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//46

            END;
            ExcelBuf.AddColumn(FORMAT(TaxGroupCode), FALSE, '', FALSE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn(FORMAT(TaxPer), FALSE, '', FALSE, FALSE, FALSE, '', 1);//48

            TAXAmountvalue := 0;// "Purch. Cr. Memo Line"."Tax Amount";
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line"."Tax Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//49

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//50
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//51
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line"."Amount To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//52

        END ELSE BEGIN
            //
            IF "Purch. Cr. Memo Line"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Unit Cost (LCY)" / "Purch. Cr. Memo Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//23
            ELSE
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Direct Unit Cost", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24
            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line"."Line Amount", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//25
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn(-1 * (freightamt / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn(-1 * (DisChgs / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(-1 * ("Purch. Cr. Memo Line"."Line Amount" / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line"."Excise Base Amount"*/ 0 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line"."GST Base Amount"*/0 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31

            ExcelBuf.AddColumn(BEDpercent, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %

            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line"."BED Amount"*/0 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line"."ADE Amount"*/0 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//38
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(-1 * (vIGST / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(-1 * (AddExcise1 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//44 GSTAmountShortage

            IF (/*"Purch. Cr. Memo Line"."Excise Amount"*/0 <> 0) AND (0/*"Purch. Cr. Memo Line"."BED Amount"*/ = 0) THEN BEGIN
                BillAmtInclusivEcx := "Purch. Cr. Memo Line"."Line Amount" + 0/* "Purch. Cr. Memo Line"."Excise Amount*/ +
                                              AddExcise1;
            END ELSE BEGIN
                BillAmtInclusivEcx := "Purch. Cr. Memo Line"."Line Amount" + 0;//"Purch. Cr. Memo Line"."BED Amount" +
                                                                               //"Purch. Cr. Memo Line"."eCess Amount" + "Purch. Cr. Memo Line"."SHE Cess Amount" +
                                                                               // "Purch. Cr. Memo Line"."ADE Amount" + AddExcise1;
            END;

            ExcelBuf.AddColumn(-1 * (BillAmtInclusivEcx / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//45

            //CLEAR(ChargeItemTaxBaseAmt);
            IF ("Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::"Charge (Item)") AND
               (("Purch. Cr. Memo Line"."No." = 'AD 01') OR ("Purch. Cr. Memo Line"."No." = 'BO 01') OR
               ("Purch. Cr. Memo Line"."No." = 'CH 01')) AND (/*"Purch. Cr. Memo Line"."Tax Amount"*/0 = 0) THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//46
                ChargeItemTaxBaseAmt += 0;// "Purch. Cr. Memo Line"."Tax Base Amount";
            END
            ELSE BEGIN
                ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line"."Tax Base Amount" */0 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//46

            END;
            ExcelBuf.AddColumn(FORMAT(TaxGroupCode), FALSE, '', FALSE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn(FORMAT(TaxPer), FALSE, '', FALSE, FALSE, FALSE, '', 1);//48
            TAXAmountvalue := 0;// "Purch. Cr. Memo Line"."Tax Amount";

            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line"."Tax Amount"*/0 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//49
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//50
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//51
            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line"."Amount To Vendor"*/0 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//52
        END;

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//53 E-Way Bill No.
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//54 E-Way Bill Date
        ExcelBuf.AddColumn(RespCtr1.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LocName15, FALSE, '', FALSE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn(LocState15, FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 15Nov2017
        ExcelBuf.AddColumn(LocStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//57 GSTState Code
        ExcelBuf.AddColumn(LocGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 GSTRegNo
        ExcelBuf.AddColumn(POno, FALSE, '', FALSE, FALSE, FALSE, '', 1);//59
        ExcelBuf.AddColumn(POdate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//60
        ExcelBuf.AddColumn(BlanketOrderNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//61
        ExcelBuf.AddColumn(BlanketOrderDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//62
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//63
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 2);//64
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//66
        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';// RecVendor."T.I.N. No.";
            CSTTINNo := '';// RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Order Address Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//69
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END;
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", "Purch. Cr. Memo Hdr."."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END
        ELSE
            IF recState.GET(/*"Purch. Cr. Memo Hdr.".State*/) THEN;

        ExcelBuf.AddColumn(recState.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//73
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//74
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//75
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//76
        ExcelBuf.AddColumn(tCommentCrMemo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//77

        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Bill of Entry No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Bill of Entry Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//83 19Oct2018
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Bill of Entry Value", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//84 19Oct2018
        /*ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Item Category Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//85 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Purchaser Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Department Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//88 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."E-Inv Irn", FALSE, '', FALSE, FALSE, FALSE, '', 1);//85 RPSLSUM04Mar20 
        ExcelBuf.AddColumn(-1 * TCSPCharges, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//86 DJ 170421
        ExcelBuf.AddColumn(ExchangeRateCrMemo, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//RSPLAM29072021

    end;

    // [Scope('Internal')]
    procedure MakePurchCrMemoTotDataGrouping()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Vendor Posting Group", FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Vtext, FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ShipStateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(ShipStateName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(BillStateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(BillStateName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(VenGSTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn(ShipGSTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Applies-to Doc. No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(PIH."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//11
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."GST Vendor Type", FALSE, '', TRUE, FALSE, FALSE, '', 1);//14 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn(-1 * TotalQtyofITEM1, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//19
        ExcelBuf.AddColumn(-1 * ExcessShortQty1, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//20
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Currency Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//23

        IF "Purch. Cr. Memo Hdr."."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(-1 * vCrLineAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(-1 * vCrExciseBase, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn(-1 * GSTBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//31 GSTBaseAmount
            GSTBaseAmt := 0;//02Aug2017

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//36 //UTGST %
            ExcelBuf.AddColumn(-1 * vCrBEDAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(-1 * vCrAEDAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0); //38
            ExcelBuf.AddColumn(-1 * TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn(-1 * TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(-1 * TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn(-1 * TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            TSGST := 0;//02Aug2017
            TCGST := 0;//02Aug2017
            TIGST := 0;//02Aug2017
            TUTGST := 0;//02Aug2017

            ExcelBuf.AddColumn(-1 * AddExcise, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//44 GSTShortageAmount

            IF (/*"Purch. Cr. Memo Line"."Excise Amount"*/0 <> 0) AND (vCrBEDAmt = 0) THEN BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + /*"Purch. Cr. Memo Line"."Excise Amount"*/0 + AddExcise;
            END ELSE BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + vCrBEDAmt + vCrAEDAmt + AddExcise;
            END;

            ExcelBuf.AddColumn(-1 * BillAmtInclusivEcx1, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//45
            ExcelBuf.AddColumn(-1 * GTaxBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//46
            GTaxBaseAmt := 0;//03Aug2017

            ExcelBuf.AddColumn(TaxGroupCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn(TaxPer, FALSE, '', TRUE, FALSE, FALSE, '', 1);//48
            ExcelBuf.AddColumn(-1 * VCrTaxAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0); //49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//51
            //"Purch. Cr. Memo Hdr.".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Hdr."."Amount to Vendor"*/0, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//52
        END ELSE BEGIN

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn(vCrLineAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//25
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(-1 * (vCrLineAmt / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(-1 * (vCrExciseBase / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//31 GSTBaseAmount
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//32
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %
            ExcelBuf.AddColumn(-1 * (vCrBEDAmt / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(-1 * (vCrAEDAmt / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//38
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(-1 * (AddExcise / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//44 GSTShortageAmount

            IF (/*"Purch. Cr. Memo Line"."Excise Amount"*/0 <> 0) AND (vCrBEDAmt = 0) THEN BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + vCrExciseBase + AddExcise;
            END ELSE BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + vCrBEDAmt + vCrAEDAmt + AddExcise;
            END;

            ExcelBuf.AddColumn(-1 * (BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//45
            ExcelBuf.AddColumn(-1 * GTaxBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//46
            GTaxBaseAmt := 0;//03Aug2017
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//48
            ExcelBuf.AddColumn(-1 * (VCrTaxAmt / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//51

            //"Purch. Cr. Memo Hdr.".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Hdr."."Amount to Vendor"*/0 / "Purch. Cr. Memo Hdr."."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//52
        END;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//53 E-Way Bill No
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//54 E-Way Bill No
        ExcelBuf.AddColumn(RespCtr1.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LocName15, FALSE, '', TRUE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn(LocState15, FALSE, '', TRUE, FALSE, FALSE, '', 1);//58 15Nov2017

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//57 GSTState Code
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 GSTRegNo
        ExcelBuf.AddColumn(POno, FALSE, '', TRUE, FALSE, FALSE, '', 1);//59
        ExcelBuf.AddColumn(POdate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//60
        ExcelBuf.AddColumn(BlanketOrderNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//61
        ExcelBuf.AddColumn(BlanketOrderDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//63
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 2);//64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//66

        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';//RecVendor."T.I.N. No.";
            CSTTINNo := '';//RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Order Address Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//69
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', TRUE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', TRUE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', TRUE, FALSE, FALSE, '', 1);//72
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END;

        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", "Purch. Cr. Memo Hdr."."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END
        ELSE
            IF recState.GET(/*"Purch. Cr. Memo Hdr.".State*/) THEN;
        ExcelBuf.AddColumn(recState.Description, FALSE, '', TRUE, FALSE, FALSE, '', 1);//73

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//74
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//75

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//76
        ExcelBuf.AddColumn(tCommentCrMemo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//77

        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Bill of Entry No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Bill of Entry Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//83 19Oct2018
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Bill of Entry Value", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//84 19Oct2018
                                                                                                                      /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//85 RSPLSUM 15Apr2020
                                                                                                                      ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//86 RSPLSUM 15Apr2020
                                                                                                                      ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//87 RSPLSUM 15Apr2020
                                                                                                                      ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//88 RSPLSUM 15Apr2020*/

        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."E-Inv Irn", FALSE, '', TRUE, FALSE, FALSE, '', 1);//85 RSPLSUM04Mar21
        ExcelBuf.AddColumn(LineTCSPCharges, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//86 170421
        ExcelBuf.AddColumn(ExchangeRateCrMemo, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//RSPLAM29072021

        //>>**Tor
        CLEAR(TotalQtyofITEM1);
        CLEAR(ExcessShortQty1);
        CLEAR(vCrLineAmt);
        CLEAR(vCrBEDAmt);
        CLEAR(vCrEcessAmt);
        CLEAR(vCrSheCessAmt);
        CLEAR(vCrAEDAmt);
        CLEAR(vCrExciseBase);
        CLEAR(VCrTaxAmt);
        CLEAR(vCrTaxBaseAmt);
        //<<

    end;

    // [Scope('Internal')]
    procedure MakePurchCrMemoGrandTotal()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//53
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //54
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //59
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //60
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //61
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //63
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //66
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //67
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //68
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //69
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //70
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //71
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //72
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //73
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //74
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //76
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //78 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //79 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //80 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //82 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
        /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//85 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//88 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//85 RSPLSUM04Mar21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//86 DJ170421

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GRAND TOTAL(Cr. Memo)', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '1', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '2', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '3', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '4', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '5', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '6', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '7', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '8', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '9', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '10', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //11 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //14 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); // 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); // 07Jul2018
        ExcelBuf.AddColumn(-1 * TotalQty1, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0); //19
        ExcelBuf.AddColumn(-1 * TotalExcessShortQty1, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0); //20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn(-1 * totalfreight, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//27 //Freight
        ExcelBuf.AddColumn(-1 * totDisChgs, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//28 //Discount
        totalfreight := 0; //03Aug2017
        totDisChgs := 0; //03Aug2017

        ExcelBuf.AddColumn(-1 * TotalAmount2, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0); //29
        ExcelBuf.AddColumn(-1 * TotalExciseBaseAmt1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//30
        ExcelBuf.AddColumn(-1 * TGSTBaseAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//31 GSTBaseAmt
        TGSTBaseAmt := 0;//02Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn(-1 * TotalBedAmt1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//37
        ExcelBuf.AddColumn(-1 * totadditionalduty1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//38
        ExcelBuf.AddColumn(-1 * TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//39
        ExcelBuf.AddColumn(-1 * TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//40
        ExcelBuf.AddColumn(-1 * TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//41
        ExcelBuf.AddColumn(-1 * TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//42
        TTSGST := 0;//02Aug2017
        TTCGST := 0;//02Aug2017
        TTIGST := 0;//02Aug2017
        TTUTGST := 0;//02Aug2017

        ExcelBuf.AddColumn(-1 * TotalAddExcise1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//43
        TotalAddExcise1 := 0;//04Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//44 GSTShortageAmount
        ExcelBuf.AddColumn(-1 * TotalBillAmtInclusivEcx1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//45
        TotalBillAmtInclusivEcx1 := 0;//04Aug2017
        //ExcelBuf.AddColumn(TotalTaxableAmount-ChargeItemTaxBaseAmt -rtTaxableAmt,FALSE,'',TRUE,FALSE,TRUE,'#,#0.00',0);//46
        ExcelBuf.AddColumn(-1 * TGTaxBaseAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//46 //03Aug2017
        TGTaxBaseAmt := 0;//04Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //48
        ExcelBuf.AddColumn(-1 * TotalTaxAmt1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//51
        ExcelBuf.AddColumn(-1 * TotalGrossValue1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//52
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//59
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
        /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//85 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//86 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//87 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//88 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//85 RSPLSUM04Mar21
        ExcelBuf.AddColumn(-1 * TotalTCSPCharges, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//86 DJ70421
        TotalTCSPCharges := 0;// DJ70421

    end;

    // [Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 19Oct2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 19Oct2018
                                                                    /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//85 RSPLSUM 15Apr2020
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//86 RSPLSUM 15Apr2020
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//87 RSPLSUM 15Apr2020
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//88 RSPLSUM 15Apr2020*/

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TotalQty2 - rtQty, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TotalAmount3 - rtAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TotalBedAmt2 - ABS(rtBEDAmt), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotaleCessAmt2 - ABS(rtEcess), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalSHECess2 - ABS(rtSheCess), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(totadditionalduty2 - ABS(rtAddDuty), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalAddExcise2, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalBillAmtInclusivEcx2 - ABS(rtBillAmt), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(totDisChgs2, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(totalfreight2, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalTaxableAmount2 - ABS(rtTaxableAmt), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TotalTaxAmt2 - ABS(rtTaxAmt), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TotalTDSAmount1, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalGrossValue2 - ABS(rtGrossAmt), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
        /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPLSUM 15Apr2020*/

    end;

    // [Scope('Internal')]
    procedure MakePurchReturnDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Return Details', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76 02Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 07Jul2018
        /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//83 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//84 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//85 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//86 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 RSPLSUM04Mar21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//84 DJ170421

        //ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('Type of Material', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('Type of Supply', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('State Code of Shipping Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('Name of Shipping Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('State Code of Billing Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('Name of Billing Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('Vendor GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('Vendor Shipping GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn('Vendor Inv.No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('Vendor Inv. Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('Party Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('Party Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('Dealer Category', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14 15Nov2017
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('Product HSN Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('Item Discription', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('GST Group Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('GST Category', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('Payment Terms', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('Qty Received', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('Excise Short Qty', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('Unit Cost', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('INR Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('FC Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('Insurance', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('Frieght', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('Discount Charges', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('Excise Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('GST Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('BED %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('SGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('CGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('IGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('UTGST %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('Additional Duty Amout', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('SGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('CGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('IGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41
        ExcelBuf.AddColumn('UTGST ', FALSE, '', TRUE, FALSE, TRUE, '', 1);//42
        ExcelBuf.AddColumn('Excise Duty On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43
        ExcelBuf.AddColumn('GST On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn('Bill Amount Inclusive Excise', FALSE, '', TRUE, FALSE, TRUE, '', 1);//45
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//46
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '', 1);//47
        ExcelBuf.AddColumn('Rate of Tax', FALSE, '', TRUE, FALSE, TRUE, '', 1);//48
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//49
        ExcelBuf.AddColumn('Other Adjustment', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('TDS Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//51
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
        ExcelBuf.AddColumn('E-Way Bill No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53
        ExcelBuf.AddColumn('E-Way Bill Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54
        ExcelBuf.AddColumn('Resp. Centre', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55
        ExcelBuf.AddColumn('Receiving Location /Place of Supply', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56
        ExcelBuf.AddColumn('Receiving State', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58 15Nov2017
        ExcelBuf.AddColumn('State Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//57
        ExcelBuf.AddColumn('GPPL GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58
        ExcelBuf.AddColumn('PO No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//59
        ExcelBuf.AddColumn('PO Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60
        ExcelBuf.AddColumn('Blanket Order No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61
        ExcelBuf.AddColumn('Blanket Order Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62
        ExcelBuf.AddColumn('GRN No. / TRR No', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63
        ExcelBuf.AddColumn('GRN Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64
        ExcelBuf.AddColumn('Category Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65
        ExcelBuf.AddColumn('Category Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66
        ExcelBuf.AddColumn('Vendor VAT TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67
        ExcelBuf.AddColumn('Vendor CST TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72
        ExcelBuf.AddColumn('State', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73
        ExcelBuf.AddColumn('Bonded Rate', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74
        ExcelBuf.AddColumn('ExBond Rate', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75
        ExcelBuf.AddColumn('Landed Cost', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '', 1); //77
        /*ExcelBuf.AddColumn('Item Category',FALSE,'',TRUE,FALSE,TRUE,'',1); //78 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Requisitioner',FALSE,'',TRUE,FALSE,TRUE,'',1); //79 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Department',FALSE,'',TRUE,FALSE,TRUE,'',1); //80 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('Deal Sheet Number',FALSE,'',TRUE,FALSE,TRUE,'',1); //81 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn('IRN', FALSE, '', TRUE, FALSE, TRUE, '', 1); //78 RSPLSUM04Mar21
        ExcelBuf.AddColumn('TCSP Charges', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 DJ170421

    end;

    //[Scope('Internal')]
    procedure MakePurchReturnDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Vendor Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Vtext, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ShipStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(ShipStateName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(BillStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(BillStateName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(VenGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn(ShipGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(PIH."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//11
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."GST Vendor Type", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14 15Nov2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."HSN/SAC Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."GST Group Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);// 07Jul2018
        GSTCat07 := '';
        GSTGroup07.RESET;
        GSTGroup07.SETRANGE(Code, "Purch. Cr. Memo Line-ForReturn"."GST Group Code");
        GSTGroup07.SETRANGE("Reverse Charge", TRUE);
        IF GSTGroup07.FINDFIRST THEN
            GSTCat07 := 'RCM'
        ELSE
            GSTCat07 := '';
        ExcelBuf.AddColumn(GSTCat07, FALSE, '', FALSE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17

        CLEAR(UintOfMeasureCRRet);
        IF "Purch. Cr. Memo Line"."Item Category Code" = 'CAT22' THEN BEGIN
            RecItemN.RESET;
            IF RecItemN.GET("Purch. Cr. Memo Line"."No.") THEN
                UintOfMeasureCRRet := RecItemN."Base Unit of Measure";
            ExcelBuf.AddColumn(UintOfMeasureCRRet, FALSE, '', FALSE, FALSE, FALSE, '', 1);//18
        END ELSE
            ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//18

        IF "Purch. Cr. Memo Line-ForReturn".Type = "Purch. Cr. Memo Line-ForReturn".Type::Item THEN BEGIN//ss
            IF "Purch. Cr. Memo Line-ForReturn"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0)//19
            ELSE
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//19
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//19

        ExcelBuf.AddColumn(-1 * ExcessShortQty1, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//20

        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        IF "Purch. Cr. Memo Line-ForReturn"."Item Category Code" = 'CAT22' THEN
            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)" / "Purch. Cr. Memo Line-ForReturn"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//22
        ELSE
            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22

        IF "Purch. Cr. Memo Hdr.-ForReturn"."Currency Code" = '' THEN BEGIN
            IF "Purch. Cr. Memo Line-ForReturn"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)" / "Purch. Cr. Memo Line-ForReturn"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//23
            ELSE
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//25
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn(-1 * freightamt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn(-1 * DisChgs, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Line Amount", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line-ForReturn"."GST Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31


            ExcelBuf.AddColumn(/*ExciseSetup."BED %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %

            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//38

            ExcelBuf.AddColumn(-1 * vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn(-1 * vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(-1 * vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn(-1 * vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(-1 * AddExcise1, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//44 GSTAmountShortage

            IF (/*"Purch. Cr. Memo Line-ForReturn"."Excise Amount"*/0 <> 0) AND (/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0 = 0) THEN BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + /*"Purch. Cr. Memo Line-ForReturn"."Excise Amount"*/0 + AddExcise1;
            END ELSE BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + 0;// "Purch. Cr. Memo Line-ForReturn"."BED Amount" +
                                                                                          //"Purch. Cr. Memo Line-ForReturn"."eCess Amount" + "Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount" +
                                                                                          //"Purch. Cr. Memo Line-ForReturn"."ADE Amount" + AddExcise1;
            END;

            ExcelBuf.AddColumn(-1 * BillAmtInclusivEcx1, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//45
            rtBillAmt += BillAmtInclusivEcx1;//04Aug2017

            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//46

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//48

            TAXAmountvalue := 0;// "Purch. Cr. Memo Line-ForReturn"."Tax Amount";
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//49


            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//51
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//52


        END ELSE BEGIN
            //
            IF "Purch. Cr. Memo Line-ForReturn"."Item Category Code" = 'CAT22' THEN
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)" / "Purch. Cr. Memo Line-ForReturn"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//23
            ELSE
                ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Direct Unit Cost", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24
            ExcelBuf.AddColumn(-1 * "Purch. Cr. Memo Line-ForReturn"."Line Amount", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//25
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn(-1 * (freightamt / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn(-1 * (DisChgs / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(-1 * ("Purch. Cr. Memo Line-ForReturn"."Line Amount" / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29


            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30

            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line-ForReturn"."GST Base Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31


            ExcelBuf.AddColumn(/*ExciseSetup."BED %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %

            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//37


            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//38


            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(-1 * (vUTGST / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(-1 * (AddExcise1 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//44 GSTAmountShortage

            IF (/*"Purch. Cr. Memo Line-ForReturn"."Excise Amount"*/0 <> 0) AND (/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0 = 0) THEN BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + 0 +//"Purch. Cr. Memo Line-ForReturn"."Excise Amount" +
                                              AddExcise1;
            END ELSE BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + 0; ///*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/ +
                /*"Purch. Cr. Memo Line-ForReturn"."eCess Amount"*/// + "Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount" +
                                                                   // "Purch. Cr. Memo Line-ForReturn"."ADE Amount" + AddExcise1;
            END;

            ExcelBuf.AddColumn(-1 * (BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//45
            rtBillAmt += BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";//04Aug2017

            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//46

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//48
            TAXAmountvalue := 0;// "Purch. Cr. Memo Line-ForReturn"."Tax Amount";

            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//49


            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//50
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//51
            ExcelBuf.AddColumn(-1 * (/*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//52

        END;

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//53 E-Way Bill No.
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//54 E-Way Bill Date
        ExcelBuf.AddColumn(RespCtr1.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LocName15, FALSE, '', FALSE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn(LocState15, FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 15Nov2017
        ExcelBuf.AddColumn(LocStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//57 GSTState Code
        ExcelBuf.AddColumn(LocGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 GSTRegNo
        ExcelBuf.AddColumn(POno, FALSE, '', FALSE, FALSE, FALSE, '', 1);//59
        ExcelBuf.AddColumn(POdate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//60
        ExcelBuf.AddColumn(BlanketOrderNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//61
        ExcelBuf.AddColumn(BlanketOrderDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//62
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//63
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 2);//64
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//66
        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';//RecVendor."T.I.N. No.";
            CSTTINNo := '';// RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//69
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END;
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.", "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END
        ELSE
            IF recState.GET(/*"Purch. Cr. Memo Hdr.-ForReturn".state*/) THEN;

        ExcelBuf.AddColumn(recState.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//73
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//74
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//75
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//76
        ExcelBuf.AddColumn(tCommentCrMemo2, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//77
                                                                                                            /* ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Item Category Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//78 RSPLSUM 15Apr2020
                                                                                                             ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Purchaser Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//78 RSPLSUM 15Apr2020
                                                                                                             ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Department Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//78 RSPLSUM 15Apr2020
                                                                                                             ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//78 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."E-Inv Irn", FALSE, '', FALSE, FALSE, FALSE, '', 1);//78 RSPLSUM04Mar21
        ExcelBuf.AddColumn(-1 * TCSPCharges, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//79 DJ170421

    end;

    // [Scope('Internal')]
    procedure MakePurchReturnTotDataGrouping()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Vendor Posting Group", FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Vtext, FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ShipStateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(ShipStateName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(BillStateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(BillStateName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(VenGSTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn(ShipGSTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//10 11Aug2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Applies-to Doc. No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(PIH."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//11
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."GST Vendor Type", FALSE, '', TRUE, FALSE, FALSE, '', 1);//14 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn(-1 * vQtyGr, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//19
        vQtyGr := 0;//04Aug2017

        ExcelBuf.AddColumn(-1 * ExcessShortQty1, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//20
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Currency Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//23

        IF "Purch. Cr. Memo Hdr.-ForReturn"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(-1 * vCrLineAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(-1 * vCrExciseBase, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn(-1 * GSTBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//31 GSTBaseAmount
            GSTBaseAmt := 0;//02Aug2017

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//32
            ExcelBuf.AddColumn(vSGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn(vCGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn(vIGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn(vUTGSTRate, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//36 //UTGST %
            ExcelBuf.AddColumn(-1 * vCrBEDAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(-1 * vCrAEDAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0); //38
            ExcelBuf.AddColumn(-1 * TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn(-1 * TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn(-1 * TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn(-1 * TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            TSGST := 0;//02Aug2017
            TCGST := 0;//02Aug2017
            TIGST := 0;//02Aug2017
            TUTGST := 0;//02Aug2017

            ExcelBuf.AddColumn(-1 * AddExcise1, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//44 GSTShortageAmount

            IF (/*"Purch. Cr. Memo Line-ForReturn"."Excise Amount"*/0 <> 0) AND (vCrBEDAmt = 0) THEN BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + vCrExciseBase + AddExcise;
            END ELSE BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + vCrBEDAmt + vCrAEDAmt + AddExcise;
            END;

            ExcelBuf.AddColumn(-1 * BillAmtInclusivEcx1, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//45

            ExcelBuf.AddColumn(-1 * GTaxBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//46
            GTaxBaseAmt := 0;//03Aug2017

            ExcelBuf.AddColumn(TaxGroupCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn(TaxPer, FALSE, '', TRUE, FALSE, FALSE, '', 1);//48
            ExcelBuf.AddColumn(-1 * VCrTaxAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0); //49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//51
                                                                                // "Purch. Cr. Memo Hdr.-ForReturn".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Hdr.-ForReturn"."Amount to Vendor"*/0, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//52
        END ELSE BEGIN

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
            ExcelBuf.AddColumn(-1 * vCrLineAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//25
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//26 Insurance
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//27 Freight
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//28 Discount Charges

            ExcelBuf.AddColumn(-1 * (vCrLineAmt / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//29
            ExcelBuf.AddColumn(-1 * (vCrEcessAmt / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//30
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//31 GSTBaseAmount
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//32
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//33 //SGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//34 //CGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//35 //IGST %
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//36 //UTGST %
            ExcelBuf.AddColumn(-1 * (vCrBEDAmt / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//37
            ExcelBuf.AddColumn(-1 * (vCrAEDAmt / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//38
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//39 //SGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//40 //CGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//41 //IGST Amt
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//42 //UTGST Amt
            ExcelBuf.AddColumn(-1 * (AddExcise1 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//44 GSTShortageAmount

            IF (/*"Purch. Cr. Memo Line-ForReturn"."Excise Amount"*/0 <> 0) AND (vCrBEDAmt = 0) THEN BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + vCrExciseBase + AddExcise1;
            END ELSE BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + vCrBEDAmt + vCrAEDAmt + AddExcise1;
            END;

            ExcelBuf.AddColumn(-1 * (BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//45
            ExcelBuf.AddColumn(-1 * GTaxBaseAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//46
            GTaxBaseAmt := 0;//03Aug2017
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//48
            ExcelBuf.AddColumn(-1 * (VCrTaxAmt / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor"), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//51

            // "Purch. Cr. Memo Hdr.-ForReturn".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(-1 * /*"Purch. Cr. Memo Hdr.-ForReturn"."Amount to Vendor"*/ 0 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//52
        END;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//53 E-Way Bill No
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//54 E-Way Bill No
        ExcelBuf.AddColumn(RespCtr1.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LocName15, FALSE, '', TRUE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn(LocState15, FALSE, '', TRUE, FALSE, FALSE, '', 1);//58 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//57 GSTState Code
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//58 GSTRegNo
        ExcelBuf.AddColumn(POno, FALSE, '', TRUE, FALSE, FALSE, '', 1);//59
        ExcelBuf.AddColumn(POdate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//60
        ExcelBuf.AddColumn(BlanketOrderNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//61
        ExcelBuf.AddColumn(BlanketOrderDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//63
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 2);//64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//66

        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';// RecVendor."T.I.N. No.";
            CSTTINNo := '';//RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//69
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', TRUE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', TRUE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', TRUE, FALSE, FALSE, '', 1);//72
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//70
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        END;

        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.", "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END
        ELSE
            IF recState.GET(/*"Purch. Cr. Memo Hdr.-ForReturn".State*/) THEN;
        ExcelBuf.AddColumn(recState.Description, FALSE, '', TRUE, FALSE, FALSE, '', 1);//73

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//74
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//75

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//76
        ExcelBuf.AddColumn(tCommentCrMemo2, FALSE, '', TRUE, FALSE, FALSE, '', 1);//77
                                                                                  /* ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//78 RSPLSUM 15Apr2020
                                                                                   ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//79 RSPLSUM 15Apr2020
                                                                                   ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//80 RSPLSUM 15Apr2020
                                                                                   ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//81 RSPLSUM 15Apr2020*/

        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."E-Inv Irn", FALSE, '', TRUE, FALSE, FALSE, '', 1);//78 RSPLSUM04Mar21
        ExcelBuf.AddColumn(LineTCSPCharges, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//79 DJ170421

        CLEAR(vCrLineAmt);
        CLEAR(vCrBEDAmt);
        CLEAR(vCrAEDAmt);
        CLEAR(vCrExciseBase);
        CLEAR(VCrTaxAmt);
        CLEAR(vCrTaxBaseAmt);
        //<<

    end;

    // [Scope('Internal')]
    procedure MakePurchReturnGrandTotal()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//52
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//53
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //54
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //59
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //60
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //61
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //63
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //66
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //67
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //68
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //69
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //70
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //71
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //72
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //73
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //74
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //76
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //78 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //79 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //80 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//81 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//82 07Jul2018
        /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//83 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//84 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//85 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//86 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//83 RSPLSUM04Mar21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//84 DJ170421

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GRAND TOTAL(Return)', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '1', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '2', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '3', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '4', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '5', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '6', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '7', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '8', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '9', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '10', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //11 11Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //14 15Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 07Jul2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//07Jul2018
        ExcelBuf.AddColumn(-1 * rtQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0); //19
        ExcelBuf.AddColumn(-1 * rtExciseQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0); //20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn(-1 * totalfreight, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//27 //Freight
        ExcelBuf.AddColumn(-1 * totDisChgs, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//28 //Discount
        totalfreight := 0; //03Aug2017
        totDisChgs := 0; //03Aug2017

        ExcelBuf.AddColumn(-1 * rtAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0); //29
        ExcelBuf.AddColumn(-1 * rtExciseBaseAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//30
        ExcelBuf.AddColumn(-1 * TGSTBaseAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//31 GSTBaseAmt
        TGSTBaseAmt := 0;//02Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn(-1 * rtBEDAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//37
        ExcelBuf.AddColumn(-1 * rtAddDuty, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//38
        ExcelBuf.AddColumn(-1 * TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//39
        ExcelBuf.AddColumn(-1 * TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//40
        ExcelBuf.AddColumn(-1 * TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//41
        ExcelBuf.AddColumn(-1 * TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//42
        TTSGST := 0;//02Aug2017
        TTCGST := 0;//02Aug2017
        TTIGST := 0;//02Aug2017
        TTUTGST := 0;//02Aug2017

        ExcelBuf.AddColumn(-1 * TotalAddExcise1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//44 GSTShortageAmount
        ExcelBuf.AddColumn(-1 * TotalBillAmtInclusivEcx1, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//45
        //ExcelBuf.AddColumn(TotalTaxableAmount-ChargeItemTaxBaseAmt -rtTaxableAmt,FALSE,'',TRUE,FALSE,TRUE,'#,#0.00',0);//46
        ExcelBuf.AddColumn(-1 * TGTaxBaseAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//46 //03Aug2017
        TGTaxBaseAmt := 0;//04Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //48
        ExcelBuf.AddColumn(-1 * rtTaxAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//51
        ExcelBuf.AddColumn(-1 * rtGrossAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//52
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//59
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//61
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//63
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//64
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//66
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//67
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//68
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//69
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80 15Nov2017
        /*ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//81 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//82 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//83 RSPLSUM 15Apr2020
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//84 RSPLSUM 15Apr2020*/
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79 RSPLSUM04Mar21
        ExcelBuf.AddColumn(-1 * TotalTCSPCharges, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//80 DJ170421
        TotalTCSPCharges := 0; //03Aug2017

    end;
}


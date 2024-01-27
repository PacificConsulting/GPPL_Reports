report 50151 "Purchase Register"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Purch. Inv. Header"; "Purch. Inv. Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "Posting Date", "Location Code", "Vendor Posting Group", "Gen. Bus. Posting Group";
            dataitem("Purch. Inv. Line"; "Purch. Inv. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Quantity = FILTER(<> 0),
                                          "No." = FILTER(<> 74012210));
                RequestFilterFields = "Gen. Prod. Posting Group";

                trigger OnAfterGetRecord()
                begin
                    //IF  (PrinttoExcel) AND  (PurchaseInvoice) THEN //05May2017
                    BEGIN  //05May2017
                        CLEAR(BEDpercent);
                        /*  recExciseEntry.RESET;
                         recExciseEntry.SETRANGE(recExciseEntry."Document No.", "Purch. Inv. Line"."Document No.");
                         IF recExciseEntry.FINDFIRST THEN BEGIN
                             BEDpercent := recExciseEntry."BED %";
                         END; */

                        recItemCategory.RESET;
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
                        //recPRH.SETRANGE(recPRH."No.", "Purch. Inv. Line"."Receipt Document No.");
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
                        END;

                        CLEAR(DisChgs);
                        /* recPostedStrOrdDetails.RESET;
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Inv. Header"."No.");
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Purch. Inv. Line"."Line No.");
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code", 'DISCOUNT');
                        IF recPostedStrOrdDetails.FINDFIRST THEN
                            REPEAT
                                DisChgs := DisChgs + recPostedStrOrdDetails.Amount
                            UNTIL recPostedStrOrdDetails.NEXT = 0; */


                        IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)" THEN
                            ExcessShortQty += ExcessShortQty + "Purch. Inv. Line".Quantity;


                        IF ("Purch. Inv. Line".Type = "Purch. Inv. Line".Type::Item) THEN
                            TotalQtyofITEM := TotalQtyofITEM + "Purch. Inv. Line".Quantity;//09Mar2017
                                                                                           //TotalQtyofITEM += TotalQtyofITEM + "Purch. Inv. Line".Quantity;

                        //>>RSPl-Rahul**Migration**test3
                        vLineAmt += "Purch. Inv. Line"."Line Amount";
                        vExciseBaseAmt += 0;//"Purch. Inv. Line"."Excise Base Amount";
                        vBEDAmt += 0;// "Purch. Inv. Line"."BED Amount";
                        vEcessAmt += 0;//"Purch. Inv. Line"."eCess Amount";
                        vSheCessAmt += 0;//"Purch. Inv. Line"."SHE Cess Amount";
                        vADEAmt += 0;//"Purch. Inv. Line"."ADE Amount";
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
                         //RecDTE.SETRANGE("Document Line No.", "Line No.");
                         IF RecDTE.FINDFIRST THEN BEGIN
                             IF RecTG.GET(RecDTE."Tax Group Code") THEN;
                             IF RecDTE."Form Code" <> '' THEN
                                 TaxGroupCode := RecTG.Description + 'WITH FORM ' + RecDTE."Form Code"
                             ELSE
                                 TaxGroupCode := RecTG.Description;
                             TaxPer := FORMAT(RecDTE."Tax %") + '%'; */

                    END;
                    iF (Vendor."Vendor Posting Group" = 'FOREIGN') /*AND ("Purch. Inv. Header"."Form Code" = '')*/ THEN BEGIN
                        TaxGroupCode := 'NoTax Ag.Export';
                    END;


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
                                            Vtext := 'LOCAL'
                                        ELSE
                                            Vtext := 'INTER STATE';
                                END ELSE
                                    IF recPurchInv."Order Address Code" = '' THEN BEGIN
                                        IF RecVendor."State Code" = LocationState THEN
                                            Vtext := 'LOCAL'
                                        ELSE
                                            Vtext := 'INTER STATE';
                                    END;
                        END;
                    END ELSE
                        IF RecVendor."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            Vtext := 'IMPORT';



                    //MESSAGE('LOcation State %1 \ OrderCode %2 \ VText %3',LocationState,orderadd.State,Vtext);

                    //RSPL-Dhananjay END

                    IF PrinttoExcel AND Detail AND PurchaseInvoice THEN BEGIN
                        MakePurchInvDataBody;
                    END;
                    //END; //05May2017

                    //>>09Mar2017
                    IF "Purch. Inv. Header"."Currency Code" = '' THEN BEGIN
                        TotalAmount1 += "Purch. Inv. Line"."Line Amount";
                        TotalBillAmtInclusivEcx += BillAmtInclusivEcx;
                        //TotalGrossValue +=  "Purch. Inv. Line"."Amount To Vendor";
                        TotalNetValue += "Purch. Inv. Line".Amount;
                        TotalBedAmt += 0;// "Purch. Inv. Line"."BED Amount";
                        TotaleCessAmt += 0;// "Purch. Inv. Line"."eCess Amount";
                        TotalSHECess += 0;// "Purch. Inv. Line"."SHE Cess Amount";
                        TotalTaxBaseAmt += "Purch. Inv. Line".Amount + 0;//"Purch. Inv. Line"."Charges To Vendor";
                        //TotalTaxAmt     +=  "Tax Amount";//Sourav  110215
                        //totadditionalduty +=  "Purch. Inv. Line"."ADE Amount";
                        totalfreight += freightamt;
                        totDisChgs += DisChgs;


                        TotalDeliveryChrgs += DeliveryChrgs;
                        TotalTDSAmount += 0;// "Purch. Inv. Line"."TDS Amount";
                    END
                    ELSE BEGIN
                        TotalAmount1 += "Purch. Inv. Line"."Line Amount" / "Purch. Inv. Header"."Currency Factor";
                        TotalBillAmtInclusivEcx += BillAmtInclusivEcx / "Purch. Inv. Header"."Currency Factor";
                        TotalGrossValue += "Purch. Inv. Line"."Amount To Vendor" / "Purch. Inv. Header"."Currency Factor";
                        TotalNetValue += "Purch. Inv. Line".Amount / "Purch. Inv. Header"."Currency Factor";
                        TotalBedAmt += 0;// "Purch. Inv. Line"."BED Amount" / "Purch. Inv. Header"."Currency Factor";
                        TotaleCessAmt += 0;// "Purch. Inv. Line"."eCess Amount" / "Purch. Inv. Header"."Currency Factor";
                        TotalSHECess += 0;//"Purch. Inv. Line"."SHE Cess Amount" / "Purch. Inv. Header"."Currency Factor";
                        TotalTaxBaseAmt += 0;// "Purch. Inv. Line".Amount + "Purch. Inv. Line"."Charges To Vendor" / "Purch. Inv. Header"."Currency Factor";
                        TotalTaxAmt += 0;//"Tax Amount" / "Purch. Inv. Header"."Currency Factor";
                        //totadditionalduty +=  "Purch. Inv. Line"."ADE Amount"/"Purch. Inv. Header"."Currency Factor";
                        totalfreight += freightamt / "Purch. Inv. Header"."Currency Factor";
                        totDisChgs += DisChgs / "Purch. Inv. Header"."Currency Factor";
                        TotalAddExcise += AddExcise / "Purch. Inv. Header"."Currency Factor";
                        TotalTaxableAmount += 0;//"Purch. Inv. Line"."Tax Base Amount" / "Purch. Inv. Header"."Currency Factor";
                        TotalDeliveryChrgs += DeliveryChrgs / "Purch. Inv. Header"."Currency Factor";
                        TotalTDSAmount += 0;//"Purch. Inv. Line"."TDS Amount" / "Purch. Inv. Header"."Currency Factor";
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
                end;
            }

            trigger OnAfterGetRecord()
            begin

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
                /*  DetTaxEntry.RESET;
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
                 END; */
                IF (Vendor."Vendor Posting Group" = 'FOREIGN') /*AND ("Purch. Inv. Header"."Form Code" = '')*/ THEN BEGIN
                    InvTaxDescription := 'NoTax Ag.Export';
                END;


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
                        // ExciseSetup.RESET;
                        //ExciseSetup.SETRANGE("Excise Bus. Posting Group", "Purch. Inv. Line"."Excise Bus. Posting Group");
                        //ExciseSetup.SETRANGE("Excise Prod. Posting Group", "Excise Prod. Posting Group");
                        // IF ExciseSetup.FINDLAST THEN;

                        recItemCategory.RESET;
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

                        //CLEAR(TotalQtyofITEM1);
                        IF "Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::Item THEN
                            TotalQtyofITEM1 += TotalQtyofITEM1 + "Purch. Cr. Memo Line".Quantity;

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
 */
                    END
                    ELSE BEGIN
                        //  TaxGroupCode := 'NoTax Ag.Form' + "Purch. Cr. Memo Hdr."."Form Code";
                    END;
                    IF (Vendor."Vendor Posting Group" = 'FOREIGN') /*AND ("Purch. Inv. Header"."Form Code" = '')*/ THEN BEGIN
                        TaxGroupCode := 'NoTax Ag.Export';
                    END;

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
                                            Vtext := 'LOCAL'
                                        ELSE
                                            Vtext := 'INTER STATE';
                                END ELSE
                                    IF RecPurchCreLine."Order Address Code" = '' THEN BEGIN
                                        IF RecVendor."State Code" = LocationState THEN
                                            Vtext := 'LOCAL'
                                        ELSE
                                            Vtext := 'INTER STATE';
                                    END
                        END
                    END
                    ELSE
                        IF RecVendor."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            Vtext := 'IMPORT';



                    IF PrinttoExcel AND Detail AND PurchaseCreditMemo THEN BEGIN
                        MakePurchCrMemoDataBody;
                    END;
                    //MESSAGE('%1',Vtext);

                    //RSPL-Dhananjay END
                    //>>TEST10
                    vCrLineAmt += "Purch. Cr. Memo Line"."Line Amount";
                    vCrExciseBase += 0;// "Purch. Cr. Memo Line"."Excise Base Amount";
                    vCrBEDAmt += 0;// "Purch. Cr. Memo Line"."BED Amount";
                    vCrEcessAmt += 0;//"Purch. Cr. Memo Line"."eCess Amount";
                    vCrSheCessAmt += 0;// "Purch. Cr. Memo Line"."SHE Cess Amount";
                    vCrAEDAmt += 0;// "Purch. Cr. Memo Line"."ADE Amount";
                    VCrTaxAmt += 0;// "Purch. Cr. Memo Line"."Tax Amount";
                    vCrTaxBaseAmt += 0;//"Purch. Cr. Memo Line"."Tax Base Amount";
                                       //<<

                END;
                //end;

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
                        //TotalGrossValue1 += "Purch. Cr. Memo Line"."Amount To Vendor";
                        TotalNetValue1 += "Purch. Cr. Memo Line".Amount;
                        TotalExciseBaseAmt1 += 0;// "Purch. Cr. Memo Line"."Excise Base Amount";
                        TotalBedAmt1 += 0;//"Purch. Cr. Memo Line"."BED Amount";
                        TotaleCessAmt1 += 0;// "Purch. Cr. Memo Line"."eCess Amount";
                        TotalSHECess1 += 0;// "Purch. Cr. Memo Line"."SHE Cess Amount";
                        TotalTaxBaseAmt1 += "Purch. Cr. Memo Line".Amount + 0;//"Purch. Cr. Memo Line"."Charges To Vendor";
                        TotalTaxAmt1 += 0;// "Purch. Cr. Memo Line"."Tax Amount";
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
                        //TotalGrossValue1 += "Purch. Cr. Memo Line"."Amount To Vendor"/ "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalNetValue1 += "Purch. Cr. Memo Line".Amount / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalExciseBaseAmt1 += 0;//"Purch. Cr. Memo Line"."Excise Base Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalBedAmt1 += 0;//"Purch. Cr. Memo Line"."BED Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotaleCessAmt1 += 0;// "Purch. Cr. Memo Line"."eCess Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalSHECess1 += 0;// "Purch. Cr. Memo Line"."SHE Cess Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalTaxBaseAmt1 += "Purch. Cr. Memo Line".Amount + 0//"Purch. Cr. Memo Line"."Charges To Vendor"
                                            / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalTaxAmt1 += 0;// "Purch. Cr. Memo Line"."Tax Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        totadditionalduty1 += additnlamt1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        totalfreight1 += freightamt1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        totDisChgs1 += DisChgs1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalAddExcise1 += AddExcise1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        //TotalTaxableAmount1  +=  "Purch. Cr. Memo Line"."Tax Base Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalDeliveryChrgs1 += DeliveryChrgs1 / "Purch. Cr. Memo Hdr."."Currency Factor";
                    END;
                end;

                trigger OnPreDataItem()
                begin
                    /*
                    IF PrinttoExcel AND PurchaseReturn THEN
                    BEGIN
                     MakePurchReturnDataHeader
                    END;
                    */

                    //>>02May2017

                    //IF "Purch. Inv. Line"."Gen. Prod. Posting Group" <> '' THEN
                    //SETFILTER("Gen. Prod. Posting Group","Purch. Inv. Line"."Gen. Prod. Posting Group");

                    //<<02May2017

                    CPINVLINE := COUNT;//05May2017

                end;
            }

            trigger OnAfterGetRecord()
            begin

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

                /*  DetTaxEntry.RESET;
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

                /* PostedStrDetail.RESET;
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
                    "Purch. Cr. Memo Hdr.".SETRANGE("Purch. Cr. Memo Hdr."."Location Code", "Purch. Inv. Header"."Location Code");
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Vendor Posting Group") <> '' THEN BEGIN
                    "Purch. Cr. Memo Hdr.".SETRANGE("Purch. Cr. Memo Hdr."."Vendor Posting Group", "Purch. Inv. Header"."Vendor Posting Group");
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Gen. Bus. Posting Group") <> '' THEN BEGIN
                    "Purch. Cr. Memo Hdr.".SETRANGE("Purch. Cr. Memo Hdr."."Gen. Bus. Posting Group", "Purch. Inv. Header"."Gen. Bus. Posting Group");
                END;

                IF PrinttoExcel AND PurchaseCreditMemo THEN BEGIN
                    MakePurchCrMemoDataHeader;
                END;
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

                    recItemCategory.RESET;
                    recItemCategory.SETRANGE(recItemCategory.Code, "Purch. Cr. Memo Line-ForReturn"."Item Category Code");
                    IF recItemCategory.FINDFIRST THEN;




                    CLEAR(TaxPer);
                    CLEAR(TaxGroupCode);
                    /*   */
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
                                Vtext := 'LOCAL'
                            ELSE
                                Vtext := 'INTER STATE';
                        END
                    END
                    ELSE
                        IF RecVendor."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            Vtext := 'IMPORT';
                    //RSPL-Dhananjay END



                    IF PrinttoExcel AND Detail AND PurchaseReturn THEN BEGIN
                        MakePurchReturnDataBody;
                    END;


                    //MESSAGE('%1',Vtext);

                    //RSPL-Dhananjay END
                end;

                trigger OnPostDataItem()
                begin
                    IF vQtyGr <> 0 THEN //02May2017
                        IF PrinttoExcel AND PurchaseReturn THEN BEGIN
                            MakePurchReturnTotDataGrouping;
                        END;
                end;

                trigger OnPreDataItem()
                begin

                    //>>02May2017

                    //IF "Purch. Inv. Line"."Gen. Prod. Posting Group" <> '' THEN
                    //SETFILTER("Gen. Prod. Posting Group","Purch. Inv. Line"."Gen. Prod. Posting Group");

                    //<<02May2017
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
                    "Purch. Cr. Memo Hdr.-ForReturn".SETRANGE("Purch. Cr. Memo Hdr.-ForReturn"."Location Code", "Purch. Inv. Header"."Location Code");
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Vendor Posting Group") <> '' THEN BEGIN
                    "Purch. Cr. Memo Hdr.-ForReturn".SETRANGE
                    ("Purch. Cr. Memo Hdr.-ForReturn"."Vendor Posting Group", "Purch. Inv. Header"."Vendor Posting Group");
                END;

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Gen. Bus. Posting Group") <> '' THEN BEGIN
                    "Purch. Cr. Memo Hdr.-ForReturn".SETRANGE("Purch. Cr. Memo Hdr.-ForReturn"."Gen. Bus. Posting Group",
                                                              "Purch. Inv. Header"."Gen. Bus. Posting Group");
                END;


                IF PrinttoExcel AND PurchaseReturn THEN BEGIN
                    MakePurchReturnDataHeader;
                    CLEAR(ExcessShortQty1);
                    CLEAR(BillAmtInclusivEcx1);
                    CLEAR(DisChgs);
                    CLEAR(freightamt);
                    //CLEAR(TotalQtyofITEM1);
                END;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(PrinttoExcel; PrinttoExcel)
                {
                    ApplicationArea = all;
                    Caption = 'Print to Excel';
                }
                field(Detail; Detail)
                {
                    ApplicationArea = all;
                    Caption = 'Detail';
                }
                field(PurchaseInvoice; PurchaseInvoice)
                {
                    ApplicationArea = all;
                    Caption = 'Purchase Invoice';
                }
                field(PurchaseCreditMemo; PurchaseCreditMemo)
                {
                    ApplicationArea = all;
                    Caption = 'Purchase Credit Memo';
                }
                field(PurchaseReturn; PurchaseReturn)
                {
                    ApplicationArea = all;
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
        IF PrinttoExcel AND PurchaseInvoice AND PurchaseCreditMemo THEN BEGIN
            MakeExcelDataFooter;
        END;


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
        //DetTaxEntry: Record 16522;
        InvTaxType: Text[30];
        CrTaxType: Text[30];
        FormCode: Text[30];
        FormCode1: Text[30];
        totalfreight: Decimal;
        totDisChgs: Decimal;
        totadditionalduty: Decimal;
        //PostedStrDetail: Record 13760;
        exciseamt: Decimal;
        freightamt: Decimal;
        additnlamt: Decimal;
        totalfreight1: Decimal;
        totDisChgs1: Decimal;
        totalfreight2: Decimal;
        totDisChgs2: Decimal;
        TaxableAmount: Decimal;
        PostingGroup: Text[30];
        //PostedStrOrdrLineDetails: Record 13798;
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
        // ExciseSetup: Record 13711;
        TAXAmountvalue: Decimal;
        ExcessShortQty: Decimal;
        TotalQtyofITEM: Decimal;
        //recPostedStrOrdDetails: Record 13798;
        //recExciseEntry: Record 13712;
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
        //RecDTE: Record 16522;
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

    //  //[Scope('Internal')]
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
        ExcelBuf.AddInfoColumn(50151, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
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

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text002,Text003,COMPANYNAME,USERID);
        /* ExcelBuf.CreateBookAndOpenExcel('', Text002, '', '', '');
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text002);
        ExcelBuf.WriteSheet(Text002, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text002, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //RSPL-290115
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //RSPL-290115
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Purchase Register', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //----
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
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //RSPL-290115
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
        //>>09Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(DateFilter, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //<<
    end;

    // //[Scope('Internal')]
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL-290115
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
        //ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor Inv.No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor Inv. Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Party Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Party Name', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Discription', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Payment Terms', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty Received', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Short Qty', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Unit Cost', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('INR Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FC Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED %', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('E cess Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('SHE cess Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Additional Duty Amout', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Duty On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Bill Amount Inclusive Excise', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Discount Charges', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Frieght', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Rate of Tax', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Other Adjustment', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('TDS Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Chapter No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Resp. Centre', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Recieving Location', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Blanket Order No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Blanket Order Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRN No. / TRR No', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRN Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Category Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Category Name', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor VAT TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor CST TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('State', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);   //SN
        ExcelBuf.AddColumn('Bonded Rate', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('ExBond Rate', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Landed Cost', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL007
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL-290115
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-Dhananjay
    end;

    // //[Scope('Internal')]
    procedure MakePurchInvDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Inv. Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Invoice No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor Name", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Purch. Inv. Line".Type = "Purch. Inv. Line".Type::Item THEN BEGIN
            ExcelBuf.AddColumn(ROUND("Purch. Inv. Line".Quantity, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', ExcelBuf."Cell Type"::Number);//09Mar2017
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(ExcessShortQty, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', ExcelBuf."Cell Type"::Number);//09Mar2017
        ExcelBuf.AddColumn("Purch. Inv. Header"."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        IF "Purch. Inv. Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND("Purch. Inv. Line"."Line Amount", 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."Excise Base Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(BEDpercent, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."BED Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."eCess Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."SHE Cess Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."ADE Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(AddExcise, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

            /* IF ("Purch. Inv. Line"."Excise Amount" <> 0) AND ("Purch. Inv. Line"."BED Amount" = 0) THEN BEGIN
                 BillAmtInclusivEcx := "Purch. Inv. Line"."Line Amount" + "Purch. Inv. Line"."Excise Amount" +
                                               AddExcise;
             END
             ELSE BEGIN
                 BillAmtInclusivEcx := "Purch. Inv. Line"."Line Amount" + "Purch. Inv. Line"."BED Amount" +
                                       "Purch. Inv. Line"."eCess Amount" + "Purch. Inv. Line"."SHE Cess Amount" +
                                       "Purch. Inv. Line"."ADE Amount" + AddExcise;
             END; */
            ExcelBuf.AddColumn(BillAmtInclusivEcx, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);


            IF ("Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)") AND
               (("Purch. Inv. Line"."No." = 'AD 01') OR ("Purch. Inv. Line"."No." = 'BO 01') OR ("Purch. Inv. Line"."No." = 'CH 01'))
               /*AND ("Purch. Inv. Line"."Tax Amount" = 0)*/ THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//31
                ChargeItemTaxBaseAmt += 0;// "Purch. Inv. Line"."Tax Base Amount";
            END
            ELSE BEGIN
                ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Tax Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//31
            END;
            ExcelBuf.AddColumn(FORMAT(TaxGroupCode), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(FORMAT(TaxPer), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

            TAXAmountvalue := 0;// "Purch. Inv. Line"."Tax Amount";
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."Tax Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."TDS Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Inv. Line"."Amount To Vendor", 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

        END ELSE BEGIN
            ExcelBuf.AddColumn("Purch. Inv. Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("Purch. Inv. Line"."Direct Unit Cost", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Inv. Line"."Line Amount", 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Inv. Line"."Line Amount", 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."Excise Base Amount"*/0, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(BEDpercent, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."BED Amount"*/0, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."eCess Amount"*/0, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."SHE Cess Amount"*/0, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."ADE Amount"*/0, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(AddExcise, 0.01, '=') / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /* IF ("Purch. Inv. Line"."Excise Amount" <> 0) AND ("Purch. Inv. Line"."BED Amount" = 0) THEN BEGIN
                BillAmtInclusivEcx := "Purch. Inv. Line"."Line Amount" + "Purch. Inv. Line"."Excise Amount" +
                                              AddExcise;
            END
            ELSE BEGIN
                BillAmtInclusivEcx := "Purch. Inv. Line"."Line Amount" + "Purch. Inv. Line"."BED Amount" +
                                      "Purch. Inv. Line"."eCess Amount" + "Purch. Inv. Line"."SHE Cess Amount" +
                                      "Purch. Inv. Line"."ADE Amount" + AddExcise;
            END; */
            ExcelBuf.AddColumn(BillAmtInclusivEcx / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs, 0.01, '=') / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt, 0.01, '=') / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

            //CLEAR(ChargeItemTaxBaseAmt);
            IF ("Purch. Inv. Line".Type = "Purch. Inv. Line".Type::"Charge (Item)") AND
               (("Purch. Inv. Line"."No." = 'AD 01') OR ("Purch. Inv. Line"."No." = 'BO 01') OR
               ("Purch. Inv. Line"."No." = 'CH 01')) /*AND ("Purch. Inv. Line"."Tax Amount" = 0)*/ THEN BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                ChargeItemTaxBaseAmt += 0;//"Purch. Inv. Line"."Tax Base Amount";
            END
            ELSE BEGIN
                ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Tax Base Amount"*/0 / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            END;
            ExcelBuf.AddColumn(FORMAT(TaxGroupCode), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(FORMAT(TaxPer), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            TAXAmountvalue := 0;// "Purch. Inv. Line"."Tax Amount";
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."Tax Amount"*/0, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Line"."TDS Amount"*/0, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Inv. Line"."Amount To Vendor", 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END;
        ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Excise Prod. Posting Group"*/0, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RespCtr.Name, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(POno, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(POdate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BlanketOrderNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BlanketOrderDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Receipt Document No."*/0, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(GRNdate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Inv. Header"."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';// RecVendor."T.I.N. No.";
            CSTTINNo := '';//RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Order Address Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;
        //sn-BEGIN
        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Inv. Header"."Buy-from Vendor No.", "Purch. Inv. Header"."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END;
        /*  ELSE
             IF recState.GET("Purch. Inv. Header".State) THEN;
  */
        ExcelBuf.AddColumn(recState.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //sn-end
        ExcelBuf.AddColumn("Purch. Inv. Line"."Bonded Rate", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//09Mar2017 56
        ExcelBuf.AddColumn("Purch. Inv. Line"."Exbond Rate", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//09Mar2017 57
        ExcelBuf.AddColumn("Purch. Inv. Line"."Landed Cost", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//09Mar2017 58 //RSPL007
        ExcelBuf.AddColumn(tComment, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL-290115
        ExcelBuf.AddColumn(Vtext, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //RSPL-DHANANJAY
    end;

    // //[Scope('Internal')]
    procedure MakePurchInvTotalDataGrouping()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Inv. Header"."Posting Date", FALSE, '1', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."No.", FALSE, '2', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Posting Group", FALSE, '3', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*"Purch. Inv. Header"."Form Code"*/'', FALSE, '4', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Invoice No.", FALSE, '5', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Document Date", FALSE, '6', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor No.", FALSE, '7', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor Name", FALSE, '8', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '9', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '10', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Payment Terms Code", FALSE, '11', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '12', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalQtyofITEM, 0.01, '='), FALSE, '13', TRUE, FALSE, FALSE, '#####0.000', ExcelBuf."Cell Type"::Number);//09Mar2017
        ExcelBuf.AddColumn(ROUND(ExcessShortQty, 0.01, '='), FALSE, '14', TRUE, FALSE, FALSE, '#####0.000', ExcelBuf."Cell Type"::Number);//09Mar2017
        ExcelBuf.AddColumn("Purch. Inv. Header"."Currency Code", FALSE, '15', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '16', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '17', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '18', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        IF "Purch. Inv. Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '19', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(vLineAmt, 0.01, '='), FALSE, '20', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vExciseBaseAmt, 0.01, '='), FALSE, '21', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '22', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(vBEDAmt, 0.01, '='), FALSE, '23', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vEcessAmt, 0.01, '='), FALSE, '24', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vSheCessAmt, 0.01, '='), FALSE, '25', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vADEAmt, 0.01, '='), FALSE, '26', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number); //test7
            ExcelBuf.AddColumn(ROUND(AddExcise, 0.01, '='), FALSE, '27', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

            /* IF ("Purch. Inv. Line"."Excise Amount" <> 0) AND (vBEDAmt = 0) THEN BEGIN
                BillAmtInclusivEcx := vLineAmt + "Purch. Inv. Line"."Excise Amount" +
                                              AddExcise;
            END */
            //ELSE BEGIN
            BillAmtInclusivEcx := vLineAmt + vBEDAmt + vEcessAmt + vSheCessAmt + AddExcise + AddExcise;


            // END;
            ExcelBuf.AddColumn(BillAmtInclusivEcx, FALSE, '28', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs, 0.01, '='), FALSE, '29', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number); //Test2
            ExcelBuf.AddColumn(ROUND(freightamt, 0.01, '='), FALSE, '30', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /*
            IF ("Purch. Inv. Line".Type = "Purch. Inv. Line".Type :: "Charge (Item)") AND
               (("Purch. Inv. Line"."No." = 'AD 01') OR ("Purch. Inv. Line"."No." = 'BO 01') OR
               ("Purch. Inv. Line"."No." = 'CH 01')) AND ("Purch. Inv. Line"."Tax Amount" = 0) THEN
            BEGIN
             ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'31',ExcelBuf."Cell Type"::Text);//31
            END
            ELSE
            BEGIN
             //ExcelBuf.AddColumn(vTaxBaseAmt - ChargeItemTaxBaseAmt,FALSE,'31',TRUE,FALSE,FALSE,'30',ExcelBuf."Cell Type"::Number);
              ExcelBuf.AddColumn(vTaxBaseAmt,FALSE,'31',TRUE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);//31
            END;
            *///Commented 09Mar2017

            ExcelBuf.AddColumn(vTaxBaseAmt, FALSE, '31', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//31
            ExcelBuf.AddColumn(TaxGroupCode, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//Sourav  110215 //32
            ExcelBuf.AddColumn(TaxPer, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//Sourav  110215 //33
            ExcelBuf.AddColumn(ROUND(vTaxAmt, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);  //Sourav  110215 //34
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//35
            ExcelBuf.AddColumn(ROUND(vTDSAmt, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//36
            //"Purch. Inv. Header".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Header"."Amount to Vendor"*/0, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn(FORMAT(ROUND(vLineAmt, 0.01, '=')), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vLineAmt, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vExciseBaseAmt, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(vBEDAmt, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vEcessAmt, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vSheCessAmt, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vADEAmt, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(AddExcise, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

            /*  IF ("Purch. Inv. Line"."Excise Amount" <> 0) AND (vBEDAmt = 0) THEN BEGIN
                 BillAmtInclusivEcx := vLineAmt + "Purch. Inv. Line"."Excise Amount" +
                                               AddExcise;
             END
             ELSE BEGIN */
            BillAmtInclusivEcx := vLineAmt + vBEDAmt +
                                  vEcessAmt + vSheCessAmt +
                                  vADEAmt + AddExcise;
            //END;
            ExcelBuf.AddColumn(BillAmtInclusivEcx / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs, 0.01, '=') / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt, 0.01, '=') / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(vTaxBaseAmt
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//31

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(vTaxAmt, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(vTDSAmt, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            // "Purch. Inv. Header".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(ROUND(/*"Purch. Inv. Header"."Amount to Vendor"*/0, 0.01, '=')
            / "Purch. Inv. Header"."Currency Factor", FALSE, '36', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RespCtr.Name, FALSE, '', TRUE, FALSE, FALSE, '38', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(POno, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(POdate, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BlanketOrderNo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BlanketOrderDate, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*"Purch. Inv. Line"."Receipt Document No."*/'', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(GRNdate, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Inv. Header"."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';// RecVendor."T.I.N. No.";
            CSTTINNo := '';//RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Order Address Code", FALSE, '', TRUE, FALSE, FALSE, '50', ExcelBuf."Cell Type"::Text);
        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '51', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '52', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '53', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '51', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '52', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '53', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;
        //sn-BEGIN
        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Inv. Header"."Buy-from Vendor No.", "Purch. Inv. Header"."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END;
        /* ELSE
            IF recState.GET("Purch. Inv. Header".State) THEN; */
        ExcelBuf.AddColumn(recState.Description, FALSE, '', TRUE, FALSE, FALSE, '', 1);//24May2017
                                                                                       //ExcelBuf.AddColumn(recState.Description,FALSE,'',FALSE,FALSE,FALSE,'54',ExcelBuf."Cell Type"::Text);
                                                                                       //sn-end
        ExcelBuf.AddColumn('', FALSE, '55', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '56', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn("Purch. Inv. Line"."Landed Cost", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//RSPL007 //24May2017
                                                                                                              //ExcelBuf.AddColumn("Purch. Inv. Line"."Landed Cost",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);//RSPL007
        ExcelBuf.AddColumn(tComment, FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPL-290115 //24May2017
                                                                           //ExcelBuf.AddColumn(tComment,FALSE,'',FALSE,FALSE,FALSE,'58',ExcelBuf."Cell Type"::Text);//RSPL-290115
        ExcelBuf.AddColumn(Vtext, FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPL-DHANANJAY //24May2017
                                                                        //ExcelBuf.AddColumn(Vtext,FALSE,'',FALSE,FALSE,FALSE,'59',ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY

        //**Migration/Clear Variables test8
        TotalQty += TotalQtyofITEM;
        TotalExcessShortQty += ExcessShortQty;
        TotalExciseBaseAmt += vExciseBaseAmt;
        TotalAddExcise += AddExcise;
        TotalTaxableAmount += vTaxBaseAmt - ChargeItemTaxBaseAmt;
        TotalTaxAmt += vTaxAmt;//Sourav  110215
        //"Purch. Inv. Header".CALCFIELDS("Amount to Vendor");
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

    // //[Scope('Internal')]
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL
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
        ExcelBuf.AddColumn('', FALSE, '11', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalQty, 0.01, '='), FALSE, '12', TRUE, FALSE, TRUE, '#####0.000', ExcelBuf."Cell Type"::Number); //09Mar2017
        ExcelBuf.AddColumn(ROUND(TotalExcessShortQty, 0.01, '='), FALSE, '13', TRUE, FALSE, TRUE, '#####0.000', ExcelBuf."Cell Type"::Number); //09Mar2017 //test5
        ExcelBuf.AddColumn('', FALSE, '14', TRUE, FALSE, TRUE, '14', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '15', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '16', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '17', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '18', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalAmount1, 0.01, '='), FALSE, '19', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number); //20
        ExcelBuf.AddColumn(ROUND(TotalExciseBaseAmt, 0.01, '='), FALSE, '20', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//21
        ExcelBuf.AddColumn('', FALSE, '21', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalBedAmt, 0.01, '='), FALSE, '22', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotaleCessAmt, 0.01, '='), FALSE, '23', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalSHECess, 0.01, '='), FALSE, '24', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(totadditionalduty, 0.01, '='), FALSE, '25', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalAddExcise, 0.01, '='), FALSE, '26', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalBillAmtInclusivEcx, 0.01, '='), FALSE, '27', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(totDisChgs, 0.01, '='), FALSE, '28', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(totalfreight, 0.01, '='), FALSE, '29', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalTaxableAmount - ChargeItemTaxBaseAmt - rtTaxableAmt, 0.01, '='), FALSE, '30', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //09Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //09Mar2017
        //ExcelBuf.AddColumn(TaxGroupCode,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);//Sourav  110215
        //ExcelBuf.AddColumn(TaxPer,FALSE,'32',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);//Sourav  110215
        ExcelBuf.AddColumn(ROUND(TotalTaxAmt, 0.01, '='), FALSE, '33', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//Sourav  110215
        ExcelBuf.AddColumn('', FALSE, '34', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalTDSAmount, 0.01, '='), FALSE, '35', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalGrossValue, 0.01, '='), FALSE, '36', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //09Mar2017
        //ExcelBuf.AddColumn(vLandedCost,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);//RSPL007
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL-DHANANJAY
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL-DHANANJAY
    end;

    // //[Scope('Internal')]
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL-290115
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL-DHANANJAY
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Applies to Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Applies to Invoice Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Party Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Party Name', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Discription', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Payment Terms', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty Received', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Short Qty', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Unit Cost', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('INR Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FC Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED %', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('E cess Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('SHE cess Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Additional Duty Amout', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Duty On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Bill Amount Inclusive Excise', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Discount Charges', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Frieght', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Rate of Tax', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Other Adjustment', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('TDS Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Chapter No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Resp. Centre', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Recieving Location', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRA No. / TRR No', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Category Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Category Name', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor VAT TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor CST TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('State', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);    //sn
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-290115
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-Dhananjay
    end;

    // //[Scope('Internal')]
    procedure MakePurchCrMemoDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Vendor Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line"."Form Code"*/0, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(PIH."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor Name", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        IF "Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::Item THEN BEGIN
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line".Quantity, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//09Mar2017
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //TotalQtyofITEM1 += "Purch. Cr. Memo Line".Quantity; //R
        ExcelBuf.AddColumn(ROUND(ExcessShortQty1, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', ExcelBuf."Cell Type"::Number);//09Mar2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

        IF "Purch. Cr. Memo Hdr."."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line"."Line Amount", 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."Excise Base Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(/*ExciseSetup."BED %"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."BED Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."eCess Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."SHE Cess Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."ADE Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(AddExcise1, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /*  IF ("Purch. Cr. Memo Line"."Excise Amount" <> 0) AND ("Purch. Cr. Memo Line"."BED Amount" = 0) THEN BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line"."Line Amount" + "Purch. Cr. Memo Line"."Excise Amount" +
                                              AddExcise1;
            END 
            ELSE BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line"."Line Amount" + "Purch. Cr. Memo Line"."BED Amount" +
                                      "Purch. Cr. Memo Line"."eCess Amount" + "Purch. Cr. Memo Line"."SHE Cess Amount" +
                                      "Purch. Cr. Memo Line"."ADE Amount" + AddExcise1;
            END; */
            ExcelBuf.AddColumn(BillAmtInclusivEcx1, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs1, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt1, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line"."Tax Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(FORMAT(TaxGroupCode), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(FORMAT(TaxPer), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."Tax Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            //"Purch. Cr. Memo Hdr.".CALCFIELDS("Amount To Vendor");
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."Amount To Vendor"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);

        END
        ELSE BEGIN
            ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Unit Cost (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Direct Unit Cost", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line"."Line Amount", 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line"."Line Amount", 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."Excise Base Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(/*ExciseSetup."BED %"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."BED Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."eCess Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."SHE Cess Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."ADE Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(AddExcise1, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /* IF ("Purch. Cr. Memo Line"."Excise Amount" <> 0) AND ("Purch. Cr. Memo Line"."BED Amount" = 0) THEN BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line"."Line Amount" + "Purch. Cr. Memo Line"."Excise Amount" +
                                              AddExcise1;
            END
            ELSE BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line"."Line Amount" + "Purch. Cr. Memo Line"."BED Amount" +
                                       "Purch. Cr. Memo Line"."eCess Amount" + "Purch. Cr. Memo Line"."SHE Cess Amount" +
                                       "Purch. Cr. Memo Line"."ADE Amount" + AddExcise1;
            END; */
            ExcelBuf.AddColumn(BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs1, 0.01, '=') / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt1, 0.01, '=') / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line"."Tax Base Amount"*/0
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number); //test10
            ExcelBuf.AddColumn(CrTaxDescription, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line"."Tax %"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."Tax Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            //"Purch. Cr. Memo Hdr.".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Hdr."."Amount to Vendor"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END;

        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line"."Excise Prod. Posting Group"*/0, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RespCtr1.Name, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';// RecVendor."T.I.N. No.";
            CSTTINNo := '';// RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Order Address Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;

        //sn-BEGIN
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", "Purch. Cr. Memo Hdr."."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END;
        /*  ELSE
             IF recState.GET("Purch. Cr. Memo Hdr.".State) THEN; */
        ExcelBuf.AddColumn(recState.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //sn-end
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(tCommentCrMemo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL-290115
        ExcelBuf.AddColumn(Vtext, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL-290115
    end;

    // //[Scope('Internal')]
    procedure MakePurchCrMemoTotDataGrouping()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Posting Date", FALSE, '1', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."No.", FALSE, '2', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Vendor Posting Group", FALSE, '3', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Hdr."."Form Code"*/'', FALSE, '4', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Applies-to Doc. No.", FALSE, '5', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(PIH."Posting Date", FALSE, '7', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", FALSE, '8', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor Name", FALSE, '9', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '10', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '11', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Payment Terms Code", FALSE, '12', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '13', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalQtyofITEM1, 0.01, '='), FALSE, '14', TRUE, FALSE, FALSE, '#####0.000', 0);//09Mar2017
        ExcelBuf.AddColumn(ROUND(ExcessShortQty1, 0.01, '='), FALSE, '15', TRUE, FALSE, FALSE, '#####0.000', 0);//09Mar2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Currency Code", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '16', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '17', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '18', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        IF "Purch. Cr. Memo Hdr."."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '19', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(vCrLineAmt, 0.01, '='), FALSE, '20', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            TotalAmount2 += vCrLineAmt;
            ExcelBuf.AddColumn(ROUND(vCrExciseBase, 0.01, '='), FALSE, '21', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '22', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(vCrBEDAmt, 0.01, '='), FALSE, '23', TRUE, FALSE, FALSE, '#,###0.00', 0);//09Mar2017
            ExcelBuf.AddColumn(ROUND(vCrEcessAmt, 0.01, '='), FALSE, '24', TRUE, FALSE, FALSE, '#,###0.00', 0);//09Mar2017
            ExcelBuf.AddColumn(ROUND(vCrSheCessAmt, 0.01, '='), FALSE, '25', TRUE, FALSE, FALSE, '#,###0.00', 0);//09Mar2017
            ExcelBuf.AddColumn(ROUND(vCrAEDAmt, 0.01, '='), FALSE, '26', TRUE, FALSE, FALSE, '#,###0.00', 0);//09Mar2017
            ExcelBuf.AddColumn(ROUND(AddExcise, 0.01, '='), FALSE, '27', TRUE, FALSE, FALSE, '#,###0.00', 0);//09Mar2017
                                                                                                             /* IF ("Purch. Cr. Memo Line"."Excise Amount" <> 0) AND (vCrBEDAmt = 0) THEN BEGIN
                                                                                                                 BillAmtInclusivEcx1 := vCrLineAmt + "Purch. Cr. Memo Line"."Excise Amount" +
                                                                                                                                               AddExcise;
                                                                                                             END
                                                                                                             ELSE BEGIN */
            BillAmtInclusivEcx1 := vCrLineAmt + vCrBEDAmt +
                                   vCrEcessAmt + vCrSheCessAmt +
                                   vCrAEDAmt + AddExcise;
            //END;
            ExcelBuf.AddColumn(BillAmtInclusivEcx1, FALSE, '28', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs, 0.01, '='), FALSE, '29', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt, 0.01, '='), FALSE, '30', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(vCrTaxBaseAmt, FALSE, '31', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            TotalTaxableAmount1 += vCrTaxBaseAmt;
            ExcelBuf.AddColumn('', FALSE, '32', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '33', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(VCrTaxAmt, 0.01, '='), FALSE, '34', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '35', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '36', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line"."Amount To Vendor"*/0, 0.01, '='), FALSE, '37', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            TotalGrossValue1 += 0;// "Purch. Cr. Memo Line"."Amount To Vendor";
        END
        ELSE BEGIN
            ExcelBuf.AddColumn(ROUND(vCrLineAmt, 0.01, '='), FALSE, '19', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vCrLineAmt, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '20', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            TotalAmount2 += ROUND(vCrLineAmt, 0.01, '=') / "Purch. Cr. Memo Hdr."."Currency Factor";
            ExcelBuf.AddColumn(ROUND(vCrExciseBase, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '21', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '22', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(vCrBEDAmt, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '23', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vCrEcessAmt, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '24', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vCrSheCessAmt, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '25', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(vCrAEDAmt, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '26', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(AddExcise, 0.01, '='), FALSE, '27', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /* IF ("Purch. Cr. Memo Line"."Excise Amount" <> 0) AND (vCrBEDAmt = 0) THEN BEGIN
                BillAmtInclusivEcx1 := vCrLineAmt + "Purch. Cr. Memo Line"."Excise Amount" +
                                              AddExcise;
            END
            ELSE BEGIN */
            BillAmtInclusivEcx1 := vCrLineAmt + vCrBEDAmt +
                                   vCrEcessAmt + vCrSheCessAmt +
                                   vCrAEDAmt + AddExcise;
            //END;
            ExcelBuf.AddColumn(BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '28', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs, 0.01, '=') / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '29', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt, 0.01, '=') / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '30', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(vCrTaxBaseAmt
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '31', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            TotalTaxableAmount1 += vCrTaxBaseAmt / "Purch. Cr. Memo Hdr."."Currency Factor";
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '32', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '33', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(VCrTaxAmt, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '34', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '35', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '35', ExcelBuf."Cell Type"::Text);
            //"Purch. Cr. Memo Hdr.".CALCFIELDS("Amount to Vendor");
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Hdr."."Amount to Vendor"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr."."Currency Factor", FALSE, '36', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END;

        ExcelBuf.AddColumn('', FALSE, '37', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RespCtr1.Name, FALSE, '38', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Location Code", FALSE, '39', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '40', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';// RecVendor."T.I.N. No.";
            CSTTINNo := '';//RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Order Address Code", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;

        //sn-BEGIN
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", "Purch. Cr. Memo Hdr."."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END;
        /*  ELSE
             IF recState.GET("Purch. Cr. Memo Hdr.".State) THEN; */
        ExcelBuf.AddColumn(recState.Description, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//09Mar2017
                                                                                                                //sn-end
                                                                                                                //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(tCommentCrMemo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL-290115
        ExcelBuf.AddColumn(Vtext, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); //RSPL-DHANANJAY

        //>>**Tor
        //>>TEST9
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

    // //[Scope('Internal')]
    procedure MakePurchCrMemoGrandTotal()
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text); //09Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GRAND TOTAL(Cr. Memo)', FALSE, '1', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '2', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '3', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '4', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '5', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '6', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '7', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '8', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '9', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '10', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '11', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '12', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalQty1, 0.01, '='), FALSE, '13', TRUE, FALSE, TRUE, '#####0.000', 0);//09Mar2017
        ExcelBuf.AddColumn(ROUND(TotalExcessShortQty1, 0.01, '='), FALSE, '14', TRUE, FALSE, TRUE, '#####0.000', 0);//09Mar2017
        ExcelBuf.AddColumn('', FALSE, '15', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '16', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '17', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '18', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '19', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalAmount2, 0.01, '='), FALSE, '20', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalExciseBaseAmt1, 0.01, '='), FALSE, '21', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '22', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalBedAmt1, 0.01, '='), FALSE, '23', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotaleCessAmt1, 0.01, '='), FALSE, '24', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalSHECess1, 0.01, '='), FALSE, '25', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(totadditionalduty1, 0.01, '='), FALSE, '26', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalAddExcise1, 0.01, '='), FALSE, '27', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalBillAmtInclusivEcx1, 0.01, '='), FALSE, '28', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(totDisChgs1, 0.01, '='), FALSE, '29', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(totalfreight1, 0.01, '='), FALSE, '30', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ROUND(TotalTaxableAmount1, 0.01, '='), FALSE, '31', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '32', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '33', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalTaxAmt1, 0.01, '='), FALSE, '34', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '35', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '36', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(TotalGrossValue1, 0.01, '='), FALSE, '37', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
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
    end;

    // //[Scope('Internal')]
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
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text); //09Mar2017
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY //09Mar2017
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY //09Mar2017
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY //09Mar2017
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
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY //09Mar2017
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY //09Mar2017
    end;

    //  //[Scope('Internal')]
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL-290115
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //RSPL-DHANANJAY
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Applies to Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Applies to Invoice Date', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Party Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Party Name', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Discription', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Payment Terms', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty Received', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Short Qty', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Unit Cost', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('INR Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FC Cost Per Ltr / KG', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Base Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED %', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('E cess Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('SHE cess Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Additional Duty Amout', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Duty On Shortage Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Bill Amount Inclusive Excise', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Discount Charges', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Frieght', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Rate of Tax', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Other Adjustment', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('TDS Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Chapter No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Resp. Centre', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Recieving Location', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRA No. / TRR No', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Category Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Category Name', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor VAT TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor CST TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('State', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//09Mar2017
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-Dhananjay
    end;

    // //[Scope('Internal')]
    procedure MakePurchReturnDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Posting Date", FALSE, '*1', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."No.", FALSE, '*2', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Vendor Posting Group", FALSE, '*3', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Hdr.-ForReturn"."Form Code"*/'', FALSE, '*4', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Applies-to Doc. No.", FALSE, '*5', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(PIH."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '*6', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.", FALSE, '*7', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor Name", FALSE, '*8', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."No.", FALSE, '', FALSE, FALSE, FALSE, '*9', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn".Description, FALSE, '', FALSE, FALSE, FALSE, '*10', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Payment Terms Code", FALSE, '*11', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Unit of Measure Code", FALSE, '*12', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Purch. Cr. Memo Line-ForReturn".Type = "Purch. Cr. Memo Line-ForReturn".Type::Item THEN BEGIN
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line-ForReturn".Quantity, 0.001, '='), FALSE, '*13', FALSE, FALSE, FALSE, '#####0.000', 0);//13 09Mar2017
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '*13', ExcelBuf."Cell Type"::Text);
        rtQty += ROUND("Purch. Cr. Memo Line-ForReturn".Quantity, 0.01, '=');
        vQtyGr += ROUND("Purch. Cr. Memo Line-ForReturn".Quantity, 0.01, '=');
        ExcelBuf.AddColumn(ROUND(ExcessShortQty1, 0.01, '='), FALSE, '*14', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        rtExciseQty += ROUND(ExcessShortQty1, 0.01, '=');
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Currency Code", FALSE, '*15', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)", FALSE, '*16', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)", FALSE, '*17', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '*18', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '*19', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line-ForReturn"."Line Amount", 0.01, '='), FALSE, '*20', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtAmt += ROUND("Purch. Cr. Memo Line-ForReturn"."Line Amount", 0.01, '=');
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0, 0.01, '='), FALSE, '*21', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtExciseBaseAmt += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0);
            ExcelBuf.AddColumn(/*ExciseSetup."BED %"*/0, FALSE, '*22', FALSE, FALSE, FALSE, '#,####0.00', 0);//09Mar2017
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0, 0.01, '='), FALSE, '*23', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtBEDAmt += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0, 0.01, '=');
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."eCess Amount"*/0, 0.01, '='), FALSE, '*24', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtEcess += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."eCess Amount"*/0, 0.01, '=');
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount"*/0, 0.01, '='), FALSE, '*25', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtSheCess += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount"*/0, 0.01, '=');
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0, 0.01, '='), FALSE, '*26', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtAddDuty += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0, 0.01, '=');
            ExcelBuf.AddColumn(ROUND(AddExcise1, 0.01, '='), FALSE, '*27', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /*  IF ("Purch. Cr. Memo Line-ForReturn"."Excise Amount" <> 0) AND ("Purch. Cr. Memo Line-ForReturn"."BED Amount" = 0) THEN BEGIN
                 BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + "Purch. Cr. Memo Line-ForReturn"."Excise Amount" +
                                               AddExcise1;
             END
             ELSE BEGIN
                 BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + "Purch. Cr. Memo Line-ForReturn"."BED Amount" +
                                        "Purch. Cr. Memo Line-ForReturn"."eCess Amount" + "Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount" +
                                        "Purch. Cr. Memo Line-ForReturn"."ADE Amount" + AddExcise1;
             END; */
            ExcelBuf.AddColumn(BillAmtInclusivEcx1, FALSE, '*28', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtBillAmt += BillAmtInclusivEcx1;
            ExcelBuf.AddColumn(ROUND(DisChgs1, 0.01, '='), FALSE, '*29', FALSE, FALSE, FALSE, '#,####0.00', 0);//09Mar2017
            ExcelBuf.AddColumn(ROUND(freightamt1, 0.01, '='), FALSE, '*30', FALSE, FALSE, FALSE, '#,####0.00', 0);//09Mar2017
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount"*/0, FALSE, '*31', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtTaxableAmt += 0;// "Purch. Cr. Memo Line-ForReturn"."Tax Base Amount";
            ExcelBuf.AddColumn(CrTaxDescription, FALSE, '*32', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line-ForReturn"."Tax %"*/0, FALSE, '*33', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"*/0, 0.01, '='), FALSE, '*34', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtTaxAmt += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"*/0, 0.01, '=');
            ExcelBuf.AddColumn('', FALSE, '*35', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '*36', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0, 0.01, '='), FALSE, '*37', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtGrossAmt += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Unit Cost (LCY)", FALSE, '*17', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Direct Unit Cost", FALSE, '*18', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line-ForReturn"."Line Amount", 0.01, '='), FALSE, '*19', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line-ForReturn"."Line Amount", 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*20', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtAmt += ROUND("Purch. Cr. Memo Line-ForReturn"."Line Amount", 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";

            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*21', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtExciseBaseAmt += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
            ExcelBuf.AddColumn(/*ExciseSetup."BED %"*/0, FALSE, '*22', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*23', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtBEDAmt += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."eCess Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*24', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtEcess += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."eCess Amount"*/0, 0.01, '=')
             / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*25', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtSheCess += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount"*/0, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*26', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtAddDuty += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
            ExcelBuf.AddColumn(ROUND(AddExcise1, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '27', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /* IF ("Purch. Cr. Memo Line-ForReturn"."Excise Amount" <> 0) AND ("Purch. Cr. Memo Line-ForReturn"."BED Amount" = 0) THEN BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + "Purch. Cr. Memo Line-ForReturn"."Excise Amount" +
                                              AddExcise1;
            END
            ELSE BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + "Purch. Cr. Memo Line-ForReturn"."BED Amount" +
                                       "Purch. Cr. Memo Line-ForReturn"."eCess Amount" + "Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount" +
                                       "Purch. Cr. Memo Line-ForReturn"."ADE Amount" + AddExcise1;
            END; */
            ExcelBuf.AddColumn(BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*28', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtBillAmt += BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
            ExcelBuf.AddColumn(ROUND(DisChgs1, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*29', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt1, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*30', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount"*/0
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*31', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtTaxableAmt += 0;// "Purch. Cr. Memo Line-ForReturn"."Tax Base Amount" / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
            ExcelBuf.AddColumn(CrTaxDescription, FALSE, '', FALSE, FALSE, FALSE, '*32', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line-ForReturn"."Tax %"*/0, FALSE, '*33', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '34', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtTaxAmt += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"*/0, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
            ExcelBuf.AddColumn('', FALSE, '*35', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '*36', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '*37', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            rtGrossAmt += ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor";
        END;

        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line-ForReturn"."Excise Prod. Posting Group"*/'', FALSE, '*38', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RespCtr1.Name, FALSE, '*39', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Location Code", FALSE, '*40', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*41', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*42', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."Item Category Code", FALSE, '*43', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '*44', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';// RecVendor."T.I.N. No.";
            CSTTINNo := '';//RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '*45', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '*46', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code", FALSE, '47', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '*48', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '*49', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '*50', ExcelBuf."Cell Type"::Text);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '*48', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '*49', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '*50', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;

        //sn-BEGIN
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.",
            "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END;
        /* ELSE
            IF recState.GET("Purch. Cr. Memo Hdr.-ForReturn".State) THEN; */
        ExcelBuf.AddColumn(recState.Description, FALSE, '*51', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //sn-end
        // ExcelBuf.AddColumn('',FALSE,'*52',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(tCommentCrMemo2, FALSE, '*53', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL-290115
        ExcelBuf.AddColumn(Vtext, FALSE, '*54', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
    end;

    // //[Scope('Internal')]
    procedure MakePurchReturnTotDataGrouping()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."No.", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Vendor Posting Group", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Hdr.-ForReturn"."Form Code"*/'', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Applies-to Doc. No.", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(PIH."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor Name", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Payment Terms Code", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ROUND(vQtyGr, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//13 09Mar2017

        vQtyGr := 0;//09Mar2017

        ExcelBuf.AddColumn(ROUND(ExcessShortQty1, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//14 09Mar2017
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Currency Code", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        IF "Purch. Cr. Memo Hdr.-ForReturn"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line-ForReturn"."Line Amount", 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."eCess Amount"*/0, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount"*/0, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(AddExcise, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /*  IF ("Purch. Cr. Memo Line-ForReturn"."Excise Amount" <> 0) AND ("Purch. Cr. Memo Line-ForReturn"."BED Amount" = 0) THEN BEGIN
                 BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + "Purch. Cr. Memo Line-ForReturn"."Excise Amount" +
                                              AddExcise;
             END
             ELSE BEGIN
                 BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + "Purch. Cr. Memo Line-ForReturn"."BED Amount" +
                                       "Purch. Cr. Memo Line-ForReturn"."eCess Amount" + "Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount" +
                                       "Purch. Cr. Memo Line-ForReturn"."ADE Amount" + AddExcise;
             END; */
            ExcelBuf.AddColumn(BillAmtInclusivEcx1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount"*/0, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"*/0, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line-ForReturn"."Line Amount", 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND("Purch. Cr. Memo Line-ForReturn"."Line Amount", 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Excise Base Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."BED Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."eCess Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."ADE Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(AddExcise, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            /* IF ("Purch. Cr. Memo Line-ForReturn"."Excise Amount" <> 0) AND ("Purch. Cr. Memo Line-ForReturn"."BED Amount" = 0) THEN BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + "Purch. Cr. Memo Line-ForReturn"."Excise Amount" +
                                             AddExcise;
            END
            ELSE BEGIN
                BillAmtInclusivEcx1 := "Purch. Cr. Memo Line-ForReturn"."Line Amount" + "Purch. Cr. Memo Line-ForReturn"."BED Amount" +
                                      "Purch. Cr. Memo Line-ForReturn"."eCess Amount" + "Purch. Cr. Memo Line-ForReturn"."SHE Cess Amount" +
                                      "Purch. Cr. Memo Line-ForReturn"."ADE Amount" + AddExcise;
            END; */
            ExcelBuf.AddColumn(BillAmtInclusivEcx1 / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(DisChgs, 0.01, '=') / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(ROUND(freightamt, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(/*"Purch. Cr. Memo Line-ForReturn"."Tax Base Amount"*/0
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Tax Amount"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(ROUND(/*"Purch. Cr. Memo Line-ForReturn"."Amount To Vendor"*/0, 0.01, '=')
            / "Purch. Cr. Memo Hdr.-ForReturn"."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        END;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RespCtr1.Name, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        VATTINNo := '';
        CSTTINNo := '';
        RecVendor.RESET;
        RecVendor.SETRANGE(RecVendor."No.", "Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.");
        IF RecVendor.FINDFIRST THEN BEGIN
            VATTINNo := '';// RecVendor."T.I.N. No.";
            CSTTINNo := '';// RecVendor."C.S.T. No.";
        END;
        ExcelBuf.AddColumn(VATTINNo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(CSTTINNo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code", FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;

        //sn-BEGIN
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            IF recOrderAddr.GET("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.",
            "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code") THEN
                IF recState.GET(recOrderAddr.State) THEN;
        END;
        /*  ELSE
             IF recState.GET("Purch. Cr. Memo Hdr.-ForReturn".State) THEN; */
        ExcelBuf.AddColumn(recState.Description, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//09Mar2017
                                                                                                                //sn-end
                                                                                                                //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(tCommentCrMemo, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//09Mar2017
        ExcelBuf.AddColumn(Vtext, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY //09Mar2017
    end;

    // //[Scope('Internal')]
    procedure MakePurchReturnGrandTotal()
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GRAND TOTAL(Return)', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*2', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*3', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*4', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*5', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*6', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*7', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*8', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*9', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*10', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*11', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*12', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(rtQty, FALSE, '*13', TRUE, FALSE, TRUE, '#####0.000', 0);//13 09Mar2017
        ExcelBuf.AddColumn(rtExciseQty, FALSE, '*14', TRUE, FALSE, TRUE, '#####0.000', 0);//13 09Mar2017
        ExcelBuf.AddColumn('', FALSE, '*15', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*16', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*17', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*18', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*19', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(rtAmt, FALSE, '*20', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(rtExciseBaseAmt, FALSE, '*21', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '*22', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(rtBEDAmt, FALSE, '*23', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(rtEcess, FALSE, '*24', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(rtSheCess, FALSE, '*25', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(rtAddDuty, FALSE, '*26', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '*27', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(rtBillAmt, FALSE, '*28', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '*29', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*30', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(rtTaxableAmt, FALSE, '*31', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '*32', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*33', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(rtTaxAmt, FALSE, '*34', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '*35', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*36', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(rtGrossAmt, FALSE, '*37', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '*38', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*39', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*40', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*41', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*42', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*43', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*44', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*45', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*46', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*47', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*48', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*49', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*50', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '*51', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text); //SN
        //ExcelBuf.AddColumn('',FALSE,'*52',TRUE,FALSE,TRUE,'',ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
        ExcelBuf.AddColumn('', FALSE, '*53', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
        ExcelBuf.AddColumn('', FALSE, '*54', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);//RSPL-DHANANJAY
    end;
}


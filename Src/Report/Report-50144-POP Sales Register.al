report 50144 "POP Sales Register"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/POPSalesRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Location; Location)
        {
            DataItemTableView = SORTING("State Code", Code);
            RequestFilterFields = "State Code", "Code";
            dataitem("Sales Invoice Header"; "Sales Invoice Header")
            {
                DataItemLink = "Location Code" = FIELD(Code);
                DataItemTableView = SORTING("Location Code", "Posting Date", "No.");
                RequestFilterFields = "Sell-to Customer No.", "Shortcut Dimension 1 Code", "Shortcut Dimension 2 Code", "Salesperson Code", "Gen. Bus. Posting Group", "Customer Posting Group"/*, "Excise Bus. Posting Group"*/, "No.";
                dataitem("Sales Invoice Line"; "Sales Invoice Line")
                {
                    DataItemLink = "Document No." = FIELD("No.");
                    DataItemTableView = SORTING("Document No.", "Line No.")
                                        WHERE(Quantity = FILTER(<> 0),
                                              "Item Category Code" = FILTER('CAT14'));
                    RequestFilterFields = Type, "No.";

                    trigger OnAfterGetRecord()
                    begin

                        //>>1

                        IF RecCust2.GET("Sales Invoice Line"."Sell-to Customer No.") THEN;
                        UPricePerLT := 0;
                        UPrice := 0;
                        FOCAmount := 0;
                        DISCOUNTAmount := 0;
                        TransportSubsidy := 0;
                        CashDiscount := 0;
                        ENTRYTax := 0;
                        CessAmount := 0;
                        FRIEGHTAmount := 0;
                        vFrieghtNonTx := 0;  //RSPL008 FRIEGHT NON Taxable Amount
                        AddTaxAmt := 0;
                        BRTNo := '';
                        MRPPricePerLt := 0;
                        ListPricePerlt := 0;
                        StateDiscount := 0;
                        SurChargeInterest := 0;
                        FCFRIEGHTAmount := 0; // EBT MILAN

                        CLEAR(ProductGroup);     //---004--RSPL
                        CLEAR(ProductGroupName);  //--004--RSPL

                        FCGROSSINV := 0;//"Sales Invoice Line"."Amount To Customer";  // EBT MILAN
                                        //totalFCGROSSINV += FCGROSSINV; // EBT MILAN
                                        //GtotalFCGROSSINV += FCGROSSINV; // EBT MILAN


                        IF "Sales Invoice Line".Type = "Sales Invoice Line".Type::"G/L Account" THEN BEGIN
                            CurrReport.SKIP;
                        END;


                        // IF "Sales Invoice Line"."Abatement %" <> 0 THEN
                        MRPPricePerLt := 0;//"Sales Invoice Line"."MRP Price";

                        //IF "Sales Invoice Line"."Price Inclusive of Tax" THEN BEGIN
                        ListPricePerlt := "Sales Invoice Line"."Unit Price Incl. of Tax" / "Sales Invoice Line"."Qty. per Unit of Measure";
                        UPricePerLT := /*"Sales Invoice Line"."MRP Price"*/0 - ((/*"Sales Invoice Line"."MRP Price"*/0 * /*"Sales Invoice Line"."Abatement %"*/0) / 100);
                        UPrice := UPricePerLT * "Sales Invoice Line"."Qty. per Unit of Measure";
                        INRprice := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Quantity (Base)" * "Sales Invoice Line".Quantity;
                        INRamt := (ListPricePerlt + ProdDiscPerlt) * "Sales Invoice Line"."Quantity (Base)" - 0;
                        //"Sales Invoice Line"."Excise Amount";
                        // END
                        // ELSE 

                        BEGIN
                            ListPricePerlt := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            UPricePerLT := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            UPrice := "Sales Invoice Line"."Unit Price";
                            INRprice := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Quantity (Base)" * "Sales Invoice Line".Quantity;
                            INRamt := UPrice * "Sales Invoice Line".Quantity;
                        END;

                        //------004-------Start--------
                        recDefaultDimension.RESET;
                        recDefaultDimension.SETRANGE(recDefaultDimension."Table ID", 27);
                        recDefaultDimension.SETRANGE(recDefaultDimension."No.", "Sales Invoice Line"."No.");
                        recDefaultDimension.SETFILTER(recDefaultDimension."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                        IF recDefaultDimension.FINDFIRST THEN BEGIN
                            ProductGroup := recDefaultDimension."Dimension Value Code";
                            recDimensionVal.RESET;
                            recDimensionVal.SETFILTER(recDimensionVal."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                            recDimensionVal.SETRANGE(recDimensionVal.Code, ProductGroup);
                            IF recDimensionVal.FINDFIRST THEN
                                ProductGroupName := recDimensionVal.Name;
                        END;

                        //-----004--------END----------


                        /*  ExciseSetup.RESET;
                         //ExciseSetup.SETRANGE("Excise Bus. Posting Group", "Sales Invoice Line"."Excise Bus. Posting Group");
                         //ExciseSetup.SETRANGE("Excise Prod. Posting Group", "Excise Prod. Posting Group");
                         IF ExciseSetup.FINDLAST THEN; */

                        /* StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'FOC', 'FOC DIS'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                FOCAmount += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2|%3|%4|%5|%6|%7|%8|%9|%10', 'ADD-DIS', 'AVP-SP-SAN',
                                                    'MONTH-DIS', 'NEW-PRI-SP', 'SPOT-DISC', 'VAT-DIF-DI', 'T2 DISC', 'DISCOUNT'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                DISCOUNTAmount += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

                        /*  StructureDetailes.RESET;
                           StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                          StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                          StructureDetailes.SETRANGE("Item No.", "No.");
                          StructureDetailes.SETRANGE("Line No.", "Line No.");
                          StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                          StructureDetailes.SETRANGE("Tax/Charge Group", 'STATE DISC'); 
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 StateDiscount += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0; */


                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'TPTSUBSIDY', 'TRANSSUBS'); 
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 TransportSubsidy += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0; */

                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'CASHDISC');
                         StructureDetailes.SETFILTER(StructureDetailes."Calculation Value", '<%1', 0); 
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 CashDiscount += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0; */

                        //Here-
                        //StructureDetailes.RESET;
                        /*  StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'CASHDISC'); 
                         StructureDetailes.SETFILTER(StructureDetailes."Calculation Value", '>%1', 0);
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 SurChargeInterest += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0; */

                        /* tructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'CCFC'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                SurChargeInterest += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'ENTRYTAX'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                ENTRYTax += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0;
 */
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'FREIGHT');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                //RSPL008 >>>> FRIEGHT NON Taxable Amount
                                //StructureDetailes.CALCFIELDS("Structure Description");
                                IF StructureDetailes."Structure Description" = 'Freight (Non-Taxable)' THEN BEGIN
                                    IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                        vFrieghtNonTx += StructureDetailes."Amount (LCY)"
                                    ELSE
                                        vFrieghtNonTx += StructureDetailes.Amount;
                                END ELSE BEGIN
                                    FRIEGHTAmount += StructureDetailes.Amount;
                                END;
                                //RSPL008 <<<< FRIEGHT NON Taxable Amount
                                FCFRIEGHTAmount += StructureDetailes.Amount;  // EBT MILAN
                            UNTIL StructureDetailes.NEXT = 0; */


                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                        CLEAR("Addl.Tax&CessRate");
                        /* recDetailedTaxEntry.RESET;
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Invoice Line"."Document No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."No.", "Sales Invoice Line"."No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Invoice Line"."Tax Area Code");
                        recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Invoice Line"."Tax Area Code");
                        IF recDetailedTaxEntry.FINDFIRST THEN BEGIN
                            "Addl.Tax&CessRate" := recDetailedTaxEntry."Tax %"
                        END; */


                        CLEAR("Addl.Tax&Cess");
                        /* recDetailedTaxEntry.RESET;
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Invoice Line"."Document No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."No.", "Sales Invoice Line"."No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Invoice Line"."Tax Area Code");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document Line No.", "Sales Invoice Line"."Line No.");
                        recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Invoice Line"."Tax Area Code");
                        IF recDetailedTaxEntry.FINDFIRST THEN
                            REPEAT
                                "Addl.Tax&Cess" := recDetailedTaxEntry."Tax Amount";
                            UNTIL recDetailedTaxEntry.NEXT = 0; */


                        CLEAR(TaxRate);
                        /* recDetailedTaxEntry.RESET;
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Invoice Line"."Document No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."No.", "Sales Invoice Line"."No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Invoice Line"."Tax Area Code");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Jurisdiction Code", "Sales Invoice Line"."Tax Area Code");
                        IF recDetailedTaxEntry.FINDFIRST THEN BEGIN
                            TaxRate := recDetailedTaxEntry."Tax %"
                        END; */


                        CLEAR(TaxAmt);
                        /* recDetailedTaxEntry.RESET;
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Invoice Line"."Document No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."No.", "Sales Invoice Line"."No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document Line No.", "Sales Invoice Line"."Line No.");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Invoice Line"."Tax Area Code");
                        recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Jurisdiction Code", "Sales Invoice Line"."Tax Area Code");
                        IF recDetailedTaxEntry.FINDFIRST THEN
                            REPEAT
                                TaxAmt := recDetailedTaxEntry."Tax Amount";
                            UNTIL recDetailedTaxEntry.NEXT = 0; */

                        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
                            TaxAmt := 0;// "Sales Invoice Line"."Tax Amount";
                            TaxRate := 0;// "Sales Invoice Line"."Tax %";
                        END;
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END

                        //CLEAR("Addl.Tax&CessRate");
                        /*StructureDetailes.RESET;
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                         StructureDetailes.SETFILTER("Tax/Charge Code", '%1|%2|%3', 'CESS', 'C1', 'ADDL TAX'); 
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            "Addl.Tax&CessRate" := StructureDetailes."Calculation Value";
                        END;*/

                        //CLEAR("Addl.Tax&Cess");
                        /*StructureDetailes.RESET;
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                         StructureDetailes.SETFILTER("Tax/Charge Code", '%1|%2', 'CESS', 'C1');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                //CessAmount += StructureDetailes.Amount;
                                "Addl.Tax&Cess" += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

                        //CLEAR("Addl.Tax&Cess");
                        /*StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                         StructureDetailes.SETRANGE("Tax/Charge Code", 'ADDL TAX'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                //AddTaxAmt += StructureDetailes.Amount;
                                "Addl.Tax&Cess" += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0;*/

                        ValueEntry.RESET;
                        ValueEntry.SETRANGE(ValueEntry."Source Code", 'SALES');
                        ValueEntry.SETRANGE(ValueEntry."Document No.", "Sales Invoice Line"."Document No.");
                        ValueEntry.SETRANGE(ValueEntry."Item No.", "Sales Invoice Line"."No.");
                        ValueEntry.SETRANGE(ValueEntry."Document Line No.", "Sales Invoice Line"."Line No.");
                        IF ValueEntry.FIND('-') THEN BEGIN
                            ILEEntryNo := ValueEntry."Item Ledger Entry No.";

                            CLEAR("ItemApplEntryNo.");
                            ItemApplEntry.RESET;
                            ItemApplEntry.SETRANGE(ItemApplEntry."Outbound Item Entry No.", ILEEntryNo);
                            IF ItemApplEntry.FINDFIRST THEN
                                "ItemApplEntryNo." := ItemApplEntry."Inbound Item Entry No.";

                            ILE.RESET;
                            ILE.SETRANGE(ILE."Entry No.", "ItemApplEntryNo.");
                            IF ILE.FINDFIRST THEN BEGIN
                                TRRNo := ILE."Document No.";
                                TransferReceipt.RESET;
                                TransferReceipt.SETRANGE(TransferReceipt."No.", TRRNo);
                                IF TransferReceipt.FINDFIRST THEN BEGIN
                                    TransferShipment.RESET;
                                    TransferShipment.SETRANGE(TransferShipment."Transfer Order No.", TransferReceipt."Transfer Order No.");
                                    IF TransferShipment.FINDFIRST THEN
                                        BRTNo := TransferShipment."No.";
                                END;
                            END;
                        END;

                        CLEAR(SalesGroup);

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Invoice Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALEGROUP');
                        IF recDimSet.FINDFIRST THEN BEGIN

                            recDimSet.CALCFIELDS("Dimension Value Name");
                            SalesGroup := recDimSet."Dimension Value Name";

                        END;
                        /*
                        recPostedDocDim.RESET;
                        recPostedDocDim.SETRANGE(recPostedDocDim."Table ID",113);
                        recPostedDocDim.SETRANGE(recPostedDocDim."Document No.","Sales Invoice Line"."Document No.");
                        recPostedDocDim.SETRANGE(recPostedDocDim."Line No.","Sales Invoice Line"."Line No.");
                        recPostedDocDim.SETFILTER(recPostedDocDim."Dimension Code",'SALEGROUP');
                        IF recPostedDocDim.FINDFIRST THEN
                        BEGIN
                         recDimVale.RESET;
                         recDimVale.SETRANGE(recDimVale."Dimension Code",'SALEGROUP');
                         recDimVale.SETRANGE(recDimVale.Code,recPostedDocDim."Dimension Value Code");
                         IF recDimVale.FINDFIRST THEN
                         BEGIN
                          SalesGroup := recDimVale.Name;
                         END;
                        END;
                        *///Commented 27Apr2017

                        CLEAR(LineSalespersonCode);
                        CLEAR(LineSalespersonName);

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Invoice Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALESPERSON');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            LineSalespersonCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            LineSalespersonName := recDimSet."Dimension Value Name";

                        END;

                        /*
                        recPostedDocumentDimension.RESET;
                        recPostedDocumentDimension.SETRANGE(recPostedDocumentDimension."Table ID",113);
                        recPostedDocumentDimension.SETRANGE(recPostedDocumentDimension."Document No.","Sales Invoice Line"."Document No.");
                        recPostedDocumentDimension.SETRANGE(recPostedDocumentDimension."Line No.","Sales Invoice Line"."Line No.");
                        recPostedDocumentDimension.SETFILTER(recPostedDocumentDimension."Dimension Code",'SALESPERSON');
                        IF recPostedDocumentDimension.FINDFIRST THEN
                        BEGIN
                        LineSalespersonCode := recPostedDocumentDimension."Dimension Value Code";
                        
                         recDimensionValue.RESET;
                         recDimensionValue.SETRANGE(recDimensionValue."Dimension Code",'SALESPERSON');
                         recDimensionValue.SETRANGE(recDimensionValue.Code,recPostedDocumentDimension."Dimension Value Code");
                         IF recDimensionValue.FINDFIRST THEN
                         BEGIN
                          LineSalespersonName := recDimensionValue.Name;
                         END;
                        END;
                        */

                        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
                            "Sales Invoice Line".Quantity := 0;
                            "Sales Invoice Line"."Quantity (Base)" := 0;
                            TotalQty := 0;
                            TotalQtyBase := 0;
                        END;

                        TotalQty += "Sales Invoice Line".Quantity;
                        TotalQtyBase += "Sales Invoice Line"."Quantity (Base)";
                        TotalAmount += "Sales Invoice Line".Amount;
                        TotalLineDisc += "Sales Invoice Line"."Line Discount Amount";
                        TotINRval += "Sales Invoice Line".Amount;
                        TotalExciseBaseAmount += 0;//"Sales Invoice Line"."Excise Base Amount";
                        TotalBEDAmount += 0;// "Sales Invoice Line"."BED Amount";
                        TotaleCessAmount += 0;//"Sales Invoice Line"."eCess Amount";
                        TotalSHECessAmount += 0;//"Sales Invoice Line"."SHE Cess Amount";
                        TotalExciseAmount += 0;// "Sales Invoice Line"."Excise Amount";
                        TotalFOCAmount += FOCAmount;
                        TotalStateDiscount += StateDiscount;
                        TotalDISCOUNTAmount += DISCOUNTAmount;
                        TotalTransportSubsidy += TransportSubsidy;
                        TotalCashDiscount += CashDiscount;
                        TotalENTRYTax += ENTRYTax;
                        TotalCessAmount += CessAmount;
                        //TotalFRIEGHTAmount += FRIEGHTAmount;
                        GfreightNonTxAmount += vFrieghtNonTx;//RSPL008
                        TotalFRIEGHTAmount += 0;// "Sales Invoice Line"."Charges To Customer";
                        TotalAddTaxAmt += AddTaxAmt;
                        TotalTaxBAse += 0;// "Sales Invoice Line"."Tax Base Amount";
                        tfreightamt += freightamt;

                        tFCFRIEGHTAmount += FCFRIEGHTAmount;  // EBT MILAN
                                                              //gtotalFCFRIEGHTAmount += tFCFRIEGHTAmount; //EBT MILAN

                        // RSPL-Ragni- Calculate total amount including excise
                        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                            IF /*("Sales Invoice Line"."Price Inclusive of Tax"= TRUE) AND */(MRPPricePerLt <> 0) THEN BEGIN
                                AmtIncludingExcise += ((UPrice * "Sales Invoice Line".Quantity) + /*"Sales Invoice Line"."Excise Amount"*/0) + ProdDisc;
                            END
                            ELSE BEGIN
                                AmtIncludingExcise += 0/* "Sales Invoice Line"."Amount Including Excise"*/ + ProdDisc;
                            END;
                        END
                        ELSE BEGIN
                            IF /*("Sales Invoice Line"."Price Inclusive of Tax" = TRUE) AND*/(MRPPricePerLt <> 0) THEN BEGIN
                                AmtIncludingExcise += (((UPrice * "Sales Invoice Line".Quantity) + /*"Sales Invoice Line"."Excise Amount"*/0) + ProdDisc)
                                / "Sales Invoice Header"."Currency Factor";
                            END
                            ELSE BEGIN
                                AmtIncludingExcise += (/*"Sales Invoice Line"."Amount Including Excise"*/0 + ProdDisc)
                                / "Sales Invoice Header"."Currency Factor";
                            END;
                        END;
                        // RSPL- Ragni- Calculate total amount including excise

                        IF "Sales Invoice Header"."Currency Code" <> '' THEN BEGIN
                            totalFCGROSSINV += FCGROSSINV;         // EBT MILAN
                            GtotalFCGROSSINV += totalFCGROSSINV;   // EBT MILAN
                        END;




                        IF "Sales Invoice Line"."Tax Liable" = TRUE THEN
                            TotalTaxBAse1 += 0;//"Sales Invoice Line"."Tax Base Amount";

                        TotalTaxAmount += 0;// "Sales Invoice Line"."Tax Amount";
                        TotalAmtCust += 0;// "Sales Invoice Line"."Amount To Customer";

                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                        TotalTaxAmt += TaxAmt;
                        "TotalAddl.Tax&Cess" += "Addl.Tax&Cess";
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END

                        TSurChargeInterest += SurChargeInterest;

                        GtotalQty += TotalQty;
                        GtotalQtyBase += TotalQtyBase;
                        GtotalAmount += TotalAmount;
                        GtotalLineDisc += TotalLineDisc;

                        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                            GTotalExciseBaseAmount += TotalExciseBaseAmount;
                        END
                        ELSE BEGIN
                            GTotalExciseBaseAmount += TotalExciseBaseAmount / "Sales Invoice Header"."Currency Factor";
                        END;

                        GTotalBEDAmount += TotalBEDAmount;
                        GTotaleCessAmount += TotaleCessAmount;
                        GTotalSHECessAmount += TotalSHECessAmount;
                        GTotalExciseAmount += TotalExciseAmount;
                        GTotalFOCAmount += TotalFOCAmount;
                        GrandTotalStateDiscount += TotalStateDiscount;
                        GTotalDISCOUNTAmount += TotalDISCOUNTAmount;
                        GTotalTransportSubsidy += TotalTransportSubsidy;
                        GTotalCashDiscount += TotalCashDiscount;
                        GTotalENTRYTax += TotalENTRYTax;
                        GTotalCessAmount += TotalCessAmount;
                        GTotalFRIEGHTAmount += TotalFRIEGHTAmount;
                        GTotalAddTaxAmt += TotalAddTaxAmt;
                        vGrpFrieghtNonTx += vFrieghtNonTx;  //RSPL008

                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                        "GTotalAddl.Tax&Cess" += "TotalAddl.Tax&Cess";
                        GTotalTaxAmt += TotalTaxAmt;
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
                        GSurChargeInterest += TSurChargeInterest;

                        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                            GTotalTaxBAse += TotalTaxBAse;
                        END
                        ELSE BEGIN
                            GTotalTaxBAse += TotalTaxBAse / "Sales Invoice Header"."Currency Factor";
                        END;

                        GTotalTaxAmount += TotalTaxAmount;

                        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                            GTotalAmtCust := GTotalAmtCust + TotalAmtCust;
                        END
                        ELSE BEGIN
                            GTotalAmtCust := GTotalAmtCust + TotalAmtCust / "Sales Invoice Header"."Currency Factor";
                        END;

                        //ebt mILAN
                        /*
                        IF "Sales Invoice Header"."Currency Code" = '' THEN
                        BEGIN
                         GfreightAmount += freightamt;
                        END
                        ELSE
                        BEGIN
                         //GfreightAmount += freightamt / "Sales Invoice Header"."Currency Factor" ;
                         GfreightAmount += freightamt;
                        END;
                        */
                        iF "Sales Invoice Header"."Currency Code" <> '' THEN BEGIN
                            gtotalFCFRIEGHTAmount += tFCFRIEGHTAmount; // EBT MILAN
                        END;


                        CLEAR(ProdDisc);
                        CLEAR(ProdDiscPerlt);
                        MrpMaster.RESET;
                        MrpMaster.SETRANGE(MrpMaster."Item No.", "Sales Invoice Line"."No.");
                        MrpMaster.SETRANGE(MrpMaster."Lot No.", "Sales Invoice Line"."Lot No.");
                        IF MrpMaster.FINDFIRST THEN BEGIN
                            ProdDiscPerlt := MrpMaster."National Discount";
                            ProdDisc := ProdDiscPerlt * "Sales Invoice Line"."Quantity (Base)";
                        END;

                        // VK robosoft - BEGIN
                        CSOAppDate := 0D;
                        recSalesApprovalEntry.RESET;
                        recSalesApprovalEntry.SETRANGE(recSalesApprovalEntry."Document No.", "Sales Invoice Header"."Order No.");
                        //recSalesApprovalEntry.SETRANGE(recSalesApprovalEntry.Approved,TRUE);
                        IF recSalesApprovalEntry.FINDFIRST THEN;
                        CSOAppDate := recSalesApprovalEntry."Approval Date";
                        // VK robosoft - END

                        DocDimValue := '';
                        DimDesc := '';
                        ItemDesc := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Invoice Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'MERCH');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DimDesc := recDimSet."Dimension Value Name";
                        END;

                        //>>28Jun2018
                        IF DimDesc = '' THEN
                            DimDesc := "Sales Invoice Line".Description;
                        //<<28Jun2018

                        /*
                        PostedDocDim.RESET;
                        PostedDocDim.SETRANGE(PostedDocDim."Table ID",113);
                        PostedDocDim.SETRANGE(PostedDocDim."Document No.","Sales Invoice Line"."Document No.");
                        PostedDocDim.SETRANGE(PostedDocDim."Line No.","Sales Invoice Line"."Line No.");
                        PostedDocDim.SETRANGE(PostedDocDim."Dimension Code","Sales Invoice Line"."Posting Group");
                        IF PostedDocDim.FINDFIRST THEN REPEAT
                          DocDimValue:=PostedDocDim."Dimension Value Code";
                        IF DimValue.GET('Merch',DocDimValue) THEN
                          DimDesc:=DimValue.Name;
                        UNTIL PostedDocDim.NEXT=0;
                        
                        *///Commented 27Apr2017

                        IF PrintToExcel AND SalesInvoice THEN BEGIN
                            MakeExcelDataBody;
                        END;
                        //<<1

                    end;

                    trigger OnPostDataItem()
                    begin

                        //>>1
                        //Sales Invoice Line, GroupFooter (1) - OnPreSection()
                        IF INVCOUNT <> 0 THEN //29Apr2017
                            IF PrintToExcel AND SHowInvTotal AND SalesInvoice THEN BEGIN
                                MakeExcelInvTotalGroupFooter;
                            END;
                        //<<1
                    end;

                    trigger OnPreDataItem()
                    begin

                        //>>1

                        CurrReport.CREATETOTALS(TotalFRIEGHTAmount, CessAmount, AddTaxAmt, INRTot, INRTot1, INRTot2);   //EBT Rakshitha

                        CurrReport.CREATETOTALS(TotalQty, TotalQtyBase, TotalAmount, TotalLineDisc, TotalExciseBaseAmount, TotalBEDAmount, TotaleCessAmount,
                        TotalSHECessAmount, TotalExciseAmount, TotalFOCAmount);
                        CurrReport.CREATETOTALS(TotalDISCOUNTAmount, TotalTransportSubsidy, TotalCashDiscount, TotalENTRYTax,
                        TotalCessAmount, TotalFRIEGHTAmount, TotalAddTaxAmt, TotalTaxBAse, TotalTaxAmount, TotalAmtCust);
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                        CurrReport.CREATETOTALS(TotalStateDiscount, TotalTaxAmt, "TotalAddl.Tax&Cess", TSurChargeInterest);
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
                        CurrReport.CREATETOTALS(ProdDisc, ProdDiscPerlt);
                        //<<1

                        INVCOUNT := COUNT;//27Apr2017
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    vGrpFrieghtNonTx := 0;  //RSPL008

                    SalesType := '';
                    IF Cust.GET("Sales Invoice Header"."Sell-to Customer No.") THEN;
                    IF LocationRec.GET("Sales Invoice Header"."Location Code") THEN BEGIN
                        LocationCode := LocationRec.Name;
                        LocationSate := LocationRec."State Code";
                    END;

                    StateCode := '';
                    recState.RESET;
                    recState.SETRANGE(recState.Code, LocationRec."State Code");
                    IF recState.FINDFIRST THEN BEGIN
                        //StateCode := recState."Std State Code";
                    END;

                    recCust.RESET;
                    recCust.SETRANGE(recCust."No.", "Sales Invoice Header"."Sell-to Customer No.");
                    IF recCust.FINDFIRST THEN BEGIN
                        CustResCenCode := recCust."Responsibility Center";
                        CustomerType := FORMAT(recCust.Type);

                        recResCenter.RESET;
                        recResCenter.SETRANGE(recResCenter.Code, CustResCenCode);
                        IF recResCenter.FINDFIRST THEN
                            ResCenName := recResCenter.Name;
                    END;

                    CLEAR(REGION);
                    recCusTable.RESET;
                    recCusTable.SETRANGE(recCusTable."No.", "Sales Invoice Header"."Sell-to Customer No.");
                    IF recCusTable.FINDFIRST THEN BEGIN
                        recStateTable.RESET;
                        recStateTable.SETRANGE(recStateTable.Code, recCusTable."State Code");
                        IF recStateTable.FINDFIRST THEN BEGIN
                            // REGION := recStateTable."Region Code";
                        END;
                    END;

                    IF "Sales Invoice Header"."Ship-to Code" <> '' THEN BEGIN
                        recShip2Address.RESET;
                        recShip2Address.SETRANGE(recShip2Address."Customer No.", "Sales Invoice Header"."Sell-to Customer No.");
                        recShip2Address.SETRANGE(recShip2Address.Code, "Sales Invoice Header"."Ship-to Code");
                        IF recShip2Address.FINDFIRST THEN BEGIN
                            Ship2State := recShip2Address.State;
                        END
                    END;

                    IF "Sales Invoice Header"."Gen. Bus. Posting Group" = 'DOMESTIC' THEN BEGIN
                        IF "Sales Invoice Header"."Ship-to Code" = '' THEN BEGIN
                            IF LocationSate = "Sales Invoice Header".State THEN
                                SalesType := 'Local'
                            ELSE
                                SalesType := 'Inter-State';
                        END
                        ELSE BEGIN
                            IF "Sales Invoice Header"."Ship-to Code" <> '' THEN BEGIN
                                IF Ship2State = LocationSate THEN
                                    SalesType := 'Local'
                                ELSE
                                    SalesType := 'Inter-State';
                            END
                        END
                    END
                    ELSE
                        IF "Sales Invoice Header"."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            SalesType := 'Export';

                    Saleperson.RESET;
                    Saleperson.SETRANGE(Saleperson.Code, "Sales Invoice Header"."Salesperson Code");
                    IF Saleperson.FINDFIRST THEN
                        SalesName := Saleperson.Name;

                    ShipmentAgentName := '';
                    RecShiptoAddress.RESET;
                    RecShiptoAddress.SETRANGE(RecShiptoAddress.Code, "Sales Invoice Header"."Shipping Agent Code");
                    IF RecShiptoAddress.FINDFIRST THEN BEGIN
                        ShipmentAgentName := RecShiptoAddress.Name;
                    END;

                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.", "Sales Invoice Header"."No.");
                    recSIH.SETRANGE(recSIH."Order No.", '');
                    recSIH.SETRANGE(recSIH."Supplimentary Invoice", FALSE);
                    IF recSIH.FINDFIRST THEN BEGIN
                        CurrReport.SKIP;
                    END;


                    CLEAR(freightamt);
                    //RSPL008 >>>>
                    /* StructureDetailes.RESET;
                    StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                    StructureDetailes.SETRANGE("Invoice No.", "No.");
                    StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                    StructureDetailes.SETRANGE("Tax/Charge Group", 'FREIGHT'); 
                    IF StructureDetailes.FINDSET THEN
                        REPEAT
                            // StructureDetailes.CALCFIELDS("Structure Description");
                            IF StructureDetailes."Structure Description" <> 'Freight (Non-Taxable)' THEN BEGIN
                                IF "Currency Factor" <> 0 THEN
                                    freightamt += StructureDetailes."Amount (LCY)"
                                ELSE
                                    freightamt += StructureDetailes.Amount;
                            END;
                        UNTIL StructureDetailes.NEXT = 0; */

                    //RSPL008 <<<<
                    IF freightamt <> 0 THEN BEGIN
                        recGLentry.RESET;
                        recGLentry.SETFILTER(recGLentry."G/L Account No.", '%1|%2', '75001010', '75001025');
                        recGLentry.SETRANGE(recGLentry."Document No.", "Sales Invoice Header"."No.");
                        IF recGLentry.FINDSET THEN
                            REPEAT
                                freightamt := freightamt + recGLentry.Amount
                            UNTIL recGLentry.NEXT = 0;
                    END;

                    CustPostSetup.RESET;
                    CustPostSetup.FINDFIRST;
                    RoundOffAcc := CustPostSetup."Invoice Rounding Account";
                    CLEAR(RoundOffAmt);
                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Document No.", "Sales Invoice Header"."No.");
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine.Type, SalesInvoiceLine.Type::"G/L Account");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."No.", RoundOffAcc);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        RoundOffAmt := (SalesInvoiceLine."Line Amount");
                    END;

                    //RSPL--
                    IF "Sales Invoice Header"."Ship-to Code" = '' THEN BEGIN
                        recState2.RESET;
                        recState2.SETRANGE(recState2.Code, "Sales Invoice Header".State);
                        IF recState2.FINDFIRST THEN
                            statedesc := recState2.Description;
                    END
                    ELSE BEGIN
                        //recShipaddress.RESET;
                        //recShipaddress.SETRANGE(recShipaddress.Code,"Sales Invoice Header"."Ship-to Code");
                        //IF recShipaddress.FINDFIRST THEN
                        recState2.RESET;
                        recState2.SETRANGE(recState2.Code, recShipaddress.State);
                        IF recState2.FINDFIRST THEN
                            statedesc := recState2.Description;
                    END;
                    //RSPL--



                    TotalQty := 0;
                    TotalQtyBase := 0;
                    TotalAmount := 0;
                    TotalLineDisc := 0;
                    TotalExciseBaseAmount := 0;
                    TotalBEDAmount := 0;
                    TotaleCessAmount := 0;
                    TotalSHECessAmount := 0;
                    TotalExciseAmount := 0;
                    TotalFOCAmount := 0;
                    TotalDISCOUNTAmount := 0;
                    TotalTransportSubsidy := 0;
                    TotalCashDiscount := 0;
                    TotalENTRYTax := 0;
                    TotalCessAmount := 0;
                    TotalFRIEGHTAmount := 0;
                    TotalTaxBAse := 0;
                    TotalTaxAmount := 0;
                    TotalAmtCust := 0;
                    TotalAddTaxAmt := 0;
                    TotINRval := 0;
                    TotalTaxBAse1 := 0;
                    tfreightamt := 0;
                    TotalStateDiscount := 0;
                    TSurChargeInterest := 0;
                    tFCFRIEGHTAmount := 0; // EBT MILAN
                    totalFCGROSSINV := 0; // EBT MILAN

                    tfreightamt += freightamt;
                    GfreightAmount += freightamt; //-EBT MILAN
                    //<<1
                end;

                trigger OnPostDataItem()
                begin

                    //>>1
                    //Sales Invoice Header, Footer (2) - OnPreSection()
                    //IF INVHEDCOUNT <> 0 THEN //27Apr2017
                    IF INVHEDCOUNT <> 0 THEN //09May2017
                        IF PrintToExcel AND SHowInvTotal AND SalesInvoice THEN BEGIN
                            MakeExcelInvGrandTotal;
                        END;

                    //IF PrintToExcel AND SalesInvoice THEN
                    //BEGIN
                    // MakeExcelDataHeader;
                    //END;


                    GtotalQty := 0;
                    GtotalQtyBase := 0;
                    GtotalAmount := 0;
                    GINRtot := 0;
                    GINRtot1 := 0;
                    GINRtot2 := 0;
                    GtotalLineDisc := 0;
                    GTotalExciseBaseAmount := 0;
                    GTotalBEDAmount := 0;
                    GTotaleCessAmount := 0;
                    GTotalSHECessAmount := 0;
                    GTotalExciseAmount := 0;
                    GTotalFOCAmount := 0;
                    GTotalDISCOUNTAmount := 0;
                    GTotalTransportSubsidy := 0;
                    GTotalCashDiscount := 0;
                    GTotalENTRYTax := 0;
                    GTotalCessAmount := 0;
                    GTotalFRIEGHTAmount := 0;
                    GTotalAddTaxAmt := 0;
                    GTotalTaxBAse := 0;
                    GTotalTaxAmount := 0;
                    GTotalAmtCust := 0;
                    GfreightAmount := 0;
                    GfreightNonTxAmount := 0; //RSPL008
                    GrandTotalStateDiscount := 0;
                    "GTotalAddl.Tax&Cess" := 0;
                    GTotalTaxAmt := 0;
                    //<<1
                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    "Sales Invoice Header".SETRANGE("Sales Invoice Header"."Posting Date", StartDate, EndDate);
                    StartDate := "Sales Invoice Header".GETRANGEMIN("Sales Invoice Header"."Posting Date");
                    EndDate := "Sales Invoice Header".GETRANGEMAX("Sales Invoice Header"."Posting Date");

                    // EBT MILAN 250314
                    Respocibilityfilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Shortcut Dimension 2 Code");
                    // EBT MILAN 250314

                    CurrReport.CREATETOTALS(freightamt);
                    CurrReport.CREATETOTALS(RoundOffAmt);
                    CurrReport.CREATETOTALS(ProdDisc, ProdDiscPerlt);
                    CurrReport.CREATETOTALS(FCFRIEGHTAmount, tFCFRIEGHTAmount);
                    //RSPL-230115
                    tComment := '';
                    CommentLine.RESET;
                    CommentLine.SETRANGE(CommentLine."Document Type", CommentLine."Document Type"::"Posted Invoice");
                    CommentLine.SETRANGE(CommentLine."No.", "No.");
                    IF CommentLine.FINDFIRST THEN
                        tComment := CommentLine.Comment;
                    //RSPL-230115
                    //<<1



                    INVHEDCOUNT := COUNT;//27Apr2017

                    //>>2
                    IF INVHEDCOUNT <> 0 THEN
                        IF PrintToExcel AND SalesInvoice THEN BEGIN
                            MakeExcelDataHeader;
                        END;
                    //<<2
                end;
            }
            dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
            {
                DataItemLink = "Location Code" = FIELD(Code);
                DataItemTableView = SORTING("Location Code", "Posting Date", "No.");
                dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
                {
                    DataItemLink = "Document No." = FIELD("No.");
                    DataItemTableView = SORTING("Document No.", "Line No.")
                                        WHERE(Type = FILTER(Item | 'Charge (Item)'),
                                              Quantity = FILTER(<> 0),
                                              "Item Category Code" = FILTER('CAT14'));

                    trigger OnAfterGetRecord()
                    begin

                        //>>1

                        IF RecCust2.GET("Sell-to Customer No.") THEN;
                        FOCAmount := 0;
                        StateDiscount := 0;
                        DISCOUNTAmount := 0;
                        TransportSubsidy := 0;
                        CashDiscount := 0;
                        ENTRYTax := 0;
                        CessAmount := 0;
                        FRIEGHTAmount := 0;
                        vFrieghtNonTx := 0; //RSPL008 FRIEGHT NON Taxable Amount to print
                        TotalAddTaxAmt := 0;
                        MRPPricePerLt := 0;
                        ListPricePerlt := 0;
                        SurChargeInterest := 0;
                        CLEAR(ProductGroup2);     //---004--RSPL
                        CLEAR(ProductGroupName2);  //--004--RSPL

                        //CrFCFRIEGHTAmount := 0; // EBT MILAN

                        IF "Sales Cr.Memo Line".Type = "Sales Cr.Memo Line".Type::"G/L Account" THEN BEGIN
                            CurrReport.SKIP;
                        END;

                        CrFCGROSSINV := 0;//"Sales Cr.Memo Line"."Amount To Customer";  // EBT MILAN
                                          //CrtotalFCGROSSINV += CrFCGROSSINV; // EBT MILAN
                                          //CrGtotalFCGROSSINV += CrFCGROSSINV; // EBT MILAN


                        //CrtFCFRIEGHTAmount += CrFCFRIEGHTAmount; // EBT MILAN
                        //CrgtotalFCFRIEGHTAmount += CrFCFRIEGHTAmount;  // EBT MILAN

                        //------004-------Start--------
                        recDefaultDimension2.RESET;
                        recDefaultDimension2.SETRANGE(recDefaultDimension2."Table ID", 27);
                        recDefaultDimension2.SETRANGE(recDefaultDimension2."No.", "Sales Cr.Memo Line"."No.");
                        recDefaultDimension2.SETFILTER(recDefaultDimension2."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                        IF recDefaultDimension2.FINDFIRST THEN BEGIN
                            ProductGroup2 := recDefaultDimension2."Dimension Value Code";
                            recDimensionVal2.RESET;
                            recDimensionVal2.SETFILTER(recDimensionVal2."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                            recDimensionVal2.SETRANGE(recDimensionVal2.Code, ProductGroup2);
                            IF recDimensionVal2.FINDFIRST THEN
                                ProductGroupName2 := recDimensionVal2.Name;
                        END;
                        //-----004--------END----------


                        IF "Sales Cr.Memo Header"."Last Year Sales Return" = TRUE THEN BEGIN
                            MRPPricePerLt := 0;//"Sales Cr.Memo Line"."MRP Price";
                        END ELSE BEGIN
                            //IF "Sales Cr.Memo Line"."Abatement %" <> 0 THEN
                            MRPPricePerLt := 0;//"Sales Cr.Memo Line"."MRP Price";
                        END;

                        //  IF "Sales Cr.Memo Line"."Price Inclusive of Tax" THEN
                        ListPricePerlt := "Sales Cr.Memo Line"."Unit Price Incl. of Tax" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
                        //ELSE
                        ListPricePerlt := "Sales Cr.Memo Line"."Unit Price" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";

                        /*  ExciseSetup.RESET;
                         //ExciseSetup.SETRANGE("Excise Bus. Posting Group", "Excise Bus. Posting Group");
                         //ExciseSetup.SETRANGE("Excise Prod. Posting Group", "Excise Prod. Posting Group");
                         IF ExciseSetup.FINDLAST THEN; */

                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'FOC', 'FOC DIS'); 
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 FOCAmount += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;
  */
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'STATE DISC'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                StateDiscount += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */


                        /* StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2|%3|%4|%5|%6|%7|%8|%9', 'ADD-DIS', 'AVP-SP-SAN',
                                                   'MONTH-DIS', 'NEW-PRI-SP', 'SPOT-DISC', 'VAT-DIF-DI', 'T2 DISC'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                DISCOUNTAmount += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

                        /* StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'TPTSUBSIDY'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                TransportSubsidy += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                          StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                          StructureDetailes.SETRANGE("Item No.", "No.");
                          StructureDetailes.SETRANGE("Line No.", "Line No.");
                          StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                          StructureDetailes.SETFILTER("Tax/Charge Group", '%1', 'CASHDISC');//RSPL001 
                         StructureDetailes.SETFILTER(StructureDetailes."Calculation Value", '<%1', 0);
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 CashDiscount += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0; */
                        /* 
                                                StructureDetailes.RESET;
                                                StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                                                StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                                                StructureDetailes.SETRANGE("Item No.", "No.");
                                                StructureDetailes.SETRANGE("Line No.", "Line No.");
                                                StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                                                StructureDetailes.SETFILTER("Tax/Charge Group", '%1', 'CCFC');//RSPL001 
                                                //StructureDetailes.SETFILTER(StructureDetailes."Calculation Value",'>%1',0);
                                                IF StructureDetailes.FINDSET THEN
                                                    REPEAT
                                                        SurChargeInterest += StructureDetailes.Amount;
                                                    UNTIL StructureDetailes.NEXT = 0; */


                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'ENTRYTAX'); 
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 ENTRYTax += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0; */

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'FREIGHT'); 
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                //RSPL008 >>>> FRIEGHT NON Taxable Amount
                                // StructureDetailes.CALCFIELDS("Structure Description");
                                IF StructureDetailes."Structure Description" = 'Freight (Non-Taxable)' THEN BEGIN
                                    IF "Sales Cr.Memo Header"."Currency Factor" <> 0 THEN
                                        vFrieghtNonTx += StructureDetailes."Amount (LCY)"
                                    ELSE
                                        vFrieghtNonTx += StructureDetailes.Amount;
                                END ELSE BEGIN
                                    FRIEGHTAmount += StructureDetailes.Amount;
                                END;
                                //RSPL008 >>>> FRIEGHT NON Taxable Amount
                                CrFCFRIEGHTAmount += StructureDetailes.Amount;  // EBT MILAN
                            UNTIL StructureDetailes.NEXT = 0; */


                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                        CLEAR("Addl.Tax&CessRate1");
                        /* recDetailedTaxEntry1.RESET;
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Document No.", "Sales Cr.Memo Line"."Document No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."No.", "Sales Cr.Memo Line"."No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                        recDetailedTaxEntry1.SETFILTER(recDetailedTaxEntry1."Tax Jurisdiction Code", '<>%1', "Sales Cr.Memo Line"."Tax Area Code");
                        IF recDetailedTaxEntry1.FINDFIRST THEN BEGIN
                            "Addl.Tax&CessRate1" := recDetailedTaxEntry1."Tax %"
                        END; */


                        CLEAR("Addl.Tax&Cess1");
                        /* recDetailedTaxEntry1.RESET;
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Document No.", "Sales Cr.Memo Line"."Document No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."No.", "Sales Cr.Memo Line"."No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Document Line No.", "Sales Cr.Memo Line"."Line No.");
                        recDetailedTaxEntry1.SETFILTER(recDetailedTaxEntry1."Tax Jurisdiction Code", '<>%1', "Sales Cr.Memo Line"."Tax Area Code");
                        IF recDetailedTaxEntry1.FINDFIRST THEN
                            REPEAT
                                "Addl.Tax&Cess1" := recDetailedTaxEntry1."Tax Amount";
                            UNTIL recDetailedTaxEntry.NEXT = 0;
 */

                        CLEAR(TaxRate1);
                        /* recDetailedTaxEntry1.RESET;
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Document No.", "Sales Cr.Memo Line"."Document No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."No.", "Sales Cr.Memo Line"."No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Tax Jurisdiction Code", "Sales Cr.Memo Line"."Tax Area Code");
                        IF recDetailedTaxEntry1.FINDFIRST THEN BEGIN
                            TaxRate1 := recDetailedTaxEntry1."Tax %"
                        END; */

                        CLEAR(TaxAmt1);
                        /*  recDetailedTaxEntry1.RESET;
                         recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Document No.", "Sales Cr.Memo Line"."Document No.");
                         recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."No.", "Sales Cr.Memo Line"."No.");
                         recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                         recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Document Line No.", "Sales Cr.Memo Line"."Line No.");
                         recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Tax Jurisdiction Code", "Sales Cr.Memo Line"."Tax Area Code");
                         IF recDetailedTaxEntry1.FINDFIRST THEN
                             REPEAT
                                 TaxAmt1 := recDetailedTaxEntry1."Tax Amount";
                             UNTIL recDetailedTaxEntry.NEXT = 0; */

                        CLEAR("Addl.Tax&CessRate1");
                        /* StructureDetailes.RESET;
                        /StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                        StructureDetailes.SETFILTER("Tax/Charge Code", '%1|%2|%3', 'CESS', 'C1', 'ADDL TAX'); 
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            "Addl.Tax&CessRate1" := StructureDetailes."Calculation Value";
                        END; */
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END

                        //CLEAR("Addl.Tax&Cess1");
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'CESS');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                //CessAmount += StructureDetailes.Amount;
                                "Addl.Tax&Cess1" += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

                        //CLEAR("Addl.Tax&Cess1");
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                        StructureDetailes.SETRANGE("Tax/Charge Code", 'ADDL TAX');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                //AddTaxAmt += StructureDetailes.Amount;
                                "Addl.Tax&Cess1" += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

                        CLEAR(SalesGroup1);

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Cr.Memo Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALEGROUP');
                        IF recDimSet.FINDFIRST THEN BEGIN

                            recDimSet.CALCFIELDS("Dimension Value Name");
                            SalesGroup1 := recDimSet."Dimension Value Name";

                        END;

                        /*
                        recPostedDocDim1.RESET;
                        recPostedDocDim1.SETRANGE(recPostedDocDim1."Table ID",115);
                        recPostedDocDim1.SETRANGE(recPostedDocDim1."Document No.","Sales Cr.Memo Line"."Document No.");
                        recPostedDocDim1.SETRANGE(recPostedDocDim1."Line No.","Sales Cr.Memo Line"."Line No.");
                        recPostedDocDim1.SETFILTER(recPostedDocDim1."Dimension Code",'SALEGROUP');
                        IF recPostedDocDim1.FINDFIRST THEN
                        BEGIN
                         recDimVale1.RESET;
                         recDimVale1.SETRANGE(recDimVale1."Dimension Code",'SALEGROUP');
                         recDimVale1.SETRANGE(recDimVale1.Code,recPostedDocDim1."Dimension Value Code");
                         IF recDimVale1.FINDFIRST THEN
                         BEGIN
                          SalesGroup1 := recDimVale1.Name;
                         END;
                        END;
                        *///Commented 27Apr2017


                        CLEAR(LineSalespersonCode1);
                        CLEAR(LineSalespersonName1);

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Cr.Memo Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALESPERSON');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            LineSalespersonCode1 := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            LineSalespersonName1 := recDimSet."Dimension Value Name";

                        END;
                        /*
                        recPostedDocumentDimension1.RESET;
                        recPostedDocumentDimension1.SETRANGE(recPostedDocumentDimension1."Table ID",115);
                        recPostedDocumentDimension1.SETRANGE(recPostedDocumentDimension1."Document No.","Sales Cr.Memo Line"."Document No.");
                        recPostedDocumentDimension1.SETRANGE(recPostedDocumentDimension1."Line No.","Sales Cr.Memo Line"."Line No.");
                        recPostedDocumentDimension1.SETFILTER(recPostedDocumentDimension1."Dimension Code",'SALESPERSON');
                        IF recPostedDocumentDimension1.FINDFIRST THEN
                        BEGIN
                        LineSalespersonCode1 := recPostedDocumentDimension1."Dimension Value Code";
                        
                         recDimensionValue1.RESET;
                         recDimensionValue1.SETRANGE(recDimensionValue1."Dimension Code",'SALESPERSON');
                         recDimensionValue1.SETRANGE(recDimensionValue1.Code,recPostedDocumentDimension1."Dimension Value Code");
                         IF recDimensionValue1.FINDFIRST THEN
                         BEGIN
                          LineSalespersonName1 := recDimensionValue1.Name;
                         END;
                        END;
                        *///Commented 27Apr2017
                          // EBT MILAN Product Disc........

                        CLEAR(crProdDisc);
                        CLEAR(crProdDiscPerlt);
                        MrpMaster.RESET;
                        MrpMaster.SETRANGE(MrpMaster."Item No.", "Sales Cr.Memo Line"."No.");
                        MrpMaster.SETRANGE(MrpMaster."Lot No.", "Sales Cr.Memo Line"."Lot No.");
                        IF MrpMaster.FINDFIRST THEN BEGIN
                            crProdDiscPerlt := MrpMaster."National Discount";
                            crProdDisc := ProdDiscPerlt * "Sales Cr.Memo Line"."Quantity (Base)";
                        END;



                        TotalQty += "Sales Cr.Memo Line".Quantity;
                        TotalQtyBase += "Sales Cr.Memo Line"."Quantity (Base)";
                        TotalAmount += "Sales Cr.Memo Line".Amount;
                        TotalLineDisc += "Sales Cr.Memo Line"."Line Discount Amount";
                        TotalExciseBaseAmount += 0;//"Sales Cr.Memo Line"."Excise Base Amount";
                        TotalBEDAmount += 0;//"Sales Cr.Memo Line"."BED Amount";
                        TotaleCessAmount += 0;// "Sales Cr.Memo Line"."eCess Amount";
                        TotalSHECessAmount += 0;//"Sales Cr.Memo Line"."SHE Cess Amount";
                        TotalExciseAmount += 0;// "Sales Cr.Memo Line"."Excise Amount";
                        TotalFOCAmount += FOCAmount;
                        TotalStateDiscount += StateDiscount;
                        TotalDISCOUNTAmount += DISCOUNTAmount;
                        TotalTransportSubsidy += TransportSubsidy;
                        TotalCashDiscount += CashDiscount;
                        TotalENTRYTax += ENTRYTax;
                        TotalCessAmount += CessAmount;
                        TotalFRIEGHTAmount += FRIEGHTAmount + vFrieghtNonTx; //RSPL008
                        TotalAddTaxAmt += AddTaxAmt;
                        //tfreightamt1 += freightamt1; EBT MILAN


                        // EBT MILAN 04102013
                        IF /*("Sales Cr.Memo Line"."Price Inclusive of Tax" = TRUE) AND*/ (MRPPricePerLt <> 0) AND ("Sales Cr.Memo Header"."Currency Code"
                        = '') THEN BEGIN
                            TotalAmtincludeExcise += (-"Sales Cr.Memo Line".Amount + -/*"Sales Cr.Memo Line"."Excise Amount"*/0);
                            // GTotalAmtincludeExcise += (-"Sales Cr.Memo Line".Amount+-"Sales Cr.Memo Line"."Excise Amount");
                        END
                        ELSE BEGIN
                            TotalAmtincludeExcise += (-/*"Sales Cr.Memo Line"."Amount Including Excise"*/0);
                            //  GTotalAmtincludeExcise += (-"Sales Cr.Memo Line"."Amount Including Excise");
                        END;


                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                        TotalTaxAmt1 += TaxAmt1;
                        "TotalAddl.Tax&Cess1" += "Addl.Tax&Cess1";
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END


                        IF "Sales Cr.Memo Header"."Currency Code" <> '' THEN BEGIN
                            CrtFCFRIEGHTAmount += CrFCFRIEGHTAmount;  // EBT MILAN
                            CrgtotalFCFRIEGHTAmount += CrtFCFRIEGHTAmount; //EBT MILAN
                        END;

                        IF "Sales Cr.Memo Header"."Currency Code" <> '' THEN BEGIN
                            CrtotalFCGROSSINV += CrFCGROSSINV;         // EBT MILAN
                            CrGtotalFCGROSSINV += CrtotalFCGROSSINV;   // EBT MILAN
                        END;


                        IF "Sales Cr.Memo Line"."Tax Liable" THEN
                            TotalTaxBAse += 0;// "Sales Cr.Memo Line"."Tax Base Amount";

                        TotalTaxAmount += 0;// "Sales Cr.Memo Line"."Tax Amount";
                        TotalAmtCust += 0;//"Sales Cr.Memo Line"."Amount To Customer";

                        TSurChargeInterest += SurChargeInterest;


                        IF PrintToExcel AND SalesCreditMemo THEN BEGIN
                            MakeCrMemoBody;
                        END;

                        GtotalQty1 += TotalQty;
                        GtotalQtyBase1 += TotalQtyBase;
                        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
                            GtotalAmount1 += TotalAmount;
                            GtotalLineDisc1 += TotalLineDisc;
                            GTotalExciseBaseAmount1 += TotalExciseBaseAmount;
                            GTotalBEDAmount1 += TotalBEDAmount;
                            GTotaleCessAmount1 += TotaleCessAmount;
                            GTotalSHECessAmount1 += TotalSHECessAmount;
                            GTotalExciseAmount1 += TotalExciseAmount;
                            GTotalFOCAmount1 += TotalFOCAmount;
                            GrandTotalStateDiscount1 += TotalStateDiscount;
                            GTotalDISCOUNTAmount1 += TotalDISCOUNTAmount;
                            GTotalTransportSubsidy1 += TotalTransportSubsidy;
                            GTotalCashDiscount1 += TotalCashDiscount;
                            GTotalENTRYTax1 += TotalENTRYTax;
                            GTotalCessAmount1 += TotalCessAmount;
                            GTotalFRIEGHTAmount1 += TotalFRIEGHTAmount;
                            GTotalAddTaxAmt1 += TotalAddTaxAmt;
                            GTotalTaxBAse1 += TotalTaxBAse;
                            GTotalTaxAmount1 += TotalTaxAmount;
                            GfreightNonTxAmount += vFrieghtNonTx;//RSPL008

                            //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                            "GTotalAddl.Tax&Cess1" += "TotalAddl.Tax&Cess1";
                            GTotalTaxAmt1 += TotalTaxAmt1;
                            //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
                            GTotalAmtCust1 += TotalAmtCust;
                            GSurChargeInterest1 += TSurChargeInterest;
                            // gfreightamt1 += freightamt1;
                            GTotalAmtincludeExcise += TotalAmtincludeExcise;
                        END
                        ELSE BEGIN
                            GtotalAmount1 += TotalAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GtotalLineDisc1 += TotalLineDisc / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalExciseBaseAmount1 += TotalExciseBaseAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalBEDAmount1 += TotalBEDAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotaleCessAmount1 += TotaleCessAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalSHECessAmount1 += TotalSHECessAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalExciseAmount1 += TotalExciseAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalFOCAmount1 += TotalFOCAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GrandTotalStateDiscount1 += TotalStateDiscount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalDISCOUNTAmount1 += TotalDISCOUNTAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalTransportSubsidy1 += TotalTransportSubsidy / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalCashDiscount1 += TotalCashDiscount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalENTRYTax1 += TotalENTRYTax / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalCessAmount1 += TotalCessAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalFRIEGHTAmount1 += TotalFRIEGHTAmount / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalAddTaxAmt1 += TotalAddTaxAmt / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalTaxBAse1 += TotalTaxBAse / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalTaxAmount1 += TotalTaxAmount / "Sales Cr.Memo Header"."Currency Factor";
                            //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                            "GTotalAddl.Tax&Cess1" += "TotalAddl.Tax&Cess1" / "Sales Cr.Memo Header"."Currency Factor";
                            GTotalTaxAmt1 += TotalTaxAmt1 / "Sales Cr.Memo Header"."Currency Factor";
                            //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
                            GTotalAmtCust1 += TotalAmtCust / "Sales Cr.Memo Header"."Currency Factor";
                            GSurChargeInterest1 += TSurChargeInterest / "Sales Cr.Memo Header"."Currency Factor";
                            //gfreightamt1 += freightamt1 / "Sales Cr.Memo Header"."Currency Factor";
                            //gfreightamt1 += freightamt1;
                            GTotalAmtincludeExcise += TotalAmtincludeExcise / "Sales Cr.Memo Header"."Currency Factor";
                        END;

                        DocDimValue := '';
                        DimDesc := '';
                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Cr.Memo Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'MERCH');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DimDesc := recDimSet."Dimension Value Name";
                        END;

                        //>>28Jun2018
                        IF DimDesc = '' THEN
                            DimDesc := "Sales Cr.Memo Line".Description;
                        //<<28Jun2018

                        /*
                        PostedDocDim.RESET;
                        PostedDocDim.SETRANGE(PostedDocDim."Table ID",115);
                        PostedDocDim.SETRANGE(PostedDocDim."Document No.","Sales Cr.Memo Line"."Document No.");
                        PostedDocDim.SETRANGE(PostedDocDim."Line No.","Sales Cr.Memo Line"."Line No.");
                        PostedDocDim.SETRANGE(PostedDocDim."Dimension Code","Sales Cr.Memo Line"."Posting Group");
                        IF PostedDocDim.FINDFIRST THEN REPEAT
                          DocDimValue:=PostedDocDim."Dimension Value Code";
                        
                        IF DimValue.GET('Merch',DocDimValue) THEN
                          DimDesc:=DimValue.Name;
                        
                        UNTIL PostedDocDim.NEXT=0;
                        *///Commented 27Apr2017
                          //<<1

                    end;

                    trigger OnPostDataItem()
                    begin

                        //>>1
                        //Sales Cr.Memo Line, GroupFooter (1) - OnPreSection()
                        IF CRCOUNT <> 0 THEN //27Apr2017
                            IF PrintToExcel AND SHowInvTotal AND SalesCreditMemo THEN BEGIN
                                MakeExcelCrMemoGrouping
                            END;
                        //<<1
                    end;

                    trigger OnPreDataItem()
                    begin

                        //>>1

                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                        CurrReport.CREATETOTALS(CessAmount, AddTaxAmt, "TotalAddl.Tax&Cess1", TotalTaxAmt1, TSurChargeInterest);
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END

                        CurrReport.CREATETOTALS(TotalQty, TotalQtyBase, TotalAmount, TotalLineDisc, TotalExciseBaseAmount, TotalBEDAmount, TotaleCessAmount,
                        TotalSHECessAmount, TotalExciseAmount, TotalFOCAmount);
                        CurrReport.CREATETOTALS(TotalDISCOUNTAmount, TotalTransportSubsidy, TotalCashDiscount, TotalENTRYTax,
                        TotalCessAmount, TotalFRIEGHTAmount, TotalAddTaxAmt, TotalTaxBAse, TotalAmtCust, TotalTaxAmount);

                        IF ("Sales Invoice Line".GETFILTER("Sales Invoice Line"."No.")) <> '' THEN BEGIN
                            "Sales Cr.Memo Line".SETRANGE("Sales Cr.Memo Line"."No.", "Sales Invoice Line".GETFILTER("Sales Invoice Line"."No."));
                        END;
                        CurrReport.CREATETOTALS(crProdDisc, crProdDiscPerlt);
                        //<<1


                        CRCOUNT := COUNT;//27Apr2017
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    vGrpFrieghtNonTx := 0;  //RSPL008

                    SalesType := '';
                    IF Cust.GET("Sales Cr.Memo Header"."Sell-to Customer No.") THEN;
                    IF LocationRec.GET("Sales Cr.Memo Header"."Location Code") THEN
                        LocationCode := LocationRec.Name;


                    StateCode := '';
                    recState.RESET;
                    recState.SETRANGE(recState.Code, LocationRec."State Code");
                    IF recState.FINDFIRST THEN BEGIN
                        // StateCode := recState."Std State Code";
                    END;

                    recCust.RESET;
                    recCust.SETRANGE(recCust."No.", "Sales Cr.Memo Header"."Sell-to Customer No.");
                    IF recCust.FINDFIRST THEN BEGIN
                        CustResCenCode := recCust."Responsibility Center";
                        CustomerType := FORMAT(recCust.Type);

                        recResCenter.RESET;
                        recResCenter.SETRANGE(recResCenter.Code, CustResCenCode);
                        IF recResCenter.FINDFIRST THEN
                            ResCenName := recResCenter.Name;
                    END;

                    CLEAR(REGION);
                    recCusTable.RESET;
                    recCusTable.SETRANGE(recCusTable."No.", "Sales Cr.Memo Header"."Sell-to Customer No.");
                    IF recCusTable.FINDFIRST THEN BEGIN
                        recStateTable.RESET;
                        recStateTable.SETRANGE(recStateTable.Code, recCusTable."State Code");
                        IF recStateTable.FINDFIRST THEN BEGIN
                            // REGION := recStateTable."Region Code";
                        END;
                    END;


                    IF "Sales Cr.Memo Header"."Gen. Bus. Posting Group" = 'DOMESTIC' THEN BEGIN
                        IF LocationSate = "Sales Cr.Memo Header".State THEN
                            SalesType := 'Local'
                        ELSE
                            SalesType := 'Inter-State';
                    END
                    ELSE
                        IF "Sales Cr.Memo Header"."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            SalesType := 'Export';

                    //EBT Rakshitha
                    Saleperson.RESET;
                    Saleperson.SETRANGE(Saleperson.Code, "Sales Cr.Memo Header"."Salesperson Code");
                    IF Saleperson.FINDFIRST THEN
                        SalesName := Saleperson.Name;
                    //EBT Rakshitha


                    CLEAR(freightamt1);
                    CLEAR(CrFCFRIEGHTAmount);
                    //RSPL008 >>>>
                    /* StructureDetailes.RESET;
                    StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                    StructureDetailes.SETRANGE("Invoice No.", "No.");
                    StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                    StructureDetailes.SETRANGE("Tax/Charge Group", 'FREIGHT');
                    IF StructureDetailes.FINDSET THEN
                        REPEAT
                            StructureDetailes.CALCFIELDS("Structure Description");
                            IF StructureDetailes."Structure Description" <> 'Freight (Non-Taxable)' THEN BEGIN
                                IF "Currency Factor" <> 0 THEN
                                    freightamt1 += StructureDetailes."Amount (LCY)"
                                ELSE
                                    freightamt1 += StructureDetailes.Amount;
                            END;
                        UNTIL StructureDetailes.NEXT = 0; */
                    //RSPL008 <<<<
                    IF freightamt1 <> 0 THEN BEGIN
                        recGLentry.RESET;
                        recGLentry.SETFILTER(recGLentry."G/L Account No.", '%1|%2', '75001010', '75001025');
                        recGLentry.SETRANGE(recGLentry."Document No.", "Sales Cr.Memo Header"."No.");
                        IF recGLentry.FINDFIRST THEN
                            REPEAT
                                freightamt1 := freightamt1 + recGLentry.Amount;
                            // CrFCFRIEGHTAmount += recGLentry.Amount; // EBT MILAN
                            UNTIL recGLentry.NEXT = 0;
                    END;

                    CustPostSetup1.RESET;
                    CustPostSetup1.FINDFIRST;
                    RoundOffAcc1 := CustPostSetup1."Invoice Rounding Account";
                    CLEAR(RoundOffAmt1);
                    SalesCrMemoLine.RESET;
                    SalesCrMemoLine.SETRANGE(SalesCrMemoLine."Document No.", "Sales Cr.Memo Header"."No.");
                    SalesCrMemoLine.SETRANGE(SalesCrMemoLine.Type, SalesCrMemoLine.Type::"G/L Account");
                    SalesCrMemoLine.SETFILTER(SalesCrMemoLine."No.", RoundOffAcc1);
                    IF SalesCrMemoLine.FINDFIRST THEN
                        REPEAT
                            RoundOffAmt1 := RoundOffAmt1 + (SalesCrMemoLine."Line Amount");
                        UNTIL SalesCrMemoLine.NEXT = 0;


                    InvoiceNo := '';
                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.", "Sales Cr.Memo Header"."Applies-to Doc. No.");
                    IF recSIH.FINDFIRST THEN BEGIN
                        InvoiceDate := recSIH."Posting Date";
                        InvoiceNo := recSIH."No.";
                    END;

                    //RSPL--
                    IF "Sales Cr.Memo Header"."Ship-to Code" = '' THEN BEGIN
                        recState2.RESET;
                        recState2.SETRANGE(recState2.Code, "Sales Cr.Memo Header".State);
                        IF recState2.FINDFIRST THEN
                            statedescCr := recState2.Description;
                    END
                    ELSE BEGIN
                        //recShipaddress.RESET;
                        //recShipaddress.SETRANGE(recShipaddress.Code,"Sales Invoice Header"."Ship-to Code");
                        //IF recShipaddress.FINDFIRST THEN
                        recState2.RESET;
                        recState2.SETRANGE(recState2.Code, recShipaddress.State);
                        IF recState2.FINDFIRST THEN
                            statedescCr := recState2.Description;
                    END;
                    //RSPL--



                    TotalQty := 0;
                    TotalQtyBase := 0;
                    TotalAmount := 0;
                    TotalLineDisc := 0;
                    TotalExciseBaseAmount := 0;
                    TotalBEDAmount := 0;
                    TotaleCessAmount := 0;
                    TotalSHECessAmount := 0;
                    TotalExciseAmount := 0;
                    TotalFOCAmount := 0;
                    TotalDISCOUNTAmount := 0;
                    TotalTransportSubsidy := 0;
                    TotalCashDiscount := 0;
                    TotalENTRYTax := 0;
                    TotalCessAmount := 0;
                    TotalFRIEGHTAmount := 0;
                    TotalTaxBAse := 0;
                    TotalTaxAmount := 0;
                    TotalAmtCust := 0;
                    TotalStateDiscount := 0;
                    TSurChargeInterest := 0;
                    tfreightamt1 := 0;

                    CrtFCFRIEGHTAmount := 0;     // EBT MILAN
                    CrtotalFCGROSSINV := 0; // EBT MILAN

                    tfreightamt1 += freightamt1;
                    gfreightamt1 += freightamt1; //-EBT MILAN

                    TotalAmtincludeExcise := 0;
                    //<<1
                end;

                trigger OnPostDataItem()
                begin

                    //>>1

                    //Sales Cr.Memo Header, Footer (2) - OnPreSection()
                    //IF CRHEDCOUNT <> 0 THEN
                    IF CRHEDCOUNT <> 0 THEN //09May2017
                        IF PrintToExcel AND SHowInvTotal AND SalesCreditMemo THEN BEGIN
                            MakeExcelCrmemoGrandTotal;
                        END;

                    GtotalQty1 := 0;
                    GtotalQtyBase1 := 0;
                    GtotalAmount1 := 0;
                    GtotalLineDisc1 := 0;
                    GTotalExciseBaseAmount1 := 0;
                    GTotalBEDAmount1 := 0;
                    GTotaleCessAmount1 := 0;
                    GTotalSHECessAmount1 := 0;
                    GTotalExciseAmount1 := 0;
                    GTotalFOCAmount1 := 0;
                    GrandTotalStateDiscount1 := 0;
                    GTotalDISCOUNTAmount1 := 0;
                    GTotalTransportSubsidy1 := 0;
                    GTotalCashDiscount1 := 0;
                    GTotalENTRYTax1 := 0;
                    GTotalCessAmount1 := 0;
                    GTotalFRIEGHTAmount1 := 0;
                    GTotalAddTaxAmt1 := 0;
                    GTotalTaxBAse1 := 0;
                    GTotalTaxAmount1 := 0;
                    GTotalAmtCust1 := 0;
                    "GTotalAddl.Tax&Cess1" := 0;
                    GTotalTaxAmt1 := 0;
                    gfreightamt1 := 0;
                    GTotalAmtincludeExcise := 0;
                    //<<1
                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Posting Date", StartDate, EndDate);
                    StartDate := "Sales Cr.Memo Header".GETRANGEMIN("Sales Cr.Memo Header"."Posting Date");
                    EndDate := "Sales Cr.Memo Header".GETRANGEMAX("Sales Cr.Memo Header"."Posting Date");

                    IF SalespersonFilter <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Salesperson Code", "Sales Invoice Header"."Salesperson Code");
                    END;

                    IF "Resp.CentreFilter" <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Shortcut Dimension 2 Code",
                                                        "Sales Invoice Header"."Shortcut Dimension 2 Code");
                    END;

                    IF Sell2CustomerFilter <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Sell-to Customer No.", "Sales Invoice Header"."Sell-to Customer No.");
                    END;

                    IF DivisionFilter <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Shortcut Dimension 1 Code", DivisionFilter);
                    END;

                    IF "Gen.Bus.PostingGroupFilter" <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Gen. Bus. Posting Group", "Sales Invoice Header"."Gen. Bus. Posting Group");
                    END;

                    IF CustomerPostingGroupFilter <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Customer Posting Group", "Sales Invoice Header"."Customer Posting Group");
                    END;

                    IF "ExciseBus.PostingGroupFilter" <> '' THEN BEGIN
                        // "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Excise Bus. Posting Group",
                        // "Sales Invoice Header"."Excise Bus. Posting Group");
                    END;

                    IF "Sales Invoice Header".GETFILTER("Sales Invoice Header"."No.") <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."No.", "Sales Invoice Header".GETFILTER("Sales Invoice Header"."No."));
                    END;

                    // EBT MILAN 250314
                    IF Respocibilityfilter <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Shortcut Dimension 2 Code", Respocibilityfilter);
                    END;
                    // EBT MILAN 250314


                    IF "Sales Cr.Memo Header"."Shortcut Dimension 2 Code" <> "Resp.CentreFilter" THEN
                        CurrReport.SKIP;

                    CurrReport.CREATETOTALS(RoundOffAmt1);
                    //CurrReport.CREATETOTALS(freightamt1); EBT MILAN
                    CurrReport.CREATETOTALS(crProdDisc, crProdDiscPerlt);

                    //RSPL-230115
                    tComment := '';
                    CommentLine.RESET;
                    CommentLine.SETRANGE(CommentLine."Document Type", CommentLine."Document Type"::"Posted Invoice");
                    CommentLine.SETRANGE(CommentLine."No.", "No.");
                    IF CommentLine.FINDFIRST THEN
                        tComment := CommentLine.Comment;
                    //RSPL-230115
                    //<<1

                    //>>2
                    CRHEDCOUNT := COUNT;//27Apr2017
                    //Sales Cr.Memo Header, Header (1) - OnPreSection()
                    IF CRHEDCOUNT <> 0 THEN //28Apr2017
                        IF PrintToExcel AND SalesCreditMemo THEN BEGIN
                            MakecrmemoHeader;
                        END;
                    //<<2
                end;
            }
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; StartDate)
                {
                    ApplicationArea = all;
                }
                field("End Date"; EndDate)
                {
                    ApplicationArea = all;
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = all;
                }
                field("Show Invoice Total"; SHowInvTotal)
                {
                    ApplicationArea = all;
                }
                field("Sales Invoice"; SalesInvoice)
                {
                    ApplicationArea = all;
                }
                field("Sales Return"; SalesCreditMemo)
                {
                    ApplicationArea = all;
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

        //>>1
        IF PrintToExcel THEN
            CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1

        IF (SalesInvoice = FALSE) AND (SalesCreditMemo = FALSE) THEN
            ERROR('You have not selected any option');

        LocationFilter := Location.GETFILTER(Location.Code);
        SalespersonFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Salesperson Code");
        Sell2CustomerFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Sell-to Customer No.");
        "Resp.CentreFilter" := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Responsibility Center");
        DivisionFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Shortcut Dimension 1 Code");
        "Gen.Bus.PostingGroupFilter" := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Gen. Bus. Posting Group");
        CustomerPostingGroupFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Customer Posting Group");
        //"ExciseBus.PostingGroupFilter" := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Excise Bus. Posting Group");

        /*
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",USERID);
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
        
         CLEAR(ResCenterofLocation);
         recLocationFilter.RESET;
         recLocationFilter.SETRANGE(recLocationFilter.Code,Location.GETFILTER(Location.Code));
         IF recLocationFilter.FINDFIRST THEN
         BEGIN
          ResCenterofLocation := recLocationFilter."Global Dimension 2 Code";
         END;
        
         CLEAR(ResCenterCode);
         ResCenterCode := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Shortcut Dimension 2 Code");
        
         CLEAR(SalesPersonCodeFilter);
         SalesPersonCodeFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Salesperson Code");
         IF SalesPersonCodeFilter <> '' THEN
         BEGIN
          CLEAR(ResCenterCodeofSalesPerson);
          recSalesPerson.RESET;
          recSalesPerson.SETFILTER(recSalesPerson.Code,SalesPersonCodeFilter);
          IF recSalesPerson.FINDFIRST THEN
          BEGIN
           ResCenterCodeofSalesPerson := recSalesPerson."Global Dimension 2 Code";
          END;
        
          IF ResCenterCodeofSalesPerson = '' THEN
          BEGIN
          ERROR('Resp. Center is Not specified for SalePerson %1',
                SalesPersonCodeFilter);
          END;
         END;
        
         IF ResCenterofLocation <> '' THEN
         BEGIN
          CSOmapping.RESET;
          CSOmapping.SETRANGE(CSOmapping."User Id",UPPERCASE(USERID));
          CSOmapping.SETRANGE(Type,CSOmapping.Type::Location);
          CSOmapping.SETRANGE(CSOmapping.Value,Location.GETFILTER(Location.Code));
          IF NOT(CSOmapping.FINDFIRST) THEN
          BEGIN
           CSOmapping.RESET;
           CSOmapping.SETRANGE(CSOmapping."User Id",UPPERCASE(USERID));
           CSOmapping.SETRANGE(CSOmapping.Type,CSOmapping.Type::"Responsibility Center");
           CSOmapping.SETRANGE(CSOmapping.Value,ResCenterofLocation);
           IF NOT(CSOmapping.FINDFIRST) THEN
            ERROR ('You are not allowed to run this report other than your location');
          END;
         END;
        
         IF ResCenterCode <> '' THEN
         BEGIN
          CSOmapping1.RESET;
          CSOmapping1.SETRANGE(CSOmapping1."User Id",UPPERCASE(USERID));
          CSOmapping1.SETRANGE(Type,CSOmapping1.Type::Location);
          CSOmapping1.SETRANGE(CSOmapping1.Value,Location.GETFILTER(Location.Code));
          IF NOT(CSOmapping1.FINDFIRST) THEN
          BEGIN
           CSOmapping1.RESET;
           CSOmapping1.SETRANGE(CSOmapping1."User Id",UPPERCASE(USERID));
           CSOmapping1.SETRANGE(CSOmapping1.Type,CSOmapping1.Type::"Responsibility Center");
           CSOmapping1.SETRANGE(CSOmapping1.Value,ResCenterCode);
           IF NOT(CSOmapping1.FINDFIRST) THEN
            ERROR ('You are not allowed to run this report other than your location');
          END;
         END;
        
         IF ResCenterCodeofSalesPerson <> '' THEN
         BEGIN
          CSOmapping2.RESET;
          CSOmapping2.SETRANGE(CSOmapping2."User Id",UPPERCASE(USERID));
          CSOmapping2.SETRANGE(Type,CSOmapping2.Type::Location);
          CSOmapping2.SETRANGE(CSOmapping2.Value,Location.GETFILTER(Location.Code));
          IF NOT(CSOmapping2.FINDFIRST) THEN
          BEGIN
           CSOmapping2.RESET;
           CSOmapping2.SETRANGE(CSOmapping2."User Id",UPPERCASE(USERID));
           CSOmapping2.SETRANGE(CSOmapping2.Type,CSOmapping2.Type::"Responsibility Center");
           CSOmapping2.SETRANGE(CSOmapping2.Value,ResCenterCodeofSalesPerson);
           IF NOT(CSOmapping2.FINDFIRST) THEN
            ERROR ('You are not allowed to run this report other than your location');
          END;
         END;
        
        END;
        *///Commented 27Apr2017


        IF StartDate = 0D THEN
            ERROR('Please enter the From Date and To Date');
        IF EndDate = 0D THEN
            ERROR('Please enter the From Date and To Date');

        "Company info".GET;

        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<1

        //>>ReportHeader 27Apr2017
        IF PrintToExcel THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
            //ExcelBuf.AddColumn("Company info".Name,FALSE,'',TRUE,FALSE,FALSE,'@');
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
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//39
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//40
            ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//41


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
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//39
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//40
            ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//41

        END;
        //<<ReportHeader 27Apr2017

    end;

    var
        StartDate: Date;
        EndDate: Date;
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        "Company info": Record 79;
        BEDPercent: Decimal;
        //ExciseSetup: Record 13711;
        //StructureDetailes: Record 13798;
        FOCAmount: Decimal;
        DISCOUNTAmount: Decimal;
        TransportSubsidy: Decimal;
        CashDiscount: Decimal;
        ENTRYTax: Decimal;
        CessAmount: Decimal;
        FRIEGHTAmount: Decimal;
        Cust: Record 18;
        SHowInvTotal: Boolean;
        TotalQty: Decimal;
        TotalQtyBase: Decimal;
        TotalAmount: Decimal;
        TotalLineDisc: Decimal;
        TotalExciseBaseAmount: Decimal;
        TotalBEDAmount: Decimal;
        TotaleCessAmount: Decimal;
        TotalSHECessAmount: Decimal;
        TotalExciseAmount: Decimal;
        TotalFOCAmount: Decimal;
        TotalDISCOUNTAmount: Decimal;
        TotalTransportSubsidy: Decimal;
        TotalCashDiscount: Decimal;
        TotalENTRYTax: Decimal;
        TotalCessAmount: Decimal;
        TotalFRIEGHTAmount: Decimal;
        TotalTaxBAse: Decimal;
        TotalTaxAmount: Decimal;
        TotalAmtCust: Decimal;
        LocationRec: Record 14;
        LocationCode: Text[50];
        TotalAddTaxAmt: Decimal;
        AddTaxAmt: Decimal;
        BRTNo: Code[20];
        //RG23D: Record 16537;
        "EBT Rakshitha---------": Integer;
        INRTot: Decimal;
        INRTot1: Decimal;
        INRTot2: Decimal;
        Saleperson: Record 13;
        SalesName: Text[50];
        UPricePerLT: Decimal;
        UPrice: Decimal;
        SalesPrice: Record 7002;
        TotINRval: Decimal;
        TotalTaxBAse1: Decimal;
        GtotalQty: Decimal;
        GtotalQtyBase: Decimal;
        GtotalAmount: Decimal;
        GINRtot: Decimal;
        GINRtot1: Decimal;
        GINRtot2: Decimal;
        GtotalLineDisc: Decimal;
        GTotalExciseBaseAmount: Decimal;
        GTotalBEDAmount: Decimal;
        GTotaleCessAmount: Decimal;
        GTotalSHECessAmount: Decimal;
        GTotalExciseAmount: Decimal;
        GTotalFOCAmount: Decimal;
        GTotalDISCOUNTAmount: Decimal;
        GTotalTransportSubsidy: Decimal;
        GTotalCashDiscount: Decimal;
        GTotalENTRYTax: Decimal;
        GTotalCessAmount: Decimal;
        GTotalFRIEGHTAmount: Decimal;
        GTotalAddTaxAmt: Decimal;
        GTotalTaxBAse: Decimal;
        GTotalTaxAmount: Decimal;
        GTotalAmtCust: Decimal;
        TotalINR: Decimal;
        TotalINR1: Decimal;
        TotalINR2: Decimal;
        TransferReceipt: Record 5746;
        TransferShipment: Record 5744;
        TRRNo: Code[20];
        GLEntry: Record 17;
        FreightAmtforExport: Decimal;
        SIH: Record 112;
        recSalesCrMemoHeader: Record 114;
        recGLentry: Record 17;
        freightamt: Decimal;
        tfreightamt: Decimal;
        GfreightAmount: Decimal;
        recResCenter: Record 5714;
        ResCenName: Text[50];
        CustResCenCode: Code[10];
        recCust: Record 18;
        SalesType: Text[30];
        LocationSate: Code[10];
        MRPPricePerLt: Decimal;
        ListPricePerlt: Decimal;
        StateDiscount: Decimal;
        TotalStateDiscount: Decimal;
        GrandTotalStateDiscount: Decimal;
        GtotalQty1: Decimal;
        GtotalQtyBase1: Decimal;
        GtotalAmount1: Decimal;
        GtotalLineDisc1: Decimal;
        GTotalExciseBaseAmount1: Decimal;
        GTotalBEDAmount1: Decimal;
        GTotaleCessAmount1: Decimal;
        GTotalSHECessAmount1: Decimal;
        GTotalExciseAmount1: Decimal;
        GTotalFOCAmount1: Decimal;
        GrandTotalStateDiscount1: Decimal;
        GTotalDISCOUNTAmount1: Decimal;
        GTotalTransportSubsidy1: Decimal;
        GTotalCashDiscount1: Decimal;
        GTotalENTRYTax1: Decimal;
        GTotalCessAmount1: Decimal;
        GTotalFRIEGHTAmount1: Decimal;
        GTotalAddTaxAmt1: Decimal;
        GTotalTaxBAse1: Decimal;
        GTotalTaxAmount1: Decimal;
        GTotalAmtCust1: Decimal;
        ShipmentAgentName: Text[100];
        InvoiceNo: Code[20];
        InvoiceDate: Date;
        GINRtotc: Decimal;
        recSIH: Record 112;
        recLocationFilter: Record 14;
        ResCenterofLocation: Code[100];
        CSOmapping: Record 50006;
        CSOmapping1: Record 50006;
        CSOmapping2: Record 50006;
        CustomerType: Text[30];
        RecShiptoAddress: Record 291;
        ResCenterCode: Code[100];
        ResCenterCodeofSalesPerson: Code[100];
        recSalesPerson: Record 13;
        SalesPersonCodeFilter: Text[100];
        "recItem-Supp": Record 27;
        LocationFilter: Text[30];
        SalespersonFilter: Text[100];
        Sell2CustomerFilter: Text[100];
        "Resp.CentreFilter": Text[100];
        SalesInvoice: Boolean;
        SalesCreditMemo: Boolean;
        DivisionFilter: Text[100];
        ValueEntry: Record 5802;
        ILEEntryNo: Integer;
        "ItemApplEntryNo.": Integer;
        ItemApplEntry: Record 339;
        ILE: Record 32;
        recDimVale: Record 349;
        SalesGroup: Text[50];
        recDimVale1: Record 349;
        SalesGroup1: Text[50];
        //recDetailedTaxEntry: Record 16522;
        "Addl.Tax&CessRate": Decimal;
        "Addl.Tax&Cess": Decimal;
        "TotalAddl.Tax&Cess": Decimal;
        "GTotalAddl.Tax&Cess": Decimal;
        TaxRate: Decimal;
        TaxAmt: Decimal;
        TotalTaxAmt: Decimal;
        GTotalTaxAmt: Decimal;
        //recDetailedTaxEntry1: Record 16522;
        "Addl.Tax&CessRate1": Decimal;
        "Addl.Tax&Cess1": Decimal;
        "TotalAddl.Tax&Cess1": Decimal;
        "GTotalAddl.Tax&Cess1": Decimal;
        TaxRate1: Decimal;
        TaxAmt1: Decimal;
        TotalTaxAmt1: Decimal;
        GTotalTaxAmt1: Decimal;
        recShip2Address: Record 222;
        Ship2State: Code[10];
        LineSalespersonCode: Code[10];
        LineSalespersonName: Text[50];
        recDimensionValue: Record 349;
        LineSalespersonCode1: Code[10];
        LineSalespersonName1: Text[50];
        recDimensionValue1: Record 349;
        CustPostSetup: Record 92;
        RoundOffAcc: Code[10];
        SalesInvoiceLine: Record 113;
        RoundOffAmt: Decimal;
        TRoundOffAmt: Decimal;
        GRoundOffAmt: Decimal;
        CustPostSetup1: Record 92;
        RoundOffAcc1: Code[10];
        SalesCrMemoLine: Record 115;
        RoundOffAmt1: Decimal;
        TRoundOffAmt1: Decimal;
        GRoundOffAmt1: Decimal;
        REGION: Code[10];
        recCusTable: Record 18;
        recStateTable: Record State;
        SurChargeInterest: Decimal;
        TSurChargeInterest: Decimal;
        GSurChargeInterest: Decimal;
        GSurChargeInterest1: Decimal;
        recState: Record State;
        StateCode: Code[10];
        "Gen.Bus.PostingGroupFilter": Text[100];
        CustomerPostingGroupFilter: Text[100];
        "ExciseBus.PostingGroupFilter": Text[100];
        recSalesIL: Record 113;
        freightamt1: Decimal;
        tfreightamt1: Decimal;
        gfreightamt1: Decimal;
        MrpMaster: Record 50013;
        ProdDisc: Decimal;
        ProdDiscPerlt: Decimal;
        crProdDisc: Decimal;
        crProdDiscPerlt: Decimal;
        INRamt: Decimal;
        FCFRIEGHTAmount: Decimal;
        tFCFRIEGHTAmount: Decimal;
        gtotalFCFRIEGHTAmount: Decimal;
        FCGROSSINV: Decimal;
        totalFCGROSSINV: Decimal;
        GtotalFCGROSSINV: Decimal;
        CrFCFRIEGHTAmount: Decimal;
        CrtFCFRIEGHTAmount: Decimal;
        CrgtotalFCFRIEGHTAmount: Decimal;
        CrFCGROSSINV: Decimal;
        CrtotalFCGROSSINV: Decimal;
        CrGtotalFCGROSSINV: Decimal;
        INRprice: Decimal;
        AmtincludeExcise: Decimal;
        TotalAmtincludeExcise: Decimal;
        GTotalAmtincludeExcise: Decimal;
        Respocibilityfilter: Text[30];
        RecCust2: Record 18;
        recSalesApprovalEntry: Record 50009;
        CSOAppDate: Date;
        recDefDim: Record 352;
        recDefDim2: Record 352;
        recDefaultDimension: Record 352;
        recDimensionVal: Record 349;
        ProductGroup: Text[30];
        ProductGroupName: Text[50];
        recDefaultDimension2: Record 352;
        recDimensionVal2: Record 349;
        ProductGroup2: Text[30];
        ProductGroupName2: Text[50];
        CommentLine: Record 44;
        tComment: Text[250];
        recState2: Record State;
        statedesc: Text[50];
        recShipaddress: Record 222;
        statedescCr: Text[50];
        "-----RSPL----": Integer;
        vFrieghtNonTx: Decimal;
        GfreightNonTxAmount: Decimal;
        vGrpFrieghtNonTx: Decimal;
        AmtIncludingExcise: Decimal;
        ItemDesc: Text[50];
        DimValue: Record 349;
        DocDimValue: Code[20];
        DimDesc: Text[50];
        Text002: Label 'Data';
        Text003: Label 'POP Sales Register';
        Text004: Label 'Company Name';
        Text005: Label 'Report No.';
        Text006: Label 'Report Name';
        Text007: Label 'User ID';
        Text008: Label 'Date';
        Text009: Label 'Location Filter';
        "-----27Apr2017-----": Integer;
        recDimSet: Record 480;
        INVCOUNT: Integer;
        INVHEDCOUNT: Integer;
        CRCOUNT: Integer;
        CRHEDCOUNT: Integer;
        GINQty: Decimal;
        GINQtyBase: Decimal;
        GCRQty: Decimal;
        GCRQtyBase: Decimal;
        NNNNNN: Integer;
        CCCCCC: Integer;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//26Apr2017
        ExcelBuf.AddInfoColumn(FORMAT(Text004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(FORMAT(Text003), FALSE, FALSE, FALSE, FALSE, '', 1);//26Apr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"POP Sales Register", FALSE, FALSE, FALSE, FALSE, '', 0);//26Apr2017
        //ExcelBuf.AddInfoColumn(REPORT::"Detailed Sales Register Statew",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.ClearNewRow;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text002,Text003,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>26Apr2017
        /* ExcelBuf.CreateBook('', Text003);
        ExcelBuf.CreateBookAndOpenExcel('', Text003, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text003);
        ExcelBuf.WriteSheet(Text003, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text003, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<26Apr2017
    end;

    // //[Scope('Internal')]
    procedure MakeExcelInvTotalGroupFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Line"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(SalesType, FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Form Code"*/'', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(StateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Header"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(ResCenName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn("Sales Invoice Header"."Customer Posting Group", FALSE, '', TRUE, FALSE, FALSE, '', 1);//9  //RSPL001
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Invoice Line"."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        //ExcelBuf.AddColumn(Cust."Full Name",FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(Cust.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//13 // Added Cust. Name not Full name EBT MILAN 130114
        ExcelBuf.AddColumn("Sales Invoice Header"."Ship-to Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn(GINQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//19 //27Apr2017
        GINQty := 0;//27Apr2017
                    //ExcelBuf.AddColumn(TotalQty,FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//19

        ExcelBuf.AddColumn(GINQtyBase, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//20 //27Apr2017
        GINQtyBase := 0;//27Apr2017
                        //ExcelBuf.AddColumn(TotalQtyBase,FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//20

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26
        ExcelBuf.AddColumn("Sales Invoice Header"."Salesperson Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//27
        ExcelBuf.AddColumn(SalesName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//28

        ExcelBuf.AddColumn("Sales Invoice Header"."Order No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn("Sales Invoice Header"."Document Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//41

        /*
        ExcelBuf.AddColumn("Sales Invoice Header"."Currency Code",FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn(INRTot1,FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn(INRTot2,FALSE,'',TRUE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn((INRTot),FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn("Sales Invoice Line".Amount,FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn(INRTot1,FALSE,'',TRUE,FALSE,FALSE,'');
          ExcelBuf.AddColumn(INRTot2,FALSE,'',TRUE,FALSE,FALSE,'');
        END;
        
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(TotalLineDisc,FALSE,'',TRUE,FALSE,FALSE,'');
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn(TotalExciseBaseAmount,FALSE,'',TRUE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(TotalExciseBaseAmount/"Sales Invoice Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
        END;
        
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(TotalBEDAmount,FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(TotaleCessAmount,FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(TotalSHECessAmount,FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(TotalExciseAmount,FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(AmtIncludingExcise,FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TotalFOCAmount),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TotalStateDiscount),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TotalDISCOUNTAmount),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TotalTransportSubsidy),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TotalCashDiscount),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TotalFOCAmount+TotalDISCOUNTAmount+TotalTransportSubsidy+TotalCashDiscount
                                              +TotalStateDiscount),FALSE,'',TRUE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn(ProdDisc,FALSE,'',TRUE,FALSE,FALSE,''); //EBT MILAN
        
        ExcelBuf.AddColumn(TSurChargeInterest,FALSE,'',TRUE,FALSE,FALSE,'');  //EBT STIVAN
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(TotalENTRYTax,FALSE,'',TRUE,FALSE,FALSE,'');
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn(ABS(TotalTaxBAse1),FALSE,'',TRUE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(ABS(TotalTaxBAse1)/"Sales Invoice Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
        END;
        
        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn(TotalTaxAmount,FALSE,'',TRUE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn(AddTaxAmt,FALSE,'',TRUE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn(CessAmount,FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(TaxRate,FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TotalTaxAmt),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn("Addl.Tax&CessRate",FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS("TotalAddl.Tax&Cess"),FALSE,'',TRUE,FALSE,FALSE,'');
        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
        
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
          ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'') // EBT MILAN
        ELSE
         ExcelBuf.AddColumn(tFCFRIEGHTAmount,FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN
        
        
        ExcelBuf.AddColumn(ABS(freightamt),FALSE,'',TRUE,FALSE,FALSE,'');
        
        //RSPL008 >>>>
        ExcelBuf.AddColumn(ABS(vGrpFrieghtNonTx),FALSE,'',TRUE,FALSE,FALSE,'');
        //RSPL008 <<<<
        
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(totalFCGROSSINV,FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN
        END;
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn((TotalAmtCust  + RoundOffAmt),FALSE,'',TRUE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn((TotalAmtCust + RoundOffAmt) / "Sales Invoice Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
        END;
        */
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN

        /*
        
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//42
        
        IF "Sales Invoice Header"."Local Driver's Mobile No." = '' THEN
        BEGIN
         ExcelBuf.AddColumn("Sales Invoice Header"."Driver's Mobile No.",FALSE,'',FALSE,FALSE,FALSE,'',1);//42
        END
        ELSE
        BEGIN
         IF "Sales Invoice Header"."Driver's Mobile No." = '' THEN
         BEGIN
          ExcelBuf.AddColumn("Sales Invoice Header"."Local Driver's Mobile No.",FALSE,'',FALSE,FALSE,FALSE,'',1);//42
         END
         ELSE
         BEGIN
         IF ("Sales Invoice Header"."Driver's Mobile No." = '') AND
            ("Sales Invoice Header"."Local Driver's Mobile No." = '') THEN
         BEGIN
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//42
         END;
        END;
        END;
        
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//43
        
        GINRtot += INRTot;
        GINRtot1 += INRTot1;
        GINRtot2 += INRTot2;
        GRoundOffAmt +=RoundOffAmt;
        
        ExcelBuf.AddColumn(REGION,FALSE,'',TRUE,FALSE,FALSE,'',1);//44
        ExcelBuf.AddColumn(FORMAT("Sales Invoice Header"."Freight Type"),FALSE,'',TRUE,FALSE,FALSE,'',1);//45
        */ //Commented 27Apr2017

    end;

    //  //[Scope('Internal')]
    procedure MakeExcelInvGrandTotal()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30 //RSPL008
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41
                                                                    /*
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//42
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//43
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//44
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//45
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//46
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//47
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//48
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//49
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//50
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//51
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//52
                                                                    *///Commented 26Apr2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn(GtotalQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//19
        ExcelBuf.AddColumn(GtotalQtyBase, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41

        /*
        ExcelBuf.AddColumn(ABS(GTotalDISCOUNTAmount),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//42
        ExcelBuf.AddColumn(ABS(GTotalTransportSubsidy),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//43
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//44
        ExcelBuf.AddColumn(ABS(GTotalCashDiscount),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//45
        ExcelBuf.AddColumn(ABS(GTotalFOCAmount+GTotalDISCOUNTAmount+GTotalTransportSubsidy+GTotalCashDiscount
                                       +GrandTotalStateDiscount),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//46
        
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,''); EBT MILAN
        //ExcelBuf.AddColumn(ProdDisc,FALSE,'',TRUE,FALSE,TRUE,'');    EBT MILAN
        
        ExcelBuf.AddColumn(GSurChargeInterest,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//47  //EBT STIVAN
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');           //rspl
        ExcelBuf.AddColumn(GTotalENTRYTax,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//48
        ExcelBuf.AddColumn(GTotalTaxBAse,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//49
        
        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.AddColumn(GTotalTaxAmount,FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.AddColumn(GTotalAddTaxAmt,FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.AddColumn(GTotalCessAmount,FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//50
        ExcelBuf.AddColumn(ABS(GTotalTaxAmt),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//51
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//52
        ExcelBuf.AddColumn(ABS("GTotalAddl.Tax&Cess"),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//53
        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
        ExcelBuf.AddColumn(gtotalFCFRIEGHTAmount,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//54   // EBT MILAN
        
        ExcelBuf.AddColumn(ABS(GfreightAmount),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//55
        ExcelBuf.AddColumn(ABS(GfreightNonTxAmount),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//56 //RSPL008
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//57
        ExcelBuf.AddColumn(GtotalFCGROSSINV,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//58  // EBT MILAN
        ExcelBuf.AddColumn(GTotalAmtCust + GRoundOffAmt,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//59
        
        
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');// EBT MILAN
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        *///Commented //26Apr2017

    end;

    // //[Scope('Internal')]
    procedure MakeCrMemoBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(SalesType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(StateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(ResCenName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code",FALSE,'',FALSE,FALSE,FALSE,'');  //RSPL001
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//9   //RSPL001
        ExcelBuf.AddColumn(InvoiceNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        //ExcelBuf.AddColumn(Cust."Full Name",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(Cust.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//13 // ADDED Name not full EBT MILAN 130114
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Ship-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn(DimDesc, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//18
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//19
        GCRQty += "Sales Cr.Memo Line".Quantity;//27Apr2017

        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//20
        GCRQtyBase += "Sales Cr.Memo Line"."Quantity (Base)";//27Apr2017

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//22
        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/'', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/'', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Excise Bus. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//26
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Salesperson Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//27
        ExcelBuf.AddColumn(SalesName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//28
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Return Order No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//30
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//35
        ExcelBuf.AddColumn(CustomerType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//36
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//37
        ExcelBuf.AddColumn(REGION, FALSE, '', FALSE, FALSE, FALSE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//39
        ExcelBuf.AddColumn(statedescCr, FALSE, '', FALSE, FALSE, FALSE, '', 1);//40
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//41

        /*
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Currency Code",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(-MRPPricePerLt,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(-ListPricePerlt,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(-("Sales Cr.Memo Line"."Unit Price"
        /"Sales Cr.Memo Line"."Qty. per Unit of Measure"),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Unit Price",FALSE,'',FALSE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Unit Price",FALSE,'',FALSE,FALSE,FALSE,'');
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line".Amount,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn("Sales Cr.Memo Line"."Price Inclusive of Tax",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Line Discount Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Excise Base Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(ExciseSetup."BED %",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."BED Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."eCess Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."SHE Cess Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Excise Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        
         IF ("Sales Cr.Memo Line"."Price Inclusive of Tax" = TRUE) AND (MRPPricePerLt <> 0) THEN
         BEGIN
          ExcelBuf.AddColumn(-"Sales Cr.Memo Line".Amount+-"Sales Cr.Memo Line"."Excise Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         END
         ELSE
         BEGIN
          ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Amount Including Excise",FALSE,'',FALSE,FALSE,FALSE,'');
         END;
        
         ExcelBuf.AddColumn(FOCAmount,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(StateDiscount,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(DISCOUNTAmount,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(TransportSubsidy,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(CashDiscount,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(FOCAmount+DISCOUNTAmount+TransportSubsidy+CashDiscount+StateDiscount,FALSE,'',FALSE,FALSE,FALSE,'');
        // ExcelBuf.AddColumn(crProdDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,''); //EBT
        // ExcelBuf.AddColumn(crProdDisc,FALSE,'',FALSE,FALSE,FALSE,''); //EBT
        
         ExcelBuf.AddColumn(SurChargeInterest,FALSE,'',FALSE,FALSE,FALSE,''); //EBT STIVAN
         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-ENTRYTax,FALSE,'',FALSE,FALSE,FALSE,'');      //SN
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Tax Base Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
         //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Tax %",FALSE,'',FALSE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Tax Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(AddTaxAmt,FALSE,'',FALSE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(CessAmount,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(TaxRate1,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TaxAmt1,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn("Addl.Tax&CessRate1",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Addl.Tax&Cess1",FALSE,'',FALSE,FALSE,FALSE,'');
         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
         //ExcelBuf.AddColumn(CrFCFRIEGHTAmount,FALSE,'',FALSE,FALSE,FALSE,''); // EBT MILAN
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line".Amount/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn("Sales Cr.Memo Line"."Price Inclusive of Tax",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Line Discount Amount"
         /"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Excise Base Amount"
         /"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(ExciseSetup."BED %",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."BED Amount"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."eCess Amount"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."SHE Cess Amount"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Excise Amount"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
        
         IF ("Sales Cr.Memo Line"."Price Inclusive of Tax" = TRUE) AND (MRPPricePerLt <> 0) THEN
         BEGIN
          ExcelBuf.AddColumn((-"Sales Cr.Memo Line".Amount+-"Sales Cr.Memo Line"."Excise Amount")
          /"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         END
         ELSE
         BEGIN
          ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Amount Including Excise"
          /"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         END;
        
         ExcelBuf.AddColumn((FOCAmount)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((StateDiscount)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((DISCOUNTAmount)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TransportSubsidy)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((CashDiscount)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((FOCAmount+DISCOUNTAmount+TransportSubsidy+CashDiscount+StateDiscount)
         /"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(crProdDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,''); //EBT
        // ExcelBuf.AddColumn(crProdDisc,FALSE,'',FALSE,FALSE,FALSE,''); //EBT commented
        
         ExcelBuf.AddColumn((SurChargeInterest)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');//EBT STIVAN
         //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-ENTRYTax/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');   //sn
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Tax Base Amount"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
         //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Tax %",FALSE,'',FALSE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Tax Amount",FALSE,'',FALSE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(AddTaxAmt,FALSE,'',FALSE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(CessAmount,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(TaxRate1,FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TaxAmt1/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn("Addl.Tax&CessRate1",FALSE,'',FALSE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"Addl.Tax&Cess1"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
         //ExcelBuf.AddColumn(CrFCFRIEGHTAmount,FALSE,'',FALSE,FALSE,FALSE,''); // EBT MILAN
        END;
        
        //ExcelBuf.AddColumn(CrFCFRIEGHTAmount,FALSE,'',FALSE,FALSE,FALSE,''); // EBT MILAN
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,''); //-EBT MILAN
        END
        ELSE
        BEGIN
          ExcelBuf.AddColumn(-CrFCFRIEGHTAmount,FALSE,'',FALSE,FALSE,FALSE,''); //-EBT MILAN      //sn
        END;
        
        
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         IF FRIEGHTAmount <> 0 THEN
         BEGIN
          ExcelBuf.AddColumn(-("Sales Cr.Memo Line"."Charges To Customer"- ENTRYTax - CessAmount - AddTaxAmt
          -(FOCAmount+DISCOUNTAmount+TransportSubsidy+CashDiscount+StateDiscount)),FALSE,'',FALSE,FALSE,FALSE,'');
         END
        END
        ELSE
        BEGIN
          ExcelBuf.AddColumn(-(("Sales Cr.Memo Line"."Charges To Customer"- ENTRYTax - CessAmount - AddTaxAmt)
          /"Sales Cr.Memo Header"."Currency Factor"),FALSE,'',FALSE,FALSE,FALSE,'');
        END;
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         IF FRIEGHTAmount = 0 THEN
         BEGIN
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
         END;
        END;
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,''); //-EBT MILAN
        END ELSE
        BEGIN
         ExcelBuf.AddColumn("Sales Cr.Memo Line"."Amount To Customer",FALSE,'',FALSE,FALSE,FALSE,''); //-EBT MILAN
        END;
        
        
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Amount To Customer",FALSE,'',FALSE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Amount To Customer"/
         "Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
        END;
        
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,''); //EBT MILAN
        ExcelBuf.AddColumn(-crProdDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,'');
        */

        /*
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//42
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//43
        ExcelBuf.AddColumn(REGION,FALSE,'',FALSE,FALSE,FALSE,'',1);//44
        ExcelBuf.AddColumn(FORMAT("Sales Cr.Memo Header"."Freight Type"),FALSE,'',FALSE,FALSE,FALSE,'',1);//45
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//46
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//47
        ExcelBuf.AddColumn(FORMAT(ProductGroupName2),FALSE,'',FALSE,FALSE,FALSE,'',1);//48  //004
        ExcelBuf.AddColumn(FORMAT(RecCust2."LBT Reg. No."),FALSE,'',FALSE,FALSE,FALSE,'',1);//49
        //RSPL--
        ExcelBuf.AddColumn(statedescCr,FALSE,'',FALSE,FALSE,FALSE,'',1);//50
        //RSPL--
        
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//51
        //RSPl-PARAG
        *///Commented 27Apr2017

    end;

    //  //[Scope('Internal')]
    procedure MakecrmemoHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('POP Sales Register', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1  //26Apr2017
        //ExcelBuf.AddColumn('Detailed Sales Register',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('For The Period', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('FROM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn(FORMAT(StartDate), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('TO', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn(FORMAT(EndDate), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41

        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//42
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//43
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//44
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//45
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//46
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//47
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//48
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//49
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//50
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//51
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//52
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,''); //RSPL008
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');  //004
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');//RSPL-230115
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        *///Commented 26Apr2017

        //ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Resp. Centre Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Resp. Centre Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        //RSPL001
        ExcelBuf.AddColumn('Customer Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        //RSPL001
        ExcelBuf.AddColumn('Cancelled Invoice', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Customer Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Customer Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Ship to Customer Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Unit of Measure', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Packing Qty per UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('UOM Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Total Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('TRR Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('BRT Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('Excise Chapter No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('Party TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Party CST No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('Excise Business Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        ExcelBuf.AddColumn('Header Salesperson Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('Header Salesperson Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('CSO Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('CSO Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('Vehicle Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('LR Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('LR Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33
        ExcelBuf.AddColumn('PO No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34
        ExcelBuf.AddColumn('Transporter Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35
        ExcelBuf.AddColumn('Customer Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//36
        ExcelBuf.AddColumn('Payment Terms Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//37
        ExcelBuf.AddColumn('Region', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//38
        ExcelBuf.AddColumn('CSO Approved Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//39
        ExcelBuf.AddColumn('Customer State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//40
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//41
        /*
        ExcelBuf.AddColumn('Date',FALSE,'',TRUE,FALSE,TRUE,'@',1);//1
        ExcelBuf.AddColumn('Type',FALSE,'',TRUE,FALSE,TRUE,'@',1);//2
        ExcelBuf.AddColumn('Form Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//3
        ExcelBuf.AddColumn('State Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//4
        ExcelBuf.AddColumn('Location Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//5
        ExcelBuf.AddColumn('Location Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//6
        ExcelBuf.AddColumn('Resp. Centre Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//7
        ExcelBuf.AddColumn('Resp. Centre Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//8
        //RSPL001
        ExcelBuf.AddColumn('Customer Posting Group',FALSE,'',TRUE,FALSE,TRUE,'@',1);//9
        //RSPL001
        ExcelBuf.AddColumn('Cancelled Invoice',FALSE,'',TRUE,FALSE,TRUE,'@',1);//10
        ExcelBuf.AddColumn('Document No.',FALSE,'',TRUE,FALSE,TRUE,'@',1);//11
        ExcelBuf.AddColumn('Customer Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//12
        ExcelBuf.AddColumn('Customer Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//13
        ExcelBuf.AddColumn('Ship to Customer Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//14
        ExcelBuf.AddColumn('Item No.',FALSE,'',TRUE,FALSE,TRUE,'@',1);//15
        ExcelBuf.AddColumn('Description',FALSE,'',TRUE,FALSE,TRUE,'@',1);//16
        ExcelBuf.AddColumn('Unit of Measure',FALSE,'',TRUE,FALSE,TRUE,'@',1);//17
        ExcelBuf.AddColumn('Packing Qty per UOM',FALSE,'',TRUE,FALSE,TRUE,'@',1);//18
        ExcelBuf.AddColumn('UOM Quantity Invoiced',FALSE,'',TRUE,FALSE,TRUE,'@',1);//19
        ExcelBuf.AddColumn('Total Quantity Invoiced',FALSE,'',TRUE,FALSE,TRUE,'@',1);//20
        ExcelBuf.AddColumn('TRR Details',FALSE,'',TRUE,FALSE,TRUE,'@',1);//21
        ExcelBuf.AddColumn('BRT Details',FALSE,'',TRUE,FALSE,TRUE,'@',1);//22
        ExcelBuf.AddColumn('Batch No',FALSE,'',TRUE,FALSE,TRUE,'@',1);//23
        //ExcelBuf.AddColumn('Cancelled Invoice No.',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Excise Chapter No',FALSE,'',TRUE,FALSE,TRUE,'@',1);//24
        ExcelBuf.AddColumn('Party TIN No.',FALSE,'',TRUE,FALSE,TRUE,'@',1);//25
        ExcelBuf.AddColumn('Party CST No',FALSE,'',TRUE,FALSE,TRUE,'@',1);//26
        ExcelBuf.AddColumn('Excise Business Posting Group',FALSE,'',TRUE,FALSE,TRUE,'@',1);//27
        ExcelBuf.AddColumn('Header Salesperson Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//28
        ExcelBuf.AddColumn('Header Salesperson Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//29
        ExcelBuf.AddColumn('Line Salesperson Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//30
        ExcelBuf.AddColumn('Line Salesperson Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//31
        ExcelBuf.AddColumn('Return Order Number',FALSE,'',TRUE,FALSE,TRUE,'@',1);//32
        ExcelBuf.AddColumn('Return Order Date',FALSE,'',TRUE,FALSE,TRUE,'@',1);//33
        ExcelBuf.AddColumn('Vehicle Number',FALSE,'',TRUE,FALSE,TRUE,'@',1);//34
        ExcelBuf.AddColumn('LR Number',FALSE,'',TRUE,FALSE,TRUE,'@',1);//35
        ExcelBuf.AddColumn('LR Date',FALSE,'',TRUE,FALSE,TRUE,'@',1);//36
        ExcelBuf.AddColumn('PO No',FALSE,'',TRUE,FALSE,TRUE,'@',1);//37
        ExcelBuf.AddColumn('Transporter Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//38
        ExcelBuf.AddColumn('Customer Type',FALSE,'',TRUE,FALSE,TRUE,'@',1);//39
        ExcelBuf.AddColumn('Payment Terms Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//40
        ExcelBuf.AddColumn('Sales Group',FALSE,'',TRUE,FALSE,TRUE,'@',1);//41
        */
        /*
        ExcelBuf.AddColumn('Drivers Mobile No.',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('Supp. Org Invoice',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Region',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Freight Type',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Product Group',FALSE,'',TRUE,FALSE,TRUE,'@'); //004
        ExcelBuf.AddColumn('Comment',FALSE,'',TRUE,FALSE,TRUE,'@');//RSPL-230115
        //RSPL--
        ExcelBuf.AddColumn('Customer state code',FALSE,'',TRUE,FALSE,TRUE,'@');
        //RSPL--
        
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn('Division',FALSE,'',TRUE,FALSE,TRUE,'@');
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        */ //Commented 26Apr2017

    end;

    // //[Scope('Internal')]
    procedure MakeExcelCrMemoGrouping()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(SalesType, FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Form Code"*/'', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(StateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(ResCenName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code",FALSE,'',TRUE,FALSE,FALSE,'');  //RSPL001
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Customer Posting Group", FALSE, '', TRUE, FALSE, FALSE, '', 1);//9     //RSPL001
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        //ExcelBuf.AddColumn(Cust."Full Name",FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(Cust.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//13 // ADDED Name not Full Name EBT MILAN 130114
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Ship-to Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn(-GCRQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//19 //27Apr2017
        GCRQty := 0; //27Apr2017
                     //ExcelBuf.AddColumn(-TotalQty,FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//19

        ExcelBuf.AddColumn(-GCRQtyBase, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//20 //27Apr2017
        GCRQtyBase := 0; //27Apr2017
                         //ExcelBuf.AddColumn(-TotalQtyBase,FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//20

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Salesperson Code", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn(SalesName, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Return Order No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Document Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//41

        /*
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Payment Terms Code",FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        {
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Currency Code",FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn(-TotalAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         GINRtotc += TotalAmount;
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalLineDisc,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalExciseBaseAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalBEDAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotaleCessAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalSHECessAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalExciseAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(TotalAmtincludeExcise,FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN041013
         ExcelBuf.AddColumn((TotalFOCAmount),FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalStateDiscount),FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalDISCOUNTAmount),FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalTransportSubsidy),FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalCashDiscount),FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalFOCAmount+TotalDISCOUNTAmount+TotalTransportSubsidy+TotalCashDiscount
                                       +TotalStateDiscount),FALSE,'',TRUE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(crProdDisc,FALSE,'',TRUE,FALSE,FALSE,'');
        
         ExcelBuf.AddColumn(TSurChargeInterest,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalENTRYTax,FALSE,'',TRUE,FALSE,FALSE,'');   //sn
         ExcelBuf.AddColumn(-TotalTaxBAse,FALSE,'',TRUE,FALSE,FALSE,'');
         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
         //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(-TotalTaxAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(AddTaxAmt,FALSE,'',TRUE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(CessAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(TaxRate1,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalTaxAmt1,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn("Addl.Tax&CessRate1",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"TotalAddl.Tax&Cess1",FALSE,'',TRUE,FALSE,FALSE,'');
         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(-TotalAmount/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         GINRtotc += TotalAmount /"Sales Cr.Memo Header"."Currency Factor";
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalLineDisc/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalExciseBaseAmount/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalBEDAmount/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotaleCessAmount/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalSHECessAmount/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalExciseAmount/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(TotalAmtincludeExcise/
         "Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN 041013
         ExcelBuf.AddColumn((TotalFOCAmount)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalStateDiscount)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalDISCOUNTAmount)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalTransportSubsidy)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
        // ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN 041013
         ExcelBuf.AddColumn((TotalCashDiscount),FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TotalFOCAmount+TotalDISCOUNTAmount+TotalTransportSubsidy+TotalCashDiscount
                                       +TotalStateDiscount)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(crProdDiscPerlt/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(crProdDisc/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn((TSurChargeInterest)/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');  //EBT STIVAN
        // ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,''); EBT MILAN 04102013
         ExcelBuf.AddColumn(-TotalENTRYTax/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');  //sn
         ExcelBuf.AddColumn(-TotalTaxBAse/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
         //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(-TotalTaxAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(AddTaxAmt,FALSE,'',TRUE,FALSE,FALSE,'');
         //ExcelBuf.AddColumn(CessAmount,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(TaxRate1,FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-TotalTaxAmt1/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn("Addl.Tax&CessRate1"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         ExcelBuf.AddColumn(-"TotalAddl.Tax&Cess1"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
        END;
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN
        END ELSE
        BEGIN
         ExcelBuf.AddColumn(-CrtFCFRIEGHTAmount,FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN     //sn
        END;
        
        
        
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn(freightamt1,FALSE,'',TRUE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(freightamt1,FALSE,'',TRUE,FALSE,FALSE,'');
        END;
        
        //RSPL008 >>>>
         ExcelBuf.AddColumn(vGrpFrieghtNonTx,FALSE,'',TRUE,FALSE,FALSE,'');
        //RSPL008 <<<<
        
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(CrtotalFCGROSSINV,FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN
        END;
        
        
        
        
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn(-(TotalAmtCust - (RoundOffAmt1)),FALSE,'',TRUE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn(-(TotalAmtCust - (RoundOffAmt1))/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'');
        END;
        }
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,''); // EBT MILAN
        
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(REGION,FALSE,'',TRUE,FALSE,FALSE,'');
        GRoundOffAmt1 += RoundOffAmt1;
        ExcelBuf.AddColumn(FORMAT("Sales Cr.Memo Header"."Freight Type"),FALSE,'',TRUE,FALSE,FALSE,'');
        *///Commented 27Apr2017

    end;

    // //[Scope('Internal')]
    procedure MakeExcelCrmemoGrandTotal()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41

        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        *///Commented 27Apr2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn(-GtotalQty1, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//19
        ExcelBuf.AddColumn(-GtotalQtyBase1, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41

        /*
        ExcelBuf.AddColumn(GTotalTransportSubsidy1,FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn(GTotalCashDiscount1,FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn(GTotalFOCAmount1+GTotalDISCOUNTAmount1+GTotalTransportSubsidy1+GTotalCashDiscount1
                                      +GrandTotalStateDiscount1,FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.AddColumn(crProdDisc,FALSE,'',TRUE,FALSE,TRUE,'');
        
        ExcelBuf.AddColumn(GSurChargeInterest1,FALSE,'',TRUE,FALSE,TRUE,'');     //EBT STIVAN
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn(-GTotalENTRYTax1,FALSE,'',TRUE,FALSE,TRUE,'');    //sn
        ExcelBuf.AddColumn(-GTotalTaxBAse1,FALSE,'',TRUE,FALSE,TRUE,'');
        
        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.AddColumn(-GTotalTaxAmount1,FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.AddColumn(GTotalAddTaxAmt1,FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.AddColumn(GTotalCessAmount1,FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn(-GTotalTaxAmt1,FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn(-"GTotalAddl.Tax&Cess1",FALSE,'',TRUE,FALSE,TRUE,'');
        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
        ExcelBuf.AddColumn(-CrgtotalFCFRIEGHTAmount,FALSE,'',TRUE,FALSE,TRUE,''); // EBT MILAN
        ExcelBuf.AddColumn(gfreightamt1,FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn(GfreightNonTxAmount,FALSE,'',TRUE,FALSE,TRUE,''); //RSPL008 FRIEGHT NON Taxable Amount to print
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn(CrGtotalFCGROSSINV,FALSE,'',TRUE,FALSE,TRUE,'');  // EBT MILAN
        ExcelBuf.AddColumn(-GTotalAmtCust1 - GRoundOffAmt1,FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        */
        //Commented 27Apr2017

    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(SalesType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(StateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Line"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(ResCenName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        //ExcelBuf.AddColumn("Sales Invoice Line"."Shortcut Dimension 1 Code",FALSE,'',FALSE,FALSE,FALSE,'');  //RSPL001
        ExcelBuf.AddColumn("Sales Invoice Header"."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//9  //RSPL001
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Invoice Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        //ExcelBuf.AddColumn(Cust."Full Name",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(Cust.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//13 // Added NAme not Full name EBT MILAN 130114
        ExcelBuf.AddColumn("Sales Invoice Header"."Ship-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14

        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
            "recItem-Supp".GET("Sales Invoice Line"."Final Item No.");
            ExcelBuf.AddColumn("Sales Invoice Line"."Final Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
            ExcelBuf.AddColumn(DimDesc, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        END ELSE BEGIN
            ExcelBuf.AddColumn("Sales Invoice Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
            ExcelBuf.AddColumn(DimDesc, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        END;

        ExcelBuf.AddColumn("Sales Invoice Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn("Sales Invoice Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//18
        ExcelBuf.AddColumn("Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//19
        GINQty += "Sales Invoice Line".Quantity;//27Apr2017

        ExcelBuf.AddColumn("Sales Invoice Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//20
        GINQtyBase += "Sales Invoice Line"."Quantity (Base)";//27Apr2017

        ExcelBuf.AddColumn(TRRNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn(BRTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//22
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/'', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/'', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Excise Bus. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//26
        ExcelBuf.AddColumn("Sales Invoice Header"."Salesperson Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//27
        ExcelBuf.AddColumn(SalesName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//28
        ExcelBuf.AddColumn("Sales Invoice Header"."Order No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn("Sales Invoice Header"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//30
        ExcelBuf.AddColumn("Sales Invoice Header"."Vehicle No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn("Sales Invoice Header"."LR/RR No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn("Sales Invoice Header"."LR/RR Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//33
        ExcelBuf.AddColumn("Sales Invoice Header"."External Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//34
        ExcelBuf.AddColumn(ShipmentAgentName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//35
        ExcelBuf.AddColumn(CustomerType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//36
        ExcelBuf.AddColumn("Sales Invoice Header"."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//37
        ExcelBuf.AddColumn(REGION, FALSE, '', FALSE, FALSE, FALSE, '', 1);//38
        ExcelBuf.AddColumn(CSOAppDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//39
        ExcelBuf.AddColumn(statedesc, FALSE, '', FALSE, FALSE, FALSE, '', 1);//40
        ExcelBuf.AddColumn("Sales Invoice Line"."Shortcut Dimension 1 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//41

        /*
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        
        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN
        BEGIN
         ExcelBuf.AddColumn("Sales Invoice Line"."Appiles to Inv.No.",FALSE,'',FALSE,FALSE,FALSE,'');
        END
        ELSE
         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(REGION,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(FORMAT("Sales Invoice Header"."Freight Type"),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(FORMAT(RecCust2."LBT Reg. No."),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(CSOAppDate,FALSE,'',FALSE,FALSE,FALSE,''); // VK Robosoft
        ExcelBuf.AddColumn(ProductGroupName,FALSE,'',FALSE,FALSE,FALSE,''); // 004
        ExcelBuf.AddColumn(tComment,FALSE,'',FALSE,FALSE,FALSE,''); // RSPL-230115
        
        //RSPL--
        ExcelBuf.AddColumn(statedesc,FALSE,'',FALSE,FALSE,FALSE,'');
        //RSPL--
        
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn("Sales Invoice Line"."Shortcut Dimension 1 Code",FALSE,'',FALSE,FALSE,FALSE,'');
        //RSPl-PARAG 02.01.2016
        //RSPl-DHANANJAY 10.08.2016
        ExcelBuf.AddColumn("Sales Invoice Header"."Customer Receipt Date",FALSE,'',FALSE,FALSE,FALSE,'');
        //RSPl-DHANANJAY 10.08.2016
        *///Commented 27Apr2017

        /*
        ExcelBuf.AddColumn("Sales Invoice Header"."Currency Code",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(MRPPricePerLt,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ListPricePerlt+ProdDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,'');
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
          ExcelBuf.AddColumn(ROUND((UPricePerLT),0.01),FALSE,'',FALSE,FALSE,FALSE,'');
          INRTot := ROUND((UPricePerLT),0.01);
          ExcelBuf.AddColumn(ROUND("Sales Invoice Line"."Unit Price"/"Sales Invoice Line"."Quantity (Base)"
          *"Sales Invoice Line".Quantity,0.01),FALSE,'',FALSE,FALSE,FALSE,'');
          ExcelBuf.AddColumn(ROUND(INRprice,0.01),FALSE,'',FALSE,FALSE,FALSE,'');
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
          ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
          ExcelBuf.AddColumn(UPrice,FALSE,'',FALSE,FALSE,FALSE,'');
          INRTot1 := UPrice;
          IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN
          BEGIN
           ExcelBuf.AddColumn("Sales Invoice Line".Amount,FALSE,'',FALSE,FALSE,FALSE,'');
           INRTot2 := "Sales Invoice Line".Amount;
          END ELSE BEGIN
           ExcelBuf.AddColumn((UPrice*"Sales Invoice Line".Quantity)+ProdDisc,FALSE,'',FALSE,FALSE,FALSE,'');
           ExcelBuf.AddColumn(INRamt,FALSE,'',FALSE,FALSE,FALSE,'');
           INRTot2 := INRamt;
          END;
        END
        ELSE
        BEGIN
          ExcelBuf.AddColumn((UPricePerLT/"Sales Invoice Header"."Currency Factor"),FALSE,'',FALSE,FALSE,FALSE,'');
          ExcelBuf.AddColumn(INRprice/"Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
          ExcelBuf.AddColumn((UPricePerLT),FALSE,'',FALSE,FALSE,FALSE,'');
          INRTot := UPricePerLT;
          ExcelBuf.AddColumn(UPrice*"Sales Invoice Line".Quantity,FALSE,'',FALSE,FALSE,FALSE,'');
          ExcelBuf.AddColumn(UPrice/"Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
          INRTot1 := UPrice/"Sales Invoice Header"."Currency Factor";
          ExcelBuf.AddColumn(((UPrice*"Sales Invoice Line".Quantity)+ProdDisc)
          "Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
          INRTot2 := ((UPrice*"Sales Invoice Line".Quantity)+ProdDisc)/"Sales Invoice Header"."Currency Factor";
          ExcelBuf.AddColumn(INRamt/"Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
           INRTot2 := INRamt/"Sales Invoice Header"."Currency Factor";
        
        END;
        
        ExcelBuf.AddColumn("Sales Invoice Line"."Price Inclusive of Tax",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn("Sales Invoice Line"."Line Discount Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn("Sales Invoice Line"."Excise Base Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn("Sales Invoice Line"."Excise Base Amount"
         "Sales Invoice Header"."Currency Factor" ,FALSE,'',FALSE,FALSE,FALSE,'');
        END;
        
        ExcelBuf.AddColumn(ExciseSetup."BED %",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn("Sales Invoice Line"."BED Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn("Sales Invoice Line"."eCess Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn("Sales Invoice Line"."SHE Cess Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn("Sales Invoice Line"."Excise Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         IF ("Sales Invoice Line"."Price Inclusive of Tax" = TRUE) AND (MRPPricePerLt <> 0)THEN
         BEGIN
          ExcelBuf.AddColumn(((UPrice*"Sales Invoice Line".Quantity)+"Sales Invoice Line"."Excise Amount")+
          ProdDisc,FALSE,'',FALSE,FALSE,FALSE,'');
         END
         ELSE
         BEGIN
          ExcelBuf.AddColumn("Sales Invoice Line"."Amount Including Excise"+ProdDisc,FALSE,'',FALSE,FALSE,FALSE,'');
         END;
        END
        ELSE
        BEGIN
         IF ("Sales Invoice Line"."Price Inclusive of Tax" = TRUE) AND (MRPPricePerLt <> 0)THEN
         BEGIN
          ExcelBuf.AddColumn((((UPrice*"Sales Invoice Line".Quantity)+"Sales Invoice Line"."Excise Amount")+ProdDisc)
          / "Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
         END
         ELSE
         BEGIN
          ExcelBuf.AddColumn(("Sales Invoice Line"."Amount Including Excise"+ProdDisc)
          / "Sales Invoice Header"."Currency Factor" ,FALSE,'',FALSE,FALSE,FALSE,'');
         END;
        END;
        
        //RSPL060814
        
        {IF "Sales Invoice Header"."Currency Code" = '' THEN
          ExcelBuf.AddColumn(((UPrice*"Sales Invoice Line".Quantity)+"Sales Invoice Line"."Excise Amount")+
          ProdDisc,FALSE,'',FALSE,FALSE,FALSE,'')
        ELSE
          ExcelBuf.AddColumn((((UPrice*"Sales Invoice Line".Quantity)+"Sales Invoice Line"."Excise Amount")+ProdDisc)
          / "Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');}
        
        //RSPL060814
        
        
        ExcelBuf.AddColumn(ABS(FOCAmount),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(StateDiscount),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(DISCOUNTAmount),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TransportSubsidy),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(CashDiscount),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(FOCAmount+DISCOUNTAmount+TransportSubsidy+CashDiscount+StateDiscount),FALSE,'',FALSE,FALSE,FALSE,'');
        
        //ExcelBuf.AddColumn(ProdDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,'');    //EBT
        //ExcelBuf.AddColumn(ProdDisc,FALSE,'',FALSE,FALSE,FALSE,'');    //EBT
        
        
        ExcelBuf.AddColumn(SurChargeInterest,FALSE,'',FALSE,FALSE,FALSE,'');    //EBT STIVAN
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ENTRYTax,FALSE,'',FALSE,FALSE,FALSE,'');
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn("Sales Invoice Line"."Tax Base Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        END ELSE BEGIN
         ExcelBuf.AddColumn("Sales Invoice Line"."Tax Base Amount"/"Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
        END;
        
        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
        //ExcelBuf.AddColumn("Sales Invoice Line"."Tax %",FALSE,'',FALSE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn("Sales Invoice Line"."Tax Amount",FALSE,'',FALSE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn(AddTaxAmt,FALSE,'',FALSE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn(CessAmount,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(TaxRate,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(TaxAmt),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn("Addl.Tax&CessRate",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS("Addl.Tax&Cess"),FALSE,'',FALSE,FALSE,FALSE,'');
        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'')  // EBT MILAN
        ELSE
        ExcelBuf.AddColumn(FCFRIEGHTAmount,FALSE,'',FALSE,FALSE,FALSE,''); // EBT MILAN
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         IF FRIEGHTAmount <> 0 THEN BEGIN
          ExcelBuf.AddColumn("Sales Invoice Line"."Charges To Customer" - ENTRYTax - CessAmount - AddTaxAmt
                             - DISCOUNTAmount-SurChargeInterest,FALSE,'',FALSE,FALSE,FALSE,'');
          TotalFRIEGHTAmount := "Sales Invoice Line"."Charges To Customer";
         END ELSE BEGIN
           ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
         END;
        END ELSE BEGIN
          ExcelBuf.AddColumn(("Sales Invoice Line"."Charges To Customer" - ENTRYTax - CessAmount - AddTaxAmt)
          /"Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'');
          TotalFRIEGHTAmount := "Sales Invoice Line"."Charges To Customer"/"Sales Invoice Header"."Currency Factor";
        END;
        
        //RSPL008 >>>>
        ExcelBuf.AddColumn(vFrieghtNonTx,FALSE,'',FALSE,FALSE,FALSE,'');
        //RSPL008 <<<<
        
        
        
        ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        
        IF "Sales Invoice Header"."Currency Code"= '' THEN
         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'') // EBT MILAN
        ELSE
          ExcelBuf.AddColumn("Sales Invoice Line"."Amount To Customer",FALSE,'',FALSE,FALSE,FALSE,''); // EBT MILAN
        
        
        
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
         ExcelBuf.AddColumn("Sales Invoice Line"."Amount To Customer",FALSE,'',FALSE,FALSE,FALSE,'');
        END
        ELSE
        BEGIN
         ExcelBuf.AddColumn("Sales Invoice Line"."Amount To Customer"/"Sales Invoice Header"."Currency Factor"
         ,FALSE,'',FALSE,FALSE,FALSE,'');
        END;
        */

    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('POP Sales Register', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1 //27Apr2017
        //ExcelBuf.AddColumn('Detailed Sales Register',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('For The Period', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('FROM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn(FORMAT(StartDate), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('TO', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn(FORMAT(EndDate), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41

        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');  //004
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');  //004
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');//RSPL-230115
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');//RSPL-DHANANJAY
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');//RSPL-DHANANJAY
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'');//RSPL-DHANANJAY
        *///Commented 27Apr2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Resp. Centre Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Resp. Centre Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        //RSPL001
        ExcelBuf.AddColumn('Customer Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        //RSPL001
        ExcelBuf.AddColumn('Cancelled Invoice', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Customer Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Customer Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Ship to Customer Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Unit of Measure', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Packing Qty per UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('UOM Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Total Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('TRR Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('BRT Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('Excise Chapter No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('Party TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Party CST No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('Excise Business Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        ExcelBuf.AddColumn('Header Salesperson Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('Header Salesperson Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('CSO Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('CSO Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('Vehicle Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('LR Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('LR Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33
        ExcelBuf.AddColumn('PO No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34
        ExcelBuf.AddColumn('Transporter Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35
        ExcelBuf.AddColumn('Customer Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//36
        ExcelBuf.AddColumn('Payment Terms Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//37
        ExcelBuf.AddColumn('Region', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//38
        ExcelBuf.AddColumn('CSO Approved Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//39
        ExcelBuf.AddColumn('Customer State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//40
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//41


        /*
        ExcelBuf.AddColumn('Drivers Mobile No.',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Supp. Org Invoice',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Region',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Freight Type',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Customer LBT No.',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('CSO Approved Date',FALSE,'',TRUE,FALSE,TRUE,'@'); // VK robosoft
        ExcelBuf.AddColumn('Product Group',FALSE,'',TRUE,FALSE,TRUE,'@'); // 004
        ExcelBuf.AddColumn('Comment',FALSE,'',TRUE,FALSE,TRUE,'@');//RSPL-230115
        //RSPL--
        ExcelBuf.AddColumn('Customer state code',FALSE,'',TRUE,FALSE,TRUE,'@');
        //RSPL--
        
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn('Division',FALSE,'',TRUE,FALSE,TRUE,'@');
        //RSPl-PARAG 02.01.2016
        //RSPL-DHANANJAY 10.08.2016
        ExcelBuf.AddColumn('Customer Receipt Date',FALSE,'',TRUE,FALSE,TRUE,'@');
        //RSPL-DHANANJAY 10.08.2016
        */ //Commented 27Apr2017

    end;
}


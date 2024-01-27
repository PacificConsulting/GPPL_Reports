report 50092 "Sales Register - GST"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 11Sep2017   RB-N         Filtering of Report on the basis of Location GSTNo.
    // 14Sep2017   RB-N         Proper Alignment of CGST & SGST Value
    // 20Sep2017   RB-N         PrintSummary Option & GST %
    // 25Oct2017   RB-N         Division Name
    // 30Dec2017   RB-N         Issue for FC Amount in Export Cancelled Invoice
    // 13Mar2018   RB-N         GST Comp Cess , TCS Amount
    // 15Jan2019   RB-N         National Discount Amount
    // 
    // 29Oct20     AR           IRN Generated, EWB No. ,and EWB Date added
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesRegisterGST.rdl';
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
                DataItemTableView = SORTING("Location Code", "Posting Date", "No.")
                                    WHERE("Debit Memo" = FILTER(false));
                RequestFilterFields = "Sell-to Customer No.", "Shortcut Dimension 1 Code", "Shortcut Dimension 2 Code", "Salesperson Code", "Gen. Bus. Posting Group", "Customer Posting Group"/*, "Excise Bus. Posting Group"*/, "No.";
                dataitem("Sales Invoice Line"; "Sales Invoice Line")
                {
                    DataItemLink = "Document No." = FIELD("No.");
                    DataItemTableView = SORTING("Document No.", "Line No.")
                                        WHERE(Quantity = FILTER(<> 0),
                                              "Item Category Code" = FILTER(<> 'CAT14'),
                                              Type = FILTER(Item | "G/L Account"),
                                              "No." = FILTER(<> '74012210'));
                    RequestFilterFields = Type, "No.";

                    trigger OnAfterGetRecord()
                    begin
                        TILCount += 1;//23Feb2017 RB-N

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


                        /*
                        IF "Sales Invoice Line".Type = "Sales Invoice Line".Type :: "G/L Account" THEN
                        BEGIN
                         CurrReport.SKIP;
                        END;
                        *///Commented 12Dec2017


                        //Sales Invoice Header, Header (1) - OnPreSection()
                        NNNNNN += 1;
                        iF NNNNNN = 1 THEN //02Aug2017
                            IF PrintToExcel AND SalesInvoice THEN BEGIN
                                MakeExcelDataHeader;
                            END;

                        CLEAR(ProdDisc);
                        CLEAR(ProdDiscPerlt);
                        /*
                        MrpMaster.RESET;
                        MrpMaster.SETRANGE(MrpMaster."Item No.","Sales Invoice Line"."No.");
                        MrpMaster.SETRANGE(MrpMaster."Lot No.","Sales Invoice Line"."Lot No.");
                        IF MrpMaster.FINDFIRST THEN
                        BEGIN
                          ProdDiscPerlt := MrpMaster."National Discount";
                          ProdDisc := ProdDiscPerlt * "Sales Invoice Line"."Quantity (Base)";
                        END;
                        */
                        //>>06Feb2019
                        //ProdDiscPerlt := "Sales Invoice Line"."National Discount Per Ltr";
                        //ProdDiscPerlt := "Sales Invoice Line"."National Discount Per Ltr";//RSPLSUM
                        //ProdDisc := "Sales Invoice Line"."National Discount Amount";//RSPLSUM
                        //<<06Feb2019

                        //RSPLSUM>>
                        ProdDiscPerlt := 0;
                        ProdDisc := 0;
                        IF "Sales Invoice Header"."Posting Date" <= 20190303D THEN BEGIN //030319D
                            ProdDisc := "Sales Invoice Line"."National Discount Amount";
                            ProdDiscPerlt := "Sales Invoice Line"."National Discount Per Ltr";
                        END ELSE BEGIN
                            ProdDisc := 0;
                            ProdDiscPerlt := 0;
                        END;
                        //RSPLSUM<<

                        //IF "Sales Invoice Line"."Abatement %" <> 0 THEN
                        MRPPricePerLt := 0;// "Sales Invoice Line"."MRP Price";

                        // IF "Sales Invoice Line"."Price Inclusive of Tax" THEN BEGIN
                        ListPricePerlt := "Sales Invoice Line"."Unit Price Incl. of Tax" / "Sales Invoice Line"."Qty. per Unit of Measure";
                        UPricePerLT := /*"Sales Invoice Line"."MRP Price"*/0 - ((/*"Sales Invoice Line"."MRP Price" * "Sales Invoice Line"."Abatement %"*/0) / 100);
                        UPrice := UPricePerLT * "Sales Invoice Line"."Qty. per Unit of Measure";
                        INRprice := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Quantity (Base)" * "Sales Invoice Line".Quantity;
                        INRamt := (ListPricePerlt + ProdDiscPerlt) * "Sales Invoice Line"."Quantity (Base)" -
                        /*"Sales Invoice Line"."Excise Amount"*/0;
                        // END ELSE BEGIN
                        ListPricePerlt := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure";
                        UPricePerLT := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure";
                        UPrice := "Sales Invoice Line"."Unit Price";
                        INRprice := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Quantity (Base)" * "Sales Invoice Line".Quantity;
                        //INRamt := UPrice*"Sales Invoice Line".Quantity;
                        INRamt := (ListPricePerlt + ProdDiscPerlt) * "Quantity (Base)";//15Jan2019
                                                                                       // END;

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

                        /*
                        ExciseSetup.RESET;
                        ExciseSetup.SETRANGE("Excise Bus. Posting Group","Sales Invoice Line"."Excise Bus. Posting Group");
                        ExciseSetup.SETRANGE("Excise Prod. Posting Group","Excise Prod. Posting Group");
                        IF ExciseSetup.FINDLAST THEN;
                        
                        StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.","Item No.","Line No.","Tax/Charge Type","Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group",'%1|%2','FOC','FOC DIS');
                        IF StructureDetailes.FINDSET THEN
                        REPEAT
                          FOCAmount += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0;
                        
                        StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.","Item No.","Line No.","Tax/Charge Type","Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group",'%1|%2|%3|%4|%5|%6|%7|%8|%9|%10','ADD-DIS','AVP-SP-SAN',
                                                   'MONTH-DIS','NEW-PRI-SP','SPOT-DISC','VAT-DIF-DI','T2 DISC','DISCOUNT');
                        IF StructureDetailes.FINDSET THEN
                        REPEAT
                          DISCOUNTAmount += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0;
                        */

                        //>>01Aug2017
                        CLEAR(AVPDiscAmt);
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        //StructureDetailes.SETFILTER("Tax/Charge Group",'AVP-SP-SAN');
                        StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'AVP-SP-SAN', 'NATSCHEME');//02Nov2018
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            AVPDiscAmt := ABS(StructureDetailes.Amount);
                        END; */

                        //IF AVPDiscAmt = 0 THEN
                        AVPDiscAmt += "National Discount Amount";//15Jan2019
                        //>>02May2019
                        IF "Sales Invoice Line"."Free of Cost" THEN
                            AVPDiscAmt += ABS("Sales Invoice Line"."Line Discount Amount");
                        //<<02May2019

                        TAVPDiscAmt += AVPDiscAmt;
                        TTAVPDiscAmt += AVPDiscAmt;

                        CLEAR(DiscAmt);
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", 'DISCOUNT');
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            DiscAmt := ABS(StructureDetailes.Amount);
                        END; */
                        TDiscAmt += DiscAmt;
                        TTDiscAmt += DiscAmt;

                        //<<01Aug2017

                        //>>13Mar2018
                        CLEAR(CessAmt);
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", 'CESS');
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            CessAmt := StructureDetailes.Amount;
                        END; */
                        TCessAmt += CessAmt;
                        TTCessAmt += CessAmt;

                        CLEAR(TCSAmt);
                        TCSAmt := 0;// "TDS/TCS Amount";
                        TTCSAmt += TCSAmt;
                        TTTCSAmt += TCSAmt;
                        //<<13Mar2018

                        //>>02Nov2018
                        CLEAR(RegDiscAmt);
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", 'REGSCHEME');
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            RegDiscAmt := ABS(StructureDetailes.Amount);
                        END; */
                        TRegDiscAmt += RegDiscAmt;
                        TTRegDiscAmt += RegDiscAmt;
                        //<<02Nov2018

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'STATE DISC');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                StateDiscount += ABS(StructureDetailes.Amount);
                            UNTIL StructureDetailes.NEXT = 0; */


                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2|%3', 'TPTSUBSIDY', 'TRANSSUBS', 'TRANSSUBS2');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT
                                TransportSubsidy += ABS(StructureDetailes.Amount);
                            UNTIL StructureDetailes.NEXT = 0; */

                        /*   */

                        //Here-
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
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

                        /* StructureDetailes.RESET;
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
                            UNTIL StructureDetailes.NEXT = 0; */

                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'FREIGHT');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 //RSPL008 >>>> FRIEGHT NON Taxable Amount
                                 StructureDetailes.CALCFIELDS("Structure Description");
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
                             UNTIL StructureDetailes.NEXT = 0;


                         //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                         CLEAR("Addl.Tax&CessRate");
                         recDetailedTaxEntry.RESET;
                         recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Invoice Line"."Document No.");
                         recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."No.", "Sales Invoice Line"."No.");
                         recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Invoice Line"."Tax Area Code");
                         recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Invoice Line"."Tax Area Code");
                         IF recDetailedTaxEntry.FINDFIRST THEN BEGIN
                             "Addl.Tax&CessRate" := recDetailedTaxEntry."Tax %"
                         END;


                         CLEAR("Addl.Tax&Cess");
                         recDetailedTaxEntry.RESET;
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
                        /*  recDetailedTaxEntry.RESET;
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

                        /*
                        //CLEAR("Addl.Tax&CessRate");
                        StructureDetailes.RESET;
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::"Other Taxes");
                        StructureDetailes.SETFILTER("Tax/Charge Code",'%1|%2|%3','CESS','C1','ADDL TAX');
                        IF StructureDetailes.FINDFIRST THEN
                        BEGIN
                         "Addl.Tax&CessRate" := StructureDetailes."Calculation Value";
                        END;
                        
                        //CLEAR("Addl.Tax&Cess");
                        StructureDetailes.RESET;
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::"Other Taxes");
                        StructureDetailes.SETFILTER("Tax/Charge Code",'%1|%2','CESS','C1');
                        IF StructureDetailes.FINDSET THEN
                        REPEAT
                          //CessAmount += StructureDetailes.Amount;
                          "Addl.Tax&Cess" += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0;
                        */

                        //CLEAR("Addl.Tax&Cess");
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
                                "Addl.Tax&Cess" += StructureDetailes.Amount;
                            UNTIL StructureDetailes.NEXT = 0; */

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
                        //>>27Feb2017
                        recDimSet.RESET;
                        //recDimSet.SETCURRENTKEY("Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Invoice Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALEGROUP');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            SalesGroup := recDimSet."Dimension Value Name";
                        END;
                        //<<
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
                        *///Commented T359..PostedDimensionValue

                        CLEAR(LineSalespersonCode);
                        CLEAR(LineSalespersonName);
                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Invoice Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALESPERSON');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            LineSalespersonCode := recDimSet."Dimension Value Code";
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
                        *///Commented T359..PostedDimensionValue

                        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
                            "Sales Invoice Line".Quantity := 0;
                            "Sales Invoice Line"."Quantity (Base)" := 0;
                            TotalQty := 0;
                            TotalQtyBase := 0;
                        END;


                        // RSPL-Ragni- Calculate total amount including excise
                        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                            IF /*("Sales Invoice Line"."Price Inclusive of Tax" = TRUE) AND*/ (MRPPricePerLt <> 0) THEN BEGIN
                                AmtIncludingExcise += ((UPrice * "Sales Invoice Line".Quantity) + 0/*"Sales Invoice Line"."Excise Amount"*/) + ProdDisc;
                            END ELSE BEGIN
                                AmtIncludingExcise += /*"Sales Invoice Line"."Amount Including Excise"*/0 + ProdDisc;
                            END;
                        END ELSE BEGIN
                            IF /*("Sales Invoice Line"."Price Inclusive of Tax"0 = TRUE) AND*/ (MRPPricePerLt <> 0) THEN BEGIN
                                AmtIncludingExcise += (((UPrice * "Sales Invoice Line".Quantity) + 0/*"Sales Invoice Line"."Excise Amount"*/) + ProdDisc)
                                / "Sales Invoice Header"."Currency Factor";
                            END ELSE BEGIN
                                AmtIncludingExcise += (/*"Sales Invoice Line"."Amount Including Excise"*/0 + ProdDisc)
                                / "Sales Invoice Header"."Currency Factor";
                            END;
                        END;
                        // RSPL- Ragni- Calculate total amount including excise

                        IF "Sales Invoice Line"."Tax Liable" = TRUE THEN
                            TotalTaxBAse1 += 0;// "Sales Invoice Line"."Tax Base Amount";

                        TotalTaxAmount += 0;//"Sales Invoice Line"."Tax Amount";
                        //TotalAmtCust += "Sales Invoice Line"."Amount To Customer";

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
                        END ELSE BEGIN
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
                        END ELSE BEGIN
                            GTotalTaxBAse += TotalTaxBAse / "Sales Invoice Header"."Currency Factor";
                        END;

                        GTotalTaxAmount += TotalTaxAmount;

                        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                            GTotalAmtCust := GTotalAmtCust + TotalAmtCust;
                        END ELSE BEGIN
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
                        IF "Sales Invoice Header"."Currency Code" <> '' THEN BEGIN
                            gtotalFCFRIEGHTAmount += tFCFRIEGHTAmount; // EBT MILAN
                        END;

                        // VK robosoft - BEGIN
                        CSOAppDate := 0D;
                        recSalesApprovalEntry.RESET;
                        recSalesApprovalEntry.SETRANGE(recSalesApprovalEntry."Document No.", "Sales Invoice Header"."Order No.");
                        //recSalesApprovalEntry.SETRANGE(recSalesApprovalEntry.Approved,TRUE);
                        IF recSalesApprovalEntry.FINDFIRST THEN;
                        CSOAppDate := recSalesApprovalEntry."Approval Date";
                        // VK robosoft - END

                        //>>31July2017 GST
                        CLEAR(vCGST);
                        CLEAR(vSGST);
                        CLEAR(vIGST);
                        CLEAR(vUTGST);
                        CLEAR(vCGSTRate);
                        CLEAR(vSGSTRate);
                        CLEAR(vIGSTRate);
                        CLEAR(GSTPer);//RB-N 20Sep2017

                        DetailGSTEntry.RESET;
                        DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                        //DetailGSTEntry.SETRANGE(Type, Type);
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
                        IF DetailGSTEntry.FINDSET THEN
                            REPEAT

                                GSTPer += DetailGSTEntry."GST %";//RB-N 20Sep2017

                                IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                    vCGST := ABS(DetailGSTEntry."GST Amount");
                                    vCGSTRate := DetailGSTEntry."GST %";
                                END;

                                IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                    vSGST := ABS(DetailGSTEntry."GST Amount");
                                    vSGSTRate := DetailGSTEntry."GST %";

                                END;

                                IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                    vUTGST := ABS(DetailGSTEntry."GST Amount");
                                    vSGSTRate := DetailGSTEntry."GST %";

                                END;


                                IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                    vIGST := ABS(DetailGSTEntry."GST Amount");
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
                        //<<31July2017 GST


                        //>>04Mar2019
                        LineDiscAmt := 0;
                        IF "Sales Invoice Header"."Posting Date" <= 20190303D THEN BEGIN //030319D
                            LineDiscAmt := "Sales Invoice Line"."Line Discount Amount";
                        END ELSE BEGIN
                            LineDiscAmt := 0;
                        END;
                        //<<04Mar2019

                        //IF PrintToExcel AND SalesInvoice THEN
                        IF PrintToExcel AND SalesInvoice AND (PrintSummary = FALSE) THEN //20Sep2017
                        BEGIN
                            MakeExcelDataBody;
                        END;
                        //<<1

                        //>>20Sep2017 GroupData
                        TQty := TQty + "Sales Invoice Line".Quantity;
                        InTQty := InTQty + "Sales Invoice Line".Quantity;//27Feb2017

                        TQtyBase := TQtyBase + "Sales Invoice Line"."Quantity (Base)";
                        InTQtyBase := InTQtyBase + "Sales Invoice Line"."Quantity (Base)";//27Feb2017

                        TTransportSubsidy += TransportSubsidy;
                        InTTransportSubsidy += TransportSubsidy;//27Feb

                        TCashDiscount += CashDiscount;
                        InTCashDiscount += CashDiscount;//27Feb

                        //TLineDisc += "Sales Invoice Line"."Line Discount Amount";
                        TLineDisc += LineDiscAmt;
                        //InLineDisc += "Sales Invoice Line"."Line Discount Amount";
                        InLineDisc += LineDiscAmt;

                        InFCFG += FCFRIEGHTAmount;
                        TInFCFG += FCFRIEGHTAmount;


                        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                            /*
                              IF "Sales Invoice Line"."GST Base Amount" <> 0 THEN
                              BEGIN
                              TTaxBAse += "Sales Invoice Line"."GST Base Amount";
                              InTTaxBAse += "Sales Invoice Line"."GST Base Amount";
                              END ELSE
                              IF "Sales Invoice Line"."Free of Cost" THEN
                              BEGIN
                                TTaxBAse += "Sales Invoice Line"."GST Base Amount";
                                InTTaxBAse += "Sales Invoice Line"."GST Base Amount";
                              END ELSE
                              BEGIN
                              TTaxBAse += "Sales Invoice Line"."Tax Base Amount";
                              InTTaxBAse += "Sales Invoice Line"."Tax Base Amount";
                              END;
                            */
                            //>>07Aug2019
                            IF "Sales Invoice Header"."Posting Date" >= 20170701D THEN BEGIN //070117D
                                TTaxBAse += 0;// "Sales Invoice Line"."GST Base Amount";
                                InTTaxBAse += 0;// "Sales Invoice Line"."GST Base Amount";
                            END ELSE BEGIN
                                TTaxBAse += 0;// "Sales Invoice Line"."Tax Base Amount";
                                InTTaxBAse += 0;// "Sales Invoice Line"."Tax Base Amount";
                            END;
                            //<<07Aug2019

                            TGrossInVAmt += 0;// "Sales Invoice Line"."Amount To Customer";//23Feb2017
                            InTGrossInVAmt += 0;//"Sales Invoice Line"."Amount To Customer";//27Feb2017

                        END;

                        IF "Sales Invoice Header"."Currency Code" <> '' THEN BEGIN
                            FCGROSSINV += 0;//"Sales Invoice Line"."Amount To Customer";
                            InTFCGROSSINV += 0;//"Sales Invoice Line"."Amount To Customer";

                            TTaxBAse += "Sales Invoice Line"."Line Amount" / "Sales Invoice Header"."Currency Factor";
                            InTTaxBAse += "Sales Invoice Line"."Line Amount" / "Sales Invoice Header"."Currency Factor";//27Feb

                            TGrossInVAmt += 0;// "Sales Invoice Line"."Amount To Customer" / "Sales Invoice Header"."Currency Factor";//23Feb2017
                            InTGrossInVAmt += 0;//"Sales Invoice Line"."Amount To Customer" / "Sales Invoice Header"."Currency Factor";//27Feb2017

                        END;
                        //<<20Sep2017

                        //>>22Feb2017 RB-N INVOICEGroupTotal

                        IF TILCount = SILCount THEN BEGIN
                            //IF PrintToExcel AND SHowInvTotal AND SalesInvoice THEN
                            IF PrintToExcel AND SalesInvoice THEN //20Sep2017
                            BEGIN
                                MakeExcelInvTotalGroupFooter;
                                TILCount := 0;
                            END;
                            //>>23Feb2017
                            TQty := 0;
                            TQtyBase := 0;
                            TLineDisc := 0;
                            TExciseBaseAmount := 0;
                            TBEDAmount := 0;
                            TeCessAmount := 0;
                            TSHECessAmount := 0;
                            TExciseAmount := 0;
                            AmtIncludingExcise := 0;//23Feb2017
                            TFOCAmount := 0;
                            TStateDiscount := 0;
                            TDISCOUNTAmount := 0;
                            TTransportSubsidy := 0;
                            TCashDiscount := 0;
                            TENTRYTax := 0;
                            TTaxBAse := 0;
                            TFRIEGHTAmount := 0;
                            TTaxAmt := 0;
                            INRTot1 := 0;
                            INRTot2 := 0;
                            INRTot := 0;
                            totalFCGROSSINV := 0; //
                            FCGROSSINV := 0;
                            TotalAmtCust := 0; //
                            TINRUnit := 0;
                            TINRAmt := 0;
                            TGrossInVAmt := 0;
                            TFCGrossInVAmt := 0;
                            InFCFG := 0;
                            InFG := 0;
                            TCGST := 0;//31July2017
                            TSGST := 0;//31July2017
                            TIGST := 0;//31July2017
                            TUTGST := 0;//31July2017

                            //<<
                        END;
                        //<<

                    end;

                    trigger OnPreDataItem()
                    begin
                        //>>1

                        CurrReport.CREATETOTALS(TotalFRIEGHTAmount, CessAmount, AddTaxAmt, INRTot, INRTot1, INRTot2);   //EBT Rakshitha

                        CurrReport.CREATETOTALS(TotalQty, TotalQtyBase, TotalAmount, TotalLineDisc, TotalExciseBaseAmount, TotalBEDAmount, TotaleCessAmount,
                        TotalSHECessAmount, TotalExciseAmount, TotalFOCAmount);
                        CurrReport.CREATETOTALS(TotalDISCOUNTAmount, TotalTransportSubsidy, TotalCashDiscount, TotalENTRYTax,
                        TotalCessAmount, TotalFRIEGHTAmount, TotalAddTaxAmt, TotalTaxBAse, TotalTaxAmount, TotalAmtCust);
                        CurrReport.CREATETOTALS(TotalStateDiscount, TotalTaxAmt, "TotalAddl.Tax&Cess", TSurChargeInterest);
                        CurrReport.CREATETOTALS(ProdDisc, ProdDiscPerlt);
                        //<<1

                        SILCount := COUNT;//23Feb2017 RB-N
                        TILCount := 0;//23Feb2017 RB-N

                        //RSPLSUM 09Oct2020>>
                        TIGST := 0;
                        TCGST := 0;
                        TSGST := 0;
                        TUTGST := 0;
                        //RSPLSUM 09Oct2020<<
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    //>>1


                    //AR
                    //IRNGeneratedTxt:='';
                    "Sales Invoice Header".CALCFIELDS(IRN, "EWB Date", "EWB No."); //AR
                    IF "Sales Invoice Header".IRN <> '' THEN
                        IRNGeneratedTxt := 'Yes'
                    ELSE
                        IRNGeneratedTxt := 'No';

                    //AR


                    //AR 25Nov20


                    VehicleNo := '';
                    DetailedEWayBillRec.RESET;
                    DetailedEWayBillRec.SETRANGE(DetailedEWayBillRec."Document No.", "Sales Invoice Header"."No.");
                    IF DetailedEWayBillRec.FINDLAST THEN
                        VehicleNo := DetailedEWayBillRec."Vehicle No."
                    ELSE
                        VehicleNo := "Sales Invoice Header"."Vehicle No.";

                    //AR 25Nov20

                    vGrpFrieghtNonTx := 0;  //RSPL008

                    SalesType := '';
                    LocGSTNo := '';//RB-N 11Sep2017
                    IF Cust.GET("Sales Invoice Header"."Sell-to Customer No.") THEN;
                    IF LocationRec.GET("Sales Invoice Header"."Location Code") THEN BEGIN
                        LocationCode := LocationRec.Name;
                        LocationSate := LocationRec."State Code";
                        LocGSTNo := LocationRec."GST Registration No.";//RB-N 11Sep2017
                    END;

                    //>>RB-N 11Sep2017 GSTNo. Filter

                    IF TempGSTNo <> '' THEN BEGIN
                        IF TempGSTNo <> LocGSTNo THEN
                            CurrReport.SKIP;

                    END;
                    //<<RB-N 11Sep2017 GSTNo. Filter


                    StateCode := '';
                    GSTStateCode := ''; //31July2017
                    OrgState := ''; //01Aug2017
                    recState.RESET;
                    recState.SETRANGE(recState.Code, LocationRec."State Code");
                    IF recState.FINDFIRST THEN BEGIN
                        StateCode := '';// recState."Std State Code";
                        GSTStateCode := recState."State Code (GST Reg. No.)";//31July2017
                        OrgState := recState.Description;//01Aug2017
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
                            REGION := '';// recStateTable."Region Code";
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
                                SalesType := 'Intra-State' //01Aug2017
                                                           //SalesType := 'Local'
                            ELSE
                                SalesType := 'Inter-State';
                        END
                        ELSE BEGIN
                            IF "Sales Invoice Header"."Ship-to Code" <> '' THEN BEGIN
                                IF Ship2State = LocationSate THEN
                                    SalesType := 'Intra-State' //01Aug2017
                                                               //SalesType := 'Local'
                                ELSE
                                    SalesType := 'Inter-State';
                            END
                        END
                    END
                    ELSE
                        IF "Sales Invoice Header"."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            SalesType := 'Export';

                    CLEAR(SPRegionCode);
                    CLEAR(SPZoneCode);//RSPLSUM 16Apr2020
                    Saleperson.RESET;
                    Saleperson.SETRANGE(Saleperson.Code, "Sales Invoice Header"."Salesperson Code");
                    IF Saleperson.FINDFIRST THEN BEGIN
                        SalesName := Saleperson.Name;
                        SPRegionCode := Saleperson."Region Code";//19Jun2019
                        SPZoneCode := Saleperson."Zone Code";//RSPLSUM 16Apr2020
                    END;

                    //RSPLSUM 23Jun2020>>
                    CLEAR(BDNNo);
                    CLEAR(BDNDate);
                    CLEAR(VesselName);
                    RecSIHAddInfo.RESET;
                    IF RecSIHAddInfo.GET("Sales Invoice Header"."No.") THEN BEGIN
                        RecSIHAddInfo.CALCFIELDS("Vessel Name");
                        BDNNo := RecSIHAddInfo."BDN No.";
                        BDNDate := RecSIHAddInfo."BDN Date";
                        VesselName := RecSIHAddInfo."Vessel Name";
                    END;
                    //RSPLSUM 23Jun2020<<

                    //RSPLSUM 29Jul2020>>-- To find GRN Vessel--
                    i := 1;
                    k := 1;
                    CLEAR(TestGRNNo);
                    CLEAR(PrintGRNNo);
                    GRNFound := FALSE;
                    txtGRN := '';
                    txtGRNNo := '';//RSPLSUM 04Aug2020
                    txtSalRetRcpt := '';//RSPLSUM 05Aug2020
                    IF ("Sales Invoice Header"."Customer Posting Group" = 'FO') OR
                        ("Sales Invoice Header"."Customer Posting Group" = 'BOILS') THEN BEGIN//RSPLSUM 04Aug2020--BOILS added--
                        RecVE.RESET;
                        RecVE.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                        IF RecVE.FINDFIRST THEN BEGIN
                            IF RecVE."Item Ledger Entry No." <> 0 THEN BEGIN
                                ILERec.RESET;
                                IF ILERec.GET(RecVE."Item Ledger Entry No.") THEN BEGIN
                                    FindGRNYSR(ILERec, txtGRN, txtGRNNo, txtSalRetRcpt);
                                END;
                            END;
                        END;
                    END;
                    //RSPLSUM 29Jul2020<<

                    ShipmentAgentName := '';
                    RecShiptoAddress.RESET;
                    RecShiptoAddress.SETRANGE(RecShiptoAddress.Code, "Sales Invoice Header"."Shipping Agent Code");
                    IF RecShiptoAddress.FINDFIRST THEN BEGIN
                        ShipmentAgentName := RecShiptoAddress.Name;
                    END;

                    /*
                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.","Sales Invoice Header"."No.");
                    recSIH.SETRANGE(recSIH."Order No.",'');
                    recSIH.SETRANGE(recSIH."Supplimentary Invoice",FALSE);
                    IF recSIH.FINDFIRST THEN
                    BEGIN
                      CurrReport.SKIP;
                    END;
                    *///RSPLSUM 04Apr2020

                    CLEAR(freightamt);
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
                    CLEAR(statedesc);//01Aug2017
                    CLEAR(FinalState);//01Aug2017
                    CLEAR(GSTPlaceofSupply);
                    CLEAR(GSTSupplyCode);
                    IF "Sales Invoice Header"."Ship-to Code" = '' THEN BEGIN
                        recState2.RESET;
                        recState2.SETRANGE(recState2.Code, "Sales Invoice Header".State);
                        IF recState2.FINDFIRST THEN BEGIN
                            statedesc := recState2.Description;
                            FinalState := recState2.Description;//01Aug2017
                            GSTPlaceofSupply := "Sales Invoice Header"."Bill-to City";
                            GSTSupplyCode := recState2."State Code (GST Reg. No.)";
                        END;
                    END ELSE BEGIN

                        recShipaddress.RESET;
                        recShipaddress.SETRANGE("Customer No.", "Bill-to Customer No.");
                        recShipaddress.SETRANGE(Code, "Sales Invoice Header"."Ship-to Code");
                        IF recShipaddress.FINDFIRST THEN BEGIN
                            recState2.RESET;
                            recState2.SETRANGE(recState2.Code, recShipaddress.State);
                            IF recState2.FINDFIRST THEN BEGIN
                                statedesc := recState2.Description;
                                FinalState := recState2.Description;//01Aug2017
                                GSTSupplyCode := recState2."State Code (GST Reg. No.)";
                            END;
                            GSTPlaceofSupply := "Sales Invoice Header"."Ship-to City";
                        END;
                    END;
                    //<<1

                end;

                trigger OnPostDataItem()
                begin
                    //>>23Feb2017 RB-N, InvoiceGrandTotal
                    IF NNNNNN <> 0 THEN //28Apr2017
                        IF PrintToExcel AND SHowInvTotal AND SalesInvoice THEN BEGIN
                            MakeExcelInvGrandTotal;
                        END;

                    NNNNNN := 0;//02Aug2017


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

                    //<<
                end;

                trigger OnPreDataItem()
                begin
                    //>>1
                    "Sales Invoice Header".SETRANGE("Sales Invoice Header"."Posting Date", StartDate, EndDate);
                    StartDate := "Sales Invoice Header".GETRANGEMIN("Sales Invoice Header"."Posting Date");
                    EndDate := "Sales Invoice Header".GETRANGEMAX("Sales Invoice Header"."Posting Date");

                    Respocibilityfilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Shortcut Dimension 2 Code");

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
                    //>>28Apr2017 RB-N
                    //NNNNNN := COUNT; //28Apr2017
                end;
            }
            dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
            {
                DataItemLink = "Location Code" = FIELD(Code);
                DataItemTableView = SORTING("Location Code", "Posting Date", "No.")
                                    WHERE("Applies-to Doc. No." = FILTER(<> ''));
                dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
                {
                    DataItemLink = "Document No." = FIELD("No.");
                    DataItemTableView = SORTING("Document No.", "Line No.")
                                        WHERE(Type = FILTER('Item|Charge (Item)|G/L Account'),
                                              Quantity = FILTER(<> 0),
                                              "Item Category Code" = FILTER(<> 'CAT14'));

                    trigger OnAfterGetRecord()
                    begin
                        TCLCount += 1;//23Feb2017 RB-N

                        IF ("Sales Cr.Memo Line".Type = "Sales Cr.Memo Line".Type::"G/L Account") AND //RSPLSUM
                            ("Sales Cr.Memo Line"."No." = '74012210') THEN//RSPLSUM
                            CurrReport.SKIP;//RSPLSUM

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
                        /*
                        IF "Sales Cr.Memo Line".Type = "Sales Cr.Memo Line".Type :: "G/L Account" THEN
                        BEGIN
                         CurrReport.SKIP;
                        END;
                        *///RSPLSUM 04Apr2020
                          //>>01Aug2017 RB-N
                        CCCCCC += 1;//01Aug2017
                        IF CCCCCC = 1 THEN //01Aug2017
                            IF PrintToExcel AND SalesCreditMemo THEN BEGIN
                                MakecrmemoHeader;
                            END;
                        //>>

                        CrFCGROSSINV := 0;//"Sales Cr.Memo Line"."Amount To Customer";  // EBT MILAN

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
                            MRPPricePerLt := 0// "Sales Cr.Memo Line"."MRP Price";
                        END ELSE BEGIN
                            //  IF "Sales Cr.Memo Line"."Abatement %" <> 0 THEN
                            MRPPricePerLt := 0;// "Sales Cr.Memo Line"."MRP Price";
                        END;

                        //  IF "Sales Cr.Memo Line"."Price Inclusive of Tax" THEN
                        ListPricePerlt := "Sales Cr.Memo Line"."Unit Price Incl. of Tax" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
                        //   ELSE
                        ListPricePerlt := "Sales Cr.Memo Line"."Unit Price" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";

                        //>>30Dec2017
                        UPricePerLT := 0;
                        UPrice := 0;
                        INRprice := 0;
                        INRamt := 0;

                        UPricePerLT := "Unit Price" / "Qty. per Unit of Measure";
                        UPrice := "Unit Price";
                        INRprice := "Unit Price" / "Quantity (Base)" * Quantity;
                        //INRamt      := UPrice * Quantity;
                        //RSPLSUM 10Nov2020--INRamt      := UPrice * Quantity +("National Discount"* "Quantity (Base)") ;//15Jan2019
                        //RSPLSUM 04Dec2020--INRamt := "Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount" + "Sales Cr.Memo Line"."Charges To Customer";//RSPLSUM 10Nov2020
                        INRamt := "Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount";//RSPLSUM 04Dec2020

                        //<<30Dec2017
                        /*
                        ExciseSetup.RESET;
                        ExciseSetup.SETRANGE("Excise Bus. Posting Group","Excise Bus. Posting Group");
                        ExciseSetup.SETRANGE("Excise Prod. Posting Group","Excise Prod. Posting Group");
                        IF ExciseSetup.FINDLAST THEN;
                        
                        StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.","Item No.","Line No.","Tax/Charge Type","Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group",'%1|%2','FOC','FOC DIS');
                        IF StructureDetailes.FINDSET THEN
                        REPEAT
                          FOCAmount += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0;
                        
                        StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.","Item No.","Line No.","Tax/Charge Type","Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group",'STATE DISC');
                        IF StructureDetailes.FINDSET THEN
                        REPEAT
                          StateDiscount += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0;
                        
                        
                        StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.","Item No.","Line No.","Tax/Charge Type","Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group",'%1|%2|%3|%4|%5|%6|%7|%8|%9','ADD-DIS','AVP-SP-SAN',
                                                   'MONTH-DIS','NEW-PRI-SP','SPOT-DISC','VAT-DIF-DI','T2 DISC');
                        IF StructureDetailes.FINDSET THEN
                        REPEAT
                          DISCOUNTAmount += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0;
                        */
                        //>>01Aug2017

                        CLEAR(AVPDiscAmt);
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        //StructureDetailes.SETFILTER("Tax/Charge Group",'AVP-SP-SAN');
                        StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'AVP-SP-SAN', 'NATSCHEME');
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            AVPDiscAmt := ABS(StructureDetailes.Amount);
                        END; */
                        //IF AVPDiscAmt = 0 THEN
                        AVPDiscAmt += ("National Discount" * "Quantity (Base)");//15Jan2019

                        //>>02May2019
                        IF "Free of Cost" THEN
                            AVPDiscAmt += ABS("Line Discount Amount");
                        //<<02May2019

                        TAVPDiscAmt += AVPDiscAmt;
                        TTAVPDiscAmt += AVPDiscAmt;

                        CLEAR(DiscAmt);
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", 'DISCOUNT');
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            DiscAmt := StructureDetailes.Amount;
                        END; */
                        TDiscAmt += DiscAmt;
                        TTDiscAmt += DiscAmt;

                        //<<01Aug2017

                        //>>13Mar2018
                        CLEAR(CessAmt);
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", 'CESS');
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            CessAmt := StructureDetailes.Amount;
                        END; */
                        TCessAmt += CessAmt;
                        TTCessAmt += CessAmt;

                        CLEAR(TCSAmt);
                        TCSAmt := 0;// "TDS/TCS Amount";
                        TTCSAmt += TCSAmt;
                        TTTCSAmt += TCSAmt;
                        //<<13Mar2018

                        //>>02Nov2018
                        CLEAR(RegDiscAmt);
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", 'REGSCHEME');
                        IF StructureDetailes.FINDFIRST THEN BEGIN
                            RegDiscAmt := StructureDetailes.Amount;
                        END; */
                        TRegDiscAmt += RegDiscAmt;
                        TTRegDiscAmt += RegDiscAmt;
                        //<<02Nov2018
                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'TPTSUBSIDY', 'TRANSSUBS', 'TRANSSUBS2');
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

                        /* StructureDetailes.RESET;
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
                            UNTIL StructureDetailes.NEXT = 0; */

                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'FREIGHT');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 //RSPL008 >>>> FRIEGHT NON Taxable Amount
                                 StructureDetailes.CALCFIELDS("Structure Description");
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
                            UNTIL recDetailedTaxEntry.NEXT = 0; */

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
                        /* recDetailedTaxEntry1.RESET;
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Document No.", "Sales Cr.Memo Line"."Document No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."No.", "Sales Cr.Memo Line"."No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Document Line No.", "Sales Cr.Memo Line"."Line No.");
                        recDetailedTaxEntry1.SETRANGE(recDetailedTaxEntry1."Tax Jurisdiction Code", "Sales Cr.Memo Line"."Tax Area Code");
                        IF recDetailedTaxEntry1.FINDFIRST THEN
                            REPEAT
                                TaxAmt1 := recDetailedTaxEntry1."Tax Amount";
                            UNTIL recDetailedTaxEntry.NEXT = 0; */

                        /*
                        CLEAR("Addl.Tax&CessRate1");
                        StructureDetailes.RESET;
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::"Other Taxes");
                        StructureDetailes.SETFILTER("Tax/Charge Code",'%1|%2|%3','CESS','C1','ADDL TAX');
                        IF StructureDetailes.FINDFIRST THEN
                        BEGIN
                         "Addl.Tax&CessRate1" := StructureDetailes."Calculation Value";
                        END;
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END
                        
                        //CLEAR("Addl.Tax&Cess1");
                        StructureDetailes.RESET;
                        StructureDetailes.SETRANGE("Invoice No.","Document No.");
                        StructureDetailes.SETRANGE("Item No.","No.");
                        StructureDetailes.SETRANGE("Line No.","Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type",StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group",'CESS');
                        IF StructureDetailes.FINDSET THEN
                        REPEAT
                          //CessAmount += StructureDetailes.Amount;
                          "Addl.Tax&Cess1" += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0;
                        */

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
                        //>>27Feb2017
                        recDimSet.RESET;
                        //recDimSet.SETCURRENTKEY("Dimension Set ID");
                        recDimSet.CALCFIELDS("Dimension Value Name");
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
                        
                        *///Commented T359..PostedDimensionValue
                        CLEAR(LineSalespersonCode1);
                        CLEAR(LineSalespersonName1);
                        //>>27Feb2017
                        recDimSet.RESET;
                        //recDimSet.SETCURRENTKEY("Dimension Set ID");
                        recDimSet.CALCFIELDS("Dimension Value Name");
                        recDimSet.SETRANGE("Dimension Set ID", "Sales Cr.Memo Line"."Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALESPERSON');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            LineSalespersonCode1 := recDimSet."Dimension Value Code";
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
                        *///Commented T359..PostedDimensionValue
                          // EBT MILAN Product Disc........

                        CLEAR(crProdDisc);
                        CLEAR(crProdDiscPerlt);
                        /*
                        MrpMaster.RESET;
                        MrpMaster.SETRANGE(MrpMaster."Item No.","Sales Cr.Memo Line"."No.");
                        MrpMaster.SETRANGE(MrpMaster."Lot No.","Sales Cr.Memo Line"."Lot No.");
                        IF MrpMaster.FINDFIRST THEN
                        BEGIN
                          crProdDiscPerlt := MrpMaster."National Discount";
                          crProdDisc := CrProdDiscPerlt * "Sales Cr.Memo Line"."Quantity (Base)";
                        END;
                        */
                        //>>06Feb2019
                        //crProdDiscPerlt := "Sales Cr.Memo Line"."National Discount" ;//RSPLSUM
                        //crProdDisc := "Sales Cr.Memo Line"."National Discount" * "Sales Cr.Memo Line"."Quantity (Base)";//RSPLSUM
                        //<<06Feb2019

                        //RSPLSUM>>
                        crProdDisc := 0;
                        crProdDiscPerlt := 0;
                        IF "Sales Cr.Memo Header"."Posting Date" <= 20190303D THEN BEGIN //030319D
                            crProdDisc := "Sales Cr.Memo Line"."National Discount" * "Sales Cr.Memo Line"."Quantity (Base)";
                            crProdDiscPerlt := "Sales Cr.Memo Line"."National Discount";
                        END ELSE BEGIN
                            crProdDisc := 0;
                            crProdDiscPerlt := 0;
                        END;
                        //RSPLSUM<<

                        TotalQty += "Sales Cr.Memo Line".Quantity;
                        TotalQtyBase += "Sales Cr.Memo Line"."Quantity (Base)";
                        TotalAmount += "Sales Cr.Memo Line".Amount;
                        TotalLineDisc += "Sales Cr.Memo Line"."Line Discount Amount";
                        TotalExciseBaseAmount += 0;// "Sales Cr.Memo Line"."Excise Base Amount";
                        TotalBEDAmount += 0;// "Sales Cr.Memo Line"."BED Amount";
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
                        IF /*("Sales Cr.Memo Line"."Price Inclusive of Tax" = TRUE) AND */(MRPPricePerLt <> 0) AND ("Sales Cr.Memo Header"."Currency Code"
                        = '') THEN BEGIN
                            TotalAmtincludeExcise += (-"Sales Cr.Memo Line".Amount + 0);//-"Sales Cr.Memo Line"."Excise Amount");
                            // GTotalAmtincludeExcise += (-"Sales Cr.Memo Line".Amount+-"Sales Cr.Memo Line"."Excise Amount");
                        END
                        ELSE BEGIN
                            TotalAmtincludeExcise += 0;// (-"Sales Cr.Memo Line"."Amount Including Excise");
                            //  GTotalAmtincludeExcise += (-"Sales Cr.Memo Line"."Amount Including Excise");
                        END;


                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure -------START
                        TotalTaxAmt1 += TaxAmt1;
                        "TotalAddl.Tax&Cess1" += "Addl.Tax&Cess1";
                        //EBT STIVAN---(08112012)---To Capture Addl. Tax / Cess or Tax Amount as per New Structure ---------END


                        IF "Sales Cr.Memo Header"."Currency Code" <> '' THEN BEGIN
                            CrtFCFRIEGHTAmount += CrFCFRIEGHTAmount;  // EBT MILAN
                            CrgtotalFCFRIEGHTAmount += CrtFCFRIEGHTAmount; //EBT MILAN
                            CrtotalFCGROSSINV += CrFCGROSSINV;         // EBT MILAN
                            CrGtotalFCGROSSINV += CrtotalFCGROSSINV;   // EBT MILAN
                        END;


                        IF "Sales Cr.Memo Line"."Tax Liable" THEN
                            TotalTaxBAse += 0;// "Sales Cr.Memo Line"."Tax Base Amount";

                        TotalTaxAmount += 0;// "Sales Cr.Memo Line"."Tax Amount";
                        TotalAmtCust += 0;// "Sales Cr.Memo Line"."Amount To Customer";

                        TSurChargeInterest += SurChargeInterest;


                        //>>31July2017 GST
                        CLEAR(vCGST);
                        CLEAR(vSGST);
                        CLEAR(vIGST);
                        CLEAR(vUTGST);
                        CLEAR(vCGSTRate);
                        CLEAR(vSGSTRate);
                        CLEAR(vIGSTRate);
                        CLEAR(GSTPer);//RB-N 20Sep2017

                        DetailGSTEntry.RESET;
                        DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                        //DetailGSTEntry.SETRANGE(Type, Type);
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
                        IF DetailGSTEntry.FINDSET THEN
                            REPEAT

                                GSTPer += DetailGSTEntry."GST %";//RB-N 20Sep2017
                                IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                    vCGST := ABS(DetailGSTEntry."GST Amount");
                                    vCGSTRate := DetailGSTEntry."GST %";
                                END;

                                IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                    vSGST := ABS(DetailGSTEntry."GST Amount");
                                    vSGSTRate := DetailGSTEntry."GST %";

                                END;

                                IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                    vUTGST := ABS(DetailGSTEntry."GST Amount");
                                    vSGSTRate := DetailGSTEntry."GST %";

                                END;


                                IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                    vIGST := ABS(DetailGSTEntry."GST Amount");
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
                        //<<31July2017 GST

                        //>>08Apr2019
                        LineDiscAmt := 0;
                        IF "Sales Cr.Memo Header"."Posting Date" <= 20190303D THEN //030319D
                        BEGIN
                            LineDiscAmt := "Line Discount Amount";
                        END ELSE BEGIN
                            LineDiscAmt := 0;
                        END;
                        //<<08Apr2019

                        //IF PrintToExcel AND SalesCreditMemo THEN
                        IF PrintToExcel AND SalesCreditMemo AND (PrintSummary = FALSE) THEN //RB-N 20Sep2017
                        BEGIN
                            MakeCrMemoBody;
                        END;


                        //>>20Sep2017 GroupData
                        CrTQty += "Sales Cr.Memo Line".Quantity;
                        InCrTQty += "Sales Cr.Memo Line".Quantity;
                        CrTQtyBase += "Sales Cr.Memo Line"."Quantity (Base)";
                        InCrTQtyBase += "Sales Cr.Memo Line"."Quantity (Base)";
                        CrTrans += TransportSubsidy;
                        InCrTrans += TransportSubsidy;
                        //CrLineDisc += "Sales Cr.Memo Line"."Line Discount Amount";
                        CrLineDisc += LineDiscAmt;//08Apr2019
                        //InCrLineDisc += "Sales Cr.Memo Line"."Line Discount Amount";
                        InCrLineDisc += LineDiscAmt;//08Apr2019

                        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
                            TFRIEGHTAmount += 0;// "Sales Cr.Memo Line"."Charges To Customer";//23Feb2017
                                                //IF "Sales Cr.Memo Line"."GST Base Amount" <> 0 THEN BEGIN
                            InCrTaxBase += 0;// "Sales Cr.Memo Line"."GST Base Amount";//27Feb
                            CrTaxBase += 0;//"Sales Cr.Memo Line"."GST Base Amount";
                                           //END ELSE BEGIN
                            InCrTaxBase += 0;// "Sales Cr.Memo Line"."Tax Base Amount";
                            CrTaxBase += 0;// "Sales Cr.Memo Line"."Tax Base Amount";
                                           // END;
                            CrAmt += 0;// "Sales Cr.Memo Line"."Amount To Customer";
                            InCrAmt += 0;// "Sales Cr.Memo Line"."Amount To Customer";
                            //RSPLSUM 10Nov2020--TINRAmt += "Sales Cr.Memo Line".Amount;//23Feb2017
                            //RSPLSUM 04Dec2020--TINRAmt += "Sales Cr.Memo Line".Amount+"Sales Cr.Memo Line"."Line Discount Amount"+"Sales Cr.Memo Line"."Charges To Customer";//RSPLSUM 10Nov2020
                            TINRAmt += "Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount";//RSPLSUM 04Dec2020
                                                                                                                 //RSPLSUM 10Nov2020--InTINRAmt += "Sales Cr.Memo Line".Amount;//27Feb2017
                                                                                                                 //RSPLSUM 04Dec2020--InTINRAmt += "Sales Cr.Memo Line".Amount+"Sales Cr.Memo Line"."Line Discount Amount"+"Sales Cr.Memo Line"."Charges To Customer";//RSPLSUM 10Nov2020
                            InTINRAmt += "Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount";//RSPLSUM 04Dec2020
                            TINRUnit += UPrice;//23Feb2017
                            InTINRUnit += UPrice;//27Feb2017
                        END;


                        IF "Sales Cr.Memo Header"."Currency Code" <> '' THEN BEGIN
                            CrFCAmt += 0;// "Sales Cr.Memo Line"."Amount To Customer";
                            InCrFCAmt += 0;//"Sales Cr.Memo Line"."Amount To Customer";

                            CrTaxBase += "Sales Cr.Memo Line"."Line Amount" / "Sales Cr.Memo Header"."Currency Factor";
                            InCrTaxBase += "Sales Cr.Memo Line"."Line Amount" / "Sales Cr.Memo Header"."Currency Factor";//27Feb

                            TFRIEGHTAmount += 0;//"Sales Cr.Memo Line"."Charges To Customer" / "Sales Cr.Memo Header"."Currency Factor";//23Feb2017

                            InFG += (/*"Sales Cr.Memo Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt) / "Sales Cr.Memo Header"."Currency Factor";
                            TInFG += (/*"Sales Cr.Memo Line"."Charges To Customer"*/ -ENTRYTax - CessAmount - AddTaxAmt) / "Sales Cr.Memo Header"."Currency Factor";

                            TINRUnit += UPrice / "Sales Cr.Memo Header"."Currency Factor";//23Feb2017
                            InTINRUnit += UPrice / "Sales Cr.Memo Header"."Currency Factor";//27Feb2017
                                                                                            //RSPLSUM 10Nov2020--TINRAmt += "Sales Cr.Memo Line".Amount /"Sales Cr.Memo Header"."Currency Factor";//23Feb2017
                                                                                            //RSPLSUM 04Dec2020--TINRAmt += ("Sales Cr.Memo Line".Amount+"Sales Cr.Memo Line"."Line Discount Amount"+"Sales Cr.Memo Line"."Charges To Customer") /"Sales Cr.Memo Header"."Currency Factor";//RSPLSUM 10Nov2020
                            TINRAmt += ("Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount") / "Sales Cr.Memo Header"."Currency Factor";//RSPLSUM 04Dec2020
                                                                                                                                                              //RSPLSUM 10Nov2020--InTINRAmt += "Sales Cr.Memo Line".Amount /"Sales Cr.Memo Header"."Currency Factor";//27Feb2017
                                                                                                                                                              //RSPLSUM 04Dec2020--InTINRAmt += ("Sales Cr.Memo Line".Amount+"Sales Cr.Memo Line"."Line Discount Amount"+"Sales Cr.Memo Line"."Charges To Customer") /"Sales Cr.Memo Header"."Currency Factor";//RSPLSUM 10Nov2020
                            InTINRAmt += ("Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount") / "Sales Cr.Memo Header"."Currency Factor";//RSPLSUM 04Dec2020
                            CrAmt += 0;// "Sales Cr.Memo Line"."Amount To Customer" / "Sales Cr.Memo Header"."Currency Factor";
                            InCrAmt += 0;// "Sales Cr.Memo Line"."Amount To Customer" / "Sales Cr.Memo Header"."Currency Factor";

                        END;
                        //<<20Sep2017 GroupData

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

                        //<<1

                        //>>23Feb2017 CreditMemoGroupTotal

                        IF TCLCount = SCLCount THEN BEGIN
                            //IF PrintToExcel AND SHowInvTotal AND SalesCreditMemo THEN
                            IF PrintToExcel AND SalesCreditMemo THEN //20Sep2017
                            BEGIN
                                MakeExcelCrMemoGrouping;
                                TCLCount := 0;//27
                            END;
                            CrTQty := 0;
                            CrTQtyBase := 0;
                            CrINRAMt := 0;
                            CrLineDisc := 0;
                            CrBaseAmt := 0;
                            CrBED := 0;
                            CrEcess := 0;
                            CrScess := 0;
                            CrExcise := 0;
                            CrAmtInc := 0;
                            CrFOC := 0;
                            CrState := 0;
                            CrDiscAmt := 0;
                            CrTrans := 0;
                            CrCash := 0;
                            CrSurC := 0;
                            CrEntry := 0;
                            CrTaxBase := 0;
                            CrTaxAMt := 0;
                            CrAddTax := 0;
                            InCrAddTax := 0;
                            CrFRAmt := 0;
                            CrFCAmt := 0;
                            CrAmt := 0;

                        END;

                        //<<

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

                        SCLCount := COUNT;//23Feb2017 RB-N
                        TCLCount := 0;//23Feb2017 RB-N

                        //RSPLSUM 09Oct20>>
                        RecSCML.RESET;
                        RecSCML.SETCURRENTKEY("Document No.", Type, Quantity, "Item Category Code");
                        RecSCML.SETRANGE("Document No.", "Sales Cr.Memo Header"."No.");
                        RecSCML.SETFILTER(Type, '%1|%2|%3', RecSCML.Type::Item, RecSCML.Type::"Charge (Item)", RecSCML.Type::"G/L Account");
                        RecSCML.SETFILTER(Quantity, '<>%1', 0);
                        RecSCML.SETFILTER("Item Category Code", '<>%1', 'CAT14');
                        IF RecSCML.FINDSET THEN
                            REPEAT
                                IF (RecSCML.Type = RecSCML.Type::"G/L Account") AND (RecSCML."No." = '74012210') THEN
                                    SCLCount -= 1;
                            UNTIL RecSCML.NEXT = 0;
                        //RSPLSUM 09Oct20<<
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    "Sales Cr.Memo Header".CALCFIELDS(IRN);

                    IF "Sales Cr.Memo Header".IRN <> '' THEN
                        IRNGeneratedTxtCrMemo := 'Yes'
                    ELSE
                        IRNGeneratedTxtCrMemo := 'No';


                    //>>1
                    vGrpFrieghtNonTx := 0;  //RSPL008

                    SalesType := '';
                    LocGSTNo := '';//RB-N 11Sep2017
                    IF Cust.GET("Sales Cr.Memo Header"."Sell-to Customer No.") THEN;
                    IF LocationRec.GET("Sales Cr.Memo Header"."Location Code") THEN BEGIN
                        LocationCode := LocationRec.Name;
                        LocGSTNo := LocationRec."GST Registration No.";//RB-N 11Sep2017
                    END;

                    //>>RB-N 11Sep2017 GSTNo. Filter

                    IF TempGSTNo <> '' THEN BEGIN
                        IF TempGSTNo <> LocGSTNo THEN
                            CurrReport.SKIP;
                    END;
                    //<<RB-N 11Sep2017 GSTNo. Filter

                    StateCode := '';
                    GSTStateCode := ''; //31July2017
                    OrgState := ''; //01Aug2017
                    recState.RESET;
                    recState.SETRANGE(recState.Code, LocationRec."State Code");
                    IF recState.FINDFIRST THEN BEGIN
                        StateCode := '';// recState."Std State Code";
                        GSTStateCode := recState."State Code (GST Reg. No.)"; //31July2017
                        OrgState := recState.Description;//01Aug2017
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
                            REGION := '';// recStateTable."Region Code";
                        END;
                    END;


                    IF "Sales Cr.Memo Header"."Gen. Bus. Posting Group" = 'DOMESTIC' THEN BEGIN
                        IF LocationSate = "Sales Cr.Memo Header".State THEN
                            SalesType := 'Intra-State' //01Aug2017
                                                       //SalesType := 'Local'
                        ELSE
                            SalesType := 'Inter-State';
                    END
                    ELSE
                        IF "Sales Cr.Memo Header"."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            SalesType := 'Export';

                    CLEAR(SPRegionCode);
                    CLEAR(SPZoneCode);//RSPLSUM 16Apr2020
                    Saleperson.RESET;
                    Saleperson.SETRANGE(Saleperson.Code, "Sales Cr.Memo Header"."Salesperson Code");
                    IF Saleperson.FINDFIRST THEN BEGIN
                        SalesName := Saleperson.Name;
                        SPRegionCode := Saleperson."Region Code";//19Jun2019
                        SPZoneCode := Saleperson."Zone Code";//RSPLSUM 16Apr2020
                    END;

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
                    CLEAR(BDNNo);//RSPLSUM 23Jun2020
                    CLEAR(BDNDate);//RSPLSUM 23Jun2020
                    CLEAR(VesselName);//RSPLSUM 23Jun2020
                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.", "Sales Cr.Memo Header"."Applies-to Doc. No.");
                    IF recSIH.FINDFIRST THEN BEGIN
                        InvoiceDate := recSIH."Posting Date";
                        InvoiceNo := recSIH."No.";

                        //RSPLSUM 23Jun2020>>
                        RecSIHAddInfo.RESET;
                        IF RecSIHAddInfo.GET(recSIH."No.") THEN BEGIN
                            RecSIHAddInfo.CALCFIELDS("Vessel Name");
                            BDNNo := RecSIHAddInfo."BDN No.";
                            BDNDate := RecSIHAddInfo."BDN Date";
                            VesselName := RecSIHAddInfo."Vessel Name";
                        END;
                        //RSPLSUM 23Jun2020<<

                    END;

                    //RSPL--
                    CLEAR(GSTPlaceofSupply);//31July2017
                    CLEAR(GSTSupplyCode);//31July2017
                    CLEAR(statedesc);
                    IF "Sales Cr.Memo Header"."Ship-to Code" = '' THEN BEGIN
                        recState2.RESET;
                        recState2.SETRANGE(recState2.Code, "Sales Cr.Memo Header".State);
                        IF recState2.FINDFIRST THEN BEGIN
                            statedesc := recState2.Description;
                            GSTPlaceofSupply := "Sales Cr.Memo Header"."Bill-to City";
                            GSTSupplyCode := recState2."State Code (GST Reg. No.)";
                        END;

                    END ELSE BEGIN
                        recShipaddress.RESET;
                        recShipaddress.SETRANGE("Customer No.", "Bill-to Customer No.");
                        recShipaddress.SETRANGE(Code, "Ship-to Code");
                        IF recShipaddress.FINDFIRST THEN BEGIN
                            recState2.RESET;
                            recState2.SETRANGE(recState2.Code, recShipaddress.State);
                            IF recState2.FINDFIRST THEN BEGIN
                                statedesc := recState2.Description;
                                GSTSupplyCode := recState2."State Code (GST Reg. No.)";
                            END;
                            GSTPlaceofSupply := "Ship-to City";
                        END;
                    END;



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
                    //>>23Feb2017 RB-N, CreditMemoGrandTotal
                    IF CCCCCC <> 0 THEN //28Apr2017
                        IF PrintToExcel AND SHowInvTotal AND SalesCreditMemo THEN BEGIN
                            MakeExcelCrmemoGrandTotal;
                        END;
                    //<<

                    CCCCCC := 0;//02Aug2017
                end;

                trigger OnPreDataItem()
                begin



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
                        //  "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Excise Bus. Posting Group",
                        // "Sales Invoice Header"."Excise Bus. Posting Group");
                    END;

                    IF "Sales Invoice Header".GETFILTER("Sales Invoice Header"."No.") <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."No.", "Sales Invoice Header".GETFILTER("Sales Invoice Header"."No."));
                    END;

                    IF Respocibilityfilter <> '' THEN BEGIN
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Shortcut Dimension 2 Code", Respocibilityfilter);
                    END;


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
                    //>>28Apr2017 RB-N
                    //CCCCCC := COUNT;//28Apr2017
                    //>>1
                    //>>GST Total Cr
                    CLEAR(TTCGST);
                    CLEAR(TTSGST);
                    CLEAR(TTIGST);
                    CLEAR(TTUTGST);
                    //<<GST Total Cr

                    CLEAR(TCessAmt);//RSPLSUM 08Jul2020
                    CLEAR(TTCSAmt);//RSPLSUM 08Jul2020

                    CLEAR(TAVPDiscAmt);
                    CLEAR(TTAVPDiscAmt);
                    CLEAR(TRegDiscAmt);
                    CLEAR(TTRegDiscAmt);
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
                field("GST Register No."; TempGSTNo)
                {
                    TableRelation = "GST Registration Nos.";
                }
                field("Start Date"; StartDate)
                {
                }
                field("End Date"; EndDate)
                {

                    trigger OnValidate()
                    begin
                        IF EndDate <> 0D THEN
                            IF EndDate < StartDate THEN
                                ERROR('End Date must be greater than Start Date');
                    end;
                }
                field("Print To Excel"; PrintToExcel)
                {
                }
                field("Print Summary"; PrintSummary)
                {
                }
                field("Show Invoice Total"; SHowInvTotal)
                {
                }
                field("Sales Invoice"; SalesInvoice)
                {
                }
                field("Sales Return"; SalesCreditMemo)
                {
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
        //<<
    end;

    trigger OnPreReport()
    begin
        //>>1

        recUserPersonal.RESET;
        recUserPersonal.SETRANGE("User ID", USERID);
        IF recUserPersonal.FINDFIRST THEN BEGIN
            IF recUserPersonal."Profile ID" = 'CFA' THEN
                ERROR('You do not have permission');
        END;

        IF (SalesInvoice = FALSE) AND (SalesCreditMemo = FALSE) THEN
            ERROR('You have not selected any option');

        LocationFilter := Location.GETFILTER(Location.Code);
        SalespersonFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Salesperson Code");
        Sell2CustomerFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Sell-to Customer No.");
        "Resp.CentreFilter" := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Responsibility Center");
        DivisionFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Shortcut Dimension 1 Code");
        "Gen.Bus.PostingGroupFilter" := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Gen. Bus. Posting Group");
        CustomerPostingGroupFilter := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Customer Posting Group");
        // "ExciseBus.PostingGroupFilter" := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Excise Bus. Posting Group");

        IF StartDate = 0D THEN
            ERROR('Please enter the From Date and To Date');
        IF EndDate = 0D THEN
            ERROR('Please enter the From Date and To Date');

        "Company info".GET;

        IF PrintToExcel THEN
            MakeExcelInfo;

        //<<
    end;

    var
        StartDate: Date;
        EndDate: Date;
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        "Company info": Record 79;
        BEDPercent: Decimal;
        //ExciseSetup: Record "13711";
        //StructureDetailes: Record "13798";
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
        LocationCode: Text;
        TotalAddTaxAmt: Decimal;
        AddTaxAmt: Decimal;
        BRTNo: Code[20];
        //RG23D: Record "16537";
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
        ResCenName: Text;
        CustResCenCode: Code[10];
        recCust: Record 18;
        SalesType: Text;
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
        y: Record 349;
        SalesGroup: Text[50];
        recDimVale1: Record 349;
        SalesGroup1: Text[50];
        // recDetailedTaxEntry: Record "16522";
        "Addl.Tax&CessRate": Decimal;
        "Addl.Tax&Cess": Decimal;
        "TotalAddl.Tax&Cess": Decimal;
        "GTotalAddl.Tax&Cess": Decimal;
        TaxRate: Decimal;
        TaxAmt: Decimal;
        TotalTaxAmt: Decimal;
        GTotalTaxAmt: Decimal;
        //recDetailedTaxEntry1: Record "16522";
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
        Text002: Label 'Data';
        Text003: Label 'Sales Register - GST';
        Text004: Label 'Company Name';
        Text005: Label 'Report No.';
        Text006: Label 'Report Name';
        Text007: Label 'User ID';
        Text008: Label 'Date';
        Text009: Label 'Location Filter';
        "--23Feb2017": Integer;
        SILCount: Integer;
        TILCount: Integer;
        SCLCount: Integer;
        TCLCount: Integer;
        TQty: Decimal;
        TQtyBase: Decimal;
        TLineDisc: Decimal;
        TExciseBaseAmount: Decimal;
        TBEDAmount: Decimal;
        TeCessAmount: Decimal;
        TSHECessAmount: Decimal;
        TExciseAmount: Decimal;
        TFOCAmount: Decimal;
        TStateDiscount: Decimal;
        TDISCOUNTAmount: Decimal;
        TTransportSubsidy: Decimal;
        TCashDiscount: Decimal;
        TENTRYTax: Decimal;
        TTaxBAse: Decimal;
        TFRIEGHTAmount: Decimal;
        TTaxAmt: Decimal;
        TINRUnit: Decimal;
        TINRAmt: Decimal;
        TGrossInVAmt: Decimal;
        TFCGrossInVAmt: Decimal;
        InTQty: Decimal;
        InTQtyBase: Decimal;
        InLineDisc: Decimal;
        InTExciseBaseAmount: Decimal;
        InTBEDAmount: Decimal;
        InTeCessAmount: Decimal;
        InTSHECessAmount: Decimal;
        InTExciseAmount: Decimal;
        InTFOCAmount: Decimal;
        InTStateDiscount: Decimal;
        InTDISCOUNTAmount: Decimal;
        InTTransportSubsidy: Decimal;
        InTCashDiscount: Decimal;
        InTENTRYTax: Decimal;
        InTTaxBAse: Decimal;
        InTFRIEGHTAmount: Decimal;
        InTTaxAmt: Decimal;
        InTINRUnit: Decimal;
        InTINRAmt: Decimal;
        InTGrossInVAmt: Decimal;
        InTFCGrossInVAmt: Decimal;
        InTFCGROSSINV: Decimal;
        InAmtInclExcise: Decimal;
        InAssVal: Decimal;
        recDimSet: Record 480;
        "----27Feb2017": Integer;
        CrTQty: Decimal;
        InCrTQty: Decimal;
        CrTQtyBase: Decimal;
        InCrTQtyBase: Decimal;
        CrINRAMt: Decimal;
        InCrINRAMt: Decimal;
        CrLineDisc: Decimal;
        InCrLineDisc: Decimal;
        CrBaseAmt: Decimal;
        InCrBaseAmt: Decimal;
        CrBED: Decimal;
        InCrBED: Decimal;
        CrEcess: Decimal;
        InCrEcess: Decimal;
        CrScess: Decimal;
        InCrScess: Decimal;
        CrExcise: Decimal;
        InCrExcise: Decimal;
        CrAmtInc: Decimal;
        InCrAmtInc: Decimal;
        CrFOC: Decimal;
        InCrFOC: Decimal;
        CrState: Decimal;
        InCrState: Decimal;
        CrDiscAmt: Decimal;
        InCrDiscAmt: Decimal;
        CrTrans: Decimal;
        InCrTrans: Decimal;
        CrCash: Decimal;
        InCrCash: Decimal;
        CrSurC: Decimal;
        InCrSurC: Decimal;
        CrEntry: Decimal;
        InCrEntry: Decimal;
        CrTaxBase: Decimal;
        InCrTaxBase: Decimal;
        CrTaxAMt: Decimal;
        InCrTaxAMt: Decimal;
        CrAddTax: Decimal;
        InCrAddTax: Decimal;
        CrFRAmt: Decimal;
        InCrFRAmt: Decimal;
        CrFCAmt: Decimal;
        InCrFCAmt: Decimal;
        CrAmt: Decimal;
        InCrAmt: Decimal;
        "----Inv27Feb2017": Integer;
        InFCFG: Decimal;
        TInFCFG: Decimal;
        InFG: Decimal;
        TInFG: Decimal;
        "----28Feb2017": Integer;
        InRound: Decimal;
        CrRound: Decimal;
        "---28Apr2017": Integer;
        NNNNNN: Integer;
        CCCCCC: Integer;
        "---31July2017": Integer;
        GSTStateCode: Code[10];
        GSTPlaceofSupply: Text[50];
        GSTSupplyCode: Code[10];
        "----------------GST": Integer;
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        vUTGST: Decimal;
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        TCGST: Decimal;
        TTCGST: Decimal;
        TSGST: Decimal;
        TTSGST: Decimal;
        TIGST: Decimal;
        TTIGST: Decimal;
        TUTGST: Decimal;
        TTUTGST: Decimal;
        FinalState: Text[50];
        OrgState: Text[50];
        InvRound: Decimal;
        AVPDiscAmt: Decimal;
        TAVPDiscAmt: Decimal;
        TTAVPDiscAmt: Decimal;
        DiscAmt: Decimal;
        TDiscAmt: Decimal;
        TTDiscAmt: Decimal;
        "------11Sep2017": Integer;
        TempGSTNo: Code[15];
        LocGSTNo: Code[15];
        "----20Sep2017": Integer;
        PrintSummary: Boolean;
        GSTPer: Decimal;
        "---------25Oct2017": Integer;
        Dim25: Record 349;
        DivisionText: Text[50];
        "----------30Dec2017": Integer;
        FcTotal: Decimal;
        GFcTotal: Decimal;
        "--------13Mar2018": Integer;
        CessAmt: Decimal;
        TCessAmt: Decimal;
        TTCessAmt: Decimal;
        TCSAmt: Decimal;
        TTCSAmt: Decimal;
        TTTCSAmt: Decimal;
        RegDiscAmt: Decimal;
        TRegDiscAmt: Decimal;
        TTRegDiscAmt: Decimal;
        "---------04Mar2019": Decimal;
        LineDiscAmt: Decimal;
        DGST04: Record "Detailed GST Ledger Entry";
        SPRegionCode: Code[10];
        SPZoneCode: Code[10];
        RecSIHAddInfo: Record 50054;
        BDNNo: Text[20];
        BDNDate: Date;
        VesselName: Text[50];
        RecVE: Record 5802;
        ILERec: Record 32;
        GRNFound: Boolean;
        txtGRN: Text;
        i: Integer;
        RecItemLE: Record 32;
        j: Integer;
        TestGRNNo: array[20] of Text;
        GRNForLoopFound: Boolean;
        PrintGRNNo: array[10] of Text;
        k: Integer;
        RecPRH: Record 120;
        GRNVessel: Text;
        txtGRNNo: Text;
        RecValEnt: Record 5802;
        RecSCMH: Record 114;
        RecSalInvHdr: Record 112;
        RecValEntNew: Record 5802;
        ILERecNew: Record 32;
        txtSalRetRcpt: Text;
        recUserPersonal: Record 2000000073;
        RecSCML: Record 115;
        IRNGeneratedTxt: Text;
        VehicleNo: Code[100];
        DetailedEWayBillRec: Record 50044;
        IRNGeneratedTxtCrMemo: Text;

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//23Feb2017
        ExcelBuf.AddInfoColumn(FORMAT(Text004), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text006), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(FORMAT(Text003), FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text005), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(50092, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text007), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.NewRow;
        ExcelBuf.ClearNewRow;

        /*
        //>>23Feb2017 RB-N
        IF PrintToExcel AND SalesInvoice THEN
        BEGIN
         MakeExcelDataHeader;
        END;
        
        //<<
        *///Commented 28Apr2017

        //>>Report Header 28Apr2017

        IF PrintToExcel THEN BEGIN

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
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//39
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//40
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//41
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//42
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//44
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//45
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//46
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//48
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//51
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//52
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//53
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//54
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//56
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//57
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//58
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//59
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//60
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//61
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//62
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//63
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//64
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//65
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//66
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//67
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//68
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//69
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//70
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//71
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//72
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//73
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//74
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//75
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//76
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//77
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//78
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//79
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//80
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//81 13Mar2018
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//82 13Mar2018
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//02Nov2018
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 16Apr2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 29Jul2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 04Aug2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 05Aug2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 11Aug2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//83
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
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//41
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//42
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//44
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//45
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//46
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//48
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//51
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//52
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//53
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//54
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//56
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//57
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//58
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//59
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//60
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//61
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//62
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//63
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//64
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//65
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//66
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//67
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//68
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//69
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//70
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//71
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//72
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//73
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//74
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//75
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//76
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//77
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//78
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//79
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//80
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//81 13Mar2018
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//82 13Mar2018
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//02Nov2018
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 16Apr2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 23Jun2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 29Jul2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 04Aug2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 05Aug2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 11Aug2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 09Nov2020
            ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//83
            ExcelBuf.NewRow;
            //<<

        END;

        //<<28Apr2017

    end;

    //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //>>22Feb2017 RB-N
        /* ExcelBuf.CreateBook('', Text003);
        //ExcelBuf.CreateBookAndOpenExcel('',Text000,Text001,COMPANYNAME,USERID);
        ExcelBuf.CreateBookAndOpenExcel('', Text003, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text003);
        ExcelBuf.WriteSheet(Text003, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text003, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<
    end;

    // [Scope('Internal')]
    procedure MakeExcelInvTotalGroupFooter()
    begin
        //..
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Line"."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//1 31July2017
        ExcelBuf.AddColumn("Sales Invoice Line"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//2 31July2017
        ExcelBuf.AddColumn("Sales Invoice Header"."Cancelled Invoice", FALSE, '', TRUE, FALSE, FALSE, '', 1);//4 31July2017
        ExcelBuf.AddColumn(SalesType, FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Line"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//5 31July2017
        ExcelBuf.AddColumn(StateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6 31July2017
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7 31July2017
        ExcelBuf.AddColumn(GSTStateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8 31July2017
        ExcelBuf.AddColumn(LocationRec."GST Registration No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//9 31July2017
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(ResCenName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Invoice Header"."Customer Posting Group", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        //ExcelBuf.AddColumn(Cust."Full Name",FALSE,'',TRUE,FALSE,FALSE,'',1);//
        ExcelBuf.AddColumn(Cust.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        //ExcelBuf.AddColumn(Cust."GST Registration No.",FALSE,'',TRUE,FALSE,FALSE,'',1);//15 31July2017
        //>>24May2019
        CLEAR(DGST04);
        DGST04.SETCURRENTKEY("Document No.");
        DGST04.SETRANGE("Document No.", "Sales Invoice Line"."Document No.");
        IF DGST04.FINDFIRST THEN;
        ExcelBuf.AddColumn(DGST04."Buyer/Seller Reg. No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
        //<<24May2019
        ExcelBuf.AddColumn("Sales Invoice Header"."Ship-to Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn(Cust."GST Registration No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//17 31July2017
        ExcelBuf.AddColumn(GSTPlaceofSupply, FALSE, '', TRUE, FALSE, FALSE, '', 1);//18 31July2017
        ExcelBuf.AddColumn(GSTSupplyCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//19 31July2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22 31July2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23 31July2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0); //26
        ExcelBuf.AddColumn(TQtyBase, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27
        ExcelBuf.AddColumn("Sales Invoice Header"."Currency Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//MRP //29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//List //30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPLSUM

        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//FCPrice //31
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//FCAmt //32
                                                                         //ExcelBuf.AddColumn(UPrice,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//INR UNIT //
            ExcelBuf.AddColumn(TINRAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//INR Amt //33
        END
        ELSE BEGIN
            ExcelBuf.AddColumn((''), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//FCPrice //31
                                                                                     //ExcelBuf.AddColumn(UPrice*"Sales Invoice Line".Quantity,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//FCAMt //32
            ExcelBuf.AddColumn(TINRUnit, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//FCAMt //32 //30Dec2017
            TINRUnit := 0;//30Dec2017
            ExcelBuf.AddColumn(TINRAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//INR Amt //33
        END;

        ExcelBuf.AddColumn(-(TAVPDiscAmt), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//34 //01Aug2017
        ExcelBuf.AddColumn(-(TRegDiscAmt), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//02Nov2018
        ExcelBuf.AddColumn(-(TTransportSubsidy), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//35 31July2017
        ExcelBuf.AddColumn(-(TCashDiscount), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//36 31July2017
        ExcelBuf.AddColumn(-(TLineDisc + TDiscAmt), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//37 01Aug2017 InvoiceDiscount
        ExcelBuf.AddColumn(-(ABS(TTransportSubsidy) + ABS(TCashDiscount) + ABS(TLineDisc) + ABS(TAVPDiscAmt)
        + ABS(TDiscAmt) + ABS(TRegDiscAmt)), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//38 31July2017
        //>>23Feb2017
        TDiscAmt := 0;//01Aug2017
        TAVPDiscAmt := 0;//01Aug2017
        TRegDiscAmt := 0;//02Nov2018

        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//39
        END ELSE BEGIN
            ExcelBuf.AddColumn(InFCFG, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); // 39

        END;

        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            IF FRIEGHTAmount <> 0 THEN BEGIN
                ExcelBuf.AddColumn(/*"Sales Invoice Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt
                                   - DISCOUNTAmount - SurChargeInterest, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//40
                TFRIEGHTAmount += 0;//"Sales Invoice Line"."Charges To Customer";//23Feb2017
                                    //>>27Feb
                                    //InFG += "Sales Invoice Line"."Charges To Customer" - ENTRYTax - CessAmount - AddTaxAmt- DISCOUNTAmount-SurChargeInterest;
                                    //TInFG += "Sales Invoice Line"."Charges To Customer" - ENTRYTax - CessAmount - AddTaxAmt- DISCOUNTAmount-SurChargeInterest;
                                    //<<
            END ELSE BEGIN
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//40
            END;
        END ELSE BEGIN
            ExcelBuf.AddColumn(InFG, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//40
                                                                                  //TFRIEGHTAmount += "Sales Invoice Line"."Charges To Customer"/"Sales Invoice Header"."Currency Factor";//23Feb2017
                                                                                  //>>27Feb
                                                                                  //<<
        END;

        ExcelBuf.AddColumn(0.0, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//41 31July2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//42 31July2017


        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(TTaxBAse, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43 31July2017

        END ELSE BEGIN
            ExcelBuf.AddColumn(TTaxBAse, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43
        END;


        //ExcelBuf.AddColumn("Sales Invoice Line"."GST %",FALSE,'',TRUE,FALSE,FALSE,'#,#0.00',0);//44 31July2017
        ExcelBuf.AddColumn(GSTPer, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//44 RB-N 20Sep2017
        ExcelBuf.AddColumn(TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//45 31July2017
        ExcelBuf.AddColumn(TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//46 31July2017
        ExcelBuf.AddColumn(TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//47 31July2017
        ExcelBuf.AddColumn(TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//48 31July2017
        ExcelBuf.AddColumn(TCessAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//49 13Mar2018
        ExcelBuf.AddColumn(TTCSAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//50 13Mar2018
        TTCSAmt := 0;
        TCessAmt := 0;

        ExcelBuf.AddColumn(RoundOffAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//49 31July2017 //RoundoffAmount
        InvRound += RoundOffAmt; //01Aug2017

        IF "Sales Invoice Header"."Currency Code" = '' THEN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1) //50
        ELSE
            ExcelBuf.AddColumn(FCGROSSINV, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0); //50
        //>>23Feb2017



        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(TGrossInVAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //51

        END
        ELSE BEGIN
            ExcelBuf.AddColumn(TGrossInVAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//51
        END;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//52
        ExcelBuf.AddColumn(TRRNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//53
        ExcelBuf.AddColumn(BRTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//54
        ExcelBuf.AddColumn(SalesName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LineSalespersonName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn("Sales Invoice Header"."Order No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//57
        ExcelBuf.AddColumn("Sales Invoice Header"."Document Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//58
        ExcelBuf.AddColumn("Sales Invoice Header"."Vehicle No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//59
        ExcelBuf.AddColumn("Sales Invoice Header"."Road Permit No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//60 //E-way BillNo
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//61 //E-way BillDate
        ExcelBuf.AddColumn("Sales Invoice Header"."LR/RR No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//62
        ExcelBuf.AddColumn("Sales Invoice Header"."LR/RR Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//63
        ExcelBuf.AddColumn("Sales Invoice Header"."External Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//64
        ExcelBuf.AddColumn(ShipmentAgentName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn(CustomerType, FALSE, '', TRUE, FALSE, FALSE, '', 1);//66
        ExcelBuf.AddColumn("Sales Invoice Header"."Payment Terms Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(SalesGroup, FALSE, '', TRUE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn("Sales Invoice Header"."Driver's Mobile No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//69
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//70 //Supp. Org Invoice
        ExcelBuf.AddColumn(REGION, FALSE, '', TRUE, FALSE, FALSE, '', 1);//71
        ExcelBuf.AddColumn(SPRegionCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//19Jun2019
        ExcelBuf.AddColumn(SPZoneCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn(FORMAT("Sales Invoice Header"."Freight Type"), FALSE, '', TRUE, FALSE, FALSE, '', 1);//72
        ExcelBuf.AddColumn(CSOAppDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//73
        ExcelBuf.AddColumn(ProductGroupName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//74
        ExcelBuf.AddColumn(tComment, FALSE, '', TRUE, FALSE, FALSE, '', 1);//75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//76
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//77
        ExcelBuf.AddColumn(OrgState, FALSE, '', TRUE, FALSE, FALSE, '', 1);//78
        ExcelBuf.AddColumn(FinalState, FALSE, '', TRUE, FALSE, FALSE, '', 1);//79
        ExcelBuf.AddColumn("Sales Invoice Header"."Nature of Supply", FALSE, '', TRUE, FALSE, FALSE, '', 1);//80
        ExcelBuf.AddColumn(BDNNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//81 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(BDNDate, FALSE, '', TRUE, FALSE, FALSE, '', 1);//82 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(VesselName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//83 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(COPYSTR(txtGRN, 2, STRLEN(txtGRN)), FALSE, '', TRUE, FALSE, FALSE, '', 1);//84 RSPLSUM 29Jul2020
        ExcelBuf.AddColumn(COPYSTR(txtGRNNo, 2, STRLEN(txtGRNNo)), FALSE, '', TRUE, FALSE, FALSE, '', 1);//85 RSPLSUM 04Aug2020
        ExcelBuf.AddColumn(COPYSTR(txtSalRetRcpt, 2, STRLEN(txtSalRetRcpt)), FALSE, '', TRUE, FALSE, FALSE, '', 1);//86 RSPLSUM 05Aug2020
        ExcelBuf.AddColumn("Sales Invoice Header"."Mintifi Channel Finance", FALSE, '', TRUE, FALSE, FALSE, '', 1);//87 RSPLSUM 11Aug2020

        //AR

        ExcelBuf.AddColumn(IRNGeneratedTxt, FALSE, '', TRUE, FALSE, FALSE, '', 1);//85
        ExcelBuf.AddColumn(/*"Sales Invoice Header"."EWB No."*/'', FALSE, '', TRUE, FALSE, FALSE, '', 1);//86
        ExcelBuf.AddColumn("Sales Invoice Header"."EWB Date", FALSE, '', TRUE, FALSE, FALSE, '', 1);//87
        //MESSAGE('no. ' +"Sales Invoice Header"."No." +'    irn ' +"Sales Invoice Header".IRN +'Ewb no '+ "Sales Invoice Header"."EWB No." +'ewb dte ' + FORMAT("Sales Invoice Header"."EWB Date"));
        //AR

        "Sales Invoice Header".CALCFIELDS("Sales Invoice Header".IRN);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn("Sales Invoice Header".IRN, FALSE, '', TRUE, FALSE, FALSE, '', 1);//88 RSPLSUM 09Nov2020
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."TDS/TCS Amount"*/0, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//89 RSPLSUM 09Nov2020
    end;

    // [Scope('Internal')]
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//42
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//51
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70 //RSPL008
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//71
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//72
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//73
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//74
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//76
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//77
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//78
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//81
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 //13Mar2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 //13Mar2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//02Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19Jun2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 29Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 04Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 05Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 11Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23Feb2017
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
        ExcelBuf.AddColumn(InTQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//26
        InTQty := 0;//01Aug2017

        ExcelBuf.AddColumn(InTQtyBase, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//27
        InTQtyBase := 0;//01Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn(InTINRAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//33
        InTINRAmt := 0;//01Aug2017

        ExcelBuf.AddColumn(-(TTAVPDiscAmt), FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//34
        ExcelBuf.AddColumn(-(TTRegDiscAmt), FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);// 02Nov2018


        ExcelBuf.AddColumn(-(InTTransportSubsidy), FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//35
        ExcelBuf.AddColumn(-(InTCashDiscount), FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//36
        ExcelBuf.AddColumn(-(InLineDisc + TTDiscAmt), FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//37
        ExcelBuf.AddColumn(-(ABS(InTTransportSubsidy) + ABS(InTCashDiscount) + ABS(TTAVPDiscAmt) + ABS(TTRegDiscAmt)
        + ABS(TTDiscAmt) + ABS(InLineDisc)), FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//38
        InTTransportSubsidy := 0;//01Aug2017
        InTCashDiscount := 0;//01Aug2017
        InLineDisc := 0;//01Aug2017
        TTAVPDiscAmt := 0;//01Aug2017
        TTDiscAmt := 0;//01Aug2017
        TTRegDiscAmt := 0;//02Nov2018

        ExcelBuf.AddColumn(TInFCFG, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//39
        ExcelBuf.AddColumn(TInFG, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//40
        TInFG := 0;//02Aug2017
        TInFCFG := 0;//02Aug2017

        ExcelBuf.AddColumn(0.0, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//41
        ExcelBuf.AddColumn(0.0, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//42
        ExcelBuf.AddColumn(InTTaxBAse, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//43
        InTTaxBAse := 0;//01Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn(TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//45
        ExcelBuf.AddColumn(TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//46
        ExcelBuf.AddColumn(TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);           //47
        ExcelBuf.AddColumn(TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//48

        TTCGST := 0;//01Aug2017
        TTSGST := 0;//01Aug2017
        TTIGST := 0;//01Aug2017
        TTUTGST := 0;//01Aug2017

        ExcelBuf.AddColumn(TTCessAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//49 //13Mar2018
        ExcelBuf.AddColumn(TTTCSAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//50 //13Mar2018
        TTCessAmt := 0;
        TTTCSAmt := 0;

        ExcelBuf.AddColumn(InvRound, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//49
        InvRound := 0;//01Aug2017

        ExcelBuf.AddColumn(InTFCGROSSINV, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//50
        InTFCGROSSINV := 0;//01Aug2017

        ExcelBuf.AddColumn(InTGrossInVAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//51
        InTGrossInVAmt := 0;//01Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//  56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);// //57
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//78
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19Jun2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 29Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 04Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 05Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 11Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);///RSPLSUM 09Nov2020
    end;

    // [Scope('Internal')]
    procedure MakeCrMemoBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1 31July2017
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2 31July2017
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3 31July2017
        ExcelBuf.AddColumn(SalesType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5 31July2017
        ExcelBuf.AddColumn(StateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6 31July2017
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7 31July2017
        ExcelBuf.AddColumn(GSTStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8 31July2017
        ExcelBuf.AddColumn(LocationRec."GST Registration No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//9 31July2017
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(ResCenName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        //ExcelBuf.AddColumn(Cust."Full Name",FALSE,'',FALSE,FALSE,FALSE,'',1);//
        ExcelBuf.AddColumn(Cust.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        //ExcelBuf.AddColumn(Cust."GST Registration No.",FALSE,'',FALSE,FALSE,FALSE,'',1);//15 31July2017
        //>>24May2019
        CLEAR(DGST04);
        DGST04.SETCURRENTKEY("Document No.");
        DGST04.SETRANGE("Document No.", "Sales Cr.Memo Line"."Document No.");
        IF DGST04.FINDFIRST THEN;
        ExcelBuf.AddColumn(DGST04."Buyer/Seller Reg. No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        //<<24May2019
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Ship-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn(Cust."GST Registration No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17 31July2017
        ExcelBuf.AddColumn(GSTPlaceofSupply, FALSE, '', FALSE, FALSE, FALSE, '', 1);//18 31July2017
        ExcelBuf.AddColumn(GSTSupplyCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//19 31July2017

        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
            "recItem-Supp".GET("Sales Cr.Memo Line"."Final Item No.");
            ExcelBuf.AddColumn("Sales Cr.Memo Line"."Final Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
            ExcelBuf.AddColumn("recItem-Supp".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        END ELSE BEGIN
            ExcelBuf.AddColumn("Sales Cr.Memo Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
            ExcelBuf.AddColumn("Sales Cr.Memo Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        END;

        ExcelBuf.AddColumn("Sales Cr.Memo Line"."HSN/SAC Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//22 31July2017
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//23 31July2017

        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//24
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//25
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0); //26


        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27


        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//28
        //>>02Aug2017
        IF "Sales Cr.Memo Header"."Currency Code" <> '' THEN
            ExcelBuf.AddColumn(1 / "Sales Cr.Memo Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//29
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//29
        //>>02Aug2017
        ExcelBuf.AddColumn(-MRPPricePerLt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//MRP //29
        //ExcelBuf.AddColumn(-ListPricePerlt+ProdDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//List //30//RSPLSUM
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Basic Price", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//RSPLSUM
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Freight/Other Chgs. Primary", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//RSPLSUM
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Freight/Other Chgs. Secondary", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//RSPLSUM
        ExcelBuf.AddColumn(-ListPricePerlt + crProdDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//List //30

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//FCPrice //31
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//FCAmt //32
                                                                          //ExcelBuf.AddColumn(UPrice,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//INR UNIT //
                                                                          //INRTot1 := UPrice;
                                                                          //RSPLSUM 10Nov2020--ExcelBuf.AddColumn(- "Sales Cr.Memo Line".Amount,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//INR Amt //33
                                                                          //RSPLSUM 04Dec2020--ExcelBuf.AddColumn(-("Sales Cr.Memo Line".Amount+"Sales Cr.Memo Line"."Line Discount Amount"+"Sales Cr.Memo Line"."Charges To Customer"),FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//INR Amt //33 RSPLSUM 10Nov2020
            ExcelBuf.AddColumn(-("Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount"), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//INR Amt //33 RSPLSUM 04Dec2020
        END ELSE BEGIN
            ExcelBuf.AddColumn(-(UPricePerLT), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//FCPrice //31
            INRTot += UPricePerLT;//23Feb2017
                                  //InINRTot += UPricePerLT;//27Feb2017
            ExcelBuf.AddColumn(-(UPrice * "Sales Cr.Memo Line".Quantity), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//FCAMt //32
            FcTotal += UPrice * "Sales Cr.Memo Line".Quantity;//30Dec2017
            GFcTotal += UPrice * "Sales Cr.Memo Line".Quantity;//30Dec2017
            ExcelBuf.AddColumn(-(INRamt / "Sales Cr.Memo Header"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//INR Amt //33
        END;

        ExcelBuf.AddColumn(AVPDiscAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//34
        ExcelBuf.AddColumn(RegDiscAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//02Nov2018
        ExcelBuf.AddColumn(TransportSubsidy, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//35 31July2017

        ExcelBuf.AddColumn(CashDiscount, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//36 31July2017
        //>>27Feb2017
        CrCash += CashDiscount;
        InCrCash += CashDiscount;

        //ExcelBuf.AddColumn(("Sales Cr.Memo Line"."Line Discount Amount"+DiscAmt),FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//37 01Aug2017
        ExcelBuf.AddColumn((LineDiscAmt + DiscAmt), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//08Apr2019
        //ExcelBuf.AddColumn((TransportSubsidy+CashDiscount+AVPDiscAmt+DiscAmt+RegDiscAmt+"Sales Cr.Memo Line"."Line Discount Amount"),FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//38 31July2017
        ExcelBuf.AddColumn((TransportSubsidy + CashDiscount + AVPDiscAmt + DiscAmt + RegDiscAmt + LineDiscAmt), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//08Apr2019

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//39
        END ELSE BEGIN
            //ExcelBuf.AddColumn(-FCFRIEGHTAmount,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0); // 39
            ExcelBuf.AddColumn(-CrFCFRIEGHTAmount, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //39 30Dec2017
        END;

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            IF FRIEGHTAmount <> 0 THEN BEGIN
                ExcelBuf.AddColumn(-(/*"Sales Cr.Memo Line"."Charges To Customer" */0 - ENTRYTax - CessAmount - AddTaxAmt
                                   - DISCOUNTAmount - SurChargeInterest), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//40
            END ELSE BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//40
            END;
        END ELSE BEGIN
            ExcelBuf.AddColumn(-(/*"Sales Cr.Memo Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt)
            / "Sales Cr.Memo Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//40

        END;

        ExcelBuf.AddColumn(0.0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//41 31July2017

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//42 31July2017


        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(-/*"Sales Cr.Memo Line"."GST Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43 31July2017
        END ELSE BEGIN
            ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Line Amount" / "Sales Cr.Memo Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//42
                                                                                                                                                             //ExcelBuf.AddColumn("Sales Cr.Memo Line"."GST Base Amount"/"Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//43
        END;
        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."GST %",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//44 31July2017
        ExcelBuf.AddColumn(GSTPer, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//44 RB-N 20Sep2017
        ExcelBuf.AddColumn(-vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//45 31July2017
        ExcelBuf.AddColumn(-vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//46 31July2017
        ExcelBuf.AddColumn(-vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//47 31July2017
        ExcelBuf.AddColumn(-vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//48 31July2017
        ExcelBuf.AddColumn(-CessAmt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//49//13Mar2018
        ExcelBuf.AddColumn(-TCSAmt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//50//13Mar2018

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//49 31July2017 //RoundoffAmount

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1) //50
        ELSE
            ExcelBuf.AddColumn(-/*"Sales Cr.Memo Line"."Amount To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //50

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(-/*"Sales Cr.Memo Line"."Amount To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); //51

        END ELSE BEGIN
            ExcelBuf.AddColumn(-/*"Sales Cr.Memo Line"."Amount To Customer"*/0 / "Sales Cr.Memo Header"."Currency Factor"
            , FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//51

        END;

        //ExcelBuf.AddColumn(ProdDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//52//RSPLSUM
        ExcelBuf.AddColumn(crProdDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//52
        ExcelBuf.AddColumn(TRRNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//53
        ExcelBuf.AddColumn(BRTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//54
        ExcelBuf.AddColumn(SalesName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LineSalespersonName1, FALSE, '', FALSE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Return Order No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//57
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//58
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//59
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//60 //E-way BillNo
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//61 //E-way BillDate
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//62
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 2);//63
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."External Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//64
        ExcelBuf.AddColumn(ShipmentAgentName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn(CustomerType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//66
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(SalesGroup1, FALSE, '', FALSE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//69
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//70 //Supp. Org Invoice
        ExcelBuf.AddColumn(REGION, FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
        ExcelBuf.AddColumn(SPRegionCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//19Jun2019
        ExcelBuf.AddColumn(SPZoneCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn(FORMAT("Sales Cr.Memo Header"."Freight Type"), FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        ExcelBuf.AddColumn(CSOAppDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//73
        ExcelBuf.AddColumn(ProductGroupName2, FALSE, '', FALSE, FALSE, FALSE, '', 1);//74
        ExcelBuf.AddColumn(tComment, FALSE, '', FALSE, FALSE, FALSE, '', 1);//75
        //>>RB-N 25Oct2017 DivisionName
        DivisionText := '';
        Dim25.RESET;
        IF Dim25.GET('DIVISION', "Sales Cr.Memo Line"."Shortcut Dimension 1 Code") THEN BEGIN
            DivisionText := Dim25.Name;
        END;
        //<<RB-N 25Oct2017 DivisionName

        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//76
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code" + ' - ' + DivisionText, FALSE, '', FALSE, FALSE, FALSE, '', 1);//76 25Oct2017 DivisionName

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//77
        ExcelBuf.AddColumn(OrgState, FALSE, '', FALSE, FALSE, FALSE, '', 1);//78
        ExcelBuf.AddColumn(FinalState, FALSE, '', FALSE, FALSE, FALSE, '', 1);//79
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Nature of Supply", FALSE, '', FALSE, FALSE, FALSE, '', 1);//80
        ExcelBuf.AddColumn(BDNNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//81 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(BDNDate, FALSE, '', FALSE, FALSE, FALSE, '', 1);//82 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(VesselName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//83 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(IRNGeneratedTxtCrMemo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//RSPL/AviMali
        ExcelBuf.AddColumn("Sales Cr.Memo Header".IRN, FALSE, '', FALSE, FALSE, FALSE, '', 1);//RSPL/AviMali
    end;

    // [Scope('Internal')]
    procedure MakecrmemoHeader()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Sales Register Cr - GST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//42
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//51
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//78
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//81
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//02Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19Jun2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLAviMali
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPL/AviMali

        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//81
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//82
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//83
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//84
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//85
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//86
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//87//004
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//88  //004
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//89 //RSPL-230115
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//90 //RSPL-DHANANJAY
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//91 //RSPL-DHANANJAY
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//92 //RSPL-DHANANJAY
        *///31July2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Cancelled Invoice', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Movement Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        /*
        ExcelBuf.AddColumn('Origin State OF Our',FALSE,'',TRUE,FALSE,TRUE,'@',1);//6 31July2017
        ExcelBuf.AddColumn('GPPL Supply Location Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//7 31Juy2017
        */
        ExcelBuf.AddColumn('Origin State', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6 01Aug2017
        ExcelBuf.AddColumn('GPPL Billing Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7 01Aug2017
        ExcelBuf.AddColumn('GPPL State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8 31Juy2017
        ExcelBuf.AddColumn('GPPL GSTIN No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9 31Juy2017
        ExcelBuf.AddColumn('Resp. Centre Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Resp. Centre Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Customer Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Customer Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Details of Customer ( Billed To )', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 31July2017
        ExcelBuf.AddColumn('Bill to Location GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15 31July2017
        ExcelBuf.AddColumn('Details of Consignee ( Shipped To)', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16 31July2017
        ExcelBuf.AddColumn('Ship to Location GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17 31July2017
        ExcelBuf.AddColumn('Place of Supply', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18 31July2017
        ExcelBuf.AddColumn('Supply State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19 31July2017
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('Product HSN Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 31July2017
        ExcelBuf.AddColumn('Batch No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 31July2017
        ExcelBuf.AddColumn('Unit of Measure', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Packing Qty per UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('UOM Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        ExcelBuf.AddColumn('Total Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('Exchange Rate', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 02Aug2017
        ExcelBuf.AddColumn('MRP Per Lt/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('Basic Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM
        ExcelBuf.AddColumn('Freight Primary', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM
        ExcelBuf.AddColumn('Freight Secondary', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM
        ExcelBuf.AddColumn('List Price per Lt/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
                                                                                         /*
                                                                                         ExcelBuf.AddColumn('Ass. Value',FALSE,'',TRUE,FALSE,TRUE,'@',1);//  //EBT MILAN
                                                                                         ExcelBuf.AddColumn('INR Price Per KG/Ltr',FALSE,'',TRUE,FALSE,TRUE,'@',1);//
                                                                                         *///31July2017

        ExcelBuf.AddColumn('FC Price Per KG/Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33
        //ExcelBuf.AddColumn('AVP DISCOUNT',FALSE,'',TRUE,FALSE,TRUE,'@',1);//34
        ExcelBuf.AddColumn('National Discount Scheme', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//02Nov2018
        ExcelBuf.AddColumn('Regional Discount Scheme', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//02Nov2018
        ExcelBuf.AddColumn('Transport Subsidy', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35
        ExcelBuf.AddColumn('Cash Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//36
        ExcelBuf.AddColumn('Invoice Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//37 01Aug2017
        ExcelBuf.AddColumn('Total Discounts', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//38
        ExcelBuf.AddColumn('FC FREIGHT', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//39
        ExcelBuf.AddColumn('FREIGHT', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//40
        ExcelBuf.AddColumn('Insurance', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//41
        ExcelBuf.AddColumn('Other Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//42
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//43
        ExcelBuf.AddColumn('Tax %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//44
        ExcelBuf.AddColumn('CGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//45 31July2017
        ExcelBuf.AddColumn('SGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//46 31July2017
        ExcelBuf.AddColumn('IGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//47 31July2017
        ExcelBuf.AddColumn('UTGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//48 31July2017
        ExcelBuf.AddColumn('GST Comp. Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//48 31July2017
        ExcelBuf.AddColumn('TCS Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//49 13Mar2018
        ExcelBuf.AddColumn('R/Off', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//50 13Mar2018
        ExcelBuf.AddColumn('FC Gross Invoice Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//51
        ExcelBuf.AddColumn('Gross Invoice Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//51
        ExcelBuf.AddColumn('Product Discount Per Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//52
        ExcelBuf.AddColumn('TRR Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//53
        ExcelBuf.AddColumn('BRT Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//54
        ExcelBuf.AddColumn('Header Salesperson Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//55
        ExcelBuf.AddColumn('Line Salesperson Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//56
        ExcelBuf.AddColumn('CSO Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//57
        ExcelBuf.AddColumn('CSO Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//58
        ExcelBuf.AddColumn('Vehicle Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//59
        ExcelBuf.AddColumn('E-Way Bill No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//60
        ExcelBuf.AddColumn('E-Way Bill Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//61
        ExcelBuf.AddColumn('LR Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//62
        ExcelBuf.AddColumn('LR Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//63
        ExcelBuf.AddColumn('PO No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//64
        ExcelBuf.AddColumn('Transporter Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//65
        ExcelBuf.AddColumn('Customer Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//66
        ExcelBuf.AddColumn('Payment Terms Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//67
        ExcelBuf.AddColumn('Sales Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//68
        ExcelBuf.AddColumn('Drivers Mobile No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//69
        ExcelBuf.AddColumn('Supp. Org Invoice', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//70
        ExcelBuf.AddColumn('Customer Region', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//71
        ExcelBuf.AddColumn('Sales Person Region', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19Jun2019
        ExcelBuf.AddColumn('Sales Person Zone', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('Freight Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//72
        ExcelBuf.AddColumn('CSO Approved Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//73
        ExcelBuf.AddColumn('Product Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//74
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//75
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//76
        ExcelBuf.AddColumn('Customer Receipt Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//77
        ExcelBuf.AddColumn('Origin State', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//78 31July2017
        ExcelBuf.AddColumn('Destination State', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//79 31July2017
        ExcelBuf.AddColumn('Nature of Transaction', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//80 31July2017
        ExcelBuf.AddColumn('BDN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//81 23Jun2020
        ExcelBuf.AddColumn('BDN Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//82 23Jun2020
        ExcelBuf.AddColumn('Vessel Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//83 23Jun2020
        ExcelBuf.AddColumn('IRN Generated', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//82 23Jun2020
        ExcelBuf.AddColumn('IRN', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//83 23Jun2020

    end;

    //[Scope('Internal')]
    procedure MakeExcelCrMemoGrouping()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//1 31July2017
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//2 31July2017
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Applies-to Doc. No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//3 31July2017
        ExcelBuf.AddColumn(SalesType, FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//5 31July2017
        ExcelBuf.AddColumn(StateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6 31July2017
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7 31July2017
        ExcelBuf.AddColumn(GSTStateCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8 31July2017
        ExcelBuf.AddColumn(LocationRec."GST Registration No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//9 31July2017
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(ResCenName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Customer Posting Group", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        //ExcelBuf.AddColumn(Cust."Full Name",FALSE,'',TRUE,FALSE,FALSE,'',1);//
        ExcelBuf.AddColumn(Cust.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn(Cust."GST Registration No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//15 31July2017
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Ship-to Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        //ExcelBuf.AddColumn(Cust."GST Registration No.",FALSE,'',TRUE,FALSE,FALSE,'',1);//17 31July2017
        //>>24May2019
        CLEAR(DGST04);
        DGST04.SETCURRENTKEY("Document No.");
        DGST04.SETRANGE("Document No.", "Sales Cr.Memo Line"."Document No.");
        IF DGST04.FINDFIRST THEN;
        ExcelBuf.AddColumn(DGST04."Buyer/Seller Reg. No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
        //<<24May2019
        ExcelBuf.AddColumn(GSTPlaceofSupply, FALSE, '', TRUE, FALSE, FALSE, '', 1);//18 31July2017
        ExcelBuf.AddColumn(GSTSupplyCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//19 31July2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22 31July2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23 31July2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
        ExcelBuf.AddColumn(-CrTQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0); //26
        ExcelBuf.AddColumn(-CrTQtyBase, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Currency Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//MRP //29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//List //30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//FCPrice //31
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//FCAmt //32
                                                                         //ExcelBuf.AddColumn(UPrice,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//INR UNIT //
                                                                         //INRTot1 := UPrice;
                                                                         //TINRUnit += UPrice;//23Feb2017
                                                                         //InTINRUnit += UPrice;//27Feb2017
            ExcelBuf.AddColumn(-TINRAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//INR Amt //33

        END ELSE BEGIN

            ExcelBuf.AddColumn(-(INRTot), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//FCPrice //31
            ExcelBuf.AddColumn(-GFcTotal, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//FCAMt //32
            ExcelBuf.AddColumn(-TINRAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//INR Amt //33
        END;

        INRTot := 0;//30Dec2017
        TINRAmt := 0;//01Aug2017
        GFcTotal := 0;//30Dec2017

        ExcelBuf.AddColumn((TAVPDiscAmt), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//34
        ExcelBuf.AddColumn((TRegDiscAmt), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//02Nov2018
        //>>23Feb2017
        //CrLineDisc += "Sales Cr.Memo Line"."Line Discount Amount";
        //InCrLineDisc += "Sales Cr.Memo Line"."Line Discount Amount";
        //<<

        ExcelBuf.AddColumn((CrTrans), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//35 31July2017
        //>>23Feb2017
        //CrTrans += TransportSubsidy;
        //InCrTrans += TransportSubsidy;
        //<<

        ExcelBuf.AddColumn((CrCash), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//36 31July2017
                                                                                     //>>27Feb2017
                                                                                     //CrCash += CashDiscount;
                                                                                     //InCrCash += CashDiscount;
                                                                                     //<<
                                                                                     //<<

        ExcelBuf.AddColumn((CrLineDisc + TDiscAmt), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//37 31July2017
        ExcelBuf.AddColumn((CrCash + CrTrans + TAVPDiscAmt + CrLineDisc + TDiscAmt + TRegDiscAmt), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//38 31July2017
        //>>23Feb2017
        TAVPDiscAmt := 0;//02Aug2017
        TDiscAmt := 0; //02Aug2017
        TRegDiscAmt := 0;//02Nov2018

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//38
        END ELSE BEGIN
            ExcelBuf.AddColumn(-FCFRIEGHTAmount, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); // 38
                                                                                                  //>>27Feb
                                                                                                  //InFCFG += FCFRIEGHTAmount;
                                                                                                  //TInFCFG += FCFRIEGHTAmount;
                                                                                                  //<<
        END;

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            IF FRIEGHTAmount <> 0 THEN BEGIN
                ExcelBuf.AddColumn(-TFRIEGHTAmount, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//39
                                                                                                    //TFRIEGHTAmount += "Sales Cr.Memo Line"."Charges To Customer";//23Feb2017
                                                                                                    //>>27Feb
                                                                                                    // InFG += "Sales Cr.Memo Line"."Charges To Customer" - ENTRYTax - CessAmount - AddTaxAmt- DISCOUNTAmount-SurChargeInterest;
                                                                                                    //TInFG += "Sales Cr.Memo Line"."Charges To Customer" - ENTRYTax - CessAmount - AddTaxAmt- DISCOUNTAmount-SurChargeInterest;
                                                                                                    //<<
            END ELSE BEGIN
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//39
            END;
        END ELSE BEGIN
            ExcelBuf.AddColumn(-TFRIEGHTAmount, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//39
                                                                                                //TFRIEGHTAmount += "Sales Cr.Memo Line"."Charges To Customer"/"Sales Cr.Memo Header"."Currency Factor";//23Feb2017
                                                                                                //>>27Feb
            InFG += (/*"Sales Cr.Memo Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt) / "Sales Cr.Memo Header"."Currency Factor";
            TInFG += (/*"Sales Cr.Memo Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt) / "Sales Cr.Memo Header"."Currency Factor";
            //<<
        END;

        ExcelBuf.AddColumn(0.0, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//40 31July2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//41 31July2017


        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(-CrTaxBase, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//42 31July2017
                                                                                        //>>23Feb2017
                                                                                        //CrTaxBase += "Sales Cr.Memo Line"."GST Base Amount";
                                                                                        //InCrTaxBase += "Sales Cr.Memo Line"."GST Base Amount";//27Feb
                                                                                        //<<
        END ELSE BEGIN
            ExcelBuf.AddColumn(-CrTaxBase, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//42
                                                                                        //>>23Feb2017
                                                                                        //CrTaxBase += "Sales Cr.Memo Line"."GST Base Amount" /"Sales Cr.Memo Header"."Currency Factor";
                                                                                        //InCrTaxBase += "Sales Cr.Memo Line"."GST Base Amount" /"Sales Cr.Memo Header"."Currency Factor";//27Feb
        END;


        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."GST %",FALSE,'',TRUE,FALSE,FALSE,'#,#0.00',0);//43 31July2017
        ExcelBuf.AddColumn(GSTPer, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//43 RB-N 20Sep2017
        ExcelBuf.AddColumn(-TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//44 31July2017
        ExcelBuf.AddColumn(-TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//45 31July2017
        ExcelBuf.AddColumn(-TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//46 31July2017
        ExcelBuf.AddColumn(-TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//47 31July2017
        CLEAR(TSGST);//31July2017
        CLEAR(TCGST);//31July2017
        CLEAR(TIGST);//31July2017
        CLEAR(TUTGST);//31July2017

        ExcelBuf.AddColumn(-TCessAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//49 13Mar2018
        ExcelBuf.AddColumn(-TTCSAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//50 13Mar2018
        TTCSAmt := 0;
        TCessAmt := 0;

        ExcelBuf.AddColumn(RoundOffAmt1, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//49 31July2017 //RoundoffAmount
        //>>01Aug2017 RoundoffAmount Credit
        CrRound += RoundOffAmt1;

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1) //50
        ELSE
            ExcelBuf.AddColumn(-CrFCAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0); //50




        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(-CrAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //50
                                                                                        //>>27Feb2017
                                                                                        //CrAmt += "Sales Cr.Memo Line"."Amount To Customer";
                                                                                        //InCrAmt += "Sales Cr.Memo Line"."Amount To Customer";
                                                                                        //<<
        END ELSE BEGIN
            ExcelBuf.AddColumn(-CrAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//50
                                                                                       //>>27Feb2017
                                                                                       //CrAmt += "Sales Cr.Memo Line"."Amount To Customer" / "Sales Cr.Memo Header"."Currency Factor";
                                                                                       //InCrAmt += "Sales Cr.Memo Line"."Amount To Customer" / "Sales Cr.Memo Header"."Currency Factor";
                                                                                       //<<
        END;

        TINRAmt := 0; //30Dec2017
        TINRUnit := 0; //30Dec2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//51
        ExcelBuf.AddColumn(TRRNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//52
        ExcelBuf.AddColumn(BRTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//53
        ExcelBuf.AddColumn(SalesName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//54
        ExcelBuf.AddColumn(LineSalespersonName1, FALSE, '', TRUE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Return Order No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Document Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//59 //E-way BillNo
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//60 //E-way BillDate
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//61
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 2);//62
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//63
        ExcelBuf.AddColumn(ShipmentAgentName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//64
        ExcelBuf.AddColumn(CustomerType, FALSE, '', TRUE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Payment Terms Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//66
        ExcelBuf.AddColumn(SalesGroup1, FALSE, '', TRUE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//69 //Supp. Org Invoice
        ExcelBuf.AddColumn(REGION, FALSE, '', TRUE, FALSE, FALSE, '', 1);//70
        ExcelBuf.AddColumn(SPRegionCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//19Jun2019
        ExcelBuf.AddColumn(SPZoneCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn(FORMAT("Sales Cr.Memo Header"."Freight Type"), FALSE, '', TRUE, FALSE, FALSE, '', 1);//71
        ExcelBuf.AddColumn(CSOAppDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//72
        ExcelBuf.AddColumn(ProductGroupName2, FALSE, '', TRUE, FALSE, FALSE, '', 1);//73
        ExcelBuf.AddColumn(tComment, FALSE, '', TRUE, FALSE, FALSE, '', 1);//74
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//77
        ExcelBuf.AddColumn(OrgState, FALSE, '', TRUE, FALSE, FALSE, '', 1);//78
        ExcelBuf.AddColumn(FinalState, FALSE, '', TRUE, FALSE, FALSE, '', 1);//79
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Nature of Supply", FALSE, '', TRUE, FALSE, FALSE, '', 1);//80
        ExcelBuf.AddColumn(BDNNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//81 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(BDNDate, FALSE, '', TRUE, FALSE, FALSE, '', 1);//82 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(VesselName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//83 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(IRNGeneratedTxtCrMemo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//RSPL/AviMali
        ExcelBuf.AddColumn("Sales Cr.Memo Header".IRN, FALSE, '', FALSE, FALSE, FALSE, '', 1);//RSPL/AviMali
    end;

    // [Scope('Internal')]
    procedure MakeExcelCrmemoGrandTotal()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//60
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//65
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//70
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//75
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//81
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 13Mar2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 13Mar2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//02Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19Jun2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPL/AviMali
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPL/AviMali

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn(-InCrTQty, FALSE, '', TRUE, FALSE, TRUE, '#####0.00', 0);//26
        InCrTQty := 0; //01Aug2017

        ExcelBuf.AddColumn(-InCrTQtyBase, FALSE, '', TRUE, FALSE, TRUE, '#####000.00', 0);//27
        InCrTQtyBase := 0;//01Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn(-FcTotal, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//32
        FcTotal := 0;//30Dec2017
        ExcelBuf.AddColumn(-InTINRAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//33
        InTINRAmt := 0;//01Aug2017

        ExcelBuf.AddColumn(TTAVPDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//34
        ExcelBuf.AddColumn(TTRegDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//02Nov2018

        ExcelBuf.AddColumn(InCrTrans, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//35
        ExcelBuf.AddColumn(InCrCash, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//36
        ExcelBuf.AddColumn((TTDiscAmt + InCrLineDisc), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//37
        ExcelBuf.AddColumn((InCrTrans + InCrCash + TTAVPDiscAmt + TTRegDiscAmt + TTDiscAmt + InCrLineDisc), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//38
        InCrTrans := 0; //01Aug2017
        InCrCash := 0;//01Aug2017
        InCrLineDisc := 0; //01Aug2017
        TTAVPDiscAmt := 0; //02Aug2017
        TTDiscAmt := 0;//02Aug2017
        TTRegDiscAmt := 0;//02Nov2018

        ExcelBuf.AddColumn(0.0, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//39
        ExcelBuf.AddColumn(0.0, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//40
        ExcelBuf.AddColumn(0.0, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//41
        ExcelBuf.AddColumn(0.0, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//42
        ExcelBuf.AddColumn(-InCrTaxBase, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//43
        InCrTaxBase := 0; //01Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn(-TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//45
        ExcelBuf.AddColumn(-TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0); //46
        ExcelBuf.AddColumn(-TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//47
        ExcelBuf.AddColumn(-TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//48

        ExcelBuf.AddColumn(-TTCessAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//49 //13Mar2018
        ExcelBuf.AddColumn(-TTTCSAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//50 //13Mar2018
        TTCessAmt := 0;
        TTTCSAmt := 0;

        ExcelBuf.AddColumn(CrRound, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//49
        CrRound := 0; //01Aug2017

        ExcelBuf.AddColumn(-InCrFCAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//50
        InCrFCAmt := 0;//01Aug2017
        ExcelBuf.AddColumn(-InCrAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//51
        InCrAmt := 0; //01Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53 -InCrAddTax
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);  //54 InCrFRAmt
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55 gfreightamt1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1); //56 GfreightNonTxAmount
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 58 InCrFCAmt
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//59 - InCrAmt - CrRound
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//78
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19Jun2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPL/AviMali
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPL/AviMali
    end;

    //  [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1 31July2017
        ExcelBuf.AddColumn("Sales Invoice Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2 31July2017
        ExcelBuf.AddColumn("Sales Invoice Header"."Cancelled Invoice", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4 31July2017
        ExcelBuf.AddColumn(SalesType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Line"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5 31July2017
        ExcelBuf.AddColumn(StateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6 31July2017
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7 31July2017
        ExcelBuf.AddColumn(GSTStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8 31July2017
        ExcelBuf.AddColumn(LocationRec."GST Registration No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//9 31July2017
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(ResCenName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Invoice Header"."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        //ExcelBuf.AddColumn(Cust."Full Name",FALSE,'',FALSE,FALSE,FALSE,'',1);//
        ExcelBuf.AddColumn(Cust.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        //ExcelBuf.AddColumn(Cust."GST Registration No.",FALSE,'',FALSE,FALSE,FALSE,'',1);//15 31July2017
        //>>24May2019
        CLEAR(DGST04);
        DGST04.SETCURRENTKEY("Document No.");
        DGST04.SETRANGE("Document No.", "Sales Invoice Line"."Document No.");
        IF DGST04.FINDFIRST THEN;
        ExcelBuf.AddColumn(DGST04."Buyer/Seller Reg. No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        //<<24May2019
        ExcelBuf.AddColumn("Sales Invoice Header"."Ship-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn(Cust."GST Registration No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17 31July2017
        ExcelBuf.AddColumn(GSTPlaceofSupply, FALSE, '', FALSE, FALSE, FALSE, '', 1);//18 31July2017
        ExcelBuf.AddColumn(GSTSupplyCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//19 31July2017

        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
            IF "recItem-Supp".GET("Sales Invoice Line"."Final Item No.") THEN;
            ExcelBuf.AddColumn("Sales Invoice Line"."Final Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
            ExcelBuf.AddColumn("recItem-Supp".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        END ELSE BEGIN
            ExcelBuf.AddColumn("Sales Invoice Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
            ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        END;

        ExcelBuf.AddColumn("Sales Invoice Line"."HSN/SAC Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//22 31July2017
        ExcelBuf.AddColumn("Sales Invoice Line"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//23 31July2017

        ExcelBuf.AddColumn("Sales Invoice Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//24
        ExcelBuf.AddColumn("Sales Invoice Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//25
        ExcelBuf.AddColumn("Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0); //26


        ExcelBuf.AddColumn("Sales Invoice Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27


        ExcelBuf.AddColumn("Sales Invoice Header"."Currency Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//28
        //>>02Aug2017
        IF "Sales Invoice Header"."Currency Code" <> '' THEN
            ExcelBuf.AddColumn(1 / "Sales Invoice Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//29
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//29
        //>>02Aug2017
        ExcelBuf.AddColumn(MRPPricePerLt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//MRP //29
        ExcelBuf.AddColumn("Sales Invoice Line"."Basic Price", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//RSPLSUM
        ExcelBuf.AddColumn("Sales Invoice Line"."Freight/Other Chgs. Primary", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//RSPLSUM
        ExcelBuf.AddColumn("Sales Invoice Line"."Freight/Other Chgs. Secondary", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//RSPLSUM
        ExcelBuf.AddColumn(ListPricePerlt + ProdDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//List //30

        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//FCPrice //31
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//FCAmt //32
            IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
                ExcelBuf.AddColumn("Sales Invoice Line".Amount, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//INR Amt //33
                TINRAmt += "Sales Invoice Line".Amount;//23Feb2017
                InTINRAmt += "Sales Invoice Line".Amount;//27Feb2017
            END ELSE BEGIN
                ExcelBuf.AddColumn(INRamt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); //33
                TINRAmt += INRamt;//23Feb2017
                InTINRAmt += INRamt;//27Feb2017
            END;
        END ELSE BEGIN

            ExcelBuf.AddColumn((UPricePerLT), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//FCPrice //31
            INRTot += UPricePerLT;//23Feb2017
            ExcelBuf.AddColumn(UPrice * "Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//FCAMt //32
                                                                                                                        //ExcelBuf.AddColumn(UPrice/"Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//INR UNIT //
                                                                                                                        //TINRUnit += UPrice/"Sales Invoice Header"."Currency Factor";//23Feb2017
            TINRUnit += UPrice * "Sales Invoice Line".Quantity;//30Dec2017
            InTINRUnit += UPrice * "Sales Invoice Line".Quantity;//30Dec2017

            ExcelBuf.AddColumn(INRamt / "Sales Invoice Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//INR Amt //33
            TINRAmt += INRamt / "Sales Invoice Header"."Currency Factor";//23Feb2017
            InTINRAmt += INRamt / "Sales Invoice Header"."Currency Factor";//27Feb2017
        END;

        ExcelBuf.AddColumn(-(AVPDiscAmt), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//34
        ExcelBuf.AddColumn(-(RegDiscAmt), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//34 02Nov2018

        ExcelBuf.AddColumn(-(TransportSubsidy), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//35 31July2017

        ExcelBuf.AddColumn(-(CashDiscount), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//36 31July2017

        ExcelBuf.AddColumn(-(LineDiscAmt + DiscAmt), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//37 01Aug2017

        ExcelBuf.AddColumn(-(ABS(TransportSubsidy) + ABS(CashDiscount) + ABS(AVPDiscAmt) + ABS(DiscAmt) +
        ABS(RegDiscAmt) + ABS(LineDiscAmt)), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//38 31July2017
                                                                                              //>>23Feb2017


        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//39
        END ELSE BEGIN
            ExcelBuf.AddColumn(FCFRIEGHTAmount, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); // 39

        END;

        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            IF FRIEGHTAmount <> 0 THEN BEGIN
                ExcelBuf.AddColumn(/*"Sales Invoice Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt
                                   - DISCOUNTAmount - SurChargeInterest, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//40
                TFRIEGHTAmount += 0;// "Sales Invoice Line"."Charges To Customer";//23Feb2017
                                    //>>27Feb
                InFG += /*"Sales Invoice Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt - DISCOUNTAmount - SurChargeInterest;
                TInFG += /*"Sales Invoice Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt - DISCOUNTAmount - SurChargeInterest;
                //<<
            END ELSE BEGIN
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//40
            END;
        END ELSE BEGIN
            ExcelBuf.AddColumn((/*"Sales Invoice Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt)
            / "Sales Invoice Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//40
            TFRIEGHTAmount += /*"Sales Invoice Line"."Charges To Customer"*/0 / "Sales Invoice Header"."Currency Factor";//23Feb2017
                                                                                                                         //>>27Feb
            InFG += (/*"Sales Invoice Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt) / "Sales Invoice Header"."Currency Factor";
            TInFG += (/*"Sales Invoice Line"."Charges To Customer"*/0 - ENTRYTax - CessAmount - AddTaxAmt) / "Sales Invoice Header"."Currency Factor";
            //<<
        END;

        ExcelBuf.AddColumn(0.0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//41 31July2017

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//42 31July2017


        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            IF "Sales Invoice Header"."Posting Date" >= 20170107D THEN //070117D
                ExcelBuf.AddColumn(/*"Sales Invoice Line"."GST Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)
            ELSE
                ExcelBuf.AddColumn(/*"Sales Invoice Line"."Tax Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);
            /*
            IF "Sales Invoice Line"."GST Base Amount" <> 0 THEN
              ExcelBuf.AddColumn("Sales Invoice Line"."GST Base Amount",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0)//43 31July2017
            ELSE
            IF "Sales Invoice Line"."Free of Cost" THEN
              ExcelBuf.AddColumn("Sales Invoice Line"."GST Base Amount",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0)
            ELSE
              ExcelBuf.AddColumn("Sales Invoice Line"."Tax Base Amount",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);
            */
        END ELSE BEGIN
            ExcelBuf.AddColumn("Sales Invoice Line"."Line Amount" / "Sales Invoice Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43 02Aug2017
                                                                                                                                                            //ExcelBuf.AddColumn("Sales Invoice Line"."GST Base Amount"/"Sales Invoice Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//43

        END;
        //ExcelBuf.AddColumn("Sales Invoice Line"."GST %",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//44 31July2017
        ExcelBuf.AddColumn(GSTPer, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//44 RB-N 20Sep2017
        ExcelBuf.AddColumn(vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//45 31July2017
        ExcelBuf.AddColumn(vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//46 31July2017
        ExcelBuf.AddColumn(vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//47 31July2017
        ExcelBuf.AddColumn(vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//48 31July2017
        ExcelBuf.AddColumn(CessAmt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//49//13Mar2018
        ExcelBuf.AddColumn(TCSAmt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//50//13Mar2018

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//49 31July2017 //RoundoffAmount

        IF "Sales Invoice Header"."Currency Code" = '' THEN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1) //50
        ELSE
            ExcelBuf.AddColumn(/*"Sales Invoice Line"."Amount To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //50



        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(/*"Sales Invoice Line"."Amount To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); //51

        END ELSE BEGIN
            ExcelBuf.AddColumn(/*"Sales Invoice Line"."Amount To Customer"*/0 / "Sales Invoice Header"."Currency Factor"
            , FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//51
        END;

        ExcelBuf.AddColumn(ProdDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//52
        ExcelBuf.AddColumn(TRRNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//53
        ExcelBuf.AddColumn(BRTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//54
        ExcelBuf.AddColumn(SalesName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//55
        ExcelBuf.AddColumn(LineSalespersonName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//56
        ExcelBuf.AddColumn("Sales Invoice Header"."Order No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//57
        ExcelBuf.AddColumn("Sales Invoice Header"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//58
        //ExcelBuf.AddColumn("Sales Invoice Header"."Vehicle No.",FALSE,'',FALSE,FALSE,FALSE,'',1);//59
        ExcelBuf.AddColumn(VehicleNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//50 //AR 25Nov2020
        ExcelBuf.AddColumn("Sales Invoice Header"."Road Permit No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//60 //E-way BillNo
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//61 //E-way BillDate
        ExcelBuf.AddColumn("Sales Invoice Header"."LR/RR No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//62
        ExcelBuf.AddColumn("Sales Invoice Header"."LR/RR Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//63
        ExcelBuf.AddColumn("Sales Invoice Header"."External Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//64
        ExcelBuf.AddColumn(ShipmentAgentName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//65
        ExcelBuf.AddColumn(CustomerType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//66
        ExcelBuf.AddColumn("Sales Invoice Header"."Payment Terms Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//67
        ExcelBuf.AddColumn(SalesGroup, FALSE, '', FALSE, FALSE, FALSE, '', 1);//68
        ExcelBuf.AddColumn("Sales Invoice Header"."Driver's Mobile No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//69
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//70 //Supp. Org Invoice
        ExcelBuf.AddColumn(REGION, FALSE, '', FALSE, FALSE, FALSE, '', 1);//71
        ExcelBuf.AddColumn(SPRegionCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//19Jun2019
        ExcelBuf.AddColumn(SPZoneCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn(FORMAT("Sales Invoice Header"."Freight Type"), FALSE, '', FALSE, FALSE, FALSE, '', 1);//72
        ExcelBuf.AddColumn(CSOAppDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//73
        ExcelBuf.AddColumn(ProductGroupName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//74
        ExcelBuf.AddColumn(tComment, FALSE, '', FALSE, FALSE, FALSE, '', 1);//75
        //>>RB-N 25Oct2017 DivisionName
        DivisionText := '';
        Dim25.RESET;
        IF Dim25.GET('DIVISION', "Sales Invoice Line"."Shortcut Dimension 1 Code") THEN BEGIN
            DivisionText := Dim25.Name;
        END;
        //<<RB-N 25Oct2017 DivisionName
        //ExcelBuf.AddColumn("Sales Invoice Line"."Shortcut Dimension 1 Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//76
        ExcelBuf.AddColumn("Sales Invoice Line"."Shortcut Dimension 1 Code" + ' - ' + DivisionText, FALSE, '', FALSE, FALSE, FALSE, '', 1);//76 25Oct2017 DivisionName
        ExcelBuf.AddColumn("Sales Invoice Header"."Customer Receipt Date", FALSE, '', FALSE, FALSE, FALSE, '', 1);//77
        ExcelBuf.AddColumn(OrgState, FALSE, '', FALSE, FALSE, FALSE, '', 1);//78
        ExcelBuf.AddColumn(FinalState, FALSE, '', FALSE, FALSE, FALSE, '', 1);//79
        ExcelBuf.AddColumn("Sales Invoice Header"."Nature of Supply", FALSE, '', FALSE, FALSE, FALSE, '', 1);//80
        ExcelBuf.AddColumn(BDNNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//81 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(BDNDate, FALSE, '', FALSE, FALSE, FALSE, '', 1);//82 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(VesselName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//83 RSPLSUM 23Jun2020
        ExcelBuf.AddColumn(COPYSTR(txtGRN, 2, STRLEN(txtGRN)), FALSE, '', FALSE, FALSE, FALSE, '', 1);//84 RSPLSUM 29Jul2020
        ExcelBuf.AddColumn(COPYSTR(txtGRNNo, 2, STRLEN(txtGRNNo)), FALSE, '', FALSE, FALSE, FALSE, '', 1);//85 RSPLSUM 04Aug2020
        ExcelBuf.AddColumn(COPYSTR(txtSalRetRcpt, 2, STRLEN(txtSalRetRcpt)), FALSE, '', FALSE, FALSE, FALSE, '', 1);//86 RSPLSUM 05Aug2020
        ExcelBuf.AddColumn("Sales Invoice Header"."Mintifi Channel Finance", FALSE, '', FALSE, FALSE, FALSE, '', 1);//87 RSPLSUM 11Aug2020

        //AR
        ExcelBuf.AddColumn(IRNGeneratedTxt, FALSE, '', FALSE, FALSE, FALSE, '', 1);//85
        ExcelBuf.AddColumn("Sales Invoice Header"."EWB No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//86
        ExcelBuf.AddColumn("Sales Invoice Header"."EWB Date", FALSE, '', FALSE, FALSE, FALSE, '', 1);//87
                                                                                                     //AR

        "Sales Invoice Header".CALCFIELDS("Sales Invoice Header".IRN);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn("Sales Invoice Header".IRN, FALSE, '', FALSE, FALSE, FALSE, '', 1);//88 RSPLSUM 09Nov2020
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."TDS/TCS Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//89 RSPLSUM 09Nov2020

    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Sales Register Invoice- GST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//42
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//51
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//78
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//79
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//80
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//81
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//82 13Mar2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 13Mar2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//83 02Nov2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19Jun2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 23Jun2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 29Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 04Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 05Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 11Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 09Nov2020
                                                                    /*
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//81
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//82
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//83
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//84
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//85
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//86
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//87//004
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//88  //004
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//89 //
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//90 //
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//91 //
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//92 //
                                                                    *///31July2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Cancelled Invoice', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Movement Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        /*
        ExcelBuf.AddColumn('Origin State OF Our',FALSE,'',TRUE,FALSE,TRUE,'@',1);//6 31July2017
        ExcelBuf.AddColumn('GPPL Supply Location Name',FALSE,'',TRUE,FALSE,TRUE,'@',1);//7 31July2017
        */
        ExcelBuf.AddColumn('Origin State', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6 01Aug2017
        ExcelBuf.AddColumn('GPPL Billing Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7 01Aug2017
        ExcelBuf.AddColumn('GPPL State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8 31Juy2017
        ExcelBuf.AddColumn('GPPL GSTIN No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9 31Juy2017
        ExcelBuf.AddColumn('Resp. Centre Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Resp. Centre Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Customer Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Customer Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Details of Customer ( Billed To )', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 31July2017
        ExcelBuf.AddColumn('Bill to Location GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15 31July2017
        ExcelBuf.AddColumn('Details of Consignee ( Shipped To)', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16 31July2017
        ExcelBuf.AddColumn('Ship to Location GSTIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17 31July2017
        ExcelBuf.AddColumn('Place of Supply', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18 31July2017
        ExcelBuf.AddColumn('Supply State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19 31July2017
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('Product HSN Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 31July2017
        ExcelBuf.AddColumn('Batch No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 31July2017
        ExcelBuf.AddColumn('Unit of Measure', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Packing Qty per UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('UOM Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        ExcelBuf.AddColumn('Total Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('Currency Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('Exchange Rate', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 02Aug2017
        ExcelBuf.AddColumn('MRP Per Lt/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('Basic Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM
        ExcelBuf.AddColumn('Freight Primary', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM
        ExcelBuf.AddColumn('Freight Secondary', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM
        ExcelBuf.AddColumn('List Price per Lt/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
                                                                                         /*
                                                                                         ExcelBuf.AddColumn('Ass. Value',FALSE,'',TRUE,FALSE,TRUE,'@',1);//  //EBT MILAN
                                                                                         ExcelBuf.AddColumn('INR Price Per KG/Ltr',FALSE,'',TRUE,FALSE,TRUE,'@',1);//
                                                                                         *///31July2017

        ExcelBuf.AddColumn('FC Price Per KG/Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('FC Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33
        //ExcelBuf.AddColumn('AVP DISCOUNT',FALSE,'',TRUE,FALSE,TRUE,'@',1);//34
        ExcelBuf.AddColumn('National Discount Scheme', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//02Nov2018
        ExcelBuf.AddColumn('Regional Discount Scheme', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//02Nov2018
        ExcelBuf.AddColumn('Transport Subsidy', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35
        ExcelBuf.AddColumn('Cash Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//36
        ExcelBuf.AddColumn('Invoice Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//37 01Aug2017
        ExcelBuf.AddColumn('Total Discounts', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//38
        ExcelBuf.AddColumn('FC FREIGHT', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//39
        ExcelBuf.AddColumn('FREIGHT', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//40
        ExcelBuf.AddColumn('Insurance', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//41
        ExcelBuf.AddColumn('Other Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//42
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//43
        ExcelBuf.AddColumn('Tax %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//44
        ExcelBuf.AddColumn('CGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//45 31July2017
        ExcelBuf.AddColumn('SGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//46 31July2017
        ExcelBuf.AddColumn('IGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//47 31July2017
        ExcelBuf.AddColumn('UTGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//48 31July2017
        ExcelBuf.AddColumn('GST Comp. Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//49 13Mar2018
        ExcelBuf.AddColumn('TCS Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//50 13Mar2018
        ExcelBuf.AddColumn('R/Off', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//51 31July2017
        ExcelBuf.AddColumn('FC Gross Invoice Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//50
        ExcelBuf.AddColumn('Gross Invoice Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//51
        ExcelBuf.AddColumn('Product Discount Per Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//52
        ExcelBuf.AddColumn('TRR Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//53
        ExcelBuf.AddColumn('BRT Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//54
        ExcelBuf.AddColumn('Header Salesperson Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//55
        ExcelBuf.AddColumn('Line Salesperson Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//56
        ExcelBuf.AddColumn('CSO Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//57
        ExcelBuf.AddColumn('CSO Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//58
        ExcelBuf.AddColumn('Vehicle Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//59
        ExcelBuf.AddColumn('E-Way Bill No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//60
        ExcelBuf.AddColumn('E-Way Bill Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//61
        ExcelBuf.AddColumn('LR Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//62
        ExcelBuf.AddColumn('LR Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//63
        ExcelBuf.AddColumn('PO No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//64
        ExcelBuf.AddColumn('Transporter Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//65
        ExcelBuf.AddColumn('Customer Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//66
        ExcelBuf.AddColumn('Payment Terms Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//67
        ExcelBuf.AddColumn('Sales Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//68
        ExcelBuf.AddColumn('Drivers Mobile No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//69
        ExcelBuf.AddColumn('Supp. Org Invoice', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//70
        ExcelBuf.AddColumn('Customer Region', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//71
        ExcelBuf.AddColumn('SalesPerson Region', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19Jun2019
        ExcelBuf.AddColumn('SalesPerson Zone', FALSE, '', TRUE, FALSE, TRUE, '', 1);//RSPLSUM 16Apr2020
        ExcelBuf.AddColumn('Freight Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//72
        ExcelBuf.AddColumn('CSO Approved Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//73
        ExcelBuf.AddColumn('Product Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//74
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//75
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//76
        ExcelBuf.AddColumn('Customer Receipt Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//77
        ExcelBuf.AddColumn('Origin State', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//78 31July2017
        ExcelBuf.AddColumn('Destination State', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//79 31July2017
        ExcelBuf.AddColumn('Nature of Transaction', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//80 31July2017
        ExcelBuf.AddColumn('BDN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//81 23Jun2020
        ExcelBuf.AddColumn('BDN Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//82 23Jun2020
        ExcelBuf.AddColumn('Vessel Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//83 23Jun2020
        ExcelBuf.AddColumn('GRN Vessel', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//84 29Jul2020
        ExcelBuf.AddColumn('GRN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//85 04Aug2020
        ExcelBuf.AddColumn('Sales Return Receipt', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//86 05Aug2020
        ExcelBuf.AddColumn('Mintifi Channel Finance', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//87 11Aug2020

        //AR 29Oct20
        ExcelBuf.AddColumn('IRN Generated', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//88
        ExcelBuf.AddColumn('EWB No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//89
        ExcelBuf.AddColumn('EWB Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//90
                                                                             //AR 29Oct20

        ExcelBuf.AddColumn('IRN', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//87 RSPLSUM 09Nov2020
        ExcelBuf.AddColumn('TCS Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//87 RSPLSUM 09Nov2020

    end;

    local procedure FindGRNYSR(varRecILeOrg: Record 32; var GRNNo: Text; var GRNNoReal: Text; var SalesRetRcpt: Text)
    var
        varRecILe: Record 32;
        cdDoc: Code[20];
        cdDoc1: Code[20];
        cdDoc2: Code[20];
        ItemApplnEntry: Record 339;
        ItemLedgEntry: Record 32;
    begin
        //Navigate
        CLEAR(cdDoc);
        CLEAR(cdDoc1);
        CLEAR(cdDoc2);

        varRecILe.RESET;
        varRecILe.SETCURRENTKEY("Document No.", "Posting Date", "Item No.");
        varRecILe.SETRANGE("Document No.", varRecILeOrg."Document No.");
        varRecILe.SETRANGE("Item No.", varRecILeOrg."Item No.");
        IF varRecILe.FINDSET THEN
            REPEAT
                IF cdDoc <> varRecILe."Document No." THEN BEGIN
                    cdDoc := varRecILe."Document No.";
                    IF varRecILe.Positive THEN BEGIN
                        ItemApplnEntry.RESET;
                        ItemApplnEntry.SETCURRENTKEY("Inbound Item Entry No.", "Outbound Item Entry No.", "Cost Application");
                        ItemApplnEntry.SETRANGE("Inbound Item Entry No.", varRecILe."Entry No.");
                        ItemApplnEntry.SETFILTER("Outbound Item Entry No.", '<>%1', 0);
                        ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                        IF ItemApplnEntry.FIND('-') THEN
                            REPEAT
                                //InsertTempEntry(ItemApplnEntry."Outbound Item Entry No.",ItemApplnEntry.Quantity);

                                ItemLedgEntry.GET(ItemApplnEntry."Outbound Item Entry No.");
                                IF cdDoc1 <> ItemLedgEntry."Document No." THEN BEGIN
                                    cdDoc1 := ItemLedgEntry."Document No.";

                                    IF NOT (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Consumption) THEN BEGIN
                                        IF ItemApplnEntry.Quantity * ItemLedgEntry.Quantity >= 0 THEN BEGIN
                                            //YSR BEGIN
                                            //---Find Transfr
                                            IF ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Transfer THEN BEGIN
                                                FindGRNYSR(ItemLedgEntry, GRNNo, GRNNoReal, SalesRetRcpt);
                                            END ELSE
                                                //YSR END
                                                IF (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Purchase) OR
                                                    (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::"Positive Adjmt.") THEN BEGIN
                                                    GRNFound := TRUE;
                                                    //GRNDocNo[i] += ',' + ItemLedgEntry."Document No.";
                                                    CLEAR(GRNVessel);
                                                    RecPRH.RESET;
                                                    IF RecPRH.GET(ItemLedgEntry."Document No.") THEN BEGIN
                                                        RecPRH.CALCFIELDS("Vessel Name");
                                                        GRNVessel := RecPRH."Vessel Name";
                                                    END;
                                                    GRNNo += ',' + GRNVessel;
                                                    GRNNoReal += ',' + RecPRH."No.";//RSPLSUM 04Aug2020
                                                                                    //MESSAGE('%1', GRNNo );

                                                    i += 1;

                                                    EXIT;
                                                END;
                                        END;
                                    END;
                                END;
                            UNTIL ItemApplnEntry.NEXT = 0;
                    END ELSE BEGIN
                        ItemApplnEntry.RESET;
                        ItemApplnEntry.SETCURRENTKEY("Outbound Item Entry No.", "Item Ledger Entry No.", "Cost Application");
                        ItemApplnEntry.SETRANGE("Outbound Item Entry No.", varRecILe."Entry No.");
                        ItemApplnEntry.SETRANGE("Item Ledger Entry No.", varRecILe."Entry No.");
                        ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                        IF ItemApplnEntry.FIND('-') THEN
                            REPEAT
                                ItemLedgEntry.GET(ItemApplnEntry."Inbound Item Entry No.");
                                IF cdDoc2 <> ItemLedgEntry."Document No." THEN BEGIN
                                    cdDoc2 := ItemLedgEntry."Document No.";

                                    IF NOT (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Consumption) THEN BEGIN
                                        IF -ItemApplnEntry.Quantity * ItemLedgEntry.Quantity >= 0 THEN BEGIN
                                            //YSR BEGIN
                                            //---Find Transfr
                                            IF ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Transfer THEN BEGIN
                                                FindGRNYSR(ItemLedgEntry, GRNNo, GRNNoReal, SalesRetRcpt);
                                            END ELSE
                                                //YSR END
                                                IF (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Purchase) OR
                                                    (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::"Positive Adjmt.") THEN BEGIN
                                                    GRNFound := TRUE;
                                                    CLEAR(GRNVessel);
                                                    RecPRH.RESET;
                                                    IF RecPRH.GET(ItemLedgEntry."Document No.") THEN BEGIN
                                                        RecPRH.CALCFIELDS("Vessel Name");
                                                        GRNVessel := RecPRH."Vessel Name";
                                                    END;
                                                    FOR j := 1 TO i DO//s
                                                    BEGIN//s
                                                        IF TestGRNNo[j] = ItemLedgEntry."Document No." THEN//s
                                                            GRNForLoopFound := TRUE;//s
                                                    END;//s
                                                    IF NOT GRNForLoopFound THEN BEGIN//s
                                                                                     //GRNNo += ',' + ItemLedgEntry."Document No.";//ysr
                                                        GRNNo += ',' + GRNVessel;//ysr
                                                        GRNNoReal += ',' + RecPRH."No.";//RSPLSUM 04Aug2020
                                                        PrintGRNNo[k] += ',' + ItemLedgEntry."Document No.";
                                                        TestGRNNo[i] := ItemLedgEntry."Document No.";//s
                                                        i += 1;
                                                        k += 1;
                                                    END;
                                                    GRNForLoopFound := FALSE;
                                                END;
                                            //RSPLSUM 04Aug2020>>Find Invoice from Sales Return Receipt--
                                            IF (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Sale) AND
                                                (ItemLedgEntry."Document Type" = ItemLedgEntry."Document Type"::"Sales Return Receipt") THEN BEGIN
                                                SalesRetRcpt += ',' + ItemLedgEntry."Document No.";//RSPLSUM 05Aug2020
                                                RecValEnt.RESET;
                                                RecValEnt.SETRANGE("Item Ledger Entry No.", ItemLedgEntry."Entry No.");
                                                RecValEnt.SETRANGE("Document Type", RecValEnt."Document Type"::"Sales Credit Memo");
                                                IF RecValEnt.FINDFIRST THEN BEGIN
                                                    RecSCMH.RESET;
                                                    IF RecSCMH.GET(RecValEnt."Document No.") THEN BEGIN
                                                        RecSalInvHdr.RESET;
                                                        IF RecSalInvHdr.GET(RecSCMH."Applies-to Doc. No.") THEN BEGIN
                                                            IF (RecSalInvHdr."Customer Posting Group" = 'FO') OR
                                                              (RecSalInvHdr."Customer Posting Group" = 'BOILS') THEN BEGIN
                                                                RecValEntNew.RESET;
                                                                RecValEntNew.SETRANGE("Document No.", RecSalInvHdr."No.");
                                                                IF RecValEntNew.FINDFIRST THEN BEGIN
                                                                    IF RecValEntNew."Item Ledger Entry No." <> 0 THEN BEGIN
                                                                        ILERecNew.RESET;
                                                                        IF ILERecNew.GET(RecValEntNew."Item Ledger Entry No.") THEN BEGIN
                                                                            FindGRNYSR(ILERecNew, GRNNo, GRNNoReal, SalesRetRcpt);
                                                                        END;
                                                                    END;
                                                                END;
                                                            END;
                                                        END;
                                                    END;
                                                END;
                                            END;
                                            //RSPLSUM 04Aug2020<<
                                        END;
                                    END;//Consmption End
                                END;

                            UNTIL ItemApplnEntry.NEXT = 0;
                    END;
                END;
            UNTIL varRecILe.NEXT = 0;
    end;
}


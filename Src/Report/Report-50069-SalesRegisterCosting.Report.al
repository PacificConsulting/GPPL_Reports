report 50069 "Sales Register Costing"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesRegisterCosting.rdl';
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
                RequestFilterFields = "Sell-to Customer No.", "Shortcut Dimension 1 Code", "Shortcut Dimension 2 Code", "Salesperson Code";
                dataitem("Sales Invoice Line"; "Sales Invoice Line")
                {
                    DataItemLink = "Document No." = FIELD("No.");
                    DataItemTableView = SORTING("Document No.", "Line No.")
                                        WHERE(Type = FILTER(Item | 'Charge (Item)'),
                                              Quantity = FILTER(<> 0));
                    RequestFilterFields = "No.";

                    trigger OnAfterGetRecord()
                    begin

                        TSILCOUNT += 1;//10Mar2017

                        //>>1

                        UPricePerLT := 0;
                        UPrice := 0;
                        FOCAmount := 0;
                        DISCOUNTAmount := 0;
                        TransportSubsidy := 0;
                        CashDiscount := 0;
                        ENTRYTax := 0;
                        CessAmount := 0;
                        FRIEGHTAmount := 0;
                        AddTaxAmt := 0;
                        BRTNo := '';
                        MRPPricePerLt := 0;
                        ListPricePerlt := 0;
                        StateDiscount := 0;

                        // IF "Sales Invoice Line"."Abatement %" <> 0 THEN
                        MRPPricePerLt := 0;//"Sales Invoice Line"."MRP Price";

                        // IF "Sales Invoice Line"."Price Inclusive of Tax" THEN
                        ListPricePerlt := "Sales Invoice Line"."Unit Price Incl. of Tax" / "Sales Invoice Line"."Qty. per Unit of Measure";
                        // ELSE
                        ListPricePerlt := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure";


                        UPricePerLT := "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure";
                        UPrice := "Sales Invoice Line"."Unit Price";


                        /*  ExciseSetup.RESET;
                         ExciseSetup.SETRANGE("Excise Bus. Posting Group", "Sales Invoice Line"."Excise Bus. Posting Group");
                         ExciseSetup.SETRANGE("Excise Prod. Posting Group", "Excise Prod. Posting Group");
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

                         StructureDetailes.RESET;
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
                             UNTIL StructureDetailes.NEXT = 0;

                         StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'STATE DISC');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 StateDiscount += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;


                         StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'TPTSUBSIDY', 'TRANSSUBS');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 TransportSubsidy += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;

                         StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'CASHDISC');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 CashDiscount += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;

                         StructureDetailes.RESET;
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

                         StructureDetailes.RESET;
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                         StructureDetailes.SETFILTER("Tax/Charge Code", '%1|%2', 'CESS', 'C1');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 CessAmount += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;

                         StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'FREIGHT');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 FRIEGHTAmount += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;


                         StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                         StructureDetailes.SETRANGE("Tax/Charge Code", 'ADDL TAX');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 AddTaxAmt += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;

                         StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'ADD TAX');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 AddTaxAmt += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;

                         StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'ADD TAX');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT
                                 AddTaxAmt += StructureDetailes.Amount;
                             UNTIL StructureDetailes.NEXT = 0;


                         RG23D.RESET;
                         RG23D.SETCURRENTKEY("Item No.", "Location Code", "Lot No.", "Transaction Type", Type);
                         RG23D.SETRANGE(RG23D."Item No.", "Sales Invoice Line"."No.");
                         RG23D.SETRANGE(RG23D."Location Code", "Sales Invoice Line"."Location Code");
                         RG23D.SETRANGE(RG23D."Lot No.", "Sales Invoice Line"."Lot No.");
                         RG23D.SETRANGE(RG23D."Transaction Type", RG23D."Transaction Type"::Purchase);
                         RG23D.SETRANGE(RG23D.Type, RG23D.Type::Transfer);
                         IF RG23D.FINDFIRST THEN BEGIN */
                        TRRNo := '';// RG23D."Document No.";
                        TransferReceipt.RESET;
                        TransferReceipt.SETRANGE(TransferReceipt."No.", TRRNo);
                        IF TransferReceipt.FINDFIRST THEN BEGIN
                            TransferShipment.RESET;
                            TransferShipment.SETRANGE(TransferShipment."Transfer Order No.", TransferReceipt."Transfer Order No.");
                            IF TransferShipment.FINDFIRST THEN
                                BRTNo := TransferShipment."No.";
                            // END;
                        END;

                        Mfgdte := 0D;
                        BlendOrderNo := '';
                        ILErec3.RESET;
                        ILErec3.SETRANGE(ILErec3."Item No.", "Sales Invoice Line"."No.");
                        ILErec3.SETRANGE(ILErec3."Lot No.", "Sales Invoice Line"."Lot No.");
                        ILErec3.SETRANGE(ILErec3."Posting Date", 0D, EndDate);
                        ILErec3.SETRANGE(ILErec3."Entry Type", ILErec3."Entry Type"::Output);
                        IF ILErec3.FIND('-') THEN BEGIN
                            Mfgdte := ILErec3."Posting Date";
                            BlendOrderNo := ILErec3."Document No.";
                        END;

                        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
                            "Sales Invoice Line".Quantity := 0;
                            "Sales Invoice Line"."Quantity (Base)" := 0;
                            TotalQty := 0;
                            TotalQtyBase := 0;
                        END;

                        CLEAR(VLEcost);
                        VLE.RESET;
                        VLE.SETRANGE(VLE."Document No.", "Sales Invoice Line"."Document No.");
                        VLE.SETRANGE(VLE."Item No.", "Sales Invoice Line"."No.");
                        VLE.SETRANGE(VLE."Document Line No.", "Sales Invoice Line"."Line No.");
                        VLE.SETRANGE(VLE."Source Code", 'SALES');
                        IF VLE.FINDFIRST THEN BEGIN
                            ILEentryNo := VLE."Item Ledger Entry No.";
                            ILE.RESET;
                            ILE.SETRANGE(ILE."Entry No.", ILEentryNo);
                            IF ILE.FINDFIRST THEN BEGIN
                                ILE.CALCFIELDS(ILE."Cost Amount (Expected)");
                                ILE.CALCFIELDS(ILE."Cost Amount (Actual)");

                                VLEcost := (ILE."Cost Amount (Expected)" + ILE."Cost Amount (Actual)");
                            END;
                        END;

                        TotalVLEcost += VLEcost;

                        //EBT STIVAN--(25/11/2013)----- Costing Code after Revaluation-------START
                        RateofItem := 0;
                        TotalCost := 0;
                        TotalQty := 0;
                        ChargeofItem := 0;
                        UNITCOST := 0;
                        UNITCOST1 := 0;
                        ValuedCost := 0;
                        ILErecEBT.RESET;
                        //MESSAGE('%1..%2',StartDate,EndDate);
                        //ILErecEBT.SETRANGE(ILErecEBT."Posting Date",StartDate,EndDate);
                        ILErecEBT.SETRANGE(ILErecEBT."Item No.", "Sales Invoice Line"."No.");
                        ILErecEBT.SETRANGE(ILErecEBT."Location Code", "Sales Invoice Line"."Location Code");
                        ILErecEBT.SETRANGE(ILErecEBT."Lot No.", "Sales Invoice Line"."Lot No.");
                        ILErecEBT.SETRANGE(ILErecEBT.Open, TRUE);
                        //ILErecEBT.SETFILTER(ILErecEBT."Posting Date",'%1..%2',StartDate,EndDate);
                        ILErecEBT.SETFILTER(ILErecEBT."Document Type", '%1|%2|%3', ILErecEBT."Document Type"::"Transfer Receipt",
                                                  ILErecEBT."Document Type"::"Sales Return Receipt", ILErecEBT."Document Type"::" ");
                        IF ILErecEBT.FINDFIRST THEN BEGIN
                            //REPEAT

                            VLE.RESET;
                            VLE.SETRANGE(VLE."Item Ledger Entry No.", ILErecEBT."Entry No.");
                            VLE.SETRANGE(VLE."Item Charge No.", '');
                            VLE.SETRANGE(VLE."Source Code", 'REVALJNL');
                            IF VLE.FINDSET THEN
                                REPEAT
                                    TotalCost1 := VLE."Cost Amount (Expected)" + VLE."Cost Amount (Actual)";
                                    //   + (UNITCOST * VLE."Valued Quantity");
                                    ValuedCost += VLE."Cost Amount (Expected)" + VLE."Cost Amount (Actual)";
                                    TotalQty1 := VLE."Valued Quantity";
                                    //     TotalQty := VLE."Item Ledger Entry Quantity";
                                    UNITCOST1 := TotalCost1 / TotalQty1;
                                UNTIL VLE.NEXT = 0;

                            //MILAN 0812


                            VLE.RESET;
                            VLE.SETRANGE(VLE."Item Ledger Entry No.", ILErecEBT."Entry No.");
                            VLE.SETFILTER(VLE."Item Charge No.", '<>%1', '');
                            IF VLE.FINDFIRST THEN
                                REPEAT
                                    ChargeofItem += VLE."Cost per Unit";
                                UNTIL VLE.NEXT = 0;

                            IF ChargeofItem < 0 THEN BEGIN
                                RateofItem := RateofItem + ChargeofItem;
                                ChargeofItem := 0;
                            END;
                            //UNTIL ILErecEBT.NEXT = 0;


                            VLE.RESET;
                            VLE.SETRANGE(VLE."Item Ledger Entry No.", ILErecEBT."Entry No.");
                            VLE.SETRANGE(VLE."Item Charge No.", '');
                            VLE.SETFILTER(VLE."Source Code", '<>%1', 'REVALJNL');
                            IF VLE.FINDSET THEN
                                REPEAT
                                    TotalCost += VLE."Cost Amount (Expected)" + VLE."Cost Amount (Actual)";
                                    ValuedCost += VLE."Cost Amount (Expected)" + VLE."Cost Amount (Actual)";
                                    TotalQty := VLE."Valued Quantity";
                                    UNITCOST := TotalCost / TotalQty;
                                UNTIL VLE.NEXT = 0;

                            UNITCOSTNEW := ABS((UNITCOST)) + UNITCOST1;
                            //  UNITCOSTNEW := ABS(((VLEcost)/("Sales Invoice Line".Quantity)/"Sales Invoice Line"."Qty. per Unit of Measure")) + UNITCOST1;
                        END;

                        //EBT STIVAN--(25/11/2013)----- Costing Code after Revaluation---------END




                        TotalQty += "Sales Invoice Line".Quantity;
                        TotalQtyBase += "Sales Invoice Line"."Quantity (Base)";
                        TotalAmount += "Sales Invoice Line".Amount;
                        TotalLineDisc += "Sales Invoice Line"."Line Discount Amount";
                        TotINRval += "Sales Invoice Line".Amount;
                        TotalExciseBaseAmount += 0;//"Sales Invoice Line"."Excise Base Amount";
                        TotalBEDAmount += 0;// "Sales Invoice Line"."BED Amount";
                        TotaleCessAmount += 0;// "Sales Invoice Line"."eCess Amount";
                        TotalSHECessAmount += 0;//"Sales Invoice Line"."SHE Cess Amount";
                        TotalExciseAmount += 0;// "Sales Invoice Line"."Excise Amount";
                        TotalFOCAmount += FOCAmount;
                        TotalStateDiscount += StateDiscount;
                        TotalDISCOUNTAmount += DISCOUNTAmount;
                        TotalTransportSubsidy += TransportSubsidy;
                        TotalCashDiscount += CashDiscount;
                        TotalENTRYTax += ENTRYTax;
                        TotalCessAmount += CessAmount;
                        TotalFRIEGHTAmount += 0;//"Sales Invoice Line"."Charges To Customer";
                        TotalAddTaxAmt += AddTaxAmt;
                        TotalTaxBAse += 0;//"Sales Invoice Line"."Tax Base Amount";
                        tfreightamt += freightamt;

                        IF "Sales Invoice Line"."Tax Liable" = TRUE THEN
                            TotalTaxBAse1 += 0;// "Sales Invoice Line"."Tax Base Amount";

                        TotalTaxAmount += 0;// "Sales Invoice Line"."Tax Amount";
                        TotalAmtCust += 0;// "Sales Invoice Line"."Amount To Customer";

                        GtotalQty += TotalQty;
                        GtotalQtyBase += TotalQtyBase;
                        GtotalAmount += TotalAmount;
                        GtotalLineDisc += TotalLineDisc;
                        GTotalExciseBaseAmount += TotalExciseBaseAmount;
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
                        GfreightAmount += freightamt;


                        //>>06Nov2017 Excel Header
                        //Sales Invoice Header, Header (1) - OnPreSection()
                        NNNNNN += 1;
                        IF NNNNNN = 1 THEN //02Aug2017
                            IF PrintToExcel AND SalesInvoice THEN BEGIN
                                MakeExcelDataHeader;
                            END;
                        //>>06Nov2017 Excel Header


                        IF PrintToExcel AND SalesInvoice THEN BEGIN
                            MakeExcelDataBody;
                        END;

                        //<<1


                        //>>10Mar2017 SalesInvoiceGrouping

                        //Sales Invoice Line, GroupFooter (1) - OnPreSection()
                        //IF PrintToExcel AND SHowInvTotal AND SalesInvoice THEN
                        IF PrintToExcel AND SHowInvTotal AND SalesInvoice AND (SILCOUNT = TSILCOUNT) THEN BEGIN
                            MakeExcelInvTotalGroupFooter;
                        END;



                        //<<10Mar2017 SalesInvoiceGrouping
                    end;

                    trigger OnPreDataItem()
                    begin

                        //>>1

                        CurrReport.CREATETOTALS(TotalFRIEGHTAmount, CessAmount, AddTaxAmt, INRTot, INRTot1, INRTot2);

                        CurrReport.CREATETOTALS(TotalQty, TotalQtyBase, TotalAmount, TotalLineDisc, TotalExciseBaseAmount, TotalBEDAmount, TotaleCessAmount,
                        TotalSHECessAmount, TotalExciseAmount, TotalFOCAmount);
                        CurrReport.CREATETOTALS(TotalDISCOUNTAmount, TotalTransportSubsidy, TotalCashDiscount, TotalENTRYTax,
                        TotalCessAmount, TotalFRIEGHTAmount, TotalAddTaxAmt, TotalTaxBAse, TotalTaxAmount, TotalAmtCust);
                        CurrReport.CREATETOTALS(TotalStateDiscount);
                        //<<1

                        SILCOUNT := COUNT;//10Mar2017
                        TSILCOUNT := 0;//10Mar2017
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    SalesType := '';
                    IF Cust.GET("Sales Invoice Header"."Sell-to Customer No.") THEN;
                    IF LocationRec.GET("Sales Invoice Header"."Location Code") THEN BEGIN
                        LocationCode := LocationRec.Name;
                        LocationSate := LocationRec."State Code";
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


                    IF "Sales Invoice Header"."Gen. Bus. Posting Group" = 'DOMESTIC' THEN BEGIN
                        IF LocationSate = "Sales Invoice Header".State THEN
                            SalesType := 'Local'
                        ELSE
                            SalesType := 'Inter-State';
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

                    CLEAR(freightamt);
                    recGLentry.RESET;
                    recGLentry.SETRANGE(recGLentry."G/L Account No.", '75001010');
                    recGLentry.SETRANGE(recGLentry."Document No.", "Sales Invoice Header"."No.");
                    IF recGLentry.FINDSET THEN
                        REPEAT
                            freightamt := freightamt + recGLentry.Amount
                        UNTIL recGLentry.NEXT = 0;


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

                    //<<1
                end;

                trigger OnPostDataItem()
                begin


                    //>>10Mar2017 SalesInvoiceTotal

                    //Sales Invoice Header, Footer (2) - OnPreSection()
                    IF NNNNNN <> 0 THEN //06Nov2017
                        IF PrintToExcel AND SHowInvTotal AND SalesInvoice THEN BEGIN
                            MakeExcelInvGrandTotal;
                        END;

                    NNNNNN := 0;//06Nov2017
                    //<<SalesInvoiceTotal
                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    "Sales Invoice Header".SETRANGE("Sales Invoice Header"."Posting Date", StartDate, EndDate);
                    StartDate := "Sales Invoice Header".GETRANGEMIN("Sales Invoice Header"."Posting Date");
                    EndDate := "Sales Invoice Header".GETRANGEMAX("Sales Invoice Header"."Posting Date");

                    CurrReport.CREATETOTALS(freightamt);
                    //<<1

                    //>>2

                    /*
                    //Sales Invoice Header, Header (1) - OnPreSection()
                    IF PrintToExcel AND SalesInvoice THEN
                    BEGIN
                     MakeExcelDataHeader;
                    END;
                    */
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
                                              Amount = FILTER(<> 0));

                    trigger OnAfterGetRecord()
                    begin

                        TSCLCOUNT += 1;//10Mar2017
                        //>>1

                        FOCAmount := 0;
                        StateDiscount := 0;
                        DISCOUNTAmount := 0;
                        TransportSubsidy := 0;
                        CashDiscount := 0;
                        ENTRYTax := 0;
                        CessAmount := 0;
                        FRIEGHTAmount := 0;
                        TotalAddTaxAmt := 0;
                        MRPPricePerLt := 0;
                        ListPricePerlt := 0;

                        IF "Sales Cr.Memo Header"."Last Year Sales Return" = TRUE THEN BEGIN
                            MRPPricePerLt := 0;//"Sales Cr.Memo Line"."MRP Price";
                        END ELSE BEGIN
                            //  IF "Sales Cr.Memo Line"."Abatement %" <> 0 THEN
                            MRPPricePerLt := 0;//"Sales Cr.Memo Line"."MRP Price";
                        END;

                        // IF "Sales Cr.Memo Line"."Price Inclusive of Tax" THEN
                        ListPricePerlt := "Sales Cr.Memo Line"."Unit Price Incl. of Tax" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
                        //  ELSE
                        ListPricePerlt := "Sales Cr.Memo Line"."Unit Price" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";


                        /* ExciseSetup.RESET;
                        ExciseSetup.SETRANGE("Excise Bus. Posting Group", "Excise Bus. Posting Group");
                        ExciseSetup.SETRANGE("Excise Prod. Posting Group", "Excise Prod. Posting Group");
                        IF ExciseSetup.FINDLAST THEN; */

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'FOC', 'FOC DIS');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT */
                        FOCAmount += 0;// StructureDetailes.Amount;
                                       //  UNTIL StructureDetailes.NEXT = 0;

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'STATE DISC');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT */
                        StateDiscount += 0;// StructureDetailes.Amount;
                                           //UNTIL StructureDetailes.NEXT = 0;


                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2|%3|%4|%5|%6|%7|%8|%9', 'ADD-DIS', 'AVP-SP-SAN',
                                                   'MONTH-DIS', 'NEW-PRI-SP', 'SPOT-DISC', 'VAT-DIF-DI', 'T2 DISC');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT */
                        DISCOUNTAmount += 0;//StructureDetailes.Amount;
                                            //  UNTIL StructureDetailes.NEXT = 0;

                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'TRANSSUBS');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT */
                        TransportSubsidy += 0;// StructureDetailes.Amount;
                                              // UNTIL StructureDetailes.NEXT = 0;

                        /*  StructureDetailes.RESET;
                         StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                         StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                         StructureDetailes.SETRANGE("Item No.", "No.");
                         StructureDetailes.SETRANGE("Line No.", "Line No.");
                         StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                         StructureDetailes.SETRANGE("Tax/Charge Group", 'CASHDISC');
                         IF StructureDetailes.FINDSET THEN
                             REPEAT */
                        CashDiscount += 0;//StructureDetailes.Amount;
                                          //UNTIL StructureDetailes.NEXT = 0;

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'ENTRYTAX');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT */
                        ENTRYTax += 0;//StructureDetailes.Amount;
                                      //UNTIL StructureDetailes.NEXT = 0;

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'CESS');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT */
                        CessAmount += 0;// StructureDetailes.Amount;
                                        //UNTIL StructureDetailes.NEXT = 0;

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                        StructureDetailes.SETRANGE("Tax/Charge Group", 'FREIGHT');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT */
                        FRIEGHTAmount += 0;//StructureDetailes.Amount;
                                           //UNTIL StructureDetailes.NEXT = 0;

                        /* StructureDetailes.RESET;
                        StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                        StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                        StructureDetailes.SETRANGE("Item No.", "No.");
                        StructureDetailes.SETRANGE("Line No.", "Line No.");
                        StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::"Other Taxes");
                        StructureDetailes.SETRANGE("Tax/Charge Code", 'ADDL TAX');
                        IF StructureDetailes.FINDSET THEN
                            REPEAT */
                        AddTaxAmt += 0;// StructureDetailes.Amount;
                                       // UNTIL StructureDetailes.NEXT = 0;

                        VLE1.RESET;
                        VLE1.SETRANGE(VLE1."Document No.", "Sales Cr.Memo Line"."Document No.");
                        VLE1.SETRANGE(VLE1."Item No.", "Sales Cr.Memo Line"."No.");
                        VLE1.SETRANGE(VLE1."Document Line No.", "Sales Cr.Memo Line"."Line No.");
                        VLE1.SETRANGE(VLE1."Source Code", 'SALES');
                        IF VLE1.FINDFIRST THEN BEGIN
                            ILEentryNo1 := VLE1."Item Ledger Entry No.";

                            ILE1.RESET;
                            ILE1.SETRANGE(ILE1."Entry No.", ILEentryNo1);
                            IF ILE1.FINDFIRST THEN
                                ILE1.CALCFIELDS(ILE1."Cost Amount (Expected)");
                            ILE1.CALCFIELDS(ILE1."Cost Amount (Actual)");

                            VLEcost1 := (ILE1."Cost Amount (Expected)" + ILE1."Cost Amount (Actual)");
                        END;


                        TotalQty += "Sales Cr.Memo Line".Quantity;
                        TotalQtyBase += "Sales Cr.Memo Line"."Quantity (Base)";
                        TotalAmount += "Sales Cr.Memo Line".Amount;
                        TotalLineDisc += "Sales Cr.Memo Line"."Line Discount Amount";
                        TotalExciseBaseAmount += 0;//"Sales Cr.Memo Line"."Excise Base Amount";
                        TotalBEDAmount += 0;// "Sales Cr.Memo Line"."BED Amount";
                        TotaleCessAmount += 0;// "Sales Cr.Memo Line"."eCess Amount";
                        TotalSHECessAmount += 0;//"Sales Cr.Memo Line"."SHE Cess Amount";
                        TotalExciseAmount += 0;//"Sales Cr.Memo Line"."Excise Amount";
                        TotalFOCAmount += FOCAmount;
                        TotalStateDiscount += StateDiscount;
                        TotalDISCOUNTAmount += DISCOUNTAmount;
                        TotalTransportSubsidy += TransportSubsidy;
                        TotalCashDiscount += CashDiscount;
                        TotalENTRYTax += ENTRYTax;
                        TotalCessAmount += CessAmount;
                        TotalFRIEGHTAmount += FRIEGHTAmount;
                        TotalAddTaxAmt += AddTaxAmt;

                        IF "Sales Cr.Memo Line"."Tax Liable" THEN
                            TotalTaxBAse += 0;// "Sales Cr.Memo Line"."Tax Base Amount";

                        TotalTaxAmount += 0;// "Sales Cr.Memo Line"."Tax Amount";
                        TotalAmtCust += 0;//"Sales Cr.Memo Line"."Amount To Customer";

                        IF PrintToExcel AND SalesCreditMemo THEN BEGIN
                            MakeCrMemoBody;
                        END;

                        GtotalQty1 += TotalQty;
                        GtotalQtyBase1 += TotalQtyBase;
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
                        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                            GTotalAmtCust1 += TotalAmtCust;
                        END ELSE
                            IF "Sales Cr.Memo Header"."Currency Factor" > 0 THEN BEGIN       //SN
                                GTotalAmtCust1 += TotalAmtCust / "Sales Cr.Memo Header"."Currency Factor";
                            END;

                        //<<1


                        //>>10Mar2017 SalesReturnGrouping

                        //Sales Cr.Memo Line, GroupFooter (1) - OnPreSection()
                        //IF PrintToExcel AND SHowInvTotal AND SalesCreditMemo THEN
                        IF PrintToExcel AND SHowInvTotal AND SalesCreditMemo AND (SCLCOUNT = TSCLCOUNT) THEN BEGIN
                            MakeExcelCrMemoGrouping;
                        END;


                        //<<10Mar2017 SalesReturnGrouping
                    end;

                    trigger OnPreDataItem()
                    begin

                        //>>1
                        CurrReport.CREATETOTALS(CessAmount, AddTaxAmt);   //EBT Rakshitha
                        //<<1

                        SCLCOUNT := COUNT;//10Mar2017
                        TSCLCOUNT := 0;//10Mar2017
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1



                    SalesType := '';
                    IF Cust.GET("Sales Cr.Memo Header"."Sell-to Customer No.") THEN;
                    IF LocationRec.GET("Sales Cr.Memo Header"."Location Code") THEN
                        LocationCode := LocationRec.Name;

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


                    IF "Sales Cr.Memo Header"."Gen. Bus. Posting Group" = 'DOMESTIC' THEN BEGIN
                        IF LocationSate = "Sales Cr.Memo Header".State THEN
                            SalesType := 'Local'
                        ELSE
                            SalesType := 'Inter-State';
                    END
                    ELSE
                        IF "Sales Cr.Memo Header"."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                            SalesType := 'Export';

                    Saleperson.RESET;
                    Saleperson.SETRANGE(Saleperson.Code, "Sales Cr.Memo Header"."Salesperson Code");
                    IF Saleperson.FINDFIRST THEN
                        SalesName := Saleperson.Name;

                    InvoiceNo := '';
                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.", "Sales Cr.Memo Header"."Applies-to Doc. No.");
                    IF recSIH.FINDFIRST THEN BEGIN
                        InvoiceDate := recSIH."Posting Date";
                        InvoiceNo := recSIH."No.";
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

                    //<<1
                end;

                trigger OnPostDataItem()
                begin

                    //>>10Mar2017 SalesReturnTotal

                    //Sales Cr.Memo Header, Footer (2) - OnPreSection()
                    IF PrintToExcel AND SHowInvTotal AND SalesCreditMemo THEN BEGIN
                        MakeExcelCrmemoGrandTotal;
                    END;

                    //<<10Mar2017 SalesReturnTotal
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
                        "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Shortcut Dimension 1 Code",
                                                        "Sales Invoice Header"."Shortcut Dimension 1 Code");
                    END;

                    IF "Sales Cr.Memo Header"."Shortcut Dimension 2 Code" <> "Resp.CentreFilter" THEN
                        CurrReport.SKIP;
                    //<<1


                    //>>2

                    //Sales Cr.Memo Header, Header (1) - OnPreSection()
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
                }
                field("End Date"; EndDate)
                {
                }
                field("Print To Excel"; PrintToExcel)
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
        *///Commented 10Mar2017
        IF StartDate = 0D THEN
            ERROR('Please enter the From Date and To Date');
        IF EndDate = 0D THEN
            ERROR('Please enter the From Date and To Date');

        "Company info".GET;

        IF PrintToExcel THEN
            MakeExcelInfo;
        //<1

    end;

    var
        StartDate: Date;
        EndDate: Date;
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        "Company info": Record 79;
        BEDPercent: Decimal;
        //ExciseSetup: Record "13711";
        // StructureDetailes: Record "13798";
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
        // RG23D: Record "16537";
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
        ShipmentAgentName: Text[50];
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
        VLE: Record 5802;
        VLEcost: Decimal;
        TotalVLEcost: Decimal;
        ILEentryNo: Integer;
        ILE: Record 32;
        VLE1: Record 5802;
        VLEcost1: Decimal;
        ILEentryNo1: Integer;
        ILE1: Record 32;
        TotalVLEcost1: Decimal;
        Mfgdte: Date;
        ILErec3: Record 32;
        recItemApplEntry: Record 339;
        BlendOrderNo: Code[20];
        recILEtable: Record 32;
        TotalCost: Decimal;
        UNITCOST: Decimal;
        ChargeofItem: Decimal;
        RateofItem: Decimal;
        ValuedCost: Decimal;
        ILErecEBT: Record 32;
        Revaluedcost: Decimal;
        UNITCOST1: Decimal;
        TotalCost1: Decimal;
        TotalQty1: Decimal;
        UNITCOSTNEW: Decimal;
        VLEQTY: Decimal;
        VLEcost2: Decimal;
        Text002: Label 'Data';
        Text003: Label 'Sales Register Costing';
        Text004: Label 'Company Name';
        Text005: Label 'Report No.';
        Text006: Label 'Report Name';
        Text007: Label 'User ID';
        Text008: Label 'Date';
        Text009: Label 'Location Filter';
        "--10Mar2017": Integer;
        SILCOUNT: Integer;
        TSILCOUNT: Integer;
        SCLCOUNT: Integer;
        TSCLCOUNT: Integer;
        GINQty: Decimal;
        TINQty: Decimal;
        GINQtyBase: Decimal;
        TINQtyBase: Decimal;
        GINRAmt: Decimal;
        TINRAmt: Decimal;
        GFOCAmt: Decimal;
        TFOCAmt: Decimal;
        "--CrMemo----": Integer;
        CrGINQty: Decimal;
        CrTINQty: Decimal;
        CrGINQtyBase: Decimal;
        CrTINQtyBase: Decimal;
        CrGINRAmt: Decimal;
        CrTINRAmt: Decimal;
        CrGFOCAmt: Decimal;
        CrTFOCAmt: Decimal;
        "------06Nov2017": Integer;
        NNNNNN: Integer;

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//09Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(FORMAT(Text003), FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Sales Register Costing", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.NewRow;
        ExcelBuf.ClearNewRow;
    end;

    //  [Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text002,Text003,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>09Mar2017 RB-N
        /*  ExcelBuf.CreateBook('', Text003);
         ExcelBuf.CreateBookAndOpenExcel('', Text003, '', '', USERID);
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text003);
        ExcelBuf.WriteSheet(Text003, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text003, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<
    end;

    //   [Scope('Internal')]
    procedure MakeExcelInvTotalGroupFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Line"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(SalesType, FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(''/*"Sales Invoice Line"."Form Code"*/, FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(LocationRec."State Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Header"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(ResCenName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8 //10Mar2017
        ExcelBuf.AddColumn("Sales Invoice Line"."Shortcut Dimension 1 Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Invoice Line"."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn(Cust."Full Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Sales Invoice Header"."Ship-to Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        //ExcelBuf.AddColumn(TotalQty,FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//19
        ExcelBuf.AddColumn(GINQty, FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//19 10Mar2017
        GINQty := 0;//10Mar2017

        //ExcelBuf.AddColumn(TotalQtyBase,FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//20
        ExcelBuf.AddColumn(GINQtyBase, FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//20 10Mar2017
        GINQtyBase := 0;//10Mar2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//22
        /*
        IF "Sales Invoice Header"."Currency Code" = '' THEN
        BEGIN
          ExcelBuf.AddColumn(INRTot2,FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//23
        END
        ELSE
        BEGIN
          ExcelBuf.AddColumn(INRTot2,FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//24
        END;
        *///Commented 10Mar2017
        ExcelBuf.AddColumn(GINRAmt, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//23 //10Mar2017
        GINRAmt := 0; //10Mar2017

        //ExcelBuf.AddColumn(ABS(TotalFOCAmount),FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//24
        ExcelBuf.AddColumn(ABS(GFOCAmt), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//24 //10Mar2017
        GFOCAmt := 0; //10Mar2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//35

        GINRtot += INRTot;
        GINRtot1 += INRTot1;
        GINRtot2 += INRTot2;

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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35 //10Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
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
        //ExcelBuf.AddColumn(GtotalQty,FALSE,'',TRUE,FALSE,TRUE,'####0.000',0);//19
        ExcelBuf.AddColumn(TINQty, FALSE, '', TRUE, FALSE, TRUE, '####0.000', 0);//19 10Mar2017
        //ExcelBuf.AddColumn(GtotalQtyBase,FALSE,'',TRUE,FALSE,TRUE,'####0.000',0);//20
        ExcelBuf.AddColumn(TINQtyBase, FALSE, '', TRUE, FALSE, TRUE, '####0.000', 0);//20 10Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        //ExcelBuf.AddColumn(GINRtot2,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//23
        ExcelBuf.AddColumn(TINRAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//23 //10Mar2017
        //ExcelBuf.AddColumn(ABS(GTotalFOCAmount),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//24
        ExcelBuf.AddColumn(ABS(TFOCAmt), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//24
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35 //10Mar2017
    end;

    //  [Scope('Internal')]
    procedure MakeCrMemoBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(SalesType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(LocationRec."State Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(ResCenName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn(InvoiceNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn(Cust."Full Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Ship-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn("Sales Cr.Memo Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//18
        ExcelBuf.AddColumn(-"Sales Cr.Memo Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//19
        //>>10Mar2017 TotalQty
        CrGINQty += "Sales Cr.Memo Line".Quantity;
        CrTINQty += "Sales Cr.Memo Line".Quantity;
        //<<10Mar2017 TotalQty

        ExcelBuf.AddColumn(-"Sales Cr.Memo Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//20
        //>>10Mar2017 TotalQtyBase
        CrGINQtyBase += "Sales Cr.Memo Line"."Quantity (Base)";
        CrTINQtyBase += "Sales Cr.Memo Line"."Quantity (Base)";
        //<<10Mar2017 TotalQtyBase

        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."MRP Price"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//21

        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(ListPricePerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//22
            ExcelBuf.AddColumn(-"Sales Cr.Memo Line".Amount, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23
                                                                                                             //>>10Mar2017 INRAmt
            CrGINRAmt += "Sales Cr.Memo Line".Amount;
            CrTINRAmt += "Sales Cr.Memo Line".Amount;
            //<<10Mar2017 INRAmt

        END
        ELSE BEGIN
            ExcelBuf.AddColumn((ListPricePerlt / "Sales Cr.Memo Header"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//22
            ExcelBuf.AddColumn(-("Sales Cr.Memo Line".Amount / "Sales Cr.Memo Header"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23
                                                                                                                                                          //>>10Mar2017 INRAmt
            CrGINRAmt += "Sales Cr.Memo Line".Amount / "Sales Cr.Memo Header"."Currency Factor";
            CrTINRAmt += "Sales Cr.Memo Line".Amount / "Sales Cr.Memo Header"."Currency Factor";
            //<<10Mar2017 INRAmt

        END;

        ExcelBuf.AddColumn(ABS(FOCAmount), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//24
        //>>10Mar2017 FOCAmt
        CrGFOCAmt += FOCAmount;
        CrTFOCAmt += FOCAmount;
        //<<10Mar2017 FOCAmt

        ExcelBuf.AddColumn(((VLEcost1) / ("Sales Cr.Memo Line".Quantity) /
        "Sales Cr.Memo Line"."Qty. per Unit of Measure"), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);  //26
        ExcelBuf.AddColumn(-VLEcost1, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27
        ExcelBuf.AddColumn(-(ListPricePerlt - ((VLEcost1) / ("Sales Cr.Memo Line".Quantity) /
        "Sales Cr.Memo Line"."Qty. per Unit of Measure")), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//28

        //ExcelBuf.AddColumn(-((ListPricePerlt - ((VLEcost1)/ ("Sales Cr.Memo Line".Quantity)/
        //"Sales Cr.Memo Line"."Qty. per Unit of Measure"))/ListPricePerlt),FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(("Sales Cr.Memo Line".Amount + FOCAmount - VLEcost1) /
                                  "Sales Cr.Memo Line".Amount + FOCAmount, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//29

        //ExcelBuf.AddColumn(-((ListPricePerlt - ((VLEcost1)/ ("Sales Cr.Memo Line".Quantity)/
        //"Sales Cr.Memo Line"."Qty. per Unit of Measure"))*"Sales Cr.Memo Line"."Quantity (Base)"),FALSE,'',FALSE,FALSE,FALSE,'');
        //

        ExcelBuf.AddColumn("Sales Cr.Memo Line".Amount + FOCAmount - VLEcost1, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//30

        //
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn(Mfgdte, FALSE, '', FALSE, FALSE, FALSE, '', 2);//32
        ExcelBuf.AddColumn(BlendOrderNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//34
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//35 //10Mar2017
        //ExcelBuf.AddColumn("Sales Cr.Memo Header"."Customer Posting Group",FALSE,'',TRUE,FALSE,TRUE,'@',1);
        //RSPl-PARAG 02.01.2016
    end;

    //  [Scope('Internal')]
    procedure MakecrmemoHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Detailed Sales Register', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19 // EBT MILAN
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35 //10Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Resp. Centre Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Resp. Centre Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
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
        ExcelBuf.AddColumn('MRP Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('Billing Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('FOC', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Cost per Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('Revalued Cost', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 // EBT MILAN
        ExcelBuf.AddColumn('Total Cost', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('GP per Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('GP %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('Total GP', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('Mfg Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('Blend Order No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33
        ExcelBuf.AddColumn('Supp. Org Inv', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn('Customer Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35
        //RSPl-PARAG 02.01.2016
    end;

    //  [Scope('Internal')]
    procedure MakeExcelCrMemoGrouping()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(SalesType, FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Form Code"*/'', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(LocationRec."State Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(ResCenName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Shortcut Dimension 1 Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn(Cust."Full Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Ship-to Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        //ExcelBuf.AddColumn(-TotalQty,FALSE,'',TRUE,FALSE,FALSE,'####0.000',0);//19
        ExcelBuf.AddColumn(-CrGINQty, FALSE, '', TRUE, FALSE, FALSE, '####0.000', 0);//19 //10Mar2017
        CrGINQty := 0; //10Mar2017

        //ExcelBuf.AddColumn(-TotalQtyBase,FALSE,'',TRUE,FALSE,FALSE,'####0.000',0);//20
        ExcelBuf.AddColumn(-CrGINQtyBase, FALSE, '', TRUE, FALSE, FALSE, '####0.000', 0);//20 //10Mar2017
        CrGINQtyBase := 0;//10Mar2017

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//22

        ExcelBuf.AddColumn(-CrGINRAmt, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//23 //10Mar2017
        CrGINRAmt := 0;//10Mar2017

        ExcelBuf.AddColumn(ABS(CrGFOCAmt), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//24 //10Mar2017
        CrGFOCAmt := 0;//10Mar2017

        /*
        IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
        BEGIN
        //ExcelBuf.AddColumn(-TotalAmount,FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//23
        ExcelBuf.AddColumn(-CrGINRAmt,FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//23 //10Mar2017
        CrGINRAmt := 0;//10Mar2017
        
        GINRtotc += TotalAmount;
        //ExcelBuf.AddColumn(ABS(TotalFOCAmount),FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//24
        ExcelBuf.AddColumn(ABS(CrGFOCAmt),FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//24 //10Mar2017
        CrGFOCAmt := 0;//10Mar2017
        END
        ELSE
        BEGIN
        //ExcelBuf.AddColumn(-TotalAmount / "Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//23
        ExcelBuf.AddColumn(-CrGINRAmt / "Sales Cr.Memo Header"."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//23
        CrGINRAmt := 0;//10Mar2017
        
        GINRtotc += TotalAmount / "Sales Cr.Memo Header"."Currency Factor";
        //ExcelBuf.AddColumn(ABS(TotalFOCAmount) / "Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//24
        ExcelBuf.AddColumn(ABS(CrGINRAmt) / "Sales Cr.Memo Header"."Currency Factor",FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//24 //10Mar2017
        CrGINRAmt := 0;//10Mar2017
        END;
        *///commented

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//30 // EBT MILAN
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//35

    end;

    //  [Scope('Internal')]
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26  // EBT MILAN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35 //10Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
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
        //ExcelBuf.AddColumn(-GtotalQty1,FALSE,'',TRUE,FALSE,TRUE,'####0.000',0);//19
        ExcelBuf.AddColumn(-CrTINQty, FALSE, '', TRUE, FALSE, TRUE, '####0.000', 0);//19 //10Mar2017
        //ExcelBuf.AddColumn(-GtotalQtyBase1,FALSE,'',TRUE,FALSE,TRUE,'####0.000',0);//20
        ExcelBuf.AddColumn(-CrTINQtyBase, FALSE, '', TRUE, FALSE, TRUE, '####0.000', 0);//20 //10Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        //ExcelBuf.AddColumn(GINRtotc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//23
        ExcelBuf.AddColumn(-CrTINRAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//23 //10Mar2017

        //ExcelBuf.AddColumn(GTotalFOCAmount1,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//24
        ExcelBuf.AddColumn(CrTFOCAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//24 //10Mar2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26 // EBT MILAN
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35 //10Mar2017
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn(SalesType, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(LocationRec."State Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Line"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(LocationRec.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(CustResCenCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(ResCenName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn("Sales Invoice Line"."Shortcut Dimension 1 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Invoice Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn(Cust."Full Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Sales Invoice Header"."Ship-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14

        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
            "recItem-Supp".GET("Sales Invoice Line"."Final Item No.");
            ExcelBuf.AddColumn("Sales Invoice Line"."Final Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
            ExcelBuf.AddColumn("recItem-Supp".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        END ELSE BEGIN
            ExcelBuf.AddColumn("Sales Invoice Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
            ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        END;

        ExcelBuf.AddColumn("Sales Invoice Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn("Sales Invoice Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//18
        ExcelBuf.AddColumn("Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//19

        //>>10Mar2017 UoMQty
        GINQty += "Sales Invoice Line".Quantity;
        TINQty += "Sales Invoice Line".Quantity;
        //<<10Mar2017 UoMQty

        ExcelBuf.AddColumn("Sales Invoice Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//20

        //>>10Mar2017 QtyBase
        GINQtyBase += "Sales Invoice Line"."Quantity (Base)";
        TINQtyBase += "Sales Invoice Line"."Quantity (Base)";
        //<<10Mar2017 QtyBase

        ExcelBuf.AddColumn(/*"Sales Invoice Line"."MRP Price"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//21

        IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
            //ExcelBuf.AddColumn(ROUND((UPricePerLT),0.01),FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//22
            ExcelBuf.AddColumn(UPricePerLT, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//22 //06Nov2017
            INRTot := ROUND((UPricePerLT), 0.01);
            INRTot1 := UPrice;
            IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
                ExcelBuf.AddColumn("Sales Invoice Line".Amount, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23
                INRTot2 := "Sales Invoice Line".Amount;
                //>>10Mar2017 INRAmount
                GINRAmt += "Sales Invoice Line".Amount;
                TINRAmt += "Sales Invoice Line".Amount;
                //<<10Mar2017 INRAmount

            END ELSE BEGIN
                ExcelBuf.AddColumn(UPrice * "Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23
                INRTot2 := UPrice * "Sales Invoice Line".Quantity;
                //>>10Mar2017 INRAmount
                GINRAmt += UPrice * "Sales Invoice Line".Quantity;
                TINRAmt += UPrice * "Sales Invoice Line".Quantity;
                //<<10Mar2017 INRAmount

            END;
        END ELSE BEGIN
            ExcelBuf.AddColumn((UPricePerLT / "Sales Invoice Header"."Currency Factor"), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//22
            INRTot := UPricePerLT;
            INRTot1 := UPrice / "Sales Invoice Header"."Currency Factor";
            ExcelBuf.AddColumn(UPrice * "Sales Invoice Line".Quantity / "Sales Invoice Header"."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23

            INRTot2 := UPrice * "Sales Invoice Line".Quantity / "Sales Invoice Header"."Currency Factor";
            //>>10Mar2017 INRAmount
            GINRAmt += UPrice * "Sales Invoice Line".Quantity / "Sales Invoice Header"."Currency Factor";
            TINRAmt += UPrice * "Sales Invoice Line".Quantity / "Sales Invoice Header"."Currency Factor";
            //<<10Mar2017 INRAmount

        END;

        //ExcelBuf.AddColumn(-ABS(FOCAmount),FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//24
        ExcelBuf.AddColumn(ABS(FOCAmount), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//24
        //>>10Mar2017 FOCAmt
        GFOCAmt += FOCAmount;
        TFOCAmt += FOCAmount;
        //<<10Mar2017 FOCAmt

        IF "Sales Invoice Header"."Supplimentary Invoice" = FALSE THEN BEGIN

            IF UNITCOST1 <> 0 THEN
                ExcelBuf.AddColumn(ROUND((UNITCOST), 0.00001), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0)//25
            ELSE
                ExcelBuf.AddColumn(-((VLEcost) / ("Sales Invoice Line".Quantity) /
                "Sales Invoice Line"."Qty. per Unit of Measure"), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25

            // ExcelBuf.AddColumn(-ROUND((-UNITCOST+ChargeofItem),0.00001),FALSE,'',FALSE,FALSE,FALSE,'');

            //Revaluedcost := -ROUND((UNITCOST+ChargeofItem)+(((VLEcost) / ("Sales Invoice Line".Quantity)/
            //  "Sales Invoice Line"."Qty. per Unit of Measure")),0.00001);

            Revaluedcost := ROUND(UNITCOSTNEW + ChargeofItem, 0.00001);

            IF UNITCOST1 <> 0 THEN
                ExcelBuf.AddColumn(UNITCOSTNEW, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0)//26
            ELSE
                ExcelBuf.AddColumn('0', FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//26

            IF (UNITCOST1 + ChargeofItem) = 0 THEN BEGIN
                ExcelBuf.AddColumn(-VLEcost, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27
                                                                                             /*
                                                                                             ExcelBuf.AddColumn(ABS((ROUND((UPricePerLT),0.01)+((VLEcost)/ ("Sales Invoice Line".Quantity)/
                                                                                             "Sales Invoice Line"."Qty. per Unit of Measure"))),FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//28
                                                                                             */

                ExcelBuf.AddColumn(ABS((((UPricePerLT)) + ((VLEcost) / ("Sales Invoice Line".Quantity) /
                    "Sales Invoice Line"."Qty. per Unit of Measure"))), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//28 //06Nov2017

                IF UPricePerLT <> 0 THEN
                    /*
                    ExcelBuf.AddColumn(ABS((ROUND((UPricePerLT),0.01)+((VLEcost)/ ("Sales Invoice Line".Quantity)/
                    "Sales Invoice Line"."Qty. per Unit of Measure"))/(ROUND((UPricePerLT),0.01))),FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0)//29
                    */

              ExcelBuf.AddColumn(ABS(((UPricePerLT) + ((VLEcost) / ("Sales Invoice Line".Quantity) /
                  "Sales Invoice Line"."Qty. per Unit of Measure")) / (UPricePerLT)), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//29 06Nov2017
                ELSE
                    ExcelBuf.AddColumn(ABS((((UPricePerLT)) + ((VLEcost) / ("Sales Invoice Line".Quantity) /
                        "Sales Invoice Line"."Qty. per Unit of Measure"))), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//29


                ExcelBuf.AddColumn(ABS("Sales Invoice Line".Amount + VLEcost), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//30
            END ELSE BEGIN

                //  ExcelBuf.AddColumn(-ValuedCost,FALSE,'',FALSE,FALSE,FALSE,'');  //
                ExcelBuf.AddColumn(((Revaluedcost) * "Sales Invoice Line"."Quantity (Base)"), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27 //

                //  ExcelBuf.AddColumn((ROUND((UPricePerLT),0.01)+((ValuedCost)/("Sales Invoice Line".Quantity)/
                //  "Sales Invoice Line"."Qty. per Unit of Measure")),FALSE,'',FALSE,FALSE,FALSE,'');

                //ExcelBuf.AddColumn((ROUND(UPricePerLT,0.01)-(Revaluedcost)),FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//28
                ExcelBuf.AddColumn((UPricePerLT - Revaluedcost), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//28 06Nov2017

                //  ExcelBuf.AddColumn((ROUND((UPricePerLT),0.01)+((ValuedCost)/("Sales Invoice Line".Quantity)/
                //  "Sales Invoice Line"."Qty. per Unit of Measure"))/(ROUND((UPricePerLT),0.01)),FALSE,'',FALSE,FALSE,FALSE,'');
                /*
                ExcelBuf.AddColumn((ROUND((UPricePerLT),0.01)-(Revaluedcost))
                /(ROUND((UPricePerLT),0.01)),FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//29
                */

                ExcelBuf.AddColumn(((UPricePerLT) - (Revaluedcost)) / ((UPricePerLT)), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29 //06Nov2017

                //  ExcelBuf.AddColumn("Sales Invoice Line".Amount - ValuedCost,FALSE,'',FALSE,FALSE,FALSE,'');  //

                ExcelBuf.AddColumn("Sales Invoice Line".Amount - (Revaluedcost *
                                                     "Sales Invoice Line"."Quantity (Base)"), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//30  //
            END;
        END ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//25
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//26
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//27
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//28
            ExcelBuf.AddColumn("Sales Invoice Line".Amount / "Sales Invoice Line".Amount, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//
            ExcelBuf.AddColumn("Sales Invoice Line".Amount, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//

            // ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
        END;

        ExcelBuf.AddColumn("Sales Invoice Line"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn(Mfgdte, FALSE, '', FALSE, FALSE, FALSE, '', 2);//32
        ExcelBuf.AddColumn(BlendOrderNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//33
        IF "Sales Invoice Header"."Supplimentary Invoice" = TRUE THEN BEGIN
            ExcelBuf.AddColumn("Sales Invoice Line"."Appiles to Inv.No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//34
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//34
        //ExcelBuf.AddColumn(UNITCOST,FALSE,'',FALSE,FALSE,FALSE,'');
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn("Sales Invoice Header"."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//35
        //ExcelBuf.AddColumn("Sales Invoice Header"."Customer Posting Group",FALSE,'',TRUE,FALSE,TRUE,'@',1);
        //RSPl-PARAG 02.01.2016

    end;

    //   [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>10Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//34
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//35

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text003, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//34
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//35
        ExcelBuf.NewRow;

        //<<

        //ExcelBuf.NewRow;
        //ExcelBuf.AddColumn("Company info".Name,FALSE,'',TRUE,FALSE,FALSE,'@',1);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Detailed Sales Register', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35 //10Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('Type', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('State Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('Resp. Centre Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('Resp. Centre Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('Cancelled Invoice', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('Customer Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('Customer Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('Ship to Customer Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('Unit of Measure', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('Packing Qty per UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('UOM Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('Total Quantity Invoiced', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('MRP Price', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('Billing Price', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('INR Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('FOC', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('Cost per Ltr', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('Revalued Cost', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('Total Cost', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('GP per Ltr', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('GP %', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('Total GP', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('Mfg Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('Blend Order No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('Supp. Org Inv', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        //RSPl-PARAG 02.01.2016
        ExcelBuf.AddColumn('Customer Posting Group', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        //RSPl-PARAG 02.01.2016
    end;
}


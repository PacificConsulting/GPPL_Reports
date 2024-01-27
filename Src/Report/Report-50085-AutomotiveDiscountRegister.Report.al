report 50085 "Automotive Discount Register"
{
    //     // 
    //     // Date        Version      Remarks
    //     // .....................................................................................
    //     // 15Nov2018   RB-N         National Discount Scheme & Regional Discount Scheme
    //     DefaultLayout = RDLC;
    //     RDLCLayout = './AutomotiveDiscountRegister.rdlc';


    //     dataset
    //     {
    //         dataitem("Sales Invoice Header"; "Sales Invoice Header")
    //         {
    //             DataItemTableView = SORTING("No.");
    //             RequestFilterFields = "Posting Date", "Location Code";
    //             dataitem("Sales Invoice Line"; "Sales Invoice Line")
    //             {
    //                 DataItemLink = "Document No." = FIELD("No.");
    //                 DataItemTableView = SORTING(Type, "Inventory Posting Group", "No.")
    //                                     WHERE(Type = CONST(Item),
    //                                           "Inventory Posting Group" = FILTER('AUTOOILS | REPSOL'));
    //                 dataitem(DataItem1000000002; 13798)
    //                 {
    //                     DataItemLink = "Invoice No." = FIELD("Document No."),
    //                                    "Item No." = FIELD("No."),
    //                                    "Line No." = FIELD("Line No.");
    //                     DataItemTableView = SORTING("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group")
    //                                         WHERE("Tax/Charge Type" = CONST(Charges));

    //                     trigger OnAfterGetRecord()
    //                     begin

    //                         //>>1

    //                         ItemRec.GET("Posted Str Order Line Details"."Item No.");
    //                         SpclFOC := 0;
    //                         SpclFOCPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'FOC DIS');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             SpclFOCPerlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             SpclFOC := recPostedStrLineDet.Amount;
    //                         END;

    //                         StateDisc := 0;
    //                         StateDiscPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'STATE DISC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             StateDisc := recPostedStrLineDet.Amount;
    //                             StateDiscPerlt := recPostedStrLineDet.Amount / "Sales Invoice Line"."Quantity (Base)";
    //                         END;

    //                         Spot := 0;
    //                         SpotPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'SPOT-DISC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             SpotPerlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             Spot := recPostedStrLineDet.Amount;
    //                         END;

    //                         AVPSanction := 0;
    //                         AVPSanctionPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'AVP-SP-SAN');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             AVPSanctionPerlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             AVPSanction := recPostedStrLineDet.Amount;
    //                         END;

    //                         NewPrice := 0;
    //                         NewPricePerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'NEW-PRI-SP');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             NewPricePerlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             NewPrice := recPostedStrLineDet.Amount;
    //                         END;


    //                         SpecDisc := 0;
    //                         SpecDiscPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'T2 DISC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             SpecDiscPerlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             SpecDisc := recPostedStrLineDet.Amount;
    //                         END;

    //                         FOC := 0;
    //                         FOCperlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'FOC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             FOCperlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             FOC := recPostedStrLineDet.Amount;
    //                         END;

    //                         MntDisc := 0;
    //                         MntDiscPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'MONTH-DIS');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             MntDiscPerlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             MntDisc := recPostedStrLineDet.Amount;
    //                         END;

    //                         AddtlDisc := 0;
    //                         AddtlDiscPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'ADD-DIS');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             AddtlDiscPerlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             AddtlDisc := recPostedStrLineDet.Amount;
    //                         END;

    //                         VATDiff := 0;
    //                         VATDiffPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'VAT-DIF-DI');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             VATDiffPerlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             VATDiff := recPostedStrLineDet.Amount;
    //                         END;

    //                         //>>15Nov2018
    //                         NATDisc := 0;
    //                         NATDiscperlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'NATSCHEME');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             NATDiscperlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             NATDisc := recPostedStrLineDet.Amount;
    //                         END;

    //                         RegDisc := 0;
    //                         RegDiscperlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Invoice Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'REGSCHEME');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             RegDiscperlt := recPostedStrLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
    //                             RegDisc := recPostedStrLineDet.Amount;
    //                         END;
    //                         //<<15Nov2018

    //                         //>>02Jun2019
    //                         //>>02Jun2019
    //                         IF "Sales Invoice Header"."Posting Date" <= 030319D THEN BEGIN
    //                             NATDisc += "Sales Invoice Line"."National Discount Amount";
    //                             NATDiscperlt += "Sales Invoice Line"."National Discount Per Ltr";
    //                         END;
    //                         //<<02Jun2019

    //                         ProdDisc := 0;
    //                         ProdDiscPerLt := 0;
    //                         MRpMstr.RESET;
    //                         MRpMstr.SETRANGE(MRpMstr."Item No.", "Sales Invoice Line"."No.");
    //                         MRpMstr.SETRANGE(MRpMstr."Lot No.", "Sales Invoice Line"."Lot No.");
    //                         IF MRpMstr.FINDFIRST THEN BEGIN
    //                             //ProdDiscPerLt := MRpMstr."National Discount";
    //                             PriceSupportPerLt := MRpMstr."Price Support";
    //                             //ProdDisc := ProdDiscPerLt * "Sales Invoice Line"."Quantity (Base)";
    //                             PriceSupport := MRpMstr."Price Support" * "Sales Invoice Line"."Quantity (Base)";
    //                         END;
    //                         //<<1

    //                         ProdDiscPerLt := "Sales Invoice Line"."National Discount Per Ltr";//02Jun2019
    //                         ProdDisc := "Sales Invoice Line"."National Discount Amount";//02Jun2019

    //                         //>>27Apr2017 GroupData
    //                         TINVCOUNT += 1;

    //                         STR25Apr.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
    //                         STR25Apr.SETRANGE("Line No.", "Line No.");
    //                         IF STR25Apr.FINDLAST THEN BEGIN
    //                             IINVCOUNT := STR25Apr.COUNT;

    //                         END;

    //                         IF (TINVCOUNT = IINVCOUNT) AND PrintToExcel THEN BEGIN

    //                             TotalDisc := SpclFOC + StateDisc + Spot + AVPSanction + NewPrice + SpecDisc + FOC + MntDisc + AddtlDisc + VATDiff + NATDisc + RegDisc;
    //                             TotalDiscPerlt := SpclFOCPerlt + StateDiscPerlt + SpotPerlt + AVPSanctionPerlt + NewPricePerlt + SpecDiscPerlt + FOCperlt + MntDiscPerlt
    //                                             + AddtlDiscPerlt + VATDiffPerlt + NATDiscperlt + RegDiscperlt;
    //                             TotListPrice += ListPrice + ProdDiscPerLt + PriceSupportPerLt;

    //                             MakeExcelInvoiceGroupfooter;

    //                             TINVCOUNT := 0;
    //                             TotalDisc := 0;

    //                             TSpclFOC += SpclFOC;//25Apr2017
    //                             TStateDisc += StateDisc;//25Apr2017
    //                             TSpot += Spot;//25Apr2017
    //                             TAVPSanction += AVPSanction;//25Apr2017
    //                             TPriceSupport += PriceSupport;//27Apr2017
    //                             TNewPrice += NewPrice;//25Apr2017
    //                             TSpecDisc += SpecDisc;//25Apr2017
    //                             TFOC += FOC;//25Apr2017
    //                             TMntDisc += MntDisc;//25Apr2017
    //                             TAddtlDisc += AddtlDisc;//25Apr2017
    //                             TVATDiff += VATDiff;//25Apr2017
    //                             TProdDisc += ProdDisc;//25Apr2017
    //                             TListPrice += ListPrice;//27Apr2017
    //                             RegDiscTotal += RegDisc;//15Nov2018
    //                             NATDiscTotal += NATDisc;//15Nov2018
    //                         END;
    //                         //<<27Apr2017 GroupData
    //                     end;

    //                     trigger OnPreDataItem()
    //                     begin

    //                         STR25Apr.COPYFILTERS("Posted Str Order Line Details");//27Apr2017
    //                         TINVCOUNT := 0; //27Apr2017
    //                     end;
    //                 }

    //                 trigger OnAfterGetRecord()
    //                 begin

    //                     //>>1

    //                     //ListPrice := "Sales Invoice Line"."Unit Price Incl. of Tax"/"Sales Invoice Line"."Qty. per Unit of Measure";
    //                     ListPrice := "Sales Invoice Line"."List Price";//02Jun2019
    //                     TotQty += "Sales Invoice Line".Quantity;
    //                     TotQtyBase += "Sales Invoice Line"."Quantity (Base)";
    //                     //<<1
    //                 end;

    //                 trigger OnPreDataItem()
    //                 begin

    //                     CurrReport.CREATETOTALS(FOC, SpclFOC, Spot, NewPrice, AVPSanction, AddtlDisc, MntDisc, VATDiff, CD, OtherDisc);
    //                     CurrReport.CREATETOTALS(StateDisc, SpecDisc, ProdDisc, PriceSupport,
    //                     ProdDiscPerLt, PriceSupportPerLt, TotalDisc, ProdDisc, TotalDiscPerlt);

    //                     ListPrice := 0;
    //                 end;
    //             }

    //             trigger OnAfterGetRecord()
    //             begin

    //                 //>>1
    //                 //TotalAmount := 0;//28Apr2017
    //                 CashDiscountAmt := 0;
    //                 recPostedStrOrdDetails.RESET;
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'CASHDISC');
    //                 IF recPostedStrOrdDetails.FINDFIRST THEN
    //                     REPEAT
    //                         CashDiscountAmt += recPostedStrOrdDetails.Amount;
    //                     UNTIL recPostedStrOrdDetails.NEXT = 0;

    //                 IF CashDiscountAmt <> 0 THEN
    //                     CashDiscText := 'Cash Discount @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%'
    //                 ELSE
    //                     CashDiscText := '';


    //                 CCFC := 0;
    //                 recPostedStrOrdDetails.RESET;
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'CCFC');
    //                 IF recPostedStrOrdDetails.FINDFIRST THEN
    //                     REPEAT
    //                         CCFC += recPostedStrOrdDetails.Amount;
    //                     UNTIL recPostedStrOrdDetails.NEXT = 0;

    //                 IF CCFC <> 0 THEN
    //                     CCFCtext := 'CCFC @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%'
    //                 ELSE
    //                     CCFCtext := '';

    //                 V_TransSub := 0;
    //                 recPostedStrOrdDetails.RESET;
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'TPTSUBSIDY');
    //                 IF recPostedStrOrdDetails.FINDFIRST THEN
    //                     REPEAT
    //                         V_TransSub += recPostedStrOrdDetails.Amount;
    //                     UNTIL recPostedStrOrdDetails.NEXT = 0;

    //                 IF V_TransSub <> 0 THEN
    //                     V_TransSubTxt := 'Transport Subsidy'
    //                 ELSE
    //                     V_TransSubTxt := '';
    //                 //<<1
    //             end;

    //             trigger OnPostDataItem()
    //             begin

    //                 IF "Sales Invoice" AND PrintToExcel THEN
    //                     MakeExcelInvoiceGrandTotal;//27Apr2017
    //             end;

    //             trigger OnPreDataItem()
    //             begin
    //                 NATDiscTotal := 0;
    //                 RegDiscTotal := 0;
    //                 //>>1
    //                 //CurrReport.CREATETOTALS(FOC,SpclFOC,Spot,NewPrice,AVPSanction,AddtlDisc,MntDisc,VATDiff,CD,OtherDisc);
    //                 //CurrReport.CREATETOTALS(StateDisc,SpecDisc,ProdDisc,ProdDiscPerLt,TotalDisc,ProdDisc,TotalDiscPerlt,TotalAmount);

    //                 StartDate := "Sales Invoice Header".GETRANGEMIN("Sales Invoice Header"."Posting Date");
    //                 EndDate := "Sales Invoice Header".GETRANGEMAX("Sales Invoice Header"."Posting Date");
    //                 //<<1

    //                 IF "Sales Invoice" AND PrintToExcel THEN
    //                     MakeExcelDataHeader;//27Apr2017
    //             end;
    //         }
    //         dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
    //         {
    //             DataItemTableView = SORTING("No.");
    //             dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
    //             {
    //                 DataItemLink = "Document No." = FIELD("No.");
    //                 DataItemTableView = SORTING(Type, "Document No.", "Line No.")
    //                                     WHERE(Type = CONST(Item),
    //                                           "Item Category Code" = FILTER('CAT02'));
    //                 dataitem("<Posted Str Ordr Ln Detail1>"; 13798)
    //                 {
    //                     DataItemLink = "Invoice No." = FIELD("Document No."),
    //                                    "Item No." = FIELD("No."),
    //                                    "Line No." = FIELD("Line No.");
    //                     DataItemTableView = SORTING("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group")
    //                                         WHERE("Tax/Charge Type" = CONST(Charges));

    //                     trigger OnAfterGetRecord()
    //                     begin

    //                         //>>1

    //                         ItemRec.GET("<Posted Str Ordr Ln Detail1>"."Item No.");
    //                         ;
    //                         CrSpclFOC := 0;
    //                         CrSpclFOCPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'FOC DIS');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrSpclFOCPerlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrSpclFOC := recPostedStrLineDet.Amount;
    //                         END;

    //                         CrStateDisc := 0;
    //                         CrStateDiscPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'STATE DISC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrStateDisc := recPostedStrLineDet.Amount;
    //                             CrStateDiscPerlt := recPostedStrLineDet.Amount / "Sales Cr.Memo Line"."Quantity (Base)";
    //                         END;

    //                         CrSpot := 0;
    //                         CrSpotPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'SPOT-DISC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrSpotPerlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrSpot := recPostedStrLineDet.Amount;
    //                         END;

    //                         CrAVPSanction := 0;
    //                         CrAVPSanctionPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'AVP-SP-SAN');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrAVPSanctionPerlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrAVPSanction := recPostedStrLineDet.Amount;
    //                         END;

    //                         CrNewPrice := 0;
    //                         CrNewPricePerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'NEW-PRI-SP');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrNewPricePerlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrNewPrice := recPostedStrLineDet.Amount;
    //                         END;


    //                         CrFOC := 0;
    //                         CrFOCperlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'FOC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrFOCperlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrFOC := recPostedStrLineDet.Amount;
    //                         END;

    //                         CrCD := 0;
    //                         CrCDPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'CASHDISC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrCD := recPostedStrLineDet.Amount;
    //                             CrCDPerlt := recPostedStrLineDet.Amount / "Sales Cr.Memo Line"."Quantity (Base)";
    //                         END;

    //                         CrSpecDisc := 0;
    //                         CrSpecDiscPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'T2 DISC');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrSpecDiscPerlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrSpecDisc := recPostedStrLineDet.Amount;
    //                         END;

    //                         CrMntDisc := 0;
    //                         CrMntDiscPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'MONTH-DIS');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrMntDiscPerlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrMntDisc := recPostedStrLineDet.Amount;
    //                         END;

    //                         CrAddtlDisc := 0;
    //                         CrAddtlDiscPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'ADD-DIS');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrAddtlDiscPerlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrAddtlDisc := recPostedStrLineDet.Amount;
    //                         END;

    //                         CrVATDiff := 0;
    //                         CrVATDiffPerlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'VAT-DIF-DI');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             CrVATDiffPerlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             CrVATDiff := recPostedStrLineDet.Amount;
    //                         END;


    //                         //>>15Nov2018
    //                         NATDisc := 0;
    //                         NATDiscperlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'NATSCHEME');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             NATDiscperlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             NATDisc := recPostedStrLineDet.Amount;
    //                         END;

    //                         RegDisc := 0;
    //                         RegDiscperlt := 0;
    //                         recPostedStrLineDet.RESET;
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Invoice No.", "Sales Cr.Memo Line"."Document No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Line No.", "Sales Cr.Memo Line"."Line No.");
    //                         recPostedStrLineDet.SETRANGE(recPostedStrLineDet."Tax/Charge Group", 'REGSCHEME');
    //                         IF recPostedStrLineDet.FINDSET THEN BEGIN
    //                             RegDiscperlt := recPostedStrLineDet."Calculation Value" / "Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                             RegDisc := recPostedStrLineDet.Amount;
    //                         END;
    //                         //<<15Nov2018

    //                         CrProdDisc := 0;
    //                         CrProdDiscperlt := 0;
    //                         MRpMstr.RESET;
    //                         MRpMstr.SETRANGE(MRpMstr."Item No.", "Sales Cr.Memo Line"."No.");
    //                         MRpMstr.SETRANGE(MRpMstr."Lot No.", "Sales Cr.Memo Line"."Lot No.");
    //                         IF MRpMstr.FINDFIRST THEN BEGIN
    //                             CrProdDiscperlt := MRpMstr."National Discount";
    //                             CrProdDiscperlt := MRpMstr."Price Support";
    //                             CrProdDisc := CrProdDiscperlt * "Sales Cr.Memo Line"."Quantity (Base)";// 28Apr2017
    //                                                                                                    //CrProdDisc := ProdDiscPerLt * "Sales Cr.Memo Line"."Quantity (Base)";//NAV2009 Bug 28Apr2017
    //                             CrPriceSupport := CrPriceSupportPerLt * "Sales Cr.Memo Line"."Quantity (Base)";
    //                         END;
    //                         //<<1

    //                         //>>02Jun2019
    //                         IF "Sales Cr.Memo Header"."Posting Date" <= 030319D THEN BEGIN
    //                             NATDisc += "Sales Cr.Memo Line"."National Discount";
    //                             NATDiscperlt += "Sales Cr.Memo Line"."National Discount" / "Sales Cr.Memo Line"."Quantity (Base)";
    //                         END;
    //                         //<<02Jun2019

    //                         //>>27Apr2017 GroupData
    //                         CTINVCOUNT += 1;

    //                         CSTR25Apr.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
    //                         //Invoice No.,Item No.,Line No.,Tax/Charge Type,Tax/Charge Group
    //                         CSTR25Apr.SETRANGE("Line No.", "Line No.");
    //                         IF CSTR25Apr.FINDLAST THEN BEGIN
    //                             CIINVCOUNT := CSTR25Apr.COUNT;

    //                         END;

    //                         IF (CTINVCOUNT = CIINVCOUNT) AND PrintToExcel THEN BEGIN


    //                             //CTotalDisc := SpclFOC+StateDisc+Spot+AVPSanction+NewPrice+SpecDisc+FOC+MntDisc+AddtlDisc+VATDiff;
    //                             //TotalDiscPerlt := SpclFOCPerlt+StateDiscPerlt+SpotPerlt+AVPSanctionPerlt+NewPricePerlt+SpecDiscPerlt+FOCperlt+MntDiscPerlt
    //                             //              +AddtlDiscPerlt+VATDiffPerlt;

    //                             //MakeExcelInvoiceGroupfooter;
    //                             CrTotalDisc := 0;
    //                             CrTotalDisc := CrSpclFOC + CrStateDisc + CrSpot + CrAVPSanction + CrNewPrice + CrSpecDisc + CrFOC + CrMntDisc + CrAddtlDisc + CrVATDiff + NATDisc + RegDisc;

    //                             CrTotalDiscPerlt := CrSpclFOCPerlt + CrStateDiscPerlt + CrSpotPerlt + CrAVPSanctionPerlt + CrNewPricePerlt
    //                                                + CrSpecDiscPerlt + CrFOCperlt + CrMntDiscPerlt + CrAddtlDiscPerlt + CrVATDiffPerlt + NATDiscperlt + RegDiscperlt;

    //                             crTotListPrice += CrListPrice + CrProdDiscperlt + CrPriceSupportPerLt;

    //                             CrTotSpclFOC += CrSpclFOC;
    //                             CrTotFoc += CrFOC;
    //                             CrTotspot += CrSpot;
    //                             CrTotMntDisc += CrMntDisc;
    //                             CrTotStateDisc += CrStateDisc;
    //                             CrTotAvpSanction += CrAVPSanction;
    //                             CrTotNewPrice += CrNewPrice;
    //                             CrTotspecDisc += CrSpecDisc;
    //                             CrTotAddtlDisc += CrAddtlDisc;
    //                             CrTotVATDiff += CrVATDiff;


    //                             MakeExcelDataGroupFooterCrMemo;

    //                             CTINVCOUNT := 0;
    //                             //TotalDisc:=0;



    //                             CTSpclFOC += SpclFOC;//25Apr2017
    //                             CTStateDisc += StateDisc;//25Apr2017
    //                             CTSpot += Spot;//25Apr2017
    //                             CTAVPSanction += AVPSanction;//25Apr2017
    //                             CTPriceSupport += PriceSupport;//27Apr2017
    //                             CTNewPrice += NewPrice;//25Apr2017
    //                             CTSpecDisc += SpecDisc;//25Apr2017
    //                             CTFOC += FOC;//25Apr2017
    //                             CTMntDisc += MntDisc;//25Apr2017
    //                             CTAddtlDisc += AddtlDisc;//25Apr2017
    //                             CTVATDiff += VATDiff;//25Apr2017
    //                             CTProdDisc += ProdDisc;//25Apr2017
    //                             CTListPrice += ListPrice;//27Apr2017
    //                             RegDiscTotal += RegDisc;//15Nov2018
    //                             NATDiscTotal += NATDisc;//15Nov2018
    //                         END;
    //                         //<<27Apr2017 GroupData
    //                     end;

    //                     trigger OnPreDataItem()
    //                     begin

    //                         CSTR25Apr.COPYFILTERS("<Posted Str Ordr Ln Detail1>");//28Apr2017
    //                         CTINVCOUNT := 0; //28Apr2017
    //                     end;
    //                 }

    //                 trigger OnAfterGetRecord()
    //                 begin

    //                     //>>1

    //                     //CrListPrice := "Sales Cr.Memo Line"."Unit Price Incl. of Tax"/"Sales Cr.Memo Line"."Qty. per Unit of Measure";
    //                     CrListPrice := "Sales Cr.Memo Line"."List Price";//02Jun2019

    //                     CrTotQty += "Sales Cr.Memo Line".Quantity;
    //                     CrQtyTotBase += "Sales Cr.Memo Line"."Quantity (Base)";
    //                     //<<1
    //                 end;

    //                 trigger OnPreDataItem()
    //                 begin

    //                     //>>1

    //                     CurrReport.CREATETOTALS(CrFOC, CrSpclFOC, CrSpot, CrNewPrice, CrAVPSanction, CrAddtlDisc, CrMntDisc, CrVATDiff, CrCD, CrOtherDisc);
    //                     CurrReport.CREATETOTALS(CrTotalDisc, CrStateDisc, CrSpecDisc, CrProdDisc, CrProdDiscperlt, CrPriceSupportPerLt, CrPriceSupport);
    //                     //<<1

    //                     CrListPrice := 0;
    //                 end;
    //             }

    //             trigger OnAfterGetRecord()
    //             begin

    //                 //>>1

    //                 //CrTotalAmount := 0; //28Apr2017
    //                 CrCashDiscountAmt := 0;
    //                 recPostedStrOrdDetails.RESET;
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Header"."No.");
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'CASHDISC');
    //                 IF recPostedStrOrdDetails.FINDFIRST THEN
    //                     REPEAT
    //                         CrCashDiscountAmt += recPostedStrOrdDetails.Amount;
    //                     UNTIL recPostedStrOrdDetails.NEXT = 0;

    //                 IF CrCashDiscountAmt <> 0 THEN
    //                     CrCashDiscText := 'Cash Discount @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%'
    //                 ELSE
    //                     CrCashDiscText := '';


    //                 CrCCFC := 0;
    //                 recPostedStrOrdDetails.RESET;
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Header"."No.");
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'CCFC');
    //                 IF recPostedStrOrdDetails.FINDFIRST THEN
    //                     REPEAT
    //                         CrCCFC += recPostedStrOrdDetails.Amount;
    //                     UNTIL recPostedStrOrdDetails.NEXT = 0;

    //                 IF CrCCFC <> 0 THEN
    //                     CrCCFCtext := 'CCFC @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%'
    //                 ELSE
    //                     CrCCFCtext := '';

    //                 //Sharath
    //                 V_TransSub1 := 0;
    //                 recPostedStrOrdDetails.RESET;
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Header"."No.");
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
    //                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'TPTSUBSIDY');
    //                 IF recPostedStrOrdDetails.FINDFIRST THEN
    //                     REPEAT
    //                         V_TransSub1 += recPostedStrOrdDetails.Amount;
    //                     UNTIL recPostedStrOrdDetails.NEXT = 0;

    //                 IF V_TransSub1 <> 0 THEN
    //                     V_TransSubTxt1 := 'Transport Subsidy'
    //                 ELSE
    //                     V_TransSubTxt1 := '';

    //                 //Sharath
    //                 //<<1
    //             end;

    //             trigger OnPostDataItem()
    //             begin

    //                 //>>1
    //                 IF "Sales Return" AND PrintToExcel THEN
    //                     MakeExcelDataFooterCrMemo;
    //                 //<<1
    //             end;

    //             trigger OnPreDataItem()
    //             begin
    //                 NATDiscTotal := 0;
    //                 RegDiscTotal := 0;

    //                 //>>1
    //                 IF "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Posting Date") <> '' THEN BEGIN
    //                     "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Posting Date", StartDate, EndDate);
    //                 END;

    //                 IF "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Location Code") <> '' THEN BEGIN
    //                     "Sales Cr.Memo Header".SETRANGE("Sales Cr.Memo Header"."Location Code",
    //                                                "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Location Code"));
    //                 END;

    //                 CurrReport.CREATETOTALS(CrFOC, CrSpot, CrNewPrice, CrAVPSanction, CrAddtlDisc, CrMntDisc, CrVATDiff, CrCD, CrOtherDisc);
    //                 CurrReport.CREATETOTALS(CrTotalDisc, CrStateDisc, CrSpecDisc, CrProdDiscperlt, CrProdDisc, CrProdDiscperlt);
    //                 //<<1

    //                 //>>2

    //                 IF "Sales Return" AND PrintToExcel THEN
    //                     MakeExcelDataHeaderCrMemo;//28Apr2017
    //                 //<<2
    //             end;
    //         }
    //     }

    //     requestpage
    //     {

    //         layout
    //         {
    //             area(content)
    //             {
    //                 field("Print to Excel"; PrintToExcel)
    //                 {
    //                 }
    //                 field("Sales Invoice"; "Sales Invoice")
    //                 {
    //                 }
    //                 field("Cancel Invoice / Sales Return"; "Sales Return")
    //                 {
    //                 }
    //             }
    //         }

    //         actions
    //         {
    //         }
    //     }

    //     labels
    //     {
    //     }

    //     trigger OnPostReport()
    //     begin

    //         //>>1

    //         IF PrintToExcel THEN
    //             CreateExcelbook;
    //         //<<1
    //     end;

    //     trigger OnPreReport()
    //     begin

    //         //>>1

    //         LocCode := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Location Code");
    //         //user.GET(USERID);

    //         /*
    //         MemberOf.RESET;
    //         MemberOf.SETRANGE(MemberOf."User ID",user."User ID");
    //         MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
    //         IF NOT(MemberOf.FINDFIRST) THEN
    //         BEGIN
    //           CLEAR(LocResC);
    //           Location.RESET;
    //           Location.SETRANGE(Location.Code,LocResC);
    //           IF Location.FINDFIRST THEN
    //           BEGIN
    //            LocResC := Location."Global Dimension 2 Code";
    //           END;

    //           CSOMapping.RESET;
    //           CSOMapping.SETRANGE(CSOMapping."User Id",UPPERCASE(USERID));
    //           CSOMapping.SETRANGE(Type,CSOMapping.Type::Location);
    //           CSOMapping.SETRANGE(CSOMapping.Value,LocCode);
    //           IF NOT (CSOMapping.FINDFIRST) THEN
    //           BEGIN
    //             CSOMapping.RESET;
    //             CSOMapping.SETRANGE(CSOMapping."User Id",UPPERCASE(USERID));
    //             CSOMapping.SETRANGE(CSOMapping.Type,CSOMapping.Type::"Responsibility Center");
    //             CSOMapping.SETRANGE(CSOMapping.Value,LocResC);
    //           IF NOT(CSOMapping.FINDFIRST) THEN
    //             ERROR ('You are not allowed to run this report other than your location');
    //           END;
    //         END;
    //         *///Commented 27Apr2017

    //         IF ("Sales Invoice" = FALSE) AND ("Sales Return" = FALSE) THEN
    //             ERROR('You have not selected any option');

    //         CompInfo.GET;

    //         IF PrintToExcel THEN
    //             MakeExcelInfo;

    //         //<<1

    //     end;

    //     var
    //         LastFieldNo: Integer;
    //         FooterPrinted: Boolean;
    //         FOC: Decimal;
    //         FOCperlt: Decimal;
    //         FOCTotal: Decimal;
    //         SpclFOC: Decimal;
    //         SpclFOCPerlt: Decimal;
    //         SpclFOCTotal: Decimal;
    //         Spot: Decimal;
    //         SpotPerlt: Decimal;
    //         SpotTotal: Decimal;
    //         NewPrice: Decimal;
    //         NewPricePerlt: Decimal;
    //         NewPriceTotal: Decimal;
    //         AVPSanction: Decimal;
    //         AVPSanctionPerlt: Decimal;
    //         AVPSanctionTotal: Decimal;
    //         AddtlDisc: Decimal;
    //         AddtlDiscPerlt: Decimal;
    //         AddtlDiscTotal: Decimal;
    //         MntDisc: Decimal;
    //         MntDiscPerlt: Decimal;
    //         MntDiscTotal: Decimal;
    //         VATDiff: Decimal;
    //         VATDiffPerlt: Decimal;
    //         VATDiffTotal: Decimal;
    //         CD: Decimal;
    //         CDPerlt: Decimal;
    //         CDTotal: Decimal;
    //         OtherDisc: Decimal;
    //         OtherDiscPerlt: Decimal;
    //         OtherDiscTotal: Decimal;
    //         TotalDisc: Decimal;
    //         recPostedStrLineDet: Record 13798;
    //         ExcelBuf: Record 370 temporary;
    //         ItemRec: Record 27;
    //         ListPrice: Decimal;
    //         StateDisc: Decimal;
    //         StateDiscPerlt: Decimal;
    //         SpecDisc: Decimal;
    //         SpecDiscPerlt: Decimal;
    //         TotalDiscPerlt: Decimal;
    //         TotalAmount: Decimal;
    //         CashDiscountAmt: Decimal;
    //         CashDiscText: Text[50];
    //         recPostedStrOrdDetails: Record 13798;
    //         CCFC: Decimal;
    //         CCFCtext: Text[30];
    //         CrFOC: Decimal;
    //         CrFOCperlt: Decimal;
    //         CrSpclFOC: Decimal;
    //         CrSpclFOCPerlt: Decimal;
    //         CrStateDisc: Decimal;
    //         CrStateDiscPerlt: Decimal;
    //         CrSpot: Decimal;
    //         CrSpotPerlt: Decimal;
    //         CrListPrice: Decimal;
    //         CrAVPSanction: Decimal;
    //         CrAVPSanctionPerlt: Decimal;
    //         CrNewPrice: Decimal;
    //         CrAddtlDisc: Decimal;
    //         CrMntDisc: Decimal;
    //         CrVATDiff: Decimal;
    //         CrCD: Decimal;
    //         CrOtherDisc: Decimal;
    //         CrTotalDisc: Decimal;
    //         CrSpecDisc: Decimal;
    //         CrNewPricePerlt: Decimal;
    //         CrCDPerlt: Decimal;
    //         CrMntDiscPerlt: Decimal;
    //         CrAddtlDiscPerlt: Decimal;
    //         CrVATDiffPerlt: Decimal;
    //         CrSpecDiscPerlt: Decimal;
    //         CrTotalDiscPerlt: Decimal;
    //         CrTotalAmount: Decimal;
    //         CrCashDiscountAmt: Decimal;
    //         CrCashDiscText: Text[50];
    //         CrCCFC: Decimal;
    //         CrCCFCtext: Text[30];
    //         CrQty: Decimal;
    //         CrTotQty: Decimal;
    //         CrQtyBase: Decimal;
    //         CrQtyTotBase: Decimal;
    //         CrTotspecDisc: Decimal;
    //         TotQty: Decimal;
    //         TotQtyBase: Decimal;
    //         TotTotalAmt: Decimal;
    //         TotTotalDiscPerlt: Decimal;
    //         TotListPrice: Decimal;
    //         CrTotSpclFOCPerlt: Decimal;
    //         CrTotFoc: Decimal;
    //         CrTotspot: Decimal;
    //         CrTotMntDisc: Decimal;
    //         CrTotTotalDisc: Decimal;
    //         CrTotTotalAmount: Decimal;
    //         CrTotspotperlt: Decimal;
    //         CrTotSpclFOC: Decimal;
    //         CrTotStateDisc: Decimal;
    //         CrTotAvpSanction: Decimal;
    //         CrTotNewPrice: Decimal;
    //         CrTotAddtlDisc: Decimal;
    //         CrTotVATDiff: Decimal;
    //         "Sales Invoice": Boolean;
    //         "Sales Return": Boolean;
    //         PrintToExcel: Boolean;
    //         StartDate: Date;
    //         EndDate: Date;
    //         CompInfo: Record 79;
    //         CSOMapping: Record 50006;
    //         LocCode: Code[10];
    //         user: Record 2000000120;
    //         Location: Record 14;
    //         LocResC: Code[10];
    //         ProdDisc: Decimal;
    //         TotProdDisc: Decimal;
    //         ProdDiscPerLt: Decimal;
    //         MRpMstr: Record 50013;
    //         Netpricetocustperltr: Decimal;
    //         CrProdDisc: Decimal;
    //         CrProdDiscperlt: Decimal;
    //         crNetpricetocustperltr: Decimal;
    //         CrTottotdiscperlt: Decimal;
    //         crTotListPrice: Decimal;
    //         TotProdDiscPerlt: Decimal;
    //         totTotalDisc: Decimal;
    //         totTotdisPerlt: Decimal;
    //         crtotTotdisPerlt: Decimal;
    //         V_TransSub: Decimal;
    //         V_TransSubTxt: Text[50];
    //         V_TransSub1: Decimal;
    //         V_TransSubTxt1: Text[50];
    //         PriceSupportPerLt: Decimal;
    //         PriceSupport: Decimal;
    //         CrPriceSupportPerLt: Decimal;
    //         CrPriceSupport: Decimal;
    //         "----------27Apr2017": Integer;
    //         TINVCOUNT: Integer;
    //         IINVCOUNT: Integer;
    //         STR25Apr: Record 13798;
    //         TQty: Decimal;
    //         TQtyBase: Decimal;
    //         TListPrice: Decimal;
    //         TTDisc: Decimal;
    //         TSpclFOC: Decimal;
    //         TStateDisc: Decimal;
    //         TSpot: Decimal;
    //         TAVPSanction: Decimal;
    //         TNewPrice: Decimal;
    //         TSpecDisc: Decimal;
    //         TFOC: Decimal;
    //         TMntDisc: Decimal;
    //         TAddtlDisc: Decimal;
    //         TVATDiff: Decimal;
    //         TProdDisc: Decimal;
    //         TPriceSupport: Decimal;
    //         TTotalAmount: Decimal;
    //         "----------28Apr2017": Integer;
    //         CTINVCOUNT: Integer;
    //         CIINVCOUNT: Integer;
    //         CSTR25Apr: Record 13798;
    //         CTQty: Decimal;
    //         CTQtyBase: Decimal;
    //         CTListPrice: Decimal;
    //         CTTDisc: Decimal;
    //         CTSpclFOC: Decimal;
    //         CTStateDisc: Decimal;
    //         CTSpot: Decimal;
    //         CTAVPSanction: Decimal;
    //         CTNewPrice: Decimal;
    //         CTSpecDisc: Decimal;
    //         CTFOC: Decimal;
    //         CTMntDisc: Decimal;
    //         CTAddtlDisc: Decimal;
    //         CTVATDiff: Decimal;
    //         CTProdDisc: Decimal;
    //         CTPriceSupport: Decimal;
    //         CTTotalAmount: Decimal;
    //         "--------15Nov2018": Integer;
    //         NATDisc: Decimal;
    //         NATDiscperlt: Decimal;
    //         NATDiscTotal: Decimal;
    //         RegDisc: Decimal;
    //         RegDiscperlt: Decimal;
    //         RegDiscTotal: Decimal;

    //     //[Scope('Internal')]
    //     procedure MakeExcelInfo()
    //     begin
    //     end;

    //     //[Scope('Internal')]
    //     procedure MakeExcelDataHeader()
    //     begin

    //         //>> ReportHeader 27Apr2017
    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27 15Nov2018
    //         ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//28


    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('Automotive Discount Register', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27 15Nov2018
    //         ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//28


    //         ExcelBuf.NewRow;
    //         ExcelBuf.NewRow;
    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('Date Filter : ' + "Sales Invoice Header".GETFILTER("Posting Date"), FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28 15Nov2018
    //         //<< ReportHeader 27Apr2017

    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
    //         ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
    //         ExcelBuf.AddColumn('Customer No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
    //         ExcelBuf.AddColumn('Customer Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
    //         ExcelBuf.AddColumn('FG Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
    //         ExcelBuf.AddColumn('Product Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
    //         ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
    //         ExcelBuf.AddColumn('Pack Size in Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
    //         ExcelBuf.AddColumn('Qty Billed', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
    //         ExcelBuf.AddColumn('Total Invoice Qty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
    //         ExcelBuf.AddColumn('List Price Billed', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
    //         ExcelBuf.AddColumn('FOC Product', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
    //         ExcelBuf.AddColumn('State Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
    //         ExcelBuf.AddColumn('Spot Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
    //         ExcelBuf.AddColumn('HO Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
    //         ExcelBuf.AddColumn('Price Support', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
    //         ExcelBuf.AddColumn('Special Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
    //         ExcelBuf.AddColumn('FOC Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
    //         ExcelBuf.AddColumn('Coupon Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
    //         ExcelBuf.AddColumn('Additional Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
    //         ExcelBuf.AddColumn('VAT Diff Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
    //         ExcelBuf.AddColumn('National Discount Scheme', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 15Nov2018
    //         ExcelBuf.AddColumn('Regional Discount Scheme', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 15Nov2018
    //         ExcelBuf.AddColumn('Total Discount Per Lt', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
    //         ExcelBuf.AddColumn('Product Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
    //         ExcelBuf.AddColumn('Total Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
    //         ExcelBuf.AddColumn('Net Price Per Liter to Customer', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
    //         ExcelBuf.AddColumn('Net Price Amount to Customer', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
    //     end;

    //     [Scope('Internal')]
    //     procedure MakeExcelDataBody()
    //     begin
    //     end;

    //     [Scope('Internal')]
    //     procedure CreateExcelbook()
    //     begin
    //         //ExcelBuf.CreateBook;
    //         //ExcelBuf.CreateSheet('Data','Discount Details',COMPANYNAME,USERID);
    //         //ExcelBuf.GiveUserControl;
    //         //ERROR('');

    //         //>>27Apr2017
    //         ExcelBuf.CreateBook('', 'Discount Details');
    //         ExcelBuf.CreateBookAndOpenExcel('', 'Discount Details', '', '', USERID);
    //         ExcelBuf.GiveUserControl;

    //         //<<27Apr2017
    //     end;

    //     [Scope('Internal')]
    //     procedure MakeExcelDataHeaderCrMemo()
    //     begin
    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('Credit Memo No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
    //         ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
    //         ExcelBuf.AddColumn('Customer No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
    //         ExcelBuf.AddColumn('Customer Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
    //         ExcelBuf.AddColumn('FG Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
    //         ExcelBuf.AddColumn('Product Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
    //         ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
    //         ExcelBuf.AddColumn('Pack Size in Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
    //         ExcelBuf.AddColumn('Qty Billed', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
    //         ExcelBuf.AddColumn('Total Invoice Qty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
    //         ExcelBuf.AddColumn('List Price Billed', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
    //         ExcelBuf.AddColumn('Spacial FOC Disc', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
    //         ExcelBuf.AddColumn('State Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
    //         ExcelBuf.AddColumn('Spot Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
    //         ExcelBuf.AddColumn('AVP Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
    //         ExcelBuf.AddColumn('Price Support', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
    //         ExcelBuf.AddColumn('Special Mgmt Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
    //         ExcelBuf.AddColumn('FOC Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
    //         ExcelBuf.AddColumn('Monthly Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
    //         ExcelBuf.AddColumn('Additional Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
    //         ExcelBuf.AddColumn('VAT Diff Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
    //         ExcelBuf.AddColumn('National Discount Scheme', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 15Nov2018
    //         ExcelBuf.AddColumn('Regional Discount Scheme', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 15Nov2018
    //         ExcelBuf.AddColumn('Total Discount Per Lt', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
    //         ExcelBuf.AddColumn('Product Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
    //         ExcelBuf.AddColumn('Total Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
    //         ExcelBuf.AddColumn('Net Price Per Liter to Customer', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
    //         ExcelBuf.AddColumn('Net Price Amount to Customer', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
    //     end;

    //     [Scope('Internal')]
    //     procedure MakeExcelDataFooterCrMemo()
    //     begin
    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('Total Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
    //         ExcelBuf.AddColumn(CrTotQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//9
    //         ExcelBuf.AddColumn(CrQtyTotBase, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//10
    //         ExcelBuf.AddColumn(crTotListPrice, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//11
    //         ExcelBuf.AddColumn(CrTotSpclFOC, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//12
    //         ExcelBuf.AddColumn(CrTotStateDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//13
    //         ExcelBuf.AddColumn(CrTotspot, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//14
    //         ExcelBuf.AddColumn(CrTotAvpSanction, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//15
    //         ExcelBuf.AddColumn(CrPriceSupportPerLt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//16
    //         ExcelBuf.AddColumn(CrTotspecDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//17
    //         ExcelBuf.AddColumn(CrTotFoc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//18
    //         ExcelBuf.AddColumn(CrTotMntDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//19
    //         ExcelBuf.AddColumn(CrAddtlDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//20
    //         ExcelBuf.AddColumn(CrVATDiff, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//21
    //         ExcelBuf.AddColumn(NATDiscTotal, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//22 15Nov2018
    //         ExcelBuf.AddColumn(RegDiscTotal, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//23 15Nov2018

    //         ExcelBuf.AddColumn(ROUND(CrTotSpclFOC + CrTotStateDisc + CrTotspot + CrTotAvpSanction
    //                + CrTotNewPrice + CrTotMntDisc + CrAddtlDisc + CrTotFoc + CrVATDiff + NATDiscTotal + VATDiffTotal, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);
    //         ExcelBuf.AddColumn(CrProdDiscperlt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//23
    //         ExcelBuf.AddColumn(ROUND(CrTotTotalDisc, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//24
    //         ExcelBuf.AddColumn(ROUND(crNetpricetocustperltr, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//25
    //         ExcelBuf.AddColumn(ROUND(CrTotalAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//26

    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn(CrCashDiscText, FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn(ABS(CrCashDiscountAmt), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26

    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn(V_TransSubTxt1, FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn(ABS(V_TransSub1), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26

    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn(CrCCFCtext, FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn(ABS(CrCCFC), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26


    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn('Net Discount Amount', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn(ROUND(CrTotTotalAmount + V_TransSub + CrCashDiscountAmt + CrCCFC, 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26
    //     end;

    //     // [Scope('Internal')]
    //     procedure MakeExcelInvoiceGroupfooter()
    //     begin
    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn("Sales Invoice Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn("Sales Invoice Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2
    //         ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn("Posted Str Order Line Details"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn(ItemRec.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn("Sales Invoice Line"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn("Sales Invoice Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//8
    //         ExcelBuf.AddColumn("Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//9
    //         ExcelBuf.AddColumn("Sales Invoice Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//10
    //         //ExcelBuf.AddColumn(ListPrice+ProdDiscPerLt+CrPriceSupportPerLt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//11
    //         //>>02Jun2019
    //         IF "Sales Invoice Header"."Posting Date" <= 030319D THEN
    //             ExcelBuf.AddColumn(ListPrice + ProdDiscPerLt + CrPriceSupportPerLt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0)
    //         ELSE
    //             ExcelBuf.AddColumn(ListPrice + CrPriceSupportPerLt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//11
    //         //<<02Jun2019
    //         ExcelBuf.AddColumn(SpclFOCPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//12
    //         ExcelBuf.AddColumn(StateDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//13
    //         ExcelBuf.AddColumn(SpotPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//14
    //         ExcelBuf.AddColumn(AVPSanctionPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//15
    //         ExcelBuf.AddColumn(-PriceSupportPerLt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//16
    //         ExcelBuf.AddColumn(SpecDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//17
    //         ExcelBuf.AddColumn(FOCperlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//18
    //         ExcelBuf.AddColumn(MntDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//19
    //         ExcelBuf.AddColumn(AddtlDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//20
    //         ExcelBuf.AddColumn(VATDiffPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//21
    //         ExcelBuf.AddColumn(NATDiscperlt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22 15Nov2018
    //         ExcelBuf.AddColumn(RegDiscperlt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23 15Nov2018
    //         ExcelBuf.AddColumn(ROUND(TotalDiscPerlt, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//22
    //         ExcelBuf.AddColumn(-ProdDiscPerLt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23
    //         ExcelBuf.AddColumn(ROUND(TotalDisc + ProdDisc + PriceSupport, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//24

    //         IF ProdDiscPerLt <> 0 THEN BEGIN
    //             IF "Sales Invoice Header"."Posting Date" <= 030319D THEN BEGIN
    //                 ExcelBuf.AddColumn(ROUND(((ListPrice + ProdDiscPerLt + PriceSupportPerLt)
    //                 - (ABS(TotalDiscPerlt + (-ProdDiscPerLt) + (-PriceSupportPerLt)))), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25

    //                 ExcelBuf.AddColumn(ROUND(((ListPrice + ProdDiscPerLt + PriceSupportPerLt)
    //                 - (ABS(TotalDiscPerlt + (-ProdDiscPerLt) + (-PriceSupportPerLt)))) *
    //                                       "Sales Invoice Line"."Quantity (Base)", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//26

    //                 TotalAmount += ((ListPrice + ProdDiscPerLt + PriceSupportPerLt)
    //                 - (ABS(TotalDiscPerlt + (-ProdDiscPerLt) + (-PriceSupportPerLt)))) * "Sales Invoice Line"."Quantity (Base)";

    //                 Netpricetocustperltr += ((ListPrice + ProdDiscPerLt + PriceSupportPerLt) - (ABS(TotalDiscPerlt + (-ProdDiscPerLt) + (-PriceSupportPerLt))));
    //             END ELSE BEGIN
    //                 //>>02Jun2019
    //                 ExcelBuf.AddColumn(ROUND(((ListPrice + PriceSupportPerLt)
    //             - (ABS(TotalDiscPerlt + (-ProdDiscPerLt) + (-PriceSupportPerLt)))), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//25

    //                 ExcelBuf.AddColumn(ROUND(((ListPrice + PriceSupportPerLt)
    //                 - (ABS(TotalDiscPerlt + (-ProdDiscPerLt) + (-PriceSupportPerLt)))) *
    //                                       "Sales Invoice Line"."Quantity (Base)", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26

    //                 TotalAmount += ((ListPrice + PriceSupportPerLt)
    //                 - (ABS(TotalDiscPerlt + (-ProdDiscPerLt) + (-PriceSupportPerLt)))) * "Sales Invoice Line"."Quantity (Base)";

    //                 Netpricetocustperltr += ((ListPrice + PriceSupportPerLt) - (ABS(TotalDiscPerlt + (-ProdDiscPerLt) + (-PriceSupportPerLt))));
    //                 //<<02Jun2019
    //             END;
    //         END ELSE BEGIN
    //             ExcelBuf.AddColumn(ROUND((ListPrice + (TotalDiscPerlt)), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25

    //             ExcelBuf.AddColumn(ROUND((ListPrice + (TotalDiscPerlt)) * "Sales Invoice Line"."Quantity (Base)", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//26

    //             TotalAmount += (ListPrice + (TotalDiscPerlt)) * "Sales Invoice Line"."Quantity (Base)";
    //             Netpricetocustperltr += (ListPrice + (TotalDiscPerlt));
    //         END;

    //         totTotdisPerlt += TotalDiscPerlt;
    //         totTotalDisc += TotalDisc + ProdDisc + PriceSupport;

    //         TTotalAmount += TotalAmount;//27Apr2017
    //     end;

    //     // [Scope('Internal')]
    //     procedure MakeExcelInvoiceGrandTotal()
    //     begin
    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('Total Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
    //         ExcelBuf.AddColumn(TotQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//9
    //         ExcelBuf.AddColumn(TotQtyBase, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//10

    //         ExcelBuf.AddColumn(TListPrice, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//11 //27Apr2017
    //         //ExcelBuf.AddColumn(TotListPrice,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//11

    //         ExcelBuf.AddColumn(TSpclFOC, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//12 //27Apr2017
    //         //TSpclFOC := 0;//27Apr2017
    //         //ExcelBuf.AddColumn(SpclFOC,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//12

    //         ExcelBuf.AddColumn(TStateDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//13 //27Apr2017
    //         //TStateDisc := 0;//27Apr2017
    //         //ExcelBuf.AddColumn(StateDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//13

    //         ExcelBuf.AddColumn(TSpot, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//14//27Apr2017
    //         //TSpot := 0; //27Apr2017
    //         //ExcelBuf.AddColumn(Spot,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//14

    //         ExcelBuf.AddColumn(TAVPSanction, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//15//27Apr2017
    //         //TAVPSanction := 0;//27Apr2017
    //         //ExcelBuf.AddColumn(AVPSanction,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//15

    //         ExcelBuf.AddColumn(TPriceSupport, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//16 //27Apr2017
    //         //TPriceSupport := 0;//27Apr2017
    //         //ExcelBuf.AddColumn(PriceSupport,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//16

    //         ExcelBuf.AddColumn(TSpecDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//17 //27Apr2017
    //         //TSpecDisc := 0;//27Apr2017
    //         //ExcelBuf.AddColumn(SpecDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//17

    //         ExcelBuf.AddColumn(TFOC, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//18 //27Apr2017
    //         //TFOC := 0;//27Apr2017
    //         //ExcelBuf.AddColumn(FOC,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//18

    //         ExcelBuf.AddColumn(TMntDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//19 //27Apr2017
    //         //TMntDisc := 0;//27Apr2017
    //         //ExcelBuf.AddColumn(MntDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//19

    //         ExcelBuf.AddColumn(TAddtlDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//20 //27Apr2017
    //         //TAddtlDisc := 0;//27Apr2017
    //         //ExcelBuf.AddColumn(AddtlDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//20

    //         ExcelBuf.AddColumn(TVATDiff, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//21 //27Apr2017
    //         //TVATDiff := 0; //27Apr2017

    //         ExcelBuf.AddColumn(NATDiscTotal, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//22 15Nov2018
    //         ExcelBuf.AddColumn(RegDiscTotal, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//23 15Nov2018

    //         ExcelBuf.AddColumn(ROUND(TSpclFOC + TStateDisc + TSpot + TAVPSanction + TNewPrice + TSpecDisc + TFOC
    //                                        + TMntDisc + TAddtlDisc + TVATDiff + NATDiscTotal + RegDiscTotal, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);

    //         //ExcelBuf.AddColumn(ROUND(SpclFOC + StateDisc + Spot+ AVPSanction + NewPrice + SpecDisc + FOC
    //         //                             + MntDisc + AddtlDisc + VATDiff,0.01),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//22 //EBT MILAN

    //         ExcelBuf.AddColumn(TProdDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//23 //27Apr2017
    //         //ExcelBuf.AddColumn(ProdDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//23

    //         ExcelBuf.AddColumn(ROUND(totTotalDisc, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//24
    //         ExcelBuf.AddColumn(ROUND(Netpricetocustperltr, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//25

    //         ExcelBuf.AddColumn(ROUND(TotalAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//26 //27Apr2017
    //         //ExcelBuf.AddColumn(ROUND(TotalAmount,0.01),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//26


    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn(CashDiscText, FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn(ABS(CashDiscountAmt), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26

    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn(V_TransSubTxt, FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn(ABS(V_TransSub), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26

    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn(CCFCtext, FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn(ABS(CCFC), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26


    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 15Nov2018
    //         ExcelBuf.AddColumn('Net Discount Amount', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn(ROUND(TotalAmount + CashDiscountAmt + CCFC, 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26

    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27 15Nov2018
    //         ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//28 15Nov2018
    //     end;

    //     // [Scope('Internal')]
    //     procedure MakeExcelDataGroupFooterCrMemo()
    //     begin
    //         ExcelBuf.NewRow;
    //         ExcelBuf.AddColumn("Sales Cr.Memo Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
    //         ExcelBuf.AddColumn("Sales Cr.Memo Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2
    //         ExcelBuf.AddColumn("Sales Cr.Memo Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
    //         ExcelBuf.AddColumn("Sales Cr.Memo Header"."Sell-to Customer Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
    //         ExcelBuf.AddColumn("<Posted Str Ordr Ln Detail1>"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
    //         ExcelBuf.AddColumn(ItemRec.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
    //         ExcelBuf.AddColumn("Sales Cr.Memo Line"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
    //         ExcelBuf.AddColumn("Sales Cr.Memo Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//8
    //         ExcelBuf.AddColumn("Sales Cr.Memo Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//9
    //         ExcelBuf.AddColumn("Sales Cr.Memo Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//10
    //         //ExcelBuf.AddColumn(CrListPrice+CrProdDiscperlt+CrPriceSupportPerLt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//11
    //         //>>02Jun2019
    //         IF "Sales Cr.Memo Header"."Posting Date" <= 030319D THEN
    //             ExcelBuf.AddColumn(CrListPrice + CrPriceSupportPerLt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)
    //         ELSE
    //             ExcelBuf.AddColumn(CrListPrice + CrProdDiscperlt + CrPriceSupportPerLt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//11
    //         //<<02Jun2019
    //         ExcelBuf.AddColumn(CrSpclFOCPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//12
    //         ExcelBuf.AddColumn(CrStateDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//13
    //         ExcelBuf.AddColumn(CrSpotPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//14
    //         ExcelBuf.AddColumn(CrAVPSanctionPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//15
    //         ExcelBuf.AddColumn(CrPriceSupportPerLt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//16
    //         ExcelBuf.AddColumn(CrSpecDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//17
    //         ExcelBuf.AddColumn(CrFOCperlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//18
    //         ExcelBuf.AddColumn(CrMntDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//19
    //         ExcelBuf.AddColumn(CrAddtlDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//20
    //         ExcelBuf.AddColumn(CrVATDiffPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//21
    //         ExcelBuf.AddColumn(NATDiscperlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//22 15Nov2018
    //         ExcelBuf.AddColumn(RegDiscperlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23 15Nov2018

    //         /*
    //         //Sharath
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'');
    //         //Sharath
    //         */
    //         ExcelBuf.AddColumn(ROUND(CrTotalDiscPerlt, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//22
    //         ExcelBuf.AddColumn(-CrProdDiscperlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23
    //         ExcelBuf.AddColumn(ROUND(CrTotalDisc + CrProdDisc + CrPriceSupport, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//24

    //         CrTotTotalDisc += CrTotalDisc + CrProdDisc + CrPriceSupport;//28Apr2017

    //         IF CrProdDiscperlt <> 0 THEN BEGIN
    //             //>>02Jun2019
    //             IF "Sales Cr.Memo Header"."Posting Date" <= 030319D THEN BEGIN
    //                 ExcelBuf.AddColumn(ROUND((CrListPrice + CrProdDiscperlt + CrPriceSupportPerLt)
    //                 - (ABS(CrTotalDiscPerlt + (-CrProdDiscperlt) + (-CrPriceSupportPerLt))), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25

    //                 ExcelBuf.AddColumn(ROUND(((CrListPrice + CrProdDiscperlt + CrPriceSupportPerLt)
    //                 - (ABS(CrTotalDiscPerlt + (-CrProdDiscperlt) + (-CrPriceSupportPerLt)))) *
    //                                     "Sales Cr.Memo Line"."Quantity (Base)", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//26

    //                 CrTotalAmount += ((CrListPrice + CrProdDiscperlt + CrPriceSupportPerLt)
    //                 - (ABS(CrTotalDiscPerlt + (-CrProdDiscperlt) + (-CrPriceSupportPerLt))))
    //                                               * "Sales Cr.Memo Line"."Quantity (Base)";

    //                 crNetpricetocustperltr += ((CrListPrice + CrProdDiscperlt + CrPriceSupportPerLt)
    //                 - (ABS(CrTotalDiscPerlt + (-CrProdDiscperlt) + (-CrPriceSupportPerLt))));
    //             END ELSE BEGIN
    //                 ExcelBuf.AddColumn(ROUND((CrListPrice + CrPriceSupportPerLt)
    //                 - (ABS(CrTotalDiscPerlt + (-CrProdDiscperlt) + (-CrPriceSupportPerLt))), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25

    //                 ExcelBuf.AddColumn(ROUND(((CrListPrice + CrPriceSupportPerLt)
    //                 - (ABS(CrTotalDiscPerlt + (-CrProdDiscperlt) + (-CrPriceSupportPerLt)))) *
    //                                     "Sales Cr.Memo Line"."Quantity (Base)", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//26

    //                 CrTotalAmount += ((CrListPrice + CrPriceSupportPerLt)
    //                 - (ABS(CrTotalDiscPerlt + (-CrProdDiscperlt) + (-CrPriceSupportPerLt))))
    //                                               * "Sales Cr.Memo Line"."Quantity (Base)";

    //                 crNetpricetocustperltr += ((CrListPrice + CrPriceSupportPerLt)
    //                 - (ABS(CrTotalDiscPerlt + (-CrProdDiscperlt) + (-CrPriceSupportPerLt))));
    //             END;
    //             //<<02Jun2019
    //         END ELSE BEGIN
    //             ExcelBuf.AddColumn(ROUND((CrListPrice + (CrTotalDiscPerlt)), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25
    //             ExcelBuf.AddColumn(ROUND((CrListPrice + (CrTotalDiscPerlt)) * "Sales Cr.Memo Line"."Quantity (Base)", 0.01),
    //                                                             FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//26

    //             CrTotalAmount += (CrListPrice + (CrTotalDiscPerlt)) * "Sales Cr.Memo Line"."Quantity (Base)";
    //             crNetpricetocustperltr += (CrListPrice + (CrTotalDiscPerlt));
    //         END;

    //         crtotTotdisPerlt += CrTotalDiscPerlt;
    //         //CrTotTotalDisc += CrTotalDisc+CrProdDisc+CrPriceSupport;

    //     end;
}


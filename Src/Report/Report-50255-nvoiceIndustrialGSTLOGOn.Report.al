report 50255 "Invoice Industrial GST--LOGO n"
{
    //     // 
    //     // Date        Version      Remarks
    //     // .....................................................................................
    //     // 14Nov2017   RB-N         Not to allowing SaveasPDF from Preview Mode
    //     DefaultLayout = RDLC;
    //     RDLCLayout = './InvoiceIndustrialGSTLOGOn.rdlc';

    //     Permissions = TableData 112 = rim;
    //     PreviewMode = PrintLayout;

    //     dataset
    //     {
    //         dataitem("Sales Invoice Header"; "Sales Invoice Header")
    //         {
    //             DataItemTableView = SORTING("No.");
    //             RequestFilterFields = "No.";
    //             column(SuppliersCode; SuppliersCode)
    //             {
    //             }
    //             column(FactoryCaption; FactoryCaption)
    //             {
    //             }
    //             column(CompName; CompanyInfo1.Name)
    //             {
    //             }
    //             column(CompPicture; CompanyInfo1.Picture)
    //             {
    //             }
    //             column(CompNamePicture; CompanyInfo1."Name Picture")
    //             {
    //             }
    //             column(CompShadedBoxPicture; CompanyInfo1."Shaded Box")
    //             {
    //             }
    //             column(FactoryLocation; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' - ' + LocFrom."Post Code" + ' ' + LocState + ' ' + ' Phone: ' + LocFrom."Phone No." + ' Email: ' + LocFrom."E-Mail")
    //             {
    //             }
    //             column(LocGSTNo; LocFrom."GST Registration No.")
    //             {
    //             }
    //             column(ExternalDocDate; "Document Date")
    //             {
    //             }
    //             column(ExternalDocDate2; "Order Date")
    //             {
    //             }
    //             column(ExternalDocNo; "External Document No.")
    //             {
    //             }
    //             column(Remarks1; "Sales Invoice Header".Remarks)
    //             {
    //             }
    //             column(Remarks2; "Sales Invoice Header".Remarks2)
    //             {
    //             }
    //             column(CurrencyFactor_Header; "Sales Invoice Header"."Currency Factor")
    //             {
    //             }
    //             column(SIH_InvoiceTime; "Sales Invoice Header"."Invoice Print Time")
    //             {
    //             }
    //             column(RoundOffAmnt; RoundOffAmnt)
    //             {
    //             }
    //             column(FinalAmt; FinalAmt)
    //             {
    //             }
    //             column(AmtinWord15; AmtinWord15[1])
    //             {
    //             }
    //             column(CGSTinWord15; CGSTinWord15[1])
    //             {
    //             }
    //             column(SGSTinWord15; SGSTinWord15[1])
    //             {
    //             }
    //             column(IGSTinWord15; IGSTinWord15[1])
    //             {
    //             }
    //             column(FullName_SIH; BillToCustomerName)
    //             {
    //             }
    //             column(Location; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' ' + LocFrom."Post Code" + ' ' + LocState + ' Phone: ' + LocFrom."Phone No." + ' ' + ' Email: ' + LocFrom."E-Mail")
    //             {
    //             }
    //             column(BillAdd1; "Bill-to Address")
    //             {
    //             }
    //             column(BillAdd2; "Bill-to Address 2")
    //             {
    //             }
    //             column(BillAdd3; "Bill-to City" + ' - ' + "Bill-to Post Code")
    //             {
    //             }
    //             column(SIH_BillAdd3; "Bill-to City" + '-' + "Bill-to Post Code" + ' ' + BillState + ', ' + BillCountry)
    //             {
    //             }
    //             column(BillPhoneNo; 'Tel : ' + BillPhoneNo)
    //             {
    //             }
    //             column(PlaceofSupply; "Bill-to City")
    //             {
    //             }
    //             column(ShipGSTStateCode; ShipGSTStateCode)
    //             {
    //             }
    //             column(ShipGSTNo; ShipGSTNo)
    //             {
    //             }
    //             column(ShipPANNo; ShipPANNo)
    //             {
    //             }
    //             column(Ship2Name; Ship2Name)
    //             {
    //             }
    //             column(ShiptoName15; ShiptoName15)
    //             {
    //             }
    //             column(BillGSTStateCode; BillGSTStateCode)
    //             {
    //             }
    //             column(BillGSTNo; BillGSTNo)
    //             {
    //             }
    //             column(BillPANNo; BillPANNo)
    //             {
    //             }
    //             column(SIH_ShipName; "Sales Invoice Header"."Ship-to Name")
    //             {
    //             }
    //             column(SIH_ShipAdd1; "Sales Invoice Header"."Ship-to Address")
    //             {
    //             }
    //             column(SIH_ShipAdd2; "Sales Invoice Header"."Ship-to Address 2")
    //             {
    //             }
    //             column(SIH_ShipAdd3; "Ship-to City" + '-' + "Ship-to Post Code" + ' ' + ShipState + ', ' + ShipCountry)
    //             {
    //             }
    //             column(ShipPhoneNo; 'Tel : ' + ShipPhoneNo)
    //             {
    //             }
    //             column(LocState; LocState)
    //             {
    //             }
    //             column(InvNo; InvNo)
    //             {
    //             }
    //             column(EccNos; '"E.C.C. Nos."."C.E. Registration No."')
    //             {
    //             }
    //             column(RangDivision; '')//"E.C.C. Nos."."C.E. Range" + '&' + "E.C.C. Nos."."C.E. Division" + '& ' + "E.C.C. Nos."."C.E. Commissionerate")
    //             {
    //             }
    //             column(CEAddr; '"E.C.C. Nos."."C.E Address1"')
    //             {
    //             }
    //             column(CEAddr2; '"E.C.C. Nos."."C.E Address2"')
    //             {
    //             }
    //             column(CECityPostCode; '')//"E.C.C. Nos."."C.E City" + '-' + "E.C.C. Nos."."C.E Post Code")
    //             {
    //             }
    //             column(TINNo; '')//Loc1."T.I.N. No." + '  &  ' + Loc1."C.S.T No.")
    //             {
    //             }
    //             column(PostingDate; "Posting Date")
    //             {
    //             }
    //             column(BRTTime; "Sales Invoice Header"."Invoice Print Time")
    //             {
    //             }
    //             column(No_Header; "No.")
    //             {
    //             }
    //             column(RoadPermitNo_Header; EWayBillNo)
    //             {
    //             }
    //             column(Name_ShipingAgent; recShippingAgent.Name)
    //             {
    //             }
    //             column(VehicleNo_Header; "Vehicle No.")
    //             {
    //             }
    //             column(LRRRDate_Header; "LR/RR Date")
    //             {
    //             }
    //             column(LRRRNo_Header; "LR/RR No.")
    //             {
    //             }
    //             column(FrieghtType_Header; "Freight Type")
    //             {
    //             }
    //             column(TransportType_Header; "Sales Invoice Header"."Transport Type")
    //             {
    //             }
    //             column(PaymentDescription; "Payment Terms".Description)
    //             {
    //             }
    //             column(Loc_Contact; 'Contact ' + ':- ' + Loc."Phone No.")
    //             {
    //             }
    //             column(Loc_Add1; Loc.Address)
    //             {
    //             }
    //             column(Loc_Add2; Loc."Address 2")
    //             {
    //             }
    //             column(Loc_Add3; Loc.City + ' - ' + Loc."Post Code")
    //             {
    //             }
    //             column(State; RecState.Description + ', ' + RecCountry.Name)
    //             {
    //             }
    //             column(GSTCity; Loc.City)
    //             {
    //             }
    //             column(GSTState; RecState.Description)
    //             {
    //             }
    //             column(GSTStateCode; RecState."State Code (GST Reg. No.)")
    //             {
    //             }
    //             column(Loc_CST; 'Loc."C.S.T No."')
    //             {
    //             }
    //             column(BuyerTIN; 'Loc."T.I.N. No."')
    //             {
    //             }
    //             column(Loc_GSTRegNo; Loc."GST Registration No.")
    //             {
    //             }
    //             column(Loc_WEF_CST; Loc."W.e.f. Date(C.S.T No.)")
    //             {
    //             }
    //             column(Buyer_WEF; Loc."W.e.f. Date(T.I.N No.)")
    //             {
    //             }
    //             column(BuyerECC; '"E.C.C. Nos.1"."C.E. Registration No."')
    //             {
    //             }
    //             column(Phone_Consignee; '')//'Contact ' + ':- ' + Loc."Phone No.")
    //             {
    //             }
    //             column(Loc_AddConsignee; Loc.Address)
    //             {
    //             }
    //             column(Loc_Add2_Consgnii; Loc."Address 2")
    //             {
    //             }
    //             column(City_Consignee; Loc.City + ' - ' + Loc."Post Code")
    //             {
    //             }
    //             column(State_Consginee; RecState.Description + ', ' + RecCountry.Name)
    //             {
    //             }
    //             column(CST_Con; 'Loc."C.S.T No."')
    //             {
    //             }
    //             column(TIN_Con; 'Loc."T.I.N. No."')
    //             {
    //             }
    //             column(WEF_Cons; Loc."W.e.f. Date(T.I.N No.)")
    //             {
    //             }
    //             column(CE_Cons; '"E.C.C. Nos.1"."C.E. Registration No."')
    //             {
    //             }
    //             column(Batch_UOM; "Batch/UOM")
    //             {
    //             }
    //             column(InsuranceNo_CompInfo; CompanyInfo1."Insurance No.")
    //             {
    //             }
    //             column(InsuranceProvoder_CompInfo; CompanyInfo1."Insurance Provider")
    //             {
    //             }
    //             column(Poilicy_ComInfo; FORMAT(CompanyInfo1."Policy Expiration Date"))
    //             {
    //             }
    //             column(BuyerTINDate; Loc."W.e.f. Date(T.I.N No.)")
    //             {
    //             }
    //             column(Comp_GSTRegNo; CompanyInfo1."GST Registration No.")
    //             {
    //             }
    //             column(CompWeb; CompanyInfo1."Home Page")
    //             {
    //             }
    //             column(CompCINNo; CompanyInfo1."Registration No.")
    //             {
    //             }
    //             column(CompPANNo; CompanyInfo1."P.A.N. No.")
    //             {
    //             }
    //             column(QRCode; "Sales Invoice Header"."QR code")
    //             {
    //             }
    //             column(IRN_SalesInvoiceHeader; IRN)
    //             {
    //             }
    //             column(TCSPercent; TCSPercent)
    //             {
    //             }
    //             column(TCSAmount; TCSAmount)
    //             {
    //             }
    //             column(DecPackingForwarding; DecPackingForwarding)
    //             {
    //             }
    //             dataitem(CopyLoop; Integer)
    //             {
    //                 DataItemTableView = SORTING(Number);
    //                 dataitem(Integer; Integer)
    //                 {
    //                     DataItemTableView = SORTING(Number)
    //                                         WHERE(Number = FILTER(1));
    //                     column(InvType; InvType)
    //                     {
    //                     }
    //                     column(InvType1; InvType1)
    //                     {
    //                     }
    //                     column(number; Number)
    //                     {
    //                     }
    //                     column(outPutNos; OutPutNo)
    //                     {
    //                     }
    //                     dataitem("Sales Invoice Line"; "Sales Invoice Line")
    //                     {
    //                         DataItemLink = "Document No." = FIELD("No.");
    //                         DataItemLinkReference = "Sales Invoice Header";
    //                         DataItemTableView = SORTING("Document No.", "Line No.")
    //                                             ORDER(Ascending)
    //                                             WHERE(Type = FILTER(Item),
    //                                                   Quantity = FILTER(<> 0));
    //                         column(TotalAmtLine; TotalAmt30)
    //                         {
    //                         }
    //                         column(CrossReferenceNo; '"Sales Invoice Line"."Cross-Reference No."')
    //                         {
    //                         }
    //                         column(LineDiscountAmt; "Sales Invoice Line"."Line Discount Amount")
    //                         {
    //                         }
    //                         column(GSTBaseAmt; 0)// "Sales Invoice Line"."GST Base Amount")
    //                         {
    //                         }
    //                         column(AmountToCustomer; 0)// "Sales Invoice Line"."Amount To Customer")
    //                         {
    //                         }
    //                         column(ChargesToCustomer; 0)// "Sales Invoice Line"."Charges To Customer")
    //                         {
    //                         }
    //                         column(GSTPer; 0)//"GST %")
    //                         {
    //                         }
    //                         column(GSTGroupCode_Line; "GST Group Code")
    //                         {
    //                         }
    //                         column(HSNSACCode_Line; "HSN/SAC Code")
    //                         {
    //                         }
    //                         column(LineNo; "Line No.")
    //                         {
    //                         }
    //                         column(DimensionName; DimensionName)
    //                         {
    //                         }
    //                         column(LineNo_Line; "Line No.")
    //                         {
    //                         }
    //                         column(DocumentNo_Line; "Document No.")
    //                         {
    //                         }
    //                         column(BatchNo; "Lot No.")
    //                         {
    //                         }
    //                         column(QtyperUnitofMeasure_Line; QtyPerPack)
    //                         {
    //                         }
    //                         column(QuantityBase_Line; TotalQtyFO)
    //                         {
    //                         }
    //                         column(Quantity_LineN; Quantity_LineN)
    //                         {
    //                         }
    //                         column(Quantity_Line; Quantity)
    //                         {
    //                         }
    //                         column(RatePerMTLTKG; RatePerMTLTKG)
    //                         {
    //                         }
    //                         column(TotalUnitPrice; Quantity * "Unit Price")
    //                         {
    //                         }
    //                         column(UnitPrice_TSL; "Unit Price")
    //                         {
    //                         }
    //                         column(BaseUnitOfMeasure; UOMCode)
    //                         {
    //                         }
    //                         column(Unitprice_QtyPerUnit; BasicRate)
    //                         {
    //                         }
    //                         column(ExciseProdPostingGroup_Line; '"Excise Prod. Posting Group"')
    //                         {
    //                         }
    //                         column(BEDpercent; BEDpercent)
    //                         {
    //                         }
    //                         column(BEDant_QtyBase; 0)//"BED Amount" / "Quantity (Base)")
    //                         {
    //                         }
    //                         column(BEDAmount_Line; 0)// "BED Amount")
    //                         {
    //                         }
    //                         column(SHE_Cess; 0)//"SHE Cess Amount" + "eCess Amount")
    //                         {
    //                         }
    //                         column(Amount_Line; Amount)
    //                         {
    //                         }
    //                         column(DescriptionLineDuty_1; DescriptionLineDuty[1])
    //                         {
    //                         }
    //                         column(DescriptionLineeCess_1; DescriptionLineeCess[1])
    //                         {
    //                         }
    //                         column(DescriptionLineSHeCess_1; DescriptionLineSHeCess[1])
    //                         {
    //                         }
    //                         column(DescriptionLineAD_1; DescriptionLineAD[1])
    //                         {
    //                         }
    //                         column(DescriptionLineTot_1; DescriptionLineTot[1])
    //                         {
    //                         }
    //                         column(BEDAmount_Line1; 0)// "BED Amount")
    //                         {
    //                         }
    //                         column(eCessAmount_Line; 0)// "eCess Amount")
    //                         {
    //                         }
    //                         column(ADEAmount_Line; 0)// "ADE Amount")
    //                         {
    //                         }
    //                         column(SHECessAmount_Line; 0)// "SHE Cess Amount")
    //                         {
    //                         }
    //                         column(FrieghtAmount; FrieghtAmount)
    //                         {
    //                         }
    //                         column(DocTotal; DocTotal)
    //                         {
    //                         }
    //                         column(Commente; recInvComment.Comment)
    //                         {
    //                         }
    //                         column(EcessPer; FORMAT(ecesspercent) + '%')
    //                         {
    //                         }
    //                         column(SheCessPer; FORMAT(SHEpercent) + '%')
    //                         {
    //                         }
    //                         column(FormCodeDesc; 'FormCodes.Description')
    //                         {
    //                         }
    //                         column(vCount; vCount)
    //                         {
    //                         }
    //                         column(CGST06; vCGST)
    //                         {
    //                         }
    //                         column(CGST06Rate; vCGSTRate)
    //                         {
    //                         }
    //                         column(SGST06; vSGST)
    //                         {
    //                         }
    //                         column(SGST06Rate; vSGSTRate)
    //                         {
    //                         }
    //                         column(IGST06; vIGST)
    //                         {
    //                         }
    //                         column(IGST06Rate; vIGSTRate)
    //                         {
    //                         }
    //                         column(DiscAmt06; DiscAmt06)
    //                         {
    //                         }
    //                         column(ChargeAmt06; ChargeAmt06)
    //                         {
    //                         }
    //                         column(CGST15; CGST15)
    //                         {
    //                         }
    //                         column(SGST15; SGST15)
    //                         {
    //                         }
    //                         column(IGST15; IGST15)
    //                         {
    //                         }
    //                         column(PCharge; PCharge)
    //                         {
    //                         }
    //                         column(NCharge; NCharge)
    //                         {
    //                         }
    //                         column(PKGFWDAmount; PKGFWDAmount)
    //                         {
    //                         }
    //                         column(FREIGHTAmount; FREIGHTAmount)
    //                         {
    //                         }

    //                         trigger OnAfterGetRecord()
    //                         begin
    //                             vCount := COUNT;
    //                             Ctr := Ctr + 1;



    //                             //RSPL
    //                             //DimensionName:='';
    //                             //Dim table not avialble
    //                             /*
    //                             IF "Inventory Posting Group"='MERCH' THEN BEGIN
    //                              recPosteddocDiemension.RESET;
    //                              recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Table ID",5745);
    //                              recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Document No.","Transfer Shipment Line"."Document No.");
    //                              recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Line No.","Transfer Shipment Line"."Line No.");
    //                              recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Dimension Code",'MERCH');
    //                               IF recPosteddocDiemension.FINDFIRST THEN;
    //                                  dimensionCode:=recPosteddocDiemension."Dimension Value Code";
    //                                //IF recDimensionValue.GET(dimensionCode) THEN
    //                                //  DimensionName:=recDimensionValue.Name;
    //                                recDimensionValue.RESET;
    //                                recDimensionValue.SETRANGE(recDimensionValue.Code,dimensionCode);
    //                                IF recDimensionValue.FINDFIRST THEN
    //                                   DimensionName:=recDimensionValue.Name;
    //                                   DimensionName2:=DimensionName;
    //                             END;
    //                             */

    //                             //>>14Apr2017 DimsetEntry

    //                             DimensionName := '';
    //                             IF "Inventory Posting Group" = 'MERCH' THEN BEGIN

    //                                 recDimSet.RESET;
    //                                 recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
    //                                 recDimSet.SETRANGE("Dimension Code", 'MERCH');
    //                                 IF recDimSet.FINDFIRST THEN BEGIN

    //                                     recDimSet.CALCFIELDS("Dimension Value Name");
    //                                     DimensionName := recDimSet."Dimension Value Name";

    //                                 END;
    //                                 DimensionName2 := DimensionName;
    //                             END;
    //                             //<<14Apr2017 DimsetEntry


    //                             IF (DimensionName = '') OR (recItemEBT."Item Category Code" = 'CAT02') OR (recItemEBT."Item Category Code" = 'CAT03') OR
    //                               (recItemEBT."Item Category Code" = 'CAT11') OR (recItemEBT."Item Category Code" = 'CAT12') OR
    //                                 (recItemEBT."Item Category Code" = 'CAT13') THEN
    //                                 IF (DimensionName = '') AND ("Item Category Code" <> 'CAT01')
    //                                     AND ("Item Category Code" <> 'CAT04') AND ("Item Category Code" <> 'CAT05')
    //                                     AND ("Item Category Code" <> 'CAT06') AND ("Item Category Code" <> 'CAT07')
    //                                     AND ("Item Category Code" <> 'CAT08') AND ("Item Category Code" <> 'CAT09')
    //                                     AND ("Item Category Code" <> 'CAT10') AND ("Item Category Code" <> 'CAT17') THEN BEGIN
    //                                     recILE.SETRANGE("Document No.", "Document No.");
    //                                     recILE.SETRANGE("Document Line No.", "Line No.");
    //                                     recILE.SETRANGE("Item No.", "Sales Invoice Line"."No.");
    //                                     IF recILE.FINDFIRST THEN
    //                                         BatchNo := recILE."Lot No.";
    //                                     DimensionName := 'IPOL ' + Description;
    //                                     //ItemDescription := 'IPOL ' +"Transfer Shipment Line".Description;
    //                                 END ELSE BEGIN
    //                                     BatchNo := "Unit of Measure Code";

    //                                     //DimensionName:="Transfer Shipment Line".Description;
    //                                     IF DimensionName <> '' THEN
    //                                         DimensionName := DimensionName2
    //                                     ELSE
    //                                         DimensionName := Description;
    //                                     //  ItemDescription := "Transfer Shipment Line".Description;
    //                                 END;

    //                             IF "Sales Invoice Line".Type = "Sales Invoice Line".Type::"G/L Account" THEN
    //                                 DimensionName := Description;

    //                             //>>21Aug2017 RepsolName

    //                             IF "Item Category Code" = 'CAT15' THEN
    //                                 DimensionName := Description;

    //                             IF "Item Category Code" = 'CAT14' THEN
    //                                 DimensionName := Description;


    //                             IF "Item Category Code" = 'CAT18' THEN
    //                                 DimensionName := Description;

    //                             //>>21Aug2017 RepsolName
    //                             //>>14Apr2017

    //                             /* recExcisePostnSetup.RESET;
    //                             recExcisePostnSetup.SETRANGE(recExcisePostnSetup."Excise Bus. Posting Group", "Excise Bus. Posting Group");
    //                             recExcisePostnSetup.SETFILTER(recExcisePostnSetup."From Date", '<=%1', "Posting Date"); //RSPL008
    //                             IF recExcisePostnSetup.FINDLAST THEN BEGIN
    //                                 ecesspercent := recExcisePostnSetup."eCess %";
    //                                 BEDpercent := recExcisePostnSetup."BED %";
    //                                 SHEpercent := recExcisePostnSetup."SHE Cess %";
    //                             END; */

    //                             /*
    //                             RecState.SETRANGE(RecState.Code,Loc."State Code");
    //                             RecState.FINDFIRST;
    //                             State:=RecState.Description;
    //                             */

    //                             IF RecItem.GET("No.") THEN;

    //                             //RSPLSUM 12Jan21>>
    //                             CLEAR(TotalQtyFO);
    //                             CLEAR(UOMCode);
    //                             CLEAR(BasicRate);//RSPLSUM 19Jan21
    //                             IF ("Sales Invoice Line"."Item Category Code" = 'CAT05') OR ("Sales Invoice Line"."Item Category Code" = 'CAT17') THEN BEGIN
    //                                 UOMCode := "Sales Invoice Line"."Unit of Measure Code";
    //                                 QtyPerPack := 1;
    //                                 TotalQtyFO := "Sales Invoice Line".Quantity;
    //                                 BasicRate := "Unit Price";//RSPLSUM 19Jan21
    //                                 Quantity_LineN := 0;
    //                                 //RatePerMTLTKG := 'Basic Rate per MT';
    //                             END ELSE BEGIN
    //                                 UOMCode := RecItem."Base Unit of Measure";
    //                                 QtyPerPack := "Sales Invoice Line"."Qty. per Unit of Measure";
    //                                 TotalQtyFO := "Sales Invoice Line"."Quantity (Base)";
    //                                 BasicRate := "Unit Price" / "Qty. per Unit of Measure";//RSPLSUM 19Jan21
    //                                 Quantity_LineN := "Sales Invoice Line".Quantity;
    //                             END;
    //                             //RSPLSUM 12Jan21<<

    //                             IF ("Sales Invoice Line"."Item Category Code" = 'CAT17') THEN BEGIN
    //                                 RatePerMTLTKG := 'Basic Rate per MT';
    //                             END ELSE BEGIN
    //                                 RatePerMTLTKG := 'Basic Rate per Lt/Kg';
    //                             END;

    //                             TotalQty += "Quantity (Base)";
    //                             TotalBR += "Unit Price";
    //                             TotalValue := TotalValue + ("Quantity (Base)" * "Unit Price");


    //                             TaxDescr := 'Tax Descr';

    //                             //>>06APr2017

    //                             FExcise += 0;// "BED Amount";
    //                             FEcess += 0;// "eCess Amount";
    //                             FShe += 0;// "SHE Cess Amount";
    //                             FInvT += Amount;

    //                             //<<06APr2017

    //                             //>>06July2017 GST
    //                             CLEAR(vCGST);
    //                             CLEAR(vSGST);
    //                             CLEAR(vIGST);
    //                             CLEAR(vCGSTRate);
    //                             CLEAR(vSGSTRate);
    //                             CLEAR(vIGSTRate);

    //                             //>>15July2017
    //                             CLEAR(CGST15);
    //                             CLEAR(SGST15);
    //                             CLEAR(IGST15);
    //                             DetailGSTEntry.RESET;
    //                             DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
    //                             DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Sales Invoice Header"."No.");
    //                             DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
    //                             DetailGSTEntry.SETRANGE(Type, DetailGSTEntry.Type::Item);
    //                             DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
    //                             IF DetailGSTEntry.FINDSET THEN
    //                                 REPEAT

    //                                     IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
    //                                         vCGST := ABS(DetailGSTEntry."GST Amount");
    //                                         vCGSTRate := DetailGSTEntry."GST %";

    //                                         IF DetailGSTEntry."GST Amount" = 0 THEN
    //                                             CGST15 := ''
    //                                         ELSE
    //                                             CGST15 := ' @ ' + FORMAT(vCGSTRate) + ' %';
    //                                     END;

    //                                     IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
    //                                         vSGST := ABS(DetailGSTEntry."GST Amount");
    //                                         vSGSTRate := DetailGSTEntry."GST %";

    //                                         IF DetailGSTEntry."GST Amount" = 0 THEN
    //                                             SGST15 := ''
    //                                         ELSE
    //                                             SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

    //                                     END;

    //                                     //>>21July2017 UTGST
    //                                     IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
    //                                         vSGST := ABS(DetailGSTEntry."GST Amount");
    //                                         vSGSTRate := DetailGSTEntry."GST %";

    //                                         IF DetailGSTEntry."GST Amount" = 0 THEN
    //                                             SGST15 := ''
    //                                         ELSE
    //                                             SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

    //                                     END;
    //                                     //<<21July2017 UTGST

    //                                     IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
    //                                         vIGST := ABS(DetailGSTEntry."GST Amount");
    //                                         vIGSTRate := DetailGSTEntry."GST %";
    //                                         //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");

    //                                         IF DetailGSTEntry."GST Amount" = 0 THEN
    //                                             IGST15 := ''
    //                                         ELSE
    //                                             IGST15 := ' @ ' + FORMAT(vIGSTRate) + ' %';

    //                                     END;

    //                                 UNTIL DetailGSTEntry.NEXT = 0;
    //                             //<<06July2017 GST


    //                             /*
    //                             IF "Total GST Amount" = 0 THEN
    //                             BEGIN
    //                               CGST15 := '';
    //                               SGST15 := '';
    //                               IGST15 := '';
    //                             END;

    //                             IF "Total GST Amount" <> 0 THEN
    //                             BEGIN
    //                               CGST15 := ' @ '+FORMAT(vCGSTRate)+' %';
    //                               SGST15 := ' @ '+FORMAT(vSGSTRate)+' %';
    //                               IGST15 := ' @ '+FORMAT(vIGSTRate)+' %';
    //                             END;
    //                             */
    //                             //<<15July2017

    //                             //>>06July2017 StructureDetails
    //                             CLEAR(DiscAmt06);
    //                             CLEAR(ChargeAmt06);

    //                             /* recStrOrderLineDet.RESET;
    //                             recStrOrderLineDet.SETRANGE("Invoice No.", "Document No.");
    //                             recStrOrderLineDet.SETRANGE("Item No.", "No.");
    //                             recStrOrderLineDet.SETRANGE("Line No.", "Line No.");
    //                             IF recStrOrderLineDet.FINDSET THEN
    //                                 REPEAT

    //                                     IF recStrOrderLineDet."Tax/Charge Group" = 'DISCOUNT' THEN BEGIN

    //                                         DiscAmt06 += recStrOrderLineDet.Amount;

    //                                     END;

    //                                     //>>AVPDiscount 03Aug2017
    //                                     IF recStrOrderLineDet."Tax/Charge Group" = 'AVP-SP-SAN' THEN BEGIN
    //                                         DiscAmt06 += recStrOrderLineDet.Amount;
    //                                     END;

    //                                     //>>TRANSSUBS 03Aug2017
    //                                     IF recStrOrderLineDet."Tax/Charge Group" = 'TRANSSUBS' THEN BEGIN
    //                                         DiscAmt06 += recStrOrderLineDet.Amount;
    //                                     END;

    //                                     //>>CashDiscount 03Aug2017
    //                                     IF recStrOrderLineDet."Tax/Charge Group" = 'CASHDISC' THEN BEGIN
    //                                         DiscAmt06 += recStrOrderLineDet.Amount;
    //                                     END;

    //                                     //>>02Nov2018
    //                                     IF recStrOrderLineDet."Tax/Charge Group" = 'NATSCHEME' THEN
    //                                         DiscAmt06 += recStrOrderLineDet.Amount;

    //                                     IF recStrOrderLineDet."Tax/Charge Group" = 'REGSCHEME' THEN
    //                                         DiscAmt06 += recStrOrderLineDet.Amount;
    //                                     //<<02Nov2018

    //                                     IF recStrOrderLineDet."Tax/Charge Group" = 'FREIGHT' THEN BEGIN

    //                                         recStrOrderLineDet.CALCFIELDS("Structure Description");//21Aug2017
    //                                         IF recStrOrderLineDet."Structure Description" = 'Freight (Taxable)' THEN//21Aug2017
    //                                             ChargeAmt06 += recStrOrderLineDet.Amount;

    //                                     END;

    //                                 UNTIL recStrOrderLineDet.NEXT = 0; */
    //                             //<<06July2017 StrucureDetails

    //                             //RSPLAM12072021  --
    //                             CLEAR(PKGFWDAmount);
    //                             /*  recStrOrderLineDet.RESET;
    //                              recStrOrderLineDet.SETRANGE("Invoice No.", "Document No.");
    //                              recStrOrderLineDet.SETRANGE("Item No.", "No.");
    //                              recStrOrderLineDet.SETRANGE("Line No.", "Line No.");
    //                              recStrOrderLineDet.SETRANGE(Type, recStrOrderLineDet.Type::Sale);
    //                              recStrOrderLineDet.SETFILTER("Tax/Charge Group", 'PKGFWD');
    //                              IF recStrOrderLineDet.FINDSET THEN
    //                                  REPEAT
    //                                      PKGFWDAmount += recStrOrderLineDet."Amount (LCY)";
    //                                  UNTIL recStrOrderLineDet.NEXT = 0; */

    //                             CLEAR(FREIGHTAmount);
    //                             /*  recStrOrderLineDet.RESET;
    //                              recStrOrderLineDet.SETRANGE("Invoice No.", "Document No.");
    //                              recStrOrderLineDet.SETRANGE("Item No.", "No.");
    //                              recStrOrderLineDet.SETRANGE("Line No.", "Line No.");
    //                              recStrOrderLineDet.SETRANGE(Type, recStrOrderLineDet.Type::Sale);
    //                              recStrOrderLineDet.SETFILTER("Tax/Charge Group", 'FREIGHT');
    //                              IF recStrOrderLineDet.FINDSET THEN
    //                                  REPEAT
    //                                      FREIGHTAmount += recStrOrderLineDet."Amount (LCY)";
    //                                  UNTIL recStrOrderLineDet.NEXT = 0; */
    //                             //RSPLAM12072021  ++

    //                             //>>30Nov2017
    //                             /* TotalAmt30 := 0;
    //                             IF "GST Base Amount" = 0 THEN
    //                                 TotalAmt30 := "Line Amount" - (FREIGHTAmount + PKGFWDAmount)
    //                             ELSE
    //                                 TotalAmt30 := "GST Base Amount" - (FREIGHTAmount + PKGFWDAmount); */
    //                             //<<30Nov2017

    //                             //>>17May2018
    //                             /*  IF "Amount To Customer" = 0 THEN
    //                                  TotalAmt30 := 0; */
    //                             //<<17May2018

    //                             //>>21Aug2017 Charges
    //                             CLEAR(PCharge);
    //                             CLEAR(NCharge);
    //                             /* 
    //                                                         IF "Sales Invoice Line"."Charges To Customer" > 0 THEN
    //                                                             PCharge := "Charges To Customer"
    //                                                         ELSE
    //                                                             NCharge := "Charges To Customer"; */

    //                             //<<21Aug2017

    //                             DocTotal += Amount + 0;// "BED Amount" + "eCess Amount"
    //                                                    //  + "SHE Cess Amount" + FrieghtAmount + "ADE Amount" + "Total GST Amount";

    //                             /*
    //                             DocTotal:=ROUND((
    //                             "Transfer Shipment Line".Amount + "Transfer Shipment Line"."BED Amount"+"Transfer Shipment Line"."eCess Amount"
    //                             +"Transfer Shipment Line"."SHE Cess Amount"+FrieghtAmount+"Transfer Shipment Line"."ADE Amount"+RoundOffAmnt),1.0,'=');
    //                             */

    //                             /*  FormCodes.SETRANGE(FormCodes.Code, "Form Code");
    //                              IF FormCodes.FINDFIRST THEN; */

    //                             vCheck.InitTextVariable;
    //                             vCheck.FormatNoText(DescriptionLineDuty, ROUND(FExcise, 0.01), ''); //06Apr2017
    //                             //vCheck.FormatNoText(DescriptionLineDuty,ROUND("BED Amount",0.01),'');
    //                             DescriptionLineDuty[1] := DELCHR(DescriptionLineDuty[1], '=', '*');

    //                             vCheck.InitTextVariable;
    //                             vCheck.FormatNoText(DescriptionLineeCess, ROUND(FEcess, 0.01), ''); //06Apr2017
    //                             //vCheck.FormatNoText(DescriptionLineeCess,ROUND("eCess Amount",0.01),'');
    //                             DescriptionLineeCess[1] := DELCHR(DescriptionLineeCess[1], '=', '*');

    //                             vCheck.InitTextVariable;
    //                             vCheck.FormatNoText(DescriptionLineSHeCess, ROUND(FShe, 0.01), ''); //06Apr2017
    //                             //vCheck.FormatNoText(DescriptionLineSHeCess,ROUND("SHE Cess Amount",0.01),'');
    //                             DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1], '=', '*');

    //                             /*
    //                             vCheck.InitTextVariable;
    //                             vCheck.FormatNoText(DescriptionLineSHeCess,ROUND("SHE Cess Amount",0.01),'');
    //                             DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1],'=','*');
    //                             */

    //                             vCheck.InitTextVariable;
    //                             vCheck.FormatNoText(DescriptionLineTot, ROUND((DocTotal), 1.0), '');
    //                             DescriptionLineTot[1] := DELCHR(DescriptionLineTot[1], '=', '*');


    //                             vCheck.InitTextVariable;
    //                             vCheck.FormatNoText(DescriptionLineAD, ROUND("ADE Amount", 0.01), '');
    //                             DescriptionLineAD[1] := DELCHR(DescriptionLineAD[1], '=', '*');

    //                         end;
    //                     }

    //                     trigger OnAfterGetRecord()
    //                     begin

    //                         DocTotal := 0; //06Apr2017
    //                         FExcise := 0; //06Apr2017
    //                         FEcess := 0; //06Apr2017
    //                         FShe := 0; //06Apr2017
    //                         FInvT := 0; //06Apr2017


    //                         i += 1;
    //                         //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START
    //                         IF ("Sales Invoice Header"."Print Invoice" = FALSE) THEN BEGIN
    //                             IF i = 1 THEN
    //                                 InvType := 'ORIGINAL FOR RECIPIENT'
    //                             //InvType := 'ORIGINAL FOR BUYER'
    //                             ELSE
    //                                 IF i = 2 THEN
    //                                     InvType := 'DUPLICATE FOR TRANSPORTER'
    //                                 ELSE
    //                                     IF i = 3 THEN
    //                                         InvType := 'TRIPLICATE FOR SUPPLIER'
    //                                     //InvType := 'TRIPLICATE FOR ASSESSEE'
    //                                     ELSE
    //                                         IF i = 4 THEN BEGIN
    //                                             InvType := 'QUADRAPLICATE FOR H.O';
    //                                             //InvType := 'QUADRAPLICATE FOR ACCOUNTS';
    //                                             //InvType1 := 'NOT FOR CENVAT';
    //                                         END ELSE
    //                                             IF i = 5 THEN BEGIN
    //                                                 InvType := 'EXTRA COPY';
    //                                                 //InvType := 'EXTRA COPY NOT FOR CENVAT';
    //                                                 InvType1 := '';
    //                                             END
    //                                             ELSE
    //                                                 InvType := '';
    //                         END;

    //                         //IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'sa') THEN
    //                         IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'GPUAE\UNNIKRISHNAN.VS') THEN //18July2017
    //                         BEGIN
    //                             IF i = 1 THEN
    //                                 InvType := 'ORIGINAL FOR RECIPIENT'
    //                             //InvType := 'ORIGINAL FOR BUYER'
    //                             //InvType := 'ORIGINAL FOR BUYER'
    //                             ELSE
    //                                 IF i = 2 THEN
    //                                     InvType := 'DUPLICATE FOR TRANSPORTER'
    //                                 ELSE
    //                                     IF i = 3 THEN
    //                                         InvType := 'TRIPLICATE FOR SUPPLIER'
    //                                     //InvType := 'TRIPLICATE FOR ASSESSEE'
    //                                     ELSE
    //                                         IF i = 4 THEN BEGIN
    //                                             InvType := 'QUADRAPLICATE FOR H.O';
    //                                             //InvType := 'QUADRAPLICATE FOR ACCOUNTS';
    //                                             //InvType1 := 'NOT FOR CENVAT';
    //                                         END ELSE
    //                                             IF i = 5 THEN BEGIN
    //                                                 InvType := 'EXTRA COPY NOT FOR CENVAT';
    //                                                 InvType1 := '';
    //                                             END
    //                                             ELSE
    //                                                 InvType := '';
    //                         END;

    //                         //IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'sa')THEN
    //                         IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'GPUAE\UNNIKRISHNAN.VS') THEN //18July2017
    //                         BEGIN
    //                             IF i = 1 THEN
    //                                 InvType := 'EXTRA COPY';
    //                             InvType1 := ' NOT FOR COMMERCIAL USE';
    //                         END;

    //                         //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
    //                     end;
    //                 }

    //                 trigger OnAfterGetRecord()
    //                 begin

    //                     OutPutNo += 1;
    //                     IF Number > 1 THEN
    //                         copytext := Text003;
    //                     //CurrReport.PAGENO := 1;
    //                 end;

    //                 trigger OnPreDataItem()
    //                 begin

    //                     DocTotal := 0; //06Apr2017
    //                     //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START

    //                     //IF NOT(USERID = 'sa') THEN
    //                     IF NOT (USERID = 'GPUAE\FAHIM.AHMAD') THEN BEGIN
    //                         IF "Sales Invoice Header"."Print Invoice" = TRUE THEN BEGIN
    //                             SETRANGE(Number, 1, 1);
    //                             //  SETRANGE(Number,1,ABS(NoOfCopies));
    //                         END ELSE BEGIN

    //                             //NoOfCopies:=4;
    //                             NoOfCopies := 1; //Fahim 17-01-23
    //                             NoOfLoops := ABS(NoOfCopies);//06July2017
    //                                                          //NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
    //                             IF NoOfLoops <= 0 THEN
    //                                 NoOfLoops := 1;
    //                             copytext := '';
    //                             SETRANGE(Number, 1, NoOfLoops);
    //                         END;
    //                     END ELSE BEGIN

    //                         //NoOfCopies:=4;
    //                         NoOfCopies := 1; //Fahim 17-01-23
    //                         NoOfLoops := ABS(NoOfCopies);//06July2017
    //                                                      //NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
    //                         IF NoOfLoops <= 0 THEN
    //                             NoOfLoops := 1;
    //                         copytext := '';
    //                         SETRANGE(Number, 1, NoOfLoops);
    //                     END;

    //                     //SETRANGE(Number,1,NoOfCopies);
    //                     //MESSAGE('%1',ABS(NoOfCopies));
    //                     //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
    //                     //i:= 1
    //                 end;
    //             }

    //             trigger OnAfterGetRecord()
    //             begin
    //                 IF NOT CurrReport.PREVIEW THEN;//RB-N 14Nov2017

    //                 //>>01Aug2017
    //                 IF "Sales Invoice Header"."Posting Date" < 20170108D THEN //010817D
    //                     ERROR('This Report is Applicable from August2017 Invoices');
    //                 //>>01Aug2017

    //                 //>>09Aug2017
    //                 SuppliersCode := '';
    //                 IF Cust09.GET("Sell-to Customer No.") THEN
    //                     SuppliersCode := Cust09."Our Account No.";
    //                 //<<09Aug2017

    //                 //>>02Aug2017 FactoryCaption
    //                 CLEAR(FactoryCaption);
    //                 IF ("Location Code" = 'PLANT01') OR ("Location Code" = 'PLANT02') THEN
    //                     FactoryCaption := 'Factory'
    //                 ELSE
    //                     FactoryCaption := 'Matrl. Sold from';
    //                 //<<02Aug2017 FactoryCaption

    //                 //>>06Apr2017
    //                 IF LocFrom.GET("Sales Invoice Header"."Location Code") THEN BEGIN
    //                     IF RecState.GET(LocFrom."State Code") THEN
    //                         LocState := RecState.Description;
    //                 END;
    //                 //<<06Apr2017

    //                 //>>17July2017 PaymentTerm
    //                 "Payment Terms".RESET;
    //                 IF "Payment Terms".GET("Sales Invoice Header"."Payment Terms Code") THEN;
    //                 //<<17July2017 PaymentTerm

    //                 InvNoLen := STRLEN("No.");
    //                 InvNo := COPYSTR("No.", (InvNoLen - 3), 4);

    //                 CurrDatetime := TIME;
    //                 TotalQty := 0;

    //                 CLEAR(FrieghtAmount);
    //                 CLEAR(AdDutyAmount);
    //                 /* PostedStrOrdrDetails.RESET;
    //                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
    //                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
    //                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'FREIGHT');
    //                 IF PostedStrOrdrDetails.FINDFIRST THEN
    //                     FrieghtAmount := PostedStrOrdrDetails."Calculation Value"; */

    //                 // EBT Start 10.08.10 start
    //                 /*   PostedStrOrdrDetails.RESET;
    //                   PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
    //                   PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails.Type, PostedStrOrdrDetails.Type::Transfer);
    //                   PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'ADDL.DUTY');
    //                   IF PostedStrOrdrDetails.FINDFIRST THEN
    //                       AdDutyAmount := PostedStrOrdrDetails."Calculation Value"; */
    //                 // EBT Start 10.08.10 end

    //                 recShippingAgent.RESET;
    //                 recShippingAgent.SETRANGE(recShippingAgent.Code, "Shipping Agent Code");
    //                 iF recShippingAgent.FINDFIRST THEN;

    //                 //RSPLSUM-TCS>>
    //                 CLEAR(TCSAmount);
    //                 SL14.RESET;
    //                 SL14.SETCURRENTKEY("Document No.", Type, Quantity);
    //                 SL14.SETRANGE("Document No.", "No.");
    //                 SL14.SETFILTER(Quantity, '<>%1', 0);
    //                 IF SL14.FINDSET THEN
    //                     REPEAT
    //                         TCSAmount += 0;// SL14."TDS/TCS Amount";
    //                     UNTIL SL14.NEXT = 0;
    //                 //RSPLSUM-TCS<<

    //                 //>>14July2017 Roundoff Amount
    //                 CLEAR(RoundOffAmnt);
    //                 SL14.RESET;
    //                 SL14.SETRANGE("Document No.", "No.");
    //                 SL14.SETRANGE(Type, SL14.Type::"G/L Account");
    //                 SL14.SETRANGE("No.", '74012210');
    //                 IF SL14.FINDFIRST THEN BEGIN
    //                     RoundOffAmnt := SL14.Amount;
    //                 END;

    //                 CLEAR(FinalAmt);
    //                 SL14.RESET;
    //                 SL14.SETRANGE("Document No.", "No.");
    //                 SL14.SETRANGE(Type, SL14.Type::Item);
    //                 IF SL14.FINDSET THEN
    //                     REPEAT
    //                         FinalAmt += 0;// SL14."Amount To Customer";
    //                     //MESSAGE(' %1 Final Amt \\  %2 LineNo.',FinalAmt,SL14."Line No.");
    //                     UNTIL SL14.NEXT = 0;

    //                 //<<14July2017 Roundoff Amount

    //                 //>>AmountinWords
    //                 vCheck.InitTextVariable;
    //                 vCheck.FormatNoText(AmtinWord15, ROUND((FinalAmt + RoundOffAmnt), 0.01), '');
    //                 AmtinWord15[1] := DELCHR(AmtinWord15[1], '=', '*');

    //                 //<<AmountinWords

    //                 //>>17July2017TotalGST Amount
    //                 CLEAR(TotalCGST);
    //                 CLEAR(TotalSGST);
    //                 CLEAR(TotalIGST);
    //                 DGST04.RESET;
    //                 DGST04.SETRANGE("Document No.", "No.");
    //                 IF DGST04.FINDSET THEN
    //                     REPEAT
    //                         IF DGST04."GST Component Code" = 'CGST' THEN BEGIN
    //                             TotalCGST += ABS(DGST04."GST Amount");
    //                         END;
    //                         IF DGST04."GST Component Code" = 'SGST' THEN BEGIN
    //                             TotalSGST += ABS(DGST04."GST Amount");
    //                         END;

    //                         //>>21July2017 UTGST
    //                         IF DGST04."GST Component Code" = 'UTGST' THEN BEGIN
    //                             TotalSGST += ABS(DGST04."GST Amount");
    //                         END;

    //                         //<<21July2017 UTGST
    //                         IF DGST04."GST Component Code" = 'IGST' THEN BEGIN
    //                             TotalIGST += ABS(DGST04."GST Amount");
    //                         END;
    //                     UNTIL DGST04.NEXT = 0;

    //                 //>>CGST AmountinWords
    //                 vCheck.InitTextVariable;
    //                 //vCheck.FormatNoText(CGSTinWord15,ROUND((TotalCGST),1.0),'');
    //                 vCheck.FormatNoText(CGSTinWord15, ROUND((TotalCGST), 0.01), '');
    //                 CGSTinWord15[1] := DELCHR(CGSTinWord15[1], '=', '*');

    //                 //>>SGST AmountinWords
    //                 vCheck.InitTextVariable;
    //                 vCheck.FormatNoText(SGSTinWord15, ROUND((TotalSGST), 0.01), '');
    //                 SGSTinWord15[1] := DELCHR(SGSTinWord15[1], '=', '*');

    //                 //>>IGST AmountinWords
    //                 vCheck.InitTextVariable;
    //                 vCheck.FormatNoText(IGSTinWord15, ROUND((TotalIGST), 0.01), '');
    //                 IGSTinWord15[1] := DELCHR(IGSTinWord15[1], '=', '*');

    //                 //<<17July2017TotalGST Amount


    //                 //>>15Jul2017
    //                 //Bill State Description
    //                 RecState.RESET;
    //                 //IF RecState.GET("Bill to Customer State") THEN
    //                 IF RecState.GET(State) THEN //09Nov2017
    //                 BEGIN
    //                     BillState := RecState.Description;
    //                     BillGSTStateCode := RecState."State Code (GST Reg. No.)";
    //                 END;


    //                 //Bill GSTNo
    //                 Cus12.RESET;
    //                 IF Cus12.GET("Sales Invoice Header"."Bill-to Customer No.") THEN
    //                     BillGSTNo := Cus12."GST Registration No.";
    //                 BillPANNo := Cus12."P.A.N. No.";
    //                 BillPhoneNo := Cus12."Phone No.";//03Sep2019

    //                 //>>24May2019
    //                 DGST04.RESET;
    //                 DGST04.SETCURRENTKEY("Document No.");
    //                 DGST04.SETRANGE("Document No.", "No.");
    //                 IF DGST04.FINDFIRST THEN
    //                     BillGSTNo := DGST04."Buyer/Seller Reg. No.";
    //                 //<<24May2019

    //                 //Bill Country Name
    //                 RecCountry.RESET;
    //                 IF RecCountry.GET("Bill-to Country/Region Code") THEN
    //                     BillCountry := RecCountry.Name;

    //                 //Ship Country Name
    //                 RecCountry.RESET;
    //                 IF RecCountry.GET("Ship-to Country/Region Code") THEN
    //                     ShipCountry := RecCountry.Name;


    //                 //Ship2Name
    //                 IF "Sales Invoice Header"."Ship-to Code" = '' THEN
    //                     Ship2Name := "Sales Invoice Header"."Full Name"
    //                 ELSE
    //                     Ship2Name := "Sales Invoice Header"."Ship-to Name";

    //                 //Shipping
    //                 IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code") THEN BEGIN
    //                     RecState.RESET;
    //                     IF RecState.GET(ShipToCode.State) THEN BEGIN
    //                         ShipState := RecState.Description;
    //                         ShipGSTStateCode := RecState."State Code (GST Reg. No.)";
    //                         ShipGSTNo := ShipToCode."GST Registration No.";
    //                         ShipPANNo := ShipToCode."P.A.N. No.";
    //                         ShipPhoneNo := ShipToCode."Phone No.";//03Sep2019
    //                     END;
    //                 END;

    //                 IF "Ship-to Code" = '' THEN BEGIN
    //                     ShipState := BillState;
    //                     ShipGSTStateCode := BillGSTStateCode;
    //                     ShipGSTNo := BillGSTNo;
    //                     ShipPANNo := BillPANNo;
    //                     ShipPhoneNo := BillPhoneNo;//03S
    //                 END;


    //                 //>>15Jul2017

    //                 //RSPLSUM 30Jul2020>>
    //                 CLEAR(EWayBillNo);
    //                 "Sales Invoice Header".CALCFIELDS("Sales Invoice Header"."EWB No.");
    //                 IF "Sales Invoice Header"."EWB No." <> '' THEN BEGIN
    //                     EWayBillNo := "Sales Invoice Header"."EWB No.";
    //                 END ELSE BEGIN
    //                     EWayBillNo := "Sales Invoice Header"."Road Permit No.";
    //                 END;
    //                 //RSPLSUM 30Jul2020<<

    //                 //RSPLSUM 10Aug2020>>
    //                 CLEAR(BillToCustomerName);
    //                 IF "Sales Invoice Header"."Sell-to Customer No." <> "Sales Invoice Header"."Bill-to Customer No." THEN BEGIN
    //                     RecCustNew.RESET;
    //                     IF RecCustNew.GET("Sales Invoice Header"."Bill-to Customer No.") THEN
    //                         BillToCustomerName := RecCustNew."Full Name";
    //                 END ELSE
    //                     BillToCustomerName := "Sales Invoice Header"."Full Name";
    //                 //RSPLSUM 10Aug2020<<

    //                 //RSPLSUM BEGIN>>
    //                 CLEAR(IRN);
    //                 "Sales Invoice Header".CALCFIELDS("Sales Invoice Header".IRN);
    //                 IF "Sales Invoice Header".IRN <> '' THEN
    //                     IRN := 'IRN: ' + "Sales Invoice Header".IRN;
    //                 //RSPLSUM END>>

    //                 //RSPLSUM-TCS>>
    //                 CLEAR(TCSPercent);
    //                 RecSIL.RESET;
    //                 RecSIL.SETRANGE("Document No.", "No.");
    //                 RecSIL.SETRANGE(Type, RecSIL.Type::Item);
    //                 RecSIL.SETFILTER(Quantity, '<>%1', 0);
    //                 IF RecSIL.FINDFIRST THEN BEGIN
    //                     /*  IF RecSIL."TDS/TCS %" <> 0 THEN
    //                          TCSPercent := FORMAT(RecSIL."TDS/TCS %") + ' %'; */
    //                 END;
    //                 //RSPLSUM-TCS<<

    //                 Ctr := 1;

    //                 "Sales Invoice Header".CALCFIELDS("Sales Invoice Header"."QR code");//ROBOEinv
    //                 //BarcodeForQuarantineLabel;//RSPLSUM // DJ Commented 06/10/2020

    //                 //DJ 15062021
    //                 CLEAR(DecPackingForwarding);
    //                 /*  CLEAR(PostedStrOrdrDetails);
    //                  PostedStrOrdrDetails.SETRANGE(Type, PostedStrOrdrDetails.Type::Sale);
    //                  PostedStrOrdrDetails.SETRANGE("Document Type", PostedStrOrdrDetails."Document Type"::Invoice);
    //                  PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
    //                  PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
    //                  PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'PKGFWD');
    //                  IF PostedStrOrdrDetails.FINDFIRST THEN
    //                      DecPackingForwarding := PostedStrOrdrDetails."Calculation Value"; */
    //                 //DJ 15062021
    //             end;
    //         }
    //     }

    //     requestpage
    //     {

    //         layout
    //         {
    //             area(content)
    //             {
    //                 field(NoOfCopies; NoOfCopies)
    //                 {
    //                     Caption = 'No.Of Copies';
    //                 }
    //                 field(ShowInternalInfo; ShowInternalInfo)
    //                 {
    //                     Caption = 'Show Internal Information';
    //                 }
    //                 field(LogInteraction; LogInteraction)
    //                 {
    //                     Caption = 'Log Interaction';
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
    //         /*
    //         //>>18July2017
    //         //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-----START
    //         IF "Sales Invoice Header"."Print Invoice" = FALSE THEN
    //         BEGIN
    //          IF NOT(CurrReport.PREVIEW) THEN
    //          BEGIN
    //           recSIH.RESET;
    //           recSIH.SETRANGE(recSIH."No.","Sales Invoice Header"."No.");
    //           recSIH.SETRANGE(recSIH."Print Invoice",FALSE);
    //           IF recSIH.FINDFIRST THEN
    //           BEGIN
    //             //RSPLSUM 04Feb2020>>
    //             SalesSetup.GET;
    //             IF SalesSetup."Email Alert on Sales Inv Post" THEN BEGIN
    //               IF (recSIH."Print Invoice" = FALSE) AND (recSIH."Salesperson Email Sent" = FALSE) THEN BEGIN
    //                   SetSalespersonEmailValues(recSIH);
    //                   SalespersonEmailNotification(recSIH);
    //                   recSIH."Salesperson Email Sent" := TRUE;
    //               END;
    //             END;
    //             //RSPLSUM 04Feb2020<<
    //            recSIH."Print Invoice" := TRUE;
    //            recSIH.MODIFY;
    //           END;
    //          END;
    //         END;
    //         //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-------END

    //         //<<18July2017
    //         */

    //     end;

    //     trigger OnPreReport()
    //     begin

    //         CompanyInfo1.GET;
    //         CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
    //         CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
    //         CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");

    //         //InitTextVariable;
    //     end;

    //     var
    //         Loc: Record 14;
    //         Loc1: Record 14;
    //         SalesPerson: Record 13;
    //         RespCentre: Record 5714;
    //         RecState: Record State;
    //         RecState1: Record State;
    //         RecCountry: Record 9;
    //         RecCountry1: Record 9;
    //         CompanyInfo1: Record 79;
    //         RecCust: Record 18;
    //         RecItem: Record 27;
    //         ShipToCode: Record 222;
    //         "Payment Terms": Record 3;
    //         "Shipment Method": Record 10;
    //         "Shipping Agent": Record 291;
    //         ILE: Record 32;
    //         ///"E.C.C. Nos.": Record "13708";
    //        // "E.C.C. Nos.1": Record "13708";
    //         CustPostSetup: Record 92;
    //         CustName: Text[100];
    //         UOM: Code[10];
    //         Qty: Decimal;
    //         QtyPerUOM: Decimal;
    //         ShiptoName: Text[80];
    //         ShipToCST: Code[30];
    //         ShiptoTIN: Code[30];
    //         ShiptoLST: Code[30];
    //         ECCNo: Code[20];
    //         OnesText: array[20] of Text[30];
    //         TensText: array[10] of Text[30];
    //         ExponentText: array[5] of Text[30];
    //         DescriptionLineDuty: array[2] of Text[132];
    //         DescriptionLineeCess: array[2] of Text[132];
    //         DescriptionLineSHeCess: array[2] of Text[132];
    //         DescriptionLineTot: array[2] of Text[132];
    //         DescriptionLineAD: array[2] of Text[132];
    //         PackDescription: Text[30];
    //         SrNo: Integer;
    //         RoundOffAcc: Code[20];
    //         RoundOffAmnt: Decimal;
    //         TransShptNo: Code[5];
    //         TransShptNoLen: Integer;
    //         "<EBT>": Integer;
    //         TaxDescr: Code[10];
    //         ECCDesc: Code[20];
    //         PackDescriptionforILE: Text[30];
    //         TaxJurisd: Record 320;
    //         TaxDescription: Text[50];
    //         Freight: Decimal;
    //         Discount: Decimal;
    //         TotalAmttoCustomer: Decimal;
    //         InvTotal: Decimal;
    //         CurrDatetime: Time;
    //         State: Text[50];
    //         Ctr: Integer;
    //         SalesInvLine: Record 113;
    //         SegManagement: Codeunit 5051;
    //         LogInteraction: Boolean;
    //         NoOfLoops: Integer;
    //         NoOfCopies: Integer;
    //         cust: Record 18;
    //         copytext: Text[30];
    //         SalesInvCountPrinted: Codeunit 315;
    //         ShowInternalInfo: Boolean;
    //         InvNoLen: Integer;
    //         InvNo: Code[10];
    //         DocTotal: Decimal;
    //         i: Integer;
    //         InvType: Text[50];
    //         InvType1: Text[50];
    //         //  FormCodes: Record "13756";
    //         TransHdr: Record 5740;
    //         ItemVarParmResFinal: Record 50015;
    //         FrieghtAmount: Decimal;
    //         AdDutyAmount: Decimal;
    //         //  PostedStrOrdrDetails: Record "13760";
    //         TotalQty: Decimal;
    //         BEDpercent: Decimal;
    //         //recExcisePostnSetup: Record "13711";
    //         BatchNo: Code[10];
    //         recILE: Record 32;
    //         ecesspercent: Decimal;
    //         SHEpercent: Decimal;
    //         TotalBR: Decimal;
    //         TotalValue: Decimal;
    //         recInvComment: Record 5748;
    //         "Batch/UOM": Text[30];
    //         recItemEBT: Record 27;
    //         ItemDescription: Text[100];
    //         recTSH: Record 5744;
    //         // recECC: Record "13708";
    //         recLocation: Record 14;
    //         recShippingAgent: Record 291;
    //         "---------------": Integer;
    //         recDimensionValue: Record 349;
    //         DimensionName: Text[100];
    //         dimensionCode: Code[20];
    //         DimensionName2: Text[100];
    //         Text003: Label 'Original for Buyer';
    //         Text026: Label 'Zero';
    //         Text027: Label 'Hundred';
    //         Text028: Label '&';
    //         Text029: Label '%1 results in a written number that is too long.';
    //         Text032: Label 'One';
    //         Text033: Label 'Two';
    //         Text034: Label 'Three';
    //         Text035: Label 'Four';
    //         Text036: Label 'Five';
    //         Text037: Label 'Six';
    //         Text038: Label 'Seven';
    //         Text039: Label 'Eight';
    //         Text040: Label 'Nine';
    //         Text041: Label 'Ten';
    //         Text042: Label 'Eleven';
    //         Text043: Label 'Twelve';
    //         Text044: Label 'Thirteen';
    //         Text045: Label 'Fourteen';
    //         Text046: Label 'Fifteen';
    //         Text047: Label 'Sixteen';
    //         Text048: Label 'Seventeen';
    //         Text049: Label 'Eighteen';
    //         Text050: Label 'Nineteen';
    //         Text051: Label 'Twenty';
    //         Text052: Label 'Thirty';
    //         Text053: Label 'Forty';
    //         Text054: Label 'Fifty';
    //         Text055: Label 'Sixty';
    //         Text056: Label 'Seventy';
    //         Text057: Label 'Eighty';
    //         Text058: Label 'Ninety';
    //         Text059: Label 'Thousand';
    //         Text060: Label 'Million';
    //         Text061: Label 'Billion';
    //         Text1280000: Label 'Lakh';
    //         Text1280001: Label 'Crore';
    //         vCheck: Report 1401;
    //         OutPutNo: Integer;
    //         vCount: Integer;
    //         "-----06Apr2017": Integer;
    //         LocFrom: Record 14;
    //         LocState: Text[80];
    //         FExcise: Decimal;
    //         FEcess: Decimal;
    //         FShe: Decimal;
    //         FAdd: Decimal;
    //         FInvT: Decimal;
    //         "--14Apr2017": Integer;
    //         recDimSet: Record 480;
    //         "---07July2017--GST": Integer;
    //         vCGST: Decimal;
    //         vCGSTRate: Decimal;
    //         vSGST: Decimal;
    //         vSGSTRate: Decimal;
    //         vIGST: Decimal;
    //         vIGSTRate: Decimal;
    //         DetailGSTEntry: Record "Detailed GST Ledger Entry";
    //         DiscAmt06: Decimal;
    //         ChargeAmt06: Decimal;
    //         //        recStrOrderLineDet: Record "13798";
    //         SL14: Record 113;
    //         "-----15July2017": Integer;
    //         ShipAdd1: Text[60];
    //         ShipAdd2: Text[60];
    //         ShipAdd3: Text[60];
    //         ShipGSTStateCode: Code[10];
    //         ShipGSTNo: Code[15];
    //         ShipPANNo: Code[10];
    //         ShiptoName15: Text[100];
    //         BillState: Text[50];
    //         BillGSTStateCode: Code[10];
    //         BillGSTNo: Code[15];
    //         BillPANNo: Code[10];
    //         BillCountry: Text[50];
    //         ShipCountry: Text[50];
    //         Ship2Name: Text[60];
    //         ShipState: Text[50];
    //         Cus12: Record 18;
    //         AmtinWord15: array[2] of Text[150];
    //         FinalAmt: Decimal;
    //         CGST15: Text[30];
    //         SGST15: Text[30];
    //         IGST15: Text[30];
    //         DGST04: Record "Detailed GST Ledger Entry";
    //         TotalCGST: Decimal;
    //         TotalSGST: Decimal;
    //         TotalIGST: Decimal;
    //         CGSTinWord15: array[2] of Text[150];
    //         SGSTinWord15: array[2] of Text[150];
    //         IGSTinWord15: array[2] of Text[150];
    //         "---18July2017": Integer;
    //         recSIH: Record 112;
    //         FactoryCaption: Text[30];
    //         "-----------09Aug2017": Integer;
    //         Cust09: Record 18;
    //         SuppliersCode: Text[20];
    //         "--------21Aug2017": Integer;
    //         PCharge: Decimal;
    //         NCharge: Decimal;
    //         "-----30Nov2017": Integer;
    //         TotalAmt30: Decimal;
    //         BillPhoneNo: Text;
    //         ShipPhoneNo: Text;
    //         SalesSetup: Record 311;
    //         GloInvoiceDate: Date;
    //         GloInvoiceNo: Code[30];
    //         GloInvoicevalue: Decimal;
    //         GloCustomerName: Text[100];
    //         GloVehicleno: Code[20];
    //         GloTransporteraName: Text[100];
    //         GloDriverName: Text[100];
    //         GloDriverMobileNo: Code[20];
    //         GloReciepientEmail: Text[80];
    //         UserSetupRec: Record 91;
    //         SalespersonPurchaser: Record 13;
    //         GloLocationName: Text[100];
    //         LocRec: Record 14;
    //         ShippingAgent: Record 291;
    //         SaleLines: Record 37;
    //         GloInvoiceQuantity: Decimal;
    //         GloExtDocNo: Text;
    //         RecSalesInvHdr: Record 112;
    //         Text1005: Label 'Email notification has been sent to salesperson.';
    //         EWayBillNo: Code[20];
    //         GloOrderNo: Code[20];
    //         BalQtyToInv: Decimal;
    //         BillToCustomerName: Text[100];
    //         RecCustNew: Record 18;
    //         RecGSTLedgEntry: Record "GST Ledger Entry";
    //         IRN: Text[255];
    //         TCSPercent: Text[10];
    //         RecSIL: Record 113;
    //         TCSAmount: Decimal;
    //         GloEwayBillNo: Code[20];
    //         GloEwayBillValidUpto: Text[30];
    //         DetailedsEWayBill: Record 50044;
    //         UOMCode: Code[20];
    //         QtyPerPack: Decimal;
    //         TotalQtyFO: Decimal;
    //         BasicRate: Decimal;
    //         //SMTPMailSetup: Record "409";
    //         DecPackingForwarding: Decimal;
    //         PKGFWDAmount: Decimal;
    //         FREIGHTAmount: Decimal;
    //         Quantity_LineN: Decimal;
    //         RatePerMTLTKG: Text[50];
    //         //  InvoiceAutomotiveGSTLOGO1: Report 50077;
    //         FileName: Text;
    //         PathName: Text;
    //         Path: Text;
    //         SalesInvHead: Record 112;
    //         ServerTempFile: Text;
    //         FileMgmt: Codeunit 419;

    //     // [Scope('Internal')]
    //     procedure SalespersonEmailNotification(recSIH1: Record 112)
    //     var
    //         // SMTPMail: Codeunit "400";
    //         SAE18: Record 50009;
    //         SA18: Record 50008;
    //         SubjectText: Text;
    //         User18: Record 91;
    //         SenderName: Text;
    //         SenderEmail: Text;
    //         Text18: Text;
    //         Cust18: Record 18;
    //         OtAmt: Decimal;
    //         CrAmt: Decimal;
    //         ODAmt: Decimal;
    //         ReceiveEmail: Text;
    //         DimEnable: Boolean;
    //         Loc21: Record 14;
    //         SL21: Record 37;
    //         AppEmail: Text;
    //         RecSIL: Record 113;
    //         RecSL: Record 37;
    //         recReportInvoice: Report 50076;
    //         FileName: Text;
    //         PathName: Text;
    //         Path: Text;
    //         EmailMsg: Codeunit "Email Message";
    //         EmailObj: Codeunit Email;
    //         RecipientType: Enum "Email Recipient Type";
    //         TempBlob: Codeunit "Temp Blob";
    //         OutStr: OutStream;
    //         Instr: InStream;
    //     begin
    //         //    SMTPMailSetup.GET;//RSPLSUM02Apr21

    //         SubjectText := '';
    //         SenderName := '';
    //         SenderEmail := '';
    //         ReceiveEmail := '';
    //         Text18 := '';
    //         //  CLEAR(SMTPMail);


    //         IF UserSetupRec.GET(USERID) THEN;
    //         //UserSetupRec.TESTFIELD("E-Mail");

    //         //RSPLAM  <<
    //         //IF recSIH1.GET("Sales Invoice Header"."No.") THEN BEGIN
    //         CLEAR(ServerTempFile);
    //         ServerTempFile := FileMgmt.ServerTempFileName('pdf');
    //         IF REPORT.SAVEASPDF(50076, ServerTempFile, "Sales Invoice Header") THEN BEGIN
    //             //RSPLAM  >>

    //             SubjectText := 'Invoice - ' + "Sales Invoice Header"."No." + ' - ' + "Sales Invoice Header"."Sell-to Customer Name";//RSPLSUM 01Feb21

    //             //>>Email Body
    //             //RSPLSUM 01Feb21--SMTPMail.CreateMessage('',UserSetupRec."E-Mail",GloReciepientEmail,'','',TRUE);
    //             //RSPLSUM02Apr21--SMTPMail.CreateMessage('',UserSetupRec."E-Mail",GloReciepientEmail,SubjectText,'',TRUE);//RSPLSUM 01Feb21
    //             //SMTPMail.CreateMessage('',SMTPMailSetup."User ID",GloReciepientEmail,SubjectText,'',TRUE);//RSPLSUM02Apr21
    //             //SMTPMail.CreateMessage('',SMTPMailSetup."User ID",'fahim.ahmad@gpglobal.com',SubjectText,'',TRUE);//RSPLSUM02Apr21
    //             //  SMTPMail.CreateMessage('', SMTPMailSetup."User ID", 'avinash.mali@robo-soft.net', SubjectText, '', TRUE);//RSPLSUM02Apr21
    //             //SMTPMail.AddRecipients(GloReciepientEmail);
    //             Emailmsg.Create(GloReciepientEmail, SubjectText, '', TRUE);


    //             SMTPMail.AppendBody('<Br>');
    //             SMTPMail.AppendBody('<Br>');
    //             SMTPMail.AppendBody('<B> New invoice dispatched from </B>' + GloLocationName);
    //             SMTPMail.AppendBody('<Br>');
    //             SMTPMail.AppendBody('<Br>');

    //             //Table Start
    //             SMTPMail.AppendBody('<Table  Border="1">');
    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Invoice No </th>');
    //             SMTPMail.AppendBody('<td>' + GloInvoiceNo + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>InvoiceDate</th>');
    //             SMTPMail.AppendBody('<td>' + FORMAT(GloInvoiceDate) + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Customer Name</th>');
    //             SMTPMail.AppendBody('<td>' + GloCustomerName + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             RecSIL.RESET;
    //             RecSIL.SETRANGE("Document No.", GloInvoiceNo);
    //             RecSIL.CALCSUMS("Quantity (Base)");
    //             SMTPMail.AppendBody('<th>Quantity</th>');
    //             SMTPMail.AppendBody('<td>' + FORMAT(RecSIL."Quantity (Base)", 0, '<Integer Thousand><Decimals,3>') + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             //RSPLSUM 07Aug2020>>
    //             SMTPMail.AppendBody('<tr>');
    //             CLEAR(BalQtyToInv);
    //             RecSL.RESET;
    //             RecSL.SETRANGE("Document No.", GloOrderNo);
    //             RecSL.SETRANGE("Document Type", RecSL."Document Type"::Order);
    //             RecSL.SETRANGE(Type, RecSL.Type::Item);
    //             IF RecSL.FINDSET THEN
    //                 REPEAT
    //                     BalQtyToInv += RecSL."Quantity (Base)" - RecSL."Qty. Invoiced (Base)";
    //                 UNTIL RecSL.NEXT = 0;

    //             SMTPMail.AppendBody('<th>Balance Quantity to Invoice</th>');
    //             SMTPMail.AppendBody('<td>' + FORMAT(BalQtyToInv, 0, '<Integer Thousand><Decimals,3>') + '</td>');
    //             SMTPMail.AppendBody('</tr>');
    //             //RSPLSUM 07Aug2020<<

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Invoice value </th>');
    //             SMTPMail.AppendBody('<td>' + FORMAT(GloInvoicevalue, 0, '<Integer Thousand><Decimals,3>') + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Purchase Order No </th>');
    //             SMTPMail.AppendBody('<td>' + GloExtDocNo + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Vehicle No </th>');
    //             SMTPMail.AppendBody('<td>' + GloVehicleno + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Transport Name </th>');
    //             SMTPMail.AppendBody('<td>' + GloTransporteraName + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Driver Name</th>');
    //             SMTPMail.AppendBody('<td>' + GloDriverName + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Driver Mobile No</th>');
    //             SMTPMail.AppendBody('<td>' + GloDriverMobileNo + '</td>');
    //             SMTPMail.AppendBody('</tr>');
    //             //RSPLSUM 19Nov2020>>
    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Eway Bill No.</th>');
    //             SMTPMail.AppendBody('<td>' + GloEwayBillNo + '</td>');
    //             SMTPMail.AppendBody('</tr>');

    //             SMTPMail.AppendBody('<tr>');
    //             SMTPMail.AppendBody('<th>Eway Bill Valid Date</th>');
    //             SMTPMail.AppendBody('<td>' + GloEwayBillValidUpto + '</td>');
    //             SMTPMail.AppendBody('</tr>');
    //             //RSPLSUM 19Nov2020<<
    //             SMTPMail.AppendBody('</table>');
    //             //Table End

    //             SMTPMail.AppendBody('<Br>');
    //             SMTPMail.AppendBody('<Br>');
    //             //SMTPMail.AddAttachment(FileName,''); // Fahim 02-08-2022
    //             IF ServerTempFile <> '' THEN
    //                 SMTPMail.AddAttachment(ServerTempFile, '');
    //             EmailObj.Send(Emailmsg, Enum::"Email Scenario"::Default);
    //             // SMTPMail.Send;
    //             MESSAGE(Text1005);
    //         END;
    //         //<<Email Body
    //     end;

    //     local procedure SetSalespersonEmailValues(SalesInvHeaderP: Record 112)
    //     var
    //         CustRec: Record 18;
    //     begin
    //         IF CustRec.GET(SalesInvHeaderP."Bill-to Customer No.") THEN;
    //         IF SalespersonPurchaser.GET(SalesInvHeaderP."Salesperson Code") THEN;
    //         IF LocRec.GET(SalesInvHeaderP."Location Code") THEN;
    //         IF ShippingAgent.GET(SalesInvHeaderP."Shipping Agent Code") THEN;
    //         //  SalesInvHeaderP.CALCFIELDS("Amount to Customer", "EWB No.", "EWB Date");//RSPLSUM 19Nov2020--Ewb No. & Ewb Date Added
    //         GloInvoiceDate := SalesInvHeaderP."Posting Date";
    //         GloInvoiceNo := SalesInvHeaderP."No.";
    //         GloInvoicevalue := 0;//SalesInvHeaderP."Amount to Customer";
    //         GloCustomerName := CustRec.Name;
    //         GloVehicleno := SalesInvHeaderP."Vehicle No.";
    //         GloTransporteraName := ShippingAgent.Name;
    //         //RSPLSUM 02Dec2020--GloDriverName := '';
    //         GloDriverName := SalesInvHeaderP."Driver's Name";//RSPLSUM 02Dec2020
    //         GloDriverMobileNo := SalesInvHeaderP."Driver's Mobile No.";
    //         GloReciepientEmail := SalespersonPurchaser."E-Mail";
    //         GloLocationName := LocRec.Name;
    //         //GloInvoiceQuantity := InvoiceQuantity;
    //         GloExtDocNo := SalesInvHeaderP."External Document No.";
    //         GloOrderNo := SalesInvHeaderP."Order No.";//RSPLSUM 07Aug2020
    //         GloEwayBillNo := SalesInvHeaderP."EWB No.";//RSPLSUM 19Nov2020

    //         //RSPLSUM 02Dec2020>>
    //         DetailedsEWayBill.RESET;
    //         DetailedsEWayBill.SETRANGE("Document No.", SalesInvHeaderP."No.");
    //         DetailedsEWayBill.SETRANGE("EWB No.", SalesInvHeaderP."EWB No.");
    //         IF DetailedsEWayBill.FINDLAST THEN
    //             GloEwayBillValidUpto := DetailedsEWayBill."EWB Valid Upto";
    //         //RSPLSUM 02Dec2020<<

    //         //RSPLSUM 02Dec2020--GloEwayBillDate := SalesInvHeaderP."EWB Date";//RSPLSUM 19Nov2020
    //     end;

    //     //  [Scope('Internal')]
    //     procedure BarcodeForQuarantineLabel()
    //     var
    //         Text5000_Ctx: Label '<xpml><page quantity=''0'' pitch=''74.1 mm''></xpml>SIZE 100 mm, 74.1 mm';
    //         Text5001_Ctx: Label 'DIRECTION 0,0';
    //         Text5002_Ctx: Label 'REFERENCE 0,0';
    //         Text5003_Ctx: Label 'OFFSET 0 mm';
    //         Text5004_Ctx: Label 'SET PEEL OFF';
    //         Text5005_Ctx: Label 'SET CUTTER OFF';
    //         Text5006_Ctx: Label 'SET PARTIAL_CUTTER OFF';
    //         Text5007_Ctx: Label '<xpml></page></xpml><xpml><page quantity=''1'' pitch=''74.1 mm''></xpml>SET TEAR ON';
    //         Text5008_Ctx: Label 'CLS';
    //         Text5009_Ctx: Label 'CODEPAGE 1252';
    //         Text5010_Ctx: Label 'TEXT 806,792,"0",180,17,14,"SCHILLER_"';
    //         Text5011_Ctx: Label 'TEXT 1093,443,"0",180,14,12,"Item"';
    //         Text5012_Ctx: Label 'TEXT 1093,303,"0",180,10,12,"Part No."';
    //         Text5013_Ctx: Label 'TEXT 1093,166,"0",180,12,12,"Sr. No."';
    //         Text5014_Ctx: Label 'TEXT 933,443,"0",180,14,12,":"';
    //         Text5015_Ctx: Label 'TEXT 933,303,"0",180,14,12,":"';
    //         Text5016_Ctx: Label 'TEXT 933,166,"0",180,14,12,":"';
    //         Text5017_Ctx: Label 'TEXT 896,303,"0",180,10,12,"%1"';
    //         Text5018_Ctx: Label 'TEXT 896,166,"0",180,10,12,"%1"';
    //         Text5019_Ctx: Label 'QRCODE 330,300,L,10,A,180,M2,S7,"%1"';
    //         Text5020_Ctx: Label 'TEXT 896,443,"0",180,10,12,"%1"';
    //         Text5021_Ctx: Label 'PRINT 1,1';
    //         Text5022_Ctx: Label '<xpml></page></xpml><xpml><end/></xpml>';
    //         QRCodeInput: Text;
    //     // TempBlob: Record "99008535";
    //     begin
    //         RecGSTLedgEntry.RESET;
    //         RecGSTLedgEntry.SETRANGE("Document No.", "Sales Invoice Header"."No.");
    //         RecGSTLedgEntry.SETRANGE("Document Type", RecGSTLedgEntry."Document Type"::Invoice);
    //         RecGSTLedgEntry.SETRANGE("Transaction Type", RecGSTLedgEntry."Transaction Type"::Sales);
    //         IF RecGSTLedgEntry.FINDFIRST THEN BEGIN
    //             QRCodeInput := STRSUBSTNO('%1', RecGSTLedgEntry."Scan QrCode E-Invoice");
    //             IF QRCodeInput <> '' THEN BEGIN
    //                 // CreateQRCode(QRCodeInput, TempBlob);
    //                 //  RecGSTLedgEntry."QR Code" := TempBlob.Blob;
    //                 RecGSTLedgEntry.MODIFY;
    //             END;
    //         END;
    //     end;

    //     /*  local procedure CreateQRCode(QRCodeInput: Text[500]; var TempBLOB: Record "99008535")
    //      var
    //          QRCodeFileName: Text[1024];
    //      begin
    //          CLEAR(TempBLOB);
    //          QRCodeFileName := GetQRCode(QRCodeInput);
    //          UploadFileBLOBImportandDeleteServerFile(TempBLOB, QRCodeFileName);
    //      end; */

    //     /* local procedure GetQRCode(QRCodeInput: Text[300]) QRCodeFileName: Text[1024]
    //     var
    //         [RunOnClient]
    //         IBarCodeProvider: DotNet IBarcodeProvider;
    //     begin
    //         GetBarCodeProvider(IBarCodeProvider);
    //         QRCodeFileName := IBarCodeProvider.GetBarcode(QRCodeInput);
    //     end; */

    //     //   [Scope('Internal')]
    //     /*  procedure UploadFileBLOBImportandDeleteServerFile(var TempBlob: Record "99008535"; FileName: Text[1024])
    //      var
    //          FileManagement: Codeunit "419";
    //      begin
    //          FileName := FileManagement.UploadFileSilent(FileName);
    //          FileManagement.BLOBImportFromServerFile(TempBlob, FileName);
    //          DeleteServerFile(FileName);
    //      end; */

    //     // [Scope('Internal')]
    //     /*  procedure GetBarCodeProvider(var IBarCodeProvider: DotNet IBarcodeProvider)
    //      var
    //          [RunOnClient]
    //          QRCodeProvider: DotNet QRCodeProvider;
    //      begin
    //          IF ISNULL(IBarCodeProvider) THEN
    //              IBarCodeProvider := QRCodeProvider.QRCodeProvider;
    //      end; */

    //     /* local procedure DeleteServerFile(ServerFileName: Text)
    //     begin
    //         IF ERASE(ServerFileName) THEN;
    //     end; */
}


report 50173 "Purchase Return Invoice GST"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 12Sep2017   RB-N         New Report Development
    // 19Dec2017   RB-N         VendorItem Description
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseReturnInvoiceGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Purchase Return Invoice GST';

    dataset
    {
        dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(SuppliersCode; SuppliersCode)
            {
            }
            column(FactoryCaption; FactoryCaption)
            {
            }
            column(CompName; CompanyInfo1.Name)
            {
            }
            column(CompPicture; CompanyInfo1.Picture)
            {
            }
            column(CompNamePicture; CompanyInfo1."Name Picture")
            {
            }
            column(CompShadedBoxPicture; CompanyInfo1."Shaded Box")
            {
            }
            column(FactoryLocation; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' - ' + LocFrom."Post Code" + ' ' + LocState + ' ' + ' Phone: ' + LocFrom."Phone No." + ' Email: ' + LocFrom."E-Mail")
            {
            }
            column(LocGSTNo; LocFrom."GST Registration No.")
            {
            }
            column(ExternalDocDate; "Document Date")
            {
            }
            column(ExternalDocNo; "Vendor Cr. Memo No.")
            {
            }
            column(Remarks1; '')
            {
            }
            column(Remarks2; '')
            {
            }
            column(CurrencyFactor_Header; "Currency Factor")
            {
            }
            column(SIH_InvoiceTime; '')
            {
            }
            column(RoundOffAmnt; RoundOffAmnt)
            {
            }
            column(FinalAmt; FinalAmt)
            {
            }
            column(AmtinWord15; AmtinWord15[1])
            {
            }
            column(CGSTinWord15; CGSTinWord15[1])
            {
            }
            column(SGSTinWord15; SGSTinWord15[1])
            {
            }
            column(IGSTinWord15; IGSTinWord15[1])
            {
            }
            column(FullName_SIH; "Full Name")
            {
            }
            column(Location; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' ' + LocFrom."Post Code" + ' ' + LocState + ' Phone: ' + LocFrom."Phone No." + ' ' + ' Email: ' + LocFrom."E-Mail")
            {
            }
            column(BillAdd1; "Buy-from Address")
            {
            }
            column(BillAdd2; "Buy-from Address 2")
            {
            }
            column(BillAdd3; "Buy-from City" + ' - ' + "Buy-from Post Code")
            {
            }
            column(SIH_BillAdd3; "Buy-from City" + '-' + "Buy-from Post Code" + ' ' + BillState + ', ' + BillCountry)
            {
            }
            column(PlaceofSupply; "Buy-from City")
            {
            }
            column(ShipGSTStateCode; OrderStateCo)
            {
            }
            column(ShipGSTNo; OrderGST)
            {
            }
            column(Ship2Name; Ship2Name)
            {
            }
            column(ShiptoName15; ShiptoName15)
            {
            }
            column(BillGSTStateCode; BillGSTStateCode)
            {
            }
            column(BillGSTNo; BillGSTNo)
            {
            }
            column(SIH_ShipName; "Pay-to Name")
            {
            }
            column(SIH_ShipAdd1; OrderAdd)
            {
            }
            column(SIH_ShipAdd2; OrderAdd2)
            {
            }
            column(SIH_ShipAdd3; OrderCity + '-' + OrderPostC + ' ' + OrderStatDesc + ', ' + OrderCoutryNM)
            {
            }
            column(LocState; LocState)
            {
            }
            column(InvNo; InvNo)
            {
            }
            column(EccNos; '"E.C.C. Nos."."C.E. Registration No."')
            {
            }
            column(RangDivision; '')//"E.C.C. Nos."."C.E. Range" + '&' + "E.C.C. Nos."."C.E. Division" + '& ' + "E.C.C. Nos."."C.E. Commissionerate")
            {
            }
            column(CEAddr; '"E.C.C. Nos."."C.E Address1')
            {
            }
            column(CEAddr2; '"E.C.C. Nos."."C.E Address2"')
            {
            }
            column(CECityPostCode; '')//"E.C.C. Nos."."C.E City" + '-' + "E.C.C. Nos."."C.E Post Code")
            {
            }
            column(TINNo; '')//Loc1."T.I.N. No." + '  &  ' + Loc1."C.S.T No.")
            {
            }
            column(PostingDate; "Posting Date")
            {
            }
            column(BRTTime; '')
            {
            }
            column(No_Header; "No.")
            {
            }
            column(RoadPermitNo_Header; '')
            {
            }
            column(Name_ShipingAgent; recShippingAgent.Name)
            {
            }
            column(VehicleNo_Header; '')
            {
            }
            column(LRRRDate_Header; '')
            {
            }
            column(LRRRNo_Header; '')
            {
            }
            column(FrieghtType_Header; "Freight Type")
            {
            }
            column(TransportType_Header; '')
            {
            }
            column(PaymentDescription; "Payment Terms".Description)
            {
            }
            column(Loc_Contact; 'Contact ' + ':- ' + Loc."Phone No.")
            {
            }
            column(Loc_Add1; Loc.Address)
            {
            }
            column(Loc_Add2; Loc."Address 2")
            {
            }
            column(Loc_Add3; Loc.City + ' - ' + Loc."Post Code")
            {
            }
            column(State; RecState.Description + ', ' + RecCountry.Name)
            {
            }
            column(GSTCity; Loc.City)
            {
            }
            column(GSTState; RecState.Description)
            {
            }
            column(GSTStateCode; RecState."State Code (GST Reg. No.)")
            {
            }
            column(Loc_CST; 'Loc."C.S.T No."')
            {
            }
            column(BuyerTIN; 'Loc."T.I.N. No."')
            {
            }
            column(Loc_GSTRegNo; Loc."GST Registration No.")
            {
            }
            column(Loc_WEF_CST; Loc."W.e.f. Date(C.S.T No.)")
            {
            }
            column(Buyer_WEF; Loc."W.e.f. Date(T.I.N No.)")
            {
            }
            column(BuyerECC; '"E.C.C. Nos.1"."C.E. Registration No."')
            {
            }
            column(Phone_Consignee; 'Contact ' + ':- ' + Loc."Phone No.")
            {
            }
            column(Loc_AddConsignee; Loc.Address)
            {
            }
            column(Loc_Add2_Consgnii; Loc."Address 2")
            {
            }
            column(City_Consignee; Loc.City + ' - ' + Loc."Post Code")
            {
            }
            column(State_Consginee; RecState.Description + ', ' + RecCountry.Name)
            {
            }
            column(CST_Con; 'Loc."C.S.T No."')
            {
            }
            column(TIN_Con; 'Loc."T.I.N. No."')
            {
            }
            column(WEF_Cons; Loc."W.e.f. Date(T.I.N No.)")
            {
            }
            column(CE_Cons; '"E.C.C. Nos.1"."C.E. Registration No."')
            {
            }
            column(Batch_UOM; "Batch/UOM")
            {
            }
            column(InsuranceNo_CompInfo; CompanyInfo1."Insurance No.")
            {
            }
            column(InsuranceProvoder_CompInfo; CompanyInfo1."Insurance Provider")
            {
            }
            column(Poilicy_ComInfo; FORMAT(CompanyInfo1."Policy Expiration Date"))
            {
            }
            column(BuyerTINDate; Loc."W.e.f. Date(T.I.N No.)")
            {
            }
            column(Comp_GSTRegNo; CompanyInfo1."GST Registration No.")
            {
            }
            column(CompWeb; CompanyInfo1."Home Page")
            {
            }
            column(CompCINNo; CompanyInfo1."Registration No.")
            {
            }
            column(CompPANNo; CompanyInfo1."P.A.N. No.")
            {
            }
            column(QRCode; "Purch. Cr. Memo Hdr."."QR Code")
            {
            }
            column(IRN_PurchCrMemoHeader; IRN)
            {
            }
            dataitem(CopyLoop; Integer)
            {
                DataItemTableView = SORTING(Number);
                dataitem(Integer; Integer)
                {
                    DataItemTableView = SORTING(Number)
                                        WHERE(Number = FILTER(1));
                    column(InvType; InvType)
                    {
                    }
                    column(InvType1; InvType1)
                    {
                    }
                    column(number; Number)
                    {
                    }
                    column(outPutNos; OutPutNo)
                    {
                    }
                    dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Purch. Cr. Memo Hdr.";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Quantity = FILTER(<> 0));
                        column(CrossReferenceNo; '"Cross-Reference No."')
                        {
                        }
                        column(LineDiscountAmt; "Line Discount Amount")
                        {
                        }
                        column(GSTBaseAmt; 0)// "GST Base Amount")
                        {
                        }
                        column(AmountToCustomer; 0)// "Amount To Vendor")
                        {
                        }
                        column(ChargesToCustomer; 0)// "Charges To Vendor")
                        {
                        }
                        column(GSTPer; 0)// "GST %")
                        {
                        }
                        column(GSTGroupCode_Line; "GST Group Code")
                        {
                        }
                        column(HSNSACCode_Line; "HSN/SAC Code")
                        {
                        }
                        column(LineNo; "Line No.")
                        {
                        }
                        column(DimensionName; DimensionName)
                        {
                        }
                        column(LineNo_Line; "Line No.")
                        {
                        }
                        column(DocumentNo_Line; "Document No.")
                        {
                        }
                        column(BatchNo; '')
                        {
                        }
                        column(QtyperUnitofMeasure_Line; "Qty. per Unit of Measure")
                        {
                        }
                        column(QuantityBase_Line; "Quantity (Base)")
                        {
                        }
                        column(Quantity_Line; Quantity)
                        {
                        }
                        column(TotalUnitPrice; Quantity * "Direct Unit Cost")
                        {
                        }
                        column(UnitPrice; "Direct Unit Cost")
                        {
                        }
                        column(BaseUnitOfMeasure; RecItem."Base Unit of Measure")
                        {
                        }
                        column(Unitprice_QtyPerUnit; "Direct Unit Cost" / "Qty. per Unit of Measure")
                        {
                        }
                        column(ExciseProdPostingGroup_Line; '"Excise Prod. Posting Group"')
                        {
                        }
                        column(BEDpercent; BEDpercent)
                        {
                        }
                        column(BEDant_QtyBase; 0)// "BED Amount" / "Quantity (Base)")
                        {
                        }
                        column(BEDAmount_Line; 0)// "BED Amount")
                        {
                        }
                        column(SHE_Cess; 0)// "SHE Cess Amount" + "eCess Amount")
                        {
                        }
                        column(Amount_Line; Amount)
                        {
                        }
                        column(DescriptionLineDuty_1; DescriptionLineDuty[1])
                        {
                        }
                        column(DescriptionLineeCess_1; DescriptionLineeCess[1])
                        {
                        }
                        column(DescriptionLineSHeCess_1; DescriptionLineSHeCess[1])
                        {
                        }
                        column(DescriptionLineAD_1; DescriptionLineAD[1])
                        {
                        }
                        column(DescriptionLineTot_1; DescriptionLineTot[1])
                        {
                        }
                        column(BEDAmount_Line1; 0)//"BED Amount")
                        {
                        }
                        column(eCessAmount_Line; 0)// "eCess Amount")
                        {
                        }
                        column(ADEAmount_Line; 0)// "ADE Amount")
                        {
                        }
                        column(SHECessAmount_Line; 0)// "SHE Cess Amount")
                        {
                        }
                        column(FrieghtAmount; FrieghtAmount)
                        {
                        }
                        column(DocTotal; DocTotal)
                        {
                        }
                        column(Commente; recInvComment.Comment)
                        {
                        }
                        column(EcessPer; FORMAT(ecesspercent) + '%')
                        {
                        }
                        column(SheCessPer; FORMAT(SHEpercent) + '%')
                        {
                        }
                        column(FormCodeDesc; 'FormCodes.Description')
                        {
                        }
                        column(vCount; vCount)
                        {
                        }
                        column(CGST06; vCGST)
                        {
                        }
                        column(CGST06Rate; vCGSTRate)
                        {
                        }
                        column(SGST06; vSGST)
                        {
                        }
                        column(SGST06Rate; vSGSTRate)
                        {
                        }
                        column(IGST06; vIGST)
                        {
                        }
                        column(IGST06Rate; vIGSTRate)
                        {
                        }
                        column(DiscAmt06; DiscAmt06)
                        {
                        }
                        column(ChargeAmt06; ChargeAmt06)
                        {
                        }
                        column(CGST15; CGST15)
                        {
                        }
                        column(SGST15; SGST15)
                        {
                        }
                        column(IGST15; IGST15)
                        {
                        }
                        column(PCharge; PCharge)
                        {
                        }
                        column(NCharge; NCharge)
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin
                            vCount := COUNT;
                            Ctr := Ctr + 1;
                            //>>12Sep2017 DimsetEntry
                            DimensionName := '';
                            /*
                            IF "Inventory Posting Group"='MERCH' THEN
                            BEGIN
                            
                              recDimSet.RESET;
                              recDimSet.SETRANGE("Dimension Set ID","Dimension Set ID");
                              recDimSet.SETRANGE("Dimension Code",'MERCH');
                              IF recDimSet.FINDFIRST THEN
                              BEGIN
                            
                                recDimSet.CALCFIELDS("Dimension Value Name");
                                DimensionName := recDimSet."Dimension Value Name";
                            
                              END;
                              DimensionName2:=DimensionName;
                            END;
                            */
                            //<<12Sep2017 DimsetEntry


                            IF (DimensionName = '') OR (recItemEBT."Item Category Code" = 'CAT02') OR (recItemEBT."Item Category Code" = 'CAT03') OR
                              (recItemEBT."Item Category Code" = 'CAT11') OR (recItemEBT."Item Category Code" = 'CAT12') OR
                                (recItemEBT."Item Category Code" = 'CAT13') THEN
                                IF (DimensionName = '') AND ("Item Category Code" <> 'CAT01')
                                    AND ("Item Category Code" <> 'CAT04') AND ("Item Category Code" <> 'CAT05')
                                    AND ("Item Category Code" <> 'CAT06') AND ("Item Category Code" <> 'CAT07')
                                    AND ("Item Category Code" <> 'CAT08') AND ("Item Category Code" <> 'CAT09')
                                    AND ("Item Category Code" <> 'CAT10') THEN BEGIN
                                    recILE.SETRANGE("Document No.", "Document No.");
                                    recILE.SETRANGE("Document Line No.", "Line No.");
                                    recILE.SETRANGE("Item No.", "No.");
                                    IF recILE.FINDFIRST THEN
                                        BatchNo := recILE."Lot No.";
                                    DimensionName := 'IPOL ' + Description;
                                END ELSE BEGIN
                                    BatchNo := "Unit of Measure Code";

                                    IF DimensionName <> '' THEN
                                        DimensionName := DimensionName2
                                    ELSE
                                        DimensionName := Description;
                                END;

                            IF Type = Type::"G/L Account" THEN
                                DimensionName := Description;

                            //>>12Sep2017 RepsolName

                            IF "Item Category Code" = 'CAT15' THEN
                                DimensionName := Description;


                            //>>19Dec2017 VendorItemNo
                            IF "Purch. Cr. Memo Line"."Vendor Item No." <> '' THEN
                                DimensionName := "Vendor Item No.";
                            //<<19Dec2017

                            //>>12Sep2017 RepsolName
                            //>>12Sep2017

                            /* recExcisePostnSetup.RESET;
                            recExcisePostnSetup.SETRANGE(recExcisePostnSetup."Excise Bus. Posting Group", "Excise Bus. Posting Group");
                            recExcisePostnSetup.SETFILTER(recExcisePostnSetup."From Date", '<=%1', "Posting Date"); //RSPL008
                            IF recExcisePostnSetup.FINDLAST THEN BEGIN
                                ecesspercent := recExcisePostnSetup."eCess %";
                                BEDpercent := recExcisePostnSetup."BED %";
                                SHEpercent := recExcisePostnSetup."SHE Cess %";
                            END; */

                            /*
                            RecState.SETRANGE(RecState.Code,Loc."State Code");
                            RecState.FINDFIRST;
                            State:=RecState.Description;
                            */

                            IF RecItem.GET("No.") THEN;
                            TotalQty += "Quantity (Base)";
                            TotalBR += "Purch. Cr. Memo Line"."Unit Cost";
                            TotalValue := TotalValue + ("Quantity (Base)" * "Unit Cost");


                            TaxDescr := 'Tax Descr';

                            //>>12Sep2017

                            FExcise += 0;// "BED Amount";
                            FEcess += 0;// "eCess Amount";
                            FShe += 0;//"SHE Cess Amount";
                            FInvT += Amount;

                            //<<12Sep2017

                            //>>12Sep2017 GST
                            CLEAR(vCGST);
                            CLEAR(vSGST);
                            CLEAR(vIGST);
                            CLEAR(vCGSTRate);
                            CLEAR(vSGSTRate);
                            CLEAR(vIGSTRate);

                            //>>12Sep2017
                            CLEAR(CGST15);
                            CLEAR(SGST15);
                            CLEAR(IGST15);
                            DetailGSTEntry.RESET;
                            DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                            DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                            DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                            //DetailGSTEntry.SETRANGE(Type, Type);
                            DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
                            IF DetailGSTEntry.FINDSET THEN
                                REPEAT

                                    IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                        vCGST := ABS(DetailGSTEntry."GST Amount");
                                        vCGSTRate := DetailGSTEntry."GST %";

                                        IF DetailGSTEntry."GST Amount" = 0 THEN
                                            CGST15 := ''
                                        ELSE
                                            CGST15 := ' @ ' + FORMAT(vCGSTRate) + ' %';
                                    END;

                                    IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                        vSGST := ABS(DetailGSTEntry."GST Amount");
                                        vSGSTRate := DetailGSTEntry."GST %";

                                        IF DetailGSTEntry."GST Amount" = 0 THEN
                                            SGST15 := ''
                                        ELSE
                                            SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                                    END;


                                    IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                        vSGST := ABS(DetailGSTEntry."GST Amount");
                                        vSGSTRate := DetailGSTEntry."GST %";

                                        IF DetailGSTEntry."GST Amount" = 0 THEN
                                            SGST15 := ''
                                        ELSE
                                            SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                                    END;

                                    IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                        vIGST := ABS(DetailGSTEntry."GST Amount");
                                        vIGSTRate := DetailGSTEntry."GST %";
                                        //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");

                                        IF DetailGSTEntry."GST Amount" = 0 THEN
                                            IGST15 := ''
                                        ELSE
                                            IGST15 := ' @ ' + FORMAT(vIGSTRate) + ' %';

                                    END;

                                UNTIL DetailGSTEntry.NEXT = 0;
                            //<<12Sep2017 GST


                            /*
                            IF "Total GST Amount" = 0 THEN
                            BEGIN
                              CGST15 := '';
                              SGST15 := '';
                              IGST15 := '';
                            END;
                            
                            IF "Total GST Amount" <> 0 THEN
                            BEGIN
                              CGST15 := ' @ '+FORMAT(vCGSTRate)+' %';
                              SGST15 := ' @ '+FORMAT(vSGSTRate)+' %';
                              IGST15 := ' @ '+FORMAT(vIGSTRate)+' %';
                            END;
                            */
                            //<<15July2017

                            //>>06July2017 StructureDetails
                            CLEAR(DiscAmt06);
                            CLEAR(ChargeAmt06);

                            /*  recStrOrderLineDet.RESET;
                             recStrOrderLineDet.SETRANGE("Invoice No.", "Document No.");
                             recStrOrderLineDet.SETRANGE("Item No.", "No.");
                             recStrOrderLineDet.SETRANGE("Line No.", "Line No.");
                             IF recStrOrderLineDet.FINDSET THEN
                                 REPEAT

                                     IF recStrOrderLineDet."Tax/Charge Group" = 'DISCOUNT' THEN BEGIN

                                         DiscAmt06 += recStrOrderLineDet.Amount;

                                     END;

                                     //>>AVPDiscount 03Aug2017
                                     IF recStrOrderLineDet."Tax/Charge Group" = 'AVP-SP-SAN' THEN BEGIN

                                         DiscAmt06 += recStrOrderLineDet.Amount;

                                     END; */

                            //>>TRANSSUBS 03Aug2017
                            /* IF recStrOrderLineDet."Tax/Charge Group" = 'TRANSSUBS' THEN BEGIN

                                DiscAmt06 += recStrOrderLineDet.Amount;

                            END; */

                            //>>CashDiscount 03Aug2017
                            /* IF recStrOrderLineDet."Tax/Charge Group" = 'CASHDISC' THEN BEGIN

                                DiscAmt06 += recStrOrderLineDet.Amount;

                            END; */

                            /*  IF recStrOrderLineDet."Tax/Charge Group" = 'FREIGHT' THEN BEGIN

                                 recStrOrderLineDet.CALCFIELDS("Structure Description");//21Aug2017
                                 IF recStrOrderLineDet."Structure Description" = 'Freight (Taxable)' THEN//21Aug2017
                                     ChargeAmt06 += recStrOrderLineDet.Amount;

                             END;

                         UNTIL recStrOrderLineDet.NEXT = 0; */
                            //<<06July2017 StrucureDetails

                            //>>21Aug2017 Charges
                            CLEAR(PCharge);
                            CLEAR(NCharge);

                            /*  IF "Charges To Vendor" > 0 THEN
                                 PCharge := "Charges To Vendor"
                             ELSE
                                 NCharge := "Charges To Vendor"; */

                            //<<21Aug2017

                            DocTotal += Amount + 0;// "BED Amount" + "eCess Amount"
                                                   // + "SHE Cess Amount" + FrieghtAmount + "ADE Amount" + "Total GST Amount";

                            /*
                            DocTotal:=ROUND((
                            "Transfer Shipment Line".Amount + "Transfer Shipment Line"."BED Amount"+"Transfer Shipment Line"."eCess Amount"
                            +"Transfer Shipment Line"."SHE Cess Amount"+FrieghtAmount+"Transfer Shipment Line"."ADE Amount"+RoundOffAmnt),1.0,'=');
                            */

                            /*  FormCodes.SETRANGE(FormCodes.Code, "Form Code");
                             IF FormCodes.FINDFIRST THEN; */

                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineDuty, ROUND(FExcise, 0.01), ''); //06Apr2017
                            //vCheck.FormatNoText(DescriptionLineDuty,ROUND("BED Amount",0.01),'');
                            DescriptionLineDuty[1] := DELCHR(DescriptionLineDuty[1], '=', '*');

                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineeCess, ROUND(FEcess, 0.01), ''); //06Apr2017
                            //vCheck.FormatNoText(DescriptionLineeCess,ROUND("eCess Amount",0.01),'');
                            DescriptionLineeCess[1] := DELCHR(DescriptionLineeCess[1], '=', '*');

                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineSHeCess, ROUND(FShe, 0.01), ''); //06Apr2017
                            //vCheck.FormatNoText(DescriptionLineSHeCess,ROUND("SHE Cess Amount",0.01),'');
                            DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1], '=', '*');

                            /*
                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineSHeCess,ROUND("SHE Cess Amount",0.01),'');
                            DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1],'=','*');
                            */

                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineTot, ROUND((DocTotal), 1.0), '');
                            DescriptionLineTot[1] := DELCHR(DescriptionLineTot[1], '=', '*');


                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineAD, ROUND(/*"ADE Amount"*/0, 0.01), '');
                            DescriptionLineAD[1] := DELCHR(DescriptionLineAD[1], '=', '*');

                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin

                        DocTotal := 0; //06Apr2017
                        FExcise := 0; //06Apr2017
                        FEcess := 0; //06Apr2017
                        FShe := 0; //06Apr2017
                        FInvT := 0; //06Apr2017


                        i += 1;

                        IF i = 1 THEN
                            InvType := 'ORIGINAL FOR RECIPIENT'
                        ELSE
                            IF i = 2 THEN
                                InvType := 'DUPLICATE FOR TRANSPORTER'
                            ELSE
                                IF i = 3 THEN
                                    InvType := 'TRIPLICATE FOR SUPPLIER'
                                ELSE
                                    IF i = 4 THEN BEGIN
                                        InvType := 'QUADRAPLICATE FOR H.O';
                                    END ELSE
                                        IF i = 5 THEN BEGIN
                                            InvType := 'EXTRA COPY';
                                            InvType1 := '';
                                        END ELSE
                                            InvType := '';
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    OutPutNo += 1;
                    IF Number > 1 THEN
                        copytext := Text003;
                    //CurrReport.PAGENO := 1;
                end;

                trigger OnPreDataItem()
                begin

                    DocTotal := 0; //12Sep2017

                    NoOfCopies := 3;
                    NoOfLoops := ABS(NoOfCopies) + 1;
                    copytext := '';
                    SETRANGE(Number, 1, NoOfLoops);
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //>>12Sep2017
                IF "Posting Date" < 20170107D THEN //010717
                    ERROR('This Report is Applicable from JULY2017 Invoices');

                //>>12Sep2017

                //>>12Sep2017
                SuppliersCode := '';
                IF Cust09.GET("Sell-to Customer No.") THEN
                    SuppliersCode := Cust09."Our Account No.";

                //<<12Sep2017


                //>>12Sep2017 FactoryCaption
                CLEAR(FactoryCaption);
                IF ("Location Code" = 'PLANT01') OR ("Location Code" = 'PLANT02') THEN
                    FactoryCaption := 'Factory'
                ELSE
                    FactoryCaption := 'Matrl. Sold from';
                //<<12Sep2017 FactoryCaption

                //>>12Sep2017
                IF LocFrom.GET("Location Code") THEN BEGIN
                    IF RecState.GET(LocFrom."State Code") THEN
                        LocState := RecState.Description;
                END;
                //<<12Sep2017

                //>>12Sep2017 PaymentTerm
                "Payment Terms".RESET;
                IF "Payment Terms".GET("Payment Terms Code") THEN;

                //<<12Sep2017 PaymentTerm



                CurrDatetime := TIME;
                TotalQty := 0;

                CLEAR(FrieghtAmount);
                CLEAR(AdDutyAmount);
                /* PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'FREIGHT');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                    FrieghtAmount := PostedStrOrdrDetails."Calculation Value"; */

                // EBT Start 10.08.10 start

                /* PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails.Type, PostedStrOrdrDetails.Type::Transfer);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'ADDL.DUTY');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                    AdDutyAmount := PostedStrOrdrDetails."Calculation Value"; */

                // EBT Start 10.08.10 end


                /*
                recShippingAgent.RESET;
                recShippingAgent.SETRANGE(recShippingAgent.Code,"Shipping Agent Code");
                IF recShippingAgent.FINDFIRST THEN;
                */

                //>>12Sep2017 Roundoff Amount
                CLEAR(RoundOffAmnt);
                SL14.RESET;
                SL14.SETRANGE("Document No.", "No.");
                SL14.SETRANGE(Type, SL14.Type::"G/L Account");
                SL14.SETRANGE("No.", '74012210');
                IF SL14.FINDFIRST THEN BEGIN
                    RoundOffAmnt := SL14.Amount;
                END;

                CLEAR(FinalAmt);
                SL14.RESET;
                SL14.SETRANGE("Document No.", "No.");
                //SL14.SETRANGE(Type,SL14.Type::Item);
                IF SL14.FINDSET THEN
                    REPEAT
                        FinalAmt += 0;// SL14."Amount To Vendor";
                    //MESSAGE(' %1 Final Amt \\  %2 LineNo.',FinalAmt,SL14."Line No.");
                    UNTIL SL14.NEXT = 0;

                //<<12Sep2017 Roundoff Amount

                //>>AmountinWords
                vCheck.InitTextVariable;
                vCheck.FormatNoText(AmtinWord15, ROUND((FinalAmt + RoundOffAmnt), 0.01), '');
                AmtinWord15[1] := DELCHR(AmtinWord15[1], '=', '*');

                //<<AmountinWords

                //>>12Sep2017TotalGST Amount
                CLEAR(TotalCGST);
                CLEAR(TotalSGST);
                CLEAR(TotalIGST);
                DGST04.RESET;
                DGST04.SETRANGE("Document No.", "No.");
                IF DGST04.FINDSET THEN
                    REPEAT
                        IF DGST04."GST Component Code" = 'CGST' THEN BEGIN
                            TotalCGST += ABS(DGST04."GST Amount");
                        END;

                        IF DGST04."GST Component Code" = 'SGST' THEN BEGIN
                            TotalSGST += ABS(DGST04."GST Amount");
                        END;

                        IF DGST04."GST Component Code" = 'UTGST' THEN BEGIN
                            TotalSGST += ABS(DGST04."GST Amount");
                        END;


                        IF DGST04."GST Component Code" = 'IGST' THEN BEGIN
                            TotalIGST += ABS(DGST04."GST Amount");
                        END;

                    UNTIL DGST04.NEXT = 0;

                //>>CGST AmountinWords
                vCheck.InitTextVariable;
                //vCheck.FormatNoText(CGSTinWord15,ROUND((TotalCGST),1.0),'');
                vCheck.FormatNoText(CGSTinWord15, ROUND((TotalCGST), 0.01), '');
                CGSTinWord15[1] := DELCHR(CGSTinWord15[1], '=', '*');

                //>>SGST AmountinWords
                vCheck.InitTextVariable;
                vCheck.FormatNoText(SGSTinWord15, ROUND((TotalSGST), 0.01), '');
                SGSTinWord15[1] := DELCHR(SGSTinWord15[1], '=', '*');

                //>>IGST AmountinWords
                vCheck.InitTextVariable;
                vCheck.FormatNoText(IGSTinWord15, ROUND((TotalIGST), 0.01), '');
                IGSTinWord15[1] := DELCHR(IGSTinWord15[1], '=', '*');

                //<<12Sep2017TotalGST Amount


                //>>12Sep2017
                //Bill State Description
                RecState.RESET;
                IF RecState.GET(State) THEN //1
                BEGIN
                    BillState := RecState.Description;
                    BillGSTStateCode := RecState."State Code (GST Reg. No.)";
                END;


                //Bill GSTNo
                Cus12.RESET;
                IF Cus12.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.") THEN
                    BillGSTNo := Cus12."GST Registration No.";

                //Bill Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Purch. Cr. Memo Hdr."."Buy-from Country/Region Code") THEN
                    BillCountry := RecCountry.Name;



                //Ship Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Purch. Cr. Memo Hdr."."Pay-to Country/Region Code") THEN
                    ShipCountry := RecCountry.Name;

                //RSPL-PA++
                //Ship2Name
                IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
                    IF OrderAddress.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", "Purch. Cr. Memo Hdr."."Order Address Code") THEN BEGIN
                        Ship2Name := OrderAddress.Name; //"Full Name"
                        OrderAdd := OrderAddress.Address;
                        OrderAdd2 := OrderAddress."Address 2";
                        OrderCity := OrderAddress.City;
                        OrderPostC := OrderAddress."Post Code";
                        OrderGST := OrderAddress."GST Registration No.";
                        IF OrderState.GET(OrderAddress.State) THEN
                            OrderStatDesc := OrderState.Description;
                        OrderStateCo := OrderState."State Code (GST Reg. No.)";
                        IF OrderCoutry.GET(OrderAddress."Country/Region Code") THEN
                            OrderCoutryNM := OrderCoutry.Name;

                    END;
                END ELSE BEGIN
                    Ship2Name := "Purch. Cr. Memo Hdr."."Pay-to Name";
                    OrderAdd := "Purch. Cr. Memo Hdr."."Pay-to Address";
                    OrderAdd2 := "Purch. Cr. Memo Hdr."."Pay-to Address 2";
                    OrderCity := "Purch. Cr. Memo Hdr."."Pay-to City";
                    OrderPostC := "Purch. Cr. Memo Hdr."."Pay-to Post Code";
                    IF PayVend.GET("Purch. Cr. Memo Hdr."."Pay-to Vendor No.") THEN
                        OrderGST := PayVend."GST Registration No.";
                    IF OrderState.GET(PayVend."State Code") THEN
                        OrderStatDesc := OrderState.Description;
                    OrderStateCo := OrderState."State Code (GST Reg. No.)";
                    IF OrderCoutry.GET("Purch. Cr. Memo Hdr."."Pay-to Country/Region Code") THEN
                        OrderCoutryNM := OrderCoutry.Name;

                END;
                //Shipping

                IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code") THEN BEGIN

                    RecState.RESET;
                    IF RecState.GET(ShipToCode.State) THEN BEGIN
                        ShipState := RecState.Description;
                        ShipGSTStateCode := RecState."State Code (GST Reg. No.)";
                        ShipGSTNo := ShipToCode."GST Registration No.";
                    END;

                END ELSE BEGIN

                    ShipState := BillState;
                    ShipGSTStateCode := BillGSTStateCode;
                    ShipGSTNo := BillGSTNo;

                END;


                //>>15Jul2017



                Ctr := 1;

                //RSPLSUM 01Feb21 BEGIN>>
                CLEAR(IRN);
                "Purch. Cr. Memo Hdr.".CALCFIELDS("Purch. Cr. Memo Hdr."."E-Inv Irn");
                IF "Purch. Cr. Memo Hdr."."E-Inv Irn" <> '' THEN
                    IRN := 'IRN: ' + "Purch. Cr. Memo Hdr."."E-Inv Irn";
                //RSPLSUM 01Feb21 END>>

                "Purch. Cr. Memo Hdr.".CALCFIELDS("Purch. Cr. Memo Hdr."."QR Code");//ROBOEinv //RSPLSUM 01Feb21

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(NoOfCopies; NoOfCopies)
                {
                    ApplicationArea = all;
                    Caption = 'No.Of Copies';
                }
                field(ShowInternalInfo; ShowInternalInfo)
                {
                    ApplicationArea = all;
                    Caption = 'Show Internal Information';
                }
                field(LogInteraction; LogInteraction)
                {
                    ApplicationArea = all;
                    Caption = 'Log Interaction';
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

    trigger OnPreReport()
    begin

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");

        //InitTextVariable;
    end;

    var
        Loc: Record 14;
        Loc1: Record 14;
        SalesPerson: Record 13;
        RespCentre: Record 5714;
        RecState: Record state;
        RecState1: Record state;
        RecCountry: Record 9;
        RecCountry1: Record 9;
        CompanyInfo1: Record 79;
        RecCust: Record 18;
        RecItem: Record 27;
        ShipToCode: Record 222;
        "Payment Terms": Record 3;
        "Shipment Method": Record 10;
        "Shipping Agent": Record 291;
        ILE: Record 32;
        //"E.C.C. Nos.": Record 13708;
        //"E.C.C. Nos.1": Record 13708;
        CustPostSetup: Record 92;
        CustName: Text[100];
        UOM: Code[10];
        Qty: Decimal;
        QtyPerUOM: Decimal;
        ShiptoName: Text[80];
        ShipToCST: Code[30];
        ShiptoTIN: Code[30];
        ShiptoLST: Code[30];
        ECCNo: Code[20];
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        DescriptionLineDuty: array[2] of Text[132];
        DescriptionLineeCess: array[2] of Text[132];
        DescriptionLineSHeCess: array[2] of Text[132];
        DescriptionLineTot: array[2] of Text[132];
        DescriptionLineAD: array[2] of Text[132];
        PackDescription: Text[30];
        SrNo: Integer;
        RoundOffAcc: Code[20];
        RoundOffAmnt: Decimal;
        TransShptNo: Code[5];
        TransShptNoLen: Integer;
        "<EBT>": Integer;
        TaxDescr: Code[10];
        ECCDesc: Code[20];
        PackDescriptionforILE: Text[30];
        TaxJurisd: Record 320;
        TaxDescription: Text[50];
        Freight: Decimal;
        Discount: Decimal;
        TotalAmttoCustomer: Decimal;
        InvTotal: Decimal;
        CurrDatetime: Time;
        State: Text[50];
        Ctr: Integer;
        SalesInvLine: Record 113;
        // SegManagement: Codeunit SegManagement;
        LogInteraction: Boolean;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        cust: Record 18;
        copytext: Text[30];
        SalesInvCountPrinted: Codeunit "Sales Inv.-Printed";
        ShowInternalInfo: Boolean;
        InvNoLen: Integer;
        InvNo: Code[10];
        DocTotal: Decimal;
        i: Integer;
        InvType: Text[50];
        InvType1: Text[50];
        //FormCodes: Record 13756;
        TransHdr: Record 5740;
        ItemVarParmResFinal: Record 50015;
        FrieghtAmount: Decimal;
        AdDutyAmount: Decimal;
        //PostedStrOrdrDetails: Record 13760;
        TotalQty: Decimal;
        BEDpercent: Decimal;
        //recExcisePostnSetup: Record 13711;
        BatchNo: Code[10];
        recILE: Record 32;
        ecesspercent: Decimal;
        SHEpercent: Decimal;
        TotalBR: Decimal;
        TotalValue: Decimal;
        recInvComment: Record 5748;
        "Batch/UOM": Text[30];
        recItemEBT: Record 27;
        ItemDescription: Text[100];
        recTSH: Record 5744;
        // recECC: Record 13708;
        recLocation: Record 14;
        recShippingAgent: Record 291;
        "---------------": Integer;
        recDimensionValue: Record 349;
        DimensionName: Text[100];
        dimensionCode: Code[20];
        DimensionName2: Text[100];
        Text003: Label 'Original for Buyer';
        Text026: Label 'Zero';
        Text027: Label 'Hundred';
        Text028: Label '&';
        Text029: Label '%1 results in a written number that is too long.';
        Text032: Label 'One';
        Text033: Label 'Two';
        Text034: Label 'Three';
        Text035: Label 'Four';
        Text036: Label 'Five';
        Text037: Label 'Six';
        Text038: Label 'Seven';
        Text039: Label 'Eight';
        Text040: Label 'Nine';
        Text041: Label 'Ten';
        Text042: Label 'Eleven';
        Text043: Label 'Twelve';
        Text044: Label 'Thirteen';
        Text045: Label 'Fourteen';
        Text046: Label 'Fifteen';
        Text047: Label 'Sixteen';
        Text048: Label 'Seventeen';
        Text049: Label 'Eighteen';
        Text050: Label 'Nineteen';
        Text051: Label 'Twenty';
        Text052: Label 'Thirty';
        Text053: Label 'Forty';
        Text054: Label 'Fifty';
        Text055: Label 'Sixty';
        Text056: Label 'Seventy';
        Text057: Label 'Eighty';
        Text058: Label 'Ninety';
        Text059: Label 'Thousand';
        Text060: Label 'Million';
        Text061: Label 'Billion';
        Text1280000: Label 'Lakh';
        Text1280001: Label 'Crore';
        vCheck: Report "Check Report";
        OutPutNo: Integer;
        vCount: Integer;
        "-----06Apr2017": Integer;
        LocFrom: Record 14;
        LocState: Text[80];
        FExcise: Decimal;
        FEcess: Decimal;
        FShe: Decimal;
        FAdd: Decimal;
        FInvT: Decimal;
        "--14Apr2017": Integer;
        recDimSet: Record 480;
        "---07July2017--GST": Integer;
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        DiscAmt06: Decimal;
        ChargeAmt06: Decimal;
        //recStrOrderLineDet: Record 13798;
        SL14: Record 125;
        "-----15July2017": Integer;
        ShipAdd1: Text[60];
        ShipAdd2: Text[60];
        ShipAdd3: Text[60];
        ShipGSTStateCode: Code[10];
        ShipGSTNo: Code[15];
        ShiptoName15: Text[100];
        BillState: Text[50];
        BillGSTStateCode: Code[10];
        BillGSTNo: Code[15];
        BillCountry: Text[50];
        ShipCountry: Text[50];
        Ship2Name: Text[60];
        ShipState: Text[50];
        Cus12: Record 23;
        AmtinWord15: array[2] of Text[150];
        FinalAmt: Decimal;
        CGST15: Text[30];
        SGST15: Text[30];
        IGST15: Text[30];
        DGST04: Record "Detailed GST Ledger Entry";
        TotalCGST: Decimal;
        TotalSGST: Decimal;
        TotalIGST: Decimal;
        CGSTinWord15: array[2] of Text[150];
        SGSTinWord15: array[2] of Text[150];
        IGSTinWord15: array[2] of Text[150];
        "---18July2017": Integer;
        recSIH: Record 112;
        FactoryCaption: Text[30];
        "-----------09Aug2017": Integer;
        Cust09: Record 23;
        SuppliersCode: Text[20];
        "--------21Aug2017": Integer;
        PCharge: Decimal;
        NCharge: Decimal;
        OrderAddress: Record 224;
        OrderAdd: Text;
        OrderAdd2: Text;
        OrderCity: Text;
        OrderPostC: Code[20];
        OrderStatDesc: Text;
        OrderState: Record State;
        PayVend: Record 23;
        OrderCoutryNM: Code[20];
        OrderCoutry: Record 9;
        OrderGST: Code[20];
        OrderStateCo: Code[10];
        IRN: Text[255];
}


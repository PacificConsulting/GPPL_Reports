report 50083 "Transfer Invoice -  GST--LOGO"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 07Nov2017   RB-N         Dispalying RespolLogo Or IpolLogo Depending upon ItemCategory
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TransferInvoiceGSTLOGO.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Transfer Invoice -  GST--LOGO';
    Permissions = TableData 5744 = rim;
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(Logo21; Logo21)
            {
            }
            column(Cat21; Cat21)
            {
            }
            column(RepsolLogo; CompanyInfo1."Repsol Logo")
            {
            }
            column(IpolLogo; CompanyInfo1."Name Picture")
            {
            }
            column(PlaceofSupply1; BillState)
            {
            }
            column(TranComment; recInvComment.Comment)
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
            column(Name; CompanyInfo1."Name Picture")
            {
            }
            column(CompShadedBox; CompanyInfo1."Shaded Box")
            {
            }
            column(FactoryLocation; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' - ' + LocFrom."Post Code" + ' ' + LocState + ' ' + ' Phone: ' + LocFrom."Phone No." + ' Email: ' + LocFrom."E-Mail")
            {
            }
            column(LocGSTNo; LocFrom."GST Registration No.")
            {
            }
            column(ExternalDocDate; '')
            {
            }
            column(ExternalDocNo; "External Document No.")
            {
            }
            column(Remarks1; '')
            {
            }
            column(Remarks2; '')
            {
            }
            column(PrintBRT; "Transfer Shipment Header"."BRT Print Time")
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
            column(FullName_SIH; '')
            {
            }
            column(Location; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' ' + LocFrom."Post Code" + ' ' + LocState + ' Phone: ' + LocFrom."Phone No." + ' ' + ' Email: ' + LocFrom."E-Mail")
            {
            }
            column(BillAdd1; "Transfer-to Address")
            {
            }
            column(BillAdd2; "Transfer-to Address 2")
            {
            }
            column(BillAdd3; "Transfer-to City" + '-' + "Transfer-to Post Code" + ' ' + BillState + ', ' + BillCountry)
            {
            }
            column(SIH_BillAdd3; '')
            {
            }
            column(PlaceofSupply; "Transfer-to City")
            {
            }
            column(ShipGSTStateCode; ShipGSTStateCode)
            {
            }
            column(ShipGSTNo; ShipGSTNo)
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
            column(SIH_ShipName; '')
            {
            }
            column(SIH_ShipAdd1; '')
            {
            }
            column(SIH_ShipAdd2; '')
            {
            }
            column(SIH_ShipAdd3; '')
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
            column(RangDivision; '')// "E.C.C. Nos."."C.E. Range" + '&' + "E.C.C. Nos."."C.E. Division" + '& ' + "E.C.C. Nos."."C.E. Commissionerate")
            {
            }
            column(CEAddr; '"E.C.C. Nos."."C.E Address1"')
            {
            }
            column(CEAddr2; '"E.C.C. Nos."."C.E Address2"')
            {
            }
            column(CECityPostCode; '')//"E.C.C. Nos."."C.E City" + '-' + "E.C.C. Nos."."C.E Post Code")
            {
            }
            column(TINNo; '')// Loc1."T.I.N. No." + '  &  ' + Loc1."C.S.T No.")
            {
            }
            column(PostingDate; "Posting Date")
            {
            }
            column(No_Header; "No.")
            {
            }
            column(RoadPermitNo_Header; EWayBillNo)
            {
            }
            column(Name_ShipingAgent; recShippingAgent.Name)
            {
            }
            column(VehicleNo_Header; "Vehicle No.")
            {
            }
            column(LRRRDate_Header; "LR/RR Date")
            {
            }
            column(LRRRNo_Header; "LR/RR No.")
            {
            }
            column(FrieghtType_Header; "Frieght Type")
            {
            }
            column(TransportType_Header; "Transport Type")
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
            column(QRCode; RecGSTLedgEntry."QR Code")
            {
            }
            column(IRN_SalesInvoiceHeader; IRN)
            {
            }
            dataitem(CopyLoop; 2000000026)
            {
                DataItemTableView = SORTING(Number);
                dataitem(DataItem1000000077; 2000000026)
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
                    dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Transfer Shipment Header";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Quantity = FILTER(<> 0));
                        column(ChargesLine; 0)//"Charges to Transfer")
                        {
                        }
                        column(TotalGSTAmt; 0)//"Total GST Amount")
                        {
                        }
                        column(BatchNo; BatchNo)
                        {
                        }
                        column(MRPPrice; 0)// "MRP Price" * "Qty. per Unit of Measure")
                        {
                        }
                        column(GSTBaseAmt; 0)// "GST Base Amount")
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
                        column(QtyperUnitofMeasure_Line; "Qty. per Unit of Measure")
                        {
                        }
                        column(QuantityBase_Line; "Quantity (Base)")
                        {
                        }
                        column(Quantity_Line; Quantity)
                        {
                        }
                        column(TotalUnitPrice; Quantity * "Unit Price")
                        {
                        }
                        column(UnitPrice_TSL; "Unit Price")
                        {
                        }
                        column(BaseUnitOfMeasure; RecItem."Base Unit of Measure")
                        {
                        }
                        column(Unitprice_QtyPerUnit; "Unit Price" / "Qty. per Unit of Measure")
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
                        column(BEDAmount_Line1; 0)// "BED Amount")
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
                        column(MRPofItem; MRPofItem)
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin
                            vCount := COUNT;
                            Ctr := Ctr + 1;




                            //RSPL
                            //DimensionName:='';
                            //Dim table not avialble
                            /*
                            IF "Inventory Posting Group"='MERCH' THEN BEGIN
                             recPosteddocDiemension.RESET;
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Table ID",5745);
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Document No.","Transfer Shipment Line"."Document No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Line No.","Transfer Shipment Line"."Line No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Dimension Code",'MERCH');
                              IF recPosteddocDiemension.FINDFIRST THEN;
                                 dimensionCode:=recPosteddocDiemension."Dimension Value Code";
                               //IF recDimensionValue.GET(dimensionCode) THEN
                               //  DimensionName:=recDimensionValue.Name;
                               recDimensionValue.RESET;
                               recDimensionValue.SETRANGE(recDimensionValue.Code,dimensionCode);
                               IF recDimensionValue.FINDFIRST THEN
                                  DimensionName:=recDimensionValue.Name;
                                  DimensionName2:=DimensionName;
                            END;
                            */

                            //>>14Apr2017 DimsetEntry

                            DimensionName := '';
                            IF "Inventory Posting Group" = 'MERCH' THEN BEGIN

                                recDimSet.RESET;
                                recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                recDimSet.SETRANGE("Dimension Code", 'MERCH');
                                IF recDimSet.FINDFIRST THEN BEGIN

                                    recDimSet.CALCFIELDS("Dimension Value Name");
                                    DimensionName := recDimSet."Dimension Value Name";

                                END;
                                DimensionName2 := DimensionName;
                            END;
                            //<<14Apr2017 DimsetEntry


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
                                    recILE.SETRANGE("Item No.", "Item No.");
                                    IF recILE.FINDFIRST THEN
                                        BatchNo := recILE."Lot No.";
                                    DimensionName := 'IPOL ' + Description;
                                    //ItemDescription := 'IPOL ' +"Transfer Shipment Line".Description;
                                END ELSE BEGIN
                                    BatchNo := "Unit of Measure Code";

                                    //DimensionName:="Transfer Shipment Line".Description;
                                    IF DimensionName <> '' THEN
                                        DimensionName := DimensionName2
                                    ELSE
                                        DimensionName := Description;
                                    //  ItemDescription := "Transfer Shipment Line".Description;
                                END;

                            IF "Transfer Shipment Line"."Item Category Code" = 'CAT15' THEN
                                DimensionName := Description;

                            IF "Item Category Code" = 'CAT14' THEN
                                DimensionName := Description;

                            IF "Item Category Code" = 'CAT18' THEN
                                DimensionName := Description;
                            //>>14Apr2017

                            /*
                            recExcisePostnSetup.RESET;
                            recExcisePostnSetup.SETRANGE(recExcisePostnSetup."Excise Bus. Posting Group","Excise Bus. Posting Group");
                            recExcisePostnSetup.SETFILTER(recExcisePostnSetup."From Date",'<=%1',"Posting Date"); //RSPL008
                            IF recExcisePostnSetup.FINDLAST THEN BEGIN
                              ecesspercent :=  recExcisePostnSetup."eCess %";
                              BEDpercent :=   recExcisePostnSetup."BED %";
                              SHEpercent := recExcisePostnSetup."SHE Cess %";
                            END;
                            */
                            /*
                            RecState.SETRANGE(RecState.Code,Loc."State Code");
                            RecState.FINDFIRST;
                            State:=RecState.Description;
                            */

                            IF RecItem.GET("Item No.") THEN;
                            TotalQty += "Quantity (Base)";
                            TotalBR += "Unit Price";
                            TotalValue := TotalValue + 0;//("Quantity (Base)" * "Unit Price");


                            TaxDescr := 'Tax Descr';

                            //>>06APr2017

                            FExcise += 0;//"BED Amount";
                            FEcess += 0;//"eCess Amount";
                            FShe += 0;//"SHE Cess Amount";
                            FInvT += Amount;

                            //<<06APr2017

                            //>>06July2017 GST
                            CLEAR(vCGST);
                            CLEAR(vSGST);
                            CLEAR(vIGST);
                            CLEAR(vCGSTRate);
                            CLEAR(vSGSTRate);
                            CLEAR(vIGSTRate);

                            //>>15July2017
                            CLEAR(CGST15);
                            CLEAR(SGST15);
                            CLEAR(IGST15);
                            DetailGSTEntry.RESET;
                            DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                            DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                            DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                            DetailGSTEntry.SETRANGE(Type, DetailGSTEntry.Type::Item);
                            DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "Item No.");
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

                                    //>>21July2017 UTGST
                                    IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                        vSGST := ABS(DetailGSTEntry."GST Amount");
                                        vSGSTRate := DetailGSTEntry."GST %";

                                        IF DetailGSTEntry."GST Amount" = 0 THEN
                                            SGST15 := ''
                                        ELSE
                                            SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                                    END;
                                    //<<21July2017 UTGST

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
                            //<<06July2017 GST


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

                            /* recStrOrderLineDet.RESET;
                            recStrOrderLineDet.SETRANGE("Invoice No.", "Document No.");
                            recStrOrderLineDet.SETRANGE("Item No.", "Item No.");
                            recStrOrderLineDet.SETRANGE("Line No.", "Line No.");
                            IF recStrOrderLineDet.FINDSET THEN
                                REPEAT

                                    IF recStrOrderLineDet."Tax/Charge Group" = 'DISCOUNT' THEN BEGIN

                                        DiscAmt06 := recStrOrderLineDet.Amount;

                                    END;

                                    IF recStrOrderLineDet."Tax/Charge Group" = 'FREIGHT' THEN BEGIN

                                        ChargeAmt06 := recStrOrderLineDet.Amount;

                                    END;

                                UNTIL recStrOrderLineDet.NEXT = 0; */
                            //<<06July2017 StrucureDetails

                            //>>27July2017
                            CLEAR(MRPofItem);
                            recMRPmaster.RESET;
                            recMRPmaster.SETRANGE(recMRPmaster."Item No.", "Transfer Shipment Line"."Item No.");
                            recMRPmaster.SETRANGE(recMRPmaster."Lot No.", BatchNo);
                            IF recMRPmaster.FINDFIRST THEN BEGIN
                                MRPofItem := recMRPmaster."MRP Price" * recMRPmaster."Qty. per Unit of Measure";
                            END;


                            //<<27July2017

                            DocTotal += Amount + 0;//"BED Amount" + "eCess Amount"
                            //+ "SHE Cess Amount" + FrieghtAmount + "ADE Amount" + "Total GST Amount";


                            /*
                            DocTotal:=ROUND((
                            "Transfer Shipment Line".Amount + "Transfer Shipment Line"."BED Amount"+"Transfer Shipment Line"."eCess Amount"
                            +"Transfer Shipment Line"."SHE Cess Amount"+FrieghtAmount+"Transfer Shipment Line"."ADE Amount"+RoundOffAmnt),1.0,'=');
                            */
                            /*
                            FormCodes.SETRANGE(FormCodes.Code,"Form Code");
                            IF FormCodes.FINDFIRST THEN;
                            */
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
                        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START
                        IF ("Transfer Shipment Header"."Print BRT" = FALSE) THEN BEGIN
                            IF i = 1 THEN
                                InvType := 'ORIGINAL FOR RECIPIENT'
                            //InvType := 'ORIGINAL FOR BUYER'
                            ELSE
                                IF i = 2 THEN
                                    InvType := 'DUPLICATE FOR TRANSPORTER'
                                ELSE
                                    IF i = 3 THEN
                                        InvType := 'TRIPLICATE FOR SUPPLIER'
                                    //InvType := 'TRIPLICATE FOR ASSESSEE'
                                    ELSE
                                        IF i = 4 THEN BEGIN
                                            InvType := 'QUADRAPLICATE FOR H.O';
                                            //InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                                            //InvType1 := 'NOT FOR CENVAT';
                                        END ELSE
                                            IF i = 5 THEN BEGIN
                                                InvType := 'EXTRA COPY';
                                                //InvType := 'EXTRA COPY NOT FOR CENVAT';
                                                InvType1 := '';
                                            END
                                            ELSE
                                                InvType := '';
                        END;

                        //IF ("Transfer Shipment Header"."Print BRT" = TRUE) AND (USERID = 'sa') THEN
                        IF ("Transfer Shipment Header"."Print BRT" = TRUE) AND (USERID = 'GPUAE\UNNIKRISHNAN.VS') THEN //18July2017
                        BEGIN
                            IF i = 1 THEN
                                InvType := 'ORIGINAL FOR RECIPIENT'
                            //InvType := 'ORIGINAL FOR BUYER'
                            //InvType := 'ORIGINAL FOR BUYER'
                            ELSE
                                IF i = 2 THEN
                                    InvType := 'DUPLICATE FOR TRANSPORTER'
                                ELSE
                                    IF i = 3 THEN
                                        InvType := 'TRIPLICATE FOR SUPPLIER'
                                    //InvType := 'TRIPLICATE FOR ASSESSEE'
                                    ELSE
                                        IF i = 4 THEN BEGIN
                                            InvType := 'QUADRAPLICATE FOR H.O';
                                            //InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                                            //InvType1 := 'NOT FOR CENVAT';
                                        END ELSE
                                            IF i = 5 THEN BEGIN
                                                InvType := 'EXTRA COPY NOT FOR CENVAT';
                                                InvType1 := '';
                                            END
                                            ELSE
                                                InvType := '';
                        END;

                        //IF ("Transfer Shipment Header"."Print BRT"= TRUE) AND (USERID <> 'sa')THEN
                        IF ("Transfer Shipment Header"."Print BRT" = TRUE) AND (USERID <> 'GPUAE\UNNIKRISHNAN.VS') THEN //18July2017
                        BEGIN
                            IF i = 1 THEN
                                InvType := 'EXTRA COPY';
                            InvType1 := ' NOT FOR COMMERCIAL USE';
                        END;

                        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
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

                    DocTotal := 0; //06Apr2017
                    //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START

                    //IF NOT(USERID = 'sa') THEN
                    IF NOT (USERID = 'GPUAE\UNNIKRISHNAN.VS') THEN BEGIN
                        IF "Transfer Shipment Header"."Print BRT" = TRUE THEN BEGIN
                            SETRANGE(Number, 1, 1);
                            //  SETRANGE(Number,1,ABS(NoOfCopies));
                        END ELSE BEGIN

                            NoOfCopies := 4;
                            NoOfLoops := ABS(NoOfCopies);//06July2017
                                                         //NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
                            IF NoOfLoops <= 0 THEN
                                NoOfLoops := 1;
                            copytext := '';
                            SETRANGE(Number, 1, NoOfLoops);
                        END;
                    END ELSE BEGIN

                        NoOfCopies := 4;
                        NoOfLoops := ABS(NoOfCopies);//06July2017
                                                     //NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
                        IF NoOfLoops <= 0 THEN
                            NoOfLoops := 1;
                        copytext := '';
                        SETRANGE(Number, 1, NoOfLoops);
                    END;

                    //SETRANGE(Number,1,NoOfCopies);
                    //MESSAGE('%1',ABS(NoOfCopies));
                    //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
                    //i:= 1
                end;
            }

            trigger OnAfterGetRecord()
            begin


                //>>03Aug2017
                IF "Posting Date" < 20170108D THEN //010817D 
                    ERROR('This Report is Applicable from August2017 Transfer Invoices');

                //>>03Aug2017

                //>>RB-N 07Nov2017 Logo Baseon ItemCategory
                CLEAR(Cat21);
                CLEAR(Logo21);
                TSL25.RESET;
                TSL25.SETRANGE("Document No.", "No.");
                IF TSL25.FINDFIRST THEN BEGIN
                    Cat21 := TSL25."Item Category Code";

                END;

                IF Cat21 = 'CAT15' THEN
                    Logo21 := 1
                ELSE
                    Logo21 := 3;
                //>>RB-N 07Nov2017 Logo Baseon ItemCategory

                //>>RB-N 24Nov2017 Merch Item Logo
                CLEAR(Mtext);
                TSL24.RESET;
                TSL24.SETRANGE("Document No.", "No.");
                IF TSL24.FINDFIRST THEN BEGIN
                    IF TSL24."Item Category Code" = 'CAT14' THEN BEGIN
                        Mtext := COPYSTR(TSL24.Description, 1, 3);
                        /*
                        IF Mtext = 'Rep' THEN
                          Logo21 := 1
                        ELSE
                          Logo21 := 3;
                        */
                    END;
                END;

                IF Mtext = 'Rep' THEN
                    Logo21 := 1;

                //<<RB-N 24Nov2017 Merch Item Logo

                //>>04Aug2017
                recInvComment.RESET;
                recInvComment.SETRANGE(recInvComment."Document Type", recInvComment."Document Type"::"Posted Transfer Shipment");
                recInvComment.SETRANGE(recInvComment."No.", "Transfer Shipment Header"."No.");
                IF recInvComment.FINDFIRST THEN

                    //<<04Aug2017


                    //>>27July2017 FactoryCaption
                    CLEAR(FactoryCaption);
                IF ("Transfer-from Code" = 'PLANT01') OR ("Transfer-from Code" = 'PLANT02') THEN
                    FactoryCaption := 'Factory'
                ELSE
                    FactoryCaption := 'Matrl. Sold from';
                //<<27July2017 FactoryCaption

                //>>06Apr2017
                IF LocFrom.GET("Transfer-from Code") THEN BEGIN
                    IF RecState.GET(LocFrom."State Code") THEN
                        LocState := RecState.Description;
                END;
                //<<06Apr2017


                //IF Loc.GET("Transfer-from Code") THEN;
                //RecState.GET(Loc."State Code");



                InvNoLen := STRLEN("No.");
                InvNo := COPYSTR("No.", (InvNoLen - 3), 4);

                /*
                //EBT STIVAN---(29012013)---Error Message POP UP,if ECC No of Transfer from Code is not Specified----START
                recLocation.RESET;
                recLocation.SETRANGE(recLocation.Code,"Transfer Shipment Header"."Transfer-from Code");
                IF recLocation.FINDFIRST THEN
                BEGIN
                 recECC.RESET;
                 recECC.SETRANGE(recECC.Code,recLocation."E.C.C. No.");
                 IF  recECC.FINDFIRST THEN
                 BEGIN
                  IF recECC.Description = '' THEN
                  ERROR('Your Location is not Excise Registered');
                 END;
                END;
                //EBT STIVAN---(29012013)---Error Message POP UP,if ECC No of Transfer from Code is not Specified------END
                */
                /*
                Loc.RESET;
                Loc.SETRANGE(Loc.Code,"Sales Invoice Header"."Location Code");
                IF Loc.FINDFIRST THEN;
                "E.C.C. Nos.1".GET(Loc."E.C.C. No.");
                */

                /*
                RecState.RESET;
                RecState.GET(Loc."State Code");
                RecCountry.RESET;
                RecCountry.GET(Loc."Country/Region Code");
                */

                /*
                Loc1.RESET;
                Loc1.SETRANGE(Loc1.Code,"Sales Invoice Header"."Location Code");
                IF Loc1.FINDFIRST THEN;
                "E.C.C. Nos.".GET(Loc1."E.C.C. No.");
                */

                CurrDatetime := TIME;
                TotalQty := 0;

                CLEAR(FrieghtAmount);
                CLEAR(AdDutyAmount);
                /* PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'FREIGHT');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                    FrieghtAmount := PostedStrOrdrDetails."Calculation Value";

                // EBT Start 10.08.10 start

                PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails.Type, PostedStrOrdrDetails.Type::Transfer);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'ADDL.DUTY');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                    AdDutyAmount := PostedStrOrdrDetails."Calculation Value"; */

                // EBT Start 10.08.10 end



                recShippingAgent.RESET;
                recShippingAgent.SETRANGE(recShippingAgent.Code, "Shipping Agent Code");
                IF recShippingAgent.FINDFIRST THEN;



                CLEAR(FinalAmt);
                SL14.RESET;
                SL14.SETRANGE("Document No.", "No.");
                IF SL14.FINDSET THEN
                    REPEAT
                        FinalAmt += SL14.Amount + 0;// SL14."Total GST Amount";
                    //MESSAGE(' %1 Final Amt \\  %2 LineNo.',FinalAmt,SL14."Line No.");
                    UNTIL SL14.NEXT = 0;


                //>>27July2017 Roundoff Amount
                CLEAR(RoundOffAmnt);
                CLEAR(InvoiceRoundOff);
                InvoiceRoundOff := ROUND(FinalAmt, 1.0);
                RoundOffAmnt := InvoiceRoundOff - FinalAmt;
                //<<27July2017 Roundoff Amount

                //>>AmountinWords
                vCheck.InitTextVariable;
                vCheck.FormatNoText(AmtinWord15, ROUND((FinalAmt + RoundOffAmnt), 0.01), '');
                AmtinWord15[1] := DELCHR(AmtinWord15[1], '=', '*');

                //<<AmountinWords

                //>>17July2017TotalGST Amount
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

                        //>>21July2017 UTGST
                        IF DGST04."GST Component Code" = 'UTGST' THEN BEGIN
                            TotalSGST += ABS(DGST04."GST Amount");
                        END;

                        //<<21July2017 UTGST
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

                //<<17July2017TotalGST Amount

                /*
                //>>15Jul2017
                //Bill State Description
                RecState.RESET;
                IF RecState.GET("Bill to Customer State") THEN
                BEGIN
                  BillState := RecState.Description;
                  BillGSTStateCode := RecState."State Code (GST Reg. No.)";
                END;
                
                
                //Bill GSTNo
                Cus12.RESET;
                IF Cus12.GET("Sales Invoice Header"."Bill-to Customer No.") THEN
                  BillGSTNo := Cus12."GST Registration No.";
                
                //Bill Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Bill-to Country/Region Code") THEN
                  BillCountry := RecCountry.Name;
                
                
                
                //Ship Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Ship-to Country/Region Code") THEN
                  ShipCountry := RecCountry.Name;
                
                
                //Ship2Name
                IF "Sales Invoice Header"."Ship-to Code" = '' THEN
                  Ship2Name := "Sales Invoice Header"."Full Name"
                ELSE
                  Ship2Name := "Sales Invoice Header"."Ship-to Name";
                
                //Shipping
                
                IF ShipToCode.GET("Sell-to Customer No.","Ship-to Code") THEN
                BEGIN
                
                      RecState.RESET;
                      IF RecState.GET(ShipToCode.State) THEN
                      BEGIN
                        ShipState := RecState.Description;
                        ShipGSTStateCode :=RecState."State Code (GST Reg. No.)";
                        ShipGSTNo := ShipToCode."GST Registration No.";
                      END;
                
                END ELSE
                BEGIN
                
                  ShipState := BillState;
                  ShipGSTStateCode := BillGSTStateCode;
                  ShipGSTNo := BillGSTNo;
                
                END;
                
                
                //>>15Jul2017
                */

                //>>27July2017 Location Address
                CLEAR(BillGSTNo);
                CLEAR(BillState);
                CLEAR(BillGSTStateCode);
                CLEAR(BillCountry);
                IF Loc1.GET("Transfer-to Code") THEN BEGIN
                    BillGSTNo := Loc1."GST Registration No.";

                    IF RecState1.GET(Loc1."State Code") THEN BEGIN
                        BillGSTStateCode := RecState1."State Code (GST Reg. No.)";
                        BillState := RecState1.Description;
                    END;

                END;

                //Bill Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Trsf.-to Country/Region Code") THEN
                    BillCountry := RecCountry.Name;
                //<<27July2017 Location Address

                //RSPLSUM 30Jul2020>>
                CLEAR(EWayBillNo);
                "Transfer Shipment Header".CALCFIELDS("Transfer Shipment Header"."EWB No.");
                IF "Transfer Shipment Header"."EWB No." <> '' THEN BEGIN
                    EWayBillNo := "Transfer Shipment Header"."EWB No.";
                END ELSE BEGIN
                    EWayBillNo := "Transfer Shipment Header"."Road Permit No.";
                END;
                //RSPLSUM 30Jul2020<<

                //RSPLSUM BEGIN>>
                CLEAR(IRN);
                "Transfer Shipment Header".CALCFIELDS("Transfer Shipment Header".IRN);
                IF "Transfer Shipment Header".IRN <> '' THEN
                    IRN := 'IRN: ' + "Transfer Shipment Header".IRN;
                //RSPLSUM END>>

                Ctr := 1;

                RecGSTLedgEntry.RESET;
                RecGSTLedgEntry.SETRANGE("Document No.", "Transfer Shipment Header"."No.");
                RecGSTLedgEntry.SETRANGE("Document Type", RecGSTLedgEntry."Document Type"::Invoice);
                RecGSTLedgEntry.SETRANGE("Transaction Type", RecGSTLedgEntry."Transaction Type"::Sales);
                IF RecGSTLedgEntry.FINDFIRST THEN BEGIN
                    RecGSTLedgEntry.CALCFIELDS(RecGSTLedgEntry."QR Code");
                END;

                //BarcodeForQuarantineLabel;//RSPLSUM

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
                    Caption = 'No.Of Copies';
                }
                field(ShowInternalInfo; ShowInternalInfo)
                {
                    Caption = 'Show Internal Information';
                }
                field(LogInteraction; LogInteraction)
                {
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

    trigger OnPostReport()
    begin

        //>>26July2017
        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START
        IF "Transfer Shipment Header"."Print BRT" = FALSE THEN BEGIN
            IF NOT (CurrReport.PREVIEW) THEN BEGIN
                recTSH.RESET;
                recTSH.SETRANGE(recTSH."No.", "Transfer Shipment Header"."No.");
                recTSH.SETRANGE(recTSH."Print BRT", FALSE);
                IF recTSH.FINDFIRST THEN BEGIN
                    recTSH."Print BRT" := TRUE;
                    recTSH.MODIFY;
                END;
            END;
        END;
        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END

        //<<26July2017
    end;

    trigger OnPreReport()
    begin

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
        CompanyInfo1.CALCFIELDS("Repsol Logo");//07Nov2017

        //InitTextVariable;
    end;

    var
        Loc: Record 14;
        Loc1: Record 14;
        SalesPerson: Record 13;
        RespCentre: Record 5714;
        RecState: Record State;
        RecState1: Record "State";
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
        //"E.C.C. Nos.": Record "13708";
        //"E.C.C. Nos.1": Record "13708";
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
        SegManagement: Codeunit SegManagement;
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
        //FormCodes: Record "13756";
        TransHdr: Record 5740;
        ItemVarParmResFinal: Record 50015;
        FrieghtAmount: Decimal;
        AdDutyAmount: Decimal;
        //PostedStrOrdrDetails: Record "13760";
        TotalQty: Decimal;
        BEDpercent: Decimal;
        // recExcisePostnSetup: Record "13711";
        BatchNo: Code[20];
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
        //recECC: Record "13708";
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
        //recStrOrderLineDet: Record "13798";
        SL14: Record 5745;
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
        Cus12: Record 18;
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
        "---27July2017----RoundingAmt": Integer;
        InvoiceRoundOff: Decimal;
        MRPofItem: Decimal;
        recMRPmaster: Record 50013;
        FactoryCaption: Text[30];
        "-------07Nov2017": Integer;
        TSL25: Record 5745;
        Cat21: Code[20];
        Logo21: Integer;
        "---24Nov2017": Integer;
        Mtext: Text;
        TSL24: Record 5745;
        EWayBillNo: Code[20];
        RecGSTLedgEntry: Record "GST Ledger Entry";
        IRN: Text[255];

    // [Scope('Internal')]
    procedure BarcodeForQuarantineLabel()
    var
        Text5000_Ctx: Label '<xpml><page quantity=''0'' pitch=''74.1 mm''></xpml>SIZE 100 mm, 74.1 mm';
        Text5001_Ctx: Label 'DIRECTION 0,0';
        Text5002_Ctx: Label 'REFERENCE 0,0';
        Text5003_Ctx: Label 'OFFSET 0 mm';
        Text5004_Ctx: Label 'SET PEEL OFF';
        Text5005_Ctx: Label 'SET CUTTER OFF';
        Text5006_Ctx: Label 'SET PARTIAL_CUTTER OFF';
        Text5007_Ctx: Label '<xpml></page></xpml><xpml><page quantity=''1'' pitch=''74.1 mm''></xpml>SET TEAR ON';
        Text5008_Ctx: Label 'CLS';
        Text5009_Ctx: Label 'CODEPAGE 1252';
        Text5010_Ctx: Label 'TEXT 806,792,"0",180,17,14,"SCHILLER_"';
        Text5011_Ctx: Label 'TEXT 1093,443,"0",180,14,12,"Item"';
        Text5012_Ctx: Label 'TEXT 1093,303,"0",180,10,12,"Part No."';
        Text5013_Ctx: Label 'TEXT 1093,166,"0",180,12,12,"Sr. No."';
        Text5014_Ctx: Label 'TEXT 933,443,"0",180,14,12,":"';
        Text5015_Ctx: Label 'TEXT 933,303,"0",180,14,12,":"';
        Text5016_Ctx: Label 'TEXT 933,166,"0",180,14,12,":"';
        Text5017_Ctx: Label 'TEXT 896,303,"0",180,10,12,"%1"';
        Text5018_Ctx: Label 'TEXT 896,166,"0",180,10,12,"%1"';
        Text5019_Ctx: Label 'QRCODE 330,300,L,10,A,180,M2,S7,"%1"';
        Text5020_Ctx: Label 'TEXT 896,443,"0",180,10,12,"%1"';
        Text5021_Ctx: Label 'PRINT 1,1';
        Text5022_Ctx: Label '<xpml></page></xpml><xpml><end/></xpml>';
        QRCodeInput: Text;
    //   TempBlob: Record "99008535";
    begin
        RecGSTLedgEntry.RESET;
        RecGSTLedgEntry.SETRANGE("Document No.", "Transfer Shipment Header"."No.");
        RecGSTLedgEntry.SETRANGE("Document Type", RecGSTLedgEntry."Document Type"::Invoice);
        RecGSTLedgEntry.SETRANGE("Transaction Type", RecGSTLedgEntry."Transaction Type"::Sales);
        IF RecGSTLedgEntry.FINDFIRST THEN BEGIN
            QRCodeInput := STRSUBSTNO('%1', RecGSTLedgEntry."Scan QrCode E-Invoice");
            IF QRCodeInput <> '' THEN BEGIN
                //CreateQRCode(QRCodeInput, TempBlob);
                // RecGSTLedgEntry."QR Code" := TempBlob.Blob;
                RecGSTLedgEntry.MODIFY;
            END;
        END;
    end;

    /*    local procedure CreateQRCode(QRCodeInput: Text[500]; var TempBLOB: Record "99008535")
       var
           QRCodeFileName: Text[1024];
       begin
           CLEAR(TempBLOB);
           QRCodeFileName := GetQRCode(QRCodeInput);
         //  UploadFileBLOBImportandDeleteServerFile(TempBLOB, QRCodeFileName);
       end; */

    /*  local procedure GetQRCode(QRCodeInput: Text[300]) QRCodeFileName: Text[1024]
     var
         [RunOnClient]
         IBarCodeProvider: DotNet IBarcodeProvider;
     begin
         GetBarCodeProvider(IBarCodeProvider);
         QRCodeFileName := IBarCodeProvider.GetBarcode(QRCodeInput);
     end; */

    // [Scope('Internal')]
    /*  procedure UploadFileBLOBImportandDeleteServerFile(var TempBlob: Record "99008535"; FileName: Text[1024])
     var
         FileManagement: Codeunit 419;
     begin
         FileName := FileManagement.UploadFileSilent(FileName);
         FileManagement.BLOBImportFromServerFile(TempBlob, FileName);
         DeleteServerFile(FileName);
     end; */

    // [Scope('Internal')]
    /* procedure GetBarCodeProvider(var IBarCodeProvider: DotNet IBarcodeProvider)
    var
        [RunOnClient]
        QRCodeProvider: DotNet QRCodeProvider;
    begin
        IF ISNULL(IBarCodeProvider) THEN
            IBarCodeProvider := QRCodeProvider.QRCodeProvider;
    end;

    local procedure DeleteServerFile(ServerFileName: Text)
    begin
        IF ERASE(ServerFileName) THEN;
    end; */
}


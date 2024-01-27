report 50042 "Delivery Challan-Invoice- GST"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 08Sep2017   RB-N         Dispalying RespolLogo Or IpolLogo Depending upon ItemCategory
    // 26Oct2017   RB-N         Print Invoice Tick
    // 14Nov2017   RB-N         Not to allowing SaveasPDF from Preview Mode
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/DeliveryChallanInvoiceGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    Permissions = TableData 112 = rim;
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            RequestFilterFields = "No.";
            column(Logo21; Logo21)
            {
            }
            column(Cat21; Cat21)
            {
            }
            column(CompName; CompanyInfo1.Name)
            {
            }
            column(CompPicture; CompanyInfo1.Picture)
            {
            }
            column(IpolLogo; CompanyInfo1."Name Picture")
            {
            }
            column(CompShadedBoxPicture; CompanyInfo1."Shaded Box")
            {
            }
            column(RepsolLogo; CompanyInfo1."Repsol Logo")
            {
            }
            column(FactoryLocation; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' - ' + LocFrom."Post Code" + ' ' + LocState + ' ' + ' Phone: ' + LocFrom."Phone No." + ' Email: ' + LocFrom."E-Mail")
            {
            }
            column(LocationCity; LocFrom.City)
            {
            }
            column(GSTSupplyCode; GSTSupplyCode)
            {
            }
            column(LocGSTNo; LocFrom."GST Registration No.")
            {
            }
            column(Remarks1; "Sales Invoice Header".Remarks)
            {
            }
            column(Remarks2; "Sales Invoice Header".Remarks2)
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
            column(FullName_SIH; "Sales Invoice Header"."Full Name")
            {
            }
            column(Location; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' ' + LocFrom."Post Code" + ' ' + LocState + ' Phone: ' + LocFrom."Phone No." + ' ' + ' Email: ' + LocFrom."E-Mail")
            {
            }
            column(BillAdd1; "Bill-to Address")
            {
            }
            column(BillAdd2; "Bill-to Address 2")
            {
            }
            column(BillAdd3; "Bill-to City" + ' - ' + "Bill-to Post Code")
            {
            }
            column(SIH_BillAdd3; "Bill-to City" + '-' + "Bill-to Post Code" + ' ' + BillState + ', ' + BillCountry)
            {
            }
            column(BillState; BillState)
            {
            }
            column(BillCity; "Bill-to City")
            {
            }
            column(ShipGSTStateCode; ShipGSTStateCode)
            {
            }
            column(ShipGSTNo; ShipGSTNo)
            {
            }
            column(ShipPANNo; ShipPANNo)
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
            column(BillPANNo; BillPANNo)
            {
            }
            column(SIH_ShipName; "Sales Invoice Header"."Ship-to Name")
            {
            }
            column(SIH_ShipAdd1; "Sales Invoice Header"."Ship-to Address")
            {
            }
            column(SIH_ShipAdd2; "Sales Invoice Header"."Ship-to Address 2")
            {
            }
            column(SIH_ShipAdd3; "Ship-to City" + '-' + "Ship-to Post Code" + ' ' + ShipState + ', ' + ShipCountry)
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
            column(RangDivision; '"E.C.C. Nos."."C.E. Range"' + '&' + '"E.C.C. Nos."."C.E. Division"' + '& ' + '"E.C.C. Nos."."C.E. Commissionerate"')
            {
            }
            column(CEAddr; '"E.C.C. Nos."."C.E Address1"')
            {
            }
            column(CEAddr2; '"E.C.C. Nos."."C.E Address2"')
            {
            }
            column(CECityPostCode; '"E.C.C. Nos."."C.E City"' + '-' + '"E.C.C. Nos."."C.E Post Code"')
            {
            }
            column(TINNo; 'Loc1."T.I.N. No." ' + '  &  ' + 'Loc1."C.S.T No."')
            {
            }
            column(PostingDate; "Posting Date")
            {
            }
            column(BRTTime; "Sales Invoice Header"."Invoice Print Time")
            {
            }
            column(DocNo; "No.")
            {
            }
            column(RoadPermitNo; EWayBillNo)
            {
            }
            column(Name_ShipingAgent; recShippingAgent.Name)
            {
            }
            column(VehicleNo; "Vehicle No.")
            {
            }
            column(LRRRDate; "LR/RR Date")
            {
            }
            column(LRRRNo; "LR/RR No.")
            {
            }
            column(FrieghtType; "Freight Type")
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
            column(CompCINNo; 'CompanyInfo1."Company Registration  No."')
            {
            }
            column(CompPANNo; CompanyInfo1."P.A.N. No.")
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
                    dataitem("Sales Invoice Line"; "Sales Invoice Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Sales Invoice Header";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Type = FILTER(Item),
                                                  Quantity = FILTER(<> 0));
                        column(TotalGSTAmount; 0)// "Sales Invoice Line"."Total GST Amount")
                        {
                        }
                        column(LineDiscountAmt; "Sales Invoice Line"."Line Discount Amount")
                        {
                        }
                        column(GSTBaseAmt; 0)// "Sales Invoice Line"."GST Base Amount")
                        {
                        }
                        column(AmountToCustomer; 0)// "Sales Invoice Line"."Amount To Customer")
                        {
                        }
                        column(ChargesToCustomer; '"Sales Invoice Line"."Charges To Customer"')
                        {
                        }
                        column(GST_Line; 0)// "GST %")
                        {
                        }
                        column(GSTGroupCode; "GST Group Code")
                        {
                        }
                        column(HSNSACCode; "HSN/SAC Code")
                        {
                        }
                        column(LineNo; "Line No.")
                        {
                        }
                        column(DimensionName; DimensionName)
                        {
                        }
                        column(LineNo1; "Line No.")
                        {
                        }
                        column(DocumentNo; "Document No.")
                        {
                        }
                        column(BatchNo; "Lot No.")
                        {
                        }
                        column(QtyperUnitofMeasure; "Qty. per Unit of Measure")
                        {
                        }
                        column(QuantityBase; "Quantity (Base)")
                        {
                        }
                        column(Quantity; Quantity)
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
                        column(ExciseProdPostingGroup; '"Excise Prod. Posting Group"')
                        {
                        }
                        column(BEDpercent; BEDpercent)
                        {
                        }
                        column(BEDant_QtyBase; 0/*"BED Amount" *// "Quantity (Base)")
                        {
                        }
                        column(BEDAmount1; 0)//"BED Amount")
                        {
                        }
                        column(SHE_Cess; 0)//"SHE Cess Amount" + "eCess Amount")
                        {
                        }
                        column(Amount; Amount)
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
                        column(BEDAmount; 0)// "BED Amount")
                        {
                        }
                        column(eCessAmount; 0)// "eCess Amount")
                        {
                        }
                        column(ADEAmount; 0)//"ADE Amount")
                        {
                        }
                        column(SHECessAmount; 0)//"SHE Cess Amount")
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

                        trigger OnAfterGetRecord()
                        begin
                            vCount := COUNT;
                            Ctr := Ctr + 1;


                            IF recItemEBT.GET("Sales Invoice Line"."No.") THEN;
                            IF (recItemEBT."Item Category Code" = 'CAT02') OR (recItemEBT."Item Category Code" = 'CAT03') OR
                            (recItemEBT."Item Category Code" = 'CAT11') OR (recItemEBT."Item Category Code" = 'CAT12') OR
                            (recItemEBT."Item Category Code" = 'CAT13') THEN BEGIN
                                "Batch/UOM" := '   Batch        No.';
                            END ELSE
                                "Batch/UOM" := '  Unit of Measure';


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
                                    recILE.SETRANGE("Item No.", "Sales Invoice Line"."No.");
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

                            IF "Sales Invoice Line".Type = "Sales Invoice Line".Type::"G/L Account" THEN
                                DimensionName := Description;

                            //>>21Aug2017 RepsolName

                            IF ("Item Category Code" = 'CAT15') OR ("Item Category Code" = 'CAT14') THEN
                                DimensionName := Description;


                            IF "Item Category Code" = 'CAT18' THEN
                                DimensionName := Description;
                            //>>21Aug2017 RepsolName
                            //>>14Apr2017

                            /*  recExcisePostnSetup.RESET;
                             recExcisePostnSetup.SETRANGE(recExcisePostnSetup."Excise Bus. Posting Group", "Excise Bus. Posting Group");
                             recExcisePostnSetup.SETFILTER(recExcisePostnSetup."From Date", '<=%1', "Posting Date"); //RSPL008
                             IF recExcisePostnSetup.FINDLAST THEN BEGIN
                                 ecesspercent := recExcisePostnSetup."eCess %";
                                 BEDpercent := recExcisePostnSetup."BED %";
                                 SHEpercent := recExcisePostnSetup."SHE Cess %";
                             END; */


                            //RecState.SETRANGE(RecState.Code,Loc."State Code");
                            //RecState.FINDFIRST;
                            //State:=RecState.Description;

                            IF RecItem.GET("No.") THEN;
                            TotalQty += "Quantity (Base)";
                            TotalBR += "Unit Price";
                            TotalValue := TotalValue + ("Quantity (Base)" * "Unit Price");


                            TaxDescr := 'Tax Descr';

                            //>>06APr2017

                            FExcise += 0;//"BED Amount";
                            FEcess += 0;// "eCess Amount";
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
                            DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Sales Invoice Header"."No.");
                            DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                            DetailGSTEntry.SETRANGE(Type, DetailGSTEntry.Type::Item);
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
                                        //>>17July2017
                                        IF FinalAmt = 0 THEN BEGIN
                                            vCGST := 0;
                                            vCGSTRate := 0;
                                            CGST15 := '';
                                        END;
                                        //<<17July2017
                                    END;

                                    IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                        vSGST := ABS(DetailGSTEntry."GST Amount");
                                        vSGSTRate := DetailGSTEntry."GST %";

                                        IF DetailGSTEntry."GST Amount" = 0 THEN
                                            SGST15 := ''
                                        ELSE
                                            SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';
                                        //>>17July2017
                                        IF FinalAmt = 0 THEN BEGIN
                                            vSGST := 0;
                                            vSGSTRate := 0;
                                            SGST15 := '';
                                        END;
                                        //<<17July2017
                                    END;

                                    IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                        vIGST := ABS(DetailGSTEntry."GST Amount");
                                        vIGSTRate := DetailGSTEntry."GST %";
                                        //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");

                                        IF DetailGSTEntry."GST Amount" = 0 THEN
                                            IGST15 := ''
                                        ELSE
                                            IGST15 := ' @ ' + FORMAT(vIGSTRate) + ' %';

                                        //>>17July2017
                                        IF FinalAmt = 0 THEN BEGIN
                                            vIGST := 0;
                                            vIGSTRate := 0;
                                            IGST15 := '';
                                        END;
                                        //<<17July2017
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

                            /*   recStrOrderLineDet.RESET;
                              recStrOrderLineDet.SETRANGE("Invoice No.", "Document No.");
                              recStrOrderLineDet.SETRANGE("Item No.", "No.");
                              recStrOrderLineDet.SETRANGE("Line No.", "Line No.");
                              IF recStrOrderLineDet.FINDSET THEN
                                  REPEAT

                                      IF recStrOrderLineDet."Tax/Charge Group" = 'DISCOUNT' THEN BEGIN

                                          DiscAmt06 := recStrOrderLineDet.Amount;

                                      END; */

                            // IF recStrOrderLineDet."Tax/Charge Group" = 'FREIGHT' THEN BEGIN

                            ChargeAmt06 := 0;//recStrOrderLineDet.Amount;

                            // END;

                            //  UNTIL recStrOrderLineDet.NEXT = 0;
                            //<<06July2017 StrucureDetails

                            DocTotal += Amount + 0;//"BED Amount" + "eCess Amount"
                                                   // + "SHE Cess Amount" + FrieghtAmount + "ADE Amount" + "Total GST Amount";

                            /*
                            DocTotal:=ROUND((
                            "Transfer Shipment Line".Amount + "Transfer Shipment Line"."BED Amount"+"Transfer Shipment Line"."eCess Amount"
                            +"Transfer Shipment Line"."SHE Cess Amount"+FrieghtAmount+"Transfer Shipment Line"."ADE Amount"+RoundOffAmnt),1.0,'=');
                            */

                            //  FormCodes.SETRANGE(FormCodes.Code, "Form Code");
                            // IF FormCodes.FINDFIRST THEN;

                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineDuty, ROUND(FExcise, 0.01), ''); //06Apr2017
                            vCheck.FormatNoText(DescriptionLineDuty, ROUND(/*"BED Amount"*/0, 0.01), '');
                            DescriptionLineDuty[1] := DELCHR(DescriptionLineDuty[1], '=', '*');

                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineeCess, ROUND(FEcess, 0.01), ''); //06Apr2017
                            vCheck.FormatNoText(DescriptionLineeCess, ROUND(/*"eCess Amount"*/0, 0.01), '');
                            DescriptionLineeCess[1] := DELCHR(DescriptionLineeCess[1], '=', '*');

                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineSHeCess, ROUND(FShe, 0.01), ''); //06Apr2017
                            vCheck.FormatNoText(DescriptionLineSHeCess, ROUND(/*"SHE Cess Amount"*/0, 0.01), '');
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
                        IF ("Sales Invoice Header"."Print Invoice" = FALSE) THEN BEGIN
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

                        //IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'sa') THEN
                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'GPUAE\UNNIKRISHNAN.VS') THEN //RB-N 26Oct2017
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

                        //IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'sa')THEN
                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'GPUAE\UNNIKRISHNAN.VS') THEN //RB-N 26Oct2017
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
                    IF NOT (USERID = 'GPUAE\UNNIKRISHNAN.VS') THEN //RB-N 26Oct2017
                    BEGIN
                        IF "Sales Invoice Header"."Print Invoice" = TRUE THEN BEGIN
                            SETRANGE(Number, 1, 1);
                            //  SETRANGE(Number,1,ABS(NoOfCopies));
                        END ELSE BEGIN
                            DocTotal := 0; //06Apr2017

                            NoOfCopies := 4;
                            NoOfLoops := ABS(NoOfCopies);//06July2017
                                                         //NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
                            IF NoOfLoops <= 0 THEN
                                NoOfLoops := 1;
                            copytext := '';
                            SETRANGE(Number, 1, NoOfLoops);
                        END
                    END ELSE BEGIN
                        DocTotal := 0; //06Apr2017

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

                IF NOT CurrReport.PREVIEW THEN;//RB-N 14Nov2017

                //>>RB-N 08Sep2017 Logo Baseon ItemCategory
                CLEAR(Cat21);
                CLEAR(Logo21);
                SL21.RESET;
                SL21.SETRANGE("Document No.", "No.");
                SL21.SETRANGE(Type, SL21.Type::Item);
                IF SL21.FINDFIRST THEN BEGIN
                    Cat21 := SL21."Item Category Code";

                END;

                IF Cat21 = 'CAT15' THEN
                    Logo21 := 1
                ELSE
                    Logo21 := 3;
                //>>RB-N 08Sep2017 Logo Baseon ItemCategory

                //>>RB-N 26Oct2017 Merch Item Logo
                CLEAR(MText);
                SL21.RESET;
                SL21.SETRANGE("Document No.", "No.");
                SL21.SETRANGE(Type, SL21.Type::Item);
                IF SL21.FINDFIRST THEN BEGIN
                    IF SL21."Item Category Code" = 'CAT14' THEN BEGIN
                        MText := COPYSTR(SL21.Description, 1, 3);

                        IF MText = 'Rep' THEN
                            Logo21 := 1
                        ELSE
                            Logo21 := 3;

                    END;
                END;

                IF MText = 'Rep' THEN
                    Logo21 := 1
                ELSE
                    Logo21 := 3;
                //<<RB-N 26Oct2017 Merch Item Logo

                //>>06Apr2017
                IF LocFrom.GET("Sales Invoice Header"."Location Code") THEN BEGIN
                    IF RecState.GET(LocFrom."State Code") THEN BEGIN
                        LocState := RecState.Description;
                        GSTSupplyCode := RecState."State Code (GST Reg. No.)";
                    END;
                END;
                //<<06Apr2017

                //IF Loc.GET("Transfer-from Code") THEN;
                //RecState.GET(Loc."State Code");

                //Loc1.GET("Transfer-to Code");
                //RecState1.GET(Loc1."State Code");

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

                Loc.RESET;
                Loc.SETRANGE(Loc.Code, "Sales Invoice Header"."Location Code");
                IF Loc.FINDFIRST THEN;
                //"E.C.C. Nos.1".GET(Loc."E.C.C. No.");

                RecState.RESET;
                RecState.GET(Loc."State Code");
                RecCountry.RESET;
                RecCountry.GET(Loc."Country/Region Code");



                Loc1.RESET;
                Loc1.SETRANGE(Loc1.Code, "Sales Invoice Header"."Location Code");
                IF Loc1.FINDFIRST THEN;
                //  "E.C.C. Nos.".GET(Loc1."E.C.C. No.");


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

                /*  PostedStrOrdrDetails.RESET;
                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails.Type, PostedStrOrdrDetails.Type::Transfer);
                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'ADDL.DUTY');
                 IF PostedStrOrdrDetails.FINDFIRST THEN
                     AdDutyAmount := PostedStrOrdrDetails."Calculation Value";
  */
                // EBT Start 10.08.10 end



                recShippingAgent.RESET;
                recShippingAgent.SETRANGE(recShippingAgent.Code, "Shipping Agent Code");
                IF recShippingAgent.FINDFIRST THEN;

                //>>14July2017 Roundoff Amount
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
                SL14.SETRANGE(Type, SL14.Type::Item);
                IF SL14.FINDSET THEN
                    REPEAT
                        FinalAmt += 0;// SL14."Amount To Customer";
                    UNTIL SL14.NEXT = 0;

                //<<14July2017 Roundoff Amount

                //>>AmountinWords
                vCheck.InitTextVariable;
                vCheck.FormatNoText(AmtinWord15, ROUND((FinalAmt + RoundOffAmnt), 1.0), '');
                AmtinWord15[1] := DELCHR(AmtinWord15[1], '=', '*');

                //<<AmountinWords

                //>>15Jul2017
                //Bill State Description
                RecState.RESET;
                //IF RecState.GET("Bill to Customer State") THEN
                IF RecState.GET(State) THEN //09Nov2017
                BEGIN
                    BillState := RecState.Description;
                    BillGSTStateCode := RecState."State Code (GST Reg. No.)";
                END;


                //Bill GSTNo
                Cus12.RESET;
                IF Cus12.GET("Sales Invoice Header"."Bill-to Customer No.") THEN
                    BillGSTNo := Cus12."GST Registration No.";
                BillPANNo := Cus12."P.A.N. No.";

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

                IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code") THEN BEGIN

                    RecState.RESET;
                    IF RecState.GET(ShipToCode.State) THEN BEGIN
                        ShipState := RecState.Description;
                        ShipGSTStateCode := RecState."State Code (GST Reg. No.)";
                        ShipGSTNo := ShipToCode."GST Registration No.";
                        ShipPANNo := ShipToCode."P.A.N. No.";
                    END;

                END ELSE BEGIN

                    ShipState := BillState;
                    ShipGSTStateCode := BillGSTStateCode;
                    ShipGSTNo := BillGSTNo;
                    ShipPANNo := BillPANNo;

                END;


                //>>15Jul2017

                //RSPLSUM 30Jul2020>>
                CLEAR(EWayBillNo);
                "Sales Invoice Header".CALCFIELDS("Sales Invoice Header"."EWB No.");
                IF "Sales Invoice Header"."EWB No." <> '' THEN BEGIN
                    EWayBillNo := "Sales Invoice Header"."EWB No.";
                END ELSE BEGIN
                    EWayBillNo := "Sales Invoice Header"."Road Permit No.";
                END;
                //RSPLSUM 30Jul2020<<

                Ctr := 1;

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
        //>>RB-N 26Oct2017
        IF "Sales Invoice Header"."Print Invoice" = FALSE THEN BEGIN
            IF NOT (CurrReport.PREVIEW) THEN BEGIN
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Header"."No.");
                recSIH.SETRANGE(recSIH."Print Invoice", FALSE);
                IF recSIH.FINDFIRST THEN BEGIN
                    //RSPLSUM 04Feb2020>>
                    SalesSetup.GET;
                    IF SalesSetup."Email Alert on Sales Inv Post" THEN BEGIN
                        IF (recSIH."Print Invoice" = FALSE) AND (recSIH."Salesperson Email Sent" = FALSE) THEN BEGIN
                            SetSalespersonEmailValues(recSIH);
                            SalespersonEmailNotification;
                            recSIH."Salesperson Email Sent" := TRUE;
                        END;
                    END;
                    //RSPLSUM 04Feb2020<<
                    recSIH."Print Invoice" := TRUE;
                    recSIH.MODIFY;
                END;
            END;
        END;
        //>>RB-N 26Oct2017
    end;

    trigger OnPreReport()
    begin

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
        CompanyInfo1.CALCFIELDS("Repsol Logo");//RB-N 08Sep2017

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
        State: Code[50];
        Ctr: Integer;
        SalesInvLine: Record 113;
        SegManagement: Codeunit SegManagement;
        LogInteraction: Boolean;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        cust: Record 18;
        copytext: Text[30];
        //SalesInvCountPrinted: Codeunit 315;
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
        // PostedStrOrdrDetails: Record 13760;
        TotalQty: Decimal;
        BEDpercent: Decimal;
        // recExcisePostnSetup: Record 13711;
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
        SL14: Record 113;
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
        GSTSupplyCode: Code[10];
        "-------08Sep2017": Integer;
        SL21: Record 113;
        Cat21: Code[20];
        Logo21: Integer;
        "---26Oct2017": Integer;
        recSIH: Record 112;
        MText: Text[30];
        SalesSetup: Record 311;
        GloInvoiceDate: Date;
        GloInvoiceNo: Code[30];
        GloInvoicevalue: Decimal;
        GloCustomerName: Text[100];
        GloVehicleno: Code[20];
        GloTransporteraName: Text[100];
        GloDriverName: Text[100];
        GloDriverMobileNo: Code[20];
        GloReciepientEmail: Text[80];
        UserSetupRec: Record 91;
        SalespersonPurchaser: Record 13;
        GloLocationName: Text[100];
        LocRec: Record 14;
        ShippingAgent: Record 291;
        SaleLines: Record 37;
        GloInvoiceQuantity: Decimal;
        GloExtDocNo: Text;
        RecSalesInvHdr: Record 112;
        Text1005: Label 'Email notification has been sent to salesperson.';
        EWayBillNo: Code[20];
        GloOrderNo: Code[20];
        BalQtyToInv: Decimal;
        // SMTPMailSetup: Record 409;
        SMTPMailSetup: Codeunit Email;
        BillPANNo: Code[10];
        ShipPANNo: Code[10];

    ////[Scope('Internal')]
    procedure SalespersonEmailNotification()
    var
        SMTPMail: Codeunit "Email Message";
        SAE18: Record 50009;
        SA18: Record 50008;
        SubjectText: Text;
        User18: Record 91;
        SenderName: Text;
        SenderEmail: Text;
        Text18: Text;
        Cust18: Record 18;
        OtAmt: Decimal;
        CrAmt: Decimal;
        ODAmt: Decimal;
        ReceiveEmail: Text;
        DimEnable: Boolean;
        Loc21: Record 14;
        SL21: Record 37;
        AppEmail: Text;
        RecSIL: Record 113;
        RecSL: Record 37;
        Emailmsg: Codeunit "Email Message";
        EmailObj: Codeunit Email;
        ReceiptType: Enum "Email Recipient Type";
    begin
        //  SMTPMailSetup.GET;//RSPLSUM02Apr21

        SubjectText := '';
        SenderName := '';
        SenderEmail := '';
        ReceiveEmail := '';
        Text18 := '';
        CLEAR(SMTPMail);
        IF UserSetupRec.GET(USERID) THEN;
        //UserSetupRec.TESTFIELD("E-Mail");

        SubjectText := 'Invoice - ' + "Sales Invoice Header"."No." + ' - ' + "Sales Invoice Header"."Sell-to Customer Name";//RSPLSUM 01Feb21

        //>>Email Body
        //RSPLSUM 01Feb21--SMTPMail.CreateMessage('',UserSetupRec."E-Mail",GloReciepientEmail,'','',TRUE);
        //RSPLSUM02Apr21--SMTPMail.CreateMessage('',UserSetupRec."E-Mail",GloReciepientEmail,SubjectText,'',TRUE);//RSPLSUM 01Feb21
        // Emailmsg.Create('', SMTPMailSetup."User ID", GloReciepientEmail, SubjectText, '', TRUE);//RSPLSUM02Apr21
        Emailmsg.Create(GloReciepientEmail, SubjectText, '', TRUE);

        //SMTPMail.AddRecipients(GloReciepientEmail);


        Emailmsg.AppendtoBody('<Br>');
        Emailmsg.AppendtoBody('<Br>');
        Emailmsg.AppendtoBody('<B> New invoice dispatched from </B>' + GloLocationName);
        Emailmsg.AppendtoBody('<Br>');
        Emailmsg.AppendtoBody('<Br>');

        //Table Start
        Emailmsg.AppendtoBody('<Table  Border="1">');
        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Invoice No </th>');
        Emailmsg.AppendtoBody('<td>' + GloInvoiceNo + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>InvoiceDate</th>');
        Emailmsg.AppendtoBody('<td>' + FORMAT(GloInvoiceDate) + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Customer Name</th>');
        Emailmsg.AppendtoBody('<td>' + GloCustomerName + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        RecSIL.RESET;
        RecSIL.SETRANGE("Document No.", GloInvoiceNo);
        RecSIL.CALCSUMS("Quantity (Base)");
        Emailmsg.AppendtoBody('<th>Quantity</th>');
        Emailmsg.AppendtoBody('<td>' + FORMAT(RecSIL."Quantity (Base)", 0, '<Integer Thousand><Decimals,3>') + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        //RSPLSUM 07Aug2020>>
        Emailmsg.AppendtoBody('<tr>');
        CLEAR(BalQtyToInv);
        RecSL.RESET;
        RecSL.SETRANGE("Document No.", GloOrderNo);
        RecSL.SETRANGE("Document Type", RecSL."Document Type"::Order);
        RecSL.SETRANGE(Type, RecSL.Type::Item);
        IF RecSL.FINDSET THEN
            REPEAT
                BalQtyToInv += RecSL."Quantity (Base)" - RecSL."Qty. Invoiced (Base)";
            UNTIL RecSL.NEXT = 0;

        Emailmsg.AppendtoBody('<th>Balance Quantity to Invoice</th>');
        Emailmsg.AppendtoBody('<td>' + FORMAT(BalQtyToInv, 0, '<Integer Thousand><Decimals,3>') + '</td>');
        Emailmsg.AppendtoBody('</tr>');
        //RSPLSUM 07Aug2020<<

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Invoice value </th>');
        Emailmsg.AppendtoBody('<td>' + FORMAT(GloInvoicevalue, 0, '<Integer Thousand><Decimals,3>') + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Purchase Order No </th>');
        Emailmsg.AppendtoBody('<td>' + GloExtDocNo + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Vehicle No </th>');
        Emailmsg.AppendtoBody('<td>' + GloVehicleno + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Transport Name </th>');
        Emailmsg.AppendtoBody('<td>' + GloTransporteraName + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Driver Name</th>');
        Emailmsg.AppendtoBody('<td>' + GloDriverName + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Driver Mobile No</th>');
        Emailmsg.AppendtoBody('<td>' + GloDriverMobileNo + '</td>');
        Emailmsg.AppendtoBody('</tr>');
        Emailmsg.AppendtoBody('</table>');
        //Table End

        Emailmsg.AppendtoBody('<Br>');
        Emailmsg.AppendtoBody('<Br>');
        EmailObj.Send(Emailmsg, Enum::"Email Scenario"::Default);
        //MESSAGE(Text1005);
        //<<Email Body
    end;

    local procedure SetSalespersonEmailValues(SalesInvHeaderP: Record 112)
    var
        CustRec: Record 18;
    begin
        IF CustRec.GET(SalesInvHeaderP."Bill-to Customer No.") THEN;
        IF SalespersonPurchaser.GET(SalesInvHeaderP."Salesperson Code") THEN;
        IF LocRec.GET(SalesInvHeaderP."Location Code") THEN;
        IF ShippingAgent.GET(SalesInvHeaderP."Shipping Agent Code") THEN;
        // SalesInvHeaderP.CALCFIELDS("Amount to Customer");
        GloInvoiceDate := SalesInvHeaderP."Posting Date";
        GloInvoiceNo := SalesInvHeaderP."No.";
        GloInvoicevalue := 0;// SalesInvHeaderP."Amount to Customer";
        GloCustomerName := CustRec.Name;
        GloVehicleno := SalesInvHeaderP."Vehicle No.";
        GloTransporteraName := ShippingAgent.Name;
        GloDriverName := '';
        GloDriverMobileNo := SalesInvHeaderP."Driver's Mobile No.";
        GloReciepientEmail := SalespersonPurchaser."E-Mail";
        GloLocationName := LocRec.Name;
        //GloInvoiceQuantity := InvoiceQuantity;
        GloExtDocNo := SalesInvHeaderP."External Document No.";
        GloOrderNo := SalesInvHeaderP."Order No.";//RSPLSUM 07Aug2020
    end;
}


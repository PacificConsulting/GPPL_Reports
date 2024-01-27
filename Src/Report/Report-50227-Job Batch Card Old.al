report 50227 "Job Batch Card Old"
{
    // BlendQty Total --> sum(Actual Qty)//02May2017
    // Production Yield -->
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 10Nov2017   RB-N         Line was Grouping for Production order Component
    //                          Secondary Order Grouping for "Prod. Order No."
    // 01Sep2018   RB-N         Standard Qty from Item BOM Components
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/JobBatchCardOld.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Production Order"; "Production Order")
        {
            DataItemTableView = SORTING("No.", Status)
                                ORDER(Ascending);
            RequestFilterFields = Status, "No.";
            column(TenBatchNo; "Production Order"."Description 2")
            {
            }
            column(ManHours; "Production Order"."Man Hours")
            {
            }
            column(MachineHours; "Production Order"."Machine Hours")
            {
            }
            column(CompanyName; COMPANYNAME)
            {
            }
            column(OrdNo; "Production Order"."No.")
            {
            }
            column(LocName; recLocation.Name)
            {
            }
            column(RoutingDscr; RoutingDscr)
            {
            }
            column(Process; 'Process: ' + FORMAT("Production Order"."Order Type"))
            {
            }
            column(Status; 'Status: ' + FORMAT("Production Order".Status))
            {
            }
            column(FinalProduct; 'Final Product: ' + "Production Order".Description)
            {
            }
            column(StartTime; CREATEDATETIME("Production Order"."Creation Date", "Production Order"."Created Time"))
            {
            }
            column(FinishTime; CREATEDATETIME("Production Order"."Finished Date", "Production Order"."Finished Time"))
            {
            }
            column(PrimaryOrder; PrimaryOrder)
            {
            }
            column(SecondaryOrder; SecondaryOrder)
            {
            }
            column(VersionDesc; VersionDesc)
            {
            }
            column(CapacityUtilization; CapacityUtilization)
            {
            }
            dataitem("Production BOM Line"; "Production BOM Line")
            {
                DataItemLink = "Production BOM No." = FIELD("Source No.");
                DataItemTableView = SORTING("Production BOM No.", "Version Code", "No.")
                                    WHERE("Version Code" = CONST('VER001'));
                column(No_ProductionBOMLine; "Production BOM Line"."No.")
                {
                }
                column(LineNo_ProductionBOMLine; "Production BOM Line"."Line No.")
                {
                }
                column(Description_ProductionBOMLine; "Production BOM Line".Description)
                {
                }
                column(UnitofMeasureCode_ProductionBOMLine; "Production BOM Line"."Unit of Measure Code")
                {
                }
                column(Quantity_ProductionBOMLine; "Production BOM Line".Quantity)
                {
                }
                column(StdQtyPer1; "Production BOM Line"."Quantity per" * 100)
                {
                }
                column(StdQty1; "Quantity per" * "Production Order".Quantity)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>03Sep2018
                    PrOrdCom.RESET;
                    PrOrdCom.SETRANGE(Status, "Production Order".Status);
                    PrOrdCom.SETRANGE("Prod. Order No.", "Production Order"."No.");
                    PrOrdCom.SETRANGE("Item No.", "Production BOM Line"."No.");
                    IF PrOrdCom.FINDFIRST THEN BEGIN
                        CurrReport.SKIP;
                    END;
                    //<<03Sep2018
                end;
            }
            dataitem("Prod. Order Component"; "Prod. Order Component")
            {
                DataItemLink = Status = FIELD(Status),
                               "Prod. Order No." = FIELD("No.");
                DataItemTableView = SORTING(Status, "Prod. Order No.", "Prod. Order Line No.", "Line No.")
                                    ORDER(Ascending);
                column(itmNo; "Item No.")
                {
                }
                column(LineNo; "Line No.")
                {
                }
                column(Description; Description)
                {
                }
                column(StdPer; StdQty * 100)
                {
                }
                column(StdQty; StdQty * "Production Order".Quantity)
                {
                }
                column(Qtyper; ActQtyPer)
                {
                }
                column(BlendQtyAdded; BlendQty)
                {
                }
                column(BlendQty; "Prod. Order Component"."Expected Quantity" - "Prod. Order Component"."Remaining Quantity")
                {
                }
                column(UOM; "Prod. Order Component"."Unit of Measure Code")
                {
                }
                column(ExpectedQty; ExpectedQty)
                {
                }
                column(ActualQty; ActualQty)
                {
                }
                column(FinishQty; ProductionOrderLine."Finished Quantity")
                {
                }
                column(PrimaryProdDesc; ProductionOrderLine.Description)
                {
                }
                column(PrimaryProdItmNo; ProductionOrderLine."Item No.")
                {
                }
                column(ConsumptionBlendOrderNo; "ConsumptionBlendOrderNo.")
                {
                }
                column(UsedQty; UsedQty)
                {
                }
                column(ProductionYield; ProductionYield)
                {
                }
                column(QCbatchno; QCbatchno)
                {
                }
                column(QCstatus; QCstatus)
                {
                }
                column(QCCertificateNo; QCCertificateNo)
                {
                }
                column(CapacityKet; CapacityKet)
                {
                }
                column(SpcGrvPrim; SpcGrvPrim)
                {
                }
                column(ComFooter1; ComFooter1)
                {
                }
                column(ComFooter2; ComFooter2)
                {
                }
                column(ComFooter3; ComFooter3)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>01Nov2018
                    TotalQtyPer := 0;
                    TotalExpQty := 0;
                    PrOrdCom.RESET;
                    PrOrdCom.SETRANGE(Status, "Production Order".Status);
                    PrOrdCom.SETRANGE("Prod. Order No.", "Production Order"."No.");
                    IF PrOrdCom.FINDFIRST THEN BEGIN
                        PrOrdCom.CALCSUMS("Quantity per", "Expected Quantity");
                        TotalQtyPer := PrOrdCom."Quantity per";
                        TotalExpQty := PrOrdCom."Expected Quantity";
                    END;

                    ActQtyPer := 0;
                    IF TotalQtyPer = 1 THEN BEGIN
                        ActQtyPer := "Prod. Order Component"."Quantity per" * 100;
                    END ELSE BEGIN
                        ActQtyPer := "Prod. Order Component"."Expected Quantity" * 100 / TotalExpQty;
                    END;
                    //<<01Nov2018

                    //>>04sep2018
                    BlendQty := 0;
                    IF Status = Status::Finished THEN
                        BlendQty := "Prod. Order Component"."Expected Quantity" - "Prod. Order Component"."Remaining Quantity"
                    ELSE
                        BlendQty := "Prod. Order Component"."Expected Quantity";
                    //<04Sep2018

                    //>>1

                    productionBOM.RESET;
                    productionBOM.SETRANGE(productionBOM."Production BOM No.", ProductionOrderLine."Production BOM No.");
                    productionBOM.SETRANGE(productionBOM."Version Code", ProductionOrderLine."Production BOM Version Code");
                    productionBOM.SETRANGE(productionBOM.Type, productionBOM.Type::Item);
                    productionBOM.SETRANGE(productionBOM."No.", "Prod. Order Component"."Item No.");
                    IF productionBOM.FINDFIRST THEN
                        TotalBOMvalue := TotalBOMvalue + productionBOM."Quantity per"
                    ELSE
                        TotalBOMvalue := 0;


                    QCCertificateNo := '';
                    ItemVersionParamterResult.RESET;
                    ItemVersionParamterResult.SETRANGE(ItemVersionParamterResult."Blend Order No", "Prod. Order Component"."Prod. Order No.");
                    IF ItemVersionParamterResult.FINDFIRST THEN BEGIN
                        QCCertificateNo := ItemVersionParamterResult."Certificate No.";
                    END;


                    CLEAR(OutputBlendOrderNo);
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Entry Type", recILE."Entry Type"::Consumption);
                    recILE.SETRANGE(recILE."Document No.", "Prod. Order Component"."Prod. Order No.");
                    recILE.SETRANGE(recILE."Prod. Order Comp. Line No.", "Prod. Order Component"."Line No.");
                    recILE.SETRANGE(recILE."Order No.", "Prod. Order Component"."Prod. Order No.");//28Apr2017
                    //recILE.SETRANGE(recILE."Prod. Order No.","Prod. Order Component"."Prod. Order No."); //nav2009
                    recILE.SETRANGE(recILE."Order Line No.", "Prod. Order Component"."Prod. Order Line No.");//28Apr2017
                    //recILE.SETRANGE(recILE."Prod. Order Line No.","Prod. Order Component"."Prod. Order Line No.");//nav2009
                    recILE.SETRANGE(recILE."Item No.", "Prod. Order Component"."Item No.");
                    recILE.SETRANGE(recILE."Item Tracking", recILE."Item Tracking"::"Lot No.");
                    IF recILE.FIND('-') THEN BEGIN
                        recILE1.RESET;
                        recILE1.SETRANGE(recILE1."Entry Type", recILE1."Entry Type"::Output);
                        recILE1.SETRANGE(recILE1."Lot No.", recILE."Lot No.");
                        IF recILE1.FINDFIRST THEN BEGIN
                            OutputBlendOrderNo := recILE1."Document No.";
                        END;
                    END;

                    // EBT MILAN (11092013)...IN Case of Semifinished Goods Report should not Show ConsumptionBlendOrderNo...........START
                    IF NOT ("Production Order"."Inventory Posting Group" = 'SEMIFG') THEN BEGIN
                        // EBT MILAN (11092013)...IN Case of Semifinished Goods Report should not Show ConsumptionBlendOrderNo...........END
                        CLEAR("ConsumptionBlendOrderNo.");
                        CLEAR(UsedQty);
                        recILE.RESET;
                        recILE.SETRANGE(recILE."Entry Type", recILE."Entry Type"::Output);
                        recILE.SETRANGE(recILE."Document No.", "Production Order"."No.");
                        recILE.SETFILTER(Quantity, '>%1', 0); //CAS-13955-H5T3T6
                        IF recILE.FINDFIRST THEN
                            REPEAT  //RSPL008 Multi Output Etries
                                recILE1.RESET;
                                recILE1.SETRANGE(recILE1."Entry Type", recILE1."Entry Type"::Consumption);
                                recILE1.SETRANGE(recILE1."Lot No.", recILE."Lot No.");
                                IF recILE1.FINDFIRST THEN
                                    REPEAT
                                        //UsedQty := UsedQty + recILE1.Quantity;
                                        recProdOrderLine.RESET;
                                        //     recProdOrderLine.SETRANGE(recProdOrderLine."Inventory Posting Group",'BULK'); // EBT MILAN 170214 To show semifinnished goods
                                        //  production order no at bottom
                                        recProdOrderLine.SETFILTER(recProdOrderLine."Inventory Posting Group", '%1|%2', 'BULK', 'SEMIFG');
                                        recProdOrderLine.SETRANGE(recProdOrderLine."Prod. Order No.", recILE1."Document No.");
                                        IF recProdOrderLine.FINDFIRST THEN BEGIN
                                            "ConsumptionBlendOrderNo." := "ConsumptionBlendOrderNo." + recILE1."Document No." + ', ';
                                            UsedQty := UsedQty + recILE1.Quantity;
                                        END;
                                    UNTIL recILE1.NEXT = 0;
                            UNTIL recILE.NEXT = 0;  //RSPL008 Multi Output Etries
                    END;

                    //<<1

                    //>>2
                    //Prod. Order Component, Body (2) - OnPreSection()
                    IF ItemDescription = TRUE THEN BEGIN
                        Description := "Prod. Order Component".Description;
                    END ELSE
                        Description := OutputBlendOrderNo;

                    ExpectedQty := ExpectedQty + (productionBOM."Quantity per" * "Production Order".Quantity);
                    ActualQty := ActualQty + "Prod. Order Component"."Expected Quantity" - "Prod. Order Component"."Remaining Quantity";
                    //<<2

                    //>>3
                    //Prod. Order Component, Footer (3) - OnPreSection()
                    IF TotalBOMvalue <> 0 THEN
                        BOMYield := 1 / TotalBOMvalue;

                    IF ActualQty <> 0 THEN //17Aug2018
                        ProductionYield := (ProductionOrderLine."Finished Quantity" / ActualQty) * 100;
                    //<<3

                    //>>29Apr2017 ComFooter1

                    CLEAR(ComFooter1);
                    IF NOT "Production Order"."QC Tested"
                      AND ("Production Order"."Order Type" = "Production Order"."Order Type"::Primary) THEN
                        ComFooter1 := TRUE;
                    //MESSAGE('Com1 %1',ComFooter1);

                    //<<29Apr2017
                    //>>4
                    //Prod. Order Component, Footer (5) - OnPreSection()
                    //CurrReport.SHOWOUTPUT(("Production Order"."QC Tested"=FALSE)
                    //AND ("Production Order"."Order Type" = "Production Order"."Order Type" :: Primary));

                    //<<4

                    //>>29Apr2017 ComFooter2

                    CLEAR(ComFooter2);
                    IF "Production Order"."QC Tested"
                      AND (ProductionOrderLine."Routing No." = '')
                      AND ("Production Order"."Order Type" = "Production Order"."Order Type"::Primary) THEN
                        ComFooter2 := TRUE;
                    //MESSAGE('Com2 %1',ComFooter2);

                    //<<29Apr2017
                    //>>5
                    //Prod. Order Component, Footer (6) - OnPreSection()
                    /*
                    IF CurrReport.SHOWOUTPUT(("Production Order"."QC Tested"=TRUE) AND (ProductionOrderLine."Routing No." = '')
                    AND ("Production Order"."Order Type" = "Production Order"."Order Type" :: Primary)) THEN
                    BEGIN
                     ILE.RESET;
                     ILE.SETRANGE(ILE."Entry Type",ILE."Entry Type"::Output);
                     ILE.SETRANGE(ILE."Document No.",ProductionOrderLine."Prod. Order No.");
                     ILE.FINDFIRST;
                    END;
                    */
                    //<<5


                    //>>29Apr2017 ComFooter2

                    CLEAR(ComFooter3);
                    IF "Production Order"."QC Tested"
                      AND (ProductionOrderLine."Routing No." <> '')
                      AND ("Production Order"."Order Type" = "Production Order"."Order Type"::Primary) THEN
                        ComFooter3 := TRUE;
                    //MESSAGE('Com3 %1',ComFooter3);

                    //<<29Apr2017
                    //>>6
                    /*
                    Prod. Order Component, Footer (7) - OnPreSection()
                    IF CurrReport.SHOWOUTPUT(("Production Order"."QC Tested"=TRUE) AND (ProductionOrderLine."Routing No." <> '')
                    AND ("Production Order"."Order Type" = "Production Order"."Order Type" :: Primary)) THEN
                    BEGIN
                     ILE.RESET;
                     ILE.SETRANGE(ILE."Entry Type",ILE."Entry Type"::Output);
                     ILE.SETRANGE(ILE."Document No.",ProductionOrderLine."Prod. Order No.");
                     ILE.FINDFIRST;
                    END;
                    */
                    //<<6

                    //>>7
                    /*
                    Prod. Order Component, Footer (8) - OnPreSection()
                    CurrReport.SHOWOUTPUT(("Production Order"."Order Type" = "Production Order"."Order Type" :: Primary));
                    */
                    //<<7

                    //>>01Sep2018
                    StdQty := 0;
                    PrBOMLine.RESET;
                    PrBOMLine.SETRANGE("Production BOM No.", "Production Order"."Source No.");
                    PrBOMLine.SETRANGE(Type, PrBOMLine.Type::Item);
                    PrBOMLine.SETRANGE("No.", "Prod. Order Component"."Item No.");
                    PrBOMLine.SETRANGE("Version Code", 'VER001');
                    IF PrBOMLine.FINDFIRST THEN BEGIN
                        StdQty := PrBOMLine."Quantity per";
                    END;


                    //<<01Sep2018

                end;
            }
            dataitem(DataItem1000000002; 50016)
            {
                DataItemLink = "Blend Order No." = FIELD("No.");
                DataItemTableView = SORTING("Certificate No.", "Item No.", "Blend Order No.")
                                    ORDER(Ascending);
                dataitem("QC Certifcate Details Second"; 50016)
                {
                    DataItemLink = "Certificate No." = FIELD("Certificate No.");
                    DataItemTableView = SORTING("Certificate No.", "Item No.", "Blend Order No.")
                                        ORDER(Ascending);
                    dataitem("Production Order Second"; "Production Order")
                    {
                        DataItemLink = "No." = FIELD("Blend Order No.");
                        DataItemTableView = SORTING("No.", Status)
                                            ORDER(Ascending)
                                            WHERE(Status = FILTER(Finished | Released),
                                                  "Order Type" = FILTER(Secondary));
                        dataitem("Prod. Order Line"; "Prod. Order Line")
                        {
                            DataItemLink = Status = FIELD(Status),
                                           "Prod. Order No." = FIELD("No.");
                            DataItemTableView = SORTING(Status, "Prod. Order No.", "Line No.")
                                                ORDER(Ascending);
                            column(ProdItmNo; "Item No.")
                            {
                            }
                            column(ProdDesc; "Prod. Order Line".Description)
                            {
                            }
                            column(QtyPerUom; QtyPerUom)
                            {
                            }
                            column(SOMQty; SOMQty)
                            {
                            }
                            column(SpecGravity; SpecGravity)
                            {
                            }
                            column(ConsQty; ABS(ConsQty))
                            {
                            }
                            column(prodQtyLtr; "Prod. Order Line"."Finished Quantity" * ItemUOM."Qty. per Unit of Measure")
                            {
                            }
                            column(ProdBatchNo; "QC Certifcate Details Second"."Batch No.")
                            {
                            }
                            column(ProdOrderNo; "QC Certifcate Details Second"."Blend Order No.")
                            {
                            }

                            trigger OnAfterGetRecord()
                            begin

                                //MESSAGE("Prod. Order Line"."Prod. Order No.");

                                //>>1

                                /*
                                CLEAR("ConsumptionBlendOrderNo.");
                                CLEAR(UsedQty);
                                recILE.RESET;
                                recILE.SETRANGE(recILE."Entry Type",recILE."Entry Type"::Output);
                                recILE.SETRANGE(recILE."Document No.","Production Order"."No.");
                                IF recILE.FINDFIRST THEN
                                BEGIN
                                 recILE1.RESET;
                                 recILE1.SETRANGE(recILE1."Entry Type",recILE1."Entry Type"::Consumption);
                                 recILE1.SETRANGE(recILE1."Lot No.",recILE."Lot No.");
                                 IF recILE1.FINDFIRST THEN
                                 REPEAT
                                   recProdOrderLine.RESET;
                                   recProdOrderLine.SETRANGE(recProdOrderLine."Inventory Posting Group",'BULK');
                                   recProdOrderLine.SETRANGE(recProdOrderLine."Prod. Order No.",recILE1."Document No.");
                                   IF recProdOrderLine.FINDFIRST THEN
                                   BEGIN
                                      "ConsumptionBlendOrderNo." := "ConsumptionBlendOrderNo." + recILE1."Document No."+', ';
                                      UsedQty := UsedQty + recILE1.Quantity;
                                   END;
                                 UNTIL recILE1.NEXT = 0;
                                END;
                                */

                                ILE.RESET;
                                ILE.SETRANGE(ILE."Entry Type", ILE."Entry Type"::Output);
                                ILE.SETRANGE(ILE."Document No.", "Prod. Order Line"."Prod. Order No.");
                                ILE.SETRANGE(ILE."Item No.", "Prod. Order Line"."Item No.");
                                IF ILE.FINDFIRST THEN
                                    SpecGravity := ILE."Density Factor";

                                ConsQty := 0;
                                ILE.RESET;
                                ILE.SETRANGE(ILE."Entry Type", ILE."Entry Type"::Consumption);
                                ILE.SETRANGE(ILE."Document No.", "Prod. Order Line"."Prod. Order No.");
                                ILE.SETRANGE(ILE."Item No.", ProductionOrderLine."Item No.");
                                IF ILE.FINDFIRST THEN
                                    REPEAT
                                        ConsQty := ConsQty + ILE.Quantity;
                                    UNTIL ILE.NEXT = 0;

                                TotConsQty += ConsQty;
                                //<<1

                                //>>2
                                //Prod. Order Line, Body (1) - OnPreSection()

                                //IF CurrReport.SHOWOUTPUT(("Production Order"."Order Type" = "Production Order"."Order Type" :: Primary)) THEN
                                //BEGIN
                                IF PrimaryOrder THEN //29Apr2017
                                BEGIN
                                    ILE.RESET;
                                    ILE.SETRANGE(ILE."Entry Type", ILE."Entry Type"::Output);
                                    ILE.SETRANGE(ILE."Document No.", "Prod. Order Line"."Prod. Order No.");
                                    ILE.FINDFIRST;
                                    ItemUOM.GET("Prod. Order Line"."Item No.", "Prod. Order Line"."Unit of Measure Code");
                                    TotalQty := TotalQty + ("Prod. Order Line".Quantity * ItemUOM."Qty. per Unit of Measure");
                                    QtyPerUom := ILE."Qty. per Unit of Measure";
                                    SOMQty := ILE.Quantity / ILE."Qty. per Unit of Measure";
                                END;
                                //END
                                //ELSE
                                //CurrReport.SHOWOUTPUT(FALSE);
                                //<<2

                            end;
                        }
                    }

                    trigger OnAfterGetRecord()
                    begin

                        //>>1
                        //QC Certifcate Details Second, Footer (2) - OnPreSection()
                        IF CurrReport.SHOWOUTPUT(("Production Order"."Order Type" = "Production Order"."Order Type"::Primary)) THEN BEGIN
                            DiffQTy := TotalQty - TotConsQty;
                        END
                        ELSE
                            CurrReport.SHOWOUTPUT(FALSE);
                        //<<1

                        //>>2
                        /*
                        QC Certifcate Details Second, Footer (3) - OnPreSection()
                        CurrReport.SHOWOUTPUT(("Production Order"."Order Type" = "Production Order"."Order Type" :: Primary));
                        */
                        //<<2

                        //>>3
                        /*
                        QC Certifcate Details, Footer (2) - OnPreSection()
                        CurrReport.SHOWOUTPUT(("Production Order"."Order Type" = "Production Order"."Order Type" :: Secondary));
                        */
                        //<<3

                    end;
                }
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                ProductionOrderLine.RESET;
                ProductionOrderLine.SETRANGE(ProductionOrderLine.Status, "Production Order".Status);
                ProductionOrderLine.SETRANGE(ProductionOrderLine."Prod. Order No.", "Production Order"."No.");
                IF ProductionOrderLine.FINDFIRST THEN BEGIN
                    RoutingLine.RESET;
                    RoutingLine.SETRANGE(RoutingLine."Routing No.", ProductionOrderLine."Routing No.");
                    RoutingLine.SETRANGE(RoutingLine.Type, RoutingLine.Type::"Machine Center");
                    IF RoutingLine.FINDFIRST THEN BEGIN
                        RoutingDscr := RoutingLine.Description;
                        IF MachinCentr.GET(RoutingLine."No.") THEN
                            CapacityKet := MachinCentr.Capacity;
                    END;
                END;

                /*
                ILE.RESET;
                ILE.SETRANGE(ILE."Entry Type",ILE."Entry Type"::Output);
                ILE.SETRANGE(ILE."Document No.","Production Order"."No.");
                IF ILE.FINDFIRST THEN
                   SpcGrvPrim := ILE."Density Factor";
                *///Code Commented 27may2019

                IF SpcGrvPrim = 0 THEN
                    SpcGrvPrim := 1;

                ProductionOrderLine.RESET;
                ProductionOrderLine.SETRANGE(ProductionOrderLine.Status, "Production Order".Status);
                ProductionOrderLine.SETRANGE(ProductionOrderLine."Prod. Order No.", "Production Order"."No.");
                ProductionOrderLine.FINDFIRST;

                QCbatchno := '';
                recQCcertificate.RESET;
                recQCcertificate.SETRANGE(recQCcertificate."Blend Order No.", "Production Order"."No.");
                IF recQCcertificate.FINDFIRST THEN BEGIN
                    QCbatchno := recQCcertificate."Batch No.";
                END
                ELSE
                    QCbatchno := "Production Order"."Description 2";
                //<<1

                //>>2
                //Production Order, Header (1) - OnPreSection()
                recLocation.GET("Production Order"."Location Code");
                //<<2

                //>>3
                //Production Order, Body (2) - OnPreSection()
                IF "Production Order"."QC Tested" = TRUE THEN
                    QCstatus := 'Approved'
                ELSE
                    QCstatus := 'Pending For Approval';
                //<<3


                //>>29Apr2017

                CLEAR(PrimaryOrder);
                IF ("Production Order"."Order Type" = "Production Order"."Order Type"::Primary) THEN
                    PrimaryOrder := TRUE;

                CLEAR(SecondaryOrder);
                IF ("Production Order"."Order Type" = "Production Order"."Order Type"::Secondary) THEN
                    SecondaryOrder := TRUE;

                //<<29Apr2017

                //>>05Sep2018
                VersionDesc := '';
                PrOrdLine01.RESET;
                PrOrdLine01.SETRANGE(Status, Status);
                PrOrdLine01.SETRANGE("Prod. Order No.", "No.");
                IF PrOrdLine01.FINDFIRST THEN BEGIN
                    PrBOMVer01.RESET;
                    IF PrBOMVer01.GET(PrOrdLine01."Production BOM No.", PrOrdLine01."Production BOM Version Code") THEN BEGIN
                        VersionDesc := PrBOMVer01."Version Code" + ' - ' + PrBOMVer01.Description;
                    END;
                END;
                //<<05Sep2018

                //RSPLSUM 13Mar2020>>
                CLEAR(TotBlendQty);
                RecProdOrdComp.RESET;
                RecProdOrdComp.SETRANGE(Status, "Production Order".Status);
                RecProdOrdComp.SETRANGE("Prod. Order No.", "Production Order"."No.");
                IF RecProdOrdComp.FINDSET THEN
                    REPEAT
                        IF "Production Order".Status = "Production Order".Status::Finished THEN
                            TotBlendQty += RecProdOrdComp."Expected Quantity" - RecProdOrdComp."Remaining Quantity"
                        ELSE
                            TotBlendQty += RecProdOrdComp."Expected Quantity";
                    UNTIL RecProdOrdComp.NEXT = 0;

                IF CapacityKet <> 0 THEN
                    CapacityUtilization := (TotBlendQty / SpcGrvPrim) / CapacityKet * 100
                ELSE
                    CapacityUtilization := 0;
                //RSPLSUM 13Mar2020<<

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                TotConsQty := 0;
                TotalQty := 0;
                ActualQty := 0;
                ExpectedQty := 0;
                TotalBOMvalue := 0;
                BOMYield := 0;
                ProductionYield := 0;
                //<<1
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Reqd. Item Description"; ItemDescription)
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

    var
        Routing: Record 5409;
        RoutingDscr: Text[100];
        productionBOM: Record 99000772;
        ProductionOrderLine: Record 5406;
        ExpectedQty: Decimal;
        ActualQty: Decimal;
        TotalBOMvalue: Decimal;
        BOMYield: Decimal;
        ProductionYield: Decimal;
        ILE: Record 32;
        TotalQty: Decimal;
        QcCert: Code[20];
        ItemUOM: Record 5404;
        CapacityKet: Decimal;
        RoutingLine: Record 99000764;
        MachinCentr: Record 99000758;
        SpecGravity: Decimal;
        ConsQty: Decimal;
        TotConsQty: Decimal;
        DiffQTy: Decimal;
        QtyPerUom: Decimal;
        SOMQty: Decimal;
        SpcGrvPrim: Decimal;
        recLocation: Record 14;
        ItemDescription: Boolean;
        Description: Text[100];
        DescriptionText: Text[30];
        ItemVersionParamterResult: Record 50015;
        QCstatus: Text[30];
        QCCertificateNo: Code[20];
        recQCcertificate: Record 50016;
        QCbatchno: Code[10];
        recProdOrderLine: Record 5406;
        OutputBlendOrderNo: Code[20];
        "ConsumptionBlendOrderNo.": Text[250];
        recILE: Record 32;
        recILE1: Record 32;
        recProdOrderComponent: Record 5407;
        RemQty: Decimal;
        UsedQty: Decimal;
        "-----29Apr2017": Integer;
        ComFooter1: Boolean;
        ComFooter2: Boolean;
        ComFooter3: Boolean;
        PrimaryOrder: Boolean;
        SecondaryOrder: Boolean;
        "--------01Sep2018": Integer;
        PrBOMLine: Record 99000772;
        StdQty: Decimal;
        PrOrdCom: Record 5407;
        BlendQty: Decimal;
        PrBOMVer01: Record 99000779;
        PrOrdLine01: Record 5406;
        VersionDesc: Text;
        "------01Nov2018": Integer;
        ActQtyPer: Decimal;
        TotalQtyPer: Decimal;
        TotalExpQty: Decimal;
        TotBlendQty: Decimal;
        RecProdOrdComp: Record 5407;
        CapacityUtilization: Decimal;
}


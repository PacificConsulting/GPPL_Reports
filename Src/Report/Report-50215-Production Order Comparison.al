report 50215 "Production Order Comparison"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 03Sep2018   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ProductionOrderComparison.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Production Order"; "Production Order")
        {
            DataItemTableView = SORTING(Status, "No.")
                                WHERE(Status = CONST(Finished));
            column(Description_ProductionOrder; "Production Order".Description)
            {
            }
            column(Description2_ProductionOrder; "Production Order"."Description 2")
            {
            }
            column(SourceNo_ProductionOrder; "Production Order"."Source No.")
            {
            }
            column(InventoryPostingGroup_ProductionOrder; "Production Order"."Inventory Posting Group")
            {
            }
            column(StartingTime_ProductionOrder; "Production Order"."Starting Time")
            {
            }
            column(StartingDate_ProductionOrder; "Production Order"."Starting Date")
            {
            }
            column(EndingTime_ProductionOrder; "Production Order"."Ending Time")
            {
            }
            column(EndingDate_ProductionOrder; "Production Order"."Ending Date")
            {
            }
            column(DueDate_ProductionOrder; "Production Order"."Due Date")
            {
            }
            column(FinishedDate_ProductionOrder; "Production Order"."Finished Date")
            {
            }
            column(No_ProductionOrder; "Production Order"."No.")
            {
            }
            column(Quantity_ProductionOrder; "Production Order".Quantity)
            {
            }
            column(VersionDesc; VersionDesc)
            {
            }
            dataitem("Production BOM Line"; "Production BOM Line")
            {
                DataItemLink = "Production BOM No." = FIELD("Source No.");
                DataItemTableView = SORTING("Production BOM No.", "Version Code", "No.")
                                    WHERE("Version Code" = CONST('VER001'));
                column(DirectCostAmount_ProdOrderComponent; "Prod. Order Component"."Direct Cost Amount")
                {
                }
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

                    /*
                    //>>07Sep2018
                    PrOrdCom.RESET;
                    PrOrdCom.SETRANGE(Status,"Production Order".Status);
                    PrOrdCom.SETRANGE("Prod. Order No.","Production Order"."No.");
                    PrOrdCom.SETRANGE("Item No.","Production BOM Line"."No.");
                    IF PrOrdCom.FINDFIRST  THEN
                    BEGIN
                      CurrReport.SKIP;
                    END;
                    //<<07Sep2018
                    */

                    //>>28Aug2019
                    IF PrintExcel THEN
                        ProdBOMData;
                    //<<28Aug2019

                end;
            }
            dataitem(ProductionBOMActual; "Production BOM Line")
            {
                DataItemLink = "Production BOM No." = FIELD("Source No.");
                DataItemTableView = SORTING("Production BOM No.", "Version Code", "No.");

                trigger OnAfterGetRecord()
                begin
                    //IF ActualBOM = 'VER001' THEN
                    // CurrReport.SKIP;

                    //>>24Sep2019
                    IF PrintExcel THEN
                        ProdBOMDataActual;
                    //<<24Sep2019
                end;

                trigger OnPreDataItem()
                begin
                    SETRANGE("Version Code", ActualBOM);
                end;
            }
            dataitem("Prod. Order Component"; "Prod. Order Component")
            {
                DataItemLink = Status = FIELD(Status),
                               "Prod. Order No." = FIELD("No.");
                DataItemTableView = SORTING(Status, "Prod. Order No.", "Prod. Order Line No.", "Line No.");
                column(Status_ProdOrderComponent; "Prod. Order Component".Status)
                {
                }
                column(ProdOrderNo_ProdOrderComponent; "Prod. Order Component"."Prod. Order No.")
                {
                }
                column(ProdOrderLineNo_ProdOrderComponent; "Prod. Order Component"."Prod. Order Line No.")
                {
                }
                column(LineNo_ProdOrderComponent; "Prod. Order Component"."Line No.")
                {
                }
                column(ItemNo_ProdOrderComponent; "Prod. Order Component"."Item No.")
                {
                }
                column(Description_ProdOrderComponent; "Prod. Order Component".Description)
                {
                }
                column(UnitofMeasureCode_ProdOrderComponent; "Prod. Order Component"."Unit of Measure Code")
                {
                }
                column(Quantity_ProdOrderComponent; "Prod. Order Component".Quantity)
                {
                }
                column(RoutingLinkCode_ProdOrderComponent; "Prod. Order Component"."Routing Link Code")
                {
                }
                column(ExpectedQuantity_ProdOrderComponent; "Prod. Order Component"."Expected Quantity")
                {
                }
                column(RemainingQuantity_ProdOrderComponent; "Prod. Order Component"."Remaining Quantity")
                {
                }
                column(LocationCode_ProdOrderComponent; "Prod. Order Component"."Location Code")
                {
                }
                column(ShortcutDimension1Code_ProdOrderComponent; "Prod. Order Component"."Shortcut Dimension 1 Code")
                {
                }
                column(ShortcutDimension2Code_ProdOrderComponent; "Prod. Order Component"."Shortcut Dimension 2 Code")
                {
                }
                column(BinCode_ProdOrderComponent; "Prod. Order Component"."Bin Code")
                {
                }
                column(CalculationFormula_ProdOrderComponent; "Prod. Order Component"."Calculation Formula")
                {
                }
                column(Quantityper_ProdOrderComponent; "Prod. Order Component"."Quantity per")
                {
                }
                column(UnitCost_ProdOrderComponent; "Prod. Order Component"."Unit Cost")
                {
                }
                column(CostAmount_ProdOrderComponent; "Prod. Order Component"."Cost Amount")
                {
                }
                column(DueDate_ProdOrderComponent; "Prod. Order Component"."Due Date")
                {
                }
                column(DueTime_ProdOrderComponent; "Prod. Order Component"."Due Time")
                {
                }
                column(QtyperUnitofMeasure_ProdOrderComponent; "Prod. Order Component"."Qty. per Unit of Measure")
                {
                }
                column(RemainingQtyBase_ProdOrderComponent; "Prod. Order Component"."Remaining Qty. (Base)")
                {
                }
                column(QuantityBase_ProdOrderComponent; "Prod. Order Component"."Quantity (Base)")
                {
                }
                column(ReservedQtyBase_ProdOrderComponent; "Prod. Order Component"."Reserved Qty. (Base)")
                {
                }
                column(ReservedQuantity_ProdOrderComponent; "Prod. Order Component"."Reserved Quantity")
                {
                }
                column(ExpectedQtyBase_ProdOrderComponent; "Prod. Order Component"."Expected Qty. (Base)")
                {
                }
                column(OriginalItemNo_ProdOrderComponent; "Prod. Order Component"."Original Item No.")
                {
                }
                column(OriginalVariantCode_ProdOrderComponent; "Prod. Order Component"."Original Variant Code")
                {
                }
                column(QtyToConsume_ProdOrderComponent; "Prod. Order Component"."Qty. To Consume")
                {
                }
                column(InputQuantity_ProdOrderComponent; "Prod. Order Component"."Input Quantity")
                {
                }
                column(OnlineRejectionQty_ProdOrderComponent; "Prod. Order Component"."Online Rejection Qty")
                {
                }
                column(DirectUnitCost_ProdOrderComponent; "Prod. Order Component"."Direct Unit Cost")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>24Sep2019
                    CLEAR(UnitCostVE);
                    CLEAR(TotalCostVE);
                    VE.RESET;
                    VE.SETCURRENTKEY("Document No.", "Item No.");
                    VE.SETRANGE("Document No.", "Prod. Order Component"."Prod. Order No.");
                    VE.SETRANGE("Item No.", "Item No.");
                    IF VE.FINDFIRST THEN BEGIN
                        VE.CALCSUMS("Cost per Unit", "Cost Amount (Actual)", "Cost Amount (Expected)");
                        UnitCostVE := ABS(VE."Cost per Unit");
                        //TotalCostVE := ABS(VE."Cost Amount (Actual)")+ABS(VE."Cost Amount (Expected)");
                    END;
                    //<<24Sep2019

                    BulkItem := FALSE;
                    RecILENew.RESET;
                    RecILENew.SETRANGE("Document No.", "Prod. Order Component"."Prod. Order No.");
                    RecILENew.SETRANGE("Item No.", "Item No.");
                    IF RecILENew.FINDFIRST THEN
                        IF RecILENew."Item Category Code" = 'CAT10' THEN
                            BulkItem := TRUE;

                    //RSPLSUM 17Dec19>>
                    RecILE.RESET;
                    RecILE.SETRANGE("Document No.", "Prod. Order Component"."Prod. Order No.");
                    RecILE.SETRANGE("Entry Type", RecILE."Entry Type"::Consumption);
                    RecILE.SETRANGE("Item No.", "Item No.");
                    IF NOT BulkItem THEN
                        RecILE.SETRANGE(Quantity, -"Prod. Order Component"."Expected Quantity");
                    IF RecILE.FINDFIRST THEN
                        REPEAT
                            RecILE.CALCFIELDS("Cost Amount (Actual)", "Cost Amount (Expected)");
                            TotalCostVE += ABS(RecILE."Cost Amount (Actual)") + ABS(RecILE."Cost Amount (Expected)");
                        UNTIL RecILE.NEXT = 0;

                    //RSPLSUM 17Dec19<<

                    //>>28Aug2019
                    IF PrintExcel THEN
                        ProdCOMData;
                    //<<28Aug2019
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>07Sep2018
                ActualBOM := '';
                VersionDesc := '';
                PrOrdLine01.RESET;
                PrOrdLine01.SETRANGE(Status, Status);
                PrOrdLine01.SETRANGE("Prod. Order No.", "No.");
                IF PrOrdLine01.FINDFIRST THEN BEGIN
                    ActualBOM := PrOrdLine01."Production BOM Version Code";
                    PrBOMVer01.RESET;
                    IF PrBOMVer01.GET(PrOrdLine01."Production BOM No.", PrOrdLine01."Production BOM Version Code") THEN BEGIN
                        VersionDesc := PrBOMVer01."Version Code" + ' - ' + PrBOMVer01.Description;
                    END;
                END;
                //<<07Sep2018

                //>>28Aug2019
                IF PrintExcel THEN
                    ProdBodyData;
                //<<28Aug2019
            end;

            trigger OnPreDataItem()
            begin
                SETRANGE("Starting Date", SDate, EDate);

                IF LocFilter <> '' THEN
                    SETFILTER("Location Code", LocFilter);

                IF ItmNo <> '' THEN
                    SETFILTER("Source No.", ItmNo);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; SDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin

                        EDate := SDate;
                    end;
                }
                field("End Date"; EDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin

                        IF EDate < SDate THEN
                            ERROR('EndDate must be greater than StartDate');
                    end;
                }
                field("Location Filter"; LocFilter)
                {
                    ApplicationArea = all;
                    //TableRelation = Location;
                }
                field("Item Filter"; ItmNo)
                {
                    ApplicationArea = all;
                    // TableRelation = Item;
                }
                field("Print To Excel"; PrintExcel)
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

        //>>28Aug2019
        IF PrintExcel THEN
            CreateExcelbook;
        //<<28Aug2019
    end;

    trigger OnPreReport()
    begin

        IF SDate = 0D THEN
            ERROR('Please Specify Start Date');

        IF EDate = 0D THEN
            ERROR('Please Specify End Date');

        //>>28Aug2019
        IF PrintExcel THEN
            MakeExcelInfo;
        //<<28Aug2019
    end;

    var
        SDate: Date;
        EDate: Date;
        LocFilter: Code[100];
        ItmNo: Code[100];
        VersionDesc: Text;
        PrOrdCom: Record 5407;
        PrOrdLine01: Record 5406;
        PrBOMVer01: Record 99000779;
        ExBuf: Record 370 temporary;
        PrintExcel: Boolean;
        VE: Record 5802;
        UnitCostVE: Decimal;
        TotalCostVE: Decimal;
        ActualBOM: Code[20];
        RecILE: Record 32;
        RecILENew: Record 32;
        BulkItem: Boolean;

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        ExBuf.SetUseInfoSheet;
        ExBuf.AddInfoColumn('Company Name', FALSE, TRUE, FALSE, FALSE, '', 1);
        ExBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExBuf.NewRow;
        ExBuf.AddInfoColumn('Report Name', FALSE, TRUE, FALSE, FALSE, '', 1);
        ExBuf.AddInfoColumn('Production Order Comparision', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExBuf.NewRow;
        ExBuf.AddInfoColumn('Report No.', FALSE, TRUE, FALSE, FALSE, '', 1);
        ExBuf.AddInfoColumn(50215, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExBuf.NewRow;
        ExBuf.AddInfoColumn('User ID', FALSE, TRUE, FALSE, FALSE, '', 1);
        ExBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExBuf.NewRow;
        ExBuf.AddInfoColumn('Date', FALSE, TRUE, FALSE, FALSE, '', 1);
        ExBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExBuf.NewRow;
        ExBuf.ClearNewRow;

        ProdHeaderData;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //>>
        /* ExBuf.CreateBook('', 'Production Order Comparision');
        ExBuf.CreateBookAndOpenExcel('', 'Production Order Comparision', '', '', USERID);
        ExBuf.GiveUserControl; */

        ExBuf.CreateNewBook('Production Order Comparision');
        ExBuf.WriteSheet('Production Order Comparision', CompanyName, UserId);
        ExBuf.CloseBook();
        ExBuf.SetFriendlyFilename(StrSubstNo('Production Order Comparision', CurrentDateTime, UserId));
        ExBuf.OpenExcel();
        //<<
    end;

    local procedure ProdHeaderData()
    begin
        ExBuf.NewRow;
        ExBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Version Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Order Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);

        ExBuf.AddColumn('Standard', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Qty Per', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);

        ExBuf.AddColumn('Selected Version', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Qty Per', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);

        ExBuf.AddColumn('Actual', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Qty Per', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Unit Cost', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Total Cost', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExBuf.AddColumn('Rejection Qty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
    end;

    local procedure ProdBodyData()
    begin

        ExBuf.NewRow;
        ExBuf.AddColumn("Production Order"."No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Production Order"."Starting Date", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Production Order"."Source No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Production Order".Description, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn(VersionDesc, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Production Order"."Description 2", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Production Order".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);
        ExBuf.AddColumn("Production Order"."Order Type", FALSE, '', FALSE, FALSE, FALSE, '', 1);
    end;

    local procedure ProdBOMData()
    begin

        ExBuf.NewRow;
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);

        ExBuf.AddColumn("Production BOM Line"."Version Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Production BOM Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Production BOM Line".Description, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Production BOM Line"."Quantity per" * 100, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00000', 0);
        ExBuf.AddColumn("Production BOM Line"."Quantity per" * "Production Order".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);
    end;

    local procedure ProdBOMDataActual()
    begin

        ExBuf.NewRow;
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);

        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);

        ExBuf.AddColumn(ProductionBOMActual."Version Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn(ProductionBOMActual."No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn(ProductionBOMActual.Description, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn(ProductionBOMActual."Quantity per" * 100, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00000', 0);
        ExBuf.AddColumn(ProductionBOMActual."Quantity per" * "Production Order".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);
    end;

    local procedure ProdCOMData()
    begin

        ExBuf.NewRow;
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);

        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);

        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);

        ExBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Prod. Order Component"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Prod. Order Component".Description, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExBuf.AddColumn("Prod. Order Component"."Quantity per" * 100, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00000', 0);
        ExBuf.AddColumn("Prod. Order Component"."Expected Quantity", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);
        //ExBuf.AddColumn("Prod. Order Component"."Unit Cost",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);
        //ExBuf.AddColumn("Prod. Order Component"."Cost Amount",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);
        ExBuf.AddColumn(UnitCostVE, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24Sep2019
        ExBuf.AddColumn(TotalCostVE, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24Sep2019
        ExBuf.AddColumn("Prod. Order Component"."Online Rejection Qty", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);
    end;
}


report 50115 "Yeild Report"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/YeildReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Production Order"; "Production Order")
        {
            DataItemTableView = SORTING(Status, "No.")
                                WHERE(Status = CONST(Finished),
                                      "Inventory Posting Group" = FILTER('BULK' | 'SEMIFG'));
            RequestFilterFields = "Date Filter", "No.", "Location Code", "Inventory Posting Group";
            column(CompName; recCompInfo.Name)
            {
            }
            column(Dateeee; FORMAT(TODAY, 0, 4))
            {
            }
            column(DateCriteria; DateCriteria)
            {
            }
            column(LocFilter; 'Location :- ' + reclocation.Name)
            {
            }
            dataitem("Prod. Order Line"; "Prod. Order Line")
            {
                DataItemLink = Status = FIELD(Status),
                               "Prod. Order No." = FIELD("No.");
                DataItemTableView = SORTING(Status, "Prod. Order No.", "Line No.");
                dataitem("Prod. Order Component"; "Prod. Order Component")
                {
                    DataItemLink = Status = FIELD(Status),
                                   "Prod. Order No." = FIELD("Prod. Order No."),
                                   "Prod. Order Line No." = FIELD("Line No.");
                    DataItemTableView = SORTING(Status, "Prod. Order No.", "Prod. Order Line No.", "Line No.");
                    column(BlendNo; "Prod. Order Component"."Prod. Order No.")
                    {
                    }
                    column(Date; "Production Order"."Last Date Modified")
                    {
                    }
                    column(BatchNo; "Production Order"."Description 2")
                    {
                    }
                    column(FGItemName; recItem.Description)
                    {
                    }
                    column(ConQTY; ConQTY)
                    {
                    }
                    column(decQty; decQty)
                    {
                    }
                    column(Yeild; ROUND(decYeild, 0.01))
                    {
                    }
                    column(ProdLoss; decQty - ConQTY)
                    {
                    }
                    column(vCapacity; vCapacity)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin
                        //>>1

                        //Prod. Order Component, Body (1) - OnPreSection()
                        //consumed quantity
                        "Prod. Order Component".SETCURRENTKEY(Status, "Prod. Order No.", "Prod. Order Line No.", "Line No.");
                        "Prod. Order Component".CALCFIELDS("Prod. Order Component"."Act. Consumption (Qty)");

                        IF recItem.GET("Prod. Order Component"."Item No.") THEN;
                        recItemUOM.RESET;
                        recItemUOM.SETRANGE(recItemUOM."Item No.", recItem."No.");
                        recItemUOM.SETRANGE(recItemUOM.Code, recItem."Sales Unit of Measure");
                        IF recItemUOM.FINDFIRST THEN
                            IF (recItem."Sales Unit of Measure" <> '') AND (recItem."Sales Unit of Measure" <> 'LTR-S') //RSPLSUM 01Sept2020 LTR-S added
                             AND ("Production Order"."Location Code" <> 'CAPCON') THEN//RSPLSUM 28Jan21 CAPCON added
                            BEGIN
                                ConQTY := ConQTY + recItemUOM."Qty. per Unit of Measure" * "Prod. Order Component"."Act. Consumption (Qty)";
                            END
                            ELSE BEGIN
                                ConQTY := ConQTY + "Prod. Order Component"."Act. Consumption (Qty)";
                            END;

                        //output quantity
                        IF recItem.GET("Production Order"."Source No.") THEN;
                        recItemUOM.RESET;
                        recItemUOM.SETRANGE(recItemUOM."Item No.", recItem."No.");
                        recItemUOM.SETRANGE(recItemUOM.Code, recItem."Sales Unit of Measure");
                        IF recItemUOM.FINDFIRST THEN
                            IF recItem."Sales Unit of Measure" <> '' THEN BEGIN
                                decQty := ROUND(recItemUOM."Qty. per Unit of Measure" * "Prod. Order Line"."Finished Quantity", 1);
                            END
                            ELSE BEGIN
                                decQty := ROUND("Prod. Order Line"."Finished Quantity", 1);
                            END;

                        //<<1

                        //>>1

                        IF "Prod. Order Component"."Unit of Measure Code" = 'NO' THEN
                            CurrReport.SKIP;
                        //<<1

                        LCOUNT += 1;//16Mar2017

                        //>>17Mar2017 C11

                        IF (CapacityKet <> 0) AND (SpcGrvPrim <> 0) THEN
                            C11 += ((("Prod. Order Component"."Expected Quantity" - "Prod. Order Component"."Remaining Quantity") / SpcGrvPrim) / CapacityKet)
                        ELSE
                            C11 += ((("Prod. Order Component"."Expected Quantity" - "Prod. Order Component"."Remaining Quantity") / SpcGrvPrim));

                        //<<17Mar2017 C11

                        //>>2
                        //Prod. Order Component, Footer (2) - OnPreSection()
                        decYeild := 0;

                        IF ConQTY <> 0 THEN
                            decYeild := (decQty / ConQTY) * 100;

                        //IF blnExporttoExcel THEN
                        IF blnExporttoExcel AND (LCOUNT = TCOUNT) THEN
                            MakeExcelDataBody;

                        IF CapacityKet <> 0 THEN
                            vCapacity := ((("Prod. Order Component"."Expected Quantity" - "Prod. Order Component"."Remaining Quantity") /
                            SpcGrvPrim) / CapacityKet) * 100
                        ELSE
                            vCapacity := ((("Prod. Order Component"."Expected Quantity" - "Prod. Order Component"."Remaining Quantity") /
                            SpcGrvPrim)) * 100


                        //<<2
                    end;

                    trigger OnPreDataItem()
                    begin

                        //>>1
                        ConQTY := 0;
                        //<<1

                        TCOUNT := COUNT; //16Mar2017
                        LCOUNT := 0;//16Mar2017
                    end;
                }
            }

            trigger OnAfterGetRecord()
            begin

                //>>1


                CapacityKet := 0;
                Kettle := '';
                ProductionOrderLine.RESET;
                ProductionOrderLine.SETRANGE(ProductionOrderLine.Status, "Production Order".Status);
                ProductionOrderLine.SETRANGE(ProductionOrderLine."Prod. Order No.", "Production Order"."No.");
                IF ProductionOrderLine.FINDFIRST THEN BEGIN
                    RoutingLine.RESET;
                    RoutingLine.SETRANGE(RoutingLine."Routing No.", ProductionOrderLine."Routing No.");
                    RoutingLine.SETRANGE(RoutingLine.Type, RoutingLine.Type::"Machine Center");
                    IF RoutingLine.FINDFIRST THEN BEGIN
                        IF MachinCentr.GET(RoutingLine."No.") THEN
                            CapacityKet := MachinCentr.Capacity;
                    END;

                    recRoutingHeader.RESET;
                    recRoutingHeader.SETRANGE(recRoutingHeader."No.", ProductionOrderLine."Routing No.");
                    IF recRoutingHeader.FINDFIRST THEN BEGIN
                        Kettle := recRoutingHeader.Description;
                    END;

                END;

                ILE.RESET;
                ILE.SETRANGE(ILE."Entry Type", ILE."Entry Type"::Output);
                ILE.SETRANGE(ILE."Document No.", "Production Order"."No.");
                IF ILE.FINDFIRST THEN
                    SpcGrvPrim := ILE."Density Factor";

                IF SpcGrvPrim = 0 THEN
                    SpcGrvPrim := 1;

                //<<1

                //>>2

                //Production Order, Header (1) - OnPreSection()
                IF "Production Order".GETFILTER("Date Filter") <> '' THEN
                    DateCriteria := 'Statement of Production Loss for the period of ' + FORMAT(GETRANGEMIN("Production Order"."Date Filter")) + ' . . ' +
                    FORMAT(GETRANGEMAX("Production Order"."Date Filter"))
                ELSE
                    DateCriteria := '';

                IF "Production Order".GETFILTER("Location Code") <> '' THEN
                    reclocation.GET("Production Order"."Location Code");

                //<<2

                //RSPLSUM 28Jan21>>
                CLEAR(SpecificGravity);
                CLEAR(StandardYield);
                CLEAR(Comments);
                RecItemLedgerEntry.RESET;
                RecItemLedgerEntry.SETRANGE("Order Type", RecItemLedgerEntry."Order Type"::Production);
                RecItemLedgerEntry.SETRANGE("Order No.", "Production Order"."No.");
                RecItemLedgerEntry.SETRANGE("Entry Type", RecItemLedgerEntry."Entry Type"::Output);
                IF RecItemLedgerEntry.FINDFIRST THEN BEGIN
                    SpecificGravity := RecItemLedgerEntry."Density Factor";
                    RecItemForYield.RESET;
                    IF RecItemForYield.GET(RecItemLedgerEntry."Item No.") THEN
                        StandardYield := RecItemForYield."Standard Yield";

                    RecProdOrdCmtLine.RESET;
                    RecProdOrdCmtLine.SETRANGE(Status, RecProdOrdCmtLine.Status::Finished);
                    RecProdOrdCmtLine.SETRANGE("Prod. Order No.", "Production Order"."No.");
                    IF RecProdOrdCmtLine.FINDSET THEN BEGIN
                        REPEAT
                            Comments += ',' + RecProdOrdCmtLine.Comment;
                        UNTIL RecProdOrdCmtLine.NEXT = 0;
                        Comments := COPYSTR(Comments, 2, STRLEN(Comments));
                    END;
                END;
                //RSPLSUM 28Jan21<<
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Production Order".SETFILTER("Production Order"."Last Date Modified", "Production Order".GETFILTER("Date Filter"));
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
                field("Print To Excel"; blnExporttoExcel)
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

        IF blnExporttoExcel THEN
            CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        ExcelBuf.DELETEALL;//12Sep2017

        //>>1

        recCompInfo.GET;
        IF blnExporttoExcel THEN
            MakeExcelInfo;
        //<<1
    end;

    var
        CompItem: Record 27;
        recItem: Record 27;
        recCompInfo: Record 79;
        decYeild: Decimal;
        decQty: Decimal;
        ExcelBuf: Record 370 temporary;
        blnExporttoExcel: Boolean;
        ConQTY: Decimal;
        DateCriteria: Text[150];
        t: Integer;
        reclocation: Record 14;
        SalesUOM: Code[10];
        recItemUOM: Record 5404;
        ProductionOrderLine: Record 5406;
        RoutingLine: Record 99000764;
        MachinCentr: Record 99000758;
        CapacityKet: Decimal;
        ILE: Record 32;
        SpcGrvPrim: Decimal;
        recRoutingHeader: Record 99000763;
        Kettle: Text[30];
        vCapacity: Decimal;
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        Text000: Label 'Data';
        Text001: Label 'Yeild Report';
        "------16Mar2017": Integer;
        TCOUNT: Integer;
        LCOUNT: Integer;
        C11: Decimal;
        RecItemLedgerEntry: Record 32;
        SpecificGravity: Decimal;
        RecItemForYield: Record 27;
        StandardYield: Decimal;
        RecProdOrdCmtLine: Record 5414;
        Comments: Text[1024];

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//16Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Yeild Report', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Yeild Report", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>15Mar2017 RB-N
        /* ExcelBuf.CreateBook('', Text001);
        ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
        ExcelBuf.GiveUserControl; */
        //<<

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
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
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('Yeild Report', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11 RSPLSUM 28Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12 RSPLSUM 28Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13 RSPLSUM 28Jan21
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        //ExcelBuf.NewRow;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12 RSPLSUM 28Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13 RSPLSUM 28Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 RSPLSUM 28Jan21
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('Blend No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('RM Consumed', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('FG Output', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Yeild %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Production Loss', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Kettle No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Kettle Capacity utilization %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Comments', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12 RSPLSUM 28Jan21
        ExcelBuf.AddColumn('Specific Gravity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13 RSPLSUM 28Jan21
        ExcelBuf.AddColumn('Standard Yield', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 RSPLSUM 28Jan21
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Prod. Order Component"."Prod. Order No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Production Order"."Last Date Modified", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn("Production Order"."Description 2", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn("Production Order"."Source No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4  //sn
        //ExcelBuf.AddColumn("Prod. Order Component"."Item No.",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(recItem.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(ConQTY, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//6
        ExcelBuf.AddColumn(decQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//7
        ExcelBuf.AddColumn(ROUND(decYeild, 0.01), FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//8 %
        ExcelBuf.AddColumn(decQty - ConQTY, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//9
        ExcelBuf.AddColumn(Kettle, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
                                                                          /*
                                                                          IF (CapacityKet<>0) AND (SpcGrvPrim<>0) THEN
                                                                             ExcelBuf.AddColumn(((("Prod. Order Component"."Expected Quantity"-"Prod. Order Component"."Remaining Quantity")/SpcGrvPrim)
                                                                            /CapacityKet)*100,FALSE,'',FALSE,FALSE,FALSE,'0.00',0)//11
                                                                          ELSE
                                                                             ExcelBuf.AddColumn(((("Prod. Order Component"."Expected Quantity"-"Prod. Order Component"."Remaining Quantity"))
                                                                            )*100,FALSE,'',FALSE,FALSE,FALSE,'0.00',0);//11

                                                                          *///Commented 16Mar2017

        IF (CapacityKet <> 0) AND (SpcGrvPrim <> 0) THEN
            ExcelBuf.AddColumn((C11) * 100, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0)//11
        ELSE
            ExcelBuf.AddColumn((C11) * 100, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//11

        C11 := 0;//16Mar2017

        ExcelBuf.AddColumn(Comments, FALSE, '', FALSE, FALSE, FALSE, '', 1);//12 RSPLSUM 28Jan21
        ExcelBuf.AddColumn(SpecificGravity, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//13 RSPLSUM 28Jan21
        ExcelBuf.AddColumn(StandardYield, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//14 RSPLSUM 28Jan21

    end;
}


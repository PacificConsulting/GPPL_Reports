report 50186 "Summary of CSO/Trf Indent"//"Summary of CSO_Trf Indent"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SummaryofCSOTrfIndent.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Sales Line"; "Sales Line")
        {
            DataItemTableView = SORTING("Unit of Measure Code")
                                WHERE(Closed = FILTER(false),
                                      "Document Type" = FILTER(Order),
                                      "Gen. Bus. Posting Group" = FILTER('DOMESTIC'));

            trigger OnAfterGetRecord()
            begin
                //R
                /*
                recSH.RESET;
                recSH.SETRANGE(recSH."Short Close",FALSE);
                recSH.SETRANGE(recSH.Status,recSH.Status::Released);
                recSH.SETRANGE(recSH."No.","Sales Line"."Document No.");
                IF NOT recSH.FINDFIRST THEN
                   CurrReport.SKIP;
                
                
                SalesHeader.RESET;
                SalesHeader.SETRANGE(SalesHeader."No.","Sales Line"."Document No.");
                IF SalesHeader.FINDFIRST THEN
                   PostingDateforSales := SalesHeader."Posting Date";
                   docdate:=SalesHeader."Document Date";
                 IF NOT ((docdate >= StartDate) AND (docdate <= EndDate)) THEN
                    CurrReport.SKIP;
                */
                //<<
                CLEAR(BinQty);
                recBinContent.RESET;
                recBinContent.SETCURRENTKEY("Location Code", "Bin Code", "Item No.", "Variant Code", "Unit of Measure Code");
                recBinContent.SETRANGE(recBinContent."Item No.", "Sales Line"."No.");
                recBinContent.SETRANGE(recBinContent."Location Code", LocCode);
                recBinContent.SETFILTER(recBinContent."Bin Code", 'V-S-INDL');
                IF recBinContent.FIND('-') THEN BEGIN
                    recBinContent.CALCFIELDS(recBinContent.Quantity);
                    BinQty := BinQty + recBinContent.Quantity;
                END;

                "Last5daysQty-S" := 0;
                recSL.RESET;
                recSL.SETCURRENTKEY("Document Type", "Sell-to Customer No.", "Shipment No.");
                recSL.SETRANGE(recSL."Shipment Date", Last5daysDate, EndDate);
                recSL.SETRANGE(recSL."Location Code", LocCode);
                recSL.SETRANGE(recSL."Unit of Measure Code", "Sales Line"."Unit of Measure Code");
                recSL.SETFILTER(recSL."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                IF recSL.FINDFIRST THEN
                    REPEAT

                        "Last5daysQty-S" := "Last5daysQty-S" + recSL.Quantity;
                    UNTIL recSL.NEXT = 0;

                IF ("Sales Line"."Unit of Measure Code" = 'KGS') OR ("Sales Line"."Unit of Measure Code" = 'LTRS') THEN BEGIN
                    "Last5daysQty-S" := 0;
                END;

                IF ("Sales Line".Quantity - "Sales Line"."Quantity Shipped") = 0 THEN BEGIN
                    CurrReport.SKIP;
                END;

                IF UnitOfMeasure4 <> "Sales Line"."Unit of Measure Code" THEN BEGIN
                    IF NOT (("Sales Line"."Unit of Measure Code" = 'KGS') OR ("Sales Line"."Unit of Measure Code" = 'LTRS')) THEN BEGIN
                        "itemqtytotal-S" := "itemqtytotal-S" + "Sales Line".Quantity;
                        "itemqtytotalinLtrs-S" := "itemqtytotalinLtrs-S" + ("Sales Line".Quantity * "Sales Line"."Qty. per Unit of Measure");
                        "totqtyship-S" := "totqtyship-S" + "Sales Line"."Quantity Shipped";
                        "totqtyshipinLtrs-S" := "totqtyshipinLtrs-S" + ("Sales Line"."Quantity Shipped" * "Sales Line"."Qty. per Unit of Measure");
                        "totqtytoship-S" := "totqtytoship-S" + ("Sales Line".Quantity - "Sales Line"."Quantity Shipped");
                        "totqtytoshipinLtrs-S" := "totqtytoshipinLtrs-S" + (("Sales Line".Quantity * "Sales Line"."Qty. per Unit of Measure") -
                                                                           ("Sales Line"."Quantity Shipped" * "Sales Line"."Qty. per Unit of Measure"));
                    END;


                    //>>Footer Setction

                    //IF "Sales Line"."Gen. Bus. Posting Group" = 'DOMESTIC' THEN
                    //BEGIN

                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn("Sales Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                    //ExcelBuf.AddColumn("Sales Line"."Qty. per Unit of Measure",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn("Sales Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Sales Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017 //3
                                                                                                          //>>27Mar2017  SordQty
                    SOrdQty += "Sales Line".Quantity;
                    //<<27Mar2017  SordQty

                    ExcelBuf.AddColumn("Sales Line".Quantity * "Sales Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017 //4
                                                                                                                                                    //>>27Mar2017  SordQtyL
                    SOrdQtyL += "Sales Line".Quantity * "Sales Line"."Qty. per Unit of Measure";
                    //<<27Mar2017  SordQtyL

                    ExcelBuf.AddColumn("Sales Line"."Quantity Shipped", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017 //5
                                                                                                                    //>>27Mar2017  SShipQty
                    SshpQty += "Sales Line"."Quantity Shipped";
                    //<<27Mar2017  SshipQty

                    ExcelBuf.AddColumn("Sales Line"."Quantity Shipped" * "Sales Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017 //6
                                                                                                                                                              //>>27Mar2017  SshpQtyL
                    SshpQtyL += "Sales Line"."Quantity Shipped" * "Sales Line"."Qty. per Unit of Measure";
                    //<<27Mar2017  SshpQtyL

                    ExcelBuf.AddColumn("Sales Line".Quantity - "Sales Line"."Quantity Shipped", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017 //7

                    ExcelBuf.AddColumn((("Sales Line".Quantity * "Sales Line"."Qty. per Unit of Measure")
                     - ("Sales Line"."Quantity Shipped" * "Sales Line"."Qty. per Unit of Measure")), FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017 //8

                    ExcelBuf.AddColumn("Last5daysQty-S", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Last5daysQty-S" * "Sales Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Sales Line"."Gen. Bus. Posting Group", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                    ExcelBuf.AddColumn('Order', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                    i += 1;
                END;
                UnitOfMeasure4 := "Sales Line"."Unit of Measure Code";

            end;

            trigger OnPostDataItem()
            begin

                //IF "Sales Line"."Gen. Bus. Posting Group" = 'DOMESTIC' THEN
                //BEGIN
                ExcelBuf.NewRow;
                ExcelBuf.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                ExcelBuf.AddColumn('Excl. Kgs & Ltrs-SL', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                //ExcelBuf.AddColumn("itemqtytotal-S",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(SOrdQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("itemqtytotalinLtrs-S",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(SOrdQtyL, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("totqtyship-S",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(SshpQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("totqtyshipinLtrs-S",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(SshpQtyL, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("totqtytoship-S",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(SOrdQty - SshpQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("totqtytoshipinLtrs-S",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(SOrdQtyL - SshpQtyL, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017


                i += 1;
                //END;
                //UnitOfMeasure4 := "Sales Line"."Unit of Measure"
            end;

            trigger OnPreDataItem()
            begin

                "Sales Line".SETRANGE("Sales Line"."Location Code", LocCode);
                "Sales Line".SETFILTER("Sales Line"."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                Last5daysDate := CALCDATE('-4D', EndDate);
            end;
        }
        dataitem("Sales Line0"; "Sales Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Closed = FILTER(false));

            trigger OnAfterGetRecord()
            begin
                //>>RSPL
                /*
                recSH.RESET;
                recSH.SETRANGE(recSH."Short Close",FALSE);
                recSH.SETRANGE(recSH.Status,recSH.Status::Released);
                recSH.SETRANGE(recSH."No.","Sales Line0"."Document No.");
                IF NOT recSH.FINDFIRST THEN
                   CurrReport.SKIP;
                
                SalesHeader.RESET;
                SalesHeader.SETRANGE(SalesHeader."No.","Sales Line0"."Document No.");
                IF SalesHeader.FINDFIRST THEN
                   PostingDateforSales := SalesHeader."Posting Date";
                   docdate:=SalesHeader."Document Date";
                 IF NOT ((docdate >= StartDate) AND (docdate <= EndDate)) THEN
                    CurrReport.SKIP;
                */
                //>>RSPL
                CLEAR(BinQty);
                recBinContent.RESET;
                recBinContent.SETCURRENTKEY(Default, "Location Code", "Item No.", "Variant Code", "Bin Code");
                recBinContent.SETRANGE(recBinContent."Item No.", "Sales Line0"."No.");
                recBinContent.SETFILTER(recBinContent."Location Code", LocCode);
                recBinContent.SETFILTER(recBinContent."Bin Code", 'V-S-INDL');
                IF recBinContent.FIND('-') THEN BEGIN
                    recBinContent.CALCFIELDS(recBinContent.Quantity);
                    BinQty := BinQty + recBinContent.Quantity;
                END;

                "Last5daysQty-E" := 0;
                recSL.RESET;
                recSL.SETRANGE(recSL."Shipment Date", Last5daysDate, EndDate);
                recSL.SETRANGE(recSL."Location Code", LocCode);
                recSL.SETRANGE(recSL."Unit of Measure Code", "Sales Line0"."Unit of Measure Code");
                recSL.SETFILTER(recSL."Gen. Bus. Posting Group", '%1', 'FOREIGN');
                IF recSL.FINDFIRST THEN
                    REPEAT
                        "Last5daysQty-E" := "Last5daysQty-E" + recSL.Quantity;
                    UNTIL recSL.NEXT = 0;

                IF ("Sales Line0"."Unit of Measure Code" = 'KGS') OR ("Sales Line0"."Unit of Measure Code" = 'LTRS') THEN BEGIN
                    "Last5daysQty-E" := 0;
                END;

                IF ("Sales Line".Quantity - "Sales Line"."Quantity Shipped") = 0 THEN BEGIN
                    CurrReport.SKIP;
                END;

                IF UnitOfMeasure <> "Unit of Measure Code" THEN BEGIN
                    IF NOT (("Sales Line0"."Unit of Measure Code" = 'KGS') OR ("Sales Line0"."Unit of Measure Code" = 'LTRS')) THEN BEGIN
                        "itemqtytotal-E" := "itemqtytotal-E" + "Sales Line0".Quantity;
                        "itemqtytotalinLtrs-E" := "itemqtytotalinLtrs-E" + ("Sales Line0".Quantity * "Sales Line0"."Qty. per Unit of Measure");
                        "totqtyship-E" := "totqtyship-E" + "Sales Line0"."Quantity Shipped";
                        "totqtyshipinLtrs-E" := "totqtyshipinLtrs-E" + ("Sales Line0"."Quantity Shipped" * "Sales Line0"."Qty. per Unit of Measure");
                        "totqtytoship-E" := "totqtytoship-E" + ("Sales Line0".Quantity - "Sales Line0"."Quantity Shipped");
                        "totqtytoshipinLtrs-E" := "totqtytoshipinLtrs-E" + (("Sales Line0".Quantity * "Sales Line0"."Qty. per Unit of Measure") -
                                                                           ("Sales Line0"."Quantity Shipped" * "Sales Line0"."Qty. per Unit of Measure"));
                    END;

                    //>>Footer Section

                    //IF "Sales Line0"."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                    //BEGIN

                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn("Sales Line0"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                    ExcelBuf.AddColumn("Sales Line0"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
                    ExcelBuf.AddColumn("Sales Line0".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
                                                                                                           //>>27Mar2017 CSordQty
                    CSOrdQty += "Sales Line0".Quantity;
                    //<<27Mar2017 CSordQty

                    ExcelBuf.AddColumn("Sales Line0".Quantity * "Sales Line0"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
                                                                                                                                                      //>>27Mar2017 CSordQtyL
                    CSOrdQtyL += "Sales Line0".Quantity * "Sales Line0"."Qty. per Unit of Measure";
                    //<<27Mar2017 CSordQtyL

                    ExcelBuf.AddColumn("Sales Line0"."Quantity Shipped", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
                                                                                                                     //>>27Mar2017 CSshipqty
                    CSshpQty += "Sales Line0"."Quantity Shipped";
                    //<<27Mar2017 CSshipqty

                    ExcelBuf.AddColumn("Sales Line0"."Quantity Shipped" * "Sales Line0"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
                                                                                                                                                                //>>27Mar2017 CSshipqtyL
                    CSshpQtyL += "Sales Line0"."Quantity Shipped" * "Sales Line0"."Qty. per Unit of Measure";
                    //<<27Mar2017 CSshipqtyL

                    ExcelBuf.AddColumn("Sales Line0".Quantity - "Sales Line0"."Quantity Shipped", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn(("Sales Line0".Quantity * "Sales Line0"."Qty. per Unit of Measure")
                    - ("Sales Line0"."Quantity Shipped" * "Sales Line0"."Qty. per Unit of Measure"), FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Last5daysQty-E", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Last5daysQty-E" * "Sales Line0"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Sales Line0"."Gen. Bus. Posting Group", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn('Order', FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    i += 1;
                END;
                UnitOfMeasure := "Sales Line0"."Unit of Measure";

            end;

            trigger OnPostDataItem()
            begin

                //IF "Sales Line0"."Gen. Bus. Posting Group" = 'FOREIGN' THEN
                //BEGIN
                //IF UnitOfMeasure2 <> "Sales Line0"."Unit of Measure Code" THEN BEGIN
                ExcelBuf.NewRow;
                ExcelBuf.AddColumn('TOTAL SL0', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                ExcelBuf.AddColumn('EXCL.Kgs & Ltrs', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                //ExcelBuf.AddColumn("itemqtytotal-E",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(CSOrdQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("itemqtytotalinLtrs-E",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(CSOrdQtyL, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("totqtyship-E",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(CSshpQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("totqtyshipinLtrs-E",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(CSshpQtyL, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("totqtytoship-E",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(CSOrdQty - CSshpQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                //ExcelBuf.AddColumn("totqtytoshipinLtrs-E",FALSE,'',TRUE,FALSE,FALSE,'0.000',0);//27Mar2017
                ExcelBuf.AddColumn(CSOrdQtyL - CSshpQtyL, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                i += 1;
                //END;
            end;

            trigger OnPreDataItem()
            begin
                "Sales Line0".SETRANGE("Sales Line0"."Location Code", LocCode);
                //"Sales Line0".SETFILTER("Sales Line0"."Gen. Bus. Posting Group",'%1','FOREIGN');
                SETRANGE("Sales Line0"."Gen. Bus. Posting Group", 'FOREIGN');
            end;
        }
        dataitem("Transfer Indent Line"; "Transfer Indent Line")
        {
            DataItemTableView = SORTING("Unit of Measure Code")
                                WHERE(Approve = FILTER(true),
                                      Closed = FILTER(false));

            trigger OnAfterGetRecord()
            begin
                //>>R
                /*
                recTIH.RESET;
                recTIH.SETRANGE(recTIH."Transfer indent Date",StartDate,EndDate);
                recTIH.SETRANGE(recTIH."No.","Transfer Indent Line"."Document No.");
                recTIH.SETRANGE(recTIH."Short Closed",FALSE);
                recTIH.SETRANGE(recTIH."Transfer-from Code",LocCode);
                IF NOT recTIH.FINDFIRST THEN
                BEGIN
                 CurrReport.SKIP;
                END;
                */

                TransferHeader.RESET;
                TransferHeader.SETRANGE(TransferHeader."No.", "Transfer Indent Line"."Document No.");
                IF TransferHeader.FINDFIRST THEN
                    PostingDateforTransfer := TransferHeader."Posting Date";

                IF NOT ((PostingDateforTransfer >= StartDate) AND (PostingDateforTransfer <= EndDate)) THEN
                    CurrReport.SKIP;

                CLEAR(QtyperUOM);
                recItemUOMEBT.RESET;
                recItemUOMEBT.SETCURRENTKEY("Item No.", Code);
                recItemUOMEBT.SETRANGE(recItemUOMEBT."Item No.", "Transfer Indent Line"."Item No.");
                recItemUOMEBT.SETRANGE(recItemUOMEBT.Code, "Transfer Indent Line"."Unit of Measure Code");
                IF recItemUOMEBT.FINDFIRST THEN BEGIN
                    QtyperUOM := recItemUOMEBT."Qty. per Unit of Measure";
                END;

                recTIH.RESET;
                recTIH.SETRANGE(recTIH."Posting Date", StartDate, EndDate);
                recTIH.SETRANGE(recTIH."No.", "Transfer Indent Line"."Document No.");
                recTIH.SETRANGE(recTIH."Short Closed", FALSE);
                recTIH.SETRANGE(recTIH."Transfer-from Code", LocCode);
                IF recTIH.FINDFIRST THEN BEGIN
                    CLEAR(QtyShip);
                    recTSL.RESET;
                    recTSL.SETRANGE(recTSL."Transfer-from Code", LocCode);
                    recTSL.SETRANGE(recTSL."Shipment Date", StartDate, EndDate);
                    //recTSL.SETRANGE(recTSL."Transfer Indent No.", recTIH."No.");
                    recTSL.SETRANGE(recTSL."Item No.", "Transfer Indent Line"."Item No.");
                    //recTSL.SETRANGE(recTSL."Transfer Indent Line No.", "Transfer Indent Line"."Line No.");
                    IF recTSL.FINDFIRST THEN
                        REPEAT
                            QtyShip := QtyShip + recTSL.Quantity;
                        UNTIL recTSL.NEXT = 0;
                END;

                recTIH.RESET;
                recTIH.SETCURRENTKEY("No.");
                recTIH.SETRANGE(recTIH."Posting Date", Last5daysDate, EndDate);
                recTIH.SETRANGE(recTIH."No.", "Transfer Indent Line"."Document No.");
                recTIH.SETRANGE(recTIH."Short Closed", FALSE);
                recTIH.SETRANGE(recTIH."Transfer-from Code", LocCode);
                IF recTIH.FINDFIRST THEN BEGIN
                    "Last5daysQty-T" := 0;
                    recTIL.RESET;
                    recTIL.SETCURRENTKEY("Document No.", "Line No.");
                    recTIL.SETRANGE(recTIL."Addition date", Last5daysDate, EndDate);
                    recTIL.SETRANGE(recTIL."Transfer-from Code", LocCode);
                    recTIL.SETRANGE(recTIL."Unit of Measure Code", "Transfer Indent Line"."Unit of Measure Code");
                    recTIL.SETRANGE(recTIL.Approve, TRUE);
                    recTIL.SETRANGE(recTIL.Closed, FALSE);
                    IF recTIL.FINDFIRST THEN
                        REPEAT
                            "Last5daysQty-T" := "Last5daysQty-T" + recTIL.Quantity;
                        UNTIL recTIL.NEXT = 0;

                    IF (recTIL."Unit of Measure Code" = 'KGS') OR (recTIL."Unit of Measure Code" = 'LTRS') THEN BEGIN
                        "Last5daysQty-T" := 0;
                    END;

                END;

                IF ("Transfer Indent Line".Quantity - "Transfer Indent Line"."Quantity Shipped") = 0 THEN BEGIN
                    CurrReport.SKIP;
                END;
                IF UnitOfMeasure3 <> "Transfer Indent Line"."Unit of Measure Code" THEN BEGIN
                    IF NOT ((recTIL."Unit of Measure Code" = 'KGS') OR (recTIL."Unit of Measure Code" = 'LTRS')) THEN BEGIN
                        "itemqtytotal-T" := "itemqtytotal-T" + "Transfer Indent Line".Quantity;
                        "itemqtytotalinLtrs-T" := "itemqtytotalinLtrs-T" + ("Transfer Indent Line".Quantity * QtyperUOM);
                        "totqtyship-T" := "totqtyship-T" + "Transfer Indent Line"."Quantity Shipped";
                        "totqtyshipinLtrs-T" := "totqtyshipinLtrs-T" + ("Transfer Indent Line"."Quantity Shipped" * QtyperUOM);
                        "totqtytoship-T" := "totqtytoship-T" + ("Transfer Indent Line".Quantity - "Transfer Indent Line"."Quantity Shipped");
                        "totqtytoshipinLtrs-T" := "totqtytoshipinLtrs-T" + (("Transfer Indent Line".Quantity * QtyperUOM) -
                                                                           ("Transfer Indent Line"."Quantity Shipped" * QtyperUOM));
                    END;
                    //>>Footer

                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn("Transfer Indent Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

                    ExcelBuf.AddColumn(QtyperUOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Transfer Indent Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Transfer Indent Line".Quantity * QtyperUOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Transfer Indent Line"."Quantity Shipped", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Transfer Indent Line"."Quantity Shipped" * QtyperUOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Transfer Indent Line".Quantity - "Transfer Indent Line"."Quantity Shipped", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn(("Transfer Indent Line".Quantity * QtyperUOM) -
                                                         ("Transfer Indent Line"."Quantity Shipped" * QtyperUOM), FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Last5daysQty-T", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn("Last5daysQty-T" * QtyperUOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

                    ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

                    ExcelBuf.AddColumn('Indent', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                    i += 1;
                END;
                UnitOfMeasure3 := "Transfer Indent Line"."Unit of Measure Code"

            end;

            trigger OnPostDataItem()
            begin
                ExcelBuf.NewRow;
                ExcelBuf.AddColumn('TOTAL Transfer', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                ExcelBuf.AddColumn('Excel. Kgs & Ltrs', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                ExcelBuf.AddColumn("itemqtytotal-T", FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                ExcelBuf.AddColumn("itemqtytotalinLtrs-T", FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                ExcelBuf.AddColumn("totqtyship-T", FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                ExcelBuf.AddColumn("totqtyshipinLtrs-T", FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                ExcelBuf.AddColumn("totqtytoship-T", FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                ExcelBuf.AddColumn("totqtytoshipinLtrs-T", FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017

                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);


                i += 1;
            end;

            trigger OnPreDataItem()
            begin

                "Transfer Indent Line".SETRANGE("Transfer Indent Line"."Addition date", StartDate, EndDate);
                "Transfer Indent Line".SETRANGE("Transfer Indent Line"."Transfer-from Code", LocCode);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(StartDate; StartDate)
                {
                    ApplicationArea = all;
                    Caption = 'Start Date';
                }
                field(EndDate; EndDate)
                {
                    ApplicationArea = all;
                    Caption = 'End Date';
                }
                field(LocCode; LocCode)
                {
                    ApplicationArea = all;
                    Caption = 'Location Code';
                    TableRelation = Location;
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

        //XLSHEET.Columns.AutoFit;
        //MESSAGE('Report Finished');
        //XLAPP.Visible(TRUE);
        /* ExcelBuf.CreateBookAndOpenExcel('', 'Summary of CSO/Trf Indent', '', '', '');
        ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook('Summary of CSO/Trf Indent');
        ExcelBuf.WriteSheet('Summary of CSO/Trf Indent', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Summary of CSO/Trf Indent', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

    end;

    trigger OnPreReport()
    begin
        /*
        //EBT STIVAN ---(290062013)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        User.GET(USERID);
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",User."User ID");
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2|%3','SUPER','REPORT VIEW','PENDING-INDENT/CSO');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.RESET;
         recLocation.SETRANGE(recLocation.Code,LocCode);
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping1.RESET;
         CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
         CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
         CSOMapping1.SETRANGE(CSOMapping1.Value,LocCode);
         IF NOT(CSOMapping1.FINDFIRST) THEN
         BEGIN
          CSOMapping1.RESET;
          CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
          CSOMapping1.SETRANGE(CSOMapping1.Type,CSOMapping1.Type::"Responsibility Center");
          CSOMapping1.SETRANGE(CSOMapping1.Value,LocResC);
          IF NOT(CSOMapping1.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        //EBT STIVAN ---(29062013)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        */
        CompInfo.GET;
        CreateXLSHEET;
        CreateHeader;
        i := 5;

    end;

    var
        StartDate: Date;
        EndDate: Date;
        LocCode: Code[10];
        QtyShip: Decimal;
        recTSL: Record 5745;
        "itemqtytotal-S": Decimal;
        "itemqtytotalinLtrs-S": Decimal;
        "totqtyship-S": Decimal;
        "totqtyshipinLtrs-S": Decimal;
        "totqtytoship-S": Decimal;
        "totqtytoshipinLtrs-S": Decimal;
        "itemqtytotal-T": Decimal;
        "itemqtytotalinLtrs-T": Decimal;
        "totqtyship-T": Decimal;
        "totqtyshipinLtrs-T": Decimal;
        "totqtytoship-T": Decimal;
        "totqtytoshipinLtrs-T": Decimal;
        "itemqtytotal-E": Decimal;
        "itemqtytotalinLtrs-E": Decimal;
        "totqtyship-E": Decimal;
        "totqtyshipinLtrs-E": Decimal;
        "totqtytoship-E": Decimal;
        "totqtytoshipinLtrs-E": Decimal;
        recBinContent: Record 7302;
        BinQty: Decimal;
        CBinQty: Decimal;
        recWarehouseEntry: Record 7312;
        recSL: Record 37;
        recILE: Record 32;
        Last5daysDate: Date;
        "Last5daysQty-S": Decimal;
        recTIH: Record 50022;
        "Last5daysQty-E": Decimal;
        "Last5daysQty-T": Decimal;
        recTIL: Record 50023;
        MaxCol: Text[3];
        CompInfo: Record 79;
        i: Integer;
        QtyperUOM: Decimal;
        recItemUOMEBT: Record 5404;
        recSH: Record 36;
        SalesHeader: Record 36;
        PostingDateforSales: Date;
        docdate: Date;
        TransferHeader: Record 50022;
        PostingDateforTransfer: Date;
        LocResC: Code[30];
        recLocation: Record 14;
        CSOMapping1: Record 50006;
        "--": Integer;
        UnitOfMeasure: Code[30];
        UnitOfMeasure2: Code[30];
        UnitOfMeasure3: Code[30];
        UnitOfMeasure4: Code[30];
        "---RSPL-Migration--": Integer;
        ExcelBuf: Record 370 temporary;
        vLocation: Record 14;
        "----27Mar2017": Integer;
        SOrdQty: Decimal;
        SOrdQtyL: Decimal;
        SshpQty: Decimal;
        SshpQtyL: Decimal;
        CSOrdQty: Decimal;
        CSOrdQtyL: Decimal;
        CSshpQty: Decimal;
        CSshpQtyL: Decimal;

    //   //[Scope('Internal')]
    procedure CreateXLSHEET()
    begin
    end;

    // //[Scope('Internal')]
    procedure CreateHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);


        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        IF vLocation.GET(LocCode) THEN;
        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn('Location : '+vLocation.Name,FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Summary of CSO/Trf Indent', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);


        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn('Summary of CSO/Trf Indent',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location : ' + vLocation.Name, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27Mar2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty per UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Order Qty in Ltrs', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Dispatched Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Dispatched Qty in Ltrs', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Balance Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Balance Qty in Ltrs', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Last 5 Days Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Last 5 Days Qty in Ltrs', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Gen. Bus. Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('type', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;
}


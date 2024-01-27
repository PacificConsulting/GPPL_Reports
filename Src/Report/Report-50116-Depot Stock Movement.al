report 50116 "Depot Stock Movement"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/DepotStockMovement.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            DataItemTableView = SORTING("Location Code", "Item No.", "Lot No.", "Posting Date", "Document No.", "Document Line No.");
            RequestFilterFields = "Location Code", "Posting Date", "Item No.";

            trigger OnAfterGetRecord()
            begin
                //IF vItemNo <> "Item Ledger Entry"."Item No." THEN BEGIN
                vReceiptsUOM := 0;
                IF recItem.GET("Item No.") THEN;
                IF recItemUOM.GET("Item No.", "Unit of Measure Code") THEN;
                IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Transfer Receipt" THEN BEGIN
                    vReceiptsQty += (Quantity);
                    vReceiptsUOM += ((Quantity * recItemUOM."Qty. per Unit of Measure"));
                END;
                IF "Entry Type" = "Entry Type"::Sale THEN BEGIN
                    vSales += ABS(Quantity);
                    vSalesUOM += ABS((Quantity * recItemUOM."Qty. per Unit of Measure"));
                END;
                IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Transfer Shipment" THEN BEGIN
                    vShipemnt += ABS(Quantity);
                    vShipmentUOM += ABS((Quantity * recItemUOM."Qty. per Unit of Measure"));
                END;


                //VK Robosoft - Begin
                RItem.RESET;
                RItem.SETRANGE(RItem."No.", "Item Ledger Entry"."Item No.");
                IF RItem.FINDFIRST THEN
                    IUOM.RESET;
                IUOM.SETRANGE(IUOM."Item No.", RItem."No.");
                IUOM.SETRANGE(IUOM.Code, RItem."Sales Unit of Measure");
                IF IUOM.FINDFIRST THEN
                    SUOM := IUOM."Qty. per Unit of Measure";
                //VK Robosoft - End


                vIndentQty := 0;
                vIndentUOM := 0;
                vDespatchedQty := 0;
                vDespatchedUOM := 0;
                vReceiptsUOM := 0;
                vStock := 0;
                vOpStock := 0;
                vStockUOM := 0;
                IF recLocation.GET("Location Code") THEN;
                //>>RSPL/Migratiom/Rahul
                vILE.RESET;
                vILE.SETRANGE("Item No.", "Item No.");
                vILE.SETRANGE("Posting Date", fromDate, toDate);
                vILE.SETRANGE("Location Code", "Location Code");
                IF vILE.FINDSET THEN
                    vCount := vILE.COUNT;

                i += 1;
                //<<

                //Calculating Indent Qty
                recTransferIndentLine.RESET;
                recTransferIndentLine.SETCURRENTKEY("Item No.", "Transfer-from Code", "Outstanding Quantity");
                recTransferIndentLine.SETRANGE(recTransferIndentLine."Item No.", "Item No.");
                recTransferIndentLine.SETRANGE(recTransferIndentLine."Transfer-to Code", "Location Code");
                IF (fromDate <> 0D) AND (toDate <> 0D) THEN
                    recTransferIndentLine.SETRANGE(recTransferIndentLine."Addition date", fromDate, toDate);
                IF recTransferIndentLine.FINDSET THEN
                    REPEAT
                        recTransferIndentLine.CALCFIELDS(recTransferIndentLine."Quantity Shipped");
                        vIndentQty += recTransferIndentLine.Quantity;
                        vDespatchedQty += recTransferIndentLine."Quantity Shipped";
                        IF recItemUOM.GET(recTransferIndentLine."Item No.", recTransferIndentLine."Unit of Measure Code") THEN;
                        vIndentUOM := vIndentQty * recItemUOM."Qty. per Unit of Measure";
                        vDespatchedUOM := vDespatchedQty * recItemUOM."Qty. per Unit of Measure";
                    UNTIL recTransferIndentLine.NEXT = 0;
                //Calculation Opening Stock
                recILE.RESET;
                //recILE.SETCURRENTKEY("Location Code", "Posting Date", "Item No.", "DSA Entry No.");
                recILE.SETRANGE(recILE."Posting Date", 0D, fromDate);
                recILE.SETRANGE(recILE."Location Code", "Location Code");
                recILE.SETRANGE(recILE."Item No.", "Item No.");
                IF recILE.FINDSET THEN
                    REPEAT
                        IF recItemUOM.GET(recILE."Item No.", recILE."Unit of Measure Code") THEN;
                        vOpStock += recILE.Quantity;
                        vOpStockUOM := (vOpStock / recItemUOM."Qty. per Unit of Measure");
                    UNTIL recILE.NEXT = 0;

                //Calculation Closing Stock
                recILE.RESET;
                // recILE.SETCURRENTKEY("Location Code", "Posting Date", "Item No.", "DSA Entry No.");
                recILE.SETRANGE(recILE."Posting Date", 0D, toDate);
                recILE.SETRANGE(recILE."Location Code", "Location Code");
                recILE.SETRANGE(recILE."Item No.", "Item No.");
                IF recILE.FINDSET THEN
                    REPEAT
                        IF recItemUOM.GET(recILE."Item No.", recILE."Unit of Measure Code") THEN;
                        vStock += recILE.Quantity;
                        vStockUOM := (vStock * recItemUOM."Qty. per Unit of Measure");
                    UNTIL recILE.NEXT = 0;
                //SN-BEGIN
                IF vCount = i THEN
                    IF (PrintToExcel) THEN BEGIN
                        RowID += 1;
                        i := 0;
                        EnterCell(RowID, 1, FORMAT(recLocation.Name), FALSE, FALSE, '@');
                        EnterCell(RowID, 2, FORMAT("Item Ledger Entry"."Item No."), FALSE, FALSE, '');
                        EnterCell(RowID, 3, FORMAT(recItem.Description), FALSE, FALSE, '');
                        EnterCell(RowID, 4, FORMAT(recItem."Sales Unit of Measure"), FALSE, FALSE, '@');
                        EnterCell(RowID, 5, FORMAT(SUOM), FALSE, FALSE, '@'); //VK
                                                                              //EnterCell(RowID,6,FORMAT(vOpStock),FALSE,FALSE,'#,####0.00');      //vOpStockUOM
                        EnterCell(RowID, 6, FORMAT(vOpStock), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                      //EnterCell(RowID,7,FORMAT(vOpStockUOM),FALSE,FALSE,'#,####0.00'); //vOpStock
                        EnterCell(RowID, 7, FORMAT(vOpStockUOM), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                         //EnterCell(RowID,8,FORMAT(vIndentUOM),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 8, FORMAT(vIndentUOM), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                        //EnterCell(RowID,9,FORMAT(vIndentQty),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 9, FORMAT(vIndentQty), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                        //EnterCell(RowID,10,FORMAT(vDespatchedUOM),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 10, FORMAT(vDespatchedUOM), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                             //EnterCell(RowID,11,FORMAT(vDespatchedQty),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 11, FORMAT(vDespatchedQty), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                             //EnterCell(RowID,12,FORMAT(vReceiptsUOM),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 12, FORMAT(vReceiptsUOM), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                           //EnterCell(RowID,13,FORMAT(vReceiptsQty),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 13, FORMAT(vReceiptsQty), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                           //EnterCell(RowID,14,FORMAT(vSalesUOM+vShipmentUOM),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 14, FORMAT(vSalesUOM + vShipmentUOM), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                                       //EnterCell(RowID,15,FORMAT(vSales+vShipemnt),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 15, FORMAT(vSales + vShipemnt), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                                 //EnterCell(RowID,16,FORMAT(vStockUOM),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 16, FORMAT(vStockUOM), FALSE, FALSE, '0.000'); //27Mar2017
                                                                                        //EnterCell(RowID,17,FORMAT(vStock),FALSE,FALSE,'#,####0.00');
                        EnterCell(RowID, 17, FORMAT(vStock), FALSE, FALSE, '0.000'); //27Mar2017

                        vReceiptsQty := 0;
                        vReceiptsUOM := 0;
                        vSalesUOM := 0;
                        vSales := 0;
                        vShipmentUOM := 0;
                        vShipemnt := 0;

                    END;
                //SN-END
                //END;
                vItemNo := "Item Ledger Entry"."Item No.";
            end;

            trigger OnPreDataItem()
            begin
                //IF "Item Ledger Entry".GETFILTER("Posting Date")<>'' THEN
                BEGIN
                    fromDate := "Item Ledger Entry".GETRANGEMIN("Item Ledger Entry"."Posting Date");
                    toDate := "Item Ledger Entry".GETRANGEMAX("Item Ledger Entry"."Posting Date");
                END;
                vLocationCode := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code");

                IF toDate = 0D THEN
                    toDate := WORKDATE;

                //IF PrintToExcel  THEN BEGIN
                //  ExcelBuf.DELETEALL(TRUE);
                //END;


                //SN-BEGIN
                IF PrintToExcel THEN BEGIN
                    RowID += 1;
                    EnterCell(RowID, 1, FORMAT(COMPANYNAME), TRUE, FALSE, '0');
                    EnterCell(RowID, 18, FORMAT(TODAY, 0, 4), TRUE, FALSE, '0');
                    RowID += 1;
                    EnterCell(RowID, 1, 'Depot Stock Movement Report', TRUE, FALSE, '0');
                    EnterCell(RowID, 3, 'From Date ' + FORMAT(fromDate), TRUE, FALSE, '0');
                    EnterCell(RowID, 4, 'To Date ' + FORMAT(toDate), TRUE, FALSE, '0');
                    EnterCell(RowID, 18, USERID, TRUE, FALSE, '0');
                    //  EnterCell(RowID,4,FORMAT(edate),TRUE,FALSE,'0');
                    RowID += 1;
                    EnterCell(RowID, 1, 'Location :', TRUE, FALSE, '0');
                    IF vLocation.GET(vLocationCode) THEN;
                    EnterCell(RowID, 2, FORMAT(vLocation.Name), TRUE, FALSE, '0');
                    //EnterCell(RowID,3,FORMAT(fromDate),TRUE,FALSE,'@');
                    //EnterCell(RowID,4,FORMAT(toDate),TRUE,FALSE,'@');

                    RowID += 1;
                END;
                //IF CurrReport.PAGENO=1 THEN
                IF PrintToExcel THEN BEGIN
                    RowID += 1;
                    EnterCell(RowID, 1, 'Location Code', TRUE, TRUE, '0');
                    EnterCell(RowID, 2, 'Item No.', TRUE, TRUE, '');
                    EnterCell(RowID, 3, 'Description', TRUE, TRUE, '');
                    EnterCell(RowID, 4, 'SKU', TRUE, TRUE, '');
                    EnterCell(RowID, 5, 'Qty per sale UOM', TRUE, TRUE, ''); //VK
                    EnterCell(RowID, 6, 'Opening Stock in Kg/Ltr', TRUE, TRUE, '');
                    EnterCell(RowID, 7, 'Opening Stock Qty in SKU', TRUE, TRUE, '');
                    EnterCell(RowID, 8, 'Intent Qty Kg/Ltr', TRUE, TRUE, '@');
                    EnterCell(RowID, 9, 'Indent Qty in SKU', TRUE, TRUE, '@');
                    EnterCell(RowID, 10, 'Despatched Qty Kg/Ltr', TRUE, TRUE, '@');
                    EnterCell(RowID, 11, 'Despatched Qty in SKU', TRUE, TRUE, '@');
                    EnterCell(RowID, 12, 'Receipts Qty Kg/Ltr', TRUE, TRUE, '@');
                    EnterCell(RowID, 13, 'Receipts Qty in SKU', TRUE, TRUE, '@');
                    EnterCell(RowID, 14, 'Sales/Transfer Kg/Ltr', TRUE, TRUE, '@');
                    EnterCell(RowID, 15, 'Sales/Transfer Qty in SKU', TRUE, TRUE, '@');
                    EnterCell(RowID, 16, 'Closing Kg/Ltr', TRUE, TRUE, '@');
                    EnterCell(RowID, 17, 'Closing Qty in SKU', TRUE, TRUE, '@');
                    EnterCell(RowID, 18, 'Reason', TRUE, TRUE, '@');
                END;
                //SN-END
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(PrintToExcel; PrintToExcel)
                {
                    Caption = 'Print to Excel';
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

        IF PrintToExcel THEN BEGIN
            //ExcelBuf.CreateBook;
            //ExcelBuf.CreateSheet('Depot Stock Movement Report','',COMPANYNAME,USERID);
            /*  ExcelBuf.CreateBookAndOpenExcel('', 'Depot Stock Movement Report', '', '', '');
             ExcelBuf.GiveUserControl; */

            ExcelBuf.CreateNewBook('Depot Stock Movement Report');
            ExcelBuf.WriteSheet('Depot Stock Movement Report', CompanyName, UserId);
            ExcelBuf.CloseBook();
            ExcelBuf.SetFriendlyFilename(StrSubstNo('Depot Stock Movement Report', CurrentDateTime, UserId));
            ExcelBuf.OpenExcel();
        END;
    end;

    var
        recItem: Record 27;
        vIndentQty: Decimal;
        vIndentUOM: Decimal;
        vDespatchedQty: Decimal;
        vDespatchedUOM: Decimal;
        vReceiptsQty: Decimal;
        vReceiptsUOM: Decimal;
        vShipemnt: Decimal;
        vShipmentUOM: Decimal;
        vSales: Decimal;
        vSalesUOM: Decimal;
        vStockUOM: Decimal;
        vStock: Decimal;
        vOpStock: Decimal;
        vOpStockUOM: Decimal;
        recTransferIndentLine: Record 50023;
        fromDate: Date;
        toDate: Date;
        recILE: Record 32;
        "----ExportToExcel-----": Integer;
        ExcelBuf: Record 370 temporary;
        RowID: Integer;
        PrintToExcel: Boolean;
        vLocationCode: Code[20];
        recLocation: Record 14;
        recItemUOM: Record 5404;
        IUOM: Record 5404;
        RItem: Record 27;
        SUOM: Decimal;
        "--": Integer;
        vItemNo: Code[30];
        vILE: Record 32;
        vCount: Integer;
        vLocation: Record 14;
        i: Integer;

    //  //[Scope('Internal')]
    procedure EnterCell(RowNo: Integer; ColumnNo: Integer; CellValue: Text[250]; Bold: Boolean; Underline: Boolean; NumberFormat: Text[30]): Boolean
    begin
        ExcelBuf.INIT;
        ExcelBuf.VALIDATE("Row No.", RowNo);
        ExcelBuf.VALIDATE("Column No.", ColumnNo);
        ExcelBuf."Cell Value as Text" := CellValue;
        ExcelBuf.Bold := Bold;
        ExcelBuf.NumberFormat := NumberFormat;
        ExcelBuf.INSERT;
    end;
}


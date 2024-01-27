report 50113 "Item Wise Sales Report"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ItemWiseSalesReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            CalcFields = "Sales Amount (Actual)", "Cost Amount (Actual)";
            DataItemTableView = SORTING("Posting Date", "Item No.")
                                ORDER(Ascending)
                                WHERE("Document Type" = FILTER("Sales Shipment" | "Transfer Shipment"));
            RequestFilterFields = "Posting Date", "Item No.", "Location Code";

            trigger OnAfterGetRecord()
            begin

                //>>1

                vPartyName := '';
                vPartyAddress := '';
                vDoc := '';
                vAmt := 0;
                //>>RSPL
                cdRetDoc := '';
                vRetQty := 0;
                vRetValue := 0;
                //<<RSPL

                IF "Document Type" = "Document Type"::"Sales Shipment" THEN BEGIN
                    recValueEntry.RESET;
                    recValueEntry.SETRANGE(recValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                    recValueEntry.SETFILTER(recValueEntry."Item Ledger Entry Type", '%1', recValueEntry."Item Ledger Entry Type"::Sale);
                    recValueEntry.SETRANGE(recValueEntry."Document Type", recValueEntry."Document Type"::"Sales Invoice");
                    IF recValueEntry.FINDFIRST THEN BEGIN
                        IF recSIH.GET(recValueEntry."Document No.") THEN BEGIN
                            vDoc := recSIH."No.";
                            recCust.GET(recSIH."Sell-to Customer No.");
                            vPartyAddress := recCust.Address + '' + recCust."Address 2" + '' + recCust.City + '' + recCust."Post Code" + '' + recCust."State Code";
                            vPartyName := recCust.Name;
                            recSIL.RESET;
                            recSIL.SETRANGE(recSIL."Document No.", recValueEntry."Document No.");
                            recSIL.SETRANGE(recSIL."Line No.", recValueEntry."Document Line No.");
                            IF recSIL.FINDFIRST THEN
                                vAmt := recSIL.Amount;
                            //>>RSPL
                            recCustLedger.RESET;
                            recCustLedger.SETRANGE(recCustLedger."Document No.", vDoc);
                            IF recCustLedger.FINDFIRST THEN BEGIN
                                recDtldCustLedger.RESET;
                                recDtldCustLedger.SETRANGE(recDtldCustLedger."Cust. Ledger Entry No.", recCustLedger."Entry No.");
                                recDtldCustLedger.SETRANGE(recDtldCustLedger."Document Type", recDtldCustLedger."Document Type"::"Credit Memo");
                                IF recDtldCustLedger.FINDFIRST THEN BEGIN
                                    cdRetDoc := recDtldCustLedger."Document No.";
                                    recSalesCredMemoLine.RESET;
                                    recSalesCredMemoLine.SETRANGE(recSalesCredMemoLine."Document No.", cdRetDoc);
                                    IF recSalesCredMemoLine.FINDSET THEN
                                        REPEAT
                                            vRetQty += recSalesCredMemoLine.Quantity;
                                            vRetValue += recSalesCredMemoLine.Amount;
                                        UNTIL recSalesCredMemoLine.NEXT = 0;
                                END;
                            END;
                            //<<RSPL
                        END
                    END;
                END
                ELSE BEGIN
                    IF recTransferShipmentHdr.GET("Item Ledger Entry"."Document No.") THEN BEGIN
                        vDoc := recTransferShipmentHdr."No.";
                        recLoc.GET(recTransferShipmentHdr."Transfer-to Code");
                        vPartyAddress := recLoc.Address + '' + recLoc."Address 2" + '' + recLoc.City + '' + recLoc."Post Code" + '' + recLoc."State Code";
                        vPartyName := recLoc.Name;
                        recTransShipmentL.RESET;
                        recTransShipmentL.SETRANGE(recTransShipmentL."Document No.", "Document No.");
                        recTransShipmentL.SETRANGE(recTransShipmentL."Line No.", "Item Ledger Entry"."Document Line No.");
                        IF recTransShipmentL.FINDFIRST THEN
                            vAmt := recTransShipmentL.Amount;
                    END;
                END;

                //<<1


                //>>2
                //Item Ledger Entry, Body (2) - OnPreSection()
                vTotQty += ABS("Item Ledger Entry".Quantity);
                IF vMonth <> DATE2DMY("Posting Date", 2) THEN BEGIN
                    IF fromDate = 0D THEN
                        fromDate := DMY2DATE(1, vMonth, DATE2DMY("Posting Date", 3));
                    toDate := DMY2DATE(LastNoofMonth(vMonth), vMonth, DATE2DMY("Posting Date", 3));
                    vTotQty -= ABS("Item Ledger Entry".Quantity);
                    Row += 1;
                    EnterCell(Row, 2, 'Quantity Dispatched from ' + FORMAT(fromDate) + ' to ' + FORMAT(toDate) + ' ->', TRUE, FALSE, '', 1);
                    EnterCell(Row, 5, FORMAT(ABS(vTotQty)), TRUE, FALSE, '#,###0.00', 0);//09May2017
                                                                                         //EnterCell(Row,5,FORMAT(ABS(vTotQty)),TRUE,FALSE,'0.00');

                    vTotQty := 0;
                    vTotQty += ABS("Item Ledger Entry".Quantity);
                    vMonth := DATE2DMY("Posting Date", 2);
                    fromDate := 0D;
                END;
                //<<2

                //>>3
                ///Item Ledger Entry, Body (3) - OnPreSection()
                Row += 1;
                EnterCell(Row, 1, FORMAT("Posting Date"), FALSE, FALSE, 'dd-mm-yyyy', 2);//09May2017
                //EnterCell(Row,1,FORMAT("Posting Date"),FALSE,FALSE,'');
                EnterCell(Row, 2, vPartyName, FALSE, FALSE, '', 1);
                EnterCell(Row, 3, vDoc, FALSE, FALSE, '', 1);
                EnterCell(Row, 4, "Item Ledger Entry"."Unit of Measure Code", FALSE, FALSE, '', 1);
                EnterCell(Row, 5, FORMAT(ABS("Item Ledger Entry".Quantity)), FALSE, FALSE, '0.000', 0);//09May2017
                TQty += ABS("Item Ledger Entry".Quantity);//09May2017

                //EnterCell(Row,5,FORMAT(ABS("Item Ledger Entry".Quantity)),FALSE,FALSE,'');
                EnterCell(Row, 6, FORMAT(vAmt), FALSE, FALSE, '#,###0.00', 0);//09May2017
                //EnterCell(Row,6,FORMAT(vAmt),FALSE,FALSE,'0.00');

                EnterCell(Row, 7, "Lot No.", FALSE, FALSE, '', 1);
                EnterCell(Row, 8, vPartyAddress, FALSE, FALSE, '', 1);
                //>>RSPL
                EnterCell(Row, 9, cdRetDoc, FALSE, FALSE, '', 1);
                EnterCell(Row, 10, FORMAT(vRetQty), FALSE, FALSE, '0.000', 0);//09May2017
                //EnterCell(Row,10,FORMAT(vRetQty),FALSE,FALSE,'0.00');
                EnterCell(Row, 11, FORMAT(vRetValue), FALSE, FALSE, '#,###0.00', 0);//09May2017
                //EnterCell(Row,11,FORMAT(vRetValue),FALSE,FALSE,'0.00');
                EnterCell(Row, 12, "Item Ledger Entry"."Item No.", FALSE, FALSE, '', 1);//09May2017
                //<<RSPL
                //<<3
            end;

            trigger OnPostDataItem()
            begin

                //>>1
                //Item Ledger Entry, Footer (5) - OnPreSection()
                vTotQty += ABS("Item Ledger Entry".Quantity); //09May2017
                IF fromDate = 0D THEN
                    fromDate := DMY2DATE(1, vMonth, DATE2DMY("Posting Date", 3));
                vTotQty -= ABS("Item Ledger Entry".Quantity);
                //>>09May2017
                Row += 1;
                EnterCell(Row, 1, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 2, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 3, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 4, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 5, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 6, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 7, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 8, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 9, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 10, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 11, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 12, '', TRUE, TRUE, '', 1);//09May2017
                //<<09May2017

                Row += 1;
                EnterCell(Row, 1, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 2, 'Quantity Dispatch from ' + FORMAT(fromDate) + ' to ' + FORMAT(toDate1) + ' ->', TRUE, TRUE, '', 1);
                EnterCell(Row, 3, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 4, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 5, FORMAT(ABS(TQty)), TRUE, TRUE, '0.000', 0);//09May2017
                TQty := 0;//09May2017
                //EnterCell(Row,5,FORMAT(ABS(vTotQty)),TRUE,FALSE,'0.00');
                EnterCell(Row, 6, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 7, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 8, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 9, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 10, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 11, '', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 12, '', TRUE, TRUE, '', 1);//09May2017

                vTotQty := 0;
                vTotQty += ABS("Item Ledger Entry".Quantity);
                vMonth := DATE2DMY("Posting Date", 2);
                fromDate := 0D;
                //<<1
            end;

            trigger OnPreDataItem()
            begin

                //>>1
                //Item Ledger Entry, Header (1) - OnPreSection()
                //>> ReportHeader 09May2017
                Row += 1;
                EnterCell(Row, 1, COMPANYNAME, TRUE, FALSE, '', 1);
                EnterCell(Row, 11, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

                Row += 1;
                EnterCell(Row, 1, 'Item Wise Sales Report', TRUE, FALSE, '', 1);
                EnterCell(Row, 11, USERID, TRUE, FALSE, '', 1);
                Row += 1;
                //<< ReportHeader 09May2017
                Row += 1;
                Row += 1;
                EnterCell(Row, 1, 'Date', TRUE, TRUE, '', 1);//09May2017
                EnterCell(Row, 2, 'Party Name', TRUE, TRUE, '', 1);
                EnterCell(Row, 3, 'Invoice No.', TRUE, TRUE, '', 1);
                EnterCell(Row, 4, 'Pack UOM', TRUE, TRUE, '', 1);
                EnterCell(Row, 5, 'Quantity', TRUE, TRUE, '', 1);
                EnterCell(Row, 6, 'Net Value', TRUE, TRUE, '', 1);
                EnterCell(Row, 7, 'Batch No.', TRUE, TRUE, '', 1);
                EnterCell(Row, 8, 'Customer/Depot. Address', TRUE, TRUE, '', 1);
                //>>RSPL
                EnterCell(Row, 9, 'Ret. Document No.', TRUE, TRUE, '', 1);
                EnterCell(Row, 10, 'Return Qty', TRUE, TRUE, '', 1);
                EnterCell(Row, 11, 'Return Value', TRUE, TRUE, '', 1);
                EnterCell(Row, 12, 'Item No.', TRUE, TRUE, '', 1);//09May2017
                //<<RSPL
                //<<1
            end;
        }
    }

    requestpage
    {

        layout
        {
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

        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet('Item wise sales Report','',COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ExcelBuf.DELETEALL(TRUE);

        //>>09May2017
        /* ExcelBuf.CreateBook('', 'Item Wise Sales Report');
        ExcelBuf.CreateBookAndOpenExcel('', 'Item Wise Sales Report', '', '', USERID);
        ExcelBuf.GiveUserControl;
 */
        ExcelBuf.CreateNewBook('Item Wise Sales Report');
        ExcelBuf.WriteSheet('Item Wise Sales Report', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Item Wise Sales Report', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1

        ExcelBuf.DELETEALL;
        vMonth := DATE2DMY("Item Ledger Entry".GETRANGEMIN("Posting Date"), 2);
        fromDate := "Item Ledger Entry".GETRANGEMIN("Posting Date");
        toDate1 := "Item Ledger Entry".GETRANGEMAX("Posting Date");
        //<<1
    end;

    var
        ExcelBuf: Record 370 temporary;
        Row: Integer;
        recValueEntry: Record 5802;
        recSIH: Record 112;
        recTransferShipmentHdr: Record 5744;
        vPartyName: Text[100];
        vPartyAddress: Text[500];
        vDoc: Code[20];
        recCust: Record 18;
        recLoc: Record 14;
        recSIL: Record 113;
        recTransShipmentL: Record 5745;
        vAmt: Decimal;
        vTotQty: Decimal;
        vMonth: Integer;
        dateLastDate: Integer;
        fromDate: Date;
        toDate: Date;
        toDate1: Date;
        cdRetDoc: Code[20];
        vRetQty: Decimal;
        vRetValue: Decimal;
        recCustLedger: Record 21;
        recDtldCustLedger: Record 379;
        recSalesCredMemoLine: Record 115;
        "---09May2017": Integer;
        TQty: Decimal;

    //  //[Scope('Internal')]
    procedure EnterCell(RowNo: Integer; ColumnNo: Integer; CellValue: Text[250]; Bold: Boolean; Underline: Boolean; NumberFormat: Text[30]; CType: Option Number,Text,Date,Time): Boolean
    begin
        ExcelBuf.INIT;
        ExcelBuf.VALIDATE("Row No.", RowNo);
        ExcelBuf.VALIDATE("Column No.", ColumnNo);
        ExcelBuf."Cell Value as Text" := CellValue;
        ExcelBuf.Bold := Bold;
        ExcelBuf.Underline := Underline;//09May2017
        ExcelBuf.NumberFormat := NumberFormat;
        ExcelBuf."Cell Type" := CType;//09May2017
        ExcelBuf.INSERT;
    end;

    // //[Scope('Internal')]
    procedure LastNoofMonth(vM: Integer) vLDate: Integer
    begin
        IF vM = 1 THEN
            EXIT(31)
        ELSE
            IF vM = 2 THEN
                EXIT(28)
            ELSE
                IF vM = 3 THEN
                    EXIT(31)
                ELSE
                    IF vM = 4 THEN
                        EXIT(30)
                    ELSE
                        IF vM = 5 THEN
                            EXIT(31)
                        ELSE
                            IF vM = 6 THEN
                                EXIT(30)
                            ELSE
                                IF vM = 7 THEN
                                    EXIT(31)
                                ELSE
                                    IF vM = 8 THEN
                                        EXIT(31)
                                    ELSE
                                        IF vM = 9 THEN
                                            EXIT(30)
                                        ELSE
                                            IF vM = 10 THEN
                                                EXIT(31)
                                            ELSE
                                                IF vM = 11 THEN
                                                    EXIT(30)
                                                ELSE
                                                    IF vM = 12 THEN
                                                        EXIT(31);
    end;
}


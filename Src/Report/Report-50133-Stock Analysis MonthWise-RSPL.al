report 50133 "Stock Analysis MonthWise-RSPL"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/StockAnalysisMonthWiseRSPL.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.")
                                WHERE(Blocked = FILTER(false),
                                      "Item Category Code" = FILTER('CAT02' | 'CAT03' | 'CAT11' | 'CAT12' | 'CAT13' | 'CAT15'));
            RequestFilterFields = "No.", "Item Category Code";
            dataitem(Location; Location)
            {
                DataItemTableView = SORTING(Code);
                RequestFilterFields = "Code";

                trigger OnAfterGetRecord()
                begin

                    //>>1


                    OpeningBalance := 0;
                    ReceiptQuantity := 0;
                    DespQuantity := 0;
                    OtherQuantity := 0;
                    ClosingBalance := 0;
                    LastMonthQuantity[1] := 0;
                    LastMonthQuantity[2] := 0;
                    LastMonthQuantity[3] := 0;
                    LastMonthQuantity[4] := 0;
                    LastMonthQuantity[5] := 0;
                    LastMonthQuantity[6] := 0;
                    LastMonthQuantity[7] := 0;
                    LastMonthQuantity[8] := 0;
                    LastMonthQuantity[9] := 0;
                    LastMonthQuantity[10] := 0;
                    LastMonthQuantity[11] := 0;
                    LastMonthQuantity[12] := 0;
                    LastMonthQuantity[13] := 0;
                    LastMonthQuantity[14] := 0;
                    LastMonthQuantity[15] := 0;
                    LastMonthQuantity[16] := 0;
                    LastMonthQuantity[17] := 0;
                    LastMonthQuantity[18] := 0;
                    LastMonthQuantity[19] := 0;
                    LastMonthQuantity[20] := 0;
                    LastMonthQuantity[21] := 0;
                    LastMonthQuantity[22] := 0;
                    LastMonthQuantity[23] := 0;
                    LastMonthQuantity[24] := 0;
                    CLEAR(LastRecQty);
                    ILE.RESET;
                    ILE.SETRANGE("Item No.", Item."No.");
                    ILE.SETFILTER(ILE."Posting Date", '<%1', RequestedDate);
                    ILE.SETRANGE(ILE."Location Code", Code);
                    IF ILE.FINDSET THEN BEGIN
                        REPEAT
                            OpeningBalance := OpeningBalance + ILE.Quantity;
                        UNTIL ILE.NEXT = 0;
                    END;

                    ILE.RESET;
                    ILE.SETRANGE("Item No.", Item."No.");
                    ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                    ILE.SETRANGE(ILE."Location Code", Code);
                    ILE.SETFILTER(ILE.Quantity, '>%1', 0);
                    IF ILE.FINDSET THEN BEGIN
                        REPEAT
                            ReceiptQuantity := ReceiptQuantity + ILE.Quantity;
                        UNTIL ILE.NEXT = 0;
                    END;

                    ILE.RESET;
                    ILE.SETRANGE("Item No.", Item."No.");
                    ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                    ILE.SETRANGE(ILE."Location Code", Code);
                    ILE.SETRANGE(ILE."Entry Type", 1);
                    ILE.SETFILTER(ILE.Quantity, '<%1', 0);
                    IF ILE.FINDSET THEN BEGIN
                        REPEAT
                            DespQuantity := DespQuantity + ABS(ILE.Quantity);
                        UNTIL ILE.NEXT = 0;
                    END;

                    ILE.RESET;
                    ILE.SETRANGE("Item No.", Item."No.");
                    ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                    ILE.SETRANGE(ILE."Location Code", Code);
                    ILE.SETFILTER(ILE."Entry Type", '<>1');
                    ILE.SETFILTER(ILE.Quantity, '<%1', 0);
                    IF ILE.FINDSET THEN BEGIN
                        REPEAT
                            OtherQuantity := OtherQuantity + ABS(ILE.Quantity);
                        UNTIL ILE.NEXT = 0;
                    END;

                    ClosingBalance := (OpeningBalance + ReceiptQuantity) - (DespQuantity + OtherQuantity);

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[1], EndDate[1]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[1] := LastMonthQuantity[1] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[2], EndDate[2]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[2] := LastMonthQuantity[2] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[3], EndDate[3]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[3] := LastMonthQuantity[3] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[4], EndDate[4]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[4] := LastMonthQuantity[4] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[5], EndDate[5]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[5] := LastMonthQuantity[5] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[6], EndDate[6]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[6] := LastMonthQuantity[6] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[7], EndDate[7]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[7] := LastMonthQuantity[7] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[8], EndDate[8]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[8] := LastMonthQuantity[8] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[9], EndDate[9]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[9] := LastMonthQuantity[9] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[10], EndDate[10]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[10] := LastMonthQuantity[10] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[11], EndDate[11]);
                    IF RequestedLocation <> '' THEN BEGIN
                        //SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code",RequestedLocation);
                        SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", RequestedLocation);
                        IF SalesInvoiceLine.FINDSET THEN BEGIN
                            REPEAT
                                LastMonthQuantity[11] := LastMonthQuantity[11] + SalesInvoiceLine."Quantity (Base)";
                            UNTIL SalesInvoiceLine.NEXT = 0;
                        END;
                    END;

                    SalesInvoiceLine.RESET;
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.", Item."No.");
                    SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date", '%1..%2', StartDate[12], EndDate[12]);
                    SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code", Code);
                    IF SalesInvoiceLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[12] := LastMonthQuantity[12] + SalesInvoiceLine."Quantity (Base)";
                        UNTIL SalesInvoiceLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[1], EndDate[1]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[13] := LastMonthQuantity[13] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[2], EndDate[2]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[14] := LastMonthQuantity[14] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[3], EndDate[3]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[15] := LastMonthQuantity[15] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[4], EndDate[4]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[16] := LastMonthQuantity[16] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[5], EndDate[5]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[17] := LastMonthQuantity[17] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[6], EndDate[6]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[18] := LastMonthQuantity[18] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[7], EndDate[7]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[19] := LastMonthQuantity[19] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[8], EndDate[8]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[20] := LastMonthQuantity[20] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[9], EndDate[9]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[21] := LastMonthQuantity[21] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[10], EndDate[10]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[22] := LastMonthQuantity[22] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;


                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[11], EndDate[11]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[23] := LastMonthQuantity[23] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    TransferShipmentLine.RESET;
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                    TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[12], EndDate[12]);
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", Code);
                    IF TransferShipmentLine.FINDSET THEN BEGIN
                        REPEAT
                            LastMonthQuantity[24] := LastMonthQuantity[24] + TransferShipmentLine."Quantity (Base)";
                        UNTIL TransferShipmentLine.NEXT = 0;
                    END;

                    //>>
                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[1], EndDate[1]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[1] := LastMonthQuantity[1] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[2], EndDate[2]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[2] := LastRecQty[2] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[3], EndDate[3]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[3] := LastRecQty[3] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[4], EndDate[4]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[4] := LastRecQty[4] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[5], EndDate[5]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[5] := LastRecQty[5] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[6], EndDate[6]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[6] := LastRecQty[6] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[7], EndDate[7]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[7] := LastRecQty[7] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[8], EndDate[8]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[8] := LastRecQty[8] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[9], EndDate[9]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[9] := LastRecQty[9] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[10], EndDate[10]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[9] := LastRecQty[9] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[10], EndDate[10]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[10] := LastRecQty[10] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[10], EndDate[10]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[10] := LastRecQty[10] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[11], EndDate[11]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[11] := LastRecQty[11] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    SalesCrMemo.RESET;
                    SalesCrMemo.SETRANGE(SalesCrMemo."No.", Item."No.");
                    SalesCrMemo.SETFILTER(SalesCrMemo."Shipment Date", '%1..%2', StartDate[12], EndDate[12]);
                    SalesCrMemo.SETRANGE(SalesCrMemo."Location Code", Code);
                    IF SalesCrMemo.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[12] := LastRecQty[12] + SalesCrMemo."Quantity (Base)";
                        UNTIL SalesCrMemo.NEXT = 0;
                    END;

                    //**TRANSFER DATA

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[1], EndDate[1]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[13] := LastRecQty[13] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;


                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[2], EndDate[2]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[14] := LastRecQty[14] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[3], EndDate[3]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[15] := LastRecQty[15] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;


                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[4], EndDate[4]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[16] := LastRecQty[16] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[5], EndDate[5]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[17] := LastRecQty[17] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[6], EndDate[6]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[18] := LastRecQty[18] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[7], EndDate[7]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[19] := LastRecQty[19] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[8], EndDate[8]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[20] := LastRecQty[20] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[9], EndDate[9]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[21] := LastRecQty[21] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[10], EndDate[10]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[22] := LastRecQty[22] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[11], EndDate[11]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[23] := LastRecQty[23] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    TransferRecLine.RESET;
                    TransferRecLine.SETRANGE(TransferRecLine."Item No.", Item."No.");
                    TransferRecLine.SETFILTER(TransferRecLine."Receipt Date", '%1..%2', StartDate[12], EndDate[12]);
                    TransferRecLine.SETRANGE(TransferRecLine."Transfer-to Code", Code);
                    IF TransferRecLine.FINDSET THEN BEGIN
                        REPEAT
                            LastRecQty[24] := LastRecQty[24] + TransferRecLine."Quantity (Base)";
                        UNTIL TransferRecLine.NEXT = 0;
                    END;

                    //<<
                    IF (OpeningBalance = 0) AND (ReceiptQuantity = 0) AND (DespQuantity = 0) AND (OtherQuantity = 0) AND (LastMonthQuantity[1] = 0)
                       AND (LastMonthQuantity[2] = 0) AND (LastMonthQuantity[3] = 0) AND (LastMonthQuantity[4] = 0) AND (LastMonthQuantity[5] = 0)
                       AND (LastMonthQuantity[6] = 0) AND (LastMonthQuantity[7] = 0) AND (LastMonthQuantity[8] = 0) AND (LastMonthQuantity[9] = 0)
                       AND (LastMonthQuantity[10] = 0) AND (LastMonthQuantity[11] = 0) AND (LastMonthQuantity[12] = 0)
                       AND (LastMonthQuantity[13] = 0) AND (LastMonthQuantity[14] = 0) AND (LastMonthQuantity[15] = 0)
                       AND (LastMonthQuantity[16] = 0) AND (LastMonthQuantity[17] = 0) AND (LastMonthQuantity[18] = 0)
                       AND (LastMonthQuantity[19] = 0) AND (LastMonthQuantity[20] = 0) AND (LastMonthQuantity[21] = 0)
                        AND (LastMonthQuantity[22] = 0) AND (LastMonthQuantity[23] = 0) AND (LastMonthQuantity[24] = 0) THEN
                        CurrReport.SKIP;

                    //<<1

                    //>>2

                    //Location, Body (1) - OnPreSection()
                    SrNo := SrNo + 1;
                    IF PrintToExcel THEN BEGIN
                        IF (OpeningBalance = 0) AND (ReceiptQuantity = 0) AND (DespQuantity = 0) AND (OtherQuantity = 0) AND (LastMonthQuantity[1] = 0)
                           AND (LastMonthQuantity[2] = 0) AND (LastMonthQuantity[3] = 0) AND (LastMonthQuantity[4] = 0) AND (LastMonthQuantity[5] = 0)
                           AND (LastMonthQuantity[6] = 0) AND (LastMonthQuantity[7] = 0) AND (LastMonthQuantity[8] = 0) AND (LastMonthQuantity[9] = 0)
                           AND (LastMonthQuantity[10] = 0) AND (LastMonthQuantity[11] = 0) AND (LastMonthQuantity[12] = 0)
                           AND (LastMonthQuantity[13] = 0) AND (LastMonthQuantity[14] = 0) AND (LastMonthQuantity[15] = 0)
                           AND (LastMonthQuantity[16] = 0) AND (LastMonthQuantity[17] = 0) AND (LastMonthQuantity[18] = 0)
                           AND (LastMonthQuantity[19] = 0) AND (LastMonthQuantity[20] = 0) AND (LastMonthQuantity[21] = 0)
                            AND (LastMonthQuantity[22] = 0) AND (LastMonthQuantity[23] = 0) AND (LastMonthQuantity[24] = 0) THEN
                            CurrReport.SKIP;
                        CreateExcelBody;
                    END;
                    //<<2
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recItemCategory.RESET;
                recItemCategory.SETRANGE(recItemCategory.Code, Item."Item Category Code");
                IF recItemCategory.FINDFIRST THEN;

                //<<1

                //>>2


                IF NOT FooterPrinted THEN
                    LastFieldNo := CurrReport.TOTALSCAUSEDBY;

                //CurrReport.SHOWOUTPUT := NOT FooterPrinted;

                FooterPrinted := TRUE;
                //<<2
            end;

            trigger OnPostDataItem()
            begin
                //>>3 06Mar2017
                //Item, Footer (5) - OnPreSection()
                IF PrintToExcel THEN
                    CreateExcelFooter;
                //<<3
            end;

            trigger OnPreDataItem()
            begin
                //>>1

                CurrReport.CREATETOTALS(OpeningBalance, ReceiptQuantity, DespQuantity, OtherQuantity, ClosingBalance);
                CurrReport.CREATETOTALS(LastMonthQuantity[1], LastMonthQuantity[2], LastMonthQuantity[3], LastMonthQuantity[4], LastMonthQuantity[5],
                                        LastMonthQuantity[6], LastMonthQuantity[7], LastMonthQuantity[8], LastMonthQuantity[9], LastMonthQuantity[10])
                ;
                CurrReport.CREATETOTALS(LastMonthQuantity[11], LastMonthQuantity[12]);
                CurrReport.CREATETOTALS(LastMonthQuantity[13], LastMonthQuantity[14], LastMonthQuantity[15], LastMonthQuantity[16],
                                        LastMonthQuantity[17], LastMonthQuantity[18], LastMonthQuantity[19], LastMonthQuantity[20],
                                        LastMonthQuantity[21], LastMonthQuantity[22]);
                CurrReport.CREATETOTALS(LastMonthQuantity[23], LastMonthQuantity[24]);

                CurrReport.CREATETOTALS(LastRecQty[1], LastRecQty[2], LastRecQty[3], LastRecQty[4], LastRecQty[5],
                                        LastRecQty[6], LastRecQty[7], LastRecQty[8], LastRecQty[9], LastRecQty[10]);

                CurrReport.CREATETOTALS(LastRecQty[11], LastRecQty[12]);
                CurrReport.CREATETOTALS(LastRecQty[13], LastRecQty[14], LastRecQty[15], LastRecQty[16], LastRecQty[17],
                                        LastRecQty[18], LastRecQty[19], LastRecQty[20], LastRecQty[21], LastRecQty[22]);

                CurrReport.CREATETOTALS(LastRecQty[23], LastRecQty[24]);
                SrNo := 0;
                IF PrintToExcel THEN
                    CreateExcelHeader;
                i := 4;

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
                field(Month; RequestedMonth)
                {
                    ApplicationArea = all;
                }
                field(year; RequestedYear)
                {
                    ApplicationArea = all;
                }
                field("Print To Excel"; PrintToExcel)
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

        //>>06Mar2017
        IF PrintToExcel THEN
            CreateExcelBook;

        //<<06Mar2017
    end;

    trigger OnPreReport()
    begin
        //>>1
        CompanyInfo.GET;
        //<<1


        //>>2
        /*
        IF PrintToExcel THEN BEGIN
          CLEAR(XLapp);
          CREATE(XLapp, TRUE, TRUE);
          Addworksheet;
        END;
        
        //*/



        IF Item.GETFILTER(Item."Item Category Code") <> '' THEN BEGIN
            IF NOT ((Item.GETFILTER(Item."Item Category Code") = 'CAT02') OR (Item.GETFILTER(Item."Item Category Code") = 'CAT03')) THEN BEGIN
                ERROR('You can not Select this Category');
            END;
        END;
        /*
        //EBT STIVAN ---(23072013)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        User.GET(USERID);
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",User."User ID");
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.SETRANGE(recLocation.Code,RequestedLocation);
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping.RESET;
         CSOMapping.SETRANGE(CSOMapping."User Id",UPPERCASE(USERID));
         CSOMapping.SETRANGE(Type,CSOMapping.Type::Location);
         CSOMapping.SETRANGE(CSOMapping.Value,RequestedLocation);
         IF NOT(CSOMapping.FINDFIRST) THEN
         BEGIN
          CSOMapping.RESET;
          CSOMapping.SETRANGE(CSOMapping."User Id",UPPERCASE(USERID));
          CSOMapping.SETRANGE(CSOMapping.Type,CSOMapping.Type::"Responsibility Center");
          CSOMapping.SETRANGE(CSOMapping.Value,LocResC);
          IF NOT(CSOMapping.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        //EBT STIVAN ---(23072013)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        */
        RequestedDate := DMY2DATE(1, RequestedMonth, RequestedYear);
        RequestedDate1 := CALCDATE('<+1M', RequestedDate) - 1;

        StartDate[1] := RequestedDate;
        EndDate[1] := CALCDATE('CM', StartDate[1]);

        StartDate[2] := CALCDATE('-1M', RequestedDate);
        EndDate[2] := CALCDATE('CM', StartDate[2]);

        StartDate[3] := CALCDATE('-2M', RequestedDate);
        EndDate[3] := CALCDATE('CM', StartDate[3]);

        StartDate[4] := CALCDATE('-3M', RequestedDate);
        EndDate[4] := CALCDATE('CM', StartDate[4]);

        StartDate[5] := CALCDATE('-4M', RequestedDate);
        EndDate[5] := CALCDATE('CM', StartDate[5]);

        StartDate[6] := CALCDATE('-5M', RequestedDate);
        EndDate[6] := CALCDATE('CM', StartDate[6]);

        StartDate[7] := CALCDATE('-6M', RequestedDate);
        EndDate[7] := CALCDATE('CM', StartDate[7]);

        StartDate[8] := CALCDATE('-7M', RequestedDate);
        EndDate[8] := CALCDATE('CM', StartDate[8]);

        StartDate[9] := CALCDATE('-8M', RequestedDate);
        EndDate[9] := CALCDATE('CM', StartDate[9]);

        StartDate[10] := CALCDATE('-9M', RequestedDate);
        EndDate[10] := CALCDATE('CM', StartDate[10]);

        StartDate[11] := CALCDATE('-10M', RequestedDate);
        EndDate[11] := CALCDATE('CM', StartDate[11]);

        StartDate[12] := CALCDATE('-11M', RequestedDate);
        EndDate[12] := CALCDATE('CM', StartDate[12]);


        Monthly[1] := 'Sale ' + FORMAT(StartDate[1], 0, '<Month text>,<Year4>');
        Monthly[2] := 'Sale ' + FORMAT(StartDate[2], 0, '<Month text>,<Year4>');
        Monthly[3] := 'Sale ' + FORMAT(StartDate[3], 0, '<Month text>,<Year4>');
        Monthly[4] := 'Sale ' + FORMAT(StartDate[4], 0, '<Month text>,<Year4>');
        Monthly[5] := 'Sale ' + FORMAT(StartDate[5], 0, '<Month text>,<Year4>');
        Monthly[6] := 'Sale ' + FORMAT(StartDate[6], 0, '<Month text>,<Year4>');
        Monthly[7] := 'Sale ' + FORMAT(StartDate[7], 0, '<Month text>,<Year4>');
        Monthly[8] := 'Sale ' + FORMAT(StartDate[8], 0, '<Month text>,<Year4>');
        Monthly[9] := 'Sale ' + FORMAT(StartDate[9], 0, '<Month text>,<Year4>');
        Monthly[10] := 'Sale ' + FORMAT(StartDate[10], 0, '<Month text>,<Year4>');
        Monthly[11] := 'Sale ' + FORMAT(StartDate[11], 0, '<Month text>,<Year4>');
        Monthly[12] := 'Sale ' + FORMAT(StartDate[12], 0, '<Month text>,<Year4>');

        Monthly[13] := 'Transfer ' + FORMAT(StartDate[1], 0, '<Month text>,<Year4>');
        Monthly[14] := 'Transfer ' + FORMAT(StartDate[2], 0, '<Month text>,<Year4>');
        Monthly[15] := 'Transfer ' + FORMAT(StartDate[3], 0, '<Month text>,<Year4>');
        Monthly[16] := 'Transfer ' + FORMAT(StartDate[4], 0, '<Month text>,<Year4>');
        Monthly[17] := 'Transfer ' + FORMAT(StartDate[5], 0, '<Month text>,<Year4>');
        Monthly[18] := 'Transfer ' + FORMAT(StartDate[6], 0, '<Month text>,<Year4>');
        Monthly[19] := 'Transfer ' + FORMAT(StartDate[7], 0, '<Month text>,<Year4>');
        Monthly[20] := 'Transfer ' + FORMAT(StartDate[8], 0, '<Month text>,<Year4>');
        Monthly[21] := 'Transfer ' + FORMAT(StartDate[9], 0, '<Month text>,<Year4>');
        Monthly[22] := 'Transfer ' + FORMAT(StartDate[10], 0, '<Month text>,<Year4>');
        Monthly[23] := 'Transfer ' + FORMAT(StartDate[11], 0, '<Month text>,<Year4>');
        Monthly[24] := 'Transfer ' + FORMAT(StartDate[12], 0, '<Month text>,<Year4>');

        //>>RSPL-rahul
        //**Code added for Sales Return and Transfer receipt
        MonthlyReturn[1] := 'Sale Return ' + FORMAT(StartDate[1], 0, '<Month text>,<Year4>');
        MonthlyReturn[2] := 'Sale Return ' + FORMAT(StartDate[2], 0, '<Month text>,<Year4>');
        MonthlyReturn[3] := 'Sale Return ' + FORMAT(StartDate[3], 0, '<Month text>,<Year4>');
        MonthlyReturn[4] := 'Sale Return ' + FORMAT(StartDate[4], 0, '<Month text>,<Year4>');
        MonthlyReturn[5] := 'Sale Return ' + FORMAT(StartDate[5], 0, '<Month text>,<Year4>');
        MonthlyReturn[6] := 'Sale Return ' + FORMAT(StartDate[6], 0, '<Month text>,<Year4>');
        MonthlyReturn[7] := 'Sale Return ' + FORMAT(StartDate[7], 0, '<Month text>,<Year4>');
        MonthlyReturn[8] := 'Sale Return ' + FORMAT(StartDate[8], 0, '<Month text>,<Year4>');
        MonthlyReturn[9] := 'Sale Return ' + FORMAT(StartDate[9], 0, '<Month text>,<Year4>');
        MonthlyReturn[10] := 'Sale Return ' + FORMAT(StartDate[10], 0, '<Month text>,<Year4>');
        MonthlyReturn[11] := 'Sale Return ' + FORMAT(StartDate[11], 0, '<Month text>,<Year4>');
        MonthlyReturn[12] := 'Sale Return ' + FORMAT(StartDate[12], 0, '<Month text>,<Year4>');

        MonthlyReturn[13] := 'Transfer receipt ' + FORMAT(StartDate[1], 0, '<Month text>,<Year4>');
        MonthlyReturn[14] := 'Transfer receipt ' + FORMAT(StartDate[2], 0, '<Month text>,<Year4>');
        MonthlyReturn[15] := 'Transfer receipt ' + FORMAT(StartDate[3], 0, '<Month text>,<Year4>');
        MonthlyReturn[16] := 'Transfer receipt ' + FORMAT(StartDate[4], 0, '<Month text>,<Year4>');
        MonthlyReturn[17] := 'Transfer receipt ' + FORMAT(StartDate[5], 0, '<Month text>,<Year4>');
        MonthlyReturn[18] := 'Transfer receipt ' + FORMAT(StartDate[6], 0, '<Month text>,<Year4>');
        MonthlyReturn[19] := 'Transfer receipt ' + FORMAT(StartDate[7], 0, '<Month text>,<Year4>');
        MonthlyReturn[20] := 'Transfer receipt ' + FORMAT(StartDate[8], 0, '<Month text>,<Year4>');
        MonthlyReturn[21] := 'Transfer receipt ' + FORMAT(StartDate[9], 0, '<Month text>,<Year4>');
        MonthlyReturn[22] := 'Transfer receipt ' + FORMAT(StartDate[10], 0, '<Month text>,<Year4>');
        MonthlyReturn[23] := 'Transfer receipt ' + FORMAT(StartDate[11], 0, '<Month text>,<Year4>');
        MonthlyReturn[24] := 'Transfer receipt ' + FORMAT(StartDate[12], 0, '<Month text>,<Year4>');

        //<<

        //<<2


        //>>07Mar2017 RB-N

        LocCode := Location.GETFILTER(Location.Code);

        //CLEAR(LocName);
        recLoc.RESET;
        recLoc.SETRANGE(Code, LocCode);
        IF recLoc.FINDSET THEN
            REPEAT
                LocName := LocName + recLoc.Name + ' ';
            UNTIL recLoc.NEXT = 0;
        //<<

    end;

    var
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        ILE: Record 32;
        OpeningBalance: Decimal;
        StartDate: array[12] of Date;
        EndDate: array[12] of Date;
        RequestedMonth: Option " ",January,February,March,April,May,June,July,August,September,October,November,December;
        RequestedYear: Integer;
        RequestedDate: Date;
        RequestedDate1: Date;
        RequestedLocation: Code[20];
        PurchRecptLine: Record 121;
        ReturnRecptLine: Record 6661;
        ReceiptQuantity: Decimal;
        DespQuantity: Decimal;
        OtherQuantity: Decimal;
        ClosingBalance: Decimal;
        SrNo: Integer;
        CompanyInfo: Record 79;
        LastMonthQuantity: array[24] of Decimal;
        SalesInvoiceLine: Record 113;
        Monthly: array[24] of Text[30];
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        recItemCategory: Record 5722;
        User: Record 2000000120;
        LocResC: Code[100];
        recLocation: Record 14;
        CSOMapping: Record 50006;
        TransferShipmentLine: Record 5745;
        ///XLRange: Automation;
        //XLSHEET: Automation;
        // XLWorkbook: Automation;
        //XLapp: Automation;
        i: Integer;
        "-----": Integer;
        MonthlyReturn: array[24] of Text[50];
        TransferRecLine: Record 5747;
        SalesCrMemo: Record 115;
        LastRecQty: array[24] of Decimal;
        Text000: Label 'Data';
        Text001: Label 'Stock Analysis Monthwise';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "----06Mar2017": Integer;
        NNN: Integer;
        recLoc: Record 14;
        LocCode: Code[20];
        LocName: Text[100];
        TOpeningBalance: Decimal;
        TReceiptQuantity: Decimal;
        TDespQuantity: Decimal;
        TOtherQuantity: Decimal;
        TClosingBalance: Decimal;
        TLastMonthQuantity: array[24] of Decimal;
        TLastRecQty: array[24] of Decimal;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        /*
        ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddInfoColumn(COMPANYNAME,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddInfoColumn('STOCK ANALYSIS',FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddInfoColumn(REPORT::"Summary of Sales/Transfer",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddInfoColumn(USERID,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddInfoColumn(TODAY,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.ClearNewRow;
         MakeExcelDataHeader;
        */

    end;

    // //[Scope('Internal')]
    procedure Addworksheet()
    begin

        /*
        IF PrintToExcel THEN BEGIN
          XLWorkbook:=XLapp.Workbooks.Add(1);
          XLSHEET:=XLWorkbook.ActiveSheet;
          XLapp.Visible:=TRUE;
        END;
        *///Commented Mar2017

    end;

    // //[Scope('Internal')]
    procedure CreateExcelHeader()
    begin
        /*
        XLSHEET.Range('A1').Value := 'Stock Analysis From';
        XLSHEET.Range('A1').Font.Bold := TRUE;
        XLSHEET.Range('B1').Value :=FORMAT(RequestedDate);
        XLSHEET.Range('B1').Font.Bold := TRUE;
        XLSHEET.Range('C1').Value := 'To';
        XLSHEET.Range('C1').Font.Bold := TRUE;
        XLSHEET.Range('D1').Value := FORMAT(RequestedDate1);
        XLSHEET.Range('D1').Font.Bold := TRUE;
        XLSHEET.Range('A2').Value := 'Location';
        XLSHEET.Range('A2').Font.Bold := TRUE;
        XLSHEET.Range('B2').Value := RequestedLocation;
        XLSHEET.Range('B2').Font.Bold := TRUE;
        XLSHEET.Range('A3').Value := 'Item No.';
        XLSHEET.Range('A3').Font.Bold := TRUE;
        XLSHEET.Range('B3').Value := 'Item Description';
        XLSHEET.Range('B3').Font.Bold := TRUE;
        XLSHEET.Range('C3').Value := 'Sale UOM';
        XLSHEET.Range('C3').Font.Bold := TRUE;
        XLSHEET.Range('D3').Value := 'Item Category';
        XLSHEET.Range('D3').Font.Bold := TRUE;
        XLSHEET.Range('E3').Value := 'Location';
        XLSHEET.Range('E3').Font.Bold := TRUE;
        XLSHEET.Range('F3').Value := 'Opening Balance';
        XLSHEET.Range('F3').Font.Bold := TRUE;
        XLSHEET.Range('G3').Value := 'Receipt Qty.';
        XLSHEET.Range('G3').Font.Bold := TRUE;
        XLSHEET.Range('H3').Value := 'Sold Qty';
        XLSHEET.Range('H3').Font.Bold := TRUE;
        XLSHEET.Range('I3').Value := 'Other Despatches';
        XLSHEET.Range('I3').Font.Bold := TRUE;
        XLSHEET.Range('J3').Value := 'Closing Balance';
        XLSHEET.Range('J3').Font.Bold := TRUE;
        
        XLSHEET.Range('K3').Value := Monthly[1];
        XLSHEET.Range('K3').Font.Bold := TRUE;
        XLSHEET.Range('L3').Value := Monthly[13];
        XLSHEET.Range('L3').Font.Bold := TRUE;
        //>>
        XLSHEET.Range('M3').Value := MonthlyReturn[1];
        XLSHEET.Range('M3').Font.Bold := TRUE;
        XLSHEET.Range('N3').Value := MonthlyReturn[13];
        XLSHEET.Range('N3').Font.Bold := TRUE;
        //<<
        
        XLSHEET.Range('O3').Value := Monthly[2];
        XLSHEET.Range('O3').Font.Bold := TRUE;
        XLSHEET.Range('P3').Value := Monthly[14];
        XLSHEET.Range('P3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('Q3').Value := MonthlyReturn[2];
        XLSHEET.Range('Q3').Font.Bold := TRUE;
        XLSHEET.Range('R3').Value := MonthlyReturn[14];
        XLSHEET.Range('R3').Font.Bold := TRUE;
        
        //<<
        
        XLSHEET.Range('S3').Value := Monthly[3];
        XLSHEET.Range('S3').Font.Bold := TRUE;
        XLSHEET.Range('T3').Value := Monthly[15];
        XLSHEET.Range('T3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('U3').Value := MonthlyReturn[3];
        XLSHEET.Range('U3').Font.Bold := TRUE;
        XLSHEET.Range('V3').Value := MonthlyReturn[15];
        XLSHEET.Range('V3').Font.Bold := TRUE;
        
        //<<
        
        XLSHEET.Range('W3').Value := Monthly[4];
        XLSHEET.Range('W3').Font.Bold := TRUE;
        XLSHEET.Range('X3').Value := Monthly[16];
        XLSHEET.Range('X3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('Y3').Value := MonthlyReturn[4];
        XLSHEET.Range('Y3').Font.Bold := TRUE;
        XLSHEET.Range('Z3').Value := MonthlyReturn[16];
        XLSHEET.Range('Z3').Font.Bold := TRUE;
        
        //<<
        
        XLSHEET.Range('AA3').Value := Monthly[5];
        XLSHEET.Range('AA3').Font.Bold := TRUE;
        XLSHEET.Range('AB3').Value := Monthly[17];
        XLSHEET.Range('AB3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('AC3').Value := MonthlyReturn[5];
        XLSHEET.Range('AC3').Font.Bold := TRUE;
        XLSHEET.Range('AD3').Value := MonthlyReturn[17];
        XLSHEET.Range('AD3').Font.Bold := TRUE;
        
        //<<
        XLSHEET.Range('AE3').Value := Monthly[6];
        XLSHEET.Range('AE3').Font.Bold := TRUE;
        XLSHEET.Range('AF3').Value := Monthly[18];
        XLSHEET.Range('AF3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('AG3').Value := MonthlyReturn[6];
        XLSHEET.Range('AG3').Font.Bold := TRUE;
        XLSHEET.Range('AH3').Value := MonthlyReturn[18];
        XLSHEET.Range('AH3').Font.Bold := TRUE;
        //<<
        
        XLSHEET.Range('AI3').Value := Monthly[7];
        XLSHEET.Range('AI3').Font.Bold := TRUE;
        XLSHEET.Range('AJ3').Value := Monthly[19];
        XLSHEET.Range('AJ3').Font.Bold := TRUE;
        //>>
        XLSHEET.Range('AK3').Value := MonthlyReturn[7];
        XLSHEET.Range('AK3').Font.Bold := TRUE;
        XLSHEET.Range('AL3').Value := MonthlyReturn[19];
        XLSHEET.Range('AL3').Font.Bold := TRUE;
        
        //<<
        
        XLSHEET.Range('AM3').Value := Monthly[8];
        XLSHEET.Range('AM3').Font.Bold := TRUE;
        XLSHEET.Range('AN3').Value := Monthly[20];
        XLSHEET.Range('AN3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('AO3').Value := MonthlyReturn[8];
        XLSHEET.Range('AO3').Font.Bold := TRUE;
        XLSHEET.Range('AP3').Value := MonthlyReturn[20];
        XLSHEET.Range('AP3').Font.Bold := TRUE;
        
        //<<
        XLSHEET.Range('AQ3').Value := Monthly[9];
        XLSHEET.Range('AQ3').Font.Bold := TRUE;
        XLSHEET.Range('AR3').Value := Monthly[21];
        XLSHEET.Range('AR3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('AS3').Value := MonthlyReturn[9];
        XLSHEET.Range('AS3').Font.Bold := TRUE;
        XLSHEET.Range('AT3').Value := MonthlyReturn[21];
        XLSHEET.Range('AT3').Font.Bold := TRUE;
        
        //<<
        XLSHEET.Range('AU3').Value := Monthly[10];
        XLSHEET.Range('AU3').Font.Bold := TRUE;
        XLSHEET.Range('AV3').Value := Monthly[22];
        XLSHEET.Range('AV3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('AW3').Value := MonthlyReturn[10];
        XLSHEET.Range('AW3').Font.Bold := TRUE;
        XLSHEET.Range('AX3').Value := MonthlyReturn[22];
        XLSHEET.Range('AX3').Font.Bold := TRUE;
        
        //<<
        XLSHEET.Range('AY3').Value := Monthly[11];
        XLSHEET.Range('AY3').Font.Bold := TRUE;
        XLSHEET.Range('AZ3').Value := Monthly[23];
        XLSHEET.Range('AZ3').Font.Bold := TRUE;
        
        //>>
        XLSHEET.Range('BA3').Value := MonthlyReturn[11];
        XLSHEET.Range('BA3').Font.Bold := TRUE;
        XLSHEET.Range('BB3').Value := MonthlyReturn[23];
        XLSHEET.Range('BB3').Font.Bold := TRUE;
        
        //<<
        
        XLSHEET.Range('BC3').Value := Monthly[12];
        XLSHEET.Range('BC3').Font.Bold := TRUE;
        XLSHEET.Range('BD3').Value := Monthly[24];
        XLSHEET.Range('BD3').Font.Bold := TRUE;
        //>>
        XLSHEET.Range('BE3').Value := MonthlyReturn[12];
        XLSHEET.Range('BE3').Font.Bold := TRUE;
        XLSHEET.Range('BF3').Value := MonthlyReturn[24];
        XLSHEET.Range('BF3').Font.Bold := TRUE;
        
        //<<
        *///Commented 06Mar2017

        //>> 06Mar2017 RB-N,
        EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '@', 1);
        EnterCell(1, 58, FORMAT(TODAY, 0, 4), TRUE, FALSE, '@', 1);
        EnterCell(2, 1, 'Stock Analysis MonthWise', TRUE, FALSE, '@', 1);
        EnterCell(2, 58, USERID, TRUE, FALSE, '@', 1);

        EnterCell(4, 1, 'Date Filter', TRUE, FALSE, '@', 1);
        EnterCell(4, 2, FORMAT(RequestedDate) + ' .. ' + FORMAT(RequestedDate1), TRUE, FALSE, '@', 1);
        //EnterCell(4,3,'To',TRUE,FALSE,'@',1);
        //EnterCell(4,4,FORMAT(RequestedDate1),TRUE,FALSE,'',2);

        EnterCell(5, 1, 'Location', TRUE, FALSE, '@', 1);
        EnterCell(5, 2, LocName, TRUE, FALSE, '@', 1);

        EnterCell(6, 1, 'Item No.', TRUE, FALSE, '@', 1);
        EnterCell(6, 2, 'Item Description', TRUE, FALSE, '@', 1);
        EnterCell(6, 3, 'Sale UOM', TRUE, FALSE, '@', 1);
        EnterCell(6, 4, 'Item Category', TRUE, FALSE, '@', 1);
        EnterCell(6, 5, 'Location', TRUE, FALSE, '@', 1);
        EnterCell(6, 6, 'Opening Balance', TRUE, FALSE, '@', 1);
        EnterCell(6, 7, 'Receipt Qty.', TRUE, FALSE, '@', 1);
        EnterCell(6, 8, 'Sold Qty', TRUE, FALSE, '@', 1);
        EnterCell(6, 9, 'Other Despatches', TRUE, FALSE, '@', 1);
        EnterCell(6, 10, 'Closing Balance', TRUE, FALSE, '@', 1);
        EnterCell(6, 11, Monthly[1], TRUE, FALSE, '@', 1);
        EnterCell(6, 12, Monthly[13], TRUE, FALSE, '@', 1);
        EnterCell(6, 13, MonthlyReturn[1], TRUE, FALSE, '@', 1);
        EnterCell(6, 14, MonthlyReturn[13], TRUE, FALSE, '@', 1);
        EnterCell(6, 15, Monthly[2], TRUE, FALSE, '@', 1);
        EnterCell(6, 16, Monthly[14], TRUE, FALSE, '@', 1);
        EnterCell(6, 17, MonthlyReturn[2], TRUE, FALSE, '@', 1);
        EnterCell(6, 18, MonthlyReturn[14], TRUE, FALSE, '@', 1);
        EnterCell(6, 19, Monthly[3], TRUE, FALSE, '@', 1);
        EnterCell(6, 20, Monthly[15], TRUE, FALSE, '@', 1);
        EnterCell(6, 21, MonthlyReturn[3], TRUE, FALSE, '@', 1);
        EnterCell(6, 22, MonthlyReturn[15], TRUE, FALSE, '@', 1);
        EnterCell(6, 23, Monthly[4], TRUE, FALSE, '@', 1);
        EnterCell(6, 24, Monthly[16], TRUE, FALSE, '@', 1);
        EnterCell(6, 25, MonthlyReturn[4], TRUE, FALSE, '@', 1);
        EnterCell(6, 26, MonthlyReturn[16], TRUE, FALSE, '@', 1);
        EnterCell(6, 27, Monthly[5], TRUE, FALSE, '@', 1);
        EnterCell(6, 28, Monthly[17], TRUE, FALSE, '@', 1);
        EnterCell(6, 29, MonthlyReturn[5], TRUE, FALSE, '@', 1);
        EnterCell(6, 30, MonthlyReturn[17], TRUE, FALSE, '@', 1);
        EnterCell(6, 31, Monthly[6], TRUE, FALSE, '@', 1);
        EnterCell(6, 32, Monthly[18], TRUE, FALSE, '@', 1);
        EnterCell(6, 33, MonthlyReturn[6], TRUE, FALSE, '@', 1);
        EnterCell(6, 34, MonthlyReturn[18], TRUE, FALSE, '@', 1);
        EnterCell(6, 35, Monthly[7], TRUE, FALSE, '@', 1);
        EnterCell(6, 36, Monthly[19], TRUE, FALSE, '@', 1);
        EnterCell(6, 37, MonthlyReturn[7], TRUE, FALSE, '@', 1);
        EnterCell(6, 38, MonthlyReturn[19], TRUE, FALSE, '@', 1);
        EnterCell(6, 39, Monthly[8], TRUE, FALSE, '@', 1);
        EnterCell(6, 40, Monthly[20], TRUE, FALSE, '@', 1);
        EnterCell(6, 41, MonthlyReturn[8], TRUE, FALSE, '@', 1);
        EnterCell(6, 42, MonthlyReturn[20], TRUE, FALSE, '@', 1);
        EnterCell(6, 43, Monthly[9], TRUE, FALSE, '@', 1);
        EnterCell(6, 44, Monthly[21], TRUE, FALSE, '@', 1);
        EnterCell(6, 45, MonthlyReturn[9], TRUE, FALSE, '@', 1);
        EnterCell(6, 46, MonthlyReturn[21], TRUE, FALSE, '@', 1);
        EnterCell(6, 47, Monthly[10], TRUE, FALSE, '@', 1);
        EnterCell(6, 48, Monthly[22], TRUE, FALSE, '@', 1);
        EnterCell(6, 49, MonthlyReturn[10], TRUE, FALSE, '@', 1);
        EnterCell(6, 50, MonthlyReturn[22], TRUE, FALSE, '@', 1);
        EnterCell(6, 51, Monthly[11], TRUE, FALSE, '@', 1);
        EnterCell(6, 52, Monthly[23], TRUE, FALSE, '@', 1);
        EnterCell(6, 53, MonthlyReturn[11], TRUE, FALSE, '@', 1);
        EnterCell(6, 54, MonthlyReturn[23], TRUE, FALSE, '@', 1);
        EnterCell(6, 55, Monthly[12], TRUE, FALSE, '@', 1);
        EnterCell(6, 56, Monthly[24], TRUE, FALSE, '@', 1);
        EnterCell(6, 57, MonthlyReturn[12], TRUE, FALSE, '@', 1);
        EnterCell(6, 58, MonthlyReturn[24], TRUE, FALSE, '@', 1);

        NNN := 7;
        //<< 06Mar2017

    end;

    // //[Scope('Internal')]
    procedure CreateExcelBody()
    begin
        /*
        XLSHEET.Range('A'+FORMAT(i)).Value := Item."No.";
        XLSHEET.Range('A'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('B'+FORMAT(i)).Value := Item.Description;
        XLSHEET.Range('B'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('C'+FORMAT(i)).Value := Item."Sales Unit of Measure";
        XLSHEET.Range('C'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('D'+FORMAT(i)).Value := recItemCategory.Description;
        XLSHEET.Range('D'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('E'+FORMAT(i)).Value := Location.Name;
        XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('F'+FORMAT(i)).Value := OpeningBalance;
        XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('G'+FORMAT(i)).Value := ReceiptQuantity;
        XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('H'+FORMAT(i)).Value := DespQuantity;
        XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('I'+FORMAT(i)).Value := OtherQuantity;
        XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('J'+FORMAT(i)).Value := ClosingBalance;
        XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('K'+FORMAT(i)).Value := LastMonthQuantity[1];
        XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('L'+FORMAT(i)).Value := LastMonthQuantity[13];
        XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('M'+FORMAT(i)).Value := LastRecQty[1];
        XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('N'+FORMAT(i)).Value := LastRecQty[13];
        XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
        //<<--E
        
        XLSHEET.Range('O'+FORMAT(i)).Value := LastMonthQuantity[2];
        XLSHEET.Range('O'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('P'+FORMAT(i)).Value := LastMonthQuantity[14];
        XLSHEET.Range('P'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>S
        XLSHEET.Range('Q'+FORMAT(i)).Value := LastRecQty[2];
        XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('R'+FORMAT(i)).Value := LastRecQty[14];
        XLSHEET.Range('R'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('R'+FORMAT(i)).EntireColumn.AutoFit;
        //<< E
        
        XLSHEET.Range('S'+FORMAT(i)).Value := LastMonthQuantity[3];
        XLSHEET.Range('S'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('S'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('T'+FORMAT(i)).Value := LastMonthQuantity[15];
        XLSHEET.Range('T'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('T'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>S
        XLSHEET.Range('U'+FORMAT(i)).Value := LastRecQty[3];
        XLSHEET.Range('U'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('U'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('V'+FORMAT(i)).Value := LastRecQty[15];
        XLSHEET.Range('V'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('V'+FORMAT(i)).EntireColumn.AutoFit;
        //<<E
        
        XLSHEET.Range('W'+FORMAT(i)).Value := LastMonthQuantity[4];
        XLSHEET.Range('W'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('W'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('X'+FORMAT(i)).Value := LastMonthQuantity[16];
        XLSHEET.Range('X'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('X'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('Y'+FORMAT(i)).Value := LastRecQty[4];
        XLSHEET.Range('Y'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('Y'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('Z'+FORMAT(i)).Value := LastRecQty[16];
        XLSHEET.Range('Z'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('Z'+FORMAT(i)).EntireColumn.AutoFit;
        //<<--E
        
        XLSHEET.Range('AA'+FORMAT(i)).Value := LastMonthQuantity[5];
        XLSHEET.Range('AA'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AA'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AB'+FORMAT(i)).Value := LastMonthQuantity[17];
        XLSHEET.Range('AB'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AB'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('AC'+FORMAT(i)).Value := LastRecQty[5];
        XLSHEET.Range('AC'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AC'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AD'+FORMAT(i)).Value := LastRecQty[17];
        XLSHEET.Range('AD'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AD'+FORMAT(i)).EntireColumn.AutoFit;
        //<< --E
        
        XLSHEET.Range('AE'+FORMAT(i)).Value := LastMonthQuantity[6];
        XLSHEET.Range('AE'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AE'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AF'+FORMAT(i)).Value := LastMonthQuantity[18];
        XLSHEET.Range('AF'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AF'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('AG'+FORMAT(i)).Value := LastRecQty[6];
        XLSHEET.Range('AG'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AH'+FORMAT(i)).Value := LastRecQty[18];
        XLSHEET.Range('AH'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AH'+FORMAT(i)).EntireColumn.AutoFit;
        //<<--E
        
        XLSHEET.Range('AI'+FORMAT(i)).Value := LastMonthQuantity[7];
        XLSHEET.Range('AI'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AI'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AJ'+FORMAT(i)).Value := LastMonthQuantity[19];
        XLSHEET.Range('AJ'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AJ'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('AK'+FORMAT(i)).Value := LastRecQty[7];
        XLSHEET.Range('AK'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AK'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AL'+FORMAT(i)).Value := LastRecQty[19];
        XLSHEET.Range('AL'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AL'+FORMAT(i)).EntireColumn.AutoFit;
        
        //<<--E
        XLSHEET.Range('AM'+FORMAT(i)).Value := LastMonthQuantity[8];
        XLSHEET.Range('AM'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AM'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AN'+FORMAT(i)).Value := LastMonthQuantity[20];
        XLSHEET.Range('AN'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AN'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('AO'+FORMAT(i)).Value := LastRecQty[8];
        XLSHEET.Range('AO'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AO'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AP'+FORMAT(i)).Value := LastRecQty[20];
        XLSHEET.Range('AP'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AP'+FORMAT(i)).EntireColumn.AutoFit;
        //<<--E
        
        XLSHEET.Range('AQ'+FORMAT(i)).Value := LastMonthQuantity[9];
        XLSHEET.Range('AQ'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AQ'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AR'+FORMAT(i)).Value := LastMonthQuantity[21];
        XLSHEET.Range('AR'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AR'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('AS'+FORMAT(i)).Value := LastRecQty[9];
        XLSHEET.Range('AS'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AS'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AT'+FORMAT(i)).Value := LastRecQty[21];
        XLSHEET.Range('AT'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AT'+FORMAT(i)).EntireColumn.AutoFit;
        //<<--E
        
        XLSHEET.Range('AU'+FORMAT(i)).Value := LastMonthQuantity[10];
        XLSHEET.Range('AU'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AU'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AV'+FORMAT(i)).Value := LastMonthQuantity[22];
        XLSHEET.Range('AV'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AV'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('AW'+FORMAT(i)).Value := LastRecQty[10];
        XLSHEET.Range('AW'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AW'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AX'+FORMAT(i)).Value := LastRecQty[22];
        XLSHEET.Range('AX'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AX'+FORMAT(i)).EntireColumn.AutoFit;
        //<<--E
        
        XLSHEET.Range('AY'+FORMAT(i)).Value := LastMonthQuantity[11];
        XLSHEET.Range('AY'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AY'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AZ'+FORMAT(i)).Value := LastMonthQuantity[23];
        XLSHEET.Range('AZ'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AZ'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('BA'+FORMAT(i)).Value := LastRecQty[11];
        XLSHEET.Range('BA'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('BA'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('BB'+FORMAT(i)).Value := LastRecQty[23];
        XLSHEET.Range('BB'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('BB'+FORMAT(i)).EntireColumn.AutoFit;
        
        //<<--E
        XLSHEET.Range('BC'+FORMAT(i)).Value := LastMonthQuantity[12];
        XLSHEET.Range('BC'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('BC'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('BD'+FORMAT(i)).Value := LastMonthQuantity[24];
        XLSHEET.Range('BD'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('BD'+FORMAT(i)).EntireColumn.AutoFit;
        
        //>>--S
        XLSHEET.Range('BE'+FORMAT(i)).Value := LastRecQty[12];
        XLSHEET.Range('BE'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('BE'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('BF'+FORMAT(i)).Value := LastRecQty[24];
        XLSHEET.Range('BF'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('BF'+FORMAT(i)).EntireColumn.AutoFit;
        
        //<<--E
        i+=1;
        
        *///Commented 06Mar2017

        //>>06Mar2017 RB-N

        EnterCell(NNN, 1, Item."No.", FALSE, FALSE, '@', 1);
        EnterCell(NNN, 2, Item.Description, FALSE, FALSE, '@', 1);
        EnterCell(NNN, 3, Item."Sales Unit of Measure", FALSE, FALSE, '@', 1);
        EnterCell(NNN, 4, recItemCategory.Description, FALSE, FALSE, '@', 1);
        EnterCell(NNN, 5, Location.Name, FALSE, FALSE, '@', 1);
        EnterCell(NNN, 6, FORMAT(OpeningBalance), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TOpeningBalance := TOpeningBalance + OpeningBalance;
        //<<

        EnterCell(NNN, 7, FORMAT(ReceiptQuantity), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TReceiptQuantity := TReceiptQuantity + ReceiptQuantity;
        //<<

        EnterCell(NNN, 8, FORMAT(DespQuantity), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TDespQuantity := TDespQuantity + DespQuantity;
        //<<

        EnterCell(NNN, 9, FORMAT(OtherQuantity), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TOtherQuantity := TOtherQuantity + OtherQuantity;
        //<<

        EnterCell(NNN, 10, FORMAT(ClosingBalance), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TClosingBalance := TClosingBalance + ClosingBalance;
        //<<

        EnterCell(NNN, 11, FORMAT(LastMonthQuantity[1]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[1] += LastMonthQuantity[1];
        //<<

        EnterCell(NNN, 12, FORMAT(LastMonthQuantity[13]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[13] += LastMonthQuantity[13];
        //<<

        EnterCell(NNN, 13, FORMAT(LastRecQty[1]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[1] += LastRecQty[1];
        //<<

        EnterCell(NNN, 14, FORMAT(LastRecQty[13]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[13] += LastRecQty[13];
        //<<

        EnterCell(NNN, 15, FORMAT(LastMonthQuantity[2]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[2] += LastMonthQuantity[2];
        //<<

        EnterCell(NNN, 16, FORMAT(LastMonthQuantity[14]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[14] += LastMonthQuantity[14];
        //<<

        EnterCell(NNN, 17, FORMAT(LastRecQty[2]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[2] += LastRecQty[2];
        //<<

        EnterCell(NNN, 18, FORMAT(LastRecQty[14]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[14] += LastRecQty[14];
        //<<
        EnterCell(NNN, 19, FORMAT(LastMonthQuantity[3]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[3] += LastMonthQuantity[3];
        //<<

        EnterCell(NNN, 20, FORMAT(LastMonthQuantity[15]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[15] += LastMonthQuantity[15];
        //<<

        EnterCell(NNN, 21, FORMAT(LastRecQty[3]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[3] += LastRecQty[3];
        //<<

        EnterCell(NNN, 22, FORMAT(LastRecQty[15]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[15] += LastRecQty[15];
        //<<

        EnterCell(NNN, 23, FORMAT(LastMonthQuantity[4]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[4] += LastMonthQuantity[4];
        //<<

        EnterCell(NNN, 24, FORMAT(LastMonthQuantity[16]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[16] += LastMonthQuantity[16];
        //<<

        EnterCell(NNN, 25, FORMAT(LastRecQty[4]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[4] += LastRecQty[4];
        //<<

        EnterCell(NNN, 26, FORMAT(LastRecQty[16]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[16] += LastRecQty[16];
        //<<

        EnterCell(NNN, 27, FORMAT(LastMonthQuantity[5]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[5] += LastMonthQuantity[5];
        //<<

        EnterCell(NNN, 28, FORMAT(LastMonthQuantity[17]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[17] += LastMonthQuantity[17];
        //<<

        EnterCell(NNN, 29, FORMAT(LastRecQty[5]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[5] += LastRecQty[5];
        //<<

        EnterCell(NNN, 30, FORMAT(LastRecQty[17]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[17] += LastRecQty[17];
        //<<

        EnterCell(NNN, 31, FORMAT(LastMonthQuantity[6]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[6] += LastMonthQuantity[6];
        //<<

        EnterCell(NNN, 32, FORMAT(LastMonthQuantity[18]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[18] += LastMonthQuantity[18];
        //<<

        EnterCell(NNN, 33, FORMAT(LastRecQty[6]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[6] += LastRecQty[6];
        //<<

        EnterCell(NNN, 34, FORMAT(LastRecQty[18]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[18] += LastRecQty[18];
        //<<

        EnterCell(NNN, 35, FORMAT(LastMonthQuantity[7]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[7] += LastMonthQuantity[7];
        //<<

        EnterCell(NNN, 36, FORMAT(LastMonthQuantity[19]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[19] += LastMonthQuantity[19];
        //<<

        EnterCell(NNN, 37, FORMAT(LastRecQty[7]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[7] += LastRecQty[7];
        //<<

        EnterCell(NNN, 38, FORMAT(LastRecQty[19]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[19] += LastRecQty[19];
        //<<

        EnterCell(NNN, 39, FORMAT(LastMonthQuantity[8]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[8] += LastMonthQuantity[8];
        //<<

        EnterCell(NNN, 40, FORMAT(LastMonthQuantity[20]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[20] += LastMonthQuantity[20];
        //<<

        EnterCell(NNN, 41, FORMAT(LastRecQty[8]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[8] += LastRecQty[8];
        //<<

        EnterCell(NNN, 42, FORMAT(LastRecQty[20]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[20] += LastRecQty[20];
        //<<

        EnterCell(NNN, 43, FORMAT(LastMonthQuantity[9]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[9] += LastMonthQuantity[9];
        //<<

        EnterCell(NNN, 44, FORMAT(LastMonthQuantity[21]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[21] += LastMonthQuantity[21];
        //<<

        EnterCell(NNN, 45, FORMAT(LastRecQty[9]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[9] += LastRecQty[9];
        //<<

        EnterCell(NNN, 46, FORMAT(LastRecQty[21]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[21] += LastRecQty[21];
        //<<

        EnterCell(NNN, 47, FORMAT(LastMonthQuantity[10]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[10] += LastMonthQuantity[10];
        //<<

        EnterCell(NNN, 48, FORMAT(LastMonthQuantity[22]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[22] += LastMonthQuantity[22];
        //<<

        EnterCell(NNN, 49, FORMAT(LastRecQty[10]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[10] += LastRecQty[10];
        //<<

        EnterCell(NNN, 50, FORMAT(LastRecQty[22]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[22] += LastRecQty[22];
        //<<

        EnterCell(NNN, 51, FORMAT(LastMonthQuantity[11]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[11] += LastMonthQuantity[11];
        //<<

        EnterCell(NNN, 52, FORMAT(LastMonthQuantity[23]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[23] += LastMonthQuantity[23];
        //<<

        EnterCell(NNN, 53, FORMAT(LastRecQty[11]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[11] += LastRecQty[11];
        //<<

        EnterCell(NNN, 54, FORMAT(LastRecQty[23]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[23] += LastRecQty[23];
        //<<

        EnterCell(NNN, 55, FORMAT(LastMonthQuantity[12]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[12] += LastMonthQuantity[12];
        //<<

        EnterCell(NNN, 56, FORMAT(LastMonthQuantity[24]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastMonthQuantity[24] += LastMonthQuantity[24];
        //<<

        EnterCell(NNN, 57, FORMAT(LastRecQty[12]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[12] += LastRecQty[12];
        //<<

        EnterCell(NNN, 58, FORMAT(LastRecQty[24]), FALSE, FALSE, '#####0.000', 0);
        //>>07Mar2017
        TLastRecQty[24] += LastRecQty[24];
        //<<


        NNN += 1;
        //<<

    end;

    // //[Scope('Internal')]
    procedure CreateExcelFooter()
    begin
        /*
        XLSHEET.Range('A'+FORMAT(i)).Value := '';
        XLSHEET.Range('A'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('B'+FORMAT(i)).Value := '';
        XLSHEET.Range('B'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('C'+FORMAT(i)).Value := '';
        XLSHEET.Range('C'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('D'+FORMAT(i)).Value := '';
        XLSHEET.Range('D'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('E'+FORMAT(i)).Value :='';
        XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('F'+FORMAT(i)).Value := OpeningBalance;
        XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('G'+FORMAT(i)).Value := ReceiptQuantity;
        XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('H'+FORMAT(i)).Value := DespQuantity;
        XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('I'+FORMAT(i)).Value := OtherQuantity;
        XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('J'+FORMAT(i)).Value := ClosingBalance;
        XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('K'+FORMAT(i)).Value := LastMonthQuantity[1];
        XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('L'+FORMAT(i)).Value := LastMonthQuantity[13];
        XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('M'+FORMAT(i)).Value := LastMonthQuantity[2];
        XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('N'+FORMAT(i)).Value := LastMonthQuantity[14];
        XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('O'+FORMAT(i)).Value := LastMonthQuantity[3];
        XLSHEET.Range('O'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('P'+FORMAT(i)).Value := LastMonthQuantity[15];
        XLSHEET.Range('P'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('Q'+FORMAT(i)).Value := LastMonthQuantity[4];
        XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('R'+FORMAT(i)).Value := LastMonthQuantity[16];
        XLSHEET.Range('R'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('R'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('S'+FORMAT(i)).Value := LastMonthQuantity[5];
        XLSHEET.Range('S'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('S'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('T'+FORMAT(i)).Value := LastMonthQuantity[17];
        XLSHEET.Range('T'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('T'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('U'+FORMAT(i)).Value := LastMonthQuantity[6];
        XLSHEET.Range('U'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('U'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('V'+FORMAT(i)).Value := LastMonthQuantity[18];
        XLSHEET.Range('V'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('V'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('W'+FORMAT(i)).Value := LastMonthQuantity[7];
        XLSHEET.Range('W'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('W'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('X'+FORMAT(i)).Value := LastMonthQuantity[19];
        XLSHEET.Range('X'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('X'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('Y'+FORMAT(i)).Value := LastMonthQuantity[8];
        XLSHEET.Range('Y'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('Y'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('Z'+FORMAT(i)).Value := LastMonthQuantity[20];
        XLSHEET.Range('Z'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('Z'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AA'+FORMAT(i)).Value := LastMonthQuantity[9];
        XLSHEET.Range('AA'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AA'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AB'+FORMAT(i)).Value := LastMonthQuantity[21];
        XLSHEET.Range('AB'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AB'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AC'+FORMAT(i)).Value := LastMonthQuantity[10];
        XLSHEET.Range('AC'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AC'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AD'+FORMAT(i)).Value := LastMonthQuantity[22];
        XLSHEET.Range('AD'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AD'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AE'+FORMAT(i)).Value := LastMonthQuantity[11];
        XLSHEET.Range('AE'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AE'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AF'+FORMAT(i)).Value := LastMonthQuantity[23];
        XLSHEET.Range('AF'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AF'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AG'+FORMAT(i)).Value := LastMonthQuantity[12];
        XLSHEET.Range('AG'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
        
        XLSHEET.Range('AH'+FORMAT(i)).Value := LastMonthQuantity[24];
        XLSHEET.Range('AH'+FORMAT(i)).NumberFormat := '@';
        XLSHEET.Range('AH'+FORMAT(i)).EntireColumn.AutoFit;
        
        i+=1;
        *///Commented 06Mar2017

        //>>06Mar2017 RB-N

        EnterCell(NNN, 1, '', TRUE, FALSE, '@', 1);
        EnterCell(NNN, 2, '', TRUE, FALSE, '@', 1);
        EnterCell(NNN, 3, '', TRUE, FALSE, '@', 1);
        EnterCell(NNN, 4, '', TRUE, FALSE, '@', 1);
        EnterCell(NNN, 5, '', TRUE, FALSE, '@', 1);
        EnterCell(NNN, 6, FORMAT(TOpeningBalance), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 7, FORMAT(TReceiptQuantity), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 8, FORMAT(TDespQuantity), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 9, FORMAT(TOtherQuantity), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 10, FORMAT(TClosingBalance), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 11, FORMAT(TLastMonthQuantity[1]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 12, FORMAT(TLastMonthQuantity[13]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 13, FORMAT(TLastRecQty[1]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 14, FORMAT(TLastRecQty[13]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 15, FORMAT(TLastMonthQuantity[2]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 16, FORMAT(TLastMonthQuantity[14]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 17, FORMAT(TLastRecQty[2]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 18, FORMAT(TLastRecQty[14]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 19, FORMAT(TLastMonthQuantity[3]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 20, FORMAT(TLastMonthQuantity[15]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 21, FORMAT(TLastRecQty[3]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 22, FORMAT(TLastRecQty[15]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 23, FORMAT(TLastMonthQuantity[4]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 24, FORMAT(TLastMonthQuantity[16]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 25, FORMAT(TLastRecQty[4]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 26, FORMAT(TLastRecQty[16]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 27, FORMAT(TLastMonthQuantity[5]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 28, FORMAT(TLastMonthQuantity[17]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 29, FORMAT(TLastRecQty[5]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 30, FORMAT(TLastRecQty[17]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 31, FORMAT(TLastMonthQuantity[6]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 32, FORMAT(TLastMonthQuantity[18]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 33, FORMAT(TLastRecQty[6]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 34, FORMAT(TLastRecQty[18]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 35, FORMAT(TLastMonthQuantity[7]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 36, FORMAT(TLastMonthQuantity[19]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 37, FORMAT(TLastRecQty[7]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 38, FORMAT(TLastRecQty[19]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 39, FORMAT(TLastMonthQuantity[8]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 40, FORMAT(TLastMonthQuantity[20]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 41, FORMAT(TLastRecQty[8]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 42, FORMAT(TLastRecQty[20]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 43, FORMAT(TLastMonthQuantity[9]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 44, FORMAT(TLastMonthQuantity[21]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 45, FORMAT(TLastRecQty[9]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 46, FORMAT(TLastRecQty[21]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 47, FORMAT(TLastMonthQuantity[10]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 48, FORMAT(TLastMonthQuantity[22]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 49, FORMAT(TLastRecQty[10]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 50, FORMAT(TLastRecQty[22]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 51, FORMAT(TLastMonthQuantity[11]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 52, FORMAT(TLastMonthQuantity[23]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 53, FORMAT(TLastRecQty[11]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 54, FORMAT(TLastRecQty[23]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 55, FORMAT(TLastMonthQuantity[12]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 56, FORMAT(TLastMonthQuantity[24]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 57, FORMAT(TLastRecQty[12]), TRUE, FALSE, '#####0.000', 0);
        EnterCell(NNN, 58, FORMAT(TLastRecQty[24]), TRUE, FALSE, '#####0.000', 0);

        //<<

    end;

    local procedure CreateExcelBook()
    begin
        //>>06Mar2017 RB-N
        /* ExcelBuf.CreateBook('', Text001);
        ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<
    end;

    ////[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        IF PrintToExcel THEN BEGIN
            ExcelBuf.INIT;
            ExcelBuf.VALIDATE("Row No.", Rowno);
            ExcelBuf.VALIDATE("Column No.", columnno);
            ExcelBuf."Cell Value as Text" := Cellvalue;
            ExcelBuf.Formula := '';
            ExcelBuf.Bold := Bold;
            ExcelBuf.Underline := Underline;
            ExcelBuf.NumberFormat := NoFormat;
            ExcelBuf."Cell Type" := CType;
            ExcelBuf.INSERT;
        END;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


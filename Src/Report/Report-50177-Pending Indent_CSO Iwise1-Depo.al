report 50177 "Pending Indent/CSO Iwise1-Depo"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 09Aug2017   RB-N         Removed filter Trading Location
    //                          LocationFilter Applied to Transfer-To Location(Transfer Indent Line)
    // 22Sep2017   RB-N         New Field Stock at Location in LTR
    // 22Dec2017   RB-N         New Filed Qty. in Ltrs
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PendingIndentCSOIwise1Depo.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING(Description);
            RequestFilterFields = "No.", "Date Filter", "Location Filter", "Item Category Code";
            dataitem("Sales Line"; "Sales Line")
            {
                DataItemLink = "No." = FIELD("No.");
                DataItemTableView = WHERE(Closed = FILTER(false),
                                          "Document Type" = FILTER(Order));

                trigger OnAfterGetRecord()
                begin

                    //RK-22-11-2017
                    CLEAR(SupplyName);
                    IF Loc_Rec.GET("Sales Line"."Location Code") THEN BEGIN
                        SupplyName := Loc_Rec.Name;
                    END;
                    //RK-22-11-2017

                    //>>RB-N 22Sep2017 Inventory at Location

                    IQty22 := 0;
                    Itm22.RESET;
                    Itm22.SETRANGE("No.", "No.");
                    Itm22.SETRANGE("Location Filter", "Sales Line"."Location Code");
                    Itm22.SETRANGE("Date Filter", 0D, maxdate);
                    IF Itm22.FINDFIRST THEN BEGIN
                        Itm22.CALCFIELDS(Inventory);

                        IQty22 := Itm22.Inventory;

                    END;
                    //>>RB-N 22Sep2017 Inventory at Location

                    //>>1

                    QuantityToSHipSale := 0;
                    CLEAR(docdate);
                    IF ShowUnApproved = FALSE THEN BEGIN //rSPL sACHIN
                        recSH.RESET;
                        recSH.SETRANGE(recSH."Short Close", FALSE);
                        recSH.SETRANGE(recSH.Status, recSH.Status::Released);
                        recSH.SETRANGE("Credit Limit Approval", recSH."Credit Limit Approval"::Approved);//RSPLSUM
                        recSH.SETRANGE(recSH."No.", "Sales Line"."Document No.");
                        IF NOT recSH.FINDFIRST THEN
                            CurrReport.SKIP;
                    END;


                    SalesHeader.RESET;
                    SalesHeader.SETRANGE(SalesHeader."No.", "Sales Line"."Document No.");
                    IF SalesHeader.FINDFIRST THEN
                        PostingDateforSales := SalesHeader."Posting Date";
                    docdate := SalesHeader."Document Date";
                    IF NOT ((docdate >= mindate) AND (docdate <= maxdate)) THEN
                        CurrReport.SKIP;


                    recSalesApprovalEntry.RESET;
                    recSalesApprovalEntry.SETRANGE(recSalesApprovalEntry."Document No.", "Sales Line"."Document No.");
                    IF recSalesApprovalEntry.FINDFIRST THEN;

                    QuantityToSHipSale := "Sales Line".Quantity - "Sales Line"."Quantity Shipped";

                    IF QuantityToSHipSale = 0 THEN
                        CurrReport.SKIP;

                    IF itemcategory.GET("Sales Line"."Item Category Code") THEN;

                    CLEAR(ProductGrade);
                    CLEAR(ProductGroup);
                    CLEAR(ProductGroupName);

                    recDefaultDimension.RESET;
                    recDefaultDimension.SETRANGE(recDefaultDimension."Table ID", 27);
                    recDefaultDimension.SETRANGE(recDefaultDimension."No.", "Sales Line"."No.");
                    recDefaultDimension.SETRANGE(recDefaultDimension."Dimension Code", 'FOCUS');
                    IF recDefaultDimension.FINDFIRST THEN BEGIN
                        ProductGrade := recDefaultDimension."Dimension Value Code";
                    END;

                    recDefaultDimension.RESET;
                    recDefaultDimension.SETRANGE(recDefaultDimension."Table ID", 27);
                    recDefaultDimension.SETRANGE(recDefaultDimension."No.", "Sales Line"."No.");
                    recDefaultDimension.SETFILTER(recDefaultDimension."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                    IF recDefaultDimension.FINDFIRST THEN BEGIN
                        ProductGroup := recDefaultDimension."Dimension Value Code";
                        recDimensionValue.RESET;
                        recDimensionValue.SETFILTER(recDimensionValue."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                        recDimensionValue.SETRANGE(recDimensionValue.Code, ProductGroup);
                        IF recDimensionValue.FINDFIRST THEN
                            ProductGroupName := recDimensionValue.Name;
                    END;

                    /*
                    CLEAR(BinName);
                    
                    CLEAR(BinQty);
                    recBinContent.RESET;
                    recBinContent.SETRANGE(recBinContent."Item No.","Sales Line"."No.");
                    recBinContent.SETFILTER(recBinContent."Location Code",'PLANT01');
                    //recBinContent.SETFILTER(recBinContent."Bin Code",'V-S-INDL');
                    recBinContent.SETFILTER(recBinContent."Bin Code",'%1|%2','V-S-INDL','V-S-AUTO');
                    recBinContent.SETFILTER(recBinContent.Quantity,'<>%1',0);
                    IF recBinContent.FIND('-') THEN
                    BEGIN
                    recBinContent.CALCFIELDS(recBinContent.Quantity);
                    BinQty := BinQty + recBinContent.Quantity;
                    
                      recBin.RESET;
                      //recBin.SETRANGE(recBin.Code,'V-S-INDL');
                      //recBin.SETFILTER(recBin.Code,'%1|%2','V-S-INDL','V-S-AUTO');
                      recBin.SETRANGE(recBin.Code,recBinContent."Bin Code");
                      IF recBin.FINDFIRST THEN
                      BinName := recBin.Description
                    END;
                    *///Commented on 29May2017


                    SalesComment := '';
                    recSalesComment.RESET;
                    recSalesComment.SETRANGE(recSalesComment."Document Type", recSalesComment."Document Type"::Order);
                    recSalesComment.SETRANGE(recSalesComment."No.", "Sales Line"."Document No.");
                    IF recSalesComment.FINDFIRST THEN BEGIN
                        SalesComment := recSalesComment.Comment;
                    END;

                    salc1 := 0;
                    salc2 := 0;
                    salc3 := 0;

                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.", "Sales Line"."No.");
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Posting Date", date1m, endd1);
                    IF recILE.FINDSET THEN
                        REPEAT
                            salc1 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.", "Sales Line"."No.");
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Posting Date", date2m, endd2);
                    IF recILE.FINDSET THEN
                        REPEAT
                            salc2 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.", "Sales Line"."No.");
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Posting Date", date3m, endd3);
                    IF recILE.FINDSET THEN
                        REPEAT
                            salc3 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    //>>Instransit Inventory to be Calculated
                    IntransitInv := 0;
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETFILTER(recILE."Remaining Quantity", '<>%1', 0);
                    recILE.SETFILTER(recILE."Posting Date", '%1..%2', mindate, maxdate);
                    recILE.SETRANGE(recILE."Entry Type", recILE."Entry Type"::Transfer);
                    recILE.SETFILTER(recILE."Location Code", '%1', 'IN-TRANS');
                    IF recILE.FINDSET THEN
                        REPEAT
                            IntransitInv += recILE."Remaining Quantity";
                        UNTIL recILE.NEXT = 0;
                    //END;
                    //<<1


                    //>>2

                    //Sales Line, Body (1) - OnPreSection()
                    SrNo := SrNo + 1;

                    IF "Sales Line"."Responsibility Center" <> '' THEN BEGIN
                        ResCenter.GET("Sales Line"."Responsibility Center");
                        ResponsibilityName := ResCenter.Name;
                    END
                    ELSE
                        ResponsibilityName := '';

                    itemqtytotal += Quantity;
                    totqtyship += "Quantity Shipped";
                    totqtytoship += QuantityToSHipSale;
                    TotalQtyBase += "Quantity (Base)";//RB-N 22Dec2017

                    CLEAR(CreditAppStatus);//RSPLSUM
                    CLEAR(vStatus);
                    RecSheader1.RESET;
                    CLEAR(PartyName);
                    RecSheader1.SETRANGE(RecSheader1."No.", "Sales Line"."Document No.");
                    IF RecSheader1.FINDFIRST THEN BEGIN
                        PartyName := RecSheader1."Sell-to Customer Name";
                        vStatus := RecSheader1."Campaign No."; //rspl sachin
                        CreditAppStatus := FORMAT(RecSheader1."Credit Limit Approval");//RSPLSUM
                    END;

                    //rSPL sACHIN
                    //RSPLSUM--IF vStatus<>'APPROVED' THEN
                    //RSPLSUM--vStatus:='NOT APPROVED';
                    //rSPL sACHIN
                    ItemNo := "Sales Line"."No.";
                    itemdesc := "Sales Line".Description;

                    CLEAR(UserName);
                    IF usersetup.GET(recSalesApprovalEntry."Approvar ID") THEN
                        UserName := usersetup.Name;

                    CulumativeQty := 0;
                    CulumativeQtyBase := 0;
                    recSIL.RESET;
                    recSIL.SETRANGE(recSIL."Sell-to Customer No.", "Sales Line"."Sell-to Customer No.");
                    recSIL.SETRANGE(recSIL."No.", "Sales Line"."No.");
                    recSIL.SETFILTER(recSIL."Posting Date", '%1..%2', mindate, maxdate);
                    recSIL.SETRANGE(recSIL."Location Code", "Sales Line"."Location Code");
                    IF recSIL.FINDFIRST THEN
                        REPEAT
                            CulumativeQty := CulumativeQty + recSIL.Quantity;
                            CulumativeQtyBase := CulumativeQtyBase + recSIL."Quantity (Base)";
                        UNTIL recSIL.NEXT = 0;

                    //>>Instransit Inventory to be Calculated
                    IntransitInv := 0;
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETFILTER(recILE."Remaining Quantity", '<>%1', 0);
                    recILE.SETFILTER(recILE."Posting Date", '%1..%2', mindate, maxdate);
                    recILE.SETRANGE(recILE."Entry Type", recILE."Entry Type"::Transfer);
                    recILE.SETFILTER(recILE."Location Code", '%1', 'IN-TRANS');
                    IF recILE.FINDSET THEN
                        REPEAT
                            IntransitInv += recILE."Remaining Quantity";
                        UNTIL recILE.NEXT = 0;
                    //END;

                    //RSPLSUM 31Aug2020>>
                    CLEAR(Zone);
                    RecSHNew.RESET;
                    RecSHNew.SETRANGE("No.", "Sales Line"."Document No.");
                    IF RecSHNew.FINDFIRST THEN BEGIN
                        RecSalesperPurchser.RESET;
                        IF RecSalesperPurchser.GET(RecSHNew."Salesperson Code") THEN
                            Zone := RecSalesperPurchser."Zone Code";
                    END;
                    //RSPLSUM 31Aug2020<<

                    IF (ExportToExcel) AND (NOT PrintSummary) THEN BEGIN
                        RowNo += 1;
                        EnterCell(RowNo, 1, FORMAT(SrNo), FALSE, FALSE, '', 0);
                        EnterCell(RowNo, 2, FORMAT("No."), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 3, FORMAT(Description), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 4, FORMAT(itemcategory.Description), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 5, FORMAT(ProductGroupName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 6, FORMAT(''), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 7, FORMAT(ProductGrade), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 8, FORMAT(ResponsibilityName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 9, FORMAT(docdate), FALSE, FALSE, '', 2);
                        EnterCell(RowNo, 10, FORMAT("Document No."), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 11, FORMAT(PartyName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 12, FORMAT(UserName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 13, FORMAT(recSalesApprovalEntry."Approval Date"), FALSE, FALSE, '', 2);
                        EnterCell(RowNo, 14, FORMAT(recSalesApprovalEntry."Approval Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        //EnterCell(RowNo,14,FORMAT(recSalesApprovalEntry."Approval Time"),FALSE,FALSE,'',3);
                        EnterCell(RowNo, 15, FORMAT("Unit of Measure"), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 16, FORMAT("Qty. per Unit of Measure"), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 17, FORMAT(Quantity), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 18, FORMAT("Quantity (Base)"), FALSE, FALSE, '0.000', 0);//RB-N 22Dec2017
                        EnterCell(RowNo, 19, FORMAT("Quantity Shipped"), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 20, FORMAT(QuantityToSHipSale), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 21, FORMAT("Document Type"), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 22, FORMAT(IQty22), FALSE, FALSE, '0.000', 0);//RB-N 22Sep2017
                                                                                       //EnterCell(RowNo,22,FORMAT("Document Type"),FALSE,FALSE,'',1);
                                                                                       //EnterCell(RowNo,23,FORMAT(BinQty/ItemUOM),FALSE,FALSE,'0.000',0);
                        EnterCell(RowNo, 23, FORMAT(SalesComment), FALSE, FALSE, '', 1);
                        //EnterCell(RowNo,24,FORMAT(SalesComment),FALSE,FALSE,'',1);
                        //Sales of last three months for customers

                        IF salc1 <> 0 THEN BEGIN
                            EnterCell(RowNo, 25, FORMAT(salc1 / "Qty. per Unit of Measure"), FALSE, FALSE, '0.000', 0); //29May2017
                        END ELSE
                            EnterCell(RowNo, 25, FORMAT(salc1), FALSE, FALSE, '0.000', 0);
                        //EnterCell(RowNo,23,FORMAT(salc1),FALSE,FALSE,'0.000',0);

                        IF salc2 <> 0 THEN BEGIN
                            EnterCell(RowNo, 26, FORMAT(salc2 / "Qty. per Unit of Measure"), FALSE, FALSE, '0.000', 0); //29May2017
                        END ELSE
                            EnterCell(RowNo, 26, FORMAT(salc2), FALSE, FALSE, '0.000', 0);
                        //EnterCell(RowNo,24,FORMAT(salc2),FALSE,FALSE,'0.000',0);

                        IF salc3 <> 0 THEN BEGIN
                            EnterCell(RowNo, 27, FORMAT(salc3 / "Qty. per Unit of Measure"), FALSE, FALSE, '0.000', 0); //29May2017
                        END ELSE
                            EnterCell(RowNo, 27, FORMAT(salc3), FALSE, FALSE, '0.000', 0);
                        //EnterCell(RowNo,25,FORMAT(salc3),FALSE,FALSE,'0.000',0);

                        //>>29May2017
                        IF IntransitInv <> 0 THEN BEGIN

                            EnterCell(RowNo, 28, FORMAT(IntransitInv / "Qty. per Unit of Measure"), FALSE, FALSE, '0.000', 0);

                        END ELSE
                            EnterCell(RowNo, 28, FORMAT(IntransitInv), FALSE, FALSE, '0.000', 0);
                        //>>29May2017

                        //EnterCell(RowNo,26,FORMAT(IntransitInv),FALSE,FALSE,'0.000',0);
                        EnterCell(RowNo, 29, FORMAT(vStatus), FALSE, FALSE, '', 1);//rspl sachin
                        EnterCell(RowNo, 30, FORMAT(SupplyName), FALSE, FALSE, '', 1);//rspl-Rk
                        EnterCell(RowNo, 31, FORMAT(Zone), FALSE, FALSE, '', 1);//RSPLSUM 31Aug2020
                        EnterCell(RowNo, 32, FORMAT(CreditAppStatus), FALSE, FALSE, '', 1);//RSPLSUM
                    END;
                    //<<2

                end;

                trigger OnPreDataItem()
                begin

                    ///
                    //>>1

                    IF locfilt <> '' THEN
                        "Sales Line".SETRANGE("Location Code", locfilt);

                    //To calc sales of last three month
                    DAY := DATE2DMY(mindate, 1);
                    MONTH := DATE2DMY(mindate, 2);
                    YEAR := DATE2DMY(mindate, 3);
                    MonthCaption := ReturnMonth(MONTH);
                    Startd := 1;
                    //For calculating start date and end date
                    CalDate := DMY2DATE(Startd, MONTH, YEAR);
                    Days := ReturnEndDate(MonthCaption);

                    //Start Date -1M
                    date1m := CALCDATE('-0M', CalDate);//start1
                    MONTH := DATE2DMY(date1m, 2);
                    ED1 := ReturnRespEndDate(MONTH);
                    endd1 := DMY2DATE(ED1, MONTH, YEAR);//end1
                    MonthCaption1 := ReturnMonth(MONTH);

                    //Start Date -2M
                    date2m := CALCDATE('-1M', CalDate);//start2
                    MONTH := DATE2DMY(date2m, 2);
                    ED2 := ReturnRespEndDate(MONTH);
                    IF MONTH > 10 THEN
                        endd2 := DMY2DATE(ED2, MONTH, YEAR - 1)//end2
                    ELSE
                        endd2 := DMY2DATE(ED2, MONTH, YEAR);//end2
                    MonthCaption2 := ReturnMonth(MONTH);

                    //Start Date -3M
                    date3m := CALCDATE('-2M', CalDate);//start3
                    MONTH := DATE2DMY(date3m, 2);

                    ED3 := ReturnRespEndDate(MONTH);
                    IF MONTH > 10 THEN
                        endd3 := DMY2DATE(ED3, MONTH, YEAR - 1)//end3
                    ELSE
                        endd3 := DMY2DATE(ED3, MONTH, YEAR);//end3

                    MonthCaption3 := ReturnMonth(MONTH);
                    //<<1
                end;
            }
            dataitem("Transfer Indent Line"; "Transfer Indent Line")
            {
                DataItemLink = "Item No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Closed = FILTER(false));

                trigger OnAfterGetRecord()
                begin

                    //>>RB-N 22Sep2017 Inventory at Location

                    IQty22 := 0;
                    Itm22.RESET;
                    Itm22.SETRANGE("No.", "Item No.");
                    Itm22.SETRANGE("Location Filter", "Transfer Indent Line"."Transfer-to Code");
                    Itm22.SETRANGE("Date Filter", 0D, maxdate);
                    IF Itm22.FINDFIRST THEN BEGIN
                        Itm22.CALCFIELDS(Inventory);

                        IQty22 := Itm22.Inventory;

                    END;
                    //>>RB-N 22Sep2017 Inventory at Location

                    //RK-22-11-2017
                    CLEAR(SupplyName);
                    TransferHeader.RESET;
                    TransferHeader.SETRANGE(TransferHeader."No.", "Transfer Indent Line"."Document No.");
                    IF TransferHeader.FINDFIRST THEN
                        IF Loc_Rec.GET(TransferHeader."Transfer-from Code") THEN BEGIN
                            SupplyName := Loc_Rec.Name;
                        END;
                    //RK-22-11-2017

                    //>>1

                    QuantityToSHipTransfer := 0;

                    TransferHeader.RESET;
                    TransferHeader.SETRANGE(TransferHeader."No.", "Transfer Indent Line"."Document No.");
                    IF TransferHeader.FINDFIRST THEN
                        PostingDateforTransfer := TransferHeader."Posting Date";

                    IF NOT ((PostingDateforTransfer >= mindate) AND (PostingDateforTransfer <= maxdate)) THEN
                        CurrReport.SKIP;

                    // EBT MILAN to Show short closed indent 290314
                    IF Showshortclosed = FALSE THEN BEGIN
                        IF TransferHeader."Short Closed" THEN
                            CurrReport.SKIP;
                    END;

                    recDefaultDimension.RESET;
                    recDefaultDimension.SETRANGE(recDefaultDimension."Table ID", 27);
                    recDefaultDimension.SETRANGE(recDefaultDimension."No.", "Transfer Indent Line"."Item No.");
                    recDefaultDimension.SETFILTER(recDefaultDimension."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                    IF recDefaultDimension.FINDFIRST THEN BEGIN
                        ProductGroup := recDefaultDimension."Dimension Value Code";
                        recDimensionValue.RESET;
                        recDimensionValue.SETFILTER(recDimensionValue."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                        recDimensionValue.SETRANGE(recDimensionValue.Code, ProductGroup);
                        IF recDimensionValue.FINDFIRST THEN
                            ProductGroupName := recDimensionValue.Name;
                    END;

                    CLEAR(QtyShip);
                    recTSL.RESET;
                    //recTSL.SETRANGE(recTSL."Transfer Indent No.", "Transfer Indent Line"."Document No.");
                    recTSL.SETRANGE(recTSL."Item No.", "Transfer Indent Line"."Item No.");
                    // recTSL.SETRANGE(recTSL."Transfer Indent Line No.", "Transfer Indent Line"."Line No.");
                    IF recTSL.FINDSET THEN
                        REPEAT
                            QtyShip := QtyShip + recTSL.Quantity;
                        UNTIL recTSL.NEXT = 0;

                    QuantityToSHipTransfer := "Transfer Indent Line".Quantity - QtyShip;

                    /*
                    CLEAR(BinName);
                    CLEAR(BinQty);
                    recBinContent.RESET;
                    recBinContent.SETRANGE(recBinContent."Item No.","Transfer Indent Line"."Item No.");
                    recBinContent.SETFILTER(recBinContent."Location Code",'PLANT01');
                    //recBinContent.SETFILTER(recBinContent."Bin Code",'V-S-INDL');
                    recBinContent.SETFILTER(recBinContent."Bin Code",'%1|%2','V-S-INDL','V-S-AUTO');
                    recBinContent.SETFILTER(recBinContent.Quantity,'<>%1',0);
                    IF recBinContent.FIND('-') THEN
                    BEGIN
                    recBinContent.CALCFIELDS(recBinContent.Quantity);
                    BinQty := BinQty + recBinContent.Quantity;
                    
                      recBin.RESET;
                      //recBin.SETRANGE(recBin.Code,'V-S-INDL');
                      //recBin.SETFILTER(recBin.Code,'%1|%2','V-S-INDL','V-S-AUTO');
                      recBin.SETRANGE(recBin.Code,recBinContent."Bin Code");
                      IF recBin.FINDFIRST THEN
                      BinName := recBin.Description
                    END;
                    *///Commented on 29May2017

                    CLEAR(QtyperUOM);
                    recItemUOMEBT.RESET;
                    recItemUOMEBT.SETRANGE(recItemUOMEBT."Item No.", "Transfer Indent Line"."Item No.");
                    recItemUOMEBT.SETRANGE(recItemUOMEBT.Code, "Transfer Indent Line"."Unit of Measure Code");
                    IF recItemUOMEBT.FINDFIRST THEN BEGIN
                        QtyperUOM := recItemUOMEBT."Qty. per Unit of Measure";
                    END;

                    CLEAR(itemCatforind);
                    RecItemCat.RESET;
                    RecItemCat.SETRANGE(RecItemCat.Code, Item."Item Category Code");
                    IF RecItemCat.FINDFIRST THEN
                        itemCatforind := RecItemCat.Description;

                    // (290314) Comented by EBT MILAN to show zero balance qty in indent in print summary option
                    /*
                    IF NOT PrintDetail THEN
                    BEGIN
                      IF QuantityToSHipTransfer = 0 THEN
                         CurrReport.SKIP;
                    END;
                    */
                    //RSPL-CAS-04663-V8J3G6
                    CLEAR(Zone);//RSPLSUM 31Aug2020
                    SalesPerson := '';
                    recSP.RESET;//RPSLSUM 31Aug2020
                    IF recSP.GET("Transfer Indent Line"."Salesperson Code") THEN BEGIN
                        SalesPerson := recSP.Name;
                        Zone := recSP."Zone Code";//RSPLSUM 31Aug2020
                    END;//RSPLSUM 31Aug2020
                        //RSPL-CAS-04663-V8J3G6

                    //RSPL-CAS-07044-N8Z9H8
                    Qty_sum := 0;
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETFILTER(recILE."Posting Date", '<=%1', maxdate);
                    IF recILE.FINDSET THEN
                        REPEAT
                            Qty_sum += recILE.Quantity;
                        UNTIL recILE.NEXT = 0;
                    qtyuom := 0;
                    recItemUOM.RESET;
                    recItemUOM.SETRANGE(recItemUOM."Item No.", TransferLine."Item No.");
                    recItemUOM.SETRANGE(recItemUOM.Code, TransferLine."Unit of Measure Code");
                    IF recItemUOM.FINDSET THEN
                        qtyuom := Qty_sum / recItemUOM."Qty. per Unit of Measure";
                    //RSPL-CAS-07044-N8Z9H8


                    q1 := 0;
                    q2 := 0;
                    q3 := 0;
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date1m, endd1);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q1 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date2m, endd2);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q2 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date3m, endd3);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q3 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;
                    /*
                    salc1:=0;
                    salc2:=0;
                    salc3:=0;
                    
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.","Sales Line"."No.");
                    recILE.SETRANGE(recILE."Location Code","Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type",recILE."Document Type"::"Transfer Shipment");
                    recILE.SETRANGE(recILE."Posting Date",date1m,endd1);
                    IF recILE.FINDSET THEN REPEAT
                      salc1+=ABS(recILE.Quantity);
                    UNTIL recILE.NEXT=0;
                    
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.","Sales Line"."No.");
                    recILE.SETRANGE(recILE."Location Code","Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type",recILE."Document Type"::"Transfer Shipment");
                    recILE.SETRANGE(recILE."Posting Date",date2m,endd2);
                    IF recILE.FINDSET THEN REPEAT
                      salc2+=ABS(recILE.Quantity);
                    UNTIL recILE.NEXT=0;
                    
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.","Sales Line"."No.");
                    recILE.SETRANGE(recILE."Location Code","Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type",recILE."Document Type"::"Transfer Shipment");
                    recILE.SETRANGE(recILE."Posting Date",date3m,endd3);
                    IF recILE.FINDSET THEN REPEAT
                      salc3+=ABS(recILE.Quantity);
                    UNTIL recILE.NEXT=0;
                     */

                    /*
                    //>>Instransit Inventory to be Calculated
                    IntransitInv:=0;
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.",Item."No.");
                    recILE.SETFILTER(recILE."Remaining Quantity",'<>%1',0);
                    recILE.SETFILTER(recILE."Posting Date",'%1..%2',mindate,maxdate);
                    recILE.SETRANGE(recILE."Entry Type",recILE."Entry Type"::Transfer);
                    recILE.SETFILTER(recILE."Location Code",'%1','IN-TRANS');
                    IF recILE.FINDSET THEN
                    REPEAT
                           IntransitInv+=recILE."Remaining Quantity";
                    UNTIL recILE.NEXT=0;
                    */
                    //END;
                    //<<1


                    //>>2
                    //Transfer Indent Line, Body (1) - OnPreSection()
                    SrNo := SrNo + 1;
                    TransferHeader.RESET;
                    TransferHeader.SETRANGE(TransferHeader."No.", "Transfer Indent Line"."Document No.");
                    IF TransferHeader.FINDFIRST THEN BEGIN
                        IF TransferHeader."Responsibility Center" <> '' THEN BEGIN
                            ResCenter.GET(TransferHeader."Responsibility Center");
                            ResponsibilityName := ResCenter.Name;
                        END
                        ELSE
                            ResponsibilityName := '';
                    END;

                    itemqtytotal += Quantity;
                    totqtyship += QtyShip;
                    totqtytoship += QuantityToSHipTransfer;
                    TotalQtyBase += Quantity * QtyperUOM;//RB-N 22Dec2017

                    itemdesc := "Transfer Indent Line".Description;
                    ItemNo := "Transfer Indent Line"."Item No.";
                    CLEAR(LocationName);
                    IF recLocation.GET(TransferHeader."Transfer-to Code") THEN
                        LocationName := recLocation.Name;

                    CLEAR(UserName);
                    IF usersetup.GET(TransferHeader."Approve User ID") THEN
                        UserName := usersetup.Name;

                    CulumativeQty := 0;
                    CulumativeQtyBase := 0;
                    recTSL.RESET;
                    recTSL.SETRANGE(recTSL."Transfer-from Code", "Transfer Indent Line"."Transfer-from Code");
                    recTSL.SETRANGE(recTSL."Item No.", "Transfer Indent Line"."Item No.");
                    recTSL.SETFILTER(recTSL."Shipment Date", '%1..%2', mindate, maxdate);
                    recTSL.SETRANGE(recTSL."Transfer-to Code", "Transfer Indent Line"."Transfer-to Code");
                    IF recTSL.FINDFIRST THEN
                        REPEAT
                            CulumativeQty := CulumativeQty + recTSL.Quantity;
                            CulumativeQtyBase := CulumativeQtyBase + recTSL."Quantity (Base)";
                        UNTIL recTSL.NEXT = 0;

                    /*
                    //>>Instransit Inventory to be Calculated
                    IntransitInv:=0;
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.",Item."No.");
                    recILE.SETFILTER(recILE."Remaining Quantity",'<>%1',0);
                    recILE.SETFILTER(recILE."Posting Date",'%1..%2',mindate,maxdate);
                    recILE.SETRANGE(recILE."Entry Type",recILE."Entry Type"::Transfer);
                    recILE.SETFILTER(recILE."Location Code",'%1','IN-TRANS');
                    IF recILE.FINDSET THEN
                    REPEAT
                           IntransitInv+=recILE."Remaining Quantity";
                    UNTIL recILE.NEXT=0;
                    //END;
                    */ //Commented 11Aug2017

                    //>>IntransitQty 11Aug2017
                    IntransitInv := 0;
                    TL11.RESET;
                    TL11.SETRANGE("Transfer-to Code", "Transfer Indent Line"."Transfer-to Code");
                    TL11.SETRANGE("Transfer-from Code", 'IN-TRANS');
                    TL11.SETRANGE("Item No.", Item."No.");
                    TL11.SETFILTER("Shipment Date", '<=%1', maxdate);
                    IF TL11.FINDSET THEN
                        REPEAT

                            IntransitInv += TL11.Quantity;

                        UNTIL TL11.NEXT = 0;
                    //<<IntransitQty 11Aug2017

                    IF ExportToExcel AND (NOT PrintSummary) THEN BEGIN
                        RowNo += 1;
                        EnterCell(RowNo, 1, FORMAT(SrNo), FALSE, FALSE, '', 0);
                        EnterCell(RowNo, 2, FORMAT("Item No."), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 3, FORMAT(Description), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 4, FORMAT(itemCatforind), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 5, FORMAT(ProductGroupName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 6, FORMAT(''), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 7, FORMAT(ProductGrade), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 8, FORMAT(ResponsibilityName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 9, FORMAT("Transfer Indent Line"."Addition date"), FALSE, FALSE, '', 2);
                        EnterCell(RowNo, 10, FORMAT("Document No."), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 11, FORMAT(LocationName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 12, FORMAT(UserName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 13, FORMAT("Transfer Indent Line"."Approval Date"), FALSE, FALSE, '', 2);

                        EnterCell(RowNo, 14, FORMAT("Transfer Indent Line"."Approval Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        //EnterCell(RowNo,14,FORMAT("Transfer Indent Line"."Approval Time"),FALSE,FALSE,'',3);
                        EnterCell(RowNo, 15, FORMAT("Unit of Measure Code"), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 16, FORMAT(QtyperUOM), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 17, FORMAT(Quantity), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 18, FORMAT(Quantity * QtyperUOM), FALSE, FALSE, '0.000', 0);//RB-N 22Dec2017
                        EnterCell(RowNo, 19, FORMAT(QtyShip), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 20, FORMAT(QuantityToSHipTransfer), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 21, FORMAT('Indent'), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 22, FORMAT(IQty22), FALSE, FALSE, '0.000', 0);//RB-N 22Sep2017
                                                                                       //EnterCell(RowNo,22,FORMAT('Indent'),FALSE,FALSE,'',1);
                                                                                       //EnterCell(RowNo,23,FORMAT(BinQty/ItemUOM),FALSE,FALSE,'0.000',0);
                                                                                       //RSPL-CAS-04663-V8J3G6
                        EnterCell(RowNo, 24, FORMAT(SalesPerson), FALSE, FALSE, '', 1);
                        //EnterCell(RowNo,25,FORMAT(SalesPerson),FALSE,FALSE,'',1);
                        //RSPL-CAS-04663-V8J3G6
                        //EnterCell(RowNo,27,FORMAT(qtyuom),FALSE,FALSE); //RSPL-CAS-07044-N8Z9H8 Rspl Sachin
                        //sales of last three months

                        //IF q1 <> 0 THEN // RSPLSID 29Sept2023 Temp Comment
                        IF QtyperUOM <> 0 THEN BEGIN
                            EnterCell(RowNo, 25, FORMAT(q1 / QtyperUOM), FALSE, FALSE, '0.000', 0); //29May2017
                        END ELSE
                            EnterCell(RowNo, 25, FORMAT(q1), FALSE, FALSE, '0.000', 0);
                        //EnterCell(RowNo,23,FORMAT(q1),FALSE,FALSE,'0.000',0);

                        //IF q2 <> 0 THEN // RSPLSID 29Sept2023 Temp Comment
                        IF QtyperUOM <> 0 THEN BEGIN
                            EnterCell(RowNo, 26, FORMAT(q2 / QtyperUOM), FALSE, FALSE, '0.000', 0); //29May2017
                        END ELSE
                            EnterCell(RowNo, 26, FORMAT(q2), FALSE, FALSE, '0.000', 0);
                        //EnterCell(RowNo,24,FORMAT(q2),FALSE,FALSE,'0.000',0);

                        //IF q3 <> 0 THEN // RSPLSID 29Sept2023 Temp Comment
                        IF QtyperUOM <> 0 THEN BEGIN
                            EnterCell(RowNo, 27, FORMAT(q3 / QtyperUOM), FALSE, FALSE, '0.000', 0); //29May2017
                        END ELSE
                            EnterCell(RowNo, 27, FORMAT(q3), FALSE, FALSE, '0.000', 0);
                        //EnterCell(RowNo,25,FORMAT(q3),FALSE,FALSE,'0.000',0);
                        //EnterCell(RowNo,31,FORMAT(salc1),FALSE,FALSE);
                        //EnterCell(RowNo,32,FORMAT(salc2),FALSE,FALSE);
                        //EnterCell(RowNo,33,FORMAT(salc3),FALSE,FALSE);

                        //>>29May2017
                        IF IntransitInv <> 0 THEN BEGIN
                            /*
                              Itm29UOM.RESET;
                              Itm29UOM.SETRANGE("Item No.","Transfer Indent Line"."Item No.");
                              Itm29UOM.SETRANGE(Code,"Transfer Indent Line"."Unit of Measure Code");
                              IF Itm29UOM.FINDFIRST THEN

                            */
                            IF QtyperUOM <> 0 THEN  //AR-281218
                                EnterCell(RowNo, 28, FORMAT(IntransitInv / QtyperUOM), FALSE, FALSE, '0.000', 0)
                            ELSE
                                EnterCell(RowNo, 28, '', FALSE, FALSE, '', 1);

                        END ELSE
                            EnterCell(RowNo, 28, FORMAT(IntransitInv), FALSE, FALSE, '0.000', 0);
                        //>>29May2017
                        //>>Parag
                        //EnterCell(RowNo,26,FORMAT(IntransitInv),FALSE,FALSE,'0.000',0);
                        //<<Parag
                        EnterCell(RowNo, 29, FORMAT(Status), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 30, FORMAT(SupplyName), FALSE, FALSE, '', 1); //rspl-RK
                        EnterCell(RowNo, 31, FORMAT(Zone), FALSE, FALSE, '', 1);//RSPLSUM 31Aug2020
                        EnterCell(RowNo, 32, '', FALSE, FALSE, '', 1);//RSPLSUM
                    END;
                    //<<2

                end;

                trigger OnPreDataItem()
                begin


                    //>>1

                    IF locfilt <> '' THEN
                        "Transfer Indent Line".SETRANGE("Transfer Indent Line"."Transfer-to Code", locfilt); //09Aug2017 Change the FilterLoction

                    //"Transfer Indent Line".SETRANGE("Transfer Indent Line"."Transfer-from Code",locfilt);
                    "Transfer Indent Line".SETRANGE("Transfer Indent Line"."Addition date", mindate, maxdate);
                    //To calc sales of last three month
                    DAY := DATE2DMY(mindate, 1);
                    MONTH := DATE2DMY(mindate, 2);
                    YEAR := DATE2DMY(mindate, 3);
                    MonthCaption := ReturnMonth(MONTH);
                    Startd := 1;
                    //For calculating start date and end date
                    CalDate := DMY2DATE(Startd, MONTH, YEAR);
                    Days := ReturnEndDate(MonthCaption);

                    //Start Date -1M
                    date1m := CALCDATE('-0M', CalDate);//start1
                    MONTH := DATE2DMY(date1m, 2);
                    ED1 := ReturnRespEndDate(MONTH);
                    endd1 := DMY2DATE(ED1, MONTH, YEAR);//end1
                    MonthCaption1 := ReturnMonth(MONTH);

                    //Start Date -2M
                    date2m := CALCDATE('-1M', CalDate);//start2
                    MONTH := DATE2DMY(date2m, 2);
                    ED2 := ReturnRespEndDate(MONTH);
                    endd2 := DMY2DATE(ED2, MONTH, YEAR);//end2
                    MonthCaption2 := ReturnMonth(MONTH);

                    //Start Date -3M
                    date3m := CALCDATE('-2M', CalDate);//start3
                    MONTH := DATE2DMY(date3m, 2);
                    ED3 := ReturnRespEndDate(MONTH);
                    endd3 := DMY2DATE(ED3, MONTH, YEAR);//end3
                    MonthCaption3 := ReturnMonth(MONTH);
                    //Rspl Sachin
                    IF ShowUnApproved = FALSE THEN
                        "Transfer Indent Line".SETFILTER("Transfer Indent Line".Approve, '%1', TRUE);
                    //<<1
                end;
            }
            dataitem(Integer; Integer)
            {
                DataItemTableView = SORTING(Number)
                                    WHERE(Number = CONST(1));

                trigger OnAfterGetRecord()
                begin

                    //>>1
                    //Integer, Body (1) - OnPreSection()

                    //IF itemqtytotal=0 THEN
                    //CurrReport.SHOWOUTPUT(FALSE);//04May2017
                    IF itemqtytotal <> 0 THEN BEGIN

                        IF ExportToExcel AND (SalesExist OR TransferExists) THEN //04May2017
                                                                                 //IF ExportToExcel AND (CurrReport.SHOWOUTPUT) AND (SalesExist OR TransferExists)   THEN
                        BEGIN
                            RowNo += 1;
                            EnterCell(RowNo, 2, FORMAT(ItemNo), TRUE, FALSE, '', 1);
                            EnterCell(RowNo, 3, FORMAT(itemdesc), TRUE, FALSE, '', 1);
                            EnterCell(RowNo, 16, FORMAT(ItemUOM), TRUE, FALSE, '0.000', 0);
                            EnterCell(RowNo, 17, FORMAT(itemqtytotal), TRUE, FALSE, '0.000', 0);
                            EnterCell(RowNo, 18, FORMAT(TotalQtyBase), TRUE, FALSE, '0.000', 0);//RB-N 22Dec2017
                            EnterCell(RowNo, 19, FORMAT(totqtyship), TRUE, FALSE, '0.000', 0);
                            EnterCell(RowNo, 20, FORMAT(totqtytoship), TRUE, FALSE, '0.000', 0);
                            //EnterCell(RowNo,20,FORMAT(BinQty/ItemUOM),TRUE,FALSE,'0.000',0);
                            //EnterCell(RowNo,21,FORMAT((BinQty/ItemUOM)-totqtytoship),TRUE,FALSE,'0.000',0);
                            //EnterCell(RowNo,22,FORMAT(BinName),TRUE,FALSE);//Rspl Sachin
                        END;
                    END;

                    printheader := FALSE;
                    //<<1
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                CLEAR(SrNo);
                CLEAR(itemqtytotal);
                CLEAR(totqtyship);
                CLEAR(totqtytoship);
                CLEAR(BINCode);
                CLEAR(BinName);
                CLEAR(TotalQtyBase);//RB-N 22Dec2017

                SalesLine.RESET;
                SalesLine.SETCURRENTKEY(SalesLine."No.");
                SalesLine.SETRANGE(SalesLine."No.", Item."No.");
                IF SalesLine.FINDFIRST THEN
                    SalesExist := TRUE;

                TransferLine.RESET;
                TransferLine.SETCURRENTKEY(TransferLine."Item No.");
                TransferLine.SETRANGE(TransferLine."Item No.", Item."No.");
                IF TransferLine.FINDFIRST THEN
                    TransferExists := TRUE;

                IF (SalesExist = FALSE) AND (TransferExists = FALSE) THEN
                    CurrReport.SKIP;
                //>>Instransit Inventory to be Calculated
                IntransitInv := 0;
                recILE.RESET;
                recILE.SETRANGE(recILE."Item No.", Item."No.");
                recILE.SETFILTER(recILE."Remaining Quantity", '<>%1', 0);
                recILE.SETFILTER(recILE."Posting Date", '%1..%2', mindate, maxdate);
                recILE.SETRANGE(recILE."Entry Type", recILE."Entry Type"::Transfer);
                recILE.SETFILTER(recILE."Location Code", '%1', 'IN-TRANS');
                IF recILE.FINDSET THEN
                    REPEAT
                        IntransitInv += recILE."Remaining Quantity";
                    UNTIL recILE.NEXT = 0;
                //END;
                //<<1


                //>>2

                //Item, Body (3) - OnPreSection()
                recitem.RESET;
                recitem.SETRANGE(recitem."No.", "No.");
                IF recitem.FIND('-') THEN BEGIN
                    recitem.SETFILTER(recitem."Location Filter", Item.GETFILTER(Item."Location Filter"));
                    recitem.CALCFIELDS(recitem.Inventory);
                END;

                CLEAR(itemdesc);

                recItemUOM.RESET;
                recItemUOM.SETRANGE(recItemUOM."Item No.", "No.");
                IF recItemUOM.FINDFIRST THEN
                    ItemUOM := recItemUOM."Qty. per Unit of Measure";
                //<<2
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                mindate := Item.GETRANGEMIN("Date Filter");
                maxdate := Item.GETRANGEMAX("Date Filter");
                //To calc sales of last three month
                DAY := DATE2DMY(mindate, 1);
                MONTH := DATE2DMY(mindate, 2);
                YEAR := DATE2DMY(mindate, 3);
                MonthCaption := ReturnMonth(MONTH);
                Startd := 1;
                //For calculating start date and end date
                CalDate := DMY2DATE(Startd, MONTH, YEAR);
                Days := ReturnEndDate(MonthCaption);

                //Start Date -1M
                date1m := CALCDATE('-0M', CalDate);//start1
                MONTH := DATE2DMY(date1m, 2);
                ED1 := ReturnRespEndDate(MONTH);
                endd1 := DMY2DATE(ED1, MONTH, YEAR);//end1
                MonthCaption1 := ReturnMonth(MONTH);

                //Start Date -2M
                date2m := CALCDATE('-1M', CalDate);//start2
                MONTH := DATE2DMY(date2m, 2);
                ED2 := ReturnRespEndDate(MONTH);
                endd2 := DMY2DATE(ED2, MONTH, YEAR);//end2
                MonthCaption2 := ReturnMonth(MONTH);

                //Start Date -3M
                date3m := CALCDATE('-2M', CalDate);//start3
                MONTH := DATE2DMY(date3m, 2);
                ED3 := ReturnRespEndDate(MONTH);
                endd3 := DMY2DATE(ED3, MONTH, YEAR);//end3
                MonthCaption3 := ReturnMonth(MONTH);
                //<<1


                //>>2
                //Item, Header (2) - OnPreSection()
                IF ExportToExcel AND printheader THEN BEGIN
                    RowNo += 1;

                    //>>04May2017
                    EnterCell(RowNo, 1, COMPANYNAME, TRUE, FALSE, '', 1);
                    EnterCell(RowNo, 32, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);//29May2017
                                                                                  //EnterCell(RowNo,30,FORMAT(TODAY,0,4),TRUE,FALSE,'',1);//Commented on 29May2017
                    RowNo += 1;

                    EnterCell(RowNo, 1, 'Pending Indent CSO Item Wise Depo', TRUE, FALSE, '', 1);
                    EnterCell(RowNo, 32, USERID, TRUE, FALSE, '', 1);//29May2017
                                                                     //EnterCell(RowNo,30,USERID,TRUE,FALSE,'',1);//Commented on 29May2017
                    RowNo += 2;
                    //<<04may2017
                    EnterCell(RowNo, 1, FORMAT('Sr.No.'), TRUE, TRUE, '', 1);//04May2017
                                                                             //EnterCell(RowNo,1,FORMAT('Sr.No.'),TRUE,FALSE,'',1);
                    EnterCell(RowNo, 2, FORMAT('Item No.'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 3, FORMAT('Item Description'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 4, FORMAT('Item Category'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 5, FORMAT('Product Group'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 6, FORMAT('Product Sub Group'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 7, FORMAT('Product Grade'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 8, FORMAT('Branch'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 9, FORMAT('Doc./Addition date'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 10, FORMAT('Document No.'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 11, FORMAT('Party/ Location Name'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 12, FORMAT('Approver Name'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 13, FORMAT('Approve Date'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 14, FORMAT('Approve time'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 15, FORMAT('Pack UOM'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 16, FORMAT('Qty per UOM'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 17, FORMAT('Qty. Order as per pack UOM'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 18, FORMAT('Qty. in Ltrs'), TRUE, TRUE, '', 1);//RB-N 22Dec2017
                    EnterCell(RowNo, 19, FORMAT('Despatched per pack UOM'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 20, FORMAT('Balance  as per pack UOM'), TRUE, TRUE, '', 1);
                    //EnterCell(RowNo,20,FORMAT('Product Availability'),TRUE,TRUE,'',1); //Commented on 29May2017
                    //EnterCell(RowNo,21,FORMAT('Excess/Short for Production'),TRUE,TRUE,'',1);//Commented on 29May2017
                    //EnterCell(RowNo,22,FORMAT('Bin'),TRUE,FALSE);Rspl Sachin
                    EnterCell(RowNo, 21, FORMAT('Type'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 22, FORMAT('Stock at Location in LTR'), TRUE, TRUE, '', 1);//RB-N 22Sep2017
                                                                                                //EnterCell(RowNo,23,FORMAT('Stock at from Location in Sales UOM'),TRUE,TRUE,'',1); //Commented on 29May2017
                    EnterCell(RowNo, 23, FORMAT('Remarks'), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 24, FORMAT('Salesperson'), TRUE, TRUE, '', 1);
                    //EnterCell(RowNo,27,FORMAT('Stock at Indented Location in Sales UOM'),TRUE,FALSE);Rspl Sachin
                    EnterCell(RowNo, 25, FORMAT('Sales for the Month ' + FORMAT(MonthCaption1)), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 26, FORMAT('Sales for the Month ' + FORMAT(MonthCaption2)), TRUE, TRUE, '', 1);
                    EnterCell(RowNo, 27, FORMAT('Sales for the Month ' + FORMAT(MonthCaption3)), TRUE, TRUE, '', 1);
                    //EnterCell(RowNo,31,FORMAT('Transfer for the Month '+FORMAT(MonthCaption1)),TRUE,FALSE);Rspl Sachin
                    //EnterCell(RowNo,32,FORMAT('Transfer for the Month '+FORMAT(MonthCaption2)),TRUE,FALSE);
                    //EnterCell(RowNo,33,FORMAT('Transfer for the Month '+FORMAT(MonthCaption3)),TRUE,FALSE);
                    //>>Parag
                    EnterCell(RowNo, 28, FORMAT('Intransit Inventory'), TRUE, TRUE, '', 1);
                    //<<Parag
                    //Rspl Sachin
                    //RSPLSUM--EnterCell(RowNo,29,FORMAT('CSO Status'),TRUE,TRUE,'',1);
                    EnterCell(RowNo, 29, FORMAT('L1 Approval Status'), TRUE, TRUE, '', 1);//RSPLSUM
                                                                                          //Rspl Sachin
                    EnterCell(RowNo, 30, FORMAT('Supply Name'), TRUE, TRUE, '', 1); //rspl-RK
                    EnterCell(RowNo, 31, FORMAT('Zone'), TRUE, TRUE, '', 1);//RSPLSUM 31Aug2020
                    EnterCell(RowNo, 32, FORMAT('Credit Approval Status'), TRUE, TRUE, '', 1);//RSPLSUM
                END;


                //<<2
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print To Excel"; ExportToExcel)
                {
                    ApplicationArea = all;
                }
                field("Print Details"; PrintDetail)
                {
                    ApplicationArea = all;
                }
                field("Print Summary"; PrintSummary)
                {
                    ApplicationArea = all;
                }
                field("Show Short Closed"; Showshortclosed)
                {
                    ApplicationArea = all;
                }
                field("Show UnApproved CSO"; ShowUnApproved)
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

        //>>1
        IF ExportToExcel THEN BEGIN

            /* TempExcelBuffer1.CreateBook('', 'Pending Indent CSO Item Wise Depo');//04May2017
                                                                                 //TempExcelBuffer1.CreateBook;
            TempExcelBuffer1.CreateBookAndOpenExcel('', 'Pending Indent CSO Item Wise Depo', '', '', USERID);//04May2017
                                                                                                             //TempExcelBuffer1.CreateSheet('Pending Indent CSO','', COMPANYNAME, USERID);
            TempExcelBuffer1.GiveUserControl; */

            TempExcelBuffer1.CreateNewBook(Text001);
            TempExcelBuffer1.WriteSheet(Text001, CompanyName, UserId);
            TempExcelBuffer1.CloseBook();
            TempExcelBuffer1.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
            TempExcelBuffer1.OpenExcel();
        END;
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1

        SrNo := 0;
        printheader := TRUE;

        SalesExist := FALSE;
        TransferExists := FALSE;
        datefilter := Item.GETFILTER(Item."Date Filter");
        locfilt := Item.GETFILTER(Item."Location Filter");

        /*
        //EBT STIVAN ---(07112012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        LocCode := Item.GETFILTER(Item."Location Filter");
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
        //EBT STIVAN ---(07112012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        *///Commented 04May2017

        //<<1

    end;

    var
        SrNo: Integer;
        SalesHeader: Record 36;
        SalesLine: Record 37;
        PostingDateforSales: Date;
        PartyName: Text[150];
        ItemNameSale: Text[30];
        QuantityToSHipSale: Decimal;
        TransferHeader: Record 50022;
        TransferLine: Record 50023;
        PostingDateforTransfer: Date;
        ItemNameTransfer: Text[30];
        QuantityToSHipTransfer: Decimal;
        ResCenter: Record 5714;
        ResponsibilityName: Text;
        ExportToExcel: Boolean;
        SalesExist: Boolean;
        TransferExists: Boolean;
        TempExcelBuffer1: Record 370 temporary;
        RowNo: Integer;
        PrintToExcel: Boolean;
        itemcategory: Record 5722;
        itemqtytotal: Decimal;
        totqtyship: Decimal;
        totqtytoship: Decimal;
        datefilter: Text[100];
        mindate: Date;
        maxdate: Date;
        printheader: Boolean;
        locfilt: Code[100];
        RecSheader: Record 36;
        recitem: Record 27;
        RecSheader1: Record 36;
        docdate: Date;
        RecUserSetup1: Record 91;
        SalesHeader1: Record 36;
        TransferHeader1: Record 5740;
        sclosed: Boolean;
        itemdesc: Text[150];
        usersetup: Record 91;
        ItemNo: Code[20];
        recLocation: Record 14;
        LocationName: Text;
        UserName: Text;
        recDefaultDimension: Record 352;
        ProductGrade: Text;
        ProductGroup: Text;
        ProductGroupName: Text[50];
        recDimensionValue: Record 349;
        recItemUOM: Record 5404;
        ItemUOM: Decimal;
        recSalesApprovalEntry: Record 50009;
        recSH: Record 36;
        PendingQTYtoShip: Decimal;
        recTL: Record 5741;
        TotalPendingQTYtoShip: Decimal;
        TSno: Code[20];
        recTSH: Record 5744;
        recTSL: Record 5745;
        QtyShip: Decimal;
        recILE: Record 32;
        recItemApplEntry: Record 339;
        recWarehouseEntry: Record 7312;
        BINCode: Code[10];
        recBin: Record 7354;
        BinName: Text[50];
        recBinContent: Record 7302;
        Qty: Decimal;
        LocCode: Code[20];
        User: Record 2000000120;
        LocResC: Code[20];
        CSOMapping1: Record 50006;
        recItemUOMEBT: Record 5404;
        QtyperUOM: Decimal;
        BinQty: Decimal;
        recSIL: Record 113;
        CulumativeQty: Decimal;
        CulumativeQtyBase: Decimal;
        PrintDetail: Boolean;
        recSalesComment: Record 44;
        SalesComment: Text[100];
        PrintSummary: Boolean;
        RecItemCat: Record 5722;
        itemCatforind: Text[100];
        Showshortclosed: Boolean;
        recSP: Record 13;
        SalesPerson: Text[50];
        Qty_sum: Decimal;
        qtyuom: Decimal;
        recItemLedger: Record 32;
        recSalesLine: Record 37;
        recItemUnitOfMEasure: Record 5404;
        DAY: Integer;
        MONTH: Integer;
        YEAR: Integer;
        Startd: Integer;
        MonthCaption: Text[30];
        Days: Integer;
        DateText: Text[30];
        CalDate: Date;
        date1m: Date;
        date2m: Date;
        date3m: Date;
        ED1: Integer;
        ED2: Integer;
        ED3: Integer;
        endd1: Date;
        endd2: Date;
        endd3: Date;
        MonthCaption1: Text[30];
        MonthCaption2: Text[30];
        MonthCaption3: Text[30];
        q1: Decimal;
        q2: Decimal;
        q3: Decimal;
        salc1: Decimal;
        salc2: Decimal;
        salc3: Decimal;
        IntransitInv: Decimal;
        recIntransit: Record 32;
        recTRO: Record 5740;
        RecValueEntry: Record 5802;
        "-------Rspl Sachin-------": Integer;
        ShowUnApproved: Boolean;
        vStatus: Code[30];
        Text000: Label 'Data';
        Text001: Label 'Pending Indent CSO Itemwise Depo';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "----29May2017": Integer;
        Itm29UOM: Record 5404;
        "----11Aug2017": Integer;
        TL11: Record 5741;
        "-----22Sep2017": Integer;
        Itm22: Record 27;
        IQty22: Decimal;
        TotalQtyBase: Decimal;
        "-------------------------RK------------------------": Integer;
        Loc_Rec: Record 14;
        SupplyName: Text;
        RecSHNew: Record 36;
        RecSalesperPurchser: Record 13;
        Zone: Code[10];
        CreditAppStatus: Text;

    // //[Scope('Internal')]
    procedure EnterCell(RowNo: Integer; ColumnNo: Integer; CellValue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin
        TempExcelBuffer1.INIT;
        TempExcelBuffer1.VALIDATE("Row No.", RowNo);
        TempExcelBuffer1.VALIDATE("Column No.", ColumnNo);
        TempExcelBuffer1."Cell Value as Text" := CellValue;
        TempExcelBuffer1.Formula := '';
        TempExcelBuffer1.Bold := Bold;
        TempExcelBuffer1.Underline := Underline;
        TempExcelBuffer1.NumberFormat := NoFormat;//04May2017
        TempExcelBuffer1."Cell Type" := CType;//04May2017
        TempExcelBuffer1.INSERT;
    end;

    // //[Scope('Internal')]
    procedure ReturnMonth(int: Integer): Text[30]
    begin
        IF int = 1 THEN
            EXIT('January');
        IF int = 2 THEN
            EXIT('February');
        IF int = 3 THEN
            EXIT('March');
        IF int = 4 THEN
            EXIT('April');
        IF int = 5 THEN
            EXIT('May');
        IF int = 6 THEN
            EXIT('June');
        IF int = 7 THEN
            EXIT('July');
        IF int = 8 THEN
            EXIT('August');
        IF int = 9 THEN
            EXIT('September');
        IF int = 10 THEN
            EXIT('October');
        IF int = 11 THEN
            EXIT('November');
        IF int = 12 THEN
            EXIT('December');
    end;

    // //[Scope('Internal')]
    procedure ReturnEndDate(Text: Text[30]): Integer
    begin
        IF Text = 'January' THEN
            EXIT(31);
        IF Text = 'February' THEN
            EXIT(29);
        IF Text = 'March' THEN
            EXIT(31);
        IF Text = 'April' THEN
            EXIT(30);
        IF Text = 'May' THEN
            EXIT(31);
        IF Text = 'June' THEN
            EXIT(30);
        IF Text = 'July' THEN
            EXIT(31);
        IF Text = 'August' THEN
            EXIT(31);
        IF Text = 'September' THEN
            EXIT(30);
        IF Text = 'October' THEN
            EXIT(31);
        IF Text = 'November' THEN
            EXIT(30);
        IF Text = 'December' THEN
            EXIT(31);
    end;

    //  //[Scope('Internal')]
    procedure ReturnRespEndDate(int: Integer): Integer
    begin
        IF int = 1 THEN
            EXIT(31);
        IF int = 2 THEN
            EXIT(28);
        IF int = 3 THEN
            EXIT(31);
        IF int = 4 THEN
            EXIT(30);
        IF int = 5 THEN
            EXIT(31);
        IF int = 6 THEN
            EXIT(30);
        IF int = 7 THEN
            EXIT(31);
        IF int = 8 THEN
            EXIT(31);
        IF int = 9 THEN
            EXIT(30);
        IF int = 10 THEN
            EXIT(31);
        IF int = 11 THEN
            EXIT(30);
        IF int = 12 THEN
            EXIT(31);
    end;
}


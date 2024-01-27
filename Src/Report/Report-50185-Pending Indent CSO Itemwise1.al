report 50185 "Pending Indent/CSO Itemwise1"
{
    // 
    // RSPL-CAS-04663-V8J3G6       Sourav Dey       Code added to print Salesperson Description
    // RSPL-CAS-07044-N8Z9H8       Nikhil           Code added to convert the sum of Qty into UOM locationwise
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 08Aug2017   RB-N         ResponsibilityCenter & SalesPerson Code Filter Added
    // 07Sep2017   RB-N         Displaying Sales Comment Lines w.r.t DocumentNo.
    // 31Oct2017   RB-N         Last 3 months sales details of Previous years,Remarks column shifted to last,Transfer qty details of last 3 months
    // 18Dec2017   RB-N         Supply Location Name
    // 22Dec2017   RB-N         New Field Qty. in Ltrs
    // 26Dec2017   RB-N         New Field Qty. Produced in UOM
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PendingIndentCSOItemwise1.rdl';
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

                    //>>08Aug2017
                    IF SaleCode08 <> '' THEN BEGIN

                        SH08.RESET;
                        SH08.SETRANGE("No.", "Document No.");
                        SH08.SETRANGE("Salesperson Code", SaleCode08);
                        IF NOT SH08.FINDFIRST THEN
                            CurrReport.SKIP;

                    END;

                    //SalesPersonName
                    CLEAR(Zone);//RSPLSUM 01Sept2020
                    SalesPerson := '';
                    SH08.RESET;
                    SH08.SETRANGE("No.", "Document No.");
                    IF SH08.FINDFIRST THEN BEGIN
                        SalesPerson := '';
                        IF recSP.GET(SH08."Salesperson Code") THEN BEGIN
                            SalesPerson := recSP.Name;
                            Zone := recSP."Zone Code";//RSPLSUM 01Sept2020
                        END;
                    END;




                    //<<08Aug2017
                    QuantityToSHipSale := 0;
                    CLEAR(docdate);
                    //RSPL-Rahul

                    recSH.RESET;
                    recSH.SETRANGE(recSH."Short Close", FALSE);
                    recSH.SETRANGE(recSH.Status, recSH.Status::Released);
                    recSH.SETRANGE(recSH."No.", "Sales Line"."Document No.");
                    IF NOT recSH.FINDFIRST THEN
                        CurrReport.SKIP;

                    //<<RSPL-Rahul*****

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

                    //<<RSPL-Rahul

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

                    CLEAR(BinName);
                    CLEAR(BinQty);
                    recBinContent.RESET;
                    recBinContent.SETRANGE(recBinContent."Item No.", "Sales Line"."No.");
                    recBinContent.SETFILTER(recBinContent."Location Code", 'PLANT01');
                    //recBinContent.SETFILTER(recBinContent."Bin Code",'V-S-INDL');
                    recBinContent.SETFILTER(recBinContent."Bin Code", '%1|%2', 'V-S-INDL', 'V-S-AUTO');
                    recBinContent.SETFILTER(recBinContent.Quantity, '<>%1', 0);
                    IF recBinContent.FIND('-') THEN BEGIN
                        recBinContent.CALCFIELDS(recBinContent.Quantity);
                        BinQty := BinQty + recBinContent.Quantity;

                        recBin.RESET;
                        //recBin.SETRANGE(recBin.Code,'V-S-INDL');
                        //recBin.SETFILTER(recBin.Code,'%1|%2','V-S-INDL','V-S-AUTO');
                        recBin.SETRANGE(recBin.Code, recBinContent."Bin Code");
                        IF recBin.FINDFIRST THEN
                            BinName := recBin.Description
                    END;

                    /*
                    SalesComment := '';
                    recSalesComment.RESET;
                    recSalesComment.SETRANGE(recSalesComment."Document Type",recSalesComment."Document Type"::Order);
                    recSalesComment.SETRANGE(recSalesComment."No.","Sales Line"."Document No.");
                    IF recSalesComment.FINDFIRST THEN
                    BEGIN
                     SalesComment := recSalesComment.Comment;
                    END;
                    */

                    //>>07Sep2017 RB-N

                    CLEAR(CommentText);
                    C07 := 0;
                    recSalesComment.RESET;
                    recSalesComment.SETRANGE(recSalesComment."Document Type", recSalesComment."Document Type"::Order);
                    recSalesComment.SETRANGE(recSalesComment."No.", "Sales Line"."Document No.");
                    IF recSalesComment.FINDSET THEN
                        REPEAT
                            C07 += 1;

                            IF C07 = 1 THEN BEGIN
                                CommentText := recSalesComment.Comment;
                            END;

                            IF C07 > 1 THEN BEGIN
                                CommentText += ' , ' + recSalesComment.Comment;
                            END;

                        UNTIL recSalesComment.NEXT = 0;

                    CommLen := 0;
                    CommLen := STRLEN(CommentText);

                    CLEAR(Comm1);
                    CLEAR(Comm2);
                    CLEAR(Comm3);
                    CLEAR(Comm4);
                    Comm1 := COPYSTR(CommentText, 1, 250);
                    Comm2 := COPYSTR(CommentText, 251, 250);
                    Comm3 := COPYSTR(CommentText, 500, 250);
                    Comm4 := COPYSTR(CommentText, 751, 250);
                    //<<07Sep2017 RB-N

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

                    //TotQtyPerUOM += "Sales Line"."Qty. per Unit of Measure";//RSPLSUM 24Feb21  //Nb commented as total not needed in report
                    TotQtyPerUOM := "Sales Line"."Qty. per Unit of Measure";//RSPLSUM 24Feb21  //Nb added

                    //FOC Fahim 20-11-21
                    IF "Sales Line"."Free of Cost" = TRUE THEN
                        FOC := true
                    ELSE
                        FOC := false;

                    RecSheader1.RESET;
                    CLEAR(PartyName);
                    RecSheader1.SETRANGE(RecSheader1."No.", "Sales Line"."Document No.");
                    IF RecSheader1.FINDFIRST THEN
                        PartyName := RecSheader1."Sell-to Customer Name";
                    ItemNo := "Sales Line"."No.";
                    itemdesc := "Sales Line".Description;

                    CLEAR(UserName);
                    IF usersetup.GET(recSalesApprovalEntry."Approvar ID") THEN
                        UserName := usersetup.Name;

                    CulumativeQty := 0;
                    CulumativeQtyBase := 0;
                    recSIL.RESET;
                    recSIL.SETCURRENTKEY("No.");//31Oct2017
                    recSIL.SETRANGE(recSIL."Sell-to Customer No.", "Sales Line"."Sell-to Customer No.");
                    recSIL.SETRANGE(recSIL."No.", "Sales Line"."No.");
                    recSIL.SETFILTER(recSIL."Posting Date", '%1..%2', mindate, maxdate);
                    recSIL.SETRANGE(recSIL."Location Code", "Sales Line"."Location Code");
                    IF recSIL.FINDFIRST THEN
                        REPEAT
                            CulumativeQty := CulumativeQty + recSIL.Quantity;
                            CulumativeQtyBase := CulumativeQtyBase + recSIL."Quantity (Base)";
                        UNTIL recSIL.NEXT = 0;

                    //Ragni----
                    DocDimValue := '';
                    DimDesc := '';
                    ItemDescription := '';

                    //>>11May2017 DimensionSetEntry
                    recDimSet.RESET;
                    recDimSet.SETRANGE("Dimension Set ID", "Sales Line"."Dimension Set ID");
                    IF recDimSet.FINDFIRST THEN BEGIN

                        IF "Sales Line"."Posting Group" = 'MERCH' THEN
                            ItemDescription := ''
                        ELSE
                            ItemDescription := "Sales Line".Description;
                    END;

                    //<<11May2017 DimensionSetEntry
                    /*
                    PostedDocDim.RESET;
                    PostedDocDim.SETRANGE(PostedDocDim."Table ID",37);
                    PostedDocDim.SETRANGE(PostedDocDim."Document No.","Sales Line"."Document No.");
                    PostedDocDim.SETRANGE(PostedDocDim."Line No.","Sales Line"."Line No.");
                    PostedDocDim.SETRANGE(PostedDocDim."Dimension Code","Sales Line"."Posting Group");
                    IF PostedDocDim.FINDFIRST THEN REPEAT
                      DocDimValue:=PostedDocDim."Dimension Value Code";
                    IF DimValue.GET('Merch',DocDimValue) THEN
                      DimDesc:=DimValue.Name;
                    IF "Sales Line"."Posting Group"='MERCH' THEN
                       ItemDescription:=''
                    ELSE
                       ItemDescription:="Sales Line".Description;
                    
                    UNTIL PostedDocDim.NEXT=0;
                    */
                    //END;

                    //>>RB-N 31Oct2017 SalesQty
                    q1 := 0;
                    q2 := 0;
                    q3 := 0;
                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date1m, endd1);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q1 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date2m, endd2);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q2 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date3m, endd3);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q3 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    //>>RB-N 31Oct2017
                    q13 := 0;
                    q14 := 0;
                    q15 := 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date13m, endd13);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q13 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date14m, endd14);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q14 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Sales Line"."Location Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date15m, endd15);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q15 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;
                    //>>RB-N 31Oct2017 SalesQty

                    //>> RB-N 18Dec2017 SupplyLocation
                    LocName18 := '';
                    Loc18.RESET;
                    IF Loc18.GET("Sales Line"."Location Code") THEN
                        LocName18 := Loc18.Name;
                    //<< RB-N 18Dec2017 SupplyLocation

                    IF (ExportToExcel) AND (NOT PrintSummary) THEN  //RSPl-Rahul
                    //IF ExportToExcel THEN
                    BEGIN

                        EnterCell(RowNo, 1, FORMAT(SrNo), FALSE, FALSE, '', 0);
                        EnterCell(RowNo, 2, FORMAT("No."), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 3, FORMAT(ItemDescription), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 4, FORMAT(itemcategory.Description), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 5, FORMAT(ProductGroupName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 6, FORMAT(''), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 7, FORMAT(ProductGrade), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 8, FORMAT(ResponsibilityName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 9, FORMAT(docdate), FALSE, FALSE, '', 2);
                        EnterCell(RowNo, 10, FORMAT("Document No."), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 11, FORMAT(PartyName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 12, FORMAT(LocName18), FALSE, FALSE, '', 1);//RB-N 18Dec2017
                        EnterCell(RowNo, 13, FORMAT(UserName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 14, FORMAT(recSalesApprovalEntry."Approval Date"), FALSE, FALSE, '', 2);
                        EnterCell(RowNo, 15, FORMAT(recSalesApprovalEntry."Approval Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(RowNo, 16, FORMAT("Unit of Measure"), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 17, FORMAT("Qty. per Unit of Measure"), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 18, FORMAT(Quantity), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 19, FORMAT("Quantity (Base)"), FALSE, FALSE, '0.000', 0);//RB-N 22Dec2017
                        EnterCell(RowNo, 20, FORMAT(''), FALSE, FALSE, '0.000', 0);//RB-N 26Dec2017
                        EnterCell(RowNo, 21, FORMAT("Quantity Shipped"), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 22, FORMAT(QuantityToSHipSale), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 26, FORMAT("Document Type"), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 27, FORMAT(BinQty / ItemUOM), FALSE, FALSE, '0.000', 0);
                        //EnterCell(RowNo,25,FORMAT(SalesComment),FALSE,FALSE,'',1);
                        EnterCell(RowNo, 28, SalesPerson, FALSE, FALSE, '', 1);//RB-N 31Oct2017

                        //Sales of last three months for customers
                        EnterCell(RowNo, 30, FORMAT(q1), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 31, FORMAT(q13), FALSE, FALSE, '0.000', 0);//RB-N 31Oct2017

                        EnterCell(RowNo, 32, FORMAT(q2), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 33, FORMAT(q14), FALSE, FALSE, '0.000', 0);//RB-N 31Oct2017

                        EnterCell(RowNo, 34, FORMAT(q3), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 35, FORMAT(q15), FALSE, FALSE, '0.000', 0);//RB-N 31Oct2017

                        EnterCell(RowNo, 40, FORMAT(Comm1), FALSE, FALSE, '', 1);//RB-N 31Oct2017
                        EnterCell(RowNo, 41, FORMAT(Comm2), FALSE, FALSE, '', 1);//RB-N 31Oct2017
                        EnterCell(RowNo, 42, FORMAT(Comm3), FALSE, FALSE, '', 1);//RB-N 31Oct2017
                        EnterCell(RowNo, 43, FORMAT(Comm4), FALSE, FALSE, '', 1);//RB-N 31Oct2017
                        EnterCell(RowNo, 44, FORMAT(Zone), FALSE, FALSE, '', 1);//RSPLSUM 01Sept2020
                        EnterCell(RowNo, 45, FORMAT("MRP/Sales Price"), FALSE, FALSE, '0.00', 0);//RSPLSUM 03Mar21
                        EnterCell(RowNo, 46, FORMAT("Last Billing Price"), FALSE, FALSE, '0.00', 0);//RSPLSUM 03Mar21
                        EnterCell(RowNo, 47, FORMAT(/*"Transfer Indent Line".Remarks*/''), FALSE, FALSE, '', 1);//Fahim 20-11-21
                        EnterCell(RowNo, 48, FORMAT(FOC), FALSE, FALSE, '', 1);//Fahim 20-11-21
                        EnterCell(RowNo, 49, FORMAT("Sales Line"."List Price"), FALSE, FALSE, '', 1);//Fahim 22-11-23

                        RowNo += 1; //27Mar2017

                    END;

                end;

                trigger OnPreDataItem()
                begin

                    //>>08Aug2017
                    IF RespCenter08 <> '' THEN
                        SETRANGE("Sales Line"."Responsibility Center", RespCenter08);

                    //<<08Aug2017

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
                end;
            }
            dataitem("Transfer Indent Line"; "Transfer Indent Line")
            {
                DataItemLink = "Item No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Approve = FILTER(true),
                                          Closed = FILTER(false));

                trigger OnAfterGetRecord()
                begin

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

                    //<<RSPL-Rahul

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
                    recTSL.SETCURRENTKEY("Item No.");//31Oct2017
                    //recTSL.SETRANGE(recTSL."Transfer Indent No.", "Transfer Indent Line"."Document No.");
                    recTSL.SETRANGE(recTSL."Item No.", "Transfer Indent Line"."Item No.");
                    //recTSL.SETRANGE(recTSL."Transfer Indent Line No.", "Transfer Indent Line"."Line No.");
                    IF recTSL.FINDSET THEN
                        REPEAT
                            QtyShip := QtyShip + recTSL.Quantity;
                        UNTIL recTSL.NEXT = 0;

                    QuantityToSHipTransfer := "Transfer Indent Line".Quantity - QtyShip;

                    CLEAR(BinName);
                    CLEAR(BinQty);
                    recBinContent.RESET;
                    recBinContent.SETRANGE(recBinContent."Item No.", "Transfer Indent Line"."Item No.");
                    recBinContent.SETFILTER(recBinContent."Location Code", 'PLANT01');
                    //recBinContent.SETFILTER(recBinContent."Bin Code",'V-S-INDL');
                    recBinContent.SETFILTER(recBinContent."Bin Code", '%1|%2', 'V-S-INDL', 'V-S-AUTO');
                    recBinContent.SETFILTER(recBinContent.Quantity, '<>%1', 0);
                    IF recBinContent.FIND('-') THEN BEGIN
                        recBinContent.CALCFIELDS(recBinContent.Quantity);
                        BinQty := BinQty + recBinContent.Quantity;

                        recBin.RESET;
                        //recBin.SETRANGE(recBin.Code,'V-S-INDL');
                        //recBin.SETFILTER(recBin.Code,'%1|%2','V-S-INDL','V-S-AUTO');
                        recBin.SETRANGE(recBin.Code, recBinContent."Bin Code");
                        IF recBin.FINDFIRST THEN
                            BinName := recBin.Description
                    END;

                    CLEAR(QtyperUOM);
                    recItemUOMEBT.RESET;
                    recItemUOMEBT.SETRANGE(recItemUOMEBT."Item No.", "Transfer Indent Line"."Item No.");
                    recItemUOMEBT.SETRANGE(recItemUOMEBT.Code, "Transfer Indent Line"."Unit of Measure Code");
                    IF recItemUOMEBT.FINDFIRST THEN BEGIN
                        QtyperUOM := recItemUOMEBT."Qty. per Unit of Measure";
                    END;

                    //TotQtyPerUOM += QtyperUOM;//RSPLSUM 24Feb21  //Nb commented as total not needed in report
                    TotQtyPerUOM := QtyperUOM;//RSPLSUM 24Feb21  //Nb added

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
                    CLEAR(Zone);//RSPLSUM 01Sept2020
                    SalesPerson := '';
                    IF recSP.GET("Transfer Indent Line"."Salesperson Code") THEN BEGIN
                        SalesPerson := recSP.Name;
                        Zone := recSP."Zone Code";//RSPLSUM 01Sept2020
                    END;//RSPLSUM 01Sept2020
                        //RSPL-CAS-04663-V8J3G6

                    //RSPL-CAS-07044-N8Z9H8
                    Qty_sum := 0;
                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
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

                    /*
                    //>>Instransit Inventory to be Calculated
                    IntransitInv:=0;
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Location Code","Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type",recILE."Document Type"::"Transfer Shipment");
                    recILE.SETRANGE(recILE."Item No.",Item."No.");
                    recILE.SETFILTER(recILE."Posting Date",'<=%1',maxdate);
                    IF recILE.FINDSET THEN REPEAT
                      IF recTRO.GET(recILE."Order No.") THEN BEGIN
                       //IF recTRO."Transfer-to Code"="Transfer Indent Line"."Transfer-to Code" THEN BEGIN
                        recIntransit.RESET;
                        recIntransit.SETRANGE("Order No.",recILE."Order No.");
                        recIntransit.SETRANGE("Location Code",'IN-TRANS');
                        IF recIntransit.FINDSET THEN REPEAT
                           IntransitInv+=recIntransit.Quantity;
                        UNTIL recIntransit.NEXT=0;
                       //END;
                      END;
                    UNTIL recILE.NEXT=0;
                    //<<End
                    *///Commented 18May2017

                    //>>IntransitQty 18May2017
                    IntransitInv := 0;
                    TransferLine18.RESET;
                    TransferLine18.SETCURRENTKEY("Item No.");
                    TransferLine18.SETRANGE("Transfer-to Code", "Transfer Indent Line"."Transfer-to Code");
                    TransferLine18.SETRANGE("Transfer-from Code", 'IN-TRANS');
                    TransferLine18.SETRANGE("Item No.", Item."No.");
                    TransferLine18.SETFILTER("Shipment Date", '<=%1', maxdate);
                    IF TransferLine18.FINDSET THEN
                        REPEAT

                            IntransitInv += TransferLine18.Quantity;

                        UNTIL TransferLine18.NEXT = 0;
                    //<<IntransitQty 18May2017

                    q1 := 0;
                    q2 := 0;
                    q3 := 0;
                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date1m, endd1);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q1 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date2m, endd2);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q2 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date3m, endd3);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q3 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    //>>RB-N 31Oct2017
                    q13 := 0;
                    q14 := 0;
                    q15 := 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date13m, endd13);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q13 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date14m, endd14);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q14 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Sales Shipment");
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Item No.", Item."No.");
                    recILE.SETRANGE(recILE."Posting Date", date15m, endd15);
                    IF recILE.FINDSET THEN
                        REPEAT
                            q15 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;
                    //>>RB-N 31Oct2017 SalesQty

                    salc1 := 0;
                    salc2 := 0;
                    salc3 := 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    //recILE.SETRANGE(recILE."Item No.","Sales Line"."No.");
                    recILE.SETRANGE(recILE."Item No.", "Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Transfer Shipment");
                    recILE.SETRANGE(recILE."Posting Date", date1m, endd1);
                    IF recILE.FINDSET THEN
                        REPEAT
                            salc1 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    //recILE.SETRANGE(recILE."Item No.","Sales Line"."No.");
                    recILE.SETRANGE(recILE."Item No.", "Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Transfer Shipment");
                    recILE.SETRANGE(recILE."Posting Date", date2m, endd2);
                    IF recILE.FINDSET THEN
                        REPEAT
                            salc2 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.");//RB-N 31Oct2017
                    //recILE.SETRANGE(recILE."Item No.","Sales Line"."No.");
                    recILE.SETRANGE(recILE."Item No.", "Item No.");//RB-N 31Oct2017
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    recILE.SETRANGE(recILE."Document Type", recILE."Document Type"::"Transfer Shipment");
                    recILE.SETRANGE(recILE."Posting Date", date3m, endd3);
                    IF recILE.FINDSET THEN
                        REPEAT
                            salc3 += ABS(recILE.Quantity);
                        UNTIL recILE.NEXT = 0;

                    IF "Transfer Indent Line"."Inventory Posting Group" = 'MERCH' THEN BEGIN
                        ItemDescription := "Transfer Indent Line"."Merch Name"
                    END ELSE
                        ItemDescription := "Transfer Indent Line".Description;


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
                    //>>RSPl-Rahul
                    itemqtytotal += Quantity;
                    totqtyship += QtyShip;
                    totqtytoship += QuantityToSHipTransfer;
                    //<<
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
                    recTSL.SETCURRENTKEY("Item No.");//31Oct2017
                    recTSL.SETRANGE(recTSL."Transfer-from Code", "Transfer Indent Line"."Transfer-from Code");
                    recTSL.SETRANGE(recTSL."Item No.", "Transfer Indent Line"."Item No.");
                    recTSL.SETFILTER(recTSL."Shipment Date", '%1..%2', mindate, maxdate);
                    recTSL.SETRANGE(recTSL."Transfer-to Code", "Transfer Indent Line"."Transfer-to Code");
                    IF recTSL.FINDFIRST THEN
                        REPEAT
                            CulumativeQty := CulumativeQty + recTSL.Quantity;
                            CulumativeQtyBase := CulumativeQtyBase + recTSL."Quantity (Base)";
                        UNTIL recTSL.NEXT = 0;


                    //>> RB-N 18Dec2017 SupplyLocation
                    LocName18 := '';
                    Loc18.RESET;
                    IF Loc18.GET("Transfer Indent Line"."Transfer-from Code") THEN
                        LocName18 := Loc18.Name;
                    //<< RB-N 18Dec2017 SupplyLocation

                    IF ExportToExcel AND (NOT PrintSummary) THEN   //Rahul
                    //IF ExportToExcel THEN
                    BEGIN

                        //EnterCell(RowNo+3,1,FORMAT(SrNo),FALSE,FALSE,'',0);
                        EnterCell(RowNo, 1, FORMAT(SrNo), FALSE, FALSE, '', 0);
                        EnterCell(RowNo, 2, FORMAT("Item No."), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 3, FORMAT(ItemDescription), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 4, FORMAT(itemCatforind), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 5, FORMAT(ProductGroupName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 6, FORMAT(''), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 7, FORMAT(ProductGrade), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 8, FORMAT(ResponsibilityName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 9, FORMAT("Transfer Indent Line"."Addition date"), FALSE, FALSE, '', 2);
                        EnterCell(RowNo, 10, FORMAT("Document No."), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 11, FORMAT(LocationName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 12, FORMAT(LocName18), FALSE, FALSE, '', 1);//RB-N 18Dec2017
                        EnterCell(RowNo, 13, FORMAT(UserName), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 14, FORMAT("Transfer Indent Line"."Approval Date"), FALSE, FALSE, '', 2);
                        EnterCell(RowNo, 15, FORMAT("Transfer Indent Line"."Approval Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(RowNo, 16, FORMAT("Unit of Measure Code"), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 17, FORMAT(QtyperUOM), FALSE, FALSE, '0.000', 0);
                        //EnterCell(RowNo+3,17,FORMAT(Quantity)+'Test',FALSE,FALSE,'0.000',0);
                        EnterCell(RowNo, 18, FORMAT(Quantity), FALSE, FALSE, '0.000', 0);//27Mar2017
                        EnterCell(RowNo, 19, FORMAT(Quantity * QtyperUOM), FALSE, FALSE, '0.000', 0);//RB-N 22Dec2017
                        EnterCell(RowNo, 20, FORMAT(''), FALSE, FALSE, '0.000', 0);//RB-N 26Dec2017
                        EnterCell(RowNo, 21, FORMAT(QtyShip), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 22, FORMAT(QuantityToSHipTransfer), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 26, FORMAT('Indent'), FALSE, FALSE, '', 1);
                        EnterCell(RowNo, 27, FORMAT(BinQty / ItemUOM), FALSE, FALSE, '0.000', 0);
                        //RSPL-CAS-04663-V8J3G6
                        EnterCell(RowNo, 28, FORMAT(SalesPerson), FALSE, FALSE, '', 1);
                        //RSPL-CAS-04663-V8J3G6
                        EnterCell(RowNo, 29, FORMAT(qtyuom), FALSE, FALSE, '0.000', 0); //RSPL-CAS-07044-N8Z9H8
                                                                                        //sales of last three months
                        EnterCell(RowNo, 30, FORMAT(q1), FALSE, FALSE, '0.000', 0);//RB-N Qty 2017
                        EnterCell(RowNo, 31, FORMAT(q13), FALSE, FALSE, '0.000', 0);//RB-N Qty 31Oct2017

                        EnterCell(RowNo, 32, FORMAT(q2), FALSE, FALSE, '0.000', 0);//RB-N Qty 2017
                        EnterCell(RowNo, 33, FORMAT(q14), FALSE, FALSE, '0.000', 0);//RB-N Qty 31Oct2017

                        EnterCell(RowNo, 34, FORMAT(q3), FALSE, FALSE, '0.000', 0);//RB-N Qty 2017
                        EnterCell(RowNo, 35, FORMAT(q15), FALSE, FALSE, '0.000', 0);//RB-N Qty 31Oct2017

                        EnterCell(RowNo, 36, FORMAT(salc1), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 37, FORMAT(salc2), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 38, FORMAT(salc3), FALSE, FALSE, '0.000', 0);

                        EnterCell(RowNo, 39, FORMAT(IntransitInv), FALSE, FALSE, '0.000', 0);
                        EnterCell(RowNo, 44, FORMAT(Zone), FALSE, FALSE, '', 1);//RSPLSUM 01Sept2020
                        EnterCell(RowNo, 47, FORMAT("Transfer Indent Line".Remarks), FALSE, FALSE, '', 1);//RSPLSUM 19Mar21
                        EnterCell(RowNo, 48, FORMAT(FOC), FALSE, FALSE, '', 1);//Fahim 20-11-21
                        EnterCell(RowNo, 49, FORMAT("Sales Line"."List Price"), FALSE, FALSE, '', 1);//Fahim 23-11-23

                        RowNo += 1; //27Mar2017
                    END;

                end;

                trigger OnPreDataItem()
                begin
                    //>>08Aug2017
                    IF RespCenter08 <> '' THEN
                        SETRANGE("Transfer Indent Line"."Shortcut Dimension 2 Code", RespCenter08);

                    IF SaleCode08 <> '' THEN
                        SETRANGE("Transfer Indent Line"."Salesperson Code", SaleCode08);
                    //<<08Aug2017

                    IF locfilt <> '' THEN
                        "Transfer Indent Line".SETRANGE("Transfer Indent Line"."Transfer-from Code", locfilt);
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
                    //MonthCaption1:=ReturnMonth(MONTH);
                    MonthCaption1 := FORMAT(ReturnMonth(MONTH)) + FORMAT(YEAR);//RB-N 31Oct2017

                    //Start Date -2M
                    date2m := CALCDATE('-1M', CalDate);//start2
                    MONTH := DATE2DMY(date2m, 2);
                    ED2 := ReturnRespEndDate(MONTH);
                    endd2 := DMY2DATE(ED2, MONTH, YEAR);//end2
                    //MonthCaption2:=ReturnMonth(MONTH);
                    MonthCaption2 := FORMAT(ReturnMonth(MONTH)) + FORMAT(YEAR);//RB-N 31Oct2017

                    //Start Date -3M
                    date3m := CALCDATE('-2M', CalDate);//start3
                    MONTH := DATE2DMY(date3m, 2);
                    ED3 := ReturnRespEndDate(MONTH);
                    endd3 := DMY2DATE(ED3, MONTH, YEAR);//end3
                    //MonthCaption3:=ReturnMonth(MONTH);
                    MonthCaption3 := FORMAT(ReturnMonth(MONTH)) + FORMAT(YEAR);//RB-N 31Oct2017
                end;
            }
            dataitem(Integer; Integer)
            {
                DataItemTableView = SORTING(Number)
                                    WHERE(Number = CONST(1));

                trigger OnAfterGetRecord()
                begin
                    //>>Section

                    // IF itemqtytotal=0 THEN
                    //   CurrReport.SHOWOUTPUT(FALSE);
                    IF itemqtytotal <> 0 THEN BEGIN
                        IF ExportToExcel AND (SalesExist OR TransferExists) THEN BEGIN

                            EnterCell(RowNo, 2, FORMAT(ItemNo), TRUE, FALSE, '', 1);
                            EnterCell(RowNo, 3, FORMAT(itemdesc), TRUE, FALSE, '', 1);
                            //RSPLSUM 24Feb21--EnterCell(RowNo,17,FORMAT(ItemUOM),TRUE,FALSE,'0.000',0);
                            EnterCell(RowNo, 17, FORMAT(TotQtyPerUOM), TRUE, FALSE, '0.000', 0);//RSPLSUM 24Feb21
                            EnterCell(RowNo, 18, FORMAT(itemqtytotal), TRUE, FALSE, '0.000', 0);//27Mar2017
                            EnterCell(RowNo, 19, FORMAT(TotalQtyBase), TRUE, FALSE, '0.000', 0);//RB-N 22Dec2017
                            EnterCell(RowNo, 20, FORMAT(TotalQtyProd), TRUE, FALSE, '0.000', 0);//RB-N 22Dec2017
                                                                                                //EnterCell(RowNo+3,17,FORMAT(itemqtytotal)+'test tot2',TRUE,FALSE,'0.000',0);
                            EnterCell(RowNo, 21, FORMAT(totqtyship), TRUE, FALSE, '0.000', 0);
                            EnterCell(RowNo, 22, FORMAT(totqtytoship), TRUE, FALSE, '0.000', 0);
                            EnterCell(RowNo, 23, FORMAT(BinQty / ItemUOM), TRUE, FALSE, '0.000', 0);
                            EnterCell(RowNo, 24, FORMAT((BinQty / ItemUOM) - totqtytoship), TRUE, FALSE, '0.000', 0);
                            EnterCell(RowNo, 25, FORMAT(BinName), TRUE, FALSE, '', 1);

                            RowNo += 1;//27Mar2017
                        END;
                    END;

                    printheader := FALSE;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                CLEAR(SrNo);
                CLEAR(itemqtytotal);
                CLEAR(totqtyship);
                CLEAR(totqtytoship);
                CLEAR(BINCode);
                CLEAR(BinName);
                CLEAR(TotalQtyBase);//RB-N 22Dec2017
                CLEAR(TotalQtyProd);//RB-N 26Dec2017
                CLEAR(TotQtyPerUOM);//RSPLSUM 24Feb21

                //>>RSPL/Migration/Rahul/**No need of code

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

                //<<
                //>>Section 1
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

                //>>Section 2

                //>> RB-N QtyinProduction  26Dec2017
                ILE26.RESET;
                ILE26.SETCURRENTKEY("Item No.");
                ILE26.SETRANGE("Item No.", "No.");
                ILE26.SETRANGE("Posting Date", mindate, maxdate);
                ILE26.SETRANGE("Location Code", 'PLANT01');
                ILE26.SETRANGE("Entry Type", ILE26."Entry Type"::Output);
                IF ILE26.FINDSET THEN
                    REPEAT
                        TotalQtyProd += ILE26.Quantity / ILE26."Qty. per Unit of Measure";
                    UNTIL ILE26.NEXT = 0;
                //<< RB-N QtyinProduction  26Dec2017
            end;

            trigger OnPreDataItem()
            begin

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
                //MonthCaption1:=ReturnMonth(MONTH);
                MonthCaption1 := FORMAT(ReturnMonth(MONTH)) + '-' + FORMAT(YEAR);//31Oct2017

                //Start Date -2M
                date2m := CALCDATE('-1M', CalDate);//start2
                MONTH := DATE2DMY(date2m, 2);
                ED2 := ReturnRespEndDate(MONTH);
                endd2 := DMY2DATE(ED2, MONTH, YEAR);//end2
                //MonthCaption2:=ReturnMonth(MONTH);
                MonthCaption2 := ReturnMonth(MONTH) + '-' + FORMAT(YEAR);//31Oct2017

                //Start Date -3M
                date3m := CALCDATE('-2M', CalDate);//start3
                MONTH := DATE2DMY(date3m, 2);
                ED3 := ReturnRespEndDate(MONTH);
                endd3 := DMY2DATE(ED3, MONTH, YEAR);//end3
                //MonthCaption3:=ReturnMonth(MONTH);
                MonthCaption3 := ReturnMonth(MONTH) + '-' + FORMAT(YEAR);//31Oct2017

                //RB-N 31Oct2017 Start Date -13M
                date13m := CALCDATE('-12M', CalDate);//start13
                MONTH := DATE2DMY(date13m, 2);
                ED13 := ReturnRespEndDate(MONTH);
                endd13 := DMY2DATE(ED13, MONTH, YEAR - 1);//end13
                MonthCaption13 := FORMAT(ReturnMonth(MONTH)) + '-' + FORMAT(YEAR - 1);//31Oct2017
                //MESSAGE('%1 Month 13',MonthCaption13);
                //RB-N 31Oct2017 Start Date -13M

                //RB-N 31Oct2017 Start Date -14M
                date14m := CALCDATE('-13M', CalDate);//start14
                MONTH := DATE2DMY(date14m, 2);
                ED14 := ReturnRespEndDate(MONTH);
                endd14 := DMY2DATE(ED14, MONTH, YEAR - 1);//end14
                MonthCaption14 := FORMAT(ReturnMonth(MONTH)) + '-' + FORMAT(YEAR - 1);//31Oct2017
                //RB-N 31Oct2017 Start Date -14M

                //RB-N 31Oct2017 Start Date -15M
                date15m := CALCDATE('-14M', CalDate);//start15
                MONTH := DATE2DMY(date15m, 2);
                ED15 := ReturnRespEndDate(MONTH);
                endd15 := DMY2DATE(ED15, MONTH, YEAR - 1);//end15
                MonthCaption15 := FORMAT(ReturnMonth(MONTH)) + '-' + FORMAT(YEAR - 1);//31Oct2017
                //RB-N 31Oct2017 Start Date -15M


                //>>Section

                IF ExportToExcel AND printheader THEN BEGIN
                    //RowNo += 1;
                    EnterCell(1, 1, FORMAT(COMPANYNAME), TRUE, FALSE, '@', 1);
                    EnterCell(1, 46, FORMAT(TODAY, 0, 4), TRUE, FALSE, '@', 1);
                    EnterCell(2, 1, 'Pending Indent/CSO Itemwise', TRUE, FALSE, '@', 1);
                    EnterCell(2, 46, FORMAT(USERID), TRUE, FALSE, '@', 1);
                    //>>27Mar2017
                    EnterCell(4, 1, 'Date Filter : ' + Item.GETFILTER("Date Filter"), TRUE, FALSE, '@', 1);

                    //LocationName
                    Loc14Code := Item.GETFILTER("Location Filter");
                    LOC14.RESET;
                    LOC14.SETRANGE(Code, Loc14Code);
                    IF LOC14.FINDFIRST THEN
                        Loc14Name := LOC14.Name;

                    EnterCell(5, 1, 'Location : ' + Loc14Name, TRUE, FALSE, '@', 1);

                    //ItemCategory
                    Cat08Char := Item.GETFILTER("Item Category Code");

                    IF Cat08Char <> '' THEN BEGIN
                        Cat08Length := STRLEN(Cat08Char);
                        ItemCat14 := COPYSTR(Cat08Char, 1, 5); //08Aug2017
                                                               //ItemCat14 := Item.GETRANGEMIN("Item Category Code");//08Aug2017

                        IF Cat08Length > 10 THEN
                            ItemCat14_08 := COPYSTR(Cat08Char, Cat08Length - 4, Cat08Length); //08Aug2017
                    END;


                    ItemCatR.RESET;
                    //ItemCatR.SETRANGE(Code,ItemCat14);
                    ItemCatR.SETRANGE(Code, ItemCat14);//08Aug2017
                    IF ItemCatR.FINDFIRST THEN
                        ItemCat14Name := ItemCatR.Description;

                    IF ItemCat14_08 <> '' THEN //08Aug2017
                        ItemCat14Name := ''; //08Aug2017

                    EnterCell(6, 1, 'Item Category : ' + ItemCat14Name, TRUE, FALSE, '@', 1);

                    //<<27Mar2017
                    EnterCell(7, 1, FORMAT('Sr.No.'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 2, FORMAT('Item No.'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 3, FORMAT('Item Description'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 4, FORMAT('Item Category'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 5, FORMAT('Product Group'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 6, FORMAT('Product Sub Group'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 7, FORMAT('Product Grade'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 8, FORMAT('Branch'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 9, FORMAT('Doc./Addition date'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 10, FORMAT('Document No.'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 11, FORMAT('Party/ Location Name'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 12, FORMAT('Supply Location'), TRUE, TRUE, '@', 1);//18Dec2017
                    EnterCell(7, 13, FORMAT('Approver Name'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 14, FORMAT('Approve Date'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 15, FORMAT('Approve time'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 16, FORMAT('Pack UOM'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 17, FORMAT('Qty per UOM'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 18, FORMAT('Qty. Order as per pack UOM'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 19, FORMAT('Qty. in Ltrs'), TRUE, TRUE, '', 1);//RB-N 22Dec2017
                    EnterCell(7, 20, FORMAT('Qty. Produced in UOM'), TRUE, TRUE, '', 1);//RB-N 26Dec2017
                    EnterCell(7, 21, FORMAT('Despatched per pack UOM'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 22, FORMAT('Balance  as per pack UOM'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 23, FORMAT('Product Availability'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 24, FORMAT('Excess/Short for Production'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 25, FORMAT('Bin'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 26, FORMAT('Type'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 27, FORMAT('Stock at from Location in Sales UOM'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 28, FORMAT('Salesperson'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 29, FORMAT('Stock at Indented Location in Sales UOM'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 30, FORMAT('Sales for the Month ' + FORMAT(MonthCaption1)), TRUE, TRUE, '@', 1);//RB-N 31Oct2017 2017Year
                    EnterCell(7, 31, FORMAT('Sales for the Month ' + FORMAT(MonthCaption13)), TRUE, TRUE, '@', 1);//RB-N 31Oct2017 2016Year

                    EnterCell(7, 32, FORMAT('Sales for the Month ' + FORMAT(MonthCaption2)), TRUE, TRUE, '@', 1);//RB-N 31Oct2017 2017Year
                    EnterCell(7, 33, FORMAT('Sales for the Month ' + FORMAT(MonthCaption14)), TRUE, TRUE, '@', 1);//RB-N 31Oct2017 2016Year

                    EnterCell(7, 34, FORMAT('Sales for the Month ' + FORMAT(MonthCaption3)), TRUE, TRUE, '@', 1);//RB-N 31Oct2017 2017Year
                    EnterCell(7, 35, FORMAT('Sales for the Month ' + FORMAT(MonthCaption15)), TRUE, TRUE, '@', 1);//RB-N 31Oct2017 2016Year

                    EnterCell(7, 36, FORMAT('Transfer for the Month ' + FORMAT(MonthCaption1)), TRUE, TRUE, '@', 1);
                    EnterCell(7, 37, FORMAT('Transfer for the Month ' + FORMAT(MonthCaption2)), TRUE, TRUE, '@', 1);
                    EnterCell(7, 38, FORMAT('Transfer for the Month ' + FORMAT(MonthCaption3)), TRUE, TRUE, '@', 1);
                    EnterCell(7, 39, FORMAT('Intransit Inventory'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 40, FORMAT('Remarks'), TRUE, TRUE, '@', 1);
                    EnterCell(7, 41, FORMAT('Remarks 2'), TRUE, TRUE, '@', 1); //07Sep2017
                    EnterCell(7, 42, FORMAT('Remarks 3'), TRUE, TRUE, '@', 1); //07Sep2017
                    EnterCell(7, 43, FORMAT('Remarks 4'), TRUE, TRUE, '@', 1); //07Sep2017
                    EnterCell(7, 44, FORMAT('Zone'), TRUE, TRUE, '@', 1);//RSPLSUM 01Sept2020
                    EnterCell(7, 45, FORMAT('Sales Price'), TRUE, TRUE, '@', 1);//RSPLSUM 03Mar21
                    EnterCell(7, 46, FORMAT('Last Billing Price'), TRUE, TRUE, '@', 1);//RSPLSUM 03Mar21
                    EnterCell(7, 47, FORMAT('Remarks'), TRUE, TRUE, '@', 1);//RSPLSUM 19Mar21
                    EnterCell(7, 48, FORMAT('FOC'), TRUE, TRUE, '@', 1);//Fahim
                    EnterCell(7, 49, FORMAT('List Price'), TRUE, TRUE, '@', 1);//Fahim 22-11-23

                    RowNo := 8;
                END;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Responsibility Center"; RespCenter08)
                {
                    ApplicationArea = ALL;
                    TableRelation = "Responsibility Center";
                }
                field("SalesPerson Code"; SaleCode08)
                {
                    ApplicationArea = ALL;
                    TableRelation = "Salesperson/Purchaser";
                }
                field(ExportToExcel; ExportToExcel)
                {
                    ApplicationArea = all;
                    Caption = 'Export to Excel';
                }
                field(PrintDetail; PrintDetail)
                {
                    ApplicationArea = all;
                    Caption = 'Print Details';
                }
                field(PrintSummary; PrintSummary)
                {
                    ApplicationArea = all;
                    Caption = 'Print Summary';
                }
                field(Showshortclosed; Showshortclosed)
                {
                    ApplicationArea = all;
                    Caption = 'Show Short Close';
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

        IF ExportToExcel THEN BEGIN
            //TempExcelBuffer1.CreateBook;
            //TempExcelBuffer1.CreateSheet('Pending Indent CSO', '', COMPANYNAME, USERID);
            /* TempExcelBuffer1.CreateBookAndOpenExcel('', 'Pending Indent CSO', '', '', '');
            TempExcelBuffer1.GiveUserControl; */
            TempExcelBuffer1.CreateNewBook(Text001);
            TempExcelBuffer1.WriteSheet(Text001, CompanyName, UserId);
            TempExcelBuffer1.CloseBook();
            TempExcelBuffer1.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
            TempExcelBuffer1.OpenExcel();
        END;
    end;

    trigger OnPreReport()
    begin

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
        */

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
        ProductGrade: Text[30];
        ProductGroup: Text[30];
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
        DocDimValue: Text[50];
        DimDesc: Text[50];
        ItemDescription: Text[80];
        DimValue: Record 349;
        Text000: Label 'Data';
        Text001: Label 'Pending Indent CSO';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "----27Mar2017": Integer;
        LOC14: Record 14;
        Loc14Code: Code[50];
        Loc14Name: Text[100];
        ItemCat14: Code[50];
        ItemCat14Name: Text[100];
        ItemCatR: Record 5722;
        "---11May2017": Integer;
        recDimSet: Record 480;
        "-----18May2017": Integer;
        TransferLine18: Record 5741;
        "----08Aug2017": Integer;
        RespCenter08: Code[20];
        SaleCode08: Code[20];
        SH08: Record 36;
        ItemCat14_08: Code[50];
        Cat08Length: Integer;
        Cat08Char: Code[50];
        "------07Sep2017": Integer;
        Comm1: Text[250];
        Comm2: Text[250];
        Comm3: Text[250];
        Comm4: Text[250];
        C07: Integer;
        CommentText: Text[1024];
        CommLen: Integer;
        "--------31Oct2017": Integer;
        date13m: Date;
        date14m: Date;
        date15m: Date;
        ED13: Integer;
        ED14: Integer;
        ED15: Integer;
        endd13: Date;
        endd14: Date;
        endd15: Date;
        MonthCaption13: Text[30];
        MonthCaption14: Text[30];
        MonthCaption15: Text[30];
        q13: Decimal;
        q14: Decimal;
        q15: Decimal;
        "----------18Dec2017": Integer;
        Loc18: Record 14;
        LocName18: Text;
        TotalQtyBase: Decimal;
        TotalQtyProd: Decimal;
        ILE26: Record 32;
        Zone: Code[10];
        TotQtyPerUOM: Decimal;
        FOC: Text[10];

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
        TempExcelBuffer1.NumberFormat := NoFormat;
        TempExcelBuffer1."Cell Type" := CType;
        TempExcelBuffer1.INSERT;
    end;

    //  //[Scope('Internal')]
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

    // //[Scope('Internal')]
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


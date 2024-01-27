report 50030 "Daily Production Memo(Excel)"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 13Mar2018   RB-N         Bulk Code & Bulk Description
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/DailyProductionMemoExcel.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            DataItemTableView = SORTING("Location Code", "Item No.", "Lot No.", "Posting Date", "Document No.")
                                WHERE("Entry Type" = CONST(Output),
                                      "Location Code" = FILTER('PLANT01 | PLANT02 | CAPCON | "DAMAN PR" | PIPAVAV PR|PLANT03|PIPAVAVFG'),
                                      "Lot No." = FILTER('<>REV.&<>REV'));
            RequestFilterFields = "Item No.", "Location Code", "Posting Date", "Item Category Code";// "Product Group Code";

            trigger OnAfterGetRecord()
            begin


                //>>10Aug2017 HSN Code
                CLEAR(HSNCode10);
                Itm10.RESET;
                IF Itm10.GET("Item Ledger Entry"."Item No.") THEN
                    HSNCode10 := Itm10."HSN/SAC Code";

                //<<10Aug2017 HSN Code

                CLEAR(Qtytotal);
                RECILE.RESET;
                RECILE.SETRANGE(RECILE."Document No.", "Item Ledger Entry"."Document No.");
                RECILE.SETRANGE(RECILE."Entry Type", "Item Ledger Entry"."Entry Type");
                RECILE.SETRANGE(RECILE."Lot No.", "Item Ledger Entry"."Lot No.");
                IF RECILE.FINDSET THEN
                    REPEAT
                        Qtytotal += RECILE.Quantity;
                    UNTIL RECILE.NEXT = 0;

                IF Qtytotal = 0 THEN
                    CurrReport.SKIP;

                companyinfo.GET;
                IF Location.GET("Location Code") THEN;

                IF Location.GET(UserSetup."Location Filter") THEN;
                LocCode := COPYSTR(Location.Name, 7, 30);


                itemname := '';
                PostingGP := '';
                Qty := 0;
                QtyInLtr := 0;
                SpecificGravity := 0;

                IF item1.GET("Item Ledger Entry"."Item No.") THEN
                    itemname := item1.Description;
                // PostingGP:=item1."Excise Prod. Posting Group"; //28dec2023

                ProductionOrder.RESET;
                ProductionOrder.SETRANGE(ProductionOrder."No.", "Item Ledger Entry"."Document No.");
                IF ProductionOrder.FINDFIRST THEN BEGIN
                    IF ProductionOrder."Order Type" = ProductionOrder."Order Type"::Secondary THEN BEGIN
                        ILE.RESET;
                        ILE.SETRANGE(ILE."Document No.", "Item Ledger Entry"."Document No.");
                        ILE.SETRANGE(ILE."Entry Type", ILE."Entry Type"::Consumption);
                        ILE.SETRANGE(ILE."Item Category Code", 'CAT10');
                        IF ILE.FINDSET THEN
                            REPEAT
                                QtyInLtr += ABS(ILE.Quantity);
                            UNTIL ILE.NEXT = 0;
                        IF QtyInLtr <> 0 THEN
                            SpecificGravity := ABS(QtyInLtr / "Item Ledger Entry".Quantity);
                    END
                    ELSE
                        IF ProductionOrder."Order Type" = ProductionOrder."Order Type"::Primary THEN BEGIN
                            ILE.RESET;
                            ILE.SETRANGE(ILE."Document No.", "Item Ledger Entry"."Document No.");
                            ILE.SETRANGE(ILE."Entry Type", ILE."Entry Type"::Consumption);
                            IF ILE.FINDSET THEN
                                REPEAT
                                    QtyInLtr += ABS(ILE.Quantity);
                                UNTIL ILE.NEXT = 0;
                            IF (ILE."Item Category Code" = 'CAT02') OR (ILE."Item Category Code" = 'CAT03') THEN BEGIN
                                IF QtyInLtr <> 0 THEN
                                    SpecificGravity := ABS(QtyInLtr / "Item Ledger Entry".Quantity);
                            END ELSE BEGIN
                                IF QtyInLtr <> 0 THEN BEGIN
                                    SpecificGravity := "Item Ledger Entry"."Density Factor";
                                    QtyInLtr := "Item Ledger Entry".Quantity / SpecificGravity;
                                END;
                            END;
                        END;
                END;


                GrandTotal += "Item Ledger Entry".Quantity / "Item Ledger Entry"."Qty. per Unit of Measure";
                TotalQty += QtyInLtr;

                //RSPL060814
                vBatchNo := '';
                ProductionOrder.RESET;
                ProductionOrder.SETRANGE(ProductionOrder."No.", "Document No.");
                IF ProductionOrder.FINDFIRST THEN
                    vBatchNo := ProductionOrder."Description 2";
                vRoutingName := '';
                LineProd.RESET;
                LineProd.SETRANGE(LineProd."Prod. Order No.", "Document No.");
                LineProd.SETRANGE(LineProd."Item No.", "Item No.");
                IF LineProd.FINDFIRST THEN
                    IF recRouting.GET(LineProd."Routing No.") THEN
                        vRoutingName := recRouting.Description;
                //RSPL


                TotalConvQty += "Item Ledger Entry".Quantity;

                //>>13Mar2018
                CLEAR(BulkCode);
                CLEAR(BulkDesc);
                VE13.RESET;
                VE13.SETCURRENTKEY("Document No.");
                VE13.SETRANGE("Document No.", "Document No.");
                VE13.SETRANGE("Source No.", "Item No.");
                VE13.SETRANGE("Item Ledger Entry Type", VE13."Item Ledger Entry Type"::Consumption);
                VE13.SETRANGE("Gen. Prod. Posting Group", 'BULK');
                IF VE13.FINDFIRST THEN BEGIN
                    BulkCode := VE13."Item No.";
                    Itm10.RESET;
                    IF Itm10.GET(BulkCode) THEN
                        BulkDesc := Itm10.Description;

                END;

                //<<13Mar2018


                IF PrinttoExcel THEN BEGIN
                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn("Item Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//1
                    ExcelBuf.AddColumn("Item Ledger Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//2
                    ExcelBuf.AddColumn(item1."No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3
                    ExcelBuf.AddColumn(itemname, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
                    ExcelBuf.AddColumn(HSNCode10, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5
                    ExcelBuf.AddColumn(BulkCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6
                    ExcelBuf.AddColumn(BulkDesc, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
                    ExcelBuf.AddColumn(vBatchNo, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//8
                    Qty := QtyInLtr / "Item Ledger Entry"."Qty. per Unit of Measure";
                    ExcelBuf.AddColumn("Item Ledger Entry".Quantity / "Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//9
                    ExcelBuf.AddColumn("Item Ledger Entry"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
                    ExcelBuf.AddColumn("Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//11
                    ExcelBuf.AddColumn("Item Ledger Entry".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//12
                    ExcelBuf.AddColumn(PostingGP, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//13
                    ExcelBuf.AddColumn(FORMAT(SpecificGravity, 5, 3), FALSE, '', FALSE, FALSE, FALSE, '#,####0.000', 0);//14
                    ExcelBuf.AddColumn(QtyInLtr, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//15
                                                                                            //RSPL-007
                    IF item1."Inventory Posting Group" = 'BULK' THEN BEGIN
                        ExcelBuf.AddColumn(vRoutingName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//16
                    END
                    ELSE BEGIN
                        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//16
                    END;

                    //>>16Jan2019
                    DivName := '';
                    DimVal.RESET;
                    IF DimVal.GET('DIVISION', "Item Ledger Entry"."Global Dimension 1 Code") THEN
                        DivName := DimVal.Name
                    ELSE
                        DivName := '';

                    ExcelBuf.AddColumn(DivName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//16

                    //<<16Jan2019
                    i += 1;
                END;
            end;

            trigger OnPostDataItem()
            begin

                IF PrinttoExcel THEN BEGIN
                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Mar2018
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Mar2018
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//5 10Aug2017
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(GrandTotal, FALSE, '', TRUE, FALSE, TRUE, '0.00', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(TotalConvQty, FALSE, '', TRUE, FALSE, TRUE, '0.00', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(TotalQty, FALSE, '', TRUE, FALSE, TRUE, '0.00', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16Jan2019
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
                field(PrinttoExcel; PrinttoExcel)
                {
                    Caption = 'Print To Excel';
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
        // ExcelBuf.CreateBookAndOpenExcel('', 'Daily Production Register', '', '', '');
        IF printtoexcel THEN BEGIN
            // ExcBuffer.CreateBook('', 'Transfer Receipt Register');
            // ExcBuffer.CreateBookAndOpenExcel('', 'Transfer Receipt Register', '', '', USERID);
            // ExcBuffer.GiveUserControl;

            ExcBuffer.CreateNewBook('Daily Production Register');
            ExcBuffer.WriteSheet('Daily Production Register', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Daily Production Register', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();
        end;
    end;

    trigger OnPreReport()
    begin

        DateFilter := 'Daily Production Register as on ' + "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Posting Date");

        Location.RESET;
        Location.SETRANGE(Location.Code, "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code"));
        IF Location.FINDFIRST THEN
            LocationName := Location.Name;


        IF PrinttoExcel THEN BEGIN
            CreateXLSHEET;
            CreateHeader;
        END;
    end;

    var
        itemname: Text[100];
        item1: Record 27;
        companyinfo: Record 79;
        UserSetup: Record 91;
        Location: Record 14;
        LocLength: Integer;
        LocCode: Text[50];
        GrandTotal: Decimal;
        "*For excel": Text[30];
        // XLAPP: Automation;
        // XLWBOOK: Automation;
        //        XLSHEET: Automation;
        // XLRANGE: Automation;
        MaxCol: Text[3];
        i: Integer;
        CompInfo: Record 79;
        Loc: Text[30];
        CompName: Text[100];
        PrinttoExcel: Boolean;
        PostingGP: Text[30];
        TotalQty: Decimal;
        TotalConvQty: Decimal;
        Qty: Decimal;
        ProductionOrder: Record 5405;
        ILE: Record 32;
        QtyInLtr: Decimal;
        SpecificGravity: Decimal;
        DateFilter: Text[100];
        LocationName: Text[100];
        Qtytotal: Decimal;
        RECILE: Record 32;
        vBatchNo: Code[20];
        LineProd: Record 5406;
        vRoutingName: Text[50];
        recRouting: Record 99000763;
        "--ExcelBu": Integer;
        ExcelBuf: Record 370 temporary;
        "--------10Aug2017": Integer;
        Itm10: Record 27;
        HSNCode10: Code[10];
        "--------13Mar2018": Integer;
        VE13: Record 5802;
        BulkCode: Code[20];
        BulkDesc: Text;
        DivName: Text;
        DimVal: Record 349;
        ExcBuffer: Record 370 temporary;

    // //[Scope('Internal')]
    procedure CreateXLSHEET()
    begin
        /*
        CLEAR(XLAPP);
        CLEAR(XLWBOOK);
        CLEAR(XLSHEET);
        CREATE(XLAPP,FALSE,TRUE);
        XLWBOOK := XLAPP.Workbooks.Add;
        XLSHEET := XLWBOOK.Worksheets.Add;
        XLSHEET.Name := 'Daily Production Register';
        MaxCol:= 'M';
        */

    end;

    //  //[Scope('Internal')]
    procedure CreateHeader()
    begin
        companyinfo.GET;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(companyinfo.Name, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        CreateHdr('', TODAY);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(DateFilter, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        CreateHdr(USERID, 0D);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location Code :' + FORMAT(LocationName), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        //ExcelBuf.NewRow;
        //ExcelBuf.AddColumn('Date :'+FORMAT("Item Ledger Entry".GETFILTER("Posting Date")),FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);


        ExcelBuf.NewRow;
        //>>
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//13Mar2018
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//13Mar2018
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//05 10Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//16Jan2019

        //<<

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Document No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('HSN Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//05 10Aug2017
        ExcelBuf.AddColumn('Bulk Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Mar2018
        ExcelBuf.AddColumn('Bulk Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13Mar2018
        ExcelBuf.AddColumn('Batch No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty per Packing', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty in Ltr/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Chapter No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Specific Gravity', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Conver Qty in Kgs/Lts', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Kettle Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Division Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//16Jan2019
        i := 7;
    end;

    local procedure CreateHdr(UserID: Code[40]; PostingDate: Date)
    begin
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//05 10Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//16Jan2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        IF UserID = '' THEN
            ExcelBuf.AddColumn(FORMAT(PostingDate, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number)
        ELSE
            ExcelBuf.AddColumn(UserID, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
    end;
}


report 50040 "Pending CSO Order Date Wise"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PendingCSOOrderDateWise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Header"; "Sales Header")
        {
            DataItemTableView = SORTING("Posting Date")
                                ORDER(Ascending)
                                WHERE("Document Type" = FILTER(Order),
                                      Status = FILTER(Released));
            RequestFilterFields = "Posting Date", "Location Code";
            dataitem("BOdy SL"; "Sales Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document Type", "Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Closed = FILTER(false));

                trigger OnAfterGetRecord()
                begin

                    CLEAR(ProductGrade);
                    CLEAR(ProductGroup);
                    CLEAR(ProductGroupName);

                    recDefaultDimension.RESET;
                    recDefaultDimension.SETRANGE(recDefaultDimension."Table ID", 27);
                    recDefaultDimension.SETRANGE(recDefaultDimension."No.", "BOdy SL"."No.");
                    recDefaultDimension.SETRANGE(recDefaultDimension."Dimension Code", 'FOCUS');
                    IF recDefaultDimension.FINDFIRST THEN BEGIN
                        ProductGrade := recDefaultDimension."Dimension Value Code";
                    END;

                    recDefaultDimension.RESET;
                    recDefaultDimension.SETRANGE(recDefaultDimension."Table ID", 27);
                    recDefaultDimension.SETRANGE(recDefaultDimension."No.", "BOdy SL"."No.");
                    recDefaultDimension.SETFILTER(recDefaultDimension."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                    IF recDefaultDimension.FINDFIRST THEN BEGIN
                        ProductGroup := recDefaultDimension."Dimension Value Code";
                        recDimensionValue.RESET;
                        recDimensionValue.SETFILTER(recDimensionValue."Dimension Code", '%1|%2', 'PRODHIERAUTO', 'PRODHIERINDL');
                        recDimensionValue.SETRANGE(recDimensionValue.Code, ProductGroup);
                        IF recDimensionValue.FINDFIRST THEN
                            ProductGroupName := recDimensionValue.Name;
                    END;

                    IF ("BOdy SL".Quantity - "BOdy SL"."Quantity Shipped") = 0 THEN
                        CurrReport.SKIP;

                    CLEAR(BinQty);
                    recBinContent.RESET;
                    recBinContent.SETRANGE(recBinContent."Item No.", "BOdy SL"."No.");
                    recBinContent.SETFILTER(recBinContent."Location Code", 'PLANT01');
                    recBinContent.SETFILTER(recBinContent."Bin Code", '%1|%2', 'V-S-INDL', 'V-S-AUTO');
                    recBinContent.SETFILTER(recBinContent.Quantity, '<>%1', 0);
                    IF recBinContent.FIND('-') THEN BEGIN
                        recBinContent.CALCFIELDS(recBinContent.Quantity);
                        BinQty := BinQty + recBinContent.Quantity;
                    END;


                    RespCtr.GET("Sales Header"."Responsibility Center");

                    IF PrinttoExcel THEN BEGIN
                        ExcelBuf.NewRow;
                        ExcelBuf.AddColumn(SrNo, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(RespCtr.Name, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn("Sales Header"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
                        IF "BOdy SL"."Promised Delivery Date" <> 0D THEN BEGIN
                            ExcelBuf.AddColumn("BOdy SL"."Promised Delivery Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
                        END ELSE
                            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);


                        ExcelBuf.AddColumn("Sales Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn("Sales Header"."Sell-to Customer Name", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn("BOdy SL".Description, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        //ExcelBuf.AddColumn(CompInfo.Name,FALSE,'',FALSE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
                        IF Itemcategory.GET("BOdy SL"."Item Category Code") THEN;
                        ExcelBuf.AddColumn(Itemcategory.Description, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(ProductGroupName, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn(ProductGrade, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn("BOdy SL"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn("BOdy SL".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.00', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn("BOdy SL"."Quantity Shipped", FALSE, '', FALSE, FALSE, FALSE, '0.00', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn("BOdy SL".Quantity - "BOdy SL"."Quantity Shipped", FALSE, '', FALSE, FALSE, FALSE, '0.00', ExcelBuf."Cell Type"::Number);
                        ExcelBuf.AddColumn(BinQty / "BOdy SL"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.00', ExcelBuf."Cell Type"::Number);
                        //Fahim 29-05-23 for Resp. Center and Sales person
                        ExcelBuf.AddColumn("Sales Header"."Responsibility Center", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        ExcelBuf.AddColumn("Sales Header"."Salesperson Code", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        //Fahim 08-11-23
                        ExcelBuf.AddColumn("BOdy SL"."List Price", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
                        i += 1;
                    END;

                    QtyOrdTot := QtyOrdTot + SalesLine."Quantity (Base)";
                    DespatchedTot := DespatchedTot + SalesLine."Qty. Shipped (Base)";
                    BalanceTot := BalanceTot + SalesLine."Qty. to Ship (Base)";
                end;
            }

            trigger OnAfterGetRecord()
            begin
                SrNo := SrNo + 1;
            end;

            trigger OnPreDataItem()
            begin

                CompInfo.GET;
                IF Location.FINDFIRST THEN
                    Loc := Location.Name;

                CurrReport.CREATETOTALS(QtyOrdTot, DespatchedTot, BalanceTot);

                //>>09Mar2017 LocName
                LocCode1 := "Sales Header".GETFILTER("Sales Header"."Location Code");
                recLOC.RESET;
                recLOC.SETRANGE(Code, LocCode1);
                IF recLOC.FINDFIRST THEN BEGIN
                    LocName := recLOC.Name;
                END;
                //<< LocName


                IF PrinttoExcel THEN BEGIN
                    //CreateXLSHEET;
                    CreateHeader;
                    i := 7;
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
        // ExcelBuf.CreateBookAndOpenExcel('', 'Pending CSO Datewise', '', '', '');
        ExcBuffer.CreateNewBook('Pending CSO Datewise');
        ExcBuffer.WriteSheet('Pending CSO Datewise', CompanyName, UserId);
        ExcBuffer.CloseBook();
        ExcBuffer.SetFriendlyFilename(StrSubstNo('Pending CSO Datewise', CurrentDateTime, UserId));
        ExcBuffer.OpenExcel();
    end;

    trigger OnPreReport()
    begin

        DateFilter := "Sales Header".GETFILTER("Sales Header"."Order Date");
        locfilt := "Sales Header".GETFILTER("Sales Header"."Location Code");
        CompInfo.GET;

        filtertext := 'Pending CSO Datewise ' + ' For the Period of  ' + DateFilter;
    end;

    var
        CompInfo: Record 79;
        Location: Record 14;
        Loc: Text[30];
        UserSetup: Record 91;
        SalesHdr: Record 36;
        SalesLine: Record 37;
        StartDate: Date;
        EndDate: Date;
        SrNo: Integer;
        QtyOrdTot: Decimal;
        DespatchedTot: Decimal;
        BalanceTot: Decimal;
        DateFilter: Text[30];
        RespCtr: Record 5714;
        PrinttoExcel: Boolean;
        /*  XLAPP: Automation;
         XLWBOOK: Automation;
         XLSHEET: Automation;
         XLRANGE: Automation; */
        MaxCol: Text[30];
        Window: Dialog;
        filtertext: Text[100];
        i: Integer;
        TotalQtyShip: Decimal;
        Itemcategory: Record 5722;
        RecUserSetup: Record 91;
        RecUserSetup1: Record 91;
        locfilt: Code[30];
        recDefaultDimension: Record 352;
        ProductGrade: Text[30];
        ProductGroup: Text[30];
        ProductGroupName: Text[50];
        recDimensionValue: Record 349;
        LocCode: Code[10];
        LocResC: Code[10];
        recLocation: Record 14;
        CSOMapping1: Record 50006;
        BinQty: Decimal;
        recBinContent: Record 7302;
        "--": Integer;
        ExcelBuf: Record 370 temporary;
        "----09Mar2017": Integer;
        LocCode1: Code[100];
        recLOC: Record 14;
        LocName: Text[250];
        LocName1: Text[100];
        ExcBuffer: Record 370 temporary;

    // //[Scope('Internal')]
    procedure CreateXLSHEET()
    begin
        /* CLEAR(XLAPP);
        CLEAR(XLWBOOK);
        CLEAR(XLSHEET);
        CREATE(XLAPP, FALSE, TRUE);
        XLWBOOK := XLAPP.Workbooks.Add;
        XLSHEET := XLWBOOK.Worksheets.Add;
        XLSHEET.Name := 'Pending CSO Datewise';
        XLAPP.Visible(TRUE);
        MaxCol := 'O'; */
    end;

    //  //[Scope('Internal')]
    procedure CreateHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.NewRow;
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        // Fahim 29-05-23 for Responsibility center and Sales person code to be added
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        // Fahim 08-11-23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('Date :' + FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('Time :'+FORMAT(TIME),FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Pending CSO Order Date Wise', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017
        //ExcelBuf.AddColumn("Sales Header".GETFILTERS,FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        // Fahim 29-05-23 for Responsibility center and Sales person code to be added
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        // Fahim 08-11-23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);


        ExcelBuf.AddColumn('User Id :' + USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;

        //>>09Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017
        ExcelBuf.AddColumn("Sales Header".GETFILTER("Sales Header"."Posting Date"), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017
        ExcelBuf.AddColumn(LocName, FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017

        ExcelBuf.NewRow;
        //<<
        ExcelBuf.AddColumn('Sr. No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Branch', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//09Mar2017
        ExcelBuf.AddColumn('CSO Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Schedule Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('CSO No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Part Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Product Group Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Product Grade', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Pack Unnit of Measure', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty Order as per Pack Uom', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Despached as per Pack Uom', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Balance as per Pack Uom', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Product Availability', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        //Fahim 29-05-23 for Resp. center and Sales person code
        ExcelBuf.AddColumn('Responsibility Center', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Sales Person', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        //Fahim 08-11-23
        ExcelBuf.AddColumn('List Price', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;
}


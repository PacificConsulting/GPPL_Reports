report 50097 "Consumption Report Summary"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ConsumptionReportSummary.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            DataItemTableView = SORTING("Item No.")
                                ORDER(Ascending)
                                WHERE(Positive = FILTER(false),
                                      "Entry Type" = FILTER(Consumption),
                                      Quantity = FILTER(< 0));
            RequestFilterFields = "Posting Date", "Location Code", "Item Category Code";// "Product Group Code";
            column(Name_CompInfo; CompInfo.Name)
            {
            }
            column(ReportFilter; 'Raw Material Consumption Report for Period:' + ReportFilter)
            {
            }
            column(SrNo; SrNo)
            {
            }
            column(Desc_Item; Item.Description)
            {
            }
            column(varitemcategory; varitemcategory)
            {
            }
            column(UnitofMeasureCode_ItemLedgerEntry; "Item Ledger Entry"."Unit of Measure Code")
            {
            }
            column(Quantity_ItemLedgerEntry; "Item Ledger Entry".Quantity)
            {
            }
            column(ExciseProd_Item; 'Item."Excise Prod. Posting Group"')
            {
            }
            column(Name_Location; Location.Name)
            {
            }
            column(ItemNo_ItemLedgerEntry; "Item Ledger Entry"."Item No.")
            {
            }

            trigger OnAfterGetRecord()
            begin

                CLEAR(VarLotNo);
                ILE.RESET();
                //ILE.SETCURRENTKEY();
                ILE.SETRANGE(ILE."Document No.", "Item Ledger Entry"."Document No.");
                ILE.SETFILTER(ILE."Entry Type", '=%1', "Item Ledger Entry"."Entry Type"::Output);
                IF ILE.FINDFIRST THEN
                    VarLotNo := ILE."Lot No.";

                //>>RSPl/Migratiom/Rahul
                CLEAR(ILEQty);
                ILE.RESET;
                ILE.SETRANGE("Item No.", "Item No.");
                ILE.SETRANGE("Posting Date", FromDate, ToDate);
                ILE.SETRANGE("Entry Type", "Entry Type"::Consumption);
                //ILE.SETRANGE("Location Code","Item Ledger Entry".GETFILTER("Location Code"));
                ILE.SETFILTER(Quantity, '<%1', 0);
                ILE.SETFILTER(Positive, '%1', FALSE);
                IF ILE.FINDSET THEN
                    REPEAT
                        ILEQty += ILE.Quantity;
                    UNTIL ILE.NEXT = 0;
                //<<
                ProductionOrder.RESET;
                ProductionOrder.SETRANGE("No.", "Document No.");
                IF ProductionOrder.FINDFIRST THEN;


                Item.GET("Item Ledger Entry"."Item No.");

                varitemcategory := '';
                IF Item."Item Category Code" <> '' THEN BEGIN
                    IF itemcategory.GET(Item."Item Category Code") THEN
                        varitemcategory := itemcategory.Description;
                END;

                varproductgroup := '';
                // IF Item."Product Group Code" <> '' THEN BEGIN
                //     productgroup.RESET;
                //     productgroup.SETRANGE(Code, Item."Product Group Code");
                //     IF productgroup.FIND('-') THEN
                //         varproductgroup := productgroup.Description;
                // END;

                IF (PrintToExcel) AND (vItemNo <> "Item Ledger Entry"."Item No.") THEN
                    MakeExcelDataBody;
                vItemNo := "Item Ledger Entry"."Item No.";
            end;

            trigger OnPostDataItem()
            begin

                IF (PrintToExcel) THEN
                    MakeExcelDataFooter;
            end;

            trigger OnPreDataItem()
            begin
                CompInfo.GET;
                ReportFilter := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Posting Date");
                FromDate := "Item Ledger Entry".GETRANGEMIN("Item Ledger Entry"."Posting Date");
                ToDate := "Item Ledger Entry".GETRANGEMAX("Item Ledger Entry"."Posting Date");

                IF PrintToExcel THEN
                    MakeExcelInfo;
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

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    var
        CompInfo: Record 79;
        Usersetup: Record 91;
        ReportFilter: Text[50];
        date: Date;
        total: Decimal;
        Loc: Text[30];
        Location: Record 14;
        SrNo: Integer;
        Item: Record 27;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        itemcategory: Record 5722;
        // productgroup: Record 5723;
        productsubgroup: Record 50015;
        productgrade: Record 50016;
        varitemcategory: Text[200];
        varproductgroup: Text[200];
        varproductsubgroup: Text[200];
        varproductgrade: Text[200];
        VarLotNo: Code[20];
        ILE: Record 32;
        ProductionOrder: Record 5405;
        Text000: Label 'Data';
        Text001: Label 'Consumption Report Summary';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "---": Integer;
        ILEQty: Decimal;
        vItemNo: Code[30];
        FromDate: Date;
        ToDate: Date;

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('Consumption Report Summary', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"Consumption Report Summary", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        CompInfo.GET;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        Location.RESET;
        Location.SETRANGE(Code, "Item Ledger Entry".GETFILTER("Location Code"));
        IF Location.FINDFIRST THEN
            Loc := Location.Name;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location :' + Loc, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Sr. No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Quantity Issued', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Chapter Heading', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        SrNo := SrNo + 1;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(SrNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Item."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Item."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(varitemcategory, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn("Item Ledger Entry".Quantity,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Number);
        //ExcelBuf.AddColumn(ILEQty,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ILEQty, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', 0);//27Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(''/*Item."Excise Prod. Posting Group"*/, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
    end;

    //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        /* ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', '');
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text000);
        ExcelBuf.WriteSheet(Text000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        "Item Ledger Entry".CALCSUMS(Quantity);
        //ExcelBuf.AddColumn("Item Ledger Entry".Quantity,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("Item Ledger Entry".Quantity, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017
    end;
}


report 50095 "RM Stock Analysis DateWise"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/RMStockAnalysisDateWise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.")
                                WHERE(Blocked = FILTER(false),
                                     "Item Category Code" = FILTER(<> 'CAT02|CAT03'));
            RequestFilterFields = "No.", "Item Category Code";

            trigger OnAfterGetRecord()
            begin

                OpeningBalance := 0;
                ReceiptQuantity := 0;
                DespQuantity := 0;
                OtherQuantity := 0;
                ClosingBalance := 0;

                recItemCategory.RESET;
                recItemCategory.SETRANGE(recItemCategory.Code, Item."Item Category Code");
                IF recItemCategory.FINDFIRST THEN;

                //RSPL
                itemDim := '';
                recDefualtDim.RESET;
                recDefualtDim.SETRANGE(recDefualtDim."Table ID", 27);
                recDefualtDim.SETRANGE(recDefualtDim."No.", "No.");
                IF recDefualtDim.FINDFIRST THEN
                    itemDim := recDefualtDim."Dimension Value Code";
                //RSPL

                ILE.RESET;
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '<%1', RequestedDate);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        OpeningBalance := OpeningBalance + ILE.Quantity;
                    UNTIL ILE.NEXT = 0;
                END;

                ILE.RESET;
                ILE.SETCURRENTKEY("Location Code", "Posting Date", "Item No.", "Entry Type");
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                ILE.SETFILTER(ILE.Quantity, '>%1', 0);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        ReceiptQuantity := ReceiptQuantity + ILE.Quantity;
                    UNTIL ILE.NEXT = 0;
                END;

                ILE.RESET;
                ILE.SETCURRENTKEY("Location Code", "Posting Date", "Item No.", "Entry Type");
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                ILE.SETRANGE(ILE."Entry Type", 5);
                ILE.SETFILTER(ILE.Quantity, '<%1', 0);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        DespQuantity := DespQuantity + ABS(ILE.Quantity);
                    UNTIL ILE.NEXT = 0;
                END;

                ILE.RESET;
                ILE.SETCURRENTKEY("Location Code", "Posting Date", "Item No.", "Entry Type");
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                ILE.SETFILTER(ILE."Entry Type", '<>5');
                ILE.SETFILTER(ILE.Quantity, '<%1', 0);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        OtherQuantity := OtherQuantity + ABS(ILE.Quantity);
                    UNTIL ILE.NEXT = 0;
                END;

                ClosingBalance := (OpeningBalance + ReceiptQuantity) - (DespQuantity + OtherQuantity);

                IF (OpeningBalance = 0) AND (ReceiptQuantity = 0) AND (DespQuantity = 0) AND (OtherQuantity = 0) THEN
                    CurrReport.SKIP;


                SrNo := SrNo + 1;

                IF PrintToExcel THEN
                    MakeExcelDataBody;
            end;

            trigger OnPreDataItem()
            begin
                SrNo := 0;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(RequestedDate; RequestedDate)
                {
                    Caption = 'Starting Date';
                }
                field(RequestedDate1; RequestedDate1)
                {
                    Caption = 'Ending Date';
                }
                field(RequestedLocation; RequestedLocation)
                {
                    Caption = 'Location';
                    TableRelation = Location;
                }
                field(PrintToExcel; PrintToExcel)
                {
                    Caption = 'Export to Excel';
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

    trigger OnPreReport()
    begin
        CompanyInfo.GET;

        IF Item.GETFILTER(Item."Item Category Code") <> '' THEN BEGIN
            IF ((Item.GETFILTER(Item."Item Category Code") = 'CAT02') OR (Item.GETFILTER(Item."Item Category Code") = 'CAT03')) THEN BEGIN
                ERROR('You can not Select this Category');
            END;
        END;

        IF RequestedLocation = '' THEN BEGIN
            ERROR('Please Select Location');
        END;

        IF PrintToExcel THEN
            MakeExcelInfo;
    end;

    var
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        ILE: Record 32;
        OpeningBalance: Decimal;
        StartDate: array[6] of Date;
        EndDate: array[6] of Date;
        RequestedMonth: Option " ",Jan,Feb,Mar,April,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec;
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
        LastMonthQuantity: array[6] of Decimal;
        SalesInvoiceLine: Record 113;
        Monthly: array[6] of Text[30];
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        recItemCategory: Record 5722;
        LocResC: Code[100];
        recLocation: Record 14;
        CSOMapping: Record 50006;
        itemDim: Code[20];
        recDefualtDim: Record 352;
        Text000: Label 'Data';
        Text001: Label 'Location Wise Sale Value';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('STOCK ANALYSIS FOR RM', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"RM Stock Analysis DateWise", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
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
    procedure MakeExcelDataHeader()
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
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Stock Analysis From:', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT(RequestedDate), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('To', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT(RequestedDate1), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(RequestedLocation, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        //ExcelBuf.AddColumn('Item Category',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);//RSPL

        ExcelBuf.AddColumn('Opening Balance', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Receipt Qty.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Consumption Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Other Despatches', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Closing Balance', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Item."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Item."Sales Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(OpeningBalance,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(OpeningBalance, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn(ReceiptQuantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn(DespQuantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn(OtherQuantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn(ClosingBalance, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(OpeningBalance, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ReceiptQuantity, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(DespQuantity, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(OtherQuantity, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ClosingBalance, FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
    end;
}


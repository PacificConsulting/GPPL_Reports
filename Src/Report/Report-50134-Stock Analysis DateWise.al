report 50134 "Stock Analysis DateWise"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/StockAnalysisDateWise.rdl';
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

            trigger OnAfterGetRecord()
            begin

                //>>1

                OpeningBalance := 0;
                ReceiptQuantity := 0;
                DespQuantity := 0;
                OtherQuantity := 0;
                ClosingBalance := 0;

                recItemCategory.RESET;
                recItemCategory.SETRANGE(recItemCategory.Code, Item."Item Category Code");
                IF recItemCategory.FINDFIRST THEN;

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
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
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
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                ILE.SETFILTER(ILE."Entry Type", '<>1');
                ILE.SETFILTER(ILE.Quantity, '<%1', 0);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        OtherQuantity := OtherQuantity + ABS(ILE.Quantity);
                    UNTIL ILE.NEXT = 0;
                END;

                ClosingBalance := (OpeningBalance + ReceiptQuantity) - (DespQuantity + OtherQuantity);

                IF (OpeningBalance = 0) AND (ReceiptQuantity = 0) AND (DespQuantity = 0) AND (OtherQuantity = 0) THEN
                    CurrReport.SKIP;
                //<<1

                //>>2

                //Item, Body (3) - OnPreSection()
                SrNo := SrNo + 1;

                IF PrintToExcel THEN
                    MakeExcelDataBody;
                //<<2

                //>>3

                //Item, GroupFooter (4) - OnPreSection()
                IF NOT FooterPrinted THEN
                    LastFieldNo := CurrReport.TOTALSCAUSEDBY;

                // CurrReport.SHOWOUTPUT := NOT FooterPrinted;

                FooterPrinted := TRUE;
                //<<3
            end;

            trigger OnPreDataItem()
            begin

                //>>1
                CurrReport.CREATETOTALS(OpeningBalance, ReceiptQuantity, DespQuantity, OtherQuantity, ClosingBalance);
                SrNo := 0;
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
                field("Start Date"; RequestedDate)
                {
                    ApplicationArea = ALL;
                }
                field("End Date"; RequestedDate1)
                {
                    ApplicationArea = ALL;
                }
                field(Location; RequestedLocation)
                {
                    ApplicationArea = ALL;
                    TableRelation = Location;
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = ALL;
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

    trigger OnInitReport()
    begin

        //>>1
        CompanyInfo.GET;
        //<<1
    end;

    trigger OnPostReport()
    begin

        //>>1
        IF PrintToExcel THEN
            CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1

        IF Item.GETFILTER(Item."Item Category Code") <> '' THEN BEGIN
            IF NOT ((Item.GETFILTER(Item."Item Category Code") = 'CAT02') OR (Item.GETFILTER(Item."Item Category Code") = 'CAT03')) THEN BEGIN
                ERROR('You can not Select this Category');
            END;
        END;

        IF RequestedLocation = '' THEN BEGIN
            ERROR('Please Select Location');
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
        *///Commented 07Mar2017

        //>>09Mar2017 LocName
        LocCode := RequestedLocation;
        recLOC.RESET;
        recLOC.SETRANGE(Code, LocCode);
        IF recLOC.FINDFIRST THEN BEGIN
            LocName := recLOC.Name;
        END;
        //<< LocName


        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<1

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
        User: Record 2000000120;
        LocResC: Code[100];
        recLocation: Record 14;
        CSOMapping: Record 50006;
        Text000: Label 'Data';
        Text001: Label 'Stock Analysis Datewise';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "----09Mar2017": Integer;
        LocCode: Code[100];
        recLOC: Record 14;
        LocName: Text[250];
        LocName1: Text[100];

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//07Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn('STOCK ANALYSIS DATEWISE', FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Stock Analysis DateWise", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>07Mar2017 RB-N
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

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>07Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('Stock Analysis Datewise', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.NewRow;
        //<<
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(FORMAT(RequestedDate) + ' .. ' + FORMAT(RequestedDate1), FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        //ExcelBuf.AddColumn('To',FALSE,'',TRUE,FALSE,FALSE,'@');
        //ExcelBuf.AddColumn(FORMAT(RequestedDate1),FALSE,'',TRUE,FALSE,FALSE,'@');
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(LocName, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        //ExcelBuf.AddColumn(RequestedLocation,FALSE,'',TRUE,FALSE,FALSE,'@',1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item No', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('Sale UOM', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('Opening Balance', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('Receipt Qty.', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('Sold Qty', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('Other Despatches', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('Closing Balance', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Item."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(Item."Sales Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(OpeningBalance, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
        ExcelBuf.AddColumn(ReceiptQuantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
        ExcelBuf.AddColumn(DespQuantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
        ExcelBuf.AddColumn(OtherQuantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
        ExcelBuf.AddColumn(ClosingBalance, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(OpeningBalance, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
        ExcelBuf.AddColumn(ReceiptQuantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
        ExcelBuf.AddColumn(DespQuantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
        ExcelBuf.AddColumn(OtherQuantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
        ExcelBuf.AddColumn(ClosingBalance, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);
    end;
}


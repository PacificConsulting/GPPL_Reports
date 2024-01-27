report 50023 "Vendor Item Analysis"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/VendorItemAnalysis.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Vendor"; "Item Vendor")
        {
            DataItemTableView = SORTING("Item No.", "Variant Code", "Vendor No.");
            RequestFilterFields = "Item No.", "Vendor No.";
            dataitem("Purch. Rcpt. Line"; "Purch. Rcpt. Line")
            {
                DataItemLink = "No." = FIELD("Item No."),
                               "Buy-from Vendor No." = FIELD("Vendor No.");
                DataItemTableView = SORTING("No.", "Buy-from Vendor No.", "Location Code")
                                    WHERE(Type = FILTER(Item));
                RequestFilterFields = "Posting Date";

                trigger OnAfterGetRecord()
                begin
                    ItmCount += 1;//23Feb2017

                    //>>1

                    recItem.RESET;
                    recItem.SETRANGE(recItem."No.", "Purch. Rcpt. Line"."No.");
                    IF recItem.FINDFIRST THEN;

                    recVendor.RESET;
                    recVendor.SETRANGE(recVendor."No.", "Purch. Rcpt. Line"."Buy-from Vendor No.");
                    IF recVendor.FINDFIRST THEN;

                    recLocation.RESET;
                    recLocation.SETRANGE(recLocation.Code, "Purch. Rcpt. Line"."Location Code");
                    IF recLocation.FINDFIRST THEN;


                    /*LastPirce := 0;
                    recPIL.RESET;
                    recPIL.SETRANGE(recPIL."No.","Purch. Rcpt. Line"."No.");
                    recPIL.SETRANGE(recPIL."Buy-from Vendor No.","Purch. Rcpt. Line"."Buy-from Vendor No.");
                    recPIL.SETRANGE(recPIL."Expected Receipt Date","Purch. Rcpt. Line"."Posting Date");
                    IF recPIL.FINDLAST THEN
                    BEGIN
                     LastPirce := recPIL."Direct Unit Cost";
                    END;*/ //EBT Commented 24/4/14
                           //EBT Added 24/4/14//
                    BlanketorderDirectCost := 0;
                    RecPurchline.RESET;
                    RecPurchline.SETRANGE(RecPurchline."Document Type", RecPurchline."Document Type"::"Blanket Order");
                    RecPurchline.SETRANGE(RecPurchline."Document No.", "Purch. Rcpt. Line"."Blanket Order No.");
                    RecPurchline.SETRANGE(RecPurchline."Line No.", "Purch. Rcpt. Line"."Blanket Order Line No.");
                    RecPurchline.SETRANGE(RecPurchline."No.", "Purch. Rcpt. Line"."No.");
                    IF RecPurchline.FINDFIRST THEN
                        REPEAT
                            BlanketorderDirectCost := RecPurchline."Direct Unit Cost";
                        UNTIL RecPurchline.NEXT = 0;
                    //EBT Added 24/4/14//
                    //<<

                    //>>3
                    MakeExcelDataBody;
                    //<<

                    //>>23Feb2017 GroupFooter
                    IF ItmCount = PRLCount THEN BEGIN
                        MakeExcelDataFooter;
                        ItmCount := 0;

                    END;

                    //<<

                end;

                trigger OnPreDataItem()
                begin
                    //MESSAGE('No.of Count %1',COUNT);
                    PRLCount := COUNT;//23Feb2017
                    ItmCount := 0;//23Feb2017
                end;
            }
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
        CreateExcelbook;
    end;

    trigger OnPreReport()
    begin
        MakeExcelInfo;
    end;

    var
        ExcelBuf: Record 370 temporary;
        recItem: Record 27;
        recVendor: Record 23;
        recLocation: Record 14;
        recPIL: Record 123;
        LastPirce: Decimal;
        RecPurchline: Record 39;
        BlanketorderDirectCost: Decimal;
        Text002: Label 'Data';
        Text003: Label 'Vendor Item Analysis';
        Text004: Label 'Company Name';
        Text005: Label 'Report No.';
        Text006: Label 'Report Name';
        Text007: Label 'User ID';
        Text008: Label 'Date';
        Text009: Label 'Location Filter';
        "--22Feb2017": Integer;
        TQty: Decimal;
        ItmCount: Integer;
        PRLCount: Integer;
        recPRL: Record 121;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text004), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text006), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(FORMAT(Text003), FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text005), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Vendor Item Analysis", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text007), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.NewRow;
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text002,Text003,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');
        //>>22Feb2017 RB-N
        // ExcelBuf.CreateBook('', Text003);
        // //ExcelBuf.CreateBookAndOpenExcel('',Text000,Text001,COMPANYNAME,USERID);
        // ExcelBuf.CreateBookAndOpenExcel('', Text003, '', '', USERID);
        // ExcelBuf.GiveUserControl;

        ExcelBuf.CreateNewBook(Text003);
        ExcelBuf.WriteSheet(Text003, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text003, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>23Feb2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text003, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.NewRow;


        //<<

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Vendor No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Vendor Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Vendor Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Blanket Order No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Blanket Order Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('GRN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('GRN Date.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Purchase Order Direct Unit Cost', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        //ExcelBuf.AddColumn('Last Price',FALSE,'',TRUE,FALSE,TRUE,'@'); //EBT Commented
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(recItem.Description, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(recVendor."Full Name", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Vendor Item No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Blanket Order No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Order Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.AddColumn(BlanketorderDirectCost, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); //EBT Added 24/4/14
        ExcelBuf.AddColumn("Purch. Rcpt. Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        //>>23Feb2017

        TQty := TQty + "Purch. Rcpt. Line".Quantity;

        //<<
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //ExcelBuf.AddColumn(LastPirce,FALSE,'',FALSE,FALSE,FALSE,'');  //EBT Commented
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(recItem.Description, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Buy-from Vendor No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(recVendor."Full Name", FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        //ExcelBuf.AddColumn("Purch. Rcpt. Line".Quantity,FALSE,'',TRUE,FALSE,FALSE,'#,####0',0);
        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, FALSE, '#,####0', 0);//23Feb2017
        TQty := 0; //23Feb2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
    end;
}


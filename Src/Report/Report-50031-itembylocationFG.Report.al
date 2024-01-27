report 50031 "item by location (FG)"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 04Jan2018   RB-N         Single Location Filter
    // 20Jul2018   RB-N         Skipping Include in Valuation Location -- False
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/itembylocationFG.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("Inventory Posting Group")
                                WHERE("Inventory Posting Group" = FILTER('INDOILS | AUTOOILS | BOILS | RPO | TOILS | REPSOL | CEPSA'),
                                      "Item Tracking Code" = FILTER('FGIL|FGAL'));
            RequestFilterFields = "No.", "Item Category Code"; //"Product Group Code";

            trigger OnAfterGetRecord()
            begin

                //>>1

                IntI := 1;

                ItemUOM.GET(Item."No.", Item."Sales Unit of Measure");
                TotQTY := 0;
                ExcelBuf.NewRow;

                ExcelBuf.AddColumn("No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn(Description, FALSE, '', TRUE, FALSE, FALSE, '', 1);

                remarks := '';
                recItemTable.RESET;
                recItemTable.SETRANGE(recItemTable."No.", "No.");
                recItemTable.SETRANGE(recItemTable.Blocked, TRUE);
                if recItemTable.FINDFIRST THEN BEGIN
                    remarks := 'Blocked';
                END;
                ExcelBuf.AddColumn(remarks, FALSE, '', TRUE, FALSE, FALSE, '', 1);

                recItem.RESET;
                recItem.SETRANGE("No.", "No.");
                IF recItem.FIND('-') THEN BEGIN

                    //calculate inventory balance by location and date
                    IF LocFilter04 = '' THEN BEGIN
                        reclocation.RESET;
                        reclocation.SETRANGE(reclocation.Closed, FALSE);
                        reclocation.SETRANGE("Include in Valuation Report", TRUE);//20Jul2018
                        IF reclocation.FIND('-') THEN
                            REPEAT
                                recItem.RESET;
                                recItem.SETRANGE("No.", "No.");
                                recItem.SETRANGE("Location Filter", reclocation.Code);
                                recItem.SETFILTER("Date Filter", '<=%1', TillDate);
                                IF recItem.FINDFIRST THEN BEGIN
                                    recItem.CALCFIELDS("Inventory Change");
                                    TotQTY := TotQTY + recItem."Inventory Change";
                                    ExcelBuf.AddColumn(recItem."Inventory Change" / ItemUOM."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);
                                    ExcelBuf.AddColumn(recItem."Inventory Change", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);
                                    ArrLoctotSOM[IntI] := ArrLoctotSOM[IntI] + recItem."Inventory Change" / ItemUOM."Qty. per Unit of Measure";
                                    ArrLoctot[IntI] := ArrLoctot[IntI] + recItem."Inventory Change";
                                    IntI := IntI + 1;
                                END;
                            UNTIL reclocation.NEXT = 0;
                    END ELSE BEGIN

                        reclocation.RESET;
                        reclocation.SETFILTER(Code, LocFilter04);
                        //reclocation.SETRANGE(reclocation.Closed,FALSE);
                        IF reclocation.FINDFIRST THEN BEGIN
                            recItem.RESET;
                            recItem.SETRANGE("No.", "No.");
                            recItem.SETRANGE("Location Filter", reclocation.Code);
                            recItem.SETFILTER("Date Filter", '<=%1', TillDate);
                            IF recItem.FINDFIRST THEN BEGIN
                                recItem.CALCFIELDS("Inventory Change");
                                TotQTY := TotQTY + recItem."Inventory Change";
                                ExcelBuf.AddColumn(recItem."Inventory Change" / ItemUOM."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);
                                ExcelBuf.AddColumn(recItem."Inventory Change", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);
                                ArrLoctotSOM[IntI] := ArrLoctotSOM[IntI] + recItem."Inventory Change" / ItemUOM."Qty. per Unit of Measure";
                                ArrLoctot[IntI] := ArrLoctot[IntI] + recItem."Inventory Change";
                                IntI := IntI + 1;
                            END;
                        END;
                    END;
                END;

                ExcelBuf.AddColumn(TotQTY / ItemUOM."Qty. per Unit of Measure", FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                ExcelBuf.AddColumn(TotQTY, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);

                //<<1
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                IntI := 1;
                GrandTotal := 0;
                ExcelBuf.NewRow;
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);

                IF LocFilter04 = '' THEN BEGIN
                    reclocation.RESET;
                    reclocation.SETRANGE(reclocation.Closed, FALSE);
                    reclocation.SETRANGE("Include in Valuation Report", TRUE);//20Jul2018
                    IF reclocation.FIND('-') THEN
                        REPEAT
                            ExcelBuf.AddColumn(ArrLoctotSOM[IntI], FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                            ExcelBuf.AddColumn(ArrLoctot[IntI], FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                            GrandTotalSOM := GrandTotalSOM + ArrLoctotSOM[IntI];
                            GrandTotal := GrandTotal + ArrLoctot[IntI];
                            IntI := IntI + 1;
                        UNTIL reclocation.NEXT = 0;
                END ELSE BEGIN
                    reclocation.RESET;
                    reclocation.SETFILTER(Code, LocFilter04);
                    //reclocation.SETRANGE(reclocation.Closed,FALSE);
                    IF reclocation.FINDFIRST THEN BEGIN
                        ExcelBuf.AddColumn(ArrLoctotSOM[IntI], FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                        ExcelBuf.AddColumn(ArrLoctot[IntI], FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                        GrandTotalSOM := GrandTotalSOM + ArrLoctotSOM[IntI];
                        GrandTotal := GrandTotal + ArrLoctot[IntI];
                        IntI := IntI + 1;
                    END;
                END;

                ExcelBuf.AddColumn(GrandTotalSOM, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                ExcelBuf.AddColumn(GrandTotal, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);

                //<<1
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                companyinfo.GET;

                CLEAR(ArrLoctot);
                CLEAR(ArrLoctotSOM);

                //get the location name for column heading
                IF LocFilter04 = '' THEN BEGIN
                    reclocation.RESET;
                    reclocation.SETRANGE(reclocation.Closed, FALSE);
                    reclocation.SETRANGE("Include in Valuation Report", TRUE);//20Jul2018
                    IF reclocation.FIND('-') THEN
                        REPEAT
                            ExcelBuf.AddColumn(reclocation.Name + 'Qty in Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
                            ExcelBuf.AddColumn(reclocation.Name + 'Qty in Base UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
                        UNTIL reclocation.NEXT = 0;
                END ELSE BEGIN
                    reclocation.RESET;
                    reclocation.SETFILTER(Code, LocFilter04);
                    //reclocation.SETRANGE(reclocation.Closed,FALSE);
                    IF reclocation.FINDFIRST THEN BEGIN
                        ExcelBuf.AddColumn(reclocation.Name + 'Qty in Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);
                        ExcelBuf.AddColumn(reclocation.Name + 'Qty in Base UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);
                    END;
                END;

                ExcelBuf.AddColumn('TOTAL QTY in Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
                ExcelBuf.AddColumn('TOTAL QTY in Base UOM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017

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
                field("Till Date"; TillDate)
                {
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

        //>>2
        CreateExcelbook;
        //<<2
    end;

    trigger OnPreReport()
    begin

        //>>RB-N 04Jan2018
        LocFilter04 := '';
        LocFilter04 := Item.GETFILTER("Location Filter");

        //<<RB-N 04Jan2018

        //>>1
        IF Item.GETFILTER("Item Category Code") <> '' THEN BEGIN
            IF itemcategory.GET(Item.GETFILTER("Item Category Code")) THEN;
            Filter := itemcategory.Description;
        END;

        /*  IF Item.GETFILTER("Product Group Code") <> '' THEN BEGIN
             productgroup.RESET;
             productgroup.SETRANGE(Code, Item.GETFILTER("Product Group Code"));
             IF productgroup.FIND('-') THEN;
             filter1 := productgroup.Description;
     END; *///28dec2023

        MakeExcelInfo;

        //<<1
    end;

    var
        reclocation: Record 14;
        recile: Record 32;
        ArrLoctot: array[30000] of Decimal;
        ArrLoctotSOM: array[3000] of Decimal;
        IntI: Integer;
        recItem: Record 27;
        vartext: Text[30];
        vardecimal: Decimal;
        ExcelBuf: Record 370 temporary;
        companyinfo: Record 79;
        TillDate: Date;
        TotQTY: Decimal;
        GrandTotalSOM: Decimal;
        GrandTotal: Decimal;
        itemcategory: Record 5722;
        // productgroup: Record 5723;//28dec2023
        productsubgroup: Record 50015;
        productgrade: Record 50016;
        "Filter": Text[100];
        filter1: Text[100];
        filter2: Text[100];
        filter3: Text[100];
        usersetup: Record 91;
        ItemUOM: Record 5404;
        ILE: Record 32;
        remarks: Text[30];
        recItemTable: Record 27;
        Text001: Label 'As of %1';
        Text000: Label 'Data';
        Text0001: Label 'Item by Location FG';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "-------------------04Jan2018": Integer;
        LocFilter04: Code[10];

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//13Apr2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Item By Location FG', FALSE, FALSE, FALSE, FALSE, '', 1);//13Apr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"item by location (FG)", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn('Item Category', FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(itemcategory.Description, FALSE, FALSE, FALSE, FALSE, '', 1);

        ExcelBuf.ClearNewRow;

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item By Location FG', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017

        //ExcelBuf.AddColumn('Item By Location as on',FALSE,'',TRUE,FALSE,FALSE,'@');
        ExcelBuf.NewRow;//13Apr2017
        ExcelBuf.NewRow;//13Apr2017
        ExcelBuf.AddColumn('As on  ' + FORMAT(TillDate), FALSE, '', TRUE, FALSE, FALSE, '', 1);//13Apr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
    end;

    ////[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
    end;

    ////[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        /*   //>>13Apr2017
          ExcelBuf.CreateBook('', Text0001);
          ExcelBuf.CreateBookAndOpenExcel('', Text0001, '', '', USERID);
          ExcelBuf.GiveUserControl; *///28dec2023
        //<<13Apr2017
        ExcelBuf.CreateNewBook(Text0001);
        ExcelBuf.WriteSheet(Text0001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text0001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
    end;

    ////[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
    end;
}


report 50032 "item by location (RM)"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/itembylocationRM.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("Inventory Posting Group")
                                WHERE("Inventory Posting Group" = FILTER('ADDITIVES | BARRELS | BASEOILS | CHEMICALS | MOULDS | SEMIFG | SMALLPACKS | BULK | REPSOL'));
            RequestFilterFields = "No.", "Item Category Code";//"Product Group Code";

            trigger OnAfterGetRecord()
            begin

                //>>1

                IntI := 1;

                //RSPL
                itemDim := '';
                recDefualtDim.RESET;
                recDefualtDim.SETRANGE(recDefualtDim."Table ID", 27);
                recDefualtDim.SETRANGE(recDefualtDim."No.", Item."No.");
                IF recDefualtDim.FINDFIRST THEN
                    itemDim := recDefualtDim."Dimension Value Code";
                //RSPL

                //skip blocked items  and raw materials
                TotQTY := 0;
                ExcelBuf.NewRow;
                ExcelBuf.AddColumn("No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn(Description, FALSE, '', TRUE, FALSE, FALSE, '', 1);
                //RSPL
                ExcelBuf.AddColumn(itemDim, FALSE, '', TRUE, FALSE, FALSE, '', 1);
                //RSPL

                recItem.RESET;
                recItem.SETRANGE("No.", "No.");
                IF recItem.FIND('-') THEN BEGIN

                    //calculate inventory balance by location and date
                    reclocation.RESET;
                    reclocation.SETRANGE(reclocation.Closed, FALSE);
                    IF reclocation.FIND('-') THEN
                        REPEAT
                            recItem.RESET;
                            recItem.SETRANGE("No.", "No.");
                            recItem.SETRANGE("Location Filter", reclocation.Code);
                            recItem.SETFILTER("Date Filter", '<=%1', TillDate);
                            IF recItem.FINDFIRST THEN BEGIN
                                recItem.CALCFIELDS("Inventory Change");
                                TotQTY := TotQTY + recItem."Inventory Change";
                                ExcelBuf.AddColumn(recItem."Inventory Change", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//13Apr2017
                                ArrLoctot[IntI] := ArrLoctot[IntI] + recItem."Inventory Change";
                                IntI := IntI + 1;
                            END;
                        UNTIL reclocation.NEXT = 0;

                END;

                ExcelBuf.AddColumn(TotQTY, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//13Apr2017
                //<<1
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                IntI := 1;
                GrandTotal := 0;
                ExcelBuf.NewRow;
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//RSPL
                ExcelBuf.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                reclocation.RESET;
                reclocation.SETRANGE(reclocation.Closed, FALSE);
                IF reclocation.FIND('-') THEN
                    REPEAT
                        ExcelBuf.AddColumn(ArrLoctot[IntI], FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//13Apr2017
                        GrandTotal := GrandTotal + ArrLoctot[IntI];
                        IntI := IntI + 1;
                    UNTIL reclocation.NEXT = 0;
                ExcelBuf.AddColumn(GrandTotal, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//13Apr2017
                //<<1
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                companyinfo.GET;

                CLEAR(ArrLoctot);
                //get the location name for column heading

                /*reclocation.SETCURRENTKEY("Responsibility Center");
                
                IF reclocation.FIND('-') THEN REPEAT
                     RespCtr:=reclocation."Responsibility Center";
                     IF OldRespCtr=RespCtr THEN
                           ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'')
                     ELSE IF OldRespCtr<>RespCtr THEN BEGIN
                         IF "Responsiblity Center".GET(RespCtr) THEN;
                           OldRespCtr:=RespCtr;
                           ExcelBuf.AddColumn("Responsiblity Center".Name,FALSE,'',TRUE,FALSE,FALSE,'');
                     END ELSE;
                UNTIL  reclocation.NEXT=0; */

                ExcelBuf.NewRow;
                ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
                ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
                ExcelBuf.AddColumn('Defualt Dimension', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017 //RSPL

                reclocation.RESET;
                reclocation.SETRANGE(reclocation.Closed, FALSE);
                IF reclocation.FIND('-') THEN
                    REPEAT
                        ExcelBuf.AddColumn(reclocation.Name, FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
                    UNTIL reclocation.NEXT = 0;
                ExcelBuf.AddColumn('TOTAL QTY', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13Apr2017
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

        //>>1
        CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin
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

         END; */

        MakeExcelInfo;
        //<<1
    end;

    var
        Text001: Label 'As of %1';
        Text000: Label 'Data';
        Text0001: Label 'Item by Location RM';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        reclocation: Record 14;
        recile: Record 32;
        ArrLoctot: array[30000] of Decimal;
        IntI: Integer;
        recItem: Record 27;
        vartext: Text[30];
        vardecimal: Decimal;
        ExcelBuf: Record 370 temporary;
        companyinfo: Record 79;
        TillDate: Date;
        TotQTY: Decimal;
        GrandTotal: Decimal;
        itemcategory: Record 5722;
        // productgroup: Record 5723;
        productsubgroup: Record 50015;
        productgrade: Record 50016;
        "Filter": Text[100];
        filter1: Text[100];
        filter2: Text[100];
        filter3: Text[100];
        OldRespCtr: Text[30];
        RespCtr: Text[30];
        "Responsiblity Center": Record 5714;
        itemDim: Code[20];
        recDefualtDim: Record 352;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//13Apr2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Item By Location RM', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"item by location (RM)", FALSE, FALSE, FALSE, FALSE, '', 0);//13Apr2017
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn('Item Category', FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(itemcategory.Description, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;

        ExcelBuf.ClearNewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item By Location RM', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017

        ExcelBuf.NewRow;//13Apr2017
        ExcelBuf.NewRow;//13Apr2017
        ExcelBuf.AddColumn('Date : ' + FORMAT(TillDate), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13Apr2017
        //ExcelBuf.AddColumn(TillDate,FALSE,'',TRUE,FALSE,FALSE,'',1);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>13Apr2017
        /*  ExcelBuf.CreateBook('', Text0001);
         ExcelBuf.CreateBookAndOpenExcel('', Text0001, '', '', USERID);
         ExcelBuf.GiveUserControl; */
        //<<13Apr2017
        ExcelBuf.CreateNewBook(Text0001);
        ExcelBuf.WriteSheet(Text0001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text0001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
    end;
}


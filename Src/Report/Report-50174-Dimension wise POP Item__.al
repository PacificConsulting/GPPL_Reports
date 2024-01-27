report 50174 "Dimension wise POP Item__"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/DimensionwisePOPItem.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Location; Location)
        {
            RequestFilterFields = "Code";
            dataitem(Item; Item)
            {
                DataItemLink = "Location Filter" = FIELD(Code);
                DataItemTableView = SORTING("No.")
                                    ORDER(Ascending)
                                    WHERE("Item Category Code" = FILTER('CAT14'));
                RequestFilterFields = "No.";
                dataitem("Dimension Set Entry"; "Dimension Set Entry")
                {
                    CalcFields = "Dimension Value Name";
                    DataItemTableView = SORTING("Dimension Value ID")
                                        ORDER(Ascending)
                                        WHERE("Dimension Code" = FILTER('MERCH'));

                    trigger OnAfterGetRecord()
                    begin

                        //>>24Apr2017
                        TCount += 1;

                        recDim24.SETCURRENTKEY("Dimension Value ID");
                        recDim24.SETRANGE("Dimension Value Code", "Dimension Value Code");
                        //recDim24.SETRANGE("Dimension Value ID","Dimension Value ID");
                        IF recDim24.FINDLAST THEN BEGIN
                            ICount := recDim24.COUNT;
                        END;
                        //<<24Apr2017


                        //>> QuantityTotal
                        ILE24Apr.RESET;
                        ILE24Apr.SETCURRENTKEY("Item No.");
                        ILE24Apr.SETRANGE("Item No.", Item."No.");
                        ILE24Apr.SETRANGE("Location Code", Location.Code);
                        ILE24Apr.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        IF ILE24Apr.FINDSET THEN
                            REPEAT

                                vQuantity += ILE24Apr.Quantity;
                            UNTIL ILE24Apr.NEXT = 0;


                        //>> QuantityTotal

                        IF TCount = ICount THEN BEGIN

                            IF vQuantity <> 0 THEN BEGIN
                                ExcelBuff.NewRow;
                                //IF recDimension.GET('MERCH',"Dimension Set Entry"."Dimension Value Code") THEN;
                                //recDimension.GET('MERCH',"Ledger Entry Dimension"."Dimension Value Code");
                                ExcelBuff.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
                                ExcelBuff.AddColumn("Dimension Set Entry"."Dimension Value Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
                                ExcelBuff.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
                                ExcelBuff.AddColumn(Location.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
                                ExcelBuff.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
                                ExcelBuff.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6

                                ExcelBuff.AddColumn(vQuantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//7

                                ExcelBuff.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8

                            END;
                            //MESSAGE('Dimension Value Id %1 \ TCOUNT %2',"Dimension Value ID",TCount);
                            TCount := 0;
                            vQuantity := 0;
                        END;
                    end;

                    trigger OnPostDataItem()
                    begin


                        //>>1
                        /*
                        //Ledger Entry Dimension, GroupFooter (1) - OnPreSection()
                        IF vtotalQty<>0 THEN BEGIN
                          ExcelBuff.NewRow;
                          IF recDimension.GET('MERCH',"Dimension Set Entry"."Dimension Value Code") THEN;
                          //recDimension.GET('MERCH',"Ledger Entry Dimension"."Dimension Value Code");
                          ExcelBuff.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//1
                          ExcelBuff.AddColumn(recDimension.Name,FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                          ExcelBuff.AddColumn(Item.Description,FALSE,'',FALSE,FALSE,FALSE,'',1);//3
                          ExcelBuff.AddColumn(Location.Name,FALSE,'',FALSE,FALSE,FALSE,'',1);//4
                          ExcelBuff.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//5
                          ExcelBuff.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//6
                        
                          ExcelBuff.AddColumn(vtotalQty,FALSE,'',FALSE,FALSE,FALSE,'0.000',0);//7
                        
                          ExcelBuff.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//8
                        END;
                        
                        vtotalQty:=0;
                        */
                        //<<1

                    end;

                    trigger OnPreDataItem()
                    begin

                        recDim24.COPYFILTERS("Dimension Set Entry");//24Apr2017
                        TCount := 0;//24Apr2017
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    Item.CALCFIELDS(Inventory);//24Apr2017
                                               //>>1

                    RecItemCategory.RESET;
                    RecItemCategory.SETRANGE(RecItemCategory.Code, Item."Item Category Code");

                    IF RecItemCategory.FINDSET THEN
                        CategoryDesc := RecItemCategory.Description;

                    opTotQty := 0;
                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.", "Location Code");
                    recILE.SETRANGE("Item No.", "No.");
                    recILE.SETRANGE("Location Code", Location.Code);
                    IF recILE.FINDSET THEN
                        REPEAT
                            opTotQty += recILE.Quantity;
                        UNTIL recILE.NEXT = 0;

                    CLEAR(NNN); //24Apr2017
                    IF opTotQty = 0 THEN
                        NNN := TRUE;//24Apr2017
                                    //CurrReport.SKIP;

                    //>>RSPL-RAHUL
                    //ILE.SETCURRENTKEY(ILE."Posting Date");
                    ILE.RESET;
                    ILE.SETCURRENTKEY("Item No.", "Location Code");//24APr2017
                    ILE.SETRANGE(ILE."Item No.", "No.");
                    ILE.SETRANGE(ILE."Location Code", Location.Code);
                    ILE.SETRANGE(ILE."Unit of Measure Code", 'NOS');
                    ILE.SETFILTER(ILE.Quantity, '>0');

                    ILE.SETRANGE(ILE.Positive, TRUE);
                    IF ILE.FINDLAST THEN
                        PostingDate := ILE."Posting Date";
                    //<<

                    //<<1


                    //>>2

                    //Item, Body (2) - OnPreSection()
                    IF NOT NNN THEN BEGIN

                        ExcelBuff.NewRow;
                        SrNo += 1;
                        ExcelBuff.AddColumn(SrNo, FALSE, '', TRUE, FALSE, FALSE, '', 0);//1
                        ExcelBuff.AddColumn("No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
                        ExcelBuff.AddColumn(Description, FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
                        ExcelBuff.AddColumn(Location.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
                                                                                                 //commented parag
                                                                                                 //ExcelBuff.AddColumn("Item Category Code",FALSE,'',TRUE,FALSE,FALSE,'');
                                                                                                 // added to show item category description
                        ExcelBuff.AddColumn(CategoryDesc, FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
                        ExcelBuff.AddColumn("Base Unit of Measure", FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
                        ExcelBuff.AddColumn(Inventory, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//7

                        ItemDesc := Description;
                        LocName := Location.Name;

                    END;
                    //<<2
                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    IF AsOnDate <> 0D THEN
                        SETRANGE("Date Filter", 0D, AsOnDate);
                    // "Item Ledger Entry".SETFILTER("Item Ledger Entry"."Posting Date",'%1..%2',0D,AsOnDate);
                    //<<1
                end;
            }

            trigger OnPreDataItem()
            begin

                //>>1
                //Location, Header (1) - OnPreSection()
                ExcelBuff.NewRow;
                ExcelBuff.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
                ExcelBuff.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//8


                ExcelBuff.NewRow;
                ExcelBuff.AddColumn('POP Materials Stock', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
                ExcelBuff.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8

                ExcelBuff.NewRow;
                ExcelBuff.NewRow;
                ExcelBuff.AddColumn('As on : ' + FORMAT(AsOnDate), FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
                ExcelBuff.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8

                ExcelBuff.NewRow;
                ExcelBuff.AddColumn('SrNo', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
                ExcelBuff.AddColumn('RM Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
                ExcelBuff.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
                ExcelBuff.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
                // commented parag
                //ExcelBuff.AddColumn('Item Category Code',FALSE,'',TRUE,FALSE,TRUE,'');
                // added coloumn item category description
                ExcelBuff.AddColumn('Item Category Name', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
                //Excelbuff.AddColumn('Bin Code',FALSE,'',TRUE,FALSE,TRUE,'');
                ExcelBuff.AddColumn('Unit Of Measure Code', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
                ExcelBuff.AddColumn('Remaining Qty in LTRS/KGS', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
                ExcelBuff.AddColumn('Receipt Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
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
                field("Print To Excel"; BoolPrintToExcel)
                {
                    ApplicationArea = all;
                }
                field("Date Filter"; AsOnDate)
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

        IF BoolPrintToExcel THEN
            createbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1
        ExcelBuff.DELETEALL;
        //<<1
    end;

    var
        recItem: Record 27;
        ExcelBuff: Record 370 temporary;
        SrNo: Integer;
        BoolPrintToExcel: Boolean;
        ItemDescription: Text[50];
        TotalQty: Decimal;
        AsOnDate: Date;
        totQty: Decimal;
        recDimension: Record 349;
        RecItemCategory: Record 5722;
        CategoryDesc: Text[50];
        User: Record 2000000120;
        LocResC: Code[20];
        recLocation: Record 14;
        CSOMapping1: Record 50006;
        recILE: Record 32;
        opTotQty: Decimal;
        "----------": Integer;
        ItemDesc: Code[70];
        LocName: Text[70];
        ILE: Record 32;
        PostingDate: Date;
        "---Rspl-Rahul": Integer;
        ItemNo: Text[100];
        vQuantity: Decimal;
        vAmount: Decimal;
        vtotalQty: Decimal;
        "---24Apr2017": Integer;
        NNN: Boolean;
        recDim24: Record 480;
        TCount: Integer;
        ICount: Integer;
        ILE24: Record 32;
        ILE24Apr: Record 32;

    // //[Scope('Internal')]
    procedure createbook()
    begin

        //ExcelBuff.CreateBook;
        //ExcelBuff.CreateSheet('Text001','Text002',COMPANYNAME,USERID);
        //ExcelBuff.GiveUserControl;
        //ERROR('');

        //>>ExcelBook
        /*  ExcelBuff.CreateBook('', 'Dimensionwise POP Item');
         ExcelBuff.CreateBookAndOpenExcel('', 'Dimensionwise POP Item', '', '', USERID);
         ExcelBuff.GiveUserControl; */

        ExcelBuff.CreateNewBook('Dimensionwise POP Item');
        ExcelBuff.WriteSheet('Dimensionwise POP Item', CompanyName, UserId);
        ExcelBuff.CloseBook();
        ExcelBuff.SetFriendlyFilename(StrSubstNo('Dimensionwise POP Item', CurrentDateTime, UserId));
        ExcelBuff.OpenExcel();
        //<<ExcelBook
    end;
}


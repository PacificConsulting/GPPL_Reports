report 50162 "Raw Material - Ageing"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 10Aug2018   RB-N         Transfer Price & Transfer Value
    // 27Aug2018   RB-N         Values for Ageing Period
    // 31Aug2018   RB-N         Item Category Description
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/RawMaterialAgeing.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Location; Location)
        {
            DataItemTableView = SORTING(Code)
                                ORDER(Ascending);
            RequestFilterFields = "Code";
            column(HDate; 'DATED :- ' + FORMAT(TODAY))
            {
            }
            column(LocName; 'RM STOCK DETAILS FOR  --------> ' + Location.Name)
            {
            }
            column(HLoc; Code)
            {
            }
            column(Date1; '0-' + FORMAT(NoOfDays))
            {
            }
            column(Date2; FORMAT(NoOfDays + 1) + '-' + FORMAT(NoOfDays1))
            {
            }
            column(Date3; FORMAT(NoOfDays1 + 1) + '-' + FORMAT(NoOfDays2))
            {
            }
            column(Date4; FORMAT(NoOfDays2 + 1) + '-' + FORMAT(NoOfDays3))
            {
            }
            column(Date5; '>' + FORMAT(NoOfDays3))
            {
            }
            dataitem("Item Ledger Entry"; "Item Ledger Entry")
            {
                DataItemLink = "Location Code" = FIELD(Code);
                DataItemTableView = SORTING("Item No.", "Location Code", "Lot No.", "Variant Code")
                                    ORDER(Ascending)
                                    WHERE("Remaining Quantity" = FILTER(> 0));
                RequestFilterFields = "Item No.", "Lot No.", "Item Category Code", "Global Dimension 1 Code";
                column(Srno; Srno)
                {
                }
                column(ItemDesc; Item.Description)
                {
                }
                column(BatchNo; "Item Ledger Entry"."Lot No.")
                {
                }
                column(PackType; "Item Ledger Entry"."Unit of Measure Code")
                {
                }
                column(PackSize; "Item Ledger Entry"."Qty. per Unit of Measure")
                {
                }
                column(NoPack; "Item Ledger Entry"."Remaining Quantity" / "Item Ledger Entry"."Qty. per Unit of Measure")
                {
                }
                column(RemQty; "NoofBarrels/Pails")
                {
                }
                column(MfgDate; "Item Ledger Entry"."Posting Date")
                {
                }
                column(Qty1; RemQty[1])
                {
                }
                column(Qty2; RemQty[2])
                {
                }
                column(Qty3; RemQty[3])
                {
                }
                column(Qty4; RemQty[4])
                {
                }
                column(Qty5; RemQty[5])
                {
                }
                column(ItmNo; "Item No.")
                {
                }
                column(ItemDim; itemDim)
                {
                }
                column(ItmLoc; "Item Ledger Entry"."Location Code")
                {
                }
                column(EntryNo; "Item Ledger Entry"."Entry No.")
                {
                }
                column(Text1; '0-' + FORMAT(NoOfDays))
                {
                }
                column(NoOfDays; NoOfDays)
                {
                }
                column(NoOfDays1; NoOfDays1)
                {
                }
                column(NoOfDays2; NoOfDays2)
                {
                }
                column(NoOfDays3; NoOfDays3)
                {
                }
                column(ExpireDate_ItemLedgerEntry; "Item Ledger Entry"."Expire Date")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>04Sep2018
                    RepItmNo := '';
                    IF "Item Category Code" = 'CAT15' THEN BEGIN
                        RepItmNo := COPYSTR("Item No.", 1, 4);
                        IF RepItmNo = 'FGRP' THEN
                            CurrReport.SKIP;
                    END;
                    //<<04Sep2018


                    //>>1
                    Item.GET("Item Ledger Entry"."Item No.");

                    IF Item.Blocked THEN
                        CurrReport.SKIP;

                    //IF ((Item."Item Category Code"='CAT02') OR (Item."Item Category Code"='CAT03') OR (Item."Item Category Code"='CAT10')) THEN //Fahim 21-05-2022
                    IF ((Item."Item Category Code" = 'CAT02') OR (Item."Item Category Code" = 'CAT03')) THEN
                        CurrReport.SKIP;

                    ItemUOM1.RESET;
                    ItemUOM1.SETRANGE(ItemUOM1."Item No.", "Item Ledger Entry"."Item No.");
                    ItemUOM1.SETRANGE(ItemUOM1.Code, "Item Ledger Entry"."Unit of Measure Code");
                    IF ItemUOM1.FINDFIRST THEN;

                    "NoofBarrels/Pails" := "Item Ledger Entry"."Remaining Quantity";
                    TotBarrels := TotBarrels + "NoofBarrels/Pails";

                    //>>10Aug2018
                    TPrice10 := 0;
                    Value10 := 0;
                    IF TPrice10 = 0 THEN BEGIN
                        CALCFIELDS("Cost Amount (Expected)", "Cost Amount (Actual)");
                        TPrice10 := ("Cost Amount (Expected)" + "Cost Amount (Actual)") / Quantity;
                        Value10 := TPrice10 * "Remaining Quantity";
                    END;
                    //<<10Aug2018

                    FOR i := 1 TO 5 DO BEGIN
                        RemQty[i] := 0;
                        RemValue[i] := 0;
                    END;

                    IF ((TODAY - "Item Ledger Entry"."Posting Date") >= 0) AND ((TODAY - "Item Ledger Entry"."Posting Date") <= NoOfDays) THEN BEGIN
                        RemQty[1] += "Item Ledger Entry"."Remaining Quantity";
                        RemValue[1] += Value10;
                    END;

                    IF ((TODAY - "Item Ledger Entry"."Posting Date") >= NoOfDays + 1) AND ((TODAY - "Item Ledger Entry"."Posting Date") <= (NoOfDays1)) THEN BEGIN
                        RemQty[2] += "Item Ledger Entry"."Remaining Quantity";
                        RemValue[2] += Value10;
                    END;

                    IF ((TODAY - "Item Ledger Entry"."Posting Date") >= (NoOfDays1 + 1)) AND ((TODAY - "Item Ledger Entry"."Posting Date") <= (NoOfDays2)) THEN BEGIN
                        RemQty[3] += "Item Ledger Entry"."Remaining Quantity";
                        RemValue[3] += Value10;
                    END;

                    IF ((TODAY - "Item Ledger Entry"."Posting Date") >= (NoOfDays2 + 1)) AND ((TODAY - "Item Ledger Entry"."Posting Date") <= (NoOfDays3)) THEN BEGIN
                        RemQty[4] += "Item Ledger Entry"."Remaining Quantity";
                        RemValue[4] += Value10;
                    END;

                    IF ((TODAY - "Item Ledger Entry"."Posting Date") > (NoOfDays3)) THEN BEGIN
                        RemQty[5] += "Item Ledger Entry"."Remaining Quantity";
                        RemValue[5] += Value10;
                    END;

                    CLEAR(BinName);
                    recWarehouseEntry.RESET;
                    recWarehouseEntry.SETRANGE(recWarehouseEntry."Item No.", "Item Ledger Entry"."Item No.");
                    //recWarehouseEntry.SETRANGE(recWarehouseEntry."Reference No.","Item Ledger Entry"."Document No."); comment by Fahim 21-05-22
                    recWarehouseEntry.SETRANGE(recWarehouseEntry."Source No.", "Item Ledger Entry"."Document No.");
                    //recWarehouseEntry.SETRANGE(recWarehouseEntry."Source Line No.","Item Ledger Entry"."Document Line No."); comment by Fahim 21-05-22
                    IF recWarehouseEntry.FINDFIRST THEN BEGIN
                        recBin.RESET;
                        recBin.SETRANGE(recBin.Code, recWarehouseEntry."Bin Code");
                        IF recBin.FINDFIRST THEN
                            BinName := recBin.Description;
                    END;

                    DocumentNo := '';
                    Date := 0D;
                    "Trf.frmCode" := '';
                    "Trf.frmName" := '';
                    IF ("Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Transfer Receipt") THEN BEGIN
                        recTSH.RESET;
                        //recTSH.SETRANGE(recTSH."Transfer Order No.","Item Ledger Entry"."Transfer Order No.");
                        recTSH.SETRANGE(recTSH."Transfer Order No.", "Item Ledger Entry"."Order No.");//17Mar2017
                        IF recTSH.FINDFIRST THEN BEGIN
                            DocumentNo := recTSH."No.";
                            Date := recTSH."Posting Date";
                            "Trf.frmCode" := recTSH."Transfer-from Code";
                        END;
                        recLoc.RESET;
                        recLoc.SETRANGE(recLoc.Code, "Trf.frmCode");
                        IF recLoc.FINDFIRST THEN BEGIN
                            "Trf.frmName" := recLoc.Name;
                        END;
                    END
                    ELSE BEGIN
                        DocumentNo := "Item Ledger Entry"."Document No.";
                        Date := "Item Ledger Entry"."Posting Date";
                    END;

                    VendorCode := '';
                    VendorName := '';
                    IF ("Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Purchase Receipt") THEN BEGIN
                        recPRH.RESET;
                        recPRH.SETRANGE(recPRH."No.", "Item Ledger Entry"."Document No.");
                        IF recPRH.FINDFIRST THEN BEGIN
                            VendorCode := recPRH."Buy-from Vendor No.";
                        END;
                        recvendor.RESET;
                        recvendor.SETRANGE(recvendor."No.", VendorCode);
                        IF recvendor.FINDFIRST THEN BEGIN
                            VendorName := recvendor."Full Name";
                        END;
                    END;

                    //RSPL
                    itemDim := '';
                    recDefualtDim.RESET;
                    recDefualtDim.SETRANGE(recDefualtDim."Table ID", 27);
                    recDefualtDim.SETRANGE(recDefualtDim."No.", "Item No.");
                    IF recDefualtDim.FINDFIRST THEN
                        itemDim := recDefualtDim."Dimension Value Code";
                    //RSPL

                    //<<1



                    //>>2
                    Srno := Srno + 1;//17Mar2017
                    //Item Ledger Entry, Body (1) - OnPreSection()
                    IF PrintToExcel THEN BEGIN
                        //Srno:=Srno+1; //17Mar2017
                        MakeExcelDataBody;
                    END;
                    //<<2


                    //>>17Mar2017 GroupFooterNew
                    LCOUNT += 1;

                    ILE17Mar.SETCURRENTKEY("Item No.", "Location Code", "Lot No.", "Variant Code", "Posting Date");
                    ILE17Mar.SETRANGE("Item No.", "Item No.");
                    ILE17Mar.SETRANGE("Location Code", "Location Code");
                    IF ILE17Mar.FINDFIRST THEN BEGIN
                        TCOUNT := ILE17Mar.COUNT;
                    END;

                    IF TCOUNT <> 0 THEN
                        IF PrintToExcel AND (TCOUNT = LCOUNT) THEN BEGIN
                            MakeExcelFooter;
                            LCOUNT := 0;

                        END;

                    //<<17Mar2017 GroupFooterNew
                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    i := 5;

                    CurrReport.CREATETOTALS("NoofBarrels/Pails");

                    //<<1

                    GCOUNT := COUNT;

                    ILE17Mar.COPYFILTERS("Item Ledger Entry");
                end;
            }

            trigger OnAfterGetRecord()
            begin


                //>>1
                //Location, Header (1) - OnPreSection()
                j := 0;

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
                field("No. Of Days"; NoOfDays)
                {
                    ApplicationArea = all;
                }
                field("No. Of Days1"; NoOfDays1)
                {
                    ApplicationArea = all;
                }
                field("No. Of Days2"; NoOfDays2)
                {
                    ApplicationArea = all;
                }
                field("No. Of Days3"; NoOfDays3)
                {
                    ApplicationArea = all;
                }
                field("Print To Excel"; PrintToExcel)
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

        IF PrintToExcel THEN
            CreateExcelbook;

        //<<1
    end;

    trigger OnPreReport()
    begin



        //>>1


        LocCode := Location.GETFILTER(Location.Code);

        /*
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",UPPERCASE(USERID));
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
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
        *///Commented 17Mar2017

        CompInfo.GET;

        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<1

        Text1 := '0-' + FORMAT(NoOfDays);//17Mar2017

    end;

    var
        Srno: Integer;
        Item: Record 27;
        ItemNo: Code[20];
        Locationcode: Code[20];
        batchno: Code[10];
        Packsize: Decimal;
        Noofbarrels: Decimal;
        OpeningStock: Decimal;
        MfgDate: Date;
        "NoofBarrels/Pails": Decimal;
        "<To Export to Excel>": Integer;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        LocationFilter: Text[50];
        ILEFilter: Text[50];
        n: Text[30];
        MaxCol: Text[3];
        i: Integer;
        DateLength: Integer;
        Date1: Text[30];
        MonthLength: Integer;
        Month1: Text[30];
        YearLength: Integer;
        Year1: Text[30];
        PostingDate: Text[30];
        ItemUOM: Record 5404;
        ItemUOM1: Record 5404;
        UserSetup: Record 91;
        UserRespCtr: Code[10];
        CompInfo: Record 79;
        TotBarrels: Decimal;
        ItemCatCode: Code[30];
        j: Integer;
        PrevItemNo: Text[50];
        RemQty: array[10] of Decimal;
        ILERec: Record 32;
        NoOfDays: Integer;
        Makefooter: Boolean;
        NoOfDays1: Integer;
        NoOfDays2: Integer;
        NoOfDays3: Integer;
        recWarehouseEntry: Record 7312;
        recBin: Record 7354;
        BinName: Text[50];
        recTSH: Record 5744;
        DocumentNo: Code[20];
        Date: Date;
        "Trf.frmCode": Code[10];
        recLoc: Record 14;
        "Trf.frmName": Text[50];
        VendorCode: Code[10];
        VendorName: Text[50];
        recPRH: Record 120;
        recvendor: Record 23;
        LocResC: Code[30];
        recLocation: Record 14;
        LocCode: Code[30];
        CSOMapping1: Record 50006;
        recBinContent: Record 7302;
        recILE: Record 32;
        itemDim: Code[20];
        recDefualtDim: Record 352;
        Text16500: Label 'As per Details';
        Text000: Label 'Data';
        Text001: Label 'Raw Material - Ageing';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "----17Mar2017": Integer;
        GCOUNT: Integer;
        GNoBar: Decimal;
        Text1: Text[50];
        "---------10Aug2018": Integer;
        TPrice10: Decimal;
        Value10: Decimal;
        ILE17Mar: Record 32;
        TCOUNT: Integer;
        LCOUNT: Integer;
        RemValue: array[10] of Decimal;
        ItmCat: Record 5722;
        ItmCatName: Text;
        RepItmNo: Code[10];

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//17Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Raw Material - Ageing', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Raw Material - Ageing", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        /*
        ExcelBuf.CreateBook;
        ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        ExcelBuf.GiveUserControl;
        ERROR('');
        *///Commented

        //>>17Mar2017 RB-N

        /*  ExcelBuf.CreateBook('', Text001);
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
        //>>17Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20 10Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21 10Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27 31Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28 RSPLSUM 25Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 22-11-21 Expiry Date
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn(Text001, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20 10Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21 10Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27 31Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28 RSPLSUM 25Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 22-11-21 Expiry Date
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.NewRow;

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21 10Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 10Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 31Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 RSPLSUM 25Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 22-11-21 Expiry Date
        //<<17Mar2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Sr.No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31Aug2018
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('BIN', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Pack Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Pack Size', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('No.Of Bar per Pails', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Remaining Qty in Ltrs/KG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Mfg.Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('0-' + FORMAT(NoOfDays), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('0-' + FORMAT(NoOfDays) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27Aug2018

        ExcelBuf.AddColumn(FORMAT(NoOfDays + 1) + '-' + FORMAT(NoOfDays1), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn(FORMAT(NoOfDays + 1) + '-' + FORMAT(NoOfDays1) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27Aug2018

        ExcelBuf.AddColumn(FORMAT(NoOfDays1 + 1) + '-' + FORMAT(NoOfDays2), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn(FORMAT(NoOfDays1 + 1) + '-' + FORMAT(NoOfDays2) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27Aug2018

        ExcelBuf.AddColumn(FORMAT(NoOfDays2 + 1) + '-' + FORMAT(NoOfDays3), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn(FORMAT(NoOfDays2 + 1) + '-' + FORMAT(NoOfDays3) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27Aug2018

        ExcelBuf.AddColumn('>' + FORMAT(NoOfDays3), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('>' + FORMAT(NoOfDays3) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27Aug2018

        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        //RSPLSUM 25Aug2020--ExcelBuf.AddColumn('Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//18
        ExcelBuf.AddColumn('Vendor/From Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18 RSPLSUM 25Aug2020
        ExcelBuf.AddColumn('Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Default Dimension', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20 //RSPL
        ExcelBuf.AddColumn('Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21 10Aug2018
        ExcelBuf.AddColumn('Transfer Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 10Aug2018
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 RSPLSUM 25Aug2020
        ExcelBuf.AddColumn('Expiry Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 22-11-21 Expiry Date
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Srno, FALSE, '', FALSE, FALSE, FALSE, '', 0);//1 17Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2 17Mar2017
        ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3 17Mar2017
        //>>31Aug2018
        ItmCatName := '';
        ItmCat.RESET;
        IF ItmCat.GET("Item Ledger Entry"."Item Category Code") THEN
            ItmCatName := ItmCat.Description;
        ExcelBuf.AddColumn(ItmCatName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //<<31Aug2018
        ExcelBuf.AddColumn("Item Ledger Entry"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4 17Mar2017
        ExcelBuf.AddColumn(BinName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5 17Mar2017
        //ExcelBuf.AddColumn(recWarehouseEntry."Bin Code",FALSE,'',FALSE,FALSE,FALSE,'@',1);
        ExcelBuf.AddColumn("Item Ledger Entry"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6 17Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//7 17Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Remaining Quantity", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//8 17Mar2017
        //>>17Mar2017 GNoBar
        GNoBar += "Item Ledger Entry"."Remaining Quantity";
        //>>17Mar2017 GNoBar

        ExcelBuf.AddColumn("NoofBarrels/Pails", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//9 17Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//10
        ExcelBuf.AddColumn(RemQty[1], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//11
        ExcelBuf.AddColumn(RemValue[1], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27Aug2018

        ExcelBuf.AddColumn(RemQty[2], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//12
        ExcelBuf.AddColumn(RemValue[2], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27Aug2018

        ExcelBuf.AddColumn(RemQty[3], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//13
        ExcelBuf.AddColumn(RemValue[3], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27Aug2018

        ExcelBuf.AddColumn(RemQty[4], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//14
        ExcelBuf.AddColumn(RemValue[4], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27Aug2018

        ExcelBuf.AddColumn(RemQty[5], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//15
        ExcelBuf.AddColumn(RemValue[5], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27Aug2018

        ExcelBuf.AddColumn(DocumentNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn(Date, FALSE, '', FALSE, FALSE, FALSE, '', 2);//17

        IF ("Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Transfer Receipt") THEN BEGIN
            ExcelBuf.AddColumn("Trf.frmCode", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
            ExcelBuf.AddColumn("Trf.frmName", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
        END;

        IF ("Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Purchase Receipt") THEN BEGIN
            ExcelBuf.AddColumn(VendorCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
            ExcelBuf.AddColumn(VendorName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
        END;

        IF NOT (("Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Transfer Receipt") OR
           ("Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Purchase Receipt")) THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
        END;
        ExcelBuf.AddColumn(itemDim, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//20 //RSPL

        ExcelBuf.AddColumn(Value10, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//21 10Aug2018
        ExcelBuf.AddColumn(TPrice10, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22 10Aug2018
        ExcelBuf.AddColumn("Item Ledger Entry"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//23 RSPLSUM 25Aug2020
        ExcelBuf.AddColumn("Item Ledger Entry"."Expire Date", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//Fahim 22-11-21 Expiry Date
    end;

    // //[Scope('Internal')]
    procedure MakeExcelFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        //ExcelBuf.AddColumn("Item Ledger Entry"."Remaining Quantity",FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//8
        ExcelBuf.AddColumn(GNoBar, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//8 //17Mar2017
        GNoBar := 0; //17Mar2017

        ExcelBuf.AddColumn(TotBarrels, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//9
        TotBarrels := 0;//17Mar2018

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20 17Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21 10Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 10Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 RSPLSUM 25Aug2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 22-11-21 Expiry Date
    end;
}


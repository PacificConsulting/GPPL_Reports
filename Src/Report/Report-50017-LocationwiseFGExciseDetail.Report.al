report 50017 "Location wise FG-Excise Detail"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/LocationwiseFGExciseDetail.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            DataItemTableView = SORTING("Location Code", "Item No.", "Lot No.", "Posting Date", "Document No.")
                                ORDER(Ascending)
                                WHERE("Item Category Code" = FILTER('CAT02' | 'CAT03' | 'CAT11' | 'CAT12' | 'CAT13'),
                                      Open = FILTER(true));
            RequestFilterFields = "Item No.", "Lot No.", "Item Category Code", "Location Code";

            trigger OnAfterGetRecord()
            begin
                I += 1;//18Feb2017
                       //>>1

                ItemDescription := '';
                Item.RESET;
                Item.SETRANGE(Item."No.", "Item Ledger Entry"."Item No.");
                IF Item.FINDFIRST THEN BEGIN
                    ItemDescription := Item.Description;
                END;

                LocationName := '';
                Lcoation.RESET;
                Lcoation.SETRANGE(Lcoation.Code, "Item Ledger Entry"."Location Code");
                IF Lcoation.FINDFIRST THEN BEGIN
                    LocationName := Lcoation.Name;
                END;

                CLEAR(BRTno);
                CLEAR(BRTdate);
                CLEAR(ExciseRate);
                CLEAR(BEDAmount);
                CLEAR(EcessAmount);
                CLEAR(SheCessAmount);
                CLEAR(TransferFromCode);
                CLEAR(TransferFromName);
                CLEAR(TrfPrice);
                CLEAR(TrfValue);
                CLEAR(AssValue);
                CLEAR(Amt);

                BRTno := "Item Ledger Entry"."Document No.";
                BRTdate := "Item Ledger Entry"."Posting Date";

                IF "Document Type" = "Document Type"::"Transfer Receipt" THEN BEGIN
                    recTSH.RESET;
                    //recTSH.SETRANGE(recTSH."Transfer Order No.","Transfer Order No.");
                    recTSH.SETRANGE("Transfer Order No.", "Order No.");//18Feb2017
                    IF recTSH.FINDFIRST THEN BEGIN
                        BRTno := recTSH."No.";
                        BRTdate := recTSH."Posting Date";
                        TransferFromCode := recTSH."Transfer-from Code";
                    END;

                    recLoc.RESET;
                    recLoc.SETRANGE(recLoc.Code, TransferFromCode);
                    IF recLoc.FINDFIRST THEN BEGIN
                        TransferFromName := recLoc.Name;
                    END;

                    recTSL.RESET;
                    recTSL.SETRANGE(recTSL."Document No.", BRTno);
                    recTSL.SETRANGE(recTSL."Item No.", "Item No.");
                    recTSL.SETRANGE(recTSL."Line No.", "Document Line No.");
                    IF recTSL.FIND('-') THEN BEGIN

                        /*  ExcPostingSetup.RESET;
                         ExcPostingSetup.SETRANGE(ExcPostingSetup."Excise Bus. Posting Group", recTSL."Excise Bus. Posting Group");
                         ExcPostingSetup.SETRANGE(ExcPostingSetup."Excise Prod. Posting Group", recTSL."Excise Prod. Posting Group");
                         IF ExcPostingSetup.FINDLAST THEN BEGIN
                             BEDperc := ExcPostingSetup."BED %";
                             eCessperc := ExcPostingSetup."eCess %";
                             SheCessperc := ExcPostingSetup."SHE Cess %";
                         END;
  */
                        TrfPrice := recTSL."Unit Price" / recTSL."Qty. per Unit of Measure";
                        TrfValue := TrfPrice * "Item Ledger Entry"."Remaining Quantity";
                        AssValue := 0;// recTSL."Assessable Value" / recTSL."Qty. per Unit of Measure";
                        Amt := AssValue * "Item Ledger Entry"."Remaining Quantity";
                        ExciseRate := 0;// recTSL."Excise Effective Rate";
                        BEDAmount := ROUND(((TrfValue * BEDperc) / 100), 0.01);
                        EcessAmount := ROUND(((BEDAmount * eCessperc) / 100), 0.01);
                        SheCessAmount := ROUND(((BEDAmount * SheCessperc) / 100), 0.01);

                    END;

                END
                ELSE BEGIN

                    /*  recRG23D.RESET;
                     recRG23D.SETRANGE(recRG23D."Document No.", "Item Ledger Entry"."Document No.");
                     recRG23D.SETRANGE(recRG23D."Item No.", "Item Ledger Entry"."Item No.");
                     recRG23D.SETRANGE(recRG23D."Lot No.", "Item Ledger Entry"."Lot No.");
                     recRG23D.SETRANGE(recRG23D."Location Code", "Item Ledger Entry"."Location Code");
                     IF recRG23D.FINDFIRST THEN BEGIN
                         TrfPrice := recRG23D."Transfer Price";
                         TrfValue := TrfPrice * "Item Ledger Entry"."Remaining Quantity";
                         AssValue := recRG23D."Excise Base Amt Per Unit";
                         Amt := AssValue * "Item Ledger Entry"."Remaining Quantity";
                         ExciseRate := "Item Ledger Entry"."BED %";
                         BEDAmount := ROUND((recRG23D."BED Amount Per Unit" * "Item Ledger Entry"."Remaining Quantity"), 0.01);
                         EcessAmount := ROUND((recRG23D."eCess Amount Per Unit" * "Item Ledger Entry"."Remaining Quantity"), 0.01);
                         SheCessAmount := ROUND((recRG23D."SHE Cess Amount Per Unit" * "Item Ledger Entry"."Remaining Quantity"), 0.01);

                     END;
  */
                END;

                MRP := 0;
                recMrpMaster.RESET;
                recMrpMaster.SETRANGE(recMrpMaster."Item No.", "Item Ledger Entry"."Item No.");
                recMrpMaster.SETRANGE(recMrpMaster."Lot No.", "Item Ledger Entry"."Lot No.");
                IF recMrpMaster.FINDFIRST THEN BEGIN
                    MRP := recMrpMaster."MRP Price";
                END;


                EBTILE.RESET;
                EBTILE.SETRANGE(EBTILE."Item No.", "Item Ledger Entry"."Item No.");
                EBTILE.SETRANGE(EBTILE."Lot No.", "Item Ledger Entry"."Lot No.");
                EBTILE.SETRANGE(EBTILE."Location Code", "Item Ledger Entry"."Location Code");
                EBTILE.SETRANGE(EBTILE."Entry Type", EBTILE."Entry Type"::Transfer);
                EBTILE.SETRANGE(EBTILE."Document No.", "Item Ledger Entry"."Document No.");
                //EBTILE.SETRANGE(EBTILE.Open,TRUE);
                IF EBTILE.FINDFIRST THEN BEGIN
                    ReceiptDate := EBTILE."Posting Date";
                END;
                /*
                IF ReceiptDate = 0D THEN
                BEGIN

                {
                  EBTILE.RESET;
                  EBTILE.SETRANGE(EBTILE."Item No.","Item Ledger Entry"."Item No.");
                  EBTILE.SETRANGE(EBTILE."Lot No.","Item Ledger Entry"."Lot No.");
                  EBTILE.SETRANGE(EBTILE."Location Code","Item Ledger Entry"."Location Code");
                  EBTILE.SETFILTER(EBTILE."Posting Date",'%1..%2',0D,TODAY);
                  EBTILE.SETFILTER(EBTILE.Quantity,'>%1',0);
                  IF EBTILE.FINDLAST THEN
                  BEGIN
                   ReceiptDate := EBTILE."Posting Date";
                  END;
                }

                END;
                */
                //<<1

                IF PrintToExcel THEN
                    MakeExcelDataBody;


                //18Feb2017 RB-N
                IF I = ILECount THEN BEGIN

                    IF PrintToExcel THEN
                        MakeExcelDataFooter;
                END;

            end;

            trigger OnPreDataItem()
            begin
                //>>1
                SETFILTER("Item Ledger Entry"."Posting Date", '%1..%2', 0D, TODAY);
                //<<1

                ILECount := COUNT;//18Feb2017
                I := 0;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print To Excel"; PrintToExcel)
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

        IF PrintToExcel THEN
            CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin
        //>>1
        /*
        //EBT STIVAN ---(09062012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        LocCode := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code");
        User.GET(USERID);
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",User."User ID");
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

         CSOMapping.RESET;
         CSOMapping.SETRANGE(CSOMapping."User Id",UPPERCASE(USERID));
         CSOMapping.SETRANGE(Type,CSOMapping.Type::Location);
         CSOMapping.SETRANGE(CSOMapping.Value,LocCode);
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
        //EBT STIVAN ---(09062012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        */
        CompanyInfo.GET;

        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<1

    end;

    var
        Item: Record 27;
        ItemDescription: Text[100];
        Lcoation: Record 14;
        LocationName: Text[50];
        recTSH: Record 5744;
        BRTno: Code[20];
        BRTdate: Date;
        recTSL: Record 5745;
        TransferFromCode: Code[10];
        TransferFromName: Text[30];
        ExciseRate: Decimal;
        BEDAmount: Decimal;
        EcessAmount: Decimal;
        SheCessAmount: Decimal;
        recLoc: Record 14;
        ExcelBuf: Record 370 temporary;
        filterdate: Date;
        CompanyInfo: Record 79;
        PrintToExcel: Boolean;
        LocCode: Code[20];
        LocResC: Code[20];
        recLocation: Record 14;
        CSOMapping: Record 50006;
        //  ExcPostingSetup: Record 13711;
        BEDperc: Decimal;
        eCessperc: Decimal;
        SheCessperc: Decimal;
        TrfPrice: Decimal;
        TrfValue: Decimal;
        AssValue: Decimal;
        Amt: Decimal;
        //recRG23D: Record 16537;
        recMrpMaster: Record 50013;
        MRP: Decimal;
        EBTILE: Record 32;
        ReceiptDate: Date;
        Text002: Label 'Data';
        Text003: Label 'Location wise FG-Excise Details';
        Text004: Label 'Company Name';
        Text005: Label 'Report No.';
        Text006: Label 'Report Name';
        Text007: Label 'User ID';
        Text008: Label 'Date';
        Text009: Label 'Location Filter';
        "--18Feb2017": Integer;
        I: Integer;
        ILECount: Integer;
        TQty: Decimal;

    ////[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text004), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(CompanyInfo.Name, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text006), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(FORMAT(Text003), FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text005), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Location wise FG-Excise Detail", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text007), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader
    end;

    ////[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text002,Text003,CompanyInfo.Name,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>18Feb2017 RB-N
        /*  ExcelBuf.CreateBook('', Text003);
         //ExcelBuf.CreateBookAndOpenExcel('',Text000,Text001,COMPANYNAME,USERID);
         ExcelBuf.CreateBookAndOpenExcel('', Text003, '', '', USERID);
         ExcelBuf.GiveUserControl; */

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
        //>> 18Feb2017 RB-N
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn(Text003, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//14
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//15
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//16
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//17
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//18
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//19
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//20
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//21
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.NewRow;


        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//14
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//15
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//16
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//17
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//18
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//19
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//20
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//21
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.NewRow;
        //<<
        //20Feb17
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.NewRow;
        //20Feb17
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Qty Per Pack', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Qty in LTRS/KGS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Lot No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('BRT N0.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('BRT Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Transfered From', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Excise Rate', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Ecess Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Shecess Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Transfer Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('ASS. Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('MRP', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Receipt Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('No. of Days', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
    end;

    ////[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Item Ledger Entry"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(ItemDescription, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn("Item Ledger Entry"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn("Item Ledger Entry"."Remaining Quantity" /
        "Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn("Item Ledger Entry"."Remaining Quantity", FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn("Item Ledger Entry"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn("Item Ledger Entry"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(LocationName, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(BRTno, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(BRTdate, FALSE, '', FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.AddColumn(TransferFromName, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(ExciseRate, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn(BEDAmount, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn(EcessAmount, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn(SheCessAmount, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn(TrfPrice, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn(TrfValue, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn(AssValue, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn(Amt, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn(MRP, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
        IF ReceiptDate <> 0D THEN BEGIN
            ExcelBuf.AddColumn(ReceiptDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);
            ExcelBuf.AddColumn(-(ReceiptDate - TODAY), FALSE, '', FALSE, FALSE, FALSE, '', 0);
        END
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);

        TQty := TQty + "Item Ledger Entry"."Remaining Quantity"; //18Feb2017
    end;

    ////[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        //ExcelBuf.AddColumn("Item Ledger Entry"."Remaining Quantity",FALSE,'',TRUE,FALSE,FALSE,'',0);
        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, FALSE, '#,####0', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
    end;
}


report 50052 "FG Movement (Locwise)"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/FGMovementLocwise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Item Category Code" = FILTER('CAT02|CAT03|CAT11|CAT12|CAT13|CAT15|CAT18'));
            RequestFilterFields = "No.", "Date Filter", "Location Filter";

            trigger OnAfterGetRecord()
            begin

                //>>1

                //SLNo += 1; //12May2017
                ItemUOM.GET(Item."No.", Item."Sales Unit of Measure");
                QtyPerUOm := ItemUOM."Qty. per Unit of Measure";
                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETFILTER("Date Filter", '<%1', PosDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETRANGE(ItemRec."Location Filter", LocFilter);
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    OpeningBalance := ItemRec."Inventory Change";
                    OpeningBalanceSOM := OpeningBalance / QtyPerUOm;
                END;

                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETRANGE("Date Filter", StartDate, EndDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETFILTER(ItemRec."Location Filter", LocFilter);
                ItemRec.SETRANGE(ItemRec."Entry Type Filter", ItemRec."Entry Type Filter"::Output);
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    MfgQty := ItemRec."Inventory Change";
                    MfgQtySOM := MfgQty / QtyPerUOm;
                END;

                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETRANGE("Date Filter", StartDate, EndDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETFILTER(ItemRec."Location Filter", LocFilter);
                ItemRec.SETRANGE(ItemRec."Entry Type Filter", ItemRec."Entry Type Filter"::Transfer);
                ItemRec.SETRANGE(ItemRec."Document Type Filter", ItemRec."Document Type Filter"::"Transfer Receipt");
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    TransferRecvQty := ItemRec."Inventory Change";
                    TransferRecvQtySOM := TransferRecvQty / QtyPerUOm;
                END;
                //.....sunil
                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETRANGE("Date Filter", StartDate, EndDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETFILTER(ItemRec."Location Filter", LocFilter);
                ItemRec.SETRANGE(ItemRec."Entry Type Filter", ItemRec."Entry Type Filter"::"Positive Adjmt.");
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    PosAdj := ItemRec."Inventory Change";
                    PosAdjQtySom := PosAdj / QtyPerUOm;
                END;
                //....sunil

                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETRANGE("Date Filter", StartDate, EndDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETFILTER(ItemRec."Location Filter", LocFilter);
                ItemRec.SETRANGE(ItemRec."Entry Type Filter", ItemRec."Entry Type Filter"::Sale);
                ItemRec.SETRANGE(ItemRec."Document Type Filter", ItemRec."Document Type Filter"::"Sales Return Receipt");
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    SalesReturnQty := ItemRec."Inventory Change";
                    SalesReturnQtySOM := SalesReturnQty / QtyPerUOm;
                END;

                TotalRecvQty := MfgQty + TransferRecvQty + SalesReturnQty + PosAdj;
                TotalRecvQtySOM := MfgQtySOM + TransferRecvQtySOM + SalesReturnQtySOM + PosAdjQtySom;

                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETRANGE("Date Filter", StartDate, EndDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETFILTER(ItemRec."Location Filter", LocFilter);
                ItemRec.SETRANGE(ItemRec."Entry Type Filter", ItemRec."Entry Type Filter"::Sale);
                ItemRec.SETRANGE(ItemRec."Document Type Filter", ItemRec."Document Type Filter"::"Sales Shipment");
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    SalesQty := ItemRec."Inventory Change";
                    SalesQtySOM := SalesQty / QtyPerUOm;
                END;

                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETRANGE("Date Filter", StartDate, EndDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETFILTER(ItemRec."Location Filter", LocFilter);
                ItemRec.SETRANGE(ItemRec."Entry Type Filter", ItemRec."Entry Type Filter"::Transfer);
                ItemRec.SETRANGE(ItemRec."Document Type Filter", ItemRec."Document Type Filter"::"Transfer Shipment");
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    TransferShipQty := ItemRec."Inventory Change";
                    TransferShipQtySOM := TransferShipQty / QtyPerUOm;
                END;

                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETRANGE("Date Filter", StartDate, EndDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETFILTER(ItemRec."Location Filter", LocFilter);
                ItemRec.SETRANGE(ItemRec."Entry Type Filter", ItemRec."Entry Type Filter"::Consumption);
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    ConsumedQty := ItemRec."Inventory Change";
                    ConsumedinSOM := ConsumedQty / QtyPerUOm;
                END;
                //;'''''sunil
                ItemRec.RESET;
                ItemRec.SETRANGE(ItemRec."No.", Item."No.");
                ItemRec.SETRANGE("Date Filter", StartDate, EndDate);
                IF LocFilter <> '' THEN
                    ItemRec.SETFILTER(ItemRec."Location Filter", LocFilter);
                ItemRec.SETRANGE(ItemRec."Entry Type Filter", ItemRec."Entry Type Filter"::"Negative Adjmt.");
                IF ItemRec.FINDFIRST THEN BEGIN
                    ItemRec.CALCFIELDS(ItemRec."Inventory Change");
                    NegAdj := (ItemRec."Inventory Change");
                    NegAdjQtySom := NegAdj / QtyPerUOm;
                END;



                TotalDespatchQty := SalesQty + TransferShipQty + ConsumedQty + NegAdj;
                TotalDespatchQtySOM := SalesQtySOM + TransferShipQtySOM + ConsumedinSOM + NegAdjQtySom;

                ClosingBalance := OpeningBalance + TotalRecvQty + TotalDespatchQty;
                ClosingBalanceSOM := OpeningBalanceSOM + TotalRecvQtySOM + TotalDespatchQtySOM;

                //EBT STIVAN ---(25062012)--- To Skip Null Values------------------------------------------------------------------------------START
                IF (OpeningBalanceSOM = 0) AND (OpeningBalance = 0) AND (MfgQtySOM = 0) AND (MfgQty = 0) AND (TransferRecvQtySOM = 0) AND
                   (TransferRecvQty = 0) AND (SalesReturnQtySOM = 0) AND (SalesReturnQty = 0) AND (TotalRecvQtySOM = 0) AND (TotalRecvQty = 0)
                   AND (ConsumedQty = 0) THEN
                    CurrReport.SKIP;
                //EBT STIVAN ---(25062012)--- To Skip Null Values--------------------------------------------------------------------------------END

                SLNo += 1; //12May2017
                MakeExcelDataBody;
                //<<1
            end;

            trigger OnPostDataItem()
            begin

                //>>1
                MakeExcelDataFooter;
                //<<1
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                PosDate := Item.GETRANGEMIN(Item."Date Filter");
                LocFilter := Item.GETFILTER(Item."Location Filter");
                StartDate := Item.GETRANGEMIN(Item."Date Filter");
                EndDate := Item.GETRANGEMAX(Item."Date Filter");

                CurrReport.CREATETOTALS(OpeningBalance, OpeningBalanceSOM, MfgQty, MfgQtySOM, TransferRecvQty, TransferRecvQtySOM, SalesReturnQty,
                                        SalesReturnQtySOM, TotalRecvQty, TotalRecvQtySOM);
                CurrReport.CREATETOTALS(SalesQty, SalesQtySOM, TransferShipQty, TransferShipQtySOM, ConsumedQty, ConsumedinSOM,
                                      TotalDespatchQty, TotalDespatchQtySOM, ClosingBalance, ClosingBalanceSOM);
                CurrReport.CREATETOTALS(PosAdj, PosAdjQtySom);
                CurrReport.CREATETOTALS(NegAdj, NegAdjQtySom);
                //<<1
            end;
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

        //>>1
        CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1
        /*
        //user.GET(USERID);
        IF (user."User ID" = 'CF24') OR (user."User ID" = 'CF09') OR (user."User ID" = 'CF11') OR (user."User ID" = 'CF12') OR
        (user."User ID" = 'CF16') OR (user."User ID" = 'CF22') OR (user."User ID" = 'CF29') OR (user."User ID" = 'CF31') OR
        (user."User ID" = 'CF32') OR (user."User ID" = 'CF34') OR (user."User ID" = 'CF35') OR (user."User ID" = 'CF36') OR
        (user."User ID" = 'CF40') OR (user."User ID" = 'CF41') OR (user."User ID" = 'CF42') OR (user."User ID" = 'CF43') OR
        (user."User ID" = 'CF44') OR (user."User ID" = 'CF46') OR (user."User ID" = 'CF47') OR (user."User ID" = 'CF48') OR
        (user."User ID" = 'CF49') OR (user."User ID" = 'CF50') OR (user."User ID" = 'CF51') OR (user."User ID" = 'CF52') OR
        (user."User ID" = 'CF53') OR (user."User ID" = 'CF55') THEN
        ERROR('You are not allowed to run this Report');
        */

        //EBT STIVAN ---(20092012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User Name", USERID);
        Memberof.SETFILTER(Memberof."Role ID", '%1|%2', 'SUPER', 'REPORT VIEW');
        IF NOT (Memberof.FINDFIRST) THEN BEGIN
            CLEAR(LocCode);
            LocCode := Item.GETFILTER(Item."Location Filter");
            CLEAR(LocResC);
            recLocation.RESET;
            recLocation.SETRANGE(recLocation.Code, LocCode);
            IF recLocation.FINDFIRST THEN BEGIN
                LocResC := recLocation."Global Dimension 2 Code";
            END;

            IF LocCode = '' THEN
                ERROR('Location Filter is Blank');

            CSOMapping.RESET;
            CSOMapping.SETRANGE(CSOMapping."User Id", UPPERCASE(USERID));
            CSOMapping.SETRANGE(Type, CSOMapping.Type::Location);
            CSOMapping.SETRANGE(CSOMapping.Value, LocCode);
            IF NOT (CSOMapping.FINDFIRST) THEN BEGIN
                CSOMapping.RESET;
                CSOMapping.SETRANGE(CSOMapping."User Id", UPPERCASE(USERID));
                CSOMapping.SETRANGE(CSOMapping.Type, CSOMapping.Type::"Responsibility Center");
                CSOMapping.SETRANGE(CSOMapping.Value, LocResC);
                IF NOT (CSOMapping.FINDFIRST) THEN
                    ERROR('You are not allowed to run this report other than your location');
            END;
        END;
        //EBT STIVAN ---(20092012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END


        CLEAR(LocCode);
        LocCode := Item.GETFILTER(Item."Location Filter");

        IF LocCode = '' THEN
            ERROR('Location Filter is Blank');

        SLNo := 0;
        LocFilter := '';
        CompanyInfo.GET;
        ExcelBuf.RESET;
        IF ExcelBuf.FINDSET THEN
            ExcelBuf.DELETEALL;
        MakeExcelInfo;

        //<<1

    end;

    var
        OpeningBalance: Decimal;
        OpeningBalanceSOM: Decimal;
        ItemRec: Record 27;
        LastDate: Date;
        PosDate: Date;
        LocFilter: Code[20];
        ItemUOM: Record 5404;
        QtyPerUOm: Decimal;
        MfgQty: Decimal;
        MfgQtySOM: Decimal;
        TransferRecvQty: Decimal;
        TransferRecvQtySOM: Decimal;
        SalesReturnQty: Decimal;
        SalesReturnQtySOM: Decimal;
        TotalRecvQty: Decimal;
        TotalRecvQtySOM: Decimal;
        SalesQty: Decimal;
        SalesQtySOM: Decimal;
        TransferShipQty: Decimal;
        TransferShipQtySOM: Decimal;
        TotalDespatchQty: Decimal;
        TotalDespatchQtySOM: Decimal;
        ClosingBalance: Decimal;
        ClosingBalanceSOM: Decimal;
        ExcelBuf: Record 370 temporary;
        SLNo: Integer;
        ILE: Record 32;
        StartDate: Date;
        EndDate: Date;
        CompanyInfo: Record 79;
        CSOMapping: Record 50006;
        LocCode: Code[10];
        LocResC: Code[10];
        recLocation: Record 14;
        user: Record 91;
        ConsumedQty: Decimal;
        ConsumedinSOM: Decimal;
        PosAdj: Decimal;
        NegAdj: Decimal;
        TotalPosAdj: Decimal;
        TotalNegAdj: Decimal;
        PosAdjQtySom: Decimal;
        NegAdjQtySom: Decimal;
        "-----12May2017": Integer;
        Loc12: Code[50];
        recLoc12: Record "14";
        TOpeningBalanceSOM: Decimal;
        TOpeningBalance: Decimal;
        TMfgQtySOM: Decimal;
        TMfgQty: Decimal;
        TTransferRecvQtySOM: Decimal;
        TTransferRecvQty: Decimal;
        TPosAdjQtySom: Decimal;
        TPosAdj: Decimal;
        TSalesReturnQtySOM: Decimal;
        TSalesReturnQty: Decimal;
        TTotalRecvQtySOM: Decimal;
        TTotalRecvQty: Decimal;
        TSalesQtySOM: Decimal;
        TSalesQty: Decimal;
        TTransferShipQtySOM: Decimal;
        TTransferShipQty: Decimal;
        TConsumedinSOM: Decimal;
        TConsumedQty: Decimal;
        TNegAdjQtySom: Decimal;
        TNegAdj: Decimal;
        TTotalDespatchQtySOM: Decimal;
        TTotalDespatchQty: Decimal;
        TClosingBalanceSOM: Decimal;
        TClosingBalance: Decimal;
        Memberof: Record 2000000053;

    //  [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//11May2017
        ExcelBuf.AddInfoColumn(FORMAT('Company Name'), FALSE, TRUE, FALSE, FALSE, '', 1);//11May2017
        //ExcelBuf.AddInfoColumn(FORMAT('Comapny Name'),FALSE,'',TRUE,FALSE,FALSE,'');
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT('Report Name'), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(FORMAT('FG Movement'), FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT('User ID'), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT('Date'), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT('Item Filter'), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(Item.GETFILTERS, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>12May2017 Report Header

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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('FG Movement', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter : ' + Item.GETFILTER(Item."Date Filter"), FALSE, '', TRUE, FALSE, TRUE, '@', 1);

        CLEAR(Loc12);
        Loc12 := Item.GETFILTER(Item."Location Filter");

        recLoc12.RESET;
        recLoc12.SETRANGE(Code, Loc12);
        IF recLoc12.FINDFIRST THEN;

        //ExcelBuf.NewRow;
        //ExcelBuf.AddColumn('Location : '+recLoc12.Name,FALSE,'',TRUE,FALSE,TRUE,'@',1);

        //<<12May2017 Report Header

        /*
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(CompanyInfo.Name,FALSE,'',TRUE,FALSE,TRUE,'@',1);//1
        
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('FG Stock Movement',FALSE,'',TRUE,FALSE,TRUE,'@',1);//1
        ExcelBuf.AddColumn('For The Location..',FALSE,'',TRUE,FALSE,TRUE,'@',1);//2
        ExcelBuf.AddColumn(Item.GETFILTER(Item."Location Filter"),FALSE,'',TRUE,FALSE,TRUE,'@',1);//3
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Movement for the Period..',FALSE,'',TRUE,FALSE,TRUE,'@',1);//1
        ExcelBuf.AddColumn(Item.GETFILTER(Item."Date Filter"),FALSE,'',TRUE,FALSE,TRUE,'@',1);//2
        */
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location : ' + recLoc12.Name, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1 //12May2017
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 //12May2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Sl No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2 //12May2017
        //ExcelBuf.AddColumn('Location Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//2
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('ITEM Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Base UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Pack Size', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Packing Qty per Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Opening Balance in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Opening Balance in LTR/KGS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Qty Mfg in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Qty Mfg in LTR/KGS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Qty Received through Transfer in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Qty Received through Transfer In LTR/KGS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Positive Adjustment In UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Positive Adjustment', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Qty Received through Sale Return In Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Qty Received through Sale Return in KGS / LTR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Total Received in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('Total Received in KGS / LTR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Qty Despatch For Sales in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('Qty Despatch For Sales in KGS / LTRS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('Qty Despatch For Branch Transfer in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('Qty Despatch For Branch Transfer in KGS/LTR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('Qty Consumed for Production in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Qty Consumed for Production in KGS/LTR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('Negative Adjustement UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        ExcelBuf.AddColumn('Negative Adjustement', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('Total Dispatched in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('Total Dispatched in KGS / LTR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('Closing Balance in Sale UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('Closing Balance in KGS / LTR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31

    end;

    // [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(SLNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);//1
        ExcelBuf.AddColumn(recLoc12.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2 //12May2017
        //ExcelBuf.AddColumn(LocFilter,FALSE,'',FALSE,FALSE,FALSE,'',1);//2
        ExcelBuf.AddColumn(Item."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(Item."Base Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(Item."Sales Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(QtyPerUOm, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//7
        ExcelBuf.AddColumn(OpeningBalanceSOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//8
        TOpeningBalanceSOM += OpeningBalanceSOM; //12May2017

        ExcelBuf.AddColumn(OpeningBalance, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//9
        TOpeningBalance += OpeningBalance; //12May2017

        ExcelBuf.AddColumn(MfgQtySOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//10
        TMfgQtySOM += MfgQtySOM; //12May2017

        ExcelBuf.AddColumn(MfgQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//11
        TMfgQty += MfgQty; //12May2017

        ExcelBuf.AddColumn(TransferRecvQtySOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//12
        TTransferRecvQtySOM += TransferRecvQtySOM;//12May2017

        ExcelBuf.AddColumn(TransferRecvQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//13
        TTransferRecvQty += TransferRecvQty; //12May2017

        ExcelBuf.AddColumn(PosAdjQtySom, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//14
        TPosAdjQtySom += PosAdjQtySom; //12May2017

        ExcelBuf.AddColumn(PosAdj, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//15
        TPosAdj += PosAdj; //12May2017

        ExcelBuf.AddColumn(SalesReturnQtySOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//16
        TSalesReturnQtySOM += SalesReturnQtySOM; //12May2017

        ExcelBuf.AddColumn(SalesReturnQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//17
        TSalesReturnQty += SalesReturnQty; //12May2017

        ExcelBuf.AddColumn(TotalRecvQtySOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//18
        TTotalRecvQtySOM += TotalRecvQtySOM; //12May2017

        ExcelBuf.AddColumn(TotalRecvQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//19
        TTotalRecvQty += TotalRecvQty; //12May2017

        ExcelBuf.AddColumn(-SalesQtySOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//20
        TSalesQtySOM += SalesQtySOM;//12May2017

        ExcelBuf.AddColumn(-SalesQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//21
        TSalesQty += SalesQty; //12May2017

        ExcelBuf.AddColumn(-TransferShipQtySOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//22
        TTransferShipQtySOM += TransferShipQtySOM;//12May2017

        ExcelBuf.AddColumn(-TransferShipQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//23
        TTransferShipQty += TransferShipQty;//12May2017

        ExcelBuf.AddColumn(-ConsumedinSOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//24
        TConsumedinSOM += ConsumedinSOM;//12May2017

        ExcelBuf.AddColumn(-ConsumedQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//25
        TConsumedQty += ConsumedQty; //12May2017

        ExcelBuf.AddColumn(ABS(NegAdjQtySom), FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//26
        TNegAdjQtySom += NegAdjQtySom; //12May2017

        ExcelBuf.AddColumn(ABS(NegAdj), FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27
        TNegAdj += NegAdj; //12May2017

        ExcelBuf.AddColumn(-TotalDespatchQtySOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//28
        TTotalDespatchQtySOM += TotalDespatchQtySOM;//12May2017

        ExcelBuf.AddColumn(-TotalDespatchQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//29
        TTotalDespatchQty += TotalDespatchQty;//12May2017

        ExcelBuf.AddColumn(ClosingBalanceSOM, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//30
        TClosingBalanceSOM += ClosingBalanceSOM; //12May2017

        ExcelBuf.AddColumn(ClosingBalance, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//31
        TClosingBalance += ClosingBalance; //12May2017
    end;

    // [Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet('Data','FG Movement',COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>12May2017 RB-N

        //ExcelBuf.CreateBook('','FG Movement');
        // ExcelBuf.CreateBookAndOpenExcel('','FG Movement','','',USERID);
        //  ExcelBuf.GiveUserControl;

        ExcelBuf.CreateNewBook('FG Movement');
        ExcelBuf.WriteSheet('FG Movement', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('FG Movement', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();


        //<<
    end;

    //  [Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 //12May2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 //12May2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total:', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn(recLoc12.Name, FALSE, '', TRUE, FALSE, TRUE, '', 1);//2 //12May2017
        //ExcelBuf.AddColumn(LocFilter,FALSE,'',TRUE,FALSE,TRUE,'',1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7 //12May2017
        //ExcelBuf.AddColumn(QtyPerUOm,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//7

        ExcelBuf.AddColumn(TOpeningBalanceSOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//8 //12May2017
        //ExcelBuf.AddColumn(OpeningBalanceSOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//8

        ExcelBuf.AddColumn(TOpeningBalance, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//9 //12May2017
        //ExcelBuf.AddColumn(OpeningBalance,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//9

        ExcelBuf.AddColumn(TMfgQtySOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//10 //12May2017
        //ExcelBuf.AddColumn(MfgQtySOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//10

        ExcelBuf.AddColumn(TMfgQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//11 //12May2017
        //ExcelBuf.AddColumn(MfgQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//11

        ExcelBuf.AddColumn(TTransferRecvQtySOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//12 //12May2017
        //ExcelBuf.AddColumn(TransferRecvQtySOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//12

        ExcelBuf.AddColumn(TTransferRecvQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//13
        //ExcelBuf.AddColumn(TransferRecvQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//13

        ExcelBuf.AddColumn(TPosAdjQtySom, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//14 //12May2017
        //ExcelBuf.AddColumn(PosAdjQtySom,FALSE,'',FALSE,FALSE,FALSE,'0.000',0);//14

        ExcelBuf.AddColumn(TPosAdj, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//15 //12May2017
        //ExcelBuf.AddColumn(PosAdj,FALSE,'',FALSE,FALSE,FALSE,'0.000',0);//15

        ExcelBuf.AddColumn(TSalesReturnQtySOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//16 //12May2017
        //ExcelBuf.AddColumn(SalesReturnQtySOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//16

        ExcelBuf.AddColumn(TSalesReturnQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//17 //12May2017
        //ExcelBuf.AddColumn(SalesReturnQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//17

        ExcelBuf.AddColumn(TTotalRecvQtySOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//18 //12May2017
        //ExcelBuf.AddColumn(TotalRecvQtySOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//18

        ExcelBuf.AddColumn(TTotalRecvQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//19 //12May2017
        //ExcelBuf.AddColumn(TotalRecvQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//19

        ExcelBuf.AddColumn(-TSalesQtySOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//20 //12May2017
        //ExcelBuf.AddColumn(-SalesQtySOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//20

        ExcelBuf.AddColumn(-TSalesQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//21 //12May2017
        //ExcelBuf.AddColumn(-SalesQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//21

        ExcelBuf.AddColumn(-TTransferShipQtySOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//22 //12May2017
        //ExcelBuf.AddColumn(-TransferShipQtySOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//22

        ExcelBuf.AddColumn(-TTransferShipQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//23 //12May2017
        //ExcelBuf.AddColumn(-TransferShipQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//23

        ExcelBuf.AddColumn(-TConsumedinSOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//24 //12May2017
        //ExcelBuf.AddColumn(-ConsumedinSOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//24

        ExcelBuf.AddColumn(-TConsumedQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//25 //12May2017
        //ExcelBuf.AddColumn(-ConsumedQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//25

        ExcelBuf.AddColumn(ABS(TNegAdjQtySom), FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//26 //12May2017
        //ExcelBuf.AddColumn(ABS(NegAdjQtySom),FALSE,'',FALSE,FALSE,FALSE,'0.000',0);//26

        ExcelBuf.AddColumn(ABS(TNegAdj), FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//27 //12May2017
        //ExcelBuf.AddColumn(ABS(NegAdj),FALSE,'',FALSE,FALSE,FALSE,'0.000',0);//27

        ExcelBuf.AddColumn(-TTotalDespatchQtySOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//28 //12May2017
        //ExcelBuf.AddColumn(-TotalDespatchQtySOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//28

        ExcelBuf.AddColumn(-TTotalDespatchQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//29 //12May2017
        //ExcelBuf.AddColumn(-TotalDespatchQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//29

        ExcelBuf.AddColumn(TClosingBalanceSOM, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//30 //12May2017
        //ExcelBuf.AddColumn(ClosingBalanceSOM,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//30

        ExcelBuf.AddColumn(TClosingBalance, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//31 //12May2017
        //ExcelBuf.AddColumn(ClosingBalance,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//31
    end;
}


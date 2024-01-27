report 50015 "In-Transit Item Valuation"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/InTransitItemValuation.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Transfer Line"; "Transfer Line")
        {
            DataItemTableView = SORTING("Transfer-to Code", "Document No.", "Line No.");
            RequestFilterFields = "Transfer-to Code", "Item No.";

            trigger OnAfterGetRecord()
            begin

                CLEAR(BRTNO);

                TransferHeader.RESET;
                TransferHeader.SETRANGE("No.", "Document No.");
                IF TransferHeader.FIND('-') THEN BEGIN

                    Location.RESET;
                    Location.SETRANGE(Location.Code, TransferHeader."Transfer-from Code");
                    IF Location.FIND('-') THEN
                        TransferFrom := Location.Name;

                    Location.RESET;
                    Location.SETRANGE(Location.Code, TransferHeader."Transfer-to Code");
                    IF Location.FIND('-') THEN
                        TransferTo := Location.Name;
                END;

                Item.RESET;
                Item.SETRANGE("No.", "Item No.");
                IF Item.FIND('-') THEN BEGIN
                    Description := Item.Description;
                    Description2 := Item.Description;
                END;


                RecTransShipHeader.RESET;
                RecTransShipHeader.SETRANGE("Transfer Order No.", "Document No.");
                IF RecTransShipHeader.FINDFIRST THEN BEGIN
                    BRTNO := RecTransShipHeader."No.";
                    BRTDATETO := RecTransShipHeader."Posting Date";
                END;

                IF Recitem.GET("Item No.") THEN;

                // EBT MILAN ADDED TO SHOW BATCH NO--------120314-----------------------------------------START
                TranslineLotno := '';
                Reservationentry.RESET;
                Reservationentry.SETRANGE(Reservationentry."Source Type", 5741);
                Reservationentry.SETRANGE(Reservationentry."Source ID", "Transfer Line"."Document No.");
                Reservationentry.SETRANGE(Reservationentry."Source Prod. Order Line", "Transfer Line"."Line No.");
                IF Reservationentry.FINDFIRST THEN
                    TranslineLotno := Reservationentry."Lot No.";

                // EBT MILAN ADDED TO SHOW BATCH NO--------120314-------------------------------------------END

                IF PrintToExcel THEN
                    MakeExcelDataBody;
            end;

            trigger OnPreDataItem()
            begin

                LastFieldNo := FIELDNO("Transfer-to Code");
                SETFILTER("Transfer Line"."Shipment Date", '<=%1', TillDate);
                SETRANGE("Transfer Line"."Transfer-from Code", 'IN-TRANS');

                //EBT STIVAN ---(23102012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
                CSOMapping2.RESET;
                CSOMapping2.SETRANGE(CSOMapping2."User Id", UPPERCASE(USERID));
                CSOMapping2.SETRANGE(Type, CSOMapping2.Type::Location);
                CSOMapping2.SETRANGE(CSOMapping2.Value, LocCode);
                IF CSOMapping2.FINDFIRST THEN
                    MARK := TRUE;
                BEGIN
                    IF LocCode <> '' THEN
                        "Transfer Line".SETRANGE("Transfer Line"."Transfer-to Code", LocCode);
                END;
                //EBT STIVAN ---(23102012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
            end;
        }
        dataitem("Transfer Receipt Header"; "Transfer Receipt Header")
        {
            DataItemTableView = SORTING("No.");
            dataitem("Transfer Receipt Line"; "Transfer Receipt Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));

                trigger OnAfterGetRecord()
                begin

                    Location.RESET;
                    Location.SETRANGE(Location.Code, "Transfer Receipt Header"."Transfer-from Code");
                    IF Location.FIND('-') THEN BEGIN
                        TransferFrom := Location.Name;
                    END;

                    Location.RESET;
                    Location.SETRANGE(Location.Code, "Transfer Receipt Header"."Transfer-to Code");
                    IF Location.FIND('-') THEN BEGIN
                        TransferTo := Location.Name;
                    END;

                    Item.RESET;
                    Item.SETRANGE("No.", "Transfer Receipt Line"."Item No.");
                    IF Item.FIND('-') THEN BEGIN
                        Description := Item.Description;
                        Description2 := Item.Description;
                    END;

                    CLEAR(BRTNO);
                    RecTransShipHeader.RESET;
                    RecTransShipHeader.SETRANGE("Transfer Order No.", "Transfer Receipt Line"."Transfer Order No.");
                    IF RecTransShipHeader.FINDFIRST THEN BEGIN
                        BRTNO := RecTransShipHeader."No.";
                        BRTDate := RecTransShipHeader."Posting Date";
                    END;


                    CLEAR(remQty);
                    ILE.RESET;
                    ILE.SETRANGE(ILE."Document No.", "Transfer Receipt Line"."Document No.");
                    ILE.SETRANGE(ILE."Item No.", "Transfer Receipt Line"."Item No.");
                    IF ILE.FINDFIRST THEN BEGIN
                        ILE1.SETRANGE(ILE1."Order No.", "Transfer Receipt Line"."Transfer Order No.");
                        ILE1.SETRANGE(ILE1."Item No.", "Transfer Receipt Line"."Item No.");
                        ILE1.SETFILTER(ILE1."Location Code", '%1', 'IN-TRANS');
                        ILE1.SETRANGE(ILE1."Document No.", "Transfer Receipt Line"."Document No.");
                        // ILE1.SETFILTER(ILE1."Posting Date",'<=%1',TillDate); // ADDED Doc. no. filter and remove posting date filter 100114  MILAN
                        ILE1.SETRANGE(ILE1."Document Line No.", "Transfer Receipt Line"."Line No.");
                        IF ILE1.FINDSET THEN
                            REPEAT
                                remQty += ILE1.Quantity;
                            UNTIL ILE1.NEXT = 0;
                    END;
                    remQty := ABS(remQty);

                    CLEAR(Shipmentqty);
                    ILE2.RESET;
                    ILE2.SETRANGE(ILE2."Order No.", "Transfer Receipt Line"."Transfer Order No.");
                    ILE2.SETRANGE(ILE2."Item No.", "Transfer Receipt Line"."Item No.");
                    ILE2.SETFILTER(ILE2."Location Code", '<>%1', 'IN-TRANS');
                    ILE2.SETFILTER(ILE2."Posting Date", '<=%1', TillDate);
                    ILE2.SETFILTER(ILE2."Document Type", 'Transfer Shipment');
                    IF ILE2.FINDSET THEN
                        REPEAT
                            Shipmentqty += ILE2.Quantity;
                            uuom := ILE2."Qty. per Unit of Measure";
                        UNTIL ILE2.NEXT = 0;

                    IF uuom <> 0 THEN
                        ShipmentQtyUOM := ILE2.Quantity / uuom;

                    /*
                    ILE.SETFILTER(ILE."Location Code",'<>%1','IN-TRANS');
                    IF ILE.FINDFIRST THEN
                    remQty := ILE."Remaining Quantity";
                    */

                    CLEAR(InvoiceRate);
                    CLEAR(GRNNo);
                    RecILE.RESET;
                    RecILE.SETRANGE(RecILE."Document No.", BRTNO);
                    RecILE.SETRANGE(RecILE."Item No.", "Transfer Receipt Line"."Item No.");
                    IF RecILE.FINDFIRST THEN BEGIN
                        ItemAppEntry.RESET;
                        ItemAppEntry.SETRANGE(ItemAppEntry."Outbound Item Entry No.", RecILE."Entry No.");
                        IF ItemAppEntry.FINDFIRST THEN BEGIN
                            RecILE1.RESET;
                            RecILE1.SETRANGE(RecILE1."Entry No.", ItemAppEntry."Inbound Item Entry No.");
                            IF RecILE1.FINDFIRST THEN BEGIN
                                IF RecILE1."Entry Type" = RecILE1."Entry Type"::Purchase THEN BEGIN
                                    GRNNo := RecILE1."Document No.";
                                    Valueentry.RESET;
                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE1."Entry No.");
                                    IF Valueentry.FINDFIRST THEN BEGIN
                                        InvoiceRate := Valueentry."Cost per Unit";
                                    END;
                                END;
                                IF RecILE1."Entry Type" = RecILE1."Entry Type"::Output THEN BEGIN
                                    GRNNo := RecILE1."Document No.";
                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE1."Entry No.");
                                    IF Valueentry.FINDSET THEN
                                        REPEAT
                                            InvoiceRate += Valueentry."Cost per Unit";
                                        UNTIL Valueentry.NEXT = 0;
                                END;
                                IF RecILE1."Entry Type" = RecILE1."Entry Type"::"Positive Adjmt." THEN BEGIN
                                    GRNNo := RecILE1."Document No.";
                                END;
                                //  ELSE
                                //  BEGIN
                                IF RecILE1."Entry Type" = RecILE1."Entry Type"::Transfer THEN BEGIN
                                    RecILE3.SETRANGE(RecILE3."Order No.", RecILE1."Order No.");
                                    RecILE3.SETRANGE(RecILE3."Item No.", "Transfer Receipt Line"."Item No.");
                                    RecILE3.SETFILTER(RecILE3."Location Code", '<>%1', 'IN-TRANSE');
                                    IF RecILE3.FINDFIRST THEN BEGIN
                                        ItemAppEntry1.RESET;
                                        ItemAppEntry1.SETRANGE(ItemAppEntry1."Outbound Item Entry No.", RecILE3."Entry No.");
                                        IF ItemAppEntry1.FINDFIRST THEN BEGIN
                                            RecILE2.RESET;
                                            RecILE2.SETRANGE(RecILE2."Entry No.", ItemAppEntry1."Inbound Item Entry No.");
                                            IF RecILE2.FINDFIRST THEN BEGIN
                                                IF RecILE2."Entry Type" = RecILE2."Entry Type"::Purchase THEN BEGIN
                                                    GRNNo := RecILE2."Document No.";
                                                    Valueentry.RESET;
                                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE2."Entry No.");
                                                    IF Valueentry.FINDFIRST THEN BEGIN
                                                        InvoiceRate := Valueentry."Cost per Unit";
                                                    END;
                                                END;
                                                IF RecILE2."Entry Type" = RecILE2."Entry Type"::Output THEN BEGIN
                                                    GRNNo := RecILE2."Document No.";
                                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE2."Entry No.");
                                                    IF Valueentry.FINDSET THEN
                                                        REPEAT
                                                            InvoiceRate += Valueentry."Cost per Unit";
                                                        UNTIL Valueentry.NEXT = 0;
                                                END;
                                                IF RecILE2."Entry Type" = RecILE2."Entry Type"::"Positive Adjmt." THEN BEGIN
                                                    GRNNo := RecILE2."Document No.";
                                                END;
                                                IF RecILE2."Entry Type" = RecILE2."Entry Type"::Transfer THEN BEGIN
                                                    RecILE4.RESET;
                                                    RecILE4.SETRANGE(RecILE4."Order No.", RecILE2."Order No.");
                                                    RecILE4.SETRANGE(RecILE4."Item No.", "Transfer Receipt Line"."Item No.");
                                                    RecILE4.SETFILTER(RecILE4."Location Code", '<>%1', 'IN-TRANSE');
                                                    IF RecILE4.FINDFIRST THEN BEGIN
                                                        ItemAppEntry2.RESET;
                                                        ItemAppEntry2.SETRANGE(ItemAppEntry2."Outbound Item Entry No.", RecILE4."Entry No.");
                                                        IF ItemAppEntry2.FINDFIRST THEN BEGIN
                                                            RecILE5.RESET;
                                                            RecILE5.SETRANGE(RecILE5."Entry No.", ItemAppEntry2."Inbound Item Entry No.");
                                                            IF RecILE5.FINDFIRST THEN BEGIN
                                                                IF RecILE5."Entry Type" = RecILE5."Entry Type"::Purchase THEN BEGIN
                                                                    GRNNo := RecILE5."Document No.";
                                                                    Valueentry.RESET;
                                                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE5."Entry No.");
                                                                    IF Valueentry.FINDFIRST THEN BEGIN
                                                                        InvoiceRate := Valueentry."Cost per Unit";
                                                                    END;
                                                                END;
                                                                IF RecILE5."Entry Type" = RecILE5."Entry Type"::Output THEN BEGIN
                                                                    GRNNo := RecILE5."Document No.";
                                                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE5."Entry No.");
                                                                    IF Valueentry.FINDSET THEN
                                                                        REPEAT
                                                                            InvoiceRate += Valueentry."Cost per Unit";
                                                                        UNTIL Valueentry.NEXT = 0;
                                                                END;
                                                                IF RecILE5."Entry Type" = RecILE5."Entry Type"::Transfer THEN BEGIN
                                                                    RecILE6.RESET;
                                                                    RecILE6.SETRANGE(RecILE6."Order No.", RecILE5."Order No.");
                                                                    RecILE6.SETRANGE(RecILE6."Item No.", "Transfer Receipt Line"."Item No.");
                                                                    RecILE6.SETFILTER(RecILE6."Location Code", '<>%1', 'IN-TRANSE');
                                                                    IF RecILE6.FINDFIRST THEN BEGIN
                                                                        ItemAppEntry3.RESET;
                                                                        ItemAppEntry3.SETRANGE(ItemAppEntry3."Outbound Item Entry No.", RecILE6."Entry No.");
                                                                        IF ItemAppEntry3.FINDFIRST THEN BEGIN
                                                                            RecILE7.RESET;
                                                                            RecILE7.SETRANGE(RecILE7."Entry No.", ItemAppEntry3."Inbound Item Entry No.");
                                                                            IF RecILE7.FINDFIRST THEN BEGIN
                                                                                IF RecILE7."Entry Type" = RecILE7."Entry Type"::Purchase THEN BEGIN
                                                                                    GRNNo := RecILE7."Document No.";
                                                                                    Valueentry.RESET;
                                                                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE7."Entry No.");
                                                                                    IF Valueentry.FINDFIRST THEN BEGIN
                                                                                        InvoiceRate := Valueentry."Cost per Unit";
                                                                                    END;
                                                                                END;
                                                                                IF RecILE7."Entry Type" = RecILE7."Entry Type"::Output THEN BEGIN
                                                                                    GRNNo := RecILE1."Document No.";
                                                                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE7."Entry No.");
                                                                                    IF Valueentry.FINDSET THEN
                                                                                        REPEAT
                                                                                            InvoiceRate += Valueentry."Cost per Unit";
                                                                                        UNTIL Valueentry.NEXT = 0;
                                                                                END;
                                                                                IF RecILE7."Entry Type" = RecILE7."Entry Type"::Transfer THEN BEGIN
                                                                                    RecILE8.RESET;
                                                                                    RecILE8.SETRANGE(RecILE8."Order No.", RecILE7."Order No.");
                                                                                    RecILE8.SETRANGE(RecILE8."Item No.", "Transfer Receipt Line"."Item No.");
                                                                                    RecILE8.SETFILTER(RecILE8."Location Code", '<>%1', 'IN-TRANSE');
                                                                                    IF RecILE8.FINDFIRST THEN BEGIN
                                                                                        ItemAppEntry4.RESET;
                                                                                        ItemAppEntry4.SETRANGE(ItemAppEntry4."Outbound Item Entry No.", RecILE8."Entry No.");
                                                                                        IF ItemAppEntry4.FINDFIRST THEN BEGIN
                                                                                            RecILE9.RESET;
                                                                                            RecILE9.SETRANGE(RecILE9."Entry No.", ItemAppEntry4."Inbound Item Entry No.");
                                                                                            IF RecILE9.FINDFIRST THEN BEGIN
                                                                                                IF RecILE9."Entry Type" = RecILE9."Entry Type"::Purchase THEN BEGIN
                                                                                                    GRNNo := RecILE9."Document No.";
                                                                                                    Valueentry.RESET;
                                                                                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE9."Entry No.");
                                                                                                    IF Valueentry.FINDFIRST THEN BEGIN
                                                                                                        InvoiceRate := Valueentry."Cost per Unit";
                                                                                                    END;
                                                                                                END;
                                                                                                IF RecILE9."Entry Type" = RecILE9."Entry Type"::Output THEN BEGIN
                                                                                                    GRNNo := RecILE1."Document No.";
                                                                                                    Valueentry.SETRANGE(Valueentry."Item Ledger Entry No.", RecILE9."Entry No.");
                                                                                                    IF Valueentry.FINDSET THEN
                                                                                                        REPEAT
                                                                                                            InvoiceRate += Valueentry."Cost per Unit";
                                                                                                        UNTIL Valueentry.NEXT = 0;
                                                                                                END;
                                                                                            END;
                                                                                        END;
                                                                                    END;
                                                                                END;
                                                                            END;
                                                                        END;
                                                                    END;
                                                                END;
                                                            END;
                                                        END;
                                                    END;
                                                END;
                                            END
                                        END;
                                    END;
                                END;
                            END;
                        END;
                    END;


                    CLEAR(Invno);
                    CLEAR(bondedrate);
                    CLEAR(ExBondedrate);
                    CLEAR(Vendcode);
                    CLEAR(VendName);
                    CLEAR("vendinvno.");
                    Purchinvline.RESET;
                    Purchinvline.SETRANGE(Purchinvline."Receipt No.", GRNNo);
                    Purchinvline.SETRANGE(Purchinvline."No.", "Transfer Receipt Line"."Item No.");
                    IF Purchinvline.FINDFIRST THEN BEGIN
                        Invno := Purchinvline."Document No.";
                        bondedrate := Purchinvline."Bonded Rate";
                        ExBondedrate := Purchinvline."Exbond Rate";
                        Vendcode := Purchinvline."Buy-from Vendor No.";
                        recven.SETRANGE(recven."No.", Purchinvline."Buy-from Vendor No.");
                        IF recven.FINDFIRST THEN
                            VendName := recven.Name;
                        PurchinvHeader.RESET;
                        PurchinvHeader.SETRANGE(PurchinvHeader."No.", Purchinvline."Document No.");
                        IF PurchinvHeader.FINDFIRST THEN BEGIN
                            "vendinvno." := PurchinvHeader."Vendor Invoice No.";
                        END;

                    END;

                    IF Recitem.GET("Item No.") THEN;

                    // EBT MILAN ADDED TO SHOW BATCH NO--------120314-----------------------------------------START
                    TranslineReceiptLotno := '';
                    ILErecc.RESET;
                    ILErecc.SETRANGE(ILErecc."Document No.", "Transfer Receipt Line"."Document No.");
                    ILErecc.SETRANGE(ILErecc."Item No.", "Transfer Receipt Line"."Item No.");
                    ILErecc.SETRANGE(ILErecc."Document Line No.", "Transfer Receipt Line"."Line No.");
                    IF ILErecc.FINDFIRST THEN
                        TranslineReceiptLotno := ILErecc."Lot No.";

                    // EBT MILAN ADDED TO SHOW BATCH NO--------120314-------------------------------------------END
                    // EBT MILAN ADDED TO SHOW Purch/Blend order Date--------120314-------------------------------------------START
                    DoccDate := 0D;
                    ILEreccc.RESET;
                    ILEreccc.SETRANGE(ILEreccc."Document No.", GRNNo);
                    ILEreccc.SETRANGE(ILEreccc."Item No.", "Transfer Receipt Line"."Item No.");
                    ILEreccc.SETFILTER(ILEreccc."Entry Type", '%1|%2|%3', ILEreccc."Entry Type"::Purchase, ILEreccc."Entry Type"::Output
                    , ILEreccc."Entry Type"::"Positive Adjmt.");
                    IF ILEreccc.FINDFIRST THEN
                        DoccDate := ILEreccc."Posting Date";
                    // EBT MILAN ADDED TO SHOW Purch/Blend order Date--------120314-------------------------------------------END

                    IF PrintToExcel THEN
                        MakeExcelTrnsfReceioptBody;

                end;
            }

            trigger OnPreDataItem()
            begin

                CalcReciptDate := CALCDATE('+1D', TillDate);
                //"Transfer Receipt Header".SETFILTER("Transfer Receipt Header"."Receipt Date",'%1..%2',CalcReciptDate,TODAY);
                "Transfer Receipt Header".SETFILTER("Transfer Receipt Header"."Shipment Date", '<=%1', TillDate);
                "Transfer Receipt Header".SETFILTER("Transfer Receipt Header"."Posting Date", '%1..%2', CalcReciptDate, TODAY);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(TillDate; TillDate)
                {
                    Caption = 'Till Date';
                }
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

    trigger OnPreReport()
    begin

        fromlocfilt := "Transfer Line".GETFILTER("Transfer Line"."Transfer-from Code");
        tolocfilt := "Transfer Line".GETFILTER("Transfer Line"."Transfer-to Code");

        RecLoc1.RESET;
        IF RecLoc1.GET(fromlocfilt) THEN
            FromLoc := RecLoc1.Name;

        RecLoc1.RESET;
        IF RecLoc1.GET(tolocfilt) THEN
            ToLoc := RecLoc1.Name;


        //EBT STIVAN ---(23102012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        LocCode := "Transfer Line".GETFILTER("Transfer Line"."Transfer-to Code");
        /*
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
        //EBT STIVAN ---(23102012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        */
        IF PrintToExcel THEN
            MakeExcelInfo;

    end;

    var
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        TransferHeader: Record 5740;
        TransferFrom: Text;
        TransferTo: Text;
        Description: Text[100];
        Location: Record 14;
        Item: Record 27;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        Description2: Text[100];
        usersetup: Record 91;
        LocFilter: Text[30];
        RecLoc: Record 14;
        RecTransHeader: Record 5740;
        RecTransShipHeader: Record 5744;
        BRTNO: Code[30];
        fromlocfilt: Text[30];
        tolocfilt: Text[30];
        RecLoc1: Record 14;
        FromLoc: Text;
        ToLoc: Text;
        datefilt: Text[30];
        Recitem: Record 27;
        LocCode: Code[100];
        LocResC: Code[20];
        recLocation: Record 14;
        CSOMapping1: Record 50006;
        CSOMapping2: Record 50006;
        TillDate: Date;
        CalcReciptDate: Date;
        RecILE: Record 32;
        ItemAppEntry: Record 339;
        RecILE1: Record 32;
        ItemAppEntry1: Record 339;
        RecILE2: Record 32;
        GRNNo: Code[20];
        OPENINGPOSITIVE: Text[30];
        RecILE3: Record 32;
        RecILE4: Record 32;
        RecILE5: Record 32;
        RecILE6: Record 32;
        GRNNo1: Code[20];
        ItemAppEntry2: Record 339;
        ItemAppEntry3: Record 339;
        RecILE7: Record 32;
        remQty: Decimal;
        ILE: Record 32;
        InvoiceRate: Decimal;
        Purchinvline: Record 123;
        Valueentry: Record 5802;
        bondedrate: Decimal;
        ExBondedrate: Decimal;
        Invno: Code[20];
        Vendcode: Code[20];
        VendName: Text[50];
        recven: Record 23;
        PurchinvHeader: Record 122;
        "vendinvno.": Code[20];
        RecILE8: Record 32;
        RecILE9: Record 32;
        ItemAppEntry4: Record 339;
        ILE1: Record 32;
        ILE2: Record 32;
        Shipmentqty: Decimal;
        ShipmentQtyUOM: Decimal;
        uuom: Decimal;
        BRTDate: Date;
        BRTDATETO: Date;
        Reservationentry: Record 337;
        TranslineLotno: Code[20];
        TranslineReceiptLotno: Code[20];
        ILErecc: Record 32;
        ILEreccc: Record 32;
        DoccDate: Date;
        Text001: Label 'In-Transit Statement';
        Text000: Label 'Data';
        Text0001: Label 'Khan';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('In-Transit Statement', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"In-Transit Item Availability", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('In-Transit Statement', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        //>>RoboSoft/Rahul
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);   // EBT MILAN ADDED 120314
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);   // EBT MILAN ADDED 120314
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);  // EBT MILAN 120314
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        //<<
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('In-Transit Valuation Report as on ' + FORMAT(TillDate), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        //>>RoboSoft/Rahul
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);   // EBT MILAN ADDED 120314
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);   // EBT MILAN ADDED 120314
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);  // EBT MILAN 120314
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        //<<

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Transfer Receipt No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Date of Subsequence GRN', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item N0.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);   // EBT MILAN ADDED 120314
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);   // EBT MILAN ADDED 120314
        ExcelBuf.AddColumn('Base UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty. per UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Shipment Qty in Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Shipment Qty in LT/KG', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('From Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('To Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BRT No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BRT Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Transfer Price Per Unit', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('eCess Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('SHE Cess Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Total Amount Incl. Excise', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRN / Blend Order No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRN / Blend Order Date.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);  // EBT MILAN 120314
        ExcelBuf.AddColumn('Remaining Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Purchase / Blend Cost', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Inventory Value', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Bonded Rate', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('ExBond Rate', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Transfer Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line"."Shipment Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn("Transfer Line"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Description2, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TranslineLotno, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Recitem."Base Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("Transfer Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TransferFrom, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TransferTo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BRTNO, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BRTDATETO, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn("Transfer Line"."Transfer Price of Base Unit", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("Transfer Line"."Transfer Price of Base Unit" * "Transfer Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("Transfer Line".Amount + 0 +
                          0 + 0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        //ExcelBuf.AddColumn(GRNNo,FALSE,'',FALSE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn(OPENINGPOSITIVE,FALSE,'',FALSE,FALSE,FALSE,'');
    end;

    ////[Scope('Internal')]
    procedure MakeExcelTrnsfReceioptBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Transfer Receipt Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn("Transfer Receipt Line"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Description2, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); // EBT MILAN ADDED 120314
        ExcelBuf.AddColumn(TranslineReceiptLotno, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text); // EBT MILAN ADDED 120314
        ExcelBuf.AddColumn(Recitem."Base Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn("Transfer Receipt Line".Quantity,FALSE,'',FALSE,FALSE,FALSE,'');
        //ExcelBuf.AddColumn("Transfer Receipt Line"."Quantity (Base)",FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(ShipmentQtyUOM), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ABS(Shipmentqty), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TransferFrom, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TransferTo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BRTNO, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BRTDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Line"."Unit Price" /
        "Transfer Receipt Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(("Transfer Receipt Line"."Unit Price" / "Transfer Receipt Line"."Qty. per Unit of Measure")
        * "Transfer Receipt Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("Transfer Receipt Line".Amount + 0 +
                           0 + 0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(Vendcode, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(VendName, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("vendinvno.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Invno, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(GRNNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(DoccDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn(remQty, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(InvoiceRate, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(InvoiceRate * remQty, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(bondedrate, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(ExBondedrate, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000 ,Text001,COMPANYNAME,USERID);
        //ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', '');
        //ExcelBuf.GiveUserControl;

        ExcelBuf.CreateNewBook(Text000);
        ExcelBuf.WriteSheet(Text000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelGroupdata()
    begin
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelGroupFooter()
    begin
    end;
}


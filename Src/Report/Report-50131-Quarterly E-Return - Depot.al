report 50131 "Quarterly E-Return - Depot"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/QuarterlyEReturnDepot.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Cancelled Invoice" = CONST(false));
            RequestFilterFields = "Posting Date", "Location Code";
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Type = FILTER(Item | 'Charge (Item)'),
                                          Quantity = FILTER(<> 0));
                RequestFilterFields = "No.";
                dataitem("Item Ledger Entry"; "Item Ledger Entry")
                {
                    DataItemLink = "Lot No." = FIELD("Lot No."),
                                   "Location Code" = FIELD("Location Code"),
                                   "Item No." = FIELD("No.");
                    DataItemTableView = SORTING("Entry No.")
                                        WHERE("Lot No." = FILTER(<> ''),
                                              "Entry Type" = FILTER(<> Sale & <> "Negative Adjmt." & <> "Positive Adjmt."));

                    trigger OnAfterGetRecord()
                    begin

                        //>>1


                        CLEAR(TSHno);
                        CLEAR(TSHdate);
                        CLEAR(TROno);
                        CLEAR(No);
                        CLEAR(ILEEntryNo);
                        ValueEntry.RESET;
                        ValueEntry.SETRANGE(ValueEntry."Source Code", 'SALES');
                        ValueEntry.SETRANGE(ValueEntry."Document No.", "Sales Invoice Line"."Document No.");
                        ValueEntry.SETRANGE(ValueEntry."Item No.", "Sales Invoice Line"."No.");
                        ValueEntry.SETRANGE(ValueEntry."Location Code", "Sales Invoice Line"."Location Code");
                        ValueEntry.SETRANGE(ValueEntry."Document Line No.", "Sales Invoice Line"."Line No.");
                        IF ValueEntry.FINDFIRST THEN BEGIN
                            ILEEntryNo := ValueEntry."Item Ledger Entry No.";

                            CLEAR("ItemApplEntryNo.");
                            ItemApplEntry.RESET;
                            ItemApplEntry.SETRANGE(ItemApplEntry."Outbound Item Entry No.", ILEEntryNo);
                            IF ItemApplEntry.FINDFIRST THEN BEGIN
                                "ItemApplEntryNo." := ItemApplEntry."Inbound Item Entry No.";

                                ILE.RESET;
                                ILE.SETRANGE(ILE."Entry No.", "ItemApplEntryNo.");
                                IF ILE.FINDFIRST THEN BEGIN
                                    //TROno := ILE."Transfer Order No.";
                                    TROno := ILE."Order No."; //03Mar2017
                                    TSHno := ILE."Document No.";
                                    TSHdate := ILE."Posting Date";
                                    No := ILE."Document No.";

                                    IF ILE."Document Type" = ILE."Document Type"::"Transfer Receipt" THEN BEGIN
                                        recTSH.RESET;
                                        recTSH.SETRANGE(recTSH."Transfer Order No.", TROno);
                                        IF recTSH.FINDFIRST THEN BEGIN
                                            TSHno := recTSH."No.";
                                            TSHdate := recTSH."Posting Date";
                                            LocCode := recTSH."Transfer-from Code";
                                        END;

                                    END;
                                END;
                            END;

                            TrfQty := 0;
                            ExciseValue := 0;
                            Exciseduty := 0;
                            ecess := 0;
                            shecess := 0;
                            CLEAR(ItemNo);
                            /* recRG23D.RESET;
                            recRG23D.SETRANGE(recRG23D."Document No.", No);
                            recRG23D.SETRANGE(recRG23D."Item No.", "Sales Invoice Line"."No.");
                            recRG23D.SETRANGE(recRG23D."Location Code", "Sales Invoice Line"."Location Code");
                            IF recRG23D.FINDFIRST THEN BEGIN
                                ItemNo := recRG23D."Item No.";
                                TrfQty := recRG23D.Quantity;
                                // EBT MILAN 060214----------------------------------------------------------------START
                                Exciseduty += (recRG23D."BED Amount Per Unit" * recRG23D.Quantity * ItemUOM);
                                ecess += (recRG23D."eCess Amount Per Unit" * recRG23D.Quantity * ItemUOM);
                                shecess += (recRG23D."SHE Cess Amount Per Unit" * recRG23D.Quantity * ItemUOM);
                                // EBT MILAN 060214------------------------------------------------------------------END
                                ExciseValue += recRG23D.Amount;
                            END; */

                            Recitemtable.RESET;
                            Recitemtable.SETRANGE(Recitemtable."No.", ItemNo);
                            IF Recitemtable.FINDFIRST THEN BEGIN
                                recItemUom.RESET;
                                recItemUom.SETRANGE(recItemUom.Code, Recitemtable."Sales Unit of Measure");
                                IF recItemUom.FINDFIRST THEN
                                    ItemUOM := recItemUom."Qty. per Unit of Measure";
                            END;

                            IF (LocCode = 'PLANT01') OR (LocCode = 'Plant02') THEN
                                IssueBy := 'Manufacturer'
                            ELSE
                                IssueBy := 'Dealer';

                            CLEAR(ECCcode);
                            CLEAR(ECCDes);
                            CLEAR(LocName);
                            Location.RESET;
                            Location.SETRANGE(Location.Code, LocCode);
                            IF Location.FINDFIRST THEN BEGIN
                                ECCcode := '';//Location."E.C.C. No.";
                                LocName := Location.Name;
                                /* recECC.RESET;
                                recECC.SETRANGE(recECC.Code, ECCcode);
                                IF recECC.FINDFIRST THEN
                                    ECCDes := recECC."C.E. Registration No."; */
                            END;

                        END;

                        //MakeExcelDataBody2;

                        Transfershipmentline.RESET;
                        Transfershipmentline.SETRANGE(Transfershipmentline."Document No.", "Item Ledger Entry"."Document No.");
                        Transfershipmentline.SETRANGE(Transfershipmentline."Item No.", "Item Ledger Entry"."Item No.");
                        Transfershipmentline.SETRANGE(Transfershipmentline."Line No.", "Item Ledger Entry"."Document Line No.");
                        Transfershipmentline.SETRANGE(Transfershipmentline."Shipment Date", "Item Ledger Entry"."Posting Date");

                        IF (Location_REC.Code = 'PLANT01') OR (Location_REC.Code = 'Plant02') THEN
                            IssueBy := 'Manufacturer'
                        ELSE
                            IssueBy := 'Dealer';

                        /*  "E.C.C.".RESET;
                         "E.C.C.".SETRANGE("E.C.C.".Code, Location_REC."E.C.C. No.");
                         IF "E.C.C.".FINDFIRST THEN */
                        CLEAR(Intermvalue);
                        CLEAR(Value);


                        RecItem.GET("Item Ledger Entry"."Item No.");

                        Qty := Quantity;
                        QtyPerUOM := "Qty. per Unit of Measure";
                        IF QtyPerUOM <> 0 THEN
                            PackDescriptionforILE := FORMAT(QtyPerUOM) + ' ' + UOM + '*' + FORMAT(Qty / QtyPerUOM);

                        Intermvalue := "Item Ledger Entry".Quantity;
                        IF Intermvalue <> 0 THEN
                            Value := (/*"Item Ledger Entry"."BED Amount"*/0 * "Qty. per Unit of Measure") / Intermvalue
                        ELSE
                            Value := 0;

                        RecItem2.GET("Item Ledger Entry"."Item No.");

                        //<<
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    CLEAR(BRTNO);
                    CLEAR(BRNo);
                    CLEAR(TROno);
                    CLEAR(BRTExciseValue);
                    CLEAR(BRTQty);

                    CLEAR(BRTExciseDuty);
                    CLEAR(BRTecess);
                    CLEAR(BRTSheCess);

                    /* RG23Drec.RESET;
                    RG23Drec.SETRANGE(RG23Drec."Transaction Type", RG23Drec."Transaction Type"::Sale);
                    RG23Drec.SETRANGE(RG23Drec.Type, RG23Drec.Type::Invoice);
                    RG23Drec.SETRANGE(RG23Drec."Document No.", "Sales Invoice Line"."Document No.");
                    RG23Drec.SETRANGE(RG23Drec."Item No.", "Sales Invoice Line"."No.");
                    RG23Drec.SETRANGE(RG23Drec."Line No.", "Sales Invoice Line"."Line No.");
                    RG23Drec.SETRANGE(RG23Drec."Lot No.", "Sales Invoice Line"."Lot No.");
                    RG23Drec.SETRANGE(RG23Drec."Location Code", "Sales Invoice Line"."Location Code");
                    RG23Drec.SETFILTER(RG23Drec.Quantity, '<>%1', 0);
                    IF RG23Drec.FINDFIRST THEN
                        REPEAT */

                    /* recRG23Dtable1.RESET;
                    recRG23Dtable1.SETRANGE(recRG23Dtable1."Entry No.", RG23Drec."Ref. Entry No.");
                    IF recRG23Dtable1.FINDFIRST THEN BEGIN
                        // EBT MILAN 080114 ADDED to SHOW CMP NO. and QTY ----------------------------------------------------------START
                        IF (recRG23Dtable1."Transaction Type" = recRG23Dtable1."Transaction Type"::Purchase) AND
                        (recRG23Dtable1.Type = recRG23Dtable1.Type::Transfer) THEN BEGIN
                            BRNo := recRG23Dtable1."Document No.";
                            TROno := recRG23Dtable1."Order No.";
                        END ELSE */
                    // milan
                    //BEGIN
                    SalesCrMemo.RESET;
                    //SalesCrMemo.SETRANGE(SalesCrMemo."Document No.", recRG23Dtable1."Document No.");
                    //SalesCrMemo.SETRANGE(SalesCrMemo."No.", recRG23Dtable1."Item No.");
                    //SalesCrMemo.SETRANGE(SalesCrMemo."Line No.", recRG23Dtable1."Line No.");
                    IF SalesCrMemo.FINDFIRST THEN BEGIN
                        BRTNO := '';//recRG23Dtable1."Document No.";
                        BRTExciseValue := 0;//SalesCrMemo."Excise Amount";
                                            // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess --------------------------------------------START
                        BRTExciseDuty := 0;// SalesCrMemo."BED Amount";
                        BRTecess := 0;// SalesCrMemo."eCess Amount";
                        BRTSheCess := 0;// SalesCrMemo."SHE Cess Amount";
                                        // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess ----------------------------------------------END
                        BRTQty := SalesCrMemo."Quantity (Base)";
                        BRDate := 0D;//recRG23Dtable1."Posting Date";
                    END;
                    // EBT MILAN 080114 ADDED to SHOW CMP NO. and QTY ----------------------------------------------------------END

                    //END;
                    //

                    recTSH.RESET;
                    recTSH.SETRANGE(recTSH."Transfer Order No.", TROno);
                    IF recTSH.FINDFIRST THEN BEGIN
                        BRTNO := recTSH."No.";
                        BRDate := recTSH."Posting Date";
                        LocCode := recTSH."Transfer-from Code";
                    END;

                    Recitemtable.RESET;
                    Recitemtable.SETRANGE(Recitemtable."No.", "Sales Invoice Line"."No.");
                    IF Recitemtable.FINDFIRST THEN BEGIN
                        recItemUom.RESET;
                        recItemUom.SETRANGE(recItemUom.Code, Recitemtable."Sales Unit of Measure");
                        IF recItemUom.FINDFIRST THEN
                            ItemUOM := recItemUom."Qty. per Unit of Measure";
                    END;


                    TransferShipment.RESET;
                    TransferShipment.SETRANGE(TransferShipment."Document No.", BRTNO);
                    TransferShipment.SETRANGE(TransferShipment."Item No.", "Sales Invoice Line"."No.");
                    IF TransferShipment.FINDFIRST THEN BEGIN
                        BRTQty := TransferShipment."Quantity (Base)";
                        BRTExciseValue := 0;//TransferShipment."Excise Amount";
                                            // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess --------------------------------------------START
                        BRTExciseDuty := 0;//TransferShipment."BED Amount";
                        BRTecess := 0;//TransferShipment."eCess Amount";
                        BRTSheCess := 0;// TransferShipment."SHE Cess Amount";
                                        // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess ----------------------------------------------END
                    END ELSE BEGIN
                        /*  recRG23D.RESET;
                         recRG23D.SETRANGE(recRG23D."Transaction Type", recRG23D."Transaction Type"::Purchase);
                         recRG23D.SETRANGE(recRG23D.Type, recRG23D.Type::Transfer);
                         recRG23D.SETRANGE(recRG23D."Lot No.", "Sales Invoice Line"."Lot No.");
                         recRG23D.SETRANGE(recRG23D."Location Code", "Sales Invoice Line"."Location Code");
                         recRG23D.SETRANGE(recRG23D."Item No.", "Sales Invoice Line"."No.");
                         recRG23D.SETRANGE(recRG23D."Document No.", BRNo);
                         IF recRG23D.FINDFIRST THEN BEGIN
                             BRTQty := recRG23D.Quantity * ItemUOM;
                             BRTExciseValue := recRG23D.Amount;
                             // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess --------------------------------------------START
                             BRTExciseDuty := recRG23D."BED Amount Per Unit" * recRG23D.Quantity * ItemUOM;
                             BRTecess := recRG23D."eCess Amount Per Unit" * recRG23D.Quantity * ItemUOM;
                             BRTSheCess := recRG23D."SHE Cess Amount Per Unit" * recRG23D.Quantity * ItemUOM;
                             // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess ----------------------------------------------END
                             LocCode := recRG23D."Location Code";
                             BRTNO := recRG23D."Document No.";
                             BRDate := recRG23D."Posting Date";
                         END; */

                    END;

                    TransferShipment.RESET;
                    TransferShipment.SETRANGE(TransferShipment."Document No.", BRTNO);
                    TransferShipment.SETRANGE(TransferShipment."Item No.", "Sales Invoice Line"."No.");
                    IF TransferShipment.COUNT > 1 THEN BEGIN
                        /* recRG23D.RESET;
                        recRG23D.SETRANGE(recRG23D."Transaction Type", recRG23D."Transaction Type"::Purchase);
                        recRG23D.SETRANGE(recRG23D.Type, recRG23D.Type::Transfer);
                        recRG23D.SETRANGE(recRG23D."Lot No.", "Sales Invoice Line"."Lot No.");
                        recRG23D.SETRANGE(recRG23D."Location Code", "Sales Invoice Line"."Location Code");
                        recRG23D.SETRANGE(recRG23D."Item No.", "Sales Invoice Line"."No.");
                        recRG23D.SETRANGE(recRG23D."Document No.", BRNo);
                        IF recRG23D.FINDFIRST THEN BEGIN
                            BRTQty := recRG23D.Quantity * ItemUOM;
                            BRTExciseValue := recRG23D.Amount;
                            // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess --------------------------------------------START
                            BRTExciseDuty := recRG23D."BED Amount Per Unit" * recRG23D.Quantity * ItemUOM;
                            BRTecess := recRG23D."eCess Amount Per Unit" * recRG23D.Quantity * ItemUOM;
                            BRTSheCess := recRG23D."SHE Cess Amount Per Unit" * recRG23D.Quantity * ItemUOM;
                            // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess ----------------------------------------------END
                        END; */
                    END;


                    CLEAR(TrfQty);
                    CLEAR(ExciseValue);
                    /* RG23Drec1.RESET;
                    RG23Drec1.SETRANGE(RG23Drec1."Transaction Type", RG23Drec1."Transaction Type"::Sale);
                    RG23Drec1.SETRANGE(RG23Drec1.Type, RG23Drec1.Type::Invoice);
                    RG23Drec1.SETRANGE(RG23Drec1."Document No.", "Sales Invoice Line"."Document No.");
                    RG23Drec1.SETRANGE(RG23Drec1."Item No.", "Sales Invoice Line"."No.");
                    RG23Drec1.SETRANGE(RG23Drec1."Line No.", "Sales Invoice Line"."Line No.");
                    RG23Drec1.SETRANGE(RG23Drec1."Location Code", "Sales Invoice Line"."Location Code");
                    RG23Drec1.SETFILTER(RG23Drec1.Quantity, '<>%1', 0);
                    RG23Drec1.SETRANGE(RG23Drec1."Ref. Entry No.", RG23Drec."Ref. Entry No.");
                    IF RG23Drec1.FINDFIRST THEN BEGIN
                        TrfQty := ABS(RG23Drec1.Quantity);
                        ExciseValue := ABS(RG23Drec1.Amount);
                        // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess ----------------------------------------------START
                        Exciseduty := ABS(RG23Drec1."BED Amount Per Unit" * RG23Drec1.Quantity * ItemUOM);
                        ecess := ABS(RG23Drec1."eCess Amount Per Unit" * RG23Drec1.Quantity * ItemUOM);
                        shecess := ABS(RG23Drec1."SHE Cess Amount Per Unit" * RG23Drec1.Quantity * ItemUOM);
                        // EBT MILAN 060214 To ADD Excies Duty , eCess SheCess ------------------------------------------------END
                    END; */

                    IF (LocCode = 'PLANT01') OR (LocCode = 'PLANT02') THEN
                        IssueBy := 'Manufacturer'
                    ELSE
                        IssueBy := 'Dealer';

                    CLEAR(ECCcode);
                    CLEAR(ECCDes);
                    CLEAR(LocName);
                    Location.RESET;
                    Location.SETRANGE(Location.Code, LocCode);
                    IF Location.FINDFIRST THEN BEGIN
                        ECCcode := '';//Location."E.C.C. No.";
                        LocName := Location.Name;

                        /* recECC.RESET;
                        recECC.SETRANGE(recECC.Code, ECCcode);
                        IF recECC.FINDFIRST THEN
                            ECCDes := recECC."C.E. Registration No."; */
                    END;

                    RecItem.RESET;
                    IF RecItem.GET("Sales Invoice Line"."No.") THEN BEGIN
                        UOM := RecItem."Base Unit of Measure";
                    END;

                    IF UOM = 'LTRS' THEN BEGIN
                        UOM := 'L'
                    END;

                    IF UOM = 'KGS' THEN BEGIN
                        UOM := 'KG'
                    END;


                    MakeExcelDataBody;
                    NewHeader := TRUE;

                    /*
                    ELSE
                    //MILAN
                     BEGIN
                       SalesCrMemo.RESET;
                       SalesCrMemo.SETRANGE(SalesCrMemo."Document No.",recRG23Dtable1."Document No.");
                       SalesCrMemo.SETRANGE(SalesCrMemo."No.",recRG23Dtable1."Item No.");
                       SalesCrMemo.SETRANGE(SalesCrMemo."Line No.",recRG23Dtable1."Line No.");
                       IF SalesCrMemo.FINDFIRST THEN
                       BEGIN
                        BRTNO := recRG23Dtable1."Document No.";
                        BRTExciseValue := SalesCrMemo."Excise Amount";
                       END;


                     END;
                       */
                    //
                END;
                //UNTIL RG23Drec.NEXT = 0;


                //>> 06Mar2017
                //IF BRTNO <> '' THEN
                //BEGIN
                //MakeExcelDataBody;//03Mar2017
                //NewHeader := TRUE;//03Mar2017
                //END;
                //<<
                /*
                //<<1

                 ExcelBuf.NewRow;
                 ExcelBuf.AddColumn("Sales Invoice Header"."No.",FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                 ExcelBuf.AddColumn('2',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                 ExcelBuf.AddColumn('3',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                 ExcelBuf.AddColumn('4',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                 ExcelBuf.AddColumn('5',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                 ExcelBuf.AddColumn('6',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                 ExcelBuf.AddColumn('7',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                 ExcelBuf.AddColumn('8',FALSE,'',FALSE,FALSE,FALSE,'',1);//2
                 ExcelBuf.AddColumn('*'+FORMAT("Sales Invoice Header"."Posting Date",0,'<Day,2>/<Month,2>/<Year4>'),FALSE,'',FALSE,FALSE,FALSE,'',1);//3
                 //ExcelBuf.AddColumn("Sales Invoice Line".Description,TRUE,'',FALSE,FALSE,FALSE,'',1);//4
                 //ExcelBuf.AddColumn("Sales Invoice Line"."Excise Prod. Posting Group",TRUE,'',FALSE,FALSE,FALSE,'',1);//5
                 //ExcelBuf.AddColumn(UOM,TRUE,'',FALSE,FALSE,FALSE,'@',1);//7

                //MESSAGE(' ...%1 \ .....%2 \ Line No. %3 \ Doc No. %4',NewHeader,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                */

                //end;



            }

            trigger OnPreDataItem()
            begin

                //>>1

                mindate := "Sales Invoice Header".GETRANGEMIN("Posting Date");
                maxdate := "Sales Invoice Header".GETRANGEMAX("Posting Date");
                //<<1
            end;
        }
        dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Item No.")
                                    WHERE(Quantity = FILTER(<> 0));
                dataitem("Transfer Receipt Line"; "Transfer Receipt Line")
                {
                    DataItemLink = "Transfer Order No." = FIELD("Transfer Order No."),
                                   "Item No." = FIELD("Item No."),
                                   "Line No." = FIELD("Line No.");
                    DataItemTableView = SORTING("Document No.", "Line No.");
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    CLEAR(TSHno1);
                    CLEAR(TSHdate1);
                    CLEAR(TROno1);
                    CLEAR(ILEEntryNo1);
                    CLEAR("ItemApplEntryNo.1");
                    CLEAR(No1);
                    ValueEntry1.RESET;
                    ValueEntry1.SETRANGE(ValueEntry1."Document No.", "Transfer Shipment Line"."Document No.");
                    ValueEntry1.SETRANGE(ValueEntry1."Item No.", "Transfer Shipment Line"."Item No.");
                    ValueEntry1.SETRANGE(ValueEntry1."Location Code", "Transfer Shipment Line"."Transfer-from Code");
                    ValueEntry1.SETRANGE(ValueEntry1."Document Line No.", "Transfer Shipment Line"."Line No.");
                    IF ValueEntry1.FINDFIRST THEN BEGIN
                        ILEEntryNo1 := ValueEntry1."Item Ledger Entry No.";

                        ItemApplEntry1.RESET;
                        ItemApplEntry1.SETRANGE(ItemApplEntry1."Item Ledger Entry No.", ILEEntryNo1);
                        IF ItemApplEntry1.FINDFIRST THEN
                            "ItemApplEntryNo.1" := ItemApplEntry1."Inbound Item Entry No.";

                        ILE1.RESET;
                        ILE1.SETRANGE(ILE1."Entry No.", "ItemApplEntryNo.1");
                        IF ILE1.FINDFIRST THEN BEGIN
                            //TROno1 := ILE1."Transfer Order No.";
                            TROno1 := ILE1."Order No.";//03Mar2017
                            TSHno1 := ILE1."Document No.";
                            TSHdate1 := ILE1."Posting Date";
                            No1 := ILE1."Document No.";
                            IF ILE1."Document Type" = ILE1."Document Type"::"Transfer Receipt" THEN BEGIN
                                recTSH1.RESET;
                                recTSH1.SETRANGE(recTSH1."Transfer Order No.", TROno1);
                                IF recTSH1.FINDFIRST THEN BEGIN
                                    TSHno1 := recTSH1."No.";
                                    TSHdate1 := recTSH1."Posting Date";
                                    //IPOL/LOCATION/001 >>>>
                                    LocName := recTSH1."Transfer-from Code";
                                    //IPOL/LOCATION/001 <<<<
                                END;
                            END;
                        END;
                    END;

                    TrfQty1 := 0;
                    ExciseValue1 := 0;
                    CLEAR(ItemNo1);
                    /* recRG23D1.RESET;
                    recRG23D1.SETRANGE(recRG23D1."Document No.", No1);
                    recRG23D1.SETRANGE(recRG23D1."Item No.", "Transfer Shipment Line"."Item No.");
                    recRG23D1.SETRANGE(recRG23D1."Location Code", "Transfer Shipment Line"."Transfer-from Code");
                    IF recRG23D1.FINDFIRST THEN BEGIN
                        ItemNo1 := recRG23D1."Item No.";
                        TrfQty1 += recRG23D1.Quantity;
                        // EBT MILAN 060214---------------------------------------------------------START
                        Exciseduty += (recRG23D1."BED Amount Per Unit" * recRG23D1.Quantity * ItemUOM);
                        ecess += (recRG23D1."eCess Amount Per Unit" * recRG23D1.Quantity * ItemUOM);
                        shecess += (recRG23D1."SHE Cess Amount Per Unit" * recRG23D1.Quantity * ItemUOM);
                        // EBT MILAN 060214-----------------------------------------------------------END
                        ExciseValue1 += recRG23D1.Amount;
                    END; */

                    Recitemtable1.RESET;
                    Recitemtable1.SETRANGE(Recitemtable1."No.", ItemNo1);
                    IF Recitemtable1.FINDFIRST THEN BEGIN
                        recItemUom1.RESET;
                        recItemUom1.SETRANGE(recItemUom1.Code, Recitemtable1."Sales Unit of Measure");
                        IF recItemUom1.FINDFIRST THEN
                            ItemUOM1 := recItemUom1."Qty. per Unit of Measure";
                    END;


                    CLEAR(IssueBy1);
                    IF (LocCode = 'PLANT01') OR (LocCode = 'PLANT02') THEN
                        IssueBy1 := 'Manufacturer'
                    ELSE
                        IssueBy1 := 'Dealer';

                    RecItem3.RESET;
                    IF RecItem3.GET(RecTransShipLine."Item No.") THEN;

                    MakeExcelDataBody4;


                    //<<1
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                ExcelBuf.NewRow;//03Mar2017

                //<<1
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Transfer Shipment Header".SETRANGE("Posting Date", mindate, maxdate);
                "Transfer Shipment Header".SETRANGE("Transfer Shipment Header"."Transfer-from Code", locfilt);
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

    trigger OnInitReport()
    begin
        //>>1

        CompInfo.GET;
        //<<1
    end;

    trigger OnPostReport()
    begin

        //>>1

        CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        ExcelBuf.RESET;//03Mar
        ExcelBuf.DELETEALL;//03Mar
                           //>>2

        /*
        //EBT STIVAN ---(22082012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",USERID);
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','Report View');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.RESET;
         recLocation.SETRANGE(recLocation.Code,"Sales Invoice Header".GETFILTER("Sales Invoice Header"."Location Code"));
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping1.RESET;
         CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
         CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
         CSOMapping1.SETRANGE(CSOMapping1.Value,"Sales Invoice Header".GETFILTER("Sales Invoice Header"."Location Code"));
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
        //EBT STIVAN ---(22082012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        *///Commented  03Mar2017

        MakeExcelInfo;
        locfilt := "Sales Invoice Header".GETFILTER("Sales Invoice Header"."Location Code");

        "SrNo." := 0;
        SrNo := 0;
        //<<2

    end;

    var
        ExcelBuf: Record 370 temporary;
        m: Integer;
        CompInfo: Record 79;
        "SrNo.": Integer;
        ItemREC: Record 27;
        BRTNoLen: BigInteger;
        BRTNumber: Code[30];
        Value: Decimal;
        Intermvalue: Decimal;
        RecItem2: Record Item;
        SrNo: Integer;
        RecItem: Record 27;
        Qty: Decimal;
        QtyPerUOM: Decimal;
        PackDescriptionforILE: Text[250];
        UOM: Code[20];
        SIL1: Record 113;
        NewHeader: Boolean;
        Location_REC: Record 14;
        IssueBy: Text[30];
        //"E.C.C.": Record 13708;
        TransferShipmentHeader: Record 5744;
        Transfershipmentline: Record 5745;
        TrfQty: Decimal;
        ExciseValue: Decimal;
        SalesInvoiceHeader: Record 112;
        salesinvoiceline: Record 113;
        itemLedgerEntry: Record 32;
        recile: Record 32;
        lineno: Integer;
        TransShipmentLine: Record 5745;
        BRTNO: Code[20];
        qnty: Decimal;
        maxdate: Date;
        mindate: Date;
        RecItem3: Record 27;
        locfilt: Code[30];
        RecLoc1: Record 14;
        //ECC: Record 13708;
        IssueBy1: Text[30];
        //RG23D: Record 16537;
        RecTransShipLine: Record 5745;
        //RG23Drec: Record 16537;
        BRNo: Code[20];
        BRDate: Date;
        recTSH: Record 5744;
        recTRH: Record 5746;
        TROno: Code[20];
        TSHno: Code[20];
        TSHnolen: BigInteger;
        TSHdate: Date;
        ItemApplEntry: Record 339;
        ValueEntry: Record 5802;
        ILEEntryNo: Integer;
        "ItemApplEntryNo.": Integer;
        ILE: Record 32;
        ILEDocType: Option;
        recTSL: Record 5745;
        //recRG23D: Record 16537;
        trfShpLine: Record 5745;
        LocCode: Code[20];
        Location: Record 14;
        ECCcode: Code[20];
        //recECC: Record 13708;
        ECCDes: Text[30];
        LocName: Text[30];
        ItemApplEntry1: Record 339;
        ValueEntry1: Record 5802;
        ILEEntryNo1: Integer;
        "ItemApplEntryNo.1": Integer;
        ILE1: Record 32;
        recTSH1: Record 5744;
        TROno1: Code[20];
        TSHno1: Code[20];
        TSHdate1: Date;
        TrfQty1: Decimal;
        ExciseValue1: Decimal;
        //recRG23D1: Record 16537;
        recTSL1: Record 5745;
        recItemUom: Record 5404;
        ItemUOM: Decimal;
        Recitemtable: Record 27;
        Recitemtable1: Record 27;
        recItemUom1: Record 5404;
        ItemUOM1: Decimal;
        No: Code[20];
        No1: Code[20];
        ItemNo: Code[20];
        ItemNo1: Code[20];
        LocResC: Code[10];
        recLocation: Record 14;
        CSOMapping1: Record 50006;
        ItemApplEntryqty: Decimal;
        recTShpLine: Record 5745;
        ExciseAmtperunit: Decimal;
        ILELineNo: Integer;
        BRTQty: Decimal;
        BRTExciseValue: Decimal;
        //recDetailedRG23D: Record 16533;
        recTrfSH: Record 5744;
        //recRG23Dtable1: Record 16537;
        TransferShipment: Record 5745;
        //RG23Drec1: Record 16537;
        SalesCrMemo: Record 115;
        BRTExciseDuty: Decimal;
        BRTecess: Decimal;
        BRTSheCess: Decimal;
        Exciseduty: Decimal;
        ecess: Decimal;
        shecess: Decimal;
        Text000: Label 'Data';
        Text001: Label 'Quarterly E-Return of Depot';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//03Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Quarterly E-Return of Depot', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Quarterly E-Return - Depot", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 1);
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

        //>>03Mar2017 RB-N
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
        //>>03Mar2017
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
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.NewRow;

        //<<


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Sr No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Des. of Goods', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('C.E. Tarrif Heading', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Qty. Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Base UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Qty.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        // EBT MILAN 060213
        ExcelBuf.AddColumn('Excise Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Edu.Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('S & H Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        // EBT MILAN 060213
        ExcelBuf.AddColumn('Amount of Duty Involved (Rs.)', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Sr No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Issue by', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Registration No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('Address', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Description of Goods', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('C.E. Tariff', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('Qty Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('Qty.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        // EBT MILAN 060213
        ExcelBuf.AddColumn('Excise Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Edu.Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('S & H Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        // EBT MILAN 060213
        ExcelBuf.AddColumn('Duty Involved', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        "SrNo." += 1;
        SrNo += 1;

        //MESSAGE(' ...%1 \ .....%2 \ Line No. %3 \ Doc No. %4',SrNo,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");

        recTrfSH.RESET;
        recTrfSH.SETRANGE(recTrfSH."No.", BRTNO);
        recTrfSH.SETRANGE(recTrfSH."BRT Cancelled", TRUE);
        IF recTrfSH.FINDFIRST THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn("SrNo.", FALSE, '', FALSE, FALSE, FALSE, '', 0);//1
            ExcelBuf.AddColumn("Sales Invoice Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
            ExcelBuf.AddColumn('*' + FORMAT("Sales Invoice Header"."Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'), FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
            ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
            ExcelBuf.AddColumn(/*"Sales Invoice Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
            ItemREC.RESET;
            IF ItemREC.GET("Sales Invoice Line"."No.") THEN;
            ExcelBuf.AddColumn("Sales Invoice Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
            ExcelBuf.AddColumn(UOM, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
                                                                            //MESSAGE('C7 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',UOM,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                            //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//7 //03Mar2017
                                                                            //IF RG23Drec.COUNT > 1 THEN BEGIN
            ExcelBuf.AddColumn(TrfQty * ItemUOM, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//8
                                                                                                  //MESSAGE('C8 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',TrfQty*ItemUOM,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                                                  //EBT MILAN 060214 ------------------------------------------------------------------START
            ExcelBuf.AddColumn(ROUND(Exciseduty, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//9
            ExcelBuf.AddColumn(ROUND(ecess, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//10
            ExcelBuf.AddColumn(ROUND(shecess, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//11
                                                                                                   //EBT MILAN 060214 ------------------------------------------------------------------START
            ExcelBuf.AddColumn(ROUND(ExciseValue, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//12
                                                                                                       //MESSAGE('C12 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',ExciseValue,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
        END ELSE BEGIN
            ExcelBuf.AddColumn(("Sales Invoice Line".Quantity *
            "Sales Invoice Line"."Qty. per Unit of Measure"), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//8
                                                                                                               //MESSAGE('C8 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',"Sales Invoice Line".Quantity*
                                                                                                               //"Sales Invoice Line"."Qty. per Unit of Measure","SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                                                               // EBT MILAN 060214-------------------------------------------------------------------------START
            ExcelBuf.AddColumn(ROUND(/*"Sales Invoice Line"."BED Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//9
            ExcelBuf.AddColumn(ROUND(/*"Sales Invoice Line"."eCess Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//10
            ExcelBuf.AddColumn(ROUND(/*"Sales Invoice Line"."SHE Cess Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//11
                                                                                                                                       // EBT MILAN 060214---------------------------------------------------------------------------END
            ExcelBuf.AddColumn(ROUND((/*"Sales Invoice Line"."BED Amount" + "Sales Invoice Line"."eCess Amount" +
                                    "Sales Invoice Line"."SHE Cess Amount"*/0), 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//12

            //MESSAGE('C12 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',"Sales Invoice Line"."BED Amount"+"Sales Invoice Line"."eCess Amount"+
            //                  "Sales Invoice Line"."SHE Cess Amount","SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
        END;

        ExcelBuf.AddColumn(SrNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn(IssueBy, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn(ECCDes, FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
                                                                                  //MESSAGE('C18 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',CompInfo.Name,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                                  //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//18 //03Mar2017
        ExcelBuf.AddColumn(LocName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
                                                                            //MESSAGE('C19 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',LocName,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                            //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//19
        ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        ItemREC.RESET;
        IF ItemREC.GET("Sales Invoice Line"."No.") THEN;
        ExcelBuf.AddColumn(ItemREC."Base Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//22
                                                                                                   //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//22
        ExcelBuf.AddColumn((BRTQty * ItemUOM), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//23
                                                                                                // EBT MILAN 060214-----------------------------------------------------------------------------START
        ExcelBuf.AddColumn((ROUND(BRTExciseDuty, 1)), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//24
        ExcelBuf.AddColumn((ROUND(BRTecess, 1)), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//25
        ExcelBuf.AddColumn((ROUND(BRTSheCess, 1)), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//26
                                                                                                    // EBT MILAN 060214-------------------------------------------------------------------------------END
        ExcelBuf.AddColumn(ROUND(BRTExciseValue, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//27

        // END ELSE BEGIN
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("SrNo.", FALSE, '', FALSE, FALSE, FALSE, '', 0);//1
        ExcelBuf.AddColumn("Sales Invoice Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('*' + FORMAT("Sales Invoice Header"."Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'), FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ItemREC.RESET;
        IF ItemREC.GET("Sales Invoice Line"."No.") THEN;
        ExcelBuf.AddColumn("Sales Invoice Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(UOM, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
                                                                        //MESSAGE('C7 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',UOM,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                        //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//7 //03Mar2017
                                                                        //IF RG23Drec.COUNT > 1 THEN BEGIN
        ExcelBuf.AddColumn(TrfQty * ItemUOM, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//8
                                                                                              //MESSAGE('C8 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',TrfQty*ItemUOM,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                                              // EBT MILAN 060214-----------------------------------------------------------------------------START
        ExcelBuf.AddColumn(ROUND(Exciseduty, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//9
        ExcelBuf.AddColumn(ROUND(ecess, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//10
        ExcelBuf.AddColumn(ROUND(shecess, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//11
                                                                                               // EBT MILAN 060214-------------------------------------------------------------------------------END
        ExcelBuf.AddColumn(ROUND(ExciseValue, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//12
                                                                                                   //END ELSE BEGIN
        ExcelBuf.AddColumn(("Sales Invoice Line".Quantity *
        "Sales Invoice Line"."Qty. per Unit of Measure"), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//8
                                                                                                           //MESSAGE('C8 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',"Sales Invoice Line".Quantity*
                                                                                                           //"Sales Invoice Line"."Qty. per Unit of Measure","SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                                                           // EBT MILAN 060214-----------------------------------------------------------------------------START
        ExcelBuf.AddColumn(ROUND(/*"Sales Invoice Line"."BED Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//9
        ExcelBuf.AddColumn(ROUND(/*"Sales Invoice Line"."eCess Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//10
        ExcelBuf.AddColumn(ROUND(/*"Sales Invoice Line"."SHE Cess Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//11
                                                                                                                                   // EBT MILAN 060214-------------------------------------------------------------------------------END
        ExcelBuf.AddColumn(ROUND((/*"Sales Invoice Line"."BED Amount" + "Sales Invoice Line"."eCess Amount" +
                                            "Sales Invoice Line"."SHE Cess Amount"*/0), 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//12
                                                                                                                                             //END;
        ExcelBuf.AddColumn(SrNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);//13
        ExcelBuf.AddColumn(BRTNO, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('*' + FORMAT(BRDate, 0, '<Day,2>/<Month,2>/<Year4>'), FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn(IssueBy, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn(ECCDes, FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
                                                                                  //MESSAGE('C18 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',CompInfo.Name,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                                  //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//18 //03Mar2017
        ExcelBuf.AddColumn(LocName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
                                                                            //MESSAGE('C19 %1 \ .....%2 \ Line No. %3 \ Doc No. %4',LocName,"SrNo.","Sales Invoice Line"."Line No.","Sales Invoice Header"."No.");
                                                                            //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//19
        ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        ItemREC.RESET;
        IF ItemREC.GET("Sales Invoice Line"."No.") THEN;

        IF ItemREC."Base Unit of Measure" = 'LTRS' THEN BEGIN
            ExcelBuf.AddColumn('L', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//22
                                                                            //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//22
        END;
        IF ItemREC."Base Unit of Measure" = 'KGS' THEN BEGIN
            ExcelBuf.AddColumn('KG', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//22
                                                                             //ExcelBuf.AddColumn('',TRUE,'',FALSE,FALSE,FALSE,'@',1);//22
        END;
        ExcelBuf.AddColumn((BRTQty), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//23
                                                                                      // EBT MILAN 060214------------------------------------------------------------------------START
        ExcelBuf.AddColumn((ROUND(BRTExciseDuty, 1)), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//24
        ExcelBuf.AddColumn((ROUND(BRTecess, 1)), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//25
        ExcelBuf.AddColumn((ROUND(BRTSheCess, 1)), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//26
                                                                                                    // EBT MILAN 060214--------------------------------------------------------------------------END
        ExcelBuf.AddColumn(ROUND(BRTExciseValue, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//27
    END;
    //end;

    // //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody2()
    begin
        IF NewHeader = FALSE THEN BEGIN
            ExcelBuf.NewRow;//03Mar2017
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//10

            ExcelBuf.AddColumn(SrNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);//11
            ExcelBuf.AddColumn(TSHno, FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
            ExcelBuf.AddColumn('*' + FORMAT(TSHdate, 0, '<Day,2>/<Month,2>/<Year4>'), FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
            ExcelBuf.AddColumn(IssueBy, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
            ExcelBuf.AddColumn(ECCDes, FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
            ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
            ExcelBuf.AddColumn(LocName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
            ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//18
            ExcelBuf.AddColumn(/*"Sales Invoice Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//19
            ItemREC.RESET;
            IF ItemREC.GET("Sales Invoice Line"."No.") THEN;

            IF ItemREC."Base Unit of Measure" = 'LTRS' THEN BEGIN
                ExcelBuf.AddColumn('L', FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
            END;
            IF ItemREC."Base Unit of Measure" = 'KGS' THEN BEGIN
                ExcelBuf.AddColumn('KG', FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
            END;

            ExcelBuf.AddColumn((TrfQty * ItemUOM), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//22
                                                                                                    // EBT MILAN 060214-----------------------------------------------------------------START
            ExcelBuf.AddColumn(ROUND(Exciseduty, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//23
            ExcelBuf.AddColumn(ROUND(ecess, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//24
            ExcelBuf.AddColumn(ROUND(shecess, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//25
                                                                                                   // EBT MILAN 060214-------------------------------------------------------------------END
            ExcelBuf.AddColumn(ROUND(ExciseValue, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//26
        END;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody3()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Transfer Shipment Header"."No.", TRUE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('*' + FORMAT("Transfer Shipment Header"."Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'),
        FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataBody4()
    begin

        "SrNo." := "SrNo." + 1;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("SrNo.", FALSE, '', TRUE, FALSE, FALSE, '', 0);//1

        //MESSAGE(' No.%1  &  SR %2',"Transfer Shipment Header"."No.","SrNo.");

        ExcelBuf.AddColumn("Transfer Shipment Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('*' + FORMAT("Transfer Shipment Header"."Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'),
        FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn("Transfer Shipment Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(/*"Transfer Shipment Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn("Transfer Shipment Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        IF RecItem3.GET("Transfer Shipment Line"."Item No.") THEN;

        IF RecItem3."Base Unit of Measure" = 'LTRS' THEN BEGIN
            ExcelBuf.AddColumn('L', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        END;
        IF RecItem3."Base Unit of Measure" = 'KGS' THEN BEGIN
            ExcelBuf.AddColumn('KG', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        END;

        ExcelBuf.AddColumn("Transfer Shipment Line".Quantity * "Transfer Shipment Line"."Qty. per Unit of Measure",
                           FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//8
        // EBT MILAN 060214--------------------------------------------------------------------------------------START
        ExcelBuf.AddColumn(ROUND(/*"Transfer Shipment Line"."BED Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//9
        ExcelBuf.AddColumn(ROUND(/*"Transfer Shipment Line"."eCess Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//10
        ExcelBuf.AddColumn(ROUND(/*"Transfer Shipment Line"."SHE Cess Amount"*/0, 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//11
        // EBT MILAN 060214----------------------------------------------------------------------------------------END
        ExcelBuf.AddColumn(ROUND((/*"Transfer Shipment Line"."BED Amount" +
                                 "Transfer Shipment Line"."eCess Amount" + "Transfer Shipment Line"."SHE Cess Amount"*/0), 1)
        , FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//12
        ExcelBuf.AddColumn("SrNo.", FALSE, '', FALSE, FALSE, FALSE, '', 0);//13
        ExcelBuf.AddColumn(TSHno1, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('*' + FORMAT(TSHdate1, 0, '<Day,2>/<Month,2>/<Year4>'), FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn(IssueBy1, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn(ECCDes, FALSE, '', FALSE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn(LocName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn(RecItem3.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn(/*"Transfer Shipment Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn(UOM, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn((TrfQty1 * ItemUOM1), FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//23
        // EBT MILAN 060214------------------------------------------------------------------------------------------START
        ExcelBuf.AddColumn(ROUND((Exciseduty), 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//24
        ExcelBuf.AddColumn(ROUND((ecess), 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//25
        ExcelBuf.AddColumn(ROUND((shecess), 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//26
        // EBT MILAN 060214--------------------------------------------------------------------------------------------END
        ExcelBuf.AddColumn(ROUND((ExciseValue1), 1), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//27
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody5()
    begin
        ExcelBuf.AddColumn("SrNo.", TRUE, '', FALSE, FALSE, FALSE, '', 0);//1
        ExcelBuf.AddColumn(TSHno, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('*' + FORMAT(TSHdate, 0, '<Day,2>/<Month,2>/<Year4>'), FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        IF RecLoc1.GET(RecTransShipLine."Transfer-from Code") THEN;
        //IF ECC.GET(RecLoc1."E.C.C. No.") THEN;
        ExcelBuf.AddColumn(IssueBy1, TRUE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ECCDes, TRUE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(LocName, TRUE, '', FALSE, FALSE, FALSE, '', 1);//7
        RecItem3.RESET;
        IF RecItem3.GET("Transfer Shipment Line"."Item No.") THEN;
        ExcelBuf.AddColumn(RecItem3.Description, TRUE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(/*RecTransShipLine."Excise Prod. Posting Group"*/0, TRUE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn(UOM, TRUE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn((RecTransShipLine.Quantity * "Sales Invoice Line"."Qty. per Unit of Measure"), TRUE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//11
        // EBT MILAN 060214 ------------------------------------------------------------------------------------------------START
        ExcelBuf.AddColumn(ROUND(/*RecTransShipLine."BED Amount"*/0, 1), TRUE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//12
        ExcelBuf.AddColumn(ROUND(/*RecTransShipLine."eCess Amount"*/0, 1), TRUE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//13
        ExcelBuf.AddColumn(ROUND(/*RecTransShipLine."SHE Cess Amount"*/0, 1), TRUE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//14
        // EBT MILAN 060214 --------------------------------------------------------------------------------------------------END
        ExcelBuf.AddColumn(ROUND((/*RecTransShipLine."BED Amount" + RecTransShipLine."eCess Amount" +
                                 RecTransShipLine."SHE Cess Amount"*/0), 1)
                                  , TRUE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//15
    end;
}


report 50099 "Ledger Entries - FG"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/LedgerEntriesFG.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.", "Location Filter", "Date Filter";
            column(ItmNo; "No.")
            {
            }
            column(ItmName; Item.Description)
            {
            }
            column(CompName; CompInfo.Name)
            {
            }
            column(DateFILTER; DateFILTER)
            {
            }
            column(Label; Label)
            {
            }
            dataitem("Item Ledger Entry"; "Item Ledger Entry")
            {
                DataItemLink = "Item No." = FIELD("No."),
                               "Posting Date" = FIELD("Date Filter"),
                               "Location Code" = FIELD("Location Filter");
                DataItemTableView = SORTING("Posting Date", "Item No.");
                column(SrNo; SrNo)
                {
                }
                column(PostingDate; "Item Ledger Entry"."Posting Date")
                {
                }
                column(OPBAL; OPBAL)
                {
                }
                column(QTYMAN; QTYMAN)
                {
                }
                column(Dispatch; QTYDES + QTYDES1)
                {
                }
                column(RejectionQty; RejectionQty)
                {
                }
                column(CLOSEBAL; CLOSEBAL)
                {
                }
                column(CLOSEBALInLtr; CLOSEBALInLtr)
                {
                }
                column(DocNo; DocNo)
                {
                }
                column(UOM; "Unit of Measure Code")
                {
                }
                column(BatchNo; "Lot No.")
                {
                }
                column(EntryNo; "Entry No.")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    QTYMAN := 0;
                    QTYDES := 0;
                    QTYDES1 := 0;
                    QTYREC := 0;
                    QTYMANInLtr := 0;
                    RejectionQty := 0;
                    CLOSEBAL := 0;
                    QTYRECInLtr := 0;
                    QTYDESInLtr := 0;
                    SrNo := SrNo + 1;

                    IF SrNo <> 1 THEN BEGIN
                        OPBAL := OpngBalforNextLine;
                    END;

                    ItemRec.GET("Item Ledger Entry"."Item No.");
                    SalesUOM := ItemRec."Sales Unit of Measure";

                    recItemUOM.RESET;
                    recItemUOM.SETRANGE(recItemUOM.Code, SalesUOM);
                    IF recItemUOM.FINDFIRST THEN
                        IF ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Output) OR
                           (("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer)
                                    AND ("Item Ledger Entry".Positive = TRUE))
                            OR ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Positive Adjmt.")
                        THEN BEGIN
                            QTYMAN := "Item Ledger Entry".Quantity / "Item Ledger Entry"."Qty. per Unit of Measure";
                            QTYMANInLtr := QTYMAN * recItemUOM."Qty. per Unit of Measure";
                        END;

                    IF ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Sale) AND
                         ("Item Ledger Entry".Positive = TRUE) THEN BEGIN
                        RejectionQty := ABS("Item Ledger Entry".Quantity / "Item Ledger Entry"."Qty. per Unit of Measure");
                        RejectionQtyInLtr := RejectionQty * recItemUOM."Qty. per Unit of Measure";
                    END;

                    IF ((("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Sale) OR
                        ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer) OR
                         ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Consumption) OR
                        ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Negative Adjmt.")) AND
                        ("Item Ledger Entry".Positive = FALSE)) THEN BEGIN
                        QTYDES := ABS("Item Ledger Entry".Quantity / "Item Ledger Entry"."Qty. per Unit of Measure");
                        QTYDESInLtr := QTYDES * recItemUOM."Qty. per Unit of Measure";
                    END;

                    IF (("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Sale) OR
                        ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer)) AND
                        ("Item Ledger Entry".Positive = TRUE) THEN BEGIN
                        QTYREC := "Item Ledger Entry".Quantity / "Item Ledger Entry"."Qty. per Unit of Measure";
                        QTYRECInLtr := QTYREC * recItemUOM."Qty. per Unit of Measure";
                    END;

                    CLOSEBAL := (((OPBAL + QTYMAN) - (QTYDES + QTYDES1)));
                    CLOSEBALInLtr := CLOSEBAL * recItemUOM."Qty. per Unit of Measure";
                    OPBALInLtr := OPBAL * recItemUOM."Qty. per Unit of Measure";

                    ValueEntry.RESET;
                    ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                    IF ValueEntry.FINDFIRST THEN BEGIN
                        DocumentNumber := ValueEntry."Document No.";
                    END;

                    InvoiceNo := '';
                    IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Sales Shipment" THEN BEGIN
                        recSSH.RESET;
                        recSSH.SETRANGE(recSSH."No.", DocumentNumber);
                        IF recSSH.FINDFIRST THEN BEGIN
                            SalesInvLine.SETRANGE(SalesInvLine."No.", "Item Ledger Entry"."Item No.");
                            SalesInvLine.SETRANGE(SalesInvLine."Lot No.", "Item Ledger Entry"."Lot No.");
                            SalesInvLine.SETRANGE(SalesInvLine."Location Code", "Item Ledger Entry"."Location Code");
                            SalesInvLine.SETRANGE(SalesInvLine."Line No.", "Item Ledger Entry"."Document Line No.");
                            IF SalesInvLine.FINDFIRST THEN BEGIN
                                InvoiceNo := SalesInvLine."Document No.";
                            END;
                        END;
                    END;

                    IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Sales Shipment" THEN BEGIN
                        DocNo := InvoiceNo;
                    END
                    ELSE BEGIN
                        DocNo := DocumentNumber;
                    END;




                    TotalOPBAL := TotalOPBAL + OPBAL;
                    TotalQTYMAN := TotalQTYMAN + QTYMAN;
                    TotalQTYREC := TotalQTYREC + QTYREC;
                    TotalQTYDES := TotalQTYDES + QTYDES + QTYDES1;
                    TotalCLOSEBAL := TotalCLOSEBAL + CLOSEBAL;
                    //<<1


                    //>>2

                    //Item Ledger Entry, Body (2) - OnPreSection()
                    SalesInvHdr.RESET;
                    SalesInvHdr.SETRANGE(SalesInvHdr."No.", DocumentNumber);
                    IF SalesInvHdr.FINDFIRST THEN BEGIN
                        PartyName := SalesInvHdr."Sell-to Customer Name";
                    END
                    ELSE BEGIN
                        TransferShptHdr.RESET;
                        TransferShptHdr.SETRANGE(TransferShptHdr."No.", DocumentNumber);
                        IF TransferShptHdr.FINDFIRST THEN BEGIN
                            PartyName := 'Sah Petroleums Ltd' + '-' + TransferShptHdr."Transfer-to Code";
                        END;
                    END;

                    //>>05May2017 ILEBody
                    IF PrinttoExcel THEN BEGIN

                        EnterCell(NNN, 1, FORMAT("Item Ledger Entry"."Posting Date"), FALSE, FALSE, '', 2);
                        EnterCell(NNN, 2, FORMAT(OPBAL), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 3, FORMAT(OPBALInLtr), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 4, FORMAT(QTYMAN), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 5, FORMAT(QTYMANInLtr), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 6, FORMAT(QTYREC), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 7, FORMAT(QTYRECInLtr), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 8, FORMAT(QTYDES), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 9, FORMAT(QTYDESInLtr), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 10, FORMAT(CLOSEBAL), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 11, FORMAT(CLOSEBALInLtr), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 12, DocumentNumber, FALSE, FALSE, '', 1);
                        EnterCell(NNN, 13, PartyName, FALSE, FALSE, '', 1);
                        EnterCell(NNN, 14, "Item Ledger Entry"."Unit of Measure Code", FALSE, FALSE, '', 1);
                        EnterCell(NNN, 15, FORMAT("Item Ledger Entry"."Entry No."), FALSE, FALSE, '', 0);
                        EnterCell(NNN, 16, FORMAT("Item Ledger Entry"."Entry Type"), FALSE, FALSE, '', 1);
                        EnterCell(NNN, 17, "Item Ledger Entry"."Lot No.", FALSE, FALSE, '', 1);
                        //rspl119-start
                        VE.RESET;
                        VE.SETRANGE(VE."Document Type", VE."Document Type"::"Sales Invoice");
                        VE.SETRANGE(VE."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                        IF VE.FIND('-') THEN
                            InvoiceNo := VE."Document No.";
                        //rspl119-end
                        EnterCell(NNN, 18, InvoiceNo, FALSE, FALSE, '', 1);

                        NNN += 1;

                    END;

                    //<<05May2017 ILEBody

                    /*
                    IF PrinttoExcel THEN
                    BEGIN
                      Date1:= FORMAT("Item Ledger Entry"."Posting Date") ;
                    
                            XLSHEET.Range('A'+FORMAT(i)).Value :=  Date1 ;
                            XLSHEET.Range('A'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('A'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('B'+FORMAT(i)).Value := OPBAL ;
                            XLSHEET.Range('B'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('B'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('B'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('C'+FORMAT(i)).Value := OPBALInLtr;
                            XLSHEET.Range('C'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('C'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('C'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('C'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('D'+FORMAT(i)).Value := QTYMAN;
                            XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('D'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('D'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('E'+FORMAT(i)).Value := QTYMANInLtr;
                            XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('E'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('F'+FORMAT(i)).Value := QTYREC;
                            XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('F'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('G'+FORMAT(i)).Value := QTYRECInLtr;
                            XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('G'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('H'+FORMAT(i)).Value :=QTYDES ;
                            XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('H'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('I'+FORMAT(i)).Value :=QTYDESInLtr;
                            XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('I'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('J'+FORMAT(i)).Value :=CLOSEBAL;
                            XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('K'+FORMAT(i)).Value :=CLOSEBALInLtr;
                            XLSHEET.Range('K'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('K'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('K'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('L'+FORMAT(i)).Value :=DocumentNumber;
                            XLSHEET.Range('L'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('L'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('M'+FORMAT(i)).Value :=PartyName  ;
                            XLSHEET.Range('M'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('M'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('N'+FORMAT(i)).Value :="Item Ledger Entry"."Unit of Measure Code";
                            XLSHEET.Range('N'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('N'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('O'+FORMAT(i)).Value :="Item Ledger Entry"."Entry No.";
                            XLSHEET.Range('O'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('O'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('O'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('P'+FORMAT(i)).Value := FORMAT("Item Ledger Entry"."Entry Type");
                            XLSHEET.Range('P'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('P'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('P'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('Q'+FORMAT(i)).Value :="Item Ledger Entry"."Lot No.";
                            XLSHEET.Range('Q'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('Q'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            //rspl119-start
                            VE.RESET;
                            VE.SETRANGE(VE."Document Type",VE."Document Type"::"Sales Invoice");
                            VE.SETRANGE(VE."Item Ledger Entry No.","Item Ledger Entry"."Entry No.");
                            IF VE.FIND('-') THEN
                            InvoiceNo:= VE."Document No.";
                            //rspl119-end
                    
                            XLSHEET.Range('R'+FORMAT(i)).Value := InvoiceNo;
                            XLSHEET.Range('R'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('R'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('R'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('R'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.Color :=65535;
                            XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 5;
                    
                       i += 1;
                    END;
                    *///Commented 05May2017
                      //<<2


                    //>>3
                    //Item Ledger Entry, Body (2) - OnPostSection()
                    OpngBalforNextLine := CLOSEBAL;
                    //<<3

                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                OPBAL := 0;
                QTYMAN := 0;
                QTYDES := 0;
                QTYREC := 0;
                CLOSEBAL := 0;

                //RSPL-CAS-06852-W6W8P4-START
                vQtySalesUom := 0;
                openbalc := 0;
                OPBALLtr := 0;

                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Location Code", LocFilter);
                RecILE.SETFILTER(RecILE."Posting Date", '%1|<%2', 0D, PosDate);
                IF RecILE.FINDSET THEN
                    REPEAT
                        openbalc += RecILE.Quantity;
                    UNTIL RecILE.NEXT = 0;

                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Location Code", LocFilter);
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Output);
                RecILE.SETFILTER(RecILE."Posting Date", '%1|<%2', 0D, PosDate);
                IF RecILE.FINDLAST THEN BEGIN
                    QTYMAN_v := RecILE.Quantity / RecILE."Qty. per Unit of Measure";
                    DocNo_v := RecILE."Document No.";
                    vBatchno := RecILE."Lot No.";
                END;

                RecItem.RESET;
                RecItem.SETRANGE("No.", Item."No.");
                IF RecItem.FINDFIRST THEN BEGIN
                    ItemUOM.RESET;
                    ItemUOM.SETRANGE(ItemUOM."Item No.", RecItem."No.");
                    ItemUOM.SETRANGE(ItemUOM.Code, RecItem."Sales Unit of Measure");
                    IF ItemUOM.FINDSET THEN BEGIN
                        vQtySalesUom := OPBALInLtr / ItemUOM."Qty. per Unit of Measure";
                        OPBALLtr := vQtySalesUom * ItemUOM."Qty. per Unit of Measure";
                    END;
                END;
                Date2 := FORMAT(PosDate);
                //RSPL-CAS-06852-W6W8P4-END

                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);

                asondate := ItemMaster.GETRANGEMAX(ItemMaster."Date Filter");
                PosDate := ItemMaster.GETRANGEMIN(ItemMaster."Date Filter");

                ItemMaster.SETFILTER("Date Filter", '<%1', PosDate);
                ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);

                IF ItemMaster.FINDSET THEN BEGIN
                    OPBALInLtr := 0;
                    ItemUOM.RESET;
                    ItemUOM.SETRANGE(ItemUOM."Item No.", Item."No.");
                    ItemUOM.SETRANGE(ItemUOM.Code, Item."Sales Unit of Measure");
                    ItemMaster.CALCFIELDS("Inventory Change");
                    IF ItemUOM.FINDFIRST THEN
                        OPBAL := ItemMaster."Inventory Change" / ItemUOM."Qty. per Unit of Measure";
                    //OPBALInLtr :=  OPBAL *  ItemUOM."Qty. per Unit of Measure";
                END;
                //<<1

                //>>2
                //Item, Header (1) - OnPreSection()
                CompInfo.GET;

                Location.SETRANGE(Location.Code, Item."Location Filter");
                IF Location.FINDFIRST THEN
                    Loc := Location.Name;

                IF Location."Trading Location" = TRUE THEN
                    Label := 'Receipt'
                ELSE
                    Label := 'Production';
                //<<2


                //>>3

                //>>ItemDataBody 05May2017
                IF PrinttoExcel THEN BEGIN

                    //rspl119-start
                    VE.RESET;
                    VE.SETRANGE(VE."Document Type", VE."Document Type"::"Sales Invoice");
                    VE.SETRANGE(VE."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                    IF VE.FIND('-') THEN
                        InvoiceNo := VE."Document No.";
                    //rspl119-end

                    EnterCell(NNN, 1, Date2, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 2, FORMAT(vQtySalesUom), FALSE, FALSE, '0.000', 0);
                    EnterCell(NNN, 3, FORMAT(OPBALLtr), FALSE, FALSE, '0.000', 0);
                    EnterCell(NNN, 4, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 5, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 6, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 7, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 8, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 9, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 10, FORMAT(vQtySalesUom), FALSE, FALSE, '0.000', 0);
                    EnterCell(NNN, 11, FORMAT(openbalc), FALSE, FALSE, '0.000', 0);
                    EnterCell(NNN, 12, DocNo_v, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 13, PartyName_v, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 14, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 15, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 16, '-', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 17, vBatchno, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 18, '-', FALSE, FALSE, '', 1);

                    NNN += 1;
                END;

                //<<ItemDataBody 05May2017
                //Item, Body (2) - OnPreSection()
                /*
                //RSPL-CAS-06852-W6W8P4-START
                IF PrinttoExcel THEN
                BEGIN
                
                        XLSHEET.Range('A'+FORMAT(i)).Value := Date2;
                        XLSHEET.Range('A'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('A'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('B'+FORMAT(i)).Value := vQtySalesUom;
                        XLSHEET.Range('B'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('B'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('B'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('C'+FORMAT(i)).Value := OPBALLtr;
                        XLSHEET.Range('C'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('C'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('C'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('C'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('D'+FORMAT(i)).Value := '-';
                        XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('D'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('D'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('E'+FORMAT(i)).Value :='-' ;
                        XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('E'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('F'+FORMAT(i)).Value := '-';
                        XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('F'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('G'+FORMAT(i)).Value := '-';
                        XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('G'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('H'+FORMAT(i)).Value :='-' ;
                        XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('H'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('I'+FORMAT(i)).Value :='-';
                        XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('I'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('J'+FORMAT(i)).Value :=vQtySalesUom;
                        XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('K'+FORMAT(i)).Value :=openbalc;
                        XLSHEET.Range('K'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('K'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('K'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('L'+FORMAT(i)).Value :=DocNo_v;
                        XLSHEET.Range('L'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('L'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('M'+FORMAT(i)).Value :=PartyName_v;
                        XLSHEET.Range('M'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('M'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('N'+FORMAT(i)).Value :='-';
                        XLSHEET.Range('N'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('N'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('O'+FORMAT(i)).Value :='-';
                        XLSHEET.Range('O'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('O'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('O'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('P'+FORMAT(i)).Value := '-';
                        XLSHEET.Range('P'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('P'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('P'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('Q'+FORMAT(i)).Value := vBatchno;
                        XLSHEET.Range('Q'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('Q'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                
                        //rspl119-start
                        VE.RESET;
                        VE.SETRANGE(VE."Document Type",VE."Document Type"::"Sales Invoice");
                        VE.SETRANGE(VE."Item Ledger Entry No.","Item Ledger Entry"."Entry No.");
                        IF VE.FIND('-') THEN
                        InvoiceNo:= VE."Document No.";
                        //rspl119-end
                
                        XLSHEET.Range('R'+FORMAT(i)).Value := '-';
                        XLSHEET.Range('R'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('R'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('R'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('R'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.Color :=65535;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 5;
                
                  i+= 1;
                END;
                *///Commented 05May2017
                  //RSPL-CAS-06852-W6W8P4-END
                  //<<3

            end;

            trigger OnPostDataItem()
            begin

                //>>1
                //Item, Footer (4) - OnPreSection()
                ItemRec.GET(Item."No.");

                //>>05May2017 ReportFooter
                IF PrinttoExcel THEN BEGIN
                    EnterCell(NNN, 1, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 2, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 3, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 4, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 5, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 6, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 7, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 8, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 9, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 10, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 11, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 12, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 13, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 14, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 15, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 16, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 17, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 18, '', TRUE, TRUE, '', 1);

                    NNN += 1;

                    EnterCell(NNN, 1, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 2, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 3, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 4, FORMAT(TotalQTYMAN), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 5, FORMAT(TotalQTYMANLtr), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 6, FORMAT(TotalQTYREC), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 7, FORMAT(TotalQTYRECLtr), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 8, FORMAT(TotalQTYDES), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 9, FORMAT(TotalQTYDESLtr), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 10, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 11, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 12, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 13, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 14, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 15, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 16, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 17, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 18, '', TRUE, TRUE, '', 1);

                    NNN += 1;

                END;

                //<<05May2017 ReportFooter

                /*
                IF PrinttoExcel THEN
                BEGIN
                        XLSHEET.Range('D'+FORMAT(i)).Value := TotalQTYMAN;
                        XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('D'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('D'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('D1',MaxCol+'1').Interior.ColorIndex := 8;
                        XLSHEET.Range('D1',MaxCol+'1').Borders.ColorIndex := 5;
                
                        XLSHEET.Range('E'+FORMAT(i)).Value := TotalQTYMANLtr;
                        XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('E'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('E1',MaxCol+'1').Interior.ColorIndex := 8;
                        XLSHEET.Range('E1',MaxCol+'1').Borders.ColorIndex := 5;
                
                        XLSHEET.Range('F'+FORMAT(i)).Value := TotalQTYREC;
                        XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('F'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('G'+FORMAT(i)).Value := TotalQTYRECLtr;
                        XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('G'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('H'+FORMAT(i)).Value :=TotalQTYDES ;
                        XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('H'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('I'+FORMAT(i)).Value :=TotalQTYDESLtr ;
                        XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('I'+FORMAT(i)).HorizontalAlignment :=-4152;
                END;
                *///Commented 05May2017
                  //<<1

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                CompInfo.GET;
                /*
                IF PrinttoExcel THEN
                BEGIN
                  CreateXLSHEET;
                  CreateHeader;
                  i := 7;
                END;
                *///Commented 05May2017
                  //<<1

                //>>05May2017 ReportHeader
                IF PrinttoExcel THEN BEGIN
                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '', 1);
                    EnterCell(1, 18, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

                    EnterCell(2, 1, 'Ledger Entries - FG', TRUE, FALSE, '', 1);
                    EnterCell(2, 18, USERID, TRUE, FALSE, '', 1);


                    EnterCell(4, 1, 'Item Code', TRUE, FALSE, '', 1);
                    EnterCell(4, 2, Item.GETFILTER("No."), TRUE, FALSE, '', 1);

                    EnterCell(5, 1, 'Item Description', TRUE, FALSE, '', 1);
                    EnterCell(5, 2, ItemFilter, TRUE, FALSE, '', 1);

                    EnterCell(6, 1, 'Location', TRUE, TRUE, '', 1);
                    EnterCell(6, 2, filtertext, TRUE, TRUE, '', 1);
                    EnterCell(6, 3, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 4, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 5, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 6, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 7, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 8, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 9, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 10, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 11, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 12, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 13, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 14, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 15, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 16, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 17, '', TRUE, TRUE, '', 1);
                    EnterCell(6, 18, '', TRUE, TRUE, '', 1);


                    EnterCell(7, 1, 'Date', TRUE, TRUE, '', 1);
                    EnterCell(7, 2, 'Opening Balance in Sales UOM', TRUE, TRUE, '', 1);
                    EnterCell(7, 3, 'Opening Balance in LTR/KG', TRUE, TRUE, '', 1);
                    EnterCell(7, 4, 'Production in Sales UOM', TRUE, TRUE, '', 1);
                    EnterCell(7, 5, 'Production in LTR/KG', TRUE, TRUE, '', 1);
                    EnterCell(7, 6, 'Rejection in Sales UOM', TRUE, TRUE, '', 1);
                    EnterCell(7, 7, 'Rejection in LTR/KG', TRUE, TRUE, '', 1);
                    EnterCell(7, 8, 'Despatch in Sales UOM', TRUE, TRUE, '', 1);
                    EnterCell(7, 9, 'Despatch in LTR/KG', TRUE, TRUE, '', 1);
                    EnterCell(7, 10, 'Closing Balance in Sales UOM', TRUE, TRUE, '', 1);
                    EnterCell(7, 11, 'Closing Balance in LTR/KG', TRUE, TRUE, '', 1);
                    EnterCell(7, 12, 'Document No.', TRUE, TRUE, '', 1);
                    EnterCell(7, 13, 'Party Name', TRUE, TRUE, '', 1);
                    EnterCell(7, 14, 'Sales UOM', TRUE, TRUE, '', 1);
                    EnterCell(7, 15, 'Entry No', TRUE, TRUE, '', 1);
                    EnterCell(7, 16, 'Entry Type', TRUE, TRUE, '', 1);
                    EnterCell(7, 17, 'Batch No.', TRUE, TRUE, '', 1);
                    EnterCell(7, 18, 'Invoice No.', TRUE, TRUE, '', 1);

                    NNN := 8;
                END;
                //<<05May2017 ReportHeader

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print to Excel"; PrinttoExcel)
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

        //>>05May2017
        IF PrinttoExcel THEN BEGIN

            /* ExcBuffer.CreateBook('', 'Ledeger Entries FG');
            ExcBuffer.CreateBookAndOpenExcel('', 'Ledeger Entries FG', '', '', USERID);
            ExcBuffer.GiveUserControl; */

            ExcBuffer.CreateNewBook('Ledeger Entries FG');
            ExcBuffer.WriteSheet('Ledeger Entries FG', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Ledeger Entries FG', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();
        END;
        //<<05May2017
    end;

    trigger OnPreReport()
    begin

        //>>1

        IF Item.GETFILTER(Item."Location Filter") = '' THEN
            ERROR('Please Select Location Code');


        DateFILTER := Item.GETFILTER("Date Filter");
        LocFilter := Item.GETFILTER(Item."Location Filter");
        ItemFilter := Item.GETFILTER(Item."No.");
        ItemMaster.GET(ItemFilter);
        ItemFilter := ItemMaster.Description;
        vitemno := ItemMaster."No.";
        Location.SETRANGE(Location.Code, LocFilter);
        IF Location.FINDFIRST THEN;
        Loc := COPYSTR(Location.Name, 1, 50);
        filtertext := Loc;
        //<<1
    end;

    var
        OPBAL: Decimal;
        QTYMAN: Decimal;
        QTYREC: Decimal;
        QTYDES: Decimal;
        QTYDES1: Decimal;
        CLOSEBAL: Decimal;
        ItemMaster: Record 27;
        DateFILTER: Text[30];
        LocFilter: Text[30];
        OpenFilter: Text[30];
        SrNo: Integer;
        PosDate: Date;
        CompInfo: Record 79;
        "<For Excel Generation>": Integer;
        MaxCol: Text[3];
        i: Integer;
        "Grand Total": Text[30];
        respcentname: Text[30];
        filtertext: Text[100];
        datefiltertext: Text[250];
        TotalOPBAL: Decimal;
        TotalQTYMAN: Decimal;
        TotalQTYREC: Decimal;
        TotalQTYDES: Decimal;
        TotalCLOSEBAL: Decimal;
        Location: Record 14;
        Loc: Text[30];
        Usersetup: Record 91;
        CompName: Text[100];
        OpngBalforNextLine: Decimal;
        DocumentNumber: Text[100];
        ValueEntry: Record 5802;
        SalesInvHdr: Record 112;
        PartyName: Text[100];
        TransferShptHdr: Record 5744;
        ItemFilter: Text[100];
        Date1: Text[30];
        PrinttoExcel: Boolean;
        SalesInvLine: Record 113;
        InvoiceNo: Code[20];
        RejectionQty: Decimal;
        FromDate: Date;
        ToDate: Date;
        OPBALInLtr: Decimal;
        QTYMANInLtr: Decimal;
        QTYDESInLtr: Decimal;
        QTYDES1InLtr: Decimal;
        RejectionQtyInLtr: Decimal;
        CLOSEBALInLtr: Decimal;
        ItemRec: Record 27;
        QTYRECInLtr: Decimal;
        TotalQTYMANLtr: Decimal;
        TotalQTYRECLtr: Decimal;
        TotalQTYDESLtr: Decimal;
        Label: Text[50];
        SalesUOM: Code[10];
        recItemUOM: Record 5404;
        UOM: Code[10];
        ItemUOM: Record 5404;
        recSSH: Record 110;
        recSIH: Record 112;
        DocNo: Code[20];
        VE: Record 5802;
        asondate: Date;
        RecILE: Record 32;
        openbalc: Decimal;
        vQtySalesUom: Decimal;
        Date2: Text[30];
        QTYMAN_v: Decimal;
        QTYMANInLtr_v: Decimal;
        DocNo_v: Text[30];
        PartyName_v: Text[30];
        QTYMAN_o: Decimal;
        QTYMANInLtr_o: Decimal;
        vBatchno: Code[10];
        RecItem: Record 27;
        OPBALLtr: Decimal;
        vitemno: Text[30];
        "---05May2017---": Integer;
        ExcBuffer: Record 370 temporary;
        NNN: Integer;

    // [Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        IF PrinttoExcel THEN BEGIN
            ExcBuffer.INIT;
            ExcBuffer.VALIDATE("Row No.", Rowno);
            ExcBuffer.VALIDATE("Column No.", columnno);
            ExcBuffer."Cell Value as Text" := Cellvalue;
            ExcBuffer.Formula := '';
            ExcBuffer.Bold := Bold;
            ExcBuffer.Underline := Underline;
            ExcBuffer.NumberFormat := NoFormat;
            ExcBuffer."Cell Type" := CType;
            ExcBuffer.INSERT;
        END;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


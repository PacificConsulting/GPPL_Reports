report 50098 "Ledger Entries - RM"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/LedgerEntriesRM.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("Item Category Code", "No.")
                                ORDER(Ascending);
            RequestFilterFields = "No.", "Date Filter", "Location Filter", "Bin Filter";
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
            dataitem("Item Ledger Entry"; "Item Ledger Entry")
            {
                DataItemLink = "Item No." = FIELD("No."),
                               "Posting Date" = FIELD("Date Filter"),
                               "Location Code" = FIELD("Location Filter");
                DataItemTableView = SORTING("Location Code", "Posting Date", "Item No.", "Entry Type")
                                    ORDER(Ascending)
                                    WHERE("Location Code" = FILTER(<> 'IN-TRANS'));
                column(SrNo; SrNo)
                {
                }
                column(PostingDate; "Item Ledger Entry"."Posting Date")
                {
                }
                column(OPBAL; OPBAL)
                {
                }
                column(Receipt; QTYREC + QTYREC1)
                {
                }
                column(Rejected; QTYREJECTED)
                {
                }
                column(Issued; QTYISSUED)
                {
                }
                column(CLOSEBAL; CLOSEBAL)
                {
                }
                column(DocNo; DocumentNumber)
                {
                }
                column(EntryNo; "Entry No.")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    QTYREJECTED := 0;
                    QTYISSUED := 0;
                    QTYREC := 0;

                    SrNo := SrNo + 1;

                    IF SrNo <> 1 THEN BEGIN
                        OPBAL := OpngBalforNextLine;
                    END;

                    IF ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Purchase) AND
                       ("Item Ledger Entry".Positive = FALSE) THEN BEGIN
                        QTYREJECTED := ABS("Item Ledger Entry".Quantity);
                    END;


                    IF ((("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Positive Adjmt.") OR
                         ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Purchase) OR
                         ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer) OR
                         ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Sale)) AND
                         ("Item Ledger Entry".Positive = TRUE)) OR
                         ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Consumption) AND
                         ("Item Ledger Entry".Positive = TRUE) THEN BEGIN
                        QTYREC := "Item Ledger Entry".Quantity;
                        RecLocation := TransferRcptHdr."Transfer-from Code";
                    END;

                    IF ((("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer) OR
                        ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Consumption) OR
                        ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::"Negative Adjmt.") OR
                        ("Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Sale)) AND
                        ("Item Ledger Entry".Positive = FALSE)) THEN BEGIN
                        QTYISSUED := ABS("Item Ledger Entry".Quantity);
                        PartyName := TransferShptHdr."Transfer-to Code";
                    END;


                    CLOSEBAL := (OPBAL + QTYREC + QTYREC1) - (QTYISSUED + QTYREJECTED);


                    ValueEntry.RESET;
                    ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                    IF ValueEntry.FINDFIRST THEN
                        DocumentNumber := ValueEntry."Document No.";


                    TotalOPBAL := TotalOPBAL + OPBAL;
                    TotalQTYREJECTED := TotalQTYREJECTED + QTYREJECTED;
                    TotalQTYREC := TotalQTYREC + QTYREC;
                    TotalQTYISSUED := TotalQTYISSUED + QTYISSUED;
                    TotalCLOSEBAL := TotalCLOSEBAL + CLOSEBAL;
                    //<<1

                    //>>2

                    //Item Ledger Entry, Body (1) - OnPreSection()
                    PartyName := '';
                    TransferShptHdr.RESET;
                    TransferShptHdr.SETRANGE(TransferShptHdr."No.", DocumentNumber);
                    IF TransferShptHdr.FINDFIRST THEN BEGIN
                        IF QTYISSUED <> 0 THEN
                            PartyName := 'Sah Petroleums Ltd' + '-' + TransferShptHdr."Transfer-to Code";
                    END
                    ELSE
                        PartyName := '';

                    RecLocation := '';
                    TransferRcptHdr.RESET;
                    TransferRcptHdr.SETRANGE(TransferRcptHdr."No.", DocumentNumber);
                    IF TransferRcptHdr.FINDFIRST THEN BEGIN
                        IF QTYREC <> 0 THEN
                            RecLocation := 'Sah Petroleums Ltd' + '-' + TransferRcptHdr."Transfer-to Code";
                    END ELSE
                        RecLocation := '';

                    /*
                    IF PrinttoExcel THEN
                    BEGIN
                            XLSHEET.Range('A'+FORMAT(i)).Value :=  "Item Ledger Entry"."Posting Date" ;
                            XLSHEET.Range('A'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('A'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('A'+FORMAT(i)).NumberFormat := 'dd/mm/yyyy';
                            XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
                    
                    
                            XLSHEET.Range('B'+FORMAT(i)).Value := OPBAL ;
                            XLSHEET.Range('B'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('B'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('B'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('C'+FORMAT(i)).Value := QTYREC;
                            XLSHEET.Range('C'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('C'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('C'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('C'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('D'+FORMAT(i)).Value := QTYREJECTED;
                            XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('D'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('D'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('E'+FORMAT(i)).Value := QTYISSUED ;
                            XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('E'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('F'+FORMAT(i)).Value :=CLOSEBAL;
                            XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('F'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                    
                            XLSHEET.Range('G'+FORMAT(i)).Value := DocumentNumber;
                            XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                    
                    
                            XLSHEET.Range('H'+FORMAT(i)).Value := PartyName  ;
                            XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('I'+FORMAT(i)).Value := RecLocation  ;
                            XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.Color :=65535;
                            XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 5;
                       i += 1;
                    END;
                    
                    *///Commented 04May2017


                    //>>ReportBody 05May2017
                    IF PrinttoExcel THEN BEGIN

                        EnterCell(NNN, 1, FORMAT("Item Ledger Entry"."Posting Date"), FALSE, FALSE, '', 2);
                        EnterCell(NNN, 2, FORMAT(OPBAL), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 3, FORMAT(QTYREC), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 4, FORMAT(QTYREJECTED), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 5, FORMAT(QTYISSUED), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 6, FORMAT(CLOSEBAL), FALSE, FALSE, '0.000', 0);
                        EnterCell(NNN, 7, DocumentNumber, FALSE, FALSE, '', 1);
                        EnterCell(NNN, 8, PartyName, FALSE, FALSE, '', 1);
                        EnterCell(NNN, 9, RecLocation, FALSE, FALSE, '', 1);

                        NNN += 1;
                    END;
                    //<<ReportBody 05May2017
                    //>>3

                    //Item Ledger Entry, Body (1) - OnPostSection()
                    OpngBalforNextLine := CLOSEBAL;
                    //<<3

                    //<<2

                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                OPBAL := 0;
                QTYREJECTED := 0;
                QTYISSUED := 0;
                QTYREC := 0;
                CLOSEBAL := 0;


                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                PosDate := ItemMaster.GETRANGEMIN(ItemMaster."Date Filter");
                ItemMaster.SETFILTER("Date Filter", '<%1', PosDate);
                ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER(ItemMaster."Bin Filter", BinFilter);
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    OPBAL := ItemMaster."Inventory Change";
                END;
                //<<1
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                //Item, Footer (6) - OnPreSection()
                /*
                IF PrinttoExcel THEN
                BEGIN
                        XLSHEET.Range('C'+FORMAT(i)).Value := TotalQTYREJECTED;
                        XLSHEET.Range('C'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('C'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('C'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('C'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('D'+FORMAT(i)).Value := TotalQTYREC;
                        XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('D'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('D'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('E'+FORMAT(i)).Value :=TotalQTYISSUED ;
                        XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('E'+FORMAT(i)).HorizontalAlignment :=-4152;
                 END;
                 *///Commented 04May2017
                   //<<1

                //>>ReportFooter 05May2017
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
                    NNN += 1;

                    EnterCell(NNN, 1, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 2, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 3, FORMAT(TotalQTYREC), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 4, FORMAT(TotalQTYREJECTED), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 5, FORMAT(TotalQTYISSUED), TRUE, TRUE, '0.000', 0);
                    EnterCell(NNN, 6, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 7, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 8, '', TRUE, TRUE, '', 1);
                    EnterCell(NNN, 9, '', TRUE, TRUE, '', 1);
                    NNN += 1;

                END;
                //<<ReportFooter 05May2017

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
                    EnterCell(1, 9, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

                    EnterCell(2, 1, 'Ledger Entries - RM', TRUE, FALSE, '', 1);
                    EnterCell(2, 9, USERID, TRUE, FALSE, '', 1);


                    EnterCell(5, 1, 'Date Filter : ' + DateFILTER, TRUE, FALSE, '', 1);
                    IF ItemFilter <> '' THEN
                        EnterCell(6, 1, 'Item Filter : ' + ItemFilter, TRUE, FALSE, '', 1);


                    EnterCell(7, 1, 'Date', TRUE, TRUE, '', 1);
                    EnterCell(7, 2, 'Opening Balance', TRUE, TRUE, '', 1);
                    EnterCell(7, 3, 'Receipt', TRUE, TRUE, '', 1);
                    EnterCell(7, 4, 'Rejection', TRUE, TRUE, '', 1);
                    EnterCell(7, 5, 'Issued', TRUE, TRUE, '', 1);
                    EnterCell(7, 6, 'Closing Balance', TRUE, TRUE, '', 1);
                    EnterCell(7, 7, 'Document No.', TRUE, TRUE, '', 1);
                    EnterCell(7, 8, 'Issued To Location', TRUE, TRUE, '', 1);
                    EnterCell(7, 9, 'Received From Location', TRUE, TRUE, '', 1);

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

        //>>04May2017
        IF PrinttoExcel THEN BEGIN

            /* ExcBuffer.CreateBook('', 'Ledeger Entries RM');
            ExcBuffer.CreateBookAndOpenExcel('', 'Ledeger Entries RM', '', '', USERID);
            ExcBuffer.GiveUserControl; */

            ExcBuffer.CreateNewBook('Ledeger Entries RM');
            ExcBuffer.WriteSheet('Ledeger Entries RM', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Ledeger Entries RM', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();
        END;
        //<<04May2017
    end;

    trigger OnPreReport()
    begin

        //>>1

        DateFILTER := Item.GETFILTER("Date Filter");
        LocFilter := Item.GETFILTER(Item."Location Filter");
        ItemFilter := Item.GETFILTER(Item."No.");
        BinFilter := Item.GETFILTER(Item."Bin Filter");
        Location.SETRANGE(Location.Code, LocFilter);

        IF Location.FINDFIRST THEN;
        Loc := COPYSTR(Location.Name, 7, 50);
        filtertext := Loc;
        datefiltertext := 'Item Ledger Entries ' + 'For' + ' ' + ItemFilter + ' for the Peroid of  ' + DateFILTER;
        //<<1
    end;

    var
        OPBAL: Decimal;
        QTYREJECTED: Decimal;
        QTYREC: Decimal;
        QTYREC1: Decimal;
        QTYISSUED: Decimal;
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
        datefiltertext: Text[100];
        TotalOPBAL: Decimal;
        TotalQTYREJECTED: Decimal;
        TotalQTYREC: Decimal;
        TotalQTYISSUED: Decimal;
        TotalCLOSEBAL: Decimal;
        Location: Record 14;
        Loc: Text[30];
        Usersetup: Record 91;
        CompName: Text[100];
        OpngBalforNextLine: Decimal;
        DocumentNumber: Text[100];
        ValueEntry: Record 5802;
        PurRcptHdr: Record 120;
        PartyName: Text[100];
        TransferShptHdr: Record 5744;
        ItemFilter: Text[50];
        PrinttoExcel: Boolean;
        RecLocation: Text[100];
        TransferRcptHdr: Record 5746;
        BinFilter: Text[50];
        "---05May2017---": Integer;
        ExcBuffer: Record 370 temporary;
        NNN: Integer;

    //  [Scope('Internal')]
    procedure CreateXLSHEET()
    begin
        /*
        CLEAR(XLAPP);
        CLEAR(XLWBOOK);
        CLEAR(XLSHEET);
        CREATE(XLAPP);
        XLWBOOK := XLAPP.Workbooks.Add;
        XLSHEET := XLWBOOK.Worksheets.Add;
        XLSHEET.Name := 'Grade Led.Entries';
        XLAPP.Visible(FALSE);
        MaxCol:= 'I';
        *///Commented 04May2017

    end;

    // [Scope('Internal')]
    procedure CreateHeader()
    begin
        /*
        XLSHEET.Activate;
        
        XLSHEET.Range('A1',MaxCol+'1').Merge := TRUE;
        XLSHEET.Range('A1',MaxCol+'1').Font.Bold := TRUE;
        XLSHEET.Range('A1',MaxCol+'1').Value :=CompInfo.Name;
        XLSHEET.Range('A1',MaxCol+'1').HorizontalAlignment := 3;
        XLSHEET.Range('A1',MaxCol+'1').Interior.ColorIndex := 8;
        XLSHEET.Range('A1',MaxCol+'1').Borders.ColorIndex := 5;
        XLSHEET.Range('A2',MaxCol+'2').Merge := TRUE;
        
        XLSHEET.Range('A2').Value := Loc;
        XLSHEET.Range('A2').Font.Bold := TRUE;
        XLSHEET.Range('A2',MaxCol+'2').HorizontalAlignment := 3;
        XLSHEET.Range('A2',MaxCol+'2').Interior.ColorIndex := 8;
        XLSHEET.Range('A2',MaxCol+'2').Borders.ColorIndex := 5;
        XLSHEET.Range('A2',MaxCol+'2').Merge := TRUE;
        
        XLSHEET.Range('A3',MaxCol+'3').Merge := TRUE;
        XLSHEET.Range('A3',MaxCol+'3').Font.Bold := TRUE;
        XLSHEET.Range('A3',MaxCol+'3').Value := datefiltertext;
        XLSHEET.Range('A3'+FORMAT(i)).HorizontalAlignment :=-4152;
        XLSHEET.Range('A3',MaxCol+'3').Interior.ColorIndex := 8;
        XLSHEET.Range('A3',MaxCol+'3').Borders.ColorIndex := 5;
        
        XLSHEET.Range('B2',MaxCol+'2').Merge := TRUE;
        
        XLSHEET.Range('A6').Value :='Date';
        XLSHEET.Range('A6').Font.Bold := TRUE;
        XLSHEET.Range('A6',MaxCol+'6').Interior.ColorIndex := 45;
        
        XLSHEET.Range('A6',MaxCol+'6').Borders.ColorIndex := 5;
        
        XLSHEET.Range('B6').Value := ' Opening Balance';
        XLSHEET.Range('B6').Font.Bold := TRUE;
        
        XLSHEET.Range('C6').Value := 'Receipt';
        XLSHEET.Range('C6').Font.Bold := TRUE;
        
        XLSHEET.Range('D6').Value := 'Rejection';
        XLSHEET.Range('D6').Font.Bold := TRUE;
        
        XLSHEET.Range('E6').Value := 'Issued';
        XLSHEET.Range('E6').Font.Bold := TRUE;
        
        XLSHEET.Range('F6').Value := 'Closing Balance';
        XLSHEET.Range('F6').Font.Bold := TRUE;
        
        XLSHEET.Range('G6').Value := 'Document No.';
        XLSHEET.Range('G6').Font.Bold := TRUE;
        
        XLSHEET.Range('H6').Value := 'Issued To Location';
        XLSHEET.Range('H6').Font.Bold := TRUE;
        
        XLSHEET.Range('I6').Value := 'Received From Location';
        XLSHEET.Range('I6').Font.Bold := TRUE;
        
        XLSHEET.Range('A6:I6').Borders.ColorIndex := 5;
        XLAPP.Visible(TRUE);
        */
        //Commented 04May2017

    end;

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


report 50147 "Consolidated RG 23D Report"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ConsolidatedRG23DReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";

            trigger OnAfterGetRecord()
            begin

                //>>1

                totalqtyrow := 0;
                totalLTKgqtyrow := 0;
                totalassesrow := 0;
                totaledonfgrow := 0;

                /*  rg23dRec.RESET;
                 rg23dRec.SETCURRENTKEY("Item No.", "Posting Date", "Transaction Type", Type, "Location Code", "Excise Prod. Posting Group", Closed,
                                 "Document No.");
                 rg23dRec.SETRANGE("Item No.", "No.");
                 rg23dRec.SETRANGE("Posting Date", 010410D, DateFilter);
                 rg23dRec.SETFILTER(Quantity, '>0');
                 IF NOT rg23dRec.FIND('-') THEN */
                CurrReport.SKIP;

                ItemUOM.GET(Item."No.", Item."Sales Unit of Measure");

                ExcelBuf.NewRow;
                ExcelBuf.AddColumn(Item."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn(Item."Sales Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', 1);

                FOR j := 1 TO i DO BEGIN
                    Qty := 0;
                    TotAssbleVal2 := 0;
                    Amt := 0;
                    /*  rg23dRec.RESET;
                     rg23dRec.SETCURRENTKEY("Item No.", "Posting Date", "Transaction Type", Type, "Location Code", "Excise Prod. Posting Group", Closed,
                                    "Document No.");
                     rg23dRec.SETRANGE("Item No.", "No.");
                     rg23dRec.SETRANGE("Posting Date", 0D, DateFilter);
                     rg23dRec.SETRANGE("Location Code", ArrLocCode[j]);
                     IF rg23dRec.FINDFIRST THEN
                         REPEAT
                             AssvValue := 0;
                             AmtRef := 0;
                             IF rg23dRec.Type = rg23dRec.Type::Invoice THEN BEGIN
                                 RGAssRec.RESET;
                                 RGAssRec.SETRANGE(RGAssRec."Entry No.", rg23dRec."Ref. Entry No.");
                                 IF RGAssRec.FINDFIRST THEN BEGIN
                                     AssvValue := RGAssRec."Excise Base Amt Per Unit";
                                     AmtRef := ABS(RGAssRec.Amount / RGAssRec.Quantity);
                                 END;
                             END;
                             Qty += rg23dRec.Quantity;
                             IF AssvValue <> 0 THEN BEGIN
                                 TotAssbleVal2 += rg23dRec.Quantity * ABS(AssvValue) * ItemUOM."Qty. per Unit of Measure";
                                 Amt += AmtRef * rg23dRec.Quantity;
                             END
                             ELSE BEGIN
                                 TotAssbleVal2 += rg23dRec.Quantity * ABS(rg23dRec."Excise Base Amt Per Unit") * ItemUOM."Qty. per Unit of Measure";
                                 Amt += rg23dRec.Amount;
                             END;
                         UNTIL rg23dRec.NEXT = 0; */

                    remarks := '';
                    recItemTable.RESET;
                    recItemTable.SETRANGE(recItemTable."No.", Item."No.");
                    recItemTable.SETRANGE(recItemTable.Blocked, TRUE);
                    IF recItemTable.FINDFIRST THEN BEGIN
                        IF Qty <> 0 THEN BEGIN
                            remarks := 'Blocked';
                        END;
                    END;

                    ExcelBuf.AddColumn(Qty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);
                    ExcelBuf.AddColumn(Qty * ItemUOM."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);
                    ExcelBuf.AddColumn(TotAssbleVal2, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);
                    ExcelBuf.AddColumn(Amt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);

                    totalqtyrow := totalqtyrow + Qty;
                    totalLTKgqtyrow := totalLTKgqtyrow + (Qty * ItemUOM."Qty. per Unit of Measure");
                    totalassesrow := totalassesrow + TotAssbleVal2;
                    totaledonfgrow := totaledonfgrow + Amt;

                    totalqtycol[j] := totalqtycol[j] + Qty;
                    QtyInLT[j] := QtyInLT[j] + (Qty * ItemUOM."Qty. per Unit of Measure");
                    totalassescol[j] := totalassescol[j] + TotAssbleVal2;
                    totaledonfgcol[j] := totaledonfgcol[j] + Amt;

                END;

                ExcelBuf.AddColumn(totalqtyrow, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);
                ExcelBuf.AddColumn(totalLTKgqtyrow, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);
                ExcelBuf.AddColumn(totalassesrow, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);
                ExcelBuf.AddColumn(totaledonfgrow, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);
                ExcelBuf.AddColumn(remarks, FALSE, '', FALSE, FALSE, FALSE, '', 1);
                //<<1
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                ExcelBuf.NewRow;
                ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);

                FOR j := 1 TO i DO BEGIN
                    ExcelBuf.AddColumn(totalqtycol[j], FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                    totalqtyrowTotal += totalqtycol[j];
                    ExcelBuf.AddColumn(QtyInLT[j], FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                    totalLTKgqtyrowTotal += QtyInLT[j];
                    ExcelBuf.AddColumn(totalassescol[j], FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);
                    totalassesrowTotal += totalassescol[j];
                    ExcelBuf.AddColumn(totaledonfgcol[j], FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);
                    totaledonfgrowTotal += totaledonfgcol[j];
                END;

                ExcelBuf.AddColumn(totalqtyrowTotal, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                ExcelBuf.AddColumn(totalLTKgqtyrowTotal, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);
                ExcelBuf.AddColumn(totalassesrowTotal, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);
                ExcelBuf.AddColumn(totaledonfgrowTotal, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);
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
                field("Date Filter"; DateFilter)
                {
                    ApplicationArea = ALL;
                }
                field(Location; Locationcode)
                {
                    ApplicationArea = ALL;
                    TableRelation = Location;
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = ALL;
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

        LocRec.RESET;
        LocRec.SETRANGE("Trading Location", TRUE);

        IF Locationcode <> '' THEN
            LocRec.SETRANGE(Code, Locationcode);

        IF LocRec.FIND('-') THEN
            REPEAT
                i := i + 1;
                ArrLoc[i] := LocRec.Name;
                ArrLocCode[i] := LocRec.Code;
            UNTIL LocRec.NEXT = 0;

        totalqtyrowTotal := 0;
        totalLTKgqtyrowTotal := 0;
        totalassesrowTotal := 0;
        totaledonfgrowTotal := 0;

        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<1
    end;

    var
        i: Integer;
        j: Integer;
        k: Integer;
        LocRec: Record 14;
        ArrLoc: array[150] of Text[150];
        PrintToExcel: Boolean;
        DateFilter: Date;
        //rg23dRec: Record 16537;
        previtem: Code[150];
        Qty: Decimal;
        TotAssbleVal2: Decimal;
        Amt: Decimal;
        ArrLocCode: array[150] of Code[10];
        ExcelBuf: Record 370 temporary;
        Locationcode: Code[20];
        //rg23drecentryno: Record 16537;
        totalqtyrow: Decimal;
        totalLTKgqtyrow: Decimal;
        totalassesrow: Decimal;
        totaledonfgrow: Decimal;
        totalqtycol: array[30000] of Decimal;
        totalassescol: array[30000] of Decimal;
        totaledonfgcol: array[30000] of Decimal;
        ItemUOM: Record 5404;
        QtyInLT: array[30000] of Decimal;
        totalqtyrowTotal: Decimal;
        totalLTKgqtyrowTotal: Decimal;
        totalassesrowTotal: Decimal;
        totaledonfgrowTotal: Decimal;
        remarks: Text[30];
        recItemTable: Record 27;
        //RGAssRec: Record 16537;
        AssvValue: Decimal;
        AmtRef: Decimal;
        Text000: Label 'Data';
        Text001: Label 'Consolidated RG 23D Report';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';

    //   //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//25Apr2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Consolidation Report For EDonFG', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Consolidated RG 23D Report", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    //  //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');


        //>>25Apr2017
        /* ExcelBuf.CreateBook('', Text001);
        ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<25Apr2017
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>25Apr2017 ReportHeader

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text001, FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);

        ExcelBuf.NewRow;
        //<<25Apr2017 ReportHeader

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('As of Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn(DateFilter, FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Product No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Product SKU', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);

        FOR j := 1 TO i DO BEGIN
            ExcelBuf.AddColumn(ArrLoc[j], FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn(ArrLoc[j], FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn(ArrLoc[j], FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn(ArrLoc[j], FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        END;

        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn(' ', FALSE, '', TRUE, FALSE, TRUE, '@', 1);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(' ', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn(' ', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn(' ', FALSE, '', TRUE, FALSE, TRUE, '@', 1);

        FOR k := 1 TO i DO BEGIN
            ExcelBuf.AddColumn('Qty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('Quantity if Lt/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('Total Assessable value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
            ExcelBuf.AddColumn('EDonFG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        END;

        ExcelBuf.AddColumn('Qty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Qty in Lt/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Total Assessable value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('EDonFG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Remarks', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
    end;
}


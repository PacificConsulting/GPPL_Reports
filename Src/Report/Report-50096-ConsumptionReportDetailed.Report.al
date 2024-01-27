report 50096 "Consumption Report Detailed"
{
    // RSPLSUM 15Jun2020---To show reversal consumption entries, removed filters Quantity < 0 and Positive - No from Item Ledger Entry dataitem filter. Object Bak Path - E:\RB-S\15jun20\Report 50096 - Consumption Report Detailed.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ConsumptionReportDetailed.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            CalcFields = "Cost Amount (Expected)", "Cost Amount (Actual)";
            DataItemTableView = SORTING("Document No.")
                                ORDER(Ascending)
                                WHERE("Entry Type" = FILTER(Consumption));
            RequestFilterFields = "Posting Date", "Location Code", "Item Category Code";// "Product Group Code";
            column(Name_CompInfo; CompInfo.Name)
            {
            }
            column(ItemNo_ItemLedgerEntry; "Item Ledger Entry"."Item No.")
            {
            }
            column(Code_Location; Location.Name)
            {
            }
            column(ReportFilter; 'Raw Material Consumption Report for Period:' + ReportFilter)
            {
            }
            column(SrNo; SrNo)
            {
            }
            column(Desc_Item; Item.Description)
            {
            }
            column(varitemcategory; varitemcategory)
            {
            }
            column(DocumentNo_ItemLedgerEntry; "Item Ledger Entry"."Document No.")
            {
            }
            column(UnitofMeasureCode_ItemLedgerEntry; "Item Ledger Entry"."Unit of Measure Code")
            {
            }
            column(Quantity_ItemLedgerEntry; "Item Ledger Entry".Quantity)
            {
            }
            column(ExcsideProdPosting; 'Item."Excise Prod. Posting Group"')
            {
            }
            column(CostAmountExpected_ItemLedgerEntry; "Item Ledger Entry"."Cost Amount (Expected)")
            {
            }
            column(CostAmountActual_ItemLedgerEntry; "Item Ledger Entry"."Cost Amount (Actual)")
            {
            }

            trigger OnAfterGetRecord()
            begin

                ProductionOrder.RESET;
                ProductionOrder.SETRANGE("No.", "Document No.");
                IF ProductionOrder.FINDFIRST THEN;

                /*
                CLEAR(VarLotNo);
                ILE.RESET();
                ILE.SETRANGE(ILE."Document No.","Item Ledger Entry"."Document No.");
                ILE.SETFILTER(ILE."Entry Type",'=%1',"Item Ledger Entry"."Entry Type"::Output);
                IF ILE.FINDFIRST THEN
                    VarLotNo:=ILE."Lot No.";
                */
                CLEAR(VarLotNo);
                VarLotNo := ProductionOrder."Description 2";

                recItem.RESET;
                recItem.SETRANGE(recItem."No.", "Item Ledger Entry"."Source No.");
                IF recItem.FINDFIRST THEN;

                SrNo := SrNo + 1;
                Item.GET("Item Ledger Entry"."Item No.");

                varitemcategory := '';
                IF Item."Item Category Code" <> '' THEN BEGIN
                    IF itemcategory.GET(Item."Item Category Code") THEN
                        varitemcategory := itemcategory.Description;
                END;

                varproductgroup := '';
                //  IF Item."Product Group Code" <> '' THEN BEGIN
                // productgroup.RESET;
                // productgroup.SETRANGE(Code, Item."Product Group Code");
                //  IF productgroup.FIND('-') THEN
                varproductgroup := '';// productgroup.Description;
                //  END;

                i := 1;
                k := 1;
                CLEAR(TestGRNNo);
                //CLEAR(GRNDocNo);
                CLEAR(TestDocNo);
                CLEAR(PrintGRNNo);
                CLEAR(TestPosDate);
                CLEAR(ExbondDate);
                GRNFound := FALSE;
                txtGRN := '';

                IF ("Item Ledger Entry"."Item Category Code" <> 'CAT10') AND ("Item Ledger Entry"."Item Category Code" <> 'CAT02') AND
                     ("Item Ledger Entry"."Item Category Code" <> 'CAT03') AND ("Item Ledger Entry"."Item Category Code" <> 'CAT11') AND
                      ("Item Ledger Entry"."Item Category Code" <> 'CAT12') AND ("Item Ledger Entry"."Item Category Code" <> 'CAT13') AND
                        ("Item Ledger Entry"."Item Category Code" <> 'CAT15') AND ("Item Ledger Entry"."Item Category Code" <> 'CAT18') AND
                          ("Item Ledger Entry"."Item Category Code" <> 'CAT19') AND ("Item Ledger Entry"."Item Category Code" <> 'CAT20') AND
                            ("Item Ledger Entry"."Item Category Code" <> 'CAT21') THEN BEGIN
                    //  FindGRN;
                    FindGRNYSR("Item Ledger Entry", txtGRN);
                    //MESSAGE(txtGRN);

                END;

                txtGRN := COPYSTR(txtGRN, 2, STRLEN(txtGRN));

                IF PrintToExcel THEN
                    MakeExcelDataBody;
                vCount := "Item Ledger Entry".COUNT;

                TotalAmt += "Item Ledger Entry"."Cost Amount (Actual)" + "Item Ledger Entry"."Cost Amount (Expected)";

            end;

            trigger OnPostDataItem()
            begin
                IF PrintToExcel THEN
                    MakeExcelDataFooter;
            end;

            trigger OnPreDataItem()
            begin
                TotalAmt := 0;
                MakeExcelDataHeader;
            end;
        }
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number);
            column(Number_Integer; Integer.Number)
            {
            }

            trigger OnPreDataItem()
            begin
                Integer.SETRANGE(Number, 1, 20 - vCount);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(PrintToExcel; PrintToExcel)
                {
                    Caption = 'Export to Excel';
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
        CompInfo.GET;
        ReportFilter := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Posting Date");

        IF PrintToExcel THEN
            MakeExcelInfo;
    end;

    var
        CompInfo: Record 79;
        Usersetup: Record 91;
        ReportFilter: Text[50];
        date: Date;
        total: Decimal;
        Loc: Text[30];
        Location: Record 14;
        SrNo: Integer;
        Item: Record 27;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        itemcategory: Record 5722;
        //productgroup: Record 5723;
        productsubgroup: Record 50015;
        productgrade: Record 50016;
        varitemcategory: Text[200];
        varproductgroup: Text[200];
        varproductsubgroup: Text[200];
        varproductgrade: Text[200];
        VarLotNo: Code[20];
        ILE: Record 32;
        ProductionOrder: Record 5405;
        recItem: Record 27;
        Text000: Label 'Data';
        Text001: Label 'Consumption Report Detailed';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "--RSPL-Mig--": Integer;
        vCount: Integer;
        AccCont_Rec: Record 2000000053;
        TotalAmt: Decimal;
        GRNFound: Boolean;
        TestInt: Integer;
        i: Integer;
        RecILE: Record 32;
        TestDocNo: Text;
        TestPosDate: Date;
        RecPRL: Record 121;
        RecItemLE: Record 32;
        RecVE: Record 5802;
        PurchInvLine: Record 123;
        ExbondDate: Text;
        j: Integer;
        txtGRN: Text;
        TestGRNNo: array[20] of Text;
        GRNForLoopFound: Boolean;
        PrintGRNNo: array[10] of Text;
        k: Integer;

    //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('Consumption Report Detailed', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"Consumption Report Detailed", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        //MakeExcelDataHeader;
    end;

    //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//07Aug2019
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        Location.RESET;
        Location.SETRANGE(Code, "Item Ledger Entry".GETFILTER("Location Code"));
        IF Location.FINDFIRST THEN
            Loc := Location.Name;

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location :' + Loc, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//07Aug2019
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Sr. No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Date of Consumption / Production', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Quantity Issued', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Consumption / Production Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//07Aug2019
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Chapter Heading', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('FPO Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        /*
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1','ITEMDESC-PRODORDER');
        IF MemberOf.FINDFIRST THEN
        BEGIN
          ExcelBuf.AddColumn('Output Item Name',FALSE,'',TRUE,FALSE,FALSE,'@');
        END;
        */

        AccCont_Rec.RESET;
        AccCont_Rec.SETRANGE(AccCont_Rec."User Name", USERID);
        AccCont_Rec.SETFILTER(AccCont_Rec."Role ID", '%1', 'ITEMDESC-PRODORDER');
        IF AccCont_Rec.FINDFIRST THEN BEGIN
            ExcelBuf.AddColumn('Output Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        END;

        ExcelBuf.AddColumn('GRN', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

    end;

    //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(SrNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Item."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Item."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(varitemcategory, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Item Ledger Entry"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn("Item Ledger Entry".Quantity,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("Item Ledger Entry".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Cost Amount (Actual)" + "Item Ledger Entry"."Cost Amount (Expected)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//07Aug2019
        ExcelBuf.AddColumn("Item Ledger Entry"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Item Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(''/*Item."Excise Prod. Posting Group"*/, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(VarLotNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ProductionOrder."Finished Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        /*
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1','ITEMDESC-PRODORDER');
        IF MemberOf.FINDFIRST THEN
        BEGIN
          ExcelBuf.AddColumn(recItem.Description,FALSE,'',FALSE,FALSE,FALSE,'');
        END;
        */

        AccCont_Rec.RESET;
        AccCont_Rec.SETRANGE(AccCont_Rec."User Name", USERID);
        AccCont_Rec.SETFILTER(AccCont_Rec."Role ID", '%1', 'ITEMDESC-PRODORDER');
        IF AccCont_Rec.FINDFIRST THEN BEGIN
            ExcelBuf.AddColumn(recItem.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        END;

        FOR j := 1 TO k - 1 DO
            ExcelBuf.AddColumn(COPYSTR(PrintGRNNo[j], 2, STRLEN(PrintGRNNo[j])), FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(txtGRN,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);

    end;

    //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        /* ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', '');
        ExcelBuf.GiveUserControl;
 */
        ExcelBuf.CreateNewBook(Text000);
        ExcelBuf.WriteSheet(Text000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        "Item Ledger Entry".CALCSUMS(Quantity);
        //ExcelBuf.AddColumn("Item Ledger Entry".Quantity,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("Item Ledger Entry".Quantity, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn(TotalAmt, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//07Aug2019
        TotalAmt := 0;
    end;

    local procedure FindGRNYSR(varRecILeOrg: Record 32; var GRNNo: Text)
    var
        varRecILe: Record 32;
        cdDoc: Code[20];
        cdDoc1: Code[20];
        cdDoc2: Code[20];
        ItemApplnEntry: Record 339;
        ItemLedgEntry: Record 32;
    begin
        //Navigate
        CLEAR(cdDoc);
        CLEAR(cdDoc1);
        CLEAR(cdDoc2);

        varRecILe.RESET;
        varRecILe.SETCURRENTKEY("Document No.", "Posting Date", "Item No.");
        varRecILe.SETRANGE("Document No.", varRecILeOrg."Document No.");
        varRecILe.SETRANGE("Item No.", varRecILeOrg."Item No.");
        IF varRecILe.FINDSET THEN
            REPEAT
                IF cdDoc <> varRecILe."Document No." THEN BEGIN
                    cdDoc := varRecILe."Document No.";
                    IF varRecILe.Positive THEN BEGIN
                        ItemApplnEntry.RESET;
                        ItemApplnEntry.SETCURRENTKEY("Inbound Item Entry No.", "Outbound Item Entry No.", "Cost Application");
                        ItemApplnEntry.SETRANGE("Inbound Item Entry No.", varRecILe."Entry No.");
                        ItemApplnEntry.SETFILTER("Outbound Item Entry No.", '<>%1', 0);
                        ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                        IF ItemApplnEntry.FIND('-') THEN
                            REPEAT
                                //InsertTempEntry(ItemApplnEntry."Outbound Item Entry No.",ItemApplnEntry.Quantity);

                                ItemLedgEntry.GET(ItemApplnEntry."Outbound Item Entry No.");
                                IF cdDoc1 <> ItemLedgEntry."Document No." THEN BEGIN
                                    cdDoc1 := ItemLedgEntry."Document No.";

                                    IF NOT (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Consumption) THEN BEGIN
                                        IF ItemApplnEntry.Quantity * ItemLedgEntry.Quantity >= 0 THEN BEGIN
                                            //YSR BEGIN
                                            //---Find Transfr
                                            IF ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Transfer THEN BEGIN
                                                FindGRNYSR(ItemLedgEntry, GRNNo);
                                            END ELSE
                                                //YSR END
                                                IF (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Purchase) OR
                                                    (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::"Positive Adjmt.") THEN BEGIN
                                                    GRNFound := TRUE;
                                                    //GRNDocNo[i] += ',' + ItemLedgEntry."Document No.";
                                                    GRNNo += ',' + ItemLedgEntry."Document No.";
                                                    //MESSAGE('%1', GRNNo );

                                                    i += 1;

                                                    EXIT;
                                                END;
                                        END;
                                    END;
                                END;
                            UNTIL ItemApplnEntry.NEXT = 0;
                    END ELSE BEGIN
                        ItemApplnEntry.RESET;
                        ItemApplnEntry.SETCURRENTKEY("Outbound Item Entry No.", "Item Ledger Entry No.", "Cost Application");
                        ItemApplnEntry.SETRANGE("Outbound Item Entry No.", varRecILe."Entry No.");
                        ItemApplnEntry.SETRANGE("Item Ledger Entry No.", varRecILe."Entry No.");
                        ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                        IF ItemApplnEntry.FIND('-') THEN
                            REPEAT
                                ItemLedgEntry.GET(ItemApplnEntry."Inbound Item Entry No.");
                                IF cdDoc2 <> ItemLedgEntry."Document No." THEN BEGIN
                                    cdDoc2 := ItemLedgEntry."Document No.";

                                    IF NOT (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Consumption) THEN BEGIN
                                        IF -ItemApplnEntry.Quantity * ItemLedgEntry.Quantity >= 0 THEN BEGIN
                                            //YSR BEGIN
                                            //---Find Transfr
                                            IF ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Transfer THEN BEGIN
                                                FindGRNYSR(ItemLedgEntry, GRNNo);
                                            END ELSE
                                                //YSR END
                                                IF (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Purchase) OR
                                                    (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::"Positive Adjmt.") THEN BEGIN
                                                    GRNFound := TRUE;
                                                    RecPRL.RESET;
                                                    IF RecPRL.GET(ItemLedgEntry."Document No.", ItemLedgEntry."Document Line No.") THEN BEGIN
                                                        RecItemLE.RESET;
                                                        RecItemLE.SETCURRENTKEY("Document No.");
                                                        RecItemLE.SETRANGE("Document No.", RecPRL."Document No.");
                                                        RecItemLE.SETRANGE("Document Type", RecItemLE."Document Type"::"Purchase Receipt");
                                                        RecItemLE.SETRANGE("Document Line No.", RecPRL."Line No.");
                                                        RecItemLE.SETFILTER("Invoiced Quantity", '<>0');
                                                        IF RecItemLE.FINDSET THEN BEGIN
                                                            RecVE.SETCURRENTKEY("Item Ledger Entry No.", "Entry Type");
                                                            RecVE.SETRANGE("Entry Type", RecVE."Entry Type"::"Direct Cost");
                                                            RecVE.SETFILTER("Invoiced Quantity", '<>0');
                                                            REPEAT
                                                                RecVE.SETRANGE("Item Ledger Entry No.", RecItemLE."Entry No.");
                                                                IF RecVE.FINDSET THEN
                                                                    REPEAT
                                                                        IF RecVE."Document Type" = RecVE."Document Type"::"Purchase Invoice" THEN
                                                                            IF PurchInvLine.GET(RecVE."Document No.", RecVE."Document Line No.") THEN BEGIN
                                                                                ExbondDate += ', ' + FORMAT(PurchInvLine."Exbond Rate", 0, 2);
                                                                            END;
                                                                    UNTIL RecVE.NEXT = 0;
                                                            UNTIL RecItemLE.NEXT = 0;
                                                        END;
                                                    END;
                                                    ExbondDate := COPYSTR(ExbondDate, 2, STRLEN(ExbondDate));
                                                    //GRNDocNo[i] += ',' + ItemLedgEntry."Document No." + ' - ' + FORMAT(ItemLedgEntry.Quantity,0,2) + ' - ' + ExbondDate;
                                                    FOR j := 1 TO i DO//s
                                                    BEGIN//s
                                                        IF TestGRNNo[j] = ItemLedgEntry."Document No." THEN//s
                                                            GRNForLoopFound := TRUE;//s
                                                    END;//s
                                                    IF NOT GRNForLoopFound THEN BEGIN//s
                                                        GRNNo += ',' + ItemLedgEntry."Document No." + ' - ' + FORMAT(ItemLedgEntry.Quantity, 0, 2) + ' - ' + ExbondDate;//ysr
                                                        PrintGRNNo[k] += ',' + ItemLedgEntry."Document No." + ' - ' + FORMAT(ItemLedgEntry.Quantity, 0, 2) + ' - ' + ExbondDate;
                                                        TestGRNNo[i] := ItemLedgEntry."Document No.";//s
                                                        i += 1;
                                                        k += 1;
                                                    END;
                                                    GRNForLoopFound := FALSE;
                                                    ExbondDate := '';
                                                END;
                                        END;
                                    END;//Consmption End
                                END;

                            UNTIL ItemApplnEntry.NEXT = 0;
                    END;
                END;
            UNTIL varRecILe.NEXT = 0;
    end;
}


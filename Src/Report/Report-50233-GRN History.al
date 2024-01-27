report 50233 "GRN History"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/GRNHistory.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            CalcFields = "Cost Amount (Expected)", "Cost Amount (Actual)";
            DataItemTableView = SORTING("Document No.")
                                ORDER(Ascending)
                                WHERE("Entry Type" = FILTER(Purchase),
                                      "Document Type" = FILTER("Purchase Receipt"));
            RequestFilterFields = "Entry No.";

            trigger OnAfterGetRecord()
            begin

                FindSale("Item Ledger Entry");
            end;

            trigger OnPreDataItem()
            begin
                MakeExcelDataHeader;
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
                    ApplicationArea = all;
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

        ReportFilterEntryNo := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Entry No.");

        IF ReportFilterEntryNo = '' THEN
            ERROR('Please enter Entry No.');

        RecItemLedgEntry.RESET;
        IF RecItemLedgEntry.GET(ReportFilterEntryNo) THEN
            IF (RecItemLedgEntry."Entry Type" <> RecItemLedgEntry."Entry Type"::Purchase) OR
                (RecItemLedgEntry."Document Type" <> RecItemLedgEntry."Document Type"::"Purchase Receipt") THEN
                ERROR('You can only search history of GRN');

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
        // productgroup: Record 5723;
        productsubgroup: Record 50015;
        productgrade: Record 50016;
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
        RecItemLedgEntry: Record 32;
        ReportFilterEntryNo: Text;
        ExcelBufDuplicate: Record 370 temporary;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('GRN History', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"GRN History", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
    end;

    // //[Scope('Internal')]
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
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Entry No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Entry Type', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document Type', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Cost Amount (Actual)', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Cost Amount (Expected)', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        /* ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', '');
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text000);
        ExcelBuf.WriteSheet(Text000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    local procedure FindSale(varRecILeOrg: Record 32)
    var
        varRecILe: Record 32;
        cdDoc: Code[20];
        cdDoc1: Code[20];
        cdDoc2: Code[20];
        ItemApplnEntry: Record 339;
        ItemLedgEntry: Record 32;
        RecExcelBuffer: Record 370;
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

                //IF varRecILe.Positive THEN BEGIN
                ExcelBufDuplicate.RESET;
                ExcelBufDuplicate.SETFILTER("Cell Value as Text", '%1', FORMAT(varRecILe."Entry No."));
                IF NOT ExcelBufDuplicate.FINDFIRST THEN BEGIN
                    ExcelBuf.NewRow;
                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn(varRecILe."Entry No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(varRecILe."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(varRecILe."Entry Type", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(varRecILe."Document Type", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(varRecILe."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(varRecILe."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(varRecILe."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(varRecILe.Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', ExcelBuf."Cell Type"::Number);
                    varRecILe.CALCFIELDS(varRecILe."Cost Amount (Actual)", varRecILe."Cost Amount (Expected)");
                    ExcelBuf.AddColumn(varRecILe."Cost Amount (Actual)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(varRecILe."Cost Amount (Expected)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', ExcelBuf."Cell Type"::Number);
                    CreateDuplicateExcelBuf(ExcelBuf);
                END;
                //END;

                IF (cdDoc <> varRecILe."Document No.") OR
                      ((varRecILe."Document Type" = varRecILe."Document Type"::"Transfer Shipment") AND (varRecILe."Location Code" = 'IN-TRANS')) OR
                         ((varRecILe."Document Type" = varRecILe."Document Type"::"Transfer Receipt") AND (varRecILe."Location Code" <> 'IN-TRANS')) THEN BEGIN

                    cdDoc := varRecILe."Document No.";

                    IF varRecILe.Positive THEN BEGIN
                        ItemApplnEntry.RESET;
                        ItemApplnEntry.SETCURRENTKEY("Inbound Item Entry No.", "Outbound Item Entry No.", "Cost Application");
                        ItemApplnEntry.SETRANGE("Inbound Item Entry No.", varRecILe."Entry No.");
                        ItemApplnEntry.SETFILTER("Outbound Item Entry No.", '<>%1', 0);
                        ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                        IF ItemApplnEntry.FIND('-') THEN
                            REPEAT

                                ItemLedgEntry.GET(ItemApplnEntry."Outbound Item Entry No.");
                                IF cdDoc1 <> ItemLedgEntry."Document No." THEN BEGIN
                                    cdDoc1 := ItemLedgEntry."Document No.";

                                    IF (ItemLedgEntry."Entry Type" <> ItemLedgEntry."Entry Type"::Consumption) AND
                                        (ItemLedgEntry."Entry Type" <> ItemLedgEntry."Entry Type"::Sale) AND
                                          (ItemLedgEntry."Entry Type" <> ItemLedgEntry."Entry Type"::"Negative Adjmt.") THEN BEGIN

                                        IF ItemApplnEntry.Quantity * ItemLedgEntry.Quantity >= 0 THEN BEGIN

                                            ExcelBuf.NewRow;
                                            ExcelBuf.NewRow;
                                            ExcelBuf.AddColumn(ItemLedgEntry."Entry No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                            ExcelBuf.AddColumn(ItemLedgEntry."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                            ExcelBuf.AddColumn(ItemLedgEntry."Entry Type", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                            ExcelBuf.AddColumn(ItemLedgEntry."Document Type", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                            ExcelBuf.AddColumn(ItemLedgEntry."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                            ExcelBuf.AddColumn(ItemLedgEntry."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                            ExcelBuf.AddColumn(ItemLedgEntry."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                            ExcelBuf.AddColumn(ItemLedgEntry.Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', ExcelBuf."Cell Type"::Number);
                                            ItemLedgEntry.CALCFIELDS(ItemLedgEntry."Cost Amount (Actual)", ItemLedgEntry."Cost Amount (Expected)");
                                            ExcelBuf.AddColumn(ItemLedgEntry."Cost Amount (Actual)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', ExcelBuf."Cell Type"::Number);
                                            ExcelBuf.AddColumn(ItemLedgEntry."Cost Amount (Expected)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', ExcelBuf."Cell Type"::Number);
                                            CreateDuplicateExcelBuf(ExcelBuf);

                                            //---Find Transfr
                                            IF ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Transfer THEN BEGIN
                                                FindSale(ItemLedgEntry);
                                            END;

                                        END;
                                    END ELSE BEGIN
                                        ExcelBuf.NewRow;
                                        ExcelBuf.NewRow;
                                        ExcelBuf.AddColumn(ItemLedgEntry."Entry No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                        ExcelBuf.AddColumn(ItemLedgEntry."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                        ExcelBuf.AddColumn(ItemLedgEntry."Entry Type", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                        ExcelBuf.AddColumn(ItemLedgEntry."Document Type", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                        ExcelBuf.AddColumn(ItemLedgEntry."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                        ExcelBuf.AddColumn(ItemLedgEntry."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                        ExcelBuf.AddColumn(ItemLedgEntry."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                                        ExcelBuf.AddColumn(ItemLedgEntry.Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', ExcelBuf."Cell Type"::Number);
                                        ItemLedgEntry.CALCFIELDS(ItemLedgEntry."Cost Amount (Actual)", ItemLedgEntry."Cost Amount (Expected)");
                                        ExcelBuf.AddColumn(ItemLedgEntry."Cost Amount (Actual)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', ExcelBuf."Cell Type"::Number);
                                        ExcelBuf.AddColumn(ItemLedgEntry."Cost Amount (Expected)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', ExcelBuf."Cell Type"::Number);
                                        CreateDuplicateExcelBuf(ExcelBuf);
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
                            /*
                            ItemLedgEntry.GET(ItemApplnEntry."Inbound Item Entry No.");
                            IF  cdDoc2 <>  ItemLedgEntry."Document No."  THEN  BEGIN
                              cdDoc2 :=  ItemLedgEntry."Document No.";

                              //IF NOT (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Consumption) THEN BEGIN
                              IF NOT (ItemLedgEntry."Entry Type" = ItemLedgEntry."Entry Type"::Purchase) THEN BEGIN
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
                                      IF RecPRL.GET(ItemLedgEntry."Document No.",ItemLedgEntry."Document Line No.") THEN BEGIN
                                        RecItemLE.RESET;
                                        RecItemLE.SETCURRENTKEY("Document No.");
                                        RecItemLE.SETRANGE("Document No.",RecPRL."Document No.");
                                        RecItemLE.SETRANGE("Document Type",RecItemLE."Document Type"::"Purchase Receipt");
                                        RecItemLE.SETRANGE("Document Line No.",RecPRL."Line No.");
                                        RecItemLE.SETFILTER("Invoiced Quantity",'<>0');
                                        IF RecItemLE.FINDSET THEN BEGIN
                                          RecVE.SETCURRENTKEY("Item Ledger Entry No.","Entry Type");
                                          RecVE.SETRANGE("Entry Type",RecVE."Entry Type"::"Direct Cost");
                                          RecVE.SETFILTER("Invoiced Quantity",'<>0');
                                          REPEAT
                                            RecVE.SETRANGE("Item Ledger Entry No.",RecItemLE."Entry No.");
                                            IF RecVE.FINDSET THEN
                                              REPEAT
                                                IF RecVE."Document Type" = RecVE."Document Type"::"Purchase Invoice" THEN
                                                  IF PurchInvLine.GET(RecVE."Document No.",RecVE."Document Line No.") THEN BEGIN
                                                    ExbondDate += ', ' + FORMAT(PurchInvLine."Exbond Rate",0,2);
                                                  END;
                                              UNTIL RecVE.NEXT = 0;
                                          UNTIL RecItemLE.NEXT = 0;
                                        END;
                                      END;
                                      ExbondDate := COPYSTR(ExbondDate,2,STRLEN(ExbondDate));
                                      //GRNDocNo[i] += ',' + ItemLedgEntry."Document No." + ' - ' + FORMAT(ItemLedgEntry.Quantity,0,2) + ' - ' + ExbondDate;
                                      FOR j := 1 TO i DO//s
                                      BEGIN//s
                                        IF TestGRNNo[j] = ItemLedgEntry."Document No." THEN//s
                                          GRNForLoopFound := TRUE;//s
                                      END;//s
                                      IF NOT GRNForLoopFound THEN BEGIN//s
                                        GRNNo += ',' + ItemLedgEntry."Document No." + ' - ' + FORMAT(ItemLedgEntry.Quantity,0,2) + ' - ' + ExbondDate;//ysr
                                        PrintGRNNo[k] += ',' + ItemLedgEntry."Document No." + ' - ' + FORMAT(ItemLedgEntry.Quantity,0,2) + ' - ' + ExbondDate;
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
                            */
                            UNTIL ItemApplnEntry.NEXT = 0;
                    END;
                END;
            UNTIL varRecILe.NEXT = 0;

    end;

    local procedure CreateDuplicateExcelBuf(ExBufP: Record 370)
    begin
        //ExcelBufDuplicate.INIT;
        ExcelBufDuplicate.TRANSFERFIELDS(ExBufP);
        ExcelBufDuplicate.INSERT;
    end;
}


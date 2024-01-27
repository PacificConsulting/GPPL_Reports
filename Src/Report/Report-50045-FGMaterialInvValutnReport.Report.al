report 50045 "FG Material Inv Valutn Report"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/FGMaterialInvValutnReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            DataItemTableView = SORTING("Location Code", "Document Type", "Posting Date", "Item Category Code", "Item No.")
                                WHERE("Document Type" = FILTER("Transfer Receipt" | "Sales Return Receipt" | ' '));
            RequestFilterFields = "Location Code", "Item Category Code";

            trigger OnAfterGetRecord()
            begin

                //>>09Mar2017 LocName1
                CLEAR(LocName1);
                recLOC.RESET;
                recLOC.SETRANGE(Code, "Item Ledger Entry"."Location Code");
                IF recLOC.FINDFIRST THEN BEGIN
                    LocName1 := recLOC.Name;
                END;
                //<< LocName1
                ItemRec.GET("Item Ledger Entry"."Item No.");
                VendorName := '';
                IF VendorRec.GET("Item Ledger Entry"."Source No.") THEN
                    VendorName := VendorRec.Name;

                RemQty := 0;
                ItemApplication.RESET;
                ItemApplication.SETRANGE(ItemApplication."Inbound Item Entry No.", "Item Ledger Entry"."Entry No.");
                ItemApplication.SETRANGE(ItemApplication."Posting Date", 0D, PostingDate);
                IF ItemApplication.FINDSET THEN
                    REPEAT
                        RemQty += ItemApplication.Quantity;
                    UNTIL ItemApplication.NEXT = 0;

                IF RemQty = 0 THEN
                    IF NOT Showall THEN
                        CurrReport.SKIP;

                RateofItem := 0;
                TotalCost := 0;
                TotalQty := 0;
                ChargeofItem := 0;
                ValueEntry.RESET;
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Charge No.", '');
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        TotalCost += ValueEntry."Cost Amount (Expected)" + ValueEntry."Cost Amount (Actual)";
                        TotalQty += ValueEntry."Item Ledger Entry Quantity";
                    UNTIL ValueEntry.NEXT = 0;
                RateofItem := TotalCost / TotalQty;
                ValueEntry.RESET;
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '<>%1', '');
                IF ValueEntry.FINDFIRST THEN
                    REPEAT
                        ChargeofItem += ValueEntry."Cost per Unit";
                    UNTIL ValueEntry.NEXT = 0;
                IF ChargeofItem < 0 THEN BEGIN
                    RateofItem := RateofItem + ChargeofItem;
                    ChargeofItem := 0;
                END;

                MakeExcelDataBody;
            end;

            trigger OnPreDataItem()
            begin
                "Item Ledger Entry".SETRANGE("Item Ledger Entry"."Posting Date", 0D, PostingDate);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(PostingDate; PostingDate)
                {
                    Caption = 'Posting Date';
                }
                field(Showall; Showall)
                {
                    Caption = 'Show Zero Rem. Qty';
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
        CreateExcelbook;
    end;

    trigger OnPreReport()
    begin


        //>>09Mar2017 LocName
        LocCode := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code");
        recLOC.RESET;
        recLOC.SETRANGE(Code, LocCode);
        IF recLOC.FINDFIRST THEN BEGIN
            LocName := recLOC.Name;
        END;
        //<< LocName

        ExcelBuf.RESET;
        ExcelBuf.DELETEALL;
        MakeExcelInfo;
    end;

    var
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        ItemRec: Record 27;
        VendorRec: Record 23;
        ValueEntry: Record 5802;
        RateofItem: Decimal;
        ChargeofItem: Decimal;
        ItemApplication: Record 339;
        RemQty: Decimal;
        PostingDate: Date;
        Showall: Boolean;
        TotalQty: Decimal;
        TotalCost: Decimal;
        VendorName: Text[60];
        UNITCOST: Decimal;
        "----09Mar2017": Integer;
        LocCode: Code[100];
        recLOC: Record 14;
        LocName: Text[250];
        LocName1: Text[100];
        ExcBuffer: Record 370 temporary;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        MakeExcelDataHeader;
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>09Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GP Petroleums Limited', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
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
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('FG Material Valuation report', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
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
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        //<<
        //ExcelBuf.AddColumn('Raw Material Valuation report',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('FG Material Valuation report',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);//09Mar2017


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Inventory As On', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(PostingDate, FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(LocName, FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn("Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code"),FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(USERID,FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//4 //09Mar2017
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn('Original Qty in SUOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//6
        ExcelBuf.AddColumn('Original Qty in Kg / Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//7
        ExcelBuf.AddColumn('Remaining Qty in SUOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//8
        ExcelBuf.AddColumn('Remaining Qty in Kg / Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//9
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//10
        ExcelBuf.AddColumn('Rate', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//11
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//12
        ExcelBuf.AddColumn('Total Rate / Ltr. Kg', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//13
        ExcelBuf.AddColumn('Total Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//14
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Item Ledger Entry"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn(ItemRec.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//2
        ExcelBuf.AddColumn("Item Ledger Entry"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//3
        ExcelBuf.AddColumn(LocName1, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//09Mar2017 //4
        //ExcelBuf.AddColumn("Item Ledger Entry"."Location Code",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Item Ledger Entry"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//5
        ExcelBuf.AddColumn("Item Ledger Entry".Quantity / "Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//6
        ExcelBuf.AddColumn("Item Ledger Entry".Quantity, FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//7
        ExcelBuf.AddColumn(RemQty / "Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//8
        ExcelBuf.AddColumn(RemQty, FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//9
        ExcelBuf.AddColumn("Item Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//10
        ExcelBuf.AddColumn(ROUND(RateofItem, 0.00001), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//11
        ExcelBuf.AddColumn(ChargeofItem, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//12
        ExcelBuf.AddColumn(ROUND((RateofItem + ChargeofItem), 0.00001), FALSE, '', FALSE, FALSE, FALSE, '#,####0.000', ExcelBuf."Cell Type"::Number);//13
        ExcelBuf.AddColumn(ROUND((RemQty * RateofItem + ChargeofItem), 0.00001), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);//14
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet('Data','Raw Material Inventory Valuation Report',COMPANYNAME,USERID);
        //ExcelBuf.CreateBookAndOpenExcel('','Raw Material Inventory Valuation Report','','','');
        // ExcelBuf.CreateBookAndOpenExcel('', 'FG Material Inventory Valuation Report', '', '', '');
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        ExcBuffer.CreateNewBook('FG Material Inventory Valuation Report');
        ExcBuffer.WriteSheet('FG Material Inventory Valuation Report', CompanyName, UserId);
        ExcBuffer.CloseBook();
        ExcBuffer.SetFriendlyFilename(StrSubstNo('FG Material Inventory Valuation Report', CurrentDateTime, UserId));
        ExcBuffer.OpenExcel();
    end;
}


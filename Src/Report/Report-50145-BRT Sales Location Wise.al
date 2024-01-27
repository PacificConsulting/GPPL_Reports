report 50145 "BRT Sales Location Wise"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BRTSalesLocationWise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Location; Location)
        {
            DataItemTableView = SORTING(Code)
                                ORDER(Ascending);
            dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
            {
                DataItemLink = "Transfer-from Code" = FIELD(Code);
                DataItemTableView = SORTING("Posting Date")
                                    ORDER(Ascending);
                RequestFilterFields = "Posting Date";
                dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
                {
                    DataItemLink = "Document No." = FIELD("No.");

                    trigger OnAfterGetRecord()
                    begin

                        Qty += "Transfer Shipment Line".Quantity;
                        "RNet Value" += "Transfer Shipment Line".Amount;
                        "RGross Value" := "Transfer Shipment Line".Amount + 0;// "Transfer Shipment Line"."BED Amount" + "Transfer Shipment Line"."eCess Amount"
                                                                              //+ "Transfer Shipment Line"."SHE Cess Amount" + "RAddl.Duty";
                        "RBed Amount" += 0;//"Transfer Shipment Line"."BED Amount";
                        "REcess Amt" += 0;// "Transfer Shipment Line"."eCess Amount";
                        "RShe Cess Amt" += 0;// "Transfer Shipment Line"."SHE Cess Amount";
                        "RAddl.Duty" += 0;//"Transfer Shipment Line"."ADE Amount";
                    end;

                    trigger OnPreDataItem()
                    begin

                        "Transfer Shipment Line".RESET;
                        "Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Document No.", "Transfer Shipment Header"."No.");
                        "Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Transfer-from Code", "Transfer Shipment Header"."Transfer-from Code");
                    end;
                }

                trigger OnPostDataItem()
                begin

                    IF PrintToExcel THEN
                        MakeExcelDataFooter1;
                end;

                trigger OnPreDataItem()
                begin

                    Qty := 0;
                    "RNet Value" := 0;
                    "RGross Value" := 0;
                    "RBed Amount" := 0;
                    "REcess Amt" := 0;
                    "RShe Cess Amt" := 0;
                    "RAddl.Duty" := 0;
                end;
            }

            trigger OnPostDataItem()
            begin

                IF PrintToExcel THEN
                    MakeExcelDataFooter2;
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

        IF PrintToExcel THEN
            MakeExcelInfo;
    end;

    var
        Qty: Decimal;
        "RNet Value": Decimal;
        "RGross Value": Decimal;
        "RBed Amount": Decimal;
        "REcess Amt": Decimal;
        "RShe Cess Amt": Decimal;
        "RAddl.Duty": Decimal;
        //PostedStrOrderLineDetails: Record 13798;
        GQty: Decimal;
        "GRNet Value": Decimal;
        "GRGross Value": Decimal;
        "GRBed Amount": Decimal;
        "GREcess Amt": Decimal;
        "GRShe Cess Amt": Decimal;
        "GRAddl.Duty": Decimal;
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        Text000: Label 'Data';
        Text001: Label 'BRT Sales Location wise';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('BRT Sales Location Wise', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"BRT Sales Location Wise", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
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
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        /* ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', '');
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GP Petroleums Limited', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('BRT Sales Location Wise', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Transfer Shipment Header".GETFILTERS, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Branches', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Net Value', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Gross Value', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('ECess Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('SHE Cess Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Addl. Duty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataFooter1()
    begin
        IF Qty + "RNet Value" + "RGross Value" + "RBed Amount" + "REcess Amt" + "RShe Cess Amt" + "RAddl.Duty" <> 0 THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-from Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
            //ExcelBuf.AddColumn(Qty,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(Qty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017

            ExcelBuf.AddColumn("RNet Value", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("RGross Value", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("RBed Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("REcess Amt", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("RShe Cess Amt", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn("RAddl.Duty", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
            GQty += Qty;
            "GRNet Value" += "RNet Value";
            "GRGross Value" += "RGross Value";
            "GRBed Amount" += "RBed Amount";
            "GREcess Amt" += "REcess Amt";
            "GRShe Cess Amt" += "RShe Cess Amt";
            "GRAddl.Duty" += "RAddl.Duty";

        END;
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataFooter2()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(GQty,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(GQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn("GRNet Value", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("GRGross Value", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("GRBed Amount", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("GREcess Amt", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("GRShe Cess Amt", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn("GRAddl.Duty", FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
    end;
}


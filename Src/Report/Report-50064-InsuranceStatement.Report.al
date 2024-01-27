report 50064 "Insurance Statement"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/InsuranceStatement.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Location; Location)
        {
            DataItemTableView = SORTING(Code);
            RequestFilterFields = "Code";
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Location Code" = FIELD(Code);
                DataItemTableView = SORTING("No.", "Sell-to Customer No.")
                                    WHERE(Type = FILTER(Item),
                                          Quantity = FILTER(<> 0));

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    SalesInvHeader1.RESET;
                    SalesInvHeader1.SETRANGE(SalesInvHeader1."No.", "Sales Invoice Line"."Document No.");
                    IF SalesInvHeader1.FINDFIRST THEN BEGIN
                        IF (SalesInvHeader1."Posting Date" < StartDate) OR (SalesInvHeader1."Posting Date" > EndDate) THEN
                            CurrReport.SKIP;
                    END;

                    reciTEMUOM.RESET;
                    reciTEMUOM.SETRANGE(reciTEMUOM."Item No.", "Sales Invoice Line"."No.");
                    IF reciTEMUOM.FINDFIRST THEN
                        LineQty := reciTEMUOM."Qty. per Unit of Measure";

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE(SalesInvHeader."No.", "Sales Invoice Line"."Document No.");
                    IF SalesInvHeader.FINDFIRST THEN;
                    ShiipingAgent.RESET;
                    IF ShiipingAgent.GET(SalesInvHeader."Shipping Agent Code") THEN;

                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line"."Document No.");
                    recSIH.SETFILTER(recSIH."Ship-to Code", '<>%1', '');
                    IF recSIH.FINDFIRST THEN;

                    MakeExcelDataBody;
                    //<<1
                end;
            }
            dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
            {
                DataItemLink = "Transfer-from Code" = FIELD(Code);
                DataItemTableView = SORTING("Item No.")
                                    WHERE(Quantity = FILTER(<> 0));

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    CLEAR(grossamt);
                    TransferHeader1.RESET;
                    TransferHeader1.SETRANGE(TransferHeader1."No.", "Transfer Shipment Line"."Document No.");
                    IF TransferHeader1.FINDFIRST THEN BEGIN
                        IF (TransferHeader1."Posting Date" < StartDate) OR (TransferHeader1."Posting Date" > EndDate) THEN
                            CurrReport.SKIP;
                    END;

                    //EBT Rakshitha
                    reciTEMUOM.RESET;
                    reciTEMUOM.SETRANGE(reciTEMUOM."Item No.", "Transfer Shipment Line"."Item No.");
                    IF reciTEMUOM.FINDFIRST THEN
                        TransferQty := reciTEMUOM."Qty. per Unit of Measure";

                    //EBT Rakshitha

                    grossamt := "Transfer Shipment Line".Amount + 0;//"Transfer Shipment Line"."Excise Amount";
                    TransferHeader.RESET;
                    TransferHeader.SETRANGE(TransferHeader."No.", "Transfer Shipment Line"."Document No.");
                    IF TransferHeader.FINDFIRST THEN;
                    ShiipingAgent1.RESET;
                    IF ShiipingAgent1.GET(TransferHeader."Shipping Agent Code") THEN;

                    IF Location1.GET("Transfer Shipment Line"."Transfer-to Code") THEN;
                    MakeExcelDataBodyForTransfer;
                    //<<1
                end;
            }
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; StartDate)
                {
                }
                field("End Date"; EndDate)
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

        //>>1
        CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1

        //usersetup.GET(USERID);
        MakeExcelInfo;
        //<<1
    end;

    var
        ExcelBuf: Record 370 temporary;
        SalesInvHeader: Record 112;
        TotalQty: Decimal;
        LineQty: Decimal;
        Itemrec: Record 27;
        InvValue: Decimal;
        ShiipingAgent: Record 291;
        ShiipingAgent1: Record 291;
        TransferQty: Decimal;
        TransferHeader: Record 5744;
        Location1: Record 14;
        StartDate: Date;
        EndDate: Date;
        SalesInvHeader1: Record 112;
        TransferHeader1: Record 5744;
        usersetup: Record 91;
        grossamt: Decimal;
        reciTEMUOM: Record 5404;
        Value: Decimal;
        "recShip2Add-SI": Record 222;
        recSIH: Record 112;
        recState: Record State;
        Text000: Label 'Data';
        Text001: Label 'Insurance Statement';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//31May2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Insurance Statement', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Insurance Statement", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // [Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        /*  ExcelBuf.CreateBook('', Text001);
         ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', '');
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
    end;

    //  [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('From Location', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('TO', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('Name of Product', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('Number Of Cases', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('Total Quantity', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('Gross Weight', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('RR / TR Number', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('Value', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('Transport Company Name in Case the Goods are sent by Lorry/tanker', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('Place', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('City', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('Ship To Name', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('Ship To City', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Location.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn(SalesInvHeader."Bill-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn("Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//4
        ExcelBuf.AddColumn(LineQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(SalesInvHeader."LR/RR No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(SalesInvHeader."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//8
        IF SalesInvHeader."Currency Code" = '' THEN BEGIN
            ExcelBuf.AddColumn(/*"Sales Invoice Line"."Amount To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,0.00', 0);//9
        END ELSE BEGIN
            ExcelBuf.AddColumn(/*"Sales Invoice Line"."Amount To Customer"*/0 / SalesInvHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,0.00', 0);//9
        END;
        ExcelBuf.AddColumn(ShiipingAgent.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(SalesInvHeader."Bill-to Address 2", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn(SalesInvHeader."Bill-to City", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn(SalesInvHeader."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn(recSIH."Ship-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn(recSIH."Ship-to City", FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataBodyForTransfer()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Location.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn(Location1.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Transfer Shipment Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Transfer Shipment Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//4
        ExcelBuf.AddColumn(TransferQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(TransferHeader."LR/RR No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(TransferHeader."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//8
        ExcelBuf.AddColumn(grossamt, FALSE, '', FALSE, FALSE, FALSE, '#,0.00', 0);//9
        ExcelBuf.AddColumn(ShiipingAgent1.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn(TransferHeader."Transfer-to Address 2", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn(Location1.City, FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Transfer Shipment Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
    end;
}


report 50142 "ISSUED FORM (Purch. Invoice)"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ISSUEDFORMPurchInvoice.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Purch. Inv. Header"; "Purch. Inv. Header")
        {
            DataItemTableView = SORTING("Pay-to Name", "No.", "Buy-from Vendor No.");
            /* WHERE("Form Code" = FILTER(<> ''),
                  "Form No." = FILTER(<> '')); */
            RequestFilterFields = "Posting Date";
            column(PIH_No; "Purch. Inv. Header"."No.")
            {
            }
            dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
            {
                DataItemLink = "Applies-to Doc. No." = FIELD("No.");
                DataItemTableView = SORTING("No.")
                                    WHERE("Return Order No." = FILTER(<> ''));
                column(PCM_No; "Purch. Cr. Memo Hdr."."No.")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    CrCount := COUNT; //15Feb2017 RB-N
                    //>>1

                    recCrMemLine.RESET;
                    recCrMemLine.SETRANGE("Document No.", "Purch. Cr. Memo Hdr."."No.");
                    //recCrMemLine.CALCSUMS("Tax Amount", "Charges To Vendor", "Tax Base Amount", "Amount To Vendor");
                    //<<1
                end;

                trigger OnPostDataItem()
                begin
                    //>>15Feb2017 RB-N
                    IF CrCount > 0 THEN BEGIN
                        netAmount := netAmountSales - netAmountSalesReturn;
                        Charges := ChargesSales - ChargesSalesReturn;
                        Taxbase := TaxbaseSales - TaxbaseSalesReturn;
                        taxamount := taxamountSales - taxamountSalesReturn;
                        //MESSAGE('ii %1, InC %2',I,InCount);
                        IF InCount = I THEN BEGIN
                            IF PrintToExcel THEN
                                MakeExcelDataFooter;
                        END;

                    END;
                end;
            }

            trigger OnAfterGetRecord()
            begin
                I += 1;//Increment
                InCount := COUNT;//15Feb2017 RB-N, No.of Records

                //>>1

                recPIL.RESET;
                recPIL.SETCURRENTKEY(Type, "Document No.");
                recPIL.SETRANGE(recPIL.Type, recPIL.Type::Item);
                recPIL.SETRANGE("Document No.", "Purch. Inv. Header"."No.");
                //recPIL.CALCSUMS("Tax Amount", "Charges To Vendor", "Tax Base Amount", "Amount To Vendor");

                Vendor.GET("Purch. Inv. Header"."Buy-from Vendor No.");
                CSTNo := '';// Vendor."C.S.T. No.";
                TinNo := '';// Vendor."T.I.N. No.";
                //<<1

                //>>2

                IF "Purch. Inv. Header"."Responsibility Center" <> '' THEN BEGIN
                    rescenter1.RESET;
                    rescenter1.SETRANGE(rescenter1.Code, "Responsibility Center");
                    IF rescenter1.FIND('-') THEN
                        respname := rescenter1.Name;
                END;

                IF PrintToExcel THEN
                    MakeExcelDataBody;
                //>>2
            end;

            trigger OnPostDataItem()
            begin

                //>> 15Feb2017 RB-N
                IF CrCount = 0 THEN BEGIN
                    netAmount := netAmountSales;
                    Charges := ChargesSales;
                    Taxbase := TaxbaseSales;
                    taxamount := taxamountSales;

                    IF InCount = I THEN BEGIN
                        IF PrintToExcel THEN
                            MakeExcelDataFooter;
                    END
                END
            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("Pay-to Name", "No.", "Buy-from Vendor No."); //15Feb2017 RB-N

                //>>1

                IF (cdeselltocustno <> '') THEN
                    SETRANGE("Buy-from Vendor No.", cdeselltocustno);

                IF (cdeformcode <> '') THEN
                    //SETRANGE("Form Code", cdeformcode);

                IF (cderespcentre <> '') THEN BEGIN
                        SETRANGE("Responsibility Center", cderespcentre);
                        respcenter.GET(cderespcentre);
                    END;

                IF (salespersoncode <> '') THEN BEGIN
                    SETRANGE("Purchaser Code", salespersoncode);
                    RecSalesperson.GET(salespersoncode);
                END;

                IF (locationcode <> '') THEN BEGIN
                    SETRANGE("Location Code", locationcode);
                    RecLoc.GET(locationcode);
                END;

                Daterange := GETFILTER("Posting Date");
                locfilter := GETFILTER("Location Code");

                IF locfilter <> '' THEN
                    Filtercriteria := 'Pending Form Location wise :- ' + RecLoc.Name
                ELSE
                    IF cdeformcode <> '' THEN
                        Filtercriteria := 'Pending ' + cdeformcode + ' Forms'
                    ELSE
                        IF cderespcentre <> '' THEN
                            Filtercriteria := 'Pending Form Responsibility wise :- ' + respcenter.Name
                        ELSE
                            IF salespersoncode <> '' THEN
                                Filtercriteria := 'Pending Form salesperson wise :- ' + RecSalesperson.Name
                            ELSE BEGIN
                                IF Vendor.GET(cdeselltocustno) THEN;
                                Filtercriteria := 'Pending Form Party wise :- ' + Vendor.Name;
                            END;

                //<<

                I := 0;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Customer No."; cdeselltocustno)
                {
                    ApplicationArea = ALL;
                    TableRelation = Customer;
                }
                field("Responsiblity Center"; cderespcentre)
                {
                    ApplicationArea = ALL;
                    TableRelation = "Responsibility Center";
                }
                field("Form Code"; cdeformcode)
                {
                    ApplicationArea = ALL;
                    //TableRelation = "Form Codes";
                }
                field("SalesPerson Code"; salespersoncode)
                {
                    ApplicationArea = ALL;
                    TableRelation = "Salesperson/Purchaser";
                }
                field(Location; locationcode)
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
        CompInfo.GET;

        netAmount := 0;
        Charges := 0;
        Taxbase := 0;
        taxamount := 0;

        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<1
    end;

    var
        cdeselltocustno: Code[10];
        cderespcentre: Code[10];
        cdeformcode: Code[10];
        salespersoncode: Code[10];
        recPIH: Record 122;
        recPIL: Record 123;
        netAmount: Decimal;
        Charges: Decimal;
        Partyname: Text[60];
        Vendor: Record 23;
        display1: Boolean;
        Daterange: Text[60];
        CompInfo: Record 79;
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        Check: Report "Check Report";
        SIH: Record 112;
        respcenter: Record 5714;
        Taxbase: Decimal;
        taxamount: Decimal;
        respname: Text[60];
        rescenter1: Record 5714;
        locfilter: Text[60];
        RecLoc: Record 14;
        locationcode: Text[40];
        LocationName: Text[100];
        recCrMemLine: Record 125;
        netAmountSales: Decimal;
        ChargesSales: Decimal;
        TaxbaseSales: Decimal;
        taxamountSales: Decimal;
        netAmountSalesReturn: Decimal;
        ChargesSalesReturn: Decimal;
        TaxbaseSalesReturn: Decimal;
        taxamountSalesReturn: Decimal;
        Filtercriteria: Text[200];
        RecSalesperson: Record 13;
        CSTNo: Code[100];
        TinNo: Code[100];
        Text000: Label 'Data';
        Text001: Label 'Issued Form';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "--15Feb17": Integer;
        CrVisib: Boolean;
        CrCount: Integer;
        InCount: Integer;
        I: Integer;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet; //15Feb2017 RB-N
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn('ISSUED FORM (Purchase Invoice)', FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(REPORT::"ISSUED FORM (Purch. Invoice)", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 0);
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

        //>>15Feb2017 RB-N
        /*  ExcelBuf.CreateBook('', Text001);
         ExcelBuf.CreateBookAndOpenExcel('', 'Issued Form Report', 'Issued Form Report', COMPANYNAME, USERID);
         ExcelBuf.GiveUserControl;
  */
        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Vendor No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Name Of Party', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Responsibility Center', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Bill No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Taxable Amt', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Form No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Pay-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(CSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(TinNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(respname, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*"Purch. Inv. Header"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recPIL."Tax Base Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recPIL."Tax Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recPIL."Charges To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(recPIL."Amount To Vendor", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*"Purch. Inv. Header"."Form No."*/'', FALSE, '', FALSE, FALSE, FALSE, '', 0);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ABS(Taxbase), FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ABS(taxamount), FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(Charges, FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(netAmount, FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataReturn()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Hdr."."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recCrMemLine."Tax Base Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recCrMemLine."Tax Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recCrMemLine."Charges To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recCrMemLine."Amount To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
    end;
}


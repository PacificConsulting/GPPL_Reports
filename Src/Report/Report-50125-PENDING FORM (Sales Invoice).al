report 50125 "PENDING FORM (Sales Invoice)"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PENDINGFORMSalesInvoice.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("Sell-to Customer Name", "No.", "Sell-to Customer No.")
                                ORDER(Ascending)
                               WHERE(/* "Form Code" = FILTER(<> ''),
                                     "Form No." = FILTER(= ''), */
                                      "Cancelled Invoice" = FILTER(false));
            RequestFilterFields = "Posting Date";
            column(SIH_No; "Sales Invoice Header"."No.")
            {
            }
            column(SIH_CustNo; "Sales Invoice Header"."Sell-to Customer No.")
            {
            }
            column(SIH_CustName; "Sales Invoice Header"."Sell-to Customer Name")
            {
            }
            column(respname; respname)
            {
            }
            column(CompName; CompInfo.Name)
            {
            }
            column(Filtercriteria; Filtercriteria)
            {
            }
            column(SIH_PostingDate; "Sales Invoice Header"."Posting Date")
            {
            }
            column(SIH_FormCode; 0)//"Sales Invoice Header"."Form Code")
            {
            }
            column(LineTaxBaseAmt; 0)// recSIL."Tax Base Amount")
            {
            }
            column(LineTaxAmt; 0)// recSIL."Tax Amount")
            {
            }
            column(LineChargeAmt; 0)// recSIL."Charges To Customer")
            {
            }
            column(LineAmt; 0)// recSIL."Amount To Customer")
            {
            }
            column(Daterange; Daterange)
            {
            }
            dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
            {
                //DataItemLink = "Applies-to Doc. No".=FIELD("No.");
                DataItemTableView = SORTING("No.")
                                    WHERE("Return Order No." = FILTER(<> ''));
                column(CrVisib; CrVisib)
                {
                }
                column(Cr_No; "Sales Cr.Memo Header"."No.")
                {
                }
                column(Cr_PostingDate; "Sales Cr.Memo Header"."Posting Date")
                {
                }
                column(Cr_LineTaxBaseAmt; 0)//recCrMemLine."Tax Base Amount")
                {
                }
                column(Cr_LineTaxAmt; 0)// recCrMemLine."Tax Amount")
                {
                }
                column(Cr_LineChargeAmt; 0)//recCrMemLine."Charges To Customer")
                {
                }
                column(Cr_LineAmt; 0)//recCrMemLine."Amount To Customer")
                {
                }
                column(Taxbase; Taxbase)
                {
                }
                column(taxamount; taxamount)
                {
                }
                column(Charges; Charges)
                {
                }
                column(netAmount; netAmount)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    CrCount := COUNT;
                    //>>13Feb2017
                    CLEAR(CrVisib);
                    CrVisib := TRUE;
                    //>>1

                    recCrMemLine.RESET;
                    recCrMemLine.SETRANGE("Document No.", "Sales Cr.Memo Header"."No.");
                    //recCrMemLine.CALCSUMS("Tax Amount", "Charges To Customer", "Tax Base Amount", "Amount To Customer");

                    //<<1
                    //>>2



                    //Sales Cr.Memo Header, GroupFooter (1) - OnPostSection()
                    netAmountSalesReturn := netAmountSalesReturn + 0;// recCrMemLine."Amount To Customer";
                    ChargesSalesReturn := ChargesSalesReturn + 0;// recCrMemLine."Charges To Customer";
                    TaxbaseSalesReturn := TaxbaseSalesReturn + 0;//recCrMemLine."Tax Base Amount";
                    taxamountSalesReturn := taxamountSalesReturn + 0;// recCrMemLine."Tax Amount";

                    IF PrintToExcel THEN
                        MakeExcelDataReturn;
                    //<<2

                    /*
                    //>>3
                    
                    netAmount :=netAmountSales-netAmountSalesReturn;
                    Charges :=ChargesSales-ChargesSalesReturn;
                    Taxbase :=TaxbaseSales-TaxbaseSalesReturn;
                    taxamount :=taxamountSales-taxamountSalesReturn;
                    
                    IF PrintToExcel THEN
                       MakeExcelDataFooter;
                    //<<3
                    */

                end;

                trigger OnPostDataItem()
                begin
                    //>>13Feb2017
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

                InCount := COUNT; //No of Count
                I += 1; //increment
                //>>1


                recSIL.RESET;
                recSIL.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                //recSIL.CALCSUMS("Tax Amount", "Charges To Customer", "Tax Base Amount", "Amount To Customer");

                cust.GET("Sales Invoice Header"."Sell-to Customer No.");
                CSTNo := '';// cust."C.S.T. No.";
                TinNo := ' ';// cust."T.I.N. No.";

                LocName := '';
                Location.RESET;
                Location.SETRANGE(Location.Code, "Sales Invoice Header"."Location Code");
                IF Location.FINDFIRST THEN BEGIN
                    LocName := Location.Name;
                END;

                SalesPersonName := '';
                RecSalesperson.RESET;
                RecSalesperson.SETRANGE(RecSalesperson.Code, "Sales Invoice Header"."Salesperson Code");
                IF RecSalesperson.FINDFIRST THEN BEGIN
                    SalesPersonName := RecSalesperson.Name;
                END;

                recCust.RESET;
                recCust.SETRANGE(recCust."No.", "Sales Invoice Header"."Sell-to Customer No.");
                IF recCust.FINDFIRST THEN BEGIN
                    CustResCen := recCust."Responsibility Center";

                    rescenter1.RESET;
                    rescenter1.SETRANGE(rescenter1.Code, CustResCen);
                    IF rescenter1.FIND('-') THEN
                        respname := rescenter1.Name;
                END;

                // EBT MILAN 25022014------------STATE DETAIL COLUMN---------------------------------------------------START
                RecState.RESET;
                //RecState.SETRANGE(RecState.Code, "Sales Invoice Header"."Bill to Customer State");
                IF RecState.FINDFIRST THEN;
                // EBT MILAN 25022014------------STATE DETAIL COLUMN-----------------------------------------------------END



                //<<1

                //>>2 Group Total
                netAmountSales := netAmountSales + 0;// recSIL."Amount To Customer";
                ChargesSales := ChargesSales + 0;// recSIL."Charges To Customer";
                TaxbaseSales := TaxbaseSales + 0;//recSIL."Tax Base Amount";
                taxamountSales := taxamountSales + 0;//recSIL."Tax Amount";

                //<<2

                IF PrintToExcel THEN
                    MakeExcelDataBody;
            end;

            trigger OnPostDataItem()
            begin
                //>> 13Feb2017
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

                //>>1
                usersetup.GET(USERID);

                IF (cdeselltocustno <> '') THEN
                    SETRANGE("Sell-to Customer No.", cdeselltocustno);

                IF (cdeformcode <> '') THEN
                    //SETRANGE("Form Code", cdeformcode);

                IF (cderespcentre <> '') THEN BEGIN
                        recCustomer.RESET;
                        recCustomer.SETRANGE(recCustomer."Responsibility Center", cderespcentre);
                        IF recCustomer.FINDFIRST THEN
                            SETRANGE("Shortcut Dimension 2 Code", recCustomer."Responsibility Center");
                        respcenter.GET(cderespcentre);
                    END;

                IF (salespersoncode <> '') THEN BEGIN
                    SETRANGE("Salesperson Code", salespersoncode);
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
                                IF cust.GET(cdeselltocustno) THEN;
                                Filtercriteria := 'Pending Form Party wise :- ' + cust.Name;
                            END;

                //<<1

                //IF PrintToExcel THEN
                //MakeExcelDataHeader;

                IF PrintToExcel THEN
                    MakeExcelInfo;

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
                    ApplicationArea = all;
                    TableRelation = Customer;
                }
                field("Responsiblity Center"; cderespcentre)
                {
                    ApplicationArea = all;
                    TableRelation = "Responsibility Center";
                }
                field("Form Code"; cdeformcode)
                {
                    ApplicationArea = all;
                    //TableRelation = "Form Codes";
                }
                field("SalesPerson Code"; salespersoncode)
                {
                    ApplicationArea = all;
                    TableRelation = "Salesperson/Purchaser";
                }
                field(Location; locationcode)
                {
                    ApplicationArea = all;
                    TableRelation = Location;
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = all;
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

        /*
        //EBT STIVAN ---(14122012)--- To Filter Report as per the Report View Role ------START
        //MemberOf_Table NotFound
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (MemberOf.FINDFIRST) THEN
        BEGIN
          IF (cderespcentre = '') AND (locationcode = '')THEN
          ERROR('Please Specify Res. Centre Filter OR Location Filter');
        
         CSOmapping.RESET;
         CSOmapping.SETRANGE(CSOmapping."User Id",UPPERCASE(USERID));
         CSOmapping.SETRANGE(Type,CSOmapping.Type::Location);
         CSOmapping.SETRANGE(CSOmapping.Value,locationcode);
         IF NOT(CSOmapping.FINDFIRST) THEN
         BEGIN
          CSOmapping.RESET;
          CSOmapping.SETRANGE(CSOmapping."User Id",UPPERCASE(USERID));
          CSOmapping.SETRANGE(CSOmapping.Type,CSOmapping.Type::"Responsibility Center");
          CSOmapping.SETFILTER(CSOmapping.Value,'%1',cderespcentre);
          IF NOT(CSOmapping.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        //EBT STIVAN ---(14122012)--- To Filter Report as per the Report View Role --------END
        */
        ExcelBuf.RESET;
        ExcelBuf.DELETEALL;

        usersetup.GET(USERID);

        CompInfo.GET;

        netAmount := 0;
        Charges := 0;
        Taxbase := 0;
        taxamount := 0;

    end;

    var
        cdeselltocustno: Code[10];
        cderespcentre: Code[10];
        cdeformcode: Code[10];
        salespersoncode: Code[10];
        recSIH: Record 112;
        recSIL: Record 113;
        netAmount: Decimal;
        Charges: Decimal;
        Partyname: Text[60];
        cust: Record 18;
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
        recCrMemLine: Record 115;
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
        usersetup: Record 91;
        LocRec: Record 14;
        RespCentre: Code[10];
        Location: Record 14;
        LocName: Text[30];
        recCust: Record 18;
        CustResCen: Code[10];
        recCustomer: Record 18;
        SalesPersonName: Text[30];
        CSOmapping: Record 50006;
        RecState: Record State;
        State: Text[30];
        RecSalesState: Record State;
        RecShipadd: Record 222;
        SaleState: Text[1024];
        Text000: Label 'Data';
        Text001: Label 'Pending Form';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "--13Feb2017": Integer;
        CrVisib: Boolean;
        CrCount: Integer;
        InCount: Integer;
        I: Integer;

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin

        ExcelBuf.SetUseInfoSheet;
        //ExcelBuf.AddInfoColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn('PENDING FORM (Sales Invoice)', FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(REPORT::"PENDING FORM (Sales Invoice)", FALSE, FALSE, FALSE, FALSE, '', 0);
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
        //ExcelBuf.CreateBook(Text000,Text001);
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');
        /*  ExcelBuf.CreateBook('', Text001);
         //ExcelBuf.CreateBookAndOpenExcel('','Pending Form Report','Pending Form Report',COMPANYNAME,USERID);
         ExcelBuf.CreateBookAndOpenExcel('', 'Pending Form Report', '', '', USERID); //17Feb2017 RB-N
         ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet('Pending Form Report', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Pending Form Report', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType) 2016
        //ExcelBuf.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat) 2009
        //>>13Feb2017 To Display UserId, Filter ,Date
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn(Filtercriteria, FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('Date Range :' + Daterange, FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        //<<
        ExcelBuf.AddColumn('Customer No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
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
        ExcelBuf.AddColumn('Supply Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Supply Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('SalesPerson Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('SalesPerson Name', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Customer State', FALSE, '', TRUE, FALSE, TRUE, '@', 0);   // EBT MILAN 25022014
        ExcelBuf.AddColumn('Ship-to-State', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer Name", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(CSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(TinNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(respname, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Sales Invoice Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Sales Invoice Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*"Sales Invoice Header"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recSIL."Tax Base Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recSIL."Tax Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recSIL."Charges To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recSIL."Amount To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Sales Invoice Header"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(LocName, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Sales Invoice Header"."Salesperson Code", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(SalesPersonName, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(RecState.Description, FALSE, '', FALSE, FALSE, FALSE, '', 0);  // EBT MILAN 25022014

        //mayuri
        SaleState := '';
        RecSalesState.RESET;
        IF "Sales Invoice Header"."Ship-to Code" <> '' THEN
            IF RecShipadd.GET("Sales Invoice Header"."Sell-to Customer No.", "Sales Invoice Header"."Ship-to Code") THEN
                IF RecSalesState.GET(RecShipadd."State Code") THEN
                    SaleState := RecSalesState.Description;
        //mayuri

        ExcelBuf.AddColumn(SaleState, FALSE, '', FALSE, FALSE, FALSE, '', 0);
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);   // EBT MILAN 25022014
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);   // EBT MILAN 25022014
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
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*"Sales Invoice Header"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recCrMemLine."Tax Base Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recCrMemLine."Tax Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recCrMemLine."Charges To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recCrMemLine."Amount To Customer"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0); // EBT MILAN 25022014
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
    end;
}


report 50062 "Sales Cr.Memo Summary Loc Wise"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesCrMemoSummaryLocWise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
        {
            DataItemTableView = SORTING("Location Code")
                                WHERE("Source Code" = FILTER('SALES'),
                                      "Return Order No." = FILTER(''));
            RequestFilterFields = "Posting Date", "Location Code";
            column(CompName; CompInfo.Name)
            {
            }
            column(HDate; FORMAT(TODAY, 0, 4))
            {
            }
            column(DocNo; "Sales Cr.Memo Header"."No.")
            {
            }
            column(PostingDate; "Sales Cr.Memo Header"."Posting Date")
            {
            }
            column(DateFilter; "Sales Cr.Memo Header".GETFILTER("Sales Cr.Memo Header"."Posting Date"))
            {
            }
            column(HLocCode; "Location Code")
            {
            }
            dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Location Code", "No.", "Sell-to Customer No.")
                                    WHERE("No." = FILTER(<> 74012210 | ''));
                column(LDocNo; "Sales Cr.Memo Line"."Document No.")
                {
                }
                column(LineNo; "Sales Cr.Memo Line"."Line No.")
                {
                }
                column(ItmNo; "Sales Cr.Memo Line"."No.")
                {
                }
                column(LLocCode; "Sales Cr.Memo Line"."Location Code")
                {
                }
                column(QtyBase; "Sales Cr.Memo Line"."Quantity (Base)")
                {
                }
                column(LocName; LocName)
                {
                }
                column(BillAmount; BillAmount)
                {
                }
                column(NetAmt; NetAmt)
                {
                }
                column(TaxAmt; TaxAmt)
                {
                }
                column(ExcDuty; ExcDuty)
                {
                }
                column(Foc; Foc)
                {
                }
                column(Charges; Charges)
                {
                }
                column(EntryTax; EntryTax)
                {
                }
                column(Addduty; Addduty)
                {
                }
                column(Cess; Cess)
                {
                }
                column(FreightAmt; FreightAmt)
                {
                }
                column(CashDiscAmt; CashDiscAmt)
                {
                }
                column(InvDiscAmt; InvDiscAmt)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    LocRec.RESET;
                    LocRec.SETRANGE(LocRec.Code, "Sales Cr.Memo Line"."Location Code");
                    IF LocRec.FINDFIRST THEN
                        LocName := LocRec.Name;
                    LocCode := LocRec."Global Dimension 2 Code";

                    decQty := decQty + "Sales Cr.Memo Line"."Quantity (Base)";
                    NetAmt := NetAmt + "Sales Cr.Memo Line"."Line Amount";
                    TaxAmt := TaxAmt + 0;//"Sales Cr.Memo Line"."Tax Amount";

                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::"Other Taxes");
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Code", '%1|%2', 'C1', 'CESS');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT */
                    Cess := Cess + 0;//PostedStrOrderRec.Amount;
                                     // UNTIL PostedStrOrderRec.NEXT = 0;

                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::"Other Taxes");
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Code", 'ADDL TAX');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT */
                    Addduty := Addduty + 0;//PostedStrOrderRec.Amount;
                                           // UNTIL PostedStrOrderRec.NEXT = 0;

                    BillAmount := NetAmt + TaxAmt + Addduty + Cess;

                    //<<1
                end;

                trigger OnPostDataItem()
                begin
                    //>> GroupFooter
                    IF SCOUNT = TCOUNT THEN BEGIN
                        MakeExcelGroupFooter;
                        TCOUNT := 0;

                    END;
                    //<< GroupFooter
                end;
            }

            trigger OnAfterGetRecord()
            begin

                TCOUNT += 1;//31Mar2017

                //>>GroupData //31Mar2017
                SCrMemo31M.SETCURRENTKEY("Location Code");
                SCrMemo31M.SETRANGE("Location Code", "Location Code");
                IF SCrMemo31M.FINDFIRST THEN BEGIN
                    SCOUNT := SCrMemo31M.COUNT;

                END;

                //<<GroupData //31Mar2017

                //>>1

                //Sales Cr.Memo Header, GroupFooter (3) - OnPreSection()
                //CurrReport.SHOWOUTPUT:= CurrReport.TOTALSCAUSEDBY=FIELDNO("Location Code");

                //IF Printtoexcel THEN
                //MakeExcelGroupFooter;
                //<<1
            end;

            trigger OnPostDataItem()
            begin

                //>>1
                //Sales Cr.Memo Header, Footer (4) - OnPreSection()
                //IF Printtoexcel THEN
                Makeexcelfooter;
                //<<1
            end;

            trigger OnPreDataItem()
            begin
                CompInfo.GET;//31Mar2017
                SCrMemo31M.COPYFILTERS("Sales Cr.Memo Header");//31Mar2017

                TCOUNT := 0;//31Mar2017
            end;
        }
    }

    requestpage
    {

        layout
        {
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
        /*
        //EBT STIVAN ---(21082012)--- To Filter Report as per the Report View Role ------START
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (MemberOf.FINDFIRST) THEN
        ERROR('You do not have permission to run the report');
        //EBT STIVAN ---(21082012)--- To Filter Report as per the Report View Role --------END
        
        Printtoexcel := TRUE;
        */
        CLEAR(ExcelBuf);
        //IF Printtoexcel THEN //31Mar2017
        MakeExcelInfo;
        //<<1

    end;

    var
        //PostedStrOrderRec: Record "13798";
        Foc: Decimal;
        EntryTax: Decimal;
        Charges: Decimal;
        AMt: Decimal;
        decQty: Decimal;
        BillAmount: Decimal;
        NetAmt: Decimal;
        ExcDuty: Decimal;
        TaxAmt: Decimal;
        InvDiscAmt: Decimal;
        TradeDisc: Decimal;
        TradeSubsidy: Decimal;
        ExcelBuf: Record 370 temporary;
        Printtoexcel: Boolean;
        gross: Decimal;
        Datefilter: Text[50];
        LocRec: Record 14;
        LocName: Text[50];
        LocCode: Code[10];
        CashDiscAmt: Decimal;
        recitem: Record 27;
        Addduty: Decimal;
        Cess: Decimal;
        SIH: Record 112;
        recSalesCrMemoHeader: Record 114;
        SIL: Record 113;
        recItem1: Record 27;
        recItemUOM: Record 5404;
        FreightAmt: Decimal;
        UPricePerLT: Decimal;
        UPrice: Decimal;
        SalesPrice: Record 7002;
        recGLentry: Record 17;
        Text16500: Label 'Sales Credit Memo Summary Location Wise';
        Text000: Label 'Data';
        Text001: Label 'Sales Credit Memo Summary Location Wise';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "---31Mar2017": Integer;
        SCrMemo31M: Record 114;
        TCOUNT: Integer;
        SCOUNT: Integer;
        NAH: Integer;
        LCOUNT: Integer;
        TdecQty: Decimal;
        TNetAmt: Decimal;
        TTaxAmt: Decimal;
        TAddduty: Decimal;
        TCess: Decimal;
        TBillAmount: Decimal;
        CompInfo: Record 79;

    //  [Scope('Internal')]
    procedure MakeExcelInfo()
    begin

        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//31Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Sales Credit Memo Summary Location Wise', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Sales Cr.Memo Summary Loc Wise", FALSE, FALSE, FALSE, FALSE, '', 0);
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

        //>>31Mar2017 RB-N
        /* ExcelBuf.CreateBook('', Text001);
        ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>31Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('SALES CREDIT MEMO SUMMARY LOCATION-WISE', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        //<<
        ExcelBuf.AddColumn('Date Filter : ' + "Sales Cr.Memo Header".GETFILTER("Posting Date"), FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Line Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Addl. Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Gross Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.NewRow;
    end;

    // [Scope('Internal')]
    procedure MakeExcelGroupFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(LocRec.Name, FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn(LocCode, FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn(decQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//3
        TdecQty += decQty;//31Mar2017
        decQty := 0;//31Mar2017

        ExcelBuf.AddColumn(NetAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//4
        TNetAmt += NetAmt;//31Mar2017
        NetAmt := 0;//31Mar2017

        ExcelBuf.AddColumn(TaxAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//5
        TTaxAmt += TaxAmt;//31Mar2017
        TaxAmt := 0;//31Mar2017

        ExcelBuf.AddColumn(Addduty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//6
        TAddduty += Addduty;//31Mar2017
        Addduty := 0;//31Mar2017

        ExcelBuf.AddColumn(Cess, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//7
        TCess += Cess;//31Mar2017
        Cess := 0;//31Mar2017

        ExcelBuf.AddColumn(BillAmount, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//8
        TBillAmount += BillAmount;//31Mar2017
        BillAmount := 0;//31Mar2017
    end;

    // [Scope('Internal')]
    procedure Makeexcelfooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(TdecQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//3 //31Mar2017
        ExcelBuf.AddColumn(TNetAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//4 //31Mar2017
        ExcelBuf.AddColumn(TTaxAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//5 //31Mar2017
        ExcelBuf.AddColumn(TAddduty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//6 //31Mar2017
        ExcelBuf.AddColumn(TCess, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//7 //31Mar2017
        ExcelBuf.AddColumn(TBillAmount, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//8 //31Mar2017
    end;
}


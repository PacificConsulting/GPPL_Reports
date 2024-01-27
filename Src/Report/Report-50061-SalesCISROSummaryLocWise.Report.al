report 50061 "Sales CI/SRO Summary Loc Wise"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesCISROSummaryLocWise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
        {
            DataItemTableView = SORTING("Location Code")
                                WHERE("Source Code" = FILTER('SALES'),
                                      "Return Order No." = FILTER(<> ''));
            RequestFilterFields = "Posting Date", "Location Code";
            column(CompName; CompInfo.Name)
            {
            }
            column(label; label)
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
            column(LocFilter; LocFilter)
            {
            }
            column(HLocCode; "Location Code")
            {
            }
            dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Location Code", "No.", "Sell-to Customer No.")
                                    WHERE(Type = FILTER(<> "G/L Account"));
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

                    LCOUNT += 1;//30Mar2017



                    //>>1

                    LocRec.RESET;
                    //LocRec.SETRANGE(LocRec.Code,"Sales Cr.Memo Line"."Location Code");
                    LocRec.SETRANGE(LocRec.Code, "Sales Cr.Memo Header"."Location Code");
                    IF LocRec.FINDFIRST THEN BEGIN
                        LocName := LocRec.Name;
                        LocCode := LocRec."Global Dimension 2 Code";
                    END;

                    //decQty := decQty + "Sales Cr.Memo Line"."Quantity (Base)";
                    CLEAR(decQty);
                    decQty := "Sales Cr.Memo Line"."Quantity (Base)";//30Mar2017

                    GdecQty += decQty;//GroupQty 30Mar2017

                    CLEAR(NetAmt);
                    CLEAR(TaxAmt);
                    CLEAR(ExcDuty);

                    IF "Sales Cr.Memo Header"."Currency Code" = '' THEN BEGIN
                        NetAmt := "Sales Cr.Memo Line"."Line Amount";
                        TaxAmt := 0;// "Sales Cr.Memo Line"."Tax Amount";
                        ExcDuty := 0;//"Sales Cr.Memo Line"."Excise Amount";
                    END ELSE BEGIN
                        NetAmt := "Sales Cr.Memo Line"."Line Amount" / "Sales Cr.Memo Header"."Currency Factor";
                        TaxAmt := 0;// "Sales Cr.Memo Line"."Tax Amount"/"Sales Cr.Memo Header"."Currency Factor";
                        ExcDuty := 0;// "Sales Cr.Memo Line"."Excise Amount"/"Sales Cr.Memo Header"."Currency Factor";
                    END;

                    GNetAmt += NetAmt;//GrossNetAmount 30Mar2017
                    GExcDuty += ExcDuty;//GrossExcDuty 30Mar2017
                    GTaxAmt += TaxAmt;//GrossTaxAmt 30Mar2017

                    CLEAR(Foc);
                    /* PostedStrOrdrRec.RESET;
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Invoice No.", "Document No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Line No.", "Line No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Type", PostedStrOrdrRec."Tax/Charge Type"::Charges);
                    PostedStrOrdrRec.SETFILTER(PostedStrOrdrRec."Tax/Charge Group", '%1|%2', 'FOC', 'FOC DIS');
                    IF PostedStrOrdrRec.FINDSET THEN
                        REPEAT */
                    //  Foc += PostedStrOrdrRec.Amount;//30Mar2017
                    // UNTIL PostedStrOrdrRec.NEXT = 0;

                    //GFoc += Foc;//GrossFoc 30Mar2017

                    CLEAR(EntryTax);
                    /* PostedStrOrdrRec.RESET;
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Invoice No.", "Document No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Line No.", "Line No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Type", PostedStrOrdrRec."Tax/Charge Type"::Charges);
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Group", 'ENTRYTAX');
                    IF PostedStrOrdrRec.FINDSET THEN
                        REPEAT */
                    EntryTax += 0;// PostedStrOrdrRec.Amount;//30Mar2017
                                  //UNTIL PostedStrOrdrRec.NEXT = 0;

                    GEntryTax += EntryTax;//GrossEntryTax 30Mar2017

                    CLEAR(TradeSubsidy);
                    /* PostedStrOrdrRec.RESET;
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Invoice No.", "Document No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Line No.", "Line No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Type", PostedStrOrdrRec."Tax/Charge Type"::Charges);
                    PostedStrOrdrRec.SETFILTER(PostedStrOrdrRec."Tax/Charge Group", '%1|%2', 'TRANSSUBS', 'TPTSUBSIDY');
                    IF PostedStrOrdrRec.FINDSET THEN
                        REPEAT */
                    TradeSubsidy += 0;//PostedStrOrdrRec.Amount;//30Mar2017
                                      //UNTIL PostedStrOrdrRec.NEXT = 0;

                    GTradeSubsidy += TradeSubsidy;//GrossTradeSubsidy 30Mar2017

                    CLEAR(Cess);
                    /* PostedStrOrdrRec.RESET;
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Invoice No.", "Document No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Line No.", "Line No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Type", PostedStrOrdrRec."Tax/Charge Type"::"Other Taxes");
                    PostedStrOrdrRec.SETFILTER(PostedStrOrdrRec."Tax/Charge Code", '%1|%2', 'C1', 'CESS');
                    IF PostedStrOrdrRec.FINDSET THEN
                        REPEAT */
                    Cess += 0;// PostedStrOrdrRec.Amount;//30Mar2017
                              // UNTIL PostedStrOrdrRec.NEXT = 0;

                    GCess += Cess;//GrossCess 30Mar2017

                    CLEAR(Addduty);
                    /* PostedStrOrdrRec.RESET;
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Invoice No.", "Document No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Line No.", "Line No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Type", PostedStrOrdrRec."Tax/Charge Type"::"Other Taxes");
                    PostedStrOrdrRec.SETFILTER(PostedStrOrdrRec."Tax/Charge Code", 'ADDL TAX');
                    IF PostedStrOrdrRec.FINDSET THEN */
                    //REPEAT
                    Addduty += 0;// PostedStrOrdrRec.Amount;//30Mar2017
                                 // UNTIL PostedStrOrdrRec.NEXT = 0;

                    GAddduty += Addduty;//GrossAddduty 30Mar2017


                    CLEAR(CashDiscAmt);
                    /* PostedStrOrdrRec.RESET;
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Invoice No.", "Document No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Line No.", "Line No.");
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Type", PostedStrOrdrRec."Tax/Charge Type"::Charges);
                    PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Group", 'CASHDISC');
                    IF PostedStrOrdrRec.FINDSET THEN
                        REPEAT */
                    CashDiscAmt += 0;// PostedStrOrdrRec.Amount;//30Mar2017
                                     //  UNTIL PostedStrOrdrRec.NEXT = 0;

                    GCashDiscAmt += CashDiscAmt;//GrossCashDiscAmt 30Mar2017

                    CLEAR(FreightAmt);
                    recGLentry.RESET;
                    recGLentry.SETFILTER(recGLentry."G/L Account No.", '%1|%2', '75001010', '75001025');
                    recGLentry.SETRANGE(recGLentry."Document No.", "Sales Cr.Memo Header"."No.");
                    IF recGLentry.FINDSET THEN
                        REPEAT
                            FreightAmt += recGLentry.Amount;
                        UNTIL recGLentry.NEXT = 0;

                    GFreightAmt += FreightAmt;//GrossFreightAmt 30Mar2017

                    CLEAR(Charges);
                    /*  PostedStrOrdrRec.RESET;
                     PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Invoice No.", "Document No.");
                     PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Line No.", "Line No.");
                     PostedStrOrdrRec.SETRANGE(PostedStrOrdrRec."Tax/Charge Type", PostedStrOrdrRec."Tax/Charge Type"::Charges);
                     PostedStrOrdrRec.SETFILTER("Tax/Charge Group", '%1|%2|%3|%4|%5|%6|%7|%8', 'ADD-DIS', 'AVP-SP-SAN',
                                                'MONTH-DIS', 'NEW-PRI-SP', 'SPOT-DISC', 'VAT-DIF-DI', 'T2 DISC', 'STATE DISC');
                     IF PostedStrOrdrRec.FINDSET THEN */
                    //   REPEAT
                    Charges := Charges + 0;// PostedStrOrdrRec.Amount;
                                           // UNTIL PostedStrOrdrRec.NEXT = 0;

                    GCharges += Charges;//GrossCharges 30Mar2017

                    CLEAR(BillAmount);//30Mar2017
                    BillAmount := /*NetAmt + TaxAmt + ExcDuty + Addduty + Cess + Foc*/0 + Charges + EntryTax + TradeSubsidy + FreightAmt + CashDiscAmt;
                    //<<1

                    GBillAmount += BillAmount;//GroupBill 30Mar2017


                    //>>30Mar2017 GroupData
                    IF (TCOUNT = SCOUNT) AND Printtoexcel AND (NAH = LCOUNT) THEN BEGIN
                        //MESSAGE('Location Name %1 \\ Loc cOde %2 \ SCount %3 \ TCount %4',LocName,"Sales Cr.Memo Header"."Location Code",SCOUNT,TCOUNT);
                        MakeExcelGroupFooter;
                        TCOUNT := 0;

                    END;

                    //<<30Mar2017 GroupData

                    /*
                    IF TCOUNT = SCOUNT THEN
                    BEGIN
                      MESSAGE('GRoup Location Code %1',"Location Code");
                      TCOUNT := 0;
                    END;
                    *///TestCode

                end;

                trigger OnPreDataItem()
                begin

                    //>>1
                    //"Sales Cr.Memo Line".SETRANGE("Sales Cr.Memo Line"."Document No.","Sales Cr.Memo Header"."No.");
                    //<<1

                    NAH := COUNT;//30Mar2017

                    LCOUNT := 0;//30Mar2017
                end;
            }

            trigger OnAfterGetRecord()
            begin

                TCOUNT += 1;//30Mar2017


                //>>30Mar2017 GroupbyLocation

                SCrMemo30M.SETCURRENTKEY("Location Code");
                SCrMemo30M.SETRANGE("Location Code", "Location Code");
                IF SCrMemo30M.FINDFIRST THEN BEGIN
                    SCOUNT := SCrMemo30M.COUNT;
                    //MESSAGE('SCOUNT %1 \\ TCOUNT %2 \\ LOcaname %3 \\ DocNo %4',SCOUNT,TCOUNT,"Location Code","No.");
                END;


                //<<30Mar2017 GroupbyLocation

                //>>1

                //Sales Cr.Memo Header, Header (1) - OnPreSection()

                IF (CancelInvoice = TRUE) AND (SalesOrderReturn = FALSE) THEN
                    label := 'Cancel Invoice Summary'
                ELSE
                    IF (CancelInvoice = FALSE) AND (SalesOrderReturn = TRUE) THEN
                        label := 'SRO Summary'
                    ELSE
                        IF (CancelInvoice = TRUE) AND (SalesOrderReturn = TRUE) THEN
                            label := 'Cancel Invoice / SRO Summary'
                        ELSE
                            IF (CancelInvoice = FALSE) AND (SalesOrderReturn = FALSE) THEN
                                label := '';
                //<<1


                //>>2

                //Sales Cr.Memo Header, GroupHeader (2) - OnPreSection()
                //CurrReport.SHOWOUTPUT:= CurrReport.TOTALSCAUSEDBY=FIELDNO("Location Code");
                /*
                IF CurrReport.SHOWOUTPUT=TRUE THEN BEGIN
                 decQty:=0;
                 BillAmount:=0;
                 NetAmt:=0;
                 ExcDuty:=0;
                 TaxAmt:=0;
                 InvDiscAmt:=0;
                 TradeDisc:=0;
                 TradeSubsidy:=0;
                 CashDiscAmt:=0;
                 EntryTax:=0;
                 Charges:=0;
                 Foc:=0;
                 //CLEAR(FOC);//30Mar2017
                 Cess:=0;
                 Addduty:=0;
                 FreightAmt := 0;
                END;
                *///Commented 30Mar2017
                  //<<2


                //>>3

                //Sales Cr.Memo Header, GroupFooter (3) - OnPreSection()
                //CurrReport.SHOWOUTPUT:= CurrReport.TOTALSCAUSEDBY=FIELDNO("Location Code");

                //IF Printtoexcel THEN
                //MakeExcelGroupFooter;
                //<<3

            end;

            trigger OnPostDataItem()
            begin

                //>>1

                //Sales Cr.Memo Header, Footer (4) - OnPreSection()
                IF Printtoexcel THEN
                    MakeExcelGroupFooter();
                //<<1
            end;

            trigger OnPreDataItem()
            begin

                LocFilter := "Sales Cr.Memo Header".GETFILTER("Sales Cr.Memo Header"."Location Code");
                //>>1

                IF (CancelInvoice = FALSE) AND (SalesOrderReturn = FALSE) THEN
                    ERROR('You have not selected any option');

                IF (CancelInvoice = TRUE) AND (SalesOrderReturn = TRUE) THEN BEGIN
                END ELSE BEGIN
                    IF CancelInvoice = TRUE THEN
                        "Sales Cr.Memo Header".SETFILTER("Sales Cr.Memo Header"."Cancelled Invoice", 'YES');

                    IF SalesOrderReturn = TRUE THEN
                        "Sales Cr.Memo Header".SETFILTER("Sales Cr.Memo Header"."Cancelled Invoice", 'NO');
                END;

                /*
                CurrReport.CREATETOTALS(decQty);
                CurrReport.CREATETOTALS(BillAmount);
                CurrReport.CREATETOTALS(NetAmt);
                CurrReport.CREATETOTALS(ExcDuty);
                CurrReport.CREATETOTALS(TaxAmt);
                CurrReport.CREATETOTALS(InvDiscAmt);
                CurrReport.CREATETOTALS(TradeDisc);
                CurrReport.CREATETOTALS(TradeSubsidy);
                CurrReport.CREATETOTALS(CashDiscAmt);
                CurrReport.CREATETOTALS(Foc);
                CurrReport.CREATETOTALS(EntryTax);
                CurrReport.CREATETOTALS(Charges);
                CurrReport.CREATETOTALS(Addduty);
                CurrReport.CREATETOTALS(Cess);
                CurrReport.CREATETOTALS(FreightAmt);
                *///Commented 30Mar2017
                  //<<1

                //>>30Mar2017 CopyRecord
                SETFILTER("Sales Cr.Memo Header"."Location Code", '<>%1', '');

                SCrMemo30M.COPYFILTERS("Sales Cr.Memo Header");
                //<<30Mar2017 CopyRecord

                //TCOUNT := 0;//30Mar2017

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print To Excel"; Printtoexcel)
                {
                }
                field("Cancel Invoice"; CancelInvoice)
                {
                }
                field("Sales Return Order"; SalesOrderReturn)
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

        IF Printtoexcel THEN
            CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1

        CompInfo.GET;

        //EBT STIVAN ---(21082012)--- To Filter Report as per the Report View Role ------START
        /*
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (MemberOf.FINDFIRST) THEN
        ERROR('You do not have permission to run the report');
        *///Commented 30Mar2017

        //EBT STIVAN ---(21082012)--- To Filter Report as per the Report View Role --------END

        IF SalesCrMemoHeader."Location Code" <> '' THEN
            recLoc.RESET;
        recLoc.SETRANGE(recLoc.Code, SalesCrMemoHeader.GETFILTER("Location Code"));
        IF recLoc.FINDFIRST THEN
            CLEAR(ExcelBuf);
        IF Printtoexcel THEN
            MakeExcelInfo;
        //<<1

    end;

    var
        Text16500: Label 'As per Details';
        Text000: Label 'Data';
        Text001: Label 'Sales Cancel Invoice / Sales Return Order Summary Location Wise';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        //  PostedStrOrdrRec: Record "13798";
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
        CancelInvoice: Boolean;
        SalesOrderReturn: Boolean;
        label: Text[50];
        CompInfo: Record 79;
        SalesCrMemoHeader: Record 114;
        recLoc: Record 14;
        "---30Mar2017": Integer;
        SCrMemo30M: Record 114;
        TCOUNT: Integer;
        SCOUNT: Integer;
        NAH: Integer;
        LCOUNT: Integer;
        GdecQty: Decimal;
        GBillAmount: Decimal;
        GNetAmt: Decimal;
        GExcDuty: Decimal;
        GTaxAmt: Decimal;
        GInvDiscAmt: Decimal;
        GCashDiscAmt: Decimal;
        GFreightAmt: Decimal;
        GTradeSubsidy: Decimal;
        GFoc: Decimal;
        GCharges: Decimal;
        GEntryTax: Decimal;
        GAddduty: Decimal;
        GCess: Decimal;
        TdecQty: Decimal;
        TBillAmount: Decimal;
        TNetAmt: Decimal;
        TExcDuty: Decimal;
        TTaxAmt: Decimal;
        TInvDiscAmt: Decimal;
        TCashDiscAmt: Decimal;
        TFreightAmt: Decimal;
        TTradeSubsidy: Decimal;
        TFoc: Decimal;
        TCharges: Decimal;
        TEntryTax: Decimal;
        TAddduty: Decimal;
        TCess: Decimal;
        LocFilter: Text[100];

    //  [Scope('Internal')]
    procedure MakeExcelInfo()
    begin

        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//30Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(CompInfo.Name, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Sales Cancel Invoice / Sales Return Order Summary Location Wise', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Sales CI/SRO Summary Loc Wise", FALSE, FALSE, FALSE, FALSE, '', 0);
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
        //ExcelBuf.CreateSheet(Text000,Text001,CompInfo.Name,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>30Mar2017 RB-N
        /*  ExcelBuf.CreateBook('', 'Sales CI/SRO Summary LocWise');
         ExcelBuf.CreateBookAndOpenExcel('', 'Sales CI/SRO Summary LocWise', '', '', USERID);
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook('Sales CI/SRO Summary LocWise');
        ExcelBuf.WriteSheet('Sales CI/SRO Summary LocWise', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Sales CI/SRO Summary LocWise', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //ExcelBuf.AddColumn(CompInfo.Name,FALSE,'',TRUE,FALSE,FALSE,'@',1);
        //ExcelBuf.NewRow;
        //>>30Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16

        ExcelBuf.NewRow;
        IF (CancelInvoice = TRUE) AND (SalesOrderReturn = FALSE) THEN
            ExcelBuf.AddColumn('Cancel Invoice Summary', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1

        IF (CancelInvoice = FALSE) AND (SalesOrderReturn = TRUE) THEN
            ExcelBuf.AddColumn('SRO Summary', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1

        IF (CancelInvoice = TRUE) AND (SalesOrderReturn = TRUE) THEN
            ExcelBuf.AddColumn('Cancel Invoice / SRO Summary', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1

        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.NewRow;



        //<<30Mar2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter :', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn("Sales Cr.Memo Header".GETFILTER("Posting Date"), FALSE, '', TRUE, FALSE, FALSE, '@', 1);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location Filter :', FALSE, '', TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(recLoc.Name, FALSE, '', TRUE, FALSE, FALSE, '@', 1);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Quantity in Ltrs/Kg.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Gross Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Net Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Excise Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Invoice Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Cash Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Freight Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Transport Subsidy', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('FOC', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Entry Tax', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Add.Tax', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16

        ExcelBuf.NewRow;
    end;

    //[Scope('Internal')]
    procedure MakeExcelGroupFooter()
    begin
        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn(LocRec.Name,FALSE,'',TRUE,FALSE,TRUE,'',1);//1
        ExcelBuf.AddColumn(LocName, FALSE, '', TRUE, FALSE, TRUE, '', 1);//1 //30Mar2017
        ExcelBuf.AddColumn(LocCode, FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        //ExcelBuf.AddColumn(decQty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//3
        ExcelBuf.AddColumn(GdecQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//3 //GroupQty 30Mar2017
        TdecQty += GdecQty;//
        GdecQty := 0;//GroupQty 30Mar2017

        //ExcelBuf.AddColumn(BillAmount,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//4
        ExcelBuf.AddColumn(GBillAmount, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//4 //GroupBill 30Mar2017
        TBillAmount += GBillAmount;//
        GBillAmount := 0;//GroupBill 30Mar2017

        //ExcelBuf.AddColumn(NetAmt,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//5
        ExcelBuf.AddColumn(GNetAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//5 //GrossNetAmount 30Mar2017
        TNetAmt += GNetAmt;//
        GNetAmt := 0;//GrossNetAmount 30Mar2017

        //ExcelBuf.AddColumn(ExcDuty,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//6
        ExcelBuf.AddColumn(GExcDuty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//6 //GrossExcDuty 30Mar2017
        TExcDuty += GExcDuty;//
        GExcDuty := 0;//GrossExcDuty 30Mar2017

        //ExcelBuf.AddColumn(TaxAmt,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//7
        ExcelBuf.AddColumn(GTaxAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//7 //GrossTaxAmt 30Mar2017
        TTaxAmt += GTaxAmt;//
        GTaxAmt := 0;//GrossTaxAmt 30Mar2017

        ExcelBuf.AddColumn(InvDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//8


        //ExcelBuf.AddColumn(CashDiscAmt,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//9
        ExcelBuf.AddColumn(GCashDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//9 //GrossCashDiscAmt 30Mar2017
        TCashDiscAmt += GCashDiscAmt;
        GCashDiscAmt := 0;//GrossCashDiscAmt 30Mar2017

        //ExcelBuf.AddColumn(FreightAmt,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//10
        ExcelBuf.AddColumn(GFreightAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//10 //GrossFreightAmt 30Mar2017
        TFreightAmt += GFreightAmt;
        GFreightAmt := 0;//GrossFreightAmt 30Mar2017

        //ExcelBuf.AddColumn(TradeSubsidy,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//11
        ExcelBuf.AddColumn(GTradeSubsidy, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//11 //GrossTradeSubsidy 30Mar2017
        TTradeSubsidy += GTradeSubsidy;//
        GTradeSubsidy := 0;//GrossTradeSubsidy 30Mar2017

        //ExcelBuf.AddColumn(Foc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//12
        ExcelBuf.AddColumn(GFoc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//12 //GrossFoc 30Mar2017
        TFoc += GFoc;//
        GFoc := 0;//GrossFoc 30Mar2017

        //ExcelBuf.AddColumn(Charges,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//13
        ExcelBuf.AddColumn(GCharges, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//13
        TCharges += GCharges;//
        GCharges := 0;//GrossCharges 30Mar2017

        //ExcelBuf.AddColumn(EntryTax,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//14
        ExcelBuf.AddColumn(GEntryTax, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//14 //GrossEntryTax 30Mar2017
        TEntryTax += GEntryTax;//
        GEntryTax := 0;//GrossEntryTax 30Mar2017

        //ExcelBuf.AddColumn(Addduty,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//15
        ExcelBuf.AddColumn(GAddduty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//15 //GrossAddduty 30Mar2017
        TAddduty += GAddduty;//
        GAddduty := 0;//GrossAddduty 30Mar2017

        //ExcelBuf.AddColumn(Cess,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//16
        ExcelBuf.AddColumn(GCess, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//16 //GrossCess 30Mar2017
        TCess += GCess;//
        GCess := 0;//GrossCess 30Mar2017
    end;

    [Scope('Internal')]
    procedure Makeexcelfooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn(TdecQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//3
        ExcelBuf.AddColumn(TBillAmount, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//4
        ExcelBuf.AddColumn(TNetAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//5
        ExcelBuf.AddColumn(TExcDuty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//6
        ExcelBuf.AddColumn(TTaxAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//7
        ExcelBuf.AddColumn(TInvDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//8
        ExcelBuf.AddColumn(TCashDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//9
        ExcelBuf.AddColumn(TFreightAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//10
        ExcelBuf.AddColumn(TTradeSubsidy, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//11
        ExcelBuf.AddColumn(TFoc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//12
        ExcelBuf.AddColumn(TCharges, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//13
        ExcelBuf.AddColumn(TEntryTax, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//14
        ExcelBuf.AddColumn(TAddduty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//15
        ExcelBuf.AddColumn(TCess, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//16
    end;
}


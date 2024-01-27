report 50109 "Sales CI - Summary"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesCISummary.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("G/L Account No.")
                                WHERE("Source Code" = FILTER('SALES'),
                                      "Document Type" = FILTER("Credit Memo"),
                                      Description = FILTER('@*Return Order*'));
            RequestFilterFields = "Posting Date";
            column(HDate; FORMAT(TODAY, 0, 4))
            {
            }
            column(DateFilter; DateFilter)
            {
            }
            column(LocationName; LocationName)
            {
            }
            column(GLAccNo; "G/L Entry"."G/L Account No.")
            {
            }
            column(GLAccName; GLAcc.Name)
            {
            }
            column(GLDrAmt; "G/L Entry"."Debit Amount")
            {
            }
            column(GLCrAmt; "G/L Entry"."Credit Amount")
            {
            }
            column(CompName; CompanyInfo.Name)
            {
            }

            trigger OnAfterGetRecord()
            begin

                TCOUNT += 1;//13APr2017


                //>>GroupingData
                GL13Apr.SETCURRENTKEY("G/L Account No.");
                GL13Apr.SETRANGE("G/L Account No.", "G/L Account No.");
                IF GL13Apr.FINDLAST THEN BEGIN

                    NNN := GL13Apr.COUNT;

                END;
                //<<GroupingData

                //>>1

                //ExcelData := TRUE;//13Apr2017
                IF LocationCode <> '' THEN BEGIN
                    SalesCrMemoheader.RESET;
                    IF SalesCrMemoheader.GET("G/L Entry"."Document No.") THEN BEGIN
                        IF SalesCrMemoheader."Location Code" <> LocationCode THEN
                            CurrReport.SKIP;
                    END;
                END;

                SalesCrMemoheader1.RESET;
                IF NOT SalesCrMemoheader1.GET("G/L Entry"."Document No.") THEN
                    CurrReport.SKIP;

                recSalesCrMemoHeader.RESET;
                IF recSalesCrMemoHeader.GET("G/L Entry"."Document No.") THEN BEGIN
                    IF (recSalesCrMemoHeader."Cancelled Invoice" = FALSE) THEN
                        CurrReport.SKIP;
                END;

                GLAcc.GET("G/L Entry"."G/L Account No.");
                //<<1

                //>>2

                //G/L Entry, Header (1) - OnPreSection()
                location.RESET;
                location.SETRANGE(location.Code, LocationCode);
                IF location.FIND('-') THEN
                    LocationName := location.Name;

                DateFilter := "G/L Entry".GETFILTER("G/L Entry"."Posting Date");

                Details := LocationName + '  Summary for  ' + DateFilter;


                //<<2


                //>>3

                //G/L Entry, GroupFooter (5) - OnPreSection()
                //IF PrintToExcel AND (("G/L Entry"."Debit Amount" <> 0) OR ("G/L Entry"."Credit Amount" <> 0)) THEN
                // MakeExcelDataBody;
                //<<3

                GCrAmt += "G/L Entry"."Credit Amount";//13APr2017
                GDrAmt += "G/L Entry"."Debit Amount";//13APr2017

                TotalCreditAmount += "G/L Entry"."Credit Amount";//13APr2017
                TotalDebitAmount += "G/L Entry"."Debit Amount";//13APr2017

                //>>GroupDataBody //13APr2017
                IF PrintToExcel AND (TCOUNT = NNN) AND ((GCrAmt <> 0) OR (GDrAmt <> 0)) THEN BEGIN

                    MakeExcelDataBody;
                    TCOUNT := 0;

                END;
                //>>GroupDataBody //13APr2017
            end;

            trigger OnPostDataItem()
            begin

                //>>4
                //G/L Entry, Footer (6) - OnPreSection()
                //TotalDebitAmount := TotalDebitAmount + "G/L Entry"."Debit Amount";
                //TotalCreditAmount := TotalCreditAmount + "G/L Entry"."Credit Amount";

                IF PrintToExcel THEN
                    MakeExcelforGrandTotal;
                //<<4
            end;

            trigger OnPreDataItem()
            begin

                GL13Apr.COPYFILTERS("G/L Entry");

                TCOUNT := 0;//13APr2017
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Location Code"; LocationCode)
                {
                    ApplicationArea = ALL;
                    TableRelation = Location;
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = ALL;
                    Visible = false;
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
        /*
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (MemberOf.FINDFIRST) THEN
           ERROR('You do not have permission to run the report');
        *///13APr2017 Commented

        CompanyInfo.GET;

        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<1

        //>>Report Header
        IF PrintToExcel THEN
            MakeExcelDataHeader;
        //<<Report Header

    end;

    var
        LocationCode: Code[20];
        SalesInvheader: Record 112;
        GLAcc: Record 15;
        CompanyInfo: Record 79;
        DateFilter: Text[30];
        GLAcc1: Record 15;
        TotalDebitAmount: Decimal;
        TotalCreditAmount: Decimal;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        Details: Text[100];
        SalesInvheader1: Record 112;
        SalesInvheader2: Record 112;
        location: Record 14;
        LocationName: Text[60];
        recSalesCrMemoHeader: Record 114;
        SalesCrMemoheader: Record 114;
        SalesCrMemoheader1: Record 114;
        Text000: Label 'Data';
        Text001: Label 'Sales Cancel invoice - Audit Summary ';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "----13APr2017": Integer;
        GL13Apr: Record 17;
        TCOUNT: Integer;
        NNN: Integer;
        GDrAmt: Decimal;
        GCrAmt: Decimal;
        ExcelData: Boolean;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//13Apr2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Sales Cancel invoice - Audit Summary ', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Sales CI - Summary", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;//13Apr2017
    end;

    //  //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>13Apr2017
        /* ExcelBuf.CreateBook('', Text001);
        ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<13Apr2017
    end;

    //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13APr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13APr2017
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13APr2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text001, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13APr2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13APr2017
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13APr2017

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location : ' + LocationName, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13APr2017
        ExcelBuf.AddColumn('Date : ' + DateFilter, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13APr2017
        //ExcelBuf.AddColumn(Details,FALSE,'',TRUE,FALSE,FALSE,'@',1);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('G/L Account', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13APr2017
        ExcelBuf.AddColumn('G/L Account Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13APr2017
        ExcelBuf.AddColumn('Debit Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13APr2017
        ExcelBuf.AddColumn('Credit Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13APr2017
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("G/L Entry"."G/L Account No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(GLAcc.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        //ExcelBuf.AddColumn("G/L Entry"."Debit Amount",FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);
        ExcelBuf.AddColumn(GDrAmt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//13APr2017
        GDrAmt := 0;//13APr2017

        //ExcelBuf.AddColumn("G/L Entry"."Credit Amount",FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);
        ExcelBuf.AddColumn(GCrAmt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//13APr2017
        GCrAmt := 0;//13APr2017
    end;

    //   //[Scope('Internal')]
    procedure MakeExcelforGrandTotal()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn(TotalDebitAmount, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);
        ExcelBuf.AddColumn(TotalCreditAmount, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);
    end;
}


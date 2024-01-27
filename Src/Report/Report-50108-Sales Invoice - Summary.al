report 50108 "Sales Invoice - Summary"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 30Aug2018   RB-N         Including Cancelled Invoice Details
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesInvoiceSummary.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("G/L Account No.")
                                WHERE("Source Code" = FILTER('SALES'),
                                      "Document Type" = FILTER(Invoice | "Credit Memo"));
            RequestFilterFields = "Posting Date";
            column(Name_ComInfo; CompanyInfo.Name)
            {
            }
            column(DateFilter; 'Summary for  ' + DateFilter)
            {
            }
            column(LocationName; LocationName)
            {
            }
            column(vToday; FORMAT(TODAY, 0, 4))
            {
            }
            column(GLAccountNo_GLEntry; "G/L Entry"."G/L Account No.")
            {
            }
            column(GLAcc_Name; GLAcc.Name)
            {
            }
            column(DebitAmount_GLEntry; "G/L Entry"."Debit Amount")
            {
            }
            column(CreditAmount_GLEntry; "G/L Entry"."Credit Amount")
            {
            }
            column(PageNo; CurrReport.PAGENO)
            {
            }
            column(UserID; USERID)
            {
            }

            trigger OnAfterGetRecord()
            begin
                /*
                IF LocationCode <> '' THEN
                BEGIN
                 SalesInvheader.RESET;
                 IF SalesInvheader.GET("G/L Entry"."Document No.") THEN
                 BEGIN
                  IF SalesInvheader."Location Code" <> LocationCode THEN
                     CurrReport.SKIP;
                 END;
                END;
                */

                IF "Document Type" = "Document Type"::Invoice THEN BEGIN
                    SalesInvheader1.RESET;
                    IF NOT SalesInvheader1.GET("G/L Entry"."Document No.") THEN
                        CurrReport.SKIP;

                    recSalesInvheader.RESET;
                    IF recSalesInvheader.GET("G/L Entry"."Document No.") THEN BEGIN
                        IF (recSalesInvheader."Order No." = '') AND (recSalesInvheader."Supplimentary Invoice" = FALSE) THEN
                            CurrReport.SKIP;
                    END;
                END;

                //>>30Aug2018
                IF "Document Type" = "Document Type"::"Credit Memo" THEN BEGIN
                    SCMH.RESET;
                    IF NOT SCMH.GET("Document No.") THEN
                        CurrReport.SKIP;

                    SCMH.RESET;
                    IF SCMH.GET("Document No.") THEN BEGIN
                        //IF SCMH."Applies-to Doc. No." = '' THEN
                        //CurrReport.SKIP;

                        IF SCMH."Applies-to Doc. No." <> '' THEN BEGIN
                            SIH.RESET;
                            IF SIH.GET(SCMH."Applies-to Doc. No.") THEN
                                IF NOT SIH."Cancelled Invoice" THEN
                                    CurrReport.SKIP;
                        END;
                    END;
                END;
                //<<30Aug2018

                GLAcc.GET("G/L Entry"."G/L Account No.");


                /*
                //>>
                IF vGLNo <> "G/L Entry"."G/L Account No." THEN
                BEGIN
                  CLEAR(Cr);
                  CLEAR(Dr);
                  //CLEAR(TotCr);
                  //CLEAR(TotDr);
                  vGLEntry.RESET;
                  vGLEntry.SETRANGE("G/L Account No.","G/L Account No.");
                  vGLEntry.SETRANGE("Posting Date",Mindate,MaxDate);
                  IF vGLEntry.FINDSET THEN
                  REPEAT
                    IF LocationCode <> '' THEN
                    BEGIN
                      SalesInvheader.RESET;
                      IF SalesInvheader.GET(vGLEntry."Document No.") THEN
                      IF (SalesInvheader."Location Code" = LocationCode) THEN
                      BEGIN
                        Cr += vGLEntry."Credit Amount";
                        Dr += vGLEntry."Debit Amount";
                      END;
                    END;
                  UNTIL vGLEntry.NEXT =0;
                  //<<
                  IF PrintToExcel AND (("G/L Entry"."Debit Amount" <> 0) OR ("G/L Entry"."Credit Amount" <> 0)) THEN
                     MakeExcelDataBody;
                END;
                
                vGLNo := "G/L Entry"."G/L Account No.";
                TotCr += "Credit Amount";
                TotDr += "Debit Amount"
                */

            end;

            trigger OnPostDataItem()
            begin

                TotalDebitAmount := TotalDebitAmount + "G/L Entry"."Debit Amount";
                TotalCreditAmount := TotalCreditAmount + "G/L Entry"."Credit Amount";

                IF PrintToExcel THEN
                    MakeExcelforGrandTotal;
            end;

            trigger OnPreDataItem()
            begin

                location.RESET;
                location.SETRANGE(location.Code, LocationCode);
                IF location.FIND('-') THEN
                    LocationName := location.Name;

                DateFilter := "G/L Entry".GETFILTER("G/L Entry"."Posting Date");

                Details := LocationName + '  Summary for  ' + DateFilter;
                Mindate := "G/L Entry".GETRANGEMIN("Posting Date");
                MaxDate := "G/L Entry".GETRANGEMAX("Posting Date");
                IF PrintToExcel THEN
                    MakeExcelDataHeader;

                //>>30Aug2018
                IF LocationCode <> '' THEN;
                //SETFILTER("G/L Entry"."Location Code", LocationCode);
                //<<30Aug2018
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
                    ApplicationArea = all;
                    Caption = 'Location Code';
                    TableRelation = Location;
                }
                field("Print to Excel"; PrintToExcel)
                {
                    ApplicationArea = all;
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

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin

        CompanyInfo.GET;

        IF PrintToExcel THEN
            MakeExcelInfo;
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
        recSalesInvheader: Record 112;
        Text000: Label 'Data';
        Text001: Label 'Sales Invoice - Audit Summary';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "---": Integer;
        Cr: Decimal;
        Dr: Decimal;
        TotCr: Decimal;
        TotDr: Decimal;
        vGLEntry: Record 17;
        Mindate: Date;
        MaxDate: Date;
        vGLNo: Code[30];
        SCMH: Record 114;
        SIH: Record 112;

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Sales Invoice - Audit Summary', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Sales Invoice - Summary", FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 1);
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ExcelBuf.CreateBookAndOpenExcel('', Text000, Text001, COMPANYNAME, USERID);
        ExcelBuf.CreateNewBook(Text000);
        ExcelBuf.WriteSheet(Text000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text000, Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(Details, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('G/L Account', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('G/L Account Name', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('Debit Amount', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('Credit Amount', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("G/L Entry"."G/L Account No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(GLAcc.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.AddColumn(Dr, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(Cr, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
    end;

    // //[Scope('Internal')]
    procedure MakeExcelforGrandTotal()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(TotDr, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotCr, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuf."Cell Type"::Number);
    end;
}


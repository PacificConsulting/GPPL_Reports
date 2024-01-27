report 50111 "Sales Credit Memo - Summary"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesCreditMemoSummary.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("G/L Account No.")
                                WHERE("Source Code" = FILTER('SALES'),
                                      "Document Type" = FILTER("Credit Memo"),
                                      Description = FILTER('Credit Memo'));
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
                //>>1

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

                //IF PrintToExcel THEN
                //MakeExcelDataHeader;
                //<<2
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
            }
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPreReport()
    begin
        //>>1
        /*
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (MemberOf.FINDFIRST) THEN
           ERROR('You do not have permission to run the report');
        
        */
        CompanyInfo.GET;
        //<<1

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
        Text001: Label 'Sales Credit Memo - Audit Summary ';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
}


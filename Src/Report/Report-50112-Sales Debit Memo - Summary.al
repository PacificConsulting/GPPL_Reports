report 50112 "Sales Debit Memo - Summary"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesDebitMemoSummary.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("G/L Account No.")
                                WHERE("Source Code" = FILTER('SALES'),
                                      "Document Type" = FILTER(Invoice),
                                      Description = FILTER(<> 'Order'));
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
                    SalesInvheader.RESET;
                    IF SalesInvheader.GET("G/L Entry"."Document No.") THEN BEGIN
                        IF SalesInvheader."Location Code" <> LocationCode THEN
                            CurrReport.SKIP;
                    END;
                END;

                SalesInvheader1.RESET;
                IF NOT SalesInvheader1.GET("G/L Entry"."Document No.") THEN
                    CurrReport.SKIP;


                recSalesInvheader.RESET;
                IF recSalesInvheader.GET("G/L Entry"."Document No.") THEN BEGIN
                    IF (recSalesInvheader."Order No." = '') AND (recSalesInvheader."Supplimentary Invoice" = TRUE) THEN
                        CurrReport.SKIP;
                END;

                GLAcc.GET("G/L Entry"."G/L Account No.");
                //<<1

                //>>2

                location.RESET;
                location.SETRANGE(location.Code, LocationCode);
                IF location.FIND('-') THEN
                    LocationName := location.Name;

                DateFilter := "G/L Entry".GETFILTER("G/L Entry"."Posting Date");

                Details := LocationName + '  Summary for  ' + DateFilter;

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
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",USERID);
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (Memberof.FINDFIRST) THEN
           ERROR('You do not have permission to run the report');
        *///Commented 17Apr2017

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
        recSalesInvheader: Record 112;
        Text000: Label 'Data';
        Text001: Label 'Sales Debit Memo - Audit Summary';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
}


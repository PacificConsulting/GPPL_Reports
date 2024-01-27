report 50110 "Sales Return - Summary"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesReturnSummary.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("G/L Account No.")
                                WHERE("Source Code" = FILTER('SALES'),
                                      "Document Type" = FILTER("Credit Memo"),
                                      Description = FILTER('Return Order'));
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

                recSalesCrMemoHeader.RESET;
                IF recSalesCrMemoHeader.GET("G/L Entry"."Document No.") THEN BEGIN
                    IF (recSalesCrMemoHeader."Cancelled Invoice" = TRUE) THEN
                        CurrReport.SKIP;
                END;

                /*
                //EBT STIVAN ---(10102011)----To make null value if SI and SCM Posting Date is Same----START
                SalesInvheader.RESET;
                SalesInvheader.SETFILTER(SalesInvheader."OverDue Balance",'%1',TRUE);
                SalesInvheader.SETFILTER(SalesInvheader."No.","G/L Entry"."Document No.");
                IF SalesInvheader.FINDFIRST THEN
                BEGIN
                 recSalesCrMemoHeader.RESET;
                 recSalesCrMemoHeader.SETFILTER("Applies-to Doc. No.",'%1', SalesInvheader."No.");
                 IF  recSalesCrMemoHeader.FINDFIRST THEN
                  IF recSalesCrMemoHeader."Posting Date" = SalesInvheader."Posting Date" THEN
                  BEGIN
                   CurrReport.SKIP;
                  END;
                END;
                //EBT STIVAN ---(10102011)----To make null value if SI and SCM Posting Date is Same------END
                */

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
        //EBT STIVAN ---(21082012)--- To Filter Report as per the Report View Role ------START
        MemberOf.RESET;
        MemberOf.SETRANGE(MemberOf."User ID",USERID);
        MemberOf.SETFILTER(MemberOf."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT (MemberOf.FINDFIRST) THEN
        ERROR('You do not have permission to run the report');
        //EBT STIVAN ---(21082012)--- To Filter Report as per the Report View Role --------END
        *///Commented 14Apr2017

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
}


report 50223 "Customer Invoice Yearly"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CustomerInvoiceYearly.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            //CalcFields = "Amount to Customer";
            DataItemTableView = SORTING("Posting Date");
            RequestFilterFields = "Posting Date";
            column(PostingDate_SalesInvoiceHeader; "Posting Date")
            {
            }
            column(AmounttoCustomer_SalesInvoiceHeader; 0)// "Amount to Customer")
            {
            }
            column(No_SalesInvoiceHeader; "No.")
            {
            }
            column(AmtLCY; AmtLCY)
            {
            }

            trigger OnAfterGetRecord()
            begin
                AmtLCY := 0;
                IF "Currency Factor" <> 0 THEN
                    AmtLCY := 0// "Amount to Customer" * (1 / "Currency Factor")
                ELSE
                    AmtLCY := 0;// "Amount to Customer";
            end;

            trigger OnPreDataItem()
            begin
                // SETFILTER("Amount to Customer", '<>0');
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

    var
        AmtLCY: Decimal;
}


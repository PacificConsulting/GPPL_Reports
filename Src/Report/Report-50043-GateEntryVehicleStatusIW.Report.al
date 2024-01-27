report 50043 "Gate Entry Vehicle Status I/W"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 30Dec2017   RB-N         New Report Development as per NAV2009 R#50043
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/GateEntryVehicleStatusIW.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Gate Entry Header"; "Gate Entry Header")
        {
            DataItemTableView = SORTING("Entry Type", "No.")
                                WHERE("Entry Type" = FILTER(Inward),
                                      "Location Code" = FILTER('PLANT01'),
                                      "Short Closed" = FILTER(false));
            RequestFilterFields = "Document Date";
            column(Datefilter; Datefilter)
            {
            }
            column(CompName; Companyinfo.Name)
            {
            }
            column(LocationName; "Vehicle For Location")
            {
            }
            column(DocNo; "No.")
            {
            }
            column(DocDate; FORMAT("Document Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(DocTime; "Document Time")
            {
            }
            column(VehicleNo; "Vehicle No.")
            {
            }
            column(Description; Description)
            {
            }
            column(ItemName; "Item Description")
            {
            }
            column(TransporterName; "Transporter Name")
            {
            }

            trigger OnPreDataItem()
            begin
                Datefilter := "Gate Entry Header".GETFILTER("Gate Entry Header"."Document Date");
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

    trigger OnPreReport()
    begin
        Companyinfo.GET;
    end;

    var
        Companyinfo: Record 79;
        Datefilter: Text[30];
}


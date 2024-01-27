report 50149 "Vehicle Loading Advice"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 19Dec2017   RB-N         New Field (Check Post)
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/VehicleLoadingAdvice.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Gate Entry Header"; "Gate Entry Header")
        {
            RequestFilterFields = "No.";
            column(DocNo; "Gate Entry Header"."No.")
            {
            }
            column(DocDate; "Document Date")
            {
            }
            column(VehicleNo; "Vehicle No.")
            {
            }
            column(TransportName; "Transporter Name")
            {
            }
            column(DocTime; "Document Time")
            {
            }
            column(DriverName; "Driver's Name")
            {
            }
            column(DriverLicNo; "Driver's License No.")
            {
            }
            column(DriverMobNo; "Driver's Mobile No.")
            {
            }
            column(Description; Description)
            {
            }
            column(VehicleBodyHeight; "Vehicle Body Height/Length")
            {
            }
            column(VehicleType; "Type of Vehicle")
            {
            }
            column(Category1; Category1)
            {
            }
            column(Category2; Category2)
            {
            }
            column(Category3; Category3)
            {
            }
            column(Category4; Category4)
            {
            }
            column(LRNo; "LR/RR No.")
            {
            }
            column(VehicleInsDate; "Vehicle Insurance Date")
            {
            }
            column(PucDate; "PUC Date")
            {
            }
            column(VehicleCapacity; "Vehicle Capacity")
            {
            }
            column(UserName; UserName)
            {
            }
            column(CheckPost; "Check Post")
            {
            }
            column(GrossWeight; "Gross Weight")
            {
            }
            column(TareWeight; "Tare Weight")
            {
            }
            column(NetWeight; "Net Weight")
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1
                // EBT MILAN (141013)............................START
                user.RESET;
                //user.SETRANGE(user."User ID","Gate Entry Header"."Release User ID");//
                user.SETRANGE(user."User Name", "Gate Entry Header"."Release User ID");//29Mar2017
                IF user.FINDFIRST THEN;
                UserName := user."Full Name";//29Mar2017

                // EBT MILAN (141013)............................END
                //<<1


                //>>2

                //Gate Entry Header, Body (3) - OnPreSection()
                IF "Gate Entry Header"."Category of Tanker 1" <> "Gate Entry Header"."Category of Tanker 1"::" " THEN
                    Category1 := '1: ' + FORMAT("Gate Entry Header"."Category of Tanker 1")
                ELSE
                    Category1 := '1: ' + FORMAT("Gate Entry Header"."Type of Truck");

                IF "Gate Entry Header"."Type of Vehicle" = "Gate Entry Header"."Type of Vehicle"::Tanker THEN BEGIN
                    Category2 := '2: ' + FORMAT("Gate Entry Header"."Category of Tanker 3");
                    Category3 := '3 :' + FORMAT("Gate Entry Header"."Category of Tanker 2");
                    Category4 := '4: ' + FORMAT("Gate Entry Header"."Category of Tanker 4");
                END;
                //<<2
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
        user: Record 2000000120;
        Category1: Text[30];
        Category2: Text[30];
        Category3: Text[30];
        Category4: Text[30];
        "----29Mar2017": Integer;
        UserName: Text[80];
}


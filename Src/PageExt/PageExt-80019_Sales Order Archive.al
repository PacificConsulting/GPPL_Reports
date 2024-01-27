pageextension 80019 SalesOrderArchiveExt extends "Sales Order Archive"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        // Add changes to page actions here
        addlast(reporting)
        {
            action("Sales Order Archives-Industrial GST")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                VAR
                    SOA: Record 5107;
                BEGIN

                    //26Sep2017
                    SOA.RESET;
                    SOA.SETRANGE("No.", rec."No.");
                    SOA.SETRANGE("Version No.", rec."Version No.");
                    IF SOA.FINDFIRST THEN
                        REPORT.RUNMODAL(50047, TRUE, TRUE, SOA);

                    //26Sep2017
                END;
            }
            action("Sales Order Archives-Automotive GST")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                VAR
                    SOA: Record 5107;
                BEGIN

                    //27Sep2017
                    SOA.RESET;
                    SOA.SETRANGE("No.", rec."No.");
                    SOA.SETRANGE("Version No.", rec."Version No.");
                    IF SOA.FINDFIRST THEN
                        REPORT.RUNMODAL(50050, TRUE, TRUE, SOA);

                    //27Sep2017
                END;
            }
            action("Sales Order Archives-Indl")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                VAR
                    SOA: Record 5107;
                BEGIN

                    //27Sep2017
                    SOA.RESET;
                    SOA.SETRANGE("No.", rec."No.");
                    SOA.SETRANGE("Version No.", rec."Version No.");
                    IF SOA.FINDFIRST THEN
                        REPORT.RUNMODAL(50053, TRUE, TRUE, SOA);

                    //27Sep2017
                END;
            }
            action("Sales Order Archives-Auto")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                VAR
                    SOA: Record 5107;
                BEGIN

                    //27Sep2017
                    SOA.RESET;
                    SOA.SETRANGE("No.", rec."No.");
                    SOA.SETRANGE("Version No.", rec."Version No.");
                    IF SOA.FINDFIRST THEN
                        REPORT.RUNMODAL(50067, TRUE, TRUE, SOA);

                    //27Sep2017
                END;
            }

        }
    }

    var
        myInt: Integer;
}
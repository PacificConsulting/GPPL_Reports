pageextension 80020 SalesOrderArchivesExt extends "Sales Order Archives"
{
    layout
    {
        // Add changes to page layout here

    }

    actions
    {
        // Add changes to page actions here
        addlast(Reporting)
        {
            action("Sales Order Archives-Industrial GST")
            {
                ApplicationArea = all;
                Visible = false;
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
                Visible = false;
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
                Visible = false;
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
                Visible = false;
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
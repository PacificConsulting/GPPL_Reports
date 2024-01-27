pageextension 90001 PosSaleInvListExtcstm extends "Posted Sales Invoices"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addlast(Reporting)
        {
            action("Export Invoice")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader2: Record "Sales Invoice Header";
                begin
                    SalesInvHeader2.RESET;
                    SalesInvHeader2.SETRANGE("No.", Rec."No.");
                    IF SalesInvHeader2.FINDFIRST THEN begin
                        REPORT.RUNMODAL(50000, TRUE, FALSE, SalesInvHeader2);
                    END;
                end;
            }
            action("Print Plant Invoice-&IL")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader2: Record "Sales Invoice Header";
                begin
                    SalesInvHeader2.RESET;
                    SalesInvHeader2.SETRANGE("No.", Rec."No.");
                    IF SalesInvHeader2.FINDFIRST THEN begin
                        REPORT.RUNMODAL(50003, TRUE, FALSE, SalesInvHeader2);
                    END;
                end;
            }
            action("Print Plant Invoice-&AL")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader2: Record "Sales Invoice Header";
                begin
                    SalesInvHeader2.RESET;
                    SalesInvHeader2.SETRANGE("No.", Rec."No.");
                    IF SalesInvHeader2.FINDFIRST THEN begin
                        REPORT.RUNMODAL(50005, TRUE, FALSE, SalesInvHeader2);
                    END;
                end;
            }
            action("Print &QC Test Report")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader2: Record "Sales Invoice Header";
                begin
                    SalesInvHeader2.RESET;
                    SalesInvHeader2.SETRANGE("No.", Rec."No.");
                    IF SalesInvHeader2.FINDFIRST THEN begin
                        REPORT.RUNMODAL(50008, TRUE, FALSE, SalesInvHeader2);
                    END;
                end;
            }

            action("Purchase Order")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader2: Record "Sales Invoice Header";
                begin
                    SalesInvHeader2.RESET;
                    SalesInvHeader2.SETRANGE("No.", Rec."No.");
                    IF SalesInvHeader2.FINDFIRST THEN begin
                        REPORT.RUNMODAL(50010, TRUE, FALSE, SalesInvHeader2);
                    END;
                end;
            }
            action("Invoice GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader2: Record "Sales Invoice Header";
                begin


                    SalesInvHeader2.RESET;
                    SalesInvHeader2.SETRANGE("No.", Rec."No.");
                    IF SalesInvHeader2.FINDFIRST THEN
                        REPORT.RUNMODAL(50025, TRUE, TRUE, SalesInvHeader2);
                end;
            }
            action("Summary SalesInvoice GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", rec."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50039, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Delivery Challan GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    TSH: Record 112;
                BEGIN
                    // CAS-22565-T8N8W6 13Aug2018
                    IF rec."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    TSH.RESET;
                    TSH.SETRANGE("No.", rec."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50042, TRUE, FALSE, TSH);
                END;
            }
            action("Sales Debit Memo")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", rec."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50044, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Print Depot Invoice-AL-&Exciseable")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50055, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Print Depot Invoice-AL-&NonExciseable")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50059, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Scrap Sales Invoice - GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", rec."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50060, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Export Invoice GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    TSH: Record 112;
                BEGIN
                    //CAS-22565-T8N8W6 13Aug2018
                    IF rec."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    TSH.RESET;
                    TSH.SETRANGE("No.", rec."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50071, TRUE, FALSE, TSH);
                END;
            }
            action("Invoice Industrial GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    // CAS-22565-T8N8W6 13Aug2018
                    IF rec."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", rec."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50072, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Invoice Automotive GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    //   CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50073, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Debit Note GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    //  CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50075, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Invoice Industrial GST--LOGO")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    //  CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50076, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Invoice Automotive GST--LOGO")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    //CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50077, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("COAL Invoice GST")
            {
                ApplicationArea = ALL;
                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    // CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50107, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Print Structure Details-AL")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50137, TRUE, FALSE, SalesInvHeader);
                END;

            }
            action("POP Invoice")
            {
                ApplicationArea = all;
                // Name =POP Invoice;
                Image = Print;
                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50146, TRUE, FALSE, SalesInvHeader);
                END;

            }
            action("Print Depot Invoice-IL")
            {
                ApplicationArea = all;
                // Name =POP Invoice;
                Image = Print;
                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50157, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Bond Transfer Invoice")
            {
                ApplicationArea = all;
                // Name =POP Invoice;
                Image = Print;
                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50181, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Bond Transfer Invoice GST")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    // CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    // RB-N 04Nov2017
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50198, TRUE, FALSE, SalesInvHeader);
                    //RB-N 04Nov2017
                END;
            }
            action("Print Depot-AL-GST")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50199, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Trading Invoice GST")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50200, TRUE, FALSE, SalesInvHeader);
                END;
            }




        }
    }

    var
        myInt: Integer;



}

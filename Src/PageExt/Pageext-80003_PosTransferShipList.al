pageextension 80003 PosTransShipListExtRep extends "Posted Transfer Shipments"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addlast(Reporting)
        {
            action("BRT - Plant - &IL")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    TransShptHeader: Record "Transfer Shipment Header";
                begin
                    TransShptHeader.RESET;
                    TransShptHeader := Rec;
                    TransShptHeader.SETRECFILTER;
                    REPORT.RUN(50006, TRUE, FALSE, TransShptHeader);
                end;
            }
            action("BRT - Plant - &AL")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    TransShptHeader: Record "Transfer Shipment Header";
                begin
                    TransShptHeader.RESET;
                    TransShptHeader := Rec;
                    TransShptHeader.SETRECFILTER;
                    REPORT.RUN(50007, TRUE, FALSE, TransShptHeader);
                end;
            }

            action("Summary TransferInvoice GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    TSH: Record "Transfer Shipment Header";
                BEGIN

                    TsH.RESET;
                    TsH.SETRANGE("No.", rec."No.");
                    IF TsH.FINDFIRST THEN
                        REPORT.RUNMODAL(50049, TRUE, FALSE, TsH);
                END;
            }
            action("BRT - Depot - AL - Non &Exciseable")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    TransShptHeader: Record "Transfer Shipment Header";
                BEGIN

                    TransShptHeader.RESET;
                    TransShptHeader := Rec;
                    TransShptHeader.SETRECFILTER;
                    REPORT.RUN(50070, TRUE, FALSE, TransShptHeader);
                END;
            }
            action("Transfer Invoice GST - New")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    TSH: Record 5744;
                BEGIN
                    //CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    TSH.RESET;
                    TSH.SETRANGE("No.", REC."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50082, TRUE, FALSE, TSH);
                END;
            }
            action("Transfer Invoice GST LOGO")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                VAR
                    TSH: Record 5744;
                BEGIN
                    //    CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    TsH.RESET;
                    TsH.SETRANGE("No.", REC."No.");
                    IF TsH.FINDFIRST THEN
                        REPORT.RUNMODAL(50083, TRUE, FALSE, TsH);
                END;
            }
            action("Transfer Invoice GST")
            {
                ApplicationArea = all;
                trigger OnAction()
                VAR
                    TSH: Record 5744;
                BEGIN
                    //CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    TSH.RESET;
                    TSH.SETRANGE("No.", REC."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50154, TRUE, FALSE, TSH);
                END;
            }
            action("BRT - Depot - I&L")
            {
                ApplicationArea = ALL;
                Image = Print;
                trigger OnAction()
                var
                    TransShptHeader: Record "Transfer Shipment Header";
                BEGIN

                    TransShptHeader.RESET;
                    TransShptHeader := Rec;
                    TransShptHeader.SETRECFILTER;
                    REPORT.RUN(50158, TRUE, FALSE, TransShptHeader);
                END;
            }
            action("BRT - &Depot - AL")
            {
                ApplicationArea = ALL;
                Image = Print;
                trigger OnAction()
                var
                    TransShptHeader: Record "Transfer Shipment Header";
                BEGIN

                    TransShptHeader.RESET;
                    TransShptHeader := Rec;
                    TransShptHeader.SETRECFILTER;
                    REPORT.RUN(50159, TRUE, FALSE, TransShptHeader);
                END;
            }
            action("Print E-Way Bill")
            {
                ApplicationArea = ALL;
                Image = Print;
                trigger OnAction()
                VAR
                    RecTSH: Record 5744;
                BEGIN
                    RecTSH.RESET;
                    RecTSH.SETCURRENTKEY("No.", "No.");
                    RecTSH.SETRANGE("No.", rec."No.");
                    IF RecTSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50234, TRUE, TRUE, RecTSH);
                END;
            }

        }
    }

    var
        myInt: Integer;
}
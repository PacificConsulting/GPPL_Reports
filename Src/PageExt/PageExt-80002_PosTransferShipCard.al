pageextension 80002 PosTransShipExtRep extends "Posted Transfer Shipment"
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
                ApplicationArea = all;
                trigger OnAction()
                var
                    TSH: Record "Transfer Shipment Header";
                BEGIN

                    TSH.RESET;
                    TSH.SETRANGE("No.", rec."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50049, TRUE, FALSE, TSH);
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
                    RecTSL: Record "Transfer Shipment Line";
                BEGIN
                    // CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    //RSPLSUM 13Jan21


                    IF REC."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(IRN);
                        RecTSL.RESET;
                        RecTSL.SETRANGE("Document No.", REC."No.");
                        RecTSL.SETFILTER(Quantity, '%1', 0);
                        // RecTSL.SETFILTER("Total GST Amount", '%1', 0); templory comment 1-jan-23 pcpl-065
                        IF RecTSL.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    TSH.RESET;
                    TSH.SETRANGE("No.", REC."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50082, TRUE, FALSE, TSH);
                END;
            }
            action("Transfer Invoice GST LOGO")
            {
                ApplicationArea = ALL;
                Image = Print;
                trigger OnAction()
                var

                    RecTSL: Record "Transfer Shipment Line";
                    TSH: Record "Transfer Shipment Header";
                BEGIN
                    //   CAS-22565-T8N8W6 13Aug2018
                    IF REC."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    //RSPLSUM 13Jan21


                    IF REC."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(IRN);
                        RecTSL.RESET;
                        RecTSL.SETRANGE("Document No.", REC."No.");
                        RecTSL.SETFILTER(Quantity, '%1', 0);
                        //  RecTSL.SETFILTER("Total GST Amount", '%1', 0); tempory comment 1-jan-23 pcpl-065
                        IF RecTSL.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    TSH.RESET;
                    TSH.SETRANGE("No.", REC."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50083, TRUE, FALSE, TSH);
                END;
            }
            action("Transfer Invoice GST")
            {
                ApplicationArea = all;
                trigger OnAction()
                VAR
                    TSH: Record 5744;
                    RecTSL: Record "Transfer Shipment Line";
                BEGIN
                    //  CAS-22565-T8N8W6 13Aug2018
                    IF rec."Road Permit No." = '' THEN
                        ERROR('E-Way Bill No. cannot be blank');
                    //

                    //RSPLSUM 13Jan21


                    IF rec."Posting Date" > 20210113D/*130121D */THEN BEGIN
                        rec.CALCFIELDS(IRN);
                        RecTSL.RESET;
                        RecTSL.SETRANGE("Document No.", REC."No.");
                        RecTSL.SETFILTER(Quantity, '%1', 0);
                        // RecTSL.SETFILTER("Total GST Amount", '%1', 0); TEMPORY COMMENT 3 JAN 23 
                        IF RecTSL.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

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
                Promoted = true;
                Visible = TRUE;
                PromotedIsBig = true;
                PromotedCategory = Report;
                trigger OnAction()

                VAR
                    RecTSH: Record 5744;
                BEGIN
                    RecTSH.RESET;
                    RecTSH.SETCURRENTKEY("No.", "No.");
                    RecTSH.SETRANGE("No.", REC."No.");
                    IF RecTSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50234, TRUE, TRUE, RecTSH);
                END;
            }

        }
    }

    var
        myInt: Integer;
}
pageextension 80000 PosSaleInvCardExtcstm extends "Posted Sales Invoice"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addlast(Reporting)
        {
            action("Tax Invoice")
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
            action("Bill of Supply")
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
                        REPORT.RUNMODAL(50001, TRUE, FALSE, SalesInvHeader2);
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

            action("Invoice GST July")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader2: Record "Sales Invoice Header";
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
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
                ApplicationArea = all;
                Image = Report;
                trigger OnAction()
                VAR
                    TSH: Record 112;
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
                BEGIN
                    //CAS-22565-T8N8W6 13Aug2018
                    IF (rec."EWB No." = '') AND (rec."EWB Date" = 0D) THEN BEGIN//RSPLSUM 30Jul2020
                        IF rec."Road Permit No." = '' THEN
                            ERROR('E-Way Bill No. cannot be blank');
                    END;//RSPLSUM 30Jul2020
                        //

                    //RSPLSUM 13Jan21


                    IF rec."Posting Date" > 20210113D /*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(REC.IRN);

                        RecSalesInvoiceLine.RESET;
                        RecSalesInvoiceLine.SETRANGE("Document No.", REC."No.");
                        RecSalesInvoiceLine.SETFILTER(Quantity, '%1', 0);
                        // RecSalesInvoiceLine.SETFILTER("Total GST Amount", '%1', 0); TEMPORY COMMENT 3-jan-23 pcpl-065
                        IF RecSalesInvoiceLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    TSH.RESET;
                    TSH.SETRANGE("No.", REC."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50042, TRUE, FALSE, TSH);
                END;
            }
            action("Scrap Sales Invoice - GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    SalesInvHeader2: Record "Sales Invoice Header";
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
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
                BEGIN
                    //CAS-22565-T8N8W6 13Aug2018
                    IF (rec."EWB No." = '') AND (rec."EWB Date" = 0D) THEN BEGIN//RSPLSUM 30Jul2020
                        IF rec."Road Permit No." = '' THEN
                            ERROR('E-Way Bill No. cannot be blank');
                    END;//RSPLSUM 30Jul2020
                        //

                    //RSPLSUM 13Jan21


                    IF rec."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        rec.CALCFIELDS(IRN);
                        RecSalesInvoiceLine.RESET;
                        RecSalesInvoiceLine.SETRANGE("Document No.", REC."No.");
                        RecSalesInvoiceLine.SETFILTER(Quantity, '%1', 0);
                        // RecSalesInvoiceLine.SETFILTER("Total GST Amount", '%1', 0); Tempory comment 2-jan-2023 pcpl-065
                        IF RecSalesInvoiceLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    TSH.RESET;
                    TSH.SETRANGE("No.", REC."No.");
                    IF TSH.FINDFIRST THEN
                        REPORT.RUNMODAL(50071, TRUE, FALSE, TSH);
                END;
            }
            action("Invoice Industrial GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
                BEGIN
                    // CAS-22565-T8N8W6 13Aug2018
                    IF (rec."EWB No." = '') AND (rec."EWB Date" = 0D) THEN BEGIN//RSPLSUM 30Jul2020
                        IF rec."Road Permit No." = '' THEN
                            ERROR('E-Way Bill No. cannot be blank');
                    END;//RSPLSUM 30Jul2020
                        //

                    //RSPLSUM 13Jan21


                    IF REC."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(IRN);
                        RecSalesInvoiceLine.RESET;
                        RecSalesInvoiceLine.SETRANGE("Document No.", REC."No.");
                        RecSalesInvoiceLine.SETFILTER(Quantity, '%1', 0);
                        // RecSalesInvoiceLine.SETFILTER("Total GST Amount", '%1', 0);Tempory comment 2-jan-23 pcpl-065
                        IF RecSalesInvoiceLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50072, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Invoice Automotive GST")
            {
                ApplicationArea = All;
                Image = Report;

                trigger OnAction()
                var
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
                BEGIN
                    //CAS-22565-T8N8W6 13Aug2018
                    IF (REC."EWB No." = '') AND (REC."EWB Date" = 0D) THEN BEGIN//RSPLSUM 30Jul2020
                        IF REC."Road Permit No." = '' THEN
                            ERROR('E-Way Bill No. cannot be blank');
                    END;//RSPLSUM 30Jul2020
                        //

                    //RSPLSUM 13Jan21


                    IF REC."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(IRN);
                        RecSalesInvoiceLine.RESET;
                        RecSalesInvoiceLine.SETRANGE("Document No.", REC."No.");
                        RecSalesInvoiceLine.SETFILTER(Quantity, '%1', 0);
                        // RecSalesInvoiceLine.SETFILTER("Total GST Amount", '%1', 0); TEMPORY COMMENT 3-jan-23 pcpl-065
                        IF RecSalesInvoiceLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50073, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Sales Debit Note GST")
            {
                ApplicationArea = ALL;
                Image = Print;
                trigger OnAction()
                var
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
                BEGIN
                    // CAS-22565-T8N8W6 13Aug2018
                    IF (REC."EWB No." = '') AND (REC."EWB Date" = 0D) THEN BEGIN//RSPLSUM 30Jul2020
                        IF REC."Road Permit No." = '' THEN
                            ERROR('E-Way Bill No. cannot be blank');
                    END;//RSPLSUM 30Jul2020
                        //

                    //RSPLSUM 13Jan21


                    IF REC."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(IRN);
                        RecSalesInvoiceLine.RESET;
                        RecSalesInvoiceLine.SETRANGE("Document No.", REC."No.");
                        RecSalesInvoiceLine.SETFILTER(Quantity, '%1', 0);
                        // RecSalesInvoiceLine.SETFILTER("Total GST Amount", '%1', 0); TEMPORY COMMENT 3-jan-23 pcpl-065
                        IF RecSalesInvoiceLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50075, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Invoice Industrial GST--LOGO")
            {
                ApplicationArea = ALL;
                Image = Print;
                trigger OnAction()
                var
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
                BEGIN
                    // CAS-22565-T8N8W6 13Aug2018
                    IF (REC."EWB No." = '') AND (REC."EWB Date" = 0D) THEN BEGIN//RSPLSUM 30Jul2020
                        IF REC."Road Permit No." = '' THEN
                            ERROR('E-Way Bill No. cannot be blank');
                    END;//RSPLSUM 30Jul2020
                        //

                    //RSPLSUM 13Jan21


                    IF REC."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(IRN);
                        RecSalesInvoiceLine.RESET;
                        RecSalesInvoiceLine.SETRANGE("Document No.", REC."No.");
                        RecSalesInvoiceLine.SETFILTER(Quantity, '%1', 0);
                        // RecSalesInvoiceLine.SETFILTER("Total GST Amount", '%1', 0); TEMPORY COMMENT 3-jan-23 pcpl-065
                        IF RecSalesInvoiceLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50076, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Invoice Automotive GST--LOGO")
            {
                ApplicationArea = ALL;
                Image = Print;
                trigger OnAction()
                var
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
                BEGIN
                    //  CAS-22565-T8N8W6 13Aug2018
                    IF (REC."EWB No." = '') AND (REC."EWB Date" = 0D) THEN BEGIN//RSPLSUM 30Jul2020
                        IF REC."Road Permit No." = '' THEN
                            ERROR('E-Way Bill No. cannot be blank');
                    END;//RSPLSUM 30Jul2020
                        //

                    //RSPLSUM 13Jan21


                    IF REC."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        REC.CALCFIELDS(IRN);
                        RecSalesInvoiceLine.RESET;
                        RecSalesInvoiceLine.SETRANGE("Document No.", REC."No.");
                        RecSalesInvoiceLine.SETFILTER(Quantity, '%1', 0);
                        //   RecSalesInvoiceLine.SETFILTER("Total GST Amount", '%1', 0); TEMPORY COMMENT 3-jan-23 pcpl-065
                        IF RecSalesInvoiceLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50077, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("COAL Invoice - GST")
            {
                ApplicationArea = all;
                trigger OnAction()
                var
                    RecSalesInvoiceLine: Record "Sales Invoice Line";
                BEGIN
                    //  CAS-22565-T8N8W6 13Aug2018
                    IF (rec."EWB No." = '') AND (rec."EWB Date" = 0D) THEN BEGIN//RSPLSUM 30Jul2020
                        IF rec."Road Permit No." = '' THEN
                            ERROR('E-Way Bill No. cannot be blank');
                    END;//RSPLSUM 30Jul2020
                        //

                    //RSPLSUM 13Jan21


                    IF rec."Posting Date" > 20210113D/*130121D*/ THEN BEGIN
                        rec.CALCFIELDS(IRN);
                        RecSalesInvoiceLine.RESET;
                        RecSalesInvoiceLine.SETRANGE("Document No.", REC."No.");
                        RecSalesInvoiceLine.SETFILTER(Quantity, '%1', 0);
                        //  RecSalesInvoiceLine.SETFILTER("Total GST Amount", '%1', 0); TEMPORY COMMENT 3-jan-23 pcpl-065
                        IF RecSalesInvoiceLine.FINDFIRST THEN BEGIN
                            IF REC.IRN = '' THEN
                                MESSAGE('Please generate IRN for this document');
                        END;
                    END;
                    //RSPLSUM 13Jan21

                    SalesInvHeader.RESET;
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50107, TRUE, TRUE, SalesInvHeader);
                END;
            }
            action("Print E-Way Bill")
            {
                ApplicationArea = all;
                trigger OnAction()
                VAR
                    EWPrint: Report 50118;
                    //  recGstLedEntry: Record 16418;
                    RecSIH: Record 112;
                BEGIN
                    //              {
                    //              recGstLedEntry.RESET;
                    // recGstLedEntry.SETRANGE("Document No.", "No.");
                    // IF recGstLedEntry.FINDFIRST THEN
                    //     REPORT.RUNMODAL(50118, TRUE, TRUE, recGstLedEntry);
                    //              }
                    RecSIH.RESET;
                    RecSIH.SETCURRENTKEY("No.", "No.");
                    RecSIH.SETRANGE("No.", rec."No.");
                    IF RecSIH.FINDFIRST THEN
                        REPORT.RUNMODAL(50118, TRUE, TRUE, RecSIH);
                END;
            }
            action("Print QR code")
            {

                ApplicationArea = ALL;
                Promoted = true;
                PromotedIsBig = true;
                trigger OnAction()

                VAR
                    RecSIH: Record 112;
                BEGIN

                    RecSIH.RESET;
                    RecSIH.SETCURRENTKEY("No.", "No.");
                    RecSIH.SETRANGE("No.", REC."No.");
                    IF RecSIH.FINDFIRST THEN
                        REPORT.RUNMODAL(50120, TRUE, TRUE, RecSIH);
                END;
            }
            action("Print Structure Details-AL")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50137, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Bond Transfer Invoice")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()

                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50181, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("GST Invoice")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()

                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50181, TRUE, FALSE, SalesInvHeader);
                END;
            }
            action("Bond Transfer Invoice-GST")
            {
                ApplicationArea = ALL;
                Image = Print;
                trigger OnAction()

                BEGIN
                    SalesInvHeader.RESET;
                    SalesInvHeader := Rec;
                    SalesInvHeader.SETRECFILTER;
                    REPORT.RUN(50198, TRUE, FALSE, SalesInvHeader);
                END;

            }
            // action("High Seas Invoice")
            // {
            //     Promoted = true;
            //     PromotedIsBig = false;
            //     Image = Print;
            //     PromotedCategory = New;
            //     trigger OnAction()

            //     BEGIN
            //         SalesInvHeader.RESET;
            //         SalesInvHeader := Rec;
            //         SalesInvHeader.SETRECFILTER;
            //         REPORT.RUN(50198, TRUE, FALSE, SalesInvHeader);
            //     END;

            // }
            action("Invoice GST Bunker")
            {
                ApplicationArea = all;
                Promoted = true;
                PromotedIsBig = false;
                Image = Print;
                PromotedCategory = New;
                trigger OnAction()

                BEGIN
                    SalesInvHeader.SETRANGE("No.", REC."No.");
                    IF SalesInvHeader.FINDFIRST THEN
                        REPORT.RUNMODAL(50232, TRUE, TRUE, SalesInvHeader);
                END;

            }
            action("High Seas Invoice")
            {
                ApplicationArea = all;
                Promoted = false;

                // PromotedIsBig = false;
                Image = Print;
                //PromotedCategory = New;

                trigger OnAction()

                VAR
                    SalesInvoiceHeaderRec: Record 112;
                BEGIN
                    SalesInvoiceHeaderRec.RESET;
                    SalesInvoiceHeaderRec.SETRANGE("No.", rec."No.");
                    IF SalesInvoiceHeaderRec.FINDFIRST THEN
                        REPORT.RUNMODAL(50241, TRUE, TRUE, SalesInvoiceHeaderRec);
                END;
            }









        }

    }

    var
        myInt: Integer;

}

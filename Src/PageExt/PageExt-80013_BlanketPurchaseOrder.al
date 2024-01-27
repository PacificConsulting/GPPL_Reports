pageextension 80013 BlanketPOExtCstmRpt extends "Blanket Purchase Order"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {
        addlast(Reporting)
        {
            action("Purchase Order - GST")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    PH: Record 38;
                    Ord05: Record "Order Address";
                    Ven05: Record Vendor;
                begin
                    //RSPLSUM 06Dec19


                    IF rec."Approved by Finance" = FALSE THEN
                        ERROR('Financial approval is required');
                    //RSPLSUM 06Dec19

                    //05Aug2017 GST Calculation with or Without OrdessCode

                    IF Rec."Order Address Code" <> '' THEN BEGIN
                        Ord05.RESET;
                        IF Ord05.GET(Rec."Buy-from Vendor No.", Rec."Order Address Code") THEN BEGIN

                            IF Ord05."GST Registration No." = '' THEN
                                ERROR('Please update GST Regisration Order Address');
                        END;

                    END ELSE BEGIN
                        Ven05.RESET;
                        IF Ven05.GET(Rec."Buy-from Vendor No.") THEN
                            IF Ven05."Gen. Bus. Posting Group" <> 'FOREIGN' THEN BEGIN
                                IF Ven05."GST Vendor Type" = Ven05."GST Vendor Type"::" " THEN
                                    ERROR('Please update the Vendor master with GST Vendor Type');
                            END;

                    END;
                    //

                    //06July2017
                    PH.RESET;
                    PH.SETRANGE("Document Type", Rec."Document Type");
                    PH.SETRANGE("No.", Rec."No.");
                    IF PH.FINDFIRST THEN
                        REPORT.RUNMODAL(50020, TRUE, TRUE, PH);

                    //
                end;
            }
            action("Purchase Order - &Import1")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                var
                    myInt: Integer;
                BEGIN
                    //RSPLSUM 06Dec19


                    IF rec."Approved by Finance" = FALSE THEN
                        ERROR('Financial approval is required');
                    //RSPLSUM 06Dec19

                    CurrPage.SETSELECTIONFILTER(Rec);
                    REPORT.RUNMODAL(50054, TRUE, FALSE, Rec)
                END;
            }
            action("Purchase Order - R&egular Detailed")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                var
                    RecPurchheader: Record "Purchase Header";
                BEGIN
                    RecPurchheader.RESET;
                    RecPurchheader := Rec;
                    RecPurchheader.SETRECFILTER;
                    REPORT.RUNMODAL(50087, TRUE, FALSE, RecPurchheader);
                END;
            }
            action("Purchase Order- GST LOGO")
            {
                ApplicationArea = all;
                Image = Print;
                trigger OnAction()
                VAR
                    PH: Record 38;
                    Ord05: Record "Order Address";
                    Ven05: Record Vendor;
                BEGIN
                    //RSPLSUM 06Dec19


                    IF REC."Approved by Finance" = FALSE THEN
                        ERROR('Financial approval is required');
                    //RSPLSUM 06Dec19

                    //17Nov2017 GST Calculation with or Without OrdessCode

                    IF REC."Order Address Code" <> '' THEN BEGIN
                        Ord05.RESET;
                        IF Ord05.GET(REC."Buy-from Vendor No.", REC."Order Address Code") THEN BEGIN

                            IF Ord05."GST Registration No." = '' THEN
                                ERROR('Please update GST Regisration Order Address');
                        END;

                    END ELSE BEGIN
                        Ven05.RESET;
                        IF Ven05.GET(REC."Buy-from Vendor No.") THEN
                            IF Ven05."Gen. Bus. Posting Group" <> 'FOREIGN' THEN BEGIN
                                IF Ven05."GST Vendor Type" = Ven05."GST Vendor Type"::" " THEN
                                    ERROR('Please update the Vendor master with GST Vendor Type');
                            END;

                    END;
                    //

                    //17Nov2017
                    PH.RESET;
                    PH.SETRANGE("Document Type", REC."Document Type");
                    PH.SETRANGE("No.", REC."No.");
                    IF PH.FINDFIRST THEN
                        REPORT.RUNMODAL(50170, TRUE, TRUE, PH);

                    //
                END;
            }
        }
    }

    var
        myInt: Integer;
}
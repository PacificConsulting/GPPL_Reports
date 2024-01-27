report 50251 "Vehicle packing list new"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/Vehiclepackinglistnew.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Gate Entry Header"; "Gate Entry Header")
        {
            RequestFilterFields = "No.";
            column(No_GateEntryHeader; "Gate Entry Header"."No.")
            {
            }
            column(VehicleNo_GateEntryHeader; "Gate Entry Header"."Vehicle No.")
            {
            }
            column(TransporterName_GateEntryHeader; "Gate Entry Header"."Transporter Name")
            {
            }
            column(DocumentDate_GateEntryHeader; FORMAT("Gate Entry Header"."Document Date"))
            {
            }
            column(CompanyName; CompanyInfo.Name)
            {
            }
            column(LocName; LocationRec.Name)
            {
            }
            column(PreparedByUsername; UserSetupRec.Name)
            {
            }
            dataitem("Gate Entry Line"; "Gate Entry Line")
            {
                DataItemLink = "Gate Entry No." = FIELD("No.");
                column(INVNo; INVNo)
                {
                }
                column(GateEntryNo_GateEntryLine; "Gate Entry Line"."Gate Entry No.")
                {
                }
                column(EWBDate; FORMAT(EWBDate))
                {
                }
                column(EWBNo; EWBNo)
                {
                }
                column(LRNo; LRNo)
                {
                }
                column(CustLocationName; TransferToName)
                {
                }
                column(PostingDate; FORMAT(PostingDate))
                {
                }
                column(QtyDec; QtyDec)
                {
                }
                column(QtyBarrels; QtyBarrels)
                {
                }
                column(Qtyjumbo; Qtyjumbo)
                {
                }
                column(QtyBuckets; QtyBuckets)
                {
                }
                column(QtyIPOL; QtyIPOL)
                {
                }
                column(QtyREPSOL; QtyREPSOL)
                {
                }
                column(QtySample; QtySample)
                {
                }
                column(IndusQtyBarrels; IndusQtyBarrels)
                {
                }
                column(IndusQtyjumbo; IndusQtyjumbo)
                {
                }
                column(IndusQtyBuckets; IndusQtyBuckets)
                {
                }
                column(IndusQtyIPOL; IndusQtyIPOL)
                {
                }
                column(IndusQtyREPSOL; IndusQtyREPSOL)
                {
                }
                column(IndusQtySample; IndusQtySample)
                {
                }
                column(IPOLQtyBarrels; IPOLQtyBarrels)
                {
                }
                column(IPOLQtyjumbo; IPOLQtyjumbo)
                {
                }
                column(IPOLQtyBuckets; IPOLQtyBuckets)
                {
                }
                column(IPOLQtyIPOL; IPOLQtyIPOL)
                {
                }
                column(IPOLQtyREPSOL; IPOLQtyREPSOL)
                {
                }
                column(IPOLQtySample; IPOLQtySample)
                {
                }
                column(REPSOLQtyBarrels; REPSOLQtyBarrels)
                {
                }
                column(REPSOLQtyjumbo; REPSOLQtyjumbo)
                {
                }
                column(REPSOLQtyBuckets; REPSOLQtyBuckets)
                {
                }
                column(REPSOLQtyIPOL; REPSOLQtyIPOL)
                {
                }
                column(REPSOLQtySample; REPSOLQtySample)
                {
                }
                column(REPSOLQtyREPSOL; REPSOLQtyREPSOL)
                {
                }
                column(OtherQtyBarrels; OtherQtyBarrels)
                {
                }
                column(OtherQtyjumbo; OtherQtyjumbo)
                {
                }
                column(OtherQtyBuckets; OtherQtyBuckets)
                {
                }
                column(OtherQtyIPOL; OtherQtyIPOL)
                {
                }
                column(OtherQtyREPSOL; OtherQtyREPSOL)
                {
                }
                column(OtherQtySample; OtherQtySample)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    CLEAR(INVNo);
                    CLEAR(PostingDate);
                    CLEAR(TransferToName);
                    CLEAR(LRNo);
                    CLEAR(EWBNo);
                    CLEAR(EWBDate);


                    IF "Source Type" = "Gate Entry Line"."Source Type"::"Transfer Shipment" THEN
                        TransSHpHdrRec.RESET;
                    //TransSHpHdrRec.SETRANGE(TransSHpHdrRec."No.","Gate Entry Line"."Gate Entry No.");
                    TransSHpHdrRec.SETRANGE(TransSHpHdrRec."No.", "Gate Entry Line"."Source No.");
                    IF TransSHpHdrRec.FINDSET THEN
                        REPEAT
                            INVNo := TransSHpHdrRec."No.";
                            PostingDate := TransSHpHdrRec."Posting Date";
                            TransferToName := TransSHpHdrRec."Transfer-to Name";
                            //LRNo           := TransSHpHdrRec."LR/RR No.";

                            DetailEWB.RESET;
                            DetailEWB.SETRANGE("Document No.", TransSHpHdrRec."No.");
                            DetailEWB.SETRANGE(Cancelled, FALSE);
                            IF DetailEWB.FINDLAST THEN BEGIN
                                EWBNo := DetailEWB."EWB No.";
                                EWBDate := DetailEWB."EWB Valid Upto";
                                LRNo := DetailEWB."Trans. Doc. No.";
                            END;



                            /*
                            CLEAR(QtyDec);
                            CLEAR(QtyBarrels);
                            CLEAR(Qtyjumbo);
                            CLEAR(QtyBuckets);
                            CLEAR(QtyIPOL);
                            CLEAR(QtyREPSOL);
                            */


                            TSLRec.RESET;
                            TSLRec.SETRANGE("Document No.", TransSHpHdrRec."No.");
                            IF TSLRec.FINDFIRST THEN
                                REPEAT

                                    IF ItemRec.GET(TSLRec."Item No.") THEN;

                                    IF TSLRec."Inventory Posting Group" <> 'AUTOOILS' THEN BEGIN
                                        IF (TSLRec."Qty. per Unit of Measure" >= 180) THEN BEGIN
                                            //QtyDec += TSLRec.Quantity;
                                            QtyBarrels += TSLRec.Quantity;

                                        END ELSE
                                            IF (TSLRec."Qty. per Unit of Measure" = 50) THEN BEGIN
                                                //QtyDec += TSLRec.Quantity;
                                                Qtyjumbo += TSLRec.Quantity;

                                            END ELSE
                                                IF ((TSLRec."Qty. per Unit of Measure" >= 6) AND (TSLRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                    //QtyDec += TSLRec.Quantity;
                                                    QtyBuckets += TSLRec.Quantity;

                                                    /*
                                                          END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                            IF (TSLRec."Qty. per Unit of Measure" < 6) THEN
                                                             //QtyDec += TSLRec.Quantity;
                                                             QtyIPOL += TSLRec.Quantity;
                                                    */

                                                END ELSE
                                                    IF TSLRec."Inventory Posting Group" = 'REPSOL' THEN BEGIN
                                                        //IF (TSLRec."Qty. per Unit of Measure" < 6) THEN BEGIN
                                                        //MESSAGE('%1',ItemRec."No of Packages"); //AR
                                                        //QtyDec += TSLRec.Quantity;
                                                        QtyREPSOL += TSLRec.Quantity / ItemRec."No of Packages";
                                                    END;
                                    END ELSE BEGIN
                                        IF TSLRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                            SalesPriceRec.RESET;
                                            SalesPriceRec.SETRANGE("Item No.", ItemRec."No.");
                                            IF SalesPriceRec.FINDLAST THEN BEGIN

                                                IF (SalesPriceRec."Qty. per Pack" >= 180) THEN BEGIN
                                                    //QtyDec += TSLRec.Quantity;
                                                    QtyBarrels += TSLRec.Quantity;

                                                END ELSE
                                                    IF (SalesPriceRec."Qty. per Pack" = 50) THEN BEGIN
                                                        //QtyDec += TSLRec.Quantity;
                                                        Qtyjumbo += TSLRec.Quantity;

                                                    END ELSE
                                                        IF ((SalesPriceRec."Qty. per Pack" >= 6) AND (SalesPriceRec."Qty. per Pack" <= 30)) THEN BEGIN
                                                            //QtyDec += TSLRec.Quantity;
                                                            QtyBuckets += TSLRec.Quantity;

                                                            /*
                                                                  END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                    IF (TSLRec."Qty. per Unit of Measure" < 6) THEN
                                                                     //QtyDec += TSLRec.Quantity;
                                                                     QtyIPOL += TSLRec.Quantity;


                                                                   END ELSE  IF TSLRec."Inventory Posting Group" = 'REPSOL' THEN BEGIN
                                                                    //IF (TSLRec."Qty. per Unit of Measure" < 6) THEN BEGIN
                                                                    //MESSAGE('%1',ItemRec."No of Packages"); //AR
                                                                     //QtyDec += TSLRec.Quantity;
                                                                     QtyREPSOL += TSLRec.Quantity/ItemRec."No of Packages";
                                                                    END;

                                                              */

                                                        END ELSE
                                                            IF (SalesPriceRec."Qty. per Pack" < 6) THEN BEGIN
                                                                QtyIPOL += TSLRec.Quantity;
                                                                //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                            END;
                                            END;
                                        END;

                                        //Sample Begin

                                        IF (TSLRec."Unit of Measure Code" = 'LTR') OR (TSLRec."Unit of Measure Code" = 'LTRS') OR
                                           (TSLRec."Unit of Measure Code" = 'KG') OR (TSLRec."Unit of Measure Code" = 'KGS') THEN BEGIN
                                            QtySample += TSLRec.Quantity;
                                            IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) THEN BEGIN
                                                IndusQtySample += TSLRec.Quantity;

                                            END ELSE
                                                IF (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') THEN BEGIN
                                                    IPOLQtySample += TSLRec.Quantity;

                                                END ELSE
                                                    IF (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') THEN BEGIN
                                                        REPSOLQtySample += TSLRec.Quantity / ItemRec."No of Packages";

                                                    END ELSE
                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND
                                              (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08')) THEN BEGIN
                                                            OtherQtySample += TSLRec.Quantity;
                                                        END;
                                        END;

                                        //Sample End

                                        IF (TSLRec."Unit of Measure Code" <> 'LTR') AND (TSLRec."Unit of Measure Code" <> 'LTRS') AND
                                           (TSLRec."Unit of Measure Code" <> 'KG') AND (TSLRec."Unit of Measure Code" <> 'KGS') THEN BEGIN

                                            IF TSLRec."Inventory Posting Group" <> 'AUTOOILS' THEN BEGIN
                                                //Industrail --DIV-03  begin
                                                IF (((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (TSLRec."Qty. per Unit of Measure" >= 180)) THEN BEGIN
                                                    IndusQtyBarrels += TSLRec.Quantity;
                                                    //IF (TSLRec."Qty. per Unit of Measure" >= 180) THEN BEGIN
                                                    //QtyBarrels += TSLRec.Quantity;

                                                END ELSE
                                                    IF (((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (TSLRec."Qty. per Unit of Measure" = 50)) THEN BEGIN
                                                        IndusQtyjumbo += TSLRec.Quantity;
                                                        //Qtyjumbo += TSLRec.Quantity;

                                                    END ELSE
                                                        IF (((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (TSLRec."Qty. per Unit of Measure" >= 6) AND (TSLRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                            IndusQtyBuckets += TSLRec.Quantity;
                                                            //QtyBuckets += TSLRec.Quantity;
                                                            /*
                                                                END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                  IF (((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-05')) AND (TSLRec."Qty. per Unit of Measure" < 6)) THEN
                                                                   IndusQtyIPOL += TSLRec.Quantity;
                                                                   //QtyIPOL += TSLRec.Quantity;
                                                            */

                                                            /*
                                                            END ELSE IF TSLRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                              IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-05')) THEN BEGIN
                                                                    SalesPriceRec.RESET;
                                                                    SalesPriceRec.SETRANGE("Item No.",ItemRec."No.");
                                                                    IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                      IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                        IndusQtyIPOL += TSLRec.Quantity;
                                                                        //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                    END;
                                                              END;

                                                            */ //AR

                                                        END ELSE
                                                            IF (TSLRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) THEN //BEGIN //AND (TSLRec."Qty. per Unit of Measure" < 6))
                                                                    IndusQtyREPSOL += TSLRec.Quantity / ItemRec."No of Packages";
                                                                //QtyREPSOL += TSLRec.Quantity;
                                                            END;
                                                //Industrail --DIV-08  end


                                                //IPOL-AUtomotive --DIV-04  begin
                                                IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (TSLRec."Qty. per Unit of Measure" >= 180)) THEN BEGIN
                                                    IPOLQtyBarrels += TSLRec.Quantity;
                                                    //IndusQtyBarrels += TSLRec.Quantity;

                                                END ELSE
                                                    IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (TSLRec."Qty. per Unit of Measure" = 50)) THEN BEGIN
                                                        IPOLQtyjumbo += TSLRec.Quantity;
                                                        //IndusQtyjumbo += TSLRec.Quantity;

                                                    END ELSE
                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (TSLRec."Qty. per Unit of Measure" >= 6) AND (TSLRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                            IPOLQtyBuckets += TSLRec.Quantity;
                                                            //QtyBuckets += TSLRec.Quantity;
                                                            /*
                                                                END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                  IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-04') AND (TSLRec."Qty. per Unit of Measure" < 6)) THEN
                                                                   IPOLQtyIPOL += TSLRec.Quantity;
                                                                   //IndusQtyIPOL += TSLRec.Quantity;
                                                                   //QtyIPOL += TSLRec.Quantity;
                                                            */
                                                            /* //AR
                                                            END ELSE IF TSLRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                              IF (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-04') THEN BEGIN
                                                                    SalesPriceRec.RESET;
                                                                    SalesPriceRec.SETRANGE("Item No.",ItemRec."No.");
                                                                    IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                      IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                        IPOLQtyIPOL += TSLRec.Quantity;
                                                                        //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                    END;
                                                              END;

                                                            */


                                                        END ELSE
                                                            IF (TSLRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                IF (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') THEN //BEGIN //AND (TSLRec."Qty. per Unit of Measure" < 6))
                                                                    IPOLQtyREPSOL += TSLRec.Quantity / ItemRec."No of Packages";
                                                                //IndusQtyREPSOL += TSLRec.Quantity;
                                                                //QtyREPSOL += TSLRec.Quantity;
                                                            END;
                                                //IPOL-AUtomotive --DIV-04  begin



                                                //REPSOL--DIV-08  begin
                                                IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (TSLRec."Qty. per Unit of Measure" >= 180)) THEN BEGIN
                                                    REPSOLQtyBarrels += TSLRec.Quantity;
                                                    //IPOLQtyBarrels += TSLRec.Quantity;
                                                    //IndusQtyBarrels += TSLRec.Quantity;

                                                END ELSE
                                                    IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (TSLRec."Qty. per Unit of Measure" = 50)) THEN BEGIN
                                                        REPSOLQtyjumbo += TSLRec.Quantity;
                                                        //IPOLQtyjumbo += TSLRec.Quantity;
                                                        //IndusQtyjumbo += TSLRec.Quantity;

                                                    END ELSE
                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (TSLRec."Qty. per Unit of Measure" >= 6) AND (TSLRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                            REPSOLQtyBuckets += TSLRec.Quantity;
                                                            //IPOLQtyBuckets += TSLRec.Quantity;
                                                            //QtyBuckets += TSLRec.Quantity;
                                                            /*
                                                                END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                  IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-08') AND (TSLRec."Qty. per Unit of Measure" < 6)) THEN
                                                                   REPSOLQtyIPOL += TSLRec.Quantity;
                                                                   //IPOLQtyIPOL += TSLRec.Quantity;
                                                                   //IndusQtyIPOL += TSLRec.Quantity;
                                                                   //QtyIPOL += TSLRec.Quantity;
                                                            */
                                                            /*
                                                            END ELSE IF TSLRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                              IF (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-08') THEN BEGIN
                                                                    SalesPriceRec.RESET;
                                                                    SalesPriceRec.SETRANGE("Item No.",ItemRec."No.");
                                                                    IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                      IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                        REPSOLQtyIPOL += TSLRec.Quantity;
                                                                        //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                    END;
                                                              END;
                                                            */ //AR


                                                        END ELSE
                                                            IF (TSLRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                IF (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') THEN //AND (TSLRec."Qty. per Unit of Measure" < 6)) THEN //BEGIN
                                                                    REPSOLQtyREPSOL += TSLRec.Quantity / ItemRec."No of Packages";
                                                                //IPOLQtyREPSOL += TSLRec.Quantity;
                                                                //IndusQtyREPSOL += TSLRec.Quantity;
                                                                //QtyREPSOL += TSLRec.Quantity;
                                                            END;
                                                //REPSOL--DIV-08  end



                                                //OTHERS--DIV-00  begin
                                                IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                                   (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND (TSLRec."Qty. per Unit of Measure" >= 180)) THEN BEGIN
                                                    OtherQtyBarrels += TSLRec.Quantity;
                                                    //REPSOLQtyBarrels += TSLRec.Quantity;
                                                    //IPOLQtyBarrels += TSLRec.Quantity;
                                                    //IndusQtyBarrels += TSLRec.Quantity;

                                                END ELSE
                                                    IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                          (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND (TSLRec."Qty. per Unit of Measure" = 50)) THEN BEGIN
                                                        OtherQtyjumbo += TSLRec.Quantity;
                                                        //REPSOLQtyjumbo += TSLRec.Quantity;
                                                        //IPOLQtyjumbo += TSLRec.Quantity;
                                                        //IndusQtyjumbo += TSLRec.Quantity;

                                                    END ELSE
                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                              (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND
                                              (TSLRec."Qty. per Unit of Measure" >= 6) AND (TSLRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                            OtherQtyBuckets += TSLRec.Quantity;
                                                            //REPSOLQtyBuckets += TSLRec.Quantity;
                                                            //IPOLQtyBuckets += TSLRec.Quantity;
                                                            //QtyBuckets += TSLRec.Quantity;

                                                            /*
                                                                END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                  IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code"<>'DIV-04') AND
                                                                  (TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-08')
                                                                  AND (TSLRec."Qty. per Unit of Measure" < 6)) THEN
                                                                   OtherQtyIPOL  += TSLRec.Quantity;
                                                                   //REPSOLQtyIPOL += TSLRec.Quantity;
                                                                   //IPOLQtyIPOL += TSLRec.Quantity;
                                                                   //IndusQtyIPOL += TSLRec.Quantity;
                                                                   //QtyIPOL += TSLRec.Quantity;
                                                            */

                                                            /*
                                                            END ELSE IF TSLRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                             IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code"<>'DIV-04') AND
                                                                  (TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-08')) THEN BEGIN
                                                                    SalesPriceRec.RESET;
                                                                    SalesPriceRec.SETRANGE("Item No.",ItemRec."No.");
                                                                    IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                      IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                        OtherQtyIPOL += TSLRec.Quantity;
                                                                        //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                    END;
                                                              END;
                                                            */ //AR

                                                        END ELSE
                                                            IF (TSLRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                                               (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08')) THEN
                                                                    //AND (TSLRec."Qty. per Unit of Measure" < 6)) THEN //BEGIN
                                                                    OtherQtyREPSOL += TSLRec.Quantity / ItemRec."No of Packages";
                                                                //REPSOLQtyREPSOL += TSLRec.Quantity;
                                                                //IPOLQtyREPSOL += TSLRec.Quantity;
                                                                //IndusQtyREPSOL += TSLRec.Quantity;
                                                                //QtyREPSOL += TSLRec.Quantity;
                                                            END;
                                                //OTHERS--DIV-00  end
                                                //END ELSE BEGIN
                                                //END;
                                                //END ELSE BEGIN  //AR


                                            END ELSE
                                                IF TSLRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                    SalesPriceRec.RESET;
                                                    SalesPriceRec.SETRANGE("Item No.", ItemRec."No.");
                                                    IF SalesPriceRec.FINDLAST THEN BEGIN
                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05') AND (SalesPriceRec."Qty. per Pack" < 6)) THEN BEGIN
                                                            IndusQtyIPOL += TSLRec.Quantity;
                                                            //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                            //          END;
                                                            // END;
                                                            //industrial

                                                            //Industrail --DIV-03  begin
                                                            IF (((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (SalesPriceRec."Qty. per Pack" >= 180)) THEN BEGIN
                                                                IndusQtyBarrels += TSLRec.Quantity;
                                                                //IF (SalesPriceRec."Qty. per Pack" >= 180) THEN BEGIN
                                                                //QtyBarrels += TSLRec.Quantity;

                                                            END ELSE
                                                                IF (((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (SalesPriceRec."Qty. per Pack" = 50)) THEN BEGIN
                                                                    IndusQtyjumbo += TSLRec.Quantity;
                                                                    //Qtyjumbo += TSLRec.Quantity;

                                                                END ELSE
                                                                    IF (((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (SalesPriceRec."Qty. per Pack" >= 6) AND (SalesPriceRec."Qty. per Pack" <= 30)) THEN BEGIN
                                                                        IndusQtyBuckets += TSLRec.Quantity;
                                                                        //QtyBuckets += TSLRec.Quantity;
                                                                        /*
                                                                            END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                              IF (((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-05')) AND (SalesPriceRec."Qty. per Pack" < 6)) THEN
                                                                               IndusQtyIPOL += TSLRec.Quantity;
                                                                               //QtyIPOL += TSLRec.Quantity;


                                                                            END ELSE IF TSLRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                                              IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-05')) THEN BEGIN
                                                                                    SalesPriceRec.RESET;
                                                                                    SalesPriceRec.SETRANGE("Item No.",ItemRec."No.");
                                                                                    IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                                      IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                                        IndusQtyIPOL += TSLRec.Quantity;
                                                                                        //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                                    END;
                                                                              END;




                                                                             END ELSE  IF (TSLRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                              IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-05')) THEN //AND (SalesPriceRec."Qty. per Pack" < 6)) THEN //BEGIN
                                                                               IndusQtyREPSOL += TSLRec.Quantity/ItemRec."No of Packages";
                                                                               //QtyREPSOL += TSLRec.Quantity;
                                                                             END;
                                                                        */

                                                                        //industrial

                                                                        //IPOL

                                                                        //IPOL-AUtomotive --DIV-04  begin

                                                                        //IPOL-AUtomotive --DIV-04  begin
                                                                    END ELSE
                                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (SalesPriceRec."Qty. per Pack" >= 180)) THEN BEGIN
                                                                            IPOLQtyBarrels += TSLRec.Quantity;
                                                                            //IndusQtyBarrels += TSLRec.Quantity;

                                                                        END ELSE
                                                                            IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (SalesPriceRec."Qty. per Pack" = 50)) THEN BEGIN
                                                                                IPOLQtyjumbo += TSLRec.Quantity;
                                                                                //IndusQtyjumbo += TSLRec.Quantity;

                                                                            END ELSE
                                                                                IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (SalesPriceRec."Qty. per Pack" >= 6) AND (SalesPriceRec."Qty. per Pack" <= 30)) THEN BEGIN
                                                                                    IPOLQtyBuckets += TSLRec.Quantity;
                                                                                    //QtyBuckets += TSLRec.Quantity;
                                                                                    /*
                                                                                        END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                                          IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-04') AND (SalesPriceRec."Qty. per Pack" < 6)) THEN
                                                                                           IPOLQtyIPOL += TSLRec.Quantity;
                                                                                           //IndusQtyIPOL += TSLRec.Quantity;
                                                                                           //QtyIPOL += TSLRec.Quantity;
                                                                                    */

                                                                                END ELSE
                                                                                    IF (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') THEN BEGIN
                                                                                        IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                                            IPOLQtyIPOL += TSLRec.Quantity;
                                                                                        //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                                        /*
                                                                                             END ELSE  IF (TSLRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                                              IF (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-04') THEN //AND (SalesPriceRec."Qty. per Pack" < 6)) THEN //BEGIN
                                                                                               IPOLQtyREPSOL += TSLRec.Quantity/ItemRec."No of Packages";
                                                                                               //IndusQtyREPSOL += TSLRec.Quantity;
                                                                                               //QtyREPSOL += TSLRec.Quantity;
                                                                                             END;
                                                                                        */
                                                                                        //IPOL-AUtomotive --DIV-04  begin


                                                                                        //END;

                                                                                        //REPSOL

                                                                                        //REPSOL--DIV-08  begin
                                                                                    END ELSE
                                                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (SalesPriceRec."Qty. per Pack" >= 180)) THEN BEGIN
                                                                                            REPSOLQtyBarrels += TSLRec.Quantity;
                                                                                            //IPOLQtyBarrels += TSLRec.Quantity;
                                                                                            //IndusQtyBarrels += TSLRec.Quantity;

                                                                                        END ELSE
                                                                                            IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (SalesPriceRec."Qty. per Pack" = 50)) THEN BEGIN
                                                                                                REPSOLQtyjumbo += TSLRec.Quantity;
                                                                                                //IPOLQtyjumbo += TSLRec.Quantity;
                                                                                                //IndusQtyjumbo += TSLRec.Quantity;

                                                                                            END ELSE
                                                                                                IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (SalesPriceRec."Qty. per Pack" >= 6) AND (SalesPriceRec."Qty. per Pack" <= 30)) THEN BEGIN
                                                                                                    REPSOLQtyBuckets += TSLRec.Quantity;
                                                                                                    //IPOLQtyBuckets += TSLRec.Quantity;
                                                                                                    //QtyBuckets += TSLRec.Quantity;
                                                                                                    /*
                                                                                                        END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                                                          IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-08') AND (SalesPriceRec."Qty. per Pack" < 6)) THEN
                                                                                                           REPSOLQtyIPOL += TSLRec.Quantity;
                                                                                                           //IPOLQtyIPOL += TSLRec.Quantity;
                                                                                                           //IndusQtyIPOL += TSLRec.Quantity;
                                                                                                           //QtyIPOL += TSLRec.Quantity;
                                                                                                    */

                                                                                                END ELSE
                                                                                                    IF (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (SalesPriceRec."Qty. per Pack" < 6) THEN BEGIN
                                                                                                        REPSOLQtyIPOL += TSLRec.Quantity;
                                                                                                        //MESSAGE('%1',SalesPriceRec."Qty. per Pack");

                                                                                                        /*
                                                                                                             END ELSE  IF (TSLRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                                                              IF (TransSHpHdrRec."Shortcut Dimension 1 Code"='DIV-08') THEN //AND (SalesPriceRec."Qty. per Pack" < 6)) THEN //BEGIN
                                                                                                               REPSOLQtyREPSOL += TSLRec.Quantity/ItemRec."No of Packages";
                                                                                                               //IPOLQtyREPSOL += TSLRec.Quantity;
                                                                                                               //IndusQtyREPSOL += TSLRec.Quantity;
                                                                                                               //QtyREPSOL += TSLRec.Quantity;
                                                                                                             END;
                                                                                                        */

                                                                                                        //REPSOL--DIV-08  end



                                                                                                        //OTHERS--DIV-00  begin
                                                                                                    END ELSE
                                                                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                                                                                   (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND (SalesPriceRec."Qty. per Pack" >= 180)) THEN BEGIN
                                                                                                            OtherQtyBarrels += TSLRec.Quantity;
                                                                                                            //REPSOLQtyBarrels += TSLRec.Quantity;
                                                                                                            //IPOLQtyBarrels += TSLRec.Quantity;
                                                                                                            //IndusQtyBarrels += TSLRec.Quantity;

                                                                                                        END ELSE
                                                                                                            IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                                                                                  (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND (SalesPriceRec."Qty. per Pack" = 50)) THEN BEGIN
                                                                                                                OtherQtyjumbo += TSLRec.Quantity;
                                                                                                                //REPSOLQtyjumbo += TSLRec.Quantity;
                                                                                                                //IPOLQtyjumbo += TSLRec.Quantity;
                                                                                                                //IndusQtyjumbo += TSLRec.Quantity;

                                                                                                            END ELSE
                                                                                                                IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                                                                                      (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND (SalesPriceRec."Qty. per Pack" >= 6) AND (SalesPriceRec."Qty. per Pack" <= 30)) THEN BEGIN
                                                                                                                    OtherQtyBuckets += TSLRec.Quantity;
                                                                                                                    //REPSOLQtyBuckets += TSLRec.Quantity;
                                                                                                                    //IPOLQtyBuckets += TSLRec.Quantity;
                                                                                                                    //QtyBuckets += TSLRec.Quantity;
                                                                                                                    /*
                                                                                                                        END ELSE  IF TSLRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                                                                                          IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code"<>'DIV-04') AND
                                                                                                                          (TransSHpHdrRec."Shortcut Dimension 1 Code"<>'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-08') AND (SalesPriceRec."Qty. per Pack" < 6)) THEN
                                                                                                                           OtherQtyIPOL  += TSLRec.Quantity;
                                                                                                                           //REPSOLQtyIPOL += TSLRec.Quantity;
                                                                                                                           //IPOLQtyIPOL += TSLRec.Quantity;
                                                                                                                           //IndusQtyIPOL += TSLRec.Quantity;
                                                                                                                           //QtyIPOL += TSLRec.Quantity;
                                                                                                                    */
                                                                                                                END ELSE
                                                                                                                    IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                                                                                                (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08')) THEN BEGIN
                                                                                                                        IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                                                                            OtherQtyIPOL += TSLRec.Quantity;
                                                                                                                        //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                                                                        /*
                                                                                                                             END ELSE  IF (TSLRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                                                                               IF ((TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code"<>'DIV-04') AND
                                                                                                                                  (TransSHpHdrRec."Shortcut Dimension 1 Code"<>'DIV-05') AND (TransSHpHdrRec."Shortcut Dimension 1 Code"<> 'DIV-08')) THEN //AND (SalesPriceRec."Qty. per Pack" < 6)) THEN //BEGIN
                                                                                                                                  OtherQtyREPSOL  += TSLRec.Quantity/ItemRec."No of Packages";
                                                                                                                               //REPSOLQtyREPSOL += TSLRec.Quantity;
                                                                                                                               //IPOLQtyREPSOL += TSLRec.Quantity;
                                                                                                                               //IndusQtyREPSOL += TSLRec.Quantity;
                                                                                                                               //QtyREPSOL += TSLRec.Quantity;
                                                                                                                             END;
                                                                                                                        //OTHERS--DIV-00  end
                                                                                                                        */

                                                                                                                    END;
                                                        END;

                                                    END;
                                                END;
                                        END;
                                    END;
                                UNTIL TSLRec.NEXT = 0;
                        //MESSAGE('%1',QtyREPSOL);
                        UNTIL TransSHpHdrRec.NEXT = 0;




                    IF ("Source Type" = "Gate Entry Line"."Source Type"::"Sales Shipment") THEN
                        SIHRec.RESET;
                    SIHRec.SETRANGE(SIHRec."No.", "Gate Entry Line"."Source No.");
                    IF SIHRec.FINDSET THEN
                        REPEAT
                            INVNo := SIHRec."No.";
                            PostingDate := SIHRec."Posting Date";
                            TransferToName := SIHRec."Sell-to Customer Name";
                            //LRNo           := SIHRec."LR/RR No.";


                            DetailEWB.RESET;
                            DetailEWB.SETRANGE("Document No.", SIHRec."No.");
                            DetailEWB.SETRANGE(Cancelled, FALSE);
                            IF DetailEWB.FINDLAST THEN BEGIN
                                EWBNo := DetailEWB."EWB No.";
                                EWBDate := DetailEWB."EWB Valid Upto";
                                LRNo := DetailEWB."Trans. Doc. No.";
                            END;

                            /*
                            CLEAR(QtyDec);
                            CLEAR(QtyBarrels);
                            CLEAR(Qtyjumbo);
                            CLEAR(QtyBuckets);
                            CLEAR(QtyIPOL);
                            CLEAR(QtyREPSOL);
                            */

                            SILRec.RESET;
                            SILRec.SETRANGE("Document No.", SIHRec."No.");
                            IF SILRec.FINDSET THEN BEGIN  //Repeat
                                IF SILRec.Type = SILRec.Type::Item THEN
                                    REPEAT

                                        IF ItemRec.GET(SILRec."No.") THEN;

                                        IF (SILRec."Unit of Measure Code" <> 'LTR') AND (SILRec."Unit of Measure Code" <> 'LTRS') AND
                                           (SILRec."Unit of Measure Code" <> 'KG') AND (SILRec."Unit of Measure Code" <> 'KGS') THEN BEGIN

                                            //Industrail --DIV-03  begin
                                            IF (((SIHRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (SIHRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (SILRec."Qty. per Unit of Measure" >= 180)) THEN BEGIN
                                                IndusQtyBarrels += SILRec.Quantity;
                                                //IF (SILRec."Qty. per Unit of Measure" >= 180) THEN BEGIN
                                                //QtyBarrels += SILRec.Quantity;

                                            END ELSE
                                                IF (((SIHRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (SIHRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (SILRec."Qty. per Unit of Measure" = 50)) THEN BEGIN
                                                    IndusQtyjumbo += SILRec.Quantity;
                                                    //Qtyjumbo += SILRec.Quantity;

                                                END ELSE
                                                    IF (((SIHRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (SIHRec."Shortcut Dimension 1 Code" = 'DIV-05')) AND (SILRec."Qty. per Unit of Measure" >= 6) AND (SILRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                        IndusQtyBuckets += SILRec.Quantity;
                                                        //QtyBuckets += SILRec.Quantity;
                                                        /*
                                                            END ELSE  IF SILRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                              IF (((SIHRec."Shortcut Dimension 1 Code"='DIV-03') OR (SIHRec."Shortcut Dimension 1 Code"='DIV-05')) AND (SILRec."Qty. per Unit of Measure" < 6)) THEN
                                                               IndusQtyIPOL += SILRec.Quantity;
                                                               //QtyIPOL += SILRec.Quantity;
                                                        */

                                                    END ELSE
                                                        IF SILRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                            IF ((SIHRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (SIHRec."Shortcut Dimension 1 Code" = 'DIV-05')) THEN BEGIN
                                                                SalesPriceRec.RESET;
                                                                SalesPriceRec.SETRANGE("Item No.", ItemRec."No.");
                                                                IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                    IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                        IndusQtyIPOL += SILRec.Quantity;
                                                                    //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                END;
                                                            END;



                                                        END ELSE
                                                            IF (SILRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                IF ((SIHRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (SIHRec."Shortcut Dimension 1 Code" = 'DIV-05')) THEN //AND (SILRec."Qty. per Unit of Measure" < 6)) THEN //BEGIN
                                                                    IndusQtyREPSOL += SILRec.Quantity / ItemRec."No of Packages";
                                                                //QtyREPSOL += SILRec.Quantity;
                                                            END;
                                            //Industrail --DIV-08  end


                                            //IPOL-AUtomotive --DIV-04  begin
                                            IF ((SIHRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (SILRec."Qty. per Unit of Measure" >= 180)) THEN BEGIN
                                                IPOLQtyBarrels += SILRec.Quantity;
                                                //IndusQtyBarrels += SILRec.Quantity;

                                            END ELSE
                                                IF ((SIHRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (SILRec."Qty. per Unit of Measure" = 50)) THEN BEGIN
                                                    IPOLQtyjumbo += SILRec.Quantity;
                                                    //IndusQtyjumbo += SILRec.Quantity;

                                                END ELSE
                                                    IF ((SIHRec."Shortcut Dimension 1 Code" = 'DIV-04') AND (SILRec."Qty. per Unit of Measure" >= 6) AND (SILRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                        IPOLQtyBuckets += SILRec.Quantity;
                                                        //QtyBuckets += SILRec.Quantity;
                                                        /*
                                                            END ELSE  IF SILRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                              IF ((SIHRec."Shortcut Dimension 1 Code"='DIV-04') AND (SILRec."Qty. per Unit of Measure" < 6)) THEN
                                                               IPOLQtyIPOL += SILRec.Quantity;
                                                               //IndusQtyIPOL += SILRec.Quantity;
                                                               //QtyIPOL += SILRec.Quantity;
                                                        */

                                                    END ELSE
                                                        IF SILRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                            IF (SIHRec."Shortcut Dimension 1 Code" = 'DIV-04') THEN BEGIN
                                                                SalesPriceRec.RESET;
                                                                SalesPriceRec.SETRANGE("Item No.", ItemRec."No.");
                                                                IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                    IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                        IPOLQtyIPOL += SILRec.Quantity;
                                                                    //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                END;
                                                            END;



                                                        END ELSE
                                                            IF (SILRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                IF (SIHRec."Shortcut Dimension 1 Code" = 'DIV-04') THEN //AND (SILRec."Qty. per Unit of Measure" < 6)) THEN //BEGIN
                                                                    IPOLQtyREPSOL += SILRec.Quantity / ItemRec."No of Packages";
                                                                //IndusQtyREPSOL += SILRec.Quantity;
                                                                //QtyREPSOL += SILRec.Quantity;
                                                            END;
                                            //IPOL-AUtomotive --DIV-04  begin



                                            //REPSOL--DIV-08  begin
                                            IF ((SIHRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (SILRec."Qty. per Unit of Measure" >= 180)) THEN BEGIN
                                                REPSOLQtyBarrels += SILRec.Quantity;
                                                //IPOLQtyBarrels += SILRec.Quantity;
                                                //IndusQtyBarrels += SILRec.Quantity;

                                            END ELSE
                                                IF ((SIHRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (SILRec."Qty. per Unit of Measure" = 50)) THEN BEGIN
                                                    REPSOLQtyjumbo += SILRec.Quantity;
                                                    //IPOLQtyjumbo += SILRec.Quantity;
                                                    //IndusQtyjumbo += SILRec.Quantity;

                                                END ELSE
                                                    IF ((SIHRec."Shortcut Dimension 1 Code" = 'DIV-08') AND (SILRec."Qty. per Unit of Measure" >= 6) AND (SILRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                        REPSOLQtyBuckets += SILRec.Quantity;
                                                        //IPOLQtyBuckets += SILRec.Quantity;
                                                        //QtyBuckets += SILRec.Quantity;
                                                        /*
                                                            END ELSE  IF SILRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                              IF ((SIHRec."Shortcut Dimension 1 Code"='DIV-08') AND (SILRec."Qty. per Unit of Measure" < 6)) THEN
                                                               REPSOLQtyIPOL += SILRec.Quantity;
                                                               //IPOLQtyIPOL += SILRec.Quantity;
                                                               //IndusQtyIPOL += SILRec.Quantity;
                                                               //QtyIPOL += SILRec.Quantity;
                                                        */

                                                    END ELSE
                                                        IF SILRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                            IF (SIHRec."Shortcut Dimension 1 Code" = 'DIV-08') THEN BEGIN
                                                                SalesPriceRec.RESET;
                                                                SalesPriceRec.SETRANGE("Item No.", ItemRec."No.");
                                                                IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                    IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                        REPSOLQtyIPOL += SILRec.Quantity;
                                                                    //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                END;
                                                            END;



                                                        END ELSE
                                                            IF (SILRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                IF (SIHRec."Shortcut Dimension 1 Code" = 'DIV-08') THEN //AND (SILRec."Qty. per Unit of Measure" < 6)) THEN //BEGIN
                                                                    REPSOLQtyREPSOL += SILRec.Quantity / ItemRec."No of Packages";
                                                                //IPOLQtyREPSOL += SILRec.Quantity;
                                                                //IndusQtyREPSOL += SILRec.Quantity;
                                                                //QtyREPSOL += SILRec.Quantity;
                                                            END;
                                            //REPSOL--DIV-08  end



                                            //OTHERS--DIV-00  begin
                                            IF ((SIHRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                               (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND (SILRec."Qty. per Unit of Measure" >= 180)) THEN BEGIN
                                                OtherQtyBarrels += SILRec.Quantity;
                                                //REPSOLQtyBarrels += SILRec.Quantity;
                                                //IPOLQtyBarrels += SILRec.Quantity;
                                                //IndusQtyBarrels += SILRec.Quantity;

                                            END ELSE
                                                IF ((SIHRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                      (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND (SILRec."Qty. per Unit of Measure" = 50)) THEN BEGIN
                                                    OtherQtyjumbo += SILRec.Quantity;
                                                    //REPSOLQtyjumbo += SILRec.Quantity;
                                                    //IPOLQtyjumbo += SILRec.Quantity;
                                                    //IndusQtyjumbo += SILRec.Quantity;

                                                END ELSE
                                                    IF ((SIHRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                          (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-08') AND (SILRec."Qty. per Unit of Measure" >= 6) AND (SILRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                        OtherQtyBuckets += SILRec.Quantity;
                                                        //REPSOLQtyBuckets += SILRec.Quantity;
                                                        //IPOLQtyBuckets += SILRec.Quantity;
                                                        //QtyBuckets += SILRec.Quantity;
                                                        /*
                                                            END ELSE  IF SILRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                              IF ((SIHRec."Shortcut Dimension 1 Code"<> 'DIV-03') AND (SIHRec."Shortcut Dimension 1 Code"<>'DIV-04') AND
                                                              (SIHRec."Shortcut Dimension 1 Code"<>'DIV-05') AND (SIHRec."Shortcut Dimension 1 Code"<> 'DIV-08') AND (SILRec."Qty. per Unit of Measure" < 6)) THEN
                                                               OtherQtyIPOL  += SILRec.Quantity;
                                                               //REPSOLQtyIPOL += SILRec.Quantity;
                                                               //IPOLQtyIPOL += SILRec.Quantity;
                                                               //IndusQtyIPOL += SILRec.Quantity;
                                                               //QtyIPOL += SILRec.Quantity;
                                                        */
                                                    END ELSE
                                                        IF SILRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                            IF ((SIHRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                                               (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-08')) THEN BEGIN
                                                                SalesPriceRec.RESET;
                                                                SalesPriceRec.SETRANGE("Item No.", ItemRec."No.");
                                                                IF SalesPriceRec.FINDLAST THEN BEGIN
                                                                    IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                        OtherQtyIPOL += SILRec.Quantity;
                                                                    //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                                END;
                                                            END;


                                                        END ELSE
                                                            IF (SILRec."Inventory Posting Group" = 'REPSOL') THEN BEGIN
                                                                IF ((SIHRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND
                                                                   (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND (SIHRec."Shortcut Dimension 1 Code" <> 'DIV-08')) THEN //AND (SILRec."Qty. per Unit of Measure" < 6)) THEN //BEGIN
                                                                    OtherQtyREPSOL += SILRec.Quantity / ItemRec."No of Packages";
                                                                //REPSOLQtyREPSOL += SILRec.Quantity;
                                                                //IPOLQtyREPSOL += SILRec.Quantity;
                                                                //IndusQtyREPSOL += SILRec.Quantity;
                                                                //QtyREPSOL += SILRec.Quantity;
                                                            END;
                                            //OTHERS--DIV-00  end
                                        END;

                                        IF (SILRec."Qty. per Unit of Measure" >= 180) THEN BEGIN
                                            //QtyDec += SILRec.Quantity;
                                            QtyBarrels += SILRec.Quantity;

                                        END ELSE
                                            IF (SILRec."Qty. per Unit of Measure" = 50) THEN BEGIN
                                                //QtyDec += SILRec.Quantity;
                                                Qtyjumbo += SILRec.Quantity;

                                            END ELSE
                                                IF ((SILRec."Qty. per Unit of Measure" >= 6) AND (SILRec."Qty. per Unit of Measure" <= 30)) THEN BEGIN
                                                    //QtyDec += SILRec.Quantity;
                                                    QtyBuckets += SILRec.Quantity;

                                                    /*
                                                        END ELSE  IF SILRec."Inventory Posting Group" = 'IPOL' THEN BEGIN
                                                          IF (SILRec."Qty. per Unit of Measure" < 6) THEN
                                                           //QtyDec += SILRec.Quantity;
                                                           QtyIPOL += SILRec.Quantity;
                                                    */

                                                END ELSE
                                                    IF SILRec."Inventory Posting Group" = 'AUTOOILS' THEN BEGIN
                                                        //      IF (SIHRec."Shortcut Dimension 1 Code"='DIV-08') THEN BEGIN
                                                        SalesPriceRec.RESET;
                                                        SalesPriceRec.SETRANGE("Item No.", ItemRec."No.");
                                                        IF SalesPriceRec.FINDLAST THEN BEGIN
                                                            IF (SalesPriceRec."Qty. per Pack" < 6) THEN
                                                                QtyIPOL += SILRec.Quantity;
                                                            //MESSAGE('%1',SalesPriceRec."Qty. per Pack");
                                                        END;

                                                    END ELSE
                                                        IF SILRec."Inventory Posting Group" = 'REPSOL' THEN BEGIN
                                                            //IF (SILRec."Qty. per Unit of Measure" < 6) THEN
                                                            //QtyDec += SILRec.Quantity;
                                                            QtyREPSOL += SILRec.Quantity / ItemRec."No of Packages";
                                                        END;

                                        IF (SILRec."Unit of Measure Code" = 'LTR') OR (SILRec."Unit of Measure Code" = 'LTRS') OR
                                           (SILRec."Unit of Measure Code" = 'KG') OR (SILRec."Unit of Measure Code" = 'KGS') THEN BEGIN
                                            QtySample += SILRec.Quantity;
                                            IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-03') OR (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-05')) THEN BEGIN
                                                IndusQtySample += SILRec.Quantity;

                                            END ELSE
                                                IF (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-04') THEN BEGIN
                                                    IPOLQtySample += SILRec.Quantity;

                                                END ELSE
                                                    IF (TransSHpHdrRec."Shortcut Dimension 1 Code" = 'DIV-08') THEN BEGIN
                                                        REPSOLQtySample += SILRec.Quantity;

                                                    END ELSE
                                                        IF ((TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-03') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-05') AND
                                              (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-04') AND (TransSHpHdrRec."Shortcut Dimension 1 Code" <> 'DIV-08')) THEN BEGIN
                                                            OtherQtySample += SILRec.Quantity;
                                                        END;
                                        END;




                                    UNTIL SILRec.NEXT = 0;
                            END;


                        //UNTIL SILRec.NEXT=0;

                        UNTIL SIHRec.NEXT = 0;


                    //END;

                end;
            }


            trigger OnAfterGetRecord()
            begin

                IF LocationRec.GET("Gate Entry Header"."Location Code") THEN
                    // MESSAGE(LocationRec.Name);


                    IF UserSetupRec.GET(USERID) THEN;
            end;

            trigger OnPreDataItem()
            begin
                CompanyInfo.GET;
                CompanyInfo.CALCFIELDS(CompanyInfo.Picture);
            end;
        }
    }





    var
        CompanyInfo: Record 79;
        LocationRec: Record 14;
        SIHRec: Record 112;
        TransSHpHdrRec: Record 5744;
        EWBDate: Text[50];
        EWBNo: Code[20];
        LRNo: Code[20];
        TransferToName: Text[50];
        PostingDate: Date;
        SILRec: Record 113;
        QtyDec: Decimal;
        TSLRec: Record 5745;
        QtyBarrels: Decimal;
        Qtyjumbo: Decimal;
        QtyBuckets: Decimal;
        QtyIPOL: Decimal;
        QtyREPSOL: Decimal;
        QtySample: Decimal;
        UserSetupRec: Record 91;
        INVNo: Code[20];
        DetailEWB: Record 50044;
        "------------------------------": Integer;
        IndusQtyBarrels: Decimal;
        IndusQtyjumbo: Decimal;
        IndusQtyBuckets: Decimal;
        IndusQtyIPOL: Decimal;
        IndusQtyREPSOL: Decimal;
        IndusQtySample: Decimal;
        "----------------------------------": Integer;
        IPOLQtyBarrels: Decimal;
        IPOLQtyjumbo: Decimal;
        IPOLQtyBuckets: Decimal;
        IPOLQtyIPOL: Decimal;
        IPOLQtyREPSOL: Decimal;
        IPOLQtySample: Decimal;
        "---------------------------": Integer;
        REPSOLQtyBarrels: Decimal;
        REPSOLQtyjumbo: Decimal;
        REPSOLQtyBuckets: Decimal;
        REPSOLQtyIPOL: Decimal;
        REPSOLQtyREPSOL: Decimal;
        REPSOLQtySample: Decimal;
        "-------------------------": Integer;
        OtherQtyBarrels: Decimal;
        OtherQtyjumbo: Decimal;
        OtherQtyBuckets: Decimal;
        OtherQtyIPOL: Decimal;
        OtherQtyREPSOL: Decimal;
        OtherQtySample: Decimal;
        ItemRec: Record 27;
        SalesPriceRec: Record 7002;
}


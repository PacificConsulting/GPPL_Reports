report 50105 "Summary of Sales/Transfer"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SummaryofSalesTransfer.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Sales Invoice Line"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL_UOM; "Sales Invoice Line"."Unit of Measure")
            {
            }
            column(SIL_Qty; "Sales Invoice Line".Quantity)
            {
            }
            column(SIL_QtyPer; "Sales Invoice Line"."Qty. per Unit of Measure")
            {
            }
            column(SIL_QtyBase; "Sales Invoice Line"."Quantity (Base)")
            {
            }
            column(SIL_Amt; "Sales Invoice Line".Amount)
            {
            }
            column(SIL_Inventory; "Sales Invoice Line"."Inventory Posting Group")
            {
            }
            column(SIL_Posting; "Sales Invoice Line"."Gen. Bus. Posting Group")
            {
            }
            column(SIL_BEDAmt; 0)// "BED Amount")
            {
            }
            column(SIL_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL_ExAmt; 0)//"Excise Amount")
            {
            }
            column(vCGST; vCGST)
            {
            }
            column(vSGST; vSGST)
            {
            }
            column(vIGST; vIGST)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line"."Document No.");
                recSIH.SETRANGE(recSIH."Supplimentary Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;
                /*//RSPL Sharath
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.","Sales Invoice Line"."Document No.");
                recSIH.SETFILTER(recSIH."Form Code",'%1','I');
                IF recSIH.FINDFIRST THEN
                BEGIN
                 CurrReport.SKIP
                END;
                *///RSPL Sharath
                  //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST);
                CLEAR(vSGST);
                CLEAR(vIGST);
                recDGST.RESET;
                recDGST.SETRANGE("Document No.", "Document No.");
                recDGST.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST.SETRANGE("No.","No.");
                IF recDGST.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST += ABS(recDGST."GST Amount");
                        END;
                        IF recDGST."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST += ABS(recDGST."GST Amount");
                        END;
                        IF recDGST."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST += ABS(recDGST."GST Amount");
                        END;
                    UNTIL recDGST.NEXT = 0;
                END;
                //RSPL Sharath<<<<<

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line".SETRANGE("Sales Invoice Line"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line".SETFILTER("Sales Invoice Line"."Inventory Posting Group", '<>%1', 'AUTOOILS');
                "Sales Invoice Line".SETFILTER("Sales Invoice Line"."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                "Sales Invoice Line".SETRANGE("Sales Invoice Line"."Location Code", LocCode);
                //<<1
            end;
        }
        dataitem("Sales Invoice Line3"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL3_UOM; "Unit of Measure")
            {
            }
            column(SIL3_Qty; Quantity)
            {
            }
            column(SIL3_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(SIL3_QtyBase; "Quantity (Base)")
            {
            }
            column(SIL3_Amt; Amount)
            {
            }
            column(SIL3_Inventory; "Inventory Posting Group")
            {
            }
            column(SIL3_Posting; "Gen. Bus. Posting Group")
            {
            }
            column(SIL3_BEDAmt; 0)// "BED Amount")
            {
            }
            column(SIL3_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL3_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL3_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST1; vCGST1)
            {
            }
            column(vSGST1; vSGST1)
            {
            }
            column(vIGST1; vIGST1)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line3"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line3"."Document No.");
                recSIH.SETRANGE(recSIH."Supplimentary Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line3"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", FALSE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;

                /*//RSPL Sharath
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.","Sales Invoice Line3"."Document No.");
                recSIH.SETFILTER(recSIH."Form Code",'%1','I');
                IF recSIH.FINDFIRST THEN
                BEGIN
                 CurrReport.SKIP
                END;
                *///RSPL Sharath
                  //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST1);
                CLEAR(vSGST1);
                CLEAR(vIGST1);
                recDGST1.RESET;
                recDGST1.SETRANGE("Document No.", "Document No.");
                recDGST1.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST1.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST1.SETRANGE("No.","No.");
                IF recDGST1.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST1."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST1 += ABS(recDGST1."GST Amount");
                        END;
                        IF recDGST1."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST1 += ABS(recDGST1."GST Amount");
                        END;
                        IF recDGST1."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST1 += ABS(recDGST1."GST Amount");
                        END;
                    UNTIL recDGST1.NEXT = 0;
                END;
                //RSPL Sharath<<<<<

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line3".SETRANGE("Sales Invoice Line3"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line3".SETFILTER("Sales Invoice Line3"."Inventory Posting Group", '<>%1', 'AUTOOILS');
                "Sales Invoice Line3".SETFILTER("Sales Invoice Line3"."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                "Sales Invoice Line3".SETRANGE("Sales Invoice Line3"."Location Code", LocCode);
                //<<1
            end;
        }
        dataitem("Sales Invoice Line5"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL5_UOM; "Unit of Measure")
            {
            }
            column(SIL5_Qty; Quantity)
            {
            }
            column(SIL5_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(SIL5_QtyBase; "Quantity (Base)")
            {
            }
            column(SIL5_Amt; Amount)
            {
            }
            column(SIL5_Inventory; "Inventory Posting Group")
            {
            }
            column(SIL5_Posting; "Gen. Bus. Posting Group")
            {
            }
            column(SIL5_BEDAmt; 0)//"BED Amount")
            {
            }
            column(SIL5_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL5_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL5_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST2; vCGST2)
            {
            }
            column(vSGST2; vSGST2)
            {
            }
            column(vIGST2; vIGST2)
            {
            }

            trigger OnAfterGetRecord()
            begin


                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line5"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line5"."Document No.");
                recSIH.SETRANGE(recSIH."Supplimentary Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line5"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;


                /*//RSPL Sharath
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.","Sales Invoice Line5"."Document No.");
                recSIH.SETFILTER(recSIH."Form Code",'<>%1','I');
                IF recSIH.FINDFIRST THEN
                BEGIN
                 CurrReport.SKIP
                END;
                *///RSPL Sharath
                  //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST2);
                CLEAR(vSGST2);
                CLEAR(vIGST2);
                recDGST2.RESET;
                recDGST2.SETRANGE("Document No.", "Document No.");
                recDGST2.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST2.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST2.SETRANGE("No.","No.");
                IF recDGST2.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST2."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST2 += ABS(recDGST2."GST Amount");
                        END;
                        IF recDGST2."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST2 += ABS(recDGST2."GST Amount");
                        END;
                        IF recDGST2."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST2 += ABS(recDGST2."GST Amount");
                        END;
                    UNTIL recDGST2.NEXT = 0;
                END;
                //RSPL Sharath<<<<<

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line5".SETRANGE("Sales Invoice Line5"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line5".SETFILTER("Sales Invoice Line5"."Inventory Posting Group", '<>%1', 'AUTOOILS');
                "Sales Invoice Line5".SETFILTER("Sales Invoice Line5"."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                "Sales Invoice Line5".SETRANGE("Sales Invoice Line5"."Location Code", LocCode);
                //<<1
            end;
        }
        dataitem("Sales Invoice Line0"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL0_UOM; "Unit of Measure")
            {
            }
            column(SIL0_Qty; Quantity)
            {
            }
            column(SIL0_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(SIL0_QtyBase; "Quantity (Base)")
            {
            }
            column(SIL0_Amt; Amount)
            {
            }
            column(Amt_00; Amt)
            {
            }
            column(SIL0_Inventory; "Inventory Posting Group")
            {
            }
            column(SIL0_Posting; "Gen. Bus. Posting Group")
            {
            }
            column(SIL0_BEDAmt; 0)// "BED Amount")
            {
            }
            column(SIL0_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL0_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL0_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST3; vCGST3)
            {
            }
            column(vSGST3; vSGST3)
            {
            }
            column(vIGST3; vIGST3)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line0"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line0"."Document No.");
                recSIH.SETRANGE(recSIH."Supplimentary Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line0"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;

                recSIH.RESET;
                recSIH.SETFILTER(recSIH."Currency Code", '<>%1', '');
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line0"."Document No.");
                IF recSIH.FINDFIRST THEN BEGIN
                    Amt := "Sales Invoice Line0".Amount / recSIH."Currency Factor";
                END ELSE
                    Amt := "Sales Invoice Line0".Amount;
                //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST3);
                CLEAR(vSGST3);
                CLEAR(vIGST3);
                recDGST3.RESET;
                recDGST3.SETRANGE("Document No.", "Document No.");
                recDGST3.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST3.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST3.SETRANGE("No.","No.");
                IF recDGST3.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST3."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST3 += ABS(recDGST3."GST Amount");
                        END;
                        IF recDGST3."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST3 += ABS(recDGST3."GST Amount");
                        END;
                        IF recDGST3."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST3 += ABS(recDGST3."GST Amount");
                        END;
                    UNTIL recDGST3.NEXT = 0;
                END;
                //RSPL Sharath<<<<<
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line0".SETRANGE("Sales Invoice Line0"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line0".SETFILTER("Sales Invoice Line0"."Inventory Posting Group", '%1|%2|%3|%4|%5', 'INDOILS', 'RPO', 'BOILS', 'TOILS',
                'REPSOL');//02May2017
                //"Sales Invoice Line0".SETFILTER("Sales Invoice Line0"."Inventory Posting Group",'%1|%2|%3|%4%5','INDOILS','RPO','BOILS','TOILS',
                //'REPSOL');
                "Sales Invoice Line0".SETFILTER("Sales Invoice Line0"."Gen. Bus. Posting Group", '%1', 'FOREIGN');
                "Sales Invoice Line0".SETRANGE("Sales Invoice Line0"."Location Code", LocCode);
                CurrReport.CREATETOTALS(Amt);
                //<<1
            end;
        }
        dataitem("Sales Invoice Line1"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL1_UOM; "Unit of Measure")
            {
            }
            column(SIL1_Qty; Quantity)
            {
            }
            column(SIL1_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(SIL1_QtyBase; "Quantity (Base)")
            {
            }
            column(SIL1_Amt; Amount)
            {
            }
            column(SIL1_Inventory; "Inventory Posting Group")
            {
            }
            column(SIL1_Posting; "Gen. Bus. Posting Group")
            {
            }
            column(SIL1_BEDAmt; 0)// "BED Amount")
            {
            }
            column(SIL1_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL1_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL1_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST4; vCGST4)
            {
            }
            column(vSGST4; vSGST4)
            {
            }
            column(vIGST4; vIGST4)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line1"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line1"."Document No.");
                recSIH.SETRANGE(recSIH."Supplimentary Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line1"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;

                /*//RSPL Sharath
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.","Sales Invoice Line1"."Document No.");
                recSIH.SETFILTER(recSIH."Form Code",'%1','I');
                IF recSIH.FINDFIRST THEN
                BEGIN
                 CurrReport.SKIP
                END;
                *///RSPL Sharath
                  //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST4);
                CLEAR(vSGST4);
                CLEAR(vIGST4);
                recDGST4.RESET;
                recDGST4.SETRANGE("Document No.", "Document No.");
                recDGST4.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST4.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST4.SETRANGE("No.","No.");
                IF recDGST4.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST4."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST4 += ABS(recDGST4."GST Amount");
                        END;
                        IF recDGST4."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST4 += ABS(recDGST4."GST Amount");
                        END;
                        IF recDGST4."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST4 += ABS(recDGST4."GST Amount");
                        END;
                    UNTIL recDGST4.NEXT = 0;
                END;
                //RSPL Sharath<<<<<

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line1".SETRANGE("Sales Invoice Line1"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line1".SETFILTER("Sales Invoice Line1"."Inventory Posting Group", '%1', 'AUTOOILS');
                "Sales Invoice Line1".SETFILTER("Sales Invoice Line1"."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                "Sales Invoice Line1".SETRANGE("Sales Invoice Line1"."Location Code", LocCode);
                //<<1
            end;
        }
        dataitem("Sales Invoice Line4"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL4_UOM; "Unit of Measure")
            {
            }
            column(SIL4_Qty; Quantity)
            {
            }
            column(SIL4_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(SIL4_QtyBase; "Quantity (Base)")
            {
            }
            column(SIL4_Amt; Amount)
            {
            }
            column(SIL4_Inventory; "Inventory Posting Group")
            {
            }
            column(SIL4_Posting; "Gen. Bus. Posting Group")
            {
            }
            column(SIL4_BEDAmt; 0)//"BED Amount")
            {
            }
            column(SIL4_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL4_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL4_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST5; vCGST5)
            {
            }
            column(vSGST5; vSGST5)
            {
            }
            column(vIGST5; vIGST5)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line4"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line4"."Document No.");
                recSIH.SETRANGE(recSIH."Supplimentary Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line4"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", FALSE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;

                /*//RSPL Sharath
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.","Sales Invoice Line4"."Document No.");
                recSIH.SETFILTER(recSIH."Form Code",'%1','I');
                IF recSIH.FINDFIRST THEN
                BEGIN
                 CurrReport.SKIP
                END;
                *///RSPL Sharath
                  //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST5);
                CLEAR(vSGST5);
                CLEAR(vIGST5);
                recDGST5.RESET;
                recDGST5.SETRANGE("Document No.", "Document No.");
                recDGST5.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST5.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST5.SETRANGE("No.","No.");
                IF recDGST5.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST5."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST5 += ABS(recDGST5."GST Amount");
                        END;
                        IF recDGST5."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST5 += ABS(recDGST5."GST Amount");
                        END;
                        IF recDGST5."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST5 += ABS(recDGST5."GST Amount");
                        END;
                    UNTIL recDGST5.NEXT = 0;
                END;
                //RSPL Sharath<<<<<

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line4".SETRANGE("Sales Invoice Line4"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line4".SETFILTER("Sales Invoice Line4"."Inventory Posting Group", '%1', 'AUTOOILS');
                "Sales Invoice Line4".SETFILTER("Sales Invoice Line4"."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                "Sales Invoice Line4".SETRANGE("Sales Invoice Line4"."Location Code", LocCode);
                //<<1
            end;
        }
        dataitem("Sales Invoice Line6"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL6_UOM; "Unit of Measure")
            {
            }
            column(SIL6_Qty; Quantity)
            {
            }
            column(SIL6_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(SIL6_QtyBase; "Quantity (Base)")
            {
            }
            column(SIL6_Amt; Amount)
            {
            }
            column(SIL6_Inventory; "Inventory Posting Group")
            {
            }
            column(SIL6_Posting; "Gen. Bus. Posting Group")
            {
            }
            column(SIL6_BEDAmt; 0)// "BED Amount")
            {
            }
            column(SIL6_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL6_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL6_ExAmt; 0)//"Excise Amount")
            {
            }
            column(vCGST6; vCGST6)
            {
            }
            column(vSGST6; vSGST6)
            {
            }
            column(vIGST6; vIGST6)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line6"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line6"."Document No.");
                recSIH.SETRANGE(recSIH."Supplimentary Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line6"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;

                /*//RSPL Sharath
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.","Sales Invoice Line6"."Document No.");
                recSIH.SETFILTER(recSIH."Form Code",'<>%1','I');
                IF recSIH.FINDFIRST THEN
                BEGIN
                 CurrReport.SKIP
                END;
                *///RSPL Sharath
                  //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST6);
                CLEAR(vSGST6);
                CLEAR(vIGST6);
                recDGST6.RESET;
                recDGST6.SETRANGE("Document No.", "Document No.");
                recDGST6.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST6.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST6.SETRANGE("No.","No.");
                IF recDGST6.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST6."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST6 += ABS(recDGST6."GST Amount");
                        END;
                        IF recDGST6."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST6 += ABS(recDGST6."GST Amount");
                        END;
                        IF recDGST6."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST6 += ABS(recDGST6."GST Amount");
                        END;
                    UNTIL recDGST6.NEXT = 0;
                END;
                //RSPL Sharath<<<<<

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line6".SETRANGE("Sales Invoice Line6"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line6".SETFILTER("Sales Invoice Line6"."Inventory Posting Group", '%1', 'AUTOOILS');
                "Sales Invoice Line6".SETFILTER("Sales Invoice Line6"."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                "Sales Invoice Line6".SETRANGE("Sales Invoice Line6"."Location Code", LocCode);
                //<<1
            end;
        }
        dataitem("Sales Invoice Line2"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL2_UOM; "Unit of Measure")
            {
            }
            column(SIL2_Qty; Quantity)
            {
            }
            column(SIL2_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(SIL2_QtyBase; "Quantity (Base)")
            {
            }
            column(SIL2_Amt; Amount)
            {
            }
            column(Amt_22; Amt)
            {
            }
            column(SIL2_Inventory; "Inventory Posting Group")
            {
            }
            column(SIL2_Posting; "Gen. Bus. Posting Group")
            {
            }
            column(SIL2_BEDAmt; 0)// "BED Amount")
            {
            }
            column(SIL2_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL2_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL2_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST7; vCGST7)
            {
            }
            column(vSGST7; vSGST7)
            {
            }
            column(vIGST7; vIGST7)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line2"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line2"."Document No.");
                recSIH.SETRANGE(recSIH."Supplimentary Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line2"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;
                //<<1

                //>>2

                //Sales Invoice Line2, GroupFooter (1) - OnPreSection()
                //IF (("Sales Invoice Line2"."Inventory Posting Group" = 'AUTOOILS') AND
                // ("Sales Invoice Line2"."Gen. Bus. Posting Group" = 'FOREIGN') AND ("Sales Invoice Line2".Quantity <> 0)) THEN
                //BEGIN

                recSIH.GET("Sales Invoice Line0"."Document No.");
                IF recSIH."Currency Code" <> '' THEN BEGIN
                    Amt := (("Sales Invoice Line2"."Unit Price" / "Sales Invoice Line2"."Qty. per Unit of Measure") / recSIH."Currency Factor")
                           * "Sales Invoice Line2"."Quantity (Base)";
                END ELSE
                    Amt := "Sales Invoice Line2".Amount;

                //<<2

                //RSPL Sharath>>>>>
                CLEAR(vCGST7);
                CLEAR(vSGST7);
                CLEAR(vIGST7);
                recDGST7.RESET;
                recDGST7.SETRANGE("Document No.", "Document No.");
                recDGST7.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST7.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST7.SETRANGE("No.","No.");
                IF recDGST7.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST7."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST7 += ABS(recDGST7."GST Amount");
                        END;
                        IF recDGST7."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST7 += ABS(recDGST7."GST Amount");
                        END;
                        IF recDGST7."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST7 += ABS(recDGST7."GST Amount");
                        END;
                    UNTIL recDGST7.NEXT = 0;
                END;
                //RSPL Sharath<<<<<
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line2".SETRANGE("Sales Invoice Line2"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line2".SETFILTER("Sales Invoice Line2"."Inventory Posting Group", '%1', 'AUTOOILS');
                "Sales Invoice Line2".SETFILTER("Sales Invoice Line2"."Gen. Bus. Posting Group", '%1', 'FOREIGN');
                "Sales Invoice Line2".SETRANGE("Sales Invoice Line2"."Location Code", LocCode);
                //<<1
            end;
        }
        dataitem("Sales Invoice Line7"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Type = FILTER(Item),
                                      Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(SIL7_UOM; "Unit of Measure")
            {
            }
            column(SIL7_Qty; Quantity)
            {
            }
            column(SIL7_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(SIL7_QtyBase; "Quantity (Base)")
            {
            }
            column(SIL7_Amt; Amount)
            {
            }
            column(SIL7_Inventory; "Inventory Posting Group")
            {
            }
            column(SIL7_Posting; "Gen. Bus. Posting Group")
            {
            }
            column(SIL7_BEDAmt; 0)//"BED Amount")
            {
            }
            column(SIL7_eAmt; 0)// "eCess Amount")
            {
            }
            column(SIL7_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(SIL7_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST8; vCGST8)
            {
            }
            column(vSGST8; vSGST8)
            {
            }
            column(vIGST8; vIGST8)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line"."Document No.");
                recSIH.SETRANGE(recSIH."Cancelled Invoice", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP
                END;

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Line"."Document No.");
                recSIH.SETRANGE(recSIH."CT3 Order", TRUE);
                IF recSIH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;

                /*//RSPL Sharath
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.","Sales Invoice Line"."Document No.");
                recSIH.SETFILTER(recSIH."Form Code",'%1','I');
                IF recSIH.FINDFIRST THEN
                BEGIN
                 CurrReport.SKIP
                END;
                *///RSPL Sharath
                  //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST8);
                CLEAR(vSGST8);
                CLEAR(vIGST8);
                recDGST8.RESET;
                recDGST8.SETRANGE("Document No.", "Document No.");
                recDGST8.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST8.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST8.SETRANGE("No.","No.");
                IF recDGST8.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST8."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST8 += ABS(recDGST8."GST Amount");
                        END;
                        IF recDGST8."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST8 += ABS(recDGST8."GST Amount");
                        END;
                        IF recDGST8."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST8 += ABS(recDGST8."GST Amount");
                        END;
                    UNTIL recDGST8.NEXT = 0;
                END;
                //RSPL Sharath<<<<<

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                "Sales Invoice Line7".SETRANGE("Sales Invoice Line7"."Posting Date", StartDate, EndDate);
                "Sales Invoice Line7".SETFILTER("Sales Invoice Line7"."Inventory Posting Group", '%1', '');
                "Sales Invoice Line7".SETFILTER("Sales Invoice Line7"."Gen. Bus. Posting Group", '%1', 'DOMESTIC');
                "Sales Invoice Line7".SETRANGE("Sales Invoice Line7"."Location Code", LocCode);
                //<<1
            end;
        }
        dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(TSL_UOM; "Unit of Measure")
            {
            }
            column(TSL_Qty; Quantity)
            {
            }
            column(TSL_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(TSL_QtyBase; "Quantity (Base)")
            {
            }
            column(TSL_Amt; Amount)
            {
            }
            column(TSL_Inventory; "Inventory Posting Group")
            {
            }
            column(TSL_BEDAmt; 0)// "BED Amount")
            {
            }
            column(TSL_eAmt; 0)// "eCess Amount")
            {
            }
            column(TSL_SAmt; 0)//"SHE Cess Amount")
            {
            }
            column(TSL_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST9; vCGST9)
            {
            }
            column(vSGST9; vSGST9)
            {
            }
            column(vIGST9; vIGST9)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                IF "Transfer Shipment Line"."Transfer-to Code" = 'CAPCON' THEN BEGIN
                    CurrReport.SKIP;
                END;

                recTSH.RESET;
                recTSH.SETRANGE(recTSH."No.", "Transfer Shipment Line"."Document No.");
                recTSH.SETRANGE(recTSH."BRT Cancelled", TRUE);
                IF recTSH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;
                //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST9);
                CLEAR(vSGST9);
                CLEAR(vIGST9);
                recDGST9.RESET;
                recDGST9.SETRANGE("Document No.", "Document No.");
                recDGST9.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST9.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST9.SETRANGE("No.","Transfer Shipment Line"."Item No.");
                IF recDGST9.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST9."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST9 += ABS(recDGST9."GST Amount");
                        END;
                        IF recDGST9."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST9 += ABS(recDGST9."GST Amount");
                        END;
                        IF recDGST9."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST9 += ABS(recDGST9."GST Amount");
                        END;
                    UNTIL recDGST9.NEXT = 0;
                END;
                //RSPL Sharath<<<<<
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                //EBT MILAN....(020913)..To Run Report Base On Location Filter Given in Request Form.................................START
                IF LocCode = 'PLANT01' THEN BEGIN
                    "Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Shipment Date", StartDate, EndDate);
                    "Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Transfer-from Code", LocCode);
                    "Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Inventory Posting Group",
                    '%1|%2|%3|%4|%5', 'INDOILS', 'RPO', 'BOILS', 'TOILS', 'REPSOL');
                    //IF EndDate<310315D THEN
                    //"Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Document No.",'%1','@*BRT*');
                    "Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Document No.", '%1|%2|%3', '@*BRT*', '*T/*', '*C/*');//01Nov2017
                END;

                IF LocCode = 'PLANT02' THEN BEGIN
                    "Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Shipment Date", StartDate, EndDate);
                    "Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Transfer-from Code", LocCode);
                    "Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Inventory Posting Group",
                    '%1|%2|%3|%4|%5', 'INDOILS', 'RPO', 'BOILS', 'TOILS', 'REPSOL');
                    //IF EndDate<310315D THEN
                    //"Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Document No.",'%1','@*BRT*');
                    "Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Document No.", '%1|%2|%3', '@*BRT*', '*T/*', '*C/*');//01Nov2017
                END;

                IF NOT ((LocCode = 'PLANT01') OR (LocCode = 'PLANT02')) THEN BEGIN
                    "Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Shipment Date", StartDate, EndDate);
                    "Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Transfer-from Code", LocCode);
                    "Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Inventory Posting Group",
                    '%1|%2|%3|%4|%5', 'INDOILS', 'RPO', 'BOILS', 'TOILS', 'REPSOL');
                END;
                //EBT MILAN....(020913)..To Run Report Base On Location Filter Given in Request Form.................................END
                //<<1
            end;
        }
        dataitem("Transfer Shipment Line0"; "Transfer Shipment Line")
        {
            DataItemTableView = SORTING("Unit of Measure")
                                WHERE(Quantity = FILTER(<> 0),
                                      "Inventory Posting Group" = FILTER(<> 'MERCH'));
            column(TSL0_UOM; "Unit of Measure")
            {
            }
            column(TSL0_Qty; Quantity)
            {
            }
            column(TSL0_QtyPer; "Qty. per Unit of Measure")
            {
            }
            column(TSL0_QtyBase; "Quantity (Base)")
            {
            }
            column(TSL0_Amt; Amount)
            {
            }
            column(TSL0_Inventory; "Inventory Posting Group")
            {
            }
            column(TSL0_BEDAmt; 0)// "BED Amount")
            {
            }
            column(TSL0_eAmt; 0)// "eCess Amount")
            {
            }
            column(TSL0_SAmt; 0)// "SHE Cess Amount")
            {
            }
            column(TSL0_ExAmt; 0)// "Excise Amount")
            {
            }
            column(vCGST10; vCGST10)
            {
            }
            column(vSGST10; vSGST10)
            {
            }
            column(vIGST10; vIGST10)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                IF "Transfer Shipment Line0"."Transfer-to Code" = 'CAPCON' THEN BEGIN
                    CurrReport.SKIP;
                END;

                recTSH.RESET;
                recTSH.SETRANGE(recTSH."No.", "Transfer Shipment Line0"."Document No.");
                recTSH.SETRANGE(recTSH."BRT Cancelled", TRUE);
                IF recTSH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP;
                END;

                //<<1

                //RSPL Sharath>>>>>
                CLEAR(vCGST10);
                CLEAR(vSGST10);
                CLEAR(vIGST10);
                recDGST10.RESET;
                recDGST10.SETRANGE("Document No.", "Document No.");
                recDGST10.SETRANGE("Document Line No.", "Line No.");// RB-N 26Oct2017
                recDGST10.SETRANGE("HSN/SAC Code", "HSN/SAC Code");// RB-N 26Oct2017
                //recDGST10.SETRANGE("No.","Transfer Shipment Line0"."Item No.");
                IF recDGST10.FINDFIRST THEN BEGIN
                    REPEAT
                        IF recDGST10."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST10 += ABS(recDGST10."GST Amount");
                        END;
                        IF recDGST10."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST10 += ABS(recDGST10."GST Amount");
                        END;
                        IF recDGST10."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST10 += ABS(recDGST10."GST Amount");
                        END;
                    UNTIL recDGST10.NEXT = 0;
                END;
                //RSPL Sharath<<<<<
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                //EBT MILAN....(020913)..To Run Report Base On Location Filter Given in Request Form.................................START
                IF LocCode = 'PLANT01' THEN BEGIN
                    "Transfer Shipment Line0".SETRANGE("Transfer Shipment Line0"."Shipment Date", StartDate, EndDate);
                    "Transfer Shipment Line0".SETRANGE("Transfer Shipment Line0"."Transfer-from Code", LocCode);
                    "Transfer Shipment Line0".SETFILTER("Transfer Shipment Line0"."Inventory Posting Group", '%1', 'AUTOOILS');
                    //IF EndDate<310315D THEN
                    //"Transfer Shipment Line0".SETFILTER("Transfer Shipment Line0"."Document No.",'%1','@*BRT*')
                    "Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Document No.", '%1|%2|%3', '@*BRT*', '*T/*', '*C/*');//01Nov2017
                END;

                IF LocCode = 'PLANT02' THEN BEGIN
                    "Transfer Shipment Line0".SETRANGE("Transfer Shipment Line0"."Shipment Date", StartDate, EndDate);
                    "Transfer Shipment Line0".SETRANGE("Transfer Shipment Line0"."Transfer-from Code", LocCode);
                    "Transfer Shipment Line0".SETFILTER("Transfer Shipment Line0"."Inventory Posting Group", '%1', 'AUTOOILS');
                    //IF EndDate<310315D THEN
                    //"Transfer Shipment Line0".SETFILTER("Transfer Shipment Line0"."Document No.",'%1','@*BRT*')
                    "Transfer Shipment Line".SETFILTER("Transfer Shipment Line"."Document No.", '%1|%2|%3', '@*BRT*', '*T/*', '*C/*');//01Nov2017
                END;

                IF NOT ((LocCode = 'PLANT01') OR (LocCode = 'PLANT02')) THEN BEGIN
                    "Transfer Shipment Line0".SETRANGE("Transfer Shipment Line0"."Shipment Date", StartDate, EndDate);
                    "Transfer Shipment Line0".SETRANGE("Transfer Shipment Line0"."Transfer-from Code", LocCode);
                    "Transfer Shipment Line0".SETFILTER("Transfer Shipment Line0"."Inventory Posting Group", '%1', 'AUTOOILS');
                END;
                //EBT MILAN....(020913)..To Run Report Base On Location Filter Given in Request Form.................................END
                //<<1
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; StartDate)
                {
                    ApplicationArea = ALL;
                }
                field("End Date"; EndDate)
                {
                    ApplicationArea = ALL;
                }
                field("Location Code"; LocCode)
                {
                    ApplicationArea = ALL;
                    TableRelation = Location;
                }
            }
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPostReport()
    begin

        //>>1

        //XLSHEET.Columns.AutoFit;
        //MESSAGE('Report Finished');
        //XLAPP.Visible(TRUE);
        //<<1

        //>>RB-N 26Oct2017 ExcelBook Creation
        IF printtoexcel THEN BEGIN
            /* ExcBuffer.CreateBook('', 'Summary of Sales/Transfer');
            ExcBuffer.CreateBookAndOpenExcel('', 'Summary of Sales/Transfer', '', '', USERID);
            ExcBuffer.GiveUserControl; */

            ExcBuffer.CreateNewBook('Summary of Sales/Transfer');
            ExcBuffer.WriteSheet('Summary of Sales/Transfer', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Summary of Sales/Transfer', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();

        END;
        //>>RB-N 26Oct2017 ExcelBook Creation
    end;

    trigger OnPreReport()
    begin

        //>>1

        /*
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",UPPERCASE(USERID));
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2|%3','SUPER','REPORT VIEW','PENDING-INDENT/CSO');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.RESET;
         recLocation.SETRANGE(recLocation.Code,LocCode);
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping1.RESET;
         CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
         CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
         CSOMapping1.SETRANGE(CSOMapping1.Value,LocCode);
         IF NOT(CSOMapping1.FINDFIRST) THEN
         BEGIN
          CSOMapping1.RESET;
          CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
          CSOMapping1.SETRANGE(CSOMapping1.Type,CSOMapping1.Type::"Responsibility Center");
          CSOMapping1.SETRANGE(CSOMapping1.Value,LocResC);
          IF NOT(CSOMapping1.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        *///Commented 02May2017

        CompInfo.GET;

        //CreateXLSHEET;
        //CreateHeader;
        i := 5;

        //<<1

        //>>RB-N 26Oct2017 ExcelHeader

        EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '', 1);
        EnterCell(2, 1, 'Summary of Sales/Transfer', TRUE, FALSE, '', 1);
        EnterCell(3, 1, 'Date Filter :', TRUE, FALSE, '', 1);
        EnterCell(3, 2, FORMAT(StartDate) + ' to ' + FORMAT(EndDate), TRUE, FALSE, '', 1);

        EnterCell(5, 1, 'UOM', TRUE, TRUE, '', 1);
        EnterCell(5, 2, 'Qty', TRUE, TRUE, '', 1);
        EnterCell(5, 3, 'Qty per UOM', TRUE, TRUE, '', 1);
        EnterCell(5, 4, 'Qty in Ltrs/Kgs', TRUE, TRUE, '', 1);
        EnterCell(5, 5, 'Net Amount', TRUE, TRUE, '', 1);
        EnterCell(5, 6, 'Inventory Posting Group', TRUE, TRUE, '', 1);
        EnterCell(5, 7, 'Gen. Bus. Posting Group', TRUE, TRUE, '', 1);
        EnterCell(5, 8, 'Type', TRUE, TRUE, '', 1);
        EnterCell(5, 9, '', TRUE, TRUE, '', 1);
        EnterCell(5, 10, 'BED Amount', TRUE, TRUE, '', 1);
        EnterCell(5, 11, 'Ecess Amount', TRUE, TRUE, '', 1);
        EnterCell(5, 12, 'SHE cess Amount', TRUE, TRUE, '', 1);
        EnterCell(5, 13, 'Excise Amount', TRUE, TRUE, '', 1);
        EnterCell(5, 14, 'CGST Amount', TRUE, TRUE, '', 1);
        EnterCell(5, 15, 'SGST Amount', TRUE, TRUE, '', 1);
        EnterCell(5, 16, 'IGST Amount', TRUE, TRUE, '', 1);

        //>>RB-N 26Oct2017 ExcelHeader

    end;

    var
        StartDate: Date;
        EndDate: Date;
        LocCode: Code[10];
        Amt: Decimal;
        recSIH: Record 112;
        MaxCol: Text[3];
        CompInfo: Record 79;
        i: Integer;
        Sales: Boolean;
        BRT: Boolean;
        LocResC: Code[10];
        recLocation: Record 14;
        CSOMapping1: Record 50006;
        recTSH: Record 5744;
        CT3: Boolean;
        TotalAmt: Decimal;
        recSIL: Record 113;
        "----06Sep17": Integer;
        recDGST: Record "Detailed GST Ledger Entry";
        vCGST: Decimal;
        vSGST: Decimal;
        vIGST: Decimal;
        "------SIL1": Integer;
        recDGST1: Record "Detailed GST Ledger Entry";
        vCGST1: Decimal;
        vSGST1: Decimal;
        vIGST1: Decimal;
        "------SIL2": Integer;
        recDGST2: Record "Detailed GST Ledger Entry";
        vCGST2: Decimal;
        vSGST2: Decimal;
        vIGST2: Decimal;
        "------SIL3": Integer;
        recDGST3: Record "Detailed GST Ledger Entry";
        vCGST3: Decimal;
        vSGST3: Decimal;
        vIGST3: Decimal;
        "------SIL4": Integer;
        recDGST4: Record "Detailed GST Ledger Entry";
        vCGST4: Decimal;
        vSGST4: Decimal;
        vIGST4: Decimal;
        "------SIL5": Integer;
        recDGST5: Record "Detailed GST Ledger Entry";
        vCGST5: Decimal;
        vSGST5: Decimal;
        vIGST5: Decimal;
        "------SIL6": Integer;
        recDGST6: Record "Detailed GST Ledger Entry";
        vCGST6: Decimal;
        vSGST6: Decimal;
        vIGST6: Decimal;
        "------SIL7": Integer;
        recDGST7: Record "Detailed GST Ledger Entry";
        vCGST7: Decimal;
        vSGST7: Decimal;
        vIGST7: Decimal;
        "------SIL8": Integer;
        recDGST8: Record "Detailed GST Ledger Entry";
        vCGST8: Decimal;
        vSGST8: Decimal;
        vIGST8: Decimal;
        "------TIL9": Integer;
        recDGST9: Record "Detailed GST Ledger Entry";
        vCGST9: Decimal;
        vSGST9: Decimal;
        vIGST9: Decimal;
        "------TIL10": Integer;
        recDGST10: Record "Detailed GST Ledger Entry";
        vCGST10: Decimal;
        vSGST10: Decimal;
        vIGST10: Decimal;
        "-----------26Oct2017--Excel": Integer;
        ExcBuffer: Record 370 temporary;
        printtoexcel: Boolean;

    // //[Scope('Internal')]
    procedure CreateXLSHEET()
    begin
        /*
        CLEAR(XLAPP);
        CLEAR(XLWBOOK);
        CLEAR(XLSHEET);
        CREATE(XLAPP);
        XLWBOOK := XLAPP.Workbooks.Add;
        XLSHEET := XLWBOOK.Worksheets.Add;
        XLAPP.Visible(FALSE);
        MaxCol:= 'M';
        */

    end;

    // //[Scope('Internal')]
    procedure CreateHeader()
    begin
        /*
        XLSHEET.Activate;
        
        XLSHEET.Range('A1',MaxCol+'1').Merge := TRUE;
        XLSHEET.Range('A1',MaxCol+'1').Font.Bold := TRUE;
        XLSHEET.Range('A1',MaxCol+'1').Value :=CompInfo.Name;
        XLSHEET.Range('A1',MaxCol+'1').HorizontalAlignment := 3;
        XLSHEET.Range('A1',MaxCol+'1').Interior.ColorIndex := 8;
        XLSHEET.Range('A1',MaxCol+'1').Borders.ColorIndex := 5;
        XLSHEET.Range('A1',MaxCol+'1').Merge := TRUE;
        XLSHEET.Range('A1','M1').Merge := TRUE;
        
        XLSHEET.Range('A3:M3').Merge := TRUE;
        XLSHEET.Range('A3','M3').Font.Bold := TRUE;
        XLSHEET.Range('A3','M3').Value := 'Date Filter: '+FORMAT(StartDate)+' to '+FORMAT(EndDate);
        XLSHEET.Range('A3','M3').Interior.ColorIndex := 8;
        XLSHEET.Range('A3','M3').Borders.ColorIndex := 5;
        
        XLSHEET.Range('A2:M2').Merge := TRUE;
        XLSHEET.Range('A2','M2').Font.Bold := TRUE;
        XLSHEET.Range('A2','M2').Value := 'Summary of Sales/Transfer';
        XLSHEET.Range('A2','M2').Interior.ColorIndex := 8;
        XLSHEET.Range('A2','M2').Borders.ColorIndex := 5;
        
        XLSHEET.Range('A4').Value := 'UOM';
        XLSHEET.Range('A4').Font.Bold := TRUE;
        
        XLSHEET.Range('B4').Value := 'Qty';
        XLSHEET.Range('B4').Font.Bold := TRUE;
        
        XLSHEET.Range('C4').Value := 'Qty per UOM';
        XLSHEET.Range('C4').Font.Bold := TRUE;
        
        XLSHEET.Range('D4').Value := 'Qty in Ltrs/Kgs';
        XLSHEET.Range('D4').Font.Bold := TRUE;
        
        XLSHEET.Range('E4').Value := 'Net Amount';
        XLSHEET.Range('E4').Font.Bold := TRUE;
        
        XLSHEET.Range('F4').Value := 'Inventory Posting Group';
        XLSHEET.Range('F4').Font.Bold := TRUE;
        
        XLSHEET.Range('G4').Value := 'Gen. Bus. Posting Group';
        XLSHEET.Range('G4').Font.Bold := TRUE;
        
        XLSHEET.Range('H4').Value := 'Type';
        XLSHEET.Range('H4').Font.Bold := TRUE;
        
        XLSHEET.Range('I4').Value := '';
        XLSHEET.Range('I4').Font.Bold := TRUE;
        
        XLSHEET.Range('J4').Value := 'BED Amount';
        XLSHEET.Range('J4').Font.Bold := TRUE;
        
        XLSHEET.Range('K4').Value := 'Ecess Amount';
        XLSHEET.Range('K4').Font.Bold := TRUE;
        
        XLSHEET.Range('L4').Value := 'SHE cess Amount';
        XLSHEET.Range('L4').Font.Bold := TRUE;
        
        XLSHEET.Range('M4').Value := 'Excise Amount';
        XLSHEET.Range('M4').Font.Bold := TRUE;
        
        XLSHEET.Range('A3:M4').Borders.ColorIndex := 5;
        XLSHEET.Range('A2:M4').Interior.ColorIndex := 45;
        */

    end;

    // //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        IF printtoexcel THEN BEGIN
            ExcBuffer.INIT;
            ExcBuffer.VALIDATE("Row No.", Rowno);
            ExcBuffer.VALIDATE("Column No.", columnno);
            ExcBuffer."Cell Value as Text" := Cellvalue;
            ExcBuffer.Formula := '';
            ExcBuffer.Bold := Bold;
            ExcBuffer.Underline := Underline;
            ExcBuffer.NumberFormat := NoFormat;
            ExcBuffer."Cell Type" := CType;
            ExcBuffer.INSERT;
        END;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


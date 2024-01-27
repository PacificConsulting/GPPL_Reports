report 50195 "Purchase Register -RSPL"
{
    // MZ_RSPL1.00 - Preetham - 2016.12.01:
    //             - MOdification for Report Optimization
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseRegisterRSPL.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number);
            MaxIteration = 1;
            column(Comname; CompInfo.Name)
            {
            }
            column(add; CompInfo.Address)
            {
            }
            column(add2; CompInfo."Address 2")
            {
            }
            column(city; CompInfo.City)
            {
            }
            column(ReceiptHeader; ReceiptHeader)
            {
            }
            column(InvHeader; InvHeader)
            {
            }
            column(CreditmemoHeader; CreditmemoHeader)
            {
            }
            column(ShipmentHeader; ShipmentHeader)
            {
            }
            column(Summary; Summary)
            {
            }

            trigger OnAfterGetRecord()
            begin
                //BlnEdit:=TRUE;
                IF Summary = TRUE THEN BEGIN
                    sum1 := TRUE;
                    sum2 := TRUE;
                END ELSE BEGIN
                    sum1 := FALSE;
                    sum2 := FALSE;
                END;

                CompInfo.SETRANGE(CompInfo.Name);
                IF CompInfo.FIND('-') THEN;
            end;
        }
        dataitem("Purch. Rcpt. Line"; "Purch. Rcpt. Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.")
                                WHERE(Type = FILTER(<> ' '));
            column(PRH_VSN; PRH."Vendor Shipment No.")
            {
            }
            column(PRH_DocDate; PRH."Document Date")
            {
            }
            column(PRH_VRN; PRH."VAT Registration No.")
            {
            }
            column(sum1; sum1)
            {
            }
            column(PRL_CL_sum1; sum1)
            {
            }
            column(ln; "Purch. Rcpt. Line"."Line No.")
            {
            }
            column(PR_PG; 'PG.Description')
            {
            }
            column(PR_IC; IC.Description)
            {
            }
            column(PR_Postingdate; "Purch. Rcpt. Line"."Posting Date")
            {
            }
            column(HSN_SAC_PRL; "Purch. Rcpt. Line"."HSN/SAC Code")
            {
            }
            column(Vend_GST_PRL; Vendor."GST Registration No.")
            {
            }
            column(PR_No; "Purch. Rcpt. Line"."Document No.")
            {
            }
            column(TaxGroupCode_PurchRcptLine; "Purch. Rcpt. Line"."Tax Group Code")
            {
            }
            column(ItemCategoryCode_PurchRcptLine; "Purch. Rcpt. Line"."Item Category Code")
            {
            }
            column(PR_VendorNo; Vendor.Name)
            {
            }
            column(PR_ItemNo; "Purch. Rcpt. Line"."No.")
            {
            }
            column(PostingDate_PurchRcptLine; "Purch. Rcpt. Line"."Posting Date")
            {
            }
            column(OrderNo_PurchRcptLine; "Purch. Rcpt. Line"."Order No.")
            {
            }
            column(OrderDate_PurchRcptLine; "Purch. Rcpt. Line"."Order Date")
            {
            }
            column(BlanketOrderNo_PurchRcptLine; "Purch. Rcpt. Line"."Blanket Order No.")
            {
            }
            column(PostingGroup_PurchRcptLine; "Purch. Rcpt. Line"."Posting Group")
            {
            }
            column(PR_ItemDesc; "Purch. Rcpt. Line".Description)
            {
            }
            column(Description2_PurchRcptLine; "Purch. Rcpt. Line"."Description 2")
            {
            }
            column(PR_CurrCode; "Purch. Rcpt. Line"."Currency Code")
            {
            }
            column(PR_Qty; "Purch. Rcpt. Line".Quantity)
            {
            }
            column(FormCode_PurchRcptLine; '"Purch. Rcpt. Line"."Form Code"')
            {
            }
            column(PR_UOM; "Purch. Rcpt. Line"."Unit of Measure")
            {
            }
            column(PR_UnitCost; PRHDUC)
            {
            }
            column(PR_BEDAmt; PRHBEDAMT)
            {
            }
            column(PR_EcessAmt; PRHECESSAMT)
            {
            }
            column(PR_SCessAmt; PRHSHEECESS)
            {
            }
            column(PR_ADCVAT; PRHADCVAT)
            {
            }
            column(PR_TaxPer; 0)//"Purch. Rcpt. Line"."Tax %")
            {
            }
            column(PR_FormCode; 'PRH."Form Code"')
            {
            }
            column(PR_FormNo; 'PRH."Form No."')
            {
            }
            column(PR_LRNo; '')
            {
            }
            column(PR_LRDate; '')
            {
            }
            column(PR_Dim1; PRH."Shortcut Dimension 1 Code")
            {
            }
            column(PR_Dim2; PRH."Shortcut Dimension 2 Code")
            {
            }
            column(PR_Exciseamt; PRHEXCIEAMT)
            {
            }
            column(PR_DiscPer; "Purch. Rcpt. Line"."Line Discount %")
            {
            }
            column(PR_ExcisePer; PRKEXCISEPER)
            {
            }
            column(PRAMT; PRAMT)
            {
            }
            column(ShpAgRcpt; recShipmentAg.Name)
            {
            }
            column(ADEAmount_PurchRcptLine; 0)// "Purch. Rcpt. Line"."ADE Amount")
            {
            }
            column(ServiceTaxSBCAmount_PurchRcptLine; 0)// "Purch. Rcpt. Line"."Service Tax SBC Amount")
            {
            }
            column(KKCessAmount_PurchRcptLine; 0)// "Purch. Rcpt. Line"."KK Cess Amount")
            {
            }
            column(CostCentrAmt; CostCentrAmt)
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
            column(vCGSTRate; vCGSTRate)
            {
            }
            column(vSGSTRate; vSGSTRate)
            {
            }
            column(vIGSTRate; vIGSTRate)
            {
            }
            column(GST_Group_Code_PRL; "Purch. Rcpt. Line"."GST Group Code")
            {
            }
            column(vUTGST; vUTGST)
            {
            }
            column(vSESS; vSESS)
            {
            }
            column(DetailGSTEntry_InvoiceType; 'DetailGSTEntry."Invoice Type"')
            {
            }

            trigger OnAfterGetRecord()
            begin
                // >> MZ_RSPL1.00
                IF NOT ReceiptHeader THEN
                    CurrReport.SKIP;
                // << MZ_RSPL1.00
                vRcptDoc := '';
                VrecNo := '';
                IF (Quantity * "Direct Unit Cost") <> 0 THEN
                    PR_ExcisePer := ROUND(/*"Excise Amount" */0 / (Quantity * "Direct Unit Cost") * 100);

                //GUNJAN-START
                CLEAR(CF);
                REC_PRH.RESET;
                REC_PRH.SETRANGE(REC_PRH."No.", "Purch. Rcpt. Line"."Document No.");
                IF REC_PRH.FINDFIRST THEN
                    CF := REC_PRH."Currency Factor";


                //RSPL-rahul
                CLEAR(CostCentrAmt);
                DimSetEntry.RESET;
                DimSetEntry.SETRANGE(DimSetEntry."Dimension Set ID", REC_PRH."Dimension Set ID");
                DimSetEntry.SETRANGE(DimSetEntry."Dimension Code", 'CC');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    CostCentrAmt := DimSetEntry."Dimension Value Name";
                END;

                IF CF <> 0 THEN BEGIN
                    PRHDUC := "Purch. Rcpt. Line"."Unit Cost" / CF;
                    PRHBEDAMT := 0;// "Purch. Rcpt. Line"."BED Amount" / CF;
                    PRHADCVAT := 0;// "Purch. Rcpt. Line"."ADC VAT Amount" / CF;
                    PRHECESSAMT := 0;// "Purch. Rcpt. Line"."eCess Amount" / CF;
                    PRHSHEECESS := 0;// "Purch. Rcpt. Line"."SHE Cess Amount" / CF;
                    PRHEXCIEAMT := 0;// "Purch. Rcpt. Line"."Excise Amount" / CF;
                    PRKEXCISEPER := PR_ExcisePer / CF;
                    PRAMT := ("Purch. Rcpt. Line".Quantity * "Purch. Rcpt. Line"."Unit Cost") / CF;
                END
                ELSE BEGIN
                    PRHDUC := "Purch. Rcpt. Line"."Direct Unit Cost";
                    PRHBEDAMT := 0;// "Purch. Rcpt. Line"."BED Amount";
                    PRHADCVAT := 0;//"Purch. Rcpt. Line"."ADC VAT Amount";
                    PRHECESSAMT := 0;// "Purch. Rcpt. Line"."eCess Amount";
                    PRHSHEECESS := 0;// "Purch. Rcpt. Line"."SHE Cess Amount";
                    PRHEXCIEAMT := 0;// "Purch. Rcpt. Line"."Excise Amount";
                    PRKEXCISEPER := PR_ExcisePer;
                    PRAMT := ("Purch. Rcpt. Line".Quantity * "Purch. Rcpt. Line"."Unit Cost");
                END;
                //GUNJAN-END



                // IF PG.GET("Product Group Code") THEN;
                IF IC.GET("Item Category Code") THEN;

                PRH.RESET;
                PRH.SETRANGE(PRH."No.", "Document No.");
                IF PRH.FIND('-') THEN
                    //VRN := PRH."VAT Registration No.";
                    Vendor.GET(PRH."Buy-from Vendor No.");

                recShipmentAg.RESET;
                recPurchRH.GET("Purch. Rcpt. Line"."Document No.");

                //>>GST
                vRcptDoc := "Purch. Rcpt. Line"."Document No.";
                VrecNo := "Purch. Rcpt. Line"."No.";
                CLEAR(vCGST);
                CLEAR(vSGST);
                CLEAR(vIGST);
                CLEAR(vUTGST);
                CLEAR(vSESS);
                CLEAR(vCGSTRate);
                CLEAR(vSGSTRate);
                CLEAR(vIGSTRate);
                DetailGSTEntry.RESET;
                //DetailGSTEntry.SETCURRENTKEY("Document No.","Document Line No.","GST Component Code");
                DetailGSTEntry.SETRANGE("Document No.", vRcptDoc);
                DetailGSTEntry.SETRANGE("No.", VrecNo);
                DetailGSTEntry.SETRANGE("Document Line No.", "Line No.");
                IF DetailGSTEntry.FINDFIRST THEN
                    REPEAT
                        IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST := ABS(DetailGSTEntry."GST Amount");
                            //MESSAGE(FORMAT(vCGST));       //RK
                            vCGSTRate := DetailGSTEntry."GST %";
                        END;
                        IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST := ABS(DetailGSTEntry."GST Amount");
                            vSGSTRate := DetailGSTEntry."GST %";
                        END;
                        IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST := ABS(DetailGSTEntry."GST Amount");
                            vIGSTRate := DetailGSTEntry."GST %";
                        END;
                        IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                            vUTGST := ABS(DetailGSTEntry."GST Amount");
                        END;
                        IF DetailGSTEntry."GST Component Code" = 'CESS' THEN BEGIN
                            vSESS := ABS(DetailGSTEntry."GST Amount");
                        END;


                    UNTIL DetailGSTEntry.NEXT = 0;

                //<<
            end;
        }
        dataitem("Purch. Inv. Line"; "Purch. Inv. Line")
        {
            column(PIH_DocDate; PIH."Document Date")
            {
            }
            column(PIH_VIN; PIH."Vendor Invoice No.")
            {
            }
            column(PIH_VRN; PIH."VAT Registration No.")
            {
            }
            column(PI_ExcisePer; ROUND(PIEXEPER))
            {
            }
            column(sum2; sum1)
            {
            }
            column(PIL_CL_SUM2; sum1)
            {
            }
            column(Vend_GST_Reg_PIL; Vendor."GST Registration No.")
            {
            }
            column(curr_code; PIH."Currency Code")
            {
            }
            column(PI_PG; 'PG.Description')
            {
            }
            column(PI_IC; IC.Description)
            {
            }
            column(PI_Postingdate; "Purch. Inv. Line"."Posting Date")
            {
            }
            column(PI_No; "Purch. Inv. Line"."Document No.")
            {
            }
            column(HSN_SAC_PIL; "Purch. Inv. Line"."HSN/SAC Code")
            {
            }
            column(BuyfromVendorNo_PurchInvLine; "Purch. Inv. Line"."Buy-from Vendor No.")
            {
            }
            column(PI_VendorNo; Vendor.Name)
            {
            }
            column(PI_VATNo; Vendor."VAT Registration No.")
            {
            }
            column(BlanketOrderNo_PurchInvLine; "Purch. Inv. Line"."Blanket Order No.")
            {
            }
            column(OrderNo_PurchInvLine; '')
            {
            }
            column(PI_ItemNo; "Purch. Inv. Line"."No.")
            {
            }
            column(PI_ItemDesc; "Purch. Inv. Line".Description)
            {
            }
            column(Description2_PurchInvLine; "Purch. Inv. Line"."Description 2")
            {
            }
            column(ItemCategoryCode_PurchInvLine; "Purch. Inv. Line"."Item Category Code")
            {
            }
            column(PI_Qty; "Purch. Inv. Line".Quantity)
            {
            }
            column(TaxGroupCode_PurchInvLine; "Purch. Inv. Line"."Tax Group Code")
            {
            }
            column(PI_UOM; "Purch. Inv. Line"."Unit of Measure")
            {
            }
            column(PI_UnitCost; PIHDUC)
            {
            }
            column(PI_BEDAmt; PIHBEDAMT)
            {
            }
            column(PI_EcessAmt; PIHECESSAMT)
            {
            }
            column(PI_SCessAmt; PIHSHEECESS)
            {
            }
            column(PI_TaxPer; 0)//"Purch. Inv. Line"."Tax %")
            {
            }
            column(PI_FormCode; 'PIH."Form Code"')
            {
            }
            column(PI_FormNo; 'PIH."Form No."')
            {
            }
            column(PI_Dim1; PIH."Shortcut Dimension 1 Code")
            {
            }
            column(PI_Dim2; PIH."Shortcut Dimension 2 Code")
            {
            }
            column(PI_ExciseAmt; PIHEXCISEAMT)
            {
            }
            column(PI_TaxAmt; PIHTAXAMT)
            {
            }
            column(PI_Amttovendor; PIHATV)
            {
            }
            column(PI_Chargestovendor; PIHCTV)
            {
            }
            column(PI_TDSAmt; PIHTDS)
            {
            }
            column(PI_ServiceTaxAmt; PIHSTA)
            {
            }
            column(PI_ADCVAT; PIHADCVAT)
            {
            }
            column(PI_DiscAmt; "Purch. Inv. Line"."Line Discount Amount")
            {
            }
            column(Amount_PurchInvLine; "Purch. Inv. Line".Amount)
            {
            }
            column(FormCode_PurchInvLine; '"Purch. Inv. Line"."Form Code"')
            {
            }
            column(FormNo_PurchInvLine; '"Purch. Inv. Line"."Form No."')
            {
            }
            column(PIHAmt; PIHAmt)
            {
            }
            column(ShpAgInv; recShipmentAg.Name)
            {
            }
            column(ADEAmount_PurchInvLine; 0)// "Purch. Inv. Line"."ADE Amount")
            {
            }
            column(ServiceTaxSBCAmount_PurchInvLine; 0)// "Purch. Inv. Line"."Service Tax SBC Amount")
            {
            }
            column(KKCessAmount_PurchInvLine; 0)// "Purch. Inv. Line"."KK Cess Amount")
            {
            }
            column(CostCentrAmt2; CostCentrAmt2)
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
            column(vCGSTRate1; vCGSTRate1)
            {
            }
            column(vSGSTRate1; vSGSTRate1)
            {
            }
            column(vIGSTRate1; vIGSTRate1)
            {
            }
            column(GST_Group_Code_PIL; "Purch. Inv. Line"."GST Group Code")
            {
            }
            column(vUTGST1; vUTGST1)
            {
            }
            column(vSESS1; vSESS1)
            {
            }
            column(DetailGSTEntry1_InvoiceType; 'DetailGSTEntry1."Invoice Type"')
            {
            }

            trigger OnAfterGetRecord()
            begin
                // >> MZ_RSPL1.00
                IF NOT InvHeader THEN
                    CurrReport.SKIP;
                // << MZ_RSPL1.00

                //GUNJAN-START
                CLEAR(CF);
                REC_PIH.RESET;
                REC_PIH.SETRANGE(REC_PIH."No.", "Purch. Inv. Line"."Document No.");
                IF REC_PIH.FINDFIRST THEN
                    CF := REC_PIH."Currency Factor";

                //RSPL-rahul
                CLEAR(CostCentrAmt2);
                DimSetEntry.RESET;
                DimSetEntry.SETRANGE(DimSetEntry."Dimension Set ID", REC_PIH."Dimension Set ID");
                DimSetEntry.SETRANGE(DimSetEntry."Dimension Code", 'CC');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    CostCentrAmt2 := DimSetEntry."Dimension Value Name";
                END;

                IF CF <> 0 THEN BEGIN
                    PIHDUC := "Purch. Inv. Line"."Unit Cost" / CF;
                    PIHAmt := ("Purch. Inv. Line".Quantity * "Purch. Inv. Line"."Unit Cost") / CF;
                    PIHBEDAMT := 0;// "Purch. Inv. Line"."BED Amount" / CF;
                    PIHADCVAT := 0;// "Purch. Inv. Line"."ADC VAT Amount" / CF;
                    PIHECESSAMT := 0;// "Purch. Inv. Line"."eCess Amount" / CF;
                    PIHSHEECESS := 0;// "Purch. Inv. Line"."SHE Cess Amount" / CF;
                    PIHEXCISEAMT := 0;// "Purch. Inv. Line"."Excise Amount" / CF;
                    PIHTAXAMT := 0;//"Purch. Inv. Line"."Tax Amount" / CF;
                    PIHATV := 0;// "Purch. Inv. Line"."Amount To Vendor" / CF;
                    PIHCTV := 0;// "Purch. Inv. Line"."Charges To Vendor" / CF;
                    PIHSTA := 0;// "Purch. Inv. Line"."Service Tax Amount" / CF;
                    PIHTDS := 0;//"Purch. Inv. Line"."TDS Amount" / CF;
                    PIEXEPER := PI_ExcisePer / CF;
                END
                ELSE BEGIN
                    PIHDUC := "Purch. Inv. Line"."Direct Unit Cost";
                    PIHBEDAMT := 0;// "Purch. Inv. Line"."BED Amount";
                    PIHADCVAT := 0;// "Purch. Inv. Line"."ADC VAT Amount";
                    PIHECESSAMT := 0;// "Purch. Inv. Line"."eCess Amount";
                    PIHSHEECESS := 0;// "Purch. Inv. Line"."SHE Cess Amount";
                    PIHEXCISEAMT := 0;//"Purch. Inv. Line"."Excise Amount";
                    PIHTAXAMT := 0;// "Purch. Inv. Line"."Tax Amount";
                    PIHATV := 0;// "Purch. Inv. Line"."Amount To Vendor";
                    PIHCTV := 0;// "Purch. Inv. Line"."Charges To Vendor";
                    PIHSTA := 0;//"Purch. Inv. Line"."Service Tax Amount";
                    PIHTDS := 0;// "Purch. Inv. Line"."TDS Amount";
                    PIEXEPER := PI_ExcisePer;
                    PIHAmt := "Purch. Inv. Line".Quantity * "Purch. Inv. Line"."Direct Unit Cost";
                END;

                //GUNJAN-END



                IF (Quantity * "Direct Unit Cost") <> 0 THEN
                    PI_ExcisePer := ROUND(/*"Excise Amount"*/0 / (Quantity * "Direct Unit Cost") * 100);

                //IF PG.GET("Product Group Code") THEN;

                IF IC.GET("Item Category Code") THEN;

                PIH.RESET;
                PIH.SETRANGE(PIH."No.", "Document No.");
                IF PIH.FIND('-') THEN
                    Vendor.GET(PIH."Buy-from Vendor No.");

                recShipmentAg.RESET;
                recPurchIH.GET("Purch. Inv. Line"."Document No.");

                //>>GST
                CLEAR(vCGST1);
                CLEAR(vSGST1);
                CLEAR(vIGST1);
                CLEAR(vSESS1);
                CLEAR(vUTGST1);
                CLEAR(vCGSTRate1);
                CLEAR(vSGSTRate1);
                CLEAR(vIGSTRate1);
                DetailGSTEntry1.RESET;
                DetailGSTEntry1.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DetailGSTEntry1.SETRANGE("Document No.", "Document No.");
                DetailGSTEntry1.SETRANGE("No.", "No.");
                DetailGSTEntry1.SETRANGE("Document Line No.", "Line No.");
                IF DetailGSTEntry1.FINDSET THEN
                    REPEAT
                        IF DetailGSTEntry1."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST1 := ABS(DetailGSTEntry1."GST Amount");
                            vCGSTRate1 := DetailGSTEntry1."GST %";
                        END;
                        IF DetailGSTEntry1."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST1 := ABS(DetailGSTEntry1."GST Amount");
                            vSGSTRate1 := DetailGSTEntry1."GST %";
                        END;
                        IF DetailGSTEntry1."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST1 := ABS(DetailGSTEntry1."GST Amount");
                            vIGSTRate1 := DetailGSTEntry1."GST %";
                        END;
                        IF DetailGSTEntry1."GST Component Code" = 'UTGST' THEN BEGIN
                            vUTGST1 := ABS(DetailGSTEntry1."GST Amount");
                        END;
                        IF DetailGSTEntry1."GST Component Code" = 'CESS' THEN BEGIN
                            vSESS1 := ABS(DetailGSTEntry1."GST Amount");
                        END;

                    UNTIL DetailGSTEntry1.NEXT = 0;

                //<<
            end;
        }
        dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
        {
            column(PCH_VCMN; PCH."Vendor Cr. Memo No.")
            {
            }
            column(PCH_DocDate; PCH."Document Date")
            {
            }
            column(PCH_VRN; PCH."VAT Registration No.")
            {
            }
            column(PC_ExcisePer; ROUND(PCEXEPER))
            {
            }
            column(sum3; sum1)
            {
            }
            column(PCML_CL_SUM3; sum1)
            {
            }
            column(PC_PG; 'PG.Description')
            {
            }
            column(PC_IC; IC.Description)
            {
            }
            column(PC_Postingdate; "Purch. Cr. Memo Line"."Posting Date")
            {
            }
            column(HSN_SAC_PCM; "Purch. Cr. Memo Line"."HSN/SAC Code")
            {
            }
            column(PC_No; "Purch. Cr. Memo Line"."Document No.")
            {
            }
            column(PC_VendorNo; Vendor.Name)
            {
            }
            column(Vend_Reg_No_PCM; Vendor."GST Registration No.")
            {
            }
            column(PC_ItemNo; "Purch. Cr. Memo Line"."No.")
            {
            }
            column(PC_ItemDesc; "Purch. Cr. Memo Line".Description)
            {
            }
            column(PC_Qty; "Purch. Cr. Memo Line".Quantity)
            {
            }
            column(PC_UOM; "Purch. Cr. Memo Line"."Unit of Measure")
            {
            }
            column(PC_ADC; PCADCVAT)
            {
            }
            column(PC_UnitCost; PCDUC)
            {
            }
            column(PC_BEDAmt; PCBEDAMT)
            {
            }
            column(PC_EcessAmt; PCECESSAMT)
            {
            }
            column(PC_SCessAmt; PCSHEECESS)
            {
            }
            column(PC_TaxPer; '"Purch. Cr. Memo Line"."Tax %"')
            {
            }
            column(PC_FormCode; 'PCH."Form Code"')
            {
            }
            column(PC_FormNo; 'PCH."Form No."')
            {
            }
            column(PC_Dim1; PCH."Shortcut Dimension 1 Code")
            {
            }
            column(PC_Dim2; PCH."Shortcut Dimension 2 Code")
            {
            }
            column(PC_ExciseAmt; PCEXCISEAMT)
            {
            }
            column(PC_TaxAmt; PCTAXAMT)
            {
            }
            column(PC_AmttoVendor; PCATV)
            {
            }
            column(PC_chargetovendor; PCCTV)
            {
            }
            column(PC_ServiceTaxAmt; PCSTA)
            {
            }
            column(PC_DiscAmt; "Purch. Cr. Memo Line"."Line Discount Amount")
            {
            }
            column(PCAmt; PCAmt)
            {
            }
            column(ShpAgCr; recShipmentAg.Name)
            {
            }
            column(ADEAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."ADE Amount")
            {
            }
            column(KKCessAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."KK Cess Amount")
            {
            }
            column(ServiceTaxSBCAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."Service Tax SBC Amount")
            {
            }
            column(AssessableValue_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."Assessable Value")
            {
            }
            column(VATType_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."VAT Type")
            {
            }
            column(SHECessAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."SHE Cess Amount")
            {
            }
            column(ServiceTaxSHECessAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."Service Tax SHE Cess Amount")
            {
            }
            column(Supplementary_PurchCrMemoLine; "Purch. Cr. Memo Line".Supplementary)
            {
            }
            column(SourceDocumentType_PurchCrMemoLine; "Purch. Cr. Memo Line"."Source Document Type")
            {
            }
            column(SourceDocumentNo_PurchCrMemoLine; "Purch. Cr. Memo Line"."Source Document No.")
            {
            }
            column(ADCVATAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."ADC VAT Amount")
            {
            }
            column(CIFAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."CIF Amount")
            {
            }
            column(BCDAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."BCD Amount")
            {
            }
            column(CVD_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line".CVD)
            {
            }
            column(NotificationSlNo_PurchCrMemoLine; '"Purch. Cr. Memo Line"."Notification Sl. No."')
            {
            }
            column(NotificationNo_PurchCrMemoLine; '"Purch. Cr. Memo Line"."Notification No."')
            {
            }
            column(CTSHNo_PurchCrMemoLine; '"Purch. Cr. Memo Line"."CTSH No."')
            {
            }
            column(RetShipmentNo_PurchCrMemoLine; '"Purch. Cr. Memo Line"."Ret. Shipment No."')
            {
            }
            column(CostCentrAmt3; CostCentrAmt3)
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
            column(vCGSTRate2; vCGSTRate2)
            {
            }
            column(vSGSTRate2; vSGSTRate2)
            {
            }
            column(vIGSTRate2; vIGSTRate2)
            {
            }
            column(GST_Group_Code_PCrMLine; "Purch. Cr. Memo Line"."GST Group Code")
            {
            }
            column(vUTGST2; vUTGST2)
            {
            }
            column(vSESS2; vSESS2)
            {
            }
            column(DetailGSTEntry2_InvoiceType; 'DetailGSTEntry2."Invoice Type"')
            {
            }

            trigger OnAfterGetRecord()
            begin
                // >> MZ_RSPL1.00
                IF NOT CreditmemoHeader THEN
                    CurrReport.SKIP;
                // << MZ_RSPL1.00

                //GUNJAN-START
                CLEAR(CF);
                REC_PC.RESET;
                REC_PC.SETRANGE(REC_PC."No.", "Purch. Cr. Memo Line"."Document No.");
                IF REC_PC.FINDFIRST THEN
                    CF := REC_PC."Currency Factor";

                //RSPL-rahul
                CLEAR(CostCentrAmt3);
                DimSetEntry.RESET;
                DimSetEntry.SETRANGE(DimSetEntry."Dimension Set ID", REC_PC."Dimension Set ID");
                DimSetEntry.SETRANGE(DimSetEntry."Dimension Code", 'CC');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    CostCentrAmt3 := DimSetEntry."Dimension Value Name";
                END;

                IF CF <> 0 THEN BEGIN
                    PCDUC := "Purch. Cr. Memo Line"."Unit Cost" / CF;
                    PCAmt := ("Purch. Cr. Memo Line".Quantity * "Purch. Cr. Memo Line"."Unit Cost") / CF;
                    PCBEDAMT := 0;//"Purch. Cr. Memo Line"."BED Amount" / CF;
                    PCADCVAT := 0;// "Purch. Cr. Memo Line"."ADC VAT Amount" / CF;
                    PCECESSAMT := 0;// "Purch. Cr. Memo Line"."eCess Amount" / CF;
                    PCSHEECESS := 0;// "Purch. Cr. Memo Line"."SHE Cess Amount" / CF;
                    PCEXCISEAMT := 0;// "Purch. Cr. Memo Line"."Excise Amount" / CF;
                    PCTAXAMT := 0;// "Purch. Cr. Memo Line"."Tax Amount" / CF;
                    PCATV := 0;// "Purch. Cr. Memo Line"."Amount To Vendor" / CF;
                    PCCTV := 0;// "Purch. Cr. Memo Line"."Charges To Vendor" / CF;
                    PCSTA := 0;// "Purch. Cr. Memo Line"."Service Tax Amount" / CF;
                    PCEXEPER := PC_ExcisePer / CF;
                END ELSE BEGIN
                    PCDUC := "Purch. Cr. Memo Line"."Direct Unit Cost";
                    PCAmt := ("Purch. Cr. Memo Line".Quantity * "Purch. Cr. Memo Line"."Unit Cost");
                    PCBEDAMT := 0;// "Purch. Cr. Memo Line"."BED Amount";
                    PCADCVAT := 0;// "Purch. Cr. Memo Line"."ADC VAT Amount";
                    PCECESSAMT := 0;// "Purch. Cr. Memo Line"."eCess Amount";
                    PCSHEECESS := 0;// "Purch. Cr. Memo Line"."SHE Cess Amount";
                    PCEXCISEAMT := 0;// "Purch. Cr. Memo Line"."Excise Amount";
                    PCTAXAMT := 0;// "Purch. Cr. Memo Line"."Tax Amount";
                    PCATV := 0;// "Purch. Cr. Memo Line"."Amount To Vendor";
                    PCCTV := 0;// "Purch. Cr. Memo Line"."Charges To Vendor";
                    PCSTA := 0;// "Purch. Cr. Memo Line"."Service Tax Amount";
                    PCEXEPER := PC_ExcisePer;
                END;
                //GUNJAN-END

                IF (Quantity * "Direct Unit Cost") <> 0 THEN;
                PC_ExcisePer := ROUND(/*"Excise Amount"*/0 / (Quantity * "Direct Unit Cost") * 100);

                //IF PG.GET("Product Group Code") THEN;
                IF IC.GET("Item Category Code") THEN;

                PCH.RESET;
                PCH.SETRANGE(PCH."No.", "Document No.");
                IF PCH.FIND('-') THEN
                    Vendor.GET(PCH."Buy-from Vendor No.");

                recShipmentAg.RESET;
                recPurchCH.GET("Purch. Cr. Memo Line"."Document No.");


                //>>GST
                CLEAR(vCGST2);
                CLEAR(vSGST2);
                CLEAR(vIGST2);
                CLEAR(vSESS2);
                CLEAR(vUTGST2);
                CLEAR(vCGSTRate2);
                CLEAR(vSGSTRate2);
                CLEAR(vIGSTRate2);
                DetailGSTEntry2.RESET;
                DetailGSTEntry2.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DetailGSTEntry2.SETRANGE("Document No.", "Document No.");
                DetailGSTEntry2.SETRANGE("No.", "No.");
                DetailGSTEntry2.SETRANGE("Document Line No.", "Line No.");
                IF DetailGSTEntry2.FINDSET THEN
                    REPEAT
                        IF DetailGSTEntry2."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST2 := ABS(DetailGSTEntry2."GST Amount");
                            vCGSTRate2 := DetailGSTEntry2."GST %";
                        END;
                        IF DetailGSTEntry2."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST2 := ABS(DetailGSTEntry2."GST Amount");
                            vSGSTRate2 := DetailGSTEntry2."GST %";
                        END;
                        IF DetailGSTEntry2."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST2 := ABS(DetailGSTEntry2."GST Amount");
                            vIGSTRate2 := DetailGSTEntry2."GST %";
                        END;
                        IF DetailGSTEntry2."GST Component Code" = 'UTGST' THEN BEGIN
                            vUTGST2 := ABS(DetailGSTEntry2."GST Amount");
                        END;
                        IF DetailGSTEntry2."GST Component Code" = 'CESS' THEN BEGIN
                            vSESS2 := ABS(DetailGSTEntry2."GST Amount");
                        END;

                    UNTIL DetailGSTEntry2.NEXT = 0;

                //<<
            end;
        }
        dataitem("Return Shipment Line"; "Return Shipment Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.");
            column(RSH_VRN; RSH."VAT Registration No.")
            {
            }
            column(RR_ExcisePer; ROUND(RSLEXEPER))
            {
            }
            column(sum4; sum1)
            {
            }
            column(RSL_CL_SUM4; sum1)
            {
            }
            column(RS_PG; 'PG.Description')
            {
            }
            column(RS_IC; IC.Description)
            {
            }
            column(Vend_Reg_No_PSL; Vendor."GST Registration No.")
            {
            }
            column(RR_Postingdate; "Return Shipment Line"."Posting Date")
            {
            }
            column(RR_No; "Return Shipment Line"."Document No.")
            {
            }
            column(RR_VendorNo; Vendor.Name)
            {
            }
            column(RR_ItemNo; "Return Shipment Line"."No.")
            {
            }
            column(HSN_SAC_PSL; '"Return Shipment Line"."HSN/SAC Code"')
            {
            }
            column(RR_ItemDesc; "Return Shipment Line".Description)
            {
            }
            column(RR_Qty; "Return Shipment Line".Quantity)
            {
            }
            column(RR_UOM; "Return Shipment Line"."Unit of Measure")
            {
            }
            column(RR_UnitCost; RSLDUC)
            {
            }
            column(RR_ADC; RSLADCVAT)
            {
            }
            column(RR_BEDAmt; RSLBEDAMT)
            {
            }
            column(RR_EcessAmt; RSLECESSAMT)
            {
            }
            column(RR_SCessAmt; RSLSHEECESS)
            {
            }
            column(RR_TaxPer; 0)//"Return Shipment Line"."Tax %")
            {
            }
            column(RR_FormCode; 'RSH."Form Code"')
            {
            }
            column(RR_FormNo; 'RSH."Form No."')
            {
            }
            column(RR_Dim1; RSH."Shortcut Dimension 1 Code")
            {
            }
            column(RR_Dim2; RSH."Shortcut Dimension 2 Code")
            {
            }
            column(RS_ExciseAmt; RSLEXCISEAMT)
            {
            }
            column(RS_TaxAmt; RSLTAXAMT)
            {
            }
            column(RS_DiscPer; "Return Shipment Line"."Line Discount %")
            {
            }
            column(RSLAmt; RSLAmt)
            {
            }
            column(ShpAgReturn; recShipmentAg.Name)
            {
            }
            column(AEDGSIAmount_ReturnShipmentLine; 0)// "Return Shipment Line"."AED(GSI) Amount")
            {
            }
            column(KKCessAmount_ReturnShipmentLine; 0)// "Return Shipment Line"."KK Cess Amount")
            {
            }
            column(ServiceTaxSBCAmount_ReturnShipmentLine; 0)//"Return Shipment Line"."Service Tax SBC Amount")
            {
            }
            column(CostCentrAmt4; CostCentrAmt4)
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
            column(vCGSTRate3; vCGSTRate3)
            {
            }
            column(vSGSTRate3; vSGSTRate3)
            {
            }
            column(vIGSTRate3; vIGSTRate3)
            {
            }
            column(GST_Group_Code_RSL; '"Return Shipment Line"."GST Group Code"')
            {
            }
            column(vUTGST3; vUTGST3)
            {
            }
            column(vSESS3; vSESS3)
            {
            }
            column(DetailGSTEntry3_InvoiceType; 'DetailGSTEntry3."Invoice Type"')
            {
            }

            trigger OnAfterGetRecord()
            begin
                // >> MZ_RSPL1.00
                IF NOT ShipmentHeader THEN
                    CurrReport.SKIP;
                // << MZ_RSPL1.00

                //GUNJAN-START
                CLEAR(CF);
                REC_RSH.RESET;
                REC_RSH.SETRANGE(REC_RSH."No.", "Return Shipment Line"."Document No.");
                IF REC_RSH.FINDFIRST THEN
                    CF := REC_RSH."Currency Factor";

                //RSPL-rahul
                CLEAR(CostCentrAmt4);
                DimSetEntry.RESET;
                DimSetEntry.SETRANGE(DimSetEntry."Dimension Set ID", REC_RSH."Dimension Set ID");
                DimSetEntry.SETRANGE(DimSetEntry."Dimension Code", 'CC');
                IF DimSetEntry.FINDFIRST THEN BEGIN
                    DimSetEntry.CALCFIELDS("Dimension Value Name");
                    CostCentrAmt4 := DimSetEntry."Dimension Value Name";
                END;

                IF CF <> 0 THEN BEGIN
                    RSLDUC := "Return Shipment Line"."Unit Cost" / CF;
                    RSLAmt := ("Return Shipment Line".Quantity * "Return Shipment Line"."Unit Cost") / CF;
                    RSLBEDAMT := 0;//"Return Shipment Line"."BED Amount" / CF;
                    RSLADCVAT := 0;// "Return Shipment Line"."ADC VAT Amount" / CF;
                    RSLECESSAMT := 0;// "Return Shipment Line"."eCess Amount" / CF;
                    RSLSHEECESS := 0;// "Return Shipment Line"."SHE Cess Amount" / CF;
                    RSLEXCISEAMT := 0;//"Return Shipment Line"."Excise Amount" / CF;
                    RSLTAXAMT := 0;//"Return Shipment Line"."Tax Amount" / CF;
                    RSLEXEPER := RR_ExcisePer / CF;
                END ELSE BEGIN
                    RSLDUC := "Return Shipment Line"."Direct Unit Cost";
                    RSLAmt := ("Return Shipment Line".Quantity * "Return Shipment Line"."Unit Cost");
                    RSLBEDAMT := 0;// "Return Shipment Line"."BED Amount";
                    RSLADCVAT := 0;// "Return Shipment Line"."ADC VAT Amount";
                    RSLECESSAMT := 0;// "Return Shipment Line"."eCess Amount";
                    RSLSHEECESS := 0;//"Return Shipment Line"."SHE Cess Amount";
                    RSLEXCISEAMT := 0;//"Return Shipment Line"."Excise Amount";
                    RSLTAXAMT := 0;//"Return Shipment Line"."Tax Amount";
                    RSLEXEPER := RR_ExcisePer;
                END;
                //GUNJAN-END


                IF (Quantity * "Direct Unit Cost") <> 0 THEN;
                RR_ExcisePer := ROUND(/*"Excise Amount" */0 / (Quantity * "Direct Unit Cost") * 100);

                //IF PG.GET("Product Group Code") THEN;
                IF IC.GET("Item Category Code") THEN;

                RSH.RESET;
                RSH.SETRANGE(RSH."No.", "Document No.");
                IF RSH.FIND('-') THEN
                    Vendor.GET(RSH."Buy-from Vendor No.");
                recShipmentAg.RESET;
                recReturnH.GET("Return Shipment Line"."Document No.");

                //>>GST
                CLEAR(vCGST3);
                CLEAR(vSGST3);
                CLEAR(vIGST3);
                CLEAR(vUTGST3);
                CLEAR(vSESS3);
                CLEAR(vCGSTRate3);
                CLEAR(vSGSTRate3);
                CLEAR(vIGSTRate3);
                DetailGSTEntry3.RESET;
                DetailGSTEntry3.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DetailGSTEntry3.SETRANGE("Document No.", "Document No.");
                DetailGSTEntry3.SETRANGE("No.", "No.");
                DetailGSTEntry3.SETRANGE("Document Line No.", "Line No.");
                IF DetailGSTEntry3.FINDSET THEN
                    REPEAT
                        IF DetailGSTEntry3."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST3 := ABS(DetailGSTEntry3."GST Amount");
                            vCGSTRate3 := DetailGSTEntry3."GST %";
                        END;
                        IF DetailGSTEntry3."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST3 := ABS(DetailGSTEntry3."GST Amount");
                            vSGSTRate3 := DetailGSTEntry3."GST %";
                        END;
                        IF DetailGSTEntry3."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST3 := ABS(DetailGSTEntry3."GST Amount");
                            vIGSTRate3 := DetailGSTEntry3."GST %";
                        END;
                        IF DetailGSTEntry3."GST Component Code" = 'UTGST' THEN BEGIN
                            vUTGST3 := ABS(DetailGSTEntry3."GST Amount");
                        END;
                        IF DetailGSTEntry3."GST Component Code" = 'CESS' THEN BEGIN
                            vSESS3 := ABS(DetailGSTEntry3."GST Amount");
                        END;

                    UNTIL DetailGSTEntry3.NEXT = 0;

                //<<
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                group(Option)
                {
                    field("Purchase Receipt "; ReceiptHeader)
                    {
                        Visible = false;

                        trigger OnValidate()
                        begin




                            IF (InvHeader = TRUE) OR (CreditmemoHeader = TRUE) OR (ShipmentHeader = TRUE) OR (Summary = TRUE) THEN BEGIN
                                //ReceiptHeader := FALSE;
                                ERROR('Select only One at a time');
                            END;
                        end;
                    }
                    field("Purchase Invoice"; InvHeader)
                    {

                        trigger OnValidate()
                        begin

                            IF (ReceiptHeader = TRUE) OR (CreditmemoHeader = TRUE) OR (ShipmentHeader = TRUE) OR (Summary = TRUE) THEN BEGIN
                                //InvHeader := FALSE;
                                ERROR('Select only One at a time');
                            END;
                        end;
                    }
                    field("Purchase Creditmemo"; CreditmemoHeader)
                    {

                        trigger OnValidate()
                        begin

                            IF (ReceiptHeader = TRUE) OR (InvHeader = TRUE) OR (ShipmentHeader = TRUE) OR (Summary = TRUE) THEN BEGIN
                                //CreditmemoHeader := FALSE;
                                ERROR('Select only One at a time');
                            END;
                        end;
                    }
                    field("Purchase Shipment"; ShipmentHeader)
                    {
                        Visible = false;

                        trigger OnValidate()
                        begin

                            IF (ReceiptHeader = TRUE) OR (InvHeader = TRUE) OR (CreditmemoHeader = TRUE) OR (Summary = TRUE) THEN BEGIN
                                //ShipmentHeader := FALSE;
                                ERROR('Select only One at a time');
                            END;
                        end;
                    }
                    field(Summary; Summary)
                    {
                        Caption = 'Summary';
                        Visible = false;

                        trigger OnValidate()
                        begin
                            /*
                            IF (ReceiptHeader = TRUE) OR (InvHeader=TRUE) OR (CreditmemoHeader=TRUE) OR (ShipmentHeader=TRUE) THEN BEGIN
                               //Summary := FALSE;
                               ERROR('Select only One at a time');
                            END;
                                    */

                        end;
                    }
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

    var
        PRH: Record 120;
        PIH: Record 122;
        PCH: Record 124;
        RSH: Record 6650;
        Vendor: Record 23;
        ReceiptHeader: Boolean;
        InvHeader: Boolean;
        CreditmemoHeader: Boolean;
        ShipmentHeader: Boolean;
        CompInfo: Record 79;
        IC: Record 5722;
        //PG: Record "5723";
        Summary: Boolean;
        sum1: Boolean;
        sum2: Boolean;
        PR_ExcisePer: Decimal;
        PI_ExcisePer: Decimal;
        PC_ExcisePer: Decimal;
        RR_ExcisePer: Decimal;
        VRN: Code[20];
        REC_PRH: Record 120;
        CF: Decimal;
        PRHDUC: Decimal;
        PRHBEDAMT: Decimal;
        PRHECESSAMT: Decimal;
        PRHSHEECESS: Decimal;
        PRHADCVAT: Decimal;
        REC_PIH: Record 122;
        PIHDUC: Decimal;
        PIHBEDAMT: Decimal;
        PIHECESSAMT: Decimal;
        PIHSHEECESS: Decimal;
        PIHEXCISEAMT: Decimal;
        PIHTAXAMT: Decimal;
        PIHATV: Decimal;
        PIHCTV: Decimal;
        PIHSTA: Decimal;
        PIHADCVAT: Decimal;
        PIHTDS: Decimal;
        REC_PC: Record 124;
        PCDUC: Decimal;
        PCBEDAMT: Decimal;
        PCADCVAT: Decimal;
        PCECESSAMT: Decimal;
        PCSHEECESS: Decimal;
        PCEXCISEAMT: Decimal;
        PCTAXAMT: Decimal;
        PCATV: Decimal;
        PCCTV: Decimal;
        PCSTA: Decimal;
        REC_RSH: Record 6650;
        RSLDUC: Decimal;
        RSLBEDAMT: Decimal;
        RSLADCVAT: Decimal;
        RSLECESSAMT: Decimal;
        RSLSHEECESS: Decimal;
        RSLEXCISEAMT: Decimal;
        RSLTAXAMT: Decimal;
        PRHEXCIEAMT: Decimal;
        PRKEXCISEPER: Decimal;
        PRAMT: Decimal;
        PIEXEPER: Decimal;
        PIHAmt: Decimal;
        PCAmt: Decimal;
        PCEXEPER: Decimal;
        RSLAmt: Decimal;
        RSLEXEPER: Decimal;
        recPurchRH: Record 120;
        recPurchIH: Record 122;
        recPurchCH: Record 124;
        recReturnH: Record 6650;
        recShipmentAg: Record 291;
        "---": Integer;
        DimSetEntry: Record 480;
        CostCentrAmt: Text;
        CostCentrAmt2: Text;
        CostCentrAmt3: Text;
        CostCentrAmt4: Text;
        CostCentrAmt5: Text;
        CostCentrAmt6: Text;
        "-------------------------------Purch. Rcpt. Line-----------": Integer;
        vRcptDoc: Code[50];
        VrecNo: Code[50];
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        vSESS: Decimal;
        vUTGST: Decimal;
        "-------------------------------Purch. Inv. Line----------------------": Integer;
        DetailGSTEntry1: Record "Detailed GST Ledger Entry";
        vCGST1: Decimal;
        vCGSTRate1: Decimal;
        vSGST1: Decimal;
        vSGSTRate1: Decimal;
        vIGST1: Decimal;
        vIGSTRate1: Decimal;
        vSESS1: Decimal;
        vUTGST1: Decimal;
        "--------------------------------Purch. Cr. Memo Line-------------": Integer;
        DetailGSTEntry2: Record "Detailed GST Ledger Entry";
        vCGST2: Decimal;
        vCGSTRate2: Decimal;
        vSGST2: Decimal;
        vSGSTRate2: Decimal;
        vIGST2: Decimal;
        vIGSTRate2: Decimal;
        vSESS2: Decimal;
        vUTGST2: Decimal;
        "-----------------------------------Return Shipment Line-----------": Integer;
        DetailGSTEntry3: Record "Detailed GST Ledger Entry";
        vCGST3: Decimal;
        vCGSTRate3: Decimal;
        vSGST3: Decimal;
        vSGSTRate3: Decimal;
        vIGST3: Decimal;
        vIGSTRate3: Decimal;
        vSESS3: Decimal;
        vUTGST3: Decimal;
}


report 50194 "Sales Register -RSPL"
{
    // Report Name:Sales Register
    // Company:KAPL
    // Prepared By: Gunjan Mittal
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesRegisterRSPL.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number);
            MaxIteration = 1;
            column(Comp_Name; comprec.Name)
            {
            }
            column(Comp_add; comprec.Address)
            {
            }
            column(Comp_Add1; comprec."Address 2")
            {
            }
            column(Comp_City; comprec.City)
            {
            }
            column(SSL; SSL)
            {
            }
            column(SIL; SIL)
            {
            }
            column(RSL; RSL)
            {
            }
            column(CML; CML)
            {
            }
            column(Summary; Summary)
            {
            }

            trigger OnAfterGetRecord()
            begin
                comprec.SETRANGE(comprec.Name);
                IF comprec.FIND('-') THEN;
            end;
        }
        dataitem("Sales Shipment Line"; "Sales Shipment Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.")
                                WHERE(Quantity = FILTER(<> 0));
            column(CF; CF)
            {
            }
            column(cust_CST; cust_CST)
            {
            }
            column(cust_VAT; cust_VAT)
            {
            }
            column(sum1; sum1)
            {
            }
            column(SSL_ADCVAT; ADCVAT)
            {
            }
            column(SSL_Postingdate; "Sales Shipment Line"."Posting Date")
            {
            }
            column(SSL_DocNo; "Sales Shipment Line"."Document No.")
            {
            }
            column(SSL_LineNo; "Sales Shipment Line"."No.")
            {
            }
            column(SHL_Description; "Sales Shipment Line".Description)
            {
            }
            column(SSL_Quantity; "Sales Shipment Line".Quantity)
            {
            }
            column(SSL_UOM; "Sales Shipment Line"."Unit of Measure")
            {
            }
            column(SSL_UP; UPLCY)
            {
            }
            column(SSL_Amount; Amt)
            {
            }
            column(SSl_BEDAmount; BEDAMT)
            {
            }
            column(SSL_ecessAmount; ECESSAMT)
            {
            }
            column(SSL_SheAmount; SHEECESS)
            {
            }
            column(SSL_ExciseAmount; EXCISEAMT)
            {
            }
            column(SSL_TaxPercentage; 0)// "Sales Shipment Line"."Tax %")
            {
            }
            column(SSL_Dim1; "Sales Shipment Line"."Shortcut Dimension 1 Code")
            {
            }
            column(SSL_Dim2; "Sales Shipment Line"."Shortcut Dimension 2 Code")
            {
            }
            column(SSL_Discount; "Sales Shipment Line"."Line Discount %")
            {
            }
            column(SSHNAME; SSHNAME)
            {
            }
            column(ICDesShp; IC.Description)
            {
            }
            column(PG_DesShp; 'PG.Description')
            {
            }
            column(SSL_CurrencyCode; "Sales Shipment Line"."Currency Code")
            {
            }
            column(SLDES; SLDES)
            {
            }
            column(FORMC; FORMC)
            {
            }
            column(FORMNO; FORMNO)
            {
            }
            column(SSL_LRNO; SSL_LRNO)
            {
            }
            column(SSL_LRDate; SSL_LRDate)
            {
            }
            column(SHL_Transporter; SHL_Transporter)
            {
            }
            column(EXPER; EXEPER)
            {
            }
            column(GST_Group_Code1; "Sales Shipment Line"."GST Group Code")
            {
            }
            column(DetailGSTEntry_InvoiceType; 'DetailGSTEntry."Invoice Type"')
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
            column(vUTGST; vUTGST)
            {
            }
            column(vSESS; vSESS)
            {
            }
            column(GST_Reg_No1; GST_Reg_No1)
            {
            }
            column(User_ID; USERID)
            {
            }

            trigger OnAfterGetRecord()
            begin
                SSBEDP := 0;
                SSECESSP := 0;
                SSSHECESSP := 0;
                SSPER := 0;
                IF Type = Type::Item THEN BEGIN
                    SalesshpLine1.RESET;
                    //SalesInvLine1.SETRANGE(SalesInvLine1.Type,SalesInvLine1.Type::Item);
                    SalesshpLine1.SETRANGE(SalesshpLine1."Document No.", "Document No.");
                    IF SalesshpLine1.FINDFIRST THEN BEGIN
                        /* ExciseEntry.RESET;
                        ExciseEntry.SETRANGE(ExciseEntry."Document Type", ExciseEntry."Document Type"::Invoice);
                        //ExciseEntry.SETRANGE(ExciseEntry.Type,ExciseEntry.Type::Sale);
                        ExciseEntry.SETRANGE(ExciseEntry."Document No.", SalesshpLine1."Document No.");
                        ExciseEntry.SETRANGE(ExciseEntry."Item No.", SalesshpLine1."No.");
                        ExciseEntry.SETRANGE(ExciseEntry."Excise Bus. Posting Group", SalesshpLine1."Excise Bus. Posting Group");
                        ExciseEntry.SETRANGE(ExciseEntry."Excise Prod. Posting Group", SalesshpLine1."Excise Prod. Posting Group");
                        IF ExciseEntry.FINDFIRST THEN BEGIN
                            IF SalesshpLine1."Excise Amount" <> 0 THEN BEGIN
                                SSBEDP := ExciseEntry."BED %";
                                SSECESSP := (ExciseEntry."BED %" * ExciseEntry."eCess %") / 100;
                                //MESSAGE('%1',ECESSP);
                                SSSHECESSP := (ExciseEntry."BED %" * ExciseEntry."SHE Cess %") / 100;
                                //MESSAGE('%1',SHECESSP);
                                SSPER := SSBEDP + SSECESSP + SSSHECESSP;
                            END;
                        END; */
                    END;
                END;



                //GUNJAN-START
                Rec_ShipHeader.RESET;
                Rec_ShipHeader.SETRANGE(Rec_ShipHeader."No.", "Sales Shipment Line"."Document No.");
                IF Rec_ShipHeader.FINDFIRST THEN
                    CF := Rec_ShipHeader."Currency Factor";

                IF CF <> 0 THEN BEGIN
                    UPLCY := "Sales Shipment Line"."Unit Price" / CF;
                    Amt := ("Sales Shipment Line".Quantity * "Sales Shipment Line"."Unit Price") / CF;
                    BEDAMT := 0;// "Sales Shipment Line"."BED Amount" / CF;
                    ADCVAT := 0;// "Sales Shipment Line"."ADC VAT Amount" / CF;
                    ECESSAMT := 0;// "Sales Shipment Line"."eCess Amount" / CF;
                    SHEECESS := 0;// "Sales Shipment Line"."SHE Cess Amount" / CF;
                    EXCISEAMT := 0;// "Sales Shipment Line"."Excise Amount" / CF;
                    EXEPER := SSPER / CF;
                END
                ELSE BEGIN
                    UPLCY := "Sales Shipment Line"."Unit Price";
                    Amt := ("Sales Shipment Line".Quantity * "Sales Shipment Line"."Unit Price");
                    BEDAMT := 0;// "Sales Shipment Line"."BED Amount";
                    ADCVAT := 0;//"Sales Shipment Line"."ADC VAT Amount";
                    ECESSAMT := 0;// "Sales Shipment Line"."eCess Amount";
                    SHEECESS := 0;// "Sales Shipment Line"."SHE Cess Amount";
                    EXCISEAMT := 0;// "Sales Shipment Line"."Excise Amount";
                    EXEPER := SSPER;
                END;
                //GUNJAN-END




                //Gunjan-Start

                //IF SSH.GET("Sales Shipment Line"."No.") THEN
                SSH.RESET;
                SSH.SETRANGE(SSH."No.", "Sales Shipment Line"."Document No.");
                IF SSH.FINDFIRST THEN
                    SSHNAME := SSH."Sell-to Customer Name";

                //IF PG.GET("Product Group Code") THEN;
                IF IC.GET("Item Category Code") THEN;

                //IF IC.GET("Sales Shipment Line"."No.") THEN
                //ICDesShp := IC.Description;

                IF IPG.GET("Sales Shipment Line"."Posting Group") THEN
                    SLDES := IPG.Description;

                //IF  ("Sales Shipment Line".Quantity*"Sales Shipment Line"."Unit Price") <>0 THEN
                //EXPER := ("Sales Shipment Line"."Excise Amount"/("Sales Shipment Line".Quantity*"Sales Shipment Line"."Unit Price"))*100;
                //EXPER := SalesLine."Excise Amount"/ "Sales Line".Amount*100;

                Sale_SH.RESET;
                Sale_SH.SETRANGE(Sale_SH."No.", "Sales Shipment Line"."Document No.");
                FORMC := '';//SSH."Form Code";
                FORMNO := '';// SSH."Form No.";
                SSL_LRNO := Sale_SH."LR/RR No.";
                SSL_LRDate := Sale_SH."LR/RR Date";






                CLEAR(SHL_Transporter);
                CLEAR(SCODE);
                Sale_SH.RESET;
                Sale_SH.SETRANGE(Sale_SH."No.", "Sales Shipment Line"."Document No.");
                IF Sale_SH.FINDFIRST THEN BEGIN
                    SCODE := Sale_SH."Shipping Agent Code";
                    IF SAC.GET(SCODE) THEN
                        SHL_Transporter := SAC.Name;
                END;


                SSH.RESET;
                SSH.SETRANGE(SSH."No.", "Sales Shipment Line"."Document No.");
                Rec_Cust.RESET;
                Rec_Cust.SETRANGE(Rec_Cust."No.", SSH."Sell-to Customer No.");
                IF Rec_Cust.FINDFIRST THEN
                    cust_CST := '';// Rec_Cust."C.S.T. No.";
                cust_VAT := '';// Rec_Cust."T.I.N. No.";
                GST_Reg_No1 := Rec_Cust."GST Registration No.";
                //Gunjan-End

                //>>GST
                CLEAR(vCGST);
                CLEAR(vSGST);
                CLEAR(vIGST);
                CLEAR(vUTGST);
                CLEAR(vSESS);
                DetailGSTEntry.RESET;
                DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DetailGSTEntry.SETRANGE("Document No.", "Document No.");
                DetailGSTEntry.SETRANGE("No.", "No.");
                DetailGSTEntry.SETRANGE("Document Line No.", "Line No.");
                IF DetailGSTEntry.FINDSET THEN
                    REPEAT
                        IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST := ABS(DetailGSTEntry."GST Amount");
                        END;
                        IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST := ABS(DetailGSTEntry."GST Amount");
                        END;
                        IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST := ABS(DetailGSTEntry."GST Amount");
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

            trigger OnPreDataItem()
            begin
                //Gunjan


                IF Summary = TRUE THEN
                    sum1 := TRUE
                ELSE
                    sum1 := FALSE;

                IF Summary = TRUE THEN
                    Sum2 := TRUE
                ELSE
                    Sum2 := FALSE;

                IF Summary = TRUE THEN
                    sum3 := TRUE
                ELSE
                    sum3 := FALSE;

                IF Summary = TRUE THEN
                    sum4 := TRUE
                ELSE
                    sum4 := FALSE;
            end;
        }
        dataitem("Sales Invoice Line"; "Sales Invoice Line")
        {
            DataItemTableView = WHERE("No." = FILTER(<> 144900),
                                      Quantity = FILTER(<> 0));
            column(SIHEXEPER; SIHEXEPER)
            {
            }
            column(PER; PER)
            {
            }
            column(amt_ty; amt_ty)
            {
            }
            column(AT2; ABS(AT2))
            {
            }
            column(CFORMDATE; CFORMDATE)
            {
            }
            column(C_VAT; C_VAT)
            {
            }
            column(C_CST; C_CST)
            {
            }
            column(Sum2; Sum2)
            {
            }
            column(IC_DecSIL; IC.Description)
            {
            }
            column(PG_DesSIL; 'PG.Description')
            {
            }
            column(SIL_ADCVAT; SIHADCVAT)
            {
            }
            column(SIL_PostingDate; "Sales Invoice Line"."Posting Date")
            {
            }
            column(SIL_DocNo; "Sales Invoice Line"."Document No.")
            {
            }
            column(SIL_LineNo; "Sales Invoice Line"."No.")
            {
            }
            column(SIL_Des; "Sales Invoice Line".Description)
            {
            }
            column(SIL_Quantity; "Sales Invoice Line".Quantity)
            {
            }
            column(HSN_SAC_Code_SIL; "Sales Invoice Line"."HSN/SAC Code")
            {
            }
            column(GST_Group_Code; "Sales Invoice Line"."GST Group Code")
            {
            }
            column(SIL_UOM; "Sales Invoice Line"."Unit of Measure")
            {
            }
            column(SIL_UnitCost; SIHUPLCY)
            {
            }
            column(SIL_Amount; SIHAmt)
            {
            }
            column(SIL_BEDAmount; SIHBEDAMT)
            {
            }
            column(SIL_ECESSAmount; SIHECESSAMT)
            {
            }
            column(SIL_SHEAmount; SIHSHEECESS)
            {
            }
            column(SIL_ExciseAmount; SIHEXCISEAMT)
            {
            }
            column(SIL_TaxPer; 0)//"Sales Invoice Line"."Tax %")
            {
            }
            column(SIL_TaxAmount; SIHTAXAMT)
            {
            }
            column(SIL_Discount; "Sales Invoice Line"."Line Discount %")
            {
            }
            column(SIL_ServiceTaxAmount; SIHSERVCETAXAMT)
            {
            }
            column(SIL_AmounttoVendor; SIHATC)
            {
            }
            column(SIL_ChargestoVendor; SIHCTC)
            {
            }
            column(ProductGroupCode_SalesInvoiceLine; '"Sales Invoice Line"."Product Group Code"')
            {
            }
            column(FOBUSD_SalesInvoiceLine; '')
            {
            }
            column(LineAmountUSD_SalesInvoiceLine; '')
            {
            }
            column(LineAmountINR_SalesInvoiceLine; '')
            {
            }
            column(LineAmount_SalesInvoiceLine; "Sales Invoice Line"."Line Amount")
            {
            }
            column(BEDAmount_SalesInvoiceLine; 0)// "Sales Invoice Line"."BED Amount")
            {
            }
            column(AmountIncludingVAT_SalesInvoiceLine; "Sales Invoice Line"."Amount Including VAT")
            {
            }
            column(AmountToCustomer_SalesInvoiceLine; 0)//"Sales Invoice Line"."Amount To Customer")
            {
            }
            column(ChargesToCustomer_SalesInvoiceLine; 0)// "Sales Invoice Line"."Charges To Customer")
            {
            }
            column(SIL_Dim1; "Sales Invoice Line"."Shortcut Dimension 1 Code")
            {
            }
            column(SIL_Dim2; "Sales Invoice Line"."Shortcut Dimension 2 Code")
            {
            }
            column(SILName; SILName)
            {
            }
            column(SILDES; SILDES)
            {
            }
            column(SIL_FormNo; SIL_FormNo)
            {
            }
            column(SIL_FormC; SIL_FormC)
            {
            }
            column(SIL_TDSAmount; SIHTDSAMT)
            {
            }
            column(SIL_LRNo; SIL_LRNo)
            {
            }
            column(SIL_LRDate; SIL_LRDate)
            {
            }
            column(SIL_Transporter; SIL_Transporter)
            {
            }
            column(SIL_CurrencyCode; SIL_CurrencyCode)
            {
            }
            column(DetailGSTEntry1_InvoiceType; 'DetailGSTEntry1."Invoice Type"')
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
            column(vUTGST1; vUTGST1)
            {
            }
            column(vSESS1; vSESS1)
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
            column(vUTGSTRate1; vUTGSTRate1)
            {
            }
            column(vSESSRate1; vSESSRate1)
            {
            }
            column(GST_Reg_No; GST_Reg_No)
            {
            }

            trigger OnAfterGetRecord()
            begin

                BEDP := 0;
                ECESSP := 0;
                SHECESSP := 0;
                PER := 0;
                IF Type = Type::Item THEN BEGIN
                    /* ExciseEntry.RESET;
                    ExciseEntry.SETRANGE(ExciseEntry."Document Type", ExciseEntry."Document Type"::Invoice);
                    //ExciseEntry.SETRANGE(ExciseEntry.Type,ExciseEntry.Type::Sale);
                    ExciseEntry.SETRANGE(ExciseEntry."Document No.", "Document No.");
                    ExciseEntry.SETRANGE(ExciseEntry."Item No.", "No.");
                    ExciseEntry.SETRANGE(ExciseEntry."Excise Bus. Posting Group", "Excise Bus. Posting Group");
                    ExciseEntry.SETRANGE(ExciseEntry."Excise Prod. Posting Group", "Excise Prod. Posting Group");
                    IF ExciseEntry.FINDFIRST THEN BEGIN
                        IF "Excise Amount" <> 0 THEN BEGIN
                            BEDP := ExciseEntry."BED %";
                            ECESSP := (ExciseEntry."BED %" * ExciseEntry."eCess %") / 100;
                            //MESSAGE('%1',ECESSP);
                            SHECESSP := (ExciseEntry."BED %" * ExciseEntry."SHE Cess %") / 100;
                            //MESSAGE('%1',SHECESSP);
                            PER := BEDP + ECESSP + SHECESSP;
                        END;
                    END; */
                END;





                //GUNJAN-START
                Rec_SIH.RESET;
                Rec_SIH.SETRANGE(Rec_SIH."No.", "Sales Invoice Line"."Document No.");
                IF Rec_SIH.FINDFIRST THEN BEGIN
                    SIHCF := Rec_SIH."Currency Factor";
                    IF SIHCF <> 0 THEN BEGIN
                        SIHUPLCY := "Sales Invoice Line"."Unit Price" / SIHCF;
                        SIHAmt := ("Sales Invoice Line".Quantity * "Sales Invoice Line"."Unit Price") / SIHCF;
                        SIHBEDAMT := 0;//"Sales Invoice Line"."BED Amount" / SIHCF;
                        SIHADCVAT := 0;// "Sales Invoice Line"."ADC VAT Amount" / SIHCF;
                        SIHECESSAMT := 0;// "Sales Invoice Line"."eCess Amount" / SIHCF;
                        SIHSHEECESS := 0;//"Sales Invoice Line"."SHE Cess Amount" / SIHCF;
                        SIHEXCISEAMT := 0;//"Sales Invoice Line"."Excise Amount" / SIHCF;
                        SIHTAXAMT := 0;// "Sales Invoice Line"."Tax Amount" / SIHCF;
                        SIHSERVCETAXAMT := 0;// "Sales Invoice Line"."Service Tax Amount" / SIHCF;
                        /*IF Rec_SIH."Export Document"=TRUE THEN
                          SIHATC :="Sales Invoice Line".Amount/SIHCF
                        ELSE
                          SIHATC :="Sales Invoice Line"."Amount To Customer"/SIHCF; */
                        SIHCTC := 0;// "Sales Invoice Line"."Charges To Customer" / SIHCF;
                        SIHTDSAMT := 0;// "Sales Invoice Line"."TDS/TCS Amount" / SIHCF;
                        //SIHEXEPER :=PER/SIHCF;
                    END
                    ELSE BEGIN
                        SIHUPLCY := "Sales Invoice Line"."Unit Price";
                        SIHAmt := ("Sales Invoice Line".Quantity * "Sales Invoice Line"."Unit Price");
                        SIHBEDAMT := 0;//"Sales Invoice Line"."BED Amount";
                        SIHADCVAT := 0;//"Sales Invoice Line"."ADC VAT Amount";
                        SIHECESSAMT := 0;// "Sales Invoice Line"."eCess Amount";
                        SIHSHEECESS := 0;// "Sales Invoice Line"."SHE Cess Amount";
                        SIHEXCISEAMT := 0;// "Sales Invoice Line"."Excise Amount";
                        SIHTAXAMT := 0;// "Sales Invoice Line"."Tax Amount";
                        SIHSERVCETAXAMT := 0;//"Sales Invoice Line"."Service Tax Amount";
                        /*IF Rec_SIH."Export Document"=TRUE THEN
                          SIHATC :="Sales Invoice Line".Amount
                        ELSE
                          SIHATC :="Sales Invoice Line"."Amount To Customer";
                        */
                        SIHCTC := 0;//= "Sales Invoice Line"."Charges To Customer";
                        SIHTDSAMT := 0;// "Sales Invoice Line"."TDS/TCS Amount";
                        //SIHEXEPER :=PER;
                    END;
                    //GUNJAN-END
                END;




                BEDA := 0;
                ECESSA := 0;
                SHECESSA := 0;
                SHADC := 0;
                AT2 := 0;
                //IF Type=Type::Item THEN
                //BEGIN
                //Rec_SIH.GET(SalesInvLine2."Document No.");
                SalesInvLine2.RESET;
                //SalesInvLine1.SETRANGE(SalesInvLine1.Type,SalesInvLine1.Type::Item);
                SalesInvLine2.SETRANGE(SalesInvLine2."Document No.", "Document No.");
                IF SalesInvLine2.FINDFIRST THEN BEGIN
                    /*  ExciseEntry2.RESET;
                     ExciseEntry2.SETRANGE(ExciseEntry2."Document Type", ExciseEntry."Document Type"::Invoice);
                     //ExciseEntry.SETRANGE(ExciseEntry.Type,ExciseEntry.Type::Sale);
                     ExciseEntry2.SETRANGE(ExciseEntry2."Document No.", SalesInvLine2."Document No.");
                     ExciseEntry2.SETRANGE(ExciseEntry2."Item No.", SalesInvLine2."No.");
                     ExciseEntry2.SETRANGE(ExciseEntry2."Excise Bus. Posting Group", SalesInvLine2."Excise Bus. Posting Group");
                     ExciseEntry2.SETRANGE(ExciseEntry2."Excise Prod. Posting Group", SalesInvLine2."Excise Prod. Posting Group");
                     IF ExciseEntry2.FINDFIRST THEN BEGIN
                         IF (SalesInvLine2."Direct Debit To PLA / RG" = TRUE) THEN BEGIN
                             BEDA := ExciseEntry2."BED Amount";
                             ECESSA := ExciseEntry2."eCess Amount";
                             //MESSAGE('%1',ECESSP);
                             SHECESSA := ExciseEntry2."SHE Cess Amount";
                             //MESSAGE('%1',SHECESSP);
                             SHADC := ExciseEntry2."ADC VAT Amount";
                             AT2 := BEDA + ECESSA + SHECESSA + SHADC;
                             SIHBEDAMT := 0;
                             SIHADCVAT := 0;
                             SIHECESSAMT := 0;
                             SIHSHEECESS := 0;
                             SIHEXCISEAMT := 0;
                             PER := 0;

                         END;
                     END; */
                END;
                //END;





                /*
                //GUNJAN-START
                AT2:=0;
                REC_SALES_INVOICE_HEADER.RESET;
                REC_SALES_INVOICE_HEADER.SETRANGE(REC_SALES_INVOICE_HEADER."No.","Sales Invoice Line"."Document No.");
                IF REC_SALES_INVOICE_HEADER.FINDFIRST THEN BEGIN
                  IF REC_SALES_INVOICE_HEADER."Export Document"=TRUE THEN
                  BEGIN
                    AT2 :=SIHBEDAMT+ SIHADCVAT+ SIHECESSAMT+ SIHSHEECESS;
                    SIHBEDAMT:=0;
                    SIHADCVAT:=0;
                    SIHECESSAMT:=0;
                    SIHSHEECESS:=0;
                    SIHEXCISEAMT:=0;
                    SIHEXEPER:=0;
                  END;
                END;
                //GUNJAN-END
                */

                SIH.RESET;
                SIH.SETRANGE(SIH."No.", "Sales Invoice Line"."Document No.");
                IF SIH.FINDFIRST THEN;
                SILName := SIH."Sell-to Customer Name";

                //IF PG.GET("Product Group Code") THEN;
                IF IC.GET("Item Category Code") THEN;

                IF IPG.GET("Sales Invoice Line"."Posting Group") THEN
                    SILDES := IPG.Description;

                //IF "Sales Invoice Line".Amount <>0 THEN
                //SIL_EXPER := "Sales Invoice Line"."Excise Amount"/"Sales Invoice Line".Amount*100;

                Sales_IH.RESET;
                Sales_IH.SETRANGE(Sales_IH."No.", "Sales Invoice Line"."Document No.");
                IF Sales_IH.FINDFIRST THEN BEGIN
                    SIL_FormC := '';// Sales_IH."Form Code";
                    SIL_FormNo := '';// Sales_IH."Form No.";
                    SIL_LRNo := Sales_IH."LR/RR No.";
                    SIL_LRDate := Sales_IH."LR/RR Date";
                    SIL_CurrencyCode := Sales_IH."Currency Code";
                    // CFORMDATE := Sales_IH."C form Date";
                    SIL_SCODE := Sales_IH."Shipping Agent Code";
                END;

                CLEAR(SIL_Transporter);
                CLEAR(SIL_SCODE);

                Sales_IH.RESET;
                Sales_IH.SETRANGE(Sales_IH."No.", "Sales Invoice Line"."Document No.");
                IF Sales_IH.FINDFIRST THEN BEGIN
                    SIL_SCODE := Sales_IH."Shipping Agent Code";
                    IF SAC.GET(SIL_SCODE) THEN
                        SIL_Transporter := SAC.Name;
                END;

                Rec_Cust.RESET;
                Rec_Cust.SETRANGE(Rec_Cust."No.", "Sales Invoice Line"."Sell-to Customer No.");
                IF Rec_Cust.FINDFIRST THEN
                    C_CST := '';// Rec_Cust."C.S.T. No.";
                C_VAT := '';//Rec_Cust."T.I.N. No.";
                GST_Reg_No := Rec_Cust."GST Registration No.";


                amt_ty := 0;
                rec_sales_inv_line.RESET;
                //rec_sales_inv_line.SETRANGE(rec_sales_inv_line."Document No.",rec_sales_inv_head."No.");
                rec_sales_inv_line.SETRANGE(rec_sales_inv_line."Document No.", "Document No.");
                IF rec_sales_inv_line.FINDFIRST THEN
                    REPEAT
                        IF rec_sales_inv_line.Type = rec_sales_inv_line.Type::"G/L Account" THEN
                            amt_ty += rec_sales_inv_line."Line Amount";
                    UNTIL rec_sales_inv_line.NEXT = 0;


                //>>GST
                CLEAR(vCGST1);
                CLEAR(vSGST1);
                CLEAR(vIGST1);
                CLEAR(vCGSTRate1);
                CLEAR(vSGSTRate1);
                CLEAR(vIGSTRate1);
                CLEAR(vSESS1);
                CLEAR(vUTGST1);
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
                            vUTGSTRate1 := DetailGSTEntry1."GST %";
                        END;
                        IF DetailGSTEntry1."GST Component Code" = 'CESS' THEN BEGIN
                            vSESS1 := ABS(DetailGSTEntry1."GST Amount");
                            vSESSRate1 := DetailGSTEntry1."GST %";
                        END;
                    UNTIL DetailGSTEntry1.NEXT = 0;

                //<<

            end;
        }
        dataitem("Return Shipment Line"; "Return Shipment Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.")
                                WHERE(Quantity = FILTER(<> 0));
            column(RSL_CST; RSL_CST)
            {
            }
            column(RSL_VAT; RSL_VAT)
            {
            }
            column(sum3; sum3)
            {
            }
            column(RSL_CurrencyCode; "Return Shipment Line"."Currency Code")
            {
            }
            column(IC_DesRSL; IC.Description)
            {
            }
            column(PG_DecRSL; 'PG.Description')
            {
            }
            column(RSL_ADCVAT; RSHADCVAT)
            {
            }
            column(GST_Group_Code2; '"Return Shipment Line"."GST Group Code"')
            {
            }
            column(RSL_psotingdate; "Return Shipment Line"."Posting Date")
            {
            }
            column(RSL_DocNo; "Return Shipment Line"."Document No.")
            {
            }
            column(RSL_ItemNo; "Return Shipment Line"."No.")
            {
            }
            column(RSL_Desc; "Return Shipment Line".Description)
            {
            }
            column(RSL_Quantity; "Return Shipment Line".Quantity)
            {
            }
            column(RSL_UOM; "Return Shipment Line"."Unit of Measure")
            {
            }
            column(RSL_UnitCost; RSHUPLCY)
            {
            }
            column(RSL_Amount; RSHAmt)
            {
            }
            column(RSL_BEDAmount; RSHBEDAMT)
            {
            }
            column(RSL_ECESSAmount; RSHECESSAMT)
            {
            }
            column(RSL_SheAmount; RSHSHEECESS)
            {
            }
            column(RSL_ExciseAmount; RSHEXCISEAMT)
            {
            }
            column(RSL_TaxPercentage; 0)// "Return Shipment Line"."Tax %")
            {
            }
            column(RSL_TaxAmount; RSHTAXAMT)
            {
            }
            column(RSL_Discount; "Return Shipment Line"."Line Discount %")
            {
            }
            column(RSL_Dim1; "Return Shipment Line"."Shortcut Dimension 1 Code")
            {
            }
            column(RSL_Dim2; "Return Shipment Line"."Shortcut Dimension 2 Code")
            {
            }
            column(RSLDES; RSLDES)
            {
            }
            column(RSLNAME; RSLNAME)
            {
            }
            column(RSL_FORMC; RSL_FORMC)
            {
            }
            column(RSL_FormNo; RSL_FormNo)
            {
            }
            column(RSL_LRNo; RSL_LRNo)
            {
            }
            column(RSL_LRDate; RSL_LRDate)
            {
            }
            column(RSL_Transporter; RSL_Transporter)
            {
            }
            column(RSL_EXPER; ROUND(RSHEXEPER))
            {
            }
            column(DetailGSTEntry2_InvoiceType; 'DetailGSTEntry2."Invoice Type"')
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
            column(vUTGST2; vUTGST2)
            {
            }
            column(vSESS2; vSESS2)
            {
            }
            column(GST_Reg_No2; GST_Reg_No2)
            {
            }

            trigger OnAfterGetRecord()
            begin
                /*
                RSBEDP:=0;
                RSECESSP:=0;
                RSSHECESSP:=0;
                RSPER:=0;
                
                RSEPS.RESET;
                RSEPS.SETRANGE(RSEPS."Excise Bus. Posting Group","Return Shipment Line"."Excise Bus. Posting Group");
                RSEPS.SETRANGE(RSEPS."Excise Prod. Posting Group","Return Shipment Line"."Excise Prod. Posting Group");
                IF RSEPS.FINDFIRST THEN BEGIN
                  IF "Return Shipment Line"."Excise Amount"<>0 THEN BEGIN
                  //REPEAT
                    RSBEDP := RSEPS."BED %";
                    RSECESSP := (RSEPS."BED %"*RSEPS."eCess %")/100;
                    //MESSAGE('%1',ECESSP);
                    RSSHECESSP :=(RSEPS."BED %"*RSEPS."SHE Cess %")/100;
                    //MESSAGE('%1',SHECESSP);
                    RSPER := RSBEDP+RSECESSP+RSSHECESSP;
                END;
                END;
                  //UNTIL EPS.NEXT=0;
                */


                RSBEDP := 0;
                RSECESSP := 0;
                RSSHECESSP := 0;
                RSPER := 0;
                IF Type = Type::Item THEN BEGIN
                    returnship1.RESET;
                    //SalesInvLine1.SETRANGE(SalesInvLine1.Type,SalesInvLine1.Type::Item);
                    returnship1.SETRANGE(returnship1."Document No.", "Document No.");
                    IF returnship1.FINDFIRST THEN BEGIN
                        /* ExciseEntry.RESET;
                        ExciseEntry.SETRANGE(ExciseEntry."Document Type", ExciseEntry."Document Type"::Invoice);
                        //ExciseEntry.SETRANGE(ExciseEntry.Type,ExciseEntry.Type::Sale);
                        ExciseEntry.SETRANGE(ExciseEntry."Document No.", returnship1."Document No.");
                        ExciseEntry.SETRANGE(ExciseEntry."Item No.", returnship1."No.");
                        ExciseEntry.SETRANGE(ExciseEntry."Excise Bus. Posting Group", returnship1."Excise Bus. Posting Group");
                        ExciseEntry.SETRANGE(ExciseEntry."Excise Prod. Posting Group", returnship1."Excise Prod. Posting Group");
                        IF ExciseEntry.FINDFIRST THEN BEGIN
                            IF returnship1."Excise Amount" <> 0 THEN BEGIN
                                RSBEDP := ExciseEntry."BED %";
                                RSECESSP := (ExciseEntry."BED %" * ExciseEntry."eCess %") / 100;
                                //MESSAGE('%1',ECESSP);
                                RSSHECESSP := (ExciseEntry."BED %" * ExciseEntry."SHE Cess %") / 100;
                                //MESSAGE('%1',SHECESSP);
                                RSPER := RSBEDP + RSECESSP + RSSHECESSP;
                            END;
                        END; */
                    END;
                END;

                //GUNJAN-START
                Rec_RSH.RESET;
                Rec_RSH.SETRANGE(Rec_RSH."No.", "Return Shipment Line"."Document No.");
                IF Rec_RSH.FINDFIRST THEN
                    RSHCF := Rec_RSH."Currency Factor";

                IF RSHCF <> 0 THEN BEGIN
                    RSHUPLCY := "Return Shipment Line"."Unit Cost" / RSHCF;
                    RSHAmt := ("Return Shipment Line".Quantity * "Return Shipment Line"."Unit Cost") / RSHCF;
                    RSHBEDAMT := 0;// "Return Shipment Line"."BED Amount" / RSHCF;
                    RSHADCVAT := 0;// "Return Shipment Line"."ADC VAT Amount" / RSHCF;
                    RSHECESSAMT := 0;// "Return Shipment Line"."eCess Amount" / RSHCF;
                    RSHSHEECESS := 0;// "Return Shipment Line"."SHE Cess Amount" / RSHCF;
                    RSHEXCISEAMT := 0;// "Return Shipment Line"."Excise Amount" / RSHCF;
                    RSHTAXAMT := 0;// "Return Shipment Line"."Tax Amount" / RSHCF;
                    RSHEXEPER := RSPER / RSHCF;
                END
                ELSE BEGIN
                    RSHUPLCY := "Return Shipment Line"."Unit Cost";
                    RSHAmt := ("Return Shipment Line".Quantity * "Return Shipment Line"."Unit Cost");
                    RSHBEDAMT := 0;// "Return Shipment Line"."BED Amount";
                    RSHADCVAT := 0;// "Return Shipment Line"."ADC VAT Amount";
                    RSHECESSAMT := 0;// "Return Shipment Line"."eCess Amount";
                    RSHSHEECESS := 0;// "Return Shipment Line"."SHE Cess Amount";
                    RSHEXCISEAMT := 0;// "Return Shipment Line"."Excise Amount";
                    RSHTAXAMT := 0;// "Return Shipment Line"."Tax Amount";
                    RSHEXEPER := RSPER;
                END;
                //GUNJAN-END


                RSH.RESET;
                RSH.SETRANGE(RSH."No.", "Return Shipment Line"."Document No.");
                IF RSH.FINDFIRST THEN
                    RSLNAME := RSH."Ship-to Name";

                //IF PG.GET("Product Group Code") THEN;
                IF IC.GET("Item Category Code") THEN;

                IF IPG.GET("Sales Invoice Line"."Posting Group") THEN
                    RSLDES := IPG.Description;

                //RSL_EXPER := "Return Shipment Line"."Excise Amount"/

                Retrun_SH.RESET;
                Retrun_SH.SETRANGE(Retrun_SH."No.", "Return Shipment Line"."Document No.");
                RSL_FORMC := '';// Retrun_SH."Form Code";
                RSL_FormNo := '';//Retrun_SH."Form No.";
                //RSL_LRNo := Retrun_SH."LR No.";
                //RSL_LRDate := Retrun_SH."LR Date";

                CLEAR(RSL_Transporter);
                CLEAR(RSL_SCODE);
                Retrun_SH.RESET;
                Retrun_SH.SETRANGE(Retrun_SH."No.", "Return Shipment Line"."Document No.");
                IF Retrun_SH.FINDFIRST THEN BEGIN
                    //RSL_SCODE:=Retrun_SH."Shipment Agent Code";
                    IF SAC.GET(RSL_SCODE) THEN
                        RSL_Transporter := SAC.Name;
                END;

                //IF ( "Return Shipment Line".Quantity* "Return Shipment Line"."Unit Price (LCY)")<>0 THEN
                //RSL_EXPER := ( "Return Shipment Line"."Excise Amount"/( "Return Shipment Line".Quantity* "Return Shipment Line"."Unit Price (LCY)"))*100;
                //EXPER := SalesLine."Excise Amount"/ "Sales Line".Amount*100;


                Rec_Cust.RESET;
                Rec_Cust.SETRANGE(Rec_Cust."No.", "Return Shipment Line"."Buy-from Vendor No.");
                IF Rec_Cust.FINDFIRST THEN
                    RSL_VAT := '';// Rec_Cust."T.I.N. No.";
                RSL_CST := '';//Rec_Cust."C.S.T. No.";
                GST_Reg_No2 := Rec_Cust."GST Registration No.";
                //>>GST
                CLEAR(vCGST2);
                CLEAR(vSGST2);
                CLEAR(vIGST2);
                CLEAR(vUTGST2);
                CLEAR(vSESS2);
                DetailGSTEntry2.RESET;
                DetailGSTEntry2.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DetailGSTEntry2.SETRANGE("Document No.", "Document No.");
                DetailGSTEntry2.SETRANGE("No.", "No.");
                DetailGSTEntry2.SETRANGE("Document Line No.", "Line No.");
                IF DetailGSTEntry2.FINDSET THEN
                    REPEAT
                        IF DetailGSTEntry2."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST2 := ABS(DetailGSTEntry2."GST Amount");
                        END;
                        IF DetailGSTEntry2."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST2 := ABS(DetailGSTEntry2."GST Amount");
                        END;
                        IF DetailGSTEntry2."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST2 := ABS(DetailGSTEntry2."GST Amount");
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
        dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
        {
            DataItemTableView = WHERE("No." = FILTER(<> 144900),
                                      Quantity = FILTER(<> 0));
            column(amt_ty2; amt_ty2)
            {
            }
            column(AT4; AT4)
            {
            }
            column(CMH_VAT; CMH_VAT)
            {
            }
            column(CMH_CST; CMH_CST)
            {
            }
            column(sum4; sum4)
            {
            }
            column(IC_DesCML; IC.Description)
            {
            }
            column(PG_DesCML; 'PG.Description')
            {
            }
            column(SCM_ADCVAT; CMHADCVAT)
            {
            }
            column(GST_Group_Code3; "Sales Cr.Memo Line"."GST Group Code")
            {
            }
            column(CM_PostingDate; "Sales Cr.Memo Line"."Posting Date")
            {
            }
            column(CM_DocNo; "Sales Cr.Memo Line"."Document No.")
            {
            }
            column(CM_ItemNo; "Sales Cr.Memo Line"."No.")
            {
            }
            column(CM_Decs; "Sales Cr.Memo Line".Description)
            {
            }
            column(CM_Quantity; "Sales Cr.Memo Line".Quantity)
            {
            }
            column(CM_UOM; "Sales Cr.Memo Line"."Unit of Measure")
            {
            }
            column(CM_UnitCost; CMHUPLCY)
            {
            }
            column(CM_Amount; CMHAmt)
            {
            }
            column(CM_BEDAmount; CMHBEDAMT)
            {
            }
            column(CM_ECESSAmount; CMHECESSAMT)
            {
            }
            column(CM_SheAmount; CMHSHEECESS)
            {
            }
            column(CM_ExciseAmount; CMHEXCISEAMT)
            {
            }
            column(CM_TaxPer; 0)//"Sales Cr.Memo Line"."Tax %")
            {
            }
            column(CM_TaxAmount; CMHTAXAMT)
            {
            }
            column(CM_DiscountAmount; "Sales Cr.Memo Line"."Line Discount Amount")
            {
            }
            column(CM_ServiceTaxAmount; CMHSERVCETAXAMT)
            {
            }
            column(CM_Amounttovendor; CMHATC)
            {
            }
            column(CM_ChargestoVendor; CMHCTC)
            {
            }
            column(CM_Dim1; "Sales Cr.Memo Line"."Shortcut Dimension 1 Code")
            {
            }
            column(CM_Dim2; "Sales Cr.Memo Line"."Shortcut Dimension 2 Code")
            {
            }
            column(CMHNAME; CMHNAME)
            {
            }
            column(CMLDES; CMLDES)
            {
            }
            column(CML_EXPER; ROUND(CMHEXEPER))
            {
            }
            column(SCML_FormC; SCML_FormC)
            {
            }
            column(SCML_FormNo; SCML_FormNo)
            {
            }
            column(SCML_TDSAmount; CMHTDSAMT)
            {
            }
            column(SCML_LRDate; SCML_LRDate)
            {
            }
            column(SCML_LRNo; SCML_LRNo)
            {
            }
            column(SCML_CurrencyCode; SCML_CurrencyCode)
            {
            }
            column(DetailGSTEntry3_InvoiceType; 'DetailGSTEntry3."Invoice Type"')
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
            column(vUTGST3; vUTGST3)
            {
            }
            column(vSESS3; vSESS3)
            {
            }
            column(GST_Reg_No3; GST_Reg_No3)
            {
            }

            trigger OnAfterGetRecord()
            begin
                /*
                SCMBEDP:=0;
                SCMECESSP:=0;
                SCMSHECESSP:=0;
                SCMPER:=0;
                
                SCMEPS.RESET;
                SCMEPS.SETRANGE(SCMEPS."Excise Bus. Posting Group","Sales Cr.Memo Line"."Excise Bus. Posting Group");
                SCMEPS.SETRANGE(SCMEPS."Excise Prod. Posting Group","Sales Cr.Memo Line"."Excise Prod. Posting Group");
                IF SCMEPS.FINDFIRST THEN BEGIN
                  IF "Sales Cr.Memo Line"."Excise Amount"<>0 THEN BEGIN
                  //REPEAT
                    SCMBEDP := SCMEPS."BED %";
                    SCMECESSP := (SCMEPS."BED %"*SCMEPS."eCess %")/100;
                    //MESSAGE('%1',ECESSP);
                    SCMSHECESSP :=(SCMEPS."BED %"*SCMEPS."SHE Cess %")/100;
                    //MESSAGE('%1',SHECESSP);
                    SCMPER := SCMBEDP+SCMECESSP+SCMSHECESSP;
                END;
                END;
                  //UNTIL EPS.NEXT=0;
                */


                SCMBEDP := 0;
                SCMECESSP := 0;
                SCMSHECESSP := 0;
                SCMPER := 0;
                IF Type = Type::Item THEN BEGIN
                    SalescrmLine1.RESET;
                    //SalesInvLine1.SETRANGE(SalesInvLine1.Type,SalesInvLine1.Type::Item);
                    SalescrmLine1.SETRANGE(SalescrmLine1."Document No.", "Document No.");
                    IF SalescrmLine1.FINDFIRST THEN BEGIN
                        /* ExciseEntry.RESET;
                        ExciseEntry.SETRANGE(ExciseEntry."Document Type", ExciseEntry."Document Type"::Invoice);
                        //ExciseEntry.SETRANGE(ExciseEntry.Type,ExciseEntry.Type::Sale);
                        ExciseEntry.SETRANGE(ExciseEntry."Document No.", SalescrmLine1."Document No.");
                        ExciseEntry.SETRANGE(ExciseEntry."Item No.", SalescrmLine1."No.");
                        ExciseEntry.SETRANGE(ExciseEntry."Excise Bus. Posting Group", SalescrmLine1."Excise Bus. Posting Group");
                        ExciseEntry.SETRANGE(ExciseEntry."Excise Prod. Posting Group", SalescrmLine1."Excise Prod. Posting Group");
                        IF ExciseEntry.FINDFIRST THEN BEGIN
                            IF SalescrmLine1."Excise Amount" <> 0 THEN BEGIN
                                SCMBEDP := ExciseEntry."BED %";
                                SCMECESSP := (ExciseEntry."BED %" * ExciseEntry."eCess %") / 100;
                                //MESSAGE('%1',ECESSP);
                                SCMSHECESSP := (ExciseEntry."BED %" * ExciseEntry."SHE Cess %") / 100;
                                //MESSAGE('%1',SHECESSP);
                                SCMPER := SCMBEDP + SCMECESSP + SCMSHECESSP;
                            END;
                        END; */
                    END;
                END;


                //GUNJAN-START
                Rec_CMH.RESET;
                Rec_CMH.SETRANGE(Rec_CMH."No.", "Sales Cr.Memo Line"."Document No.");
                IF Rec_CMH.FINDFIRST THEN
                    CMHCF := Rec_SIH."Currency Factor";

                IF CMHCF <> 0 THEN BEGIN
                    CMHUPLCY := "Sales Cr.Memo Line"."Unit Cost" / CMHCF;
                    CMHAmt := ("Sales Cr.Memo Line".Quantity * "Sales Cr.Memo Line"."Unit Cost") / CMHCF;
                    CMHBEDAMT := 0;// "Sales Cr.Memo Line"."BED Amount" / CMHCF;
                    CMHADCVAT := 0;// "Sales Cr.Memo Line"."ADC VAT Amount" / CMHCF;
                    CMHECESSAMT := 0;//"Sales Cr.Memo Line"."eCess Amount" / CMHCF;
                    CMHSHEECESS := 0;// "Sales Cr.Memo Line"."SHE Cess Amount" / CMHCF;
                    CMHEXCISEAMT := 0;// "Sales Cr.Memo Line"."Excise Amount" / CMHCF;
                    CMHTAXAMT := 0;// "Sales Cr.Memo Line"."Tax Amount" / CMHCF;
                    CMHSERVCETAXAMT := 0;// "Sales Cr.Memo Line"."Service Tax Amount" / CMHCF;
                    CMHATC := 0;// "Sales Cr.Memo Line"."Amount To Customer" / CMHCF;
                    CMHCTC := 0;// "Sales Cr.Memo Line"."Charges To Customer" / CMHCF;
                    CMHTDSAMT := 0;// "Sales Cr.Memo Line"."TDS/TCS Amount" / CMHCF;
                    CMHEXEPER := SCMPER / CMHCF;
                END
                ELSE BEGIN
                    CMHUPLCY := "Sales Cr.Memo Line"."Unit Cost";
                    CMHAmt := ("Sales Cr.Memo Line".Quantity * "Sales Cr.Memo Line"."Unit Cost");
                    CMHBEDAMT := 0;// "Sales Cr.Memo Line"."BED Amount";
                    CMHADCVAT := 0;//"Sales Cr.Memo Line"."ADC VAT Amount";
                    CMHECESSAMT := 0;// "Sales Cr.Memo Line"."eCess Amount";
                    CMHSHEECESS := 0;// "Sales Cr.Memo Line"."SHE Cess Amount";
                    CMHEXCISEAMT := 0;// "Sales Cr.Memo Line"."Excise Amount";
                    CMHTAXAMT := 0;//"Sales Cr.Memo Line"."Tax Amount";
                    CMHSERVCETAXAMT := 0;// "Sales Cr.Memo Line"."Service Tax Amount";
                    CMHATC := 0;//= "Sales Cr.Memo Line"."Amount To Customer";
                    CMHCTC := 0;// "Sales Cr.Memo Line"."Charges To Customer";
                    CMHTDSAMT := 0;//"Sales Cr.Memo Line"."TDS/TCS Amount";
                    CMHEXEPER := SCMPER;
                END;
                //GUNJAN-END


                //GUNJAN-START
                AT4 := 0;
                REC_SALES_CREDIT_MEMO_HEADER.RESET;
                REC_SALES_CREDIT_MEMO_HEADER.SETRANGE(REC_SALES_CREDIT_MEMO_HEADER."No.", "Sales Cr.Memo Line"."Document No.");
                IF REC_SALES_CREDIT_MEMO_HEADER.FINDFIRST THEN BEGIN
                    //IF REC_SALES_CREDIT_MEMO_HEADER."Export Document"=TRUE THEN
                    //BEGIN
                    AT4 := CMHBEDAMT + CMHADCVAT + CMHECESSAMT + CMHSHEECESS;
                    CMHBEDAMT := 0;
                    CMHADCVAT := 0;
                    CMHECESSAMT := 0;
                    CMHSHEECESS := 0;
                    CMHEXCISEAMT := 0;
                    CMHEXEPER := 0;
                    //END;
                END;
                //GUNJAN-END


                CMH.RESET;
                CMH.SETRANGE(CMH."No.", "Sales Cr.Memo Line"."Document No.");
                IF CMH.FINDFIRST THEN
                    CMHNAME := CMH."Sell-to Customer Name";

                //IF PG.GET("Product Group Code") THEN;
                IF IC.GET("Item Category Code") THEN;

                IF IPG.GET("Sales Invoice Line"."Posting Group") THEN
                    CMLDES := IPG.Description;


                //IF "Sales Cr.Memo Line".Amount <>0 THEN
                //CML_EXPER := "Sales Cr.Memo Line"."Excise Amount"/"Sales Cr.Memo Line".Amount*100;


                Sales_CMH.RESET;
                Sales_CMH.SETRANGE(Sales_CMH."No.", "Sales Cr.Memo Line"."Document No.");
                SCML_FormC := '';// Sales_CMH."Form Code";
                SCML_FormNo := '';// Sales_CMH."Form No.";
                // SCML_LRNo := Sales_CMH."LR/RR No.";
                //  SCML_LRDate := Sales_CMH.Specification;
                SCML_CurrencyCode := Sales_CMH."Currency Code";

                Rec_Cust.RESET;
                Rec_Cust.SETRANGE(Rec_Cust."No.", CMH."Sell-to Customer No.");
                IF Rec_Cust.FINDFIRST THEN
                    CMH_VAT := '';// Rec_Cust."T.I.N. No.";
                CMH_CST := '';// Rec_Cust."C.S.T. No.";
                GST_Reg_No3 := Rec_Cust."GST Registration No.";
                /*
                Sales_CMH.RESET;
                Sales_CMH.SETRANGE(Sales_CMH."No.","Sales Cr.Memo Line"."Document No.");
                IF Sales_CMH.FINDFIRST THEN
                CML_SCODE:=Sales_CMH."Shipment Agent Code";
                
                
                IF SAC.GET(SCODE) THEN;
                 CML_Transporter := SAC.Name;
                */


                amt_ty2 := 0;
                rec_sales_credit_memo_line.RESET;
                //rec_sales_inv_line.SETRANGE(rec_sales_inv_line."Document No.",rec_sales_inv_head."No.");
                rec_sales_credit_memo_line.SETRANGE(rec_sales_credit_memo_line."Document No.", "Document No.");
                IF rec_sales_credit_memo_line.FINDFIRST THEN
                    REPEAT
                        IF rec_sales_credit_memo_line.Type = rec_sales_credit_memo_line.Type::"G/L Account" THEN
                            amt_ty2 += rec_sales_credit_memo_line."Line Amount";
                    UNTIL rec_sales_credit_memo_line.NEXT = 0;

                //>>GST
                CLEAR(vCGST3);
                CLEAR(vSGST3);
                CLEAR(vIGST3);
                CLEAR(vUTGST3);
                CLEAR(vSESS3);
                DetailGSTEntry3.RESET;
                DetailGSTEntry3.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DetailGSTEntry3.SETRANGE("Document No.", "Document No.");
                DetailGSTEntry3.SETRANGE("No.", "No.");
                DetailGSTEntry3.SETRANGE("Document Line No.", "Line No.");
                IF DetailGSTEntry3.FINDSET THEN
                    REPEAT
                        IF DetailGSTEntry3."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST3 := ABS(DetailGSTEntry3."GST Amount");
                        END;
                        IF DetailGSTEntry3."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST3 := ABS(DetailGSTEntry3."GST Amount");
                        END;
                        IF DetailGSTEntry3."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST3 := ABS(DetailGSTEntry3."GST Amount");
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

            trigger OnPreDataItem()
            begin
                //CLEAR(CML_Transporter);
                //CLEAR(SCODE);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                group(GROUP1)
                {
                    field(SSL; SSL)
                    {
                        Caption = 'ShipmentLine';
                        Visible = false;
                        ApplicationArea = all;
                        trigger OnValidate()
                        begin


                            IF (CML = TRUE) OR (SIL = TRUE) OR (RSL = TRUE) OR (Summary = TRUE) THEN BEGIN
                                SSL := FALSE;
                                ERROR('Select only One at a time');
                            END;
                        end;
                    }
                    field(SIL; SIL)
                    {
                        Caption = 'InvoiceLine';
                        ApplicationArea = all;
                        trigger OnValidate()
                        begin
                            IF (CML = TRUE) OR (SSL = TRUE) OR (RSL = TRUE) OR (Summary = TRUE) THEN BEGIN
                                SIL := FALSE;
                                ERROR('Select only One at a time');
                            END;
                        end;
                    }
                    field(RSL; RSL)
                    {
                        Caption = 'ReturnShipmentLine';
                        HideValue = true;
                        Visible = false;
                        ApplicationArea = all;
                        trigger OnValidate()
                        begin
                            IF (CML = TRUE) OR (SSL = TRUE) OR (SIL = TRUE) OR (Summary = TRUE) THEN BEGIN
                                RSL := FALSE;
                                ERROR('Select only One at a time');
                            END;
                        end;
                    }
                    field(CML; CML)
                    {
                        Caption = 'CreditMemoLine';
                        ApplicationArea = all;
                        trigger OnValidate()
                        begin
                            IF (RSL = TRUE) OR (SSL = TRUE) OR (SIL = TRUE) OR (Summary = TRUE) THEN BEGIN
                                CML := FALSE;
                                ERROR('Select only One at a time');
                            END;
                        end;
                    }
                    field(Summary; Summary)
                    {
                        Caption = 'Summary';
                        Visible = false;
                        ApplicationArea = all;
                        trigger OnValidate()
                        begin
                            IF (RSL = TRUE) OR (SSL = TRUE) OR (SIL = TRUE) OR (CML = TRUE) THEN BEGIN
                                Summary := FALSE;
                                ERROR('Select only One at a time');
                            END;
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
        SSH: Record 110;
        SSHNAME: Text[100];
        SSL: Boolean;
        SIL: Boolean;
        RSL: Boolean;
        CML: Boolean;
        SIH: Record 112;
        SILName: Text[100];
        RSH: Record 6650;
        RSLNAME: Text[100];
        CMH: Record 114;
        CMHNAME: Text[100];
        comprec: Record 79;
        IC: Record 5722;
        ICDesShp: Text[30];
        IPG: Record 94;
        SLDES: Text[100];
        //PG: Record 5723;
        SILDES: Text[100];
        RSLDES: Text[100];
        CMLDES: Text[100];
        SIL_EXPER: Decimal;
        CML_EXPER: Decimal;
        Sale_SH: Record 110;
        FORMC: Code[20];
        FORMNO: Code[30];
        Sales_IH: Record 112;
        SIL_FormNo: Code[30];
        SIL_FormC: Code[20];
        Retrun_SH: Record 6650;
        RSL_FORMC: Code[20];
        RSL_FormNo: Code[30];
        Sales_CMH: Record 114;
        SCML_FormC: Code[20];
        SCML_FormNo: Code[30];
        SSL_LRNO: Code[20];
        SSL_LRDate: Date;
        SIL_LRNo: Code[20];
        SIL_LRDate: Date;
        RSL_LRNo: Code[20];
        RSL_LRDate: Date;
        SCML_LRNo: Code[20];
        SCML_LRDate: Date;
        SAC: Record 291;
        SHL_Transporter: Text[100];
        SCODE: Code[20];
        SIL_SCODE: Code[20];
        SIL_Transporter: Text[100];
        RSL_SCODE: Code[20];
        RSL_Transporter: Text[100];
        CML_Transporter: Text[100];
        CML_SCODE: Code[20];
        Summary: Boolean;
        SIL_CurrencyCode: Code[20];
        SCML_CurrencyCode: Code[20];
        EXPER: Decimal;
        RSL_EXPER: Decimal;
        sum1: Boolean;
        Sum2: Boolean;
        sum3: Boolean;
        sum4: Boolean;
        Rec_Cust: Record 18;
        cust_CST: Code[40];
        cust_VAT: Code[20];
        C_CST: Code[40];
        C_VAT: Code[20];
        RSL_VAT: Code[20];
        RSL_CST: Code[40];
        CMH_VAT: Code[20];
        CMH_CST: Code[40];
        Rec_ShipLine: Record 111;
        Rec_ShipHeader: Record 110;
        CF: Decimal;
        UPLCY: Decimal;
        Amt: Decimal;
        BEDAMT: Decimal;
        ADCVAT: Decimal;
        ECESSAMT: Decimal;
        SHEECESS: Decimal;
        EXCISEAMT: Decimal;
        Rec_SIH: Record 112;
        SIHCF: Decimal;
        SIHUPLCY: Decimal;
        SIHAmt: Decimal;
        SIHBEDAMT: Decimal;
        SIHADCVAT: Decimal;
        SIHTDSAMT: Decimal;
        SIHECESSAMT: Decimal;
        SIHSHEECESS: Decimal;
        SIHEXCISEAMT: Decimal;
        SIHTAXAMT: Decimal;
        SIHSERVCETAXAMT: Decimal;
        SIHATC: Decimal;
        SIHCTC: Decimal;
        Rec_RSH: Record 6650;
        RSHCF: Decimal;
        RSHUPLCY: Decimal;
        RSHBEDAMT: Decimal;
        RSHADCVAT: Decimal;
        RSHAmt: Decimal;
        RSHSHEECESS: Decimal;
        RSHECESSAMT: Decimal;
        RSHEXCISEAMT: Decimal;
        RSHTAXAMT: Decimal;
        Rec_CMH: Record 114;
        CMHCF: Decimal;
        CMHUPLCY: Decimal;
        CMHAmt: Decimal;
        CMHBEDAMT: Decimal;
        CMHADCVAT: Decimal;
        CMHECESSAMT: Decimal;
        CMHSHEECESS: Decimal;
        CMHEXCISEAMT: Decimal;
        CMHTAXAMT: Decimal;
        CMHSERVCETAXAMT: Decimal;
        CMHATC: Decimal;
        CMHCTC: Decimal;
        CMHTDSAMT: Decimal;
        SIHEXEPER: Decimal;
        RSHEXEPER: Decimal;
        CMHEXEPER: Decimal;
        EXEPER: Decimal;
        CFORMDATE: Date;
        REC_SALES_SHIPMENT_HEADER: Record 110;
        AT: Decimal;
        REC_SALES_INVOICE_HEADER: Record 112;
        AT2: Decimal;
        REC_SALES_CREDIT_MEMO_HEADER: Record 114;
        AT4: Decimal;
        rec_sales_inv_head: Record 112;
        rec_sales_inv_line: Record 113;
        amt_ty: Decimal;
        rec_sales_credit_memo_line: Record 115;
        amt_ty2: Decimal;
        // EBPG: Record 13709;
        //EPPG: Record 13710;
        //EPS: Record 13711;
        BEDP: Decimal;
        ECESSP: Decimal;
        SHECESSP: Decimal;
        PER: Decimal;
        SSBEDP: Decimal;
        SSECESSP: Decimal;
        SSSHECESSP: Decimal;
        SSPER: Decimal;
        // SSEPS: Record 13711;
        RSBEDP: Decimal;
        RSECESSP: Decimal;
        RSSHECESSP: Decimal;
        RSPER: Decimal;
        //RSEPS: Record 13711;
        SCMBEDP: Decimal;
        SCMECESSP: Decimal;
        SCMSHECESSP: Decimal;
        SCMPER: Decimal;
        // SCMEPS: Record 13711;
        SalesInvLine1: Record 113;
        //  ExciseEntry: Record 13712;
        SalesshpLine1: Record 111;
        returnship1: Record 6651;
        SalescrmLine1: Record 115;
        //ExciseEntry2: Record 13712;
        SalesInvLine2: Record 113;
        BEDA: Decimal;
        ECESSA: Decimal;
        SHECESSA: Decimal;
        SHADC: Decimal;
        AMTCUS: Decimal;
        "-------------------------------Sales. Inv. Line----------------------": Integer;
        DetailGSTEntry1: Record "Detailed GST Ledger Entry";
        vCGST1: Decimal;
        vCGSTRate1: Decimal;
        vSGST1: Decimal;
        vSGSTRate1: Decimal;
        vIGST1: Decimal;
        vIGSTRate1: Decimal;
        GST_Reg_No: Code[20];
        vSESS1: Decimal;
        vSESSRate1: Decimal;
        vUTGSTRate1: Decimal;
        vUTGST1: Decimal;
        "-------------------------------SalesShipLine----------------------": Integer;
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        vCGST: Decimal;
        vSGST: Decimal;
        vIGST: Decimal;
        GST_Reg_No1: Code[20];
        vSESS: Decimal;
        vUTGST: Decimal;
        "-------------------------------ReturnShipLine----------------------": Integer;
        DetailGSTEntry2: Record "Detailed GST Ledger Entry";
        vCGST2: Decimal;
        vSGST2: Decimal;
        vIGST2: Decimal;
        GST_Reg_No2: Code[20];
        vSESS2: Decimal;
        vUTGST2: Decimal;
        "-------------------------------SalesCreditLine----------------------": Integer;
        DetailGSTEntry3: Record "Detailed GST Ledger Entry";
        vCGST3: Decimal;
        vSGST3: Decimal;
        vIGST3: Decimal;
        GST_Reg_No3: Code[20];
        vSESS3: Decimal;
        vUTGST3: Decimal;
}


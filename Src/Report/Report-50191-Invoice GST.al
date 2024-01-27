report 50191 "Invoice GST"
{
    DefaultLayout = RDLC;

    RDLCLayout = 'Src/Report Layout/InvoiceGST.rdl';
    Caption = 'Invoice GST';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Permissions = TableData 112 = rm,
                  TableData 113 = r;
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(TaxInvoice15; TaxInvoice15)
            {
            }
            column(Comp_CIN_NO; Rec_Comp."Registration No.")
            {
            }
            column(Comp_Picture; Rec_Comp.Picture)
            {
            }
            column(Comp_Loc_name; Rec_Comp.Name + ' - ' + Location.Name)
            {
            }
            column(Comp_Name; Rec_Comp.Name)
            {
            }
            column(GSTINCOmInfo; Location."GST Registration No.")
            {
            }
            column(LOC_Name; Location.Name)
            {
            }
            column(LOC_ADD; Location.Address + ', ' + Location."Address 2" + ', ' + Location.City + ', ' + Loc_State + ', ' + Location."Post Code" + ', ' + 'INDIA')
            {
            }
            column(LocAddress; Rec_Comp.Address + ', ' + Rec_Comp."Address 2" + ', ' + Rec_Comp.City + ', ' + Rec_Comp."Post Code" + ', ' + 'INDIA')
            {
            }
            column(No_SalesInvoiceHeader; "Sales Invoice Header"."No.")
            {
            }
            column(Comp_Address; Rec_Comp.Address + ', ' + Rec_Comp."Address 2" + ', ' + Rec_Comp.City + ', Maharashtra, ' + Rec_Comp."Post Code" + ', ' + 'INDIA')
            {
            }
            column(Comp_Website; Rec_Comp."Home Page")
            {
            }
            column(Comp_Phone; Rec_Comp."Phone No.")
            {
            }
            column(SIH_No; "Sales Invoice Header"."No.")
            {
            }
            column(OrderNo_SalesInvoiceHeader; "Sales Invoice Header"."Order No.")
            {
            }
            column(PostingDate_SalesInvoiceHeader; "Sales Invoice Header"."Posting Date")
            {
            }
            column(OrderDate_SalesInvoiceHeader; "Sales Invoice Header"."Order Date")
            {
            }
            column(ExternalDocumentNo_SalesInvoiceHeader; "Sales Invoice Header"."External Document No.")
            {
            }
            column(LOC_ECC; 'Location."E.C.C. No."')
            {
            }
            column(LOC_range; 'Location."C.E. Range"')
            {
            }
            column(LOC_Division; 'Location."C.E. Division"')
            {
            }
            column(LOc_Comm; 'Location."C.E. Commissionerate"')
            {
            }
            column(LOC_ServiceTaxNO; 'Location."Service Tax Registration No."')
            {
            }
            column(LOC_TIN; 'Location."L.S.T. No."')
            {
            }
            column(LOC_CST; 'Location."C.S.T No."')
            {
            }
            column(COMP_PAN; Rec_Comp."P.A.N. No.")
            {
            }
            column(lcdate; lcdate)
            {
            }
            column(lcissuebk; lcissuebk)
            {
            }
            column(lcno; lcno)
            {
            }
            column(roadpermitno; roadpermitno)
            {
            }
            column(SIH_VEHNO; "Sales Invoice Header"."Vehicle No.")
            {
            }
            column(LR_NO; "Sales Invoice Header"."LR/RR No." + ' / ' + FORMAT("Sales Invoice Header"."LR/RR Date"))
            {
            }
            column(SIH_PostingDate; "Sales Invoice Header"."Posting Date")
            {
            }
            column(SIH_TimeOfRemoval; "Sales Invoice Header"."Time of Removal")
            {
            }
            column(SIH_DueDate; "Sales Invoice Header"."Due Date")
            {
            }
            column(ShipAgent_Name; shippingagent.Name)
            {
            }
            column(Pay_Desc; PaymentTerms.Description)
            {
            }
            column(Bill_toName; "Sales Invoice Header"."Bill-to Name" + "Sales Invoice Header"."Bill-to Name 2")
            {
            }
            column(Bill_toAddress; "Sales Invoice Header"."Bill-to Address" + ', ' + "Sales Invoice Header"."Bill-to Address 2" + ', ' + "Sales Invoice Header"."Bill-to City" + ' - ' + "Sales Invoice Header"."Bill-to Post Code" + ', ' + Cust_Country)
            {
            }
            column(Cust_State; Cust_State)
            {
            }
            column(CustStateCode; CustStateCode)
            {
            }
            column(VatAct; VatAct)
            {
            }
            column(VPAN; VPAN)
            {
            }
            column(VLST; VLST)
            {
            }
            column(VCST; VCST)
            {
            }
            column(VECC; VECC)
            {
            }
            column(Ship_toName; "Sales Invoice Header"."Ship-to Name" + ', ' + "Sales Invoice Header"."Ship-to Name 2")
            {
            }
            column(Ship_tAddress; "Sales Invoice Header"."Ship-to Address" + ', ' + "Sales Invoice Header"."Ship-to Address 2" + ', ' + "Sales Invoice Header"."Ship-to City" + ', ' + ' -  ' + "Sales Invoice Header"."Ship-to Post Code" + ', ' + Ship_Country)
            {
            }
            column(shiptin; shiptin)
            {
            }
            column(Shippan; Shippan)
            {
            }
            column(shipcst; shipcst)
            {
            }
            column(shipEccNo; shipEccNo)
            {
            }
            column(DebitEntryNo; '')//"Sales Invoice Header"."PLA Entry No." + ' ' + "Sales Invoice Header"."RG 23 A Entry No." + ' ' + "Sales Invoice Header"."RG 23 C Entry No.")
            {
            }
            column(EntryDate; EntryDate)
            {
            }
            column(ComInfo_MarinCargoSalesPolicy; Rec_Comp.Name)
            {
            }
            column(TotNoPack2; TotNoPack2)
            {
            }
            column(vShipGST; vShipGST)
            {
            }
            column(vCustGST; vCustGST)
            {
            }
            column(ShipStateCode; ShipStateCode)
            {
            }
            column(Ship_State; Ship_State)
            {
            }
            column(AmounttoCustomer_SalesInvoiceHeader; 0)// "Sales Invoice Header"."Amount to Customer")
            {
            }
            column(IsSupplimentory; IsSupplimentory)
            {
            }
            column(IsDeliveryChallan; IsDeliveryChallan)
            {
            }
            dataitem(CopyLoop; Integer)
            {
                DataItemTableView = SORTING(Number);
                dataitem(Pageloop; Integer)
                {
                    DataItemTableView = SORTING(Number)
                                        WHERE(Number = CONST(1));
                    column(RType; RType)
                    {
                    }
                    column(Number; Number)
                    {
                    }
                    column(OutputNo; OutputNo)
                    {
                    }
                    dataitem("Sales Invoice Line"; "Sales Invoice Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Sales Invoice Header";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            WHERE(Quantity = FILTER(> 0),
                                                  Type = FILTER(Item | 'Charge (Item)' | "Fixed Asset" | "G/L Account"),
                                                  "No." = FILTER(<> 43720010));
                        column(SrNo; SrNo)
                        {
                        }
                        column(HSNSACCode_SalesInvoiceLine; "Sales Invoice Line"."HSN/SAC Code")
                        {
                        }
                        column(DocumentNo_SalesInvoiceLine; "Sales Invoice Line"."Document No.")
                        {
                        }
                        column(SIL_Description; "Sales Invoice Line".Description)
                        {
                        }
                        column(SIL_TypeOfPacking; "Sales Invoice Line".Type)
                        {
                        }
                        column(SIL_UnitPerParcle; "Sales Invoice Line"."Units per Parcel")
                        {
                        }
                        column(SIL_Quantity; "Sales Invoice Line".Quantity)
                        {
                        }
                        column(SIL_UnitPrice; "Sales Invoice Line"."Unit Price")
                        {
                        }
                        column(SIL_LineAmount; "Sales Invoice Line"."Line Amount")
                        {
                        }
                        column(SIL_VATAmount; 0)// "Sales Invoice Line"."ADC VAT Amount")
                        {
                        }
                        column(SIL_TaxAmount; 0)// ROUND("Sales Invoice Line"."Tax Amount"))
                        {
                        }
                        column(SIL_lineNo; "Sales Invoice Line"."Line No.")
                        {
                        }
                        column(Bed; Bed)
                        {
                        }
                        column(ECess; ECess)
                        {
                        }
                        column(SheCess; SheCess)
                        {
                        }
                        column(DstAmtTotNoPacge; DstAmt * TotNoPacge)
                        {
                        }
                        column(FrtAmt; FrtAmt)
                        {
                        }
                        column(SIL_AmountToCust; 0)//"Sales Invoice Line"."Amount To Customer")
                        {
                        }
                        column(NoText; AmountINWords + ' ' + NoText[2])
                        {
                        }
                        column(NoTextExcise; NoTextExcise[1])
                        {
                        }
                        column(BEDPER; BEDPER)
                        {
                        }
                        column(EcessPER; 0)// ExciseEntry."eCess %")
                        {
                        }
                        column(SheCessPER; 0)// ExciseEntry."SHE Cess %")
                        {
                        }
                        column(TotNoPacge; TotNoPacge)
                        {
                        }
                        column(Total_SL; Total_SL)
                        {
                        }
                        column(taxdec; taxdec)
                        {
                        }
                        column(TotalPcg1; TotalPcg1)
                        {
                        }
                        column(TaxAmtTotal; TaxAmtTotal)
                        {
                        }
                        column(TotalAmt; TotalAmt)
                        {
                        }
                        column(ExciseProdPostingGroup_SalesInvoiceLine; 0)// "Sales Invoice Line"."Excise Prod. Posting Group")
                        {
                        }
                        column(TotNoOfPkgs; TotNoOfPkgs)
                        {
                        }
                        column(SupplimentoryDetail; SuppliDetail + ' ' + SupliMentInvNo + ' ' + DatedOn)
                        {
                        }
                        column(SourceNo; SourceNo)
                        {
                        }
                        column(ADCVAT; ADCVAT)
                        {
                        }
                        column(GenBusPostingGroup_SalesInvoiceLine; "Sales Invoice Line"."Gen. Bus. Posting Group")
                        {
                        }
                        column(Note; 'Note' + ':' + ' ' + Note)
                        {
                        }
                        column(Discount_SalesInvoiceLine; "Sales Invoice Line"."Line Discount Amount")
                        {
                        }
                        column(LineAmount_SalesInvoiceLine; "Sales Invoice Line"."Line Amount")
                        {
                        }
                        column(TaxAmount_SalesInvoiceLine; 0)// "Sales Invoice Line"."Tax Amount")
                        {
                        }
                        column(vCGST; vCGST)
                        {
                        }
                        column(Freight_SalesInvoiceLine; "Sales Invoice Line".Type)
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
                        dataitem("Value Entry"; "Value Entry")
                        {
                            DataItemLink = "Document No." = FIELD("Document No."),
                                           "Document Line No." = FIELD("Line No.");
                            DataItemLinkReference = "Sales Invoice Line";
                            DataItemTableView = WHERE("Item Ledger Entry Type" = FILTER(Sale),
                                                      "Document Type" = FILTER("Sales Invoice"),
                                                      "Source Code" = FILTER('SALES'));
                            dataitem("Item Ledger Entry"; "Item Ledger Entry")
                            {
                                DataItemLink = "Entry No." = FIELD("Item Ledger Entry No.");
                                column(ILE_Lot_No; "Item Ledger Entry"."Lot No.")
                                {
                                }
                                column(ILE_InvoiceQuantity; "Item Ledger Entry"."Invoiced Quantity")
                                {
                                }
                                column(V_ExpDate; V_ExpDate)
                                {
                                }

                                trigger OnAfterGetRecord()
                                begin

                                    IF ("Sales Invoice Line"."Tax Group Code" = 'ST') OR ("Sales Invoice Line"."Tax Group Code" = 'STW') THEN
                                        V_ExpDate := FORMAT("Item Ledger Entry"."Expiration Date", 0, 4)
                                    ELSE
                                        V_ExpDate := '';
                                end;

                                trigger OnPreDataItem()
                                begin
                                    vILECount += "Item Ledger Entry".COUNT;
                                end;
                            }
                        }

                        trigger OnAfterGetRecord()
                        begin
                            V_Count := "Sales Invoice Line".COUNT;
                            SrNo := SrNo + 1;
                            //IF SrNo > 10 THEN BEGIN
                            //END;
                            //SNo+=1;
                            PostedShipmentDate := 0D;
                            IF Quantity <> 0 THEN
                                PostedShipmentDate := FindPostedShipmentDate;

                            IF (Type = Type::"G/L Account") AND (NOT ShowInternalInfo) THEN
                                "No." := '';

                            VATAmountLine.INIT;
                            VATAmountLine."VAT Identifier" := "VAT Identifier";
                            VATAmountLine."VAT Calculation Type" := "VAT Calculation Type";
                            VATAmountLine."Tax Group Code" := "Tax Group Code";
                            VATAmountLine."VAT %" := "VAT %";
                            VATAmountLine."VAT Base" := Amount;
                            VATAmountLine."Amount Including VAT" := "Amount Including VAT";
                            VATAmountLine."Line Amount" := "Line Amount";
                            IF "Allow Invoice Disc." THEN
                                VATAmountLine."Inv. Disc. Base Amount" := "Line Amount";
                            VATAmountLine."Invoice Discount Amount" := "Inv. Discount Amount";
                            VATAmountLine.InsertLine;

                            TotalTCSAmount += 0;// "Total TDS/TCS Incl. SHE CESS";

                            IF ISSERVICETIER THEN BEGIN
                                TotalSubTotal += "Line Amount";
                                TotalInvoiceDiscountAmount -= "Inv. Discount Amount";
                                TotalAmount += Amount;
                                TotalAmountVAT += "Amount Including VAT" - Amount;
                                // TotalAmountInclVAT += "Amount Including VAT";
                                TotalPaymentDiscountOnVAT += -("Line Amount" - "Inv. Discount Amount" - "Amount Including VAT");
                                TotalAmountInclVAT += 0;// "Amount To Customer";
                                TotalExciseAmt += 0;//"Excise Amount";
                                TotalTaxAmt += 0;//"Tax Amount";
                                ServiceTaxAmount += 0;//"Service Tax Amount";
                                ServiceTaxeCessAmount += 0;//"Service Tax eCess Amount";
                                ServiceTaxSHECessAmount += 0;// "Service Tax SHE Cess Amount";
                            END;

                            /* StructureLineDetails.RESET;
                            StructureLineDetails.SETRANGE(Type, StructureLineDetails.Type::Sale);
                            StructureLineDetails.SETRANGE("Document Type", StructureLineDetails."Document Type"::Invoice);
                            StructureLineDetails.SETRANGE("Invoice No.", "Document No.");
                            StructureLineDetails.SETRANGE("Item No.", "No.");
                            StructureLineDetails.SETRANGE("Line No.", "Line No.");
                            IF StructureLineDetails.FIND('-') THEN
                                REPEAT
                                    IF NOT StructureLineDetails."Payable to Third Party" THEN BEGIN
                                        IF StructureLineDetails."Tax/Charge Type" = StructureLineDetails."Tax/Charge Type"::Charges THEN
                                            ChargesAmount := ChargesAmount + ABS(StructureLineDetails.Amount);
                                        IF StructureLineDetails."Tax/Charge Type" = StructureLineDetails."Tax/Charge Type"::"Other Taxes" THEN
                                            OtherTaxesAmount := OtherTaxesAmount + ABS(StructureLineDetails.Amount);
                                    END;
                                UNTIL StructureLineDetails.NEXT = 0; */
                            /* IF ISSERVICETIER THEN BEGIN
                                IF "Sales Invoice Header"."Transaction No. Serv. Tax" <> 0 THEN BEGIN
                                    ServiceTaxEntry.RESET;
                                    ServiceTaxEntry.SETRANGE(Type, ServiceTaxEntry.Type::Sale);
                                    ServiceTaxEntry.SETRANGE("Document Type", ServiceTaxEntry."Document Type"::Invoice);
                                    ServiceTaxEntry.SETRANGE("Document No.", "Document No.");
                                    IF ServiceTaxEntry.FINDFIRST THEN BEGIN

                                         IF "Sales Invoice Header"."Currency Code" <> '' THEN BEGIN
                                            ServiceTaxEntry."Service Tax Amount" :=
                                              ROUND(CurrExchRate.ExchangeAmtLCYToFCY(
                                                "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                                ServiceTaxEntry."Service Tax Amount", "Sales Invoice Header"."Currency Factor"));
                                            ServiceTaxEntry."eCess Amount" :=
                                              ROUND(CurrExchRate.ExchangeAmtLCYToFCY(
                                                "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                                ServiceTaxEntry."eCess Amount", "Sales Invoice Header"."Currency Factor"));
                                            ServiceTaxEntry."SHE Cess Amount" :=
                                              ROUND(CurrExchRate.ExchangeAmtLCYToFCY(
                                                "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                                ServiceTaxEntry."SHE Cess Amount", "Sales Invoice Header"."Currency Factor"));
                                        END; 


                                        ServiceTaxAmt := ABS(ServiceTaxEntry."Service Tax Amount");
                                        ServiceTaxECessAmt := ABS(ServiceTaxEntry."eCess Amount");
                                        ServiceTaxSHECessAmt := ABS(ServiceTaxEntry."SHE Cess Amount");
                                        AppliedServiceTaxAmt := ServiceTaxAmount - ABS(ServiceTaxEntry."Service Tax Amount");
                                        AppliedServiceTaxECessAmt := ServiceTaxeCessAmount - ABS(ServiceTaxEntry."eCess Amount");
                                        AppliedServiceTaxSHECessAmt := ServiceTaxSHECessAmount - ABS(ServiceTaxEntry."SHE Cess Amount");
                                    END ELSE BEGIN
                                        AppliedServiceTaxAmt := ServiceTaxAmount;
                                        AppliedServiceTaxECessAmt := ServiceTaxeCessAmount;
                                        AppliedServiceTaxSHECessAmt := ServiceTaxSHECessAmount;
                                    END;
                                END ELSE BEGIN
                                    ServiceTaxAmt := ServiceTaxAmount;
                                    ServiceTaxECessAmt := ServiceTaxeCessAmount;
                                    ServiceTaxSHECessAmt := ServiceTaxSHECessAmount;
                                END;
                            END;*/


                            // to print chaperheading no as Tarrif No
                            Item.RESET;
                            /* ExciseProdPostingGroup.RESET;
                            Item.SETFILTER(Item."No.", "Sales Invoice Line"."No.");
                            IF Item.FIND('-') THEN
                                ExciseProdPostingGroup.SETFILTER(ExciseProdPostingGroup.Code, Item."Excise Prod. Posting Group");
                            IF ExciseProdPostingGroup.FINDFIRST THEN BEGIN
                                REPEAT
                                    TarrifNo := ExciseProdPostingGroup."Chapter No.";
                                UNTIL ExciseProdPostingGroup.NEXT = 0;
                            END ELSE BEGIN
                                TarrifNo := '';
                            END; */

                            // to print chaperheading no as Tarrif No


                            //print percentage of excise cess and hecess
                            /*  IF "Sales Invoice Line".Type = "Sales Invoice Line".Type::Item THEN BEGIN
                                 ExciseEntry.RESET;
                                 ExciseEntry.SETRANGE(ExciseEntry."Document Type", ExciseEntry."Document Type"::Invoice);
                                 ExciseEntry.SETRANGE(ExciseEntry."Document No.", "Sales Invoice Line"."Document No.");
                                 ExciseEntry.SETRANGE(ExciseEntry."Item No.", "Sales Invoice Line"."No.");
                                 ExciseEntry.SETRANGE(ExciseEntry."Excise Bus. Posting Group", "Sales Invoice Line"."Excise Bus. Posting Group");
                                 ExciseEntry.SETRANGE(ExciseEntry."Excise Prod. Posting Group", "Sales Invoice Line"."Excise Prod. Posting Group");
                                 IF ExciseEntry.FINDFIRST THEN;
                             END; */

                            //to calculate total no of packages as Qty/Unit per parcel
                            //PB
                            //TotalPcg1:=0;

                            TotNoOfPkgs := 0;
                            IF "Sales Invoice Line"."Units per Parcel" > 0 THEN BEGIN
                                TotNoOfPkgs := "Sales Invoice Line".Quantity / "Sales Invoice Line"."Units per Parcel";
                            END;
                            TotalPcg1 += TotNoOfPkgs;


                            //to calculate total no of packages as Qty/Unit per parcel

                            /*
                            TotNoPacge:=0;
                            RecsalesLines.RESET;
                            RecsalesLines.SETRANGE(RecsalesLines."Document No.","Sales Invoice Line"."Document No.");
                            RecsalesLines.SETRANGE(RecsalesLines.Type,"Sales Invoice Line".Type::Item);
                            IF RecsalesLines.FINDFIRST THEN REPEAT
                              TotNoPacge += RecsalesLines.Quantity;
                            UNTIL RecsalesLines.NEXT = 0;
                            
                            */


                            //PB
                            //LOT NO
                            VET.RESET;
                            itl.RESET;
                            VET.SETFILTER(VET."Document No.", "Sales Invoice Line"."Document No.");
                            IF VET.FINDFIRST THEN BEGIN
                                LotNO := itl."Lot No.";
                            END ELSE BEGIN
                                LotNO := '';
                            END;

                            //LOT NO
                            //AMK
                            /*
                            IF StrHead.GET("Sales Invoice Header".Structure) THEN
                            BEGIN
                               StrDetails.RESET;
                               StrDetails.SETRANGE(StrDetails.Code,StrHead.Code);
                               IF StrDetails.FINDFIRST THEN
                               REPEAT
                                  IF StrDetails."Tax/Charge Group" = 'DISCOUNT' THEN
                                  BEGIN
                                    recstrordline.RESET;
                                    recstrordline.SETRANGE("Invoice No.","Sales Invoice Header"."No.");
                                    recstrordline.SETRANGE("Structure Code",StrDetails.Code);
                                    recstrordline.SETRANGE("Tax/Charge Group",StrDetails."Tax/Charge Group");
                                    recstrordline.SETRANGE("Line No.","Sales Invoice Line"."Line No.");
                                    IF recstrordline.FINDFIRST THEN
                                       //Freight := ABS(Str OrdDetails."Calculation Value");
                                       //Discount := "Sales Line"."Charges To Customer" - Freight;
                                         Discount:=Discount + recstrordline.Amount;
                                   END  ELSE
                                   IF StrDetails."Tax/Charge Group" = 'FREIGHT'  THEN
                                   BEGIN
                                    recstrord.RESET;
                                    recstrord.SETRANGE("No.","Sales Invoice Header"."No.");
                                    recstrord.SETRANGE("Structure Code",StrDetails.Code);
                                    recstrord.SETRANGE("Tax/Charge Group",StrDetails."Tax/Charge Group");
                                    IF recstrord.FINDFIRST THEN
                                        //Freight := ABS(recstrord."Calculation Value");
                                   END;
                                   END;
                                UNTIL StrDetails.NEXT=0;
                            END;
                            //AMK
                             */



                            //to calculate total no of packages as Qty/Unit per parcel

                            TotNoOfPkgs := 0;
                            IF "Sales Invoice Line"."Units per Parcel" > 0 THEN BEGIN
                                TotNoOfPkgs := "Sales Invoice Line".Quantity / "Sales Invoice Line"."Units per Parcel";
                            END;
                            //to calculate total no of packages as Qty/Unit per parcel

                            Bed := 0;
                            ECess := 0;
                            SheCess := 0;
                            /* ExciseEntry.RESET;
                            ExciseEntry.SETRANGE(ExciseEntry."Document No.", "Sales Invoice Line"."Document No."); //AMK Start
                            IF ExciseEntry.FINDFIRST THEN
                                REPEAT
                                    Bed := ROUND(Bed + ABS(ExciseEntry."BED Amount"), 1);
                                    ECess := ROUND(ECess + ABS(ExciseEntry."eCess Amount"), 1);
                                    SheCess := ROUND(SheCess + ABS(ExciseEntry."SHE Cess Amount"), 1);
                                UNTIL ExciseEntry.NEXT = 0; */
                            ExciseAmount := ADCVAT + Bed;//Rspl Sachin



                            RepCheck1.InitTextVariable;
                            RepCheck1.FormatNoText(NoTextExcise, ROUND(ExciseAmount, 1), "Sales Invoice Header"."Currency Code");
                            DutyText_2 := DELCHR(NoTextExcise[1], '=', '*');


                            //TotNoPacge:
                            TotNoPacge := 0;
                            RecsalesLines.RESET;
                            RecsalesLines.SETRANGE(RecsalesLines."Document No.", "Sales Invoice Line"."Document No.");
                            RecsalesLines.SETRANGE(RecsalesLines.Type, "Sales Invoice Line".Type::Item);
                            IF RecsalesLines.FINDFIRST THEN
                                REPEAT
                                    TotNoPacge += RecsalesLines.Quantity;
                                UNTIL RecsalesLines.NEXT = 0;
                            //BED % :
                            /* ExciseEntry.RESET;
                            ExciseEntry.SETRANGE(ExciseEntry."Document No.", "Sales Invoice Line"."Document No.");
                            IF ExciseEntry.FINDFIRST THEN
                                BEDPER := ExciseEntry."BED %"; */


                            /*  IF "Sales Invoice Line".Supplementary = TRUE THEN
                                 SuppliDetail := 'Supplementory Invoice of Tax Invoice No.'
                             ELSE
                                 SuppliDetail := '';

                             IF "Sales Invoice Line".Supplementary = TRUE THEN
                                 DatedOn := 'Dated On'
                             ELSE
                                 DatedOn := '';

                             IF "Sales Invoice Line".Supplementary = TRUE THEN
                                 SupliMentInvNo := "Sales Invoice Line"."Source Document No."
                             ELSE
                                 SupliMentInvNo := ''; */


                            SourceNoInvDet.RESET;
                            SourceNoInvDet.SETPERMISSIONFILTER;//AMK
                                                               // SourceNoInvDet.SETRANGE(SourceNoInvDet."No.", "Sales Invoice Line"."Source Document No.");
                            IF SourceNoInvDet.FINDFIRST THEN
                                SourceNo := SourceNoInvDet."Posting Date";

                            //Total_SL+="Sales Invoice Line"."Line Amount";
                            //TaxAmtTotal +="Sales Invoice Line"."Tax Amount";
                            //TotalAmt+="Sales Invoice Line"."Amount To Customer";

                            CLEAR(TaxAmtTotal);
                            CLEAR(TotalAmt);
                            CLEAR(Total_SL);
                            Rec_Sil.RESET;
                            Rec_Sil.SETRANGE(Rec_Sil."Document No.", "Sales Invoice Line"."Document No.");
                            IF Rec_Sil.FINDFIRST THEN BEGIN
                                REPEAT
                                    Total_SL += Rec_Sil."Line Amount";
                                    TaxAmtTotal += 0;// Rec_Sil."Tax Amount";
                                    TotalAmt += 0;// Rec_Sil."Amount To Customer";
                                UNTIL Rec_Sil.NEXT = 0;
                            END;

                            //  "Sales Invoice Header".CALCFIELDS("Amount to Customer");
                            RepCheck.InitTextVariable;
                            RepCheck.FormatNoText(NoText, /*"Sales Invoice Header"."Amount to Customer"*/0, "Sales Invoice Header"."Currency Code");
                            AmountINWords := DELCHR(NoText[1], '=', '*');
                            //MESSAGE('%1',NoText[2]);
                            //MESSAGE('%1,%2',TaxAmtTotal,TotalAmt);
                            //Rspl Sachin
                            IF ("Location Code" = '10001200') OR ("Location Code" = '10001400') OR ("Location Code" = '10001600') OR ("Location Code" = '10003700') THEN BEGIN
                                IF "Sales Invoice Line".Type = "Sales Invoice Line".Type::Item THEN BEGIN
                                    RecItem.RESET;
                                    RecItem.SETRANGE(RecItem."No.", "No.");
                                    IF RecItem.FINDFIRST THEN BEGIN
                                        IF RecItem."Item Category Code" = 'RM' THEN BEGIN
                                            Note := 'Inputs Cleared as such under Rule 3 (5) of Cenvet Credit Rule, 2004'
                                        END ELSE BEGIN
                                            Note := '';
                                        END;
                                    END;
                                END;
                            END;

                            //>>GST
                            CLEAR(vCGST);
                            CLEAR(vSGST);
                            CLEAR(vIGST);
                            CLEAR(vCGSTRate);
                            CLEAR(vSGSTRate);
                            CLEAR(vIGSTRate);
                            DetailGSTEntry.RESET;
                            DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                            DetailGSTEntry.SETRANGE("Document No.", "Document No.");
                            DetailGSTEntry.SETRANGE("No.", "No.");
                            IF DetailGSTEntry.FINDSET THEN
                                REPEAT
                                    IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                        vCGST := ABS(DetailGSTEntry."GST Amount");
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
                                UNTIL DetailGSTEntry.NEXT = 0;

                            //<<

                        end;

                        trigger OnPostDataItem()
                        begin
                            /*
                            IF V_Count <= 1 THEN
                              V_BlankLine := 8-V_Count
                             ELSE IF V_Count > 5 THEN
                              V_BlankLine := 10-V_Count;
                            */

                        end;

                        trigger OnPreDataItem()
                        begin
                            SETPERMISSIONFILTER;
                            CLEAR(TotalPcg1);
                            SrNo := 0;
                        end;
                    }
                    dataitem(Integer; Integer)
                    {
                        DataItemTableView = SORTING(Number);
                        column(V_BlankLine; Integer.Number)
                        {
                        }

                        trigger OnPreDataItem()
                        begin
                            V_BlankLine := V_Count + vILECount;

                            Integer.SETRANGE(Integer.Number, 1, 12 - V_BlankLine); //19Jun2017
                            //Integer.SETRANGE(Integer.Number,1,10-V_BlankLine);
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin
                        //Narration
                        IF ("Sales Invoice Header"."Location Code" = '10001200') OR ("Sales Invoice Header"."Location Code" = '10001400') THEN BEGIN
                            Narration := Text1;
                            VatAct := Text2;
                        END
                        ELSE
                            IF ("Sales Invoice Header"."Location Code" = '10001500') OR ("Sales Invoice Header"."Location Code" = '10001600') THEN BEGIN
                                VatAct := Text3;
                                Narration := '';
                            END;
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    //IF Number >= 1 THEN BEGIN
                    vILECount := 0;
                    RType := '';
                    /*IF noofloops = 4 THEN BEGIN
                      IF Number = 1 THEN
                       RType := 'DUPLICATE (FOR TRANSPORTER)';
                      IF Number = 2 THEN
                       RType := 'TRIPLICATE (FOR Central Excise)';
                      IF Number = 3 THEN
                       RType := 'QUADRUPLICATE FOR(H.O COPY)';
                      IF Number = 4 THEN
                       RType := 'EXTRA COPY';
                     END ELSE BEGIN
                     */
                    IF Number = 1 THEN
                        RType := 'ORIGINAL For RECIPIENT';
                    IF Number = 2 THEN
                        RType := 'DUPLICATE FOR TRANSPORTER';
                    IF Number = 3 THEN
                        RType := 'TRIPLICATE FOR Supplier';
                    IF Number = 4 THEN
                        RType := 'QUADRUPLICATE FOR H.O';
                    IF Number = 5 THEN
                        RType := 'EXTRA COPY';

                    //END;
                    OutputNo += 1;
                    CurrReport.PAGENO := 1;

                end;

                trigger OnPreDataItem()
                begin
                    noofloops := noofcopies;
                    IF noofloops <= 0 THEN
                        noofloops := 1;
                    SETRANGE(Number, 1, noofloops);
                    //IF noofcopies =4 THEN
                    //  OutputNo := 2
                    //ELSE
                    OutputNo := 1;
                    cnt := 1;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                Rec_Comp.GET;
                Rec_Comp.CALCFIELDS(Rec_Comp.Picture);

                IF Location.GET("Sales Invoice Header"."Location Code") THEN;
                States.RESET;
                States.SETRANGE(States.Code, Location."State Code");
                IF States.FINDFIRST THEN
                    Loc_State := States.Description;
                county_regin.RESET;
                county_regin.SETRANGE(county_regin.Code, Location."Country/Region Code");
                IF county_regin.FINDFIRST THEN
                    Loc_Country := county_regin.Name;




                //lcdetails start
                /*  lcdetails.RESET;
                 lcdetails.SETRANGE(lcdetails."No.", "Sales Invoice Header"."LC No.");
                 IF lcdetails.FINDFIRST THEN BEGIN
                     lcdate := lcdetails."Date of Issue";
                     //lcissuebk:=lcdetails."Issuing Bank";
                     lcno := lcdetails."LC No.";
                 END; */
                //lcdetails end;



                //to print road Permit
                /* TransitDocumentOrderDetails.SETFILTER(TransitDocumentOrderDetails."Vendor / Customer Ref.",
                  "Sales Invoice Header"."Sell-to Customer No.");
                TransitDocumentOrderDetails.SETFILTER(TransitDocumentOrderDetails."PO / SO No.", "Sales Invoice Header"."Order No.");
                IF TransitDocumentOrderDetails.FIND('-') THEN
                    roadpermitno := TransitDocumentOrderDetails."Form No."; */


                //get transporter name
                shippingagent.RESET;
                IF "Shipping Agent Code" <> '' THEN
                    IF shippingagent.GET("Shipping Agent Code") THEN
                        transporternm := shippingagent.Name;


                //get payment term description
                PaymentTerms.RESET;
                IF "Payment Terms Code" <> '' THEN
                    IF PaymentTerms.GET("Payment Terms Code") THEN
                        paydesc := PaymentTerms.Description;


                //cust detail information

                IF Cust.GET("Sales Invoice Header"."Sell-to Customer No.") THEN BEGIN
                    VTIN := '';// Cust."T.I.N. No.";
                    VCST := '';// Cust."C.S.T. No.";
                    //VCST:=Cust."L.S.T. No.";
                    // VLST :=Cust."L.S.T. No.";//"T.I.N. No.";
                    VLST := '';// Cust."T.I.N. No.";
                    VECC := '';// Cust."E.C.C. No.";
                    VPAN := Cust."P.A.N. No.";
                    VRANGE := '';// Cust.Range;
                    VDIV := '';//Cust.Collectorate;
                    VCOMM := '';//Cust.Collectorate;
                    vCustGST := Cust."GST Registration No.";
                    //CustStateCode  := ;
                END;
                //cust detail information end;


                //Discount Amount
                /* DiscAmts.RESET;
                DiscAmts.SETRANGE(DiscAmts."No.", "Sales Invoice Header"."No.");
                IF DiscAmts.FINDSET THEN
                    REPEAT
                        IF DiscAmts."Tax/Charge Group" = 'DISCOUNT' THEN BEGIN
                            //DstAmt := ChargeAmount + SalesInvLine."Charges To Customer";
                            DstAmt := DiscAmts."Calculation Value";
                        END;
                    UNTIL DiscAmts.NEXT = 0; */

                //FreightAmt :
                /*  FreightAmt.RESET;
                 FreightAmt.SETRANGE(FreightAmt."No.", "Sales Invoice Header"."No.");
                 IF FreightAmt.FINDSET THEN
                     REPEAT
                         IF FreightAmt."Tax/Charge Group" = 'FREIGHT' THEN BEGIN
                             FrtAmt := FreightAmt."Calculation Value";
                         END;
                     UNTIL FreightAmt.NEXT = 0; */


                //RSPL-RAHUL-START
                //**Code added to get customer and Consignee state and county
                customer.RESET;
                customer.SETRANGE(customer."No.", "Sales Invoice Header"."Bill-to Customer No.");
                IF customer.FINDFIRST THEN BEGIN
                    county_regin.RESET;
                    county_regin.SETRANGE(county_regin.Code, customer."Country/Region Code");
                    IF county_regin.FINDFIRST THEN
                        Cust_Country := county_regin.Name;
                    States.RESET;
                    States.SETRANGE(States.Code, customer."State Code");
                    IF States.FINDFIRST THEN
                        Cust_State := States.Description;
                    CustStateCode := States."State Code (GST Reg. No.)"
                END;
                ShiptoAddress.RESET;
                ShiptoAddress.SETRANGE(ShiptoAddress.Code, "Sales Invoice Header"."Ship-to Code");
                ShiptoAddress.SETRANGE(ShiptoAddress.Name, "Sales Invoice Header"."Ship-to Name");
                IF ShiptoAddress.FINDFIRST THEN BEGIN
                    county_regin.RESET;
                    county_regin.SETRANGE(county_regin.Code, ShiptoAddress."Country/Region Code");
                    IF county_regin.FINDFIRST THEN
                        Ship_Country := county_regin.Name;
                    States.RESET;
                    States.SETRANGE(States.Code, ShiptoAddress.State);
                    IF States.FINDFIRST THEN
                        Ship_State := States.Description;
                    ShipStateCode := States."State Code (GST Reg. No.)";
                    ;
                END;
                ///RSPL-RAHUL-END

                //AMK
                /* IF ("Sales Invoice Header"."PLA Entry No." <> '') OR ("Sales Invoice Header"."RG 23 A Entry No." <> '') OR
                   ("Sales Invoice Header"."RG 23 C Entry No." <> '') THEN
                    EntryDate := "Sales Invoice Header"."Posting Date"
                ELSE */
                EntryDate := 0D;
                TotNoPack2 := 0;
                CLEAR(ADCVAT);
                SIL.RESET;
                SIL.SETRANGE("Document No.", "No.");
                //SIL.SETRANGE("Document Type","Document Type");
                IF SIL.FINDSET THEN
                    REPEAT
                        IF "Sales Invoice Line"."Units per Parcel" > 0 THEN
                            TotNoPack2 += "Sales Invoice Line".Quantity / "Sales Invoice Line"."Units per Parcel";
                        ADCVAT += 0;// SIL."ADC VAT Amount";
                                    //IsSupplimentory :=   SIL.Supplementary;
                    UNTIL SIL.NEXT = 0;


                //>>19Jun2017 InvoiceType
                CLEAR(TaxInvoice15);
                IF "Invoice Type" = "Invoice Type"::Taxable THEN
                    TaxInvoice15 := 'Tax Invoice';
                IF "Invoice Type" = "Invoice Type"::"Bill of Supply" THEN
                    TaxInvoice15 := 'Bill of Supply';
                IF "Invoice Type" = "Invoice Type"::Export THEN
                    TaxInvoice15 := 'Export Invoice';
                IF "Invoice Type" = "Invoice Type"::Supplementary THEN
                    TaxInvoice15 := 'Supplementary Invoice';
                IF "Invoice Type" = "Invoice Type"::"Debit Note" THEN
                    TaxInvoice15 := 'Debit Note';
                IF "Invoice Type" = "Invoice Type"::"Non-GST" THEN
                    TaxInvoice15 := 'Non-GST Invoice';

                //<<19Jun2017 InvoiceType
            end;

            trigger OnPreDataItem()
            begin
                "Sales Invoice Header".SETPERMISSIONFILTER;
                IF noofcopies = 0 THEN
                    noofcopies := 4;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(noofcopies; noofcopies)
                {
                    ApplicationArea = ALL;
                    Caption = 'No. of Copies';
                }
                field(ShowInternalInfo; ShowInternalInfo)
                {
                    ApplicationArea = ALL;
                    Caption = 'Show Internal Information';
                }
                field(LogInteraction; LogInteraction)
                {
                    ApplicationArea = ALL;
                    Caption = 'Log Interaction';
                    Enabled = LogInteractionEnable;
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
        Rec_Comp: Record 79;
        Location: Record 14;
        //lcdetails: Record 16300;
        lcdate: Date;
        lcissuebk: Text[20];
        lcno: Text;
        // TransitDocumentOrderDetails: Record 13768;
        roadpermitno: Code[10];
        shippingagent: Record 291;
        transporternm: Text;
        PaymentTerms: Record 3;
        paydesc: Text;
        Cust: Record 18;
        VLST: Text[30];
        VTIN: Text[50];
        VCST: Text[50];
        VECC: Text[50];
        VRANGE: Text[50];
        VDIV: Text[50];
        VCOMM: Text[50];
        VPAN: Text[50];
        ship: Record 222;
        shiptin: Code[30];
        shipcst: Code[30];
        shipEccNo: Code[20];
        Shiprange: Text[30];
        Shipdiv: Text[30];
        Shippan: Text[30];
        Shipcomm: Text[30];
        TotNoOfPkgs: Decimal;
        //ExciseEntry: Record 13712;
        Bed: Decimal;
        ECess: Decimal;
        SheCess: Decimal;
        ExciseAmount: Decimal;
        RepCheck: Report "Check Report";
        NoTextExcise: array[2] of Text[80];
        NoText: array[2] of Text[80];
        RepCheck1: Report "Check Report";
        //DiscAmts: Record 13760;
        DstAmt: Decimal;
        RecsalesLines: Record 113;
        TotNoPacge: Decimal;
        //FreightAmt: Record 13760;
        FrtAmt: Decimal;
        BEDPER: Decimal;
        SupliMentInvNo: Text[30];
        SuppliDetail: Text[50];
        DatedOn: Text[30];
        RType: Code[150];
        noofloops: Integer;
        noofcopies: Integer;
        OutputNo: Integer;
        cnt: Integer;
        Narration: Text[250];
        VatAct: Text[30];
        Text1: Label 'SALES TAX EXEMPTED UNDER NOTIFICATION NO. 5/4/87 FIN(R&C) 2 DATED 23/04/1987 PUBLISHED IN OFFICIAL GAZZATTE IN GOA , DAMAN & DIU SERIES II ON 4 PAGE 34,35 OF DATED 23/04/1987 ';
        Text2: Label 'Daman & Diu VAT Act, 2005';
        Text3: Label 'U.K. VAT Act, 2005';
        ShowInternalInfo: Boolean;
        LogInteraction: Boolean;
        [InDataSet]
        LogInteractionEnable: Boolean;
        "------Mail Functionality------": Integer;
        recsalesinvoiceheader: Record 112;
        //SMTPMailSetup: Record 409;
        //SMTPMail: Codeunit 400;
        Mail: Codeunit 397;
        Selection: Integer;
        Options: Text[30];
        fl: File;
        WTEXT: Text[30];
        window: Dialog;
        emailid: Code[50];
        recCust: Record 18;
        emailid2: Code[50];
        test: Code[100];
        isSendMail: Boolean;
        vUser: Text[100];
        "------RSPl--RAHUL----": Integer;
        States: Record State;
        county_regin: Record 9;
        Loc_State: Text;
        Loc_Country: Code[30];
        ShiptoAddress: Record 222;
        Ship_Country: Text;
        Cust_Country: Text;
        Ship_State: Text;
        customer: Record 18;
        Cust_State: Text;
        AmountINWords: Text;
        DutyText_2: Text;
        taxdec: Text;
        //  SegManagement: Codeunit SegManagement;
        NextEntryNo: Integer;
        SalesShipmentBuffer: Record 7190 temporary;
        FirstValueEntryNo: Integer;
        TaxGroups: Record 321;
        Text004: Label 'Sales - Invoice %1';
        Text005: Label 'Page %1';
        Text006: Label 'Total %1 Excl. VAT';
        Text007: Label 'VAT Amount Specification in ';
        Text008: Label 'Local Currency';
        Text009: Label 'Exchange rate: %1/%2';
        Text010: Label 'Sales - Prepayment Invoice %1';
        Text13700: Label 'Total %1 Incl. Taxes';
        Text13701: Label 'Total %1 Excl. Taxes';
        Text16500: Label 'Supplementary Invoice';
        Text12: Label '"I/We hereby certify that my/our Registration certificate under the Maharashtra Value Added Tax Act , 2002 is in force on the date on which the sale of the goods specified in this tax invoice is made me/us and that the transaction of sale covered by this tax invoice has been affected by me/us and it shall be accounted for in the turnover of sales while filing of return and the due tax. If any payable on the sale has been paid or shall be paid."';
        Text13: Label 'Certified that the particulars given above are true & correct & the amount indicated represents the price actually charged and that there is no flow of additional consideration directly or indirectly from the buyer.';
        Text16502: Label 'Against';
        Text16503: Label 'Form';
        Text4: Label 'I/We hereby certify that my/ our Registration certificate under the ';
        Text5: Label ', is in force on the date on which the sales of goods specified in this tax invoice is made by me/us and the transaction of sales covered by this tax invoice has been effected by me/us and it shall be accounted for in the turnover of sales while filing return and the due tax, if any, payable on the sale has been paid or shall be paid.';
        Text6: Label '1) Interest @ 24% p.a. will be charged on overdue bills. Remittance should be made payble in mumbai by A/c Payee cheques/Draft/RTGS. We reserve the right to demand the payment of this bill at anytime, subject to ';
        SrNo: Integer;
        PostedShipmentDate: Date;
        VATAmountLine: Record 290 temporary;
        TotalTCSAmount: Decimal;
        TotalSubTotal: Decimal;
        TotalInvoiceDiscountAmount: Decimal;
        TotalAmount: Decimal;
        TotalAmountVAT: Decimal;
        TotalPaymentDiscountOnVAT: Decimal;
        TotalAmountInclVAT: Decimal;
        TotalExciseAmt: Decimal;
        TotalTaxAmt: Decimal;
        ServiceTaxAmount: Decimal;
        ServiceTaxeCessAmount: Decimal;
        ServiceTaxSHECessAmount: Decimal;
        // StructureLineDetails: Record 13798;
        ChargesAmount: Decimal;
        OtherTaxesAmount: Decimal;
        //ServiceTaxEntry: Record 16473;
        CurrExchRate: Record 330;
        ServiceTaxAmt: Decimal;
        ServiceTaxECessAmt: Decimal;
        ServiceTaxSHECessAmt: Decimal;
        AppliedServiceTaxAmt: Decimal;
        AppliedServiceTaxECessAmt: Decimal;
        AppliedServiceTaxSHECessAmt: Decimal;
        Item: Record 27;
        //ExciseProdPostingGroup: Record 13710;
        TarrifNo: Code[11];
        VET: Record 5802;
        itl: Record 32;
        LotNO: Text[50];
        Excisedec: Text[250];
        Tmtext: array[2] of Text[50];
        //StrHead: Record 13792;
        //StrDetails: Record 13793;
        //StrOrdDetails: Record 13794;
        //recstrordline: Record 13798;
        //recstrord: Record 13760;
        FinalTotal: Decimal;
        Discount: Decimal;
        "------": Integer;
        Total_SL: Decimal;
        TotalPcg1: Decimal;
        TaxAmtTotal: Decimal;
        TotalAmt: Decimal;
        V_BlankLine: Integer;
        V_Count: Integer;
        Rec_Sil: Record 113;
        EntryDate: Date;
        V_ExpDate: Code[30];
        RecBank: Record 287;
        RecSIH: Record 112;
        SIL: Record 113;
        TotNoPack2: Decimal;
        SourceNo: Date;
        SourceNoInvDet: Record 112;
        "--": Integer;
        Cust1: Record 18;
        ADCVAT: Decimal;
        "-----------------Rspl Sachin------------": Integer;
        RecItem: Record 27;
        Note: Text[500];
        vCustGST: Code[20];
        vShipGST: Code[20];
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        CustStateCode: Code[20];
        ShipStateCode: Code[20];
        vILECount: Integer;
        IsSupplimentory: Boolean;
        IsDeliveryChallan: Boolean;
        "------19Jun2017----": Integer;
        TaxInvoice15: Text[50];

    // //[Scope('Internal')]
    procedure InitLogInteraction()
    begin
        //  LogInteraction := SegManagement.FindInteractTmplCode(4) <> '';
    end;

    // //[Scope('Internal')]
    procedure FindPostedShipmentDate(): Date
    var
        SalesShipmentHeader: Record 110;
        SalesShipmentBuffer2: Record 7190 temporary;
    begin
        NextEntryNo := 1;
        IF "Sales Invoice Line"."Shipment No." <> '' THEN
            IF SalesShipmentHeader.GET("Sales Invoice Line"."Shipment No.") THEN
                EXIT(SalesShipmentHeader."Posting Date");

        IF "Sales Invoice Header"."Order No." = '' THEN
            EXIT("Sales Invoice Header"."Posting Date");

        CASE "Sales Invoice Line".Type OF
            "Sales Invoice Line".Type::Item:
                GenerateBufferFromValueEntry("Sales Invoice Line");
            "Sales Invoice Line".Type::"G/L Account", "Sales Invoice Line".Type::Resource,
          "Sales Invoice Line".Type::"Charge (Item)", "Sales Invoice Line".Type::"Fixed Asset":
                GenerateBufferFromShipment("Sales Invoice Line");
            "Sales Invoice Line".Type::" ":
                EXIT(0D);
        END;

        SalesShipmentBuffer.RESET;
        SalesShipmentBuffer.SETRANGE("Document No.", "Sales Invoice Line"."Document No.");
        SalesShipmentBuffer.SETRANGE("Line No.", "Sales Invoice Line"."Line No.");
        IF SalesShipmentBuffer.FIND('-') THEN BEGIN
            SalesShipmentBuffer2 := SalesShipmentBuffer;
            IF SalesShipmentBuffer.NEXT = 0 THEN BEGIN
                SalesShipmentBuffer.GET(
                  SalesShipmentBuffer2."Document No.", SalesShipmentBuffer2."Line No.", SalesShipmentBuffer2."Entry No.");
                SalesShipmentBuffer.DELETE;
                EXIT(SalesShipmentBuffer2."Posting Date");
            END;
            SalesShipmentBuffer.CALCSUMS(Quantity);
            IF SalesShipmentBuffer.Quantity <> "Sales Invoice Line".Quantity THEN BEGIN
                SalesShipmentBuffer.DELETEALL;
                EXIT("Sales Invoice Header"."Posting Date");
            END;
        END ELSE
            EXIT("Sales Invoice Header"."Posting Date");
    end;

    // //[Scope('Internal')]
    procedure GenerateBufferFromValueEntry(SalesInvoiceLine2: Record 113)
    var
        ValueEntry: Record 5802;
        ItemLedgerEntry: Record 32;
        TotalQuantity: Decimal;
        Quantity: Decimal;
    begin
        TotalQuantity := SalesInvoiceLine2."Quantity (Base)";
        ValueEntry.SETCURRENTKEY("Document No.");
        ValueEntry.SETRANGE("Document No.", SalesInvoiceLine2."Document No.");
        ValueEntry.SETRANGE("Posting Date", "Sales Invoice Header"."Posting Date");
        ValueEntry.SETRANGE("Item Charge No.", '');
        ValueEntry.SETFILTER("Entry No.", '%1..', FirstValueEntryNo);
        IF ValueEntry.FIND('-') THEN
            REPEAT
                IF ItemLedgerEntry.GET(ValueEntry."Item Ledger Entry No.") THEN BEGIN
                    IF SalesInvoiceLine2."Qty. per Unit of Measure" <> 0 THEN
                        Quantity := ValueEntry."Invoiced Quantity" / SalesInvoiceLine2."Qty. per Unit of Measure"
                    ELSE
                        Quantity := ValueEntry."Invoiced Quantity";
                    AddBufferEntry(
                      SalesInvoiceLine2,
                      -Quantity,
                      ItemLedgerEntry."Posting Date");
                    TotalQuantity := TotalQuantity + ValueEntry."Invoiced Quantity";
                END;
                FirstValueEntryNo := ValueEntry."Entry No." + 1;
            UNTIL (ValueEntry.NEXT = 0) OR (TotalQuantity = 0);
        //TAX GROUPS PRINT ON REPORT START
        TaxGroups.RESET;
        TaxGroups.SETRANGE(TaxGroups.Code, "Sales Invoice Line"."Tax Group Code");
        IF TaxGroups.FINDFIRST THEN
            taxdec := TaxGroups.Description;
        //TAX GROUPS Print on report end;

        //lc date, lc issuing bank, lc no;

        //lcdetails.RESET;
    end;

    //  //[Scope('Internal')]
    procedure GenerateBufferFromShipment(SalesInvoiceLine: Record 113)
    var
        SalesInvoiceHeader: Record 112;
        SalesInvoiceLine2: Record 113;
        SalesShipmentHeader: Record 110;
        SalesShipmentLine: Record 111;
        TotalQuantity: Decimal;
        Quantity: Decimal;
    begin
        TotalQuantity := 0;
        SalesInvoiceHeader.SETCURRENTKEY("Order No.");
        SalesInvoiceHeader.SETFILTER("No.", '..%1', "Sales Invoice Header"."No.");
        SalesInvoiceHeader.SETRANGE("Order No.", "Sales Invoice Header"."Order No.");
        SalesInvoiceHeader.SETPERMISSIONFILTER;
        IF SalesInvoiceHeader.FIND('-') THEN
            REPEAT
                SalesInvoiceLine2.SETRANGE("Document No.", SalesInvoiceHeader."No.");
                SalesInvoiceLine2.SETRANGE("Line No.", SalesInvoiceLine."Line No.");
                SalesInvoiceLine2.SETRANGE(Type, SalesInvoiceLine.Type);
                SalesInvoiceLine2.SETRANGE("No.", SalesInvoiceLine."No.");
                SalesInvoiceLine2.SETRANGE("Unit of Measure Code", SalesInvoiceLine."Unit of Measure Code");
                IF SalesInvoiceLine2.FIND('-') THEN
                    REPEAT
                        TotalQuantity := TotalQuantity + SalesInvoiceLine2.Quantity;
                    UNTIL SalesInvoiceLine2.NEXT = 0;
            UNTIL SalesInvoiceHeader.NEXT = 0;

        SalesShipmentLine.SETCURRENTKEY("Order No.", "Order Line No.");
        SalesShipmentLine.SETRANGE("Order No.", "Sales Invoice Header"."Order No.");
        SalesShipmentLine.SETRANGE("Order Line No.", SalesInvoiceLine."Line No.");
        SalesShipmentLine.SETRANGE("Line No.", SalesInvoiceLine."Line No.");
        SalesShipmentLine.SETRANGE(Type, SalesInvoiceLine.Type);
        SalesShipmentLine.SETRANGE("No.", SalesInvoiceLine."No.");
        SalesShipmentLine.SETRANGE("Unit of Measure Code", SalesInvoiceLine."Unit of Measure Code");
        SalesShipmentLine.SETFILTER(Quantity, '<>%1', 0);

        IF SalesShipmentLine.FIND('-') THEN
            REPEAT
                IF "Sales Invoice Header"."Get Shipment Used" THEN
                    CorrectShipment(SalesShipmentLine);
                IF ABS(SalesShipmentLine.Quantity) <= ABS(TotalQuantity - SalesInvoiceLine.Quantity) THEN
                    TotalQuantity := TotalQuantity - SalesShipmentLine.Quantity
                ELSE BEGIN
                    IF ABS(SalesShipmentLine.Quantity) > ABS(TotalQuantity) THEN
                        SalesShipmentLine.Quantity := TotalQuantity;
                    Quantity :=
                      SalesShipmentLine.Quantity - (TotalQuantity - SalesInvoiceLine.Quantity);

                    TotalQuantity := TotalQuantity - SalesShipmentLine.Quantity;
                    SalesInvoiceLine.Quantity := SalesInvoiceLine.Quantity - Quantity;

                    IF SalesShipmentHeader.GET(SalesShipmentLine."Document No.") THEN
                        SalesShipmentHeader.SETPERMISSIONFILTER;
                    BEGIN
                        AddBufferEntry(
                          SalesInvoiceLine,
                          Quantity,
                          SalesShipmentHeader."Posting Date");
                    END;
                END;
            UNTIL (SalesShipmentLine.NEXT = 0) OR (TotalQuantity = 0);
    end;

    // //[Scope('Internal')]
    procedure CorrectShipment(var SalesShipmentLine: Record 111)
    var
        SalesInvoiceLine: Record 113;
    begin
        SalesInvoiceLine.SETCURRENTKEY("Shipment No.", "Shipment Line No.");
        SalesInvoiceLine.SETRANGE("Shipment No.", SalesShipmentLine."Document No.");
        SalesInvoiceLine.SETRANGE("Shipment Line No.", SalesShipmentLine."Line No.");
        IF SalesInvoiceLine.FIND('-') THEN
            REPEAT
                SalesShipmentLine.Quantity := SalesShipmentLine.Quantity - SalesInvoiceLine.Quantity;
            UNTIL SalesInvoiceLine.NEXT = 0;
    end;

    //  //[Scope('Internal')]
    procedure AddBufferEntry(SalesInvoiceLine: Record "113"; QtyOnShipment: Decimal; PostingDate: Date)
    begin
        SalesShipmentBuffer.SETRANGE("Document No.", SalesInvoiceLine."Document No.");
        SalesShipmentBuffer.SETRANGE("Line No.", SalesInvoiceLine."Line No.");
        SalesShipmentBuffer.SETRANGE("Posting Date", PostingDate);
        IF SalesShipmentBuffer.FIND('-') THEN BEGIN
            SalesShipmentBuffer.Quantity := SalesShipmentBuffer.Quantity + QtyOnShipment;
            SalesShipmentBuffer.MODIFY;
            EXIT;
        END;

        WITH SalesShipmentBuffer DO BEGIN
            "Document No." := SalesInvoiceLine."Document No.";
            "Line No." := SalesInvoiceLine."Line No.";
            "Entry No." := NextEntryNo;
            Type := SalesInvoiceLine.Type;
            "No." := SalesInvoiceLine."No.";
            Quantity := QtyOnShipment;
            "Posting Date" := PostingDate;
            INSERT;
            NextEntryNo := NextEntryNo + 1
        END;
    end;

    local procedure DocumentCaption(): Text[250]
    begin
        IF "Sales Invoice Header"."Prepayment Invoice" THEN
            EXIT(Text010);
        EXIT(Text004);
    end;
}


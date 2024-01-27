report 50046 "Supplimentary Plant Inv--IL"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SupplimentaryPlantInvIL.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            column(Location; Loc.Address + ' ' + Loc."Address 2" + ' ' + Loc.City + ' ' + Loc."Post Code" + ' ' + RecState1.Description + ' Phone: ' + Loc."Phone No." + ' ' + ' Email: ' + Loc."E-Mail")
            {
            }
            column(CompInfo; 'Website :' + RecCompInfo."Home Page" + '   C.I.N No. :' + 'RecCompInfo."Company Registration  No."' + '  PAN No. :' + RecCompInfo."P.A.N. No.")
            {
            }
            column(InvNo; InvNo)
            {
            }
            column(ECCDesc; ECCDesc)
            {
            }
            column(OrderDate_SalesInvoiceHeader; "Sales Invoice Header"."Order Date")
            {
            }
            column(RangeDivisionCommissionarate; '"E.C.C. Nos."."C.E. Range"' + ' - ' + '"E.C.C. Nos."."C.E. Division"' + ' - ' + '"E.C.C. Nos."."C.E. Commissionerate"')
            {
            }
            column(CEAddr1; '"E.C.C. Nos."."C.E Address1"')
            {
            }
            column(CEAddr2; '"E.C.C. Nos."."C.E Address2"')
            {
            }
            column(ECCPostCode; '"E.C.C. Nos."."C.E City"' + ' - ' + '"E.C.C. Nos."."C.E Post Code"')
            {
            }
            column(ECCLOC_CST; 'Loc."T.I.N. No."' + ' & ' + 'Loc."C.S.T No."')
            {
            }
            column(No_SalesInvoiceHeader; "Sales Invoice Header"."No.")
            {
            }
            column(PostingDate_SalesInvoiceHeader; "Sales Invoice Header"."Posting Date")
            {
            }
            column(AppliedInvoiceNo; AppliedInvoiceNo)
            {
            }
            column(SalesInvHdr_PostingDate; SalesInvHdr."Posting Date")
            {
            }
            column(RoadPermitNo_SalesInvoiceHeader; "Sales Invoice Header"."Road Permit No.")
            {
            }
            column(ExternalDocumentNo_SalesInvoiceHeader; "Sales Invoice Header"."External Document No.")
            {
            }
            column(VehOwnersName; VehOwnersName)
            {
            }
            column(VehicleNo_SalesInvoiceHeader; "Sales Invoice Header"."Vehicle No.")
            {
            }
            column(LRRRNo_SalesInvoiceHeader; "Sales Invoice Header"."LR/RR No.")
            {
            }
            column(LRRRDate_SalesInvoiceHeader; "Sales Invoice Header"."LR/RR Date")
            {
            }
            column(ResponsibilityCenter_SalesInvoiceHeader; "Sales Invoice Header"."Responsibility Center")
            {
            }
            column(Desc_PaymentTerm; "Payment Terms".Description)
            {
            }
            column(FreightType_SalesInvoiceHeader; "Sales Invoice Header"."Freight Type")
            {
            }
            column(LBTRegNo; RecCust."LBT Reg. No.")
            {
            }
            column(BilltoName_SalesInvoiceHeader; "Sales Invoice Header"."Bill-to Name")
            {
            }
            column(BilltoAddress_SalesInvoiceHeader; "Sales Invoice Header"."Bill-to Address")
            {
            }
            column(BilltoAddress2_SalesInvoiceHeader; "Sales Invoice Header"."Bill-to Address 2")
            {
            }
            column(BillToCity; "Bill-to City" + ' - ' + "Bill-to Post Code")
            {
            }
            column(BillToSate; RecState.Description + ', ' + RecCountry.Name)
            {
            }
            column(SuppliersCode; SuppliersCode)
            {
            }
            column(BillToCST; 'RecCust."C.S.T. No."')
            {
            }
            column(BillToECC; 'RecCust."E.C.C. No."')
            {
            }
            column(BillToTIN; 'RecCust."T.I.N. No."')
            {
            }
            column(ShiptoName_SalesInvoiceHeader; "Sales Invoice Header"."Ship-to Name")
            {
            }
            column(ShiptoAddress_SalesInvoiceHeader; "Sales Invoice Header"."Ship-to Address")
            {
            }
            column(ShiptoAddress2_SalesInvoiceHeader; "Sales Invoice Header"."Ship-to Address 2")
            {
            }
            column(ShipToCity; "Ship-to City" + ' - ' + "Ship-to Post Code")
            {
            }
            column(ShipToState; RecState1.Description + ', ' + RecCountry1.Name)
            {
            }
            column(ShiptoLST; ShiptoLST)
            {
            }
            column(ShipToCST; ShipToCST)
            {
            }
            column(ECCNo; ECCNo)
            {
            }
            column(ShiptoTIN; ShiptoTIN)
            {
            }
            column(InvoicePrintTime_SalesInvoiceHeader; "Sales Invoice Header"."Invoice Print Time")
            {
            }
            column(IsuranceNo; CompanyInfo1."Insurance No." + ' of the ' + CompanyInfo1."Insurance Provider")
            {
            }
            column(Validafor; 'valid up to ' + FORMAT(CompanyInfo1."Policy Expiration Date"))
            {
            }
            column(AmounttoCustomer_SalesInvoiceHeader; 0)//"Sales Invoice Header"."Amount to Customer")
            {
            }
            column(BEDAmount; BEDAmount)
            {
            }
            column(EcessAmt; EcessAmt)
            {
            }
            column(SheCess; SheCess)
            {
            }
            column(AddlDuty; AddlDuty)
            {
            }
            column(State1; State1)
            {
            }
            dataitem(CopyLoop; Integer)
            {
                DataItemTableView = SORTING(Number);
                dataitem(Integer; Integer)
                {
                    DataItemTableView = SORTING(Number)
                                        WHERE(Number = FILTER(1));
                    column(InvType; InvType)
                    {
                    }
                    column(InvType1; InvType1)
                    {
                    }
                    column(number; Number)
                    {
                    }
                    column(outPutNo; OutPutNo)
                    {
                    }
                    dataitem("Sales Invoice Line"; "Sales Invoice Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Sales Invoice Header";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Quantity = FILTER(<> 0));
                        // "Excise Prod. Posting Group" = FILTER(<> ''));
                        column(item_Desc; 'IPOL ' + RecItem.Description)
                        {
                        }
                        column(DocumentNo_SalesInvoiceLine; "Sales Invoice Line"."Document No.")
                        {
                        }
                        column(LineNo_SalesInvoiceLine; "Sales Invoice Line"."Line No.")
                        {
                        }
                        column(CrossRefNo; 'SalesInvLine1."Cross-Reference No."')
                        {
                        }
                        column(ExciseProdPostingGroup_SalesInvoiceLine; '"Sales Invoice Line"."Excise Prod. Posting Group"')
                        {
                        }
                        column(SIl_lotNo; recSIL."Lot No.")
                        {
                        }
                        column(PackDescription; PackDescription)
                        {
                        }
                        column(QuantityBase_SalesInvoiceLine; "Sales Invoice Line"."Quantity (Base)")
                        {
                        }
                        column(QtyPerUnit; "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure")
                        {
                        }
                        column(BEDPer; 0)// recExcisePostingSetup."BED %")
                        {
                        }
                        column(ExciseAmt; 0/*"Sales Invoice Line"."Excise Amount"*/ / "Sales Invoice Line"."Quantity (Base)")
                        {
                        }
                        column(BEDAmount_SalesInvoiceLine; 0/*"Sales Invoice Line"."BED Amount"*/)
                        {
                        }
                        column(SheCessEcess; /*"Sales Invoice Line"."SHE Cess Amount"*/0 + 0/* "Sales Invoice Line"."eCess Amount"*/)
                        {
                        }
                        column(DescountAmt; "Sales Invoice Line"."Line Discount Amount" + "Sales Invoice Line"."Inv. Discount Amount")
                        {
                        }
                        column(TaxAmount_SalesInvoiceLine; 0 /*"Sales Invoice Line"."Tax Amount"*/)
                        {
                        }
                        column(ExciseAmount_SalesInvoiceLine; ROUND(0/*"Sales Invoice Line"."Excise Amount"*/, 0.01))
                        {
                        }
                        column(LineAmount_SalesInvoiceLine; "Sales Invoice Line"."Line Amount")
                        {
                        }
                        column(Freight; Freight)
                        {
                        }
                        column(cessvalue; cessvalue)
                        {
                        }
                        column(ADDTAX; ADDTAX)
                        {
                        }
                        column(RoundOffAmnt; RoundOffAmnt)
                        {
                        }
                        column(TotAmt; ROUND("Sales Invoice Line"."Line Amount" - "Sales Invoice Line"."Inv. Discount Amount" + 0/*"Sales Invoice Line"."BED Amount"*/ + 0/*"Sales Invoice Line"."eCess Amount"*/ + 0/*"Sales Invoice Line"."SHE Cess Amount" */+ 0/*"Sales Invoice Line"."Tax Amount"*/ + Freight + RoundOffAmnt, 1))
                        {
                        }
                        column(Rupees; Rupees)
                        {
                        }
                        column(vCount; vCount)
                        {
                        }
                        column(DescriptionLineDuty_1; DescriptionLineDuty[1])
                        {
                        }
                        column(DescriptionLineeCess_1; DescriptionLineeCess[1])
                        {
                        }
                        column(DescriptionLineSHeCess_1; DescriptionLineSHeCess[1])
                        {
                        }
                        column(DescriptionLineTot_1; DescriptionLineTot[1])
                        {
                        }
                        column(BEDAmount_SIL1; 0)// "Sales Invoice Line"."BED Amount")
                        {
                        }
                        column(EcessAmount_SIL1; 0)//"Sales Invoice Line"."eCess Amount")
                        {
                        }
                        column(ShecCessAmt_SIL1; 0)// "Sales Invoice Line"."SHE Cess Amount")
                        {
                        }
                        column(TaxDescription; TaxDescription)
                        {
                        }
                        column(Cesstext; Cesstext)
                        {
                        }
                        column(ADDTAXText; ADDTAXText)
                        {
                        }
                        dataitem(SIL; "Sales Invoice Line")
                        {
                            DataItemLink = "Document No." = FIELD("No.");
                            DataItemLinkReference = "Sales Invoice Header";
                            DataItemTableView = SORTING("Document No.", "Line No.")
                                                ORDER(Ascending)
                                                WHERE(Quantity = FILTER(<> 0));
                            //"Excise Prod. Posting Group" = FILTER(<> ''));
                            column(Ctr; Ctr)
                            {
                            }
                            column(DocumentNo_SIL; SIL."Document No.")
                            {
                            }
                            column(LineNo_SIL; SIL."Line No.")
                            {
                            }
                            column(Item_Desc_SIL2; 'IPOL ' + RecItem.Description)
                            {
                            }
                            column(InvoiceNo_SIL2; InvoiceNo)
                            {
                            }
                            column(LotNo_SIL2; recSIL."Lot No.")
                            {
                            }
                            column(QuantityBase_SIL; SIL."Quantity (Base)")
                            {
                            }
                            column(MRPPrice_SIL; 0)//SIL."MRP Price")
                            {
                            }
                            column(RateTobeChange; 0)//"Sales Invoice Line"."MRP Price" + SalesInvLine."MRP Price")
                            {
                            }
                            column(UnitPrice; 0)// "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure")
                            {
                            }
                            column(MRPPrice2; 0)// SalesInvLine."MRP Price")
                            {
                            }

                            trigger OnAfterGetRecord()
                            begin

                                SalesInvLine.RESET;
                                SalesInvLine.SETRANGE(SalesInvLine."Document No.", SIL."Appiles to Inv.No.");
                                SalesInvLine.SETRANGE(SalesInvLine."No.", SIL."Final Item No.");
                                IF SalesInvLine.FINDFIRST THEN
                                    Ctr += 1;
                                CLEAR(ItemQty);
                                ItemREC.RESET;
                                IF ItemREC.GET(SIL."No.") THEN;
                            end;

                            trigger OnPreDataItem()
                            begin
                                Ctr := 0;
                            end;
                        }

                        trigger OnAfterGetRecord()
                        begin

                            Ctr := Ctr + 1;

                            /*  recExcisePostingSetup.RESET;
                             recExcisePostingSetup.SETRANGE(recExcisePostingSetup."Excise Bus. Posting Group",
                                                           "Sales Invoice Line"."Excise Bus. Posting Group");
                             recExcisePostingSetup.SETRANGE(recExcisePostingSetup."Excise Prod. Posting Group",
                                                           "Sales Invoice Line"."Excise Prod. Posting Group");
                             IF recExcisePostingSetup.FINDFIRST THEN */
                            recSIL.RESET;
                            recSIL.SETRANGE(recSIL."Document No.", "Sales Invoice Line"."Appiles to Inv.No.");
                            recSIL.SETRANGE(recSIL."No.", "Sales Invoice Line"."Final Item No.");
                            recSIL.SETRANGE(recSIL."Line No.", "Sales Invoice Line"."Final Line No.");
                            iF recSIL.FINDFIRST THEN
                                RecItem.RESET;
                            iF RecItem.GET("Sales Invoice Line"."Final Item No.") THEN;


                            CLEAR(cessvalue);
                            /*   PostedStrOrdDetails.RESET;
                              PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Invoice Line"."Document No.");
                              PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                              PostedStrOrdDetails1.SETFILTER(PostedStrOrdDetails1."Tax/Charge Code", '%1|%2', 'CESS', 'C1');
                              IF PostedStrOrdDetails.FINDFIRST THEN
                                  REPEAT
                                      cessvalue := cessvalue + PostedStrOrdDetails.Amount;
                                  UNTIL PostedStrOrdDetails.NEXT = 0;

                              IF cessvalue <> 0 THEN BEGIN
                                  IF PostedStrOrdDetails."Calculation Value" <> 0 THEN
                                      Cesstext := 'Cess @' + FORMAT(PostedStrOrdDetails."Calculation Value") + ' ' + '%' */
                            // ELSE
                            Cesstext := 'Cess';
                            //  END ELSE
                            Cesstext := 'Cess';

                            CLEAR(ADDTAX);
                            /*   PostedStrOrdDetails1.RESET;
                              PostedStrOrdDetails1.SETRANGE(PostedStrOrdDetails1."Invoice No.", "Sales Invoice Line"."Document No.");
                              PostedStrOrdDetails1.SETRANGE(PostedStrOrdDetails1."Tax/Charge Type", PostedStrOrdDetails1."Tax/Charge Type"::"Other Taxes");
                              PostedStrOrdDetails1.SETRANGE(PostedStrOrdDetails1."Tax/Charge Code", 'ADDL TAX');
                              PostedStrOrdDetails1.SETRANGE(PostedStrOrdDetails1."Calculation Type", PostedStrOrdDetails1."Calculation Type"::Percentage);
                              IF PostedStrOrdDetails1.FINDFIRST THEN
                                  ADDTAXText := 'Addl Tax @' + FORMAT(PostedStrOrdDetails1."Calculation Value") + ' ' + '%';
                              REPEAT
                                  ADDTAX := ADDTAX + PostedStrOrdDetails1.Amount;
                              UNTIL PostedStrOrdDetails1.NEXT = 0;
   */

                            Qty := Quantity;
                            QtyPerUOM := "Qty. per Unit of Measure";

                            iF "Sales Invoice Line"."Unit of Measure Code" = 'BRL' THEN BEGIN
                                PackDescription := FORMAT(Qty) + '*' +
                                FORMAT("Sales Invoice Line"."Qty. per Unit of Measure") + '' + "Sales Invoice Line"."Unit of Measure Code"
                                + '-' + RecItem."Base Unit of Measure";
                            END ELSE BEGIN
                                PackDescription := FORMAT(Qty) + '*' + "Sales Invoice Line"."Unit of Measure Code" + '-' + RecItem."Base Unit of Measure";
                            END;


                            SalesInvLine.RESET;
                            SalesInvLine.SETRANGE(SalesInvLine."Document No.", SalesInvHdr."No.");
                            IF SalesInvLine.FINDFIRST THEN;

                            SalesInvLine1.RESET;
                            SalesInvLine1.SETRANGE(SalesInvLine1."Document No.", "Sales Invoice Line"."Appiles to Inv.No.");
                            SalesInvLine1.SETFILTER(SalesInvLine1.Quantity, '<>%1', 0);
                            SalesInvLine1.SETRANGE(SalesInvLine1."No.", "Sales Invoice Line"."Final Item No.");
                            SalesInvLine1.SETRANGE(SalesInvLine1."Line No.", "Sales Invoice Line"."Final Line No.");
                            IF SalesInvLine1.FINDFIRST THEN;

                            IF ItemREC.GET("Sales Invoice Line"."No.") THEN BEGIN
                                ItemQty := "Sales Invoice Line"."Quantity (Base)";
                                TotalItemQty += ItemQty;
                            END;


                            TaxDescr := 'Tax Descr';

                            /*  DetailedTaxEntry.RESET;
                             DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Sales Invoice Header"."No.");
                             IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                                 IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN
                                     FormCode := 'with Form ' + DetailedTaxEntry."Form Code";
                                 END;
                                 TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                             END ELSE BEGIN
                                 TaxDescription := 'NoTax Ag.Form' + "Sales Invoice Header"."Form Code";
                             END;
                             IF (RecCust."Customer Posting Group" = 'FOREIGN') AND ("Sales Invoice Header"."Form Code" = '') THEN BEGIN
                                 TaxDescription := 'NoTax Ag.Export';
                             END; */
                            vCount := "Sales Invoice Line".COUNT;


                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineDuty, ROUND(BEDAmount, 0.01), '');
                            vCheck.FormatNoText(DescriptionLineeCess, ROUND(EcessAmt, 0.01), '');
                            // "Sales Invoice Header".CALCFIELDS("Amount to Customer");
                            vCheck.FormatNoText(DescriptionLineTot, ROUND((/*"Sales Invoice Header"."Amount to Customer"*/0 + RoundOffAmnt), 0.01), '');
                            vCheck.FormatNoText(DescriptionLineSHeCess, ROUND(SheCess, 0.01), '');

                            DescriptionLineDuty[1] := DELCHR(DescriptionLineDuty[1], '=', '*');
                            DescriptionLineeCess[1] := DELCHR(DescriptionLineeCess[1], '=', '*');
                            DescriptionLineTot[1] := DELCHR(DescriptionLineTot[1], '=', '*');
                            DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1], '=', '*');


                            Loc2.GET("Location Code");

                            RecState.SETRANGE(RecState.Code, Loc."State Code");
                            IF RecState.FINDFIRST THEN
                                State1 := RecState.Description;
                        end;

                        trigger OnPreDataItem()
                        begin
                            //Ctr:=1;
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin
                        CLEAR(TotalItemQty);
                        i += 1;
                        //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-----START
                        IF ("Sales Invoice Header"."Print Invoice" = FALSE) THEN BEGIN
                            IF i = 1 THEN
                                InvType := 'ORIGINAL FOR BUYER'
                            ELSE
                                IF i = 2 THEN
                                    InvType := 'DUPLICATE FOR TRANSPORTER'
                                ELSE
                                    IF i = 3 THEN
                                        InvType := 'TRIPLICATE FOR ASSESSEE'
                                    ELSE
                                        IF i = 4 THEN BEGIN
                                            InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                                            InvType1 := 'NOT FOR CENVAT';
                                        END ELSE
                                            IF i = 5 THEN BEGIN
                                                InvType := 'EXTRA COPY NOT FOR CENVAT';
                                                InvType1 := '';
                                            END
                                            ELSE
                                                InvType := '';
                        END;

                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'sa') THEN BEGIN
                            IF i = 1 THEN
                                InvType := 'ORIGINAL FOR BUYER'
                            ELSE
                                IF i = 2 THEN
                                    InvType := 'DUPLICATE FOR TRANSPORTER'
                                ELSE
                                    IF i = 3 THEN
                                        InvType := 'TRIPLICATE FOR ASSESSEE'
                                    ELSE
                                        IF i = 4 THEN BEGIN
                                            InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                                            InvType1 := 'NOT FOR CENVAT';
                                        END ELSE
                                            IF i = 5 THEN BEGIN
                                                InvType := 'EXTRA COPY NOT FOR CENVAT';
                                                InvType1 := '';
                                            END
                                            ELSE
                                                InvType := '';
                        END;

                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'sa') THEN BEGIN
                            IF i = 1 THEN
                                InvType := 'EXTRA COPY';
                            InvType1 := ' NOT FOR COMMERCIAL USE';
                        END;
                        //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-------END

                        InvNoLen := STRLEN("Sales Invoice Header"."No.");
                        InvNo := COPYSTR("Sales Invoice Header"."No.", (InvNoLen - 3), 4);
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    OutPutNo += 1;
                    IF Number > 1 THEN
                        copytext := Text003;
                    CurrReport.PAGENO := 1;

                end;

                trigger OnPreDataItem()
                begin

                    //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-----START
                    IF NOT (USERID = 'sa') THEN BEGIN
                        IF "Sales Invoice Header"."Print Invoice" = TRUE THEN BEGIN
                            SETRANGE(Number, 1, 1);
                        END
                        ELSE BEGIN
                            NoOfCopies := 4;
                            NoOfLoops := ABS(NoOfCopies) + 1;
                            copytext := '';
                            SETRANGE(Number, 1, NoOfLoops);
                        END
                    END
                    ELSE BEGIN
                        NoOfCopies := 4;
                        NoOfLoops := ABS(NoOfCopies) + 1;
                        copytext := '';
                        SETRANGE(Number, 1, NoOfLoops);
                    END
                    //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-------END
                end;
            }

            trigger OnAfterGetRecord()
            begin

                SalesInvLine.RESET;
                SalesInvLine.SETRANGE(SalesInvLine."Document No.", "Sales Invoice Header"."No.");
                SalesInvLine.SETRANGE(SalesInvLine.Type, SalesInvLine.Type::Item);
                IF SalesInvLine.FINDFIRST THEN
                    REPEAT
                        Freight := Freight + 0;// SalesInvLine."Charges To Customer";
                    UNTIL SalesInvLine.NEXT = 0;

                "Payment Terms".GET("Sales Invoice Header"."Payment Terms Code");
                "Shipment Method".GET("Sales Invoice Header"."Shipment Method Code");
                IF "Shipping Agent".GET("Sales Invoice Header"."Shipping Agent Code") THEN;

                CustPostSetup.RESET;
                CustPostSetup.FINDFIRST;
                RoundOffAcc := CustPostSetup."Invoice Rounding Account";

                SalesInvLine.RESET;
                SalesInvLine.SETRANGE(SalesInvLine."Document No.", "Sales Invoice Header"."No.");
                IF SalesInvLine.FINDFIRST THEN
                    AppliedInvoiceNo := SalesInvLine."Appiles to Inv.No.";

                InvoiceNoLen := STRLEN(SalesInvLine."Appiles to Inv.No.");
                //InvoiceNo := COPYSTR(SalesInvLine."Appiles to Inv.No.",(InvoiceNoLen-3),4);

                SalesInvHdr.RESET;
                SalesInvHdr.SETRANGE(SalesInvHdr."No.", SalesInvLine."Appiles to Inv.No.");
                IF SalesInvHdr.FINDFIRST THEN
                    IF "Sales Invoice Header"."Currency Code" = '' THEN BEGIN
                        Rupees := 'INR';
                    END
                    ELSE BEGIN
                        Rupees := "Sales Invoice Header"."Currency Code";
                    END;


                IF RecCompInfo.GET THEN;


                RespCentre.GET("Responsibility Center");
                Loc.GET("Location Code");
                IF RecState1.GET(Loc."State Code") THEN;

                IF RecState.GET(State) THEN;

                InvNoLen := STRLEN("Sales Invoice Header"."No.");
                InvNo := COPYSTR("Sales Invoice Header"."No.", (InvNoLen - 3), 4);


                Loc.GET("Sales Invoice Header"."Location Code");

                SalesPerson.RESET;
                SalesPerson.SETRANGE(SalesPerson.Code, "Salesperson Code");
                IF SalesPerson.FINDFIRST THEN;

                IF "Bill-to Country/Region Code" <> ''
                  THEN
                    RecCountry.GET("Bill-to Country/Region Code")
                ELSE
                    RecCountry.Name := '';

                IF RecState.GET(State) THEN;

                RecCust.GET("Sell-to Customer No.");
                SuppliersCode := RecCust."Our Account No.";

                IF "Sales Invoice Header"."Ship-to Code" = '' THEN
                    Ship2Name := "Sales Invoice Header"."Full Name"
                ELSE
                    Ship2Name := "Sales Invoice Header"."Ship-to Name";


                IF RecState1.GET(State) THEN;


                IF NOT RecCountry1.GET("Ship-to Country/Region Code")
                  THEN
                    RecCountry1.GET("Bill-to Country/Region Code");

                // "E.C.C. Nos.".SETRANGE("E.C.C. Nos.".Code, Loc."E.C.C. No.");
                //  "E.C.C. Nos.".FINDFIRST;
                ECCDesc := '';//"E.C.C. Nos.".Description; 

                IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code")
                  THEN BEGIN
                    ECCNo := ShipToCode."E.C.C. No.";
                    ShiptoName := ShipToCode.Name;
                    ShipToCST := ShipToCode."C.S.T. No.";
                    ShiptoLST := ShipToCode."L.S.T. No.";
                    ShiptoTIN := ShipToCode."T.I.N. No.";
                END
                ELSE BEGIN
                    ECCNo := '';// RecCust."E.C.C. No.";
                    ShiptoName := CustName;
                    ShipToCST := '';// RecCust."C.S.T. No.";
                    ShiptoLST := '';// RecCust."L.S.T. No.";
                    ShiptoTIN := '';//RecCust."T.I.N. No.";
                END;
                IF "Shipping Agent".GET("Sales Invoice Header"."Shipping Agent Code") THEN;
                VehOwnersName := "Shipping Agent".Name;
                CurrDatetime := TIME;

                IF "Sales Invoice Header"."Ship-to Code" <> '' THEN BEGIN
                    ShipToCode.SETRANGE(ShipToCode.Code, "Ship-to Code");
                    IF ShipToCode.FINDFIRST THEN;
                END
                ELSE BEGIN
                    ShipToCode.SETRANGE(ShipToCode."Customer No.", "Sell-to Customer No.");
                    IF ShipToCode.FINDFIRST THEN;
                END;

                SalesInvHdr.RESET;
                SalesInvHdr.SETRANGE(SalesInvHdr."No.", "Applies-to Doc. No.");
                IF SalesInvHdr.FINDFIRST THEN;
                IF SalesInvHdr."No." <> '' THEN BEGIN
                    InvNoLen := STRLEN(SalesInvHdr."No.");
                    InvNo := COPYSTR(SalesInvHdr."No.", (InvNoLen - 3), 4);
                END;


                Loc.GET("Location Code");

                RecState.SETRANGE(RecState.Code, Loc."State Code");
                IF RecState.FINDFIRST THEN
                    State1 := RecState.Description;

                SalesInvLine.RESET;
                SalesInvLine.SETRANGE(SalesInvLine."Document No.", "No.");
                SalesInvLine.SETRANGE(SalesInvLine.Type, SalesInvLine.Type::"G/L Account");
                SalesInvLine.SETRANGE(SalesInvLine."No.", RoundOffAcc);
                IF SalesInvLine.FINDFIRST THEN BEGIN
                    RoundOffAmnt := SalesInvLine."Line Amount";
                END;

                CLEAR(BEDAmount);
                CLEAR(EcessAmt);
                CLEAR(SheCess);
                CLEAR(AddlDuty);
                vSalesInv.RESET;
                vSalesInv.SETRANGE("Document No.", "No.");
                IF vSalesInv.FINDSET THEN
                    REPEAT
                        BEDAmount += 0;// vSalesInv."BED Amount";
                        EcessAmt += 0;// vSalesInv."eCess Amount";
                        SheCess += 0;// vSalesInv."Amount To Customer" + RoundOffAmnt;
                        AddlDuty += 0;// vSalesInv."SHE Cess Amount";
                    UNTIL vSalesInv.NEXT = 0;
            end;

            trigger OnPreDataItem()
            begin

                CompanyInfo1.GET;
                CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
                CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
                CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
            end;
        }
    }

    requestpage
    {

        layout
        {
        }

        actions
        {
        }
    }

    labels
    {
    }

    var
        Loc: Record 14;
        SalesPerson: Record 13;
        RespCentre: Record 5714;
        RecState: Record State;
        RecState1: Record State;
        RecCountry: Record 9;
        RecCountry1: Record 9;
        CompanyInfo1: Record 79;
        RecCust: Record 18;
        RecCust1: Record 18;
        RecItem: Record 27;
        ShipToCode: Record 222;
        "Payment Terms": Record 3;
        "Shipment Method": Record 10;
        "Shipping Agent": Record 291;
        ILE: Record 32;
        // "E.C.C. Nos.": Record 13708;
        CustPostSetup: Record 92;
        CustName: Text[100];
        UOM: Code[10];
        Qty: Decimal;
        QtyPerUOM: Decimal;
        ShiptoName: Text[80];
        ShipToCST: Code[50];
        ShiptoTIN: Code[50];
        ShiptoLST: Code[50];
        ECCNo: Code[50];
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        DescriptionLineDuty: array[2] of Text[132];
        DescriptionLineeCess: array[2] of Text[132];
        DescriptionLineSHeCess: array[2] of Text[132];
        DescriptionLineTot: array[2] of Text[132];
        PackDescription: Text[30];
        SrNo: Integer;
        InvNo: Code[20];
        InvNoLen: Integer;
        RoundOffAcc: Code[20];
        RoundOffAmnt: Decimal;
        "<EBT>": Integer;
        ItemVarParmResFinal: Record 50015;
        TaxDescr: Code[10];
        ECCDesc: Code[30];
        PackDescriptionforILE: Text[30];
        TaxJurisd: Record 320;
        TaxDescription: Text[50];
        Freight: Decimal;
        Discount: Decimal;
        TotalAmttoCustomer: Decimal;
        InvTotal: Decimal;
        CurrDatetime: Time;
        State: Code[30];
        Ctr: Integer;
        SalesInvLine: Record 113;
        SegManagement: Codeunit SegManagement;
        LogInteraction: Boolean;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        cust: Record 18;
        copytext: Text[30];
        //SalesInvCountPrinted: Codeunit 315;
        ShowInternalInfo: Boolean;
        InvType: Text[50];
        InvType1: Text[30];
        i: Integer;
        Value: Decimal;
        Intermvalue: Decimal;
        State1: Text[30];
        // DetailedTaxEntry: Record 16522;
        FormCode: Code[30];
        SuppliersCode: Text[30];
        VehOwnersName: Text[100];
        st: Code[20];
        RefInvDate: Date;
        SalesInvHdr: Record 112;
        RefInvNo: Text[30];
        Rate: Decimal;
        ItemREC: Record 27;
        ItemQty: Decimal;
        TotalItemQty: Decimal;
        recSIL: Record 113;
        //  recExcisePostingSetup: Record 13711;
        // PostedStrOrdDetails: Record 13798;
        cessvalue: Decimal;
        Cesstext: Text[30];
        ADDTAX: Decimal;
        //  PostedStrOrdDetails1: Record 13798;
        ADDTAXText: Text[30];
        AppliedInvoiceNo: Code[20];
        recSIH: Record 112;
        Rupees: Code[10];
        InvoiceNoLen: Integer;
        InvoiceNo: Code[10];
        Ship2Name: Text[100];
        SalesInvLine1: Record 113;
        RecCompInfo: Record 79;
        Text003: Label 'Original for Buyer';
        Text026: Label 'Zero';
        Text027: Label 'Hundred';
        Text028: Label '&';
        Text029: Label '%1 results in a written number that is too long.';
        Text032: Label 'One';
        Text033: Label 'Two';
        Text034: Label 'Three';
        Text035: Label 'Four';
        Text036: Label 'Five';
        Text037: Label 'Six';
        Text038: Label 'Seven';
        Text039: Label 'Eight';
        Text040: Label 'Nine';
        Text041: Label 'Ten';
        Text042: Label 'Eleven';
        Text043: Label 'Twelve';
        Text044: Label 'Thirteen';
        Text045: Label 'Fourteen';
        Text046: Label 'Fifteen';
        Text047: Label 'Sixteen';
        Text048: Label 'Seventeen';
        Text049: Label 'Eighteen';
        Text050: Label 'Nineteen';
        Text051: Label 'Twenty';
        Text052: Label 'Thirty';
        Text053: Label 'Forty';
        Text054: Label 'Fifty';
        Text055: Label 'Sixty';
        Text056: Label 'Seventy';
        Text057: Label 'Eighty';
        Text058: Label 'Ninety';
        Text059: Label 'Thousand';
        Text060: Label 'Million';
        Text061: Label 'Billion';
        Text1280000: Label 'Lakh';
        Text1280001: Label 'Crore';
        "--": Integer;
        vCount: Integer;
        vCheck: Report "Check Report";
        BEDAmount: Decimal;
        EcessAmt: Decimal;
        SheCess: Decimal;
        AddlDuty: Decimal;
        vSalesInv: Record 113;
        OutPutNo: Integer;
        Loc2: Record 14;
}


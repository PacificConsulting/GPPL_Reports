report 50000 "Export Invoice"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ExportInvoice.rdl';
    ApplicationArea = all;
    UsageCategory = ReportsAndAnalysis;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            column(Loc_Address; recLocation.Address + '  ' + recLocation."Address 2" + '  ' + recLocation.City + '  ' + recLocation."Post Code" + '  ' + recState.Description + '  ' + ' Phone No ' + recLocation."Phone No." + '  ' + 'Email:' + recLocation."E-Mail")
            {

            }
            column(No_SalesInvoiceHeader2; "Sales Invoice Header"."No.")
            {
            }
            column(Loc_Address2; recLocation."Address 2")
            {
            }
            column(Loc_city; recLocation.City)
            {
            }
            column(Loc_postcode; recLocation."Post Code")
            {
            }
            column(Loc_phoneNo; recLocation."Phone No.")
            {
            }
            column(Loc_Email; recLocation."E-Mail")
            {
            }
            column(State_Description; recState.Description)
            {
            }
            column(recLocTinNo; 'recLocation."T.I.N. No." ' + ' & ' + 'recLocation."C.S.T No."')
            {
            }
            column(Company_homepage_regisNo_PANNO; 'Wesite:' + recCompanyInfo."Home Page" + ' CIN No. : ' + 'recCompanyInfo."Company Registration  No."' + ' PAN NO. :  ' + recCompanyInfo."P.A.N. No.")
            {
            }
            column(Company_registrationNo; 'recCompanyInfo."Company Registration  No."')
            {
            }
            column(Company_PANNo; recCompanyInfo."P.A.N. No.")
            {
            }
            column(EccDescription; EccDescription)
            {
            }
            column(EccNo_CERange_EccNoCommsion_EccNoDivision; 'recEccNo."C.E. Range" ' + '& ' + ' recEccNo."C.E. Division" ' + '&' + ' recEccNo."C.E. Commissionerate"')
            {
            }
            column(EccNo_CommisionRate; 'recEccNo."C.E. Commissionerate"')
            {
            }
            column(EccNo_Division; 'recEccNo."C.E. Division"')
            {
            }
            column(EccNo_Address1; 'recEccNo."C.E Address1"')
            {
            }
            column(EccNo_Address2; 'recEccNo."C.E Address2"')
            {
            }
            column(EccNoCity_PostCode; 'recEccNo."C.E City"' + ' -' + ' recEccNo."C.E Post Code"')
            {
            }
            column(PaymentMethod_Description; recPaymentMethod.Description)
            {
            }
            column(VehOwnersName; VehOwnersName)
            {
            }
            column(recState_desc; recState.Description + ',' + recCountry.Name)
            {
            }
            column(recCountry_Name; recCountry.Name)
            {
            }
            column(SuppliersCode; SuppliersCode)
            {
            }
            column(recCustomer_EccNo; 'recCustomer."E.C.C. No."')
            {
            }
            column(recCustomer_CSTNo; 'recCustomer."C.S.T. No."')
            {
            }
            column(recCustomer_TINNo; 'recCustomer."T.I.N. No."')
            {
            }
            column(Ship2Name; Ship2Name)
            {
            }
            column(SH_InvoiceNo; COPYSTR("Sales Invoice Header"."No.", 16, 4))
            {
            }
            column(SH_Posting_Date; "Sales Invoice Header"."Posting Date")
            {
            }
            column(SH_ExternalDocmentNo; "Sales Invoice Header"."External Document No.")
            {
            }
            column(SH_Order_Date; "Sales Invoice Header"."Order Date")
            {
            }
            column(SH_VechiNo; "Sales Invoice Header"."Vehicle No.")
            {
            }
            column(SH_LRNo; "Sales Invoice Header"."LR/RR No.")
            {
            }
            column(SH_LrDate; "Sales Invoice Header"."LR/RR Date")
            {
            }
            column(SH_FrightType; "Sales Invoice Header"."Freight Type")
            {
            }
            column(SH_InvoicePrintTime; "Sales Invoice Header"."Invoice Print Time")
            {
            }
            column(SH_fullName; "Sales Invoice Header"."Full Name")
            {
            }
            column(SH_BillToAddress; "Sales Invoice Header"."Bill-to Address")
            {
            }
            column(SH_BillToAddress2; "Sales Invoice Header"."Bill-to Address 2")
            {
            }
            column(SH_BillToCity; "Sales Invoice Header"."Bill-to City" + ' -' + "Sales Invoice Header"."Bill-to Post Code")
            {
            }
            column(SH_ShipToAddress; "Sales Invoice Header"."Ship-to Address")
            {
            }
            column(SH_ShipToAddress2; "Sales Invoice Header"."Ship-to Address 2")
            {
            }
            column(SH_ShipToCity_ShipToPostcode; "Sales Invoice Header"."Ship-to City" + ' -' + "Sales Invoice Header"."Ship-to Post Code")
            {
            }
            column(SH_CurrancyFactor; "Sales Invoice Header"."Currency Factor")
            {
            }
            column(ShipToCST; ShipToCST)
            {
            }
            column(ShiptoTIN; ShiptoTIN)
            {
            }
            column(ShiptoLST; ShiptoLST)
            {
            }
            column(EccNo; EccNo)
            {
            }
            column(reccountry1_Name; recCountry1.Name)
            {
            }
            column(Reamark1; "Sales Invoice Header".Remarks)
            {
            }
            column(Remark2; "Sales Invoice Header".Remarks2)
            {
            }
            column(State1; State1)
            {
            }
            column(InsuCompanyInf; CompanyInfo1."Insurance No.")
            {
            }
            column(a; CompanyInfo1."Insurance Provider")
            {
            }
            column(Freight; Freight)
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
                        DataItemTableView = WHERE(Quantity = FILTER('<> 0'));
                        //"Excise Prod. Posting Group"=FILTER(<>''));
                        column(SL_ChapterNo; '"Sales Invoice Line"."Excise Prod. Posting Group"')
                        {
                        }
                        column(DocumentNo_SalesInvoiceLine; "Sales Invoice Line"."Document No.")
                        {
                        }
                        column(SL_QtyBAse; "Sales Invoice Line"."Quantity (Base)")
                        {
                        }
                        column(SL_BedAmt; '"Sales Invoice Line"."BED Amount"')
                        {
                        }
                        column(SL_eCessAmount; '"Sales Invoice Line"."eCess Amount"')
                        {
                        }
                        column(SL_SHECESSAmt; '"Sales Invoice Line"."SHE Cess Amount"')
                        {
                        }
                        column(SL_SheCessEcess; '"Sales Invoice Line"."SHE Cess Amount" + "Sales Invoice Line"."eCess Amount"')
                        {
                        }
                        column(SL_UOM; SalesInvLine."Qty. per Unit of Measure")
                        {
                        }
                        column(BatchNo; BatchNo)
                        {
                        }
                        column(PackDescription; PackDescription)
                        {
                        }
                        column(BEDpercent; BEDpercent)
                        {
                        }
                        column(DutyAmtperunit; DutyAmtperunit)
                        {
                        }
                        column(ecesspercent; ecesspercent)
                        {
                        }
                        column(SHEpercent; SHEpercent)
                        {
                        }
                        column(SL_TaxAmount; '"Sales Invoice Line"."Tax Amount"')
                        {
                        }
                        column(RoundOffAmnt; RoundOffAmnt)
                        {
                        }
                        column(InvoiceGrandTotal; InvoiceGrandTotal)
                        {
                        }
                        column(AddlDutyAmount; AddlDutyAmount)
                        {
                        }
                        column(DiscountCharge; DiscountCharge)
                        {
                        }
                        column(DimensionName; DimensionName)
                        {
                        }
                        column(BasicRate2; BasicRate)
                        {
                        }
                        column(AmoutTOWords; AmoutTOWords2)
                        {
                        }
                        column(LineAmount_SalesInvoiceLine; LineAmt)
                        {
                        }
                        column(BasicRate; "Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure" / "Sales Invoice Header"."Currency Factor")
                        {
                        }
                        column(DescriptionLineDuty; DescriptionLineDuty[1])
                        {
                        }
                        column(DescriptionLineeCess; DescriptionLineeCess[1])
                        {
                        }
                        column(DescriptionLineSHeCess; DescriptionLineSHeCess[1])
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin
                            // Code to get Package Description.
                            Qty := Quantity;
                            QtyPerUOM := "Sales Invoice Line"."Qty. per Unit of Measure";

                            IF "Sales Invoice Line"."Unit of Measure Code" = 'BRL' THEN BEGIN
                                PackDescription := FORMAT(Qty) + '*' +
                                FORMAT("Sales Invoice Line"."Qty. per Unit of Measure") + '' + "Sales Invoice Line"."Unit of Measure Code";
                            END ELSE BEGIN
                                PackDescription := FORMAT(Qty) + '*' + "Sales Invoice Line"."Unit of Measure Code";
                            END;


                            // code to get batch no.
                            recVLE.RESET;
                            recVLE.SETRANGE(recVLE."Document No.", "Sales Invoice Line"."Document No.");
                            recVLE.SETRANGE(recVLE."Item No.", "Sales Invoice Line"."No.");
                            recVLE.SETRANGE(recVLE."Document Line No.", "Sales Invoice Line"."Line No.");
                            IF recVLE.FINDFIRST THEN
                                IF recILE.GET(recVLE."Item Ledger Entry No.") THEN
                                    BatchNo := recILE."Lot No.";
                            // code to get batch no.

                            /*
                            //
                            
                            recExcisePostnSetup.RESET;
                            recExcisePostnSetup.SETRANGE("Excise Bus. Posting Group","Sales Invoice Line"."Excise Bus. Posting Group");
                            recExcisePostnSetup.SETRANGE("Excise Prod. Posting Group","Sales Invoice Line"."Excise Prod. Posting Group");
                            IF recExcisePostnSetup.FINDFIRST THEN BEGIN
                              BEDpercent := recExcisePostnSetup."BED %";
                              ecesspercent := recExcisePostnSetup."eCess %";
                              SHEpercent := recExcisePostnSetup."SHE Cess %";
                            END;
                            
                            
                            
                            IF "Sales Invoice Line"."BED Amount" = 0 THEN
                            BEGIN
                             BEDpercent := 0;
                             DutyAmtperunit := 0
                            END ELSE
                            DutyAmtperunit := (("Sales Invoice Line"."Unit Price"/"Sales Invoice Line"."Qty. per Unit of Measure")*BEDpercent)/100;
                            //
                            
                            */
                            //RSPL--

                            DimensionName := '';
                            /*
                            IF "Inventory Posting Group"='MERCH' THEN BEGIN
                             recPosteddocDiemension.RESET;
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Table ID",113);
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Document No.","Sales Invoice Line"."Document No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Line No.","Sales Invoice Line"."Line No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Dimension Code",'MERCH');
                              IF recPosteddocDiemension.FINDFIRST THEN
                                 dimensionCode:=recPosteddocDiemension."Dimension Value Code";
                            
                               recDimensionValue.RESET;
                               recDimensionValue.SETRANGE(recDimensionValue.Code,dimensionCode);
                               IF recDimensionValue.FINDFIRST THEN
                                 DimensionName:=recDimensionValue.Name;
                                 DimensionName2:=DimensionName;
                            
                            END;
                            */
                            /*IF DimensionName=''THEN
                              DimensionName:='IPOL '+Description;  */

                            //RSPL

                            //RSPL For Category show only item description
                            IF (DimensionName = '') AND ("Sales Invoice Line"."Item Category Code" <> 'CAT01')
                            AND ("Sales Invoice Line"."Item Category Code" <> 'CAT04') AND ("Sales Invoice Line"."Item Category Code" <> 'CAT05')
                            AND ("Sales Invoice Line"."Item Category Code" <> 'CAT06') AND ("Sales Invoice Line"."Item Category Code" <> 'CAT07')
                            AND ("Sales Invoice Line"."Item Category Code" <> 'CAT08') AND ("Sales Invoice Line"."Item Category Code" <> 'CAT09')
                            AND ("Sales Invoice Line"."Item Category Code" <> 'CAT10') THEN
                                DimensionName := 'IPOL ' + Description
                            //ELSE
                            //  DimensionName:=Description;

                            ELSE
                                IF DimensionName <> '' THEN
                                    DimensionName := DimensionName2
                                ELSE
                                    DimensionName := Description;
                            //RSPL


                            // code to calculate roundoff amount

                            SalesInvLine.RESET;
                            SalesInvLine.SETRANGE(SalesInvLine."Document No.", "Sales Invoice Line"."Document No.");
                            SalesInvLine.SETRANGE(SalesInvLine.Type, SalesInvLine.Type::"G/L Account");
                            SalesInvLine.SETRANGE(SalesInvLine."No.", RoundOffAcc);
                            IF SalesInvLine.FINDFIRST THEN BEGIN
                                RoundOffAmnt := SalesInvLine."Line Amount";
                            END;
                            // code to calculate roundoff amount


                            // to calculate grand total.
                            /*
                            IF "Sales Invoice Header"."Currency Factor" > 0 THEN BEGIN
                                InvoiceGrandTotal += (("Sales Invoice Line"."Line Amount" / "Sales Invoice Header"."Currency Factor")
                                                       - "Sales Invoice Line"."Inv. Discount Amount" + "Sales Invoice Line"."BED Amount"
                                                      + "Sales Invoice Line"."eCess Amount" + "Sales Invoice Line"."SHE Cess Amount"
                                                       + "Sales Invoice Line"."Tax Amount" +
                                                       RoundOffAmnt);
                                LineAmt := "Sales Invoice Line"."Line Amount" / "Sales Invoice Header"."Currency Factor";
                            END ELSE BEGIN
                                InvoiceGrandTotal += (("Sales Invoice Line"."Line Amount")
                                                       - "Sales Invoice Line"."Inv. Discount Amount" + "Sales Invoice Line"."BED Amount"
                                                      + "Sales Invoice Line"."eCess Amount" + "Sales Invoice Line"."SHE Cess Amount"
                                                       + "Sales Invoice Line"."Tax Amount" +
                                                       +RoundOffAmnt);
                                "Sales Invoice Line".gstr
                                LineAmt := "Sales Invoice Line"."Line Amount"
                            END;
*/
                            // to calculate grand total.
                            //FormatNoText(DescriptionLineTot,ROUND((InvoiceGrandTotal+RoundOffAmnt),0.01),'');
                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(AmoutTOWords, ROUND((InvoiceGrandTotal + RoundOffAmnt + (Freight / "Sales Invoice Header"."Currency Factor")), 0.01), "Sales Invoice Header"."Currency Code");//17Apr2017
                            //vCheck.FormatNoText(AmoutTOWords,ROUND((InvoiceGrandTotal+RoundOffAmnt+Freight),0.01),"Sales Invoice Header"."Currency Code");
                            AmoutTOWords[1] := DELCHR(AmoutTOWords[1], '=', '*');
                            AmoutTOWords2 := AmoutTOWords[1];

                            /*
                                                        vCheck.InitTextVariable;
                                                        vCheck.FormatNoText(DescriptionLineDuty, ROUND("BED Amount", 0.01), '');
                                                        DescriptionLineDuty[1] := DELCHR(DescriptionLineDuty[1], '=', '*');

                                                        vCheck.InitTextVariable;
                                                        vCheck.FormatNoText(DescriptionLineeCess, ROUND("eCess Amount", 0.01), '');
                                                        DescriptionLineeCess[1] := DELCHR(DescriptionLineeCess[1], '=', '*');

                                                        vCheck.InitTextVariable;
                                                        vCheck.FormatNoText(DescriptionLineSHeCess, ROUND("SHE Cess Amount", 0.01), '');
                                                        DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1], '=', '*');

                            */
                        end;

                        trigger OnPreDataItem()
                        begin


                            // Code to get Package Description.
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin



                        i += 1;
                        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START
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

                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'SA') THEN BEGIN
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

                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'SA') THEN BEGIN
                            IF i = 1 THEN
                                InvType := 'EXTRA COPY';
                            InvType1 := ' NOT FOR COMMERCIAL USE';
                        END;
                        //MESSAGE(InvType);
                        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
                    end;

                    trigger OnPreDataItem()
                    begin
                        InvoiceGrandTotal := 0;
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    OutPutNo += 1;
                    IF Number > 1 THEN
                        copytext := Text003;
                    //CurrReport.PAGENO := 1;
                end;

                trigger OnPreDataItem()
                begin

                    //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START

                    IF NOT (USERID = 'sa') THEN BEGIN
                        IF "Sales Invoice Header"."Print Invoice" = TRUE THEN BEGIN
                            SETRANGE(Number, 1, 1);
                            //SETRANGE(Number,1,3);
                        END
                        ELSE BEGIN
                            NoOfCopies := 4;
                            NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
                            IF NoOfLoops <= 0 THEN
                                NoOfLoops := 1;
                            copytext := '';
                            SETRANGE(Number, 1, NoOfLoops);
                        END
                    END
                    ELSE BEGIN
                        NoOfCopies := 4;
                        NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
                        IF NoOfLoops <= 0 THEN
                            NoOfLoops := 1;
                        copytext := '';
                        SETRANGE(Number, 1, NoOfLoops);
                    END;

                    //SETRANGE(Number,1,NoOfCopies);
                    //MESSAGE('%1',ABS(NoOfCopies));
                    //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
                    //i:= 1
                end;
            }

            trigger OnAfterGetRecord()
            begin


                //to get location details.
                recLocation.GET("Sales Invoice Header"."Location Code");
                //to get location details.

                //code to get state details.
                //recState.GET(recLocation."State Code");
                IF recState.GET("Sales Invoice Header".State) THEN;
                //code to get state details.

                //code to get Country description.
                IF "Sales Invoice Header"."Bill-to Country/Region Code" <> ''
                  THEN
                    recCountry.GET("Sales Invoice Header"."Bill-to Country/Region Code")
                ELSE
                    recCountry.Name := '';
                //code to get Country description.

                // code to get RoundAccount.
                CustPostSetup.RESET;
                CustPostSetup.FINDFIRST;
                RoundOffAcc := CustPostSetup."Invoice Rounding Account";
                // code to get RoundAccount.

                //code to get ECC NoS.
                /*
                recEccNo.SETRANGE(recEccNo.Code, recLocation."E.C.C. No.");
                IF recEccNo.FINDFIRST THEN
                    EccDescription := recEccNo.Description;
                    */
                //code to get ECC NoS.

                //code to get the payment description
                //recPaymentMethod.GET("Sales Invoice Header"."Payment Terms Code");
                //code to get the payment description

                //code to get the Transport details.
                IF recShippingAgent.GET("Sales Invoice Header"."Shipping Agent Code") THEN
                    VehOwnersName := recShippingAgent.Name;
                //code to get the Transport details.

                //code to get customer no.
                IF recCustomer.GET("Sales Invoice Header"."Bill-to Customer No.") THEN
                    SuppliersCode := recCustomer."Our Account No.";
                //code to get customer no.

                //Code to get Ship to Name.
                IF "Sales Invoice Header"."Ship-to Code" <> '' THEN
                    Ship2Name := "Sales Invoice Header"."Full Name"
                ELSE
                    Ship2Name := "Sales Invoice Header"."Ship-to Name";
                //Code to get Ship to Name.

                //Code to get Ship to Country name and state description
                IF NOT recCountry1.GET("Sales Invoice Header"."Ship-to Country/Region Code") THEN
                    recCountry1.GET("Sales Invoice Header"."Bill-to Country/Region Code");
                //Code to get Ship to Country name and state

                //code to get ship to code details.
                IF recShiptoCode.GET("Sales Invoice Header"."Sell-to Customer No.", "Sales Invoice Header"."Ship-to Code") THEN BEGIN
                    EccNo := recShiptoCode."E.C.C. No.";
                    ShiptoName := recShiptoCode.Name;
                    ShipToCST := recShiptoCode."C.S.T. No.";
                    ShiptoTIN := recShiptoCode."T.I.N. No.";
                    ShiptoLST := recShiptoCode."L.S.T. No.";
                END
                ELSE BEGIN
                    EccNo := 'recCustomer."E.C.C. No."';
                    ShiptoName := CustName;
                    ShipToCST := recShiptoCode."C.S.T. No.";
                    ShiptoTIN := recShiptoCode."T.I.N. No.";
                    ShiptoLST := recShiptoCode."L.S.T. No.";
                END;
                //code to get ship to code details.
                CLEAR(AddlDutyAmount);

                //code to calcluate Freight 15Mar2017
                /*
                                PostedStrOrdrLineDetails.RESET;
                                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Sales Invoice Header"."No.");
                                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                                PostedStrOrdrLineDetails.SETFILTER(PostedStrOrdrLineDetails."Tax/Charge Group", 'FREIGHT');
                                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                                    REPEAT
                                        Freight := +Freight + PostedStrOrdrLineDetails.Amount;
                                    UNTIL PostedStrOrdrLineDetails.NEXT = 0;
                */
                //code to calcluate Freight 15Mar2017

                /*
                //code to calculate additional duty amount.
                //Code commented becasue PostedStrOrdrLineDetails."Invoice No." not available in table.
                
                PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.","Sales Invoice Header"."No.");
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type",
                                                  PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Group",'ADDL.DUTY');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                REPEAT
                AddlDutyAmount:=AddlDutyAmount+PostedStrOrdrLineDetails.Amount;
                UNTIL PostedStrOrdrLineDetails.NEXT=0;
                
                //code to calculate additional duty amount.
                
                //code to calcluate Freight
                
                PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.","Sales Invoice Header"."No.");
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type",PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrLineDetails.SETFILTER(PostedStrOrdrLineDetails."Tax/Charge Group",'FREIGHT');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                REPEAT
                Freight:=+Freight+PostedStrOrdrLineDetails.Amount;
                UNTIL PostedStrOrdrLineDetails.NEXT=0;
                
                //code to calcluate Freight
                
                // code to get Discount Amount
                DiscountCharge := 0;
                PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.","Sales Invoice Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type",PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group",'DISCOUNT');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                DiscountCharge:=PostedStrOrdrDetails."Calculation Value";
                // code to get Discount Amount
                */
                Loc.GET("Location Code");
                recState2.RESET;
                recState2.SETRANGE(recState2.Code, Loc."State Code");
                IF recState2.FINDFIRST THEN
                    State1 := recState.Description;

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(NoOfCopies; NoOfCopies)
                {
                    Caption = 'No.Of Copies';
                }
                field(ShowInternalInfo; ShowInternalInfo)
                {
                    Caption = 'Show Internal Information';
                    Visible = false;
                }
                field(LogInteraction; LogInteraction)
                {
                    Caption = 'Log Interaction';
                    Visible = false;
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

    trigger OnPreReport()
    begin
        // code to get company information
        recCompanyInfo.GET;
        CompanyInfo1.GET;
        // code to get company information
    end;

    var
        recCompanyInfo: Record 79;
        recLocation: Record 14;
        recState: Record State;
        //recEccNo: Record eccno;
        EccDescription: Text[30];
        recPaymentMethod: Record 289;
        recShippingAgent: Record 291;
        VehOwnersName: Text[100];
        recCountry: Record 9;
        recCountry1: Record 9;
        recCustomer: Record 18;
        SuppliersCode: Text[30];
        Ship2Name: Text[100];
        recShiptoCode: Record 222;
        EccNo: Code[50];
        ShiptoName: Text[80];
        ShipToCST: Code[50];
        ShiptoTIN: Code[50];
        ShiptoLST: Code[50];
        CustName: Text[100];
        recILE: Record 32;
        recVLE: Record 5802;
        BatchNo: Code[10];
        PackDescription: Text[30];
        Qty: Decimal;
        QtyPerUOM: Decimal;
        //recExcisePostnSetup: Record "13711";
        BEDpercent: Decimal;
        ecesspercent: Decimal;
        SHEpercent: Decimal;
        DutyAmtperunit: Decimal;
        AddlDutyAmount: Decimal;
        //PostedStrOrdrLineDetails: Record "13798";
        DiscountCharge: Decimal;
        //PostedStrOrdrDetails: Record "Posted Structure Order Details";
        Freight: Decimal;
        RoundOffAmnt: Decimal;
        SalesInvLine: Record 113;
        CustPostSetup: Record 92;
        RoundOffAcc: Code[20];
        InvoiceGrandTotal: Decimal;
        DimensionName: Text[100];
        recDimensionValue: Record 349;
        dimensionCode: Code[20];
        DimensionName2: Text[100];
        BasicRate: Decimal;
        // "--": Integer;
        vCheck: Report "Check Report";
        AmoutTOWords2: Text;
        AmoutTOWords: array[5] of Text;
        Loc: Record 14;
        State1: Text;
        CompanyInfo1: Record 79;
        LineAmt: Decimal;
        DescriptionLineDuty: array[5] of Text;
        DescriptionLineeCess: array[5] of Text;
        DescriptionLineSHeCess: array[5] of Text;
        recState2: Record State;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        cust: Record 18;
        copytext: Text[30];
        OutPutNo: Integer;
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
        i: Integer;
        InvType: Text[50];
        InvType1: Text[30];
        ShowInternalInfo: Boolean;
        LogInteraction: Boolean;
        BillPANNo: Code[10];
        ShipPANNo: Code[10];

}


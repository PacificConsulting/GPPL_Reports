report 50199 "Depot AL GST"
{
    // 23May2017 : Depot Inv_ALNonExciseable(#50059) --> Saves as Depot Inv_ALNonExciseable GST(#50199)
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/DepotALGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Permissions = TableData 112 = rim;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(SIH_No; "Sales Invoice Header"."No.")
            {
            }
            column(SIH_PostingDate; "Sales Invoice Header"."Posting Date")
            {
            }
            column(SIH_ExDocNo; "Sales Invoice Header"."External Document No.")
            {
            }
            column(SIH_OrderDate; "Sales Invoice Header"."Order Date")
            {
            }
            column(SIH_RoadPermitNo; "Sales Invoice Header"."Road Permit No.")
            {
            }
            column(SIH_LRNo; "Sales Invoice Header"."LR/RR No.")
            {
            }
            column(SIH_LRDate; "Sales Invoice Header"."LR/RR Date")
            {
            }
            column(SIH_VehNo; "Sales Invoice Header"."Vehicle No.")
            {
            }
            column(SIH_Freight; "Sales Invoice Header"."Freight Type")
            {
            }
            column(SIH_InvoiceTime; "Sales Invoice Header"."Invoice Print Time")
            {
            }
            column(SIH_Remark; "Sales Invoice Header".Remarks)
            {
            }
            column(SIH_Remark2; "Sales Invoice Header".Remarks2)
            {
            }
            column(SIH_Payment; recPayment.Description)
            {
            }
            column(SIH_Transport; recShipping.Name)
            {
            }
            column(SIH_FullName; "Sales Invoice Header"."Full Name")
            {
            }
            column(SIH_BillAdd1; "Sales Invoice Header"."Bill-to Address")
            {
            }
            column(SIH_BillAdd2; "Sales Invoice Header"."Bill-to Address 2")
            {
            }
            column(SIH_BillAdd3; "Bill-to City" + '-' + "Bill-to Post Code" + ' ' + BillState + ', ' + BillCountry + '         PAN No.:' + recCustomer."P.A.N. No.")
            {
            }
            column(Ship2Name; Ship2Name)
            {
            }
            column(SIH_ShipName; "Sales Invoice Header"."Ship-to Name")
            {
            }
            column(SIH_ShipAdd1; "Sales Invoice Header"."Ship-to Address")
            {
            }
            column(SIH_ShipAdd2; "Sales Invoice Header"."Ship-to Address 2")
            {
            }
            column(SIH_ShipAdd3; "Ship-to City" + '-' + "Ship-to Post Code" + ' ' + ShipState + ', ' + ShipCountry + '     PAN No.:' + recCustomer."P.A.N. No.")
            {
            }
            column(CustLBTNo; recCustomer."LBT Reg. No.")
            {
            }
            column(CustRespone; recCustomer."Responsibility Center")
            {
            }
            column(CustECCNo; 'recCustomer."E.C.C. No."')
            {
            }
            column(CustCSTNo; 'recCustomer."C.S.T. No."')
            {
            }
            column(CustTINNo; 'recCustomer."T.I.N. No."')
            {
            }
            column(CustPANNo; recCustomer."P.A.N. No.")
            {
            }
            column(CustLSTNo; 'recCustomer."L.S.T. No."')
            {
            }
            column(CustKms; recCustomer."Customer Kms")
            {
            }
            column(LocAdd1; recLocation.Address)
            {
            }
            column(LocAdd2; recLocation."Address 2")
            {
            }
            column(LocAdd3; recLocation.City + ' - ' + recLocation."Post Code")
            {
            }
            column(LocState; LocState)
            {
            }
            column(LocPhoneNo; recLocation."Phone No.")
            {
            }
            column(LocEmail; recLocation."E-Mail")
            {
            }
            column(LocTINNo; 'recLocation."T.I.N. No."')
            {
            }
            column(LocCSTNo; 'recLocation."C.S.T No."')
            {
            }
            column(CompWeb; recCompany."Home Page")
            {
            }
            column(CompCINNo; recCompany."Registration No.")
            {
            }
            column(ComPANNo; recCompany."P.A.N. No.")
            {
            }
            column(ECCDesc; 'recECCNo.Description')
            {
            }
            column(ECCRange; 'recECCNo."C.E. Range"')
            {
            }
            column(ECCDivision; 'recECCNo."C.E. Division"')
            {
            }
            column(ECCComm; 'recECCNo."C.E. Commissionerate"')
            {
            }
            column(ECCAdd1; 'recECCNo."C.E Address1"')
            {
            }
            column(ECCAdd2; 'recECCNo."C.E Address2"')
            {
            }
            column(ECCAdd3; '')//recECCNo."C.E City" + '-' + recECCNo."C.E Post Code")
            {
            }
            column(CompanyPANNo; recCompany."P.A.N. No.")
            {
            }
            column(InvNo; InvNo)
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
            column(ShipECCNo; ECCNo)
            {
            }
            column(AddlDutyAmount; AddlDutyAmount)
            {
            }
            column(Freight; Freight)
            {
            }
            column(EntryTax; EntryTax)
            {
            }
            column(DiscountCharge; DiscountCharge)
            {
            }
            column(TaxDescription; TaxDescription)
            {
            }
            column(BedPer; BedPer)
            {
            }
            column(eCessPer; eCessPer)
            {
            }
            column(SheCessPer; SheCessPer)
            {
            }
            column(DutyAmtperunit; DutyAmtperunit)
            {
            }
            column(CompanyInsNo; recCompany."Insurance No.")
            {
            }
            column(CompanyInsProv; recCompany."Insurance Provider")
            {
            }
            column(CompanyInsExp; recCompany."Policy Expiration Date")
            {
            }
            column(CurrCode; CurrCode)
            {
            }
            column(ItmQtyPer1; ItmQtyPer1)
            {
            }
            column(ItmQty1; ItmQty1)
            {
            }
            column(ItmSheAmt1; ItmSheAmt1)
            {
            }
            column(ItmeCessAmt1; ItmeCessAmt1)
            {
            }
            column(ItmExciseAmt1; ItmExciseAmt1)
            {
            }
            column(ItmQtyBase1; ItmQtyBase1)
            {
            }
            column(ItmFinalAmt; AmtFinal)
            {
            }
            column(ItmLineAmt1; ItmLineAmt1)
            {
            }
            column(ItmLineDiscAmt1; ItmLineDiscAmt1)
            {
            }
            column(TPTSUBSIDYtext; TPTSUBSIDYtext)
            {
            }
            column(TPTSUBSIDY; TPTSUBSIDY)
            {
            }
            column(CashDiscText; CashDiscText)
            {
            }
            column(CashDiscountAmt; CashDiscountAmt)
            {
            }
            column(CCFCtext; CCFCtext)
            {
            }
            column(CCFC; CCFC)
            {
            }
            column(BranchOfficeAddress; BranchOfficeAddress)
            {
            }
            column(BranchOffText; BranchOffText)
            {
            }
            column(AddlTaxCessDesc; "AddlTax/CessDesc")
            {
            }
            column(AddlTaxCess; "AddlTax/Cess")
            {
            }
            column(TransportSubsidy; TransportSubsidy)
            {
            }
            column(CessText; CessText)
            {
            }
            column(CessValue; CessValue)
            {
            }
            column(EntryOctroi; "Entry/Octroi")
            {
            }
            column(ADDTAX; ADDTAX)
            {
            }
            column(ADDTAXText; ADDTAXText)
            {
            }
            column(CCFCLineamt; CCFCLineamt)
            {
            }
            column(TaxAmount; TaxAmount)
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
                    column(TaxInvoice; TaxInvoice)
                    {
                    }
                    dataitem("Sales Invoice Line"; "Sales Invoice Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Sales Invoice Header";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Quantity = FILTER(<> 0),
                                                  // "Excise Prod. Posting Group" = FILTER(<> ''),
                                                  Type = FILTER(Item));
                        column(LineNo; "Line No.")
                        {
                        }
                        column(ItmDocNo; "Sales Invoice Line"."Document No.")
                        {
                        }
                        column(ItmNo; "Sales Invoice Line"."No.")
                        {
                        }
                        column(ItmDesc1; 'IPOL  ' + "Sales Invoice Line".Description)
                        {
                        }
                        column(ItmDesc; DimensionName)
                        {
                        }
                        column(ItmDesc3; ItmDesc)
                        {
                        }
                        column(ItmBaseUOM; recItem."Base Unit of Measure")
                        {
                        }
                        column(ItmCrossRef; '"Sales Invoice Line"."Cross-Reference No."')
                        {
                        }
                        column(BatchNo; BatchNo)
                        {
                        }
                        column(MRPPrice; MRPPrice)
                        {
                        }
                        column(AssValue; AssValue)
                        {
                        }
                        column(ProdDisc; ProdDisc)
                        {
                        }
                        column(billingprice; billingprice)
                        {
                        }
                        column(PriceSupport; PriceSupport)
                        {
                        }
                        column(ItmListPrice; "Sales Invoice Line"."List Price")
                        {
                        }
                        column(chargeamttotal; chargeamttotal)
                        {
                        }
                        column(ItmQtyPer; "Sales Invoice Line"."Qty. per Unit of Measure")
                        {
                        }
                        column(ItmQty; "Sales Invoice Line".Quantity)
                        {
                        }
                        column(ItmUnitPrice; "Sales Invoice Line"."Unit Price")
                        {
                        }
                        column(ItmQtyBase; "Sales Invoice Line"."Quantity (Base)")
                        {
                        }
                        column(ItmExcise; 0)// "Sales Invoice Line"."Excise Prod. Posting Group")
                        {
                        }
                        column(ItmExciseRate; 0)// "Sales Invoice Line"."Excise Effective Rate")
                        {
                        }
                        column(ItmBEDAmt; 0)// "Sales Invoice Line"."BED Amount")
                        {
                        }
                        column(ItmeCessAmt; 0)// "Sales Invoice Line"."eCess Amount")
                        {
                        }
                        column(ItmSHeAmt; 0)//"Sales Invoice Line"."SHE Cess Amount")
                        {
                        }
                        column(ItmLineAmt; "Sales Invoice Line"."Line Amount")
                        {
                        }
                        column(ItmTaxAmt; 0)// "Sales Invoice Line"."Tax Amount")
                        {
                        }
                        column(ItmLineDiscAmt; "Sales Invoice Line"."Line Discount Amount")
                        {
                        }
                        column(ItmInvDiscAmt; "Sales Invoice Line"."Inv. Discount Amount")
                        {
                        }
                        column(RoundOffAmnt; RoundOffAmnt)
                        {
                        }
                        column(AmtinWordExcise; AmtinWordExcise[1])
                        {
                        }
                        column(AmtinWordeCess; AmtinWordeCess[1])
                        {
                        }
                        column(AmtinWordShe; AmtinWordSHe[1])
                        {
                        }
                        column(AmtinWordTotal; AmtinWordTotal[1])
                        {
                        }
                        column(ItmFinalAmt1; 0)// "Sales Invoice Line"."Amount To Customer")
                        {
                        }
                        column(TaxBaseAmt; 0)// "Sales Invoice Line"."Tax Base Amount")
                        {
                        }
                        column(TaxAmount_SalesInvoiceLine; 0)// "Sales Invoice Line"."Tax Amount")
                        {
                        }
                        column(TaxableAmt; 0)//"Sales Invoice Line"."Line Amount" + "Sales Invoice Line"."Excise Amount" + chargeamttotal + TPTSUBSIDY + CCFC + CashDiscountAmt)
                        {
                        }
                        column(NNN; NNN)
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin

                            //Item BaseUOM
                            IF recItem.GET("Sales Invoice Line"."No.") THEN;

                            //>>14Apr2017 DimsetEntry

                            DimensionName := '';
                            IF "Inventory Posting Group" = 'MERCH' THEN BEGIN

                                recDimSet.RESET;
                                recDimSet.SETRANGE("Dimension Set ID", "Sales Invoice Line"."Dimension Set ID");
                                recDimSet.SETRANGE("Dimension Code", 'MERCH');
                                IF recDimSet.FINDFIRST THEN BEGIN

                                    recDimSet.CALCFIELDS("Dimension Value Name");
                                    DimensionName := recDimSet."Dimension Value Name";

                                END;
                                DimensionName2 := DimensionName;
                            END;

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

                            //>>19Apr2017 RepsolName
                            IF "Sales Invoice Line"."Item Category Code" = 'CAT15' THEN
                                DimensionName := Description;

                            //<<14Apr2017 DimsetEntry



                            /*
                            //RSPL008 >>>> Item Category wise IPOL Description has print
                            ItmDesc :='';
                            IF ("Sales Invoice Line"."Item Category Code"='CAT01')
                            OR ("Sales Invoice Line"."Item Category Code"<>'CAT04') OR ("Sales Invoice Line"."Item Category Code"<>'CAT05')
                            OR ("Sales Invoice Line"."Item Category Code"<>'CAT06') OR ("Sales Invoice Line"."Item Category Code"<>'CAT07')
                            OR ("Sales Invoice Line"."Item Category Code"<>'CAT08') OR ("Sales Invoice Line"."Item Category Code"<>'CAT09')
                            OR ("Sales Invoice Line"."Item Category Code"<>'CAT10') OR ("Sales Invoice Line"."Item Category Code"<>'CAT15') THEN
                              ItmDesc := 'IPOL '+Description
                             ELSE
                              ItmDesc := Description;
                            //RSPL008 <<<<
                            
                            
                            //RSPL
                            //RSPL For Category show only item description
                            IF (DimensionName='') AND ("Sales Invoice Line"."Item Category Code"<>'CAT01')
                            AND ("Sales Invoice Line"."Item Category Code"<>'CAT04') AND("Sales Invoice Line"."Item Category Code"<>'CAT05')
                            AND ("Sales Invoice Line"."Item Category Code"<>'CAT06') AND("Sales Invoice Line"."Item Category Code"<>'CAT07')
                            AND("Sales Invoice Line"."Item Category Code"<>'CAT08') AND("Sales Invoice Line"."Item Category Code"<>'CAT09')
                            AND("Sales Invoice Line"."Item Category Code"<>'CAT10') AND ("Sales Invoice Line"."Item Category Code"<>'CAT15') THEN
                              DimensionName:='IPOL '+Description                                                               //RSPL008 CAT15 ADDED
                            //ELSE
                            //  DimensionName:=Description;
                            ELSE IF DimensionName<>'' THEN
                              DimensionName:=DimensionName2
                            ELSE
                                DimensionName:=Description;
                            //RSPL
                            */


                            //BatchNo
                            recVLE.RESET;
                            recVLE.SETRANGE(recVLE."Document No.", "Sales Invoice Line"."Document No.");
                            recVLE.SETRANGE(recVLE."Item No.", "Sales Invoice Line"."No.");
                            recVLE.SETRANGE(recVLE."Document Line No.", "Sales Invoice Line"."Line No.");
                            IF recVLE.FINDFIRST THEN
                                IF recILE.GET(recVLE."Item Ledger Entry No.") THEN
                                    BatchNo := recILE."Lot No.";

                            //Tax Details

                            /* recExcisePostnSetup.RESET;
                            recExcisePostnSetup.SETRANGE("Excise Bus. Posting Group", "Sales Invoice Line"."Excise Bus. Posting Group");
                            recExcisePostnSetup.SETRANGE("Excise Prod. Posting Group", "Sales Invoice Line"."Excise Prod. Posting Group");
                            IF recExcisePostnSetup.FINDFIRST THEN BEGIN
                                BedPer := recExcisePostnSetup."BED %";
                                eCessPer := recExcisePostnSetup."eCess %";
                                SheCessPer := recExcisePostnSetup."SHE Cess %";
                            END; */



                            //ChargeAmount


                            chargeamttotal := 0;
                            /* PostedStrOrdDetails.RESET;
                            PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                            PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Line No.", "Line No.");
                            PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::Charges);
                            //PostedStrOrdDetails.SETFILTER(PostedStrOrdDetails."Tax/Charge Group",'<>%1&<>%2&<>%3','CASHDISC','CCFC','TPTSUBSIDY');
                            PostedStrOrdDetails.SETFILTER(PostedStrOrdDetails."Tax/Charge Group", '<>%1&<>%2&<>%3', 'CASHDISC', 'CCFC', 'TRANSSUBS');  //mayuri
                            IF PostedStrOrdDetails.FINDFIRST THEN
                                REPEAT
                                    chargeamttotal += PostedStrOrdDetails.Amount;
                                UNTIL PostedStrOrdDetails.NEXT = 0; */

                            TPTSUBSIDY := 0;
                            TPTSUBSIDYtext := '';
                            /* PostedStrOrdDetails.RESET;
                            PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                            PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::Charges);
                            //PostedStrOrdDetails.SETFILTER(PostedStrOrdDetails."Tax/Charge Group",'%1','TPTSUBSIDY');
                            PostedStrOrdDetails.SETFILTER(PostedStrOrdDetails."Tax/Charge Group", '%1|%2', 'TPTSUBSIDY', 'TRANSSUBS'); //mayuri added 'TRANSSUBS'
                            IF PostedStrOrdDetails.FINDFIRST THEN
                                REPEAT
                                    TPTSUBSIDY += PostedStrOrdDetails.Amount;
                                    TPTSUBSIDYtext := 'Transport Subsidy';
                                UNTIL PostedStrOrdDetails.NEXT = 0; */

                            CCFCLineamt := 0;
                            /* PostedStrOrdDetails.RESET;
                            PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                            PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Line No.", "Line No.");
                            PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::Charges);
                            PostedStrOrdDetails.SETFILTER(PostedStrOrdDetails."Tax/Charge Group", '%1', 'CCFC');
                            IF PostedStrOrdDetails.FINDFIRST THEN
                                REPEAT
                                    CCFCLineamt += PostedStrOrdDetails.Amount;
                                UNTIL PostedStrOrdDetails.NEXT = 0; */



                            //MRPPrice 08Feb
                            //IF "Sales Invoice Line"."Price Inclusive of Tax" THEN
                            // IF "Sales Invoice Line"."Abatement %" = 0 THEN BEGIN
                            MRPPrice := 0;
                            AssValue := 0;// "Sales Invoice Line"."MRP Price";
                                          // END ELSE BEGIN
                            MRPPrice := 0;// "Sales Invoice Line"."MRP Price" * "Sales Invoice Line"."Qty. per Unit of Measure";

                            recMRPMaster.RESET;
                            recMRPMaster.SETRANGE(recMRPMaster."Item No.", "Sales Invoice Line"."No.");
                            recMRPMaster.SETRANGE(recMRPMaster."Lot No.", "Sales Invoice Line"."Lot No.");
                            IF recMRPMaster.FINDFIRST THEN BEGIN
                                AssValue := recMRPMaster."Assessable Value";
                            END;
                            // END;

                            // Charges Dis. added 020114
                            //IF "Sales Invoice Line"."Posting Date" > 071813D THEN
                            IF "Sales Invoice Line"."Posting Date" > 20180713D THEN BEGIN //180713D
                                CLEAR(ProdDisc);
                                MRPMaster.RESET;
                                MRPMaster.SETRANGE(MRPMaster."Item No.", "Sales Invoice Line"."No.");
                                MRPMaster.SETRANGE(MRPMaster."Lot No.", "Sales Invoice Line"."Lot No.");
                                IF MRPMaster.FINDFIRST THEN BEGIN
                                    ProdDisc := MRPMaster."National Discount";
                                END;
                            END;
                            //RSPL-Sourav140415
                            ProdDisc := "National Discount Per Ltr";
                            PriceSupport := "Price Support Per Ltr";
                            //RSPL-Sourav140415

                            //



                            /* // Commented by milan 301113
                            
                            billingprice := ("Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure");
                            */


                            billingprice := ("Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure")
                                            + (/*"Sales Invoice Line"."Excise Amount"*/0 / ("Sales Invoice Line"."Qty. per Unit of Measure" *
                                            "Sales Invoice Line".Quantity));

                            //   IF "Sales Invoice Line"."BED Amount" = 0 THEN BEGIN
                            BedPer := 0;
                            DutyAmtperunit := 0;
                            //  END ELSE
                            DutyAmtperunit := (("Sales Invoice Line"."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure") * BedPer) / 100;


                            //AmtFinal := AmtFinal+"Amount To Customer";
                            //AmountinWords
                            Check.InitTextVariable();
                            Check.FormatNoText(AmtinWordExcise, ROUND(ItmExciseAmt1, 0.01), '');//Excise
                            //Check.FormatNoText(AmtinWordExcise,ROUND("BED Amount",0.01),'');//Excise
                            AmtinWordExcise[1] := DELCHR(AmtinWordExcise[1], '=', '*');
                            //FormatNoText(AmtinWordExcise,ROUND("BED Amount",0.01),'');//Excise

                            Check.InitTextVariable();
                            Check.FormatNoText(AmtinWordeCess, ROUND(ItmeCessAmt1, 0.01), '');//eCess
                            //Check.FormatNoText(AmtinWordeCess,ROUND("eCess Amount",0.01),'');//eCess
                            AmtinWordeCess[1] := DELCHR(AmtinWordeCess[1], '=', '*');
                            //FormatNoText(AmtinWordeCess,ROUND("eCess Amount",0.01),'');//eCess

                            Check.InitTextVariable();
                            Check.FormatNoText(AmtinWordTotal, ROUND((AmtFinal + RoundOffAmnt), 0.01), '');//Total
                            AmtinWordTotal[1] := DELCHR(AmtinWordTotal[1], '=', '*');
                            //FormatNoText(AmtinWordTotal,ROUND(("Amount To Customer"+RoundOffAmnt),0.01),'');//Total

                            Check.InitTextVariable();
                            Check.FormatNoText(AmtinWordSHe, ROUND(ItmSheAmt1, 0.01), '');//She
                            //Check.FormatNoText(AmtinWordSHe,ROUND("SHE Cess Amount",0.01),'');//She
                            AmtinWordSHe[1] := DELCHR(AmtinWordSHe[1], '=', '*');
                            //FormatNoText(AmtinWordSHe,ROUND("SHE Cess Amount",0.01),'');//She

                        end;

                        trigger OnPreDataItem()
                        begin

                            NNN := COUNT;//12Apr2017
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin


                        /*
                        i+= 1 ;
                        
                        //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-----START
                        IF ("Sales Invoice Header"."Print Invoice" = FALSE) THEN
                        BEGIN
                        IF i = 1 THEN
                           InvType := 'ORIGINAL FOR BUYER'
                        ELSE IF i = 2 THEN
                           InvType := 'DUPLICATE FOR TRANSPORTER'
                        ELSE IF i = 3 THEN
                           InvType := 'TRIPLICATE FOR ASSESSEE'
                        ELSE IF i = 4 THEN BEGIN
                           InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                           InvType1 := 'NOT FOR CENVAT';
                        END ELSE IF i=5 THEN BEGIN
                           InvType := 'EXTRA COPY NOT FOR CENVAT';
                           InvType1 := '';
                        END
                        ELSE
                        InvType:='' ;
                        END;
                        
                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'sa') THEN
                        //IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'sa') THEN
                        BEGIN
                        IF i = 1 THEN
                           InvType := 'ORIGINAL FOR BUYER'
                        ELSE IF i = 2 THEN
                           InvType := 'DUPLICATE FOR TRANSPORTER'
                        ELSE IF i = 3 THEN
                           InvType := 'TRIPLICATE FOR ASSESSEE'
                        ELSE IF i = 4 THEN BEGIN
                           InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                           InvType1 := 'NOT FOR CENVAT';
                        END ELSE IF i=5 THEN BEGIN
                           InvType := 'EXTRA COPY NOT FOR CENVAT';
                           InvType1 := '';
                        END
                        ELSE
                        InvType:='' ;
                        END;
                        
                        
                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'sa')THEN
                        BEGIN
                          IF i = 1 THEN
                            InvType := 'EXTRA COPY';
                            InvType1 := ' NOT FOR COMMERCIAL USE';
                        END;
                        //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-------END
                        */

                        /* DetailedTaxEntry.RESET;
                        DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Sales Invoice Header"."No.");
                        IF DetailedTaxEntry.FINDFIRST THEN;
                        IF ("Sales Invoice Header"."Location Code" = 'BR0009') AND (DetailedTaxEntry."Tax Type" = DetailedTaxEntry."Tax Type"::CST) THEN BEGIN
                            TaxInvoice := 'RETAIL INVOICE';
                        END ELSE BEGIN
                            TaxInvoice := 'TAX INVOICE';
                        END; */

                        // EBT MILAN 030214 To Show Invoice instand of TAX invoice in case of UP State Invoice -------------------------START
                        RecLocation2.RESET;
                        RecLocation2.SETRANGE(RecLocation2.Code, "Sales Invoice Header"."Location Code");
                        IF RecLocation2.FINDFIRST THEN BEGIN
                            IF (RecLocation2."State Code" = 'UP') AND ("Sales Invoice Header"."Posting Date" >= 20140111D) THEN //011114D
                                TaxInvoice := 'INVOICE';
                        END;
                        // EBT MILAN 030214 To Show Invoice instand of TAX invoice in case of UP State Invoice ---------------------------END;


                        i += 1;

                        IF i = 1 THEN
                            InvType := 'ORIGINAL FOR BUYER'
                        ELSE
                            IF i = 2 THEN
                                InvType := 'DUPLICATE FOR TRANSPORTER'
                            ELSE
                                IF i = 3 THEN
                                    InvType := 'TRIPLICATE FOR ASSESSEE'
                                ELSE
                                    IF i = 4 THEN
                                        InvType := 'QUADRAPLICATE FOR ACCOUNTS'
                                    ELSE
                                        IF i = 5 THEN
                                            InvType := 'EXTRA COPY FOR ACCOUNTS'
                                        ELSE
                                            InvType := '';

                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    OutPutNo += 1;
                    /*

                   IF Number > 1 THEN
                     copytext := Text003;
                   //CurrReport.PAGENO := 1;
                   */

                    IF Number > 1 THEN
                        copytext := Text003;
                    CurrReport.PAGENO := 1;

                end;

                trigger OnPostDataItem()
                begin
                    /*
                    //Commented
                    IF NOT CurrReport.PREVIEW THEN
                      SalesInvCountPrinted.RUN("Sales Invoice Header");
                      */

                end;

                trigger OnPreDataItem()
                begin
                    //Commented by EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-----START
                    /*
                    NoOfCopies:=4;
                    NoOfLoops := ABS(NoOfCopies)+ 1;
                    copytext := '';
                    SETRANGE(Number,1,NoOfLoops);
                    */
                    //Commented by EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-------END
                    /*
                    //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-----START
                    IF NOT(USERID = 'sa') THEN
                    BEGIN
                     IF "Sales Invoice Header"."Print Invoice" = TRUE THEN
                     BEGIN
                      SETRANGE(Number,1,1);
                      //SETRANGE(Number,1,4);
                     END
                     ELSE
                     BEGIN
                       NoOfCopies:=4;
                       NoOfLoops := ABS(NoOfCopies)+ 1;
                       copytext := '';
                       SETRANGE(Number,1,NoOfLoops);
                       //SETRANGE(Number,1,4);
                     END
                    END
                    ELSE
                    BEGIN
                       NoOfCopies:=4;
                       NoOfLoops := ABS(NoOfCopies)+ 1;
                       copytext := '';
                       SETRANGE(Number,1,NoOfLoops);
                    END
                    
                    //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-------END
                    */

                    NoOfCopies := 4;
                    NoOfLoops := ABS(NoOfCopies) + 1;
                    copytext := '';
                    SETRANGE(Number, 1, NoOfLoops);

                end;
            }

            trigger OnAfterGetRecord()
            begin
                //Currency Code
                IF "Sales Invoice Header"."Currency Code" = '' THEN
                    CurrCode := 'INR'
                ELSE
                    CurrCode := "Currency Code";

                //BILLCustomer
                IF recCustomer.GET("Bill-to Customer No.") THEN;

                //12Apr2017 recCustomer."P.A.N. No."

                //SHIPCustomer
                //IF recCustomer1.GET();

                //PaymentTerms
                IF recPayment.GET("Payment Terms Code") THEN;

                //ShippingAgent
                IF recShipping.GET("Shipping Agent Code") THEN;

                //Location Code
                IF recLocation.GET("Location Code") THEN;

                //State Description
                recState.RESET;
                IF recState.GET(recLocation."State Code") THEN
                    LocState := recState.Description;

                //Bill State Description
                recState.RESET;
                //  IF recState.GET("Bill to Customer State") THEN
                BillState := recState.Description;

                //Bill Country Name
                recCountry.RESET;
                IF recCountry.GET("Bill-to Country/Region Code") THEN
                    BillCountry := recCountry.Name;



                //Ship Country Name
                recCountry.RESET;
                IF recCountry.GET("Ship-to Country/Region Code") THEN
                    ShipCountry := recCountry.Name;


                //Ship2Name
                IF "Sales Invoice Header"."Ship-to Code" = '' THEN
                    Ship2Name := "Sales Invoice Header"."Full Name"
                ELSE
                    Ship2Name := "Sales Invoice Header"."Ship-to Name";

                //Shipping

                IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code") THEN BEGIN
                    ECCNo := ShipToCode."E.C.C. No.";
                    //ShiptoName := ShipToCode.Name;
                    ShipToCST := ShipToCode."C.S.T. No.";
                    ShiptoLST := ShipToCode."L.S.T. No.";
                    ShiptoTIN := ShipToCode."T.I.N. No.";
                    //Ship State Description
                    recState.RESET;
                    IF recState.GET(ShipToCode.State) THEN
                        ShipState := recState.Description;

                END ELSE BEGIN
                    ECCNo := '';//recCustomer."E.C.C. No.";
                    ShipToCST := '';// recCustomer."C.S.T. No.";
                    ShiptoLST := '';// recCustomer."L.S.T. No.";
                    ShiptoTIN := '';// recCustomer."T.I.N. No.";
                    ShipState := BillState;
                END;



                //ECC No.
                // IF recECCNo.GET(recLocation."E.C.C. No.") THEN;


                //InvNo
                InvNoLen := STRLEN("Sales Invoice Header"."No.");
                InvNo := COPYSTR("Sales Invoice Header"."No.", (InvNoLen - 3), 4);
                /*
                //EBT STIVAN---(29012013)---Error Message POP UP,if ECC No of Location Code is not Specified----START
                recLocation.RESET;
                recLocation.SETRANGE(recLocation.Code,"Sales Invoice Header"."Location Code");
                IF recLocation.FINDFIRST THEN
                BEGIN
                 recECC.RESET;
                 recECC.SETRANGE(recECC.Code,recLocation."E.C.C. No.");
                 IF  recECC.FINDFIRST THEN
                 BEGIN
                  IF recECC.Description = '' THEN
                  ERROR('Your Location is not Excise Registered');
                 END;
                END;
                //EBT STIVAN---(29012013)---Error Message POP UP,if ECC No of Location Code is not Specified------END
                *///Commented 07APr2017

                //TaxCalculation

                //>>
                CustPostSetup.RESET;
                CustPostSetup.FINDFIRST;
                RoundOffAcc := CustPostSetup."Invoice Rounding Account";

                SIL.RESET;
                SIL.SETRANGE("Document No.", "No.");
                SIL.SETRANGE(Type, SIL.Type::"G/L Account");
                SIL.SETRANGE("No.", RoundOffAcc);
                IF SIL.FINDFIRST THEN
                    REPEAT
                        RoundOffAmnt := RoundOffAmnt + SIL."Line Amount";
                    UNTIL SIL.NEXT = 0;
                //<<

                //>>
                /* PostedStrOrdDetails.RESET;
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Code", 'CESS');
                IF PostedStrOrdDetails.FINDFIRST THEN
                    REPEAT
                        CessValue := CessValue + PostedStrOrdDetails.Amount;
                    UNTIL PostedStrOrdDetails.NEXT = 0;
                IF CessValue <> 0 THEN
                    CessText := 'Cess @' + FORMAT(PostedStrOrdDetails."Calculation Value") + ' ' + '%'
                ELSE
                    CessText := 'Cess %'; */

                /* PostedStrOrdDetails.RESET;
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::Charges);
                IF PostedStrOrdDetails.FINDFIRST THEN
                    REPEAT
                        IF PostedStrOrdDetails."Tax/Charge Group" = 'ENTRYTAX' THEN
                            "Entry/Octroi" := "Entry/Octroi" + PostedStrOrdDetails.Amount;
                    UNTIL PostedStrOrdDetails.NEXT = 0; */


                /* PostedStrOrdDetails.RESET;
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Code", 'ADDL TAX');
                IF PostedStrOrdDetails.FINDFIRST THEN BEGIN
                    ADDTAXText := 'Addl Tax @' + FORMAT(PostedStrOrdDetails."Calculation Value") + ' ' + '%';
                    REPEAT
                        ADDTAX := ADDTAX + PostedStrOrdDetails.Amount;
                    UNTIL PostedStrOrdDetails.NEXT = 0;
                END; */

                /* PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "Sales Invoice Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'FOC');
                IF PostedStrOrdrDetails.FINDSET THEN
                    REPEAT
                        FOCAMOUNT := FOCAMOUNT + PostedStrOrdDetails.Amount;
                    UNTIL PostedStrOrdDetails.NEXT = 0; */

                CashDiscountAmt := 0;
                /* recPostedStrOrdDetails.RESET;
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'CASHDISC');
                IF recPostedStrOrdDetails.FINDFIRST THEN
                    REPEAT
                        CashDiscountAmt += recPostedStrOrdDetails.Amount;
                    UNTIL recPostedStrOrdDetails.NEXT = 0;

                IF CashDiscountAmt <> 0 THEN
                    CashDiscText := 'Cash Discount @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%'
                ELSE
                    CessText := ''; */

                CCFCtext := '';
                CCFC := 0;
                /* recPostedStrOrdDetails.RESET;
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'CCFC');
                IF recPostedStrOrdDetails.FINDFIRST THEN BEGIN
                    CCFCtext := 'CCFC @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%';
                    REPEAT
                        CCFC += recPostedStrOrdDetails.Amount;
                    UNTIL recPostedStrOrdDetails.NEXT = 0;
                END; */

                CLEAR(TaxAmount);
                /* DetailedTaxEntry.RESET;
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "No.");
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Tax Area Code", "Tax Area Code");
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Tax Jurisdiction Code", "Tax Area Code");
                IF DetailedTaxEntry.FINDSET THEN
                    REPEAT
                        TaxAmount := TaxAmount + DetailedTaxEntry."Tax Amount";
                    UNTIL DetailedTaxEntry.NEXT = 0; */


                //<<
                /*
                //CCFC 08Feb2017
                
                CLEAR(AddlDutyAmount);
                DiscountCharge := 0;
                PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.","Sales Invoice Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type",PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group",'CASHDISC');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                DiscountCharge:=PostedStrOrdrDetails."Calculation Value";
                
                PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.","Sales Invoice Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type",PostedStrOrdrDetails."Tax/Charge Type"::"Other Taxes");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group",'ADDL.DUTY');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                AddlDutyAmount:=PostedStrOrdrDetails."Calculation Value";
                
                
                
                
                
                
                CashDiscountAmt := 0;
                recPostedStrOrdDetails.RESET;
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.","Sales Invoice Header"."No.");
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type",recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group",'CASHDISC');
                IF recPostedStrOrdDetails.FINDFIRST THEN
                REPEAT
                  CashDiscountAmt += recPostedStrOrdDetails.Amount;
                UNTIL recPostedStrOrdDetails.NEXT=0;
                
                IF CashDiscountAmt <> 0 THEN
                CashDiscText := 'Cash Discount @ '+ FORMAT(-recPostedStrOrdDetails."Calculation Value")+'%'
                ELSE
                 CashDiscText:='';
                
                
                
                CCFCtext := '';
                CCFC := 0;
                recPostedStrOrdDetails.RESET;
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.","Sales Invoice Header"."No.");
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type",recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group",'CCFC');
                IF recPostedStrOrdDetails.FINDFIRST THEN
                BEGIN
                 CCFCtext := 'CCFC @ '+ FORMAT(-recPostedStrOrdDetails."Calculation Value")+' '+'%';
                  REPEAT
                   CCFC += recPostedStrOrdDetails.Amount;
                  UNTIL recPostedStrOrdDetails.NEXT=0;
                END;
                
                */

                TPTSUBSIDY := 0;
                TPTSUBSIDYtext := '';
                /* PostedStrOrdDetails.RESET;
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Invoice Header"."No.");
                PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::Charges);
                PostedStrOrdDetails.SETFILTER(PostedStrOrdDetails."Tax/Charge Group", '%1', 'TPTSUBSIDY');
                IF PostedStrOrdDetails.FINDFIRST THEN
                    REPEAT
                        TPTSUBSIDY += PostedStrOrdDetails.Amount;
                        TPTSUBSIDYtext := 'Transport Subsidy';
                    UNTIL PostedStrOrdDetails.NEXT = 0; */
                //<<

                //CLEAR(AddlDutyAmount);



                /* PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Sales Invoice Header"."No.");
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrLineDetails.SETFILTER(PostedStrOrdrLineDetails."Tax/Charge Group", 'FREIGHT');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                    REPEAT
                        Freight := +Freight + PostedStrOrdrLineDetails.Amount;
                    UNTIL PostedStrOrdrLineDetails.NEXT = 0; */


                CLEAR(EntryTax);
                /*  PostedStrOrdrLineDetails.RESET;
                 PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Sales Invoice Header"."No.");
                 PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                 PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Group", 'ENTRYTAX');
                 IF PostedStrOrdrLineDetails.FINDFIRST THEN
                     REPEAT
                         EntryTax := EntryTax + PostedStrOrdrLineDetails.Amount;
                     UNTIL PostedStrOrdrLineDetails.NEXT = 0; */


                DiscountCharge := 0;
                /* PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "Sales Invoice Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'DISCOUNT');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                    DiscountCharge := PostedStrOrdrDetails."Calculation Value"; */

                /* PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Sales Invoice Header"."No.");
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type",
                                                  PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Group", 'ADDL.DUTY');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                    REPEAT
                        AddlDutyAmount := AddlDutyAmount + PostedStrOrdrLineDetails.Amount;
                    UNTIL PostedStrOrdrLineDetails.NEXT = 0; */

                //>>TaxDescription
                /*
                DetailedTaxEntry.RESET;
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.","Sales Invoice Header"."No.");
                IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                IF DetailedTaxEntry."Form Code" <> '' THEN
                BEGIN
                FormCode:='with Form '+DetailedTaxEntry."Form Code";
                END;
                TaxDescription:=FORMAT(DetailedTaxEntry."Tax Type")+FORMAT(DetailedTaxEntry."Tax %")+'%'+FormCode;
                END ELSE BEGIN
                TaxDescription:='NoTax Ag.Form'+"Sales Invoice Header"."Form Code";
                END;
                IF (recCustomer."Customer Posting Group" ='FOREIGN') AND ("Sales Invoice Header"."Form Code"='') THEN BEGIN
                TaxDescription:='NoTax Ag.Export';
                END;
                //<<
                */
                //TaxDescription


                /* DetailedTaxEntry.RESET;
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Sales Invoice Header"."No.");
                IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                    IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN
                        FormCode := 'with Form' + DetailedTaxEntry."Form Code";
                    END; */
                // EBT MILAN 030214 To Show TAX instand of VAT in case of UP State Invoice -------------------------START
                /* RecLocation2.RESET;
                RecLocation2.SETRANGE(RecLocation2.Code, "Sales Invoice Header"."Location Code");
                IF RecLocation2.FINDFIRST THEN BEGIN
                    IF (RecLocation2."State Code" = 'UP') AND ("Sales Invoice Header"."Posting Date" >= 011114D) THEN BEGIN
                        IF DetailedTaxEntry."Tax Type" = DetailedTaxEntry."Tax Type"::VAT THEN BEGIN
                            TaxDescription := 'TAX' + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                        END ELSE
                            TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                    END ELSE
                        TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                END;
                // EBT MILAN 030214 To Show TAX instand of VAT in case of UP State Invoice ---------------------------END
                // TaxDescription:=FORMAT(DetailedTaxEntry."Tax Type")+FORMAT(DetailedTaxEntry."Tax %")+'%'+FormCode;
            END;
*/

                /* DetailedTaxEntry.RESET;
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Sales Invoice Header"."No.");
                //DetailedTaxEntry.SETRANGE(DetailedTaxEntry."No.","Sales Invoice Line"."No.");
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Tax Area Code", "Tax Area Code");
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Tax Jurisdiction Code", "Tax Area Code");
                IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                    IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN
                        FormCode := 'with Form ' + DetailedTaxEntry."Form Code";
                    END; */
                // EBT MILAN 030214 To Show TAX instand of VAT in case of UP State Invoice -------------------------START
                /*  RecLocation2.RESET;
                 RecLocation2.SETRANGE(RecLocation2.Code, "Sales Invoice Header"."Location Code");
                 IF RecLocation2.FINDFIRST THEN BEGIN
                     IF (RecLocation2."State Code" = 'UP') AND ("Sales Invoice Header"."Posting Date" >= 011114D) THEN BEGIN
                         IF DetailedTaxEntry."Tax Type" = DetailedTaxEntry."Tax Type"::VAT THEN BEGIN
                             TaxDescription := 'TAX' + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                         END ELSE
                             TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                     END ELSE
                         TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                 END;
                 // EBT MILAN 030214 To Show TAX instand of VAT in case of UP State Invoice ---------------------------END

                 // TaxDescription:=FORMAT(DetailedTaxEntry."Tax Type")+FORMAT(DetailedTaxEntry."Tax %")+'%'+ FormCode;
             END; */


                /* DetailedTaxEntry.RESET;
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Sales Invoice Header"."No.");
                //DetailedTaxEntry.SETRANGE(DetailedTaxEntry."No.","Sales Invoice Line"."No.");
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Tax Area Code", "Tax Area Code");
                DetailedTaxEntry.SETFILTER(DetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Tax Area Code");
                IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                    "AddlTax/CessDesc" := 'Addl.Tax/Cess ' + FORMAT(DetailedTaxEntry."Tax %") + '%';
                END; */

                CLEAR("AddlTax/Cess");
                /* DetailedTaxEntry.RESET;
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Sales Invoice Header"."No.");
                DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Tax Area Code", "Tax Area Code");
                DetailedTaxEntry.SETFILTER(DetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Tax Area Code");
                IF DetailedTaxEntry.FINDFIRST THEN
                    REPEAT
                        "AddlTax/Cess" := "AddlTax/Cess" + DetailedTaxEntry."Tax Amount";
                    UNTIL DetailedTaxEntry.NEXT = 0; */



                // To Show thw Value of Transport Subsidy/Trade Discount
                SIL.RESET;
                SIL.SETRANGE(SIL."Document No.", "Sales Invoice Header"."No.");
                SIL.SETRANGE(SIL.Type, SIL.Type::"G/L Account");
                SIL.SETFILTER(SIL."No.", '%1', '750105');
                IF SIL.FINDFIRST THEN BEGIN
                    TransportSubsidy := SIL."Line Amount";
                END;

                //<<

                //>>BranchOfficeAddress

                Loc.GET("Location Code");

                IF NOT (Loc."State Code" = 'MAH') THEN BEGIN
                    IF RespCentre.GET(Loc."Global Dimension 2 Code") THEN BEGIN
                        RecState1.GET(Loc."State Code");
                        BranchOfficeAddress :=
                        RespCentre.Address + ' ' + RespCentre."Address 2" + ' ' + RespCentre.City + ' ' + RespCentre."Post Code"
                        + ' ' + RecState1.Description + ' Phone: ' + RespCentre."Phone No." + ' ' + ' Email: ' + RespCentre."E-Mail";
                        BranchOffText := 'Branch Office :';
                    END;
                END
                ELSE
                    BranchOfficeAddress := '';
                BranchOffText := '';

                //IF RecState.GET(State) THEN ;
                //IF RecCompInfo.GET THEN;
                //<<


                //AMount Details Total
                SIL.RESET;
                SIL.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETFILTER(SIL.Quantity, '<>%1', 0);
                //SIL.SETFILTER(SIL."Excise Prod. Posting Group", '<>%1', '');
                IF SIL.FINDSET THEN
                    REPEAT
                        //ItmQtyPer1 := ItmQtyPer1 + SIL."Qty. per Unit of Measure";
                        ItmQtyPer1 += SIL."Qty. per Unit of Measure";
                        //ItmQty1    := ItmQty1 + SIL.Quantity;
                        ItmQty1 += SIL.Quantity;
                        //ItmSheAmt1 := ItmSheAmt1 + SIL."SHE Cess Amount";
                        ItmSheAmt1 += 0;// SIL."SHE Cess Amount";
                        //ItmeCessAmt1 := ItmeCessAmt1 + SIL."eCess Amount";
                        ItmeCessAmt1 += 0;// SIL."eCess Amount";
                        //ItmExciseAmt1 := ItmExciseAmt1 + SIL."BED Amount";
                        ItmExciseAmt1 += 0;// SIL."BED Amount";
                        //ItmQtyBase1 := ItmQtyBase1 + SIL."Quantity (Base)";
                        ItmQtyBase1 += SIL."Quantity (Base)";
                        //AmtFinal := AmtFinal + SIL."Amount To Customer";
                        AmtFinal += 0;// SIL."Amount To Customer";
                        //ItmLineAmt1 := ItmLineAmt1 + SIL."Line Amount";
                        ItmLineAmt1 += SIL."Line Amount";
                        ItmLineDiscAmt1 += SIL."Line Discount Amount";

                    UNTIL SIL.NEXT = 0;

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
                    ApplicationArea = all;
                    Caption = 'No.Of Copies';
                }
                /* field('';'')
                {
                  ApplicationArea = all;
                  Caption = 'Show Internal Information';
                    Visible = false;
                } */
                field(LogInteraction; LogInteraction)
                {
                    ApplicationArea = all;
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

    trigger OnPostReport()
    var
        recSIH: Record 112;
    begin


        //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-----START
        IF "Sales Invoice Header"."Print Invoice" = FALSE THEN BEGIN
            IF NOT (CurrReport.PREVIEW) THEN BEGIN
                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Invoice Header"."No.");
                recSIH.SETRANGE(recSIH."Print Invoice", FALSE);
                IF recSIH.FINDFIRST THEN BEGIN
                    //RSPLSUM 04Feb2020>>
                    SalesSetup.GET;
                    IF SalesSetup."Email Alert on Sales Inv Post" THEN BEGIN
                        IF (recSIH."Print Invoice" = FALSE) AND (recSIH."Salesperson Email Sent" = FALSE) THEN BEGIN
                            SetSalespersonEmailValues(recSIH);
                            SalespersonEmailNotification;
                            recSIH."Salesperson Email Sent" := TRUE;
                        END;
                    END;
                    //RSPLSUM 04Feb2020<<
                    recSIH."Print Invoice" := TRUE;
                    recSIH.MODIFY;
                END;
            END;
        END;
        //EBT STIVAN---(29112012)--- To Print Invoice Only Once after First Print out of 5 pages are Taken-------END
    end;

    trigger OnPreReport()
    begin
        //CompanyInformation
        recCompany.GET();
    end;

    var
        recLocation: Record 14;
        LocState: Text[50];
        recCountry: Record 9;
        BillState: Text[50];
        BillCountry: Text[50];
        ShipState: Text[50];
        ShipCountry: Text[50];
        recCompany: Record 79;
        recState: Record State;
        // recECCNo: Record 13708;
        InvNo: Code[4];
        InvNoLen: Integer;
        recCustomer: Record 18;
        recCustomer1: Record 18;
        Ship2Name: Text[100];
        ShipToCST: Code[50];
        ShiptoTIN: Code[50];
        ShiptoLST: Code[50];
        ECCNo: Code[50];
        ShipToCode: Record 222;
        BatchNo: Code[20];
        recILE: Record 32;
        recVLE: Record 5802;
        recPayment: Record 3;
        recShipping: Record 291;
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
        "--noText": Integer;
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        LogInteraction: Boolean;
        SegManagement: Codeunit SegManagement;
        "--tax": Integer;
        // recExcisePostnSetup: Record 13711;
        SalesInvLine: Record 113;
        BedPer: Decimal;
        eCessPer: Decimal;
        SheCessPer: Decimal;
        RoundOffAmnt: Decimal;
        DutyAmtperunit: Decimal;
        RoundOffAcc: Code[20];
        //  PostedStrOrdrLineDetails: Record 13798;
        // PostedStrOrdrDetails: Record 13760;
        CustPostSetup: Record 92;
        AddlDutyAmount: Decimal;
        Freight: Decimal;
        EntryTax: Decimal;
        DiscountCharge: Decimal;
        //   DetailedTaxEntry: Record 16522;
        FormCode: Code[30];
        TaxDescription: Text[50];
        CurrCode: Code[10];
        AmtinWordExcise: array[2] of Text[100];
        AmtinWordeCess: array[2] of Text[100];
        AmtinWordSHe: array[2] of Text[100];
        AmtinWordTotal: array[2] of Text[100];
        Check: Report "Check Report";
        AmtFinal: Decimal;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        copytext: Text[30];
        i: Integer;
        InvType: Text[50];
        InvType1: Text[30];
        OutPutNo: Integer;
        SIL: Record 113;
        ItmQtyPer1: Decimal;
        ItmQty1: Decimal;
        ItmSheAmt1: Decimal;
        ItmeCessAmt1: Decimal;
        ItmExciseAmt1: Decimal;
        ItmQtyBase1: Decimal;
        ItmLineAmt1: Decimal;
        ItmLineDiscAmt1: Decimal;
        ItmDesc: Text[100];
        MRPPrice: Decimal;
        AssValue: Decimal;
        recMRPMaster: Record 50013;
        CCFCtext: Text[30];
        CCFC: Decimal;
        TPTSUBSIDYtext: Text[30];
        TPTSUBSIDY: Decimal;
        ProdDisc: Decimal;
        MRPMaster: Record 50013;
        billingprice: Decimal;
        recItem: Record 27;
        chargeamttotal: Decimal;
        // PostedStrOrdDetails: Record 13798;
        CashDiscountAmt: Decimal;
        // recPostedStrOrdDetails: Record 13798;
        CashDiscText: Text[30];
        FOCAMOUNT: Decimal;
        SalesInvCountPrinted: Codeunit "Sales Inv.-Printed";
        TaxInvoice: Text[35];
        RecLocation2: Record 14;
        Loc: Record 14;
        RespCentre: Record 5714;
        RecState1: Record State;
        BranchOfficeAddress: Text[1024];
        BranchOffText: Text[30];
        PriceSupport: Decimal;
        "AddlTax/CessDesc": Text[50];
        "AddlTax/Cess": Decimal;
        TransportSubsidy: Decimal;
        //recECC: Record 13708;
        CessText: Text[50];
        CessValue: Decimal;
        "Entry/Octroi": Decimal;
        ADDTAX: Decimal;
        ADDTAXText: Text[30];
        CCFCLineamt: Decimal;
        TaxAmount: Decimal;
        "--12Apr2017": Integer;
        NNN: Integer;
        DimensionName: Text[60];
        DimensionName2: Text[60];
        recDimSet: Record 480;
        SalesSetup: Record 311;
        GloInvoiceDate: Date;
        GloInvoiceNo: Code[30];
        GloInvoicevalue: Decimal;
        GloCustomerName: Text[100];
        GloVehicleno: Code[20];
        GloTransporteraName: Text[100];
        GloDriverName: Text[100];
        GloDriverMobileNo: Code[20];
        GloReciepientEmail: Text[80];
        UserSetupRec: Record 91;
        SalespersonPurchaser: Record 13;
        GloLocationName: Text[100];
        LocRec: Record 14;
        ShippingAgent: Record 291;
        SaleLines: Record 37;
        GloInvoiceQuantity: Decimal;
        GloExtDocNo: Text;
        RecSalesInvHdr: Record 112;
        Text1005: Label 'Email notification has been sent to salesperson.';
        GloOrderNo: Code[20];
        BalQtyToInv: Decimal;
    // SMTPMailSetup: Record 409;

    //   //[Scope('Internal')]
    procedure FormatNoText(var NoText: array[2] of Text[80]; No: Decimal; CurrencyCode: Code[10])
    var
        PrintExponent: Boolean;
        Ones: Integer;
        Tens: Integer;
        Hundreds: Integer;
        Exponent: Integer;
        NoTextIndex: Integer;
        Currency: Record 4;
        TensDec: Integer;
        OnesDec: Integer;
    begin
        CLEAR(NoText);
        NoTextIndex := 1;
        NoText[1] := '';

        IF No < 1 THEN
            AddToNoText(NoText, NoTextIndex, PrintExponent, Text026)
        ELSE BEGIN
            FOR Exponent := 4 DOWNTO 1 DO BEGIN
                PrintExponent := FALSE;
                IF No > 99999 THEN BEGIN
                    Ones := No DIV (POWER(100, Exponent - 1) * 10);
                    Hundreds := 0;
                END ELSE BEGIN
                    Ones := No DIV POWER(1000, Exponent - 1);
                    Hundreds := Ones DIV 100;
                END;
                Tens := (Ones MOD 100) DIV 10;
                Ones := Ones MOD 10;
                IF Hundreds > 0 THEN BEGIN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Hundreds]);
                    AddToNoText(NoText, NoTextIndex, PrintExponent, Text027);
                END;
                IF Tens >= 2 THEN BEGIN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, TensText[Tens]);
                    IF Ones > 0 THEN
                        AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Ones]);
                END ELSE
                    IF (Tens * 10 + Ones) > 0 THEN
                        AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Tens * 10 + Ones]);
                IF PrintExponent AND (Exponent > 1) THEN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, ExponentText[Exponent]);
                IF No > 99999 THEN
                    No := No - (Hundreds * 100 + Tens * 10 + Ones) * POWER(100, Exponent - 1) * 10
                ELSE
                    No := No - (Hundreds * 100 + Tens * 10 + Ones) * POWER(1000, Exponent - 1);
            END;
        END;

        IF CurrencyCode <> '' THEN BEGIN
            Currency.GET(CurrencyCode);
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + ''/* Currency."Currency Numeric Description"*/);
        END ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, 'Rupees');

        AddToNoText(NoText, NoTextIndex, PrintExponent, Text028);

        TensDec := ((No * 100) MOD 100) DIV 10;
        OnesDec := (No * 100) MOD 10;
        IF TensDec >= 2 THEN BEGIN
            AddToNoText(NoText, NoTextIndex, PrintExponent, TensText[TensDec]);
            IF OnesDec > 0 THEN
                AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[OnesDec]);
        END ELSE
            IF (TensDec * 10 + OnesDec) > 0 THEN
                AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[TensDec * 10 + OnesDec])
            ELSE
                AddToNoText(NoText, NoTextIndex, PrintExponent, Text026);
        IF (CurrencyCode <> '') THEN
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + ''/*Currency."Currency Decimal Description"*/ + ' ONLY')
        ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, 'Paisa Only');
    end;

    local procedure AddToNoText(var NoText: array[2] of Text[80]; var NoTextIndex: Integer; var PrintExponent: Boolean; AddText: Text[30])
    begin
        PrintExponent := TRUE;

        WHILE STRLEN(NoText[NoTextIndex] + ' ' + AddText) > MAXSTRLEN(NoText[1]) DO BEGIN
            NoTextIndex := NoTextIndex + 1;
            IF NoTextIndex > ARRAYLEN(NoText) THEN
                ERROR(Text029, AddText);
        END;

        NoText[NoTextIndex] := DELCHR(NoText[NoTextIndex] + ' ' + AddText, '<');
    end;

    // //[Scope('Internal')]
    procedure InitTextVariable()
    begin
        OnesText[1] := Text032;
        OnesText[2] := Text033;
        OnesText[3] := Text034;
        OnesText[4] := Text035;
        OnesText[5] := Text036;
        OnesText[6] := Text037;
        OnesText[7] := Text038;
        OnesText[8] := Text039;
        OnesText[9] := Text040;
        OnesText[10] := Text041;
        OnesText[11] := Text042;
        OnesText[12] := Text043;
        OnesText[13] := Text044;
        OnesText[14] := Text045;
        OnesText[15] := Text046;
        OnesText[16] := Text047;
        OnesText[17] := Text048;
        OnesText[18] := Text049;
        OnesText[19] := Text050;

        TensText[1] := '';
        TensText[2] := Text051;
        TensText[3] := Text052;
        TensText[4] := Text053;
        TensText[5] := Text054;
        TensText[6] := Text055;
        TensText[7] := Text056;
        TensText[8] := Text057;
        TensText[9] := Text058;

        ExponentText[1] := '';
        ExponentText[2] := Text059;
        ExponentText[3] := Text1280000;
        ExponentText[4] := Text1280001;
    end;

    //[Scope('Internal')]
    procedure InitLogInteraction()
    begin
        LogInteraction := SegManagement.FindInteractTmplCode(4) <> '';
    end;

    //[Scope('Internal')]
    procedure SalespersonEmailNotification()
    var
        // SMTPMail: Codeunit 400;
        EmailMsg: Codeunit "Email Message";
        EmailObj: Codeunit Email;
        RecipientType: Enum "Email Recipient Type";
        SAE18: Record 50009;
        SA18: Record 50008;
        SubjectText: Text;
        User18: Record 91;
        SenderName: Text;
        SenderEmail: Text;
        Text18: Text;
        Cust18: Record 18;
        OtAmt: Decimal;
        CrAmt: Decimal;
        ODAmt: Decimal;
        ReceiveEmail: Text;
        DimEnable: Boolean;
        Loc21: Record 14;
        SL21: Record 37;
        AppEmail: Text;
        RecSIL: Record 113;
        RecSL: Record 37;
    begin
        // SMTPMailSetup.GET;//RSPLSUM02Apr21

        SubjectText := '';
        SenderName := '';
        SenderEmail := '';
        ReceiveEmail := '';
        Text18 := '';
        //CLEAR(SMTPMail);
        IF UserSetupRec.GET(USERID) THEN;
        //UserSetupRec.TESTFIELD("E-Mail");

        SubjectText := 'Invoice - ' + "Sales Invoice Header"."No." + ' - ' + "Sales Invoice Header"."Sell-to Customer Name";//RSPLSUM 01Feb21

        //>>Email Body
        //RSPLSUM 01Feb21--SMTPMail.CreateMessage('',UserSetupRec."E-Mail",GloReciepientEmail,'','',TRUE);
        //RSPLSUM02Apr21--SMTPMail.CreateMessage('',UserSetupRec."E-Mail",GloReciepientEmail,SubjectText,'',TRUE);//RSPLSUM 01Feb21
        //SMTPMail.CreateMessage('', SMTPMailSetup."User ID", GloReciepientEmail, SubjectText, '', TRUE);//RSPLSUM02Apr21
        Emailmsg.Create(GloReciepientEmail, SubjectText, '', TRUE);
        //SMTPMail.AddRecipients(GloReciepientEmail);


        Emailmsg.AppendtoBody('<Br>');
        Emailmsg.AppendtoBody('<Br>');
        Emailmsg.AppendtoBody('<B> New invoice dispatched from </B>' + GloLocationName);
        Emailmsg.AppendtoBody('<Br>');
        Emailmsg.AppendtoBody('<Br>');

        //Table Start
        Emailmsg.AppendtoBody('<Table  Border="1">');
        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Invoice No </th>');
        Emailmsg.AppendtoBody('<td>' + GloInvoiceNo + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>InvoiceDate</th>');
        Emailmsg.AppendtoBody('<td>' + FORMAT(GloInvoiceDate) + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Customer Name</th>');
        Emailmsg.AppendtoBody('<td>' + GloCustomerName + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        RecSIL.RESET;
        RecSIL.SETRANGE("Document No.", GloInvoiceNo);
        RecSIL.CALCSUMS("Quantity (Base)");
        Emailmsg.AppendtoBody('<th>Quantity</th>');
        Emailmsg.AppendtoBody('<td>' + FORMAT(RecSIL."Quantity (Base)", 0, '<Integer Thousand><Decimals,3>') + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        //RSPLSUM 07Aug2020>>
        Emailmsg.AppendtoBody('<tr>');
        CLEAR(BalQtyToInv);
        RecSL.RESET;
        RecSL.SETRANGE("Document No.", GloOrderNo);
        RecSL.SETRANGE("Document Type", RecSL."Document Type"::Order);
        RecSL.SETRANGE(Type, RecSL.Type::Item);
        IF RecSL.FINDSET THEN
            REPEAT
                BalQtyToInv += RecSL."Quantity (Base)" - RecSL."Qty. Invoiced (Base)";
            UNTIL RecSL.NEXT = 0;

        Emailmsg.AppendtoBody('<th>Balance Quantity to Invoice</th>');
        Emailmsg.AppendtoBody('<td>' + FORMAT(BalQtyToInv, 0, '<Integer Thousand><Decimals,3>') + '</td>');
        Emailmsg.AppendtoBody('</tr>');
        //RSPLSUM 07Aug2020<<

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Invoice value </th>');
        Emailmsg.AppendtoBody('<td>' + FORMAT(GloInvoicevalue, 0, '<Integer Thousand><Decimals,3>') + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Purchase Order No </th>');
        Emailmsg.AppendtoBody('<td>' + GloExtDocNo + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Vehicle No </th>');
        Emailmsg.AppendtoBody('<td>' + GloVehicleno + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Transport Name </th>');
        Emailmsg.AppendtoBody('<td>' + GloTransporteraName + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Driver Name</th>');
        Emailmsg.AppendtoBody('<td>' + GloDriverName + '</td>');
        Emailmsg.AppendtoBody('</tr>');

        Emailmsg.AppendtoBody('<tr>');
        Emailmsg.AppendtoBody('<th>Driver Mobile No</th>');
        Emailmsg.AppendtoBody('<td>' + GloDriverMobileNo + '</td>');
        Emailmsg.AppendtoBody('</tr>');
        Emailmsg.AppendtoBody('</table>');
        //Table End

        Emailmsg.AppendtoBody('<Br>');
        Emailmsg.AppendtoBody('<Br>');

        EmailObj.Send(Emailmsg, Enum::"Email Scenario"::Default);
        //SMTPMail.Send;
        //MESSAGE(Text1005);
        //<<Email Body
    end;

    local procedure SetSalespersonEmailValues(SalesInvHeaderP: Record 112)
    var
        CustRec: Record 18;
    begin
        IF CustRec.GET(SalesInvHeaderP."Bill-to Customer No.") THEN;
        IF SalespersonPurchaser.GET(SalesInvHeaderP."Salesperson Code") THEN;
        IF LocRec.GET(SalesInvHeaderP."Location Code") THEN;
        IF ShippingAgent.GET(SalesInvHeaderP."Shipping Agent Code") THEN;
        // SalesInvHeaderP.CALCFIELDS("Amount to Customer");
        GloInvoiceDate := SalesInvHeaderP."Posting Date";
        GloInvoiceNo := SalesInvHeaderP."No.";
        GloInvoicevalue := 0;//SalesInvHeaderP."Amount to Customer";
        GloCustomerName := CustRec.Name;
        GloVehicleno := SalesInvHeaderP."Vehicle No.";
        GloTransporteraName := ShippingAgent.Name;
        GloDriverName := '';
        GloDriverMobileNo := SalesInvHeaderP."Driver's Mobile No.";
        GloReciepientEmail := SalespersonPurchaser."E-Mail";
        GloLocationName := LocRec.Name;
        //GloInvoiceQuantity := InvoiceQuantity;
        GloExtDocNo := SalesInvHeaderP."External Document No.";
        GloOrderNo := SalesInvHeaderP."Order No.";//RSPLSUM 07Aug2020
    end;
}


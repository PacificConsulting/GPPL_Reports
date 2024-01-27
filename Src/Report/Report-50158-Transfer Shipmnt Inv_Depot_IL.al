report 50158 "Transfer Shipmnt Inv_Depot_IL"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TransferShipmntInvDepotIL.rdl';
    Permissions = TableData 5744 = rim;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
        {
            column(Location; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' ' + LocFrom."Post Code" + ' ' + st2 + ' Phone: ' + LocFrom."Phone No." + ' ' + ' Email: ' + LocFrom."E-Mail")
            {
            }
            column(InvNo; InvNo)
            {
            }
            column(EccNos; '"E.C.C. Nos."."C.E. Registration No."')
            {
            }
            column(RangDivision; '')//"E.C.C. Nos."."C.E. Range" + '&' + "E.C.C. Nos."."C.E. Division" + '& ' + "E.C.C. Nos."."C.E. Commissionerate")
            {
            }
            column(CEAddr; '"E.C.C. Nos."."C.E Address1"')
            {
            }
            column(CEAddr2; '"E.C.C. Nos."."C.E Address2"')
            {
            }
            column(CECityPostCode; '')// "E.C.C. Nos."."C.E City" + '-' + "E.C.C. Nos."."C.E Post Code")
            {
            }
            column(TINNo; '')//Loc1."T.I.N. No." + '  &  ' + Loc1."C.S.T No.")
            {
            }
            column(PostingDate; "Posting Date")
            {
            }
            column(BRTTime; "Transfer Shipment Header"."BRT Print Time")
            {
            }
            column(No_TransferShipmentHeader; "Transfer Shipment Header"."No.")
            {
            }
            column(TransferOrderNo_TransferShipmentHeader; "Transfer Shipment Header"."Transfer Order No.")
            {
            }
            column(RoadPermitNo_TransferShipmentHeader; "Transfer Shipment Header"."Road Permit No.")
            {
            }
            column(Name_ShipingAgent; recShippingAgent.Name)
            {
            }
            column(VehicleNo_TransferShipmentHeader; "Transfer Shipment Header"."Vehicle No.")
            {
            }
            column(LRRRDate_TransferShipmentHeader; "Transfer Shipment Header"."LR/RR Date")
            {
            }
            column(LRRRNo_TransferShipmentHeader; "Transfer Shipment Header"."LR/RR No.")
            {
            }
            column(FrieghtType_TransferShipmentHeader; "Transfer Shipment Header"."Frieght Type")
            {
            }
            column(Loc_Contact; 'Contact ' + ':- ' + Loc."Phone No.")
            {
            }
            column(Loc_Add1; Loc.Address)
            {
            }
            column(Loc_Add2; Loc."Address 2")
            {
            }
            column(Loc_Add3; Loc.City + ' - ' + Loc."Post Code")
            {
            }
            column(State; RecState.Description + ', ' + RecCountry.Name)
            {
            }
            column(Loc_CST; '')// Loc."C.S.T No." + '  ' + FORMAT(Loc."W.e.f. Date(C.S.T No.)"))
            {
            }
            column(BuyerTIN; 'Loc."T.I.N. No."')
            {
            }
            column(Loc_WEF_CST; Loc."W.e.f. Date(C.S.T No.)")
            {
            }
            column(Buyer_WEF; Loc."W.e.f. Date(T.I.N No.)")
            {
            }
            column(BuyerECC; '"E.C.C. Nos.1"."C.E. Registration No."')
            {
            }
            column(Phone_Consignee; 'Contact ' + ':- ' + Loc."Phone No.")
            {
            }
            column(Loc_AddConsignee; Loc.Address)
            {
            }
            column(Loc_Add2_Consgnii; Loc."Address 2")
            {
            }
            column(City_Consignee; Loc.City + ' - ' + Loc."Post Code")
            {
            }
            column(State_Consginee; RecState.Description + ', ' + RecCountry.Name)
            {
            }
            column(CST_Con; 'Loc."C.S.T No."')
            {
            }
            column(TIN_Con; 'Loc."T.I.N. No."')
            {
            }
            column(WEF_Cons; Loc."W.e.f. Date(T.I.N No.)")
            {
            }
            column(CE_Cons; '"E.C.C. Nos.1"."C.E. Registration No."')
            {
            }
            column(DriversLicenseNo_TransferShipmentHeader; "Transfer Shipment Header"."Driver's License No.")
            {
            }
            column(Batch_UOM; "Batch/UOM")
            {
            }
            column(InsuranceNo_CompInfo; CompanyInfo1."Insurance No.")
            {
            }
            column(InsuranceProvoder_CompInfo; CompanyInfo1."Insurance Provider")
            {
            }
            column(Poilicy_ComInfo; FORMAT(CompanyInfo1."Policy Expiration Date"))
            {
            }
            column(TransferOrderDate_TransferShipmentHeader; "Transfer Shipment Header"."Transfer Order Date")
            {
            }
            column(BuyerTINDate; Loc."W.e.f. Date(T.I.N No.)")
            {
            }
            column(st2; st2)
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
                    column(outPutNos; OutPutNo)
                    {
                    }
                    dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Transfer Shipment Header";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Quantity = FILTER(<> 0));
                        //"Excise Prod. Posting Group" = FILTER(<> ''));
                        column(DimensionName; DimensionName)
                        {
                        }
                        column(LineNo_TransferShipmentLine; "Transfer Shipment Line"."Line No.")
                        {
                        }
                        column(DocumentNo_TransferShipmentLine; "Transfer Shipment Line"."Document No.")
                        {
                        }
                        column(BatchNo; BatchNo)
                        {
                        }
                        column(QtyperUnitofMeasure_TransferShipmentLine; "Transfer Shipment Line"."Qty. per Unit of Measure")
                        {
                        }
                        column(QuantityBase_TransferShipmentLine; "Transfer Shipment Line"."Quantity (Base)")
                        {
                        }
                        column(Quantity_TransferShipmentLine; "Transfer Shipment Line".Quantity)
                        {
                        }
                        column(BaseUnitOfMeasure; RecItem."Base Unit of Measure")
                        {
                        }
                        column(Unitprice_QtyPerUnit; "Transfer Shipment Line"."Unit Price" / "Transfer Shipment Line"."Qty. per Unit of Measure")
                        {
                        }
                        column(ExciseProdPostingGroup_TransferShipmentLine; '"Transfer Shipment Line"."Excise Prod. Posting Group"')
                        {
                        }
                        column(BEDpercent; BEDpercent)
                        {
                        }
                        column(BEDant_QtyBase; 0)//"Transfer Shipment Line"."BED Amount" / "Transfer Shipment Line"."Quantity (Base)")
                        {
                        }
                        column(BEDAmount_TransferShipmentLine; 0)// "Transfer Shipment Line"."BED Amount")
                        {
                        }
                        column(SHE_Cess; 0)//"Transfer Shipment Line"."SHE Cess Amount" + "Transfer Shipment Line"."eCess Amount")
                        {
                        }
                        column(Amount_TransferShipmentLine; "Transfer Shipment Line".Amount)
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
                        column(DescriptionLineAD_1; DescriptionLineAD[1])
                        {
                        }
                        column(DescriptionLineTot_1; DescriptionLineTot[1])
                        {
                        }
                        column(BEDAmount_TransferShipmentLine1; 0)// "Transfer Shipment Line"."BED Amount")
                        {
                        }
                        column(eCessAmount_TransferShipmentLine; 0)//"Transfer Shipment Line"."eCess Amount")
                        {
                        }
                        column(ADEAmount_TransferShipmentLine; 0)//"Transfer Shipment Line"."ADE Amount")
                        {
                        }
                        column(SHECessAmount_TransferShipmentLine; 0)// "Transfer Shipment Line"."SHE Cess Amount")
                        {
                        }
                        column(FrieghtAmount; FrieghtAmount)
                        {
                        }
                        column(RoundOffAmnt; RoundOffAmnt)
                        {
                        }
                        column(DocTotal; DocTotal)
                        {
                        }
                        column(Commente; recInvComment.Comment)
                        {
                        }
                        column(EcessPer; FORMAT(ecesspercent) + '%')
                        {
                        }
                        column(SheCessPer; FORMAT(SHEpercent) + '%')
                        {
                        }
                        column(FormCodeDesc; 'FormCodes.Description')
                        {
                        }
                        column(vCount; vCount)
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin
                            vCount := "Transfer Shipment Line".COUNT;
                            Ctr := Ctr + 1;

                            IF recItemEBT.GET("Transfer Shipment Line"."Item No.") THEN;

                            IF recItemEBT.GET("Transfer Shipment Line"."Item No.") THEN;
                            IF (recItemEBT."Item Category Code" = 'CAT02') OR (recItemEBT."Item Category Code" = 'CAT03') OR
                            (recItemEBT."Item Category Code" = 'CAT11') OR (recItemEBT."Item Category Code" = 'CAT12') OR
                            (recItemEBT."Item Category Code" = 'CAT13') THEN BEGIN
                                "Batch/UOM" := '   Batch        No.';
                            END ELSE
                                "Batch/UOM" := '  Unit of Measure';

                            //>>14Apr2017 DimsetEntry

                            DimensionName := '';
                            IF "Inventory Posting Group" = 'MERCH' THEN BEGIN

                                recDimSet.RESET;
                                recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                recDimSet.SETRANGE("Dimension Code", 'MERCH');
                                IF recDimSet.FINDFIRST THEN BEGIN

                                    recDimSet.CALCFIELDS("Dimension Value Name");
                                    DimensionName := recDimSet."Dimension Value Name";

                                END;
                                DimensionName2 := DimensionName;
                            END;

                            IF (DimensionName = '') AND ("Item Category Code" <> 'CAT01')
                            AND ("Item Category Code" <> 'CAT04') AND ("Item Category Code" <> 'CAT05')
                            AND ("Item Category Code" <> 'CAT06') AND ("Item Category Code" <> 'CAT07')
                            AND ("Item Category Code" <> 'CAT08') AND ("Item Category Code" <> 'CAT09')
                            AND ("Item Category Code" <> 'CAT10') THEN
                                DimensionName := 'IPOL ' + Description
                            //ELSE
                            //  DimensionName:=Description;
                            ELSE
                                IF DimensionName <> '' THEN
                                    DimensionName := DimensionName2
                                ELSE
                                    DimensionName := Description;


                            //<<14Apr2017 DimsetEntry

                            /*
                            //RSPL
                            DimensionName:='';
                            //Dim table not avialble
                            
                            IF "Inventory Posting Group"='MERCH' THEN BEGIN
                             recPosteddocDiemension.RESET;
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Table ID",5745);
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Document No.","Transfer Shipment Line"."Document No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Line No.","Transfer Shipment Line"."Line No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Dimension Code",'MERCH');
                              IF recPosteddocDiemension.FINDFIRST THEN;
                                 dimensionCode:=recPosteddocDiemension."Dimension Value Code";
                               //IF recDimensionValue.GET(dimensionCode) THEN
                               //  DimensionName:=recDimensionValue.Name;
                               recDimensionValue.RESET;
                               recDimensionValue.SETRANGE(recDimensionValue.Code,dimensionCode);
                               IF recDimensionValue.FINDFIRST THEN
                                  DimensionName:=recDimensionValue.Name;
                                  DimensionName2:=DimensionName;
                            END;
                            
                            //RSPL
                            
                            IF (DimensionName='') OR (recItemEBT."Item Category Code" = 'CAT02') OR (recItemEBT."Item Category Code" = 'CAT03') OR
                              (recItemEBT."Item Category Code" = 'CAT11') OR (recItemEBT."Item Category Code" = 'CAT12') OR
                                (recItemEBT."Item Category Code" = 'CAT13')THEN
                              IF (DimensionName='') AND ("Transfer Shipment Line"."Item Category Code"<>'CAT01')
                                AND ("Transfer Shipment Line"."Item Category Code"<>'CAT04') AND("Transfer Shipment Line"."Item Category Code"<>'CAT05')
                                AND ("Transfer Shipment Line"."Item Category Code"<>'CAT06') AND("Transfer Shipment Line"."Item Category Code"<>'CAT07')
                                AND ("Transfer Shipment Line"."Item Category Code"<>'CAT08') AND("Transfer Shipment Line"."Item Category Code"<>'CAT09')
                                AND ("Transfer Shipment Line"."Item Category Code"<>'CAT10') THEN
                            
                            BEGIN
                             recILE.SETRANGE("Document No.","Transfer Shipment Line"."Document No.");
                             recILE.SETRANGE("Document Line No.","Transfer Shipment Line"."Line No.");
                             recILE.SETRANGE("Item No.","Transfer Shipment Line"."Item No.");
                             IF recILE.FINDFIRST THEN
                              BatchNo := recILE."Lot No.";
                              DimensionName := 'IPOL ' +"Transfer Shipment Line".Description;
                              //ItemDescription := 'IPOL ' +"Transfer Shipment Line".Description;
                             END ELSE BEGIN
                                BatchNo := "Transfer Shipment Line"."Unit of Measure Code";
                                //DimensionName:="Transfer Shipment Line".Description;
                                IF DimensionName<>'' THEN
                                  DimensionName:=DimensionName2
                                ELSE
                                  DimensionName:= "Transfer Shipment Line".Description;
                            //  ItemDescription := "Transfer Shipment Line".Description;
                            END;
                            */
                            /* recExcisePostnSetup.RESET;
                            recExcisePostnSetup.SETRANGE(recExcisePostnSetup."Excise Bus. Posting Group", "Transfer Shipment Line"."Excise Bus. Posting Group");
                            recExcisePostnSetup.SETRANGE(recExcisePostnSetup."Excise Prod. Posting Group",
                            "Transfer Shipment Line"."Excise Prod. Posting Group");
                            recExcisePostnSetup.SETFILTER(recExcisePostnSetup."From Date", '<=%1', "Transfer Shipment Header"."Posting Date"); //RSPL008
                            IF recExcisePostnSetup.FINDLAST THEN BEGIN
                                ecesspercent := recExcisePostnSetup."eCess %";
                                BEDpercent := recExcisePostnSetup."BED %";
                                SHEpercent := recExcisePostnSetup."SHE Cess %";
                            END; */

                            /*
                            RecState2.SETRANGE(RecState2.Code,Loc."State Code");
                            IF RecState2.FINDFIRST THEN
                             State:=RecState2.Description;
                            */
                            IF RecItem.GET("Transfer Shipment Line"."Item No.") THEN;
                            TotalQty += "Transfer Shipment Line"."Quantity (Base)";
                            TotalBR += "Transfer Shipment Line"."Unit Price";
                            TotalValue := TotalValue + ("Transfer Shipment Line"."Quantity (Base)" * "Transfer Shipment Line"."Unit Price");


                            TaxDescr := 'Tax Descr';

                            DocTotal +=
                            "Transfer Shipment Line".Amount + 0;//"Transfer Shipment Line"."BED Amount" + "Transfer Shipment Line"."eCess Amount"
                            //+ "Transfer Shipment Line"."SHE Cess Amount" + FrieghtAmount + "Transfer Shipment Line"."ADE Amount" + RoundOffAmnt;
                            /*
                            DocTotal +=ROUND((
                            "Transfer Shipment Line".Amount + "Transfer Shipment Line"."BED Amount"+"Transfer Shipment Line"."eCess Amount"
                            +"Transfer Shipment Line"."SHE Cess Amount"+FrieghtAmount+"Transfer Shipment Line"."ADE Amount"+RoundOffAmnt),1.0,'=');
                            */

                            vBED += 0;//"BED Amount";//10May2017

                            /* FormCodes.SETRANGE(FormCodes.Code, "Transfer Shipment Header"."Form Code");
                            IF FormCodes.FINDFIRST THEN; */
                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineDuty, ROUND(vBED, 0.01), '');//10May2017
                            //vCheck.FormatNoText(DescriptionLineDuty,ROUND("BED Amount",0.01),'');
                            DescriptionLineDuty[1] := DELCHR(DescriptionLineDuty[1], '=', '*');
                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineeCess, ROUND(/*"eCess Amount"*/0, 0.01), '');
                            DescriptionLineeCess[1] := DELCHR(DescriptionLineeCess[1], '=', '*');
                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineSHeCess, ROUND(/*"SHE Cess Amount"*/0, 0.01), '');
                            DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1], '=', '*');
                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineSHeCess, ROUND(/*"SHE Cess Amount"*/0, 0.01), '');
                            DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1], '=', '*');
                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineTot, ROUND((DocTotal), 1.0), '');
                            DescriptionLineTot[1] := DELCHR(DescriptionLineTot[1], '=', '*');
                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineAD, ROUND(/*"Transfer Shipment Line"."ADE Amount"*/0, 0.01), '');
                            DescriptionLineAD[1] := DELCHR(DescriptionLineAD[1], '=', '*');

                        end;

                        trigger OnPreDataItem()
                        begin

                            DocTotal := 0; //08Jun2017 CAS-16972-S5Q0Y3
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin



                        i += 1;
                        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START
                        IF ("Transfer Shipment Header"."Print BRT" = FALSE) THEN BEGIN
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
                        //MESSAGE(USERID);
                        IF ("Transfer Shipment Header"."Print BRT" = TRUE) AND (USERID = 'sa') THEN BEGIN
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

                        IF ("Transfer Shipment Header"."Print BRT" = TRUE) AND (USERID <> 'sa') THEN BEGIN
                            IF i = 1 THEN
                                InvType := 'EXTRA COPY';
                            InvType1 := ' NOT FOR COMMERCIAL USE';
                        END;

                        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
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
                        IF "Transfer Shipment Header"."Print BRT" = TRUE THEN BEGIN
                            SETRANGE(Number, 1, 1);
                            // SETRANGE(Number,1,ABS(NoOfCopies));
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

                //>>06Apr2017
                IF LocFrom.GET("Transfer-from Code") THEN BEGIN
                    IF RecState.GET(LocFrom."State Code") THEN
                        st2 := RecState.Description;
                END;
                //<<06Apr2017


                //IF Loc.GET("Transfer-from Code") THEN;
                //RecState.GET(Loc."State Code");
                //st2 := RecState.Description;
                Loc1.GET("Transfer-to Code");
                RecState1.GET(Loc1."State Code");

                InvNoLen := STRLEN("Transfer Shipment Header"."No.");
                InvNo := COPYSTR("Transfer Shipment Header"."No.", (InvNoLen - 3), 4);

                //EBT STIVAN---(29012013)---Error Message POP UP,if ECC No of Transfer from Code is not Specified----START
                recLocation.RESET;
                recLocation.SETRANGE(recLocation.Code, "Transfer Shipment Header"."Transfer-from Code");
                IF recLocation.FINDFIRST THEN BEGIN
                    /* recECC.RESET;
                    recECC.SETRANGE(recECC.Code, recLocation."E.C.C. No.");
                    IF recECC.FINDFIRST THEN BEGIN
                        IF recECC.Description = '' THEN
                            ERROR('Your Location is not Excise Registered');
                    END; */
                END;
                //EBT STIVAN---(29012013)---Error Message POP UP,if ECC No of Transfer from Code is not Specified------END


                Loc.RESET;
                Loc.SETRANGE(Loc.Code, "Transfer-to Code");
                IF Loc.FINDFIRST THEN;
                // "E.C.C. Nos.1".GET(Loc."E.C.C. No.");

                RecState.RESET;
                RecState.GET(Loc."State Code");
                RecCountry.RESET;
                RecCountry.GET(Loc."Country/Region Code");



                Loc1.RESET;
                Loc1.SETRANGE(Loc1.Code, "Transfer-from Code");
                IF Loc1.FINDFIRST THEN;
                //"E.C.C. Nos.".GET(Loc1."E.C.C. No.");


                CurrDatetime := TIME;
                TotalQty := 0;

                CLEAR(FrieghtAmount);
                CLEAR(AdDutyAmount);
                /* PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "Transfer Shipment Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'FREIGHT');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                    FrieghtAmount := PostedStrOrdrDetails."Calculation Value"; */

                // EBT Start 10.08.10 start

                /* PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "Transfer Shipment Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails.Type, PostedStrOrdrDetails.Type::Transfer);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'ADDL.DUTY');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                    AdDutyAmount := PostedStrOrdrDetails."Calculation Value"; */

                // EBT Start 10.08.10 end

                //EBT STIVAN ---(23/04/2012)--- TO capture Comment from posted transfer Shipment ---------START
                recInvComment.RESET;
                recInvComment.SETRANGE(recInvComment."Document Type", recInvComment."Document Type"::"Posted Transfer Shipment");
                recInvComment.SETRANGE(recInvComment."No.", "Transfer Shipment Header"."No.");
                IF recInvComment.FINDFIRST THEN
                    //EBT STIVAN ---(23/04/2012)--- TO capture Comment from posted transfer Shipment -----------END

                    recShippingAgent.RESET;
                recShippingAgent.SETRANGE(recShippingAgent.Code, "Transfer Shipment Header"."Shipping Agent Code");
                IF recShippingAgent.FINDFIRST THEN;


                Ctr := 1;
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
                field(ShowInternalInfo; ShowInternalInfo)
                {
                    ApplicationArea = all;
                    Caption = 'Show Internal Information';
                }
                field(LogInteraction; LogInteraction)
                {
                    ApplicationArea = all;
                    Caption = 'Log Interaction';
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

        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START
        IF "Transfer Shipment Header"."Print BRT" = FALSE THEN BEGIN
            IF NOT (CurrReport.PREVIEW) THEN BEGIN
                recTSH.RESET;
                recTSH.SETRANGE(recTSH."No.", "Transfer Shipment Header"."No.");
                recTSH.SETRANGE(recTSH."Print BRT", FALSE);
                IF recTSH.FINDFIRST THEN BEGIN
                    recTSH."Print BRT" := TRUE;
                    recTSH.MODIFY;
                END;
            END;
        END;
        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
    end;

    trigger OnPreReport()
    begin

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");

        //InitTextVariable;
    end;

    var
        Loc: Record Location;
        Loc1: Record Location;
        SalesPerson: Record "Salesperson/Purchaser";
        RespCentre: Record "Responsibility Center";
        RecState: Record State;
        RecState1: Record State;
        RecCountry: Record 9;
        RecCountry1: Record 9;
        CompanyInfo1: Record 79;
        RecCust: Record 18;
        RecItem: Record 27;
        ShipToCode: Record 222;
        "Payment Terms": Record 3;
        "Shipment Method": Record 10;
        "Shipping Agent": Record 291;
        ILE: Record 32;
        //"E.C.C. Nos.": Record 13708;
        //"E.C.C. Nos.1": Record 13708;
        CustPostSetup: Record 92;
        CustName: Text[100];
        UOM: Code[10];
        Qty: Decimal;
        QtyPerUOM: Decimal;
        ShiptoName: Text[80];
        ShipToCST: Code[30];
        ShiptoTIN: Code[30];
        ShiptoLST: Code[30];
        ECCNo: Code[20];
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        DescriptionLineDuty: array[2] of Text[132];
        DescriptionLineeCess: array[2] of Text[132];
        DescriptionLineSHeCess: array[2] of Text[132];
        DescriptionLineTot: array[2] of Text[132];
        DescriptionLineAD: array[2] of Text[132];
        PackDescription: Text[30];
        SrNo: Integer;
        RoundOffAcc: Code[20];
        RoundOffAmnt: Decimal;
        TransShptNo: Code[5];
        TransShptNoLen: Integer;
        "<EBT>": Integer;
        TaxDescr: Code[10];
        ECCDesc: Code[20];
        PackDescriptionforILE: Text[30];
        TaxJurisd: Record 320;
        TaxDescription: Text[50];
        Freight: Decimal;
        Discount: Decimal;
        TotalAmttoCustomer: Decimal;
        InvTotal: Decimal;
        CurrDatetime: Time;
        State: Code[50];
        Ctr: Integer;
        SalesInvLine: Record 113;
        //SegManagement: Codeunit SegManagement;
        LogInteraction: Boolean;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        cust: Record 18;
        copytext: Text[30];
        SalesInvCountPrinted: Codeunit "Sales Inv.-Printed";
        ShowInternalInfo: Boolean;
        InvNoLen: Integer;
        InvNo: Code[10];
        DocTotal: Decimal;
        i: Integer;
        InvType: Text[50];
        InvType1: Text[50];
        //FormCodes: Record 13756;
        TransHdr: Record 5740;
        ItemVarParmResFinal: Record 50015;
        FrieghtAmount: Decimal;
        AdDutyAmount: Decimal;
        //PostedStrOrdrDetails: Record 13760;
        TotalQty: Decimal;
        BEDpercent: Decimal;
        //recExcisePostnSetup: Record 13711;
        BatchNo: Code[10];
        recILE: Record 32;
        ecesspercent: Decimal;
        SHEpercent: Decimal;
        TotalBR: Decimal;
        TotalValue: Decimal;
        recInvComment: Record 5748;
        "Batch/UOM": Text[30];
        recItemEBT: Record 27;
        ItemDescription: Text[100];
        recTSH: Record 5744;
        //recECC: Record 13708;
        recLocation: Record 14;
        recShippingAgent: Record 291;
        "---------------": Integer;
        recDimensionValue: Record 349;
        DimensionName: Text[100];
        dimensionCode: Code[20];
        DimensionName2: Text[100];
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
        vCheck: Report "Check Report";
        OutPutNo: Integer;
        vCount: Integer;
        RecState2: Record State;
        st2: Text[100];
        "---13Apr2017": Integer;
        LocFrom: Record 14;
        recDimSet: Record 480;
        vBED: Decimal;
}


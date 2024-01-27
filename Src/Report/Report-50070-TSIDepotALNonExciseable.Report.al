report 50070 "TSI_Depot_AL_Non-Exciseable"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TSIDepotALNonExciseable.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
        {
            RequestFilterFields = "No.";
            column(Location; Loc2.Address + ' ' + Loc2."Address 2" + ' ' + Loc2.City + ' ' + Loc2."Post Code" + ' ' + st2 + ' Phone: ' + Loc2."Phone No." + ' ' + ' Email: ' + Loc2."E-Mail")
            {
            }
            column(st2; st2)
            {
            }
            column(InvNo; InvNo)
            {
            }
            column(LocTINNo; 'Loc1."T.I.N. No."')
            {
            }
            column(LocCSTNo; 'Loc1."C.S.T No."')
            {
            }
            column(PostingDate; "Posting Date")
            {
            }
            column(BRTTime; "Transfer Shipment Header"."BRT Print Time")
            {
            }
            column(DocNo; "Transfer Shipment Header"."No.")
            {
            }
            column(TOrderNo; "Transfer Shipment Header"."Transfer Order No.")
            {
            }
            column(TOrderDate; "Transfer Shipment Header"."Transfer Order Date")
            {
            }
            column(RoadPermitNo; "Transfer Shipment Header"."Road Permit No.")
            {
            }
            column(ShippingAgentName; recShippingAgent.Name)
            {
            }
            column(VehicleNo; "Transfer Shipment Header"."Vehicle No.")
            {
            }
            column(LRNo; "Transfer Shipment Header"."LR/RR No.")
            {
            }
            column(LRDate; "Transfer Shipment Header"."LR/RR Date")
            {
            }
            column(FreightType; "Transfer Shipment Header"."Frieght Type")
            {
            }
            column(BuyAdd1; Loc.Address)
            {
            }
            column(BuyAdd2; Loc."Address 2")
            {
            }
            column(BuyAdd3; Loc.City + ' - ' + Loc."Post Code" + '  ' + RecState.Description + ', ' + RecCountry.Name)
            {
            }
            column(BuyCSTNo; 'Loc."C.S.T No."')
            {
            }
            column(BuyCSTDate; Loc."W.e.f. Date(C.S.T No.)")
            {
            }
            column(BuyTINNo; 'Loc."T.I.N. No."')
            {
            }
            column(BuyTINDate; Loc."W.e.f. Date(T.I.N No.)")
            {
            }
            column(ShipAdd1; "Transfer Shipment Header"."Transfer-to Address")
            {
            }
            column(ShipAdd2; "Transfer Shipment Header"."Transfer-to Address 2")
            {
            }
            column(ShipAdd3; "Transfer-to City" + ' - ' + "Transfer-to Post Code" + '  ' + RecState.Description + ', ' + RecCountry.Name)
            {
            }
            column(ECCRegNo; '"E.C.C. Nos.1"."C.E. Registration No."')
            {
            }
            column(EccNos; '"E.C.C. Nos."."C.E. Registration No."')
            {
            }
            column(RangDivision; '')// "E.C.C. Nos."."C.E. Range" + '&' + "E.C.C. Nos."."C.E. Division" + '& ' + "E.C.C. Nos."."C.E. Commissionerate")
            {
            }
            column(CEAddr; '"E.C.C. Nos."."C.E Address1"')
            {
            }
            column(CEAddr2; '"E.C.C. Nos."."C.E Address2"')
            {
            }
            column(CECityPostCode; '')//"E.C.C. Nos."."C.E City" + '-' + "E.C.C. Nos."."C.E Post Code")
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
            column(DriverLicNo; "Transfer Shipment Header"."Driver's License No.")
            {
            }
            column(Comments; recInvComment.Comment)
            {
            }
            column(RoundOffAmnt; RoundOffAmnt)
            {
            }
            column(FrieghtAmount; FrieghtAmount)
            {
            }
            column(AmtFinal; AmtFinal)
            {
            }
            dataitem("Copy Loop"; 2000000026)
            {
                DataItemTableView = SORTING(Number);
                dataitem(PageLoop; 2000000026)
                {
                    DataItemTableView = SORTING(Number);
                    MaxIteration = 1;
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
                        // Excise Prod. Posting Group=FILTER(<>''));
                        column(DimensionName; DimensionName)
                        {
                        }
                        column(LineNo; "Transfer Shipment Line"."Line No.")
                        {
                        }
                        column(LinDocNo; "Transfer Shipment Line"."Document No.")
                        {
                        }
                        column(BatchNo; BatchNo)
                        {
                        }
                        column(QtyUOM; "Transfer Shipment Line"."Qty. per Unit of Measure")
                        {
                        }
                        column(QtyBase; "Transfer Shipment Line"."Quantity (Base)")
                        {
                        }
                        column(Qty; "Transfer Shipment Line".Quantity)
                        {
                        }
                        column(ItmBaseUOM; RecItem."Base Unit of Measure")
                        {
                        }
                        column(Ctr; Ctr)
                        {
                        }
                        column(UOMCode; "Unit of Measure Code")
                        {
                        }
                        column(MRPofItem; MRPofItem)
                        {
                        }
                        column(TransferPrice; ("Unit Price" / "Qty. per Unit of Measure") + /*"Transfer Shipment Line"."Excise Amount"*/0 / "Transfer Shipment Line"."Quantity (Base)")
                        {
                        }
                        column(TransferValue; "Transfer Shipment Line".Amount + 0)//"Transfer Shipment Line"."BED Amount" + "Transfer Shipment Line"."eCess Amount" + "Transfer Shipment Line"."SHE Cess Amount")
                        {
                        }
                        column(vCount; vCount)
                        {
                        }
                        column(FormCDesc; 'FormCodes.Description')
                        {
                        }
                        column(AmtinWords; DescriptionLineTot[1])
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin

                            //>>1

                            Ctr := Ctr + 1; //24Apr2017

                            UOM := "Transfer Shipment Line"."Unit of Measure Code";

                            //pratyusha
                            recILE.RESET;
                            recILE.SETRANGE("Document No.", "Transfer Shipment Line"."Document No.");
                            recILE.SETRANGE(recILE."Item No.", "Transfer Shipment Line"."Item No.");
                            recILE.SETRANGE(recILE."Document Line No.", "Transfer Shipment Line"."Line No.");
                            IF recILE.FINDFIRST THEN
                                BatchNo := recILE."Lot No.";


                            /*  recExcisePostnSetUp.RESET;
                             recExcisePostnSetUp.SETRANGE(recExcisePostnSetUp."Excise Bus. Posting Group", "Transfer Shipment Line"."Excise Bus. Posting Group");
                             recExcisePostnSetUp.SETRANGE(recExcisePostnSetUp."Excise Prod. Posting Group",
                             "Transfer Shipment Line"."Excise Prod. Posting Group");
                             IF recExcisePostnSetUp.FINDFIRST THEN BEGIN */
                            ecesspercent := 0;// recExcisePostnSetUp."eCess %";
                            BEDpercent := 0;//recExcisePostnSetUp."BED %";
                            SHEcess := 0;// recExcisePostnSetUp."SHE Cess %";
                                         //  END;
                                         //pratyusha

                            CLEAR(MRPofItem);
                            recMRPmaster.RESET;
                            recMRPmaster.SETRANGE(recMRPmaster."Item No.", "Transfer Shipment Line"."Item No.");
                            recMRPmaster.SETRANGE(recMRPmaster."Lot No.", BatchNo);
                            IF recMRPmaster.FINDFIRST THEN BEGIN
                                MRPofItem := recMRPmaster."MRP Price" * recMRPmaster."Qty. per Unit of Measure";
                            END;

                            //RSPL for showing diemension name.
                            DimensionName := '';
                            IF "Inventory Posting Group" = 'MERCH' THEN BEGIN

                                //>>24Apr2017 DimSet
                                recDimSet.RESET;
                                recDimSet.SETRANGE("Dimension Set ID", "Transfer Shipment Line"."Dimension Set ID");
                                recDimSet.SETRANGE("Dimension Code", 'MERCH');
                                IF recDimSet.FINDFIRST THEN BEGIN

                                    recDimSet.CALCFIELDS("Dimension Value Name");
                                    DimensionName := recDimSet."Dimension Value Name";
                                END;
                                DimensionName2 := DimensionName;

                            END;

                            //<<24Apr Dimset
                            /*
                             recPosteddocDiemension.RESET;
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Table ID",5745);
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Document No.","Transfer Shipment Line"."Document No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Line No.","Transfer Shipment Line"."Line No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Dimension Code",'MERCH');
                              IF recPosteddocDiemension.FINDFIRST THEN
                                 dimensionCode:=recPosteddocDiemension."Dimension Value Code";
                               //IF recDimensionValue.GET(dimensionCode) THEN
                               //  DimensionName:=recDimensionValue.Name;
                               recDimensionValue.RESET;
                               recDimensionValue.SETRANGE(recDimensionValue.Code,dimensionCode);
                               IF recDimensionValue.FINDFIRST THEN
                                  DimensionName:=recDimensionValue.Name;
                                  DimensionName2:=DimensionName;
                            END;
                            *///Commented 24Apr2017
                              /*//RSPL
                              IF DimensionName=''THEN
                                DimensionName:='IPOL '+Description;
                              */

                            //RSPL For Category show only item description
                            IF (DimensionName = '') AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT01')
                            AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT04') AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT05')
                            AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT06') AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT07')
                            AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT08') AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT09')
                            AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT10') AND ("Transfer Shipment Line"."Item Category Code" <> 'CAT15') THEN
                                DimensionName := 'IPOL ' + Description
                            //ELSE
                            //  DimensionName:=Description;

                            ELSE
                                IF DimensionName <> '' THEN
                                    DimensionName := DimensionName2
                                ELSE
                                    DimensionName := Description;
                            //RSPL
                            //<<1

                            //>>2

                            //Transfer Shipment Line, Header (1) - OnPreSection()
                            //Ctr:=1;
                            //<<2

                            //>>3

                            //Transfer Shipment Line, Body (2) - OnPreSection()
                            Qty := Quantity;
                            QtyPerUOM := "Qty. per Unit of Measure";
                            RecItem.GET("Transfer Shipment Line"."Item No.");
                            TotalQty += "Transfer Shipment Line"."Quantity (Base)";
                            //<<3


                            //>>4

                            //Transfer Shipment Line, Footer (4) - OnPreSection()
                            TaxDescr := 'Tax Descr';
                            DocTotal := ROUND((
                            "Transfer Shipment Line".Amount + 0 +/*"Transfer Shipment Line"."BED Amount" + "Transfer Shipment Line"."eCess Amount"*/
                            0/*+ "Transfer Shipment Line"."SHE Cess Amount"*/ + FrieghtAmount + AddlDutyAmount + RoundOffAmnt), 1.0, '=');

                            // FormCodes.SETRANGE(FormCodes.Code, "Transfer Shipment Header"."Form Code");
                            // IF FormCodes.FINDFIRST THEN;
                            //<<4

                            //>>24Apr2017 AmtinWords
                            RepCheck.InitTextVariable;
                            RepCheck.FormatNoText(DescriptionLineTot, ROUND((AmtFinal + FrieghtAmount + RoundOffAmnt), 1.0), '');//26Apr2017
                            //RepCheck.FormatNoText(DescriptionLineTot,ROUND((AmtFinal+FrieghtAmount+RoundOffAmnt),0.01),'');
                            DescriptionLineTot[1] := DELCHR(DescriptionLineTot[1], '=', '*');
                            //<<24Apr2017 AmtinWords

                        end;

                        trigger OnPreDataItem()
                        begin

                            vCount := COUNT;//24Apr2017

                            Ctr := 0;//10May2017
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin

                        //>>1
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
                                    IF i = 4 THEN BEGIN
                                        InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                                        InvType1 := 'NOT FOR CENVAT';
                                    END ELSE
                                        IF i = 5 THEN BEGIN
                                            InvType := 'EXTRA COPY NOT FOR CENVAT';
                                            InvType1 := ''
                                        END ELSE
                                            IF i = 6 THEN BEGIN
                                                InvType := 'EXTRA COPY FOR BUYER ';
                                                InvType1 := 'NOT FOR CENVAT'
                                            END ELSE
                                                InvType := '';
                        //<<1
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    OutPutNo += 1;//24Apr2017
                    //>>1

                    IF Number > 1 THEN
                        copytext := Text003;
                    //CurrReport.PAGENO := 1;
                    //<<1
                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    NoOfCopies := 4;
                    NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
                    IF NoOfLoops <= 0 THEN
                        NoOfLoops := 1;
                    copytext := '';
                    SETRANGE(Number, 1, NoOfLoops);
                    //<<1
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                InvNoLen := STRLEN("Transfer Shipment Header"."No.");
                InvNo := COPYSTR("Transfer Shipment Header"."No.", (InvNoLen - 3), 4);

                CLEAR(FrieghtAmount);
                CLEAR(AddlDutyAmount);
                /*  PostedStrOrdrDetails.RESET;
                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "Transfer Shipment Header"."No.");
                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                 PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'FREIGHT');
                 IF PostedStrOrdrDetails.FINDFIRST THEN */
                FrieghtAmount := 0;//PostedStrOrdrDetails."Calculation Value";

                /* PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "Transfer Shipment Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::"Other Taxes");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'ADDL.DUTY');
                IF PostedStrOrdrDetails.FINDFIRST THEN */
                AddlDutyAmount := 0;// PostedStrOrdrDetails."Calculation Value";

                //EBT STIVAN ---(23/04/2012)--- TO capture Comment from posted transfer Shipment ---------START
                recInvComment.RESET;
                recInvComment.SETRANGE(recInvComment."Document Type", recInvComment."Document Type"::"Posted Transfer Shipment");
                recInvComment.SETRANGE(recInvComment."No.", "Transfer Shipment Header"."No.");
                IF recInvComment.FINDFIRST THEN
                    //EBT STIVAN ---(23/04/2012)--- TO capture Comment from posted transfer Shipment -----------END

                    recShippingAgent.RESET;
                recShippingAgent.SETRANGE(recShippingAgent.Code, "Transfer Shipment Header"."Shipping Agent Code");
                IF recShippingAgent.FINDFIRST THEN;
                //<<1

                //>>2

                //Transfer Shipment Header, Header (1) - OnPreSection()
                IF Loc2.GET("Transfer-from Code") THEN;
                IF RecState.GET(Loc2."State Code") THEN
                    st2 := RecState.Description;

                IF Loc3.GET("Transfer-to Code") THEN;
                IF RecState1.GET(Loc3."State Code") THEN;
                //<<2


                //>>3

                //Transfer Shipment Header, Header (3) - OnPreSection()
                Loc1.RESET;
                Loc1.GET("Transfer-from Code");
                // "E.C.C. Nos.".GET(Loc1."E.C.C. No.");

                Loc.GET("Transfer-to Code");
                //  "E.C.C. Nos.1".GET(Loc."E.C.C. No.");
                RecState.RESET;
                RecState.GET(Loc."State Code");
                RecCountry.RESET;
                RecCountry.GET(Loc."Country/Region Code");


                CurrDatetime := TIME;
                TotalQty := 0;
                //<<3

                //>>24Apr2017 AmtFinal

                recTSL.RESET;
                recTSL.SETRANGE("Document No.", "Transfer Shipment Header"."No.");
                recTSL.SETFILTER(Quantity, '<>%1', 0);
                //  recTSL.SETFILTER("Excise Prod. Posting Group", '<>%1', '');
                IF recTSL.FINDSET THEN
                    REPEAT

                        AmtFinal += recTSL.Amount + 0;// recTSL."BED Amount" + recTSL."eCess Amount" + recTSL."SHE Cess Amount";
                    UNTIL recTSL.NEXT = 0;

                //<<24Apr2017 AmtFinal
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("No of Copies"; NoOfCopies)
                {
                }
                field("Show Internal Information"; ShowInternalInfo)
                {
                }
                field("Log Interaction"; LogInteraction)
                {
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
        //>>1

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
        //<<1
    end;

    var
        Loc: Record 14;
        Loc1: Record 14;
        SalesPerson: Record 13;
        RespCentre: Record 5714;
        RecState: Record State;
        RecState1: Record "state";
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
        //"E.C.C. Nos.": Record "13708";
        // "E.C.C. Nos.1": Record "13708";
        CustPostSetup: Record 92;
        CustName: Text[100];
        UOM: Code[20];
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
        DescriptionLineAddlDuty: array[2] of Text[132];
        DescriptionLineTot: array[2] of Text[132];
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
        SegManagement: Codeunit SegManagement;
        LogInteraction: Boolean;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        cust: Record 18;
        copytext: Text[30];
        // SalesInvCountPrinted: Codeunit 315;
        ShowInternalInfo: Boolean;
        InvNoLen: Integer;
        InvNo: Code[10];
        DocTotal: Decimal;
        i: Integer;
        InvType: Text[50];
        InvType1: Text[50];
        // FormCodes: Record "13756";
        TransHdr: Record 5740;
        MRPofItem: Decimal;
        SalesPrice: Record 7002;
        ItemVarParmResFinal: Record 50000;
        IUM: Record 5404;
        // PostedStrOrdrDetails: Record "13760";
        FrieghtAmount: Decimal;
        AddlDutyAmount: Decimal;
        TotalQty: Decimal;
        ShiptoLoc: Record 14;
        BEDpercent: Decimal;
        // recExcisePostnSetUp: Record "13711";
        BatchNo: Code[10];
        recILE: Record 32;
        ecesspercent: Decimal;
        SHEcess: Decimal;
        recInvComment: Record 5748;
        recMRPmaster: Record 50013;
        recShippingAgent: Record 291;
        "---------------RSPL": Integer;
        recDimensionValue: Record 349;
        DimensionName: Text[100];
        dimensionCode: Code[20];
        DimensionName2: Text[100];
        "---24Apr2017": Integer;
        recDimSet: Record 480;
        Text003: Label 'Original for Buyer';
        OutPutNo: Integer;
        Loc2: Record 14;
        st2: Text[60];
        Loc3: Record 14;
        vCount: Integer;
        recTSL: Record 5745;
        AmtFinal: Decimal;
        RepCheck: Report "Check Report";
}


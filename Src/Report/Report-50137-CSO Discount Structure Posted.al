report 50137 "CSO Discount Structure Posted"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 06Feb2018   RB-N         New Column BaseAmount
    // 16Jan2019   RB-N         National Discount Scheme & Regional Discount Scheme
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CSODiscountStructurePosted.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'CSO Discount Structure Posted';

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.", "Posting Date", "Location Code";
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING(Type, "Inventory Posting Group", "No.")
                                    WHERE(Type = CONST(Item),
                                          "Inventory Posting Group" = FILTER('AUTOOILS' | 'REPSOL'));
                /* dataitem(DataItem1000000002; 13798)
                {
                    DataItemLink = "Invoice No." = FIELD("Document No."),
                                   "Item No." = FIELD("No."),
                                   "Line No." = FIELD("Line No.");
                    DataItemTableView = SORTING("Invoice No.,Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group")
                                        WHERE("Tax/Charge Type" = CONST(Charges));

                    trigger OnAfterGetRecord()
                    begin
                        //>>1

                        ItemRec.GET("Posted Str Order Line Details"."Item No.");
                        SpclFOC := 0;
                        SpclFOCPerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'FOC DIS');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            SpclFOCPerlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //SpclFOC += recStrOrdLineDet.Amount;//25Apr2017
                            SpclFOC := recStrOrdLineDet.Amount;
                        END;

                        StateDisc := 0;
                        StateDiscPerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'STATE DISC');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            //StateDisc += recStrOrdLineDet.Amount;//25Apr2017
                            StateDisc := recStrOrdLineDet.Amount;
                            StateDiscPerlt := recStrOrdLineDet.Amount / "Sales Invoice Line"."Quantity (Base)";
                        END;

                        Spot := 0;
                        SpotPerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'SPOT-DISC');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            SpotPerlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //Spot += recStrOrdLineDet.Amount;//25Apr2017
                            Spot := recStrOrdLineDet.Amount;
                        END;

                        AVPSanction := 0;
                        AVPSanctionPerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        //recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group",'AVP-SP-SAN');
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'NATSCHEME'); //15Nov2018
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            AVPSanctionPerlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //AVPSanction += recStrOrdLineDet.Amount;//25Apr2017
                            AVPSanction := recStrOrdLineDet.Amount;
                        END;

                        NewPrice := 0;
                        NewPricePerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'NEW-PRI-SP');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            NewPricePerlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //NewPrice += recStrOrdLineDet.Amount;//25Apr2017
                            NewPrice := recStrOrdLineDet.Amount;
                        END;

                        SpecDisc := 0;
                        SpecDiscPerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        //recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group",'T2 DISC');
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'REGSCHEME');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            SpecDiscPerlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //SpecDisc += recStrOrdLineDet.Amount;//25Apr2017
                            SpecDisc := recStrOrdLineDet.Amount;
                        END;

                        FOC := 0;
                        FOCperlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'FOC');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            FOCperlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //FOC += recStrOrdLineDet.Amount;//25Apr2017
                            FOC := recStrOrdLineDet.Amount;
                        END;


                        MntDisc := 0;
                        MntDiscPerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'MONTH-DIS');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            MntDiscPerlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //MntDisc += recStrOrdLineDet.Amount;//25Apr2017
                            MntDisc := recStrOrdLineDet.Amount;
                        END;

                        AddtlDisc := 0;
                        AddtlDiscPerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'ADD-DIS');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            AddtlDiscPerlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //AddtlDisc += recStrOrdLineDet.Amount;
                            AddtlDisc := recStrOrdLineDet.Amount;
                        END;

                        VATDiff := 0;
                        VATDiffPerlt := 0;
                        recStrOrdLineDet.RESET;
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Invoice No.", "Sales Invoice Line"."Document No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Line No.", "Sales Invoice Line"."Line No.");
                        recStrOrdLineDet.SETRANGE(recStrOrdLineDet."Tax/Charge Group", 'VAT-DIF-DI');
                        IF recStrOrdLineDet.FINDSET THEN BEGIN
                            VATDiffPerlt := recStrOrdLineDet."Calculation Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                            //VATDiff += recStrOrdLineDet.Amount;
                            VATDiff := recStrOrdLineDet.Amount;
                        END;

                        ProdDisc := 0;
                        ProdDiscPerlt := 0;
                        /*
                        MRpMstr.RESET;
                        MRpMstr.SETRANGE(MRpMstr."Item No.","Sales Invoice Line"."No.");
                        MRpMstr.SETRANGE(MRpMstr."Lot No.","Sales Invoice Line"."Lot No.");
                        IF MRpMstr.FINDFIRST THEN
                        BEGIN
                          ProdDiscPerlt := MRpMstr."National Discount";
                          ProdDisc := ProdDiscPerlt * "Sales Invoice Line"."Quantity (Base)";
                        END;
                        *///Code Commented
                          //>>02Jun2019
                          /*  ProdDiscPerlt := "Sales Invoice Line"."National Discount Per Ltr";
                           ProdDisc := "Sales Invoice Line"."National Discount Amount"; */
                          //<<02Jun2019

                //<<1

                //>>2
                //Structure Order Line Details, GroupFoote - OnPreSection()
                //TotalDisc:=0;
                //TotalDisc := SpclFOC+StateDisc+Spot+AVPSanction+NewPrice+SpecDisc+FOC+MntDisc+AddtlDisc+VATDiff;
                //TotalDiscPerlt := SpclFOCPerlt+StateDiscPerlt+SpotPerlt+AVPSanctionPerlt+NewPricePerlt+SpecDiscPerlt+FOCperlt+MntDiscPerlt
                //                +AddtlDiscPerlt+VATDiffPerlt;
                //MakeExcelDataGroupFooter;
                //<<2

                //>>25Apr2017 GroupData
                // TCOUNT += 1;

                /* STR25Apr.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                STR25Apr.SETRANGE("Line No.", "Line No.");
                IF STR25Apr.FINDLAST THEN BEGIN
                            ICOUNT := STR25Apr.COUNT;
                        END;

                        IF TCOUNT = ICOUNT THEN BEGIN

                            TotalDisc := SpclFOC + StateDisc + Spot + AVPSanction + NewPrice + SpecDisc + FOC + MntDisc + AddtlDisc + VATDiff;
                            TotalDiscPerlt := SpclFOCPerlt + StateDiscPerlt + SpotPerlt + AVPSanctionPerlt + NewPricePerlt + SpecDiscPerlt + FOCperlt + MntDiscPerlt
                                            + AddtlDiscPerlt + VATDiffPerlt;

                            MakeExcelDataGroupFooter;

                            TCOUNT := 0;
                            TotalDisc := 0;

                            TSpclFOC += SpclFOC;//25Apr2017
                            TStateDisc += StateDisc;//25Apr2017
                            TSpot += Spot;//25Apr2017
                            TAVPSanction += AVPSanction;//25Apr2017
                            TNewPrice += NewPrice;//25Apr2017
                            TSpecDisc += SpecDisc;//25Apr2017
                            TFOC += FOC;//25Apr2017
                            TMntDisc += MntDisc;//25Apr2017
                            TAddtlDisc += AddtlDisc;//25Apr2017
                            TVATDiff += VATDiff;//25Apr2017
                            TProdDisc += ProdDisc;//25Apr2017
                        END;
                        //<<25Apr2017 GroupData

                    end;

                    trigger OnPreDataItem()
                    begin

                        STR25Apr.COPYFILTERS("Posted Str Order Line Details");//25Apr2017

                        TCOUNT := 0; //25Apr2017
                    end;
                } */

                trigger OnAfterGetRecord()
                begin

                    //>>1
                    //ListPrice := "Sales Invoice Line"."Unit Price Incl. of Tax"/"Sales Invoice Line"."Qty. per Unit of Measure";
                    //<<1

                    //>>07Sep2017 ListPrice
                    CLEAR(ListPrice);
                    ListPrice := "Sales Invoice Line"."List Price";
                    //>>07Sep2017 ListPrice
                end;

                trigger OnPreDataItem()
                begin

                    //>>1
                    //CurrReport.CREATETOTALS(FOC,SpclFOC,Spot,NewPrice,AVPSanction,AddtlDisc,MntDisc,VATDiff,CD,OtherDisc);
                    //CurrReport.CREATETOTALS(TotalDisc,StateDisc,SpecDisc,ProdDisc,ProdDiscPerlt,ListPrice);
                    //<<1
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                TotalAmount := 0;
                // EBT MILAN...(22072013)........Added Case Discount and CCFC Detail.................START

                CashDiscountAmt := 0;
                /* recStructOrdrlnDetail.RESET;
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Invoice No.", "Sales Invoice Header"."No.");
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Tax/Charge Type", recStructOrdrlnDetail."Tax/Charge Type"::Charges);
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Tax/Charge Group", 'CASHDISC');
                IF recStructOrdrlnDetail.FINDFIRST THEN
                    REPEAT
                        CashDiscountAmt += recStructOrdrlnDetail.Amount;
                    UNTIL recStructOrdrlnDetail.NEXT = 0; */

                IF CashDiscountAmt <> 0 THEN
                    CashDiscText := 'Cash Discount @ ' + ''/*FORMAT(-recStructOrdrlnDetail."Calculation Value")*/ + ' ' + '%'
                ELSE
                    CashDiscText := '';

                CCFC := 0;
                /* recStructOrdrlnDetail.RESET;
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Invoice No.", "Sales Invoice Header"."No.");
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Tax/Charge Type", recStructOrdrlnDetail."Tax/Charge Type"::Charges);
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Tax/Charge Group", 'CCFC');
                IF recStructOrdrlnDetail.FINDFIRST THEN
                    REPEAT
                        CCFC += recStructOrdrlnDetail.Amount;
                    UNTIL recStructOrdrlnDetail.NEXT = 0; */

                IF CCFC <> 0 THEN
                    CCFCtext := 'Cash Discount @ ' + /*FORMAT(-recStructOrdrlnDetail."Calculation Value")*/ '' + ' ' + '%'
                ELSE
                    CCFCtext := '';

                //Sharath
                V_Transsub := 0;
                /* recStructOrdrlnDetail.RESET;
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Invoice No.", "Sales Invoice Header"."No.");
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Tax/Charge Type", recStructOrdrlnDetail."Tax/Charge Type"::Charges);
                recStructOrdrlnDetail.SETRANGE(recStructOrdrlnDetail."Tax/Charge Group", 'TRANSSUBS');
                IF recStructOrdrlnDetail.FINDFIRST THEN
                    REPEAT
                        V_Transsub += recStructOrdrlnDetail.Amount;
                    UNTIL recStructOrdrlnDetail.NEXT = 0; */

                IF V_Transsub <> 0 THEN
                    V_TranssubText := 'Transport Subsidy'
                ELSE
                    V_TranssubText := '';

                //Sharath


                // EBT MILAN...(22072013)........Added Case Discount and CCFC Detail.................END

                V_FocDis := '';
                /* Rec_SD.RESET;
                //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'FOC DIS');
                IF Rec_SD.FINDFIRST THEN
                    V_FocDis := Rec_SD."Tax/Charge Group Description"; */

                V_Statedesc := '';
                /*  Rec_SD.RESET;
                 //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                 Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'STATE DISC');
                 IF Rec_SD.FINDFIRST THEN
                     V_Statedesc := Rec_SD."Tax/Charge Group Description"; */

                V_Spot := '';
                /* Rec_SD.RESET;
                //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'SPOT-DISC');
                IF Rec_SD.FINDFIRST THEN
                    V_Spot := Rec_SD."Tax/Charge Group Description" */
                //ELSE
                V_Spot := 'Spot Discount';

                V_AVP := '';
                /* Rec_SD.RESET;
                //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                //Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group",'AVP-SP-SAN');
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'NATSCHEME'); //15Nov2018
                IF Rec_SD.FINDFIRST THEN
                    V_AVP := Rec_SD."Tax/Charge Group Description"; */

                V_NEWPRI := '';
                /* Rec_SD.RESET;
                //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'NEW-PRI-SP');
                IF Rec_SD.FINDFIRST THEN
                    V_NEWPRI := Rec_SD."Tax/Charge Group Description"; */

                V_T2D := '';
                /* Rec_SD.RESET;
                //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                //Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group",'T2 DISC');
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'REGSCHEME');//15Nov2018
                IF Rec_SD.FINDFIRST THEN
                    V_T2D := Rec_SD."Tax/Charge Group Description"; */

                V_Foc := '';
                /* Rec_SD.RESET;
                // Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'FOC');
                IF Rec_SD.FINDFIRST THEN
                    V_Foc := Rec_SD."Tax/Charge Group Description"; */

                V_Monthdis := '';
                /* Rec_SD.RESET;
                //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'MONTH-DIS');
                IF Rec_SD.FINDFIRST THEN
                    V_Monthdis := Rec_SD."Tax/Charge Group Description"; */

                V_Adddis := '';
                /* Rec_SD.RESET;
                //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'ADD-DIS');
                IF Rec_SD.FINDFIRST THEN
                    V_Adddis := Rec_SD."Tax/Charge Group Description";
 */
                V_Vatdifdi := '';
                /* Rec_SD.RESET;
                //Rec_SD.SETRANGE(Rec_SD.Code, "Sales Invoice Header".Structure);
                Rec_SD.SETRANGE(Rec_SD."Tax/Charge Group", 'VAT-DIF-DI');
                IF Rec_SD.FINDFIRST THEN
                    V_Vatdifdi := Rec_SD."Tax/Charge Group Description"; */

                //Sharath

                MakeExcelInfo;
                //<<1
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                //Sales Header, Footer (2) - OnPreSection()
                MakeExcelDataFooter;
                //<<1
            end;

            trigger OnPreDataItem()
            begin

                //>>1
                //CurrReport.CREATETOTALS(FOC,SpclFOC,Spot,NewPrice,AVPSanction,AddtlDisc,MntDisc,VATDiff,CD,OtherDisc);
                //CurrReport.CREATETOTALS(TotalDisc,StateDisc,SpecDisc,ProdDisc,ProdDiscPerlt,ListPrice);
                //<<1

                //>>25Apr2017 ReportHeader
                ExcelBuf.NewRow;
                ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17 06Feb2018
                /*
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//18
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//19
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//20
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//21
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//22
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//23
                *///Commented 21Aug2017
                ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//24


                ExcelBuf.NewRow;
                ExcelBuf.AddColumn('CSO Discount Structure Posted', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
                ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17 06Feb2018
                /*
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//18
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//19
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//20
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//21
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//22
                ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//23
                *///Commented 21Aug2017
                ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//24
                ExcelBuf.NewRow;
                ExcelBuf.NewRow;
                //<<25Apr2017 ReportHeader

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

    trigger OnPostReport()
    begin

        //>>1
        CreateExcelbook;
        //<<1
    end;

    var
        FOC: Decimal;
        SpclFOC: Decimal;
        Spot: Decimal;
        NewPrice: Decimal;
        AVPSanction: Decimal;
        AddtlDisc: Decimal;
        MntDisc: Decimal;
        VATDiff: Decimal;
        CD: Decimal;
        OtherDisc: Decimal;
        TotalDisc: Decimal;
        StateDisc: Decimal;
        SpecDisc: Decimal;
        TotalAmount: Decimal;
        ListPrice: Decimal;
        SpclFOCPerlt: Decimal;
        StateDiscPerlt: Decimal;
        SpotPerlt: Decimal;
        AVPSanctionPerlt: Decimal;
        NewPricePerlt: Decimal;
        CDPerlt: Decimal;
        SpecDiscPerlt: Decimal;
        FOCperlt: Decimal;
        MntDiscPerlt: Decimal;
        AddtlDiscPerlt: Decimal;
        VATDiffPerlt: Decimal;
        TotalDiscPerlt: Decimal;
        ItemRec: Record 27;
        //recStrOrdLineDet: Record 13798;
        ExcelBuf: Record 370 temporary;
        ProdDisc: Decimal;
        ProdDiscPerlt: Decimal;
        CashDiscText: Text[30];
        CashDiscountAmt: Decimal;
        CCFCtext: Text[30];
        CCFC: Decimal;
        //recStructOrdrlnDetail: Record 13798;
        MRpMstr: Record 50013;
        tottotaldiscPerlt: Decimal;
        Netpricetocustperltr: Decimal;
        V_Transsub: Decimal;
        V_TranssubText: Text[30];
        //Rec_SD: Record 13793;
        V_FocDis: Text[50];
        V_Statedesc: Text[50];
        V_Spot: Text[50];
        V_AVP: Text[50];
        V_NEWPRI: Text[50];
        V_T2D: Text[50];
        V_Foc: Text[50];
        V_Monthdis: Text[50];
        V_Adddis: Text[50];
        V_Vatdifdi: Text[50];
        "----------25Apr2017": Integer;
        TCOUNT: Integer;
        ICOUNT: Integer;
        //STR25Apr: Record 13798;
        TQty: Decimal;
        TQtyBase: Decimal;
        TListPrice: Decimal;
        TTDisc: Decimal;
        TSpclFOC: Decimal;
        TStateDisc: Decimal;
        TSpot: Decimal;
        TAVPSanction: Decimal;
        TNewPrice: Decimal;
        TSpecDisc: Decimal;
        TFOC: Decimal;
        TMntDisc: Decimal;
        TAddtlDisc: Decimal;
        TVATDiff: Decimal;
        TProdDisc: Decimal;
        "--------06Feb2018": Integer;
        TBaseAmt06: Decimal;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        /* ExcelBuf.CreateBook('', 'Auto Discount Details Posted');
        ExcelBuf.CreateBookAndOpenExcel('', 'Auto Discount Details Posted', '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook('Auto Discount Details Posted');
        ExcelBuf.WriteSheet('Auto Discount Details Posted', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('Auto Discount Details Posted', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
    end;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('FG Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Product Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Pack Size in Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Qty Billed', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Total Invoice Qty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('List Price Billed', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Base Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10 06Feb2018
        /*
        ExcelBuf.AddColumn('Spacial FOC Disc',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('State Discount',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Spot Discount',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('AVP Discount',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Price Support',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Special Mgmt Discount',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('FOC Discount',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Monthly Discount',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('Additional Discount',FALSE,'',TRUE,FALSE,TRUE,'@');
        ExcelBuf.AddColumn('VAT Diff Discount',FALSE,'',TRUE,FALSE,TRUE,'@');
        */
        //Sharath
        //ExcelBuf.AddColumn(V_FocDis,FALSE,'',TRUE,FALSE,TRUE,'@',1);//10 //Commented 21Aug2017
        //ExcelBuf.AddColumn(V_Statedesc,FALSE,'',TRUE,FALSE,TRUE,'@',1);//11 //Commented 21Aug2017
        //ExcelBuf.AddColumn(V_Spot,FALSE,'',TRUE,FALSE,TRUE,'@',1);//12 //Commented 21Aug2017
        ExcelBuf.AddColumn(V_AVP, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Price Support', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn(V_T2D, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        /*
        ExcelBuf.AddColumn('FOC Discount',FALSE,'',TRUE,FALSE,TRUE,'@',1);//16
        ExcelBuf.AddColumn(V_Monthdis,FALSE,'',TRUE,FALSE,TRUE,'@',1);//17
        ExcelBuf.AddColumn('Additional Discount',FALSE,'',TRUE,FALSE,TRUE,'@',1);//18
        ExcelBuf.AddColumn('VAT Diff Discount',FALSE,'',TRUE,FALSE,TRUE,'@',1);//19
        *///Commented 21Aug2017
        //Sharath
        ExcelBuf.AddColumn('Total Discount Per Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('Product Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('Total Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('Net Price Per Liter to Customer', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('Net Price Amount to Customer', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24

    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataGroupFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Sales Invoice Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn(/*"Posted Str Order Line Details"."Item No."*/'', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(ItemRec.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Line"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn("Sales Invoice Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//6
        ExcelBuf.AddColumn("Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//7
        TQty += "Sales Invoice Line".Quantity;//25Apr2017

        ExcelBuf.AddColumn("Sales Invoice Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//8
        TQtyBase += "Sales Invoice Line"."Quantity (Base)";//25Apr2017

        IF "Sales Invoice Header"."Posting Date" <= 20190303D THEN BEGIN //030319D
            ExcelBuf.AddColumn(ListPrice + ProdDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//9
            ListPrice := ListPrice + ProdDiscPerlt;//25Apr2017

            ExcelBuf.AddColumn("Sales Invoice Line"."Quantity (Base)" * (ListPrice + ProdDiscPerlt), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//10 06Feb2018
            TBaseAmt06 += "Sales Invoice Line"."Quantity (Base)" * (ListPrice + ProdDiscPerlt);//06Feb2018
        END ELSE BEGIN
            //>>02Jun2019
            ExcelBuf.AddColumn(ListPrice, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//9
                                                                                        //ListPrice += ListPrice;

            ExcelBuf.AddColumn("Sales Invoice Line"."Quantity (Base)" * (ListPrice), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//10
            TBaseAmt06 += "Sales Invoice Line"."Quantity (Base)" * (ListPrice);
            //<<02Jun2019
        END;
        //ExcelBuf.AddColumn(SpclFOCPerlt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//10 //Commented 21Aug2017
        //ExcelBuf.AddColumn(StateDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//11 //Commented 21Aug2017
        //ExcelBuf.AddColumn(SpotPerlt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//12 //Commented 21Aug2017
        ExcelBuf.AddColumn(AVPSanctionPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//13
        ExcelBuf.AddColumn(NewPricePerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//14
        ExcelBuf.AddColumn(SpecDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//15
        /*
        ExcelBuf.AddColumn(FOCperlt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//16
        ExcelBuf.AddColumn(MntDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//17
        ExcelBuf.AddColumn(AddtlDiscPerlt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//18
        ExcelBuf.AddColumn(VATDiffPerlt,FALSE,'',FALSE,FALSE,FALSE,'#,###0.00',0);//19
        *///Commented 21Aug2017
        ExcelBuf.AddColumn(ROUND(TotalDiscPerlt, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//20
        // EBT MILAN (22072013)......Added Product Discount..............START
        ExcelBuf.AddColumn(-ProdDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//21
                                                                                           // EBT MILAN (22072013)......Added Product Discount..............END

        ExcelBuf.AddColumn(ROUND(TotalDisc + ProdDisc, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//22
        TTDisc += TotalDisc + ProdDisc;//25Apr2017

        IF ProdDiscPerlt <> 0 THEN BEGIN
            IF "Sales Invoice Header"."Posting Date" <= 20190303D/*030319D*/ THEN BEGIN
                ExcelBuf.AddColumn(ROUND(((ListPrice + ProdDiscPerlt) - (ABS(TotalDiscPerlt + (-ProdDiscPerlt)))), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
                ExcelBuf.AddColumn(ROUND(((ListPrice + ProdDiscPerlt) - (ABS(TotalDiscPerlt + (-ProdDiscPerlt))))
                                        * "Sales Invoice Line"."Quantity (Base)", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24

                TotalAmount += (((ListPrice + ProdDiscPerlt) - (ABS(TotalDiscPerlt + (-ProdDiscPerlt)))) * "Sales Invoice Line"."Quantity (Base)");
                Netpricetocustperltr += ((ListPrice + ProdDiscPerlt) - (ABS(TotalDiscPerlt + (-ProdDiscPerlt))));
            END ELSE BEGIN
                //>>02Jun2019
                ExcelBuf.AddColumn(ListPrice - (ABS(TotalDiscPerlt + (-ProdDiscPerlt))), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
                ExcelBuf.AddColumn(((ListPrice) - (ABS(TotalDiscPerlt + (-ProdDiscPerlt))))
                                        * "Sales Invoice Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24

                TotalAmount += (((ListPrice) - (ABS(TotalDiscPerlt + (-ProdDiscPerlt)))) * "Sales Invoice Line"."Quantity (Base)");
                Netpricetocustperltr += ((ListPrice) - (ABS(TotalDiscPerlt + (-ProdDiscPerlt))));
                //<<02Jun2019
            END;
        END ELSE BEGIN
            ExcelBuf.AddColumn(ROUND((ListPrice + (TotalDiscPerlt)), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
            ExcelBuf.AddColumn(ROUND((ListPrice + (TotalDiscPerlt)) * "Sales Invoice Line"."Quantity (Base)", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24

            TotalAmount += (ListPrice + (TotalDiscPerlt)) * "Sales Invoice Line"."Quantity (Base)";
            Netpricetocustperltr += (ListPrice + (TotalDiscPerlt));
        END;

        tottotaldiscPerlt += TotalDiscPerlt;

    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//19
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//20
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//21
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//22
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//23
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//24
        *///Commented 21Aug2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//7
        //ExcelBuf.AddColumn("Sales Invoice Line".Quantity,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//7
        TQty := 0;//25Apr2017

        ExcelBuf.AddColumn(TQtyBase, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//8
        //ExcelBuf.AddColumn("Sales Invoice Line"."Quantity (Base)",FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//8
        TQtyBase := 0;//25Apr2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//9 //25Apr2017
        //ExcelBuf.AddColumn(TListPrice,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//9 //25Apr2017
        TListPrice := 0;//25Apr2017

        ExcelBuf.AddColumn(TBaseAmt06, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//10 06Feb2018
        TBaseAmt06 := 0;//06Feb2018
                        //ExcelBuf.AddColumn(TSpclFOC,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//10 //Commented 21Aug2017
                        //ExcelBuf.AddColumn(SpclFOC,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//10

        //ExcelBuf.AddColumn(TStateDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//11 //Commented 21Aug2017
        //ExcelBuf.AddColumn(StateDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//11


        //ExcelBuf.AddColumn(TSpot,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//12 //Commented 21Aug2017
        //ExcelBuf.AddColumn(Spot,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//12

        ExcelBuf.AddColumn(TAVPSanction, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//13
                                                                                       //ExcelBuf.AddColumn(AVPSanction,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//13

        ExcelBuf.AddColumn(TNewPrice, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//14
                                                                                    //ExcelBuf.AddColumn(NewPrice,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//14

        ExcelBuf.AddColumn(TSpecDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//15
                                                                                    //ExcelBuf.AddColumn(SpecDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//15

        //ExcelBuf.AddColumn(TFOC,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//16 //Commented 21Aug2017
        //ExcelBuf.AddColumn(FOC,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//16

        //ExcelBuf.AddColumn(TMntDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//17 //Commented 21Aug2017
        //ExcelBuf.AddColumn(MntDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//17

        //ExcelBuf.AddColumn(TAddtlDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//18//Commented 21Aug2017
        //ExcelBuf.AddColumn(AddtlDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//18

        //ExcelBuf.AddColumn(TVATDiff,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//19//Commented 21Aug2017
        //ExcelBuf.AddColumn(VATDiff,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//19

        ExcelBuf.AddColumn(TSpclFOC + TStateDisc + TSpot + TAVPSanction + TNewPrice + TSpecDisc + TFOC + TMntDisc
         + TAddtlDisc + TVATDiff, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//20
                                                                                //ExcelBuf.AddColumn(SpclFOC + StateDisc + Spot+ AVPSanction + NewPrice + SpecDisc + FOC + MntDisc
                                                                                //+ AddtlDisc + VATDiff,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//20

        //EBT MILAN
        ExcelBuf.AddColumn(-TProdDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//21
                                                                                     //ExcelBuf.AddColumn(-ProdDisc,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//21
                                                                                     //EBT MILAN

        ExcelBuf.AddColumn(TTDisc, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//22 //25Apr2017
        //ExcelBuf.AddColumn(ROUND(TotalDisc-(-ProdDisc),0.01),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//22
        TTDisc := 0; //25Apr2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//23
                                                                             //ExcelBuf.AddColumn(ROUND(Netpricetocustperltr,0.01),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//23

        ExcelBuf.AddColumn(ROUND(TotalAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//24
        // EBT MILAN (22072013)..............ADDED CASE Discount and CCFC Detail.....................END
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16 06Feb2018
        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//17
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//18
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//19
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//20
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//21
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//22
        *///Commented 21Aug2017
        ExcelBuf.AddColumn(CashDiscText, FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn(ABS(CashDiscountAmt), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//24
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//17
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//18
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//19
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//20
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//21
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//22
        *///Commented 21Aug2017
        ExcelBuf.AddColumn(CCFCtext, FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn(ABS(CCFC), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//24
        //Sharath
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//17
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//18
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//19
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//20
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//21
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//22
        *///Commented 21Aug2017
        ExcelBuf.AddColumn(V_TranssubText, FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn(ABS(V_Transsub), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//24
        //Sharath
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//17
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//18
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//19
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//20
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//21
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'',1);//22
        *///Commented 21Aug2017
        ExcelBuf.AddColumn('Net Discount Amount', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn(TotalAmount + CashDiscountAmt + CCFC + V_Transsub, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//24
        // EBT MILAN (22072013)..............ADDED CASE Discount and CCFC Detail.....................END

    end;
}


report 50091 "FG Valutn (After Revaluation)"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 21Sep2017   RB-N         New Field BRT No. from TransferShipmentHeader
    // 07Oct2017   RB-N         Applies to Doc No. for Retrun Shipment
    // 09Nov2017   RB-N         New Field (Transfer,Excise & GST related Field)
    // 14Nov2017   RB-N         Assessable Value
    // 16Nov2017   RB-N         Skipping Negative Remaining Qty
    // 04Sep2018   RB-N         Allowing only FG ItemCode for REPSOL--CAT15
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/FGValutnAfterRevaluation.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            DataItemTableView = SORTING("Location Code", "Document Type", "Posting Date", "Item Category Code", "Item No.")
                                WHERE("Document Type" = FILTER('Purchase Receipt|Transfer Shipment|Transfer Receipt|Sales Return Receipt' | ' '),
                                      "Item Category Code" = FILTER('CAT02|CAT03|CAT11|CAT12|CAT13|CAT15|CAT18|CAT19|CAT20|CAT22'));
            RequestFilterFields = "Location Code", "Item Category Code", "Item No.";

            trigger OnAfterGetRecord()
            begin

                //>>04Sep2018
                RepItmNo := '';
                IF "Item Category Code" = 'CAT15' THEN BEGIN
                    RepItmNo := COPYSTR("Item No.", 1, 4);
                    IF (RepItmNo <> 'FGRP') AND (RepItmNo <> 'FGRE') THEN
                        CurrReport.SKIP;
                END;
                //<<04Sep2018

                //>>05Sep2018
                IF "Document Type" = "Document Type"::"Transfer Shipment" THEN
                    IF "Location Code" <> 'IN-TRANS' THEN
                        CurrReport.SKIP;

                //<<05Sep2018

                //>>09Mar2017 LocName1
                CLEAR(LocName1);
                recLOC.RESET;
                recLOC.SETRANGE(Code, "Item Ledger Entry"."Location Code");
                IF recLOC.FINDFIRST THEN BEGIN
                    LocName1 := recLOC.Name;
                END;
                //<< LocName1


                //>>RB-N 05Sep2018 Skipping Closed Location
                recLOC.RESET;
                recLOC.SETRANGE(Code, "Location Code");
                recLOC.SETRANGE("Include in Valuation Report", FALSE);
                IF recLOC.FINDFIRST THEN
                    CurrReport.SKIP;

                recLOC.RESET;
                recLOC.SETRANGE(Code, "Location Code");
                recLOC.SETRANGE(Closed, TRUE);
                IF recLOC.FINDFIRST THEN
                    CurrReport.SKIP;
                //<<RB-N 05Sep2018 Skipping Closed Location
                //>>1
                IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::" " THEN
                    IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Transfer THEN
                        IF "Item Ledger Entry"."Document No." = 'BRT/43/AL/1314/0003' THEN
                            CurrReport.SKIP;

                ItemRec.GET("Item Ledger Entry"."Item No.");
                VendorName := '';
                IF VendorRec.GET("Item Ledger Entry"."Source No.") THEN
                    VendorName := VendorRec.Name;

                RemQty := 0;
                ItemApplication.RESET;
                ItemApplication.SETRANGE(ItemApplication."Inbound Item Entry No.", "Item Ledger Entry"."Entry No.");
                ItemApplication.SETRANGE(ItemApplication."Posting Date", 0D, PostingDate);
                IF ItemApplication.FINDSET THEN
                    REPEAT
                        RemQty += ItemApplication.Quantity;
                    UNTIL ItemApplication.NEXT = 0;

                /*IF RemQty = 0 THEN
                   IF NOT Showall THEN
                      CurrReport.SKIP;*/

                //>>RB-N 16Nov2017 Skipping Negative Remaining Qty
                IF RemQty < 0 THEN
                    CurrReport.SKIP;

                //>>RB-N 16Nov2017 Skipping Negative Remaining Qty

                CLEAR(DenFactor);
                RateofItem := 0;
                TotalCost1 := 0;
                TotalQty1 := 0;
                ChargeofItem := 0;
                UNITCOST1 := 0;
                ValueEntry.RESET;
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Charge No.", '');
                ValueEntry.SETFILTER(ValueEntry."Source Code", '%1', 'REVALJNL');
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        TotalCost1 += ValueEntry."Cost Amount (Expected)" + ValueEntry."Cost Amount (Actual)";
                        TotalQty1 := ValueEntry."Valued Quantity";
                        UNITCOST1 := TotalCost1 / TotalQty1;
                    UNTIL ValueEntry.NEXT = 0;


                CLEAR(TotalCost);
                ValueEntry.RESET;
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETRANGE(ValueEntry."Item Charge No.", '');
                ValueEntry.SETFILTER(ValueEntry."Source Code", '<>%1', 'REVALJNL');
                IF ValueEntry.FINDSET THEN
                    REPEAT
                        TotalCost += ValueEntry."Cost Amount (Expected)" + ValueEntry."Cost Amount (Actual)";
                        TotalQty := ValueEntry."Valued Quantity";
                        UNITCOST := TotalCost / TotalQty;
                    UNTIL ValueEntry.NEXT = 0;

                UNITCOSTNEW := UNITCOST1 + ABS(UNITCOST);

                //Code Commented 11Dec2018


                //>>New Code 11Dec2018
                CLEAR(DenFactor);
                ILE11.RESET;
                ILE11.SETCURRENTKEY("Item No.");
                ILE11.SETRANGE("Item No.", "Item No.");
                ILE11.SETRANGE("Lot No.", "Lot No.");
                ILE11.SETRANGE("Entry Type", ILE11."Entry Type"::Output);
                ILE11.SETRANGE("Posting Date", 0D, PostingDate);
                IF ILE11.FINDLAST THEN BEGIN
                    DenFactor := ILE11."Density Factor";
                    /*
                    VE11.RESET;
                    VE11.SETCURRENTKEY("Item Ledger Entry No.","Entry Type");
                    VE11.SETRANGE("Item Ledger Entry No.",ILE11."Entry No.");
                    VE11.SETRANGE("Item Ledger Entry Type",VE11."Item Ledger Entry Type"::Output);
                    VE11.SETRANGE("Item No.",ILE11."Item No.");
                    IF VE11.FINDFIRST THEN
                    BEGIN
                      VE11.CALCSUMS("Cost per Unit");
                      UNITCOST :=  ABS(VE11."Cost per Unit");
                    END;
                    */
                END;

                IF DenFactor = 0 THEN
                    DenFactor := 1;

                IF DenFactor <> 1 THEN
                    UNITCOST := UNITCOST / DenFactor;
                //<<New Code 11Dec2018

                CLEAR(VLECOST);
                CALCFIELDS("Item Ledger Entry"."Cost Amount (Actual)");
                //CALCFIELDS("Item Ledger Entry"."Cost Amount (Expected)");     //Expected ADDED MILAN 030114 Removed on 130114
                VLECOST := "Item Ledger Entry"."Cost Amount (Actual)" / "Item Ledger Entry".Quantity;
                // ADDED MILAN 030114 Removed on 130114

                ValueEntry.RESET;
                ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
                ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '<>%1', '');
                IF ValueEntry.FINDFIRST THEN
                    REPEAT
                        ChargeofItem += ValueEntry."Cost per Unit";
                    UNTIL ValueEntry.NEXT = 0;
                IF ChargeofItem < 0 THEN BEGIN
                    RateofItem := RateofItem + ChargeofItem;
                    ChargeofItem := 0;
                END;

                //
                DocDate := 0D;
                DocDate := "Item Ledger Entry"."Posting Date";
                //
                //IF "Item Ledger Entry"."Document No." = 'TRR/18/1819/0138' THEN
                //MESSAGE('test');

                CLEAR(BlenOrdrNo);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", "Item Ledger Entry"."Item No.");
                RecILE.SETRANGE(RecILE."Lot No.", "Item Ledger Entry"."Lot No.");
                RecILE.SETRANGE(RecILE."Posting Date", 0D, PostingDate);
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Output);
                IF RecILE.FIND('-') THEN BEGIN
                    BlenOrdrNo := RecILE."Document No.";
                    //  DocDate := RecILE."Posting Date";
                END;

                IF BlenOrdrNo = '' THEN BEGIN
                    RecILE.RESET;
                    RecILE.SETRANGE(RecILE."Item No.", "Item Ledger Entry"."Item No.");
                    RecILE.SETRANGE(RecILE."Lot No.", "Item Ledger Entry"."Lot No.");
                    RecILE.SETRANGE(RecILE."Posting Date", 0D, PostingDate);
                    RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::"Positive Adjmt.");
                    IF RecILE.FIND('-') THEN BEGIN
                        BlenOrdrNo := 'OPENING ADJUSTMENT';
                        DocDate := RecILE."Posting Date";
                    END;
                END;


                //
                IF BlenOrdrNo = '' THEN BEGIN
                    RecILE.RESET;
                    RecILE.SETRANGE(RecILE."Item No.", "Item Ledger Entry"."Item No.");
                    RecILE.SETRANGE(RecILE."Lot No.", "Item Ledger Entry"."Lot No.");
                    RecILE.SETRANGE(RecILE."Posting Date", 0D, PostingDate);
                    IF RecILE.FIND('-') THEN BEGIN
                        BlenOrdrNo := 'POSITIVE ADJUSTMENT';
                        DocDate := RecILE."Posting Date";
                    END;
                END;

                //

                //RB-N 21Sep2017 To Capture BRT No. from TransferShipmentHeader
                TSH21.RESET;
                TSH21.SETRANGE("Transfer Order No.", "Order No.");
                IF TSH21.FINDFIRST THEN
                    BRTNo := TSH21."No."
                ELSE
                    BRTNo := '';
                //RB-N 21Sep2017 To Capture BRT No. from TransferShipmentHeader

                //RB-N 07Oct2017 To Applies to Doc No. from ReturnShipmentHeader
                RRH07.RESET;
                RRH07.SETRANGE("No.", "Item Ledger Entry"."Document No.");
                IF RRH07.FINDFIRST THEN
                    BRTNo := RRH07."Applies-to Doc. No.";
                //RB-N 07Oct2017 To Applies to Doc No. from ReturnShipmentHeader


                //RB-N 09Nov2017 Transfer Details
                TRPrice09 := 0;
                BedRate09 := 0;
                CGSTRate10 := 0;
                SGSTRate10 := 0;
                IGSTRate10 := 0;
                AssVal14 := 0;
                IF "Document Type" = "Document Type"::"Transfer Receipt" THEN BEGIN
                    //MESSAGE('TRL found %1  \\ %2 \\ %3',"Document Type","Document No.","Item No.");

                    TRL09.RESET;
                    TRL09.SETRANGE("Document No.", "Document No.");
                    TRL09.SETRANGE("Line No.", "Document Line No.");
                    TRL09.SETRANGE("Item No.", "Item No.");
                    IF TRL09.FINDFIRST THEN BEGIN

                        //TRPrice09 := TRL09."Unit Price" / TRL09."Quantity (Base)";
                        TRPrice09 := TRL09."Unit Price" / TRL09."Qty. per Unit of Measure";//14Nov2017
                        BedRate09 := 0;// TRL09."BED Amount" / TRL09."Quantity (Base)";
                        AssVal14 := 0;//TRL09."Assessable Value" / TRL09."Qty. per Unit of Measure";//14Nov2017
                                      //MESSAGE('TRL found');

                    END;

                    //GST Details
                    TSL10.RESET;
                    TSL10.SETRANGE("Document No.", BRTNo);
                    TSL10.SETRANGE("Item No.", "Item No.");
                    TSL10.SETFILTER(Quantity, '<>%1', 0);
                    IF TSL10.FINDFIRST THEN BEGIN
                        DGST10.RESET;
                        DGST10.SETRANGE(Type, DGST10.Type::Item);
                        DGST10.SETRANGE("Document No.", TSL10."Document No.");
                        DGST10.SETRANGE("No.", TSL10."Item No.");
                        DGST10.SETRANGE("Document Line No.", TSL10."Line No.");
                        IF DGST10.FINDSET THEN
                            REPEAT

                                IF DGST10."GST Component Code" = 'CGST' THEN
                                    CGSTRate10 := ABS(DGST10."GST Amount") / TSL10."Quantity (Base)";

                                IF DGST10."GST Component Code" = 'SGST' THEN
                                    SGSTRate10 := ABS(DGST10."GST Amount") / TSL10."Quantity (Base)";

                                IF DGST10."GST Component Code" = 'IGST' THEN
                                    IGSTRate10 := ABS(DGST10."GST Amount") / TSL10."Quantity (Base)";

                                IF DGST10."GST Component Code" = 'UTGST' THEN
                                    SGSTRate10 := ABS(DGST10."GST Amount") / TSL10."Quantity (Base)";

                            UNTIL DGST10.NEXT = 0;
                    END;
                END;

                IF "Document Type" = "Document Type"::"Sales Return Receipt" THEN BEGIN

                    ILE10.RESET;
                    ILE10.SETCURRENTKEY("Item No.");
                    ILE10.SETRANGE("Item No.", "Item No.");
                    ILE10.SETRANGE("Lot No.", "Lot No.");
                    ILE10.SETRANGE("Location Code", "Location Code");
                    ILE10.SETRANGE("Document Type", ILE10."Document Type"::"Transfer Receipt");
                    IF ILE10.FINDLAST THEN BEGIN

                        TRL09.RESET;
                        TRL09.SETRANGE("Document No.", ILE10."Document No.");
                        TRL09.SETRANGE("Line No.", ILE10."Document Line No.");
                        TRL09.SETRANGE("Item No.", ILE10."Item No.");
                        IF TRL09.FINDFIRST THEN BEGIN

                            //TRPrice09 := TRL09."Unit Price" / TRL09."Quantity (Base)";
                            TRPrice09 := TRL09."Unit Price" / TRL09."Qty. per Unit of Measure";//14Nov2017
                            BedRate09 := 0;// TRL09."BED Amount" / TRL09."Quantity (Base)";
                            AssVal14 := 0;//TRL09."Assessable Value" / TRL09."Qty. per Unit of Measure";//14Nov2017

                        END;
                    END;

                    //GST Details
                    SIL10.RESET;
                    SIL10.SETRANGE("Document No.", BRTNo);
                    SIL10.SETRANGE("No.", "Item No.");
                    SIL10.SETFILTER(Quantity, '<>%1', 0);
                    IF SIL10.FINDFIRST THEN BEGIN
                        DGST10.RESET;
                        DGST10.SETRANGE(Type, DGST10.Type::Item);
                        DGST10.SETRANGE("Document No.", SIL10."Document No.");
                        DGST10.SETRANGE("No.", SIL10."No.");
                        DGST10.SETRANGE("Document Line No.", SIL10."Line No.");
                        IF DGST10.FINDSET THEN
                            REPEAT

                                IF DGST10."GST Component Code" = 'CGST' THEN
                                    CGSTRate10 := ABS(DGST10."GST Amount") / SIL10."Quantity (Base)";

                                IF DGST10."GST Component Code" = 'SGST' THEN
                                    SGSTRate10 := ABS(DGST10."GST Amount") / SIL10."Quantity (Base)";

                                IF DGST10."GST Component Code" = 'IGST' THEN
                                    IGSTRate10 := ABS(DGST10."GST Amount") / SIL10."Quantity (Base)";

                                IF DGST10."GST Component Code" = 'UTGST' THEN
                                    SGSTRate10 := ABS(DGST10."GST Amount") / SIL10."Quantity (Base)";

                            UNTIL DGST10.NEXT = 0;
                    END;

                END;
                //RB-N 09Nov2017 Transfer Details

                MakeExcelDataBody;
                //<<1

            end;

            trigger OnPreDataItem()
            begin

                //>>1
                "Item Ledger Entry".SETRANGE("Item Ledger Entry"."Posting Date", 0D, PostingDate);
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
                field("As on Date"; PostingDate)
                {
                }
                field("Show Zero Rem. Qty"; Showall)
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

    trigger OnPostReport()
    begin
        //>>1
        CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin
        //>>1


        ExcelBuf.RESET;
        ExcelBuf.DELETEALL;
        //>>09Mar2017 LocName
        LocCode := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code");
        recLOC.RESET;
        recLOC.SETRANGE(Code, LocCode);
        IF recLOC.FINDFIRST THEN BEGIN
            LocName := recLOC.Name;
        END;
        //<< LocName

        MakeExcelInfo;
        //<<1
    end;

    var
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        ItemRec: Record 27;
        VendorRec: Record 23;
        ValueEntry: Record 5802;
        RateofItem: Decimal;
        ChargeofItem: Decimal;
        ItemApplication: Record 339;
        RemQty: Decimal;
        PostingDate: Date;
        Showall: Boolean;
        TotalQty: Decimal;
        TotalCost: Decimal;
        VendorName: Text[60];
        UNITCOST: Decimal;
        UNITCOST1: Decimal;
        TotalCost1: Decimal;
        TotalQty1: Decimal;
        UNITCOSTNEW: Decimal;
        VLECOST: Decimal;
        BlenOrdrNo: Code[20];
        RecILE: Record 32;
        DocDate: Date;
        "----09Mar2017": Integer;
        LocCode: Code[100];
        recLOC: Record 14;
        LocName: Text[250];
        LocName1: Text[100];
        "-----21Sep2017": Integer;
        BRTNo: Code[20];
        TSH21: Record 5744;
        "----07Oct2017": Integer;
        RRH07: Record 6660;
        "---------------09Nov2017": Integer;
        TRL09: Record 5747;
        TRPrice09: Decimal;
        BedRate09: Decimal;
        ILE10: Record 32;
        DGST10: Record "Detailed GST Ledger Entry";
        CGSTRate10: Decimal;
        SGSTRate10: Decimal;
        IGSTRate10: Decimal;
        SIL10: Record 113;
        TSL10: Record 5745;
        "-----------14Nov2017": Integer;
        AssVal14: Decimal;
        RepItmNo: Code[10];
        "------------------11Dec2018": Integer;
        ILE11: Record 32;
        VE11: Record 5802;
        DenFactor: Decimal;

    //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        MakeExcelDataHeader;
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>02Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18 RB-N 21Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30 RB-N 10Nov2017
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18 RB-N 21Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29 RB-N 10Nov2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30 RB-N 10Nov2017
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.NewRow;

        //<<
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('FG Material Valuation report', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Inventory As On', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn(PostingDate, FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn(LocName, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//09Mar2017
        //ExcelBuf.AddColumn("Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code"),FALSE,'',TRUE,FALSE,TRUE,'@',1);
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4 //09Mar2017
        //ExcelBuf.AddColumn('Location Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//4
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Production Order No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Production Order Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Original Qty in SUOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Original Qty in Kg / Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Remaining Qty in SUOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Remaining Qty in Kg / Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Document Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13 // EBT MILAN ADDED 180314
        ExcelBuf.AddColumn('BRT No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 RB-N 21Sep2017
        ExcelBuf.AddColumn('Cost Per KG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Cost Per Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Total Cost Amount(Ltr)', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('Total Cost Amount(KG)', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19

        ExcelBuf.AddColumn('Ass. Value / Ltr', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20 RB-N 09Nov2017
        ExcelBuf.AddColumn('Total Ass. Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21 RB-N 09Nov2017
        ExcelBuf.AddColumn('Transfer Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 RB-N 09Nov2017
        ExcelBuf.AddColumn('Transfer Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 RB-N 09Nov2017
        ExcelBuf.AddColumn('ED Rate', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24 RB-N 09Nov2017
        ExcelBuf.AddColumn('ED Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25 RB-N 09Nov2017
        ExcelBuf.AddColumn('CGST Rate', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 RB-N 09Nov2017
        ExcelBuf.AddColumn('CGST Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 RB-N 10Nov2017
        ExcelBuf.AddColumn('SGST/UTGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 RB-N 09Nov2017
        ExcelBuf.AddColumn('SGST/UTGST Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 RB-N 10Nov2017
        ExcelBuf.AddColumn('IGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 RB-N 09Nov2017
        ExcelBuf.AddColumn('IGST Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 RB-N 09Nov2017
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Item Ledger Entry"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn(ItemRec.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Item Ledger Entry"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(LocName1, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        //ExcelBuf.AddColumn("Item Ledger Entry"."Location Code",FALSE,'',FALSE,FALSE,FALSE,'',1);//4
        ExcelBuf.AddColumn("Item Ledger Entry"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(BlenOrdrNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(DocDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//7
        IF "Item Ledger Entry"."Qty. per Unit of Measure" <> 0 THEN
            ExcelBuf.AddColumn("Item Ledger Entry".Quantity / "Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0)
        ELSE
            ExcelBuf.AddColumn("Item Ledger Entry".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//8

        ExcelBuf.AddColumn("Item Ledger Entry".Quantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//9
        IF "Item Ledger Entry"."Qty. per Unit of Measure" <> 0 THEN
            ExcelBuf.AddColumn(RemQty / "Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0)
        ELSE
            ExcelBuf.AddColumn(RemQty, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//10

        ExcelBuf.AddColumn(RemQty, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//11
        ExcelBuf.AddColumn("Item Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Item Ledger Entry"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//13 // EBT MILAN ADDED 180314

        ExcelBuf.AddColumn(BRTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14 RB-N 21Sep2017

        //ExcelBuf.AddColumn(BlenOrdrNo,FALSE,'',FALSE,FALSE,FALSE,''); // Comented by milan 110314
        //ExcelBuf.AddColumn(ROUND(RateofItem,0.00001),FALSE,'',FALSE,FALSE,FALSE,'');
        //MILAN ExcelBuf.AddColumn(ROUND(UNITCOST,0.00001),FALSE,'',FALSE,FALSE,FALSE,'');

        //ExcelBuf.AddColumn(VLECOST,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ABS(UNITCOST), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//15
        ExcelBuf.AddColumn(ABS(UNITCOST * DenFactor), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//11Dec2018

        /*
        IF UNITCOST1 <> 0 THEN
         ExcelBuf.AddColumn(UNITCOSTNEW,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0) //16
        ELSE
         ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',1);//16
        */
        ExcelBuf.AddColumn(ChargeofItem, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//17

        ExcelBuf.AddColumn(ABS(UNITCOST * RemQty + ChargeofItem), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//11Dec2018
        ExcelBuf.AddColumn(ABS(UNITCOST * DenFactor * RemQty + ChargeofItem), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//11Dec2018
                                                                                                                            /*
                                                                                                                            //ExcelBuf.AddColumn(ROUND((RateofItem+ChargeofItem),0.00001),FALSE,'',FALSE,FALSE,FALSE,'');
                                                                                                                            IF UNITCOST1 <> 0 THEN
                                                                                                                            BEGIN
                                                                                                                              ExcelBuf.AddColumn(ROUND((UNITCOSTNEW+ChargeofItem),0.00001),FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//18
                                                                                                                              ExcelBuf.AddColumn(ROUND((RemQty*UNITCOSTNEW+ChargeofItem),0.00001),FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//19
                                                                                                                            END ELSE
                                                                                                                            BEGIN
                                                                                                                              ExcelBuf.AddColumn(ROUND((VLECOST+ChargeofItem),0.00001),FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//18
                                                                                                                              ExcelBuf.AddColumn(ROUND((RemQty*VLECOST+ChargeofItem),0.00001),FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',0);//19
                                                                                                                            END;
                                                                                                                            */

        ExcelBuf.AddColumn(AssVal14, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//20 RB-N 14Nov2017 Ass. Value
        ExcelBuf.AddColumn((AssVal14 * RemQty), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//21 RB-N 14Nov2017 Ass Price

        ExcelBuf.AddColumn(TRPrice09, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22 RB-N 09Nov2017 TransferPrice
        ExcelBuf.AddColumn((TRPrice09 * RemQty), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23 RB-N 09Nov2017 TransferValue
        ExcelBuf.AddColumn(BedRate09, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24 RB-N 09Nov2017 EDRate
        ExcelBuf.AddColumn((BedRate09 * RemQty), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//25 RB-N 09Nov2017 EDValue
        ExcelBuf.AddColumn(CGSTRate10, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 RB-N 09Nov2017 CGST
        ExcelBuf.AddColumn((CGSTRate10 * RemQty), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 RB-N 10Nov2017 CGST
        ExcelBuf.AddColumn(SGSTRate10, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28 RB-N 09Nov2017 SGST
        ExcelBuf.AddColumn((SGSTRate10 * RemQty), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29 RB-N 10Nov2017 SGST
        ExcelBuf.AddColumn(IGSTRate10, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30 RB-N 09Nov2017 IGST
        ExcelBuf.AddColumn((IGSTRate10 * RemQty), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31 RB-N 10Nov2017 IGST

    end;

    // [Scope('Internal')]
    procedure CreateExcelbook()
    begin

        //>>02Mar2017 RB-N
        /* ExcelBuf.CreateBook('', 'FG Material Valuation Report');
        ExcelBuf.CreateBookAndOpenExcel('', 'FG Material Valuation Report', '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook('FG Material Valuation Report');
        ExcelBuf.WriteSheet('FG Material Valuation Report', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo('FG Material Valuation Report', CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<
    end;
}


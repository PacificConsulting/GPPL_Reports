report 50102 "Excise Duty Differences"
{
    // 17May2017 : Increase the Length Value of the code 250 --->OldValue (100)
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ExciseDutyDifferences.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Location; Location)
        {
            DataItemTableView = SORTING(Code);
            RequestFilterFields = "Code";
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Location Code" = FIELD(Code);
                DataItemTableView = SORTING("Document No.", "No.")
                                    WHERE(Type = FILTER(Item | 'Charge (Item)'),
                                          Quantity = FILTER(<> 0));
                RequestFilterFields = "Posting Date";

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    dec_Mrp := 0;
                    BRTExciseBaseAmount := 0;
                    BRTAssValue := 0;
                    GoodsValue := 0;
                    ExciseDuty := 0;
                    CessDuty := 0;
                    ShessDuty := 0;
                    ExciseDutyPaid := 0;
                    CessDutyPaid := 0;
                    ShessDutyPaid := 0;

                    Item.RESET;
                    IF Item.GET("Sales Invoice Line"."No.") THEN BEGIN
                        IF "Sales Invoice Line"."Qty. per Unit of Measure" > 20 THEN BEGIN
                            dec_Mrp := 0;
                            TotalAmount := 0;// "Sales Invoice Line"."MRP Price";
                        END
                        ELSE BEGIN
                            dec_Mrp := 0;// "Sales Invoice Line"."MRP Price";
                            TotalAmount := ((/*"Sales Invoice Line"."MRP Price"*/0 - ((/*"Sales Invoice Line"."MRP Price"*/0 * /*"Sales Invoice Line"."Abatement %"*/0) / 100))
                                            * ("Sales Invoice Line"."Qty. per Unit of Measure")) / ("Sales Invoice Line"."Qty. per Unit of Measure");
                        END;
                    END;

                    ExciseBaseAmount := TotalAmount * "Sales Invoice Line"."Quantity (Base)";
                    CustPostSetup.RESET;
                    CustPostSetup.FINDFIRST;
                    RoundOffAcc := CustPostSetup."Invoice Rounding Account";
                    CLEAR(RoundOffAmt);
                    recSIL.RESET;
                    recSIL.SETRANGE(recSIL."Document No.", "Sales Invoice Line"."Document No.");
                    recSIL.SETRANGE(recSIL.Type, recSIL.Type::"G/L Account");
                    recSIL.SETFILTER(recSIL."No.", RoundOffAcc);
                    IF recSIL.FINDSET THEN BEGIN
                        RoundOffAmt := (recSIL."Line Amount");
                    END;

                    CLEAR(BRNo);
                    CLEAR(BRTNO);
                    CLEAR(BRDate);
                    CLEAR("TRONo.");
                    CLEAR(TotalQtyfromRG);

                    /* RG23Drec.RESET;
                    RG23Drec.SETRANGE(RG23Drec."Transaction Type", RG23Drec."Transaction Type"::Sale);
                    RG23Drec.SETRANGE(RG23Drec.Type, RG23Drec.Type::Invoice);
                    RG23Drec.SETRANGE(RG23Drec."Document No.", "Sales Invoice Line"."Document No.");
                    RG23Drec.SETRANGE(RG23Drec."Item No.", "Sales Invoice Line"."No.");
                    RG23Drec.SETRANGE(RG23Drec."Line No.", "Sales Invoice Line"."Line No.");
                    RG23Drec.SETRANGE(RG23Drec."Lot No.", "Sales Invoice Line"."Lot No.");
                    RG23Drec.SETRANGE(RG23Drec."Location Code", "Sales Invoice Line"."Location Code");
                    RG23Drec.SETFILTER(RG23Drec.Quantity, '<>%1', 0);
                    IF RG23Drec.FINDFIRST THEN BEGIN
                        recRG23Dtable1.RESET;
                        recRG23Dtable1.SETRANGE(recRG23Dtable1."Entry No.", RG23Drec."Ref. Entry No.");
                        IF recRG23Dtable1.FINDFIRST THEN BEGIN
                            BRNo := recRG23Dtable1."Document No.";
                            "TRONo." := recRG23Dtable1."Order No.";
                        END;
                    END; */

                    recTSH.RESET;
                    recTSH.SETRANGE(recTSH."Transfer Order No.", "TRONo.");
                    IF recTSH.FINDFIRST THEN BEGIN
                        BRTNO := recTSH."No.";
                        BRDate := recTSH."Posting Date";
                    END;


                    /* ExcPostingSetup.RESET;
                    ExcPostingSetup.SETRANGE(ExcPostingSetup."Excise Bus. Posting Group", "Sales Invoice Line"."Excise Bus. Posting Group");
                    ExcPostingSetup.SETRANGE(ExcPostingSetup."Excise Prod. Posting Group", "Sales Invoice Line"."Excise Prod. Posting Group");
                    ExcPostingSetup.SETFILTER(ExcPostingSetup."From Date", '<=%1', "Sales Invoice Line"."Posting Date");//RSPL-Sourav160315
                    IF ExcPostingSetup.FINDLAST THEN BEGIN
                        BEDperc := ExcPostingSetup."BED %";
                        eCessperc := ExcPostingSetup."eCess %";
                        SheCessperc := ExcPostingSetup."SHE Cess %";
                    END; */

                    InvExciseDuty := ROUND(((ExciseBaseAmount * BEDperc) / 100), 0.01);
                    InvCessDuty := ROUND(((InvExciseDuty * eCessperc) / 100), 0.01);
                    InvSheCessDuty := ROUND(((InvExciseDuty * SheCessperc) / 100), 0.01);

                    TransferShipment.RESET;
                    TransferShipment.SETRANGE(TransferShipment."Document No.", BRTNO);
                    TransferShipment.SETRANGE(TransferShipment."Item No.", "Sales Invoice Line"."No.");
                    TransferShipment.SETFILTER(TransferShipment.Quantity, '<>%1', 0);
                    IF TransferShipment.FINDFIRST THEN BEGIN
                        TransferPrice := TransferShipment."Unit Price" / "Sales Invoice Line"."Qty. per Unit of Measure";
                        GoodsValue := TransferPrice * "Sales Invoice Line".Quantity * "Sales Invoice Line"."Qty. per Unit of Measure";
                        BRTAssValue := 0;//TransferShipment."Assessable Value" / "Sales Invoice Line"."Qty. per Unit of Measure";
                                         // IF TransferShipment."Excise Base Amount" <> 0 THEN
                        BRTExciseBaseAmount := ROUND(((/*TransferShipment."Excise Base Amount"*/0
                        * "Sales Invoice Line".Quantity) / TransferShipment.Quantity), 0.01);
                        //IF TransferShipment."BED Amount" <> 0 THEN
                        ExciseDuty := ROUND(((/*TransferShipment."BED Amount"*/0 * "Sales Invoice Line".Quantity) / TransferShipment.Quantity), 0.01);
                        //IF TransferShipment."eCess Amount" <> 0 THEN
                        CessDuty := ROUND(((/*TransferShipment."eCess Amount"*/0 * "Sales Invoice Line".Quantity) / TransferShipment.Quantity), 0.01);
                        //IF TransferShipment."SHE Cess Amount" <> 0 THEN
                        ShessDuty := ROUND(((/*TransferShipment."SHE Cess Amount"*/0 * "Sales Invoice Line".Quantity) / TransferShipment.Quantity), 0.01);
                    END;

                    IF BRTNO = '' THEN BEGIN
                        /* recRG23D.RESET;
                        recRG23D.SETRANGE(recRG23D."Transaction Type", recRG23D."Transaction Type"::Purchase);
                        recRG23D.SETRANGE(recRG23D.Type, recRG23D.Type::Transfer);
                        recRG23D.SETRANGE(recRG23D."Lot No.", "Sales Invoice Line"."Lot No.");
                        recRG23D.SETRANGE(recRG23D."Location Code", "Sales Invoice Line"."Location Code");
                        recRG23D.SETRANGE(recRG23D."Item No.", "Sales Invoice Line"."No.");
                        recRG23D.SETRANGE(recRG23D."Document No.", BRNo);
                        IF recRG23D.FINDFIRST THEN BEGIN */
                        GoodsValue := 0/*recRG23D."Excise Base Amt Per Unit"*/ * "Sales Invoice Line".Quantity *
                                     "Sales Invoice Line"."Qty. per Unit of Measure";
                        TransferPrice := 0;//recRG23D."Excise Base Amt Per Unit";
                        ExciseDuty := ROUND(((GoodsValue * BEDperc) / 100), 0.01);
                        CessDuty := ROUND(((ExciseDuty * eCessperc) / 100), 0.01);
                        ShessDuty := ROUND(((ExciseDuty * SheCessperc) / 100), 0.01);
                        //END;
                    END;


                    /* RG23Drec1.RESET;
                   RG23Drec1.SETRANGE(RG23Drec1."Transaction Type", RG23Drec1."Transaction Type"::Sale);
                   RG23Drec1.SETRANGE(RG23Drec1.Type, RG23Drec1.Type::Invoice);
                   RG23Drec1.SETRANGE(RG23Drec1."Document No.", "Sales Invoice Line"."Document No.");
                   RG23Drec1.SETRANGE(RG23Drec1."Item No.", "Sales Invoice Line"."No.");
                   RG23Drec1.SETRANGE(RG23Drec1."Line No.", "Sales Invoice Line"."Line No.");
                   RG23Drec1.SETRANGE(RG23Drec1."Lot No.", "Sales Invoice Line"."Lot No.");
                   RG23Drec1.SETRANGE(RG23Drec1."Location Code", "Sales Invoice Line"."Location Code");
                   RG23Drec1.SETFILTER(RG23Drec1.Quantity, '<>%1', 0);
                   IF RG23Drec1.COUNT > 1 THEN BEGIN  */

                    CLEAR(Exciseperunit);
                    CLEAR(BEDperUnit);
                    CLEAR(EcessPerUnit);
                    CLEAR(SheCessPerunit);
                    CLEAR(GoodsValue);
                    CLEAR(ExciseDuty);
                    CLEAR(CessDuty);
                    CLEAR(ShessDuty);
                    CLEAR(BRNo);
                    CLEAR(BRTNO);
                    CLEAR(BRDate);
                    CLEAR(TransferPrice);
                    CLEAR(TotalQtyfromRG);


                    /* RG23Drec1.RESET;
                    RG23Drec1.SETRANGE(RG23Drec1."Transaction Type", RG23Drec1."Transaction Type"::Sale);
                    RG23Drec1.SETRANGE(RG23Drec1.Type, RG23Drec1.Type::Invoice);
                    RG23Drec1.SETRANGE(RG23Drec1."Document No.", "Sales Invoice Line"."Document No.");
                    RG23Drec1.SETRANGE(RG23Drec1."Item No.", "Sales Invoice Line"."No.");
                    RG23Drec1.SETRANGE(RG23Drec1."Line No.", "Sales Invoice Line"."Line No.");
                    RG23Drec1.SETRANGE(RG23Drec1."Lot No.", "Sales Invoice Line"."Lot No.");
                    RG23Drec1.SETRANGE(RG23Drec1."Location Code", "Sales Invoice Line"."Location Code");
                    RG23Drec1.SETFILTER(RG23Drec1.Quantity, '<>%1', 0);
                    IF RG23Drec1.FINDFIRST THEN
                        REPEAT

                            RG23drec2.RESET;
                            RG23drec2.SETRANGE(RG23drec2."Entry No.", RG23Drec1."Ref. Entry No.");
                            IF RG23drec2.FINDFIRST THEN BEGIN */

                    CLEAR(UOM);
                    recITEMUOM.RESET;
                    //recITEMUOM.SETRANGE(recITEMUOM."Item No.", RG23Drec1."Item No.");
                    //recITEMUOM.SETRANGE(recITEMUOM.Code, RG23Drec1."Unit of Measure");
                    IF recITEMUOM.FINDFIRST THEN BEGIN
                        UOM := recITEMUOM."Qty. per Unit of Measure";
                        // END;

                        IF UOM = 0 THEN BEGIN
                            CLEAR(UOM);
                            recItemtable.RESET;
                            // recItemtable.SETRANGE(recItemtable."No.", RG23Drec1."Item No.");
                            IF recItemtable.FINDFIRST THEN BEGIN
                                recITEMUOM.RESET;
                                recITEMUOM.SETRANGE(recITEMUOM."Item No.", recItemtable."No.");
                                recITEMUOM.SETRANGE(recITEMUOM.Code, recItemtable."Sales Unit of Measure");
                                IF recITEMUOM.FINDFIRST THEN BEGIN
                                    UOM := recITEMUOM."Qty. per Unit of Measure";
                                END;
                            END;
                        END;

                        CLEAR(QtyfromRG);
                        CLEAR(Exciseperunit);
                        /*  RG23drec3.RESET;
                         RG23drec3.SETRANGE(RG23drec3."Transaction Type", RG23Drec1."Transaction Type"::Sale);
                         RG23drec3.SETRANGE(RG23drec3.Type, RG23Drec1.Type::Invoice);
                         RG23drec3.SETRANGE(RG23drec3."Document No.", "Sales Invoice Line"."Document No.");
                         RG23drec3.SETRANGE(RG23drec3."Item No.", "Sales Invoice Line"."No.");
                         RG23drec3.SETRANGE(RG23drec3."Line No.", "Sales Invoice Line"."Line No.");
                         RG23drec3.SETRANGE(RG23drec3."Lot No.", "Sales Invoice Line"."Lot No.");
                         RG23drec3.SETRANGE(RG23drec3."Location Code", "Sales Invoice Line"."Location Code");
                         RG23drec3.SETFILTER(RG23drec3.Quantity, '<>%1', 0);
                         RG23drec3.SETRANGE(RG23drec3."Ref. Entry No.", RG23Drec1."Ref. Entry No.");
                         IF RG23drec3.FINDFIRST THEN; */
                        BEGIN
                            QtyfromRG := QtyfromRG + 0;// RG23drec3.Quantity;
                        END;

                        TRRno := '';// RG23drec2."Document No.";
                        TOno := '';// RG23drec2."Order No.";

                        recTSH.RESET;
                        recTSH.SETRANGE(recTSH."Transfer Order No.", TOno);
                        IF recTSH.FINDFIRST THEN BEGIN
                            TSHno := recTSH."No.";
                            TSHDate := recTSH."Posting Date";
                        END;

                        TransferPrice := 0;// RG23drec2."Excise Base Amt Per Unit";
                        Exciseperunit := Exciseperunit + 0;// RG23drec2."Excise Base Amt Per Unit";
                        BEDperUnit := 0;// RG23drec2."BED Amount Per Unit";
                        EcessPerUnit := 0;//RG23drec2."eCess Amount Per Unit";
                        SheCessPerunit := 0;// RG23drec2."SHE Cess Amount Per Unit";

                        //END;

                        TotalQtyfromRG := TotalQtyfromRG + FORMAT(-QtyfromRG) + ',';
                        BRNo := BRNo + TRRno + ',';
                        BRTNO := BRTNO + TSHno + ',';

                        ExciseBaseAmt := QtyfromRG * UOM * Exciseperunit;
                        BED := QtyfromRG * UOM * BEDperUnit;
                        Cess := QtyfromRG * UOM * EcessPerUnit;
                        SheCess := QtyfromRG * UOM * SheCessPerunit;
                        GoodsValue := ROUND((GoodsValue + ABS(ExciseBaseAmt)), 0.01);
                        ExciseDuty := ROUND((ExciseDuty + ABS(BED)), 0.01);
                        CessDuty := ROUND((CessDuty + ABS(Cess)), 0.01);
                        ShessDuty := ROUND((ShessDuty + ABS(SheCess)), 0.01);

                        //UNTIL RG23Drec1.NEXT = 0;
                        //END;

                        ExciseDutyPaid := ROUND((InvExciseDuty - ExciseDuty), 0.01);
                        CessDutyPaid := ROUND((InvCessDuty - CessDuty), 0.01);
                        ShessDutyPaid := ROUND((InvSheCessDuty - ShessDuty), 0.01);

                        Totalqty := Totalqty + "Sales Invoice Line"."Qty. per Unit of Measure" * "Sales Invoice Line".Quantity;

                        recSIH.GET("Sales Invoice Line"."Document No.");
                        IF recSIH."Supplimentary Invoice" = TRUE THEN BEGIN
                            "Sales Invoice Line".Quantity := 0;
                            Totalqty := 0;
                        END;

                        //IPOL/BRT/001 >>>>  BRT Information Pick
                        IF (BRTNO = '') AND (BRNo <> '') THEN BEGIN
                            BRTNO := '';
                            BRDate := 0D;
                            recTTrRcpHd.RESET;
                            recTTrRcpHd.SETCURRENTKEY("No.");
                            recTTrRcpHd.SETFILTER("No.", BRNo);
                            IF recTTrRcpHd.FINDFIRST THEN BEGIN
                                recTrshipHd.RESET;
                                recTrshipHd.SETCURRENTKEY("No.");
                                recTrshipHd.SETRANGE("Transfer Order No.", recTTrRcpHd."Transfer Order No.");
                                IF recTrshipHd.FINDFIRST THEN BEGIN
                                    BRTNO := recTrshipHd."No.";
                                    BRDate := recTTrRcpHd."Posting Date";
                                END;
                            END;
                        END;
                        //IPOL/BRT/001 <<<< BRT Information Pick
                        //<<1


                        //>>2

                        //>>06July2017 NewDataGroupHeader
                        TCount += 1;
                        IF TCount = 1 THEN
                            MakeExcelDataGroupHeader;
                        //<<06July2017 NewDataGroupHeader

                        //Sales Invoice Line, Body (1) - OnPreSection()
                        MakeExcelDataBody;
                        //<<2
                    end;
                end;


                trigger OnPostDataItem()
                begin
                    //>>3

                    //Sales Invoice Line, Footer (2) - OnPreSection()
                    IF TCount <> 0 THEN
                        MakeExcelDataFooter;
                    //<<3
                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    CurrReport.CREATETOTALS(dec_Mrp, GoodsValue, ExciseDuty, CessDuty, ShessDuty, ExciseDutyPaid, CessDutyPaid, ShessDutyPaid);
                    CurrReport.CREATETOTALS(InvExciseDuty, InvSheCessDuty, InvCessDuty, Totalqty, TotalAmount, ExciseBaseAmount, BRTExciseBaseAmount);
                    //<<1

                    //>>06July2017

                    SIL07.COPYFILTERS("Sales Invoice Line");
                    TCount := 0;
                    ICount := COUNT;
                    //<<06July2017
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>06July2017 SKIPPING PLANT01 and PLANT02

                IF (Location.Code = 'PLANT01') OR (Location.Code = 'PLANT02') THEN
                    CurrReport.SKIP;

                //<<06July2017 SKIPPING PLANT01 and PLANT02

                //>>2
                /*
                CSOMapping2.RESET;
                CSOMapping2.SETRANGE(CSOMapping2."User Id",UPPERCASE(USERID));
                CSOMapping2.SETRANGE(Type,CSOMapping2.Type::Location);
                CSOMapping2.SETRANGE(CSOMapping2.Value,Location.GETFILTER(Location.Code));
                IF CSOMapping2.FINDFIRST THEN
                MARK := TRUE;
                BEGIN
                IF Location.GETFILTER(Location.Code) <> '' THEN
                  SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code",Location.Code);
                  Flag := TRUE;
                END;
                */
                //<<2

                //>>3

                //Location, GroupHeader (3) - OnPreSection()
                //CurrReport.SHOWOUTPUT :=
                //CurrReport.TOTALSCAUSEDBY = Location.FIELDNO(Code);

                //IF Flag THEN
                // MakeExcelDataGroupHeader;
                //<<3

                //>>4

                //Location, GroupFooter (5) - OnPreSection()
                Flag := FALSE;
                //<<4

            end;

            trigger OnPreDataItem()
            begin

                //>>1
                LastFieldNo := FIELDNO(Code);
                //<<1
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

    trigger OnPreReport()
    begin

        //>>1

        /*
        recLoc.RESET;
        recLoc.SETRANGE(recLoc.Code,Location.GETFILTER(Location.Code));
        IF recLoc.FINDFIRST THEN
        BEGIN
         IF NOT(recLoc."Trading Location") THEN
         ERROR('You are not allowed to run the report for other than Trading Location');
        END;
        */
        //Commented 06July2017

        /*
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",USERID);
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','Report View');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.RESET;
         recLocation.SETRANGE(recLocation.Code,Location.GETFILTER(Location.Code));
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping1.RESET;
         CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
         CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
         CSOMapping1.SETRANGE(CSOMapping1.Value,Location.GETFILTER(Location.Code));
         IF NOT(CSOMapping1.FINDFIRST) THEN
         BEGIN
          CSOMapping1.RESET;
          CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
          CSOMapping1.SETRANGE(CSOMapping1.Type,CSOMapping1.Type::"Responsibility Center");
          CSOMapping1.SETRANGE(CSOMapping1.Value,LocResC);
          IF NOT(CSOMapping1.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        *///Commented 26Apr2017

        ExcelBuf.RESET;
        ExcelBuf.DELETEALL;

        MakeExcelInfo;
        //<<1

    end;

    var
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        Item: Record 27;
        dec_Mrp: Decimal;
        TransferShipment: Record 5745;
        GoodsValue: Decimal;
        ExciseDuty: Decimal;
        CessDuty: Decimal;
        ShessDuty: Decimal;
        ExciseDutyPaid: Decimal;
        CessDutyPaid: Decimal;
        ShessDutyPaid: Decimal;
        ExcelBuf: Record 370 temporary;
        InvExciseDuty: Decimal;
        InvCessDuty: Decimal;
        InvSheCessDuty: Decimal;
        Flag: Boolean;
        SalesInvoiceLine: Record 113;
        TotalAmount: Decimal;
        Totalqty: Decimal;
        //RG23Drec: Record 16537;
        BRNo: Code[250];
        BRDate: Date;
        //ExcPostingSetup: Record 13711;
        BEDperc: Decimal;
        eCessperc: Decimal;
        SheCessperc: Decimal;
        //recRG23D: Record 16537;
        //recRG23D1: Record 16537;
        recTSH: Record 5744;
        "TRONo.": Code[250];
        BRTNO: Code[250];
        LocResC: Code[10];
        recLocation: Record 14;
        CSOMapping1: Record 50006;
        CSOMapping2: Record 50006;
        EDP: Decimal;
        EDR: Decimal;
        CessP: Decimal;
        CessR: Decimal;
        SheCessP: Decimal;
        SheCessR: Decimal;
        recSIH: Record 112;
        SupplInvAmt: Decimal;
        SupplInvED: Decimal;
        SupplInvCess: Decimal;
        SupplInvEcess: Decimal;
        billingprice: Decimal;
        recLoc: Record 14;
        recLoc1: Record 14;
        ExciseBaseAmount: Decimal;
        CustPostSetup: Record 92;
        RoundOffAcc: Code[10];
        recSIL: Record 113;
        RoundOffAmt: Decimal;
        ValueEntry: Record 5802;
        ILEEntryNo: Integer;
        "ItemApplEntryNo.": Integer;
        ItemApplEntry: Record 339;
        ILE: Record 32;
        TransferReceipt: Record 5746;
        TransferShipment1: Record 5744;
        //recRG23Dtable1: Record 16537;
        //recRG23Dtable2: Record 16537;
        recRRH: Record 6660;
        TOno: Code[250];
        TRRno: Code[250];
        TRHno: Code[250];
        TSHno: Code[250];
        TSHDate: Date;
        QtyfromRG: Decimal;
        recITEMUOM: Record 5404;
        UOM: Decimal;
        BEDperUnit: Decimal;
        EcessPerUnit: Decimal;
        SheCessPerunit: Decimal;
        Exciseperunit: Decimal;
        //RG23Drec1: Record 16537;
        //RG23drec2: Record 16537;
        ExciseBaseAmt: Decimal;
        BED: Decimal;
        Cess: Decimal;
        SheCess: Decimal;
        //RG23drec3: Record 16537;
        recItemtable: Record 27;
        TotalQtyfromRG: Text[30];
        TransferPrice: Decimal;
        BRTExciseBaseAmount: Decimal;
        BRTAssValue: Decimal;
        "-----IPOL/BRT/001-----": Integer;
        recTrshipHd: Record 5744;
        recTTrRcpHd: Record 5746;
        Text000: Label 'Data';
        Text001: Label 'Excise Duty Differences';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "-------26Apr2017": Integer;
        TQty: Decimal;
        TExciseBaseAmount: Decimal;
        TInvExciseDuty: Decimal;
        TInvCessDuty: Decimal;
        TInvSheCessDuty: Decimal;
        TGoodsValue: Decimal;
        TBRTExciseBaseAmount: Decimal;
        TExciseDuty: Decimal;
        TCessDuty: Decimal;
        TShessDuty: Decimal;
        "----07July2017": Integer;
        SIL07: Record 113;
        TCount: Integer;
        ICount: Integer;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//26Apr2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Excise Duty Differences', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Excise Duty Differences", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    //  //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>26Apr2017
        /* ExcelBuf.CreateBook('', Text001);
        ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<26Apr2017
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>26Apr2017 ReportHeader
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//35
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Text001, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//35
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        //<<26Apr2017 ReportHeader


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Sales Bill No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Description of Goods', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Unit Of Measure Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Pkg Size', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('No of Packs', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Total Qty in Ltr/KG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Billing Price per Ltr/KG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('MRP per Ltr/KG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Abatement %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Chapter Heading', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Assessable Value per Ltr/KG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Excise Base Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Excise Duty ON ASS VALUE', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('2% Cess on Ex.Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('1% Higher Cess on Ex. Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Transfer Receipt No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('BRT No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('BRT Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('Transfer Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('Goods Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Assessable Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('Excise Base Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        ExcelBuf.AddColumn('Excise Duty Paid', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('2% Cess on Ex. Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('1% Higher Cess on Ex.Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('ED Payable', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('ED Receivable', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('Cess Payable', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('Cess Receivable', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33
        ExcelBuf.AddColumn('SheCess Payable', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34
        ExcelBuf.AddColumn('SheCess Receivable', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35
        ExcelBuf.AddColumn('Qty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//36
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        recLoc1.GET("Sales Invoice Line"."Location Code");
        ExcelBuf.AddColumn(recLoc1.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Sales Invoice Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Sales Invoice Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//3
        ExcelBuf.AddColumn("Sales Invoice Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn("Sales Invoice Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn("Sales Invoice Line"."Lot No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn("Sales Invoice Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//8
        ExcelBuf.AddColumn("Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//9
        ExcelBuf.AddColumn("Sales Invoice Line"."Qty. per Unit of Measure" * "Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//10
        TQty += "Sales Invoice Line"."Qty. per Unit of Measure" * "Sales Invoice Line".Quantity;//26Apr2017

        //IF "Sales Invoice Line"."Price Inclusive of Tax" = TRUE THEN BEGIN
        ExcelBuf.AddColumn(ROUND(("Sales Invoice Line"."List Price"), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//11
                                                                                                                             //END
                                                                                                                             //ELSE BEGIN
        ExcelBuf.AddColumn(ROUND((/*"Sales Invoice Line"."MRP Price"*/0), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//11
        //END;

        ExcelBuf.AddColumn(ROUND((dec_Mrp), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//12
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Abatement %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//13
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Excise Prod. Posting Group"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14

        recSIH.GET("Sales Invoice Line"."Document No.");
        IF recSIH."Supplimentary Invoice" = TRUE THEN BEGIN
            ExcelBuf.AddColumn(ROUND(("Sales Invoice Line".Amount), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//15
            ExcelBuf.AddColumn(ROUND((/*"Sales Invoice Line"."BED Amount"*/0), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//16
            ExcelBuf.AddColumn(ROUND((/*"Sales Invoice Line"."eCess Amount"*/0), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//17
            ExcelBuf.AddColumn(ROUND((/*"Sales Invoice Line"."SHE Cess Amount"*/0), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//18
            SupplInvAmt := SupplInvAmt + "Sales Invoice Line".Amount;
            SupplInvED := SupplInvED + 0;//"Sales Invoice Line"."BED Amount";
            SupplInvCess := SupplInvCess + 0;// "Sales Invoice Line"."eCess Amount";
            SupplInvEcess := SupplInvEcess + 0;//"Sales Invoice Line"."SHE Cess Amount";
        END
        ELSE BEGIN
            ExcelBuf.AddColumn(ROUND((TotalAmount), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//15
            ExcelBuf.AddColumn(ROUND((ExciseBaseAmount), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//16
            TExciseBaseAmount += ExciseBaseAmount;//26Apr2017

            ExcelBuf.AddColumn(ROUND((InvExciseDuty), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//17
            TInvExciseDuty += InvExciseDuty;//26Apr2017

            ExcelBuf.AddColumn(ROUND((InvCessDuty), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//18
            TInvCessDuty += InvCessDuty;//26Apr2017

            ExcelBuf.AddColumn(ROUND((InvSheCessDuty), 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//19
            TInvSheCessDuty += InvSheCessDuty;//26Apr2017

        END;

        ExcelBuf.AddColumn(BRNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn(BRTNO, FALSE, '', FALSE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn(BRDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//22
        ExcelBuf.AddColumn(TransferPrice, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23
        ExcelBuf.AddColumn(GoodsValue, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//24
        TGoodsValue += GoodsValue;//26Apr2017

        ExcelBuf.AddColumn(BRTAssValue, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25
        ExcelBuf.AddColumn(BRTExciseBaseAmount, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//26
        TBRTExciseBaseAmount += BRTExciseBaseAmount; //26Apr2017

        ExcelBuf.AddColumn(ExciseDuty, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27
        TExciseDuty += ExciseDuty;//26Apr2017

        ExcelBuf.AddColumn(CessDuty, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//28
        TCessDuty += CessDuty;//26Apr2017

        ExcelBuf.AddColumn(ShessDuty, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//29
        TShessDuty += ShessDuty;//26Apr2017

        IF ExciseDutyPaid > 0 THEN BEGIN
            ExcelBuf.AddColumn(ExciseDutyPaid, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//30
            EDP := EDP + ExciseDutyPaid;
        END
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//30

        IF ExciseDutyPaid < 0 THEN BEGIN
            ExcelBuf.AddColumn(ExciseDutyPaid, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//31
            EDR := EDR + ExciseDutyPaid;
        END
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//31

        IF CessDutyPaid > 0 THEN BEGIN
            ExcelBuf.AddColumn(CessDutyPaid, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//32
            CessP := CessP + CessDutyPaid;
        END ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);//32

        IF CessDutyPaid < 0 THEN BEGIN
            ExcelBuf.AddColumn(CessDutyPaid, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//33
            CessR := CessR + CessDutyPaid;
        END
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);//33

        IF ShessDutyPaid > 0 THEN BEGIN
            ExcelBuf.AddColumn(ShessDutyPaid, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//34
            SheCessP := SheCessP + ShessDutyPaid;
        END
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//34

        IF ShessDutyPaid < 0 THEN BEGIN
            ExcelBuf.AddColumn(ShessDutyPaid, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//35
            SheCessR := SheCessR + ShessDutyPaid;
        END
        ELSE
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//35

        ExcelBuf.AddColumn(TotalQtyfromRG, FALSE, '', FALSE, FALSE, FALSE, '', 1);//36
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//36

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

        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//10 //26Apr2017
        TQty := 0;//26Apr2017
        //ExcelBuf.AddColumn(Totalqty,FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//10

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn(ROUND((TExciseBaseAmount), 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//16
        TExciseBaseAmount := 0;//26Apr2017
        //ExcelBuf.AddColumn(ROUND((ExciseBaseAmount),0.01),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//16

        ExcelBuf.AddColumn(ROUND((TInvExciseDuty + SupplInvED), 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//17 //26Apr2017
        //ExcelBuf.AddColumn(ROUND((InvExciseDuty + SupplInvED),0.01) ,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//17
        TInvExciseDuty := 0;//26Apr2017
        SupplInvED := 0;//26Apr2017

        ExcelBuf.AddColumn(ROUND((TInvCessDuty + SupplInvCess), 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//18 //26Apr2017
        TInvCessDuty := 0;//26Apr2017
        SupplInvCess := 0;//26Apr2017
        //ExcelBuf.AddColumn(ROUND((InvCessDuty + SupplInvCess),0.01) ,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//18

        ExcelBuf.AddColumn(ROUND((TInvSheCessDuty + SupplInvEcess), 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//19 //26Apr2017
        TInvSheCessDuty := 0;//26Apr2017
        SupplInvEcess := 0;//26Apr2017
        //ExcelBuf.AddColumn(ROUND((InvSheCessDuty + SupplInvEcess),0.01),FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//19

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn(TGoodsValue, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//24 //26Apr2017
        TGoodsValue := 0;//26Apr2017
        //ExcelBuf.AddColumn(GoodsValue,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//24

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn(TBRTExciseBaseAmount, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//26 //26Apr2017
        TBRTExciseBaseAmount := 0;//26Apr2017
        //ExcelBuf.AddColumn(BRTExciseBaseAmount,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//26

        ExcelBuf.AddColumn(TExciseDuty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//27 //26Apr2017
        TExciseDuty := 0;//26Apr2017
        //ExcelBuf.AddColumn(ExciseDuty,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//27

        ExcelBuf.AddColumn(TCessDuty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//28 //26Apr2017
        TCessDuty := 0;//26Apr2017
        //ExcelBuf.AddColumn(CessDuty,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//28

        ExcelBuf.AddColumn(TShessDuty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//29 //26Apr2017
        TShessDuty := 0;//26Apr2017
        //ExcelBuf.AddColumn(ShessDuty,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//29

        ExcelBuf.AddColumn(EDP, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//30
        EDP := 0; //26Apr2017

        ExcelBuf.AddColumn(EDR, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//31
        EDR := 0; //26Apr2017

        ExcelBuf.AddColumn(CessP, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//32
        CessP := 0; //26Apr2017

        ExcelBuf.AddColumn(CessR, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//33
        CessR := 0; //26Apr2017

        ExcelBuf.AddColumn(SheCessP, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//34
        SheCessP := 0; //26Apr2017

        ExcelBuf.AddColumn(SheCessR, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//35
        SheCessR := 0;  //26Apr2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//36
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataGroupHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn(Location.Code, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn(Location.Name, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
    end;
}


report 50086 "BRT Register GST"
{
    // 05Aug2017 :: RB-N, New Report for Branch Transfer Register --GST
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 11Sep2017   RB-N         Filtering of Report on the basis of Location GSTNo.
    // 18Sep2017   RB-N         New Column Added (Transfer Receipt No. & Date)
    // 20Sep2017   RB-N         PrintSummary Option & GST %
    DefaultLayout = RDLC;
    RDLCLayout = 'src/GPPL/Report Layout/BRTRegisterGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.", "Posting Date", "Transfer-from Code", "Transfer-to Code";
            dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));

                trigger OnAfterGetRecord()
                begin



                    //>>BatchNo
                    CLEAR(BatchNo);
                    recILE.SETRANGE("Document No.", "Document No.");
                    recILE.SETRANGE("Document Line No.", "Line No.");
                    recILE.SETRANGE("Item No.", "Item No.");
                    IF recILE.FINDFIRST THEN
                        BatchNo := recILE."Lot No.";

                    //>>Structure Discount
                    //>>05Aug2017

                    CLEAR(AVPDiscAmt);
                    /*  StructureDetailes.RESET;
                     StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                     StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                     StructureDetailes.SETRANGE("Item No.", "Item No.");
                     StructureDetailes.SETRANGE("Line No.", "Line No.");
                     StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                     StructureDetailes.SETFILTER("Tax/Charge Group", 'AVP-SP-SAN');
                     IF StructureDetailes.FINDFIRST THEN BEGIN
                         AVPDiscAmt := StructureDetailes.Amount;
                     END; */
                    TAVPDiscAmt += AVPDiscAmt;
                    TTAVPDiscAmt += AVPDiscAmt;

                    CLEAR(DiscAmt);
                    /* StructureDetailes.RESET;
                    StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                    StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                    StructureDetailes.SETRANGE("Item No.", "Item No.");
                    StructureDetailes.SETRANGE("Line No.", "Line No.");
                    StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                    StructureDetailes.SETFILTER("Tax/Charge Group", 'DISCOUNT');
                    IF StructureDetailes.FINDFIRST THEN BEGIN
                        DiscAmt := StructureDetailes.Amount;
                    END; */
                    TDiscAmt += DiscAmt;
                    TTDiscAmt += DiscAmt;

                    //<<05Aug2017



                    /* StructureDetailes.RESET;
                    StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                    StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                    StructureDetailes.SETRANGE("Item No.", "Item No.");
                    StructureDetailes.SETRANGE("Line No.", "Line No.");
                    StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                    StructureDetailes.SETFILTER("Tax/Charge Group", '%1|%2', 'TPTSUBSIDY', 'TRANSSUBS');
                    IF StructureDetailes.FINDSET THEN
                        REPEAT
                            TransportSubsidy += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0; */

                    /* StructureDetailes.RESET;
                    StructureDetailes.SETCURRENTKEY("Invoice No.", "Item No.", "Line No.", "Tax/Charge Type", "Tax/Charge Group");
                    StructureDetailes.SETRANGE("Invoice No.", "Document No.");
                    StructureDetailes.SETRANGE("Item No.", "Item No.");
                    StructureDetailes.SETRANGE("Line No.", "Line No.");
                    StructureDetailes.SETRANGE("Tax/Charge Type", StructureDetailes."Tax/Charge Type"::Charges);
                    StructureDetailes.SETRANGE("Tax/Charge Group", 'CASHDISC');
                    IF StructureDetailes.FINDSET THEN
                        REPEAT
                            CashDiscount += StructureDetailes.Amount;
                        UNTIL StructureDetailes.NEXT = 0; */
                    //<<Structure Discount

                    //>>05Aug2017 GST
                    CLEAR(vCGST);
                    CLEAR(vSGST);
                    CLEAR(vIGST);
                    CLEAR(vUTGST);
                    CLEAR(vCGSTRate);
                    CLEAR(vSGSTRate);
                    CLEAR(vIGSTRate);
                    CLEAR(vUTGSTRate);
                    CLEAR(GSTPer);//20Sep2017

                    DetailGSTEntry.RESET;
                    DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT
                            GSTPer += DetailGSTEntry."GST %";//20Sep2017

                            IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                vCGST += ABS(DetailGSTEntry."GST Amount");
                                vCGSTRate := DetailGSTEntry."GST %";
                            END;

                            IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                vSGST += ABS(DetailGSTEntry."GST Amount");
                                vSGSTRate := DetailGSTEntry."GST %";

                            END;

                            IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                vUTGST += ABS(DetailGSTEntry."GST Amount");
                                vUTGSTRate := DetailGSTEntry."GST %";

                            END;


                            IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                vIGST += ABS(DetailGSTEntry."GST Amount");
                                vIGSTRate := DetailGSTEntry."GST %";
                            END;

                        UNTIL DetailGSTEntry.NEXT = 0;

                    TCGST += vCGST;
                    TTCGST += vCGST;

                    TSGST += vSGST;
                    TTSGST += vSGST;

                    TIGST += vIGST;
                    TTIGST += vIGST;

                    TUTGST += vUTGST;
                    TTUTGST += vUTGST;
                    //<<05Aug2017 GST

                    NNNNN += 1;

                    //>>DataHeader
                    IF (NNNNN = 1) AND PrintToExcel THEN
                        MakeExcelDataHeader;

                    //>>DataBody
                    IF PrintToExcel THEN
                        IF PrintToExcel AND (PrintSummary = FALSE) THEN//20Sep2017
                            MakeExcelDataBody;

                    //GroupTotalQty
                    TQty += Quantity;
                    TTQty += Quantity;
                    TQtyBase += "Quantity (Base)";
                    TTQtyBase += "Quantity (Base)";

                    TAmt += Quantity * "Unit Price";
                    TTAmt += Quantity * "Unit Price";
                    TFinalAmt += Amount + 0;//"Total GST Amount";
                    TTFinalAmt += Amount + 0;// "Total GST Amount";
                    GSTBaseAmt08 += 0;//"GST Base Amount";//08Aug2017

                    //>>RoundOffAmount

                    FinalAmt += Amount + 0;//"Total GST Amount";//RoundOff


                    //>>Group
                    TSL.RESET;
                    TSL.SETRANGE("Document No.", "Document No.");
                    TSL.SETFILTER(Quantity, '<>%1', 0);
                    IF TSL.FINDFIRST THEN BEGIN
                        LLLL := TSL.COUNT;
                    END;

                    IIII += 1;

                    IF (LLLL = IIII) AND PrintToExcel THEN BEGIN
                        /*
                        //> Roundoff Amount
                        CLEAR(RoundOffAmnt);
                        CLEAR(InvoiceRoundOff);
                        InvoiceRoundOff := ROUND(FinalAmt,1.0);
                        RoundOffAmnt := InvoiceRoundOff - FinalAmt;
                        */
                        MakeExcelInvTotalGroupFooter;

                        IIII := 0;

                    END;

                    //TRoundOffAmnt += RoundOffAmnt;

                end;

                trigger OnPreDataItem()
                begin
                    IIII := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //
                "Transfer Shipment Header".CALCFIELDS(IRN, "EWB Date", "EWB No."); //AR
                IF "Transfer Shipment Header".IRN <> '' THEN
                    IRNGeneratedTxt := 'Yes'
                ELSE
                    IRNGeneratedTxt := 'No';
                //


                "Transfer Shipment Header".CALCFIELDS("Transfer Shipment Header"."EWB No.",//RSPLSUM 04Dec2020
                  "Transfer Shipment Header"."EWB Date", "Transfer Shipment Header".IRN);//RSPLSUM 04Dec2020

                //>> TransferFrom
                CLEAR(OriginState);
                CLEAR(OriginStateCode);
                CLEAR(OriginStateGSTNo);

                Loc05.RESET;
                IF Loc05.GET("Transfer-from Code") THEN BEGIN
                    OriginStateGSTNo := Loc05."GST Registration No.";

                    IF State05.GET(Loc05."State Code") THEN BEGIN
                        OriginState := State05.Description;
                        OriginStateCode := State05."State Code (GST Reg. No.)";
                    END;
                END;

                //>>RB-N 11Sep2017 GSTNo. Filter

                IF TempGSTNo <> '' THEN BEGIN
                    IF TempGSTNo <> OriginStateGSTNo THEN
                        CurrReport.SKIP;

                END;
                //<<RB-N 11Sep2017 GSTNo. Filter


                //>> TransferTo
                CLEAR(DestState);
                CLEAR(DestStateCode);
                CLEAR(DestStateGSTNo);

                Loc05.RESET;
                IF Loc05.GET("Transfer-to Code") THEN BEGIN
                    DestStateGSTNo := Loc05."GST Registration No.";

                    IF State05.GET(Loc05."State Code") THEN BEGIN
                        DestState := State05.Description;
                        DestStateCode := State05."State Code (GST Reg. No.)";
                    END;
                END;

                //>>MovementType
                CLEAR(MovementType);
                IF OriginState = DestState THEN
                    MovementType := 'INTRA STATE'
                ELSE
                    MovementType := 'INTER STATE';

                //>>05Aug2017 Comment
                recInvComment.RESET;
                recInvComment.SETRANGE(recInvComment."Document Type", recInvComment."Document Type"::"Posted Transfer Shipment");
                recInvComment.SETRANGE(recInvComment."No.", "Transfer Shipment Header"."No.");
                IF recInvComment.FINDFIRST THEN;

                //<<05Aug2017
            end;


            trigger OnPostDataItem()
            begin

                IF (NNNNN <> 0) AND PrintToExcel THEN
                    MakeExcelInvGrandTotal;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("GST Register No."; TempGSTNo)
                {
                    TableRelation = "GST Registration Nos.";
                }
                field("Print To Excel"; PrintToExcel)
                {
                }
                field("Print Summary"; PrintSummary)
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

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin

        NNNNN := 0;
        StartDate := "Transfer Shipment Header".GETRANGEMIN("Posting Date");
        EndDate := "Transfer Shipment Header".GETRANGEMAX("Posting Date");
    end;

    var
        Text002: Label 'Data';
        Text003: Label 'BRT  Register - GST';
        Text004: Label 'Company Name';
        Text005: Label 'Report No.';
        Text006: Label 'Report Name';
        Text007: Label 'User ID';
        Text008: Label 'Date';
        Text009: Label 'Location Filter';
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        GSTStateCode: Code[10];
        GSTPlaceofSupply: Text[50];
        GSTSupplyCode: Code[10];
        "----------------GST": Integer;
        GSTPer: Decimal;
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        vUTGST: Decimal;
        vUTGSTRate: Decimal;
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        TCGST: Decimal;
        TTCGST: Decimal;
        TSGST: Decimal;
        TTSGST: Decimal;
        TIGST: Decimal;
        TTIGST: Decimal;
        TUTGST: Decimal;
        TTUTGST: Decimal;
        FinalState: Text[50];
        OrgState: Text[50];
        InvRound: Decimal;
        AVPDiscAmt: Decimal;
        TAVPDiscAmt: Decimal;
        TTAVPDiscAmt: Decimal;
        DiscAmt: Decimal;
        TDiscAmt: Decimal;
        TTDiscAmt: Decimal;
        InTTransportSubsidy: Decimal;
        TransportSubsidy: Decimal;
        InTCashDiscount: Decimal;
        StartDate: Date;
        EndDate: Date;
        // StructureDetailes: Record 13798;
        CashDiscount: Decimal;
        State05: Record State;
        Loc05: Record 14;
        OriginState: Text[50];
        OriginStateCode: Code[10];
        OriginStateGSTNo: Code[15];
        DestState: Text[50];
        DestStateCode: Code[10];
        DestStateGSTNo: Code[15];
        MovementType: Text[20];
        NNNNN: Integer;
        BatchNo: Code[20];
        recILE: Record 32;
        recInvComment: Record 5748;
        TQty: Decimal;
        TTQty: Decimal;
        TQtyBase: Decimal;
        TTQtyBase: Decimal;
        IIII: Integer;
        LLLL: Integer;
        TSL: Record 5745;
        TAmt: Decimal;
        TTAmt: Decimal;
        TFinalAmt: Decimal;
        TTFinalAmt: Decimal;
        "---RoundOff": Integer;
        FinalAmt: Decimal;
        SL: Record 5745;
        RoundOffAmnt: Decimal;
        InvoiceRoundOff: Decimal;
        TRoundOffAmnt: Decimal;
        GSTBaseAmt08: Decimal;
        "------11Sep2017": Integer;
        TempGSTNo: Code[15];
        "------18Sep2017": Integer;
        TRH: Record 5746;
        "----20Sep2017": Integer;
        PrintSummary: Boolean;
        IRNGeneratedTxt: Text;

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text004), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text006), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(FORMAT(Text003), FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text005), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(50086, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text007), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.NewRow;
        ExcelBuf.ClearNewRow;



        //>>Report Header 04Apr2017

        IF PrintToExcel THEN BEGIN

            //>>04Apr2017
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
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//39
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//40
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//41
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//42
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//44
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//45
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//46
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//48
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//51
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//52
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//53
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//54
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//56
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//57
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//58 RSPLSUM 04Dec2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//59 RSPLSUM 04Dec2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//60 RSPLSUM 04Dec2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//61 RSPLSUM 27Jan21
            ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//58


            /*
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//58
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//59
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//60
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//61
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//62
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//63
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//64
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//65
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//66
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//67
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//68
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//69
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//70
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//71
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//72
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//73
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//74
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//75
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//76
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//77
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//78
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//79
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//80
            */

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
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//39
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//40
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//41
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//42
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//43
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//44
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//45
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//46
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//47
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//48
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//49
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//50
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//51
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//52
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//53
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//54
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//56
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//57
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//58 RSPLSUM 04Dec2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//59 RSPLSUM 04Dec2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//60 RSPLSUM 04Dec2020
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//60 RSPLSUM 27Jan21
            ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//58
            ExcelBuf.NewRow;
            /*
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//58
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//59
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//60
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//61
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//62
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//63
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//64
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//65
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//66
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//67
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//68
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//69
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//70
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//71
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//72
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//73
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//74
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//75
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//76
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//77
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//78
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//79
            ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//80
            */

            //<<

        END;

        //<<04Aug2017

    end;

    //[Scope('Internal')]
    procedure CreateExcelbook()
    begin

        //>>04Aug2017 RB-N
        /* ExcelBuf.CreateBook('', Text003);
        ExcelBuf.CreateBookAndOpenExcel('', Text003, '', '', USERID);
        ExcelBuf.GiveUserControl; */
        //<<

        ExcelBuf.CreateNewBook(Text003);
        ExcelBuf.WriteSheet(Text003, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text003, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

    end;

    //[Scope('Internal')]
    procedure MakeExcelInvTotalGroupFooter()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Transfer Shipment Header"."No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn("Transfer Shipment Header"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn("Transfer Shipment Header"."BRT Cancelled", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3 Cancelled
        ExcelBuf.AddColumn(MovementType, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-from Code", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-from Name", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn(OriginState, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn(OriginStateCode, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn(OriginStateGSTNo, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-to Code", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-to Name", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn(DestState, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn(DestStateCode, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn(DestStateGSTNo, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-to City", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//16  Supply State Code //08Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//22
        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//23
        TQty := 0;
        ExcelBuf.AddColumn(TQtyBase, FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//24
        TQtyBase := 0;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//25
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//26 //08Aug2017
        ExcelBuf.AddColumn(TAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//27
        TAmt := 0;
        ExcelBuf.AddColumn(TAVPDiscAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//28 AVP DISCOUNT
        ExcelBuf.AddColumn(TransportSubsidy, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//29 Transport Subsidy
        ExcelBuf.AddColumn(CashDiscount, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//30 Cash Discount Amount
        ExcelBuf.AddColumn(TDiscAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//31 Transfer Discount
        ExcelBuf.AddColumn((TAVPDiscAmt + TransportSubsidy + CashDiscount + TDiscAmt), FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//32 Total Discount
        TDiscAmt := 0;
        TAVPDiscAmt := 0;

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//33 FREIGHT
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'#,#0.00',0);//34 Insurance //08Aug2017
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'#,#0.00',0);//35 Other Charges //08Aug2017
        ExcelBuf.AddColumn(GSTBaseAmt08, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//36 Taxable Amount //08Aug2017
        GSTBaseAmt08 := 0;//08Aug2017
        //ExcelBuf.AddColumn("Transfer Shipment Line"."GST %",FALSE,'',TRUE,FALSE,FALSE,'0.00',0);//37 Tax % //08Aug2017
        ExcelBuf.AddColumn(GSTPer, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//37 Tax % //20Sep2017

        ExcelBuf.AddColumn(TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//38
        ExcelBuf.AddColumn(TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//39
        ExcelBuf.AddColumn(TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//40
        ExcelBuf.AddColumn(TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//41
        TCGST := 0;
        TSGST := 0;
        TIGST := 0;
        TUTGST := 0;

        //ExcelBuf.AddColumn(RoundOffAmnt,FALSE,'',TRUE,FALSE,FALSE,'0.00',0);//42 R/Off //08Aug2017
        ExcelBuf.AddColumn(TFinalAmt, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//43
        TFinalAmt := 0;
        RoundOffAmnt := 0;

        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//44 Product Discount Per Ltr //08Aug2017
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer Order No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//45
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer Order Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//46
        ExcelBuf.AddColumn("Transfer Shipment Header"."Vehicle No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//47
        //ExcelBuf.AddColumn("Transfer Shipment Header"."Road Permit No.",FALSE,'',TRUE,FALSE,FALSE,'@',1);//48 E-Way Bill No.
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//49 E-Way Bill Date
        ExcelBuf.AddColumn("Transfer Shipment Header"."LR/RR No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//50
        ExcelBuf.AddColumn("Transfer Shipment Header"."LR/RR Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//51
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transporter Name", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//52
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//53 Payment Terms Code //08Aug2017
        ExcelBuf.AddColumn("Transfer Shipment Header"."Driver's Mobile No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//54
        ExcelBuf.AddColumn("Transfer Shipment Header"."Frieght Type", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//56
        ExcelBuf.AddColumn(recInvComment.Comment, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//57 Comment
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//59 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//60 RSPLSUM 04Dec2020
        IF "Transfer Shipment Header"."EWB No." <> '' THEN
            ExcelBuf.AddColumn("Transfer Shipment Header"."EWB No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1)//61 RSPLSUM 04Dec2020
        ELSE
            ExcelBuf.AddColumn("Transfer Shipment Header"."Road Permit No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//61 RSPLSUM 04Dec2020

        ExcelBuf.AddColumn("Transfer Shipment Header"."EWB Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//62 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn("Transfer Shipment Header".IRN, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//63 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn(IRNGeneratedTxt, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn("Transfer Shipment Line"."Quantity (Base)", FALSE, '', TRUE, FALSE, FALSE, '0.000', 0);//RSPLSUM 27Jan21
        ExcelBuf.AddColumn(/*"Transfer Shipment Line"."Total GST Amount"*/0, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//RSPLSUM 27Jan21
    end;

    // [Scope('Internal')]
    procedure MakeExcelInvGrandTotal()
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//42
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//51
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 27Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 27Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 27Jan21
        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//54
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//55
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//56
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//57
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//58
        *///08Aug2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '', 1);
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
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn(TTQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//23
        ExcelBuf.AddColumn(TTQtyBase, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'0.000',0);//26
        ExcelBuf.AddColumn(TTAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//27
        TTAmt := 0;

        ExcelBuf.AddColumn(TTAVPDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//28
        ExcelBuf.AddColumn(InTTransportSubsidy, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//29
        ExcelBuf.AddColumn(InTCashDiscount, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//30
        ExcelBuf.AddColumn(TTDiscAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//31
        ExcelBuf.AddColumn((InTTransportSubsidy + InTCashDiscount + TTAVPDiscAmt + TTDiscAmt), FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//32
        InTTransportSubsidy := 0;//05Aug2017
        InTCashDiscount := 0;//05Aug2017
        TTAVPDiscAmt := 0;//05Aug2017
        TTDiscAmt := 0;//05Aug2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//33
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'#,#0.00',0);//34
        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'#,#0.00',0);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//37
        ExcelBuf.AddColumn(TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//38
        ExcelBuf.AddColumn(TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//39
        ExcelBuf.AddColumn(TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//40
        ExcelBuf.AddColumn(TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//41
        TTCGST := 0;//05Aug2017
        TTSGST := 0;//05Aug2017
        TTIGST := 0;//05Aug2017
        TTUTGST := 0;//05Aug2017

        //ExcelBuf.AddColumn(TRoundOffAmnt,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',0);//42
        ExcelBuf.AddColumn(TTFinalAmt, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//43
        TTFinalAmt := 0;
        TRoundOffAmnt := 0;

        //ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);           //47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//51
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);// 55
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//  56
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);// //57
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//57 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58 RSPLSUM 27Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58 RSPLSUM 27Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//58 RSPLSUM 27Jan21

    end;

    //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Transfer Shipment Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn("Transfer Shipment Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn("Transfer Shipment Header"."BRT Cancelled", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3 Cancelled //08Aug2017
        ExcelBuf.AddColumn(MovementType, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-from Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-from Name", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn(OriginState, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn(OriginStateCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn(OriginStateGSTNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-to Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-to Name", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn(DestState, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn(DestStateCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn(DestStateGSTNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer-to City", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//15
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'@',1);//16  Supply State Code //08Aug2017
        ExcelBuf.AddColumn("Transfer Shipment Line"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn("Transfer Shipment Line".Description, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn("Transfer Shipment Line"."HSN/SAC Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn(BatchNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn("Transfer Shipment Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn("Transfer Shipment Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//22
        ExcelBuf.AddColumn("Transfer Shipment Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//23
        ExcelBuf.AddColumn("Transfer Shipment Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//24
        ExcelBuf.AddColumn(/*"Transfer Shipment Line"."MRP Price"*/0 * "Transfer Shipment Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//25
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'@',1);//26 //08Aug2017
        ExcelBuf.AddColumn("Transfer Shipment Line".Amount, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27
        ExcelBuf.AddColumn(AVPDiscAmt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28 AVP DISCOUNT
        ExcelBuf.AddColumn(TransportSubsidy, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//29 Transport Subsidy
        ExcelBuf.AddColumn(CashDiscount, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30 Cash Discount Amount
        ExcelBuf.AddColumn(DiscAmt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31 Transfer Discount
        ExcelBuf.AddColumn((AVPDiscAmt + TransportSubsidy + CashDiscount + DiscAmt), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//32 Total Discount
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//33 FREIGHT
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//34 Insurance //08Aug2017
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//35 Other Charges //08Aug2017
        ExcelBuf.AddColumn(/*"Transfer Shipment Line"."GST Base Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//36 Taxable Amount //08Aug2017
        //ExcelBuf.AddColumn("Transfer Shipment Line"."GST %",FALSE,'',FALSE,FALSE,FALSE,'#,#0.00',0);//37 Tax % //08Aug2017
        ExcelBuf.AddColumn(GSTPer, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//37 Tax % //20Sep2017

        ExcelBuf.AddColumn(vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//38
        ExcelBuf.AddColumn(vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//39
        ExcelBuf.AddColumn(vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//40
        ExcelBuf.AddColumn(vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//41
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',0);//42 R/Off //08Aug2017
        ExcelBuf.AddColumn("Transfer Shipment Line".Amount + 0/*"Transfer Shipment Line"."Total GST Amount"*/, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//43
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'@',1);//44 Product Discount Per Ltr //08Aug2017
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer Order No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//45
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transfer Order Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//46
        ExcelBuf.AddColumn("Transfer Shipment Header"."Vehicle No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//47
        //ExcelBuf.AddColumn("Transfer Shipment Header"."Road Permit No.",FALSE,'',FALSE,FALSE,FALSE,'@',1);//48 E-Way Bill No.
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'@',1);//49 E-Way Bill Date
        ExcelBuf.AddColumn("Transfer Shipment Header"."LR/RR No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//50
        ExcelBuf.AddColumn("Transfer Shipment Header"."LR/RR Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//51
        ExcelBuf.AddColumn("Transfer Shipment Header"."Transporter Name", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//52
        //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'@',1);//53 Payment Terms Code //08Aug2017
        ExcelBuf.AddColumn("Transfer Shipment Header"."Driver's Mobile No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//54
        ExcelBuf.AddColumn("Transfer Shipment Header"."Frieght Type", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//55
        ExcelBuf.AddColumn("Transfer Shipment Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//56
        ExcelBuf.AddColumn(recInvComment.Comment, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//57 Comment
        ExcelBuf.AddColumn("Transfer Shipment Header"."Shortcut Dimension 1 Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//58

        //>>RB-N 18Sep2017 Transfer Receipt Details
        TRH.RESET;
        TRH.SETRANGE("Transfer Order No.", "Transfer Shipment Header"."Transfer Order No.");
        IF TRH.FINDFIRST THEN BEGIN
            ExcelBuf.AddColumn(TRH."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//52
            ExcelBuf.AddColumn(TRH."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//53

        END ELSE BEGIN

            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//52
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 2);//53

        END;

        //>>RB-N 18Sep2017 Transfer Receipt Details

        IF "Transfer Shipment Header"."EWB No." <> '' THEN
            ExcelBuf.AddColumn("Transfer Shipment Header"."EWB No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1)//54 RSPLSUM 04Dec2020
        ELSE
            ExcelBuf.AddColumn("Transfer Shipment Header"."Road Permit No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//54 RSPLSUM 04Dec2020

        ExcelBuf.AddColumn("Transfer Shipment Header"."EWB Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//55 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn("Transfer Shipment Header".IRN, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//56 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn(IRNGeneratedTxt, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn("Transfer Shipment Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//RSPLSUM 27Jan21
        ExcelBuf.AddColumn(/*"Transfer Shipment Line"."Total GST Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//RSPLSUM 27Jan21
        //MESSAGE('%1',"Transfer Shipment Line"."Total GST Amount");
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//41
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//42
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//50

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//51 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//52 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//53 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//54 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55 RSPLSUM 27Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55 RSPLSUM 27Jan21
        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//53
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//54
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//55
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//56
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//57
        *///08Aug2017
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//51

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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//41
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//42
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//51 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//52 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//53 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//54 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55 RSPLSUM 27Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//55 RSPLSUM 27Jan21
        /*
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//53
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//54
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//55
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//56
        ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//57
        *///08Aug2017
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//51

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('BRT Register GST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('For The Period', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('FROM', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn(FORMAT(StartDate), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('TO', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn(FORMAT(EndDate), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//37
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//38
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//40
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//41
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//42
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//44
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//45
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//46
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//47
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//48
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//49
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//50
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//51
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//52 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//53 RB-N 18Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//54 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//55 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 27Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 27Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//56 RSPLSUM 27Jan21
                                                                    /*
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//52
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//53
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//54
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//55
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//56
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//57
                                                                    ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'',1);//58
                                                                    *///08Aug2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Cancelled', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Movement Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Transfer-From Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Transfer-From Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Origin State', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Origin State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Origin GSTIN No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Transfer-To Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Transfer-To Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Destination State', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Destination State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Destination State GSTIN No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Place of Supply', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        //ExcelBuf.AddColumn('Supply State Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//16 //08Aug2017
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('Product HSN Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('Unit of Measure', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('Packing Qty per UOM', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('UOM Quantity Transfer', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('Total Quantity Transfer', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('MRP Per Lt/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        //ExcelBuf.AddColumn('List Price per Lt/Kg',FALSE,'',TRUE,FALSE,TRUE,'@',1);//26 //08Aug2017
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('AVP DISCOUNT', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('Transport Subsidy', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('Cash Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('Transfer Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('Total Discount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('FREIGHT', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33
        //ExcelBuf.AddColumn('Insurance',FALSE,'',TRUE,FALSE,TRUE,'@',1);//34 //08Aug2017
        //ExcelBuf.AddColumn('Other Charges',FALSE,'',TRUE,FALSE,TRUE,'@',1);//35 //08Aug2017
        ExcelBuf.AddColumn('Taxable Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//36
        ExcelBuf.AddColumn('TAX %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//37
        ExcelBuf.AddColumn('CGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//38
        ExcelBuf.AddColumn('SGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//39
        ExcelBuf.AddColumn('IGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//40
        ExcelBuf.AddColumn('UTGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//41
        //ExcelBuf.AddColumn('R/Off',FALSE,'',TRUE,FALSE,TRUE,'@',1);//42 //08Aug2017
        ExcelBuf.AddColumn('Gross  Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//43
        //ExcelBuf.AddColumn('Product Discount Per Ltr',FALSE,'',TRUE,FALSE,TRUE,'@',1);//44 //08Aug2017
        ExcelBuf.AddColumn('TRO Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//45
        ExcelBuf.AddColumn('TRO Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//46
        ExcelBuf.AddColumn('Vehicle Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//47
        //ExcelBuf.AddColumn('E-Way Bill No.',FALSE,'',TRUE,FALSE,TRUE,'@',1);//48
        //ExcelBuf.AddColumn('E-Way Bill Date',FALSE,'',TRUE,FALSE,TRUE,'@',1);//49
        ExcelBuf.AddColumn('LR Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//50
        ExcelBuf.AddColumn('LR Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//51
        ExcelBuf.AddColumn('Transporter Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//52
        //ExcelBuf.AddColumn('Payment Terms Code',FALSE,'',TRUE,FALSE,TRUE,'@',1);//53 //08Aug2017
        ExcelBuf.AddColumn('Drivers Mobile No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//54
        ExcelBuf.AddColumn('Freight Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//55
        ExcelBuf.AddColumn('Product Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//56
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//57
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//58
        ExcelBuf.AddColumn('Transfer Receipt No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//59 RB-N 18Sep2017
        ExcelBuf.AddColumn('Transfer Receipt Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//60 RB-N 18Sep2017
        ExcelBuf.AddColumn('Eway Bill No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//61 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('Eway Bill Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//62 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('IRN', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//63 RSPLSUM 04Dec2020
        ExcelBuf.AddColumn('IRN Generated', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Qty  Received against TRR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 27Jan21
        ExcelBuf.AddColumn('IGST on TRR', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 27Jan21

    end;
}


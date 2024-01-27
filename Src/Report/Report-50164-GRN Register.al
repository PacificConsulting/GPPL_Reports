report 50164 "GRN Register"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 16Sep2017   RB-N         GST Related Filed Added in the Report
    // 27Aug2018   RB-N         BasicPrice, GrossAmount Calculation
    // 07Oct2019   RB-N         Location Name
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/GRNRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Purch. Rcpt. Header"; "Purch. Rcpt. Header")
        {
            DataItemTableView = SORTING("No.")
                                ORDER(Ascending);
            RequestFilterFields = "Posting Date", "Location Code";
            dataitem("Purch. Rcpt. Line"; "Purch. Rcpt. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Quantity = FILTER(<> 0),
                                          Correction = FILTER(false));
                RequestFilterFields = Type, "No.", "Item Category Code";

                trigger OnAfterGetRecord()
                begin

                    BaseOil := 0;
                    Addetives := 0;
                    Packing := 0;
                    Chemicals := 0;
                    Others := 0;
                    CLEAR(DensityFactor);//DJ 25012021
                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT05' THEN
                        BaseOil := "Purch. Rcpt. Line".Quantity;
                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT01' THEN
                        Addetives := "Purch. Rcpt. Line".Quantity;
                    IF (("Purch. Rcpt. Line"."Item Category Code" = 'CAT04') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT08')) THEN
                        Packing := "Purch. Rcpt. Line".Quantity;
                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT06' THEN
                        Chemicals := "Purch. Rcpt. Line".Quantity;
                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT07' THEN
                        Others := "Purch. Rcpt. Line".Quantity;

                    /*  ExcisePosting.RESET;
                     ExcisePosting.SETRANGE(ExcisePosting."Excise Bus. Posting Group", "Purch. Rcpt. Line"."Excise Bus. Posting Group");
                     ExcisePosting.SETRANGE(ExcisePosting."Excise Prod. Posting Group", "Purch. Rcpt. Line"."Excise Prod. Posting Group");
                     IF ExcisePosting.FINDLAST THEN; */

                    CLEAR(VatAmt);
                    recPIL.RESET;
                    //recPIL.SETRANGE(recPIL."Receipt Document No.", "Purch. Rcpt. Line"."Document No.");
                    IF recPIL.FINDSET THEN
                        REPEAT
                            VatAmt := VatAmt + 0;// recPIL."Tax Amount";
                        UNTIL recPIL.NEXT = 0;

                    CLEAR(FreightAmount);
                    recPIL.RESET;
                    //recPIL.SETRANGE(recPIL."Receipt Document No.", "Purch. Rcpt. Line"."Document No.");
                    recPIL.SETRANGE(recPIL.Type, recPIL.Type::"Charge (Item)");
                    recPIL.SETFILTER(recPIL."No.", '%1|%2|%3|%4|%5|%3', 'FRTADD', 'FRTBASEOIL', 'FRTBASEOILIMP', 'FRTCHEM', 'FRTCONSU', 'FRTPKGMAT');
                    IF recPIL.FINDFIRST THEN BEGIN
                        FreightAmount := FreightAmount + recPIL.Amount;
                    END;

                    //RSPLSUM 20Jan21--CLEAR(VendorQty);
                    //DJ Comment 25012021
                    /*
                    recPIL1.RESET;
                    recPIL1.SETRANGE(recPIL1."Receipt Document No.","Purch. Rcpt. Line"."Document No.");
                    recPIL1.SETRANGE(recPIL1."Receipt Document Line No.","Purch. Rcpt. Line"."Line No.");
                    recPIL1.SETRANGE(recPIL1.Type,recPIL1.Type::Item);
                    recPIL1.SETRANGE(recPIL1."No.","Purch. Rcpt. Line"."No.");
                    IF recPIL1.FINDSET THEN
                    REPEAT
                     //RSPLSUM 20Jan21--VendorQty := recPIL1."Vendor Quantity";
                     DensityFactor := recPIL1."Density Factor";
                    UNTIL recPIL1.NEXT = 0;
                    */
                    //DJ Comment 25012021

                    //RSPLSUM 20Jan21>>
                    CLEAR(VendorQty);
                    CLEAR(ExpQty09);
                    IF ("Location Code" = 'PLANT01') AND (Type = Type::"G/L Account") THEN BEGIN
                        VendorQty := Quantity;
                        ExpQty09 := 0;
                        DensityFactor := "Density Factor";//DJ 25012021
                    END;

                    IF ("Location Code" <> 'PLANT01') AND ("Vendor Quantity" = 0) THEN BEGIN
                        VendorQty := Quantity;
                        ExpQty09 := 0;
                        DensityFactor := "Density Factor";//DJ 25012021
                    END;

                    IF ("Location Code" <> 'PLANT01') AND ("Vendor Quantity" <> 0) THEN BEGIN
                        VendorQty := "Purch. Rcpt. Line"."Vendor Quantity";
                        ExpQty09 := ROUND(("Purch. Rcpt. Line"."Vendor Quantity" * "Purch. Rcpt. Line"."Density Factor"), 1.0, '=');
                        DensityFactor := "Density Factor";//DJ 25012021
                    END;

                    IF ("Purch. Rcpt. Line".Correction = TRUE) THEN BEGIN
                        PWRL09.RESET;
                        PWRL09.SETCURRENTKEY("Posted Source No.");
                        PWRL09.SETRANGE("Posted Source Document", PWRL09."Posted Source Document"::"Posted Receipt");
                        PWRL09.SETRANGE("Posted Source No.", "Document No.");
                        PWRL09.SETRANGE("Line No.", "Line No.");
                        PWRL09.SETRANGE("Item No.", "No.");//29Jan2018
                        PWRL09.SETRANGE(Quantity, Quantity);//29Jan2018
                        IF PWRL09.FINDFIRST THEN BEGIN
                            PWRH09.RESET;
                            IF PWRH09.GET(PWRL09."No.") THEN
                                IF PWRH09."Vendor Quantity" <> 0 THEN BEGIN
                                    VendorQty := PWRL09.Quantity;
                                    ExpQty09 := PWRL09.Quantity * PWRH09."Density Factor";
                                    DensityFactor := PWRH09."Density Factor";//DJ 25012021
                                END;
                        END;
                    END;

                    IF ("Purch. Rcpt. Line".Correction = FALSE) THEN BEGIN
                        PWRL09.RESET;
                        PWRL09.SETCURRENTKEY("Posted Source No.");
                        PWRL09.SETRANGE("Posted Source Document", PWRL09."Posted Source Document"::"Posted Receipt");
                        PWRL09.SETRANGE("Posted Source No.", "Document No.");
                        PWRL09.SETRANGE("Source Line No.", "Line No.");
                        IF PWRL09.FINDFIRST THEN BEGIN
                            PWRH09.RESET;
                            IF PWRH09.GET(PWRL09."No.") THEN
                                IF PWRH09."Vendor Quantity" <> 0 THEN BEGIN
                                    VendorQty := PWRH09."Vendor Quantity";
                                    ExpQty09 := PWRH09."Vendor Quantity" * PWRH09."Density Factor";
                                    DensityFactor := PWRH09."Density Factor";//DJ 25012021
                                END;
                        END;
                    END;

                    PWRL09.RESET;
                    PWRL09.SETCURRENTKEY("Posted Source No.");
                    PWRL09.SETRANGE("Posted Source Document", PWRL09."Posted Source Document"::"Posted Receipt");
                    PWRL09.SETRANGE("Posted Source No.", "Document No.");
                    PWRL09.SETRANGE("Source Line No.", "Line No.");
                    IF PWRL09.FINDFIRST THEN BEGIN
                        PWRH09.RESET;
                        IF PWRH09.GET(PWRL09."No.") THEN BEGIN
                            IF PWRH09."Vendor Quantity" = 0 THEN BEGIN
                                VendorQty := PWRL09.Quantity;
                                ExpQty09 := PWRL09.Quantity * PWRH09."Density Factor";
                                DensityFactor := PWRH09."Density Factor";//DJ 25012021
                            END;
                        END;
                    END;

                    IF "Purch. Rcpt. Header"."Work Order" THEN BEGIN
                        PL19.RESET;
                        PL19.SETRANGE("Document Type", PL19."Document Type"::"Blanket Order");
                        PL19.SETRANGE("Document No.", "Blanket Order No.");
                        PL19.SETRANGE("Line No.", "Blanket Order Line No.");
                        IF PL19.FINDFIRST THEN BEGIN
                            VendorQty := PL19."Line Amount";
                            ExpQty09 := "Unit Cost (LCY)" * Quantity;
                        END;
                    END;
                    //RSPLSUM 20Jan21<<

                    //DJ 25012021 Start
                    IF "Purch. Rcpt. Header"."Work Order" THEN BEGIN
                        PRL19.RESET;
                        PRL19.SETCURRENTKEY("Blanket Order No.", "Blanket Order Line No.");
                        PRL19.SETRANGE("Blanket Order No.", "Blanket Order No.");
                        PRL19.SETRANGE("Blanket Order Line No.", "Blanket Order Line No.");
                        PRL19.SETRANGE("Document No.", '', "Document No.");
                        IF PRL19.FINDSET THEN
                            DensityFactor := 0;
                        REPEAT
                            IF (PRL19."Document No." <> "Document No.") THEN BEGIN
                                DensityFactor += PRL19."Unit Cost (LCY)" * PRL19.Quantity;
                            END;
                        UNTIL PRL19.NEXT = 0;
                    END;
                    //DJ 25012021 END

                    CLEAR("InvoiceNo.");
                    CLEAR(InvoiceDate);
                    recPIL1.RESET;
                    //recPIL1.SETRANGE(recPIL1."Receipt Document No.", "Purch. Rcpt. Line"."Document No.");
                    //recPIL1.SETRANGE(recPIL1."Receipt Document Line No.", "Purch. Rcpt. Line"."Line No.");
                    recPIL1.SETRANGE(Type, Type);// RB-N 19Sep2017
                    recPIL1.SETRANGE(recPIL1."No.", "Purch. Rcpt. Line"."No.");
                    //IF recPIL1.FINDSET THEN
                    IF recPIL1.FINDFIRST THEN
                    //REPEAT
                    BEGIN
                        "InvoiceNo." := recPIL1."Document No.";
                        InvoiceDate := recPIL1."Posting Date";
                        //UNTIL recPIL1.NEXT = 0;
                    END;


                    // Blanket order detail Fahim 14-04-2022
                    CLEAR(BPODate);
                    CLEAR(MRDate);
                    CLEAR(DealSheetNo);
                    CLEAR(DealSheetDate);
                    CLEAR(DeliveryDate);
                    RecBlPurHeader.RESET;
                    RecBlPurHeader.SETRANGE(RecBlPurHeader."No.", "Purch. Rcpt. Header"."Blanket Order No.");
                    IF RecBlPurHeader.FINDFIRST THEN
                    //MESSAGE('');
                    //REPEAT
                    BEGIN
                        BPODate := RecBlPurHeader."Posting Date";
                        DealSheetNo := RecBlPurHeader."Deal Sheet No.";
                        DealSheetDate := RecBlPurHeader."Deal Sheet Date";
                        DeliveryDate := RecBlPurHeader."Promised Receipt Date";
                        //UNTIL recPIL1.NEXT = 0;
                    END;
                    RecBlPurLine.RESET;
                    RecBlPurLine.SETRANGE(RecBlPurLine."Document No.", RecBlPurHeader."No.");
                    IF RecBlPurLine.FINDFIRST THEN BEGIN
                        MRNo := RecBlPurLine."MR No";
                        MRDate := RecBlPurLine."MR Date";
                    END;
                    //End Blanket order detail Fahim 14-04-2022

                    IF "InvoiceNo." <> '' THEN BEGIN
                        CLEAR(ReturnInvoiceNo);
                        CLEAR(ReturnInvoiceDate);
                        recPCH.RESET;
                        recPCH.SETRANGE(recPCH."Applies-to Doc. Type", recPCH."Applies-to Doc. Type"::Invoice);
                        recPCH.SETRANGE(recPCH."Applies-to Doc. No.", "InvoiceNo.");
                        IF recPCH.FINDFIRST THEN BEGIN
                            ReturnInvoiceNo := recPCH."No.";
                            ReturnInvoiceDate := recPCH."Posting Date";
                        END;
                    END;


                    //>>19Sep2017 InvoiceDetail
                    CLEAR(InvNo);
                    CLEAR(InvLinNo);
                    CLEAR(InvGSTAmt);

                    PIL19.RESET;
                    //PIL19.SETRANGE("Receipt Document No.", "Document No.");
                    //PIL19.SETRANGE("Receipt Document Line No.", "Line No.");
                    PIL19.SETRANGE(Type, Type);
                    PIL19.SETRANGE("No.", "No.");
                    IF PIL19.FINDFIRST THEN BEGIN
                        InvNo := PIL19."Document No.";
                        InvLinNo := PIL19."Line No.";
                        InvGSTAmt := 0;// PIL19."Total GST Amount";
                    END;

                    //>>19Sep2017 InvoiceDetail

                    //>>19Sep2017 GST
                    CLEAR(GSTPer);
                    CLEAR(vCGST);
                    CLEAR(vSGST);
                    CLEAR(vIGST);
                    CLEAR(vUTGST);
                    CLEAR(vCGSTRate);
                    CLEAR(vSGSTRate);
                    CLEAR(vIGSTRate);
                    CLEAR(vUTGSTRate);

                    DetailGSTEntry.RESET;
                    DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", InvNo);
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", InvLinNo);
                    //DetailGSTEntry.SETRANGE(Type, Type);
                    DetailGSTEntry.SETRANGE("No.", "No.");
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT

                            GSTPer += DetailGSTEntry."GST %";

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
                    //<<19Sep2017 GST

                    IF Printtoexcel THEN
                        MakeExcelDataBody;

                    TotalBaseOil += BaseOil;
                    TotalAddetives += Addetives;
                    TotalPacking += Packing;
                    TotalOthers += Others;
                    TotalChemicals += Chemicals;

                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT05' THEN
                        TotalAssValue := TotalAssValue + (/*"Purch. Rcpt. Line"."Assessable Value"*/0 * BaseOil);
                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT01' THEN
                        TotalAssValue := TotalAssValue + (/*"Purch. Rcpt. Line"."Assessable Value"*/0 * Addetives);
                    IF (("Purch. Rcpt. Line"."Item Category Code" = 'CAT04') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT08')) THEN
                        TotalAssValue := TotalAssValue + (/*"Purch. Rcpt. Line"."Assessable Value"*/0 * Packing);
                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT06' THEN
                        TotalAssValue := TotalAssValue + (/*"Purch. Rcpt. Line"."Assessable Value"*/0 * Chemicals);
                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT07' THEN
                        TotalAssValue := TotalAssValue + (/*"Purch. Rcpt. Line"."Assessable Value"*/0 * Others);

                    TotalVatAmt += VatAmt;
                    TotalFreightAmount += FreightAmount;
                    //>>RSPL/Migration/Rahul
                    vBEDAmt += 0;// "Purch. Rcpt. Line"."BED Amount";
                    vEsheCess += 0;//"Purch. Rcpt. Line"."eCess Amount";
                    vSheCess += 0;//"Purch. Rcpt. Line"."SHE Cess Amount";

                    IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT22' THEN
                        vQty += "Purch. Rcpt. Line"."Quantity (Base)"
                    ELSE
                        vQty += "Purch. Rcpt. Line".Quantity;

                    vADEAmt += 0;// "Purch. Rcpt. Line"."ADE Amount";
                    //<<

                end;
            }

            trigger OnAfterGetRecord()
            begin

                "E.C.C.No" := '';
                recVendor.GET("Purch. Rcpt. Header"."Buy-from Vendor No.");

                "E.C.C.No" := '';//recVendor."E.C.C. No.";

                //RSPLSUM 11Jan21>>
                CLEAR(GateEntryNo);
                PGEL.RESET;
                PGEL.SETCURRENTKEY("Entry Type", "Source Type", "Source No.");
                PGEL.SETRANGE("Entry Type", PGEL."Entry Type"::Inward);
                PGEL.SETRANGE("Source Type", PGEL."Source Type"::"Purchase Order");
                PGEL.SETRANGE("Source No.", "Purch. Rcpt. Header"."Order No.");
                IF PGEL.FINDFIRST THEN
                    GateEntryNo := PGEL."Gate Entry No.";
                //RSPLSUM 11Jan21<<
            end;

            trigger OnPostDataItem()
            begin

                IF Printtoexcel THEN
                    MakeExcelDataFooter;
            end;

            trigger OnPreDataItem()
            begin

                DateFilter := "Purch. Rcpt. Header".GETFILTER("Purch. Rcpt. Header"."Posting Date");
                LocationFiletr := "Purch. Rcpt. Header".GETFILTER("Purch. Rcpt. Header"."Location Code");

                TotalBaseOil := 0;
                TotalAddetives := 0;
                TotalPacking := 0;
                TotalOthers := 0;
                TotalChemicals := 0;
                TotalAssValue := 0;
                TotalVatAmt := 0;
                TotalFreightAmount := 0;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(Printtoexcel; Printtoexcel)
                {
                    ApplicationArea = all;
                    Caption = 'Print to Excel';
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

        IF Printtoexcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin

        LocationFiletr := "Purch. Rcpt. Header".GETFILTER("Purch. Rcpt. Header"."Location Code");
        DateFilter := FORMAT("Purch. Rcpt. Header"."Posting Date");

        //>>24Mar2017 LocName
        LocCode := "Purch. Rcpt. Header".GETFILTER("Purch. Rcpt. Header"."Location Code");

        recLoc.RESET;
        recLoc.SETRANGE(Code, LocCode);
        IF recLoc.FINDFIRST THEN
            LocName := recLoc.Name;
        //<<24Mar2017 LocName

        IF Printtoexcel THEN
            MakeExcelInfo;
    end;

    var
        Printtoexcel: Boolean;
        LocationRec: Record 14;
        LocaName: Text;
        // ECCNo: Record 13708;
        BaseOil: Decimal;
        Addetives: Decimal;
        Packing: Decimal;
        Others: Decimal;
        Chemicals: Decimal;
        DateFilter: Text[50];
        LocationFiletr: Code[20];
        //ExcisePosting: Record 13711;
        ExcelBuf: Record 370 temporary;
        TotalBaseOil: Decimal;
        TotalAddetives: Decimal;
        TotalPacking: Decimal;
        TotalOthers: Decimal;
        TotalChemicals: Decimal;
        recVendor: Record 23;
        "E.C.C.No": Code[20];
        TotalAssValue: Decimal;
        VatAmt: Decimal;
        TotalVatAmt: Decimal;
        recPIL: Record 123;
        FreightAmount: Decimal;
        TotalFreightAmount: Decimal;
        TotalADEamount: Decimal;
        VendorQty: Decimal;
        DensityFactor: Decimal;
        recPIL1: Record 123;
        "InvoiceNo.": Code[20];
        InvoiceDate: Date;
        recPCH: Record 124;
        ReturnInvoiceNo: Code[20];
        ReturnInvoiceDate: Date;
        Text0004: Label 'Company Name';
        Text0005: Label 'GRN Register for the Location ';
        Text0006: Label 'GRN For The Period of ';
        Text000: Label 'Data';
        Text001: Label 'GRN Register';
        "--": Integer;
        vBEDAmt: Decimal;
        vEsheCess: Decimal;
        vSheCess: Decimal;
        vADEAmt: Decimal;
        vQty: Decimal;
        "----24Mar2017": Integer;
        LocCode: Code[50];
        recLoc: Record 14;
        LocName: Text[100];
        TAssValue: Decimal;
        TBedAmt: Decimal;
        TeCess: Decimal;
        TSheCess: Decimal;
        TVatAmt: Decimal;
        TGrossAmt: Decimal;
        "--------------19Sep2017----GST": Integer;
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
        PIL19: Record 123;
        InvNo: Code[20];
        InvLinNo: Integer;
        InvGSTAmt: Decimal;
        RecItemN: Record 27;
        UOMBase: Code[20];
        PGEL: Record "Posted Gate Entry Line";
        GateEntryNo: Code[20];
        PWRL09: Record 7319;
        PWRH09: Record 7318;
        CompanyInfo1: Record 79;
        RespCentre: Record 5714;
        PL19: Record 39;
        ExpQty09: Decimal;
        PRL19: Record 121;
        RecBlPurHeader: Record 38;
        RecBlPurLine: Record 39;
        "BPONo.": Text;
        BPODate: Date;
        MRDate: Date;
        MRNo: Text;
        DealSheetNo: Text;
        DealSheetDate: Date;
        DeliveryDate: Date;

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(LocationFiletr, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(DateFilter, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 11Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 20Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        //>>

        //>>24Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GRN Register', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 11Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//RSPLSUM 20Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter :' + "Purch. Rcpt. Header".GETFILTER("Posting Date"), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        //<<24Mar2017

        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn("Purch. Rcpt. Header".GETFILTERS,FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location : ' + LocName, FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//24Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 11Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 20Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        //<<
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Receipt Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Receipt No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Party Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('HSN Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6 RB-N 19Sep2017
        ExcelBuf.AddColumn('UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Type of Doc', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Challan No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Challan Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Basic Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 27Aug2018
        ExcelBuf.AddColumn('Invoice QTY.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Density Factor', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Expected Qty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 20Jan21
        ExcelBuf.AddColumn('Base Oils', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Addetives', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Packing/Barrels', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Chemicals', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Others(Consumables)', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Mfg/Dir', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Rate of Duty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Asse Value', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('eCess', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('SHE Cess', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('VAT Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Freight/Insurance', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Addl Duty', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('GST %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 RB-N 19Sep2017
        ExcelBuf.AddColumn('CGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 RB-N 19Sep2017
        ExcelBuf.AddColumn('SGST %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 RB-N 19Sep2017
        ExcelBuf.AddColumn('IGST %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32 RB-N 19Sep2017
        ExcelBuf.AddColumn('UTGST %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33 RB-N 19Sep2017
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//34
        ExcelBuf.AddColumn('Chapter No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('ECC No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('PO Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Invoice Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Return Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Return Invoice Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vendor Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Blanket Order No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('Gate Entry Number', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 11Jan21

        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 13-01-22
        ExcelBuf.AddColumn('Responsibility Center', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 13-01-22

        ExcelBuf.AddColumn('BPO Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-22
        ExcelBuf.AddColumn('MR Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-22
        ExcelBuf.AddColumn('DealSheet No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-22
        ExcelBuf.AddColumn('DealSheet Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-22
        ExcelBuf.AddColumn('Delivery Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-22
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn("Purch. Rcpt. Header"."Posting Date",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Rcpt. Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//24Mar2017 //1
        ExcelBuf.AddColumn("Purch. Rcpt. Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Purch. Rcpt. Header"."Buy-from Vendor Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Purch. Rcpt. Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."HSN/SAC Code", FALSE, '', FALSE, FALSE, FALSE, '', 0);//6 RB-N 19Sep2017
        CLEAR(UOMBase);
        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT22' THEN BEGIN
            RecItemN.RESET;
            IF RecItemN.GET("Purch. Rcpt. Line"."No.") THEN
                UOMBase := RecItemN."Base Unit of Measure";
            ExcelBuf.AddColumn(UOMBase, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        END ELSE
            ExcelBuf.AddColumn("Purch. Rcpt. Line"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(FORMAT("Purch. Rcpt. Line".Type), FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn("Purch. Rcpt. Header"."Vendor Shipment No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
        //ExcelBuf.AddColumn("Purch. Rcpt. Header"."Order Date",FALSE,'',FALSE,FALSE,FALSE,'',2);//24mar2017//11
        ExcelBuf.AddColumn("Purch. Rcpt. Header"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//03Oct2019
        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT22' THEN
            ExcelBuf.AddColumn("Purch. Rcpt. Line"."Direct Unit Cost" / "Purch. Rcpt. Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0)//RB-N 27Aug2018
        ELSE
            ExcelBuf.AddColumn("Purch. Rcpt. Line"."Direct Unit Cost", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//RB-N 27Aug2018
        //ExcelBuf.AddColumn("Purch. Rcpt. Line".Quantity,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT22' THEN
            ExcelBuf.AddColumn("Purch. Rcpt. Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0)//24Mar2017 //12
        ELSE
            ExcelBuf.AddColumn("Purch. Rcpt. Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//24Mar2017 //12
        //ExcelBuf.AddColumn(VendorQty,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(VendorQty, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//24Mar2017 //13
        //ExcelBuf.AddColumn(DensityFactor,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(DensityFactor, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//24Mar2017 //14
        ExcelBuf.AddColumn(ExpQty09, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//RSPLSUM 20Jan21
        ExcelBuf.AddColumn(BaseOil, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); //15
        ExcelBuf.AddColumn(Addetives, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//16
        ExcelBuf.AddColumn(Packing, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//17
        ExcelBuf.AddColumn(Chemicals, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//18
        //ExcelBuf.AddColumn(Others,FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Others, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//24Mar2017 //19
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//20

        // IF "Purch. Rcpt. Line"."BED Amount" <> 0 THEN
        ExcelBuf.AddColumn(/*ExcisePosting."BED %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//24Mar2017 //21
                                                                                                  //ELSE
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//27Aug2018 //21
        //>>Assvalue 24Mar2017

        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT05' THEN BEGIN
            ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."Assessable Value"*/0 * BaseOil, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22
            TAssValue += 0;// "Purch. Rcpt. Line"."Assessable Value" * BaseOil;
        END;

        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT01' THEN BEGIN
            ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."Assessable Value"*/0 * Addetives, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22
            TAssValue += 0;// "Purch. Rcpt. Line"."Assessable Value" * Addetives;
        END;

        IF (("Purch. Rcpt. Line"."Item Category Code" = 'CAT04') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT08')) THEN BEGIN
            ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."Assessable Value"*/0 * Packing, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22
            TAssValue += 0;//"Purch. Rcpt. Line"."Assessable Value" * Packing;
        END;

        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT06' THEN BEGIN
            ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."Assessable Value"*/0 * Chemicals, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22
            TAssValue += 0;// "Purch. Rcpt. Line"."Assessable Value" * Chemicals;
        END;

        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT07' THEN BEGIN
            ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."Assessable Value"*/0 * Others, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//22
            TAssValue += 0;// "Purch. Rcpt. Line"."Assessable Value" * Others;
        END;

        //>>24Mar2017
        IF ("Purch. Rcpt. Line"."Item Category Code" = 'CAT14') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT15')
           OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT02') THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//22
        END;
        //>>24Mar2017

        IF "Purch. Rcpt. Line".Type = "Purch. Rcpt. Line".Type::"Charge (Item)" THEN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//22

        //>>24Mar2017 GLAccount
        IF "Purch. Rcpt. Line".Type = "Purch. Rcpt. Line".Type::"G/L Account" THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//22
                                                                          //ExcelBuf.AddColumn('',FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        END;
        //<<24Mar2017 GLAccount

        //AssValue  24Mar2017

        ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."BED Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//23
        //>>24Mar2017 BED
        TBedAmt += 0;//"Purch. Rcpt. Line"."BED Amount";
        //<<24Mar2017 BED
        ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."eCess Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24
        //>>24Mar2017 eCess
        TeCess += 0;//"Purch. Rcpt. Line"."eCess Amount";
        //<<24Mar2017 eCess
        ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."SHE Cess Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//25
        //>>24Mar2017 SheCess
        TSheCess += 0;// "Purch. Rcpt. Line"."SHE Cess Amount";
        //MESSAGE('She Cess %1',TSheCess);
        //<<24Mar2017 SheCess
        ExcelBuf.AddColumn(VatAmt, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26
        //>>24Mar2017 VAT
        TVatAmt += VatAmt;
        //<<24Mar2017 VAT
        ExcelBuf.AddColumn(FreightAmount, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27
        ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."ADE Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28

        ExcelBuf.AddColumn(GSTPer, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//29 RB-N 19Sep2017 GST%
        ExcelBuf.AddColumn(vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//30 RB-N 19Sep2017 CGST
        ExcelBuf.AddColumn(vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//31 RB-N 19Sep2017 SGST
        ExcelBuf.AddColumn(vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//32 RB-N 19Sep2017 IGST
        ExcelBuf.AddColumn(vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//33 RB-N 19Sep2017 UTGST

        IF ("Purch. Rcpt. Line"."Item Category Code" = 'CAT14') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT15') THEN BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);//34
        END;


        //>>24Mar2017 GrossAmount
        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT05' THEN BEGIN
            /* IF "Purch. Rcpt. Line"."Assessable Value" <> 0 THEN BEGIN
                ExcelBuf.AddColumn((("Purch. Rcpt. Line"."Assessable Value" * BaseOil) + "Purch. Rcpt. Line"."BED Amount" +
                                 "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                 "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//34

                TGrossAmt += ("Purch. Rcpt. Line"."Assessable Value" * BaseOil) + "Purch. Rcpt. Line"."BED Amount" +
                                 "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                 "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt;
            END; */
        END;

        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT01' THEN BEGIN
            /* IF "Purch. Rcpt. Line"."Assessable Value" <> 0 THEN BEGIN
                ExcelBuf.AddColumn((("Purch. Rcpt. Line"."Assessable Value" * Addetives) + "Purch. Rcpt. Line"."BED Amount" +
                                     "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                     "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//28

                TGrossAmt += ("Purch. Rcpt. Line"."Assessable Value" * Addetives) + "Purch. Rcpt. Line"."BED Amount" +
                                     "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                     "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt;
            END; */
        END;

        IF (("Purch. Rcpt. Line"."Item Category Code" = 'CAT04') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT08')) THEN BEGIN
            /*  IF "Purch. Rcpt. Line"."Assessable Value" <> 0 THEN BEGIN
                 ExcelBuf.AddColumn((("Purch. Rcpt. Line"."Assessable Value" * Packing) + "Purch. Rcpt. Line"."BED Amount" +
                                      "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                      "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//28

                 TGrossAmt += ("Purch. Rcpt. Line"."Assessable Value" * Packing) + "Purch. Rcpt. Line"."BED Amount" +
                                      "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                      "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt;
             END; */
        END;

        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT06' THEN BEGIN
            /*  IF "Purch. Rcpt. Line"."Assessable Value" <> 0 THEN BEGIN
                 ExcelBuf.AddColumn((("Purch. Rcpt. Line"."Assessable Value" * Chemicals) + "Purch. Rcpt. Line"."BED Amount" +
                                      "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                      "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//28

                 TGrossAmt += ("Purch. Rcpt. Line"."Assessable Value" * Chemicals) + "Purch. Rcpt. Line"."BED Amount" +
                                      "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                      "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt;
             END; */
        END;

        IF "Purch. Rcpt. Line"."Item Category Code" = 'CAT07' THEN BEGIN
            /*  IF "Purch. Rcpt. Line"."Assessable Value" <> 0 THEN BEGIN
                 ExcelBuf.AddColumn((("Purch. Rcpt. Line"."Assessable Value" * Others) + "Purch. Rcpt. Line"."BED Amount" +
                                      "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                      "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28

                 TGrossAmt += ("Purch. Rcpt. Line"."Assessable Value" * Others) + "Purch. Rcpt. Line"."BED Amount" +
                                      "Purch. Rcpt. Line"."eCess Amount" + "Purch. Rcpt. Line"."SHE Cess Amount" + VatAmt + FreightAmount +
                                      "Purch. Rcpt. Line"."ADE Amount" + InvGSTAmt;
             END; */
        END;
        //<<24Mar2017 GrossAmount

        //>>27Aug2018 GrossAmount
        //IF "Purch. Rcpt. Line"."Assessable Value" = 0 THEN
        IF ("Purch. Rcpt. Line"."Item Category Code" = 'CAT07') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT06')
        OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT04') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT08')
        OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT01') OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT05')
        OR ("Purch. Rcpt. Line"."Item Category Code" = 'CAT02') THEN BEGIN
            ExcelBuf.AddColumn((("Purch. Rcpt. Line".Quantity * "Purch. Rcpt. Line"."Direct Unit Cost")
                                + InvGSTAmt), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28

            TGrossAmt += ("Purch. Rcpt. Line".Quantity * "Purch. Rcpt. Line"."Direct Unit Cost") + InvGSTAmt;
        END;
        //>>27Aug2018 GrossAmount


        IF "Purch. Rcpt. Line".Type = "Purch. Rcpt. Line".Type::"Charge (Item)" THEN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//34
        IF "Purch. Rcpt. Line".Type = "Purch. Rcpt. Line".Type::"G/L Account" THEN BEGIN
            ExcelBuf.AddColumn((("Purch. Rcpt. Line".Quantity * "Purch. Rcpt. Line"."Direct Unit Cost")
                                + InvGSTAmt), FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28
            TGrossAmt += ("Purch. Rcpt. Line".Quantity * "Purch. Rcpt. Line"."Direct Unit Cost") + InvGSTAmt;
        END;

        ExcelBuf.AddColumn(/*"Purch. Rcpt. Line"."Excise Prod. Posting Group"*/'', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("E.C.C.No", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Rcpt. Header"."Order No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Rcpt. Header"."Order Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//24Mar2017
        ExcelBuf.AddColumn("InvoiceNo.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(InvoiceDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//24Mar2017
        ExcelBuf.AddColumn(ReturnInvoiceNo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(ReturnInvoiceDate,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(ReturnInvoiceDate, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//24Mar2017
        IF "Purch. Rcpt. Line"."Vendor Item No." <> '' THEN
            ExcelBuf.AddColumn("Purch. Rcpt. Line"."Vendor Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text)
        ELSE
            ExcelBuf.AddColumn("Purch. Rcpt. Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Purch. Rcpt. Header"."Blanket Order No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        //>>07Oct2019
        LocaName := '';
        LocationRec.RESET;
        IF LocationRec.GET("Purch. Rcpt. Line"."Location Code") THEN
            LocaName := LocationRec.Name;
        ExcelBuf.AddColumn("Purch. Rcpt. Line"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(LocaName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //<<07Oct2019

        ExcelBuf.AddColumn(GateEntryNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//RSPLSUM 11Jan21

        ExcelBuf.AddColumn(GateEntryNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//Fahim 13-01-22
        ExcelBuf.AddColumn(GateEntryNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//Fahim 13-01-22

        ExcelBuf.AddColumn(BPODate, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//Fahim 14-04-22
        ExcelBuf.AddColumn(MRDate, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//Fahim 14-04-22
        ExcelBuf.AddColumn(DealSheetNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//Fahim 14-04-22
        ExcelBuf.AddColumn(DealSheetDate, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//Fahim 14-04-22
        ExcelBuf.AddColumn(DeliveryDate, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//Fahim 14-04-22
    end;

    //   //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        /* ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', '');
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text000);
        ExcelBuf.WriteSheet(Text000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//24Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//24Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 27Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 11Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 20Jan21

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//7 RB-N 19Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 27Aug2017
        //ExcelBuf.AddColumn(vQty,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(vQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', ExcelBuf."Cell Type"::Number);//24Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//RSPLSUM 20Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TotalBaseOil, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalAddetives, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalPacking, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalChemicals, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TotalOthers, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn(TotalAssValue,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TAssValue, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//24Mar2017 //21
        //"Purch. Rcpt. Line".CALCSUMS("BED Amount");
        //ExcelBuf.AddColumn(vBEDAmt,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TBedAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//24Mar2017 //22
        //ExcelBuf.AddColumn(vEsheCess,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TeCess, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//24Mar2017 //23
        //ExcelBuf.AddColumn(vSheCess,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TSheCess, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//24Mar2017 //24
        //ExcelBuf.AddColumn(TotalVatAmt,FALSE,'',TRUE,FALSE,TRUE,'#,####0.00',ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn(TVatAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//24Mar2017 //25

        ExcelBuf.AddColumn(TotalFreightAmount, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//26
        ExcelBuf.AddColumn(vADEAmt, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', ExcelBuf."Cell Type"::Number);//27
                                                                                                              //ExcelBuf.AddColumn(TotalAssValue+vBEDAmt+vEsheCess
                                                                                                              //                 +vSheCess+TotalVatAmt+TotalFreightAmount+
                                                                                                              //               vADEAmt,FALSE,'',TRUE,FALSE,TRUE,'@',ExcelBuf."Cell Type"::Number);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//29 RB-N 19Sep2017 GST%
        ExcelBuf.AddColumn(TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//30 RB-N 19Sep2017 CGST
        ExcelBuf.AddColumn(TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//31 RB-N 19Sep2017 SGST
        ExcelBuf.AddColumn(TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//32 RB-N 19Sep2017 IGST
        ExcelBuf.AddColumn(TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//33 RB-N 19Sep2017 UTGST

        ExcelBuf.AddColumn(TGrossAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//24Mar2017 //34

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(/*FORMAT(ECCNo.Description)*/'', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//24Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);//24Mar2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RB-N 07Oct2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//RSPLSUM 11Jan21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 14-04-2022
    end;
}


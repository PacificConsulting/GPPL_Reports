report 50196 "Purchase Order Statement"
{
    // Date         Version        Remarks
    // ----------------------------------------------------
    // 02Jun2018    RB-N           GT Code, Payment Terms
    // 19Oct2018    RB-N           Print Summary Option
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseOrderStatement.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Blanket Purchase Line"; "Purchase Line")
        {
            DataItemTableView = SORTING("Document Type", "Document No.", "Line No.")
                                WHERE("Document Type" = FILTER("Blanket Order"));
            RequestFilterFields = "Document No.", "Order Date";
            dataitem("Purchase Line"; "Purchase Line")
            {
                DataItemLink = "Blanket Order No." = FIELD("Document No."),
                               "Blanket Order Line No." = FIELD("Line No.");
                DataItemTableView = SORTING("Document Type", "Document No.", "Line No.")
                                    WHERE("Document Type" = FILTER(Order),
                                          Quantity = FILTER(<> 0));

                trigger OnAfterGetRecord()
                begin

                    //>>19Oct2018
                    IF Summary THEN
                        CurrReport.SKIP;
                    //<<19Oct2018

                    TCOUNT += 1; //20Mar2017

                    //>>1

                    CLEAR(BlanketExcisePer);
                    // IF "Blanket Purchase Line"."Excise Amount" <> 0 THEN BEGIN
                    /*  RecEPS.RESET;
                     // RecEPS.SETRANGE(RecEPS."Excise Bus. Posting Group", "Blanket Purchase Line"."Excise Bus. Posting Group");
                     // RecEPS.SETRANGE(RecEPS."Excise Prod. Posting Group", "Blanket Purchase Line"."Excise Prod. Posting Group");
                     IF RecEPS.FINDFIRST THEN
                         BlanketExcisePer := RecEPS."BED %"; */
                    //  END;

                    CLEAR(grnNO);
                    CLEAR(GRNDATE);
                    RecPRL.RESET;
                    RecPRL.SETRANGE(RecPRL."Order No.", "Purchase Line"."Document No.");
                    RecPRL.SETRANGE(RecPRL."No.", "Purchase Line"."No.");
                    RecPRL.SETRANGE(RecPRL."Line No.", "Purchase Line"."Line No.");
                    IF RecPRL.FINDFIRST THEN BEGIN
                        grnNO := RecPRL."Document No.";
                        GRNDATE := RecPRL."Posting Date";
                    END;

                    //RSPLSUM 11Jan21>>
                    CLEAR(GateEntryNo);
                    RecPRH.RESET;
                    IF RecPRH.GET(grnNO) THEN BEGIN
                        /*  PGEL.RESET;
                         // PGEL.SETCURRENTKEY("Entry Type", "Source Type", "Source No.");
                         // PGEL.SETRANGE("Entry Type", PGEL."Entry Type"::Inward);
                         //PGEL.SETRANGE("Source Type", PGEL."Source Type"::"Purchase Order");
                         // PGEL.SETRANGE("Source No.", RecPRH."Order No.");
                         IF PGEL.FINDFIRST THEN
                             GateEntryNo := PGEL."Gate Entry No."; */
                    END;
                    //RSPLSUM 11Jan21<<

                    CLEAR(InvNo);
                    CLEAR(InvDate);
                    recILE.RESET;
                    recILE.SETRANGE(recILE."Document No.", grnNO);
                    recILE.SETRANGE(recILE."Item No.", "No.");
                    IF recILE.FINDFIRST THEN BEGIN
                        recVal.RESET;
                        recVal.SETRANGE(recVal."Item Ledger Entry No.", recILE."Entry No.");
                        recVal.SETRANGE(recVal."Document Type", recVal."Document Type"::"Purchase Invoice");
                        IF recVal.FINDFIRST THEN BEGIN
                            InvNo := recVal."Document No.";
                            InvDate := recVal."Posting Date";
                        END;
                    END;

                    CLEAR(PayVoucher);
                    CLEAR(PayDate);
                    recVLE.RESET;
                    recVLE.SETRANGE(recVLE."Document No.", InvNo);
                    IF recVLE.FINDFIRST THEN BEGIN
                        recDVLE.RESET;
                        recDVLE.SETRANGE(recDVLE."Vendor Ledger Entry No.", recVLE."Entry No.");
                        recDVLE.SETRANGE(recDVLE."Entry Type", recDVLE."Entry Type"::Application);
                        recDVLE.SETRANGE(recDVLE."Initial Document Type", recDVLE."Initial Document Type"::Invoice);
                        IF recDVLE.FINDFIRST THEN BEGIN
                            PayVoucher := recDVLE."Document No.";
                            PayDate := recDVLE."Posting Date";
                        END;
                    END;
                    //<<1

                    //>>20Mar2017 PurchaseLineBody

                    EnterCell(NN, 17, "Document No.", FALSE, FALSE, '', 1);
                    EnterCell(NN, 18, FORMAT("Order Date"), FALSE, FALSE, '', 2);
                    EnterCell(NN, 19, FORMAT(Quantity), FALSE, FALSE, '0.000', 0);
                    //>>20Mar2017 GQty
                    GQty += Quantity;
                    //<<20Mar2017 GQty

                    EnterCell(NN, 20, FORMAT("Quantity Received"), FALSE, FALSE, '0.000', 0);
                    //>>20Mar2017 GRQty
                    GRQty += "Quantity Received";
                    //<<20Mar2017 GRQty

                    EnterCell(NN, 21, FORMAT("Outstanding Quantity"), FALSE, FALSE, '0.000', 0);
                    //>>20Mar2017 GOQty
                    GOQty += "Outstanding Quantity";
                    //<<20Mar2017 GOQty

                    EnterCell(NN, 22, FORMAT("Unit of Measure"), FALSE, FALSE, '', 1);
                    EnterCell(NN, 23, FORMAT("Direct Unit Cost"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 24, FORMAT("Line Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 25, FORMAT(BlanketExcisePer), FALSE, FALSE, '0.00', 0);
                    EnterCell(NN, 26, FORMAT(/*"Excise Amount"*/0), FALSE, FALSE, '#,###0.00', 0);
                    //>>20Mar2017 GExcise
                    GExcise += 0;// "Excise Amount";
                    //<<20Mar2017 GExcise

                    EnterCell(NN, 27, FORMAT("Tax Amount"), FALSE, FALSE, '#,###0.00', 0);
                    //>>20Mar2017 GVat
                    GVat += "Tax Amount";
                    //<<20Mar2017 GVat

                    EnterCell(NN, 28, FORMAT(/*"Charges To Vendor"*/0), FALSE, FALSE, '#,###0.00', 0);
                    //>>20Mar2017 GCharge
                    GCharge += 0;// "Charges To Vendor";
                    //<<20Mar2017 GCharge

                    EnterCell(NN, 29, FORMAT("Amount To Vendor"), FALSE, FALSE, '#,###0.00', 0);
                    //>>20Mar2017 GTotal
                    GTotal += "Amount To Vendor";
                    //<<20Mar2017 GTotal

                    EnterCell(NN, 30, grnNO, FALSE, FALSE, '', 1);
                    EnterCell(NN, 31, FORMAT(GRNDATE), FALSE, FALSE, '', 2);
                    EnterCell(NN, 32, RecLoc.Name, FALSE, FALSE, '', 1);
                    EnterCell(NN, 35, InvNo, FALSE, FALSE, '', 1);
                    EnterCell(NN, 36, FORMAT(InvDate), FALSE, FALSE, '', 2);
                    EnterCell(NN, 37, PayVoucher, FALSE, FALSE, '', 1);
                    EnterCell(NN, 38, FORMAT(PayDate), FALSE, FALSE, '', 2);
                    EnterCell(NN, 53, GateEntryNo, FALSE, FALSE, '', 2);//RSPLSUM 11Jan21
                    EnterCell(NN, 54, PL19."Vendor Item No.", FALSE, FALSE, '@', 2); //Fahim 20-12-21
                    EnterCell(NN, 55, PL19."Item Category Code", FALSE, FALSE, '@', 2); //Fahim 11-02-22

                    NN += 1;

                    //<<20Mar2017 PurchaseLineBody
                    //>>2
                    //Purchase Line, Body (1) - OnPreSection()
                    /*
                    IF ExportToExcel  THEN BEGIN
                    row+=1;

                    XlsWorkSheet.Range('O' + FORMAT(row)).Value :=FORMAT("Purchase Line"."Document No.");
                    XlsWorkSheet.Range('P' + FORMAT(row)).Value :=FORMAT("Purchase Line"."Order Date");
                    XlsWorkSheet.Range('Q' + FORMAT(row)).Value :=FORMAT("Purchase Line".Quantity);
                    XlsWorkSheet.Range('R' + FORMAT(row)).Value := FORMAT("Purchase Line"."Quantity Received");
                    XlsWorkSheet.Range('S' + FORMAT(row)).Value :=FORMAT("Purchase Line"."Outstanding Quantity");
                    XlsWorkSheet.Range('T' + FORMAT(row)).Value := FORMAT("Purchase Line"."Unit of Measure");
                    XlsWorkSheet.Range('U' + FORMAT(row)).Value := FORMAT("Purchase Line"."Direct Unit Cost");
                    XlsWorkSheet.Range('V' + FORMAT(row)).Value := FORMAT("Purchase Line"."Line Amount");
                    XlsWorkSheet.Range('W' + FORMAT(row)).Value :=FORMAT(BlanketExcisePer);
                    XlsWorkSheet.Range('X' + FORMAT(row)).Value := FORMAT("Purchase Line"."Excise Amount");
                    XlsWorkSheet.Range('Y' + FORMAT(row)).Value :=FORMAT("Purchase Line"."Tax Amount");
                    XlsWorkSheet.Range('Z' + FORMAT(row)).Value := FORMAT("Purchase Line"."Charges To Vendor");
                    XlsWorkSheet.Range('AA' + FORMAT(row)).Value :=FORMAT("Purchase Line"."Amount To Vendor");
                    XlsWorkSheet.Range('AB' + FORMAT(row)).Value := FORMAT(grnNO);
                    XlsWorkSheet.Range('AC' + FORMAT(row)).Value := FORMAT(GRNDATE);
                    XlsWorkSheet.Range('AD' + FORMAT(row)).Value := FORMAT(RecLoc.Name);
                    XlsWorkSheet.Range('AG' + FORMAT(row)).Value := FORMAT(InvNo);
                    XlsWorkSheet.Range('AH' + FORMAT(row)).Value := FORMAT(InvDate);
                    XlsWorkSheet.Range('AI' + FORMAT(row)).Value := FORMAT(PayVoucher);
                    XlsWorkSheet.Range('AJ' + FORMAT(row)).Value := FORMAT(PayDate);
                    END;

              //<<2
              */
                    //>>3

                    /*
                    //Purchase Line, Footer (2) - OnPreSection()
                          IF ExportToExcel  THEN BEGIN
                          row+=1;
                    
                          XlsWorkSheet.Range('Q' + FORMAT(row)).Value :=FORMAT("Purchase Line".Quantity);
                          XlsWorkSheet.Range('R' + FORMAT(row)).Value := FORMAT("Purchase Line"."Quantity Received");
                          XlsWorkSheet.Range('S' + FORMAT(row)).Value :=FORMAT("Purchase Line"."Outstanding Quantity");
                          XlsWorkSheet.Range('X' + FORMAT(row)).Value := FORMAT("Purchase Line"."Excise Amount");
                          XlsWorkSheet.Range('Y' + FORMAT(row)).Value :=FORMAT("Purchase Line"."Tax Amount");
                          XlsWorkSheet.Range('Z' + FORMAT(row)).Value := FORMAT("Purchase Line"."Charges To Vendor");
                          XlsWorkSheet.Range('AA' + FORMAT(row)).Value :=FORMAT("Purchase Line"."Amount To Vendor");
                            XlsWorkSheet.Range('A'+FORMAT(row)+':'+'AA'+FORMAT(row)).Font.Bold := 'True';
                          END;
                    //<<3
                    *///Commented Automation 20Mar2017


                    //>>20Mar2017 PurchaseGroupFooter
                    IF (NNCOUNT <> 0) AND (NNCOUNT = TCOUNT) THEN BEGIN

                        EnterCell(NN, 19, FORMAT(GQty), TRUE, FALSE, '0.000', 0);
                        GQty := 0;
                        EnterCell(NN, 20, FORMAT(GRQty), TRUE, FALSE, '0.000', 0);
                        GRQty := 0;
                        EnterCell(NN, 21, FORMAT(GOQty), TRUE, FALSE, '0.000', 0);
                        GOQty := 0;
                        EnterCell(NN, 26, FORMAT(GExcise), TRUE, FALSE, '#,###0.00', 0);
                        GExcise := 0;
                        EnterCell(NN, 27, FORMAT(GVat), TRUE, FALSE, '#,###0.00', 0);
                        GVat := 0;
                        EnterCell(NN, 28, FORMAT(GCharge), TRUE, FALSE, '#,###0.00', 0);
                        GCharge := 0;
                        EnterCell(NN, 29, FORMAT(GTotal), TRUE, FALSE, '#,###0.00', 0);
                        GTotal := 0;

                        NN += 1;
                        TCOUNT := 0;

                    END;

                    //<<20Mar2017 PurchaseGroupFooter

                end;

                trigger OnPreDataItem()
                begin

                    NNCOUNT := COUNT; //20Mar2017
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>02Jul2018
                IF SPCode02 <> '' THEN BEGIN
                    PH02.RESET;
                    PH02.SETRANGE("Document Type", "Document Type");
                    PH02.SETRANGE("No.", "Document No.");
                    IF PH02.FINDFIRST THEN
                        IF PH02."Purchaser Code" <> SPCode02 THEN
                            CurrReport.SKIP;
                END;
                //<<02Jul2018

                //>>1
                IF ASONDATE <> 0D THEN BEGIN
                    BlanketHeader.RESET;
                    BlanketHeader.SETRANGE("No.", "Blanket Purchase Line"."Document No.");
                    IF BlanketHeader.FINDFIRST THEN
                        IF BlanketHeader.Closed THEN
                            CurrReport.SKIP;
                END;

                Companyinfo.GET;
                IF RecVend.GET("Blanket Purchase Line"."Buy-from Vendor No.") THEN;
                CLEAR(BlanketExcisePer);
                // IF "Blanket Purchase Line"."Excise Amount" <> 0 THEN BEGIN
                /* RecEPS.RESET;
                //RecEPS.SETRANGE(RecEPS."Excise Bus. Posting Group", "Blanket Purchase Line"."Excise Bus. Posting Group");
                //RecEPS.SETRANGE(RecEPS."Excise Prod. Posting Group", "Blanket Purchase Line"."Excise Prod. Posting Group");
                IF RecEPS.FINDFIRST THEN
                    BlanketExcisePer := RecEPS."BED %"; */
                // END;

                IF "Order Date" = 0D THEN
                    CurrReport.SKIP;

                IF RecLoc.GET("Location Code") THEN;

                PurchHdr.RESET;
                PurchHdr.SETRANGE("Document Type", PurchHdr."Document Type"::"Blanket Order");
                PurchHdr.SETRANGE("No.", "Blanket Purchase Line"."Document No.");
                IF PurchHdr.FINDFIRST THEN BEGIN
                    VendPOGrp := PurchHdr."Vendor Posting Group";
                END;
                //<<1

                //>>2
                //Blanket Purchase Line, Header (1) - OnPreSection()
                /*
                IF ExportToExcel AND (CurrReport.PAGENO = 1) THEN BEGIN
                  row+=1;
                  XlsWorkSheet.Range('A' + FORMAT(row)).Value := FORMAT(COMPANYNAME);
                  XlsWorkSheet.Range('A'+FORMAT(row)+':'+'A'+FORMAT(row)).Font.Bold := 'True';
                    row+=1;
                  XlsWorkSheet.Range('A' + FORMAT(row)).Value := 'PURCHASE ORDER STATEMENT (Detailed)';
                  XlsWorkSheet.Range('A'+FORMAT(row)+':'+'A'+FORMAT(row)).Font.Bold := 'True';
                    row+=1;
                  IF Requestformdate <> '' THEN
                  XlsWorkSheet.Range('A' + FORMAT(row)).Value := 'Date Filter: '+GETFILTER("Order Date")
                  ELSE
                  XlsWorkSheet.Range('A' + FORMAT(row)).Value := 'Date Filter: '+FORMAT(ASONDATE);
                
                  XlsWorkSheet.Range('A'+FORMAT(row)+':'+'A'+FORMAT(row)).Font.Bold := 'True';
                    row+=1;
                  XlsWorkSheet.Range('A' + FORMAT(row)).Value := 'User ID :'+FORMAT(USERID);
                
                  XlsWorkSheet.Range('A'+FORMAT(row)+':'+'A'+FORMAT(row)).Font.Bold := 'True';
                    row+=1;
                
                  XlsWorkSheet.Range('A'+FORMAT(row)+':'+'L'+FORMAT(row)).Font.Bold := 'True';
                END;
                      IF ExportToExcel AND (CurrReport.PAGENO = 1) THEN BEGIN
                      row+=1;
                
                      XlsWorkSheet.Range('A' + FORMAT(row)).Value := 'Blanket Order No.';
                      XlsWorkSheet.Range('B' + FORMAT(row)).Value := 'Vendor Code';
                      XlsWorkSheet.Range('C' + FORMAT(row)).Value := 'Vendor Name';
                      XlsWorkSheet.Range('D' + FORMAT(row)).Value := 'Item Code';
                      XlsWorkSheet.Range('E' + FORMAT(row)).Value :='Item Description';
                      XlsWorkSheet.Range('F' + FORMAT(row)).Value :='Blanket Order Date';
                      XlsWorkSheet.Range('G' + FORMAT(row)).Value :='Qty';
                      XlsWorkSheet.Range('H' + FORMAT(row)).Value :='Rate';
                      XlsWorkSheet.Range('I' + FORMAT(row)).Value :='Basic Amount';
                      XlsWorkSheet.Range('J' + FORMAT(row)).Value := 'Excise %';
                      XlsWorkSheet.Range('K' + FORMAT(row)).Value := 'Excise';
                      XlsWorkSheet.Range('L' + FORMAT(row)).Value := 'Taxes';
                      XlsWorkSheet.Range('M' + FORMAT(row)).Value := 'Charges';
                      XlsWorkSheet.Range('N' + FORMAT(row)).Value := 'Total Amount';
                      XlsWorkSheet.Range('O' + FORMAT(row)).Value :='Make Order No.';
                      XlsWorkSheet.Range('P' + FORMAT(row)).Value :='Make Order Date';
                      XlsWorkSheet.Range('Q' + FORMAT(row)).Value :='Make Order Qty';
                      XlsWorkSheet.Range('R' + FORMAT(row)).Value := 'GRN Qty';
                      XlsWorkSheet.Range('S' + FORMAT(row)).Value :='Balance Qty';
                      XlsWorkSheet.Range('T' + FORMAT(row)).Value := 'Unit of Measure';
                      XlsWorkSheet.Range('U' + FORMAT(row)).Value := 'Basic Rate';
                      XlsWorkSheet.Range('V' + FORMAT(row)).Value := 'Basic Amount';
                      XlsWorkSheet.Range('W' + FORMAT(row)).Value :='Excise %';
                      XlsWorkSheet.Range('X' + FORMAT(row)).Value := 'Excise Amount';
                      XlsWorkSheet.Range('Y' + FORMAT(row)).Value :='Vat Amount';
                      XlsWorkSheet.Range('Z' + FORMAT(row)).Value := 'Charges';
                      XlsWorkSheet.Range('AA' + FORMAT(row)).Value :='Total Amount';
                      XlsWorkSheet.Range('AB' + FORMAT(row)).Value := 'GRN No.';
                      XlsWorkSheet.Range('AC' + FORMAT(row)).Value := 'GRN Date';
                      XlsWorkSheet.Range('AD' + FORMAT(row)).Value := 'Receiving Location';
                      XlsWorkSheet.Range('AE' + FORMAT(row)).Value := 'Closed';
                      XlsWorkSheet.Range('AF' + FORMAT(row)).Value := 'Vendor Item Name';
                      XlsWorkSheet.Range('AG' + FORMAT(row)).Value := 'Invoice No.';
                      XlsWorkSheet.Range('AH' + FORMAT(row)).Value := 'Invoice Date';
                      XlsWorkSheet.Range('AI' + FORMAT(row)).Value := 'Payment Voucher No.';
                      XlsWorkSheet.Range('AJ' + FORMAT(row)).Value := 'Payment Date';
                      XlsWorkSheet.Range('AK' + FORMAT(row)).Value := 'Amount To Vendor';
                      XlsWorkSheet.Range('AL' + FORMAT(row)).Value := 'Vendor Posting Group';
                      XlsWorkSheet.Range('A'+FORMAT(row)+':'+'AL'+FORMAT(row)).Font.Bold := 'True';
                      END;
                //<<2
                *///Commented Automation 20Mar2017

                //>>25May2018
                CLEAR(InitialQty);
                CLEAR(InitialRate);
                PHArc.RESET;
                PHArc.SETRANGE("Document Type", "Document Type");
                PHArc.SETRANGE("Document No.", "Document No.");
                PHArc.SETRANGE(Type, Type);
                PHArc.SETRANGE("No.", "No.");
                PHArc.SETRANGE("Line No.", "Line No.");
                IF PHArc.FINDFIRST THEN BEGIN
                    InitialQty := PHArc.Quantity;
                    InitialRate := PHArc."Direct Unit Cost";
                END;
                //<<25May2018

                //>>3
                //Blanket Purchase Line, Body (2) - OnPreSection()
                BlanketHeader.RESET;
                BlanketHeader.SETRANGE("No.", "Blanket Purchase Line"."Document No.");
                IF BlanketHeader.FINDFIRST THEN;

                //>>20Mar2017 BlanketDataBody
                IF NOT Summary THEN BEGIN
                    EnterCell(NN, 1, "Document No.", FALSE, FALSE, '@', 1);
                    EnterCell(NN, 2, "Buy-from Vendor No.", FALSE, FALSE, '@', 1);
                    EnterCell(NN, 3, RecVend.Name, FALSE, FALSE, '@', 1);
                    EnterCell(NN, 4, "No.", FALSE, FALSE, '@', 1);
                    EnterCell(NN, 5, Description, FALSE, FALSE, '@', 1);
                    EnterCell(NN, 6, FORMAT("Order Date"), FALSE, FALSE, '', 2);
                    //>>25May2018
                    IF InitialQty <> 0 THEN
                        IF InitialQty <> Quantity THEN
                            EnterCell(NN, 7, FORMAT(InitialQty), FALSE, FALSE, '0.000', 0)
                        ELSE
                            EnterCell(NN, 7, '', FALSE, FALSE, '', 1);
                    //<<25May2018
                    EnterCell(NN, 8, FORMAT(Quantity), FALSE, FALSE, '0.000', 0);
                    //>>25May2018
                    IF InitialRate <> 0 THEN
                        IF InitialRate <> "Direct Unit Cost" THEN
                            EnterCell(NN, 9, FORMAT(InitialRate), FALSE, FALSE, '#,###0.00', 0)
                        ELSE
                            EnterCell(NN, 9, '', FALSE, FALSE, '', 1);
                    //<<25May2018
                    EnterCell(NN, 10, FORMAT("Direct Unit Cost"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 11, FORMAT("Line Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 12, FORMAT(BlanketExcisePer), FALSE, FALSE, '0.00', 0);
                    EnterCell(NN, 13, FORMAT(/*"Excise Amount"*/0), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 14, FORMAT("Tax Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 15, FORMAT(/*"Charges To Vendor"*/0), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 16, FORMAT("Amount To Vendor"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 33, FORMAT(BlanketHeader.Closed), FALSE, FALSE, '@', 1);
                    EnterCell(NN, 34, "Vendor Item No.", FALSE, FALSE, '@', 1);
                    EnterCell(NN, 39, FORMAT("Amount To Vendor"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 40, VendPOGrp, FALSE, FALSE, '@', 1);

                    //>>02Jul2018
                    CLEAR(PayDesc);
                    CLEAR(SPName);

                    PH02.RESET;
                    PH02.SETRANGE("Document Type", "Document Type");
                    PH02.SETRANGE("No.", "Document No.");
                    IF PH02.FINDFIRST THEN BEGIN
                        PayTerm.RESET;
                        IF PayTerm.GET(PH02."Payment Terms Code") THEN
                            PayDesc := PayTerm.Description;

                        SP02.RESET;
                        IF SP02.GET(PH02."Purchaser Code") THEN
                            SPName := PH02."Purchaser Code" + ' - ' + SP02.Name;
                    END;
                    //<<02Jul2018
                    EnterCell(NN, 41, SPName, FALSE, FALSE, '@', 1);//02Jul2018
                    EnterCell(NN, 42, PayDesc, FALSE, FALSE, '@', 1);//02Jul2018
                    EnterCell(NN, 43, "Blanket Purchase Line"."Item Category Code", FALSE, FALSE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(NN, 44, BlanketHeader."Purchaser Code", FALSE, FALSE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(NN, 45, FORMAT(BlanketHeader."Department Code"), FALSE, FALSE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(NN, 46, BlanketHeader."Deal Sheet No.", FALSE, FALSE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(NN, 47, "Blanket Purchase Line"."MR No", FALSE, FALSE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(NN, 48, FORMAT("Blanket Purchase Line"."MR Date"), FALSE, FALSE, '', 2);//RSPLSUM 06Aug2020
                    EnterCell(NN, 49, FORMAT("Blanket Purchase Line"."MR Quantity"), FALSE, FALSE, '0.000', 0);//RSPLSUM 06Aug2020
                    EnterCell(NN, 50, FORMAT(BlanketHeader."Deal Sheet Date"), FALSE, FALSE, '', 2);//RSPLSUM 10Aug2020
                    EnterCell(NN, 51, FORMAT("Blanket Purchase Line"."Landed Cost"), FALSE, FALSE, '#,###0.00', 0);//RSPLSUM 10Aug2020
                    EnterCell(NN, 52, FORMAT("Blanket Purchase Line"."Cost Avoidance"), FALSE, FALSE, '#,###0.00', 0);//RSPLSUM 10Aug2020
                    EnterCell(NN, 53, "Blanket Purchase Line"."Vendor Item No.", FALSE, FALSE, '@', 1); //Fahim 20-12-21
                    EnterCell(NN, 54, "Blanket Purchase Line"."Item Category Code", FALSE, FALSE, '@', 1); //Fahim 11-02-22

                    NN += 1;
                END ELSE BEGIN
                    EnterCell(NN, 1, "Document No.", FALSE, FALSE, '@', 1);
                    EnterCell(NN, 2, "Buy-from Vendor No.", FALSE, FALSE, '@', 1);
                    EnterCell(NN, 3, RecVend.Name, FALSE, FALSE, '@', 1);
                    EnterCell(NN, 4, "No.", FALSE, FALSE, '@', 1);
                    EnterCell(NN, 5, Description, FALSE, FALSE, '@', 1);
                    EnterCell(NN, 6, FORMAT("Order Date"), FALSE, FALSE, '', 2);
                    EnterCell(NN, 7, FORMAT(Quantity), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 8, "Unit of Measure", FALSE, FALSE, '', 1);
                    EnterCell(NN, 9, FORMAT("Direct Unit Cost"), FALSE, FALSE, '#,#0.00', 0);
                    //>>19Oct2018
                    CLEAR(OrdQty19);
                    CLEAR(GRNQty19);
                    CLEAR(OutQty19);
                    PL19.RESET;
                    PL19.SETCURRENTKEY("Document Type", "Blanket Order No.", "Blanket Order Line No.");
                    PL19.SETRANGE("Document Type", PL19."Document Type"::Order);
                    PL19.SETRANGE("Blanket Order No.", "Document No.");
                    PL19.SETRANGE("Blanket Order Line No.", "Line No.");
                    IF PL19.FINDFIRST THEN BEGIN
                        PL19.CALCSUMS(Quantity, "Quantity Received", "Outstanding Quantity");
                        OrdQty19 := PL19.Quantity;
                        GRNQty19 := PL19."Quantity Received";
                        //OutQty19 := PL19."Outstanding Quantity";
                    END;
                    OutQty19 := Quantity - GRNQty19;
                    //<<19Oct2018
                    EnterCell(NN, 10, FORMAT(OrdQty19), FALSE, FALSE, '0.000', 0);//OrderQty
                    EnterCell(NN, 11, FORMAT(GRNQty19), FALSE, FALSE, '0.000', 0);//GRNQty
                    EnterCell(NN, 12, FORMAT(OutQty19), FALSE, FALSE, '0.000', 0);//BalanceQty
                    EnterCell(NN, 13, "Blanket Purchase Line"."MR No", FALSE, FALSE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(NN, 14, FORMAT("Blanket Purchase Line"."MR Date"), FALSE, FALSE, '', 2);//RSPLSUM 06Aug2020
                    EnterCell(NN, 15, FORMAT("Blanket Purchase Line"."MR Quantity"), FALSE, FALSE, '0.000', 0);//RSPLSUM 06Aug2020
                    EnterCell(NN, 16, FORMAT(BlanketHeader."Deal Sheet Date"), FALSE, FALSE, '', 2);//RSPLSUM 10Aug2020
                    EnterCell(NN, 17, FORMAT("Blanket Purchase Line"."Landed Cost"), FALSE, FALSE, '#,###0.00', 0);//RSPLSUM 10Aug2020
                    EnterCell(NN, 18, FORMAT("Blanket Purchase Line"."Cost Avoidance"), FALSE, FALSE, '#,###0.00', 0);//RSPLSUM 10Aug2020
                    EnterCell(NN, 19, "Blanket Purchase Line"."Vendor Item No.", FALSE, FALSE, '@', 0); //Fahim 20-12-21
                    EnterCell(NN, 20, "Blanket Purchase Line"."Item Category Code", FALSE, FALSE, '@', 0); //Fahim 11-02-22
                    NN += 1;

                END;

                //<<20Mar2017 BlanketDataBody
                /*
                      IF ExportToExcel  THEN BEGIN
                      row+=1;
                
                      XlsWorkSheet.Range('A' + FORMAT(row)).Value := FORMAT("Document No.");;
                      XlsWorkSheet.Range('B' + FORMAT(row)).Value := FORMAT("Blanket Purchase Line"."Buy-from Vendor No.");
                      XlsWorkSheet.Range('C' + FORMAT(row)).Value := FORMAT(RecVend.Name);;
                      XlsWorkSheet.Range('D' + FORMAT(row)).Value := FORMAT("Blanket Purchase Line"."No.");
                      XlsWorkSheet.Range('E' + FORMAT(row)).Value :=FORMAT("Blanket Purchase Line".Description);
                      XlsWorkSheet.Range('F' + FORMAT(row)).Value :=FORMAT("Blanket Purchase Line"."Order Date");
                      XlsWorkSheet.Range('G' + FORMAT(row)).Value :=FORMAT("Blanket Purchase Line".Quantity);
                      XlsWorkSheet.Range('H' + FORMAT(row)).Value :=FORMAT("Blanket Purchase Line"."Direct Unit Cost");
                      XlsWorkSheet.Range('I' + FORMAT(row)).Value := FORMAT("Blanket Purchase Line"."Line Amount");
                      XlsWorkSheet.Range('J' + FORMAT(row)).Value := FORMAT(BlanketExcisePer);
                      XlsWorkSheet.Range('K' + FORMAT(row)).Value := FORMAT("Blanket Purchase Line"."Excise Amount");
                      XlsWorkSheet.Range('L' + FORMAT(row)).Value := FORMAT("Blanket Purchase Line"."Tax Amount");
                      XlsWorkSheet.Range('M' + FORMAT(row)).Value := FORMAT("Blanket Purchase Line"."Charges To Vendor");
                      XlsWorkSheet.Range('N' + FORMAT(row)).Value := FORMAT("Blanket Purchase Line"."Amount To Vendor");
                      XlsWorkSheet.Range('AE' + FORMAT(row)).Value :=FORMAT(BlanketHeader.Closed) ;
                      XlsWorkSheet.Range('AF' + FORMAT(row)).Value :=FORMAT("Blanket Purchase Line"."Vendor Item No.");
                      XlsWorkSheet.Range('AK' + FORMAT(row)).Value :=FORMAT("Blanket Purchase Line"."Amount To Vendor");
                      XlsWorkSheet.Range('AL' + FORMAT(row)).Value :=FORMAT(VendPOGrp);
                      END;
                //<<3
                *///Commented Automation 20Mar2017

            end;

            trigger OnPreDataItem()
            begin

                IF ASONDATE <> 0D THEN
                    SETFILTER("Order Date", '..%1', ASONDATE);

                //>>20Mar2017 ReportHeader
                IF NOT Summary THEN BEGIN
                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '@', 1);
                    EnterCell(1, 53, FORMAT(TODAY, 0, 4), TRUE, FALSE, '@', 1);

                    EnterCell(2, 1, 'PURCHASE ORDER STATEMENT (Detailed)', TRUE, FALSE, '@', 1);
                    EnterCell(2, 53, USERID, TRUE, FALSE, '@', 1);

                    IF Requestformdate <> '' THEN
                        EnterCell(5, 1, 'Date Filter: ' + FORMAT(Requestformdate), TRUE, TRUE, '@', 1)
                    ELSE
                        EnterCell(5, 1, 'Date Filter: ' + FORMAT(ASONDATE), TRUE, TRUE, '@', 1);

                    EnterCell(5, 2, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 3, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 4, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 5, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 6, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 7, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 8, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 9, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 10, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 11, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 12, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 13, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 14, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 15, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 16, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 17, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 18, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 19, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 20, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 21, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 22, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 23, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 24, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 25, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 26, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 27, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 28, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 29, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 30, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 31, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 32, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 33, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 34, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 35, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 36, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 37, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 38, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 39, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 40, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 41, '', TRUE, TRUE, '@', 1);//02Jul2018
                    EnterCell(5, 42, '', TRUE, TRUE, '@', 1);//02Jul2018
                    EnterCell(5, 43, '', TRUE, TRUE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(5, 44, '', TRUE, TRUE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(5, 45, '', TRUE, TRUE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(5, 46, '', TRUE, TRUE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(5, 47, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(5, 48, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(5, 49, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(5, 50, '', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(5, 51, '', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(5, 52, '', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(5, 53, '', TRUE, TRUE, '@', 1);//RSPLSUM 11jan21
                    EnterCell(5, 54, '', TRUE, TRUE, '@', 1);//Fahim 20-12-2021
                    EnterCell(5, 55, '', TRUE, TRUE, '@', 1);//Fahim 11-02-2022

                    EnterCell(6, 1, 'Blanket Order No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 2, 'Vendor Code', TRUE, TRUE, '@', 1);
                    EnterCell(6, 3, 'Vendor Name', TRUE, TRUE, '@', 1);
                    EnterCell(6, 4, 'Item Code', TRUE, TRUE, '@', 1);
                    EnterCell(6, 5, 'Item Description', TRUE, TRUE, '@', 1);
                    EnterCell(6, 6, 'Blanket Order Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 7, 'Initial Qty', TRUE, TRUE, '@', 1);
                    EnterCell(6, 8, 'Revised Qty', TRUE, TRUE, '@', 1);//25May2018
                    EnterCell(6, 9, 'Initial Rate', TRUE, TRUE, '@', 1);
                    EnterCell(6, 10, 'Revised Rate', TRUE, TRUE, '@', 1);//25May2018
                    EnterCell(6, 11, 'Basic Amount', TRUE, TRUE, '@', 1);
                    EnterCell(6, 12, 'Excise %', TRUE, TRUE, '@', 1);
                    EnterCell(6, 13, 'Excise', TRUE, TRUE, '@', 1);
                    EnterCell(6, 14, 'Taxes', TRUE, TRUE, '@', 1);
                    EnterCell(6, 15, 'Charges', TRUE, TRUE, '@', 1);
                    EnterCell(6, 16, 'Total Amount', TRUE, TRUE, '@', 1);
                    EnterCell(6, 17, 'Make Order No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 18, 'Make Order Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 19, 'Make Order Qty', TRUE, TRUE, '@', 1);
                    EnterCell(6, 20, 'GRN Qty', TRUE, TRUE, '@', 1);
                    EnterCell(6, 21, 'Balance Qty', TRUE, TRUE, '@', 1);
                    EnterCell(6, 22, 'Unit of Measure', TRUE, TRUE, '@', 1);
                    EnterCell(6, 23, 'Basic Rate', TRUE, TRUE, '@', 1);
                    EnterCell(6, 24, 'Basic Amount', TRUE, TRUE, '@', 1);
                    EnterCell(6, 25, 'Excise %', TRUE, TRUE, '@', 1);
                    EnterCell(6, 26, 'Excise Amount', TRUE, TRUE, '@', 1);
                    EnterCell(6, 27, 'Vat Amount', TRUE, TRUE, '@', 1);
                    EnterCell(6, 28, 'Charges', TRUE, TRUE, '@', 1);
                    EnterCell(6, 29, 'Total Amount', TRUE, TRUE, '@', 1);
                    EnterCell(6, 30, 'GRN No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 31, 'GRN Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 32, 'Receiving Location', TRUE, TRUE, '@', 1);
                    EnterCell(6, 33, 'Closed', TRUE, TRUE, '@', 1);
                    EnterCell(6, 34, 'Vendor Item Name', TRUE, TRUE, '@', 1);
                    EnterCell(6, 35, 'Invoice No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 36, 'Invoice Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 37, 'Payment Voucher No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 38, 'Payment Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 39, 'Amount To Vendor', TRUE, TRUE, '@', 1);
                    EnterCell(6, 40, 'Vendor Posting Group', TRUE, TRUE, '@', 1);
                    EnterCell(6, 41, 'GT Code', TRUE, TRUE, '@', 1);//02Jul2018
                    EnterCell(6, 42, 'Payment Description', TRUE, TRUE, '@', 1);//02Jul2018
                    EnterCell(6, 43, 'Item Category', TRUE, TRUE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(6, 44, 'Requisitioner', TRUE, TRUE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(6, 45, 'Department', TRUE, TRUE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(6, 46, 'Deal Sheet No.', TRUE, TRUE, '@', 1);//RSPLSUM 15Apr2020
                    EnterCell(6, 47, 'MR No.', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(6, 48, 'MR Date', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(6, 49, 'MR Quantity', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(6, 50, 'Deal Sheet Date', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(6, 51, 'Landed Cost', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(6, 52, 'Cost Avoidance', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(6, 53, 'Gate Entry Number', TRUE, TRUE, '@', 1);//RSPLSUM 11Jan21
                    EnterCell(6, 54, 'Vendor Item No.', TRUE, TRUE, '@', 1);//Fahim 20-12-21
                    EnterCell(6, 55, 'Item Category', TRUE, TRUE, '@', 1);//Fahim 11-02-22
                    NN := 7;
                END ELSE BEGIN

                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '@', 1);
                    EnterCell(1, 18, FORMAT(TODAY, 0, 4), TRUE, FALSE, '@', 1);

                    EnterCell(2, 1, 'PURCHASE ORDER STATEMENT (Summary)', TRUE, FALSE, '@', 1);
                    EnterCell(2, 18, USERID, TRUE, FALSE, '@', 1);

                    IF Requestformdate <> '' THEN
                        EnterCell(5, 1, 'Date Filter: ' + FORMAT(Requestformdate), TRUE, TRUE, '@', 1)
                    ELSE
                        EnterCell(5, 1, 'Date Filter: ' + FORMAT(ASONDATE), TRUE, TRUE, '@', 1);

                    EnterCell(5, 2, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 3, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 4, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 5, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 6, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 7, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 8, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 9, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 10, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 11, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 12, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 13, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(5, 14, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(5, 15, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(5, 16, '', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(5, 17, '', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(5, 18, '', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(5, 19, '', TRUE, TRUE, '@', 1);//Fahim 20-12-21
                    EnterCell(5, 20, '', TRUE, TRUE, '@', 1);//Fahim 11-02-22

                    EnterCell(6, 1, 'Blanket Order No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 2, 'Vendor Code', TRUE, TRUE, '@', 1);
                    EnterCell(6, 3, 'Vendor Name', TRUE, TRUE, '@', 1);
                    EnterCell(6, 4, 'Item Code', TRUE, TRUE, '@', 1);
                    EnterCell(6, 5, 'Item Description', TRUE, TRUE, '@', 1);
                    EnterCell(6, 6, 'Blanket Order Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 7, 'Blanket Order Quantity', TRUE, TRUE, '@', 1);
                    EnterCell(6, 8, 'UOM', TRUE, TRUE, '@', 1);
                    EnterCell(6, 9, 'Rate', TRUE, TRUE, '@', 1);
                    EnterCell(6, 10, 'Make Order Qty', TRUE, TRUE, '@', 1);
                    EnterCell(6, 11, 'GRN Qty', TRUE, TRUE, '@', 1);
                    EnterCell(6, 12, 'Balance Qty', TRUE, TRUE, '@', 1);
                    EnterCell(6, 13, 'MR No.', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(6, 14, 'MR Date', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(6, 15, 'MR Quantity', TRUE, TRUE, '@', 1);//RSPLSUM 06Aug2020
                    EnterCell(6, 16, 'Deal Sheet Date', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(6, 17, 'Landed Cost', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(6, 18, 'Cost Avoidance', TRUE, TRUE, '@', 1);//RSPLSUM 10Aug2020
                    EnterCell(6, 19, 'Vendor Item No.', TRUE, TRUE, '@', 1);//Fahim 20-12-21
                    EnterCell(6, 20, 'Item Category', TRUE, TRUE, '@', 1);//Fahim 11-02-22

                    NN := 7;
                END;
                //<<20Mar2017 ReportHeader
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print Summary"; Summary)
                {
                    ApplicationArea = all;
                }
                field("As on Date"; ASONDATE)
                {
                    ApplicationArea = all;
                }
                field("GT Code"; SPCode02)
                {
                    ApplicationArea = all;
                    TableRelation = "Salesperson/Purchaser".Code;
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
        /*
        //>>1
        
        IF ExportToExcel THEN
          XlsApp.Visible:=TRUE;
        
        //<<1
        *///Commented Automation 20Mar2017


        //>>20Mar2017 RB-N
        /*   ExcBuffer.CreateBook('', 'Blanket PO');
          ExcBuffer.CreateBookAndOpenExcel('', 'Blanket PO', '', '', USERID);
          ExcBuffer.GiveUserControl; */

        ExcBuffer.CreateNewBook('Blanket PO');
        ExcBuffer.WriteSheet('Blanket PO', CompanyName, UserId);
        ExcBuffer.CloseBook();
        ExcBuffer.SetFriendlyFilename(StrSubstNo('Blanket PO', CurrentDateTime, UserId));
        ExcBuffer.OpenExcel();
        //<<20Mar2017

    end;

    trigger OnPreReport()
    begin
        //>>1
        /*
        ExportToExcel:=TRUE;
        IF ExportToExcel THEN BEGIN
          IF NOT CREATE(XlsApp) THEN
            ERROR('Excel Not Created');
          xlsWorkBook := XlsApp.Workbooks.Add(-4167);
          XlsWorkSheet := XlsApp.ActiveSheet;
          XlsWorkSheet.Name := 'Blanket PO';
        END;
        *///Commented Automation 20Mar2017



        Requestformdate := "Blanket Purchase Line".GETFILTER("Blanket Purchase Line"."Order Date");
        IF (ASONDATE <> 0D) AND (Requestformdate <> '') THEN
            ERROR('Please Enter the Date Either in Range or As on date');

        //<<1

    end;

    var
        RecVend: Record 23;
        ExportToExcel: Boolean;
        row: Integer;
        Companyinfo: Record 79;
        // RecEPS: Record 13711;
        BlanketExcisePer: Decimal;
        RecPRL: Record 121;
        grnNO: Code[20];
        GRNDATE: Date;
        BlanketHeader: Record 38;
        ASONDATE: Date;
        Requestformdate: Text[100];
        RecLoc: Record 14;
        InvNo: Code[20];
        InvDate: Date;
        PayVoucher: Code[20];
        PayDate: Date;
        recILE: Record 32;
        recVal: Record 5802;
        recVLE: Record 25;
        recDVLE: Record 380;
        PurchHdr: Record 38;
        VendPOGrp: Code[20];
        "-----20Mar2017": Integer;
        ExcBuffer: Record 370 temporary;
        NN: Integer;
        TCOUNT: Integer;
        NNCOUNT: Integer;
        GQty: Decimal;
        GRQty: Decimal;
        GOQty: Decimal;
        GExcise: Decimal;
        GVat: Decimal;
        GCharge: Decimal;
        GTotal: Decimal;
        "------------25May2018": Integer;
        PHArc: Record 5110;
        InitialQty: Decimal;
        InitialRate: Decimal;
        "-------------02Jul2018": Decimal;
        SPName: Text;
        PayDesc: Text;
        PayTerm: Record 3;
        SP02: Record 13;
        PH02: Record 38;
        SPCode02: Code[20];
        "---------19oct2018": Integer;
        Summary: Boolean;
        PL19: Record 39;
        GRNQty19: Decimal;
        OrdQty19: Decimal;
        OutQty19: Decimal;
        // PGEL: Record 16556;
        GateEntryNo: Code[20];
        RecPRH: Record 120;

    //  //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        ExcBuffer.INIT;
        ExcBuffer.VALIDATE("Row No.", Rowno);
        ExcBuffer.VALIDATE("Column No.", columnno);
        ExcBuffer."Cell Value as Text" := Cellvalue;
        ExcBuffer.Formula := '';
        ExcBuffer.Bold := Bold;
        ExcBuffer.Underline := Underline;
        ExcBuffer.NumberFormat := NoFormat;
        ExcBuffer."Cell Type" := CType;
        ExcBuffer.INSERT;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


report 50178 "Cancel Invoice/SRO Register"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 11Sep2017   RB-N         Filtering of Report on the basis of Location GSTNo.
    //                          GST Related Fields added in the Reports
    //                          .
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CancelInvoiceSRORegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem(SalesCrMemoHeader; "Sales Cr.Memo Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "Posting Date", "Location Code", "No.";
            column(SCMH_No; "No.")
            {
            }
            column(CompName; CompInfo.Name)
            {
            }
            column(DateRange; DateRange)
            {
            }
            column(LocFilter; LocFilter)
            {
            }
            column(SMCH_PostingDate; SalesCrMemoHeader."Posting Date")
            {
            }
            column(CustName; CustName)
            {
            }
            column(Applies_DocNo; SalesCrMemoHeader."Applies-to Doc. No.")
            {
            }
            column(AppDocDate; AppDocDate)
            {
            }
            column(Label; label)
            {
            }
            dataitem(SalesCrMemoLine; "Sales Cr.Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.");
                column(SCML_DocNo; "Document No.")
                {
                }
                column(ItemNo; SalesCrMemoLine."No.")
                {
                }
                column(ItemName; ItemName)
                {
                }
                column(LineAmt; SalesCrMemoLine."Line Amount")
                {
                }
                column(LineDiscAmt; SalesCrMemoLine."Line Discount Amount")
                {
                }
                column(LineQty; SalesCrMemoLine.Quantity)
                {
                }
                column(LineQtyBase; SalesCrMemoLine."Quantity (Base)")
                {
                }
                column(tnetsales; tnetsales)
                {
                }

                trigger OnAfterGetRecord()
                begin


                    //>>11Sep2017 GST
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
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                    // DetailGSTEntry.SETRANGE(Type, Type);
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT

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
                                //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");
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
                    //<<11Sep2017 GST

                    //>>1

                    grossamt := 0;
                    Charges := 0;
                    freight := 0;

                    IF Item.GET("No.") THEN;
                    ItemName := Item.Description;


                    //>>
                    NetSales1 := NetSales1 + "Line Discount Amount" + "Line Amount"; //>>16Feb2017 RB-N, G1
                    Qty1 := Qty1 + Quantity; //>>16Feb2017 RB-N, G2
                    BedAmt1 := BedAmt1 + 0;// SalesCrMemoLine."BED Amount"; //>>16Feb2017 RB-N, G5
                    EcessAmt1 := EcessAmt1 + 0;//SalesCrMemoLine."eCess Amount";//>>16Feb2017 RB-N, G6
                    SheAmt1 := SheAmt1 + 0;// SalesCrMemoLine."SHE Cess Amount";//>>16Feb2017 RB-N, G7
                    TaxableAmt1 := TaxableAmt1 + 0;// SalesCrMemoLine."Tax Base Amount";//>>16Feb2017 RB-N, G8
                    TaxAmt1 := TaxAmt1 + 0;// SalesCrMemoLine."Tax Amount";//>>16Feb2017 RB-N, G9

                    //<<
                    reciTEMUOM.RESET;
                    reciTEMUOM.SETRANGE(reciTEMUOM."Item No.", SalesCrMemoLine."No.");
                    IF reciTEMUOM.FINDFIRST THEN
                        QuantityBase := SalesCrMemoLine.Quantity * reciTEMUOM."Qty. per Unit of Measure";

                    QtyBase1 := QtyBase1 + QuantityBase; //>>16Feb2017 RB-N, G3


                    TotQtyBase := TotQtyBase + QuantityBase;

                    ItemCategoryCode := Item."Item Category Code";

                    RecItemCategory.RESET;
                    RecItemCategory.SETRANGE(RecItemCategory.Code, ItemCategoryCode);
                    IF RecItemCategory.FINDFIRST THEN
                        ItemCategoryName := RecItemCategory.Description;

                    Amt := Amt + SalesCrMemoLine.Amount;
                    QtyBase := QtyBase + QuantityBase;

                    IF SalesCrMemoLine.Quantity = 0 THEN
                        CurrReport.SKIP;

                    //>>16Feb2017 RB-N

                    L += 1;//Increment
                    CrLCount := COUNT; //No.of Records
                                       //<<

                    CLEAR(EntryTax);
                    /* PostedStrOrderRec.RESET;
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                    PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Type", 'Charges');
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Group", 'ENTRYTAX');
                    IF PostedStrOrderRec.FINDSET THEN
                        REPEAT
                            EntryTax := EntryTax + PostedStrOrderRec.Amount;
                        UNTIL PostedStrOrderRec.NEXT = 0;
 */
                    EntryTax1 := EntryTax1 + EntryTax;//>>16Feb2017 RB-N, G20

                    CLEAR(TradeSubsidy);
                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Group", '%1|%2', 'TRANSSUBS', 'TPTSUBSIDY');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT
                             TradeSubsidy += PostedStrOrderRec.Amount;
                         UNTIL PostedStrOrderRec.NEXT = 0; */

                    TradeSubsidy1 := TradeSubsidy1 + TradeSubsidy;//>>16Feb2017 RB-N, G11

                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Type", 'Charges');
                     PostedStrOrderRec.SETFILTER("Tax/Charge Group", '%1|%2|%3|%4|%5|%6|%7|%8', 'ADD-DIS', 'AVP-SP-SAN',
                                                'MONTH-DIS', 'NEW-PRI-SP', 'SPOT-DISC', 'VAT-DIF-DI', 'T2 DISC', 'STATE DISC');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT
                             Charges += PostedStrOrderRec.Amount;
                         UNTIL PostedStrOrderRec.NEXT = 0; */

                    Charges1 := Charges1 + Charges; //>>16Feb2017 RB-N, G12

                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Type", 'Charges');
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Group", '%1', 'FREIGHT');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT
                             freight += PostedStrOrderRec.Amount;
                         UNTIL PostedStrOrderRec.NEXT = 0; */

                    freight1 := freight1 + freight; //>>16Feb2017 RB-N, G19

                    CLEAR(CashDiscAmt);
                    /* PostedStrOrderRec.RESET;
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Group", 'CASHDISC');
                    IF PostedStrOrderRec.FINDSET THEN
                        REPEAT
                            CashDiscAmt += PostedStrOrderRec.Amount;
                        UNTIL PostedStrOrderRec.NEXT = 0; */

                    CashDiscAmt1 := CashDiscAmt1 + CashDiscAmt; //>>16Feb2017 RB-N, G10

                    CLEAR(CESSAMT);
                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::"Other Taxes");
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Code", '%1|%2', 'CESS', 'C1');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT
                             CESSAMT += PostedStrOrderRec.Amount;
                         UNTIL PostedStrOrderRec.NEXT = 0;
  */
                    CESSAMT1 := CESSAMT1 + CESSAMT; //>>16Feb2017 RB-N, G13

                    CLEAR(FOCAMT);
                    /*  PostedStrOrderRec.RESET;
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                     PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                     PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Group", '%1|%2', 'FOC', 'FOC DIS');
                     IF PostedStrOrderRec.FINDSET THEN
                         REPEAT
                             FOCAMT += PostedStrOrderRec.Amount;
                         UNTIL PostedStrOrderRec.NEXT = 0; */

                    FOCAMT1 := FOCAMT1 + FOCAMT; //>>16Feb2017 RB-N, G15

                    CLEAR(ADDTAX);
                    /* PostedStrOrderRec.RESET;
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::"Other Taxes");
                    PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Code", 'ADDL TAX');
                    IF PostedStrOrderRec.FINDSET THEN
                        REPEAT
                            ADDTAX += PostedStrOrderRec.Amount;
                        UNTIL PostedStrOrderRec.NEXT = 0; */

                    ADDTAX1 := ADDTAX1 + ADDTAX; //>>16Feb2017 RB-N, G14

                    // EBT MILAN

                    INRPriperltkg := SalesCrMemoLine."Unit Price" / SalesCrMemoLine."Quantity (Base)" *
                                           SalesCrMemoLine.Quantity;

                    INRPriperltkg1 := INRPriperltkg1 + INRPriperltkg; //>>16Feb2017 RB-N, G4

                    // EBT MILAN to Add Product Disc
                    CLEAR(ProdDiscPerlt);
                    MrpMaster.RESET;
                    MrpMaster.SETRANGE(MrpMaster."Item No.", SalesCrMemoLine."No.");
                    MrpMaster.SETRANGE(MrpMaster."Lot No.", SalesCrMemoLine."Lot No.");
                    IF MrpMaster.FINDFIRST THEN BEGIN
                        ProdDiscPerlt := MrpMaster."National Discount";
                    END;

                    ProdDiscPerlt1 := ProdDiscPerlt1 + ProdDiscPerlt; //>>16Feb2017 RB-N, G18
                    // EBT MILAN to add CCFC Detail
                    CLEAR(CCFC);
                    /* PostedStrOrderRec.RESET;
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Invoice No.", "Document No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Line No.", "Line No.");
                    PostedStrOrderRec.SETRANGE(PostedStrOrderRec."Tax/Charge Type", PostedStrOrderRec."Tax/Charge Type"::Charges);
                    PostedStrOrderRec.SETFILTER(PostedStrOrderRec."Tax/Charge Group", '%1', 'CCFC');
                    IF PostedStrOrderRec.FINDSET THEN
                        REPEAT
                            CCFC += PostedStrOrderRec.Amount;
                        UNTIL PostedStrOrderRec.NEXT = 0; */

                    CCFC1 := CCFC1 + CCFC; //>>16Feb2017 RB-N, G16
                                           /*
                                           grossamt:=(ROUND((SalesCrMemoLine."Line Amount"+SalesCrMemoLine."Line Discount Amount"+freight+EntryTax+
                                           +SalesCrMemoLine."BED Amount"+SalesCrMemoLine."eCess Amount"+SalesCrMemoLine."SHE Cess Amount"+SalesCrMemoLine."Tax Amount"+
                                           +CashDiscAmt+TradeSubsidy+(Charges)+CESSAMT+ADDTAX+FOCAMT+CCFC),0.01,'='));
                                           *///Commented 11Sep2017

                    grossamt := 0;// SalesCrMemoLine."Amount To Customer";

                    grossamt1 := grossamt1 + grossamt;//>>16Feb2017 RB-N, G17
                    IF PrintToExcel AND PrintDetail THEN
                        MakeExcelDataBody;

                    //<<1


                    //>>2
                    IF SalesCrMemoHeader."Currency Code" = '' THEN BEGIN
                        tnetsales += SalesCrMemoLine."Line Amount" + SalesCrMemoLine."Line Discount Amount";
                        TotalFreight += freight;
                        TEntrytax += EntryTax;
                        TBEDAmount += 0;// SalesCrMemoLine."BED Amount";
                        TECessAmount += 0;// SalesCrMemoLine."eCess Amount";
                        TSHEAmount += 0;// SalesCrMemoLine."SHE Cess Amount";
                        TtaxBaseAmt += 0;//SalesCrMemoLine."Tax Base Amount";
                        TTaxAmount += 0;//SalesCrMemoLine."Tax Amount";
                        TInvDisAmount += SalesCrMemoLine."Inv. Discount Amount";
                        TLineDisAmount += SalesCrMemoLine."Line Discount Amount";
                        TTradeSubsidy += TradeSubsidy;
                        TCharges += Charges;
                        TCessAmount += CESSAMT;
                        Taddtax += ADDTAX;
                        TCashDiscAmt += CashDiscAmt;
                        TGrossamount += grossamt;
                        TFocAmount += FOCAMT;
                        TQty += SalesCrMemoLine.Quantity;
                        totProdDiscPerlt += ProdDiscPerlt; // EBT MILAN
                        totINRPriperltkg += INRPriperltkg; // EBT MILAN
                        totCCFC += CCFC; // EBT MILAN
                    END
                    ELSE BEGIN
                        tnetsales += (SalesCrMemoLine."Line Amount" + SalesCrMemoLine."Line Discount Amount") / SalesCrMemoHeader."Currency Factor";
                        TotalFreight += freight / SalesCrMemoHeader."Currency Factor";
                        TEntrytax += EntryTax / SalesCrMemoHeader."Currency Factor";
                        TBEDAmount += 0;// SalesCrMemoLine."BED Amount" / SalesCrMemoHeader."Currency Factor";
                        TECessAmount += 0;//SalesCrMemoLine."eCess Amount" / SalesCrMemoHeader."Currency Factor";
                        TSHEAmount += 0;// SalesCrMemoLine."SHE Cess Amount" / SalesCrMemoHeader."Currency Factor";
                        TtaxBaseAmt += 0;// SalesCrMemoLine."Tax Base Amount" / SalesCrMemoHeader."Currency Factor";
                        TTaxAmount += 0;// SalesCrMemoLine."Tax Amount" / SalesCrMemoHeader."Currency Factor";
                        TInvDisAmount += SalesCrMemoLine."Inv. Discount Amount" / SalesCrMemoHeader."Currency Factor";
                        TLineDisAmount += SalesCrMemoLine."Line Discount Amount" / SalesCrMemoHeader."Currency Factor";
                        TTradeSubsidy += TradeSubsidy / SalesCrMemoHeader."Currency Factor";
                        TCharges += Charges / SalesCrMemoHeader."Currency Factor";
                        TCessAmount += CESSAMT / SalesCrMemoHeader."Currency Factor";
                        Taddtax += ADDTAX / SalesCrMemoHeader."Currency Factor";
                        TCashDiscAmt += CashDiscAmt / SalesCrMemoHeader."Currency Factor";
                        TGrossamount += grossamt / SalesCrMemoHeader."Currency Factor";
                        TFocAmount += FOCAMT / SalesCrMemoHeader."Currency Factor";
                        TQty += SalesCrMemoLine.Quantity;
                        totProdDiscPerlt += ProdDiscPerlt; //EBT MILAN
                        totINRPriperltkg += INRPriperltkg / SalesCrMemoHeader."Currency Factor"; // EBT MILAN
                        totCCFC += CCFC / SalesCrMemoHeader."Currency Factor"; // EBT MILAN
                    END;

                    //<<2

                    //>>16Feb2017 RB-N, Group with DocumentNo.
                    //---
                    //IF CrLCode = "Document No." THEN
                    //BEGIN
                    //IF TempCount = CrLCount THEN
                    //BEGIN
                    IF PrintToExcel AND (TempCount = L) THEN BEGIN
                        MakeExcelGroupFooter;

                        NetSales1 := 0;//g1
                        Qty1 := 0;//g2
                        QtyBase1 := 0;//g3
                        INRPriperltkg1 := 0;//g4
                        BedAmt1 := 0;//g5
                        EcessAmt1 := 0;//g6
                        SheAmt1 := 0;//g7
                        TaxableAmt1 := 0;//g8
                        TaxAmt1 := 0;//g9
                        CashDiscAmt1 := 0;//g10
                        TradeSubsidy1 := 0;//g11
                        Charges1 := 0;//g12
                        CESSAMT1 := 0;//g13
                        ADDTAX1 := 0;//g14
                        FOCAMT1 := 0;//g15
                        CCFC1 := 0;//g16
                        grossamt1 := 0;//g17
                        ProdDiscPerlt1 := 0;//g18
                        freight1 := 0; //g19
                        EntryTax1 := 0;//g20
                        L := 0; //11Sep2017
                    END;

                    //END;
                    //---
                    //<<
                    //MESSAGE('Doc No. %1 \ TempCount %2\ CrlCount %3',CrLCode,TempCount,CrLCount);

                end;

                trigger OnPreDataItem()
                begin
                    //>>1

                    SETRANGE(Type, Type::Item);

                    CurrReport.CREATETOTALS(freight, TradeSubsidy, QuantityBase, Charges, EntryTax, grossamt, CESSAMT, FOCAMT, ADDTAX, CashDiscAmt);
                    CurrReport.CREATETOTALS(ProdDiscPerlt, INRPriperltkg, CCFC);
                    //<<1

                    L := 0; //16Feb2017 RB-N
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //>> 16Feb2017 RB-N

                I += 1;
                CrCount := COUNT;

                //<<

                //>>1

                CustName := "Bill-to Name";
                OrderDate := "Posting Date";

                Cust.GET(SalesCrMemoHeader."Sell-to Customer No.");

                SalesInvHeader.SETRANGE("No.", SalesCrMemoHeader."Applies-to Doc. No.");
                IF SalesInvHeader.FIND('-') THEN;
                AppDocDate := SalesInvHeader."Posting Date";


                LocationName := '';
                LocGSTNo := '';//RB-N 11Sep2017
                recLcoationTable.RESET;
                recLcoationTable.SETRANGE(recLcoationTable.Code, SalesCrMemoHeader."Location Code");
                IF recLcoationTable.FINDFIRST THEN BEGIN
                    LocationName := recLcoationTable.Name;
                    LocGSTNo := recLcoationTable."GST Registration No.";
                END;

                //GSTStateCode RB-N 11Sep2017
                LocStateCode := '';
                State07.RESET;
                IF State07.GET(recLcoationTable."State Code") THEN
                    LocStateCode := State07."State Code (GST Reg. No.)";

                //>>RB-N 11Sep2017 GSTNo. Filter

                IF TempGSTNo <> '' THEN BEGIN
                    IF TempGSTNo <> LocGSTNo THEN
                        CurrReport.SKIP;

                END;
                //<<RB-N 11Sep2017 GSTNo. Filter


                NOofDays := 0;
                IF SalesCrMemoHeader."Applies-to Doc. No." <> '' THEN BEGIN
                    IF AppDocDate <> 0D THEN //RSPL_PRANJALI
                        NOofDays := SalesCrMemoHeader."Posting Date" - AppDocDate;
                END;
                //<<1


                //>>16Feb2017 RB-N, Group with DocumentNo.

                recSCrL.RESET;
                recSCrL.SETRANGE("Document No.", SalesCrMemoHeader."No.");
                recSCrL.SETRANGE(recSCrL.Type, recSCrL.Type::Item);
                recSCrL.SETFILTER(Quantity, '<>%1', 0);//11Sep2017
                IF recSCrL.FINDFIRST THEN BEGIN
                    CrLCode := recSCrL."Document No.";
                    TempCount := recSCrL.COUNT;
                END;
                //<<
            end;

            trigger OnPostDataItem()
            begin

                //>>16Feb2017 RB-N, CreditMemo Total


                IF CrCount = I THEN BEGIN
                    IF PrintToExcel THEN
                        MakeExcelDataFooter;
                END;

                //<<
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                IF (CancelInvoice = FALSE) AND (SalesOrderReturn = FALSE) THEN
                    ERROR('You have not selected any option');

                IF (CancelInvoice = TRUE) AND (SalesOrderReturn = TRUE) THEN BEGIN
                END ELSE BEGIN
                    IF CancelInvoice = TRUE THEN
                        SalesCrMemoHeader.SETFILTER(SalesCrMemoHeader."Cancelled Invoice", 'YES');

                    IF SalesOrderReturn = TRUE THEN
                        SalesCrMemoHeader.SETFILTER(SalesCrMemoHeader."Cancelled Invoice", 'NO');

                END;
                //<<1

                //>>2
                IF (CancelInvoice = TRUE) AND (SalesOrderReturn = FALSE) THEN
                    label := 'Cancel Invoice Register'
                ELSE
                    IF (CancelInvoice = FALSE) AND (SalesOrderReturn = TRUE) THEN
                        label := 'Sales Return Order Register'
                    ELSE
                        IF (CancelInvoice = TRUE) AND (SalesOrderReturn = TRUE) THEN
                            label := 'Cancel Invoice / SRO Register'
                        ELSE
                            IF (CancelInvoice = FALSE) AND (SalesOrderReturn = FALSE) THEN
                                label := '';
                //>>2

                //..17Feb2017 RB-N
                DateRange := GETFILTER("Posting Date");
                LocFilter := GETFILTER("Location Code");
                //..

                //>> 16Feb2017 LocationName
                LocCode := SalesCrMemoHeader.GETFILTER("Location Code");
                recLocation.RESET;
                recLocation.SETRANGE(Code, LocCode);
                IF recLocation.FINDFIRST THEN BEGIN
                    LocName := recLocation.Name;
                    //MESSAGE('LLL %1',LocName);
                END;

                IF PrintToExcel THEN
                    MakeExcelDataHeader;

                I := 0; //16Feb2017 RB-N
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
                    ApplicationArea = all;
                    TableRelation = "GST Registration Nos.";
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = all;
                }
                field("Print Details"; PrintDetail)
                {
                    ApplicationArea = all;
                }
                field("Cancel Invoice"; CancelInvoice)
                {
                    ApplicationArea = all;
                }
                field("Sales Order Return"; SalesOrderReturn)
                {
                    ApplicationArea = all;
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

        IF PrintToExcel THEN
            CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin
        //>>1

        CompInfo.GET;

        IF PrintToExcel THEN
            MakeExcelInfo;


        IF SalesCrMemoHeader."Location Code" <> '' THEN
            recLoc.RESET;
        recLoc.SETRANGE(recLoc.Code, SalesCrMemoHeader.GETFILTER("Location Code"));
        IF recLoc.FINDFIRST THEN;

        //<<1
    end;

    var
        OrderDate: Date;
        CustName: Text[100];
        ItemName: Text[100];
        Item: Record 27;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        QuantityBase: Decimal;
        SalesInvHeader: Record 112;
        AppDocDate: Date;
        CancelInvoice: Boolean;
        SalesOrderReturn: Boolean;
        QtyBase: Decimal;
        Amt: Decimal;
        Charges: Decimal;
        EntryTax: Decimal;
        //PostedStrOrderRec: Record 13798;
        CustRec: Record 18;
        CustCode: Code[50];
        freight: Decimal;
        grossamt: Decimal;
        label: Text[30];
        RecsalesLine: Record 37;
        usersetup: Record 91;
        CESSAMT: Decimal;
        FOCAMT: Decimal;
        TotalFreight: Decimal;
        TEntrytax: Decimal;
        TBEDAmount: Decimal;
        TSHEAmount: Decimal;
        TTaxAmount: Decimal;
        TInvDisAmount: Decimal;
        TLineDisAmount: Decimal;
        TTransSubsidy: Decimal;
        TECessAmount: Decimal;
        TCessAmount: Decimal;
        TCharges: Decimal;
        TGrossamount: Decimal;
        TFocAmount: Decimal;
        TtaxBaseAmt: Decimal;
        ADDTAX: Decimal;
        Taddtax: Decimal;
        tnetsales: Decimal;
        ItemCategoryCode: Code[10];
        RecItemCategory: Record 5722;
        ItemCategoryName: Text[50];
        reciTEMUOM: Record 5404;
        TotQtyBase: Decimal;
        TradeSubsidy: Decimal;
        TTradeSubsidy: Decimal;
        CashDiscAmt: Decimal;
        TCashDiscAmt: Decimal;
        TQty: Decimal;
        PrintDetail: Boolean;
        CompInfo: Record 79;
        recLoc: Record 14;
        Cust: Record 18;
        recLcoationTable: Record 14;
        LocationName: Text[100];
        NOofDays: Integer;
        ProdDiscPerlt: Decimal;
        totProdDiscPerlt: Decimal;
        INRPriperltkg: Decimal;
        totINRPriperltkg: Decimal;
        MrpMaster: Record 50013;
        CCFC: Decimal;
        totCCFC: Decimal;
        Text001: Label 'Cancel Invoice / SRO Register';
        Text000: Label 'Data';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        Text0009: Label 'Filters';
        "--15Feb2017": Integer;
        DateRange: Text[100];
        LocFilter: Text[100];
        I: Integer;
        CrCount: Integer;
        recSCrL: Record 115;
        CrLCode: Code[20];
        L: Integer;
        CrLCount: Integer;
        TempCount: Integer;
        "--16Feb2017": Integer;
        NetSales1: Decimal;
        Qty1: Decimal;
        QtyBase1: Decimal;
        EntryTax1: Decimal;
        TradeSubsidy1: Decimal;
        Charges1: Decimal;
        freight1: Decimal;
        CashDiscAmt1: Decimal;
        CESSAMT1: Decimal;
        FOCAMT1: Decimal;
        ADDTAX1: Decimal;
        INRPriperltkg1: Decimal;
        ProdDiscPerlt1: Decimal;
        CCFC1: Decimal;
        grossamt1: Decimal;
        BedAmt1: Decimal;
        EcessAmt1: Decimal;
        SheAmt1: Decimal;
        TaxableAmt1: Decimal;
        TaxAmt1: Decimal;
        recLocation: Record 14;
        LocCode: Code[20];
        LocName: Text[100];
        "---11Sep2017---GST": Integer;
        State07: Record State;
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
        "------11Sep2017": Integer;
        TempGSTNo: Code[15];
        LocGSTNo: Code[15];
        LocStateCode: Code[10];

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Cancel Invoice / SRO Register', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Cancel Invoice/SRO Register", FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0009), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(SalesCrMemoHeader.GETFILTERS, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.ClearNewRow;
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, TRUE, '@', 1);

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        IF (CancelInvoice = TRUE) AND (SalesOrderReturn = FALSE) THEN
            ExcelBuf.AddColumn('Cancel Invoice Register', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        IF (CancelInvoice = FALSE) AND (SalesOrderReturn = TRUE) THEN
            ExcelBuf.AddColumn('SRO Register', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        IF (CancelInvoice = TRUE) AND (SalesOrderReturn = TRUE) THEN
            ExcelBuf.AddColumn('Cancel Invoice / SRO Register', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        //>>17Feb2016 RB-N
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//37 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//38 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//39 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//40 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//41 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//42 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//43 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//44 RB-N 11Sep2017
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//44

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date Filter' + ' :- ' + SalesCrMemoHeader.GETFILTER("Posting Date"), FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        //>>17Feb2016 RB-N
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//34
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//37 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//38 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//39 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//40 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//41 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//42 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//43 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//44 RB-N 11Sep2017
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//44
        //<<

        ExcelBuf.NewRow;
        //ExcelBuf.AddColumn('Location Filter' + ' :- ' + recLoc.Name,FALSE,'',TRUE,FALSE,TRUE,'@',1);
        ExcelBuf.AddColumn('Location Filter :- ' + LocName, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17Feb2017 RB-N

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//37 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//38 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//39 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//40 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//41 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//42 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//43 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//44 RB-N 11Sep2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Document No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Item Caterogy Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Customer Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Party Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Location StateCode', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10 RB-N 11Sep2017
        ExcelBuf.AddColumn('Location GSTIN', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11 RB-N 11Sep2017
        ExcelBuf.AddColumn('Resp. Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Net Sale', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Quantity in Ltrs', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Applies-to Doc. No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Applies-to Doc. Date.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('INR Price Per Lt/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1); //18
        ExcelBuf.AddColumn('Freight', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Entry Tax', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('BED', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('ECess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('SHE CESS', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('Taxable Amt', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('Tax Amt', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        ExcelBuf.AddColumn('Cash Discount Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('Trade Subsidy', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('CESS Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('Addl Tax Amt', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('FOC', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('CCFC', FALSE, '', TRUE, FALSE, TRUE, '@', 1); //33

        ExcelBuf.AddColumn('GST %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 RB-N 11Sep2017
        ExcelBuf.AddColumn('CGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35 RB-N 11Sep2017
        ExcelBuf.AddColumn('SGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//36 RB-N 11Sep2017
        ExcelBuf.AddColumn('IGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//37 RB-N 11Sep2017
        ExcelBuf.AddColumn('UTGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//38 RB-N 11Sep2017

        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//39
        ExcelBuf.AddColumn('Product Disc. Per Lt', FALSE, '', TRUE, FALSE, TRUE, '@', 1); //40
        ExcelBuf.AddColumn('TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//41
        ExcelBuf.AddColumn('CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//42
        ExcelBuf.AddColumn('GST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//43 RB-N 11Sep2017
        ExcelBuf.AddColumn('No of Days', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//44
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        IF SalesCrMemoHeader."Currency Code" = '' THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn(SalesCrMemoLine."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
            ExcelBuf.AddColumn(OrderDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//2
            ExcelBuf.AddColumn(SalesCrMemoLine."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
            ExcelBuf.AddColumn(ItemName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
            ExcelBuf.AddColumn(ItemCategoryName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
            ExcelBuf.AddColumn(SalesCrMemoLine."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
            ExcelBuf.AddColumn(CustName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
            ExcelBuf.AddColumn(SalesCrMemoHeader."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
            ExcelBuf.AddColumn(LocationName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//9
            ExcelBuf.AddColumn(LocStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 RB-N 11Sep2017
            ExcelBuf.AddColumn(LocGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11 RB-N 11Sep2017
            ExcelBuf.AddColumn(SalesCrMemoLine."Responsibility Center", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
                                                                                                               //ExcelBuf.AddColumn(SalesCrMemoLine."Line Amount"+SalesCrMemoLine."Line Discount Amount",FALSE,'',FALSE,FALSE,FALSE,'',1);
            ExcelBuf.AddColumn((SalesCrMemoLine."Line Amount" + SalesCrMemoLine."Line Discount Amount"), FALSE, '', FALSE, FALSE, FALSE, '#,0.00', 0);//13
                                                                                                                                                      //0.01,'=')
                                                                                                                                                      //.AddColumn("Purchase Line"."Unit Cost",FALSE,'',FALSE,FALSE,FALSE,'#,####0.00',ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(SalesCrMemoLine.Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//14
            ExcelBuf.AddColumn(QuantityBase, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//15
            ExcelBuf.AddColumn(SalesCrMemoHeader."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
            ExcelBuf.AddColumn(AppDocDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//17
            ExcelBuf.AddColumn(ROUND(INRPriperltkg, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//18
            ExcelBuf.AddColumn(ROUND(freight, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//19
            ExcelBuf.AddColumn(ROUND(EntryTax, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//20
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."BED Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//21
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."eCess Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//22
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."SHE Cess Amount"*/0, 0.01, '='), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//23
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."Tax Base Amount"*/0, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//24
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."Tax Amount"*/0, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//25
                                                                                                                                //ExcelBuf.AddColumn(PostedStrOrderRec."Tax/Charge Type",FALSE,'',FALSE,FALSE,FALSE,'@',1);
            ExcelBuf.AddColumn(/*SalesCrMemoLine."Tax %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//26
            ExcelBuf.AddColumn(ROUND(CashDiscAmt, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//27
            ExcelBuf.AddColumn(ROUND(TradeSubsidy, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//28
            ExcelBuf.AddColumn(ROUND(Charges, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//29
            ExcelBuf.AddColumn(ROUND(CESSAMT, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//30
            ExcelBuf.AddColumn(ROUND(ADDTAX, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//31
            ExcelBuf.AddColumn(ROUND(FOCAMT, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//32
            ExcelBuf.AddColumn(ROUND(CCFC, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); //33
            ExcelBuf.AddColumn(/*SalesCrMemoLine."GST %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0); //34 RB-N 11Sep2017
            ExcelBuf.AddColumn(vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //35 RB-N 11Sep2017
            ExcelBuf.AddColumn(vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //36 RB-N 11Sep2017
            ExcelBuf.AddColumn(vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //37 RB-N 11Sep2017
            ExcelBuf.AddColumn(vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //38 RB-N 11Sep2017

            ExcelBuf.AddColumn(ROUND(grossamt, 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); //39
            ExcelBuf.AddColumn(ProdDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0); //40
            ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/0, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//41
            ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/0, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//42
            ExcelBuf.AddColumn(Cust."GST Registration No.", FALSE, '', FALSE, FALSE, FALSE, '', 1); //43 RB-N 11Sep2017
            ExcelBuf.AddColumn(NOofDays, FALSE, '', FALSE, FALSE, FALSE, '', 0);//44
        END
        ELSE BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn(SalesCrMemoLine."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(OrderDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);
            ExcelBuf.AddColumn(SalesCrMemoLine."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(ItemName, FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(ItemCategoryName, FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(SalesCrMemoLine."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(CustName, FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(SalesCrMemoHeader."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(LocationName, FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(LocStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 RB-N 11Sep2017
            ExcelBuf.AddColumn(LocGSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);//11 RB-N 11Sep2017
            ExcelBuf.AddColumn(SalesCrMemoLine."Responsibility Center", FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(ROUND((SalesCrMemoLine."Line Amount" + SalesCrMemoLine."Line Discount Amount")
             / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(SalesCrMemoLine.Quantity, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(QuantityBase, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(SalesCrMemoHeader."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(AppDocDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);
            ExcelBuf.AddColumn(ROUND(INRPriperltkg / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);// EBT MILAN
            ExcelBuf.AddColumn(ROUND(freight / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(EntryTax / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."BED Amount"*/0, 0.01, '=')
             / SalesCrMemoHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."eCess Amount"*/0, 0.01, '=')
             / SalesCrMemoHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."SHE Cess Amount"*/0, 0.01, '=')
             / SalesCrMemoHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."Tax Base Amount"*/0 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(/*SalesCrMemoLine."Tax Amount"*/0 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            //ExcelBuf.AddColumn(PostedStrOrderRec."Tax/Charge Type",FALSE,'',FALSE,FALSE,FALSE,'@',1);
            ExcelBuf.AddColumn(/*SalesCrMemoLine."Tax %"*/0, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);
            ExcelBuf.AddColumn(CashDiscAmt / SalesCrMemoHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(ROUND(TradeSubsidy / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(Charges / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(CESSAMT / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(ADDTAX / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(FOCAMT / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ROUND(CCFC / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0); //

            ExcelBuf.AddColumn(/*SalesCrMemoLine."GST %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0); //34 RB-N 11Sep2017
            ExcelBuf.AddColumn(vCGST / SalesCrMemoHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //35 RB-N 11Sep2017
            ExcelBuf.AddColumn(vSGST / SalesCrMemoHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //36 RB-N 11Sep2017
            ExcelBuf.AddColumn(vIGST / SalesCrMemoHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //37 RB-N 11Sep2017
            ExcelBuf.AddColumn(vUTGST / SalesCrMemoHeader."Currency Factor", FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0); //38 RB-N 11Sep2017
            ExcelBuf.AddColumn(ROUND(grossamt / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
            ExcelBuf.AddColumn(ProdDiscPerlt, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0); // 40
            ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/0, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/0, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(FORMAT(Cust."GST Registration No."), FALSE, '', FALSE, FALSE, FALSE, '@', 1);//43 RB-N  11Sep2017
            ExcelBuf.AddColumn(NOofDays, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        END;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,CompInfo.Name,Text001,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>15Feb2017 RB-N
        /* ExcelBuf.CreateBook('', Text001);
        //ExcelBuf.CreateBookAndOpenExcel('',Text000,Text001,COMPANYNAME,USERID);
        ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', USERID);
        ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//36
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//37 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//38 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//39 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//40 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//41 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//42 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//43 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//44 RB-N 11Sep2017


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn(ROUND(tnetsales, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, TRUE, '0.00', 0);
        ExcelBuf.AddColumn(TotQtyBase, FALSE, '', TRUE, FALSE, TRUE, '0.00', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(ROUND(totINRPriperltkg, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0); // EBT MILAN
        ExcelBuf.AddColumn(ROUND(TotalFreight, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TEntrytax, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TBEDAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TECessAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TSHEAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TtaxBaseAmt, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TTaxAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(ROUND(TCashDiscAmt, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TTradeSubsidy, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TCharges, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TCessAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(Taddtax, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(TFocAmount, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);
        ExcelBuf.AddColumn(ROUND(totCCFC, 0.01), FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0); //
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//34 RB-N 11Sep2017
        ExcelBuf.AddColumn(TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//35 RB-N 11Sep2017
        ExcelBuf.AddColumn(TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//36 RB-N 11Sep2017
        ExcelBuf.AddColumn(TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//37 RB-N 11Sep2017
        ExcelBuf.AddColumn(TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);//38 RB-N 11Sep2017

        ExcelBuf.AddColumn(TGrossamount, FALSE, '', TRUE, FALSE, TRUE, '#,#0.00', 0);
        ExcelBuf.AddColumn(totProdDiscPerlt, FALSE, '', TRUE, FALSE, TRUE, '0.00', 0); //39
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//43 RB-N 11Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelGroupdata()
    begin
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelGroupFooter()
    begin
        IF SalesCrMemoHeader."Currency Code" = '' THEN BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn(SalesCrMemoLine."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(OrderDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(ItemCategoryName, FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(SalesCrMemoLine."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(CustName, FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(SalesCrMemoHeader."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(LocationName, FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(LocStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 RB-N 11Sep2017
            ExcelBuf.AddColumn(LocGSTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//11 RB-N 11Sep2017
            ExcelBuf.AddColumn(SalesCrMemoLine."Responsibility Center", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            //ExcelBuf.AddColumn(SalesCrMemoLine."Line Amount"+SalesCrMemoLine."Line Discount Amount",FALSE,'',TRUE,FALSE,FALSE,'',1);//G1 INR_NetSales
            ExcelBuf.AddColumn(NetSales1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G1 INR_NetSales
                                                                                          //ExcelBuf.AddColumn(SalesCrMemoLine.Quantity,FALSE,'',TRUE,FALSE,FALSE,'',1);//G2 INR_Qty
            ExcelBuf.AddColumn(Qty1, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//G2 INR_Qty
                                                                               //ExcelBuf.AddColumn(QuantityBase,FALSE,'',TRUE,FALSE,FALSE,'',1);//G3 INR_QtyBase
            ExcelBuf.AddColumn(QtyBase1, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//G3 INR_QtyBase
            ExcelBuf.AddColumn(SalesCrMemoHeader."Applies-to Doc. No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(AppDocDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);
            //ExcelBuf.AddColumn(ROUND(INRPriperltkg,0.01),FALSE,'',TRUE,FALSE,FALSE,'',1); // EBT MILAN //G4 INR_INRPricPer
            ExcelBuf.AddColumn(ROUND(INRPriperltkg1, 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G4 INR_INRPricPer
                                                                                                             //ExcelBuf.AddColumn(freight,FALSE,'',TRUE,FALSE,FALSE,'',1);//G19
            ExcelBuf.AddColumn(freight1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G19
                                                                                         //ExcelBuf.AddColumn(EntryTax,FALSE,'',TRUE,FALSE,FALSE,'',1);//G20
            ExcelBuf.AddColumn(EntryTax1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G20
                                                                                          //ExcelBuf.AddColumn(ROUND(SalesCrMemoLine."BED Amount",0.01,'='),FALSE,'',TRUE,FALSE,FALSE,'',1); //G5 INR_BED
            ExcelBuf.AddColumn(ROUND(BedAmt1, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G5 INR_BED
                                                                                                           //ExcelBuf.AddColumn(ROUND(SalesCrMemoLine."eCess Amount",0.01,'='),FALSE,'',TRUE,FALSE,FALSE,'',1);//G6 INR_Ecess
            ExcelBuf.AddColumn(ROUND(EcessAmt1, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G6 INR_Ecess
                                                                                                            //ExcelBuf.AddColumn(ROUND(SalesCrMemoLine."SHE Cess Amount",0.01,'='),FALSE,'',TRUE,FALSE,FALSE,'',1);//G7 INR_She
            ExcelBuf.AddColumn(ROUND(SheAmt1, 0.01, '='), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G7 INR_She
                                                                                                          //ExcelBuf.AddColumn(SalesCrMemoLine."Tax Base Amount",FALSE,'',TRUE,FALSE,FALSE,'',1);//G8 INR_TaxBase
            ExcelBuf.AddColumn(TaxableAmt1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G8 INR_TaxBase
                                                                                            //ExcelBuf.AddColumn(SalesCrMemoLine."Tax Amount",FALSE,'',TRUE,FALSE,FALSE,'',1);//G9 INR_TaxAmount
            ExcelBuf.AddColumn(TaxAmt1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G9 INR_TaxAmount
                                                                                        //ExcelBuf.AddColumn(PostedStrOrderRec."Tax/Charge Type",FALSE,'',TRUE,FALSE,FALSE,'@',1);
            ExcelBuf.AddColumn(/*SalesCrMemoLine."Tax %"*/0, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//26
                                                                                                       //ExcelBuf.AddColumn(CashDiscAmt,FALSE,'',TRUE,FALSE,FALSE,'',1);//G10 INR_CashDiscount
            ExcelBuf.AddColumn(CashDiscAmt1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G10 INR_CashDiscount
                                                                                             //ExcelBuf.AddColumn(TradeSubsidy,FALSE,'',TRUE,FALSE,FALSE,'',1);//G11 INR_TradeSub
            ExcelBuf.AddColumn(TradeSubsidy1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G11 INR_TradeSub
                                                                                              //ExcelBuf.AddColumn(Charges,FALSE,'',TRUE,FALSE,FALSE,'',1);//G12 INR_Charges
            ExcelBuf.AddColumn(Charges1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G12 INR_Charges
                                                                                         //ExcelBuf.AddColumn(CESSAMT,FALSE,'',TRUE,FALSE,FALSE,'',1);//G13 INR_CessAmt
            ExcelBuf.AddColumn(CESSAMT1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G13 INR_CessAmt
                                                                                         //ExcelBuf.AddColumn(ADDTAX,FALSE,'',TRUE,FALSE,FALSE,'',1);//G14 INR_AddTax
            ExcelBuf.AddColumn(ADDTAX1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G14 INR_AddTax
                                                                                        //ExcelBuf.AddColumn(FOCAMT,FALSE,'',TRUE,FALSE,FALSE,'',1);//G15 INR_FOCAMT
            ExcelBuf.AddColumn(FOCAMT1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G15 INR_FOCAMT
                                                                                        //ExcelBuf.AddColumn(CCFC,FALSE,'',TRUE,FALSE,FALSE,'',1); // EBT MILAN //G16 INR_CCFC
            ExcelBuf.AddColumn(CCFC1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); // EBT MILAN //G16 INR_CCFC
                                                                                       //ExcelBuf.AddColumn(grossamt,FALSE,'',TRUE,FALSE,FALSE,'',1); //G17 INR_Gross

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//34 RB-N 11Sep2017
            ExcelBuf.AddColumn(TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//35 RB-N 11Sep2017
            ExcelBuf.AddColumn(TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//36 RB-N 11Sep2017
            ExcelBuf.AddColumn(TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//37 RB-N 11Sep2017
            ExcelBuf.AddColumn(TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//38 RB-N 11Sep2017

            ExcelBuf.AddColumn(grossamt1, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G17 INR_Gross
                                                                                           //ExcelBuf.AddColumn(ProdDiscPerlt,FALSE,'',TRUE,FALSE,FALSE,'',1);  // EBT MILAN //G18 INR_ProdPerlt
            ExcelBuf.AddColumn(ProdDiscPerlt1, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);  // EBT MILAN //G18 INR_ProdPerlt
            ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/0, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/0, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(Cust."GST Registration No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//43 RB-N 11Sep2017
            ExcelBuf.AddColumn(NOofDays, FALSE, '', TRUE, FALSE, FALSE, '', 0);

        END ELSE BEGIN
            ExcelBuf.NewRow;
            ExcelBuf.AddColumn(SalesCrMemoLine."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(OrderDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(ItemCategoryName, FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(SalesCrMemoLine."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(CustName, FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(SalesCrMemoHeader."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(LocationName, FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(LocStateCode, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 RB-N 11Sep2017
            ExcelBuf.AddColumn(LocGSTNo, FALSE, '', TRUE, FALSE, FALSE, '', 1);//10 RB-N 11Sep2017
            ExcelBuf.AddColumn(SalesCrMemoLine."Responsibility Center", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            //ExcelBuf.AddColumn((SalesCrMemoLine."Line Amount"+SalesCrMemoLine."Line Discount Amount")
            /// SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G1 NetSales
            ExcelBuf.AddColumn(ROUND(NetSales1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G1 NetSales
                                                                                                                                              //ExcelBuf.AddColumn(SalesCrMemoLine.Quantity,FALSE,'',TRUE,FALSE,FALSE,'',1);//G2 INR_Qty
            ExcelBuf.AddColumn(Qty1, FALSE, '', TRUE, FALSE, FALSE, '#,####0', 0);//G2 INR_Qty
                                                                                  //ExcelBuf.AddColumn(QuantityBase,FALSE,'',TRUE,FALSE,FALSE,'',1);//G3 INR_QtyBase
            ExcelBuf.AddColumn(QtyBase1, FALSE, '', TRUE, FALSE, FALSE, '#,####0', 0);//G3 INR_QtyBase
            ExcelBuf.AddColumn(SalesCrMemoHeader."Applies-to Doc. No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);
            ExcelBuf.AddColumn(AppDocDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);
            //ExcelBuf.AddColumn(ROUND(INRPriperltkg / SalesCrMemoHeader."Currency Factor",0.01),FALSE,'',TRUE,FALSE,FALSE,'',1); // EBT MILAN //G4 INR_INRPricPer
            ExcelBuf.AddColumn(ROUND(INRPriperltkg1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); // EBT MILAN //G4 INR_INRPricPer
                                                                                                                                                   //ExcelBuf.AddColumn(freight  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //
            ExcelBuf.AddColumn(freight1 / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G19
                                                                                                                                //ExcelBuf.AddColumn(EntryTax  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1);//
            ExcelBuf.AddColumn(EntryTax1 / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G20
                                                                                                                                //ExcelBuf.AddColumn(ROUND(SalesCrMemoLine."BED Amount",0.01,'=')
                                                                                                                                /// SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G5 BED
            ExcelBuf.AddColumn(ROUND(BedAmt1, 0.01, '=') / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G5 BED
                                                                                                                                                 //ExcelBuf.AddColumn(ROUND(SalesCrMemoLine."eCess Amount",0.01,'=')
                                                                                                                                                 // / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1);//G6 Ecess
            ExcelBuf.AddColumn(ROUND(EcessAmt1, 0.01, '=') / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G6 Ecess
                                                                                                                                                  //ExcelBuf.AddColumn(ROUND(SalesCrMemoLine."SHE Cess Amount",0.01,'=')
                                                                                                                                                  // / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1);//G7 INR_She
            ExcelBuf.AddColumn(ROUND(SheAmt1, 0.01, '=') / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G7 INR_She
                                                                                                                                                //ExcelBuf.AddColumn(SalesCrMemoLine."Tax Base Amount"/ SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1);//G8 TaxBase
            ExcelBuf.AddColumn(ROUND(TaxableAmt1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G8 TaxBase
                                                                                                                                               //ExcelBuf.AddColumn(SalesCrMemoLine."Tax Amount"  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1);//G9 TaxAmount
            ExcelBuf.AddColumn(ROUND(TaxAmt1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//G9 TaxAmount
                                                                                                                                           //ExcelBuf.AddColumn(PostedStrOrderRec."Tax/Charge Type",FALSE,'',TRUE,FALSE,FALSE,'@',1);
            ExcelBuf.AddColumn(/*SalesCrMemoLine."Tax %"*/0, FALSE, '', TRUE, FALSE, FALSE, '', 0);
            //ExcelBuf.AddColumn(CashDiscAmt  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G10 CashDiscount
            ExcelBuf.AddColumn(ROUND(CashDiscAmt1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G10 CashDiscount
                                                                                                                                                 //ExcelBuf.AddColumn(TradeSubsidy  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G11 TradeSub
            ExcelBuf.AddColumn(ROUND(TradeSubsidy1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G11 TradeSub
                                                                                                                                                  //ExcelBuf.AddColumn(Charges  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G12 Charges
            ExcelBuf.AddColumn(ROUND(Charges1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G12 Charges
                                                                                                                                             //ExcelBuf.AddColumn(CESSAMT  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G13 CessAmt
            ExcelBuf.AddColumn(ROUND(CESSAMT1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G13 CessAmt
                                                                                                                                             //ExcelBuf.AddColumn(ADDTAX  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G14 AddTax
            ExcelBuf.AddColumn(ROUND(ADDTAX1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G14 AddTax
                                                                                                                                            //ExcelBuf.AddColumn(FOCAMT  / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G15 FOCAMT
            ExcelBuf.AddColumn(ROUND(FOCAMT1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G15 FOCAMT
                                                                                                                                            //ExcelBuf.AddColumn(CCFC / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //EBT MILAN //G16 CCFC
            ExcelBuf.AddColumn(ROUND(CCFC1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //EBT MILAN //G16 CCFC

            ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//34 RB-N 11Sep2017
            ExcelBuf.AddColumn(TCGST / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//35 RB-N 11Sep2017
            ExcelBuf.AddColumn(TSGST / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//36 RB-N 11Sep2017
            ExcelBuf.AddColumn(TIGST / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//37 RB-N 11Sep2017
            ExcelBuf.AddColumn(TUTGST / SalesCrMemoHeader."Currency Factor", FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//38 RB-N 11Sep2017

            //ExcelBuf.AddColumn(grossamt / SalesCrMemoHeader."Currency Factor",FALSE,'',TRUE,FALSE,FALSE,'',1); //G17 Gross
            ExcelBuf.AddColumn(ROUND(grossamt1 / SalesCrMemoHeader."Currency Factor", 0.01), FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0); //G17 Gross
                                                                                                                                              //ExcelBuf.AddColumn(ProdDiscPerlt,FALSE,'',TRUE,FALSE,FALSE,'',1);  // EBT MILAN //G18 INR_ProdPerlt
            ExcelBuf.AddColumn(ProdDiscPerlt1, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);  // EBT MILAN //G18 INR_ProdPerlt
            ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/0, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/0, FALSE, '', TRUE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(Cust."GST Registration No.", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//43 RB-N 11Sep2017
            ExcelBuf.AddColumn(NOofDays, FALSE, '', TRUE, FALSE, FALSE, '', 0);
        END;

        TCGST := 0;//11Sep2017
        TSGST := 0;//11Sep2017
        TIGST := 0;//11Sep2017
        TUTGST := 0;//11Sep2017
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelFooter()
    begin
    end;
}


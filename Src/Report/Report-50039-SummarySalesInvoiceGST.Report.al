report 50039 "Summary SalesInvoice GST"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 20Sep2017   RB-N         New Report for SalesInvoice Summary w.r.t HSN
    // 01Feb2018   RB-N         Issue in Excel Report Value
    // 02Apr2018   RB-N         No. of Packs
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SummarySalesInvoiceGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            column(No_SalesInvoiceHeader; "Sales Invoice Header"."No.")
            {
            }
            column(BilltoCustomerNo_SalesInvoiceHeader; "Sales Invoice Header"."Bill-to Customer No.")
            {
            }
            column(BilltoName_SalesInvoiceHeader; "Sales Invoice Header"."Bill-to Name")
            {
            }
            column(PostingDate_SalesInvoiceHeader; "Sales Invoice Header"."Posting Date")
            {
            }
            column(VehicleNo; "Sales Invoice Header"."Vehicle No.")
            {
            }
            column(LRRRNo; "Sales Invoice Header"."LR/RR No.")
            {
            }
            column(LRRRDate; "Sales Invoice Header"."LR/RR Date")
            {
            }
            column(TransporterName; ShipAgent.Name)
            {
            }
            column(BillAdd1; "Bill-to Address")
            {
            }
            column(BillAdd2; "Bill-to Address 2")
            {
            }
            column(BillAdd3; "Bill-to City" + '-' + "Bill-to Post Code" + ' ' + BillState + ', ' + BillCountry)
            {
            }
            column(BillGSTNo; BillGSTNo)
            {
            }
            column(BillPANNo; BillPANNo)
            {
            }
            column(ShipAdd1; "Ship-to Address")
            {
            }
            column(ShipAdd2; "Ship-to Address 2")
            {
            }
            column(ShipAdd3; "Ship-to City" + '-' + "Ship-to Post Code" + ' ' + ShipState + ', ' + ShipCountry)
            {
            }
            column(ShipGSTNo; ShipGSTNo)
            {
            }
            column(ShipPANNo; ShipPANNo)
            {
            }
            column(Cartons; Cartons)
            {
            }
            column(Buckets; Buckets)
            {
            }
            column(DrumsBarrels; DrumsBarrels)
            {
            }
            column(Jumbo; Jumbo)
            {
            }
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Type = FILTER(Item),
                                          Quantity = FILTER(<> 0));
                column(DocumentNo_SalesInvoiceLine; "Sales Invoice Line"."Document No.")
                {
                }
                column(LineNo_SalesInvoiceLine; "Sales Invoice Line"."Line No.")
                {
                }
                column(Type_SalesInvoiceLine; "Sales Invoice Line".Type)
                {
                }
                column(No_SalesInvoiceLine; "Sales Invoice Line"."No.")
                {
                }
                column(GSTBaseAmount_SalesInvoiceLine; 0)// "Sales Invoice Line"."GST Base Amount")
                {
                }
                column(GST_SalesInvoiceLine; 0)// "Sales Invoice Line"."GST %")
                {
                }
                column(TotalGSTAmount_SalesInvoiceLine; 0)//"Sales Invoice Line"."Total GST Amount")
                {
                }
                column(HSNCode; "Sales Invoice Line"."HSN/SAC Code")
                {
                }
                column(GSTPer; GSTPer)
                {
                }
                column(vCGST; vCGST)
                {
                }
                column(vSGST; vSGST)
                {
                }
                column(vIGST; vIGST)
                {
                }
                column(AmtToCustomer; 0)//"Sales Invoice Line"."Amount To Customer")
                {
                }
                column(Quantity; "Sales Invoice Line".Quantity)
                {
                }
                column(QuantityBase; "Sales Invoice Line"."Quantity (Base)")
                {
                }
                column(TaxBaseAmount_SalesInvoiceLine; 0)// "Sales Invoice Line"."Tax Base Amount")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>20Sep2017 GST
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
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                    //DetailGSTEntry.SETRANGE(Type, "Sales Invoice Line".Type);
                    DetailGSTEntry.SETRANGE("No.", "No.");
                    DetailGSTEntry.SETRANGE("HSN/SAC Code", "HSN/SAC Code");
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
                                vSGST += ABS(DetailGSTEntry."GST Amount");
                                vSGSTRate := DetailGSTEntry."GST %";

                            END;


                            IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                vIGST += ABS(DetailGSTEntry."GST Amount");
                                vIGSTRate := DetailGSTEntry."GST %";
                            END;

                        UNTIL DetailGSTEntry.NEXT = 0;


                    //<<20Sep2017 GST

                    //>>Excel Body
                    HSNInserted := FALSE;
                    IF PrintToExcel THEN
                        IF TempHSN <> "HSN/SAC Code" THEN BEGIN
                            TempHSN := "HSN/SAC Code";

                            ExBuf11.SETRANGE("Cell Value as Text", "HSN/SAC Code");
                            IF NOT ExBuf11.FINDFIRST THEN BEGIN

                                i11 += 1;
                                ExBuf11.INIT;
                                ExBuf11."Row No." := i11;
                                ExBuf11."Column No." := 1;
                                ExBuf11."Cell Value as Text" := "HSN/SAC Code";
                                ExBuf11.INSERT;

                                HSNInserted := TRUE;

                            END;

                        END;

                    IF HSNInserted THEN BEGIN
                        CLEAR(CGST20);
                        CLEAR(SGST20);
                        CLEAR(IGST20);
                        CLEAR(TaxBaseAmt);
                        CLEAR(QtyinLtr);

                        //>>01Feb2018
                        SIL01.RESET;
                        SIL01.SETRANGE("Document No.", "Document No.");
                        SIL01.SETRANGE("HSN/SAC Code", "HSN/SAC Code");
                        SIL01.SETFILTER(Quantity, '<>%1', 0);
                        IF SIL01.FINDFIRST THEN BEGIN
                            SIL01.CALCSUMS("Quantity (Base)"/*"GST Base Amount"*/, Quantity);
                            TaxBaseAmt := 0;// SIL01."GST Base Amount";
                            QtyinLtr := SIL01."Quantity (Base)";
                            QtyPack := SIL01.Quantity; //02Apr2018
                        END;
                        //<<01Feb2018

                        DetailGSTEntry.RESET;
                        DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                        DetailGSTEntry.SETRANGE(Type, DetailGSTEntry.Type::Item);
                        DetailGSTEntry.SETRANGE("HSN/SAC Code", "HSN/SAC Code");
                        IF DetailGSTEntry.FINDSET THEN
                            REPEAT

                                //TaxBaseAmt += ABS(DetailGSTEntry."GST Base Amount");
                                //QtyinLtr   += ABS(DetailGSTEntry.Quantity);

                                IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                    CGST20 += ABS(DetailGSTEntry."GST Amount");

                                END;

                                IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                    SGST20 += ABS(DetailGSTEntry."GST Amount");

                                END;

                                IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                    SGST20 += ABS(DetailGSTEntry."GST Amount");

                                END;


                                IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                    IGST20 += ABS(DetailGSTEntry."GST Amount");

                                END;

                            UNTIL DetailGSTEntry.NEXT = 0;

                        EnterCell(ii, 1, "Document No.", FALSE, FALSE, '', 1);
                        EnterCell(ii, 2, "HSN/SAC Code", FALSE, FALSE, '', 1);
                        EnterCell(ii, 3, FORMAT(QtyPack), FALSE, FALSE, '0.00', 0);//02Apr2018
                        EnterCell(ii, 4, FORMAT(QtyinLtr), FALSE, FALSE, '0.00', 0);
                        EnterCell(ii, 5, FORMAT(TaxBaseAmt), FALSE, FALSE, '#,#0.00', 0);
                        EnterCell(ii, 6, FORMAT(GSTPer), FALSE, FALSE, '0.00', 0);
                        EnterCell(ii, 7, FORMAT(CGST20), FALSE, FALSE, '#,#0.00', 0);
                        EnterCell(ii, 8, FORMAT(SGST20), FALSE, FALSE, '#,#0.00', 0);
                        EnterCell(ii, 9, FORMAT(IGST20), FALSE, FALSE, '#,#0.00', 0);
                        EnterCell(ii, 10, FORMAT(TaxBaseAmt + CGST20 + SGST20 + IGST20), FALSE, FALSE, '#,#0.00', 0);
                        ii += 1;

                        TTaxBaseAmt += TaxBaseAmt;
                        TQtyinLtr += QtyinLtr;
                        TQtyPack += QtyPack;
                        TCGST20 += CGST20;
                        TSGST20 += SGST20;
                        TIGST20 += IGST20;

                    END;

                    //<<Excel Body
                end;

                trigger OnPostDataItem()
                begin
                    IF PrintToExcel THEN BEGIN
                        EnterCell(ii, 1, '', TRUE, TRUE, '', 1);
                        EnterCell(ii, 2, '', TRUE, TRUE, '', 1);
                        EnterCell(ii, 3, FORMAT(TQtyPack), TRUE, TRUE, '0.00', 0);//02Apr2018
                        EnterCell(ii, 4, FORMAT(TQtyinLtr), TRUE, TRUE, '0.00', 0);
                        EnterCell(ii, 5, FORMAT(TTaxBaseAmt), TRUE, TRUE, '#,#0.00', 0);
                        EnterCell(ii, 6, '', TRUE, TRUE, '', 1);
                        EnterCell(ii, 7, FORMAT(TCGST20), TRUE, TRUE, '#,#0.00', 0);
                        EnterCell(ii, 8, FORMAT(TSGST20), TRUE, TRUE, '#,#0.00', 0);
                        EnterCell(ii, 9, FORMAT(TIGST20), TRUE, TRUE, '#,#0.00', 0);
                        EnterCell(ii, 10, FORMAT(TTaxBaseAmt + TCGST20 + TSGST20 + TIGST20), TRUE, TRUE, '#,#0.00', 0);
                        ii += 1;

                        //RSPLSUM09Apr21>>
                        ii += 1;
                        EnterCell(ii, 1, 'Summary', TRUE, FALSE, '', 1);
                        ii += 1;
                        EnterCell(ii, 1, 'Cartons/Nos', FALSE, FALSE, '', 1);
                        EnterCell(ii, 2, FORMAT(Cartons), FALSE, FALSE, '#,#0.00', 0);
                        ii += 1;
                        EnterCell(ii, 1, 'Buckets', FALSE, FALSE, '', 1);
                        EnterCell(ii, 2, FORMAT(Buckets), FALSE, FALSE, '#,#0.00', 0);
                        ii += 1;
                        EnterCell(ii, 1, 'Jumbo', FALSE, FALSE, '', 1);
                        EnterCell(ii, 2, FORMAT(Jumbo), FALSE, FALSE, '#,#0.00', 0);
                        ii += 1;
                        EnterCell(ii, 1, 'Barrels/Drums', FALSE, FALSE, '', 1);
                        EnterCell(ii, 2, FORMAT(DrumsBarrels), FALSE, FALSE, '#,#0.00', 0);
                        ii += 1;
                        //RSPLSUM09Apr21<<
                    END;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>ShipAgent
                ShipAgent.RESET;
                ShipAgent.GET("Sales Invoice Header"."Shipping Agent Code");

                //>>29Aug2018
                //Bill State Description
                RecState.RESET;
                IF RecState.GET(State) THEN
                    BillState := RecState.Description;

                //Bill GSTNo
                Cust.RESET;
                IF Cust.GET("Bill-to Customer No.") THEN
                    BillGSTNo := Cust."GST Registration No.";
                BillPANNo := Cust."P.A.N. No.";
                //Bill Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Bill-to Country/Region Code") THEN
                    BillCountry := RecCountry.Name;

                //Ship Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Ship-to Country/Region Code") THEN
                    ShipCountry := RecCountry.Name;

                //Shipping
                IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code") THEN BEGIN
                    RecState.RESET;
                    IF RecState.GET(ShipToCode.State) THEN BEGIN
                        ShipState := RecState.Description;
                        ShipGSTNo := ShipToCode."GST Registration No.";
                        ShipPANNo := ShipToCode."P.A.N. No.";
                    END;
                END;

                IF "Ship-to Code" = '' THEN BEGIN
                    ShipState := BillState;
                    ShipGSTNo := BillGSTNo;
                    ShipPANNo := BillPANNo;
                END;
                //<<29Aug2018

                //RSPLSUM09Apr21>>
                CLEAR(Cartons);
                CLEAR(Buckets);
                CLEAR(Jumbo);
                CLEAR(DrumsBarrels);
                RecSalesInvLine.RESET;
                RecSalesInvLine.SETCURRENTKEY("Document No.", Type, Quantity);
                RecSalesInvLine.SETRANGE("Document No.", "No.");
                RecSalesInvLine.SETRANGE(Type, RecSalesInvLine.Type::Item);
                RecSalesInvLine.SETFILTER(Quantity, '>%1', 0);
                IF RecSalesInvLine.FINDSET THEN
                    REPEAT
                        IF (RecSalesInvLine."Qty. per Unit of Measure" >= 0) AND (RecSalesInvLine."Qty. per Unit of Measure" <= 5) THEN
                            Cartons += RecSalesInvLine.Quantity
                        ELSE
                            IF (RecSalesInvLine."Qty. per Unit of Measure" >= 6) AND (RecSalesInvLine."Qty. per Unit of Measure" <= 49) THEN
                                Buckets += RecSalesInvLine.Quantity
                            ELSE
                                IF RecSalesInvLine."Qty. per Unit of Measure" = 50 THEN
                                    Jumbo += RecSalesInvLine.Quantity
                                ELSE
                                    IF RecSalesInvLine."Qty. per Unit of Measure" > 50 THEN
                                        DrumsBarrels += RecSalesInvLine.Quantity;
                    UNTIL RecSalesInvLine.NEXT = 0;
                //RSPLSUM09Apr21<<

                //Excel Header
                IF PrintToExcel THEN BEGIN
                    EnterCell(1, 1, 'Location', TRUE, TRUE, '', 1);
                    EnterCell(1, 2, "Sales Invoice Header"."Bill-to Name", TRUE, TRUE, '', 1);

                    EnterCell(2, 1, 'Billing Address', TRUE, TRUE, '', 1);
                    EnterCell(2, 2, "Bill-to Address", TRUE, TRUE, '', 1);
                    EnterCell(2, 3, "Bill-to Address 2", TRUE, TRUE, '', 1);
                    EnterCell(2, 4, "Bill-to City" + '-' + "Bill-to Post Code" + ' ' + BillState + ', ' + BillCountry, TRUE, TRUE, '', 1);
                    EnterCell(2, 5, 'GSTIN : ' + BillGSTNo, TRUE, TRUE, '', 1);
                    EnterCell(2, 6, 'PAN No. : ' + BillPANNo, TRUE, TRUE, '', 1);

                    EnterCell(3, 1, 'Shipping Address', TRUE, TRUE, '', 1);
                    EnterCell(3, 2, "Ship-to Address", TRUE, TRUE, '', 1);
                    EnterCell(3, 3, "Ship-to Address 2", TRUE, TRUE, '', 1);
                    EnterCell(3, 4, "Ship-to City" + '-' + "Ship-to Post Code" + ' ' + ShipState + ', ' + ShipCountry, TRUE, TRUE, '', 1);
                    EnterCell(3, 5, 'GSTIN : ' + ShipGSTNo, TRUE, TRUE, '', 1);
                    EnterCell(3, 6, 'GSTIN : ' + ShipPANNo, TRUE, TRUE, '', 1);

                    EnterCell(4, 1, 'Transporter', TRUE, TRUE, '', 1);
                    EnterCell(4, 2, ShipAgent.Name, TRUE, TRUE, '', 1);

                    EnterCell(5, 1, 'L.R No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 2, "LR/RR No.", TRUE, TRUE, '', 1);

                    EnterCell(6, 1, 'L.R Date', TRUE, TRUE, '', 1);
                    EnterCell(6, 2, FORMAT("LR/RR Date"), TRUE, TRUE, '', 2);

                    EnterCell(7, 1, 'Vehicle No.', TRUE, TRUE, '', 1);
                    EnterCell(7, 2, "Vehicle No.", TRUE, TRUE, '', 1);

                    EnterCell(8, 1, 'Invoice No.', TRUE, TRUE, '', 1);
                    EnterCell(8, 2, 'HSN/SAC', TRUE, TRUE, '', 1);
                    EnterCell(8, 3, 'No. of Packs', TRUE, TRUE, '', 1);//02Apr2018
                    EnterCell(8, 4, 'Qty in Ltr/Kg', TRUE, TRUE, '', 1);
                    EnterCell(8, 5, 'Taxable Value', TRUE, TRUE, '', 1);
                    EnterCell(8, 6, 'GST %', TRUE, TRUE, '', 1);
                    EnterCell(8, 7, 'CGST', TRUE, TRUE, '', 1);
                    EnterCell(8, 8, 'SGST/UTGST', TRUE, TRUE, '', 1);
                    EnterCell(8, 9, 'IGST', TRUE, TRUE, '', 1);
                    EnterCell(8, 10, 'Gross Total', TRUE, TRUE, '', 1);

                    ii := 9;
                END;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print To Excel"; PrintToExcel)
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
        IF PrintToExcel THEN BEGIN
            /*  ExBuf.CreateBook('', 'Summary of GST Invoice');
             ExBuf.CreateBookAndOpenExcel('', 'Summary of GST Invoice', '', '', USERID);
             ExBuf.GiveUserControl; */

            ExcBuffer.CreateNewBook('Summary of GST Invoice');
            ExcBuffer.WriteSheet('Summary of GST Invoice', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Summary of GST Invoice', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();



        END;
    end;

    trigger OnPreReport()
    begin
        ExBuf11.DELETEALL;
    end;

    var
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
        ShipAgent: Record 291;
        "----------PrinttoExcel": Integer;
        PrintToExcel: Boolean;
        ExBuf: Record 370 temporary;
        ii: Integer;
        TSL: Record 5745;
        TempHSN: Code[10];
        HSNInserted: Boolean;
        ExBuf11: Record 370 temporary;
        i11: Integer;
        CGST20: Decimal;
        TCGST20: Decimal;
        SGST20: Decimal;
        TSGST20: Decimal;
        IGST20: Decimal;
        TIGST20: Decimal;
        TaxBaseAmt: Decimal;
        TTaxBaseAmt: Decimal;
        QtyinLtr: Decimal;
        TQtyinLtr: Decimal;
        "-------01Feb2018": Integer;
        SIL01: Record 113;
        QtyPack: Decimal;
        TQtyPack: Decimal;
        "-------29Aug2018": Integer;
        BillGSTNo: Code[15];
        BillState: Text[50];
        BillCountry: Text[50];
        Loc: Record 14;
        Loc1: Record 14;
        ShipGSTNo: Code[15];
        ShipCountry: Text[50];
        ShipState: Text[50];
        RecState: Record State;
        Cust: Record 18;
        RecCountry: Record 9;
        ShipToCode: Record 222;
        RecSalesInvLine: Record 113;
        Cartons: Decimal;
        Buckets: Decimal;
        DrumsBarrels: Decimal;
        Jumbo: Decimal;
        BillPANNo: Code[10];
        ShipPANNo: Code[10];
        ExcBuffer: Record 370 temporary;

    //  //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        IF PrintToExcel THEN BEGIN
            ExBuf.INIT;
            ExBuf.VALIDATE("Row No.", Rowno);
            ExBuf.VALIDATE("Column No.", columnno);
            ExBuf."Cell Value as Text" := Cellvalue;
            ExBuf.Formula := '';
            ExBuf.Bold := Bold;
            ExBuf.Underline := Underline;
            ExBuf.NumberFormat := NoFormat;
            ExBuf."Cell Type" := CType;
            ExBuf.INSERT;
        END;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


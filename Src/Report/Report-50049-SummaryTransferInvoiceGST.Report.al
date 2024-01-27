report 50049 "Summary TransferInvoice GST"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 20Sep2017   RB-N         New Report for TransferShipment Summary w.r.t HSN
    // 02Apr2018   RB-N         No. of Packs
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SummaryTransferInvoiceGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
        {
            column(No; "No.")
            {
            }
            column(PostingDate; "Posting Date")
            {
            }
            column(VehicleNo; "Vehicle No.")
            {
            }
            column(LRRRNo; "LR/RR No.")
            {
            }
            column(LRRRDate; "LR/RR Date")
            {
            }
            column(TransporterName; "Transfer Shipment Header"."Transporter Name")
            {
            }
            column(TransfertoName; "Transfer Shipment Header"."Transfer-to Name")
            {
            }
            column(BillGSTNo; BillGSTNo)
            {
            }
            column(ShipGSTNo; ShipGSTNo)
            {
            }
            column(ShipAdd1; RecLocation.Address)
            {
            }
            column(ShipAdd2; RecLocation."Address 2")
            {
            }
            column(ShipAdd3; RecLocation.City + '-' + RecLocation."Post Code" + ' ' + RecState.Description + ', ' + RecCountry.Name)
            {
            }
            dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));
                column(DocNo; "Document No.")
                {
                }
                column(LineNo; "Line No.")
                {
                }
                column(ItmNo; "Item No.")
                {
                }
                column(GSTBaseAmount; 0)// "GST Base Amount")
                {
                }
                column(TotalGSTAmount; 0)//"Total GST Amount")
                {
                }
                column(HSNCode; "HSN/SAC Code")
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
                column(Quantity; Quantity)
                {
                }
                column(QuantityBase; "Quantity (Base)")
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
                    DetailGSTEntry.SETRANGE(Type, DetailGSTEntry.Type::Item);
                    DetailGSTEntry.SETRANGE("No.", "Item No.");
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
                        CLEAR(QtyPack);

                        //>>02Apr2018
                        TSL01.RESET;
                        TSL01.SETRANGE("Document No.", "Document No.");
                        TSL01.SETRANGE("HSN/SAC Code", "HSN/SAC Code");
                        TSL01.SETFILTER(Quantity, '<>%1', 0);
                        IF TSL01.FINDFIRST THEN BEGIN
                            TSL01.CALCSUMS(/*"GST Base Amount"*/ Quantity, "Quantity (Base)");
                            TaxBaseAmt := 0;//TSL01."GST Base Amount";
                            QtyinLtr := TSL01.Quantity;
                            QtyPack := TSL01."Quantity (Base)";

                        END;
                        //<<02Apr2018

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
                        EnterCell(ii, 3, FORMAT(QtyPack), FALSE, FALSE, '0.00', 0);//12Apr2018
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
                        TQtyPack += QtyPack;//02Apr2018
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
                        EnterCell(ii, 3, FORMAT(TQtyPack), TRUE, TRUE, '0.00', 0);//02APr2018
                        EnterCell(ii, 4, FORMAT(TQtyinLtr), TRUE, TRUE, '0.00', 0);
                        EnterCell(ii, 5, FORMAT(TTaxBaseAmt), TRUE, TRUE, '#,#0.00', 0);
                        EnterCell(ii, 6, '', TRUE, TRUE, '', 1);
                        EnterCell(ii, 7, FORMAT(TCGST20), TRUE, TRUE, '#,#0.00', 0);
                        EnterCell(ii, 8, FORMAT(TSGST20), TRUE, TRUE, '#,#0.00', 0);
                        EnterCell(ii, 9, FORMAT(TIGST20), TRUE, TRUE, '#,#0.00', 0);
                        EnterCell(ii, 10, FORMAT(TTaxBaseAmt + TCGST20 + TSGST20 + TIGST20), TRUE, TRUE, '#,#0.00', 0);
                        ii += 1;


                    END;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>30Aug2018
                BillGSTNo := '';
                Loc.RESET;
                IF Loc.GET("Transfer Shipment Header"."Transfer-from Code") THEN
                    BillGSTNo := Loc."GST Registration No.";

                ShipGSTNo := '';
                Loc.RESET;
                IF Loc.GET("Transfer Shipment Header"."Transfer-to Code") THEN
                    ShipGSTNo := Loc."GST Registration No.";
                //<<30Aug2018

                RecLocation.RESET;//RSPLSUM
                IF RecLocation.GET("Transfer Shipment Header"."Transfer-to Code") THEN;//RSPLSUM

                RecState.RESET;
                IF RecState.GET(RecLocation."State Code") THEN;

                RecCountry.RESET;
                IF RecCountry.GET(RecLocation."Country/Region Code") THEN;

                IF PrintToExcel THEN BEGIN
                    EnterCell(1, 1, 'Location', TRUE, TRUE, '', 1);
                    EnterCell(1, 2, "Transfer Shipment Header"."Transfer-to Name", TRUE, TRUE, '', 1);

                    EnterCell(2, 1, 'Billing GSTIN', TRUE, TRUE, '', 1);
                    EnterCell(2, 2, BillGSTNo, TRUE, TRUE, '', 1);

                    EnterCell(3, 1, 'Shipping GSTIN', TRUE, TRUE, '', 1);
                    EnterCell(3, 2, ShipGSTNo, TRUE, TRUE, '', 1);

                    EnterCell(4, 1, 'Transporter', TRUE, TRUE, '', 1);
                    EnterCell(4, 2, "Transfer Shipment Header"."Transporter Name", TRUE, TRUE, '', 1);

                    EnterCell(5, 1, 'L.R No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 2, "Transfer Shipment Header"."LR/RR No.", TRUE, TRUE, '', 1);

                    EnterCell(6, 1, 'L.R Date', TRUE, TRUE, '', 1);
                    EnterCell(6, 2, FORMAT("Transfer Shipment Header"."LR/RR Date"), TRUE, TRUE, '', 2);


                    EnterCell(7, 1, 'Vehicle No.', TRUE, TRUE, '', 1);
                    EnterCell(7, 2, "Transfer Shipment Header"."Vehicle No.", TRUE, TRUE, '', 1);

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
        ExcBuffer: Record 370 temporary;
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
        "------02Apr2018": Integer;
        QtyPack: Decimal;
        TQtyPack: Decimal;
        TSL01: Record 5745;
        BillGSTNo: Code[15];
        ShipGSTNo: Code[15];
        Loc: Record 14;
        RecLocation: Record 14;
        RecState: Record State;
        RecCountry: Record 9;

    // //[Scope('Internal')]
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


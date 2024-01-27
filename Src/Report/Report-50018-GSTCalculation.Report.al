report 50018 "GST Calculation"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 14Jul2017   RB-N         GST Related Fields added in the Report.
    // 06Feb2018   RB-N         Actual GST % as per GSTLedger
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/GSTCalculation.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'GST Calculation';
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Transfer Header"; "Transfer Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(No_TransferHeader; "Transfer Header"."No.")
            {
            }
            column(TransferfromCode_TransferHeader; "Transfer Header"."Transfer-from Code")
            {
            }
            column(TransferfromName_TransferHeader; "Transfer Header"."Transfer-from Name")
            {
            }
            column(TransfertoCode_TransferHeader; "Transfer Header"."Transfer-to Code")
            {
            }
            column(TransfertoName_TransferHeader; "Transfer Header"."Transfer-to Name")
            {
            }
            dataitem("Transfer Line"; "Transfer Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE("In-Transit Code" = FILTER('IN-TRANS'));
                column(TransferPriceBase; "Transfer Line"."Transfer Price of Base Unit")
                {
                }
                column(Qty; Quantity)
                {
                }
                column(QuantityBase; "Transfer Line"."Quantity (Base)")
                {
                }
                column(LineNo; "Line No.")
                {
                }
                column(ItemNo_TransferLine; "Transfer Line"."Item No.")
                {
                }
                column(Description_TransferLine; "Transfer Line".Description)
                {
                }
                column(MRPrate; MRPrate)
                {
                }
                column(AssessableValu; 0)// "Assessable Value" / "Qty. per Unit of Measure")
                {
                }
                column(Transfer_Amount; "Transfer Line"."Quantity (Base)" * "Transfer Line"."Transfer Price of Base Unit")
                {
                }
                column(BEDAmount_TransferLine; 0)// "Transfer Line"."BED Amount")
                {
                }
                column(eCessAmount_TransferLine; 0)//"Transfer Line"."eCess Amount")
                {
                }
                column(SHECessAmount_TransferLine; 0)// "Transfer Line"."SHE Cess Amount")
                {
                }
                column(ExciseAmt_Transfer; 0)// "Transfer Line"."BED Amount" + "Transfer Line"."eCess Amount" + "Transfer Line"."SHE Cess Amount")
                {
                }
                column(ADEAmount_TransferLine; 0)//"Transfer Line"."ADE Amount")
                {
                }
                column(GrossAmt_TransferLine; ("Transfer Line"."Quantity (Base)" * "Transfer Line"."Transfer Price of Base Unit") + (0 + 0 + 0) + 0 + FrieghtValue)
                {
                }
                column(GSTBaseAmount; 0)// "Transfer Line"."GST Base Amount")
                {
                }
                column(GSTPer; GSTPer06)
                {
                }
                column(TotalGSTAmount; 0)// "Transfer Line"."Total GST Amount")
                {
                }
                column(GSTGroupCode_TransferLine; "Transfer Line"."GST Group Code")
                {
                }
                column(TransferPrice; "Transfer Line"."Transfer Price")
                {
                }
                column(HSNSACCode_TransferLine; "Transfer Line"."HSN/SAC Code")
                {
                }
                column(CGST07; CGST07)
                {
                }
                column(CGST07Per; CGST07Per)
                {
                }
                column(SGST07; SGST07)
                {
                }
                column(SGST07Per; SGST07Per)
                {
                }
                column(IGST07; IGST07)
                {
                }
                column(IGST07Per; IGST07Per)
                {
                }
                column(CustomDuty27; CustomDuty27)
                {
                }
                column(CustomCess27; CustomCess27)
                {
                }
                column(CompensationCess; CompensationCess)
                {
                }
                dataitem("Item Ledger Entry"; "Item Ledger Entry")
                {
                    DataItemLink = "Order No." = FIELD("Document No."),
                                   "Item No." = FIELD("Item No."),
                                   "Location Code" = FIELD("Transfer-from Code"),
                                   "Document Line No." = FIELD("Line No.");
                    DataItemTableView = SORTING("Document No.", "Posting Date", "Item No.");
                    column(ItemNo_ItemLedgerEntry; "Item Ledger Entry"."Item No.")
                    {
                    }
                    column(MRPrateILE; MRPrate)
                    {
                    }
                    column(AssessableValue_ILE; 0)// "Assessable Value" / "Qty. per Unit of Measure")
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        IF ("Transfer Line"."Inventory Posting Group" = 'AUTOOILS') AND
                          ("Transfer Line".Status = "Transfer Line".Status::Released) AND
                          ("Transfer Line"."Lot No." = '') THEN BEGIN
                            recMRPmaster.RESET;
                            recMRPmaster.SETRANGE(recMRPmaster."Item No.", "Item Ledger Entry"."Item No.");
                            recMRPmaster.SETRANGE(recMRPmaster."Lot No.", "Item Ledger Entry"."Lot No.");
                            recMRPmaster.SETFILTER(recMRPmaster."Qty. per Unit of Measure", '<=%1', 25);
                            IF recMRPmaster.FINDFIRST THEN BEGIN
                                MRPrate := recMRPmaster."MRP Price";
                            END;
                        END;
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    MRPrate := 0;
                    LotNo := '';

                    IF ("Transfer Line"."Inventory Posting Group" = 'AUTOOILS') AND
                       ("Transfer Line".Status = "Transfer Line".Status::Open) AND
                       ("Transfer Line"."Lot No." = '') THEN BEGIN
                        recSalesPrice.RESET;
                        recSalesPrice.SETRANGE(recSalesPrice."Item No.", "Transfer Line"."Item No.");
                        IF recSalesPrice.FINDLAST THEN BEGIN
                            MRPrate := 0;// recSalesPrice."MRP Price";
                        END;
                    END;

                    IF ("Transfer Line"."Inventory Posting Group" = 'AUTOOILS') AND
                      ("Transfer Line".Status = "Transfer Line".Status::Open) AND
                      ("Transfer Line"."Lot No." <> '') THEN BEGIN
                        recMRPmaster.RESET;
                        recMRPmaster.SETRANGE(recMRPmaster."Item No.", "Transfer Line"."Item No.");
                        recMRPmaster.SETRANGE(recMRPmaster."Lot No.", "Transfer Line"."Lot No.");
                        recMRPmaster.SETFILTER(recMRPmaster."Qty. per Unit of Measure", '<=%1', 25);
                        IF recMRPmaster.FINDFIRST THEN BEGIN
                            MRPrate := recMRPmaster."MRP Price";
                        END;
                    END;


                    IF ("Transfer Line"."Inventory Posting Group" = 'AUTOOILS') AND
                      ("Transfer Line".Status = "Transfer Line".Status::Released) AND
                      ("Transfer Line"."Lot No." <> '') THEN BEGIN
                        recMRPmaster.RESET;
                        recMRPmaster.SETRANGE(recMRPmaster."Item No.", "Transfer Line"."Item No.");
                        recMRPmaster.SETRANGE(recMRPmaster."Lot No.", "Transfer Line"."Lot No.");
                        recMRPmaster.SETFILTER(recMRPmaster."Qty. per Unit of Measure", '<=%1', 25);
                        IF recMRPmaster.FINDFIRST THEN BEGIN
                            MRPrate := recMRPmaster."MRP Price";
                        END;
                    END;





                    LocationRec.GET("Transfer Line"."Transfer-from Code");
                    IF LocationRec."Trading Location" THEN BEGIN
                        BEDAmount := 0;//"Transfer Line"."BED Amount" * "Transfer Line"."Qty. per Unit of Measure";
                        eCessAmount := 0;//"Transfer Line"."eCess Amount" * "Transfer Line"."Qty. per Unit of Measure";
                        SheCessAmount := 0;// "Transfer Line"."SHE Cess Amount" * "Transfer Line"."Qty. per Unit of Measure";
                    END
                    ELSE BEGIN
                        BEDAmount := 0;//"Transfer Line"."BED Amount";
                        eCessAmount := 0;//"Transfer Line"."eCess Amount";
                        SheCessAmount := 0;//"Transfer Line"."SHE Cess Amount";
                    END;

                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail.Type, FORMAT(RecStrOrederLineDetail.Type::Transfer));
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Document No.", "Transfer Line"."Document No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Item No.", "Transfer Line"."Item No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Line No.", FORMAT("Transfer Line"."Line No."));
                    // RecStrOrederLineDetail.SETRANGE(RecStrOrederLineDetail."Tax/Charge Group", 'ADDL.DUTY');
                    // IF RecStrOrederLineDetail.FIND('-') THEN BEGIN
                    //     AddDuty := RecStrOrederLineDetail."Calculation Value";
                    // END;


                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail.Type, FORMAT(RecStrOrederLineDetail.Type::Transfer));
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Document No.", "Transfer Line"."Document No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Item No.", "Transfer Line"."Item No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Line No.", FORMAT("Transfer Line"."Line No."));
                    // RecStrOrederLineDetail.SETRANGE(RecStrOrederLineDetail."Tax/Charge Group", 'FREIGHT');
                    // IF RecStrOrederLineDetail.FIND('-') THEN BEGIN
                    //     FrieghtValue := RecStrOrederLineDetail."Calculation Value";
                    // END;
                    AddlDutyTotal := AddlDutyTotal + AddDuty;
                    FrieghtTotal := FrieghtTotal + FrieghtValue;

                    //>>27July2017 CustomDuty & CustomCess
                    // CLEAR(CustomDuty27);
                    // RecStrOrederLineDetail.RESET;
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail.Type, FORMAT(RecStrOrederLineDetail.Type::Transfer));
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Document No.", "Transfer Line"."Document No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Item No.", "Transfer Line"."Item No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Line No.", FORMAT("Transfer Line"."Line No."));
                    // RecStrOrederLineDetail.SETRANGE(RecStrOrederLineDetail."Tax/Charge Group", 'CUSTOM');
                    // IF RecStrOrederLineDetail.FINDFIRST THEN BEGIN

                    //     CustomDuty27 := RecStrOrederLineDetail.Amount;
                    // END;

                    // CLEAR(CustomCess27);
                    // RecStrOrederLineDetail.RESET;
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail.Type, FORMAT(RecStrOrederLineDetail.Type::Transfer));
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Document No.", "Transfer Line"."Document No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Item No.", "Transfer Line"."Item No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Line No.", FORMAT("Transfer Line"."Line No."));
                    // RecStrOrederLineDetail.SETRANGE(RecStrOrederLineDetail."Tax/Charge Group", 'CUST-CESS');
                    // IF RecStrOrederLineDetail.FINDFIRST THEN BEGIN

                    //     CustomCess27 := RecStrOrederLineDetail.Amount;
                    // END;
                    //<<27July2017 CustomDuty & CustomCess

                    //RSPLSUM 01Jul2020>>--Compensation Cess--
                    // CLEAR(CompensationCess);
                    // RecStrOrederLineDetail.RESET;
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail.Type, FORMAT(RecStrOrederLineDetail.Type::Transfer));
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Document No.", "Transfer Line"."Document No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Item No.", "Transfer Line"."Item No.");
                    // RecStrOrederLineDetail.SETFILTER(RecStrOrederLineDetail."Line No.", FORMAT("Transfer Line"."Line No."));
                    // RecStrOrederLineDetail.SETRANGE(RecStrOrederLineDetail."Tax/Charge Group", 'CESS');
                    // IF RecStrOrederLineDetail.FINDFIRST THEN BEGIN
                    //     CompensationCess := RecStrOrederLineDetail.Amount;
                    // END;
                    //RSPLSUM 01Jun2020<<

                    TotalAmount := TotalAmount + ("Transfer Line"."Quantity (Base)" * "Transfer Line"."Transfer Price of Base Unit");


                    //>>14July2017 GST Amount

                    CLEAR(CGST07);
                    CLEAR(CGST07Per);
                    CLEAR(SGST07);
                    CLEAR(SGST07Per);
                    CLEAR(IGST07);
                    CLEAR(IGST07Per);
                    CLEAR(GSTPer06);

                    DetailGST.RESET;
                    DetailGST.SETRANGE("Transaction Type", DetailGST."Transaction Type"::Transfer);
                    DetailGST.SETRANGE("Document No.", "Document No.");
                    DetailGST.SETRANGE(Type, DetailGST.Type::Item);
                    DetailGST.SETRANGE("No.", "Item No.");
                    DetailGST.SETRANGE("Line No.", "Line No.");
                    IF DetailGST.FINDSET THEN
                        REPEAT
                            GSTPer06 += DetailGST."GST %";

                            IF DetailGST."GST Component Code" = 'CGST' THEN BEGIN

                                CGST07 := ABS(DetailGST."GST Amount");
                                CGST07Per := DetailGST."GST %";

                            END;

                            IF DetailGST."GST Component Code" = 'SGST' THEN BEGIN

                                SGST07 := ABS(DetailGST."GST Amount");
                                SGST07Per := DetailGST."GST %";

                            END;
                            //>>27JUly2017 UTGST
                            IF DetailGST."GST Component Code" = 'UTGST' THEN BEGIN

                                SGST07 := ABS(DetailGST."GST Amount");
                                SGST07Per := DetailGST."GST %";

                            END;
                            //<<27July2017 UTGST

                            IF DetailGST."GST Component Code" = 'IGST' THEN BEGIN

                                IGST07 := ABS(DetailGST."GST Amount");
                                IGST07Per := DetailGST."GST %";

                            END;

                        UNTIL DetailGST.NEXT = 0;
                    //<<14July2017 GST Amount
                end;

                trigger OnPreDataItem()
                begin

                    AddlDutyTotal := 0;
                    FrieghtTotal := 0;
                end;
            }
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

    var
        amount: Decimal;
        "Bed amount": Decimal;
        "ecess amount": Decimal;
        "shecess amount": Decimal;
        SerialNo: Integer;
        TotalExciseAmt: Decimal;
        //RecStrOrederLineDetail: Record 13795;
        AddDuty: Decimal;
        FrieghtValue: Decimal;
        AddlDutyTotal: Decimal;
        FrieghtTotal: Decimal;
        TotalAmount: Decimal;
        BEDAmount: Decimal;
        eCessAmount: Decimal;
        SheCessAmount: Decimal;
        LocationRec: Record 14;
        MRPrate: Decimal;
        recMRPmaster: Record 50013;
        recSalesPrice: Record 7002;
        recILE: Record 32;
        recTSL: Record 5745;
        LotNo: Code[10];
        "----------14July2017": Integer;
        DetailGST: Record "Detailed GST Entry Buffer";
        CGST07: Decimal;
        CGST07Per: Decimal;
        SGST07: Decimal;
        SGST07Per: Decimal;
        IGST07: Decimal;
        IGST07Per: Decimal;
        "---27July2017": Integer;
        CustomDuty27: Decimal;
        CustomCess27: Decimal;
        GSTPer06: Decimal;
        CompensationCess: Decimal;
}


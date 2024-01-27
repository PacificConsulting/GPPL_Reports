report 50246 "Prod. Order - Mat.RequisitionT"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 14Nov2017   RB-N         BatchNo.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ProdOrderMatRequisitionT.rdl';
    Caption = 'Prod. Order - Mat. Requisition';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Production Order"; "Production Order")
        {
            DataItemTableView = SORTING(Status, "No.");
            PrintOnlyIfDetail = true;
            RequestFilterFields = Status, "No.", "Source Type", "Source No.";
            column(RequestedBy; USERID)
            {
            }
            column(TodayFormatted; FORMAT(TODAY, 0, 4))
            {
            }
            column(BatchNo; "Production Order"."Description 2")
            {
            }
            column(CompanyName; COMPANYNAME)
            {
            }
            column(ProdOrderTableCaptionFilter; TABLECAPTION + ':' + ProdOrderFilter)
            {
            }
            column(No_ProdOrder; "No.")
            {
            }
            column(Desc_ProdOrder; Description)
            {
            }
            column(SourceNo_ProdOrder; "Source No.")
            {
                IncludeCaption = true;
            }
            column(Status_ProdOrder; Status)
            {
            }
            column(Qty_ProdOrder; Quantity)
            {
                IncludeCaption = true;
            }
            column(Filter_ProdOrder; ProdOrderFilter)
            {
            }
            column(ProdOrderMaterialRqstnCapt; ProdOrderMaterialRqstnCaptLbl)
            {
            }
            column(CurrReportPageNoCapt; CurrReportPageNoCaptLbl)
            {
            }
            dataitem("Prod. Order Component"; "Prod. Order Component")
            {
                DataItemLink = Status = FIELD(Status),
                               "Prod. Order No." = FIELD("No.");
                DataItemTableView = SORTING(Status, "Prod. Order No.", "Prod. Order Line No.", "Line No.");
                column(ItemNo_ProdOrderComp; "Item No.")
                {
                    IncludeCaption = true;
                }
                column(Desc_ProdOrderComp; Description)
                {
                    IncludeCaption = true;
                }
                column(Qtyper_ProdOrderComp; "Quantity per")
                {
                    IncludeCaption = true;
                }
                column(UOMCode_ProdOrderComp; "Unit of Measure Code")
                {
                    IncludeCaption = true;
                }
                column(RemainingQty_ProdOrderComp; "Remaining Quantity")
                {
                    IncludeCaption = true;
                }
                column(Scrap_ProdOrderComp; "Scrap %")
                {
                    IncludeCaption = true;
                }
                column(DueDate_ProdOrderComp; FORMAT("Due Date"))
                {
                    IncludeCaption = false;
                }
                column(LocationCode_ProdOrderComp; "Location Code")
                {
                    IncludeCaption = true;
                }
                column(InventoryAvailable; InventoryAvailable)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    WITH ReservationEntry DO BEGIN
                        SETCURRENTKEY("Source ID", "Source Ref. No.", "Source Type", "Source Subtype");

                        SETRANGE("Source Type", DATABASE::"Prod. Order Component");
                        SETRANGE("Source ID", "Production Order"."No.");
                        SETRANGE("Source Ref. No.", "Line No.");
                        SETRANGE("Source Subtype", Status);
                        SETRANGE("Source Batch Name", '');
                        SETRANGE("Source Prod. Order Line", "Prod. Order Line No.");

                        IF FINDSET THEN BEGIN
                            RemainingQtyReserved := 0;
                            REPEAT
                                IF ReservationEntry2.GET("Entry No.", NOT Positive) THEN
                                    IF (ReservationEntry2."Source Type" = DATABASE::"Prod. Order Line") AND
                                       (ReservationEntry2."Source ID" = "Prod. Order Component"."Prod. Order No.")
                                    THEN
                                        RemainingQtyReserved += ReservationEntry2."Quantity (Base)";
                            UNTIL NEXT = 0;
                            IF "Prod. Order Component"."Remaining Qty. (Base)" = RemainingQtyReserved THEN
                                CurrReport.SKIP;
                        END;
                    END;

                    //RSPLSUM 11Feb21>>
                    CLEAR(InventoryAvailable);
                    RecItem.RESET;
                    IF RecItem.GET("Prod. Order Component"."Item No.") THEN BEGIN
                        RecItem.CALCFIELDS(Inventory);
                        InventoryAvailable := RecItem.Inventory;
                    END;
                    //RSPLSUM 11Feb21<<
                end;

                trigger OnPreDataItem()
                begin
                    SETFILTER("Remaining Quantity", '<>0');
                end;
            }

            trigger OnPreDataItem()
            begin
                ProdOrderFilter := GETFILTERS;
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
        ProdOrderCompDueDateCapt = 'Due Date';
    }

    var
        ReservationEntry: Record 337;
        ReservationEntry2: Record 337;
        ProdOrderFilter: Text;
        RemainingQtyReserved: Decimal;
        ProdOrderMaterialRqstnCaptLbl: Label 'Prod. Order - Material Requisition';
        CurrReportPageNoCaptLbl: Label 'Page';
        RecItem: Record 27;
        InventoryAvailable: Decimal;
}


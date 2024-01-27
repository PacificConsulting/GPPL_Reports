report 50214 "Invoice Wise Profitability"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 29Aug2018   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/InvoiceWiseProfitability.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            column(SelltoCustomerNo_SalesInvoiceHeader; "Sales Invoice Header"."Sell-to Customer No.")
            {
            }
            column(No_SalesInvoiceHeader; "Sales Invoice Header"."No.")
            {
            }
            column(SelltoCustomerName_SalesInvoiceHeader; "Sales Invoice Header"."Sell-to Customer Name")
            {
            }
            column(ShipmentDate_SalesInvoiceHeader; "Sales Invoice Header"."Shipment Date")
            {
            }
            column(DueDate_SalesInvoiceHeader; "Sales Invoice Header"."Due Date")
            {
            }
            column(OrderDate_SalesInvoiceHeader; "Sales Invoice Header"."Order Date")
            {
            }
            column(PostingDate_SalesInvoiceHeader; "Sales Invoice Header"."Posting Date")
            {
            }
            column(LocationCode_SalesInvoiceHeader; "Sales Invoice Header"."Location Code")
            {
            }
            column(CurrencyCode_SalesInvoiceHeader; "Sales Invoice Header"."Currency Code")
            {
            }
            column(LocName; LocName)
            {
            }
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Type = CONST(Item),
                                          "Item Category Code" = FILTER(<> 'CAT14'),
                                          Quantity = FILTER(<> 0));
                column(LineType; LineType)
                {
                }
                column(No_SalesInvoiceLine; "Sales Invoice Line"."No.")
                {
                }
                column(Type_SalesInvoiceLine; "Sales Invoice Line".Type)
                {
                }
                column(LineNo_SalesInvoiceLine; "Sales Invoice Line"."Line No.")
                {
                }
                column(DocumentNo_SalesInvoiceLine; "Sales Invoice Line"."Document No.")
                {
                }
                column(Description_SalesInvoiceLine; "Sales Invoice Line".Description)
                {
                }
                column(UnitofMeasureCode_SalesInvoiceLine; "Sales Invoice Line"."Unit of Measure Code")
                {
                }
                column(Quantity_SalesInvoiceLine; "Sales Invoice Line".Quantity)
                {
                }
                column(QuantityBase_SalesInvoiceLine; "Sales Invoice Line"."Quantity (Base)")
                {
                }
                column(LineAmount_SalesInvoiceLine; "Sales Invoice Line"."Line Amount")
                {
                }
                column(UnitPrice_SalesInvoiceLine; "Sales Invoice Line"."Unit Price")
                {
                }
                column(UnitCostLCY_SalesInvoiceLine; "Sales Invoice Line"."Unit Cost (LCY)")
                {
                }
                column(SalesAmt; SalesAmt)
                {
                }
                column(CostAmountLCY; ActualAmt)
                {
                }
                column(ProfitAmt; ProfitAmt)
                {
                }
                column(ProfitPer; ProfitPer)
                {
                }
                column(UnitPrice; UnitPrice)
                {
                }
                column(CostPerUnit; CostPerUnit)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    LineType := 'Invoice';

                    ActualAmt := 0;
                    SalesAmt := 0;
                    CostPerUnit := 0;

                    VE.RESET;
                    VE.SETCURRENTKEY("Document No.");
                    VE.SETRANGE("Document No.", "Document No.");
                    VE.SETRANGE("Document Line No.", "Line No.");
                    VE.SETRANGE("Item No.", "No.");
                    VE.SETRANGE("Document Type", VE."Document Type"::"Sales Invoice");
                    IF VE.FINDFIRST THEN BEGIN
                        VE.CALCSUMS("Cost Amount (Actual)", "Sales Amount (Actual)");
                        ActualAmt := ABS(VE."Cost Amount (Actual)");
                        SalesAmt := ABS(VE."Sales Amount (Actual)");
                    END;

                    //CostPerUnit := "Unit Cost (LCY)" / "Qty. per Unit of Measure";

                    CostPerUnit := ActualAmt / "Quantity (Base)";
                    //ActualAmt := Quantity * "Unit Cost (LCY)";

                    /*
                    CLE.RESET;
                    CLE.SETCURRENTKEY("Document No.");
                    CLE.SETRANGE("Document No.","Document No.");
                    CLE.SETRANGE("Document Type",CLE."Document Type"::Invoice);
                    CLE.SETRANGE("Customer No.","Bill-to Customer No.");
                    IF CLE.FINDFIRST THEN
                      SalesAmt := CLE."Sales (LCY)"
                    ELSE
                      SalesAmt := "Line Amount";
                    */
                    UnitPrice := 0;
                    IF "Sales Invoice Header"."Currency Factor" = 0 THEN
                        UnitPrice := "Unit Price" / "Qty. per Unit of Measure"
                    ELSE
                        UnitPrice := ("Unit Price" / "Sales Invoice Header"."Currency Factor") / "Qty. per Unit of Measure";

                    ProfitAmt := 0;
                    ProfitAmt := SalesAmt - ActualAmt;

                    ProfitPer := 0;
                    IF SalesAmt <> 0 THEN
                        ProfitPer := -1 * ((ActualAmt / SalesAmt) - 1) * 100
                    ELSE
                        ProfitPer := -ActualAmt;

                end;
            }

            trigger OnAfterGetRecord()
            begin

                LocName := '';
                Loc.RESET;
                IF Loc.GET("Location Code") THEN
                    LocName := Loc.Name;
            end;

            trigger OnPreDataItem()
            begin

                SETRANGE("Posting Date", SDate, EDate);

                IF LocFilter <> '' THEN
                    SETFILTER("Location Code", LocFilter);

                IF CusFilter <> '' THEN
                    SETFILTER("Sell-to Customer No.", CusFilter);
            end;
        }
        dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
        {
            DataItemTableView = SORTING("No.");
            column(IncludeSalesCrMemo; IncludeSalesCrMemo)
            {
            }
            column(SelltoCustomerNo_SalesCrMemoHeader; "Sales Cr.Memo Header"."Sell-to Customer No.")
            {
            }
            column(No_SalesCrMemoHeader; "Sales Cr.Memo Header"."No.")
            {
            }
            column(SelltoCustomerName_SalesCrMemoHeader; "Sales Cr.Memo Header"."Sell-to Customer Name")
            {
            }
            column(ShipmentDate_SalesCrMemoHeader; "Sales Cr.Memo Header"."Shipment Date")
            {
            }
            column(DueDate_SalesCrMemoHeader; "Sales Cr.Memo Header"."Due Date")
            {
            }
            column(PostingDate_SalesCrMemoHeader; "Sales Cr.Memo Header"."Posting Date")
            {
            }
            column(LocationCode_SalesCrMemoHeader; "Sales Cr.Memo Header"."Location Code")
            {
            }
            column(CurrencyCode_SalesCrMemoHeader; "Sales Cr.Memo Header"."Currency Code")
            {
            }
            column(LocName_Cr; LocName)
            {
            }
            dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Type = CONST(Item),
                                          "Item Category Code" = FILTER(<> 'CAT14'),
                                          Quantity = FILTER(<> 0));
                column(LineType_Cr; LineType)
                {
                }
                column(No_SalesCrMemoLine; "Sales Cr.Memo Line"."No.")
                {
                }
                column(Type_SalesCrMemoLine; "Sales Cr.Memo Line".Type)
                {
                }
                column(LineNo_SalesCrMemoLine; "Sales Cr.Memo Line"."Line No.")
                {
                }
                column(DocumentNo_SalesCrMemoLine; "Sales Cr.Memo Line"."Document No.")
                {
                }
                column(Description_SalesCrMemoLine; "Sales Cr.Memo Line".Description)
                {
                }
                column(UnitofMeasureCode_SalesCrMemoLine; "Sales Cr.Memo Line"."Unit of Measure Code")
                {
                }
                column(Quantity_SalesCrMemoLine; -"Sales Cr.Memo Line".Quantity)
                {
                }
                column(QuantityBase_SalesCrMemoLine; -"Sales Cr.Memo Line"."Quantity (Base)")
                {
                }
                column(LineAmount_SalesCrMemoLine; -"Sales Cr.Memo Line"."Line Amount")
                {
                }
                column(UnitPrice_SalesCrMemoLine; -"Sales Cr.Memo Line"."Unit Price")
                {
                }
                column(UnitCostLCY_SalesCrMemoLine; -"Sales Cr.Memo Line"."Unit Cost (LCY)")
                {
                }
                column(SalesAmt_Cr; -SalesAmt)
                {
                }
                column(CostAmountLCY_Cr; -ActualAmt)
                {
                }
                column(ProfitAmt_Cr; -ProfitAmt)
                {
                }
                column(ProfitPer_Cr; ProfitPer)
                {
                }
                column(UnitPrice_Cr; -UnitPrice)
                {
                }
                column(CostPerUnit_Cr; -CostPerUnit)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    LineType := 'Credit Memo';

                    ActualAmt := 0;
                    SalesAmt := 0;
                    CostPerUnit := 0;

                    VE.RESET;
                    VE.SETCURRENTKEY("Document No.");
                    VE.SETRANGE("Document No.", "Document No.");
                    VE.SETRANGE("Document Line No.", "Line No.");
                    VE.SETRANGE("Item No.", "No.");
                    VE.SETRANGE("Document Type", VE."Document Type"::"Sales Credit Memo");
                    IF VE.FINDFIRST THEN BEGIN
                        VE.CALCSUMS("Cost Amount (Actual)", "Sales Amount (Actual)");
                        ActualAmt := ABS(VE."Cost Amount (Actual)");
                        SalesAmt := ABS(VE."Sales Amount (Actual)");
                    END;

                    //CostPerUnit := "Unit Cost (LCY)" / "Qty. per Unit of Measure";

                    CostPerUnit := ActualAmt / "Quantity (Base)";
                    //ActualAmt := Quantity * "Unit Cost (LCY)";

                    /*
                    CLE.RESET;
                    CLE.SETCURRENTKEY("Document No.");
                    CLE.SETRANGE("Document No.","Document No.");
                    CLE.SETRANGE("Document Type",CLE."Document Type"::Invoice);
                    CLE.SETRANGE("Customer No.","Bill-to Customer No.");
                    IF CLE.FINDFIRST THEN
                      SalesAmt := CLE."Sales (LCY)"
                    ELSE
                      SalesAmt := "Line Amount";
                    */
                    UnitPrice := 0;
                    IF "Sales Cr.Memo Header"."Currency Factor" = 0 THEN
                        UnitPrice := "Unit Price" / "Qty. per Unit of Measure"
                    ELSE
                        UnitPrice := ("Unit Price" / "Sales Cr.Memo Header"."Currency Factor") / "Qty. per Unit of Measure";

                    ProfitAmt := 0;
                    ProfitAmt := SalesAmt - ActualAmt;

                    ProfitPer := 0;
                    IF SalesAmt <> 0 THEN
                        ProfitPer := -1 * ((ActualAmt / SalesAmt) - 1) * 100
                    ELSE
                        ProfitPer := -ActualAmt;

                end;
            }

            trigger OnAfterGetRecord()
            begin

                LocName := '';
                Loc.RESET;
                IF Loc.GET("Location Code") THEN
                    LocName := Loc.Name;
            end;

            trigger OnPreDataItem()
            begin
                IF NOT IncludeSalesCrMemo THEN
                    CurrReport.BREAK;

                SETRANGE("Posting Date", SDate, EDate);

                IF LocFilter <> '' THEN
                    SETFILTER("Location Code", LocFilter);

                IF CusFilter <> '' THEN
                    SETFILTER("Sell-to Customer No.", CusFilter);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; SDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin

                        EDate := SDate;
                    end;
                }
                field("End Date"; EDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin

                        IF EDate < SDate THEN
                            ERROR('EndDate must be greater than StartDate');
                    end;
                }
                field("Location Filter"; LocFilter)
                {
                    ApplicationArea = all;
                    TableRelation = Location;
                }
                field("Customer Filter"; CusFilter)
                {
                    ApplicationArea = all;
                    TableRelation = Customer;
                }
                field("Include Sales Credit Memo"; IncludeSalesCrMemo)
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

    trigger OnPreReport()
    begin

        IF SDate = 0D THEN
            ERROR('Please Specify Start Date');

        IF EDate = 0D THEN
            ERROR('Please Specify End Date');
    end;

    var
        SDate: Date;
        EDate: Date;
        LocFilter: Code[100];
        CusFilter: Code[100];
        VE: Record 5802;
        SalesAmt: Decimal;
        ActualAmt: Decimal;
        ProfitAmt: Decimal;
        ProfitPer: Decimal;
        CLE: Record 21;
        Loc: Record 14;
        LocName: Text;
        UnitPrice: Decimal;
        CostPerUnit: Decimal;
        IncludeSalesCrMemo: Boolean;
        LineType: Text;
}


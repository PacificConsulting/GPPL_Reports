report 50163 "Item Sales Price History"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 06Jan2018   RB-N         New Report Development
    // 12Jan2018   RB-N         RegionWise Grouping Option
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ItemSalesPriceHistory.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Invoice Line"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("No.", "Sell-to Customer No.")
                                WHERE(Type = FILTER(Item),
                                      "Gen. Bus. Posting Group" = FILTER('DOMESTIC'),
                                      Quantity = FILTER(<> 0));
            column(RegionGrouping; Region12)
            {
            }
            column(PostingDate; "Sales Invoice Line"."Posting Date")
            {
            }
            column(Amount; Amount + AVPDisc + TSDisc + CDDisc + InvDisc)
            {
            }
            column(UnitPrice; "Sales Invoice Line"."Unit Price")
            {
            }
            column(No; "Sales Invoice Line"."No.")
            {
            }
            column(QuantityBase; "Sales Invoice Line"."Quantity (Base)")
            {
            }
            column(ItmDesc; ItmDesc)
            {
            }
            column(ItmUoM; ItmUoM)
            {
            }
            column(Qty; Quantity)
            {
            }
            column(RegionName; RegionName)
            {
            }
            column(AVPDisc; AVPDisc)
            {
            }
            column(TSDisc; TSDisc)
            {
            }
            column(CDDisc; CDDisc)
            {
            }
            column(InvDisc; InvDisc)
            {
            }
            column(ListPrice; ListPrice)
            {
            }

            trigger OnAfterGetRecord()
            begin

                Itm.RESET;
                IF Itm.GET("No.") THEN BEGIN
                    ItmDesc := Itm.Description;
                    ItmUoM := Itm."Sales Unit of Measure";
                END;

                //>>RegionName
                RegionName := '';
                LocName := '';//LocationName
                Loc20.RESET;
                IF Loc20.GET("Location Code") THEN BEGIN
                    LocName := Loc20.Name;//LocationName
                    State20.RESET;
                    IF State20.GET(Loc20."State Code") THEN
                        RegionName := ''// State20."Region Code"
                    ELSE
                        RegionName := '';
                END;
                //<<RegionName

                IF RegionName = '' THEN
                    CurrReport.SKIP;

                //>>AVP Discount
                AVPDisc := 0;

                /* PStr22.RESET;
                PStr22.SETCURRENTKEY("Invoice No.", "Item No.");
                PStr22.SETRANGE("Invoice No.", "Document No.");
                PStr22.SETRANGE("Line No.", "Line No.");
                PStr22.SETRANGE("Item No.", "No.");
                PStr22.SETRANGE("Tax/Charge Group", 'AVP-SP-SAN');
                IF PStr22.FINDFIRST THEN BEGIN
                    AVPDisc := PStr22.Amount;
                END; */
                //<<AVP Discount

                //>>Transport Discount
                TSDisc := 0;

                /*  PStr22.RESET;
                 PStr22.SETCURRENTKEY("Invoice No.", "Item No.");
                 PStr22.SETRANGE("Invoice No.", "Document No.");
                 PStr22.SETRANGE("Line No.", "Line No.");
                 PStr22.SETRANGE("Item No.", "No.");
                 PStr22.SETRANGE("Tax/Charge Group", 'TRANSSUBS');
                 IF PStr22.FINDFIRST THEN BEGIN
                     TSDisc := PStr22.Amount;
                 END; */
                //<<Transport Discount

                //>>Cash Discount
                CDDisc := 0;

                /*  PStr22.RESET;
                 PStr22.SETCURRENTKEY("Invoice No.", "Item No.");
                 PStr22.SETRANGE("Invoice No.", "Document No.");
                 PStr22.SETRANGE("Line No.", "Line No.");
                 PStr22.SETRANGE("Item No.", "No.");
                 PStr22.SETRANGE("Tax/Charge Group", 'CASHDISC');
                 IF PStr22.FINDFIRST THEN BEGIN
                     CDDisc := PStr22.Amount;
                 END; */
                //<<Cash Discount

                //>>Invoice Discount
                InvDisc := 0;

                /*  PStr22.RESET;
                 PStr22.SETCURRENTKEY("Invoice No.", "Item No.");
                 PStr22.SETRANGE("Invoice No.", "Document No.");
                 PStr22.SETRANGE("Line No.", "Line No.");
                 PStr22.SETRANGE("Item No.", "No.");
                 PStr22.SETRANGE("Tax/Charge Group", 'DISCOUNT');
                 IF PStr22.FINDFIRST THEN BEGIN
                     InvDisc := PStr22.Amount;
                 END; */
                //<<Invoice Discount

                //>>ListPrice
                ListPrice := 0;
                SPrice22.RESET;
                SPrice22.SETCURRENTKEY("Item No.");
                SPrice22.SETRANGE("Item No.", "No.");
                SPrice22.SETRANGE("Ending Date", 0D);
                IF SPrice22.FINDLAST THEN BEGIN
                    ListPrice := SPrice22."Basic Price";
                END;
                //<<ListPrice
            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("No.");
                IF InvPostFilter <> '' THEN
                    SETFILTER("Inventory Posting Group", InvPostFilter);
                SETRANGE("Posting Date", StartDate12, EndDate12);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Inventory Posting Group"; InvPostFilter)
                {
                    Style = Attention;
                    StyleExpr = TRUE;
                    TableRelation = "Inventory Posting Group";
                    ApplicationArea = all;
                    trigger OnValidate()
                    var
                        DotText: Text;
                        DotInt: Integer;
                        PipeText: Text;
                        PipeInt: Integer;
                    begin
                        DotInt := STRPOS(InvPostFilter, '.');
                        IF DotInt <> 0 THEN
                            ERROR('Select 1 Posting Group at at time');
                        PipeInt := STRPOS(InvPostFilter, '|');
                        IF PipeInt <> 0 THEN
                            ERROR('Select 1 Posting Group at at time');
                    end;
                }
                field("Start Date"; StartDate12)
                {
                    ApplicationArea = all;
                }
                field("End Date"; EndDate12)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin
                        IF EndDate12 < StartDate12 THEN
                            ERROR('End Date: %1 must greather than Start Date: %2', EndDate12, StartDate12);
                    end;
                }
                field("Region Grouping"; Region12)
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
        IF InvPostFilter = '' THEN
            ERROR('Select Inventory Posting Group ');

        IF StartDate12 = 0D THEN
            ERROR('Please Enter the Start Date');

        IF EndDate12 = 0D THEN
            ERROR('Please Enter the End Date');
    end;

    var
        Itm: Record 27;
        ItmDesc: Text;
        ItmUoM: Text;
        InvPostFilter: Code[20];
        "----------Region--12Jan2018": Integer;
        RegionName: Text;
        LocName: Text;
        Loc20: Record 14;
        State20: Record State;
        StartDate12: Date;
        EndDate12: Date;
        Region12: Boolean;
        "-----------22Jun2018": Integer;
        // PStr22: Record 13798;
        AVPDisc: Decimal;
        TSDisc: Decimal;
        CDDisc: Decimal;
        InvDisc: Decimal;
        SPrice22: Record 7002;
        ListPrice: Decimal;
}


report 50161 "RPO Items Price History"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 05Jan2018   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/RPOItemsPriceHistory.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.");
            column(No_Item; Item."No.")
            {
            }
            column(No2_Item; Item."No. 2")
            {
            }
            column(Description; Item.Description)
            {
            }
            column(SalesUoM; Item."Sales Unit of Measure")
            {
            }
            column(JAN17Qty; JAN17Qty)
            {
            }
            column(JAN17Price; JAN17Price)
            {
            }
            column(FEB17Qty; FEB17Qty)
            {
            }
            column(FEB17Price; FEB17Price)
            {
            }
            column(MAR17QTY; MAR17QTY)
            {
            }
            column(MAR17Price; MAR17Price)
            {
            }
            column(APR17QTY; APR17QTY)
            {
            }
            column(APR17Price; APR17Price)
            {
            }
            column(MAY17QTY; MAY17QTY)
            {
            }
            column(MAY17Price; MAY17Price)
            {
            }
            column(JUN17QTY; JUN17QTY)
            {
            }
            column(JUN17Price; JUN17Price)
            {
            }
            column(JUL17Qty; JUL17Qty)
            {
            }
            column(JUL17Price; JUL17Price)
            {
            }
            column(AUG17Qty; AUG17Qty)
            {
            }
            column(AUG17Price; AUG17Price)
            {
            }
            column(SEP17Qty; SEP17Qty)
            {
            }
            column(SEP17Price; SEP17Price)
            {
            }
            column(OCT17Qty; OCT17Qty)
            {
            }
            column(OCT17Price; OCT17Price)
            {
            }
            column(NOV17Qty; NOV17Qty)
            {
            }
            column(NOV17Price; NOV17Price)
            {
            }
            column(DEC17Qty; DEC17Qty)
            {
            }
            column(DEC17Price; DEC17Price)
            {
            }
            column(T17Qty; T17Qty)
            {
            }
            column(T17Price; T17Price)
            {
            }

            trigger OnAfterGetRecord()
            begin


                CLEAR(ItmFound);
                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170101D/*010117D*/, 20171231D/*311217D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    ItmFound := TRUE;

                END ELSE
                    ItmFound := FALSE;

                IF NOT ItmFound THEN
                    CurrReport.SKIP;

                CLEAR(T17Qty);
                CLEAR(TempPrice17);
                CLEAR(T17Price);


                //>>JAN2018
                CLEAR(JAN17Qty);
                CLEAR(JAN17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170101D/*010117D*/, 20170131D/*310117D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    JAN17Qty := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018
                    IF JAN17Qty <> 0 THEN
                        JAN17Price := SIL.Amount / JAN17Qty
                    ELSE
                        JAN17Price := 0;
                END;
                //MESSAGE(' %1  JANQTY \\ %2 JANPRICE ',JAN17Qty,JAN17Price);
                //<<JAN2018

                //>>FEB2018
                CLEAR(FEB17Qty);
                CLEAR(FEB17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170201D/*010217D*/, 20170228D /*280217D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    FEB17Qty := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF FEB17Qty <> 0 THEN
                        FEB17Price := SIL.Amount / FEB17Qty
                    ELSE
                        FEB17Price := 0;
                END;
                //<<FEB2018

                //>>MAR2018
                CLEAR(MAR17QTY);
                CLEAR(MAR17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170301D /*010317D*/, 20170331D/*310317D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    MAR17QTY := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF MAR17QTY <> 0 THEN
                        MAR17Price := SIL.Amount / MAR17QTY
                    ELSE
                        MAR17Price := 0;
                END;
                //<<MAR2018

                //>>APR2018
                CLEAR(APR17QTY);
                CLEAR(APR17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170401D/*010417D*/, 20170430D/*300417D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    APR17QTY := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF APR17QTY <> 0 THEN
                        APR17Price := SIL.Amount / APR17QTY
                    ELSE
                        APR17Price := 0;
                END;
                //<<APR2018

                //>>MAY017
                CLEAR(MAY17QTY);
                CLEAR(MAY17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170501D/*010517D*/, 20170531D/* 310517D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    MAY17QTY := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF MAY17QTY <> 0 THEN
                        MAY17Price := SIL.Amount / MAY17QTY
                    ELSE
                        MAY17Price := 0;
                END;
                //<<MAY2018

                //>>JUN2018
                CLEAR(JUN17QTY);
                CLEAR(JUN17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170601D/*010617D*/, 20170630D/*300617D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    JUN17QTY := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF JUN17QTY <> 0 THEN
                        JUN17Price := SIL.Amount / JUN17QTY
                    ELSE
                        JUN17Price := 0;
                END;
                //<<JUN2018

                //>>JUL2018
                CLEAR(JUL17Qty);
                CLEAR(JUL17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170701D/*010717D*/, 20170731D/* 310717D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    JUL17Qty := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF JUL17Qty <> 0 THEN
                        JUL17Price := SIL.Amount / JUL17Qty
                    ELSE
                        JUL17Price := 0;
                END;
                //<<JUL2018

                //>>AUG2018
                CLEAR(AUG17Qty);
                CLEAR(AUG17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170801D/*010817D*/, 20170831D/*310817D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    AUG17Qty := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF AUG17Qty <> 0 THEN
                        AUG17Price := SIL.Amount / AUG17Qty
                    ELSE
                        AUG17Price := 0;
                END;
                //<<AUG2018

                //>>SEP2018
                CLEAR(SEP17Qty);
                CLEAR(SEP17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20170901D/*010917D*/, 20170930D /*300917D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    SEP17Qty := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF SEP17Qty <> 0 THEN
                        SEP17Price := SIL.Amount / SEP17Qty
                    ELSE
                        SEP17Price := 0;
                END;
                //<<SEP2018

                //>>OCT2018
                CLEAR(OCT17Qty);
                CLEAR(OCT17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20171001D/*011017D*/, 20171031D/*311017D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    OCT17Qty := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF OCT17Qty <> 0 THEN
                        OCT17Price := SIL.Amount / OCT17Qty
                    ELSE
                        OCT17Price := 0;
                END;
                //<<OCT2018

                //>>NOV2018
                CLEAR(NOV17Qty);
                CLEAR(NOV17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20171101D/*011117D*/, 20171130D/*301117D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    NOV17Qty := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF NOV17Qty <> 0 THEN
                        NOV17Price := SIL.Amount / NOV17Qty
                    ELSE
                        NOV17Price := 0;
                END;
                //<<NOV2018

                //>>DEC2018
                CLEAR(DEC17Qty);
                CLEAR(DEC17Price);

                SIL.RESET;
                SIL.SETCURRENTKEY("No.");
                SIL.SETRANGE(Type, SIL.Type::Item);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Posting Date", 20171201D/*011217D*/, 20171231D/*311217D*/);
                SIL.SETRANGE("Gen. Bus. Posting Group", 'DOMESTIC');//06Jan2018
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    SIL.CALCSUMS("Quantity (Base)", Amount);
                    DEC17Qty := SIL."Quantity (Base)";
                    T17Qty += SIL."Quantity (Base)";//06Jan2018
                    TempPrice17 += SIL.Amount;//06Jan2018

                    IF DEC17Qty <> 0 THEN
                        DEC17Price := SIL.Amount / DEC17Qty
                    ELSE
                        DEC17Price := 0;
                END;
                //<<DEC2018


                //>>Total2018
                T17Price := TempPrice17 / T17Qty;
                //<<Total2018
            end;

            trigger OnPreDataItem()
            begin
                SETFILTER("Inventory Posting Group", InvPostFilter);
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
    end;

    var
        SIL: Record "Sales Invoice Line";
        JAN17Qty: Decimal;
        JAN17Price: Decimal;
        FEB17Qty: Decimal;
        FEB17Price: Decimal;
        MAR17QTY: Decimal;
        MAR17Price: Decimal;
        APR17QTY: Decimal;
        APR17Price: Decimal;
        MAY17QTY: Decimal;
        MAY17Price: Decimal;
        JUN17QTY: Decimal;
        JUN17Price: Decimal;
        JUL17Qty: Decimal;
        JUL17Price: Decimal;
        AUG17Qty: Decimal;
        AUG17Price: Decimal;
        SEP17Qty: Decimal;
        SEP17Price: Decimal;
        OCT17Qty: Decimal;
        OCT17Price: Decimal;
        NOV17Qty: Decimal;
        NOV17Price: Decimal;
        DEC17Qty: Decimal;
        DEC17Price: Decimal;
        ItmFound: Boolean;
        T17Qty: Decimal;
        T17Price: Decimal;
        TempPrice17: Decimal;
        InvPostFilter: Code[20];
}


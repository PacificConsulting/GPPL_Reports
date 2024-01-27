report 50201 "Liquidation Plan"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 07Mar2018   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/LiquidationPlan.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Sales Invoice Line"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Location Code", "No.", "Sell-to Customer No.")
                                WHERE(Type = CONST(Item),
                                      Quantity = FILTER(<> 0),
                                      "Item Category Code" = FILTER('CAT02' | 'CAT03' | 'CAT11' | 'CAT12' | 'CAT13' | 'CAT15' | 'CAT18'));
            column(ItmNo; "Sales Invoice Line"."No.")
            {
            }
            column(LocationCode; "Sales Invoice Line"."Location Code")
            {
            }
            column(Description; "Sales Invoice Line".Description)
            {
            }
            column(UoM; "Sales Invoice Line"."Unit of Measure")
            {
            }
            column(Quantity; Quantity)
            {
            }
            column(QuantityBase; "Quantity (Base)")
            {
            }
            column(UnitPrice; "Sales Invoice Line"."Unit Price")
            {
            }
            column(ItmName; ItmName)
            {
            }
            column(LocName; LocName)
            {
            }
            column(RegName; RegName)
            {
            }
            column(SDate; SDate)
            {
            }
            column(EDate; EDate)
            {
            }
            column(StartDate1; StartDate[1])
            {
            }
            column(EndDate1; EndDate[1])
            {
            }
            column(StartDate2; StartDate[2])
            {
            }
            column(EndDate2; EndDate[2])
            {
            }
            column(StartDate3; StartDate[3])
            {
            }
            column(EndDate3; EndDate[3])
            {
            }
            column(StartDate4; StartDate[4])
            {
            }
            column(EndDate4; EndDate[4])
            {
            }
            column(StartDate5; StartDate[5])
            {
            }
            column(EndDate5; EndDate[5])
            {
            }
            column(StartDate6; StartDate[6])
            {
            }
            column(EndDate6; EndDate[6])
            {
            }
            column(StartDate7; StartDate[7])
            {
            }
            column(EndDate7; EndDate[7])
            {
            }
            column(StartDate8; StartDate[8])
            {
            }
            column(EndDate8; EndDate[8])
            {
            }
            column(StartDate9; StartDate[9])
            {
            }
            column(EndDate9; EndDate[9])
            {
            }
            column(StartDate10; StartDate[10])
            {
            }
            column(EndDate10; EndDate[10])
            {
            }
            column(StartDate11; StartDate[11])
            {
            }
            column(EndDate11; EndDate[11])
            {
            }
            column(StartDate12; StartDate[12])
            {
            }
            column(EndDate12; EndDate[12])
            {
            }
            column(Qty1; Qty[1])
            {
            }
            column(Qty2; Qty[2])
            {
            }
            column(Qty3; Qty[3])
            {
            }
            column(Qty4; Qty[4])
            {
            }
            column(Qty5; Qty[5])
            {
            }
            column(Qty6; Qty[6])
            {
            }
            column(Qty7; Qty[7])
            {
            }
            column(Qty8; Qty[8])
            {
            }
            column(Qty9; Qty[9])
            {
            }
            column(Qty10; Qty[10])
            {
            }
            column(Qty11; Qty[11])
            {
            }
            column(Qty12; Qty[12])
            {
            }
            column(StockQty; StockQty)
            {
            }
            column(IndentQty; IndentQty)
            {
            }
            column(PendingOrderQty; PendingOrderQty)
            {
            }
            column(MaxQty; MaxQty)
            {
            }
            column(IndentQty2; IndentQty2)
            {
            }

            trigger OnAfterGetRecord()
            begin

                ItmName := '';
                Itm.RESET;
                IF Itm.GET("No.") THEN BEGIN
                    ItmName := Itm.Description;
                    //ItmUoM   := Itm."Sales Unit of Measure";
                END;

                LocName := '';
                RegName := '';
                Loc.RESET;
                IF Loc.GET("Location Code") THEN BEGIN
                    LocName := Loc.Name;//LocationName
                    St.RESET;
                    IF St.GET(Loc."State Code") THEN
                        RegName := ''//St."Region Code"
                    ELSE
                        RegName := '';
                END;

                CLEAR(Qty);
                //>>Qty1
                IF ("Posting Date" >= StartDate[1]) AND ("Posting Date" <= EndDate[1]) THEN
                    Qty[1] := "Quantity (Base)";
                TQty[1] += Qty[1];
                //<<Qty1

                //>>Qty2
                IF ("Posting Date" >= StartDate[2]) AND ("Posting Date" <= EndDate[2]) THEN
                    Qty[2] := "Quantity (Base)";
                TQty[2] += Qty[2];
                //<<Qty2

                //>>Qty3
                IF ("Posting Date" >= StartDate[3]) AND ("Posting Date" <= EndDate[3]) THEN
                    Qty[3] := "Quantity (Base)";
                TQty[3] += Qty[3];
                //<<Qty3

                //>>Qty4
                IF ("Posting Date" >= StartDate[4]) AND ("Posting Date" <= EndDate[4]) THEN
                    Qty[4] := "Quantity (Base)";
                TQty[4] += Qty[4];
                //<<Qty4

                //>>Qty5
                IF ("Posting Date" >= StartDate[5]) AND ("Posting Date" <= EndDate[5]) THEN
                    Qty[5] := "Quantity (Base)";
                TQty[5] += Qty[5];
                //<<Qty5

                //>>Qty6
                IF ("Posting Date" >= StartDate[6]) AND ("Posting Date" <= EndDate[6]) THEN
                    Qty[6] := "Quantity (Base)";
                TQty[6] += Qty[6];
                //<<Qty6

                //>>Qty7
                IF ("Posting Date" >= StartDate[7]) AND ("Posting Date" <= EndDate[7]) THEN
                    Qty[7] := "Quantity (Base)";
                TQty[7] += Qty[7];
                //<<Qty7

                //>>Qty8
                IF ("Posting Date" >= StartDate[8]) AND ("Posting Date" <= EndDate[8]) THEN
                    Qty[8] := "Quantity (Base)";
                TQty[8] += Qty[8];
                //<<Qty8

                //>>Qty9
                IF ("Posting Date" >= StartDate[9]) AND ("Posting Date" <= EndDate[9]) THEN
                    Qty[9] := "Quantity (Base)";
                TQty[9] += Qty[9];
                //<<Qty9

                //>>Qty10
                IF ("Posting Date" >= StartDate[10]) AND ("Posting Date" <= EndDate[10]) THEN
                    Qty[10] := "Quantity (Base)";
                TQty[10] += Qty[10];
                //<<Qty10

                //>>Qty11
                IF ("Posting Date" >= StartDate[11]) AND ("Posting Date" <= EndDate[11]) THEN
                    Qty[11] := "Quantity (Base)";
                TQty[11] += Qty[11];
                //<<Qty11

                //>>Qty12
                IF ("Posting Date" >= StartDate[12]) AND ("Posting Date" <= EndDate[12]) THEN
                    Qty[12] := "Quantity (Base)";
                TQty[12] += Qty[12];
                //<<Qty12


                //>>
                SIL.RESET;
                SIL.SETCURRENTKEY("No.", "Location Code");
                SIL.SETRANGE(Type, Type);
                SIL.SETRANGE("No.", "No.");
                SIL.SETRANGE("Location Code", "Location Code");
                SIL.SETRANGE("Posting Date", StartDate[12], EndDate[1]);
                SIL.SETFILTER(Quantity, '<>%1', 0);
                IF SIL.FINDFIRST THEN BEGIN
                    III := SIL.COUNT;
                END;

                JJJ += 1;
                //<<

                CLEAR(StockQty);
                CLEAR(PendingOrderQty);
                CLEAR(IndentQty);
                CLEAR(IndentQty2);
                CLEAR(MaxQty);
                CLEAR(QtyPer);
                IF III = JJJ THEN BEGIN

                    //>>StockQty
                    CLEAR(StockQty);
                    Itm.RESET;
                    Itm.SETRANGE("No.", "No.");
                    Itm.SETFILTER("Location Filter", "Location Code");
                    Itm.SETFILTER("Date Filter", '%1..%2', 0D, ReqDate);
                    IF Itm.FINDFIRST THEN BEGIN
                        Itm.CALCFIELDS(Inventory);
                        StockQty := Itm.Inventory;
                    END;
                    //<<StockQty

                    //>>PendingOrderQty
                    CLEAR(PendingOrderQty);

                    SL.RESET;
                    SL.SETCURRENTKEY("No.");
                    SL.SETRANGE("Document Type", SL."Document Type"::Order);
                    SL.SETRANGE("No.", "No.");
                    SL.SETRANGE(Type, Type);
                    SL.SETRANGE("Location Code", "Location Code");
                    SL.SETRANGE("Posting Date", TODAY - 30, TODAY);
                    //SL.SETFILTER("Qty. Shipped (Base)",'<>%1',0);
                    IF SL.FINDSET THEN
                        REPEAT
                            SH.RESET;
                            IF SH.GET(SL."Document Type", SL."Document No.") THEN BEGIN
                                CLEAR(Temp1);
                                IF SH."Short Close" THEN
                                    Temp1 := FALSE;
                                IF SH.Status = SH.Status::Released THEN
                                    Temp1 := TRUE;

                                IF Temp1 THEN
                                    PendingOrderQty += SL."Outstanding Qty. (Base)"
                            END;
                        UNTIL SL.NEXT = 0;
                    //<<PendingOrderQty

                    //>>IndentQty
                    CLEAR(QtyPer);
                    CLEAR(IndentQty);
                    CLEAR(IndentQty2);
                    CLEAR(Temp2);
                    TIL.SETCURRENTKEY("Item No.");
                    TIL.SETRANGE("Item No.", "No.");
                    TIL.SETRANGE("Transfer-to Code", "Location Code");
                    TIL.SETRANGE(Status, TIL.Status::Released);
                    TIL.SETRANGE("Addition date", TODAY - 50, TODAY);
                    TIL.SETFILTER("Outstanding Quantity", '<>%1', 0);
                    IF TIL.FINDSET THEN
                        REPEAT
                            ItmUoM.RESET;
                            IF ItmUoM.GET(TIL."Item No.", TIL."Unit of Measure Code") THEN
                                QtyPer := ItmUoM."Qty. per Unit of Measure";

                            TIH.RESET;
                            IF TIH.GET(TIL."Document No.") THEN BEGIN
                                Temp2 := TRUE;
                                IF TIH."Short Closed" THEN
                                    Temp2 := FALSE;
                                IF Temp2 THEN BEGIN
                                    IF TIL."Transfer-from Code" = 'PLANT01' THEN
                                        IndentQty += TIL."Outstanding Quantity" * QtyPer;

                                    IF TIL."Transfer-from Code" = 'PLANT02' THEN
                                        IndentQty2 += TIL."Outstanding Quantity" * QtyPer;
                                END;
                            END;
                        UNTIL TIL.NEXT = 0;
                    //<<IndentQty

                    //>>MaxQty
                    CLEAR(MaxQty);
                    IF TQty[1] <= TQty[2] THEN
                        MaxQty := TQty[2]
                    ELSE
                        MaxQty := TQty[1];

                    IF MaxQty <= TQty[3] THEN
                        MaxQty := TQty[3];

                    IF MaxQty <= TQty[4] THEN
                        MaxQty := TQty[4];

                    IF MaxQty <= TQty[5] THEN
                        MaxQty := TQty[5];

                    IF MaxQty <= TQty[6] THEN
                        MaxQty := TQty[6];

                    IF MaxQty <= TQty[7] THEN
                        MaxQty := TQty[7];

                    IF MaxQty <= TQty[8] THEN
                        MaxQty := TQty[8];

                    IF MaxQty <= TQty[9] THEN
                        MaxQty := TQty[9];

                    IF MaxQty <= TQty[10] THEN
                        MaxQty := TQty[10];

                    IF MaxQty <= TQty[11] THEN
                        MaxQty := TQty[11];

                    IF MaxQty <= TQty[12] THEN
                        MaxQty := TQty[12];

                    //>>MaxQty
                    CLEAR(TQty);
                    JJJ := 0;
                END;
            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("No.", "Location Code");
                SETRANGE("Posting Date", StartDate[12], EndDate[1]);
                //SETRANGE("No.",'FGIL0133');
                //SETRANGE("Location Code",'BR0012');
                IF CatOption = CatOption::"CAT02-AUTOOILS" THEN
                    SETRANGE("Item Category Code", 'CAT02');
                IF CatOption = CatOption::"CAT03-INDOILS" THEN
                    SETRANGE("Item Category Code", 'CAT03');
                IF CatOption = CatOption::"CAT11-RPO" THEN
                    SETRANGE("Item Category Code", 'CAT11');
                IF CatOption = CatOption::"CAT12-BOILS" THEN
                    SETRANGE("Item Category Code", 'CAT12');
                IF CatOption = CatOption::"CAT13-TOILS" THEN
                    SETRANGE("Item Category Code", 'CAT13');
                IF CatOption = CatOption::"CAT15-REPSOL" THEN
                    SETRANGE("Item Category Code", 'CAT15');
                IF CatOption = CatOption::"CAT18-CEPSA" THEN
                    SETRANGE("Item Category Code", 'CAT18');

                III := 0;
                JJJ := 0;
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
                    Visible = false;
                }
                field("End Date"; EDate)
                {
                    Visible = false;
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin
                        IF EDate < SDate THEN
                            ERROR('End Date: %1 must greather than Start Date: %2', EDate, SDate);
                    end;
                }
                field("Requested Date"; ReqDate)
                {
                    ApplicationArea = all;
                }
                field("Item Category Code"; CatOption)
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

        IF ReqDate = 0D THEN
            ERROR('Please Enter the Start Date');

        ReqMonth := DATE2DMY(ReqDate, 2);
        ReqMonth1 := ReqMonth - 1;

        ReqYear := DATE2DMY(ReqDate, 3);

        IF ReqMonth1 = 0 THEN
            ReqMonth1 := 12;

        IF ReqMonth1 = 12 THEN
            ReqYear := ReqYear - 1;

        ReqDate1 := DMY2DATE(1, ReqMonth1, ReqYear);

        CLEAR(StartDate);
        CLEAR(EndDate);

        StartDate[1] := ReqDate1;
        EndDate[1] := CALCDATE('CM', StartDate[1]);

        StartDate[2] := CALCDATE('-1M', StartDate[1]);
        EndDate[2] := CALCDATE('CM', StartDate[2]);

        StartDate[3] := CALCDATE('-1M', StartDate[2]);
        EndDate[3] := CALCDATE('CM', StartDate[3]);

        StartDate[4] := CALCDATE('-1M', StartDate[3]);
        EndDate[4] := CALCDATE('CM', StartDate[4]);

        StartDate[5] := CALCDATE('-1M', StartDate[4]);
        EndDate[5] := CALCDATE('CM', StartDate[5]);

        StartDate[6] := CALCDATE('-1M', StartDate[5]);
        EndDate[6] := CALCDATE('CM', StartDate[6]);

        StartDate[7] := CALCDATE('-1M', StartDate[6]);
        EndDate[7] := CALCDATE('CM', StartDate[7]);

        StartDate[8] := CALCDATE('-1M', StartDate[7]);
        EndDate[8] := CALCDATE('CM', StartDate[8]);

        StartDate[9] := CALCDATE('-1M', StartDate[8]);
        EndDate[9] := CALCDATE('CM', StartDate[9]);

        StartDate[10] := CALCDATE('-1M', StartDate[9]);
        EndDate[10] := CALCDATE('CM', StartDate[10]);

        StartDate[11] := CALCDATE('-1M', StartDate[10]);
        EndDate[11] := CALCDATE('CM', StartDate[11]);

        StartDate[12] := CALCDATE('-1M', StartDate[11]);
        EndDate[12] := CALCDATE('CM', StartDate[12]);
    end;

    var
        SDate: Date;
        EDate: Date;
        Loc: Record 14;
        Itm: Record 27;
        St: Record State;
        ItmName: Text;
        LocName: Text;
        RegName: Text;
        ReqDate: Date;
        ReqDate1: Date;
        ReqMonth: Integer;
        ReqMonth1: Integer;
        ReqYear: Integer;
        ReqYear1: Integer;
        StartDate: array[12] of Date;
        EndDate: array[12] of Date;
        TQty: array[12] of Decimal;
        Qty: array[12] of Decimal;
        MaxQty: Decimal;
        SIL: Record 113;
        StockQty: Decimal;
        PendingOrderQty: Decimal;
        SH: Record 36;
        SL: Record 37;
        III: Integer;
        JJJ: Integer;
        Temp1: Boolean;
        Temp2: Boolean;
        TIH: Record 50022;
        TIL: Record 50023;
        IndentQty: Decimal;
        IndentQty2: Decimal;
        CatOption: Option " ","CAT02-AUTOOILS","CAT03-INDOILS","CAT11-RPO","CAT12-BOILS","CAT13-TOILS","CAT15-REPSOL","CAT18-CEPSA";
        ItmUoM: Record 5404;
        QtyPer: Decimal;
}


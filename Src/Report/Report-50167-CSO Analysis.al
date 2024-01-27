report 50167 "CSO Analysis"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 20Dec2017   RB-N         New Report Development
    // 06Feb2018   RB-N         Skipping MERCH Category, Displaying SalesPersons & OrderDate
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CSOAnalysis.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'CSO Analysis';

    dataset
    {
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number);
            MaxIteration = 1;
            column(PrintSummary; PrintSummary)
            {
            }
        }
        dataitem("Sales Header Archive"; "Sales Header Archive")
        {
            DataItemTableView = SORTING("Document Type", "No.", "Doc. No. Occurrence", "Version No.")
                                WHERE("Document Type" = FILTER(Order));
            column(DocumentType; "Document Type")
            {
            }
            column(DocNo; "No.")
            {
            }
            column(StartDate; StartDate)
            {
            }
            column(EndDate; EndDate)
            {
            }
            dataitem("Sales Line Archive"; "Sales Line Archive")
            {
                DataItemLink = "Document Type" = FIELD("Document Type"),
                               "Document No." = FIELD("No."),
                               "Version No." = FIELD("Version No.");
                DataItemTableView = SORTING("Document Type", "Document No.", "Line No.")
                                    WHERE("Item Category Code" = FILTER(<> 'CAT14'));

                trigger OnAfterGetRecord()
                begin
                    /*
                    //>>Quantity Dispatch
                    QtyDispatch := 0;
                    //IF "Quantity Invoiced" = 0 THEN
                    IF "Qty. Invoiced (Base)" = 0 THEN
                    BEGIN
                      SIH20.RESET;
                      SIH20.SETCURRENTKEY("Order No.");
                      SIH20.SETRANGE("Order No.","Document No.");
                      IF SIH20.FINDSET THEN
                      REPEAT
                    
                        SIL20.RESET;
                        SIL20.SETRANGE("Document No.",SIH20."No.");
                        SIL20.SETRANGE("Line No.","Line No.");
                        SIL20.SETRANGE(Type,Type);
                        SIL20.SETRANGE("No.","No.");
                        SIL20.SETFILTER(Quantity,'<>%1',0);
                        IF SIL20.FINDFIRST THEN
                        BEGIN
                            //QtyDispatch := SIL20.Quantity;
                          QtyDispatch := SIL20."Quantity (Base)";
                        END;
                      UNTIL SIH20.NEXT = 0;
                    END ELSE
                    BEGIN
                      //QtyDispatch := "Quantity Invoiced";
                      QtyDispatch := "Qty. Invoiced (Base)";
                    END;
                    //<<Quantity Dispatch
                    
                    
                    //>>Item Description
                    ItmName := '';
                    IF Type = Type::Item THEN
                    BEGIN
                      Itm20.RESET;
                      IF Itm20.GET("No.") THEN
                        ItmName := Itm20.Description;
                    
                    END ELSE
                    BEGIN
                      ItmName := Description;
                    END;
                    
                    //<<Item Description
                    */
                    //

                    //>>TempSalesLineArc
                    TempSalesLineArc.INIT;
                    TempSalesLineArc.TRANSFERFIELDS("Sales Line Archive");
                    TempSalesLineArc.INSERT;
                    //<<TempSalesLineArc

                end;
            }

            trigger OnAfterGetRecord()
            begin


                //>>19Jul2019
                SH19.RESET;
                SH19.SETRANGE("Document Type", "Document Type");
                SH19.SETRANGE("No.", "No.");
                IF SH19.FINDFIRST THEN
                    CurrReport.SKIP;
                //<<19Jul2019

                //>>20Dec2017 Skipping Other Version
                TempSAH.RESET;
                TempSAH.SETCURRENTKEY("Document Type", "No.", "Version No.");
                TempSAH.SETRANGE("Document Type", "Document Type");
                TempSAH.SETRANGE("No.", "No.");
                TempSAH.SETRANGE("Posting Date", StartDate, EndDate);
                IF TempSAH.FINDLAST THEN BEGIN
                    IF TempSAH."Version No." <> "Version No." THEN
                        CurrReport.SKIP;
                END;

                IF "Shortcut Dimension 1 Code" = '' THEN
                    CurrReport.SKIP;
                //<<20Dec2017 Skipping Other Version

                /*
                //>>RegionName
                RegionName := '';
                LocName := '';//LocationName
                Loc20.RESET;
                IF Loc20.GET("Location Code") THEN
                BEGIN
                  LocName := Loc20.Name;//LocationName
                  State20.RESET;
                  IF State20.GET(Loc20."State Code") THEN
                    RegionName := State20."Region Code"
                  ELSE
                    RegionName := '';
                END;
                //<<RegionName
                
                //>>DivisionName
                DimName := '';
                DimVal20.RESET;
                IF DimVal20.GET('DIVISION',"Shortcut Dimension 1 Code") THEN
                BEGIN
                  DimName := DimVal20.Name
                END;
                //<<DivisionName
                */

            end;

            trigger OnPreDataItem()
            begin

                SETCURRENTKEY("Shortcut Dimension 1 Code");
                SETRANGE("Date Archived", StartDate, EndDate);
                IF DivFilter <> '' THEN
                    SETFILTER("Shortcut Dimension 1 Code", DivFilter);
                IF LocFilter <> '' THEN
                    SETFILTER("Location Code", LocFilter);

                //SETRANGE("No.",'CSO/14/1920/0129');
            end;
        }
        dataitem("Sales Header"; "Sales Header")
        {
            DataItemTableView = SORTING("Document Type", "No.")
                                WHERE("Document Type" = FILTER(Order));
            dataitem("Sales Line"; "Sales Line")
            {
                DataItemLink = "Document Type" = FIELD("Document Type"),
                               "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document Type", "Document No.", "Line No.")
                                    WHERE("Item Category Code" = FILTER(<> 'CAT14'));

                trigger OnAfterGetRecord()
                begin

                    //>>TempSalesLineArc
                    TempSalesLineArc.INIT;
                    TempSalesLineArc.TRANSFERFIELDS("Sales Line");
                    TempSalesLineArc."Version No." := 1;
                    TempSalesLineArc."Doc. No. Occurrence" := 1;
                    TempSalesLineArc.INSERT;
                    //<<TempSalesLineArc
                end;
            }

            trigger OnAfterGetRecord()
            begin

                SHCOUNT += 1;
                /*
                IF SHCOUNT = 1 THEN
                  MESSAGE(' %1 Count',COUNT);
                  */

            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("Shortcut Dimension 1 Code");
                SETRANGE("Created Date", StartDate, EndDate);
                IF DivFilter <> '' THEN
                    SETFILTER("Shortcut Dimension 1 Code", DivFilter);
                IF LocFilter <> '' THEN
                    SETFILTER("Location Code", LocFilter);

                //SETRANGE("No.",'CSO/14/1920/0129');

                SHCOUNT := 0;
            end;
        }
        dataitem(TempSalesLineArc; "Sales Line Archive")
        {
            DataItemTableView = SORTING("Document Type", "Document No.", "Doc. No. Occurrence", "Version No.", "Line No.");
            UseTemporary = true;
            column(DivisionCode; "Shortcut Dimension 1 Code")
            {
            }
            column(DimName; DimName)
            {
            }
            column(VersionNo; "Version No.")
            {
            }
            column(LocationCode; LocName)
            {
            }
            column(RegionName; RegionName)
            {
            }
            column(LineDocNo; "Document No.")
            {
            }
            column(OrderDate; OrdDate)
            {
            }
            column(DocumentDate; DocDate)
            {
            }
            column(LineNo; "Line No.")
            {
            }
            column(ItmNo; "No.")
            {
            }
            column(ItmName; ItmName)
            {
            }
            column(SPName; SPName)
            {
            }
            column(QtyShipped; "Quantity Shipped")
            {
            }
            column(QtyDispatch; QtyDispatch)
            {
            }
            column(QtyOrder; "Quantity (Base)")
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>RegionName
                RegionName := '';
                LocName := '';//LocationName
                Loc20.RESET;
                IF Loc20.GET("Location Code") THEN BEGIN
                    LocName := Loc20.Name;//LocationName
                    State20.RESET;
                    IF State20.GET(Loc20."State Code") THEN
                        RegionName := '' //State20."Region Code"
                    ELSE
                        RegionName := '';
                END;
                //<<RegionName

                //>>DivisionName
                DimName := '';
                DimVal20.RESET;
                IF DimVal20.GET('DIVISION', "Shortcut Dimension 1 Code") THEN BEGIN
                    DimName := DimVal20.Name
                END;
                //<<DivisionName

                //>>Quantity Dispatch
                QtyDispatch := 0;
                //IF "Quantity Invoiced" = 0 THEN
                IF "Qty. Invoiced (Base)" = 0 THEN BEGIN
                    SIH20.RESET;
                    SIH20.SETCURRENTKEY("Order No.");
                    SIH20.SETRANGE("Order No.", "Document No.");
                    IF SIH20.FINDSET THEN
                        REPEAT

                            SIL20.RESET;
                            SIL20.SETRANGE("Document No.", SIH20."No.");
                            //SIL20.SETRANGE("Line No.","Line No.");
                            SIL20.SETRANGE(Type, Type);
                            SIL20.SETRANGE("No.", "No.");
                            SIL20.SETFILTER(Quantity, '<>%1', 0);
                            //IF "Lot No." <> '' THEN
                            //SIL20.SETRANGE("Lot No.","Lot No.");
                            IF SIL20.FINDSET THEN
                                REPEAT
                                    //QtyDispatch := SIL20.Quantity;
                                    QtyDispatch += SIL20."Quantity (Base)";//02Jan2018
                                UNTIL SIL20.NEXT = 0;
                        UNTIL SIH20.NEXT = 0;
                    //>>05Jan2019
                    IF QtyDispatch > TempSalesLineArc."Quantity (Base)" THEN
                        QtyDispatch := TempSalesLineArc."Quantity (Base)";
                    //<<05Jan2019
                END ELSE BEGIN
                    //QtyDispatch := "Quantity Invoiced";
                    QtyDispatch := "Qty. Invoiced (Base)";
                END;
                //<<Quantity Dispatch


                //>>Item Description
                ItmName := '';
                IF Type = Type::Item THEN BEGIN
                    Itm20.RESET;
                    IF Itm20.GET("No.") THEN
                        ItmName := Itm20.Description;

                END ELSE BEGIN
                    ItmName := Description;
                END;

                //<<Item Description

                //>>SalesPerson Name
                SPName := '';
                CLEAR(OrdDate);
                CLEAR(DocDate);
                SHA06.RESET;
                SHA06.SETCURRENTKEY("Document Type", "No.");
                SHA06.SETRANGE("Document Type", "Document Type");
                SHA06.SETRANGE("No.", "Document No.");
                IF SHA06.FINDFIRST THEN BEGIN
                    OrdDate := SHA06."Order Date";
                    DocDate := SHA06."Document Date";
                    SalesPersons.RESET;
                    IF SalesPersons.GET(SHA06."Salesperson Code") THEN BEGIN
                        SPName := SalesPersons.Name;
                    END;
                END;

                IF SPName = '' THEN BEGIN
                    SH06.RESET;
                    SH06.SETCURRENTKEY("Document Type", "No.");
                    SH06.SETRANGE("Document Type", "Document Type");
                    SH06.SETRANGE("No.", "Document No.");
                    IF SH06.FINDFIRST THEN BEGIN
                        OrdDate := SH06."Order Date";
                        DocDate := SH06."Document Date";
                        SalesPersons.RESET;
                        IF SalesPersons.GET(SH06."Salesperson Code") THEN BEGIN
                            SPName := SalesPersons.Name;
                        END;
                    END;
                END;
                //<<SalesPerson Name
            end;

            trigger OnPreDataItem()
            begin

                SETCURRENTKEY("Shortcut Dimension 1 Code");
            end;
        }
        dataitem(SalesHeaderLastYear; "Sales Header Archive")
        {
            DataItemTableView = SORTING("Document Type", "No.", "Doc. No. Occurrence", "Version No.")
                                WHERE("Document Type" = FILTER(Order));
            column(DocumentType1; "Document Type")
            {
            }
            column(DocNo1; "No.")
            {
            }
            column(DivisionCode1; "Shortcut Dimension 1 Code")
            {
            }
            column(DimName1; DimName)
            {
            }
            column(VersionNo1; "Version No.")
            {
            }
            column(LocationCode1; LocName)
            {
            }
            column(RegionName1; RegionName)
            {
            }
            column(StartDate1; StartDate1)
            {
            }
            column(EndDate1; EndDate1)
            {
            }
            column(SPName1; SPName1)
            {
            }
            column(OrderDate1; "Order Date")
            {
            }
            dataitem(SalesLineLastYear; "Sales Line Archive")
            {
                DataItemLink = "Document Type" = FIELD("Document Type"),
                               "Document No." = FIELD("No."),
                               "Version No." = FIELD("Version No.");
                DataItemTableView = SORTING("Document Type", "Document No.", "Line No.")
                                    WHERE("Item Category Code" = FILTER(<> 'CAT14'));
                column(LineDocNo1; "Document No.")
                {
                }
                column(LineNo1; "Line No.")
                {
                }
                column(ItmNo1; "No.")
                {
                }
                column(ItmName1; ItmName)
                {
                }
                column(QtyShipped1; "Quantity Shipped")
                {
                }
                column(QtyDispatch1; QtyDispatch)
                {
                }
                column(QtyOrder1; "Quantity (Base)")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>Quantity Dispatch
                    QtyDispatch := 0;
                    //IF "Quantity Invoiced" = 0 THEN
                    IF "Qty. Invoiced (Base)" = 0 THEN BEGIN
                        SIH20.RESET;
                        SIH20.SETCURRENTKEY("Order No.");
                        SIH20.SETRANGE("Order No.", "Document No.");
                        IF SIH20.FINDSET THEN
                            REPEAT

                                SIL20.RESET;
                                SIL20.SETRANGE("Document No.", SIH20."No.");
                                SIL20.SETRANGE("Line No.", "Line No.");
                                SIL20.SETRANGE(Type, Type);
                                SIL20.SETRANGE("No.", "No.");
                                SIL20.SETFILTER(Quantity, '<>%1', 0);
                                IF SIL20.FINDFIRST THEN BEGIN
                                    //QtyDispatch := SIL20.Quantity;
                                    QtyDispatch := SIL20."Quantity (Base)";
                                END;
                            UNTIL SIH20.NEXT = 0;
                    END ELSE BEGIN
                        //QtyDispatch := "Quantity Invoiced";
                        QtyDispatch := "Qty. Invoiced (Base)";
                    END;
                    //<<Quantity Dispatch


                    //>>Item Description
                    ItmName := '';
                    IF Type = Type::Item THEN BEGIN
                        Itm20.RESET;
                        IF Itm20.GET("No.") THEN
                            ItmName := Itm20.Description;

                    END ELSE BEGIN
                        ItmName := Description;
                    END;

                    //<<Item Description
                end;
            }

            trigger OnAfterGetRecord()
            begin

                IF PrintSummary THEN
                    CurrReport.SKIP;

                //>>21Dec2017 Skipping Other Version
                TempSAH.RESET;
                TempSAH.SETCURRENTKEY("Document Type", "No.", "Version No.");
                TempSAH.SETRANGE("Document Type", "Document Type");
                TempSAH.SETRANGE("No.", "No.");
                TempSAH.SETRANGE("Posting Date", StartDate1, EndDate1);
                IF TempSAH.FINDLAST THEN BEGIN

                    IF TempSAH."Version No." <> "Version No." THEN
                        CurrReport.SKIP;
                END;

                IF "Shortcut Dimension 1 Code" = '' THEN
                    CurrReport.SKIP;
                //<<21Dec2017 Skipping Other Version

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

                //>>DivisionName
                DimName := '';
                DimVal20.RESET;
                IF DimVal20.GET('DIVISION', "Shortcut Dimension 1 Code") THEN BEGIN
                    DimName := DimVal20.Name
                END;
                //<<DivisionName

                //>>SalesPerson Name
                SPName1 := '';
                SalesPersons.RESET;
                IF SalesPersons.GET("Salesperson Code") THEN BEGIN
                    SPName1 := SalesPersons.Name;
                END;
                //>>SalesPerson Name
            end;

            trigger OnPreDataItem()
            begin

                SETCURRENTKEY("Shortcut Dimension 1 Code");
                SETRANGE("Date Archived", StartDate1, EndDate1);
                IF DivFilter <> '' THEN
                    SETFILTER("Shortcut Dimension 1 Code", DivFilter);
                IF LocFilter <> '' THEN
                    SETFILTER("Location Code", LocFilter);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; StartDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin
                        EndDate := StartDate;
                    end;
                }
                field("End Date"; EndDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin
                        IF EndDate < StartDate THEN
                            ERROR('End Date %1 must be Greater Than Start Date %2', EndDate, StartDate);
                    end;
                }
                field(Division; DivFilter)
                {
                    ApplicationArea = all;
                    TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(1));
                }
                field(Location; LocFilter)
                {
                    ApplicationArea = all;
                    TableRelation = Location;
                }
                field(Summary; PrintSummary)
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

        TempSalesLineArc.DELETEALL;//05Jan2018
    end;

    trigger OnPreReport()
    begin

        TempSalesLineArc.DELETEALL;//05Jan2018

        IF StartDate = 0D THEN
            ERROR('Please Enter Start Date');

        IF EndDate = 0D THEN
            ERROR('Please Enter End Date');

        //>>LastYear StartDate
        CLEAR(I1);
        CLEAR(Day1);
        CLEAR(Month1);
        CLEAR(Year1);

        Year1 := DATE2DMY(StartDate, 3);
        Month1 := DATE2DMY(StartDate, 2);
        Day1 := DATE2DMY(StartDate, 1);

        I1 := (Year1 - 1) MOD 4;

        IF (I1 = 0) AND (Month1 = 2) THEN
            StartDate1 := StartDate - 366
        ELSE
            StartDate1 := StartDate - 365;
        //<<LastYear StartDate

        //>>LastYear EndDate
        CLEAR(I1);
        CLEAR(Day1);
        CLEAR(Month1);
        CLEAR(Year1);

        Year1 := DATE2DMY(EndDate, 3);
        Month1 := DATE2DMY(EndDate, 2);
        Day1 := DATE2DMY(EndDate, 1);

        I1 := (Year1 - 1) MOD 4;

        IF (I1 = 0) AND (Month1 = 2) THEN
            EndDate1 := EndDate - 366
        ELSE
            EndDate1 := EndDate - 365;
        //<<LastYear EndDate
    end;

    var
        DivFilter: Text;
        LocFilter: Text;
        StartDate: Date;
        EndDate: Date;
        TempSAH: Record 5107;
        Loc20: Record 14;
        State20: Record State;
        RegionName: Code[20];
        SIH20: Record 112;
        SIL20: Record 113;
        QtyDispatch: Decimal;
        DimVal20: Record 349;
        DimName: Code[50];
        LocName: Code[50];
        Itm20: Record 27;
        ItmName: Text[50];
        "--------------------LastYear": Integer;
        StartDate1: Date;
        EndDate1: Date;
        Day1: Integer;
        Month1: Integer;
        Year1: Integer;
        I1: Integer;
        PrintSummary: Boolean;
        "---------------": Integer;
        SHCOUNT: Integer;
        "---------05Feb2018": Integer;
        SalesPersons: Record 13;
        SH06: Record 36;
        SHA06: Record 5107;
        SPName: Text;
        SPName1: Text;
        OrdDate: Date;
        DocDate: Date;
        SH19: Record 36;
}


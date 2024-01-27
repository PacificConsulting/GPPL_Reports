report 50222 "Binwise Consumption Report"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 10Sep2019   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BinwiseConsumptionReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem(WareEntry; "Warehouse Entry")
        {
            DataItemTableView = SORTING("Location Code", "Item No.", "Variant Code", "Zone Code", "Bin Code", "Lot No.");
            column(ToBinCode; ToBinCode)
            {
            }
            column(COMPANYNAME; COMPANYNAME)
            {
            }
            column(FilterText; FilterText)
            {
            }
            column(Itm_Description; Itm.Description)
            {
            }
            column(Itm_Category; Itm."Item Category Code")
            {
            }
            column(EntryNo_WareEntry; WareEntry."Entry No.")
            {
            }
            column(JournalBatchName_WareEntry; WareEntry."Journal Batch Name")
            {
            }
            column(LineNo_WareEntry; WareEntry."Line No.")
            {
            }
            column(RegisteringDate_WareEntry; FORMAT("Registering Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(LocationCode_WareEntry; WareEntry."Location Code")
            {
            }
            column(ZoneCode_WareEntry; WareEntry."Zone Code")
            {
            }
            column(BinCode_WareEntry; WareEntry."Bin Code")
            {
            }
            column(Description_WareEntry; WareEntry.Description)
            {
            }
            column(ItemNo_WareEntry; WareEntry."Item No.")
            {
            }
            column(Quantity_WareEntry; WareEntry.Quantity)
            {
            }
            column(QtyBase_WareEntry; WareEntry."Qty. (Base)")
            {
            }
            column(SourceType_WareEntry; WareEntry."Source Type")
            {
            }
            column(SourceSubtype_WareEntry; WareEntry."Source Subtype")
            {
            }
            column(SourceNo_WareEntry; WareEntry."Source No.")
            {
            }
            column(SourceLineNo_WareEntry; WareEntry."Source Line No.")
            {
            }
            column(SourceSublineNo_WareEntry; WareEntry."Source Subline No.")
            {
            }
            column(SourceDocument_WareEntry; WareEntry."Source Document")
            {
            }
            column(SourceCode_WareEntry; WareEntry."Source Code")
            {
            }
            column(ReasonCode_WareEntry; WareEntry."Reason Code")
            {
            }
            column(NoSeries_WareEntry; WareEntry."No. Series")
            {
            }
            column(BinTypeCode_WareEntry; WareEntry."Bin Type Code")
            {
            }
            column(Cubage_WareEntry; WareEntry.Cubage)
            {
            }
            column(Weight_WareEntry; WareEntry.Weight)
            {
            }
            column(JournalTemplateName_WareEntry; WareEntry."Journal Template Name")
            {
            }
            column(WhseDocumentNo_WareEntry; WareEntry."Whse. Document No.")
            {
            }
            column(WhseDocumentType_WareEntry; WareEntry."Whse. Document Type")
            {
            }
            column(WhseDocumentLineNo_WareEntry; WareEntry."Whse. Document Line No.")
            {
            }
            column(EntryType_WareEntry; WareEntry."Entry Type")
            {
            }
            column(ReferenceDocument_WareEntry; WareEntry."Reference Document")
            {
            }
            column(ReferenceNo_WareEntry; WareEntry."Reference No.")
            {
            }
            column(UserID_WareEntry; WareEntry."User ID")
            {
            }
            column(VariantCode_WareEntry; WareEntry."Variant Code")
            {
            }
            column(QtyperUnitofMeasure_WareEntry; WareEntry."Qty. per Unit of Measure")
            {
            }
            column(UnitofMeasureCode_WareEntry; WareEntry."Unit of Measure Code")
            {
            }
            column(SerialNo_WareEntry; WareEntry."Serial No.")
            {
            }
            column(LotNo_WareEntry; WareEntry."Lot No.")
            {
            }
            column(WarrantyDate_WareEntry; WareEntry."Warranty Date")
            {
            }
            column(ExpirationDate_WareEntry; WareEntry."Expiration Date")
            {
            }
            column(PhysInvtCountingPeriodCode_WareEntry; WareEntry."Phys Invt Counting Period Code")
            {
            }
            column(PhysInvtCountingPeriodType_WareEntry; WareEntry."Phys Invt Counting Period Type")
            {
            }
            column(Dedicated_WareEntry; WareEntry.Dedicated)
            {
            }

            trigger OnAfterGetRecord()
            begin

                CLEAR(Itm);
                IF Itm.GET("Item No.") THEN;

                CLEAR(ToBinCode);
                WE.RESET;
                WE.SETCURRENTKEY("Reference No.", "Registering Date");
                WE.SETRANGE("Reference No.", "Reference No.");
                //WE.SETRANGE("Registering Date","Registering Date");
                WE.SETRANGE("Location Code", "Location Code");
                WE.SETRANGE("Item No.", "Item No.");
                WE.SETRANGE(Quantity, (-1 * Quantity));
                IF WE.FINDFIRST THEN
                    ToBinCode := WE."Bin Code";

            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("Location Code", "Item No.", "Bin Code", "Registering Date");

                SETRANGE("Registering Date", SDate, EDate);
                IF Loc <> '' THEN
                    SETFILTER("Location Code", Loc);
                IF BinCode <> '' THEN
                    SETFILTER("Bin Code", BinCode);
                IF ItemCode <> '' THEN
                    SETFILTER("Item No.", ItemCode);

                FilterText := '';
                FilterText += 'Date filter :' + FORMAT(SDate, 0, '<Day,2>/<Month,2>/<Year4>') + ' .. ' + FORMAT(EDate, 0, '<Day,2>/<Month,2>/<Year,4>');
                IF Loc <> '' THEN
                    FilterText += '  , Location Filter : ' + Loc;
                IF BinCode <> '' THEN
                    FilterText += '  , BinCode Filter : ' + BinCode;
                IF ItemCode <> '' THEN
                    FilterText += '  , ItemNo Filter : ' + ItemCode + '.';
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

                    trigger OnValidate()
                    begin
                        IF EDate <> 0D THEN
                            IF SDate > EDate THEN
                                ERROR('EndDate %1 Must Be Greater Than StartDate %2', EDate, SDate);
                        EDate := SDate;
                    end;
                }
                field("End Date"; EDate)
                {

                    trigger OnValidate()
                    begin

                        IF EDate <> 0D THEN
                            IF SDate <> 0D THEN
                                IF SDate > EDate THEN
                                    ERROR('EndDate %1 Must Be Greater Than StartDate %2', EDate, SDate);
                    end;
                }
                field("Location Filter"; Loc)
                {
                    ApplicationArea = all;
                    TableRelation = Location WHERE("Bin Mandatory" = CONST(true));
                }
                field("BIN Code Filter"; BinCode)
                {
                    ApplicationArea = all;
                    TableRelation = Bin.Code;
                }
                field("Item No. Filter"; ItemCode)
                {
                    ApplicationArea = all;
                    TableRelation = Item;
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
            ERROR('Please Enter the Start Date');

        IF EDate = 0D THEN
            ERROR('Please Enter the End Date');

        IF Loc = '' THEN
            ERROR('Please select the Location');
    end;

    var
        SDate: Date;
        EDate: Date;
        Loc: Code[50];
        BinCode: Code[50];
        ItemCode: Code[50];
        Itm: Record 27;
        FilterText: Text;
        WE: Record 7312;
        ToBinCode: Code[50];
}


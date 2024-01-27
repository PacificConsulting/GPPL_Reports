report 50152 "BOM Explosion Prod. Planning"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BOMExplosionProdPlanning.rdl';
    Caption = 'BOM Explosion Prod. Planning';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Expected Output Quantity" = FILTER(<> 0));
            RequestFilterFields = "No.", "Search Description", "Inventory Posting Group";
            column(AsOfCalcDate; Text000 + FORMAT(CalculateDate))
            {
            }
            column(CompanyName; COMPANYNAME)
            {
            }
            column(TodayFormatted; FORMAT(TODAY, 0, 4))
            {
            }
            column(ItemTableCaptionFilter; TABLECAPTION + ': ' + ItemFilter)
            {
            }
            column(ItemFilter; ItemFilter)
            {
            }
            column(No_Item; "No.")
            {
            }
            column(Desc_Item; Description)
            {
            }
            column(QtyExplosionofBOMCapt; QtyExplosionofBOMCaptLbl)
            {
            }
            column(CurrReportPageNoCapt; CurrReportPageNoCaptLbl)
            {
            }
            column(BOMQtyCaption; BOMQtyCaptionLbl)
            {
            }
            column(BomCompLevelQtyCapt; BomCompLevelQtyCaptLbl)
            {
            }
            column(BomCompLevelDescCapt; BomCompLevelDescCaptLbl)
            {
            }
            column(BomCompLevelNoCapt; BomCompLevelNoCaptLbl)
            {
            }
            column(LevelCapt; LevelCaptLbl)
            {
            }
            column(BomCompLevelUOMCodeCapt; BomCompLevelUOMCodeCaptLbl)
            {
            }
            column(ExpectedOutputQtyLbl; ExpectedOutputQtyLbl)
            {
            }
            column(ExpectedOutputQuantity; Item."Expected Output Quantity")
            {
            }
            dataitem(BOMLoop; 2000000026)
            {
                DataItemTableView = SORTING(Number);
                dataitem(DataItem5444; 2000000026)
                {
                    DataItemTableView = SORTING(Number);
                    MaxIteration = 1;
                    column(BomCompLevelNo; BomComponent[Level]."No.")
                    {
                    }
                    column(BomCompLevelDesc; BomComponent[Level].Description)
                    {
                    }
                    column(BOMQty; BOMQty)
                    {
                        DecimalPlaces = 0 : 5;
                    }
                    column(FormatLevel; PADSTR('', Level, ' ') + FORMAT(Level))
                    {
                    }
                    column(BomCompLevelQty; BomComponent[Level].Quantity)
                    {
                        DecimalPlaces = 0 : 5;
                    }
                    column(BomCompLevelUOMCode; BomComponent[Level]."Unit of Measure Code")
                    {
                        //DecimalPlaces = 0 : 5;
                    }
                    column(ChildQuantity; ChildQuantity[Level])
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin
                        //BOMQty := Quantity[Level] * QtyPerUnitOfMeasure * BomComponent[Level].Quantity;//DJ Comment 29/01/2021
                        //DJ 29/01/2021
                        BOMQty := ChildQuantity[Level] * QtyPerUnitOfMeasure * Item."Expected Output Quantity";
                        //DJ 29/01/2021
                    end;

                    trigger OnPostDataItem()
                    begin
                        Level := NextLevel;
                    end;
                }

                trigger OnAfterGetRecord()
                var
                    BomItem: Record "27";
                begin
                    //DJ 28/01/2021
                    //IF Level <= 1 THEN
                    ChildQuantity[Level] := 0;
                    //DJ 28/01/2021
                    WHILE BomComponent[Level].NEXT = 0 DO BEGIN
                        Level := Level - 1;
                        IF Level < 1 THEN
                            CurrReport.BREAK;
                    END;

                    NextLevel := Level;
                    CLEAR(CompItem);
                    QtyPerUnitOfMeasure := 1;
                    CASE BomComponent[Level].Type OF
                        BomComponent[Level].Type::Item:
                            BEGIN
                                //DJ 28/01/2021
                                //IF Level > 1 THEN BEGIN
                                ChildLevel := Level;
                                //IF ChildLevel > 0 THEN BEGIN
                                //REPEAT
                                IF ChildLevel = 1 THEN
                                    ChildQuantity[Level] := BomComponent[Level].Quantity
                                ELSE
                                    ChildQuantity[Level] := BomComponent[Level].Quantity * ChildQuantity[Level - 1];
                                //ChildQuantity[Level] := ChildQuantity[Level] * ChildQuantity[Level-1];
                                // ChildLevel := ChildLevel-1;
                                //UNTIL ChildLevel = 0;
                                //ChildQuantity[Level] := ChildQuantity[Level] * BomComponent[1].Quantity;
                                //END;
                                //END;
                                /*
                                FOR I := 1 TO ChildLevel DO
                                  BEGIN
                                    IF I = Level THEN
                                      ChildQuantity[Level] := BomComponent[ChildLevel].Quantity * BomComponent[ChildLevel-1].Quantity
                                    ELSE
                                      ChildQuantity[ChildLevel] := ChildQuantity[Level] * BomComponent[ChildLevel-1].Quantity;
                                    I := ChildLevel-1;
                                  END;
                                  */
                                //DJ 28/01/2021
                                CompItem.GET(BomComponent[Level]."No.");
                                IF CompItem."Production BOM No." <> '' THEN BEGIN
                                    ProdBOM.GET(CompItem."Production BOM No.");
                                    IF ProdBOM.Status = ProdBOM.Status::Closed THEN
                                        CurrReport.SKIP;
                                    NextLevel := Level + 1;
                                    IF Level > 1 THEN
                                        IF (NextLevel > 50) OR (BomComponent[Level]."No." = NoList[Level - 1]) THEN
                                            ERROR(ProdBomErr, 50, Item."No.", NoList[Level], Level);
                                    CLEAR(BomComponent[NextLevel]);
                                    NoListType[NextLevel] := NoListType[NextLevel] ::Item;
                                    NoList[NextLevel] := CompItem."No.";
                                    VersionCode[NextLevel] :=
                                      VersionMgt.GetBOMVersion(CompItem."Production BOM No.", CalculateDate, TRUE);
                                    BomComponent[NextLevel].SETRANGE("Production BOM No.", CompItem."Production BOM No.");
                                    BomComponent[NextLevel].SETRANGE("Version Code", VersionCode[NextLevel]);
                                    BomComponent[NextLevel].SETFILTER("Starting Date", '%1|..%2', 0D, CalculateDate);
                                    BomComponent[NextLevel].SETFILTER("Ending Date", '%1|%2..', 0D, CalculateDate);
                                END;
                                IF Level > 1 THEN
                                    IF BomComponent[Level - 1].Type = BomComponent[Level - 1].Type::Item THEN
                                        IF BomItem.GET(BomComponent[Level - 1]."No.") THEN
                                            QtyPerUnitOfMeasure :=
                                              UOMMgt.GetQtyPerUnitOfMeasure(BomItem, BomComponent[Level - 1]."Unit of Measure Code") /
                                              UOMMgt.GetQtyPerUnitOfMeasure(
                                                BomItem, VersionMgt.GetBOMUnitOfMeasure(BomItem."Production BOM No.", VersionCode[Level]));
                            END;
                        BomComponent[Level].Type::"Production BOM":
                            BEGIN
                                //DJ 28/01/2021
                                IF ChildLevel = 1 THEN
                                    ChildQuantity[Level] := BomComponent[Level].Quantity
                                ELSE
                                    ChildQuantity[Level] := BomComponent[Level].Quantity * ChildQuantity[Level - 1];
                                //DJ 29/01/2021

                                ProdBOM.GET(BomComponent[Level]."No.");
                                IF ProdBOM.Status = ProdBOM.Status::Closed THEN
                                    CurrReport.SKIP;
                                NextLevel := Level + 1;
                                IF Level > 1 THEN
                                    IF (NextLevel > 50) OR (BomComponent[Level]."No." = NoList[Level - 1]) THEN
                                        ERROR(ProdBomErr, 50, Item."No.", NoList[Level], Level);
                                CLEAR(BomComponent[NextLevel]);
                                NoListType[NextLevel] := NoListType[NextLevel] ::"Production BOM";
                                NoList[NextLevel] := ProdBOM."No.";
                                VersionCode[NextLevel] := VersionMgt.GetBOMVersion(ProdBOM."No.", CalculateDate, TRUE);
                                BomComponent[NextLevel].SETRANGE("Production BOM No.", NoList[NextLevel]);
                                BomComponent[NextLevel].SETRANGE("Version Code", VersionCode[NextLevel]);
                                BomComponent[NextLevel].SETFILTER("Starting Date", '%1|..%2', 0D, CalculateDate);
                                BomComponent[NextLevel].SETFILTER("Ending Date", '%1|%2..', 0D, CalculateDate);
                            END;
                    END;

                    IF NextLevel <> Level THEN
                        Quantity[NextLevel] := BomComponent[NextLevel - 1].Quantity * QtyPerUnitOfMeasure * Quantity[Level];

                end;

                trigger OnPreDataItem()
                begin
                    Level := 1;

                    ProdBOM.GET(Item."Production BOM No.");

                    VersionCode[Level] := VersionMgt.GetBOMVersion(Item."Production BOM No.", CalculateDate, TRUE);
                    CLEAR(BomComponent);
                    BomComponent[Level]."Production BOM No." := Item."Production BOM No.";
                    BomComponent[Level].SETRANGE("Production BOM No.", Item."Production BOM No.");
                    BomComponent[Level].SETRANGE("Version Code", VersionCode[Level]);
                    BomComponent[Level].SETFILTER("Starting Date", '%1|..%2', 0D, CalculateDate);
                    BomComponent[Level].SETFILTER("Ending Date", '%1|%2..', 0D, CalculateDate);
                    NoListType[Level] := NoListType[Level] ::Item;
                    NoList[Level] := Item."No.";
                    Quantity[Level] :=
                      UOMMgt.GetQtyPerUnitOfMeasure(Item, Item."Base Unit of Measure") /
                      UOMMgt.GetQtyPerUnitOfMeasure(
                        Item,
                        VersionMgt.GetBOMUnitOfMeasure(
                          Item."Production BOM No.", VersionCode[Level]));
                end;
            }

            trigger OnPreDataItem()
            begin
                ItemFilter := GETFILTERS;

                SETFILTER("Production BOM No.", '<>%1', '');
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                group(Options)
                {
                    Caption = 'Options';
                    field(CalculateDate; CalculateDate)
                    {
                        ApplicationArea = all;
                        Caption = 'Calculation Date';
                    }
                }
            }
        }

        actions
        {
        }

        trigger OnInit()
        begin
            CalculateDate := WORKDATE;
        end;
    }

    labels
    {
    }

    trigger OnPostReport()
    begin
        ItemRec.RESET;
        ItemRec.SETCURRENTKEY("No.");
        ItemRec.SETFILTER("Expected Output Quantity", '<>%1', 0);
        ItemRec.MODIFYALL("Expected Output Quantity", 0, TRUE);
    end;

    var
        Text000: Label 'As of ';
        ProdBOM: Record 99000771;
        BomComponent: array[99] of Record 99000772;
        CompItem: Record 27;
        UOMMgt: Codeunit 5402;
        VersionMgt: Codeunit 99000756;
        ItemFilter: Text;
        CalculateDate: Date;
        NoList: array[99] of Code[20];
        VersionCode: array[99] of Code[20];
        Quantity: array[99] of Decimal;
        QtyPerUnitOfMeasure: Decimal;
        Level: Integer;
        NextLevel: Integer;
        BOMQty: Decimal;
        QtyExplosionofBOMCaptLbl: Label 'BOM Explosion Production Planning';
        CurrReportPageNoCaptLbl: Label 'Page';
        BOMQtyCaptionLbl: Label 'Total Quantity';
        BomCompLevelQtyCaptLbl: Label 'BOM Quantity';
        BomCompLevelDescCaptLbl: Label 'Description';
        BomCompLevelNoCaptLbl: Label 'No.';
        LevelCaptLbl: Label 'Level';
        BomCompLevelUOMCodeCaptLbl: Label 'Unit of Measure Code';
        NoListType: array[99] of Option " ",Item,"Production BOM";
        ProdBomErr: Label 'The maximum number of BOM levels, %1, was exceeded. The process stopped at item number %2, BOM header number %3, BOM level %4.';
        ChildQuantity: array[99] of Decimal;
        I: Integer;
        ChildLevel: Integer;
        ExpectedOutputQtyLbl: Label 'Expected Output Quantity';
        ItemRec: Record 27;
}


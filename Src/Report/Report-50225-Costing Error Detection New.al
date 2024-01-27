report 50225 "Costing Error Detection New"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CostingErrorDetectionNew.rdl';
    Caption = 'Costing Error Detection';
    PreviewMode = PrintLayout;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(ErrorGroup; Integer)
        {
            DataItemTableView = SORTING(Number);
            PrintOnlyIfDetail = true;
            column(CompanyName; COMPANYNAME)
            {
            }
            column(Today; FORMAT(TODAY, 0, 4))
            {
            }
            column(UserID; USERID)
            {
            }
            column(ShowCaption; ShowCaption)
            {
            }
            column(Number_ErrorGroup; Number)
            {
            }
            dataitem(Item; Item)
            {
                PrintOnlyIfDetail = true;
                RequestFilterFields = "No.";
                column(No_Item; "No.")
                {
                }
                column(Description_Item; Description)
                {
                    //  DecimalPlaces = 0 : 0;
                    IncludeCaption = true;
                }
                dataitem(ItemCheck; 27)
                {
                    DataItemLink = "No." = FIELD("No.");
                    DataItemTableView = SORTING("No.");
                    dataitem(ItemErrors; Integer)
                    {
                        DataItemTableView = SORTING(Number);
                        column(ErrorText_ItemErrors; ErrorText[Number])
                        {
                        }
                        column(Number_ItemErrors; Number)
                        {
                        }

                        trigger OnPreDataItem()
                        begin
                            SETRANGE(Number, 1, COMPRESSARRAY(ErrorText));
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin
                        ClearErrorText;

                        CASE ErrorGroupIndex OF
                            0:
                                CheckItem("No.");
                        END;

                        IF COMPRESSARRAY(ErrorText) = 0 THEN
                            CurrReport.SKIP;
                    end;
                }
                dataitem("Item Ledger Entry"; "Item Ledger Entry")
                {
                    DataItemLink = "Item No." = FIELD("No.");
                    DataItemTableView = SORTING("Item No.");
                    column(EntryNo_ItemLedgerEntry; "Entry No.")
                    {
                        IncludeCaption = true;
                    }
                    column(EntryType_ItemLedgerEntry; "Entry Type")
                    {
                        IncludeCaption = true;
                    }
                    column(EntryTypeFormat_ItemLedgerEntry; FORMAT("Entry Type"))
                    {
                    }
                    column(ItemNo_ItemLedgerEntry; "Item No.")
                    {
                        IncludeCaption = true;
                    }
                    column(Quantity_ItemLedgerEntry; Quantity)
                    {
                        IncludeCaption = true;
                    }
                    column(RemainingQuantity_ItemLedgerEntry; "Remaining Quantity")
                    {
                        IncludeCaption = true;
                    }
                    column(Positive_ItemLedgerEntry; Positive)
                    {
                        IncludeCaption = true;
                    }
                    column(Open_ItemLedgerEntry; Open)
                    {
                        IncludeCaption = true;
                    }
                    column(PostingDate_ItemLedgerEntry; "Posting Date")
                    {
                        IncludeCaption = true;
                    }
                    dataitem(Errors; Integer)
                    {
                        DataItemTableView = SORTING(Number);
                        column(ErrorText_Errors; ErrorText[Number])
                        {
                        }
                        column(Number_Errors; Number)
                        {
                        }

                        trigger OnPreDataItem()
                        begin
                            SETRANGE(Number, 1, COMPRESSARRAY(ErrorText))
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin
                        Window.UPDATE(3, "Entry No.");
                        ClearErrorText;
                        CASE ErrorGroupIndex OF
                            0:
                                CheckBasicData;
                            1:
                                CheckItemLedgEntryQty;
                            2:
                                CheckApplicationQty;
                            4:
                                CheckValuedByAverageCost;
                            5:
                                CheckValuationDate;
                            6:
                                CheckRemainingExpectedAmount;
                            7:
                                CheckOutputCompletelyInvdDate;

                        END;

                        IF COMPRESSARRAY(ErrorText) = 0 THEN
                            CurrReport.SKIP;
                    end;

                    trigger OnPreDataItem()
                    begin
                        CASE ErrorGroupIndex OF
                            0:
                                ;//Basic Data Test
                            1://Qty. Check Item Entry <--> Item Appl. Entry
                                BEGIN
                                    ItemApplEntry.RESET;
                                    ItemApplEntry.SETCURRENTKEY("Item Ledger Entry No.");
                                END;
                            2://Application Qty. Check
                                BEGIN
                                    SETRANGE(Positive, TRUE);
                                    ItemApplEntry.RESET;
                                    ItemApplEntry.SETCURRENTKEY("Inbound Item Entry No.");
                                END;
                            3://Check Related Inb. - Outb. Value Entry
                                BEGIN
                                    SETRANGE(Positive, TRUE);
                                    ItemApplEntry.RESET;
                                    ItemApplEntry.SETCURRENTKEY("Item Ledger Entry No.");
                                END;
                            4://Check Valued By Average Cost
                                BEGIN
                                    SETRANGE(Positive, FALSE);
                                    ItemApplEntry.RESET;
                                    ItemApplEntry.SETCURRENTKEY("Item Ledger Entry No.");
                                END;
                            5://Check Valuation Date
                                BEGIN
                                    SETRANGE(Positive, TRUE);
                                    ItemApplEntry.RESET;
                                    ItemApplEntry.SETCURRENTKEY("Inbound Item Entry No.");
                                END;
                            6:// Check remaining expected cost on closed item ledger entry
                                BEGIN
                                    ItemApplEntry.RESET;
                                    ItemApplEntry.SETCURRENTKEY("Outbound Item Entry No.", "Item Ledger Entry No.");
                                END;
                            7:// Check Output Completely Invd. Date' in Table 339
                                BEGIN

                                END;
                        END;
                    end;
                }
                dataitem("Value Entry"; "Value Entry")
                {
                    DataItemLink = "Item No." = FIELD("No.");
                    DataItemTableView = SORTING("Item No.", "Posting Date", "Item Ledger Entry Type", "Entry Type", "Variance Type", "Item Charge No.", "Location Code", "Variant Code");
                    column(EntryNo_ValueEntry; "Entry No.")
                    {
                    }
                    column(ItemNo_ValueEntry; "Item No.")
                    {
                    }
                    column(ItemLedgerEntryType_ValueEntry; "Item Ledger Entry Type")
                    {
                    }
                    column(PostingDate_ValueEntry; "Posting Date")
                    {
                    }
                    column(ItemLedgerEntryQuantity_ValueEntry; "Item Ledger Entry Quantity")
                    {
                    }

                    trigger OnAfterGetRecord()
                    var
                        ItemLedgEntry: Record "32";
                    begin
                        IF ItemLedgEntry.GET("Item Ledger Entry No.") THEN
                            CurrReport.SKIP;

                        IF Item2.GET("Item No.") THEN BEGIN
                            ItemTemp := Item2;
                            IF ItemTemp.INSERT THEN;
                        END;
                    end;

                    trigger OnPreDataItem()
                    begin
                        IF (ErrorGroupIndex <> 8) OR (Item."No." = '') THEN
                            CurrReport.BREAK;

                        IF NOT CheckItemLedgEntryExists THEN
                            CurrReport.BREAK;
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    IF Type = Type::Service THEN
                        CurrReport.SKIP;

                    Window.UPDATE(2, "No.");
                end;

                trigger OnPreDataItem()
                begin
                    IF ErrorGroupIndex IN [3, 4] THEN BEGIN
                        IF CostingMethodFiltered THEN
                            CurrReport.BREAK
                        ELSE
                            SETRANGE("Costing Method", "Costing Method"::Average);
                    END ELSE
                        SETRANGE("Costing Method");
                end;
            }

            trigger OnAfterGetRecord()
            begin
                ErrorGroupIndex := Number;
                CASE ErrorGroupIndex OF
                    0:
                        IF NOT BasicDataTest THEN
                            CurrReport.SKIP
                        ELSE
                            Window.UPDATE(1, Text023);
                    1:
                        IF NOT QtyCheckItemLedgEntry THEN
                            CurrReport.SKIP
                        ELSE
                            Window.UPDATE(1, Text002);
                    2:
                        IF NOT ApplicationQtyCheck THEN
                            CurrReport.SKIP
                        ELSE
                            Window.UPDATE(1, Text003);
                    4:
                        IF NOT ValuedByAverageCheck THEN
                            CurrReport.SKIP
                        ELSE
                            Window.UPDATE(1, Text006);
                    5:
                        IF NOT ValuationDateCheck THEN
                            CurrReport.SKIP
                        ELSE
                            Window.UPDATE(1, Text005);
                    6:
                        IF NOT RemExpectedOnClosedEntry THEN
                            CurrReport.SKIP
                        ELSE
                            Window.UPDATE(1, Text031);
                    7:
                        IF NOT OutputCompletelyInvd THEN
                            CurrReport.SKIP
                        ELSE
                            Window.UPDATE(1, Text033);
                    8:
                        IF NOT CheckItemLedgEntryExists THEN
                            CurrReport.SKIP
                        ELSE
                            Window.UPDATE(1, Text036);
                END;
            end;

            trigger OnPreDataItem()
            begin
                SETRANGE(Number, 0, 8);
            end;
        }
        dataitem(Summary; Integer)
        {
            DataItemTableView = SORTING(Number);
            column(ItemSummaryCaption; Text047)
            {
            }
            column(No_Summary; ItemTemp."No.")
            {
            }
            column(Description_Summary; ItemTemp.Description)
            {
            }
            column(Number_Summary; Number)
            {
            }

            trigger OnAfterGetRecord()
            begin
                IF Number = 1 THEN
                    ItemTemp.FINDFIRST
                ELSE
                    ItemTemp.NEXT;
            end;

            trigger OnPreDataItem()
            begin
                SETRANGE(Number, 1, ItemTemp.COUNT);
            end;
        }
    }

    requestpage
    {
        SaveValues = true;

        layout
        {
            area(content)
            {
                field(BasicDataTest; BasicDataTest)
                {
                    ApplicationArea = all;
                    Caption = 'Basic Data Test';
                }
                field(QtyCheckItemLedgEntry; QtyCheckItemLedgEntry)
                {
                    ApplicationArea = all;
                    Caption = 'Item Ledger Entry - Item Application Entry Quantity Check';
                }
                field(ApplicationQtyCheck; ApplicationQtyCheck)
                {
                    ApplicationArea = all;
                    Caption = 'Application Quantity Check';
                }
                field(ValuedByAverageCheck; ValuedByAverageCheck)
                {
                    ApplicationArea = all;
                    Caption = 'Check Valued By Average Cost';
                }
                field(ValuationDateCheck; ValuationDateCheck)
                {
                    ApplicationArea = all;
                    Caption = 'Check Valuation Date';
                }
                field(RemExpectedOnClosedEntry; RemExpectedOnClosedEntry)
                {
                    ApplicationArea = all;
                    Caption = 'Check Expected Cost on completely invoiced entry';
                }
                field(CheckItemLedgEntryExists; CheckItemLedgEntryExists)
                {
                    ApplicationArea = all;
                    Caption = 'Check Item ledg. Entry No. from Value Entries';
                }
            }
        }

        actions
        {
        }
    }

    labels
    {
        ReportName = 'Costing Error Detection';
        PageNoCaption = 'Page';
    }

    trigger OnPostReport()
    begin
        Window.CLOSE;
    end;

    trigger OnPreReport()
    var
        TempItem: Record 27 temporary;
    begin
        IF NOT (BasicDataTest OR QtyCheckItemLedgEntry OR ApplicationQtyCheck OR
                ValuedByAverageCheck OR ValuationDateCheck OR RemExpectedOnClosedEntry OR
                OutputCompletelyInvd OR CheckItemLedgEntryExists)
        THEN
            ERROR(Text000);

        IF NOT "Item Ledger Entry".FIND('-') THEN
            ERROR(Text001);

        Window.OPEN(Text020 + Text021 + Text022);

        TempItem.SETRANGE("Costing Method", TempItem."Costing Method"::Average);
        IF Item.GETFILTER("Costing Method") <> '' THEN
            IF Item.GETFILTER("Costing Method") <> TempItem.GETFILTER("Costing Method") THEN
                CostingMethodFiltered := TRUE;
    end;

    var
        ItemApplEntry: Record 339;
        Item2: Record 27;
        ItemTemp: Record 27 temporary;
        ErrorGroupIndex: Integer;
        ErrorIndex: Integer;
        BasicDataTest: Boolean;
        QtyCheckItemLedgEntry: Boolean;
        ApplicationQtyCheck: Boolean;
        ValuedByAverageCheck: Boolean;
        ValuationDateCheck: Boolean;
        Text000: Label 'You must select one or more of the data checks.';
        Text001: Label 'There are no Item Ledger Entries to check.';
        Text002: Label 'Item Ledger Entry - Item Appl. Entry Qty. Check';
        Text003: Label 'Application Quantity Check';
        Text005: Label 'Check Valuation Date';
        Text006: Label 'Check Valued By Average Cost';
        CostingMethodFiltered: Boolean;
        ErrorText: array[20] of Text[250];
        Text007: Label 'An %1 exists even though the %2 and %3 are the same in this negative Item Ledger Entry.';
        Text008: Label 'There is more than one %1 with the same combination of %2, %3 and %4.';
        Text009: Label 'The %1 is greater than the %1.';
        Text010: Label 'The summed %1 of the linked Item Application Entries is not the same as the difference between %2 and %3.';
        Text011: Label 'The summed %1 of the linked Item Application Entries is not the same as the %2.';
        Text012: Label 'The sign of the %1 in one or more of the linked Item Application Entries must be the opposite (positive must be negative or negative must be positive).';
        Text013: Label 'The summed %1 of the linked Item Application Entries is different than the %2 of the %3.';
        Text014: Label 'There are no Value Entries for this %1.';
        Text015: Label 'The summed %1 of the Item Application Entries for this transfer do not add up to zero as they should.';
        Text016: Label 'The Value Entries linked to the Item Application Entries do not all have the same %1 or %2.';
        Text017: Label 'The %1 in Value Entries applied to this %2 is earlier than the %1 in the %3 for this %2.';
        Text018: Label 'The value of the %1 field in the corresponding Value Entries is not correct.';
        Text019: Label 'The value of the %1 field in the corresponding Item Application Entries is not correct.';
        Window: Dialog;
        Text020: Label 'Function          #1##########################\';
        Text021: Label 'Item              #2##########################\';
        Text022: Label 'Item Ledger Entry #3##########################';
        Text023: Label 'Basic Data Test';
        Text024: Label 'The %1 must not be 0.';
        Text025: Label '%1 must be %2.';
        Text026: Label '%1 must be %2.';
        Text027: Label 'The linked Value Entries do not all have the same value of %1.';
        Text028: Label 'A linked %1 should not have the %2 %3 for a positive entry with %4 Consumption.';
        Text029: Label '%1 of Value Entries must not be Yes for any %2 other than Average.';
        Text030: Label 'There are no Item Application Entries for this %1.';
        RemExpectedOnClosedEntry: Boolean;
        Text031: Label 'Check Expected Cost on a closed entry';
        Text032: Label 'Expected Cost Amount was not 0 on an invoiced entry';
        OutputCompletelyInvd: Boolean;
        Text033: Label 'Check Output Completely Invoiced Date';
        Text034: Label 'At least one linked Item Application Entry has an unspecified %1.';
        Text035: Label 'A blank date in %1 was found on an Item Ledger Entry which has been invoiced';
        CheckItemLedgEntryExists: Boolean;
        Text036: Label 'Value Entries with missing Item Ledger Entry';
        Text037: Label 'Check Exptd. Cost on Completely Invoiced entries';
        Text038: Label 'Check Completely Invoiced date';
        Text039: Label '%1 was different to %2 on a linked Item Application Entry.';
        Text040: Label '%1 was specified on a linked Item Application Entry but the Item Ledger Entry is not Completely Invoiced.';
        Text041: Label '%1 was not equal to %2 on a linked Item Application Entry.';
        Text042: Label 'The summed %1 of the linked Value Entries is not the same as the %2.';
        Text043: Label '%1 or %2 was not 0 on a linked Value Entry. However this could be because report 1002 has not been run yet.';
        Text044: Label '%1 must be 0 when %2 is true.';
        Text045: Label '%1 or %2 is not 0 on a Completely Invoiced Item Ledger Entry';
        Text046: Label '%1 must equal %2 when %3 is true.';
        Text047: Label 'Item Summary';
        Text048: Label '%1 must be 0 on %2 %3 when %4 is true on a %5';
        Text049: Label '%1 must be false when %2 > 0 and %3 is %4 on %5 %6';
        Text050: Label '%1 on a %2 is not equal to %3 on %4 %5';
        Text051: Label '%1 on a %2 must not be smaller than or equal to 0.';
        Text052: Label '%1 must not be blank on a %2.';

    //  //[Scope('Internal')]
    procedure ShowCaption(): Text[50]
    begin
        //ShowCaption
        CASE ErrorGroupIndex OF
            0:
                EXIT(Text023);
            1:
                EXIT(Text002);
            2:
                EXIT(Text003);
            4:
                EXIT(Text006);
            5:
                EXIT(Text005);
            6:
                EXIT(Text037);
            7:
                EXIT(Text038);
            8:
                EXIT(Text036);
        END;
    end;

    //  //[Scope('Internal')]
    procedure ClearErrorText()
    begin
        //ClearErrorText
        IF COMPRESSARRAY(ErrorText) <> 0 THEN
            FOR ErrorIndex := 1 TO COMPRESSARRAY(ErrorText) DO
                ErrorText[ErrorIndex] := '';
        ErrorIndex := 1;
    end;

    //  //[Scope('Internal')]
    procedure CheckBasicData()
    begin
        //CheckBasicData
        BasicCheckItemLedgEntry;
        BasicCheckValueEntry;
    end;

    //  //[Scope('Internal')]
    procedure BasicCheckItemLedgEntry()
    var
        ValueEntry: Record 5802;
    begin
        //BasicCheckItemLedgEntry
        WITH "Item Ledger Entry" DO BEGIN
            IF "Entry No." <= 0 THEN
                AddError(STRSUBSTNO(Text051, FIELDCAPTION("Entry No."), TABLECAPTION), "Item No.");
            IF Quantity = 0 THEN BEGIN
                AddError(STRSUBSTNO(Text024, FIELDCAPTION(Quantity)), "Item No.");
            END ELSE BEGIN
                IF (Quantity * "Remaining Quantity") < 0 THEN BEGIN
                    AddError(STRSUBSTNO(Text009, FIELDCAPTION("Remaining Quantity"), FIELDCAPTION(Quantity)), "Item No.");
                END ELSE
                    IF ABS("Remaining Quantity") > ABS(Quantity) THEN BEGIN
                        AddError(STRSUBSTNO(Text009, FIELDCAPTION("Remaining Quantity"), FIELDCAPTION(Quantity)), "Item No.");
                    END;
                IF (Quantity > 0) <> Positive THEN BEGIN
                    AddError(STRSUBSTNO(Text025, FIELDCAPTION(Positive), NOT Positive), "Item No.");
                END;
            END;
            IF ("Remaining Quantity" = 0) = Open THEN BEGIN
                AddError(STRSUBSTNO(Text026, FIELDCAPTION(Open), NOT Open), "Item No.");
                ;
            END;

            IF "Completely Invoiced" THEN BEGIN
                IF "Invoiced Quantity" <> Quantity THEN BEGIN
                    AddError(
                      STRSUBSTNO(Text046, FIELDCAPTION("Invoiced Quantity"), FIELDCAPTION(Quantity), FIELDCAPTION("Completely Invoiced")),
                      "Item No.");
                END;

                ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
                ValueEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
                ValueEntry.CALCSUMS("Invoiced Quantity");
                IF "Invoiced Quantity" <> ValueEntry."Invoiced Quantity" THEN BEGIN
                    AddError(
                      STRSUBSTNO(Text042, ValueEntry.FIELDCAPTION("Invoiced Quantity"), FIELDCAPTION("Invoiced Quantity")),
                      "Item No.");
                END;
            END;
        END;
    end;

    //  //[Scope('Internal')]
    procedure BasicCheckValueEntry()
    var
        ValueEntry: Record 5802;
        ValuationDate: Date;
        ConsumptionDate: Date;
        ValuedByAverageCost: Boolean;
        Continue: Boolean;
        Compare: Boolean;
    begin
        //BasicCheckValueEntry
        WITH "Item Ledger Entry" DO BEGIN
            ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
            ValueEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
            ValueEntry.SETRANGE(Inventoriable, TRUE);
            IF NOT ValueEntry.FIND('-') THEN BEGIN
                AddError(STRSUBSTNO(Text014, TABLECAPTION), "Item No.");
            END ELSE BEGIN
                ValuedByAverageCost := ValueEntry."Valued By Average Cost";
                REPEAT
                    IF ValueEntry.Adjustment THEN BEGIN
                        IF ValueEntry."Invoiced Quantity" <> 0 THEN BEGIN
                            AddError(STRSUBSTNO(Text044, ValueEntry.FIELDCAPTION("Invoiced Quantity"), ValueEntry.FIELDCAPTION(Adjustment)),
                            "Item No.");
                            Continue := TRUE;
                        END;
                        IF ValueEntry."Item Ledger Entry Quantity" <> 0 THEN BEGIN
                            AddError(
                              STRSUBSTNO(Text048, "Item Ledger Entry".FIELDCAPTION(Quantity),
                                "Item Ledger Entry".TABLECAPTION, "Item Ledger Entry"."Entry No.",
                                ValueEntry.FIELDCAPTION(Adjustment), ValueEntry.TABLECAPTION),
                                "Item Ledger Entry"."Item No.");
                        END;
                    END;
                    IF ("Entry Type" = "Entry Type"::Consumption) AND Positive AND
                       (ValueEntry."Valuation Date" = DMY2DATE(31, 12, 9999))
                    THEN BEGIN
                        ConsumptionDate := DMY2DATE(31, 12, 9999);
                        AddError(
                          STRSUBSTNO(Text028, ValueEntry.TABLECAPTION, ValueEntry.FIELDCAPTION("Valuation Date"), ConsumptionDate,
                          FIELDCAPTION("Entry Type")),
                          "Item No.");
                        Continue := TRUE;
                    END ELSE BEGIN
                        IF (NOT Compare) AND (ValueEntry."Valuation Date" <> 0D) AND
                           NOT (ValueEntry."Entry Type" IN [ValueEntry."Entry Type"::Rounding, ValueEntry."Entry Type"::Revaluation])
                        THEN BEGIN
                            ValuationDate := ValueEntry."Valuation Date";
                            Compare := TRUE;
                        END;
                        IF Compare THEN
                            IF (ValueEntry."Valuation Date" <> ValuationDate) AND (ValueEntry."Valuation Date" <> 0D) AND
                               NOT (ValueEntry."Entry Type" IN [ValueEntry."Entry Type"::Rounding, ValueEntry."Entry Type"::Revaluation])
                            THEN BEGIN
                                AddError(STRSUBSTNO(Text027, ValueEntry.FIELDCAPTION("Valuation Date")), "Item No.");
                                Continue := TRUE;
                            END;
                    END;
                    IF (ValueEntry."Valued By Average Cost") AND
                       (Item."Costing Method" <> Item."Costing Method"::Average)
                    THEN BEGIN
                        AddError(STRSUBSTNO(Text029, ValueEntry.FIELDCAPTION("Valued By Average Cost"), Item.FIELDCAPTION("Costing Method")),
                        "Item No.");
                        Continue := TRUE;
                    END ELSE
                        IF ValueEntry."Valued By Average Cost" <> ValuedByAverageCost THEN BEGIN
                            AddError(STRSUBSTNO(Text027, ValueEntry.FIELDCAPTION("Valued By Average Cost")), "Item No.");
                            Continue := TRUE;
                        END;

                    IF (ValueEntry."Valued By Average Cost") AND
                       (Item."Costing Method" = Item."Costing Method"::Average) AND
                       (NOT Correction)
                     THEN
                        IF ValueEntry."Valued Quantity" > 0 THEN BEGIN
                            AddError(STRSUBSTNO(Text049,
                              ValueEntry.FIELDCAPTION("Valued By Average Cost"),
                              ValueEntry.FIELDCAPTION("Valued Quantity"),
                              Item.FIELDCAPTION("Costing Method"),
                              FORMAT(Item."Costing Method"),
                              ValueEntry.TABLECAPTION,
                              ValueEntry."Entry No."),
                              "Item No.");
                        END;

                    IF ValueEntry."Item Charge No." = '' THEN
                        IF ValueEntry."Item Ledger Entry Type" <> "Entry Type" THEN BEGIN
                            AddError(STRSUBSTNO(Text050,
                              ValueEntry.FIELDCAPTION("Item Ledger Entry Type"),
                              ValueEntry.TABLECAPTION,
                              FIELDCAPTION("Entry Type"),
                              "Item Ledger Entry".TABLECAPTION,
                              "Item Ledger Entry"."Entry No."),
                              "Item Ledger Entry"."Item No.");
                        END;
                UNTIL (ValueEntry.NEXT = 0) OR Continue;
            END;
        END;
    end;

    // //[Scope('Internal')]
    procedure CheckItemLedgEntryQty()
    begin
        //CheckItemLedgEntryQty
        WITH "Item Ledger Entry" DO BEGIN
            IF (NOT Positive) AND Open THEN
                CheckNegOpenILEQty
            ELSE
                CheckILEQty;
            SearchInbOutbCombination;
        END;
    end;

    // //[Scope('Internal')]
    procedure CheckNegOpenILEQty()
    var
        ApplQty: Decimal;
    begin
        //CheckNegOpenILEQty
        WITH "Item Ledger Entry" DO BEGIN
            IF Quantity = "Remaining Quantity" THEN BEGIN
                ItemApplEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
                IF ItemApplEntry.FIND('-') THEN BEGIN
                    AddError(STRSUBSTNO(Text007, ItemApplEntry.TABLECAPTION, FIELDCAPTION("Remaining Quantity"), FIELDCAPTION(Quantity)),
                      "Item No.");
                END;
            END ELSE BEGIN
                ItemApplEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
                IF ItemApplEntry.FIND('-') THEN BEGIN
                    ApplQty := 0;
                    REPEAT
                        ApplQty := ApplQty + ItemApplEntry.Quantity;
                    UNTIL ItemApplEntry.NEXT = 0;
                    IF ApplQty <> (Quantity - "Remaining Quantity") THEN BEGIN
                        AddError(
                          STRSUBSTNO(
                            Text010, ItemApplEntry.FIELDCAPTION(Quantity), FIELDCAPTION(Quantity), FIELDCAPTION("Remaining Quantity")),
                          "Item No.");
                    END;
                END;
            END;
        END;
    end;

    //  //[Scope('Internal')]
    procedure CheckILEQty()
    var
        ApplQty: Decimal;
    begin
        //CheckILEQty
        WITH "Item Ledger Entry" DO BEGIN
            ItemApplEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
            IF ItemApplEntry.FIND('-') THEN BEGIN
                ApplQty := 0;
                REPEAT
                    ApplQty := ApplQty + ItemApplEntry.Quantity;
                UNTIL ItemApplEntry.NEXT = 0;
                IF ApplQty <> Quantity THEN BEGIN
                    AddError(STRSUBSTNO(Text011, ItemApplEntry.FIELDCAPTION(Quantity), FIELDCAPTION(Quantity)), "Item No.");
                END;
            END ELSE BEGIN
                AddError(STRSUBSTNO(Text030, TABLECAPTION), "Item No.");
            END;
        END;
    end;

    // //[Scope('Internal')]
    procedure SearchInbOutbCombination()
    var
        ItemApplEntry2: Record 339;
        Continue: Boolean;
    begin
        //SearchInbOutbCombination
        WITH "Item Ledger Entry" DO BEGIN
            IF ItemApplEntry.FIND('-') THEN
                REPEAT
                    ItemApplEntry2.SETCURRENTKEY(
                      "Item Ledger Entry No.", "Inbound Item Entry No.", "Outbound Item Entry No.");
                    ItemApplEntry2.SETRANGE("Item Ledger Entry No.", ItemApplEntry."Item Ledger Entry No.");
                    ItemApplEntry2.SETRANGE("Inbound Item Entry No.", ItemApplEntry."Inbound Item Entry No.");
                    ItemApplEntry2.SETRANGE("Outbound Item Entry No.", ItemApplEntry."Outbound Item Entry No.");
                    ItemApplEntry2.SETFILTER("Entry No.", '<>%1', ItemApplEntry."Entry No.");
                    IF ItemApplEntry2.FIND('-') THEN BEGIN
                        AddError(
                          STRSUBSTNO(
                            Text008, ItemApplEntry2.TABLECAPTION, ItemApplEntry2.FIELDCAPTION("Item Ledger Entry No."),
                            ItemApplEntry2.FIELDCAPTION("Inbound Item Entry No."), ItemApplEntry2.FIELDCAPTION("Outbound Item Entry No.")),
                          "Item No.");
                        Continue := TRUE;
                    END;
                UNTIL (ItemApplEntry.NEXT = 0) OR Continue;
        END;
    end;

    // //[Scope('Internal')]
    procedure CheckApplicationQty()
    var
        ApplQty: Decimal;
        Continue: Boolean;
    begin
        //CheckApplicationQty
        WITH "Item Ledger Entry" DO BEGIN
            ItemApplEntry.SETRANGE("Inbound Item Entry No.", "Entry No.");
            IF ItemApplEntry.FIND('-') THEN BEGIN
                REPEAT
                    IF ((ItemApplEntry."Item Ledger Entry No." = ItemApplEntry."Inbound Item Entry No.") AND
                        (ItemApplEntry.Quantity < 0)) OR
                       ((ItemApplEntry."Item Ledger Entry No." <> ItemApplEntry."Inbound Item Entry No.") AND
                        (ItemApplEntry.Quantity > 0)) OR
                       ((ItemApplEntry."Item Ledger Entry No." <> ItemApplEntry."Outbound Item Entry No.") AND
                        (ItemApplEntry.Quantity < 0))
                    THEN BEGIN
                        AddError(STRSUBSTNO(Text012, ItemApplEntry.FIELDCAPTION(Quantity)), "Item No.");
                        Continue := TRUE;
                    END;
                UNTIL Continue OR (ItemApplEntry.NEXT = 0);
                ItemApplEntry.FIND('-');
                ApplQty := 0;
                REPEAT
                    ApplQty := ApplQty + ItemApplEntry.Quantity;
                UNTIL ItemApplEntry.NEXT = 0;
                IF ApplQty <> "Remaining Quantity" THEN BEGIN
                    AddError(
                      STRSUBSTNO(Text013, ItemApplEntry.FIELDCAPTION(Quantity), FIELDCAPTION("Remaining Quantity"), TABLECAPTION),
                      "Item No.");
                END;
            END;
        END;
    end;

    //  //[Scope('Internal')]
    procedure CheckValuedByAverageCost()
    begin
        //CheckValuedByAverageCost
        CheckVEValuedBySetting;
        CheckItemApplCostApplSetting;
    end;

    // //[Scope('Internal')]
    procedure CheckVEValuedBySetting()
    var
        ValueEntry: Record 5802;
        ValueEntry2: Record 5802;
        Continue: Boolean;
    begin
        //CheckVEValuedBySetting
        WITH "Item Ledger Entry" DO BEGIN
            ValueEntry2.SETCURRENTKEY("Item Ledger Entry No.");
            ValueEntry2.SETRANGE(Inventoriable, TRUE);
            ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
            ValueEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
            ValueEntry.SETRANGE(Inventoriable, TRUE);
            IF ValueEntry.FIND('-') THEN
                REPEAT
                    IF ValueEntry."Item Ledger Entry Type" <> ValueEntry."Item Ledger Entry Type"::Output THEN
                        IF "Applies-to Entry" <> 0 THEN BEGIN
                            ValueEntry2.SETRANGE("Item Ledger Entry No.", "Applies-to Entry");
                            IF ValueEntry2.FIND('-') THEN
                                IF ValueEntry."Valued By Average Cost" <> ValueEntry2."Valued By Average Cost" THEN BEGIN
                                    AddError(STRSUBSTNO(Text018, ValueEntry.FIELDCAPTION("Valued By Average Cost")), "Item No.");
                                    Continue := TRUE;
                                END;
                        END ELSE
                            IF (NOT ValueEntry."Valued By Average Cost") AND
                               (ValueEntry."Valuation Date" <> 0D)
                               //NAV2013+.begin
                               // Starting with Dynamics NAV 2013, an outbound Transfer with <blank> Applies-to Entry field
                               // is an accurate combination according to new design.
                               AND
                               ("Entry Type" <> "Entry Type"::Transfer) AND
                               (Quantity < 0)
                            //NAV2013+.end
                            THEN BEGIN
                                AddError(STRSUBSTNO(Text018, ValueEntry.FIELDCAPTION("Valued By Average Cost")), "Item No.");
                                Continue := TRUE;
                            END;
                UNTIL (ValueEntry.NEXT = 0) OR Continue;
        END;
    end;

    // //[Scope('Internal')]
    procedure CheckItemApplCostApplSetting()
    var
        Continue: Boolean;
    begin
        //CheckItemApplCostApplSetting
        WITH "Item Ledger Entry" DO BEGIN
            ItemApplEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
            IF ItemApplEntry.FIND('-') THEN
                REPEAT
                    IF ItemApplEntry.Quantity > 0 THEN BEGIN
                        IF NOT ItemApplEntry."Cost Application" THEN BEGIN
                            AddError(STRSUBSTNO(Text019, ItemApplEntry.FIELDCAPTION("Cost Application")), "Item No.");
                            Continue := TRUE;
                        END;
                    END ELSE
                        IF "Applies-to Entry" <> 0 THEN BEGIN
                            IF NOT ItemApplEntry."Cost Application" THEN BEGIN
                                AddError(STRSUBSTNO(Text019, ItemApplEntry.FIELDCAPTION("Cost Application")), "Item No.");
                                Continue := TRUE;
                            END;
                        END ELSE
                            //NAV2013+.begin
                            // Starting with Dynamics NAV 2013, an outbound Transfer with Cost Application set to TRUE
                            // is an accurate combination according to new design.
                            //IF ItemApplEntry."Cost Application" THEN BEGIN
                            IF (ItemApplEntry."Cost Application") AND ("Entry Type" <> "Entry Type"::Transfer) THEN BEGIN
                                //NAV2013+.end
                                AddError(STRSUBSTNO(Text019, ItemApplEntry.FIELDCAPTION("Cost Application")), "Item No.");
                                Continue := TRUE;
                            END;
                UNTIL (ItemApplEntry.NEXT = 0) OR Continue;
        END;
    end;

    // //[Scope('Internal')]
    procedure CheckValuationDate()
    var
        ValueEntry: Record 5802;
        ValuationDate: Date;
        Continue: Boolean;
    begin
        //CheckValuationDate
        WITH "Item Ledger Entry" DO BEGIN
            ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
            ValueEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
            ValueEntry.SETRANGE(Inventoriable, TRUE);
            ValueEntry.SETRANGE("Partial Revaluation", FALSE);
            IF ValueEntry.FIND('-') THEN BEGIN
                ValuationDate := ValueEntry."Valuation Date";
                ItemApplEntry.SETRANGE("Inbound Item Entry No.", "Entry No.");
                IF ItemApplEntry.FIND('-') THEN
                    REPEAT
                        IF ItemApplEntry.Quantity < 0 THEN BEGIN
                            ValueEntry.SETRANGE("Item Ledger Entry No.", ItemApplEntry."Item Ledger Entry No.");
                            IF ValueEntry.FIND('-') THEN
                                REPEAT
                                    IF ValueEntry."Valuation Date" < ValuationDate THEN BEGIN
                                        AddError(STRSUBSTNO(Text017, ValueEntry.FIELDCAPTION("Valuation Date"), TABLECAPTION, ValueEntry.TABLECAPTION),
                                          "Item No.");
                                        Continue := TRUE;
                                    END;
                                UNTIL (ValueEntry.NEXT = 0) OR Continue;
                        END;
                    UNTIL (ItemApplEntry.NEXT = 0) OR Continue;
            END;
        END;
    end;

    //   //[Scope('Internal')]
    procedure CheckRemainingExpectedAmount()
    var
        ValueEntry: Record "5802";
        TotalExpectedCostToGL: Decimal;
        TotalExpectedCostToGLACY: Decimal;
    begin
        IF NOT "Item Ledger Entry"."Completely Invoiced" THEN
            EXIT;

        ValueEntry.SETCURRENTKEY("Item Ledger Entry No.");
        ValueEntry.SETRANGE("Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");

        "Item Ledger Entry".CALCFIELDS("Cost Amount (Expected)", "Cost Amount (Expected) (ACY)");
        IF
          ("Item Ledger Entry"."Cost Amount (Expected)" = 0) AND
          ("Item Ledger Entry"."Cost Amount (Expected) (ACY)" = 0)
        THEN BEGIN
            IF ValueEntry.FINDSET THEN
                REPEAT
                    TotalExpectedCostToGL := TotalExpectedCostToGL + ValueEntry."Expected Cost Posted to G/L";
                    TotalExpectedCostToGLACY := TotalExpectedCostToGLACY + ValueEntry."Exp. Cost Posted to G/L (ACY)";
                UNTIL ValueEntry.NEXT = 0;
            IF (TotalExpectedCostToGL = 0) AND (TotalExpectedCostToGLACY = 0) THEN
                EXIT;
        END;

        IF ValueEntry.FINDSET THEN
            REPEAT
                IF (ValueEntry."Expected Cost Posted to G/L" <> 0) OR (ValueEntry."Exp. Cost Posted to G/L (ACY)" <> 0) THEN BEGIN
                    AddError(
                      STRSUBSTNO(Text043,
                        ValueEntry.FIELDCAPTION("Expected Cost Posted to G/L"), ValueEntry.FIELDCAPTION("Exp. Cost Posted to G/L (ACY)")),
                      ValueEntry."Item No.");

                END;
            UNTIL ValueEntry.NEXT = 0;

        IF ErrorIndex > 0 THEN BEGIN
            AddError(
              STRSUBSTNO(
                Text045,
                "Item Ledger Entry".FIELDCAPTION("Cost Amount (Expected)"), "Item Ledger Entry".FIELDCAPTION("Cost Amount (Expected) (ACY)")),
              "Item Ledger Entry"."Item No.");
        END;
    end;

    // //[Scope('Internal')]
    procedure CheckOutputCompletelyInvdDate()
    var
        ItemApplicationEntry: Record 339;
        ZeroDateFound: Boolean;
    begin
        IF "Item Ledger Entry"."Invoiced Quantity" <> 0 THEN
            IF "Item Ledger Entry"."Last Invoice Date" = 0D THEN BEGIN
                AddError(STRSUBSTNO(Text035, "Item Ledger Entry".FIELDCAPTION("Last Invoice Date")),
                  "Item Ledger Entry"."Item No.");
            END;

        ItemApplicationEntry.SETCURRENTKEY("Item Ledger Entry No.", "Output Completely Invd. Date");
        ItemApplicationEntry.SETRANGE("Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
        IF ItemApplicationEntry.FINDSET THEN
            REPEAT
                IF ItemApplicationEntry.Quantity > 0 THEN BEGIN
                    IF ItemApplicationEntry."Output Completely Invd. Date" <> ItemApplicationEntry."Posting Date" THEN BEGIN
                        ZeroDateFound := TRUE;
                        AddError(
                          STRSUBSTNO(Text039,
                            ItemApplicationEntry.FIELDCAPTION("Posting Date"), ItemApplicationEntry.FIELDCAPTION("Output Completely Invd. Date")),
                          "Item Ledger Entry"."Item No.");
                    END;
                END ELSE BEGIN
                    IF "Item Ledger Entry".Quantity = "Item Ledger Entry"."Invoiced Quantity" THEN BEGIN
                        IF ItemApplicationEntry."Output Completely Invd. Date" <> "Item Ledger Entry"."Last Invoice Date" THEN BEGIN
                            ZeroDateFound := TRUE;
                            AddError(
                              STRSUBSTNO(Text041,
                                ItemApplicationEntry.FIELDCAPTION("Output Completely Invd. Date"),
                                "Item Ledger Entry".FIELDCAPTION("Last Invoice Date")),
                              "Item Ledger Entry"."Item No.");
                        END;
                    END ELSE
                        IF ItemApplicationEntry."Output Completely Invd. Date" <> 0D THEN BEGIN
                            ZeroDateFound := TRUE;
                            AddError(
                              STRSUBSTNO(Text040, ItemApplicationEntry.FIELDCAPTION("Output Completely Invd. Date")),
                              "Item Ledger Entry"."Item No.");

                        END;
                END;
            UNTIL ZeroDateFound OR (ItemApplicationEntry.NEXT = 0);
    end;

    // //[Scope('Internal')]
    procedure CheckItem(ItemNo: Code[20])
    begin
        IF ItemNo = '' THEN
            AddError(STRSUBSTNO(Text052, Item.FIELDCAPTION("No."), Item.TABLECAPTION), ItemNo);
    end;

    // //[Scope('Internal')]
    procedure AddError(ErrorMessage: Text[250]; ItemNo: Code[20])
    begin
        ErrorText[ErrorIndex] := ErrorMessage;
        ErrorIndex := ErrorIndex + 1;

        IF Item2.GET(ItemNo) THEN BEGIN
            ItemTemp := Item2;
            IF ItemTemp.INSERT THEN;
        END;
    end;
}


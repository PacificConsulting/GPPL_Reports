report 50212 "Inventory Valuation - New"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 13Jul2018   RB-N         New Report Development from R#1001
    // 20Jul2018   RB-N         Skipping Closed Location Inventory
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/InventoryValuationNew.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Inventory Valuation - New';
    EnableHyperlinks = true;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("Inventory Posting Group")
                                WHERE(Type = CONST(Inventory));
            RequestFilterFields = "No.", "Inventory Posting Group", "Statistics Group";
            column(StartDate; StartDate)
            {
            }
            column(EndDate; EndDate)
            {
            }
            column(BoM_Text; BoM_TextLbl)
            {
            }
            column(COMPANYNAME; COMPANYNAME)
            {
            }
            column(CurrReport_PAGENO; CurrReport.PAGENO)
            {
            }
            column(STRSUBSTNO___1___2__Item_TABLECAPTION_ItemFilter_; STRSUBSTNO('%1: %2', TABLECAPTION, ItemFilter))
            {
            }
            column(STRSUBSTNO_Text005_StartDateText_; STRSUBSTNO(Text005, StartDateText))
            {
            }
            column(STRSUBSTNO_Text005_FORMAT_EndDate__; STRSUBSTNO(Text005, FORMAT(EndDate)))
            {
            }
            column(ShowExpected; ShowExpected)
            {
            }
            column(ItemFilter; ItemFilter)
            {
            }
            column(Inventory_ValuationCaption; Inventory_ValuationCaptionLbl)
            {
            }
            column(CurrReport_PAGENOCaption; CurrReport_PAGENOCaptionLbl)
            {
            }
            column(This_report_includes_entries_that_have_been_posted_with_expected_costs_Caption; This_report_includes_entries_that_have_been_posted_with_expected_costs_CaptionLbl)
            {
            }
            column(ItemNoCaption; ValueEntry.FIELDCAPTION("Item No."))
            {
            }
            column(ItemDescriptionCaption; FIELDCAPTION(Description))
            {
            }
            column(IncreaseInvoicedQtyCaption; IncreaseInvoicedQtyCaptionLbl)
            {
            }
            column(DecreaseInvoicedQtyCaption; DecreaseInvoicedQtyCaptionLbl)
            {
            }
            column(QuantityCaption; QuantityCaptionLbl)
            {
            }
            column(ValueCaption; ValueCaptionLbl)
            {
            }
            column(QuantityCaption_Control31; QuantityCaption_Control31Lbl)
            {
            }
            column(QuantityCaption_Control40; QuantityCaption_Control40Lbl)
            {
            }
            column(InvCostPostedToGL_Control53Caption; InvCostPostedToGL_Control53CaptionLbl)
            {
            }
            column(QuantityCaption_Control58; QuantityCaption_Control58Lbl)
            {
            }
            column(TotalCaption; TotalCaptionLbl)
            {
            }
            column(Expected_Cost_IncludedCaption; Expected_Cost_IncludedCaptionLbl)
            {
            }
            column(Expected_Cost_Included_TotalCaption; Expected_Cost_Included_TotalCaptionLbl)
            {
            }
            column(Expected_Cost_TotalCaption; Expected_Cost_TotalCaptionLbl)
            {
            }
            column(GetUrlForReportDrilldown; GetUrlForReportDrilldown("No."))
            {
            }
            column(ItemNo; "No.")
            {
            }
            column(ItemDescription; Description)
            {
            }
            column(ItemBaseUnitofMeasure; "Base Unit of Measure")
            {
            }
            column(Item_Inventory_Posting_Group; "Inventory Posting Group")
            {
            }
            column(StartingInvoicedValue; StartingInvoicedValue)
            {
                AutoFormatType = 1;
            }
            column(StartingInvoicedQty; StartingInvoicedQty)
            {
                DecimalPlaces = 0 : 5;
            }
            column(StartingExpectedValue; StartingExpectedValue)
            {
                AutoFormatType = 1;
            }
            column(StartingExpectedQty; StartingExpectedQty)
            {
                DecimalPlaces = 0 : 5;
            }
            column(IncreaseInvoicedValue; IncreaseInvoicedValue)
            {
                AutoFormatType = 1;
            }
            column(IncreaseInvoicedQty; IncreaseInvoicedQty)
            {
                DecimalPlaces = 0 : 5;
            }
            column(IncreaseExpectedValue; IncreaseExpectedValue)
            {
                AutoFormatType = 1;
            }
            column(IncreaseExpectedQty; IncreaseExpectedQty)
            {
                DecimalPlaces = 0 : 5;
            }
            column(DecreaseInvoicedValue; DecreaseInvoicedValue)
            {
                AutoFormatType = 1;
            }
            column(DecreaseInvoicedQty; DecreaseInvoicedQty)
            {
                DecimalPlaces = 0 : 5;
            }
            column(DecreaseExpectedValue; DecreaseExpectedValue)
            {
                AutoFormatType = 1;
            }
            column(DecreaseExpectedQty; DecreaseExpectedQty)
            {
                DecimalPlaces = 0 : 5;
            }
            column(EndingInvoicedValue; StartingInvoicedValue + IncreaseInvoicedValue - DecreaseInvoicedValue)
            {
            }
            column(EndingInvoicedQty; StartingInvoicedQty + IncreaseInvoicedQty - DecreaseInvoicedQty)
            {
            }
            column(EndingExpectedValue; StartingExpectedValue + IncreaseExpectedValue - DecreaseExpectedValue)
            {
                AutoFormatType = 1;
            }
            column(EndingExpectedQty; StartingExpectedQty + IncreaseExpectedQty - DecreaseExpectedQty)
            {
            }
            column(CostPostedToGL; CostPostedToGL)
            {
                AutoFormatType = 1;
            }
            column(InvCostPostedToGL; InvCostPostedToGL)
            {
                AutoFormatType = 1;
            }
            column(ExpCostPostedToGL; ExpCostPostedToGL)
            {
                AutoFormatType = 1;
            }

            trigger OnAfterGetRecord()
            begin
                CALCFIELDS("Assembly BOM");

                IF EndDate = 0D THEN
                    EndDate := 99991231D/*31129999D*/;

                StartingInvoicedValue := 0;
                StartingExpectedValue := 0;
                StartingInvoicedQty := 0;
                StartingExpectedQty := 0;
                IncreaseInvoicedValue := 0;
                IncreaseExpectedValue := 0;
                IncreaseInvoicedQty := 0;
                IncreaseExpectedQty := 0;
                DecreaseInvoicedValue := 0;
                DecreaseExpectedValue := 0;
                DecreaseInvoicedQty := 0;
                DecreaseExpectedQty := 0;
                InvCostPostedToGL := 0;
                CostPostedToGL := 0;
                ExpCostPostedToGL := 0;

                IsEmptyLine := TRUE;

                //>>1 Condition
                IF NOT IncValuation THEN BEGIN
                    ValueEntry.RESET;
                    ValueEntry.SETRANGE("Item No.", "No.");
                    ValueEntry.SETFILTER("Variant Code", GETFILTER("Variant Filter"));
                    ValueEntry.SETFILTER("Location Code", GETFILTER("Location Filter"));
                    ValueEntry.SETFILTER("Global Dimension 1 Code", GETFILTER("Global Dimension 1 Filter"));
                    ValueEntry.SETFILTER("Global Dimension 2 Code", GETFILTER("Global Dimension 2 Filter"));

                    IF StartDate > 0D THEN BEGIN
                        ValueEntry.SETRANGE("Posting Date", 0D, CALCDATE('<-1D>', StartDate));
                        ValueEntry.CALCSUMS("Item Ledger Entry Quantity", "Cost Amount (Actual)", "Cost Amount (Expected)", "Invoiced Quantity");
                        AssignAmounts(ValueEntry, StartingInvoicedValue, StartingInvoicedQty, StartingExpectedValue, StartingExpectedQty, 1);
                        IsEmptyLine := IsEmptyLine AND ((StartingInvoicedValue = 0) AND (StartingInvoicedQty = 0));
                        IF ShowExpected THEN
                            IsEmptyLine := IsEmptyLine AND ((StartingExpectedValue = 0) AND (StartingExpectedQty = 0));
                    END;
                END ELSE BEGIN

                    Loc20.RESET;
                    Loc20.SETRANGE(Closed, FALSE);
                    Loc20.SETRANGE("Include in Valuation Report", TRUE);
                    IF Loc20.FINDSET THEN
                        REPEAT
                            VE20.RESET;
                            VE20.SETCURRENTKEY("Item No.");
                            VE20.SETRANGE("Item No.", "No.");
                            VE20.SETFILTER("Variant Code", GETFILTER("Variant Filter"));
                            VE20.SETFILTER("Location Code", Loc20.Code);
                            VE20.SETFILTER("Global Dimension 1 Code", GETFILTER("Global Dimension 1 Filter"));
                            VE20.SETFILTER("Global Dimension 2 Code", GETFILTER("Global Dimension 2 Filter"));

                            IF StartDate > 0D THEN BEGIN
                                VE20.SETRANGE("Posting Date", 0D, CALCDATE('<-1D>', StartDate));
                                VE20.CALCSUMS("Item Ledger Entry Quantity", "Cost Amount (Actual)", "Cost Amount (Expected)", "Invoiced Quantity");
                                AssignAmounts(VE20, StartingInvoicedValue, StartingInvoicedQty, StartingExpectedValue, StartingExpectedQty, 1);
                                IsEmptyLine := IsEmptyLine AND ((StartingInvoicedValue = 0) AND (StartingInvoicedQty = 0));
                                IF ShowExpected THEN
                                    IsEmptyLine := IsEmptyLine AND ((StartingExpectedValue = 0) AND (StartingExpectedQty = 0));
                            END;
                        UNTIL Loc20.NEXT = 0;
                END;
                //<<1 Condition


                //>>2 Condition
                IF NOT IncValuation THEN BEGIN
                    ValueEntry.SETRANGE("Posting Date", StartDate, EndDate);
                    ValueEntry.SETFILTER(
                      "Item Ledger Entry Type", '%1|%2|%3|%4',
                      ValueEntry."Item Ledger Entry Type"::Purchase,
                      ValueEntry."Item Ledger Entry Type"::"Positive Adjmt.",
                      ValueEntry."Item Ledger Entry Type"::Output,
                      ValueEntry."Item Ledger Entry Type"::"Assembly Output");
                    ValueEntry.CALCSUMS("Item Ledger Entry Quantity", "Cost Amount (Actual)", "Cost Amount (Expected)", "Invoiced Quantity");
                    AssignAmounts(ValueEntry, IncreaseInvoicedValue, IncreaseInvoicedQty, IncreaseExpectedValue, IncreaseExpectedQty, 1);
                END ELSE BEGIN
                    Loc20.RESET;
                    Loc20.SETRANGE(Closed, FALSE);
                    Loc20.SETRANGE("Include in Valuation Report", TRUE);
                    IF Loc20.FINDSET THEN
                        REPEAT
                            VE20.RESET;
                            VE20.SETCURRENTKEY("Item No.");
                            VE20.SETRANGE("Item No.", "No.");
                            IF GETFILTER("Variant Filter") <> '' THEN
                                VE20.SETFILTER("Variant Code", GETFILTER("Variant Filter"));
                            VE20.SETFILTER("Location Code", Loc20.Code);
                            IF GETFILTER("Global Dimension 2 Filter") <> '' THEN
                                VE20.SETFILTER("Global Dimension 1 Code", GETFILTER("Global Dimension 1 Filter"));
                            IF GETFILTER("Global Dimension 2 Filter") <> '' THEN
                                VE20.SETFILTER("Global Dimension 2 Code", GETFILTER("Global Dimension 2 Filter"));
                            VE20.SETRANGE("Posting Date", StartDate, EndDate);
                            VE20.SETFILTER("Item Ledger Entry Type", '%1|%2|%3|%4',
                                VE20."Item Ledger Entry Type"::Purchase,
                                VE20."Item Ledger Entry Type"::"Positive Adjmt.",
                                VE20."Item Ledger Entry Type"::Output,
                                VE20."Item Ledger Entry Type"::"Assembly Output");
                            IF VE20.FINDFIRST THEN BEGIN
                                VE20.CALCSUMS("Item Ledger Entry Quantity", "Cost Amount (Actual)", "Cost Amount (Expected)", "Invoiced Quantity");
                                AssignAmounts(VE20, IncreaseInvoicedValue, IncreaseInvoicedQty, IncreaseExpectedValue, IncreaseExpectedQty, 1);
                            END;
                        UNTIL Loc20.NEXT = 0;
                END;
                //MESSAGE('%1 ttt',IncreaseInvoicedValue);

                //<<2 Condition


                //>>3 Condition
                IF NOT IncValuation THEN BEGIN
                    ValueEntry.SETRANGE("Posting Date", StartDate, EndDate);
                    ValueEntry.SETFILTER(
                      "Item Ledger Entry Type", '%1|%2|%3|%4',
                      ValueEntry."Item Ledger Entry Type"::Sale,
                      ValueEntry."Item Ledger Entry Type"::"Negative Adjmt.",
                      ValueEntry."Item Ledger Entry Type"::Consumption,
                      ValueEntry."Item Ledger Entry Type"::"Assembly Consumption");
                    ValueEntry.CALCSUMS("Item Ledger Entry Quantity", "Cost Amount (Actual)", "Cost Amount (Expected)", "Invoiced Quantity");
                    AssignAmounts(ValueEntry, DecreaseInvoicedValue, DecreaseInvoicedQty, DecreaseExpectedValue, DecreaseExpectedQty, -1);
                END ELSE BEGIN
                    Loc20.RESET;
                    Loc20.SETRANGE(Closed, FALSE);
                    Loc20.SETRANGE("Include in Valuation Report", TRUE);
                    IF Loc20.FINDSET THEN
                        REPEAT
                            VE20.RESET;
                            VE20.SETCURRENTKEY("Item No.");
                            VE20.SETRANGE("Item No.", "No.");
                            IF GETFILTER("Variant Filter") <> '' THEN
                                VE20.SETFILTER("Variant Code", GETFILTER("Variant Filter"));
                            VE20.SETFILTER("Location Code", Loc20.Code);
                            IF GETFILTER("Global Dimension 1 Filter") <> '' THEN
                                VE20.SETFILTER("Global Dimension 1 Code", GETFILTER("Global Dimension 1 Filter"));
                            IF GETFILTER("Global Dimension 2 Filter") <> '' THEN
                                VE20.SETFILTER("Global Dimension 2 Code", GETFILTER("Global Dimension 2 Filter"));
                            VE20.SETRANGE("Posting Date", StartDate, EndDate);
                            VE20.SETFILTER("Item Ledger Entry Type", '%1|%2|%3|%4',
                              VE20."Item Ledger Entry Type"::Sale,
                              VE20."Item Ledger Entry Type"::"Negative Adjmt.",
                              VE20."Item Ledger Entry Type"::Consumption,
                              VE20."Item Ledger Entry Type"::"Assembly Consumption");
                            VE20.CALCSUMS("Item Ledger Entry Quantity", "Cost Amount (Actual)", "Cost Amount (Expected)", "Invoiced Quantity");
                            AssignAmounts(VE20, DecreaseInvoicedValue, DecreaseInvoicedQty, DecreaseExpectedValue, DecreaseExpectedQty, -1);
                        UNTIL Loc20.NEXT = 0;

                END;
                //<<3 Condition


                //>>4 Condition
                IF NOT IncValuation THEN BEGIN
                    ValueEntry.SETRANGE("Posting Date", StartDate, EndDate);
                    ValueEntry.SETRANGE("Item Ledger Entry Type", ValueEntry."Item Ledger Entry Type"::Transfer);
                    IF ValueEntry.FINDSET THEN
                        REPEAT
                            IF TRUE IN [ValueEntry."Valued Quantity" < 0, NOT GetOutboundItemEntry(ValueEntry."Item Ledger Entry No.")] THEN
                                AssignAmounts(ValueEntry, DecreaseInvoicedValue, DecreaseInvoicedQty, DecreaseExpectedValue, DecreaseExpectedQty, -1)
                            ELSE
                                AssignAmounts(ValueEntry, IncreaseInvoicedValue, IncreaseInvoicedQty, IncreaseExpectedValue, IncreaseExpectedQty, 1);
                        UNTIL ValueEntry.NEXT = 0;
                END ELSE BEGIN
                    Loc20.RESET;
                    Loc20.SETRANGE(Closed, FALSE);
                    Loc20.SETRANGE("Include in Valuation Report", TRUE);
                    IF Loc20.FINDSET THEN
                        REPEAT
                            VE20.RESET;
                            VE20.SETCURRENTKEY("Item No.");
                            VE20.SETRANGE("Item No.", "No.");
                            IF GETFILTER("Variant Filter") <> '' THEN
                                VE20.SETFILTER("Variant Code", GETFILTER("Variant Filter"));
                            VE20.SETFILTER("Location Code", Loc20.Code);
                            IF GETFILTER("Global Dimension 1 Filter") <> '' THEN
                                VE20.SETFILTER("Global Dimension 1 Code", GETFILTER("Global Dimension 1 Filter"));
                            IF GETFILTER("Global Dimension 2 Filter") <> '' THEN
                                VE20.SETFILTER("Global Dimension 2 Code", GETFILTER("Global Dimension 2 Filter"));
                            VE20.SETRANGE("Posting Date", StartDate, EndDate);
                            VE20.SETRANGE("Item Ledger Entry Type", VE20."Item Ledger Entry Type"::Transfer);
                            IF VE20.FINDSET THEN
                                REPEAT
                                    IF TRUE IN [VE20."Valued Quantity" < 0, NOT GetOutboundItemEntrywithLocation(VE20."Item Ledger Entry No.", Loc20.Code)] THEN
                                        AssignAmounts(VE20, DecreaseInvoicedValue, DecreaseInvoicedQty, DecreaseExpectedValue, DecreaseExpectedQty, -1)
                                    ELSE
                                        AssignAmounts(VE20, IncreaseInvoicedValue, IncreaseInvoicedQty, IncreaseExpectedValue, IncreaseExpectedQty, 1);
                                UNTIL VE20.NEXT = 0;
                        UNTIL Loc20.NEXT = 0;

                END;
                //>>4 Condition

                /*
                //>>17July2018
                IF StartingInvoicedQty = 0 THEN
                  StartingInvoicedValue := 0;
                
                //IF IncreaseInvoicedQty = 0 THEN
                  //IncreaseInvoicedValue := 0;
                
                IF DecreaseInvoicedQty = 0 THEN
                  DecreaseInvoicedValue := 0;
                //<<17July2018
                *///Code Commented 25Sep2018

                IsEmptyLine := IsEmptyLine AND ((IncreaseInvoicedValue = 0) AND (IncreaseInvoicedQty = 0));
                IsEmptyLine := IsEmptyLine AND ((DecreaseInvoicedValue = 0) AND (DecreaseInvoicedQty = 0));
                IF ShowExpected THEN BEGIN
                    IsEmptyLine := IsEmptyLine AND ((IncreaseExpectedValue = 0) AND (IncreaseExpectedQty = 0));
                    IsEmptyLine := IsEmptyLine AND ((DecreaseExpectedValue = 0) AND (DecreaseExpectedQty = 0));
                END;

                //>>5 Condition
                IF NOT IncValuation THEN BEGIN
                    ValueEntry.SETRANGE("Posting Date", 0D, EndDate);
                    ValueEntry.SETRANGE("Item Ledger Entry Type");
                    ValueEntry.CALCSUMS("Cost Posted to G/L", "Expected Cost Posted to G/L");
                    ExpCostPostedToGL += ValueEntry."Expected Cost Posted to G/L";
                    InvCostPostedToGL += ValueEntry."Cost Posted to G/L";
                END ELSE BEGIN
                    Loc20.RESET;
                    Loc20.SETRANGE(Closed, FALSE);
                    Loc20.SETRANGE("Include in Valuation Report", TRUE);
                    IF Loc20.FINDSET THEN
                        REPEAT
                            VE20.RESET;
                            VE20.SETCURRENTKEY("Item No.");
                            VE20.SETRANGE("Item No.", "No.");
                            IF GETFILTER("Variant Filter") <> '' THEN
                                VE20.SETFILTER("Variant Code", GETFILTER("Variant Filter"));
                            VE20.SETFILTER("Location Code", Loc20.Code);
                            IF GETFILTER("Global Dimension 1 Filter") <> '' THEN
                                VE20.SETFILTER("Global Dimension 1 Code", GETFILTER("Global Dimension 1 Filter"));
                            IF GETFILTER("Global Dimension 2 Filter") <> '' THEN
                                VE20.SETFILTER("Global Dimension 2 Code", GETFILTER("Global Dimension 2 Filter"));
                            VE20.SETRANGE("Posting Date", 0D, EndDate);
                            IF VE20.FINDFIRST THEN BEGIN
                                VE20.CALCSUMS("Cost Posted to G/L", "Expected Cost Posted to G/L");
                                ExpCostPostedToGL += VE20."Expected Cost Posted to G/L";
                                InvCostPostedToGL += VE20."Cost Posted to G/L";
                            END;
                        UNTIL Loc20.NEXT = 0;

                END;
                //<<5 Condition

                StartingExpectedValue += StartingInvoicedValue;
                IncreaseExpectedValue += IncreaseInvoicedValue;
                DecreaseExpectedValue += DecreaseInvoicedValue;
                CostPostedToGL := ExpCostPostedToGL + InvCostPostedToGL;

                IF IsEmptyLine THEN
                    CurrReport.SKIP;

            end;

            trigger OnPreDataItem()
            begin
                CurrReport.CREATETOTALS(
                  StartingExpectedQty, IncreaseExpectedQty, DecreaseExpectedQty,
                  StartingInvoicedQty, IncreaseInvoicedQty, DecreaseInvoicedQty);
                CurrReport.CREATETOTALS(
                  StartingExpectedValue, IncreaseExpectedValue, DecreaseExpectedValue,
                  StartingInvoicedValue, IncreaseInvoicedValue, DecreaseInvoicedValue,
                  CostPostedToGL, ExpCostPostedToGL, InvCostPostedToGL);
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
                group(Options)
                {
                    Caption = 'Options';
                    field(StartingDate; StartDate)
                    {
                        ApplicationArea = ALL;
                        Caption = 'Starting Date';
                    }
                    field(EndingDate; EndDate)
                    {
                        ApplicationArea = ALL;
                        Caption = 'Ending Date';
                    }
                    field(IncludeExpectedCost; ShowExpected)
                    {
                        ApplicationArea = ALL;
                        Caption = 'Include Expected Cost';
                        Visible = false;
                    }
                    field("Include Valuation Location Only"; IncValuation)
                    {
                        Visible = false;
                        ApplicationArea = ALL;
                        trigger OnValidate()
                        begin

                            IF IncValuation = TRUE THEN BEGIN
                                Item."Location Filter" := '';
                            END;
                        end;
                    }
                }
            }
        }

        actions
        {
        }

        trigger OnOpenPage()
        begin
            IF (StartDate = 0D) AND (EndDate = 0D) THEN
                EndDate := WORKDATE;

            //>>17Jun2018
            IncValuation := TRUE;
            Item."Location Filter" := '';
            //<<17Jun2018
        end;
    }

    labels
    {
        Inventory_Posting_Group_NameCaption = 'Inventory Posting Group Name';
        Expected_CostCaption = 'Expected Cost';
    }

    trigger OnPreReport()
    begin
        IF (StartDate = 0D) AND (EndDate = 0D) THEN
            EndDate := WORKDATE;

        IF StartDate IN [0D, 00000101D/*01010000D*/] THEN
            StartDateText := ''
        ELSE
            StartDateText := FORMAT(StartDate - 1);

        ItemFilter := Item.GETFILTERS;

        //>>17Jul2018
        IF Item.GETFILTER("Location Filter") <> '' THEN
            IncValuation := FALSE
        ELSE
            IncValuation := TRUE;

        IF IncValuation THEN
            IncludeOnlyValuationLocation;
        //<<17Jul2018
    end;

    var
        Text005: Label 'As of %1';
        ValueEntry: Record 5802;
        StartDate: Date;
        EndDate: Date;
        ShowExpected: Boolean;
        ItemFilter: Text;
        StartDateText: Text[10];
        StartingInvoicedValue: Decimal;
        StartingExpectedValue: Decimal;
        StartingInvoicedQty: Decimal;
        StartingExpectedQty: Decimal;
        IncreaseInvoicedValue: Decimal;
        IncreaseExpectedValue: Decimal;
        IncreaseInvoicedQty: Decimal;
        IncreaseExpectedQty: Decimal;
        DecreaseInvoicedValue: Decimal;
        DecreaseExpectedValue: Decimal;
        DecreaseInvoicedQty: Decimal;
        DecreaseExpectedQty: Decimal;
        BoM_TextLbl: Label 'Base UoM';
        Inventory_ValuationCaptionLbl: Label 'Inventory Valuation';
        CurrReport_PAGENOCaptionLbl: Label 'Page';
        This_report_includes_entries_that_have_been_posted_with_expected_costs_CaptionLbl: Label 'This report includes entries that have been posted with expected costs.';
        IncreaseInvoicedQtyCaptionLbl: Label 'Increases (LCY)';
        DecreaseInvoicedQtyCaptionLbl: Label 'Decreases (LCY)';
        QuantityCaptionLbl: Label 'Quantity';
        ValueCaptionLbl: Label 'Value';
        QuantityCaption_Control31Lbl: Label 'Quantity';
        QuantityCaption_Control40Lbl: Label 'Quantity';
        InvCostPostedToGL_Control53CaptionLbl: Label 'Cost Posted to G/L';
        QuantityCaption_Control58Lbl: Label 'Quantity';
        TotalCaptionLbl: Label 'Total';
        Expected_Cost_Included_TotalCaptionLbl: Label 'Expected Cost Included Total';
        Expected_Cost_TotalCaptionLbl: Label 'Expected Cost Total';
        Expected_Cost_IncludedCaptionLbl: Label 'Expected Cost Included';
        InvCostPostedToGL: Decimal;
        CostPostedToGL: Decimal;
        ExpCostPostedToGL: Decimal;
        IsEmptyLine: Boolean;
        "---------17Jul2018": Integer;
        IncValuation: Boolean;
        Loc17: Record 14;
        LocFilter17: Code[500];
        I17: Integer;
        "-----20Jul2018": Integer;
        Loc20: Record 14;
        VE20: Record 5802;

    local procedure AssignAmounts(ValueEntry: Record 5802; var InvoicedValue: Decimal; var InvoicedQty: Decimal; var ExpectedValue: Decimal; var ExpectedQty: Decimal; Sign: Decimal)
    begin
        InvoicedValue += ValueEntry."Cost Amount (Actual)" * Sign;
        InvoicedQty += ValueEntry."Invoiced Quantity" * Sign;
        ExpectedValue += ValueEntry."Cost Amount (Expected)" * Sign;
        ExpectedQty += ValueEntry."Item Ledger Entry Quantity" * Sign;
    end;

    local procedure GetOutboundItemEntry(ItemLedgerEntryNo: Integer): Boolean
    var
        ItemApplnEntry: Record 339;
        ItemLedgEntry: Record 32;
    begin
        ItemApplnEntry.SETCURRENTKEY("Item Ledger Entry No.");
        ItemApplnEntry.SETRANGE("Item Ledger Entry No.", ItemLedgerEntryNo);
        IF NOT ItemApplnEntry.FINDFIRST THEN
            EXIT(TRUE);

        ItemLedgEntry.SETRANGE("Item No.", Item."No.");
        ItemLedgEntry.SETFILTER("Variant Code", Item.GETFILTER("Variant Filter"));
        //ItemLedgEntry.SETFILTER("Location Code",Item.GETFILTER("Location Filter"));
        //>>17Jul2018
        IF IncValuation THEN
            ItemLedgEntry.SETFILTER("Location Code", LocFilter17)
        ELSE
            ItemLedgEntry.SETFILTER("Location Code", Item.GETFILTER("Location Filter"));
        //<<17Jul2018
        ItemLedgEntry.SETFILTER("Global Dimension 1 Code", Item.GETFILTER("Global Dimension 1 Filter"));
        ItemLedgEntry.SETFILTER("Global Dimension 2 Code", Item.GETFILTER("Global Dimension 2 Filter"));
        ItemLedgEntry."Entry No." := ItemApplnEntry."Outbound Item Entry No.";
        EXIT(NOT ItemLedgEntry.FIND);
    end;

    // //[Scope('Internal')]
    procedure SetStartDate(DateValue: Date)
    begin
        StartDate := DateValue;
    end;

    //  //[Scope('Internal')]
    procedure SetEndDate(DateValue: Date)
    begin
        EndDate := DateValue;
    end;

    //   //[Scope('Internal')]
    procedure InitializeRequest(NewStartDate: Date; NewEndDate: Date; NewShowExpected: Boolean)
    begin
        StartDate := NewStartDate;
        EndDate := NewEndDate;
        ShowExpected := NewShowExpected;
    end;

    local procedure GetUrlForReportDrilldown(ItemNumber: Code[20]): Text
    begin
        // Generates a URL to the report which sets tab "Item" and field "Field1" on the request page, such as
        // dynamicsnav://hostname:port/instance/company/runreport?report=5801<&Tenant=tenantId>&filter=Item.Field1:1100.
        // TODO
        // Eventually leverage parameters 5 and 6 of GETURL by adding ",Item,TRUE)" and
        // use filter Item.SETFILTER("No.",'=%1',ItemNumber);.
        EXIT(GETURL(CURRENTCLIENTTYPE, COMPANYNAME, OBJECTTYPE::Report, REPORT::"Invt. Valuation - Cost Spec.") +
          STRSUBSTNO('&filter=Item.Field1:%1', ItemNumber));
    end;

    local procedure IncludeOnlyValuationLocation()
    begin
        CLEAR(LocFilter17);
        CLEAR(I17);
        Loc17.RESET;
        Loc17.SETRANGE("Include in Valuation Report", FALSE);
        IF Loc17.FINDSET THEN
            REPEAT
                I17 += 1;
                IF I17 = 1 THEN
                    LocFilter17 := '<>' + Loc17.Code;
                IF I17 > 1 THEN
                    LocFilter17 += '&<>' + Loc17.Code;
            UNTIL Loc17.NEXT = 0;
    end;

    local procedure GetOutboundItemEntrywithLocation(ItemLedgerEntryNo: Integer; LocationCode: Code[20]): Boolean
    var
        ItemApplnEntry: Record 339;
        ItemLedgEntry: Record 32;
    begin
        ItemApplnEntry.SETCURRENTKEY("Item Ledger Entry No.");
        ItemApplnEntry.SETRANGE("Item Ledger Entry No.", ItemLedgerEntryNo);
        IF NOT ItemApplnEntry.FINDFIRST THEN
            EXIT(TRUE);

        ItemLedgEntry.SETRANGE("Item No.", Item."No.");
        ItemLedgEntry.SETFILTER("Variant Code", Item.GETFILTER("Variant Filter"));
        ItemLedgEntry.SETFILTER("Location Code", LocationCode);
        ItemLedgEntry.SETFILTER("Global Dimension 1 Code", Item.GETFILTER("Global Dimension 1 Filter"));
        ItemLedgEntry.SETFILTER("Global Dimension 2 Code", Item.GETFILTER("Global Dimension 2 Filter"));
        ItemLedgEntry."Entry No." := ItemApplnEntry."Outbound Item Entry No.";
        EXIT(NOT ItemLedgEntry.FIND);
    end;
}


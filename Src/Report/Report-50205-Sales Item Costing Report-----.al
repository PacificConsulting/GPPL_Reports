report 50205 "Sales Item Costing Report-----"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 21Jul2018   RB-N         New Report for Sales Costing.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesItemCostingReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Sales Invoice Line"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.")
                                WHERE(Type = CONST(Item));
            RequestFilterFields = "No.", "Posting Date", "Lot No.", "Document No.";

            trigger OnAfterGetRecord()
            begin

                //>>
                CLEAR(RMCost);
                CLEAR(PMCost);
                CLEAR(TempAmt);

                ILE_Output.RESET;
                ILE_Output.SETCURRENTKEY("Item No.");
                ILE_Output.SETRANGE("Item No.", "No.");
                ILE_Output.SETRANGE("Lot No.", "Lot No.");
                ILE_Output.SETRANGE("Entry Type", ILE_Output."Entry Type"::Output);
                IF ILE_Output.FINDLAST THEN BEGIN

                    //>>Output Amount
                    VE_Actual.RESET;
                    VE_Actual.SETCURRENTKEY("Document No.");
                    VE_Actual.SETRANGE("Document No.", ILE_Output."Document No.");
                    VE_Actual.SETRANGE("Source No.", ILE_Output."Item No.");
                    VE_Actual.SETRANGE("Item Ledger Entry Type", VE_Actual."Item Ledger Entry Type"::Output);
                    IF VE_Actual.FINDFIRST THEN BEGIN
                        VE_Actual.CALCSUMS("Cost Amount (Actual)");
                        TempAmt := ABS(VE_Actual."Cost Amount (Actual)");
                    END;
                    //>>Output Amount

                    //>>Packing Material Cost
                    VE_Actual.RESET;
                    VE_Actual.SETCURRENTKEY("Document No.");
                    VE_Actual.SETRANGE("Document No.", ILE_Output."Document No.");
                    VE_Actual.SETRANGE("Source No.", ILE_Output."Item No.");
                    VE_Actual.SETRANGE("Item Ledger Entry Type", VE_Actual."Item Ledger Entry Type"::Consumption);
                    VE_Actual.SETFILTER("Gen. Prod. Posting Group", '<>%1', 'BULK');
                    IF VE_Actual.FINDFIRST THEN BEGIN
                        VE_Actual.CALCSUMS("Cost Amount (Actual)");
                        PMCost := ABS(VE_Actual."Cost Amount (Actual)");
                    END;
                    //>>Packing Material Cost

                    //>>Raw Material Cost
                    ILE01.RESET;
                    ILE01.SETCURRENTKEY("Document No.", "Posting Date", "Item No.");
                    //ILE01.SETRANGE("Item No.",ILE_Output."Item No.");
                    ILE01.SETRANGE("Document No.", ILE_Output."Document No.");
                    ILE01.SETRANGE("Entry Type", ILE01."Entry Type"::Consumption);
                    ILE01.SETRANGE("Item Tracking", ILE01."Item Tracking"::"Lot No.");
                    IF ILE01.FINDSET THEN
                        REPEAT
                            //MESSAGE('Entry No. %1',ILE01."Entry No.");
                            FindAppliedEntry(ILE01);


                        UNTIL ILE01.NEXT = 0;

                    /*
                      VE_Actual.RESET;
                      VE_Actual.SETCURRENTKEY("Document No.");
                      VE_Actual.SETRANGE("Document No.",ILE_Output."Document No.");
                      VE_Actual.SETRANGE("Source No.",ILE_Output."Item No.");
                      VE_Actual.SETRANGE("Item Ledger Entry Type",VE_Actual."Item Ledger Entry Type"::Consumption);
                      VE_Actual.SETRANGE("Gen. Prod. Posting Group",'BULK');
                      IF VE_Actual.FINDFIRST THEN
                      BEGIN
                        VE_Actual.CALCSUMS("Cost Amount (Actual)");
                        RMCost := ABS(VE_Actual."Cost Amount (Actual)" );
                      END;
                    */
                    //>>Raw Material Cost


                END;
                //<<

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
    }

    trigger OnPreReport()
    begin
        TempItemEntry.DELETEALL;
    end;

    var
        RMCost: Decimal;
        PMCost: Decimal;
        TempAmt: Decimal;
        ILE_Output: Record 32;
        VE_Actual: Record 5802;
        ILE01: Record 32;
        ILE02: Record 32;
        ItmAppEntry1: Record 339;
        ItmAppEntry2: Record 339;
        TempItemEntry: Record 32 temporary;
        NewILEEntryNo: Decimal;

    local procedure FindAppliedEntry(ItemLedgEntry: Record 32)
    var
        ItemApplnEntry: Record 339;
    begin
        WITH ItemLedgEntry DO
            IF Positive THEN BEGIN
                ItemApplnEntry.RESET;
                ItemApplnEntry.SETCURRENTKEY("Inbound Item Entry No.", "Outbound Item Entry No.", "Cost Application");
                ItemApplnEntry.SETRANGE("Inbound Item Entry No.", "Entry No.");
                ItemApplnEntry.SETFILTER("Outbound Item Entry No.", '<>%1', 0);
                ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                IF ItemApplnEntry.FIND('-') THEN
                    REPEAT
                        InsertNewCost(ItemApplnEntry."Outbound Item Entry No.", ItemApplnEntry.Quantity);
                    UNTIL ItemApplnEntry.NEXT = 0;
            END ELSE BEGIN
                ItemApplnEntry.RESET;
                ItemApplnEntry.SETCURRENTKEY("Outbound Item Entry No.", "Item Ledger Entry No.", "Cost Application");
                ItemApplnEntry.SETRANGE("Outbound Item Entry No.", "Entry No.");
                ItemApplnEntry.SETRANGE("Item Ledger Entry No.", "Entry No.");
                ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                IF ItemApplnEntry.FIND('-') THEN
                    REPEAT
                        InsertNewCost(ItemApplnEntry."Inbound Item Entry No.", -ItemApplnEntry.Quantity);
                    UNTIL ItemApplnEntry.NEXT = 0;
            END;
    end;

    local procedure InsertNewCost(EntryNo: Integer; AppliedQty: Decimal)
    var
        ItemLedgEntry: Record 32;
    begin
        ItemLedgEntry.GET(EntryNo);
        NewILEEntryNo := ItemLedgEntry."Entry No.";
        /*
        IF AppliedQty * ItemLedgEntry.Quantity < 0 THEN
          EXIT;
        
        IF NOT TempItemEntry.GET(EntryNo) THEN BEGIN
          TempItemEntry.INIT;
          TempItemEntry := ItemLedgEntry;
          TempItemEntry.Quantity := AppliedQty;
          TempItemEntry.INSERT;
        END ELSE BEGIN
          TempItemEntry.Quantity := TempItemEntry.Quantity + AppliedQty;
          TempItemEntry.MODIFY;
        END;
        */

    end;
}


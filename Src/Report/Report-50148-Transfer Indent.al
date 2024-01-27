report 50148 "Transfer Indent"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TransferIndent.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Transfer Indent Header"; "Transfer Indent Header")
        {
            RequestFilterFields = "No.";
            column(CompName; CompInfo.Name)
            {
            }
            column(HDate; FORMAT(TODAY, 0, 4))
            {
            }
            column(DocNo; "Transfer Indent Header"."No.")
            {
            }
            column(PostingDate; FORMAT("Transfer Indent Header"."Posting Date", 0, 4))
            {
            }
            column(TranToAdd1; TransferToAddr[1])
            {
            }
            column(TranToAdd2; TransferToAddr[2])
            {
            }
            column(TranToAdd3; TransferToAddr[3])
            {
            }
            column(TranToAdd4; TransferToAddr[4])
            {
            }
            column(TranToAdd5; TransferToAddr[5])
            {
            }
            column(TranToAdd6; TransferToAddr[6])
            {
            }
            column(TranToAdd7; TransferToAddr[7])
            {
            }
            column(TranToAdd8; TransferToAddr[8])
            {
            }
            column(TranFromAdd1; TransferFromAddr[1])
            {
            }
            column(TranFromAdd2; TransferFromAddr[2])
            {
            }
            column(TranFromAdd3; TransferFromAddr[3])
            {
            }
            column(TranFromAdd4; TransferFromAddr[4])
            {
            }
            column(TranFromAdd5; TransferFromAddr[5])
            {
            }
            column(TranFromAdd6; TransferFromAddr[6])
            {
            }
            column(TranToR; "TransferToRes.C")
            {
            }
            column(TranFromR; "TransferFromRes.C")
            {
            }
            column(DivName; 'DIVISION :- ' + recDimsionValue.Name)
            {
            }
            column(IsApproved; IsApproved)
            {
            }
            column(ShipDescription; ShipmentMethod.Description)
            {
            }
            column(Capacity; Capacity)
            {
            }
            dataitem("Transfer Indent Line"; "Transfer Indent Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = WHERE(Approve = FILTER(false),
                                          Closed = FILTER(false));
                column(SrNo; SrNo)
                {
                }
                column(LDocNo; "Transfer Indent Line"."Document No.")
                {
                }
                column(LineNo; "Transfer Indent Line"."Line No.")
                {
                }
                column(CreationDate; "Addition date")
                {
                }
                column(ItemDesc; ItemDesc)
                {
                }
                column(Qty; Quantity)
                {
                }
                column(UoM; "Transfer Indent Line"."Unit of Measure Code")
                {
                }
                column(TotTransferQty; TotTransferQty)
                {
                }
                column(TransferPrice; "Transfer Indent Line"."Transfer Price")
                {
                }
                column(InvCDate; InventoryatCurDate)
                {
                }
                column(UserName; UserName)
                {
                }
                column(ApprovalName; ApprovalName)
                {
                }
                column(ApprovalDate; "Transfer Indent Header"."Approval Date")
                {
                }
                column(ApprovalTime; "Transfer Indent Header"."Approval Time")
                {
                }
                column(NAH; NAH)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    ILEQTY := 0;
                    InventoryatCurDate := 0;

                    recItemUOM.RESET;
                    recItemUOM.SETRANGE(recItemUOM."Item No.", "Transfer Indent Line"."Item No.");
                    recItemUOM.SETRANGE(recItemUOM.Code, "Transfer Indent Line"."Unit of Measure Code");
                    IF recItemUOM.FINDFIRST THEN
                        TotTransferQty := recItemUOM."Qty. per Unit of Measure" * "Transfer Indent Line".Quantity;
                    /*
                    DocDim2.SETRANGE("Table ID",DATABASE::"Transfer Indent Line");
                    DocDim2.SETRANGE("Document No.","Transfer Indent Line"."Document No.");
                    DocDim2.SETRANGE("Line No.","Transfer Indent Line"."Line No.");
                    
                    *///Commented T#357 29Mar2017

                    SrNo := SrNo + 1;

                    recILE.RESET;
                    recILE.SETRANGE(recILE."Item No.", "Transfer Indent Line"."Item No.");
                    recILE.SETRANGE(recILE."Location Code", "Transfer Indent Line"."Transfer-to Code");
                    IF recILE.FINDFIRST THEN
                        REPEAT
                            ILEQTY := ILEQTY + recILE.Quantity;
                        UNTIL recILE.NEXT = 0;

                    IF ILEQTY <> 0 THEN BEGIN
                        InventoryatCurDate := ILEQTY / recILE."Qty. per Unit of Measure";
                    END;

                    IF "Transfer Indent Line"."Inventory Posting Group" = 'MERCH' THEN BEGIN
                        ItemDesc := "Transfer Indent Line"."Merch Name"
                    END ELSE
                        ItemDesc := "Transfer Indent Line".Description;
                    //<<1

                    //>>2

                    //Transfer Indent Line, Body (2) - OnPreSection()
                    TotTransferQty1 += TotTransferQty;
                    //<<2

                    //>>3

                    //Transfer Indent Line, Footer (5) - OnPreSection()
                    UserRec.RESET;
                    //UserRec.SETRANGE(UserRec."User ID","Transfer Indent Header"."USER ID");
                    UserRec.SETRANGE(UserRec."User Name", "Transfer Indent Header"."USER ID");//29Mar2017
                    IF UserRec.FINDFIRST THEN
                        UserName := UserRec."Full Name";//29Mar2017
                                                        //UserName:=UserRec.Name;

                    UserRec1.RESET;
                    //UserRec1.SETRANGE(UserRec1."User ID","Transfer Indent Header"."Approve User ID");
                    UserRec1.SETRANGE(UserRec1."User Name", "Transfer Indent Header"."Approve User ID");//29Mar2017
                    IF UserRec1.FINDFIRST THEN
                        ApprovalName := UserRec1."Full Name";//29Mar2017
                                                             //ApprovalName:=UserRec1.Name;
                                                             //<<3

                end;

                trigger OnPreDataItem()
                begin

                    NAH := COUNT;
                    //MESSAGE('No of Count %1',NAH);
                end;
            }
            dataitem("Item Ledger Entry"; "Item Ledger Entry")
            {
                DataItemLink = "Location Code" = FIELD("Transfer-to Code");
                DataItemTableView = SORTING("Global Dimension 1 Code", "Location Code", "Item Category Code", "Item No.", "Lot No.", "Posting Date", "Document No.", "Document Line No.")
                                    ORDER(Ascending)
                                    WHERE("Item Category Code" = FILTER('CAT02' | 'CAT03' | 'CAT11' | 'CAT12' | 'CAT13'));
                column(Qtyindustrial; Qtyindustrial)
                {
                }
                column(QtyAuto; QtyAuto)
                {
                }
                column(QtyRPO; QtyRPO)
                {
                }
                column(QtyBOILS; QtyBOILS)
                {
                }
                column(QtyWhiteOils; QtyWhiteOils)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    IF "Item Ledger Entry"."Global Dimension 1 Code" = 'DIV-03' THEN
                        Qtyindustrial += "Item Ledger Entry".Quantity;
                    IF "Item Ledger Entry"."Global Dimension 1 Code" = 'DIV-04' THEN
                        QtyAuto += "Item Ledger Entry".Quantity;

                    //RSPL008-START
                    IF "Item Ledger Entry"."Global Dimension 1 Code" = 'DIV-05' THEN
                        QtyRPO += "Item Ledger Entry".Quantity;
                    IF "Item Ledger Entry"."Global Dimension 1 Code" = 'DIV-06' THEN
                        QtyBOILS += "Item Ledger Entry".Quantity;
                    IF "Item Ledger Entry"."Global Dimension 1 Code" = 'DIV-07' THEN
                        QtyWhiteOils += "Item Ledger Entry".Quantity;
                    //RSPL008-END

                    //<<1
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //>>1
                /*
                DocDim1.SETRANGE("Table ID",DATABASE::"Transfer Indent Header");
                DocDim1.SETRANGE("Document No.","Transfer Indent Header"."No.");
                *///Commented T#357 29Mar2017

                //FormatAddr.TransferIndHeaderTransferFrom(TransferFromAddr, "Transfer Indent Header");
                ///FormatAddr.TransferIndHeaderTransferTo(TransferToAddr, "Transfer Indent Header");

                IF ShipmentMethod.GET("Shipment Method Code") THEN;
                //ShipmentMethod.INIT;//29Mar2017

                ResponcibilityCenterRec1.RESET;
                ResponcibilityCenterRec1.SETRANGE(ResponcibilityCenterRec1."Location Code", "Transfer Indent Header"."Transfer-to Code");
                IF ResponcibilityCenterRec1.FINDFIRST THEN
                    "TransferToRes.C" := 'Responsibility Center :- ' + ResponcibilityCenterRec1.Code + ' - ' + ResponcibilityCenterRec1.Name;

                ResponcibilityCenterRec2.RESET;
                ResponcibilityCenterRec2.SETRANGE(ResponcibilityCenterRec2."Location Code", "Transfer Indent Header"."Transfer-from Code");
                IF ResponcibilityCenterRec2.FINDFIRST THEN
                    "TransferFromRes.C" := 'Responsibility Center :- ' + ResponcibilityCenterRec2.Code + ' - ' + ResponcibilityCenterRec2.Name;

                RecLoc.RESET;
                RecLoc.SETRANGE(RecLoc.Code, "Transfer Indent Header"."Transfer-to Code");
                IF RecLoc.FINDFIRST THEN
                    Capacity := RecLoc."Strorage Capacity Industrial";
                //<<1

                //>>2

                //Transfer Indent Header, Header (1) - OnPreSection()
                IF "Transfer Indent Header".Approve = TRUE THEN
                    IsApproved := 'APPROVED'
                ELSE
                    IsApproved := '';


                recDimsionValue.RESET;
                recDimsionValue.SETRANGE(recDimsionValue.Code, "Transfer Indent Header"."Shortcut Dimension 1 Code");
                IF recDimsionValue.FINDFIRST THEN;
                //<<2


                //>>3

                //Transfer Indent Header, Body (2) - OnPreSection()
                CLEAR(TotTransferQty1);
                //<<3

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

        //>>1
        CompInfo.GET;
        //<<1
    end;

    var
        ShipmentMethod: Record 10;
        FormatAddr: Codeunit 365;
        TransferFromAddr: array[8] of Text[50];
        TransferToAddr: array[8] of Text[50];
        NoOfCopies: Integer;
        NoOfLoops: Integer;
        CopyText: Text[30];
        DimText: Text[120];
        OldDimText: Text[75];
        ShowInternalInfo: Boolean;
        Continue: Boolean;
        SrNo: Integer;
        CompInfo: Record 79;
        ResponcibilityCenterRec1: Record 5714;
        ResponcibilityCenterRec2: Record 5714;
        Location1: Record 14;
        Location2: Record 14;
        IsApproved: Text[30];
        ItmRec: Record 27;
        TotTransferQty: Decimal;
        UserRec: Record 2000000120;
        UserName: Text[80];
        UserRec1: Record 2000000120;
        ApprovalName: Text[80];
        TotTransferQty1: Decimal;
        recItemUOM: Record 5404;
        "TransferToRes.C": Text[100];
        "TransferFromRes.C": Text[100];
        InventoryatCurDate: Decimal;
        recILE: Record 32;
        ILEQTY: Decimal;
        Capacity: Code[10];
        RecLoc: Record 14;
        Qtyindustrial: Decimal;
        QtyAuto: Decimal;
        QtyRPO: Decimal;
        QtyBOILS: Decimal;
        QtyWhiteOils: Decimal;
        recDimsionValue: Record 349;
        ItemDesc: Text[50];
        Text000: Label 'COPY';
        Text001: Label 'Transfer Order %1';
        Text002: Label 'Page %1';
        "----29Mar2017": Integer;
        NAH: Integer;
}


report 50160 "Finished Goods - Ageing-N"
{
    // //RSPL Sharath Added code on 06-Sep-2017 for 'BRT no', location name for 'IN-TRANS', and bin code for 'PLANT01'.
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 08Sep2017   RB-N         Proper Alignment of TotalQty & GrandTotalQty In Excel
    //                          Commented TradingLoction Filter and New Code added for PLANT Locations
    //                          4 New Fields added(IndentNo,ApprovalDate,ApprovalTime,ApproverName)
    // 
    // 12Sep2017   RB-N         Value Amount retireved from IF Transfer -- Transfer Shipment Line,
    //                          Value Amount retireved from IF Sales Return -- Sales Cr Memo Line,
    //                          New Field (Transfer Price)
    // 
    // 06Oct2017   RB-N         IF IndentNo is Blank then Displaying TransferOrderNo. & BRT No.
    // 15May2018   RB-N         Issue in the No.of Days Calculation
    // 22May2018   RB-N         If ManufacturingDate is Blank , Then GRNDate will be considered as ManufacturingDate.
    // 10Aug2018   RB-N         Issue for Plant01 Transfer Price & Transfer Value
    // 27Aug2018   RB-N         Values for Ageing Period
    // 31Aug2018   RB-N         Item Category Description
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/FinishedGoodsAgeingN.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Location; Location)
        {
            DataItemTableView = SORTING(Code)
                                ORDER(Ascending)
                                WHERE(Code = FILTER(<> 'INVADJ0001'));
            RequestFilterFields = "Code";
            column(HDate; 'DATED :- ' + FORMAT(TODAY))
            {
            }
            column(LocName; 'FG STOCK DETAILS FOR  --------> ' + Location.Name)
            {
            }
            dataitem("Item Ledger Entry"; "Item Ledger Entry")
            {
                DataItemLink = "Location Code" = FIELD(Code);
                DataItemTableView = SORTING("Item No.", "Location Code", "Lot No.", "Variant Code", "Posting Date")
                                    ORDER(Ascending)
                                    WHERE("Lot No." = FILTER(<> 'REV'),
                                          "Item Category Code" = FILTER('CAT02' | 'CAT03' | 'CAT11' | 'CAT12' | 'CAT13' | 'CAT15' | 'CAT18' | 'CAT20' | 'CAT22'),
                                          "Remaining Quantity" = FILTER(<> 0));
                RequestFilterFields = "Item No.", "Lot No.", "Item Category Code", "Global Dimension 1 Code";
                column(Srno; Srno)
                {
                }
                column(ItemDesc; Item.Description)
                {
                }
                column(BatchNo; "Item Ledger Entry"."Lot No.")
                {
                }
                column(PackType; "Item Ledger Entry"."Unit of Measure Code")
                {
                }
                column(PackSize; "Item Ledger Entry"."Qty. per Unit of Measure")
                {
                }
                column(GNoPack; GNoPack)
                {
                }
                column(TNoPack; TNoPack)
                {
                }
                column(NoPack; "Item Ledger Entry"."Remaining Quantity" / "Item Ledger Entry"."Qty. per Unit of Measure")
                {
                }
                column(RemQty; "NoofBarrels/Pails")
                {
                }
                column(MfgDate; vMfgDate)
                {
                }
                column(RcptDate; vRcptDate)
                {
                }
                column(Qty1; RemQty[1])
                {
                }
                column(Qty2; RemQty[2])
                {
                }
                column(Qty3; RemQty[3])
                {
                }
                column(Qty4; RemQty[4])
                {
                }
                column(Qty5; RemQty[5])
                {
                }
                column(Date1; '0-' + FORMAT(NoOfDays))
                {
                }
                column(Date2; FORMAT(NoOfDays + 1) + '-' + FORMAT(NoOfDays1))
                {
                }
                column(Date3; FORMAT(NoOfDays1 + 1) + '-' + FORMAT(NoOfDays2))
                {
                }
                column(Date4; FORMAT(NoOfDays2 + 1) + '-' + FORMAT(NoOfDays3))
                {
                }
                column(Date5; '>' + FORMAT(NoOfDays3))
                {
                }
                column(ItmNo; "Item No.")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>04Sep2018
                    RepItmNo := '';
                    IF "Item Category Code" = 'CAT15' THEN BEGIN
                        RepItmNo := COPYSTR("Item No.", 1, 4);
                        IF RepItmNo <> 'FGRP' THEN
                            CurrReport.SKIP;
                    END;
                    //<<04Sep2018

                    //>>1


                    Item.GET("Item Ledger Entry"."Item No.");
                    //MESSAGE('Test');

                    //IF Item.Blocked THEN   //Sourav21122015
                    //CurrReport.SKIP;
                    /*
                    //IPOL/001 >>>> FG Item entries has to show
                    IF Item."Item Tracking Code"<>'FGIL|FGAL' THEN
                       CurrReport.SKIP;
                    //IPOL/001 <<<< FG Item entries has to show
                    */
                    ItemUOM1.RESET;
                    ItemUOM1.SETRANGE(ItemUOM1."Item No.", "Item Ledger Entry"."Item No.");
                    ItemUOM1.SETRANGE(ItemUOM1.Code, "Item Ledger Entry"."Unit of Measure Code");
                    IF ItemUOM1.FINDFIRST THEN;
                    //MESSAGE('%1',Value);
                    "NoofBarrels/Pails" := "Item Ledger Entry"."Remaining Quantity";
                    TotBarrels := TotBarrels + "NoofBarrels/Pails";

                    //RSPL Sharath>>>>>
                    IF "Item Ledger Entry"."Location Code" = 'PLANT01' THEN BEGIN
                        CLEAR(BinName);
                        recWarehouseEntry.RESET;
                        recWarehouseEntry.SETRANGE(recWarehouseEntry."Item No.", "Item Ledger Entry"."Item No.");
                        recWarehouseEntry.SETRANGE(recWarehouseEntry."Reference No.", "Item Ledger Entry"."Document No.");
                        IF recWarehouseEntry.FINDFIRST THEN BEGIN
                            recRWAL.RESET;
                            recRWAL.SETRANGE("Whse. Document No.", recWarehouseEntry."Whse. Document No.");
                            recRWAL.SETRANGE("Item No.", recWarehouseEntry."Item No.");
                            recRWAL.SETRANGE(recRWAL."Action Type", recRWAL."Action Type"::Place);
                            IF recRWAL.FINDFIRST THEN BEGIN
                                REPEAT
                                    recBin.RESET;
                                    recBin.SETRANGE(recBin.Code, recRWAL."Bin Code");
                                    IF recBin.FINDFIRST THEN
                                        BinName := recBin.Description;
                                UNTIL recRWAL.NEXT = 0;
                            END;
                        END;
                    END ELSE BEGIN
                        //RSPL Sharath<<<<<
                        CLEAR(BinName);
                        recWarehouseEntry.RESET;
                        recWarehouseEntry.SETRANGE(recWarehouseEntry."Item No.", "Item Ledger Entry"."Item No.");
                        recWarehouseEntry.SETRANGE(recWarehouseEntry."Reference No.", "Item Ledger Entry"."Document No.");
                        recWarehouseEntry.SETRANGE(recWarehouseEntry."Source Line No.", "Item Ledger Entry"."Document Line No.");
                        IF recWarehouseEntry.FINDFIRST THEN BEGIN
                            recBin.RESET;
                            recBin.SETRANGE(recBin.Code, recWarehouseEntry."Bin Code");
                            IF recBin.FINDFIRST THEN
                                BinName := recBin.Description;
                        END;
                    END;

                    //RSPL Sharath>>>>
                    IF BinName = '' THEN BEGIN
                        IF "Item Ledger Entry"."Location Code" = 'PLANT01' THEN BEGIN
                            CLEAR(BinName);
                            recWarehouseEntry.RESET;
                            recWarehouseEntry.SETRANGE(recWarehouseEntry."Item No.", "Item Ledger Entry"."Item No.");
                            recWarehouseEntry.SETRANGE(recWarehouseEntry."Reference No.", "Item Ledger Entry"."Document No.");
                            IF recWarehouseEntry.FINDFIRST THEN BEGIN
                                recBin.RESET;
                                recBin.SETRANGE(recBin.Code, recWarehouseEntry."Bin Code");
                                IF recBin.FINDFIRST THEN
                                    BinName := recBin.Description;
                            END;
                        END;
                    END;

                    /*
                    IF recTRH.GET("Item Ledger Entry"."Document No.") THEN;
                    TransferShipmentHeader.RESET;
                    TransferShipmentHeader.SETRANGE(TransferShipmentHeader."Transfer Order No.",recTRH."Transfer Order No.");
                    IF TransferShipmentHeader.FIND('-') THEN
                     //BRTNO := TransferShipmentHeader."No."
                     BRTNO := "Document No."
                    ELSE
                     BRTNO := '';
                    //RSPL Sharath<<<<<
                    */

                    //>>08Sep2017 RB-N
                    CLEAR(IndentNo);
                    CLEAR(TrFCode);
                    CLEAR(TrFName);
                    CLEAR(TrTCode);
                    CLEAR(TrTName);
                    recTRH.RESET;
                    recTRH.SETRANGE("No.", "Document No.");
                    IF recTRH.FINDFIRST THEN BEGIN
                        TransferShipmentHeader.RESET;
                        TransferShipmentHeader.SETRANGE("Transfer Order No.", recTRH."Transfer Order No.");
                        IF TransferShipmentHeader.FINDFIRST THEN BEGIN
                            BRTNO := TransferShipmentHeader."No.";
                            IndentNo := TransferShipmentHeader."Transfer Indent No.";
                            TrFCode := TransferShipmentHeader."Transfer-from Code";
                            TrFName := TransferShipmentHeader."Transfer-from Name";
                            TrTCode := TransferShipmentHeader."Transfer-to Code";
                            TrTName := TransferShipmentHeader."Transfer-to Name";
                        END ELSE BEGIN
                            BRTNO := '';
                            IndentNo := '';
                            TrFCode := '';
                            TrFName := '';
                            TrTCode := '';
                            TrTName := '';
                        END;
                    END;

                    //Indent Details
                    CLEAR(AppTime);
                    CLEAR(AppDate);
                    CLEAR(ApproverName);
                    IF IndentNo <> '' THEN BEGIN

                        TIH.RESET;
                        TIH.SETRANGE("No.", IndentNo);
                        IF TIH.FINDFIRST THEN BEGIN
                            AppDate := TIH."Approval Date";
                            AppTime := TIH."Approval Time";

                            User08.RESET;
                            User08.SETRANGE("User ID", TIH."Approve User ID");
                            IF User08.FINDFIRST THEN
                                ApproverName := User08.Name;
                        END;

                    END;

                    IF ("Location Code" = 'PLANT01') OR ("Location Code" = 'PLANT02') OR ("Location Code" = 'PLANT03') THEN
                        BRTNO := "Document No.";
                    //>>08Sep2017 RB-N

                    //RSPLSUM 27Jul2020>>
                    CLEAR(MRP);
                    CLEAR(ListPrice);
                    RecMRPMaster.RESET;
                    RecMRPMaster.SETCURRENTKEY("Posting Date", "Item No.", "Lot No.");
                    RecMRPMaster.ASCENDING(FALSE);
                    RecMRPMaster.SETRANGE("Item No.", "Item Ledger Entry"."Item No.");
                    RecMRPMaster.SETRANGE("Lot No.", "Item Ledger Entry"."Lot No.");
                    IF RecMRPMaster.FINDFIRST THEN BEGIN
                        //RSPLSUM 29Jul2020--MRP := RecMRPMaster."MRP Price";
                        MRP := RecMRPMaster."MRP Price" * RecMRPMaster."Qty. per Unit of Measure";//RSPLSUM 29Jul2020
                        ListPrice := RecMRPMaster."Sales price" / RecMRPMaster."Qty. per Unit of Measure";
                    END ELSE BEGIN
                        RecSalesPrice.RESET;
                        RecSalesPrice.SETCURRENTKEY("Starting Date", "Item No.");
                        RecSalesPrice.ASCENDING(FALSE);
                        RecSalesPrice.SETRANGE("Item No.", "Item Ledger Entry"."Item No.");
                        RecSalesPrice.SETRANGE("Ending Date", 0D);
                        IF RecSalesPrice.FINDFIRST THEN BEGIN
                            MRP := 0;
                            ListPrice := RecSalesPrice."Basic Price";
                        END;
                    END;

                    CLEAR(RegionCode);
                    RecLocNew.RESET;
                    IF RecLocNew.GET("Item Ledger Entry"."Location Code") THEN BEGIN
                        //RSPLSUM 29Jul2020>>
                        IF "Item Ledger Entry"."Location Code" = 'IN-TRANS' THEN BEGIN
                            IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Transfer Receipt" THEN BEGIN
                                RecTranRecHdr.RESET;
                                IF RecTranRecHdr.GET("Item Ledger Entry"."Document No.") THEN BEGIN
                                    RecLocationNew.RESET;
                                    IF RecLocationNew.GET(RecTranRecHdr."Transfer-to Code") THEN BEGIN
                                        RecState.RESET;
                                        IF RecState.GET(RecLocationNew."State Code") THEN
                                            RegionCode := '';// RecState."Region Code";
                                    END;
                                END;
                            END;
                            IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Transfer Shipment" THEN BEGIN
                                RecTranShptHdr.RESET;
                                IF RecTranShptHdr.GET("Item Ledger Entry"."Document No.") THEN BEGIN
                                    RecLocationNew.RESET;
                                    IF RecLocationNew.GET(RecTranShptHdr."Transfer-to Code") THEN BEGIN
                                        RecState.RESET;
                                        IF RecState.GET(RecLocationNew."State Code") THEN
                                            RegionCode := '';// RecState."Region Code";
                                    END;
                                END;
                            END;
                        END ELSE BEGIN//RSPLSUM 29Jul2020<<
                            RecState.RESET;
                            IF RecState.GET(RecLocNew."State Code") THEN
                                RegionCode := '';// RecState."Region Code";
                        END;//RSPLSUM 29Jul2020
                    END;
                    //RSPLSUM 27Jul2020<<

                    //RSPl Sharath>>>>>
                    IF "Item Ledger Entry"."Location Code" = 'IN-TRANS' THEN BEGIN
                        recTSH.RESET;
                        recTSH.SETRANGE("No.", "Item Ledger Entry"."Document No.");
                        IF recTSH.FINDFIRST THEN
                            recLoc.GET(recTSH."Transfer-to Code");
                    END ELSE
                        //RSPL Sharath<<<<<
                        recLoc.GET("Item Ledger Entry"."Location Code");

                    /*
                    IF Loc.GET("Item Ledger Entry"."Location Code") THEN BEGIN
                      IF Loc."Trading Location"=TRUE THEN BEGIN
                        RG23D.RESET;
                        RG23D.SETRANGE(RG23D."Lot No.","Item Ledger Entry"."Lot No.");
                        RG23D.SETRANGE(RG23D."Location Code","Item Ledger Entry"."Location Code");
                        RG23D.SETRANGE(RG23D."Transaction Type",RG23D."Transaction Type"::Purchase);
                        IF RG23D.FINDSET THEN
                          Value:=RG23D."Transfer Price"*"Item Ledger Entry"."Remaining Quantity";
                     END ELSE IF Location."Trading Location"=FALSE THEN BEGIN
                        SalesPrice.RESET;
                        SalesPrice.SETRANGE(SalesPrice."Item No.","Item Ledger Entry"."Item No.");
                        IF SalesPrice.FINDSET THEN //BEGIN
                          Value:=SalesPrice."Transfer Price"*"Item Ledger Entry"."Remaining Quantity";
                     END;
                    END;
                    */

                    //12Sep2017 RB-N NewCode for ValueCode
                    CLEAR(Value);
                    CLEAR(TPrice);
                    IF BRTNO <> '' THEN BEGIN
                        TSL12.RESET;
                        TSL12.SETRANGE("Document No.", BRTNO);
                        TSL12.SETRANGE("Line No.", "Item Ledger Entry"."Document Line No.");
                        TSL12.SETRANGE("Item No.", "Item Ledger Entry"."Item No.");
                        IF TSL12.FINDFIRST THEN BEGIN
                            TPrice := TSL12."Unit Price" / "Qty. per Unit of Measure";
                            Value := TSL12."Unit Price" * ("Remaining Quantity" / "Qty. per Unit of Measure");
                            //Value  := TSL12.Amount;
                        END;
                    END;


                    IF BRTNO = '' THEN BEGIN
                        TempBRTNO := '';
                        RRH12.RESET;
                        RRH12.SETRANGE("No.", "Item Ledger Entry"."Document No.");
                        IF RRH12.FINDFIRST THEN BEGIN
                            TempBRTNO := RRH12."Return Order No.";
                        END;

                        IF TempBRTNO <> '' THEN BEGIN
                            SCH12.RESET;
                            SCH12.SETRANGE("Return Order No.", TempBRTNO);
                            IF SCH12.FINDFIRST THEN BEGIN
                                BRTNO := SCH12."No.";
                            END;
                        END;

                        SCL12.RESET;
                        SCL12.SETRANGE("Document No.", BRTNO);
                        SCL12.SETRANGE("Line No.", "Item Ledger Entry"."Document Line No.");
                        SCL12.SETRANGE("No.", "Item Ledger Entry"."Item No.");
                        SCL12.SETRANGE("Location Code", "Location Code");
                        IF SCL12.FINDFIRST THEN BEGIN
                            TPrice := SCL12."Unit Price" / "Qty. per Unit of Measure";
                            Value := SCL12."Unit Price" * ("Remaining Quantity" / "Qty. per Unit of Measure");
                            //Value := SCL12.Amount;
                        END;
                    END;
                    //12Sep2017 RB-N NewCode for ValueCode

                    //>>10Aug2018
                    IF TPrice = 0 THEN BEGIN
                        CALCFIELDS("Cost Amount (Expected)", "Cost Amount (Actual)");
                        TPrice := ("Cost Amount (Expected)" + "Cost Amount (Actual)") / Quantity;
                        Value := TPrice * "Remaining Quantity";

                    END;
                    //<<10Aug2018
                    //<<1

                    //>>06Oct2017

                    LocName := '';
                    recLocation.RESET;
                    IF recLocation.GET("Location Code") THEN
                        LocName := recLocation.Name
                    ELSE
                        LocName := '';

                    //>>01Nov2017 LocationName for In-Trans
                    IF "Location Code" = 'IN-TRANS' THEN BEGIN
                        recTSH.RESET;
                        recTSH.SETRANGE("No.", "Document No.");
                        IF recTSH.FINDFIRST THEN BEGIN
                            recLocation.RESET;
                            IF recLocation.GET(recTSH."Transfer-to Code") THEN
                                LocName := recLocation.Name;
                        END;
                    END;

                    //<<


                    IF "Document Type" = "Document Type"::"Transfer Receipt" THEN
                        IF IndentNo = '' THEN BEGIN
                            IndentNo := "Item Ledger Entry"."Order No.";

                            TransferShipmentHeader.RESET;
                            TransferShipmentHeader.SETRANGE("Transfer Order No.", "Item Ledger Entry"."Order No.");
                            IF TransferShipmentHeader.FINDFIRST THEN
                                BRTNO := TransferShipmentHeader."No.";

                        END;

                    //<<06Oct2017


                    //>>2


                    //Item Ledger Entry, GroupFooter (1) - OnPreSection()
                    Item.GET("Item Ledger Entry"."Item No.");

                    IF NOT Item.Blocked THEN BEGIN

                        IF PrevItemNo <> "Item Ledger Entry"."Item No." THEN
                            Srno := 0;

                        PrevItemNo := "Item Ledger Entry"."Item No.";

                        //CurrReport.SHOWOUTPUT:=CurrReport.TOTALSCAUSEDBY=FIELDNO("Item Ledger Entry"."Lot No.");
                        //CurrReport.SHOWOUTPUT:=CurrReport.TOTALSCAUSEDBY=FIELDNO("Item Ledger Entry"."Posting Date");
                        //IF CurrReport.TOTALSCAUSEDBY=FIELDNO("Item Ledger Entry"."Lot No.") THEN BEGIN
                        //IF CurrReport.TOTALSCAUSEDBY=FIELDNO("Item Ledger Entry"."Posting Date") THEN BEGIN //16Mar2017

                        FOR i := 1 TO 5 DO BEGIN
                            RemQty[i] := 0;
                            OpenQty[i] := 0;
                            TrQty[i] := 0;
                            RemValue[i] := 0;//27Aug2018
                        END;

                        /*
                        vMfgDate:=0D;
                        vRcptDate:=0D;
                        IF recLoc."Trading Location"THEN
                        BEGIN
                          recILE.RESET;
                          recILE.SETRANGE("Item No.","Item No.");
                          recILE.SETRANGE(recILE."Lot No.","Lot No.");
                          recILE.SETRANGE(recILE."Entry Type",recILE."Entry Type"::Output);
                          IF recILE.FINDFIRST THEN
                            vMfgDate:=recILE."Posting Date";
                            vRcptDate:="Posting Date";
                        END ELSE
                        BEGIN
                          IF "Entry Type"="Entry Type"::Output THEN
                          vMfgDate:="Posting Date";
                          vRcptDate:=0D;
                          recILE.RESET;
                          recILE.SETRANGE(recILE."Item No.","Item No.");
                          recILE.SETFILTER(recILE."Entry Type",'<>%1',recILE."Entry Type"::Output);
                          recILE.SETFILTER(recILE."Remaining Quantity",'<>%1',0);
                          recILE.SETRANGE(recILE."Lot No.","Lot No.");
                          recILE.SETRANGE("Location Code","Location Code");
                          IF recILE.FINDFIRST THEN
                            vRcptDate:=recILE."Posting Date";

                        END;
                        *///Commented 08Sep2017

                        //>>RB-N NewCode 08Sep2017
                        vMfgDate := 0D;
                        vRcptDate := 0D;
                        IF ("Location Code" <> 'PLANT01') OR ("Location Code" <> 'PLANT02') OR ("Location Code" <> 'PLANT03') THEN BEGIN
                            recILE.RESET;
                            recILE.SETRANGE("Item No.", "Item No.");
                            recILE.SETRANGE(recILE."Lot No.", "Lot No.");
                            recILE.SETRANGE(recILE."Entry Type", recILE."Entry Type"::Output);
                            IF recILE.FINDFIRST THEN
                                vMfgDate := recILE."Posting Date";
                            vRcptDate := "Posting Date";
                        END ELSE BEGIN
                            IF "Entry Type" = "Entry Type"::Output THEN
                                vMfgDate := "Posting Date";
                            vRcptDate := 0D;
                            recILE.RESET;
                            recILE.SETRANGE(recILE."Item No.", "Item No.");
                            recILE.SETFILTER(recILE."Entry Type", '<>%1', recILE."Entry Type"::Output);
                            recILE.SETFILTER(recILE."Remaining Quantity", '<>%1', 0);
                            recILE.SETRANGE(recILE."Lot No.", "Lot No.");
                            recILE.SETRANGE("Location Code", "Location Code");
                            IF recILE.FINDFIRST THEN
                                vRcptDate := recILE."Posting Date";
                        END;
                        //<<RB-N NewCode 08Sep2017

                        //>>22May2018
                        IF vMfgDate = 0D THEN
                            vMfgDate := vRcptDate;
                        //<<22May2018


                        //IF recLoc."Trading Location"THEN
                        IF ("Location Code" <> 'PLANT01') OR ("Location Code" <> 'PLANT02') OR ("Location Code" <> 'PLANT03') THEN //RB-N NewCode 08Sep2017
                        BEGIN
                            IF isAgeByMfg = FALSE THEN BEGIN
                                //IF (NoOfDays>=0) OR (NoOfDays<30) THEN BEGIN
                                IF ((TODAY - "Item Ledger Entry"."Posting Date") >= 0) AND ((TODAY - "Item Ledger Entry"."Posting Date") <= NoOfDays) THEN BEGIN
                                    //    RemQty[1]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"  //ORG
                                    RemQty[1] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[1] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=30) AND (NoOfDays<60) THEN BEGIN
                                IF ((TODAY - "Posting Date") >= NoOfDays + 1) AND ((TODAY - "Posting Date") <= (NoOfDays1)) THEN BEGIN
                                    // RemQty[2]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"     //ORG
                                    RemQty[2] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[2] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=60) AND (NoOfDays<90) THEN BEGIN
                                IF ((TODAY - "Posting Date") >= (NoOfDays1 + 1)) AND ((TODAY - "Posting Date") <= (NoOfDays2)) THEN BEGIN
                                    // RemQty[3]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"    //ORG
                                    RemQty[3] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[3] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=90) AND (NoOfDays<120) THEN BEGIN
                                IF ((TODAY - "Posting Date") >= (NoOfDays2 + 1)) AND ((TODAY - "Posting Date") <= (NoOfDays3)) THEN BEGIN
                                    //  RemQty[4]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"    //ORG
                                    RemQty[4] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[4] += Value;//27Aug2018
                                                         //*"Item Ledger Entry"."Qty. per Unit of Measure"
                                END;

                                //IF (NoOfDays>=120) THEN BEGIN
                                IF ((TODAY - "Item Ledger Entry"."Posting Date") > (NoOfDays3)) THEN BEGIN
                                    //   RemQty[5]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"  //ORG
                                    RemQty[5] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[5] += Value;//27Aug2018
                                END;
                            END;

                            IF (isAgeByMfg) AND (vMfgDate <> 0D) THEN BEGIN
                                //IF (NoOfDays>=0) OR (NoOfDays<30) THEN BEGIN
                                IF ((TODAY - vMfgDate) >= 0) AND ((TODAY - vMfgDate) <= NoOfDays) THEN BEGIN
                                    //    RemQty[1]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"  //ORG
                                    RemQty[1] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[1] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=30) AND (NoOfDays<60) THEN BEGIN
                                IF ((TODAY - vMfgDate) >= NoOfDays + 1) AND ((TODAY - vMfgDate) <= (NoOfDays1)) THEN BEGIN
                                    // RemQty[2]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"     //ORG
                                    RemQty[2] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[2] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=60) AND (NoOfDays<90) THEN BEGIN
                                IF ((TODAY - vMfgDate) >= (NoOfDays1 + 1)) AND ((TODAY - vMfgDate) <= (NoOfDays2)) THEN BEGIN
                                    // RemQty[3]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"    //ORG
                                    RemQty[3] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[3] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=90) AND (NoOfDays<120) THEN BEGIN
                                IF ((TODAY - vMfgDate) >= (NoOfDays2 + 1)) AND ((TODAY - vMfgDate) <= (NoOfDays3)) THEN BEGIN
                                    //  RemQty[4]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"    //ORG
                                    RemQty[4] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[4] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=120) THEN BEGIN
                                IF ((TODAY - vMfgDate) > (NoOfDays3)) THEN BEGIN
                                    //   RemQty[5]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"  //ORG
                                    RemQty[5] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[5] += Value;//27Aug2018
                                END;
                            END;

                        END ELSE BEGIN
                            IF vRcptDate = 0D THEN BEGIN
                                //IF (NoOfDays>=0) OR (NoOfDays<30) THEN BEGIN
                                IF ((TODAY - "Item Ledger Entry"."Posting Date") >= 0) AND ((TODAY - "Item Ledger Entry"."Posting Date") <= NoOfDays) THEN BEGIN
                                    //    RemQty[1]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"  //ORG
                                    RemQty[1] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[1] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=30) AND (NoOfDays<60) THEN BEGIN
                                IF ((TODAY - "Posting Date") >= NoOfDays + 1) AND ((TODAY - "Posting Date") <= (NoOfDays1)) THEN BEGIN
                                    // RemQty[2]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"     //ORG
                                    RemQty[2] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[2] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=60) AND (NoOfDays<90) THEN BEGIN
                                IF ((TODAY - "Posting Date") >= (NoOfDays1 + 1)) AND ((TODAY - "Posting Date") <= (NoOfDays2)) THEN BEGIN
                                    // RemQty[3]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"    //ORG
                                    RemQty[3] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[3] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=90) AND (NoOfDays<120) THEN BEGIN
                                IF ((TODAY - "Posting Date") >= (NoOfDays2 + 1)) AND ((TODAY - "Posting Date") <= (NoOfDays3)) THEN BEGIN
                                    //  RemQty[4]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"    //ORG
                                    RemQty[4] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[4] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=120) THEN BEGIN
                                IF ((TODAY - "Item Ledger Entry"."Posting Date") > (NoOfDays3)) THEN BEGIN
                                    //   RemQty[5]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"  //ORG
                                    RemQty[5] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[5] += Value;//27Aug2018
                                END;
                            END;

                            IF vRcptDate <> 0D THEN BEGIN
                                //IF (NoOfDays>=0) OR (NoOfDays<30) THEN BEGIN
                                IF ((TODAY - vRcptDate) >= 0) AND ((TODAY - vRcptDate) <= NoOfDays) THEN BEGIN
                                    //    RemQty[1]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"  //ORG
                                    RemQty[1] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[1] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=30) AND (NoOfDays<60) THEN BEGIN
                                IF ((TODAY - vRcptDate) >= NoOfDays + 1) AND ((TODAY - vRcptDate) <= (NoOfDays1)) THEN BEGIN
                                    // RemQty[2]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"     //ORG
                                    RemQty[2] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[2] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=60) AND (NoOfDays<90) THEN BEGIN
                                IF ((TODAY - vRcptDate) >= (NoOfDays1 + 1)) AND ((TODAY - vRcptDate) <= (NoOfDays2)) THEN BEGIN
                                    // RemQty[3]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"    //ORG
                                    RemQty[3] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[3] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=90) AND (NoOfDays<120) THEN BEGIN
                                //IF ((TODAY-vRcptDate)>=(NoOfDays2+1)) AND ((TODAY-vRcptDate)<(NoOfDays3)) THEN
                                IF ((TODAY - vRcptDate) >= (NoOfDays2 + 1)) AND ((TODAY - vRcptDate) <= (NoOfDays3)) THEN BEGIN
                                    //  RemQty[4]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"    //ORG
                                    RemQty[4] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[4] += Value;//27Aug2018
                                END;

                                //IF (NoOfDays>=120) THEN BEGIN
                                IF ((TODAY - vRcptDate) > (NoOfDays3)) THEN BEGIN
                                    //   RemQty[5]+="Item Ledger Entry"."Remaining Quantity"*Item."Packing Qty per UOM"  //ORG
                                    RemQty[5] += "Item Ledger Entry"."Remaining Quantity";
                                    RemValue[5] += Value;//27Aug2018
                                END;
                            END;

                        END;

                        Srno := Srno + 1;
                        i += 1;

                        IF PrintToExcel THEN
                            MakeExcelDataGroupFooter1;
                        Makefooter := TRUE;

                        //END;//16Mar2017
                    END;

                    //<<2


                    //>>3

                    //Item Ledger Entry, GroupFooter (2) - OnPreSection()
                    //CurrReport.SHOWOUTPUT:=CurrReport.TOTALSCAUSEDBY=FIELDNO("Item Ledger Entry"."Item No.");

                    IF PrevItemNo <> "Item Ledger Entry"."Item No." THEN
                        Srno := 0;
                    /*
                    IF Srno<>0 THEN
                    BEGIN
                      IF CurrReport.TOTALSCAUSEDBY=FIELDNO("Item Ledger Entry"."Item No.") THEN
                      BEGIN
                       IF  PrintToExcel THEN
                       BEGIN
                        i += 1;
                        MakeExcelDataGroupFooter2;
                       END;
                      END;
                    END;
                    
                    //<<3
                    *///Commented GroupFooter 16Mar2017



                    //>>17Mar2017 GroupFooterNew
                    LCOUNT += 1;

                    ILE17Mar.SETCURRENTKEY("Item No.", "Location Code", "Lot No.", "Variant Code", "Posting Date");
                    ILE17Mar.SETRANGE("Item No.", "Item No.");
                    ILE17Mar.SETRANGE("Location Code", "Location Code");
                    IF ILE17Mar.FINDFIRST THEN BEGIN
                        TCOUNT := ILE17Mar.COUNT;
                    END;

                    IF TCOUNT <> 0 THEN
                        IF PrintToExcel AND (TCOUNT = LCOUNT) THEN BEGIN
                            MakeExcelDataGroupFooter2;
                            LCOUNT := 0;

                        END;

                    //<<17Mar2017 GroupFooterNew

                    //>>4

                    GT08 += 1; //08Sep2017 RB-N

                end;

                trigger OnPostDataItem()
                begin

                    //Item Ledger Entry, Footer (3) - OnPreSection()
                    IF GT08 <> 0 THEN
                        IF PrintToExcel THEN BEGIN
                            MakeExcelFooter;
                            i += 1;
                        END;
                    //<<4
                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    i := 5;

                    CurrReport.CREATETOTALS("NoofBarrels/Pails");
                    //<<1

                    ILE17Mar.COPYFILTERS("Item Ledger Entry");
                end;
            }

            trigger OnAfterGetRecord()
            begin
                GT08 := 0; //08Sep2017

                //>>1

                //UserSetup.GET(USERID);
                /*
                IF Location."Trading Location"=TRUE THEN BEGIN
                  ILE.RESET;
                  ILE.SETRANGE(ILE."Location Code",Location.Code);
                  ILE.SETFILTER(ILE."Lot No.",'<>%1','REV');
                  ILE.SETFILTER(ILE."Item Category Code",'%1|%2','CAT02','CAT03');
                  ILE.SETFILTER(ILE."Item Category Code",'%1|%2','CAT1','CAT12');
                  ILE.SETFILTER(ILE."Item Category Code",'%1|%2','CAT13','CAT15');
                  ILE.SETFILTER(ILE.Quantity,'<>%1',0);
                  IF ILE.FINDFIRST THEN REPEAT
                    RG23D.RESET;
                    RG23D.SETRANGE(RG23D."Lot No.",ILE."Lot No.");
                    RG23D.SETRANGE(RG23D."Location Code",ILE."Location Code");
                    RG23D.SETRANGE(RG23D."Transaction Type",RG23D."Transaction Type"::Purchase);
                    IF RG23D.FINDSET THEN
                      //RG23D.CALCFIELDS(RG23D."Transfer Price");
                      Value:=RG23D."Transfer Price"*ILE."Remaining Quantity";
                  UNTIL ILE.NEXT=0;
                 END ELSE IF Location."Trading Location"=FALSE THEN
                  ILE.RESET;
                  ILE.SETRANGE(ILE."Location Code",Location.Code);
                  ILE.SETFILTER(ILE."Lot No.",'<>%1','REV');
                  ILE.SETFILTER(ILE."Item Category Code",'%1|%2','CAT02','CAT03');
                  ILE.SETFILTER(ILE."Item Category Code",'%1|%2','CAT1','CAT12');
                  ILE.SETFILTER(ILE."Item Category Code",'%1|%2','CAT13','CAT15');
                  ILE.SETFILTER(ILE.Quantity,'<>%1',0);
                  IF ILE.FINDFIRST THEN REPEAT
                    SalesPrice.RESET;
                    SalesPrice.SETRANGE(SalesPrice."Item No.",ILE."Item No.");
                    IF SalesPrice.FINDSET THEN
                    //SalesPrice.CALCFIELDS(SalesPrice."Transfer Price");
                      Value:=SalesPrice."Transfer Price"*ILE."Remaining Quantity";
                  UNTIL ILE.NEXT=0;
                
                MESSAGE('%1',Value);
                 */
                //<<1

                //>>2

                //Location, Header (1) - OnPreSection()
                j := 0;
                //<<2

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Age By Mfg Date"; isAgeByMfg)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin
                        /*
                        IF isAgeByMfg AND isAgeByRcpt THEN
                        BEGIN
                          ERROR('Choose any of one option from Mfg Date and Rcpt Date');
                        
                        END;
                        */

                    end;
                }
                field("Age By Rcpt. Date"; isAgeByRcpt)
                {

                    trigger OnValidate()
                    begin
                        /*
                        IF isAgeByMfg AND isAgeByRcpt THEN
                        BEGIN
                          ERROR('Choose any of one option from Mfg Date and Rcpt Date');
                        
                        END;
                        */

                    end;
                }
                field("No. Of Days"; NoOfDays)
                {
                    ApplicationArea = all;
                }
                field("No. Of Days1"; NoOfDays1)
                {
                    ApplicationArea = all;
                }
                field("No. Of Days2"; NoOfDays2)
                {
                    ApplicationArea = all;
                }
                field("No. Of Days3"; NoOfDays3)
                {
                    ApplicationArea = all;
                }
                field("Print To Excel"; PrintToExcel)
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

        //>>1

        IF PrintToExcel THEN
            CreateExcelbook;

        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1

        LocCode := Location.GETFILTER(Location.Code);

        /*
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",UPPERCASE(USERID));
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.RESET;
         recLocation.SETRANGE(recLocation.Code,LocCode);
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping1.RESET;
         CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
         CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
         CSOMapping1.SETRANGE(CSOMapping1.Value,LocCode);
         IF NOT(CSOMapping1.FINDFIRST) THEN
         BEGIN
          CSOMapping1.RESET;
          CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
          CSOMapping1.SETRANGE(CSOMapping1.Type,CSOMapping1.Type::"Responsibility Center");
          CSOMapping1.SETRANGE(CSOMapping1.Value,LocResC);
          IF NOT(CSOMapping1.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        *///Commented 16Mar2017

        CompInfo.GET;

        IF PrintToExcel THEN
            MakeExcelInfo;

        //<<1

    end;

    var
        Srno: Integer;
        Item: Record 27;
        ItemNo: Code[20];
        Locationcode: Code[20];
        batchno: Code[10];
        Packsize: Decimal;
        Noofbarrels: Decimal;
        OpeningStock: Decimal;
        MfgDate: Date;
        "NoofBarrels/Pails": Decimal;
        "<To Export to Excel>": Integer;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        LocationFilter: Text[50];
        ILEFilter: Text[50];
        n: Text[30];
        MaxCol: Text[3];
        i: Integer;
        DateLength: Integer;
        Date1: Text[30];
        MonthLength: Integer;
        Month1: Text[30];
        YearLength: Integer;
        Year1: Text[30];
        PostingDate: Text[30];
        ItemUOM: Record 5404;
        ItemUOM1: Record 5404;
        UserSetup: Record 91;
        UserRespCtr: Code[10];
        CompInfo: Record 79;
        TotBarrels: Decimal;
        ItemCatCode: Code[30];
        j: Integer;
        PrevItemNo: Text[50];
        RemQty: array[10] of Decimal;
        ILERec: Record 32;
        NoOfDays: Integer;
        Makefooter: Boolean;
        NoOfDays1: Integer;
        NoOfDays2: Integer;
        NoOfDays3: Integer;
        recWarehouseEntry: Record 7312;
        recBin: Record 7354;
        BinName: Text[50];
        LocResC: Code[30];
        recLocation: Record 14;
        LocCode: Text;
        CSOMapping1: Record 50006;
        recLoc: Record 14;
        AgeingDate: Date;
        recILE: Record 32;
        OpenQty: array[5] of Decimal;
        TrQty: array[5] of Decimal;
        vMfgDate: Date;
        vRcptDate: Date;
        isAgeByMfg: Boolean;
        isAgeByRcpt: Boolean;
        vPostingDate: Date;
        remark: Text[30];
        LocationName: Text[30];
        //RG23D: Record 16537;
        Value: Decimal;
        SalesPrice: Record 7002;
        Loc: Record 14;
        ILE: Record 32;
        Text16500: Label 'As per Details';
        Text000: Label 'Data';
        Text001: Label 'Finished Goods- Ageing';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "-----17Mar2017": Integer;
        TCOUNT: Integer;
        LCOUNT: Integer;
        ILE17Mar: Record 32;
        GRmQty: Decimal;
        GNoPack: Decimal;
        TNoPack: Decimal;
        "-------06Sep17": Integer;
        recTSH: Record 5744;
        recRWAL: Record 5773;
        recTRH: Record 5746;
        TransferShipmentHeader: Record 5744;
        BRTNO: Code[30];
        "---08Sep17": Integer;
        TIH: Record 50022;
        AppDate: Date;
        AppTime: Time;
        ApproverName: Text[100];
        IndentNo: Code[30];
        User08: Record 91;
        GT08: Integer;
        "----12Sep2017": Integer;
        TSL12: Record 5745;
        RRH12: Record 6660;
        SCL12: Record 115;
        TempBRTNO: Code[20];
        SCH12: Record 114;
        TPrice: Decimal;
        "---06Oct2017": Integer;
        LocName: Text[100];
        RemValue: array[10] of Decimal;
        ItmCat: Record 5722;
        ItmCatName: Text;
        RepItmNo: Code[10];
        RecMRPMaster: Record 50013;
        MRP: Decimal;
        ListPrice: Decimal;
        RecSalesPrice: Record 7002;
        RecLocNew: Record 14;
        RecState: Record State;
        RegionCode: Code[10];
        RecTranRecHdr: Record 5746;
        RecLocationNew: Record 14;
        RecTranShptHdr: Record 5744;
        TrFCode: Code[50];
        TrFName: Code[50];
        TrTCode: Code[50];
        TrTName: Code[50];

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//16Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Finished Goods-Ageing', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Finished Goods - Ageing-N", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>16Mar2017 RB-N

        /*  ExcelBuf.CreateBook('', Text001);
         ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
         ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>17Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        //RSPLSUM 15dec2020--ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19                 //RSPL Sharath
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20 12Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30 31Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33 RSPLSUM 27Jul2020

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//35 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38 Fahim 24-Jan-22

        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn(Text001, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        //RSPLSUM 15dec2020--ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,FALSE,'@',1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19                 //RSPL Sharath
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20 12Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30 31Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33 RSPLSUM 27Jul2020

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//35 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38 Fahim 24-Jan-22

        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.NewRow;

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        //RSPLSUM 15dec2020--ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19            //RSPL Sharath
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25 12Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 31Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 RSPLSUM 27Jul2020

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//35 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//36 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//37 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//38 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//39 Fahim 24-Jan-22

        //<<17Mar2017

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Sr.No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '@', 1);// 31Aug2018
        ExcelBuf.AddColumn('BRT No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4    //RSPL Sharath 06-Sep-2017
        ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        //RSPLSUM 15dec2020--ExcelBuf.AddColumn('BIN',FALSE,'',TRUE,FALSE,TRUE,'@',1);//8
        ExcelBuf.AddColumn('Pack Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Pack Size', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('No.Of Packs', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Remaining Qty in Ltrs/KG', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Mfg.Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Receipt Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('0-' + FORMAT(NoOfDays), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('0-' + FORMAT(NoOfDays) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);// 27Aug2018 RB-N

        ExcelBuf.AddColumn(FORMAT(NoOfDays + 1) + '-' + FORMAT(NoOfDays1), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn(FORMAT(NoOfDays + 1) + '-' + FORMAT(NoOfDays1) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);// 27Aug2018 RB-N

        ExcelBuf.AddColumn(FORMAT(NoOfDays1 + 1) + '-' + FORMAT(NoOfDays2), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn(FORMAT(NoOfDays1 + 1) + '-' + FORMAT(NoOfDays2) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);// 27Aug2018 RB-N

        ExcelBuf.AddColumn(FORMAT(NoOfDays2 + 1) + '-' + FORMAT(NoOfDays3), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn(FORMAT(NoOfDays2 + 1) + '-' + FORMAT(NoOfDays3) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);// 27Aug2018 RB-N

        ExcelBuf.AddColumn('>' + FORMAT(NoOfDays3), FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('>' + FORMAT(NoOfDays3) + ' Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);// 27Aug2018 RB-N

        ExcelBuf.AddColumn('Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('Transfer Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21 12Sep2017 RB-N
        ExcelBuf.AddColumn('Indent No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 08Sep2017 RB-N
        ExcelBuf.AddColumn('Approval Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 08Sep2017 RB-N
        ExcelBuf.AddColumn('Approval Time', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24 08Sep2017 RB-N
        ExcelBuf.AddColumn('Approver Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25 08Sep2017 RB-N
        ExcelBuf.AddColumn('MRP', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('List Price', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('Region', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 RSPLSUM 27Jul2020

        ExcelBuf.AddColumn('Transfer-from Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 Fahim 20-Jan-22
        ExcelBuf.AddColumn('Transfer-from Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 Fahim 20-Jan-22
        ExcelBuf.AddColumn('Transfer-to Code ', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 Fahim 20-Jan-22
        ExcelBuf.AddColumn('Transfer-to Name ', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32 Fahim 20-Jan-22
        ExcelBuf.AddColumn('BIN ', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33 Fahim 24-Jan-22
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataGroupFooter1()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Srno, FALSE, '', FALSE, FALSE, FALSE, '', 0);//1 16Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2 16Mar2017
        ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3 16Mar2017
        //>>31Aug2018
        ItmCatName := '';
        ItmCat.RESET;
        IF ItmCat.GET("Item Ledger Entry"."Item Category Code") THEN
            ItmCatName := ItmCat.Description;
        ExcelBuf.AddColumn(ItmCatName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //<<31Aug2018
        ExcelBuf.AddColumn(BRTNO, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4 //RSPL Sharath 06-Sep-2017
        CLEAR(BRTNO);//08Sep2017 RB-N

        ExcelBuf.AddColumn("Item Ledger Entry"."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5 16Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6 16Mar2017
        ExcelBuf.AddColumn(LocName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7 16Mar2017
        //RSPLSUM 15dec2020--ExcelBuf.AddColumn(BinName,FALSE,'',FALSE,FALSE,FALSE,'@',1);//8 16Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Unit of Measure Code", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9 16Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//10 16Mar2017
        ExcelBuf.AddColumn("Item Ledger Entry"."Remaining Quantity" /
        "Item Ledger Entry"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//11 16Mar2017
        //>>17Mar2017 GNoPack
        GNoPack += "Item Ledger Entry"."Remaining Quantity" / "Item Ledger Entry"."Qty. per Unit of Measure";
        TNoPack += "Item Ledger Entry"."Remaining Quantity" / "Item Ledger Entry"."Qty. per Unit of Measure";
        //<<17Mar2017 GNoPack

        ExcelBuf.AddColumn("NoofBarrels/Pails", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//12 16Mar2017
        //>>17Mar2017 GRmQty
        GRmQty += "NoofBarrels/Pails";
        //<<17Mar2017 GRmQty

        ExcelBuf.AddColumn(vMfgDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//13
        ExcelBuf.AddColumn(vRcptDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//14
        ExcelBuf.AddColumn(RemQty[1], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//15
        ExcelBuf.AddColumn(RemValue[1], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);// RB-N 27Aug2018

        ExcelBuf.AddColumn(RemQty[2], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//16
        ExcelBuf.AddColumn(RemValue[2], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);// RB-N 27Aug2018

        ExcelBuf.AddColumn(RemQty[3], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//17
        ExcelBuf.AddColumn(RemValue[3], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);// RB-N 27Aug2018

        ExcelBuf.AddColumn(RemQty[4], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//18
        ExcelBuf.AddColumn(RemValue[4], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);// RB-N 27Aug2018

        ExcelBuf.AddColumn(RemQty[5], FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//19
        ExcelBuf.AddColumn(RemValue[5], FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);// RB-N 27Aug2018

        ExcelBuf.AddColumn(Value, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//20
        ExcelBuf.AddColumn(TPrice, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//21 12Sep2017 RB-N
        ExcelBuf.AddColumn(IndentNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//22 08Sep2017 RB-N
        ExcelBuf.AddColumn(AppDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//23 08Sep2017 RB-N
        ExcelBuf.AddColumn(AppTime, FALSE, '', FALSE, FALSE, FALSE, 'hh:mm:ss AM/PM', 3);//24 08Sep2017 RB-N
        ExcelBuf.AddColumn(ApproverName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//25 08Sep2017 RB-N
        ExcelBuf.AddColumn(MRP, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn(ListPrice, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn(RegionCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//28 RSPLSUM 27Jul2020

        ExcelBuf.AddColumn(TrFCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//29 Fahim 20-Jan-22
        ExcelBuf.AddColumn(TrFName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//30 Fahim 20-Jan-22
        ExcelBuf.AddColumn(TrTCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//31 Fahim 20-Jan-22
        ExcelBuf.AddColumn(TrTName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//32 Fahim 20-Jan-22
        ExcelBuf.AddColumn(BinName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//33 Fahim 24-Jan-22
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataGroupFooter2()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);// 31Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        //RSPLSUM 15dec2020--ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',1);//7
        ExcelBuf.AddColumn('TOTAL', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//10 //08Sep2017
        ExcelBuf.AddColumn(GNoPack, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//11 //08Sep2017
        GNoPack := 0; //08Sep2017

        ExcelBuf.AddColumn(GRmQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//12 08Sep2017
        GRmQty := 0; //08Sep2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19 16Mar2017


        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20 06Sep2017          //RSPL Sharath
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25 12Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33 RSPLSUM 27Jul2020

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35 Fahim 24-Jan-22
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31Aug2018
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        //RSPLSUM 15dec2020--ExcelBuf.AddColumn('',FALSE,'',TRUE,FALSE,TRUE,'@',1);//7
        ExcelBuf.AddColumn('GRAND TOTAL', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//10 //08Sep2017
        ExcelBuf.AddColumn(TNoPack, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//11 /08Sep2017
        ExcelBuf.AddColumn(TotBarrels, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//12 08Sep2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18 16Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19 16Mar2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20 06Sep2017         //RSPL Sharath
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24 08Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25 12Sep2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 27Aug2018 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32 RSPLSUM 27Jul2020
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33 RSPLSUM 27Jul2020

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//34 Fahim 20-Jan-22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//35 Fahim 24-Jan-22
    end;
}


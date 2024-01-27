report 50002 "Daily Branch Transfer Register"
{
    DefaultLayout = RDLC;
    // RDLCLayout = './DailyBranchTransferRegister.rdlc';
    RDLCLayout = 'Src/Report Layout/DailyBranchTransferRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
        {
            DataItemTableView = SORTING("Posting Date")
                                ORDER(Ascending)
                                WHERE("No. Series" = FILTER('<> BRTCAN'),
                                      "BRT Cancelled" = CONST(false));
            RequestFilterFields = "Posting Date", "Transfer-from Code", "Transfer-to Code";
            dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Quantity = FILTER(<> 0));
                RequestFilterFields = "Item No.", "Shortcut Dimension 1 Code";

                trigger OnAfterGetRecord()
                begin
                    "Transfer Shipment Header".CALCFIELDS("Transfer Shipment Header"."EWB No.",//RSPLSUM 01Dec2020
                                  "Transfer Shipment Header"."EWB Date", "Transfer Shipment Header".IRN);//RSPLSUM 01Dec2020

                    CLEAR(PackedUOM);
                    retItem.RESET;
                    retItem.SETRANGE(retItem."No.", "Transfer Shipment Line"."Item No.");
                    IF retItem.FINDFIRST THEN BEGIN
                        PackedUOM := retItem."Sales Unit of Measure";
                    END;

                    IF "Transfer Shipment Line".Quantity = 0 THEN
                        CurrReport.SKIP;
                    //<<R
                    CLEAR(RcvdAgInvNo);
                    CLEAR(RcvdAgInvDate);
                    CLEAR(BatchNo);
                    reILE.RESET;
                    reILE.SETRANGE(reILE."Document No.", "Transfer Shipment Line"."Document No.");
                    reILE.SETRANGE(reILE."Location Code", "Transfer Shipment Line"."Transfer-from Code");
                    reILE.SETRANGE(reILE."Item No.", "Transfer Shipment Line"."Item No.");
                    reILE.SETRANGE(reILE."Document Line No.", "Transfer Shipment Line"."Line No.");
                    IF reILE.FINDFIRST THEN BEGIN
                        BatchNo := BatchNo + reILE."Lot No.";
                        //RSPLSUM 27Jun19>>
                        IF reILE.Positive THEN BEGIN
                            ItemApplnEntry.RESET;
                            ItemApplnEntry.SETCURRENTKEY("Inbound Item Entry No.", "Outbound Item Entry No.", "Cost Application");
                            ItemApplnEntry.SETRANGE("Inbound Item Entry No.", reILE."Entry No.");
                            ItemApplnEntry.SETFILTER("Outbound Item Entry No.", '<>%1', 0);
                            ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                            IF ItemApplnEntry.FINDFIRST THEN BEGIN
                                ILERec.RESET;
                                IF ILERec.GET(ItemApplnEntry."Outbound Item Entry No.") THEN BEGIN
                                    IF ILERec."Entry Type" = ILERec."Entry Type"::Transfer THEN BEGIN
                                        RecTrRcptHdr.RESET;
                                        IF RecTrRcptHdr.GET(ILERec."Document No.") THEN BEGIN
                                            TransferShipmentHeader.RESET;
                                            TransferShipmentHeader.SETRANGE(TransferShipmentHeader."Transfer Order No.", RecTrRcptHdr."Transfer Order No.");
                                            IF TransferShipmentHeader.FINDFIRST THEN
                                                BRTNO := TransferShipmentHeader."No."
                                            ELSE
                                                BRTNO := '';

                                            RcvdAgInvNo := BRTNO;
                                            RcvdAgInvDate := RecTrRcptHdr."Posting Date";
                                        END;
                                    END;
                                    IF ILERec."Entry Type" <> ILERec."Entry Type"::Transfer THEN BEGIN
                                        RcvdAgInvNo := ILERec."Document No.";
                                        RcvdAgInvDate := ILERec."Posting Date";
                                    END;
                                END;
                            END;
                        END ELSE BEGIN
                            ItemApplnEntry.RESET;
                            ItemApplnEntry.SETCURRENTKEY("Outbound Item Entry No.", "Item Ledger Entry No.", "Cost Application");
                            ItemApplnEntry.SETRANGE("Outbound Item Entry No.", reILE."Entry No.");
                            ItemApplnEntry.SETRANGE("Item Ledger Entry No.", reILE."Entry No.");
                            ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                            IF ItemApplnEntry.FINDFIRST THEN BEGIN
                                ILERec.RESET;
                                IF ILERec.GET(ItemApplnEntry."Inbound Item Entry No.") THEN BEGIN
                                    IF ILERec."Entry Type" = ILERec."Entry Type"::Transfer THEN BEGIN
                                        RecTrRcptHdr.RESET;
                                        IF RecTrRcptHdr.GET(ILERec."Document No.") THEN BEGIN
                                            TransferShipmentHeader.RESET;
                                            TransferShipmentHeader.SETRANGE(TransferShipmentHeader."Transfer Order No.", RecTrRcptHdr."Transfer Order No.");
                                            IF TransferShipmentHeader.FINDFIRST THEN
                                                BRTNO := TransferShipmentHeader."No."
                                            ELSE
                                                BRTNO := '';

                                            RcvdAgInvNo := BRTNO;
                                            RcvdAgInvDate := RecTrRcptHdr."Posting Date";
                                        END;
                                    END;
                                    IF ILERec."Entry Type" <> ILERec."Entry Type"::Transfer THEN BEGIN
                                        RcvdAgInvNo := ILERec."Document No.";
                                        RcvdAgInvDate := ILERec."Posting Date";
                                    END;
                                END;
                            END;
                        END;
                        //RSPLSUM 27Jun19<<
                    END;

                    CLEAR(TROno);
                    CLEAR(TRRno);
                    CLEAR(TSHno);
                    CLEAR(TSHdate);
                    // recRG23D.RESET;
                    // recRG23D.SETRANGE(recRG23D."Entry No.", "Transfer Shipment Line"."Applies-to Entry (RG 23 D)");
                    // IF recRG23D.FINDFIRST THEN BEGIN
                    //     TRRno := recRG23D."Document No.";

                    //     recTRH.RESET;
                    //     recTRH.SETRANGE(recTRH."No.", TRRno);
                    //     IF recTRH.FINDFIRST THEN BEGIN
                    //         TROno := recTRH."Transfer Order No.";
                    //     END;

                    //     recTSH.RESET;
                    //     recTSH.SETRANGE(recTSH."Transfer Order No.", TROno);
                    //     IF recTSH.FINDFIRST THEN BEGIN
                    //         TSHno := recTSH."No.";
                    //         TSHdate := recTSH."Posting Date";
                    //     END;

                    //     IF TSHno = '' THEN BEGIN
                    //         recRG23D.RESET;
                    //         recRG23D.SETRANGE(recRG23D."Entry No.", "Transfer Shipment Line"."Applies-to Entry (RG 23 D)");
                    //         IF recRG23D.FINDFIRST THEN BEGIN
                    //             TSHno := recRG23D."Document No.";
                    //             TSHdate := recRG23D."Posting Date";
                    //         END;
                    //     END;
                    // END;

                    TotalNetValue := TotalNetValue + "Transfer Shipment Line".Amount;
                    TotalBedAmt := TotalBedAmt + 0;// "Transfer Shipment Line"."BED Amount";
                    TotaleCessAmt := TotaleCessAmt + 0;//"Transfer Shipment Line"."eCess Amount";
                    TotalSHECess := TotalSHECess + 0;//"Transfer Shipment Line"."SHE Cess Amount";
                    TotAddDuty := TotAddDuty + 0;// "Transfer Shipment Line"."ADE Amount";
                    AMT := AMT1 + AMT2;
                    Grpfooterqty := Grpfooterqty + "Transfer Shipment Line".Quantity;
                    Grpamt := Grpamt + 0;// "Transfer Shipment Line".Amount;
                    Grpbed := Grpbed + 0;//"Transfer Shipment Line"."BED Amount";
                    GrpeCess := GrpeCess + 0;// "Transfer Shipment Line"."eCess Amount";
                    GrpSHE := GrpSHE + 0;// "Transfer Shipment Line"."SHE Cess Amount";
                    totalnoofpacks += "Transfer Shipment Line".Quantity;

                    Item.GET("Transfer Shipment Line"."Item No.");

                    //RSPL
                    vReceiptDate := 0D;
                    recTransRcptLine.RESET;
                    recTransRcptLine.SETRANGE(recTransRcptLine."Transfer Order No.", "Transfer Shipment Line"."Transfer Order No.");
                    recTransRcptLine.SETRANGE(recTransRcptLine."Line No.", "Transfer Shipment Line"."Line No.");
                    IF recTransRcptLine.FINDFIRST THEN
                        vReceiptDate := recTransRcptLine."Receipt Date";
                    //RSPL
                    //>>Rahul

                    IF RecTraShipHeader.GET("Transfer Shipment Line"."Document No.") THEN
                        RecItemCat.RESET;
                    RecItemCat.SETRANGE(RecItemCat.Code, "Transfer Shipment Line"."Item Category Code");
                    IF RecItemCat.FINDFIRST THEN
                        ItemCatName := RecItemCat.Description;

                    TotalBRTQty := TotalBRTQty + ("Transfer Shipment Line".Quantity * "Transfer Shipment Line"."Qty. per Unit of Measure");
                    ShipmentTotalQty := Quantity * "Transfer Shipment Line"."Qty. per Unit of Measure";
                    GroupQty += ShipmentTotalQty;
                    vTotNetValue += "Transfer Shipment Line".Amount;
                    vTotBEDAmt += 0;//"Transfer Shipment Line"."BED Amount";
                    vTotEcess += 0;//"Transfer Shipment Line"."eCess Amount";
                    vTotSheCess += 0;//"Transfer Shipment Line"."SHE Cess Amount";
                    vTotAddl += 0;// "Transfer Shipment Line"."ADE Amount";
                    vTotNoPacks += "Transfer Shipment Line".Quantity;
                    /*
                    AMT1:= ("Transfer Shipment Line".Amount+"Transfer Shipment Line"."BED Amount");
                            AMT2:= ("Transfer Shipment Line"."eCess Amount"+"Transfer Shipment Line"."SHE Cess Amount");
                    */
                    /*
                    IF NOT Detail THEN
                    CurrReport.SHOWOUTPUT(FALSE);
                    */

                    Grpfooterqty := 0;

                    //"Transfer Shipment Line".RESET;
                    //"Transfer Shipment Line".SETRANGE("Transfer Shipment Line"."Document No.","Transfer Shipment Header"."No.");
                    //"Transfer Shipment Line".SETRANGE("Item No.","Transfer Shipment Line"."Item No.");
                    RecItemCat.RESET;
                    RecItemCat.SETRANGE(RecItemCat.Code, "Transfer Shipment Line"."Item Category Code");
                    IF RecItemCat.FINDFIRST THEN
                        ItemCatName := RecItemCat.Description;

                    //<r

                    //IF printtoexcel AND ( CurrReport.SHOWOUTPUT("Transfer Shipment Line".Quantity <> 0) ) THEN BEGIN
                    /*
                    IF printtoexcel THEN BEGIN
                            XLSHEET1.Range('A'+FORMAT(i)).Value := "Transfer Shipment Header"."Posting Date" ;
                            XLSHEET1.Range('A'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('A'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('A'+FORMAT(i)).NumberFormat := 'dd/mm/yyyy';
                            XLSHEET1.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('B'+FORMAT(i)).Value := "Transfer Shipment Header"."No.";
                            XLSHEET1.Range('B'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('B'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('B'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('C'+FORMAT(i)).Value :="Transfer Shipment Header"."Transfer-from Code" ;
                            XLSHEET1.Range('C'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('C'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('C'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('D'+FORMAT(i)).Value :="Transfer Shipment Header"."Transfer-from Name";
                            XLSHEET1.Range('D'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('D'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('D'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('E'+FORMAT(i)).Value :="Transfer Shipment Header"."Transfer-to Code";
                            XLSHEET1.Range('E'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('E'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('E'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('F'+FORMAT(i)).Value :="Transfer Shipment Header"."Transfer-to Name";
                            XLSHEET1.Range('F'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('F'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('F'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('G'+FORMAT(i)).Value := "Transfer Shipment Line"."Item No.";
                            XLSHEET1.Range('G'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('G'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('G'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('H'+FORMAT(i)).Value := Item.Description;
                            XLSHEET1.Range('H'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('H'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('H'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('I'+FORMAT(i)).Value := PackedUOM;
                            XLSHEET1.Range('I'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('I'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('I'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('J'+FORMAT(i)).Value :=FORMAT(ShipmentTotalQty);
                            //XLSHEET1.Range('J'+FORMAT(i)).Value :=FORMAT(GroupQty);//17Feb2017
                            XLSHEET1.Range('J'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('J'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('J'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET1.Range('K'+FORMAT(i)).Value := FORMAT(ROUND("Unit Price"/"Qty. per Unit of Measure",0.01));
                            XLSHEET1.Range('K'+FORMAT(i)).RowHeight := 18;
                            //XLSHEET1.Range('K'+FORMAT(i)).
                            XLSHEET1.Range('K'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('K'+FORMAT(i)).NumberFormat := '#,####0.00';
                            XLSHEET1.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('L'+FORMAT(i)).Value := FORMAT("Assessable Value"/"Qty. per Unit of Measure");
                            XLSHEET1.Range('L'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('L'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('L'+FORMAT(i)).NumberFormat := '#,####0.00';
                            XLSHEET1.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('M'+FORMAT(i)).Value := FORMAT("Transfer Shipment Line".Amount);
                            XLSHEET1.Range('M'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('M'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('M'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('M'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            AMT1:= ("Transfer Shipment Line".Amount+"Transfer Shipment Line"."BED Amount");
                            AMT2:= ("Transfer Shipment Line"."eCess Amount"+"Transfer Shipment Line"."SHE Cess Amount");
                            AMT:= AMT1+AMT2+Freight+"Transfer Shipment Line"."ADE Amount";
                            vTotGross += AMT;
                            XLSHEET1.Range('N'+FORMAT(i)).Value := FORMAT(AMT)  ;
                            XLSHEET1.Range('N'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('N'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('N'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('N'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET1.Range('O'+FORMAT(i)).Value :=FORMAT("Transfer Shipment Line"."BED Amount")  ;
                            XLSHEET1.Range('O'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('O'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('O'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('O'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET1.Range('P'+FORMAT(i)).Value :=FORMAT("Transfer Shipment Line"."eCess Amount") ;
                            XLSHEET1.Range('P'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('P'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('P'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('P'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET1.Range('Q'+FORMAT(i)).Value := FORMAT("Transfer Shipment Line"."SHE Cess Amount");
                            XLSHEET1.Range('Q'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('Q'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('Q'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('Q'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET1.Range('R'+FORMAT(i)).Value :='';//r
                            XLSHEET1.Range('R'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('R'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('R'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('R'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('S'+FORMAT(i)).Value := ''; //R
                            XLSHEET1.Range('S'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('S'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('S'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('S'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('T'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('T'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('T'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('T'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('T'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('U'+FORMAT(i)).Value := "Transfer Shipment Line"."Excise Prod. Posting Group";
                            XLSHEET1.Range('U'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('U'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('U'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('U'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('V'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('V'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('V'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('V'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('V'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('W'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('W'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('W'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('W'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('W'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('W'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET1.Range('X'+FORMAT(i)).Value :=FORMAT("Transfer Shipment Line"."ADE Amount");
                            XLSHEET1.Range('X'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('X'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('X'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('X'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('X'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET1.Range('Y'+FORMAT(i)).Value :='';
                            XLSHEET1.Range('Y'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('Y'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('Y'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('Y'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET1.Range('Y'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET1.Range('Z'+FORMAT(i)).Value := ReceiptNo;
                            XLSHEET1.Range('Z'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('Z'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('Z'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('Z'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AA'+FORMAT(i)).Value := "Transfer Shipment Line"."Item Category Code";
                            XLSHEET1.Range('AA'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AA'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AA'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AA'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AB'+FORMAT(i)).Value := ItemCatName;
                            XLSHEET1.Range('AB'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AB'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AB'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AB'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AC'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('AC'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AC'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AC'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AC'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AD'+FORMAT(i)).Value := Trf2CodeTIN;
                            XLSHEET1.Range('AD'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AD'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AD'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AD'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AE'+FORMAT(i)).Value := TrffrmCodeTIN;
                            XLSHEET1.Range('AE'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AE'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AE'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AE'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AF'+FORMAT(i)).Value := CST;
                            XLSHEET1.Range('AF'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AF'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AF'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AF'+FORMAT(i)).EntireColumn.AutoFit;
                            {
                            IF "Transfer Shipment Header"."Local Driver's Mobile No." = '' THEN
                            BEGIN
                            XLSHEET1.Range('AG'+FORMAT(i)).Value := "Transfer Shipment Header"."Driver's Mobile No.";
                            XLSHEET1.Range('AG'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AG'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AG'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
                            END;
                            IF "Transfer Shipment Header"."Driver's Mobile No." = '' THEN
                            BEGIN
                            XLSHEET1.Range('AG'+FORMAT(i)).Value := "Transfer Shipment Header"."Local Driver's Mobile No.";
                            XLSHEET1.Range('AG'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AG'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AG'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
                            END;
                            IF ("Transfer Shipment Header"."Local Driver's Mobile No." = '') AND
                               ("Transfer Shipment Header"."Driver's Mobile No." = '') THEN
                            BEGIN
                            XLSHEET1.Range('AG'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('AG'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AG'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AG'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
                            END;
                            }
                            //>>RSPL-Rahul
                            IF ("Transfer Shipment Header"."Local Driver's Mobile No." = '') AND ("Transfer Shipment Header"."Driver's Mobile No." = '') THEN
                            BEGIN
                            XLSHEET1.Range('AG'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('AG'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AG'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AG'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
                            END;
                            IF "Transfer Shipment Header"."BRT Cancelled" = TRUE THEN
                            BEGIN
                            XLSHEET1.Range('AG'+FORMAT(i)).Value := 'Cancelled';
                            XLSHEET1.Range('AG'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AG'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AG'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
                            END ELSE
                            BEGIN
                            XLSHEET1.Range('AG'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('AG'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AG'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AG'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
                            END;
                    
                            //<<RSPl-Rahul
                            IF "Transfer Shipment Header"."BRT Cancelled" = TRUE THEN
                            BEGIN
                            XLSHEET1.Range('AH'+FORMAT(i)).Value := 'Cancelled';
                            XLSHEET1.Range('AH'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AH'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AH'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AH'+FORMAT(i)).EntireColumn.AutoFit;
                            END ELSE BEGIN
                            XLSHEET1.Range('AH'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('AH'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AH'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AH'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AH'+FORMAT(i)).EntireColumn.AutoFit;
                            END;
                    
                            XLSHEET1.Range('AI'+FORMAT(i)).Value := "Transfer Shipment Header"."Transfer Order No.";
                            XLSHEET1.Range('AI'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AI'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AI'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AI'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AJ'+FORMAT(i)).Value := Type;
                            XLSHEET1.Range('AJ'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AJ'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AJ'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AJ'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AK'+FORMAT(i)).Value := TransferFrm;
                            XLSHEET1.Range('AK'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AK'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AK'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AK'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AL'+FORMAT(i)).Value := "TrffrmCode-StateName";
                            XLSHEET1.Range('AL'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AL'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AL'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AL'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AM'+FORMAT(i)).Value := TransferTo;
                            XLSHEET1.Range('AM'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AM'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AM'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AM'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AN'+FORMAT(i)).Value := "TrfToCode-StateName";
                            XLSHEET1.Range('AN'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AN'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AN'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AN'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AO'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('AO'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AO'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AO'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AO'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET1.Range('AP'+FORMAT(i)).Value := '';
                            XLSHEET1.Range('AP'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AP'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AP'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET1.Range('AP'+FORMAT(i)).EntireColumn.AutoFit;
                    
                    
                            XLSHEET1.Range('AQ'+FORMAT(i)).Value := FORMAT("Transfer Shipment Line".Quantity); // EBT MILAN 030314
                            XLSHEET1.Range('AQ'+FORMAT(i)).RowHeight := 18;
                            XLSHEET1.Range('AQ'+FORMAT(i)).Font.Size := 10;
                            XLSHEET1.Range('AQ'+FORMAT(i)).NumberFormat := '#,####0';
                            XLSHEET1.Range('AQ'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            //RSPL
                            IF vReceiptDate<>0D THEN
                            BEGIN
                              XLSHEET1.Range('AR'+FORMAT(i)).Value := vReceiptDate;
                              XLSHEET1.Range('AR'+FORMAT(i)).RowHeight := 18;
                              XLSHEET1.Range('AR'+FORMAT(i)).Font.Size := 10;
                              XLSHEET1.Range('AR'+FORMAT(i)).NumberFormat :='dd/mm/yyyy';
                              XLSHEET1.Range('AR'+FORMAT(i)).EntireColumn.AutoFit;
                            END
                            ELSE
                            BEGIN
                              XLSHEET1.Range('AR'+FORMAT(i)).Value := '';
                              XLSHEET1.Range('AR'+FORMAT(i)).RowHeight := 18;
                              XLSHEET1.Range('AR'+FORMAT(i)).Font.Size := 10;
                              XLSHEET1.Range('AR'+FORMAT(i)).NumberFormat :='@';
                              XLSHEET1.Range('AR'+FORMAT(i)).EntireColumn.AutoFit;
                            END;
                            //RSPL
                             //RSPL-280115
                             XLSHEET1.Range('AS'+FORMAT(i)).Value := tComment;
                             XLSHEET1.Range('AS'+FORMAT(i)).RowHeight := 18;
                             XLSHEET1.Range('AS'+FORMAT(i)).Font.Size := 10;
                             XLSHEET1.Range('AS'+FORMAT(i)).NumberFormat :='@';
                             XLSHEET1.Range('AS'+FORMAT(i)).EntireColumn.AutoFit;
                             //RSPL-280115
                           XLSHEET1.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 5;
                           XLSHEET1.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.Color :=65535;
                         // XLSHEET1.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.Color :=0;
                    
                       i += 1;
                    END;
                    *///20Feb2017COMMENTED Because Automation Not Required in ColourFormat Excel

                    //>>20Feb2017 LineDetails
                    IF printtoexcel THEN BEGIN


                        EnterCell(NN, 1, FORMAT("Transfer Shipment Header"."Posting Date"), FALSE, FALSE, 'dd/mm/yyyy', 2);
                        EnterCell(NN, 2, "Transfer Shipment Header"."No.", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 3, "Transfer Shipment Header"."Transfer-from Code", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 4, "Transfer Shipment Header"."Transfer-from Name", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 5, "Transfer Shipment Header"."Transfer-to Code", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 6, "Transfer Shipment Header"."Transfer-to Name", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 7, "Transfer Shipment Line"."Item No.", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 8, Item.Description, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 9, PackedUOM, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 10, FORMAT(ShipmentTotalQty), FALSE, FALSE, '0.000', 0);//11APr2017
                        EnterCell(NN, 11, FORMAT(ROUND("Unit Price" / "Qty. per Unit of Measure", 0.01)), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 12, FORMAT('"Assessable Value" / "Qty. per Unit of Measure"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 13, FORMAT("Transfer Shipment Line".Amount), FALSE, FALSE, '#,####0.00', 0);

                        AMT1 := ("Transfer Shipment Line".Amount + 0);// "Transfer Shipment Line"."BED Amount");
                        AMT2 := 0;//("Transfer Shipment Line"."eCess Amount" + "Transfer Shipment Line"."SHE Cess Amount");
                        AMT := AMT1 + AMT2 + Freight + 0;//"Transfer Shipment Line"."ADE Amount";
                        vTotGross += AMT;

                        EnterCell(NN, 14, FORMAT(AMT), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 15, FORMAT('"Transfer Shipment Line"."BED Amount"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 16, FORMAT('"Transfer Shipment Line"."eCess Amount"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 17, FORMAT('"Transfer Shipment Line"."SHE Cess Amount"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 18, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 19, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 20, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 21, '"Transfer Shipment Line"."Excise Prod. Posting Group"', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 22, BatchNo, FALSE, FALSE, '@', 1);//11Apr2017
                                                                         //EnterCell(NN,22,'',FALSE,FALSE,'@',1);
                        EnterCell(NN, 23, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 24, FORMAT('"Transfer Shipment Line"."ADE Amount"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 25, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 26, ReceiptNo, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 27, "Transfer Shipment Line"."Item Category Code", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 28, ItemCatName, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 29, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 30, Trf2CodeTIN, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 31, TrffrmCodeTIN, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 32, CST, FALSE, FALSE, '@', 1);
                        IF ("Transfer Shipment Header"."Local Driver's Mobile No." = '') AND ("Transfer Shipment Header"."Driver's Mobile No." = '') THEN BEGIN
                            EnterCell(NN, 33, '', FALSE, FALSE, '@', 1);
                        END;
                        IF "Transfer Shipment Header"."BRT Cancelled" = TRUE THEN BEGIN
                            EnterCell(NN, 34, 'CANCELLED', FALSE, FALSE, '@', 1);
                        END ELSE BEGIN
                            EnterCell(NN, 34, '', FALSE, FALSE, '@', 1);
                        END;
                        EnterCell(NN, 35, "Transfer Shipment Header"."Transfer Order No.", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 36, Type, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 37, TransferFrm, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 38, "TrffrmCode-StateName", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 39, TransferTo, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 40, "TrfToCode-StateName", FALSE, FALSE, '@', 1);
                        //EnterCell(NN,41,'',FALSE,FALSE,'@',1);
                        //EnterCell(NN,42,'',FALSE,FALSE,'@',1);
                        EnterCell(NN, 41, RcvdAgInvNo, FALSE, FALSE, '@', 1);//RSPLSUM 27Jun19
                        EnterCell(NN, 42, FORMAT(RcvdAgInvDate), FALSE, FALSE, 'dd/mm/yyyy', 2);//RSPLSUM 27Jun19
                        EnterCell(NN, 43, FORMAT("Transfer Shipment Line".Quantity), FALSE, FALSE, '0.000', 0);
                        IF vReceiptDate <> 0D THEN BEGIN
                            EnterCell(NN, 44, FORMAT(vReceiptDate), FALSE, FALSE, 'dd/mm/yyyy', 2);
                        END ELSE BEGIN
                            EnterCell(NN, 44, '', FALSE, FALSE, '', 1);
                        END;
                        EnterCell(NN, 45, tComment, FALSE, FALSE, '@', 1);

                        EnterCell(NN, 46, "Transfer Shipment Header"."WH Bill Entry No.", FALSE, FALSE, '@', 1); //18May2017
                        EnterCell(NN, 47, "Transfer Shipment Header"."Debond Bill Entry No.", FALSE, FALSE, '@', 1);//18May2017
                        EnterCell(NN, 48, "Transfer Shipment Header"."EWB No.", FALSE, FALSE, '@', 1);//RSPLSUM 01Dec2020
                        EnterCell(NN, 49, FORMAT("Transfer Shipment Header"."EWB Date"), FALSE, FALSE, 'dd/mm/yyyy', 2);//RSPLSUM 01Dec2020
                        EnterCell(NN, 50, FORMAT(EWBValidity), FALSE, FALSE, 'dd/mm/yyyy hh:mm:ss', 2);//Fahim 12-Nov-21
                        EnterCell(NN, 51, "Transfer Shipment Header".IRN, FALSE, FALSE, '@', 1);//RSPLSUM 01Dec2020

                        NN += 1; //20Feb2017
                    END;


                    /*
                     XLSHEET1.Range('A'+FORMAT(i)).Value := "Transfer Shipment Header"."Posting Date" ;
                     XLSHEET1.Range('A'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('A'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('A'+FORMAT(i)).NumberFormat := 'dd/mm/yyyy';
                     XLSHEET1.Range('A'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('B'+FORMAT(i)).Value := "Transfer Shipment Header"."No.";
                     XLSHEET1.Range('B'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('B'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('B'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('B'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('C'+FORMAT(i)).Value :="Transfer Shipment Header"."Transfer-from Code" ;
                     XLSHEET1.Range('C'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('C'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('C'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('C'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('D'+FORMAT(i)).Value :="Transfer Shipment Header"."Transfer-from Name";
                     XLSHEET1.Range('D'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('D'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('D'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('D'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('E'+FORMAT(i)).Value :="Transfer Shipment Header"."Transfer-to Code";
                     XLSHEET1.Range('E'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('E'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('E'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('E'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('F'+FORMAT(i)).Value :="Transfer Shipment Header"."Transfer-to Name";
                     XLSHEET1.Range('F'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('F'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('F'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('F'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('G'+FORMAT(i)).Value :='';
                     XLSHEET1.Range('G'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('G'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('G'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('G'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('H'+FORMAT(i)).Value := '' ;
                     XLSHEET1.Range('H'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('H'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('H'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('I'+FORMAT(i)).Value := '';
                     XLSHEET1.Range('I'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('I'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('I'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('I'+FORMAT(i)).EntireColumn.AutoFit;

                    // XLSHEET1.Range('J'+FORMAT(i)).Value :=Quantity*"Transfer Shipment Line"."Qty. per Unit of Measure";
                     XLSHEET1.Range('J'+FORMAT(i)).Value := FORMAT(GroupQty);
                     XLSHEET1.Range('J'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('J'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('J'+FORMAT(i)).NumberFormat := '#,####0';
                     XLSHEET1.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                     GroupQty :=0;

                     XLSHEET1.Range('K'+FORMAT(i)).Value := '';
                     XLSHEET1.Range('K'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('K'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('K'+FORMAT(i)).NumberFormat := '';
                     XLSHEET1.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('K'+FORMAT(i)).HorizontalAlignment :=-4152;

                     XLSHEET1.Range('L'+FORMAT(i)).Value := '';
                     XLSHEET1.Range('L'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('L'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('L'+FORMAT(i)).NumberFormat := '#,####0.00';
                     XLSHEET1.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('L'+FORMAT(i)).HorizontalAlignment :=-4152;

                     //XLSHEET1.Range('M'+FORMAT(i)).Value :="Transfer Shipment Line".Amount;
                     XLSHEET1.Range('M'+FORMAT(i)).Value := FORMAT(vTotNetValue);
                     XLSHEET1.Range('M'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('M'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('M'+FORMAT(i)).NumberFormat := '#,####0';
                     XLSHEET1.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('M'+FORMAT(i)).HorizontalAlignment :=-4152;
                     CLEAR(vTotNetValue);

                     AMT1:= ("Transfer Shipment Line".Amount+"Transfer Shipment Line"."BED Amount");
                     AMT2:= ("Transfer Shipment Line"."eCess Amount"+"Transfer Shipment Line"."SHE Cess Amount");
                     AMT:= AMT1+AMT2+Freight+"Transfer Shipment Line"."ADE Amount";
                     //XLSHEET1.Range('N'+FORMAT(i)).Value := AMT  ;
                     XLSHEET1.Range('N'+FORMAT(i)).Value := FORMAT(vTotGross);
                     XLSHEET1.Range('N'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('N'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('N'+FORMAT(i)).NumberFormat := '#,####0';
                     XLSHEET1.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('N'+FORMAT(i)).HorizontalAlignment :=-4152;
                     CLEAR(vTotGross);

                     //XLSHEET1.Range('O'+FORMAT(i)).Value := "Transfer Shipment Line"."BED Amount" ;
                     XLSHEET1.Range('O'+FORMAT(i)).Value := FORMAT(vTotBEDAmt);
                     XLSHEET1.Range('O'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('O'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('O'+FORMAT(i)).NumberFormat := '#,####0';
                     XLSHEET1.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('O'+FORMAT(i)).HorizontalAlignment :=-4152;
                     CLEAR(vTotBEDAmt);

                     //XLSHEET1.Range('P'+FORMAT(i)).Value := "Transfer Shipment Line"."eCess Amount" ;
                     XLSHEET1.Range('P'+FORMAT(i)).Value := FORMAT(vTotEcess);
                     XLSHEET1.Range('P'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('P'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('P'+FORMAT(i)).NumberFormat := '#,####0';
                     XLSHEET1.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('P'+FORMAT(i)).HorizontalAlignment :=-4152;
                     CLEAR(vTotEcess);

                     //XLSHEET1.Range('Q'+FORMAT(i)).Value := "Transfer Shipment Line"."SHE Cess Amount";
                     XLSHEET1.Range('Q'+FORMAT(i)).Value := FORMAT(vTotSheCess);
                     XLSHEET1.Range('Q'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('Q'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('Q'+FORMAT(i)).NumberFormat := '#,####0';
                     XLSHEET1.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('Q'+FORMAT(i)).HorizontalAlignment :=-4152;
                     CLEAR(vTotSheCess);

                     XLSHEET1.Range('R'+FORMAT(i)).Value := LrNo;
                     XLSHEET1.Range('R'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('R'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('R'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('R'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('S'+FORMAT(i)).Value := VehicleNo;
                     XLSHEET1.Range('S'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('S'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('S'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('S'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('T'+FORMAT(i)).Value := TtrasporterName;
                     XLSHEET1.Range('T'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('T'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('T'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('T'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('U'+FORMAT(i)).Value := '';
                     XLSHEET1.Range('U'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('U'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('U'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('U'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('V'+FORMAT(i)).Value := BatchNo;
                     XLSHEET1.Range('V'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('V'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('V'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('V'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('V'+FORMAT(i)).HorizontalAlignment :=-4152;

                     XLSHEET1.Range('W'+FORMAT(i)).Value := '';
                     XLSHEET1.Range('W'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('W'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('W'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('W'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('W'+FORMAT(i)).HorizontalAlignment :=-4152;

                     //XLSHEET1.Range('X'+FORMAT(i)).Value :="Transfer Shipment Line"."ADE Amount";
                     XLSHEET1.Range('X'+FORMAT(i)).Value := FORMAT(vTotAddl);
                     XLSHEET1.Range('X'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('X'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('X'+FORMAT(i)).NumberFormat := '#,####0';
                     XLSHEET1.Range('X'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('X'+FORMAT(i)).HorizontalAlignment :=-4152;
                     CLEAR(vTotAddl);

                     XLSHEET1.Range('Y'+FORMAT(i)).Value := Freight;
                     XLSHEET1.Range('Y'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('Y'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('Y'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('Y'+FORMAT(i)).EntireColumn.AutoFit;
                     XLSHEET1.Range('Y'+FORMAT(i)).HorizontalAlignment :=-4152;

                     XLSHEET1.Range('Z'+FORMAT(i)).Value :=ReceiptNo;
                     XLSHEET1.Range('Z'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('Z'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('Z'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('Z'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AA'+FORMAT(i)).Value := '';
                     XLSHEET1.Range('AA'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AA'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AA'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AA'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AB'+FORMAT(i)).Value := '';
                     XLSHEET1.Range('AB'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AB'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AB'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AB'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AC'+FORMAT(i)).Value := "Transfer Shipment Header"."Form Code";
                     XLSHEET1.Range('AC'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AC'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AC'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AC'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AD'+FORMAT(i)).Value := Trf2CodeTIN;
                     XLSHEET1.Range('AD'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AD'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AD'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AD'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AE'+FORMAT(i)).Value := TrffrmCodeTIN;
                     XLSHEET1.Range('AE'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AE'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AE'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AE'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AF'+FORMAT(i)).Value := CST;
                     XLSHEET1.Range('AF'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AF'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AF'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AF'+FORMAT(i)).EntireColumn.AutoFit;

                   IF "Transfer Shipment Header"."Local Driver's Mobile No." = '' THEN BEGIN
                     XLSHEET1.Range('AG'+FORMAT(i)).Value := "Transfer Shipment Header"."Driver's Mobile No.";
                     XLSHEET1.Range('AG'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AG'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AG'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
                   END;
                   //>>
                   IF "Transfer Shipment Header"."Driver's Mobile No." = '' THEN BEGIN
                     XLSHEET1.Range('AG'+FORMAT(i)).Value := "Transfer Shipment Header"."Local Driver's Mobile No.";
                     XLSHEET1.Range('AG'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AG'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AG'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AG'+FORMAT(i)).EntireColumn.AutoFit;
                   END;
                   //<<
                     XLSHEET1.Range('AH'+FORMAT(i)).Value := '';
                     XLSHEET1.Range('AH'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AH'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AH'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AH'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AI'+FORMAT(i)).Value := "Transfer Shipment Header"."Transfer Order No.";
                     XLSHEET1.Range('AI'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AI'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AI'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AI'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AJ'+FORMAT(i)).Value := Type;
                     XLSHEET1.Range('AJ'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AJ'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AJ'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AJ'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AK'+FORMAT(i)).Value := TransferFrm;
                     XLSHEET1.Range('AK'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AK'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AK'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AK'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AL'+FORMAT(i)).Value := "TrffrmCode-StateName";
                     XLSHEET1.Range('AL'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AL'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AL'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AL'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AM'+FORMAT(i)).Value := TransferTo;
                     XLSHEET1.Range('AM'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AM'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AM'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AM'+FORMAT(i)).EntireColumn.AutoFit;

                     XLSHEET1.Range('AN'+FORMAT(i)).Value := "TrfToCode-StateName";
                     XLSHEET1.Range('AN'+FORMAT(i)).RowHeight := 18;
                     XLSHEET1.Range('AN'+FORMAT(i)).Font.Size := 10;
                     XLSHEET1.Range('AN'+FORMAT(i)).NumberFormat := '@';
                     XLSHEET1.Range('AN'+FORMAT(i)).EntireColumn.AutoFit;

                     recRG23D.RESET;
                     recRG23D.SETRANGE(recRG23D."Entry No.","Transfer Shipment Line"."Applies-to Entry (RG 23 D)");
                     IF NOT(recRG23D.FINDFIRST) THEN
                     BEGIN
                      XLSHEET1.Range('AO'+FORMAT(i)).Value := '';
                      XLSHEET1.Range('AO'+FORMAT(i)).RowHeight := 18;
                      XLSHEET1.Range('AO'+FORMAT(i)).Font.Size := 10;
                      XLSHEET1.Range('AO'+FORMAT(i)).NumberFormat := '@';
                      XLSHEET1.Range('AO'+FORMAT(i)).EntireColumn.AutoFit;

                      XLSHEET1.Range('AP'+FORMAT(i)).Value := '';
                      XLSHEET1.Range('AP'+FORMAT(i)).RowHeight := 18;
                      XLSHEET1.Range('AP'+FORMAT(i)).Font.Size := 10;
                      XLSHEET1.Range('AP'+FORMAT(i)).NumberFormat := '@';
                      XLSHEET1.Range('AP'+FORMAT(i)).EntireColumn.AutoFit;
                     END
                     ELSE
                     BEGIN
                      XLSHEET1.Range('AO'+FORMAT(i)).Value := TSHno;
                      XLSHEET1.Range('AO'+FORMAT(i)).RowHeight := 18;
                      XLSHEET1.Range('AO'+FORMAT(i)).Font.Size := 10;
                      XLSHEET1.Range('AO'+FORMAT(i)).NumberFormat := '@';
                      XLSHEET1.Range('AO'+FORMAT(i)).EntireColumn.AutoFit;

                      XLSHEET1.Range('AP'+FORMAT(i)).Value := TSHdate;
                      XLSHEET1.Range('AP'+FORMAT(i)).RowHeight := 18;
                      XLSHEET1.Range('AP'+FORMAT(i)).Font.Size := 10;
                      XLSHEET1.Range('AP'+FORMAT(i)).NumberFormat := 'dd/mm/yyyy';
                      XLSHEET1.Range('AP'+FORMAT(i)).EntireColumn.AutoFit;
                     END;

                     // XLSHEET1.Range('AQ'+FORMAT(i)).Value := "Transfer Shipment Line".Quantity;  // ADDED MILAN 030314
                      XLSHEET1.Range('AQ'+FORMAT(i)).Value := FORMAT(vTotNoPacks);
                      XLSHEET1.Range('AQ'+FORMAT(i)).RowHeight := 18;
                      XLSHEET1.Range('AQ'+FORMAT(i)).Font.Size := 10;
                      XLSHEET1.Range('AQ'+FORMAT(i)).NumberFormat := '#,####0';
                      XLSHEET1.Range('AQ'+FORMAT(i)).EntireColumn.AutoFit;
                      CLEAR(vTotNoPacks);
                      //RSPL
                      IF vReceiptDate<>0D THEN
                      BEGIN
                      XLSHEET1.Range('AR'+FORMAT(i)).Value := vReceiptDate;
                      XLSHEET1.Range('AR'+FORMAT(i)).RowHeight := 18;
                      XLSHEET1.Range('AR'+FORMAT(i)).Font.Size := 10;
                      XLSHEET1.Range('AR'+FORMAT(i)).NumberFormat :='dd/mm/yyyy';
                      XLSHEET1.Range('AR'+FORMAT(i)).EntireColumn.AutoFit;
                      END
                      ELSE
                      BEGIN
                      XLSHEET1.Range('AR'+FORMAT(i)).Value := '';
                      XLSHEET1.Range('AR'+FORMAT(i)).RowHeight := 18;
                      XLSHEET1.Range('AR'+FORMAT(i)).Font.Size := 10;
                      XLSHEET1.Range('AR'+FORMAT(i)).NumberFormat :='@';
                      XLSHEET1.Range('AR'+FORMAT(i)).EntireColumn.AutoFit;

                      END;
                      //RSPLi
                      //RSPL-280115
                      XLSHEET1.Range('AS'+FORMAT(i)).Value := tComment;
                      XLSHEET1.Range('AS'+FORMAT(i)).RowHeight := 18;
                      XLSHEET1.Range('AS'+FORMAT(i)).Font.Size := 10;
                      XLSHEET1.Range('AS'+FORMAT(i)).NumberFormat :='@';
                      XLSHEET1.Range('AS'+FORMAT(i)).EntireColumn.AutoFit;
                      //RSPL-280115

                     XLSHEET1.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 5;
                     //XLSHEET1.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.Color :=65535;

                 *///20Feb2017COMMENTED Because Automation Not Required in ColourFormat Excel
                   //i += 1;//20FEb2017
                   //>>
                   //20Feb2017 Header Group Details
                    DocNo += 1;
                    IF DocNo = "Transfer Shipment Line".COUNT THEN
                        IF printtoexcel AND Detail THEN BEGIN
                            DocNo := 0;
                            //NN += 1; //20Feb2017
                            EnterCell(NN, 1, FORMAT("Transfer Shipment Header"."Posting Date"), TRUE, FALSE, 'dd/mm/yyyy', 2);
                            EnterCell(NN, 2, "Transfer Shipment Header"."No.", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 3, "Transfer Shipment Header"."Transfer-from Code", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 4, "Transfer Shipment Header"."Transfer-from Name", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 5, "Transfer Shipment Header"."Transfer-to Code", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 6, "Transfer Shipment Header"."Transfer-to Name", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 7, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 8, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 9, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 10, FORMAT(GroupQty), TRUE, FALSE, '0.000', 0);
                            GroupQty := 0;
                            EnterCell(NN, 11, '', TRUE, FALSE, '', 1);
                            EnterCell(NN, 12, '', TRUE, FALSE, '', 1);
                            EnterCell(NN, 13, FORMAT(vTotNetValue), TRUE, FALSE, '#,####0.00', 0);
                            vTotNetValue := 0;

                            AMT1 := ("Transfer Shipment Line".Amount + 0);// "Transfer Shipment Line"."BED Amount");
                            AMT2 := 0;// ("Transfer Shipment Line"."eCess Amount" + "Transfer Shipment Line"."SHE Cess Amount");
                            AMT := AMT1 + AMT2 + Freight + 0;//"Transfer Shipment Line"."ADE Amount";
                            //vTotGross += AMT; //Commented 04May2017

                            EnterCell(NN, 14, FORMAT(vTotGross), TRUE, FALSE, '#,####0.00', 0);
                            vTotGross := 0;
                            EnterCell(NN, 15, FORMAT(vTotBEDAmt), TRUE, FALSE, '#,####0.00', 0);
                            vTotBEDAmt := 0;
                            EnterCell(NN, 16, FORMAT(vTotEcess), TRUE, FALSE, '#,####0.00', 0);
                            vTotEcess := 0;
                            EnterCell(NN, 17, FORMAT(vTotSheCess), TRUE, FALSE, '#,####0.00', 0);
                            vTotSheCess := 0;
                            EnterCell(NN, 18, LrNo, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 19, VehicleNo, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 20, TtrasporterName, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 21, '', TRUE, FALSE, '@', 1);

                            //EnterCell(NN,22,BatchNo,TRUE,FALSE,'@',1);//11Apr2017
                            EnterCell(NN, 23, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 24, FORMAT(vTotAddl), TRUE, FALSE, '#,####0.00', 0);
                            vTotAddl := 0;
                            EnterCell(NN, 25, FORMAT(Freight), TRUE, FALSE, '#,####0.00', 0);
                            EnterCell(NN, 26, ReceiptNo, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 27, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 28, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 29, '"Transfer Shipment Header"."Form Code"', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 30, Trf2CodeTIN, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 31, TrffrmCodeTIN, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 32, CST, TRUE, FALSE, '@', 1);
                            IF ("Transfer Shipment Header"."Local Driver's Mobile No." = '') THEN BEGIN
                                EnterCell(NN, 33, "Transfer Shipment Header"."Driver's Mobile No.", TRUE, FALSE, '@', 1);
                            END ELSE
                                IF "Transfer Shipment Header"."Driver's Mobile No." = '' THEN BEGIN
                                    EnterCell(NN, 33, "Transfer Shipment Header"."Local Driver's Mobile No.", TRUE, FALSE, '@', 1);
                                END;

                            EnterCell(NN, 34, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 35, "Transfer Shipment Header"."Transfer Order No.", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 36, Type, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 37, TransferFrm, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 38, "TrffrmCode-StateName", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 39, TransferTo, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 40, "TrfToCode-StateName", TRUE, FALSE, '@', 1);

                            EnterCell(NN, 41, '', TRUE, FALSE, '@', 1);//RSPLSUM 28Jun19
                            EnterCell(NN, 42, '', TRUE, FALSE, '@', 1);//RSPLSUM 28Jun19
                                                                       /*
                                                                       recRG23D.RESET;
                                                                       recRG23D.SETRANGE(recRG23D."Entry No.","Transfer Shipment Line"."Applies-to Entry (RG 23 D)");
                                                                       IF NOT(recRG23D.FINDFIRST) THEN
                                                                       BEGIN
                                                                       EnterCell(NN,41,'',TRUE,FALSE,'@',1);
                                                                       EnterCell(NN,42,'',TRUE,FALSE,'@',1);
                                                                       END ELSE
                                                                       BEGIN
                                                                       EnterCell(NN,41,TSHno,TRUE,FALSE,'@',1);
                                                                       EnterCell(NN,42,FORMAT(TSHdate),TRUE,FALSE,'dd/mm/yyyy',2);
                                                                       END;
                                                                       */
                            EnterCell(NN, 43, FORMAT(vTotNoPacks), TRUE, FALSE, '0.000', 0);
                            vTotNoPacks := 0;

                            IF vReceiptDate <> 0D THEN BEGIN
                                EnterCell(NN, 44, FORMAT(vReceiptDate), TRUE, FALSE, 'dd/mm/yyyy', 2);
                            END ELSE BEGIN
                                EnterCell(NN, 44, '', TRUE, FALSE, '', 1);
                            END;
                            EnterCell(NN, 45, tComment, TRUE, FALSE, '@', 1);

                            EnterCell(NN, 46, "Transfer Shipment Header"."WH Bill Entry No.", TRUE, FALSE, '@', 1); //18May2017
                            EnterCell(NN, 47, "Transfer Shipment Header"."Debond Bill Entry No.", TRUE, FALSE, '@', 1);//18May2017
                            EnterCell(NN, 48, "Transfer Shipment Header"."EWB No.", TRUE, FALSE, '@', 1);//RSPLSUM 01Dec2020
                            EnterCell(NN, 49, FORMAT("Transfer Shipment Header"."EWB Date"), TRUE, FALSE, 'dd/mm/yyyy', 2);//RSPLSUM 01Dec2020
                            EnterCell(NN, 50, FORMAT(EWBValidity), TRUE, FALSE, 'dd/mm/yyyy hh:mm:ss', 2);//Fahim 12-Nov-21
                            EnterCell(NN, 51, "Transfer Shipment Header".IRN, TRUE, FALSE, '@', 1);//RSPLSUM 01Dec2020

                            NN += 1; //20Feb2017
                                     //<<
                        END;
                    //DocNo := "Transfer Shipment Line"."Document No."; //R

                    //<<R
                    /*
                    //>>20Feb2017
                    IF NCount = NCount1 THEN
                    BEGIN
                      IF printtoexcel THEN
                      BEGIN
                      //NN += 1;
                      EnterCell(NN,1,'TOTAL',TRUE,FALSE,'@',1);
                      EnterCell(NN,10,FORMAT(TotalBRTQty),TRUE,FALSE,'#,####0',0);
                      EnterCell(NN,13,FORMAT(TotalNetValue),TRUE,FALSE,'#,####0',0);
                      EnterCell(NN,14,FORMAT(TotalNetValue+TotalBedAmt+TotaleCessAmt+TotalSHECess+TotAddDuty),TRUE,FALSE,'#,####0',0);
                      EnterCell(NN,15,FORMAT(TotalBedAmt),TRUE,FALSE,'#,####0',0);
                      EnterCell(NN,16,FORMAT(TotaleCessAmt),TRUE,FALSE,'#,####0',0);
                      EnterCell(NN,17,FORMAT(TotalSHECess),TRUE,FALSE,'#,####0',0);
                      EnterCell(NN,24,FORMAT(TotAddDuty),TRUE,FALSE,'#,####0',0);
                      EnterCell(NN,25,FORMAT(TotalFreight),TRUE,FALSE,'#,####0',0);
                      EnterCell(NN,43,FORMAT(totalnoofpacks),TRUE,FALSE,'#,####0',0);
                      END;
                    
                    END;
                    */

                end;

                trigger OnPostDataItem()
                begin

                    //}

                    //>>R


                    /*
                          XLSHEET1.Range('A'+FORMAT(i)).Value := 'TOTAL';
                          XLSHEET1.Range('A'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('A'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('A'+FORMAT(i)).Borders.ColorIndex := 5;

                          XLSHEET1.Range('J'+FORMAT(i)).Value := FORMAT(TotalBRTQty) ;
                          XLSHEET1.Range('J'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('J'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('J'+FORMAT(i)).NumberFormat := '#,####0';
                          XLSHEET1.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                          XLSHEET1.Range('J'+FORMAT(i)).Interior.ColorIndex := 45;

                          XLSHEET1.Range('K'+FORMAT(i)).Value := '';
                          XLSHEET1.Range('K'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('K'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('K'+FORMAT(i)).NumberFormat := '';
                          XLSHEET1.Range('K'+FORMAT(i)).EntireColumn.AutoFit;

                          XLSHEET1.Range('L'+FORMAT(i)).Value := '';
                          XLSHEET1.Range('L'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('L'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('L'+FORMAT(i)).NumberFormat := '#,####0.00';
                          XLSHEET1.Range('L'+FORMAT(i)).EntireColumn.AutoFit;

                          XLSHEET1.Range('M'+FORMAT(i)).Value := FORMAT(TotalNetValue) ;
                          XLSHEET1.Range('M'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('M'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('M'+FORMAT(i)).NumberFormat := '#,####0';
                          XLSHEET1.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('M'+FORMAT(i)).HorizontalAlignment :=-4152;
                          XLSHEET1.Range('M'+FORMAT(i)).Interior.ColorIndex := 45;

                          XLSHEET1.Range('N'+FORMAT(i)).Value := FORMAT(TotalNetValue+TotalBedAmt+TotaleCessAmt+TotalSHECess+
                                                                TotAddDuty);
                          XLSHEET1.Range('N'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('N'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('N'+FORMAT(i)).NumberFormat := '#,####0';
                          XLSHEET1.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('N'+FORMAT(i)).HorizontalAlignment :=-4152;
                          XLSHEET1.Range('N'+FORMAT(i)).Borders.ColorIndex := 5;

                          XLSHEET1.Range('O'+FORMAT(i)).Value := FORMAT(TotalBedAmt) ;
                          XLSHEET1.Range('O'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('O'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('O'+FORMAT(i)).NumberFormat := '#,####0';
                          XLSHEET1.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('O'+FORMAT(i)).HorizontalAlignment :=-4152;
                          XLSHEET1.Range('O'+FORMAT(i)).Borders.ColorIndex := 5;

                          XLSHEET1.Range('P'+FORMAT(i)).Value := FORMAT(TotaleCessAmt) ;
                          XLSHEET1.Range('P'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('P'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('P'+FORMAT(i)).NumberFormat := '#,####0';
                          XLSHEET1.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('P'+FORMAT(i)).HorizontalAlignment :=-4152;
                          XLSHEET1.Range('P'+FORMAT(i)).Borders.ColorIndex := 5;

                          XLSHEET1.Range('Q'+FORMAT(i)).Value := FORMAT(TotalSHECess) ;
                          XLSHEET1.Range('Q'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('Q'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('Q'+FORMAT(i)).NumberFormat := '#,####0';
                          XLSHEET1.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('Q'+FORMAT(i)).HorizontalAlignment :=-4152;
                          XLSHEET1.Range('Q'+FORMAT(i)).Borders.ColorIndex := 5;

                          XLSHEET1.Range('X'+FORMAT(i)).Value :=FORMAT(TotAddDuty);
                          XLSHEET1.Range('X'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('X'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('X'+FORMAT(i)).NumberFormat := '@';
                          XLSHEET1.Range('X'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('X'+FORMAT(i)).HorizontalAlignment :=-4152;

                          XLSHEET1.Range('Y'+FORMAT(i)).Value :=FORMAT(TotalFreight) ;
                          XLSHEET1.Range('Y'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('Y'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('Y'+FORMAT(i)).NumberFormat := '#,####0';
                          XLSHEET1.Range('Y'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('Y'+FORMAT(i)).HorizontalAlignment :=-4152;

                          XLSHEET1.Range('AQ'+FORMAT(i)).Value := FORMAT(totalnoofpacks);
                          XLSHEET1.Range('AQ'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('AQ'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('AQ'+FORMAT(i)).NumberFormat := '#,####0';
                          XLSHEET1.Range('AQ'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('AQ'+FORMAT(i)).HorizontalAlignment :=-4152;
                          {
                          //RSPL
                          XLSHEET1.Range('AR'+FORMAT(i)).Value := totalnoofpacks;
                          XLSHEET1.Range('AR'+FORMAT(i)).RowHeight := 18;
                          XLSHEET1.Range('AR'+FORMAT(i)).Font.Size := 10;
                          XLSHEET1.Range('AR'+FORMAT(i)).NumberFormat := 'dd/mm/yyyy';
                          XLSHEET1.Range('AR'+FORMAT(i)).EntireColumn.AutoFit;
                          XLSHEET1.Range('AR'+FORMAT(i)).HorizontalAlignment :=-4152;
                          //RSPL
                            }
                          XLSHEET1.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.ColorIndex :=8;


                  XLSHEET1.Range('A:AR').Borders.ColorIndex := 5;
                  //XLSHEET1.Range('A:AR').Borders.ColorIndex := 0;
                  XLAPP1.Visible(TRUE);
                  END;
                  *///20Feb2017COMMENTED Because Automation Not Required in ColourFormat Excel

                end;
            }

            trigger OnAfterGetRecord()
            begin

                NCount1 += 1;//20Feb2017

                IF NOT includeinvadj THEN
                    IF "Transfer Shipment Header"."Transfer-to Code" = 'INVADJ0001' THEN
                        CurrReport.SKIP;
                //<<R
                reclocation1.GET("Transfer Shipment Header"."Transfer-from Code");
                TransferFrm := reclocation1."State Code";
                reclocation2.GET("Transfer Shipment Header"."Transfer-to Code");
                TransferTo := reclocation2."State Code";

                IF TransferFrm <> TransferTo THEN BEGIN
                    Type := 'InterState'
                END
                ELSE
                    Type := 'Local';

                TransReceptHeader.RESET;
                TransReceptHeader.SETRANGE("Transfer Order No.", "Transfer Order No.");
                IF TransReceptHeader.FINDFIRST THEN
                    ReceiptNo := TransReceptHeader."No."
                ELSE
                    ReceiptNo := 'IN-TRANSIT';

                //EWB Validity Fahim 12-Nov-21
                CLEAR(EWBValidity);
                RecEWB.RESET;
                RecEWB.SETRANGE("Document No.", "Transfer Shipment Header"."No.");
                IF RecEWB.FINDFIRST THEN
                    EWBValidity := RecEWB."EWB Valid Upto";
                //ELSE
                //ReceiptNo:='IN-TRANSIT';
                // EWB Validity

                TH.RESET;
                TH.SETRANGE(TH."Vehicle For Location", "Transfer Shipment Header"."No.");
                IF TH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP; //<<R
                END ELSE BEGIN
                    LrNo := "Transfer Shipment Header"."LR/RR No.";
                    VehicleNo := "Transfer Shipment Header"."Vehicle No.";
                    TtrasporterName := "Transfer Shipment Header"."Transporter Name";

                    InvNoLen := STRLEN("Transfer Shipment Header"."No.");
                    InvNo := 'BRT / ' + COPYSTR("Transfer Shipment Header"."No.", (InvNoLen - 3), 4);
                END;



                CLEAR(TIN);
                CLEAR(CST);
                CLEAR(Trf2CodeTIN);
                reclocation.RESET;
                reclocation.SETRANGE(reclocation.Code, "Transfer Shipment Header"."Transfer-to Code");
                IF reclocation.FINDFIRST THEN BEGIN
                    TIN := 'reclocation."T.I.N. No."';
                    Trf2CodeTIN := 'reclocation."T.I.N. No."';
                    CST := 'reclocation."C.S.T No."';
                END;

                CLEAR(TrffrmCodeTIN);
                reclocation.RESET;
                reclocation.SETRANGE(reclocation.Code, "Transfer Shipment Header"."Transfer-from Code");
                IF reclocation.FINDFIRST THEN BEGIN
                    TrffrmCodeTIN := ' reclocation."T.I.N. No."';
                END;

                CLEAR("TrffrmCode-StateName");
                RecState.RESET;
                RecState.SETRANGE(RecState.Code, TransferFrm);
                IF RecState.FINDFIRST THEN BEGIN
                    "TrffrmCode-StateName" := RecState.Description;
                END;


                CLEAR("TrfToCode-StateName");
                RecState1.RESET;
                RecState1.SETRANGE(RecState1.Code, TransferTo);
                IF RecState1.FINDFIRST THEN BEGIN
                    "TrfToCode-StateName" := RecState1.Description;
                END;


                CLEAR(Freight);
                CLEAR(AddlDutyAmount);
                // PostedStrOrdrDetails.RESET;
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "Transfer Shipment Header"."No.");
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'FREIGHT');
                // IF PostedStrOrdrDetails.FINDFIRST THEN
                //     Freight := PostedStrOrdrDetails."Calculation Value";

                // PostedStrOrdrDetails.RESET;
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "Transfer Shipment Header"."No.");
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'ADDL.DUTY');
                // IF PostedStrOrdrDetails.FINDFIRST THEN
                //     AddlDutyAmount := PostedStrOrdrDetails."Calculation Value";

                TotalFreight := TotalFreight + Freight;
                TotalAddlDutyAmount := TotalAddlDutyAmount + AddlDutyAmount;

                ShipmentTotalQty := 0;

                //GroupQty := 0;//17Feb2017

                //RSPL-280115
                tComment := '';
                CommentLine.RESET;
                CommentLine.SETRANGE(CommentLine."Document Type", CommentLine."Document Type"::"Posted Transfer Shipment");
                CommentLine.SETRANGE(CommentLine."No.", "No.");
                IF CommentLine.FINDFIRST THEN
                    tComment := CommentLine.Comment;
                //RSPL-280115
            end;

            trigger OnPostDataItem()
            begin
                //>>20Feb2017
                //IF NCount = NCount1 THEN
                //BEGIN
                IF printtoexcel THEN BEGIN
                    //NN += 1;
                    EnterCell(NN, 1, 'ACTIVE TOTAL', TRUE, FALSE, '@', 1);
                    EnterCell(NN, 10, FORMAT(TotalBRTQty), TRUE, FALSE, '0.000', 0);//11APr2017
                    EnterCell(NN, 13, FORMAT(TotalNetValue), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 14, FORMAT(TotalNetValue + TotalBedAmt + TotaleCessAmt + TotalSHECess + TotAddDuty), TRUE, FALSE, '#,####0', 0); //09May2017
                                                                                                                                                   //EnterCell(NN,14,FORMAT(TotalNetValue+TotalBedAmt+TotaleCessAmt+TotalSHECess+TotAddDuty),TRUE,FALSE,'#,####0.00',0);
                    EnterCell(NN, 15, FORMAT(TotalBedAmt), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 16, FORMAT(TotaleCessAmt), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 17, FORMAT(TotalSHECess), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 24, FORMAT(TotAddDuty), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 25, FORMAT(TotalFreight), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 43, FORMAT(totalnoofpacks), TRUE, FALSE, '0.000', 0);

                    NN += 1;
                END;

                //END;
            end;

            trigger OnPreDataItem()
            begin

                CompInfo.GET;
                UserSetup.GET(USERID);
                Location.SETFILTER(Location.Code, UserSetup."Sales Resp. Ctr. Filter");
                IF Location.FINDFIRST THEN
                    Loc := COPYSTR(Location.Name, 7, 50);
                filtertext := Loc;


                NCount := COUNT;//20Feb2017
                NCount1 := 0; //20Feb2017
                IF printtoexcel THEN BEGIN
                    //CreateXLSHEET;  //rspl//20Feb2017
                    //CreateHeader;//20Feb2017
                    //i := 7;//20Feb2017
                    //>>20Feb2017
                    ExcBuffer.DELETEALL;
                    CLEAR(ExcBuffer);

                    EnterCell(1, 1, FORMAT(COMPANYNAME), TRUE, FALSE, '@', 1);
                    EnterCell(2, 1, filtertext, TRUE, FALSE, '@', 1);
                    EnterCell(1, 47, FORMAT(FORMAT(TODAY, 0, 4)), TRUE, FALSE, '', 1);
                    EnterCell(2, 47, FORMAT(USERID), TRUE, FALSE, '', 1);
                    //EnterCell(5,1,datefiltertext,TRUE,FALSE,'@',1);
                    EnterCell(4, 1, 'Daily Branch Transfer Register', TRUE, FALSE, '@', 1);
                    EnterCell(5, 1, 'All Active', TRUE, FALSE, '@', 1);
                    EnterCell(6, 1, 'Date Filter : ' + DateFilter, TRUE, FALSE, '@', 1);
                    EnterCell(7, 1, 'Date', TRUE, TRUE, '@', 1);
                    EnterCell(7, 2, 'Invoice', TRUE, TRUE, '@', 1);
                    EnterCell(7, 3, 'Transfer From Code', TRUE, TRUE, '@', 1);
                    EnterCell(7, 4, 'Transfer From Name', TRUE, TRUE, '@', 1);
                    EnterCell(7, 5, 'Transfer To Code', TRUE, TRUE, '@', 1);
                    EnterCell(7, 6, 'Transfer To Name', TRUE, TRUE, '@', 1);
                    EnterCell(7, 7, 'Item Code', TRUE, TRUE, '@', 1);
                    EnterCell(7, 8, 'Item Description', TRUE, TRUE, '@', 1);
                    EnterCell(7, 9, 'Pack UOM', TRUE, TRUE, '@', 1);
                    EnterCell(7, 10, 'Quantity', TRUE, TRUE, '@', 1);
                    EnterCell(7, 11, 'Trf Price Per Ltr', TRUE, TRUE, '@', 1);
                    EnterCell(7, 12, 'Ass. Value Per Ltr', TRUE, TRUE, '@', 1);
                    EnterCell(7, 13, 'Net Value', TRUE, TRUE, '@', 1);
                    EnterCell(7, 14, 'Gross Value', TRUE, TRUE, '@', 1);
                    EnterCell(7, 15, 'BED Amount', TRUE, TRUE, '@', 1);
                    EnterCell(7, 16, 'e Cess Amount', TRUE, TRUE, '@', 1);
                    EnterCell(7, 17, 'SHE Cess Amt.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 18, 'LR No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 19, 'Vehicle No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 20, 'Transporter', TRUE, TRUE, '@', 1);
                    EnterCell(7, 21, 'Chapter No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 22, 'Batch No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 23, 'Charges', TRUE, TRUE, '@', 1);
                    EnterCell(7, 24, 'Addl.Duty', TRUE, TRUE, '@', 1);
                    EnterCell(7, 25, 'Freight', TRUE, TRUE, '@', 1);
                    EnterCell(7, 26, 'Transfer Recpt.No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 27, 'Item Category Code', TRUE, TRUE, '@', 1);
                    EnterCell(7, 28, 'Item Category Name', TRUE, TRUE, '@', 1);
                    EnterCell(7, 29, 'Form Code', TRUE, TRUE, '@', 1);
                    EnterCell(7, 30, 'Trf to Code TIN No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 31, 'Trf from Code TIN No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 32, 'CST No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 33, 'Drivers Mobile No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 34, 'Status', TRUE, TRUE, '@', 1);
                    EnterCell(7, 35, 'Transfer Order No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 36, 'Type', TRUE, TRUE, '@', 1);
                    EnterCell(7, 37, 'Trf from Code - State Code', TRUE, TRUE, '@', 1);
                    EnterCell(7, 38, 'Trf from Code - State Name', TRUE, TRUE, '@', 1);
                    EnterCell(7, 39, 'Trf to Code - State Code', TRUE, TRUE, '@', 1);
                    EnterCell(7, 40, 'Trf to Code - State Name', TRUE, TRUE, '@', 1);
                    EnterCell(7, 41, 'Received Against BRT No.', TRUE, TRUE, '@', 1);
                    EnterCell(7, 42, 'Received Against BRT Date', TRUE, TRUE, '@', 1);
                    EnterCell(7, 43, 'No. of Packs', TRUE, TRUE, '@', 1);
                    EnterCell(7, 44, 'Transfer Receipt Date', TRUE, TRUE, '@', 1);
                    EnterCell(7, 45, 'Comment', TRUE, TRUE, '@', 1);
                    EnterCell(7, 46, 'WH Bill Entry No.', TRUE, TRUE, '@', 1); //18May2017
                    EnterCell(7, 47, 'Debond Bill Entry No.', TRUE, TRUE, '@', 1);//18May2017
                    EnterCell(7, 48, 'Eway Bill No.', TRUE, TRUE, '@', 1);//RSPLSUM 01Dec2020
                    EnterCell(7, 49, 'Eway Bill Date', TRUE, TRUE, '@', 1);//RSPLSUM 01Dec2020
                    EnterCell(7, 50, 'Eway Bill Validity', TRUE, TRUE, '@', 1);//Fahim 12-Nov-21
                    EnterCell(7, 51, 'IRN', TRUE, TRUE, '@', 1);//RSPLSUM 01Dec2020

                    NN := 8;

                    //<<20Feb2017
                END;
            end;
        }
        dataitem("TSH BRT CANCELLED"; "Transfer Shipment Header")
        {
            DataItemTableView = SORTING("Posting Date")
                                ORDER(Ascending)
                                WHERE("No. Series" = FILTER('<> BRTCAN'),
                                      "BRT Cancelled" = CONST(true));
            dataitem("TSL BRT CANCELLED"; "Transfer Shipment Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Quantity = FILTER(<> 0));

                trigger OnAfterGetRecord()
                begin

                    //>>25May2017

                    CLEAR(PackedUOM);
                    retItem.RESET;
                    retItem.SETRANGE(retItem."No.", "Item No.");
                    IF retItem.FINDFIRST THEN BEGIN
                        PackedUOM := retItem."Sales Unit of Measure";
                    END;

                    IF Quantity = 0 THEN
                        CurrReport.SKIP;
                    //<<R
                    CLEAR(BatchNo);
                    CLEAR(RcvdAgInvNo);
                    CLEAR(RcvdAgInvDate);
                    reILE.RESET;
                    reILE.SETRANGE(reILE."Document No.", "Document No.");
                    reILE.SETRANGE(reILE."Location Code", "Transfer-from Code");
                    reILE.SETRANGE(reILE."Item No.", "Item No.");
                    reILE.SETRANGE(reILE."Document Line No.", "Line No.");
                    IF reILE.FINDFIRST THEN BEGIN
                        BatchNo := BatchNo + reILE."Lot No.";
                        //RSPLSUM 27Jun19>>
                        IF reILE.Positive THEN BEGIN
                            ItemApplnEntry.RESET;
                            ItemApplnEntry.SETCURRENTKEY("Inbound Item Entry No.", "Outbound Item Entry No.", "Cost Application");
                            ItemApplnEntry.SETRANGE("Inbound Item Entry No.", reILE."Entry No.");
                            ItemApplnEntry.SETFILTER("Outbound Item Entry No.", '<>%1', 0);
                            ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                            IF ItemApplnEntry.FINDFIRST THEN BEGIN
                                ILERec.RESET;
                                IF ILERec.GET(ItemApplnEntry."Outbound Item Entry No.") THEN BEGIN
                                    IF ILERec."Entry Type" = ILERec."Entry Type"::Transfer THEN BEGIN
                                        RecTrRcptHdr.RESET;
                                        IF RecTrRcptHdr.GET(ILERec."Document No.") THEN BEGIN
                                            TransferShipmentHeader.RESET;
                                            TransferShipmentHeader.SETRANGE(TransferShipmentHeader."Transfer Order No.", RecTrRcptHdr."Transfer Order No.");
                                            IF TransferShipmentHeader.FINDFIRST THEN
                                                BRTNO := TransferShipmentHeader."No."
                                            ELSE
                                                BRTNO := '';

                                            RcvdAgInvNo := BRTNO;
                                            RcvdAgInvDate := RecTrRcptHdr."Posting Date";
                                        END;
                                    END;
                                    IF ILERec."Entry Type" <> ILERec."Entry Type"::Transfer THEN BEGIN
                                        RcvdAgInvNo := ILERec."Document No.";
                                        RcvdAgInvDate := ILERec."Posting Date";
                                    END;
                                END;
                            END;
                        END ELSE BEGIN
                            ItemApplnEntry.RESET;
                            ItemApplnEntry.SETCURRENTKEY("Outbound Item Entry No.", "Item Ledger Entry No.", "Cost Application");
                            ItemApplnEntry.SETRANGE("Outbound Item Entry No.", reILE."Entry No.");
                            ItemApplnEntry.SETRANGE("Item Ledger Entry No.", reILE."Entry No.");
                            ItemApplnEntry.SETRANGE("Cost Application", TRUE);
                            IF ItemApplnEntry.FINDFIRST THEN BEGIN
                                ILERec.RESET;
                                IF ILERec.GET(ItemApplnEntry."Inbound Item Entry No.") THEN BEGIN
                                    IF ILERec."Entry Type" = ILERec."Entry Type"::Transfer THEN BEGIN
                                        RecTrRcptHdr.RESET;
                                        IF RecTrRcptHdr.GET(ILERec."Document No.") THEN BEGIN
                                            TransferShipmentHeader.RESET;
                                            TransferShipmentHeader.SETRANGE(TransferShipmentHeader."Transfer Order No.", RecTrRcptHdr."Transfer Order No.");
                                            IF TransferShipmentHeader.FINDFIRST THEN
                                                BRTNO := TransferShipmentHeader."No."
                                            ELSE
                                                BRTNO := '';

                                            RcvdAgInvNo := BRTNO;
                                            RcvdAgInvDate := RecTrRcptHdr."Posting Date";
                                        END;
                                    END;
                                    IF ILERec."Entry Type" <> ILERec."Entry Type"::Transfer THEN BEGIN
                                        RcvdAgInvNo := ILERec."Document No.";
                                        RcvdAgInvDate := ILERec."Posting Date";
                                    END;
                                END;
                            END;
                        END;
                        //RSPLSUM 27Jun19<<
                    END;

                    CLEAR(TROno);
                    CLEAR(TRRno);
                    CLEAR(TSHno);
                    CLEAR(TSHdate);
                    // recRG23D.RESET;
                    // recRG23D.SETRANGE(recRG23D."Entry No.", "Applies-to Entry (RG 23 D)");
                    // IF recRG23D.FINDFIRST THEN BEGIN
                    //     TRRno := recRG23D."Document No.";

                    //     recTRH.RESET;
                    //     recTRH.SETRANGE(recTRH."No.", TRRno);
                    //     IF recTRH.FINDFIRST THEN BEGIN
                    //         TROno := recTRH."Transfer Order No.";
                    //     END;

                    //     recTSH.RESET;
                    //     recTSH.SETRANGE(recTSH."Transfer Order No.", TROno);
                    //     IF recTSH.FINDFIRST THEN BEGIN
                    //         TSHno := recTSH."No.";
                    //         TSHdate := recTSH."Posting Date";
                    //     END;

                    //     IF TSHno = '' THEN BEGIN
                    //         recRG23D.RESET;
                    //         recRG23D.SETRANGE(recRG23D."Entry No.", "Applies-to Entry (RG 23 D)");
                    //         IF recRG23D.FINDFIRST THEN BEGIN
                    //             TSHno := recRG23D."Document No.";
                    //             TSHdate := recRG23D."Posting Date";
                    //         END;
                    //     END;
                    // END;

                    BRTotalNetValue := BRTotalNetValue + Amount;
                    BRTotalBedAmt := BRTotalBedAmt + 0;//"BED Amount";
                    BRTotaleCessAmt := BRTotaleCessAmt + 0;// "eCess Amount";
                    BRTotalSHECess := BRTotalSHECess + 0;//"SHE Cess Amount";
                    BRTotAddDuty := BRTotAddDuty + 0;// "ADE Amount";
                    //AMT := AMT1 + AMT2;

                    Grpfooterqty := Grpfooterqty + Quantity;
                    Grpamt := Grpamt + Amount;
                    Grpbed := Grpbed + 0;// "BED Amount";
                    GrpeCess := GrpeCess + 0;// "eCess Amount";
                    GrpSHE := GrpSHE + 0;//"SHE Cess Amount";
                    BRtotalnoofpacks += Quantity;

                    Item.GET("Item No.");

                    //RSPL
                    vReceiptDate := 0D;
                    recTransRcptLine.RESET;
                    recTransRcptLine.SETRANGE(recTransRcptLine."Transfer Order No.", "Transfer Order No.");
                    recTransRcptLine.SETRANGE(recTransRcptLine."Line No.", "Line No.");
                    IF recTransRcptLine.FINDFIRST THEN
                        vReceiptDate := recTransRcptLine."Receipt Date";
                    //RSPL
                    //>>Rahul

                    IF RecTraShipHeader.GET("Document No.") THEN
                        RecItemCat.RESET;
                    RecItemCat.SETRANGE(RecItemCat.Code, "Item Category Code");
                    IF RecItemCat.FINDFIRST THEN
                        ItemCatName := RecItemCat.Description;

                    //nnnn
                    BRTotalBRTQty := BRTotalBRTQty + (Quantity * "Qty. per Unit of Measure");
                    BRShipmentTotalQty := Quantity * "Qty. per Unit of Measure";
                    BRGroupQty += BRShipmentTotalQty;
                    BRvTotNetValue += Amount;
                    BRvTotBEDAmt += 0;//"BED Amount";
                    BRvTotEcess += 0;//"eCess Amount";
                    BRvTotSheCess += 0;// "SHE Cess Amount";
                    BRvTotAddl += 0;// "ADE Amount";
                    BRvTotNoPacks += Quantity;

                    /*
                    AMT1:= ("Transfer Shipment Line".Amount+"Transfer Shipment Line"."BED Amount");
                            AMT2:= ("Transfer Shipment Line"."eCess Amount"+"Transfer Shipment Line"."SHE Cess Amount");
                    */
                    /*
                    IF NOT Detail THEN
                    CurrReport.SHOWOUTPUT(FALSE);
                    */

                    Grpfooterqty := 0;

                    //<r

                    BRTLINE += 1; //25May2017
                                  //>>25May2017 BRTHeader

                    IF printtoexcel AND (BRTLINE = 1) THEN BEGIN

                        NN += 2;
                        EnterCell(NN, 1, 'Cancelled Invoice', TRUE, TRUE, '@', 1);
                        NN += 1;

                        EnterCell(NN, 1, 'Date Filter : ' + DateFilter, TRUE, FALSE, '@', 1);
                        NN += 1;

                        EnterCell(NN, 1, 'Date', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 2, 'Invoice', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 3, 'Transfer From Code', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 4, 'Transfer From Name', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 5, 'Transfer To Code', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 6, 'Transfer To Name', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 7, 'Item Code', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 8, 'Item Description', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 9, 'Pack UOM', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 10, 'Quantity', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 11, 'Trf Price Per Ltr', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 12, 'Ass. Value Per Ltr', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 13, 'Net Value', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 14, 'Gross Value', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 15, 'BED Amount', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 16, 'e Cess Amount', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 17, 'SHE Cess Amt.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 18, 'LR No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 19, 'Vehicle No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 20, 'Transporter', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 21, 'Chapter No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 22, 'Batch No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 23, 'Charges', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 24, 'Addl.Duty', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 25, 'Freight', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 26, 'Transfer Recpt.No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 27, 'Item Category Code', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 28, 'Item Category Name', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 29, 'Form Code', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 30, 'Trf to Code TIN No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 31, 'Trf from Code TIN No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 32, 'CST No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 33, 'Drivers Mobile No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 34, 'Status', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 35, 'Transfer Order No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 36, 'Type', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 37, 'Trf from Code - State Code', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 38, 'Trf from Code - State Name', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 39, 'Trf to Code - State Code', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 40, 'Trf to Code - State Name', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 41, 'Received Against BRT No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 42, 'Received Against BRT Date', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 43, 'No. of Packs', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 44, 'Transfer Receipt Date', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 45, 'Comment', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 46, 'WH Bill Entry No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 47, 'Debond Bill Entry No.', TRUE, TRUE, '@', 1);
                        EnterCell(NN, 48, 'Eway Bill No.', TRUE, TRUE, '@', 1);//RSPLSUM 01Dec2020
                        EnterCell(NN, 49, 'Eway Bill Date', TRUE, TRUE, '@', 1);//RSPLSUM 01Dec2020
                        EnterCell(NN, 50, 'Eway Bill Validity', TRUE, TRUE, '@', 1);//RSPLSUM 01Dec2020
                        EnterCell(NN, 51, 'IRN', TRUE, TRUE, '@', 1);//RSPLSUM 01Dec2020

                        NN += 1;

                        //<<20Feb2017
                    END;

                    //<<25May2017 BRTHeader

                    //>>25May2017 LineDetails
                    IF printtoexcel THEN BEGIN


                        EnterCell(NN, 1, FORMAT("TSH BRT CANCELLED"."Posting Date"), FALSE, FALSE, 'dd/mm/yyyy', 2);
                        EnterCell(NN, 2, "TSH BRT CANCELLED"."No.", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 3, "TSH BRT CANCELLED"."Transfer-from Code", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 4, "TSH BRT CANCELLED"."Transfer-from Name", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 5, "TSH BRT CANCELLED"."Transfer-to Code", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 6, "TSH BRT CANCELLED"."Transfer-to Name", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 7, "Item No.", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 8, Item.Description, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 9, PackedUOM, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 10, FORMAT(-BRShipmentTotalQty), FALSE, FALSE, '0.000', 0);
                        EnterCell(NN, 11, FORMAT((-1) * "Unit Price" / "Qty. per Unit of Measure"), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 12, FORMAT('(-1) * "Assessable Value" / "Qty. per Unit of Measure"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 13, FORMAT(-Amount), FALSE, FALSE, '#,####0.00', 0);

                        //BRT nnn

                        AMT1 := (Amount + 0);
                        AMT2 := (0 + 0);
                        AMT := AMT1 + AMT2 + Freight + 0;
                        BRvTotGross += AMT;

                        EnterCell(NN, 14, FORMAT(-AMT), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 15, FORMAT('-"BED Amount"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 16, FORMAT('-"eCess Amount"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 17, FORMAT('-"SHE Cess Amount"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 18, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 19, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 20, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 21, '"Excise Prod. Posting Group"', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 22, BatchNo, FALSE, FALSE, '@', 1);//11Apr2017
                                                                         //EnterCell(NN,22,'',FALSE,FALSE,'@',1);
                        EnterCell(NN, 23, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 24, FORMAT('-"ADE Amount"'), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(NN, 25, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 26, ReceiptNo, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 27, "Item Category Code", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 28, ItemCatName, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 29, '', FALSE, FALSE, '@', 1);
                        EnterCell(NN, 30, Trf2CodeTIN, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 31, TrffrmCodeTIN, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 32, CST, FALSE, FALSE, '@', 1);
                        IF ("TSH BRT CANCELLED"."Local Driver's Mobile No." = '') AND ("TSH BRT CANCELLED"."Driver's Mobile No." = '') THEN BEGIN
                            EnterCell(NN, 33, '', FALSE, FALSE, '@', 1);
                        END;
                        IF "TSH BRT CANCELLED"."BRT Cancelled" = TRUE THEN BEGIN
                            EnterCell(NN, 34, 'CANCELLED', FALSE, FALSE, '@', 1);
                        END ELSE BEGIN
                            EnterCell(NN, 34, '', FALSE, FALSE, '@', 1);
                        END;
                        EnterCell(NN, 35, "TSH BRT CANCELLED"."Transfer Order No.", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 36, Type, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 37, TransferFrm, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 38, "TrffrmCode-StateName", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 39, TransferTo, FALSE, FALSE, '@', 1);
                        EnterCell(NN, 40, "TrfToCode-StateName", FALSE, FALSE, '@', 1);
                        //EnterCell(NN,41,'',FALSE,FALSE,'@',1);
                        //EnterCell(NN,42,'',FALSE,FALSE,'@',1);
                        EnterCell(NN, 41, RcvdAgInvNo, FALSE, FALSE, '@', 1);//RSPLSUM 27Jun19
                        EnterCell(NN, 42, FORMAT(RcvdAgInvDate), FALSE, FALSE, 'dd/mm/yyyy', 2);//RSPLSUM 27Jun19
                        EnterCell(NN, 43, FORMAT(-Quantity), FALSE, FALSE, '0.000', 0);
                        IF vReceiptDate <> 0D THEN BEGIN
                            EnterCell(NN, 44, FORMAT(vReceiptDate), FALSE, FALSE, 'dd/mm/yyyy', 2);
                        END ELSE BEGIN
                            EnterCell(NN, 44, '', FALSE, FALSE, '', 1);
                        END;
                        EnterCell(NN, 45, tComment, FALSE, FALSE, '@', 1);

                        EnterCell(NN, 46, "TSH BRT CANCELLED"."WH Bill Entry No.", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 47, "TSH BRT CANCELLED"."Debond Bill Entry No.", FALSE, FALSE, '@', 1);
                        EnterCell(NN, 48, "TSH BRT CANCELLED"."EWB No.", FALSE, FALSE, '@', 1);//RSPLSUM 01Dec2020
                        EnterCell(NN, 49, FORMAT("TSH BRT CANCELLED"."EWB Date"), FALSE, FALSE, 'dd/mm/yyyy', 2);//RSPLSUM 01Dec2020
                        EnterCell(NN, 50, FORMAT(EWBValidity), FALSE, FALSE, 'dd/mm/yyyy hh:mm:ss', 2);//Fahim 12-Nov-21
                        EnterCell(NN, 51, "TSH BRT CANCELLED".IRN, FALSE, FALSE, '@', 1);//RSPLSUM 01Dec2020

                        NN += 1; //20Feb2017
                    END;
                    //>>LineDetails


                    //25May2017 Header Group Details
                    DocNo += 1;
                    IF DocNo = COUNT THEN
                        IF printtoexcel AND Detail THEN BEGIN
                            DocNo := 0;
                            EnterCell(NN, 1, FORMAT("TSH BRT CANCELLED"."Posting Date"), TRUE, FALSE, 'dd/mm/yyyy', 2);
                            EnterCell(NN, 2, "TSH BRT CANCELLED"."No.", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 3, "TSH BRT CANCELLED"."Transfer-from Code", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 4, "TSH BRT CANCELLED"."Transfer-from Name", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 5, "TSH BRT CANCELLED"."Transfer-to Code", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 6, "TSH BRT CANCELLED"."Transfer-to Name", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 7, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 8, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 9, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 10, FORMAT(-BRGroupQty), TRUE, FALSE, '0.000', 0);
                            BRGroupQty := 0;
                            EnterCell(NN, 11, '', TRUE, FALSE, '', 1);
                            EnterCell(NN, 12, '', TRUE, FALSE, '', 1);
                            EnterCell(NN, 13, FORMAT(-BRvTotNetValue), TRUE, FALSE, '#,####0.00', 0);
                            BRvTotNetValue := 0;
                            /*
                            AMT1:= (Amount + "BED Amount");
                            AMT2:= ("eCess Amount" + "SHE Cess Amount");
                            AMT:= AMT1 + AMT2+ Freight+ "ADE Amount";
                            BRvTotGross += AMT;
                            *///Commented

                            EnterCell(NN, 14, FORMAT(-BRvTotGross), TRUE, FALSE, '#,####0.00', 0);
                            BRvTotGross := 0;

                            EnterCell(NN, 15, FORMAT(-BRvTotBEDAmt), TRUE, FALSE, '#,####0.00', 0);
                            BRvTotBEDAmt := 0;

                            EnterCell(NN, 16, FORMAT(-BRvTotEcess), TRUE, FALSE, '#,####0.00', 0);
                            BRvTotEcess := 0;

                            EnterCell(NN, 17, FORMAT(-BRvTotSheCess), TRUE, FALSE, '#,####0.00', 0);
                            BRvTotSheCess := 0;

                            EnterCell(NN, 18, LrNo, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 19, VehicleNo, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 20, TtrasporterName, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 21, '', TRUE, FALSE, '@', 1);

                            //EnterCell(NN,22,BatchNo,TRUE,FALSE,'@',1);//11Apr2017
                            EnterCell(NN, 23, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 24, FORMAT(-BRvTotAddl), TRUE, FALSE, '#,####0.00', 0);
                            BRvTotAddl := 0;

                            EnterCell(NN, 25, FORMAT(-Freight), TRUE, FALSE, '#,####0.00', 0);
                            EnterCell(NN, 26, ReceiptNo, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 27, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 28, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 29, '"Transfer Shipment Header"."Form Code"', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 30, Trf2CodeTIN, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 31, TrffrmCodeTIN, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 32, CST, TRUE, FALSE, '@', 1);
                            IF ("TSH BRT CANCELLED"."Local Driver's Mobile No." = '') THEN BEGIN
                                EnterCell(NN, 33, "TSH BRT CANCELLED"."Driver's Mobile No.", TRUE, FALSE, '@', 1);
                            END ELSE
                                IF "TSH BRT CANCELLED"."Driver's Mobile No." = '' THEN BEGIN
                                    EnterCell(NN, 33, "TSH BRT CANCELLED"."Local Driver's Mobile No.", TRUE, FALSE, '@', 1);
                                END;

                            EnterCell(NN, 34, '', TRUE, FALSE, '@', 1);
                            EnterCell(NN, 35, "TSH BRT CANCELLED"."Transfer Order No.", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 36, Type, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 37, TransferFrm, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 38, "TrffrmCode-StateName", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 39, TransferTo, TRUE, FALSE, '@', 1);
                            EnterCell(NN, 40, "TrfToCode-StateName", TRUE, FALSE, '@', 1);

                            EnterCell(NN, 41, '', TRUE, FALSE, '@', 1);//RSPLSUM 28Jun19
                            EnterCell(NN, 42, '', TRUE, FALSE, '@', 1);//RSPLSUM 28Jun19
                                                                       /*
                                                                       recRG23D.RESET;
                                                                       recRG23D.SETRANGE(recRG23D."Entry No.","Applies-to Entry (RG 23 D)");
                                                                       IF NOT(recRG23D.FINDFIRST) THEN
                                                                       BEGIN
                                                                       EnterCell(NN,41,'',TRUE,FALSE,'@',1);
                                                                       EnterCell(NN,42,'',TRUE,FALSE,'@',1);
                                                                       END ELSE
                                                                       BEGIN
                                                                       EnterCell(NN,41,TSHno,TRUE,FALSE,'@',1);
                                                                       EnterCell(NN,42,FORMAT(TSHdate),TRUE,FALSE,'dd/mm/yyyy',2);
                                                                       END;
                                                                       */
                            EnterCell(NN, 43, FORMAT(-BRvTotNoPacks), TRUE, FALSE, '0.000', 0);
                            BRvTotNoPacks := 0;

                            IF vReceiptDate <> 0D THEN BEGIN
                                EnterCell(NN, 44, FORMAT(vReceiptDate), TRUE, FALSE, 'dd/mm/yyyy', 2);
                            END ELSE BEGIN
                                EnterCell(NN, 44, '', TRUE, FALSE, '', 1);
                            END;
                            EnterCell(NN, 45, tComment, TRUE, FALSE, '@', 1);

                            EnterCell(NN, 46, "TSH BRT CANCELLED"."WH Bill Entry No.", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 47, "TSH BRT CANCELLED"."Debond Bill Entry No.", TRUE, FALSE, '@', 1);
                            EnterCell(NN, 48, "TSH BRT CANCELLED"."EWB No.", TRUE, FALSE, '@', 1);//RSPLSUM 01Dec2020
                            EnterCell(NN, 49, FORMAT("TSH BRT CANCELLED"."EWB Date"), TRUE, FALSE, 'dd/mm/yyyys', 2);//RSPLSUM 01Dec2020
                            EnterCell(NN, 50, FORMAT(EWBValidity), TRUE, FALSE, 'dd/mm/yyyy hh:mm:ss', 2);//Fahim 12-Nov-21
                            EnterCell(NN, 51, "TSH BRT CANCELLED".IRN, TRUE, FALSE, '@', 1);//RSPLSUM 01Dec2020

                            NN += 1; //
                                     //<<
                        END;
                    //Grouping
                    //<<25May2017

                end;
            }

            trigger OnAfterGetRecord()
            begin
                "TSH BRT CANCELLED".CALCFIELDS("TSH BRT CANCELLED"."EWB No.",//RSPLSUM 01Dec2020
                        "TSH BRT CANCELLED"."EWB Date", "TSH BRT CANCELLED".IRN);//RSPLSUM 01Dec2020

                //>>25May2017
                IF NOT includeinvadj THEN
                    IF "Transfer-to Code" = 'INVADJ0001' THEN
                        CurrReport.SKIP;
                //<<R
                CLEAR(TransferFrm);
                CLEAR(TransferTo);
                reclocation1.GET("Transfer-from Code");
                TransferFrm := reclocation1."State Code";
                reclocation2.GET("Transfer-to Code");
                TransferTo := reclocation2."State Code";

                IF TransferFrm <> TransferTo THEN BEGIN
                    Type := 'InterState'
                END
                ELSE
                    Type := 'Local';

                TransReceptHeader.RESET;
                TransReceptHeader.SETRANGE("Transfer Order No.", "Transfer Order No.");
                IF TransReceptHeader.FINDFIRST THEN
                    ReceiptNo := TransReceptHeader."No."
                ELSE
                    ReceiptNo := 'IN-TRANSIT';


                TH.RESET;
                TH.SETRANGE(TH."Vehicle For Location", "No.");
                IF TH.FINDFIRST THEN BEGIN
                    CurrReport.SKIP; //<<R
                END ELSE BEGIN
                    LrNo := "LR/RR No.";
                    VehicleNo := "Vehicle No.";
                    TtrasporterName := "Transporter Name";

                    InvNoLen := STRLEN("No.");
                    InvNo := 'BRT / ' + COPYSTR("No.", (InvNoLen - 3), 4);
                END;

                CLEAR(TIN);
                CLEAR(CST);
                CLEAR(Trf2CodeTIN);
                reclocation.RESET;
                reclocation.SETRANGE(reclocation.Code, "Transfer-to Code");
                IF reclocation.FINDFIRST THEN BEGIN
                    TIN := 'reclocation."T.I.N. No."';
                    Trf2CodeTIN := 'reclocation."T.I.N. No."';
                    CST := 'reclocation."C.S.T No."';
                END;

                CLEAR(TrffrmCodeTIN);
                reclocation.RESET;
                reclocation.SETRANGE(reclocation.Code, "Transfer-from Code");
                IF reclocation.FINDFIRST THEN BEGIN
                    TrffrmCodeTIN := 'reclocation."T.I.N. No."';
                END;

                CLEAR("TrffrmCode-StateName");
                RecState.RESET;
                RecState.SETRANGE(RecState.Code, TransferFrm);
                IF RecState.FINDFIRST THEN BEGIN
                    "TrffrmCode-StateName" := RecState.Description;
                END;


                CLEAR("TrfToCode-StateName");
                RecState1.RESET;
                RecState1.SETRANGE(RecState1.Code, TransferTo);
                IF RecState1.FINDFIRST THEN BEGIN
                    "TrfToCode-StateName" := RecState1.Description;
                END;


                CLEAR(Freight);
                CLEAR(AddlDutyAmount);
                // PostedStrOrdrDetails.RESET;
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'FREIGHT');
                // IF PostedStrOrdrDetails.FINDFIRST THEN
                //     Freight := PostedStrOrdrDetails."Calculation Value";

                // PostedStrOrdrDetails.RESET;
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.", "No.");
                // PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'ADDL.DUTY');
                // IF PostedStrOrdrDetails.FINDFIRST THEN
                //     AddlDutyAmount := PostedStrOrdrDetails."Calculation Value";

                BRTotalFreight := BRTotalFreight + Freight;
                BRTotalAddlDutyAmount := BRTotalAddlDutyAmount + AddlDutyAmount;

                BRShipmentTotalQty := 0;

                //GroupQty := 0;//17Feb2017

                //RSPL-280115
                tComment := '';
                CommentLine.RESET;
                CommentLine.SETRANGE(CommentLine."Document Type", CommentLine."Document Type"::"Posted Transfer Shipment");
                CommentLine.SETRANGE(CommentLine."No.", "No.");
                IF CommentLine.FINDFIRST THEN
                    tComment := CommentLine.Comment;
                //RSPL-280115
            end;

            trigger OnPostDataItem()
            begin
                //>>20Feb2017
                //IF NCount = NCount1 THEN
                //BEGIN
                IF printtoexcel THEN BEGIN
                    //NN += 1;
                    EnterCell(NN, 1, ' CANCELLED TOTAL', TRUE, FALSE, '@', 1);
                    EnterCell(NN, 10, FORMAT(-BRTotalBRTQty), TRUE, FALSE, '0.000', 0);//11APr2017
                    EnterCell(NN, 13, FORMAT(-BRTotalNetValue), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 14, FORMAT((-1) * BRTotalNetValue + BRTotalBedAmt + BRTotaleCessAmt + BRTotalSHECess + BRTotAddDuty), TRUE, FALSE, '#,####0', 0); //09May2017
                                                                                                                                                                    //EnterCell(NN,14,FORMAT(TotalNetValue+TotalBedAmt+TotaleCessAmt+TotalSHECess+TotAddDuty),TRUE,FALSE,'#,####0.00',0);
                    EnterCell(NN, 15, FORMAT(-BRTotalBedAmt), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 16, FORMAT(-BRTotaleCessAmt), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 17, FORMAT(-BRTotalSHECess), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 24, FORMAT(-BRTotAddDuty), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 25, FORMAT(-BRTotalFreight), TRUE, FALSE, '#,####0.00', 0);
                    EnterCell(NN, 43, FORMAT(-BRtotalnoofpacks), TRUE, FALSE, '0.000', 0);

                    NN += 1;
                END;

                //END;
            end;

            trigger OnPreDataItem()
            begin

                //>>25May2017
                SETRANGE("Posting Date", SDATE, EDATE);
                IF BRTFromLoc <> '' THEN
                    SETRANGE("Transfer-from Code", BRTFromLoc);

                IF BRTToLoc <> '' THEN
                    SETRANGE("Transfer-to Code", BRTToLoc);

                //<<25May2017
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(printtoexcel; printtoexcel)
                {
                    Caption = 'Print to Excel';
                }
                field(Detail; Detail)
                {
                    Caption = 'Detail';
                }
                field(includeinvadj; includeinvadj)
                {
                    Caption = 'Include Inv. Adj';
                }
            }
        }

        actions
        {
        }
    }



    trigger OnPostReport()
    begin
        //>>25May2017 GrandTotal

        IF printtoexcel THEN BEGIN
            NN += 3;
            EnterCell(NN, 1, 'GRAND TOTAL', TRUE, FALSE, '@', 1);
            EnterCell(NN, 10, FORMAT(TotalBRTQty - BRTotalBRTQty), TRUE, FALSE, '0.000', 0);//11APr2017
            EnterCell(NN, 13, FORMAT(TotalNetValue - BRTotalNetValue), TRUE, FALSE, '#,####0.00', 0);
            EnterCell(NN, 14, FORMAT(TotalNetValue + TotalBedAmt + TotaleCessAmt + TotalSHECess + TotAddDuty
                                   - BRTotalNetValue - BRTotalBedAmt - BRTotaleCessAmt - BRTotalSHECess - BRTotAddDuty), TRUE, FALSE, '#,####0', 0);
            //EnterCell(NN,14,FORMAT(TotalNetValue+TotalBedAmt+TotaleCessAmt+TotalSHECess+TotAddDuty),TRUE,FALSE,'#,####0.00',0);
            EnterCell(NN, 15, FORMAT(TotalBedAmt - BRTotalBedAmt), TRUE, FALSE, '#,####0.00', 0);
            EnterCell(NN, 16, FORMAT(TotaleCessAmt - BRTotaleCessAmt), TRUE, FALSE, '#,####0.00', 0);
            EnterCell(NN, 17, FORMAT(TotalSHECess - BRTotalSHECess), TRUE, FALSE, '#,####0.00', 0);
            EnterCell(NN, 24, FORMAT(TotAddDuty - BRTotAddDuty), TRUE, FALSE, '#,####0.00', 0);
            EnterCell(NN, 25, FORMAT(TotalFreight - BRTotalFreight), TRUE, FALSE, '#,####0.00', 0);
            EnterCell(NN, 43, FORMAT(totalnoofpacks - BRtotalnoofpacks), TRUE, FALSE, '0.000', 0);

            NN += 1;
        END;



        //<<25May2017 GrandTotal

        IF printtoexcel THEN BEGIN
            // ExcBuffer.CreateBook('', 'Daily Branch Transfer Register');
            // ExcBuffer.CreateBookAndOpenExcel('', 'Data', '', '', USERID);
            // ExcBuffer.GiveUserControl;

            ExcBuffer.CreateNewBook('Daily Branch Transfer Register');
            ExcBuffer.WriteSheet('Daily Branch Transfer Register', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Daily Branch Transfer Register', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();
        END;
    end;

    trigger OnPreReport()
    begin


        DateFilter := "Transfer Shipment Header".GETFILTER("Posting Date");
        TransferLocFilter := "Transfer Shipment Header".GETFILTER("Transfer-to Code");
        IF TransferLocFilter <> '' THEN BEGIN
            TransferLocFilter := 'For' + TransferLocFilter;
        END ELSE BEGIN
            TransferLocFilter := TransferLocFilter;
        END;

        datefiltertext := 'Daily Branch Transfer Register' + ' for the Peroid of  ' + DateFilter;

        SDATE := "Transfer Shipment Header".GETRANGEMIN("Posting Date");//25May2017
        EDATE := "Transfer Shipment Header".GETRANGEMAX("Posting Date");//25May2017
        BRTFromLoc := "Transfer Shipment Header".GETFILTER("Transfer-from Code");//25May2017
        BRTToLoc := "Transfer Shipment Header".GETFILTER("Transfer-to Code");//25May2017

        //EBT STIVAN ---(16072012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        /*
        User.GET(USERID);
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",User."User ID");
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         IF ("Transfer Shipment Header".GETFILTER("Transfer-from Code") = '') AND
            ("Transfer Shipment Header".GETFILTER("Transfer-to Code") = '') THEN
             ERROR('Please Specify Transfer From Code or Transfer To Code');
        
         CLEAR(TrffromResCode);
         CLEAR(TrftoResCode);
         CLEAR(TrffromCodeFilter);
         CLEAR(TrftoCodeFilter);
        
         recLoc1.RESET;
         recLoc1.SETRANGE(recLoc1.Code,"Transfer Shipment Header".GETFILTER("Transfer Shipment Header"."Transfer-from Code"));
         IF recLoc1.FINDFIRST THEN
         BEGIN
          TrffromResCode := recLoc1."Global Dimension 2 Code";
          TrffromCodeFilter := "Transfer Shipment Header".GETFILTER("Transfer Shipment Header"."Transfer-from Code");
         END;
        
         recLoc2.RESET;
         recLoc2.SETRANGE(recLoc2.Code,"Transfer Shipment Header".GETFILTER("Transfer Shipment Header"."Transfer-to Code"));
         IF recLoc2.FINDFIRST THEN
         BEGIN
          TrftoResCode := recLoc2."Global Dimension 2 Code";
          TrftoCodeFilter := "Transfer Shipment Header".GETFILTER("Transfer Shipment Header"."Transfer-to Code");
         END;
        
         IF TrffromCodeFilter <> '' THEN
         BEGIN
          CSOMapping1.RESET;
          CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
          CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
          CSOMapping1.SETRANGE(CSOMapping1.Value,TrffromCodeFilter);
          IF NOT(CSOMapping1.FINDFIRST) THEN
          BEGIN
           CSOMapping1.RESET;
           CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
           CSOMapping1.SETRANGE(CSOMapping1.Type,CSOMapping1.Type::"Responsibility Center");
           CSOMapping1.SETRANGE(CSOMapping1.Value,TrffromResCode);
           IF NOT(CSOMapping1.FINDFIRST) THEN
            ERROR ('You are not allowed to run this report other than your location');
          END;
         END;
        
         IF TrftoCodeFilter <> '' THEN
         BEGIN
          CSOMapping2.RESET;
          CSOMapping2.SETRANGE(CSOMapping2."User Id",UPPERCASE(USERID));
          CSOMapping2.SETRANGE(Type,CSOMapping2.Type::Location);
          CSOMapping2.SETRANGE(CSOMapping2.Value,TrftoCodeFilter);
          IF NOT(CSOMapping2.FINDFIRST) THEN
          BEGIN
           CSOMapping2.RESET;
           CSOMapping2.SETRANGE(CSOMapping2."User Id",UPPERCASE(USERID));
           CSOMapping2.SETRANGE(CSOMapping2.Type,CSOMapping2.Type::"Responsibility Center");
           CSOMapping2.SETRANGE(CSOMapping2.Value,TrftoResCode);
           IF NOT(CSOMapping2.FINDFIRST) THEN
            ERROR ('You are not allowed to run this report other than your location');
          END;
         END;
        END;
        */
        //EBT STIVAN ---(16072012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END

    end;

    var
        TotalBRTQty: Decimal;
        LrNo: Code[20];
        VehicleNo: Code[20];
        TtrasporterName: Code[100];
        InvNoLen: Integer;
        InvNo: Code[10];
        TotalNetValue: Decimal;
        TotalGrossValue: Decimal;
        TotalBedAmt: Decimal;
        TotaleCessAmt: Decimal;
        TotalSHECess: Decimal;
        DateFilter: Text[30];
        TransferLocFilter: Text[100];
        InvNoLen1: Integer;
        "--ebt": Integer;
        TH: Record 5744;
        a: Code[20];
        b: Code[20];
        filtertext: Text[100];
        datefiltertext: Text[100];
        LocFilter: Text[30];
        // XLAPP1: Automation;
        // XLWBOOK1: Automation;
        // XLSHEET1: Automation;
        // XLRANGE1: Automation;
        MaxCol: Text[3];
        i: Integer;
        CompInfo: Record 79;
        Loc: Text[30];
        CompName: Text[100];
        AMT1: Decimal;
        AMT2: Decimal;
        AMT: Decimal;
        Location: Record 14;
        UserSetup: Record 91;
        SM: Code[10];
        Grpfooterqty: Decimal;
        Grpamt: Decimal;
        Grpbed: Decimal;
        GrpeCess: Decimal;
        GrpSHE: Decimal;
        printtoexcel: Boolean;
        Customer: Record 18;
        //PostedStrOrdrLineDetails: Record 13798;
        //PostedStrOrdrDetails: Record 13760;
        AddlDutyAmount: Decimal;
        Freight: Decimal;
        Item: Record 27;
        TotalFreight: Decimal;
        TotalAddlDutyAmount: Decimal;
        Location1: Record 14;
        Print: Boolean;
        UserRespCtr: Code[10];
        Loc1: Record 14;
        Loc2: Record 14;
        ShipmentTotalQty: Decimal;
        RecLoc: Record 14;
        recresp: Record 5714;
        TransReceptHeader: Record 5746;
        ReceiptNo: Code[20];
        Detail: Boolean;
        RecTraShipHeader: Record 5744;
        RecItemCat: Record 5722;
        ItemCatName: Text[60];
        retItem: Record 27;
        PackedUOM: Text[30];
        reILE: Record 32;
        BatchNo: Code[20];
        TotAddDuty: Decimal;
        recLoc1: Record 14;
        recLoc2: Record 14;
        TrffromCodeFilter: Text[100];
        TrffromResCode: Code[100];
        TrftoCodeFilter: Text[100];
        TrftoResCode: Code[100];
        CSOMapping1: Record 50006;
        CSOMapping2: Record 50006;
        CSOMapping3: Record 50006;
        reclocation: Record 14;
        TIN: Text[30];
        CST: Text[30];
        reclocation1: Record 14;
        reclocation2: Record 14;
        TransferFrm: Code[10];
        TransferTo: Code[10];
        Type: Text[30];
        Trf2CodeTIN: Text[30];
        TrffrmCodeTIN: Text[30];
        //recRG23D: Record 16537;
        TROno: Code[20];
        recTRH: Record 5746;
        TRRno: Code[20];
        recTSH: Record 5744;
        TSHno: Code[20];
        TSHdate: Date;
        RecState: Record State;
        "TrffrmCode-StateName": Text[30];
        RecState1: Record State;
        "TrfToCode-StateName": Text[30];
        includeinvadj: Boolean;
        totalnoofpacks: Decimal;
        vReceiptDate: Date;
        recTransRcptLine: Record 5747;
        "---Comments---": Integer;
        CommentLine: Record 5748;
        tComment: Text[250];
        User: Record 2000000120;
        "---r": Integer;
        DocNo: Integer;
        GroupQty: Decimal;
        vTotNetValue: Decimal;
        vTotGross: Decimal;
        vTotBEDAmt: Decimal;
        vTotEcess: Decimal;
        vTotSheCess: Decimal;
        vTotAddl: Decimal;
        vTotFreight: Decimal;
        vTotNoPacks: Decimal;
        "--20Feb2017": Integer;
        ExcBuffer: Record 370 temporary;
        NN: Integer;
        NCount: Integer;
        NCount1: Integer;
        "--25May2017": Integer;
        BRTotalNetValue: Decimal;
        BRTotalBedAmt: Decimal;
        BRTotaleCessAmt: Decimal;
        BRTotalSHECess: Decimal;
        BRTotAddDuty: Decimal;
        BRtotalnoofpacks: Decimal;
        BRTotalBRTQty: Decimal;
        BRShipmentTotalQty: Decimal;
        BRGroupQty: Decimal;
        BRvTotNetValue: Decimal;
        BRvTotBEDAmt: Decimal;
        BRvTotEcess: Decimal;
        BRvTotSheCess: Decimal;
        BRvTotAddl: Decimal;
        BRvTotNoPacks: Decimal;
        BRvTotGross: Decimal;
        BRTLINE: Decimal;
        BRTotalFreight: Decimal;
        BRTotalAddlDutyAmount: Decimal;
        SDATE: Date;
        EDATE: Date;
        BRTFromLoc: Code[100];
        BRTToLoc: Code[100];
        ItemApplnEntry: Record 339;
        ILERec: Record 32;
        RcvdAgInvNo: Code[20];
        RcvdAgInvDate: Date;
        RecTrRcptHdr: Record 5746;
        TransferShipmentHeader: Record 5744;
        BRTNO: Code[100];
        RecEWB: Record 50044;
        EWBValidity: Code[50];

    // //[Scope('Internal')]
    procedure CreateXLSHEET()
    begin
        /*
        CLEAR(XLAPP1);
        CLEAR(XLWBOOK1);
        CLEAR(XLSHEET1);
        CREATE(XLAPP1,FALSE,TRUE);
        XLWBOOK1 := XLAPP1.Workbooks.Add;
        XLSHEET1 := XLWBOOK1.Worksheets.Add;
        XLAPP1.Visible(FALSE);
        MaxCol:= 'AS';   // ADDED MILAN 030314 from AP to AQ
        *///20Feb2017COMMENTED Because Automation Not Required in ColourFormat Excel

    end;

    // //[Scope('Internal')]
    procedure CreateHeader()
    begin
        /*
        XLSHEET1.Activate;
        
        CompInfo.GET;
        XLSHEET1.Range('A1',MaxCol+'1').Merge := TRUE;
        XLSHEET1.Range('A1',MaxCol+'1').Font.Bold := TRUE;
        XLSHEET1.Range('A1',MaxCol+'1').Value :=CompInfo.Name;
        //XLSHEET1.Range('AB',MaxCol+'1').Value :='Test2';
        XLSHEET1.Range('A1',MaxCol+'1').HorizontalAlignment := 1;
        XLSHEET1.Range('A1',MaxCol+'1').Interior.ColorIndex := 8;
        XLSHEET1.Range('A1',MaxCol+'1').Borders.ColorIndex := 5;
        XLSHEET1.Range('A1',MaxCol+'1').Merge := TRUE;
        XLSHEET1.Range('A1','A2').Merge := TRUE;
        
        XLSHEET1.Range('A2',MaxCol+'2').Value := filtertext;
        XLSHEET1.Range('A2',MaxCol+'2').Font.Bold := TRUE;
        XLSHEET1.Range('A2',MaxCol+'2').HorizontalAlignment := 1;
        XLSHEET1.Range('A2',MaxCol+'2').Interior.ColorIndex := 8;
        XLSHEET1.Range('A2',MaxCol+'2').Borders.ColorIndex := 5;
        XLSHEET1.Range('A2',MaxCol+'2').Merge := TRUE;
        //>>RSPl-Rahul
        XLSHEET1.Range('A3',MaxCol+'3').Merge := TRUE;
        XLSHEET1.Range('A3',MaxCol+'3').Merge := TRUE;
        XLSHEET1.Range('A3',MaxCol+'3').Font.Bold := TRUE;
        XLSHEET1.Range('A3',MaxCol+'3').Value :=FORMAT(TODAY,0,4);
        XLSHEET1.Range('A3',MaxCol+'3').HorizontalAlignment := -4152;
        XLSHEET1.Range('A3',MaxCol+'3').Interior.ColorIndex := 8;
        XLSHEET1.Range('A3',MaxCol+'3').Borders.ColorIndex := 5;
        XLSHEET1.Range('A3',MaxCol+'3').Merge := TRUE;
        //XLSHEET1.Range('A3','A4').Merge := TRUE;
        XLSHEET1.Range('A4',MaxCol+'4').Merge := TRUE;
        XLSHEET1.Range('A4',MaxCol+'4').Merge := TRUE;
        XLSHEET1.Range('A4',MaxCol+'4').Font.Bold := TRUE;
        XLSHEET1.Range('A4',MaxCol+'4').Value :=USERID;
        XLSHEET1.Range('A4',MaxCol+'4').HorizontalAlignment := -4152;
        XLSHEET1.Range('A4',MaxCol+'4').Interior.ColorIndex := 8;
        XLSHEET1.Range('A4',MaxCol+'4').Borders.ColorIndex := 5;
        
        //<<
        
        
        XLSHEET1.Range('A5',MaxCol+'5').Merge := TRUE;
        XLSHEET1.Range('A5',MaxCol+'5').Font.Bold := TRUE;
        XLSHEET1.Range('A5',MaxCol+'5').Value := datefiltertext;
        XLSHEET1.Range('A5',MaxCol+'5').Interior.ColorIndex := 8;
        XLSHEET1.Range('A5',MaxCol+'5').Borders.ColorIndex := 5;
        XLSHEET1.Range('A5',MaxCol+'5').Merge := TRUE;
        
        
        //XLSHEET1.Range('A5',MaxCol+'2').Merge := TRUE;
        
        
        XLSHEET1.Range('A6').Value :='Date';
        XLSHEET1.Range('A6').Font.Bold := TRUE;
        XLSHEET1.Range('A6',MaxCol+'6').Interior.ColorIndex := 45;
        XLSHEET1.Range('A6',MaxCol+'6').Borders.ColorIndex := 5;
        
        XLSHEET1.Range('B6').Value := 'Invoice No.';
        XLSHEET1.Range('B6').Font.Bold := TRUE;
        
        XLSHEET1.Range('C6').Value := 'Transfer From Code';
        XLSHEET1.Range('C6').Font.Bold := TRUE;
        
        XLSHEET1.Range('D6').Value := 'Transfer From Name';
        XLSHEET1.Range('D6').Font.Bold := TRUE;
        
        XLSHEET1.Range('E6').Value := 'Transfer To Code';
        XLSHEET1.Range('E6').Font.Bold := TRUE;
        
        XLSHEET1.Range('F6').Value := 'Transfer To Name';
        XLSHEET1.Range('F6').Font.Bold := TRUE;
        
        XLSHEET1.Range('G6').Value := 'Item Code';
        XLSHEET1.Range('G6').Font.Bold := TRUE;
        
        XLSHEET1.Range('H6').Value := 'Item Description';
        XLSHEET1.Range('H6').Font.Bold := TRUE;
        
        XLSHEET1.Range('I6').Value := 'Pack UOM';
        XLSHEET1.Range('I6').Font.Bold := TRUE;
        
        XLSHEET1.Range('J6').Value := 'Quantity';
        XLSHEET1.Range('J6').Font.Bold := TRUE;
        
        XLSHEET1.Range('K6').Value := 'Trf Price per Ltr';
        XLSHEET1.Range('K6').Font.Bold := TRUE;
        
        XLSHEET1.Range('L6').Value := 'Ass. Value Per Ltr';
        XLSHEET1.Range('L6').Font.Bold := TRUE;
        
        XLSHEET1.Range('M6').Value := 'Net Value';
        XLSHEET1.Range('M6').Font.Bold := TRUE;
        
        XLSHEET1.Range('N6').Value := 'Gross Value';
        XLSHEET1.Range('N6').Font.Bold := TRUE;
        
        XLSHEET1.Range('O6').Value := 'BED Amount';
        XLSHEET1.Range('O6').Font.Bold := TRUE;
        
        XLSHEET1.Range('P6').Value := 'e Cess Amount';
        XLSHEET1.Range('P6').Font.Bold := TRUE;
        
        XLSHEET1.Range('Q6').Value := 'SHE Cess Amt.';
        XLSHEET1.Range('Q6').Font.Bold := TRUE;
        
        XLSHEET1.Range('R6').Value := 'LR No.';
        XLSHEET1.Range('R6').Font.Bold := TRUE;
        
        XLSHEET1.Range('S6').Value := 'Vehicle No.';
        XLSHEET1.Range('S6').Font.Bold := TRUE;
        
        XLSHEET1.Range('T6').Value := 'Transporter';
        XLSHEET1.Range('T6').Font.Bold := TRUE;
        
        XLSHEET1.Range('U6').Value := 'Chapter No.';
        XLSHEET1.Range('U6').Font.Bold := TRUE;
        
        XLSHEET1.Range('V6').Value := 'Batch No.';
        XLSHEET1.Range('V6').Font.Bold := TRUE;
        
        XLSHEET1.Range('W6').Value := 'Charges';
        XLSHEET1.Range('W6').Font.Bold := TRUE;
        
        XLSHEET1.Range('X6').Value := 'Addl.Duty';
        XLSHEET1.Range('X6').Font.Bold := TRUE;
        
        XLSHEET1.Range('Y6').Value := 'Freight';
        XLSHEET1.Range('Y6').Font.Bold := TRUE;
        
        XLSHEET1.Range('Z6').Value := 'Transfer Recpt. No.';
        XLSHEET1.Range('Z6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AA6').Value := 'Item Caterogy Code';
        XLSHEET1.Range('AA6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AB6').Value := 'Item Caterogy Name';
        XLSHEET1.Range('AB6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AC6').Value := 'Form Code';
        XLSHEET1.Range('AC6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AD6').Value := 'Trf to Code TIN No.';
        XLSHEET1.Range('AD6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AE6').Value := 'Trf frm Code TIN No.';
        XLSHEET1.Range('AE6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AF6').Value := 'CST No.';
        XLSHEET1.Range('AF6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AG6').Value := 'Drivers Mobile No.';
        XLSHEET1.Range('AG6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AH6').Value := 'Status';
        XLSHEET1.Range('AH6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AI6').Value := 'Transfer Order No.';
        XLSHEET1.Range('AI6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AJ6').Value := 'Type';
        XLSHEET1.Range('AJ6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AK6').Value := 'Trf from Code - State Code';
        XLSHEET1.Range('AK6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AL6').Value := 'Trf from Code - State Name';
        XLSHEET1.Range('AL6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AM6').Value := 'Trf to Code - State Code';
        XLSHEET1.Range('AM6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AN6').Value := 'Trf to Code - State Name';
        XLSHEET1.Range('AN6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AO6').Value := 'Received Against BRT No.';
        XLSHEET1.Range('AO6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AP6').Value := 'Received Against BRT Date';
        XLSHEET1.Range('AP6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AQ6').Value := 'No. of Packs';  // ADDED MILAN 030314
        XLSHEET1.Range('AQ6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AR6').Value := 'Transfer Receipt Date';  //RSPL
        XLSHEET1.Range('AR6').Font.Bold := TRUE;
        
        XLSHEET1.Range('AS6').Value := 'Comment';  //RSPL-280115
        XLSHEET1.Range('AS6').Font.Bold := TRUE;
        
        
        XLSHEET1.Range('A6:AS6').Borders.ColorIndex := 5;
        XLAPP1.Visible(TRUE);
        *///20Feb2017COMMENTED Because Automation Not Required in ColourFormat Excel

    end;

    /// //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        IF printtoexcel THEN BEGIN
            ExcBuffer.INIT;
            ExcBuffer.VALIDATE("Row No.", Rowno);
            ExcBuffer.VALIDATE("Column No.", columnno);
            ExcBuffer."Cell Value as Text" := Cellvalue;
            ExcBuffer.Formula := '';
            ExcBuffer.Bold := Bold;
            ExcBuffer.Underline := Underline;
            ExcBuffer.NumberFormat := NoFormat;
            ExcBuffer."Cell Type" := CType;
            ExcBuffer.INSERT;
        END;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


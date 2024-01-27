report 50051 "RM Movement (Locwise)"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/RMMovementLocwise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("Item Category Code", "No.")
                                ORDER(Ascending)
                                WHERE("Item Category Code" = FILTER('CAT01|CAT04|CAT05|CAT06|CAT07|CAT08|CAT09|CAT10|CAT15'),
                                      Blocked = FILTER(false));
            RequestFilterFields = "No.", "Date Filter", "Item Category Code", "Location Filter";

            trigger OnAfterGetRecord()
            begin


                TCOUNT += 1;//15Mar2017

                //>>15Mar2017 GroupTotal
                NDate := Item.GETRANGEMAX(Item."Date Filter");

                CLEAR(TempItem);
                Item1403.SETCURRENTKEY("Item Category Code", "No.");
                Item1403.SETRANGE("Item Category Code", "Item Category Code");
                //Item1403.SETFILTER("Date Filter",'<%1',NDate);//testing
                //Item1403.SETFILTER("Inventory Change",'<>%1',0);
                IF Item1403.FINDLAST THEN BEGIN
                    //Item1403.CALCFIELDS("Inventory Change");
                    TempItem := Item1403."No.";
                    ICOUNT := Item1403.COUNT;
                    //MESSAGE('Srno %1 \ TCount %2 \ ICount %3 \\ No. %4 \\ F Item No. %5 ',SrNo,TCOUNT,ICOUNT,Item."No.",Item1403."No.");
                END;
                //<<15Mar2017 GroupTotal



                //>>1


                OPBAL := 0;
                QTYMAN := 0;
                QTYISSUED := 0;
                QTYRECEIVED := 0;
                QTYPRODUCED := 0;
                CLOSEBAL := 0;
                ManualNegAdj := 0;
                QTYISSUED1 := 0;
                PurQty := 0;
                PurchRetQty := 0;
                TotalReceipt := 0;
                Posadj := 0;
                Negadj := 0;

                PosDate := Item.GETRANGEMIN(Item."Date Filter");
                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", '<%1', PosDate);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    OPBAL := ItemMaster."Inventory Change";
                END;

                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER("Entry Type Filter", 'Output');
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    QTYPRODUCED := ItemMaster."Inventory Change"
                END;


                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER("Entry Type Filter", 'Positive Adjmt.');
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    Posadj := ItemMaster."Inventory Change"
                END;


                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER("Entry Type Filter", 'Purchase');
                ItemMaster.SETFILTER("Document Type Filter", 'Purchase Receipt');
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    PurQty := ItemMaster."Inventory Change";
                END;
                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER("Entry Type Filter", 'Negative Adjmt.');
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    Negadj := ABS(ItemMaster."Inventory Change");
                END;

                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER("Entry Type Filter", 'Transfer');
                ItemMaster.SETFILTER(ItemMaster."Document Type Filter", 'Transfer Receipt');
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    QTYRECEIVED := ItemMaster."Inventory Change"
                END;

                TotalReceipt := Posadj + PurQty + QTYRECEIVED + QTYPRODUCED;

                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER("Entry Type Filter", 'Consumption');
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    QTYISSUED := ABS(ItemMaster."Inventory Change")
                END;

                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER("Entry Type Filter", 'Transfer');
                ItemMaster.SETFILTER(ItemMaster."Document Type Filter", 'Transfer Shipment');
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    QTYISSUED1 := ABS(ItemMaster."Inventory Change");
                END;

                ItemMaster.RESET;
                ItemMaster.SETRANGE("No.", "No.");
                ItemMaster.SETFILTER("Date Filter", DateFILTER);
                IF LocFilter <> '' THEN
                    ItemMaster.SETRANGE(ItemMaster."Location Filter", LocFilter);
                ItemMaster.SETFILTER("Entry Type Filter", 'Purchase');
                ItemMaster.SETFILTER(ItemMaster."Document Type Filter", 'Purchase Return Shipment');
                IF ItemMaster.FINDFIRST THEN BEGIN
                    ItemMaster.CALCFIELDS("Inventory Change");
                    PurchRetQty := ABS(ItemMaster."Inventory Change");
                END;

                TotalIssued := Negadj + QTYISSUED + QTYISSUED1 + PurchRetQty;

                CLOSEBAL := OPBAL + TotalReceipt - TotalIssued;

                /*
                IF ((OPBAL=0) AND (QTYRECEIVED=0)AND(QTYISSUED=0)AND(QTYISSUED1=0) AND (PurQty =0))  THEN
                BEGIN
                 CurrReport.SKIP;
                END;
                *///Commented 15Mar2017

                TotalOPBAL := TotalOPBAL + OPBAL;

                TotalPurqty += PurQty;
                TotalProduced += QTYPRODUCED;
                TotalPosAdj += Posadj;
                TotalNegAdj += Negadj;
                TotalQTYRECEIVED += QTYRECEIVED;
                TotalReceipt1 += TotalReceipt;
                TotalQTYISSUED += QTYISSUED;
                TotalQtyIssued1 += QTYISSUED1;
                TotalPurchRetQty += PurchRetQty;
                TotalIssued1 += TotalIssued;
                TotalCLOSEBAL += CLOSEBAL;

                GroupOpBal += OPBAL;
                GroupPerQty += PurQty;
                GroupProduced += QTYPRODUCED;
                GroupPosAdj += Posadj;
                GroupNegAdj += Negadj;
                GroupQtyReceive += QTYRECEIVED;
                GroupRecep1 += TotalReceipt;
                GroupQctIssuid += QTYISSUED;
                GroupQcyIssuid1 += QTYISSUED1;
                GroupPurchretQty += PurchRetQty;
                groupClosBal += CLOSEBAL;
                GroupIssud1 += TotalIssued;

                //RSPL
                itemDim := '';
                recDefualtDim.RESET;
                recDefualtDim.SETRANGE(recDefualtDim."Table ID", 27);
                recDefualtDim.SETRANGE(recDefualtDim."No.", Item."No.");
                IF recDefualtDim.FINDFIRST THEN
                    itemDim := recDefualtDim."Dimension Value Code";
                //RSPL
                //<<1




                //>>2
                CLEAR(NAH);
                //Item, Body (3) - OnPreSection()
                IF (OPBAL = 0) AND (CLOSEBAL = 0) AND (QTYMAN = 0) AND (QTYRECEIVED = 0) AND (QTYISSUED = 0) AND (QTYISSUED1 = 0) AND (ManualNegAdj = 0)
                AND (PurQty = 0) THEN BEGIN
                    //CurrReport.SKIP; //17Mar2017
                    NAH := TRUE; //17Mar2017

                END;

                //>>17Mar2017 Report DataBody
                IF PrinttoExcel THEN
                    DataBody;

                //<<17Mar2017 Report DataBody


                //>>17Mar2017 Report GroupDataBody
                IF PrinttoExcel AND (TCOUNT = ICOUNT) THEN BEGIN
                    GroupDataBody;
                END;

                //<<17Mar2017 Report GroupDataBody




                //<<15Mar2017 ReportDataGroup

                /*
                IF PrinttoExcel THEN BEGIN
                        XLSHEET.Range('A'+FORMAT(i)).Value := SrNo;
                        XLSHEET.Range('A'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('A'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('B'+FORMAT(i)).Value :=  Item."No." ;
                        XLSHEET.Range('B'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('B'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('C'+FORMAT(i)).Value :=  Item.Description ;
                        XLSHEET.Range('C'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('C'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('C'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
                
                        ItemCaterogy.RESET;
                        ItemCaterogy.SETRANGE(ItemCaterogy.Code,Item."Item Category Code");
                        IF ItemCaterogy.FINDFIRST THEN
                        BEGIN
                         ItemCaterogyName := ItemCaterogy.Description
                        END;
                        XLSHEET.Range('D'+FORMAT(i)).Value :=  Item."Item Category Code" +' - ' + ItemCaterogyName;
                        XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('E'+FORMAT(i)).Value :=  "Base Unit of Measure";
                        XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('F'+FORMAT(i)).Value :=  OPBAL ;
                        XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('F'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('G'+FORMAT(i)).Value :=  QTYPRODUCED ;
                        XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('G'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                
                
                        XLSHEET.Range('H'+FORMAT(i)).Value :=  PurQty ;
                        XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('H'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('I'+FORMAT(i)).Value :=  Posadj ;
                        XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('I'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                
                
                        XLSHEET.Range('J'+FORMAT(i)).Value :=  QTYRECEIVED ;
                        XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('J'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('K'+FORMAT(i)).Value :=  TotalReceipt ;
                        XLSHEET.Range('K'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('K'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('K'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('L'+FORMAT(i)).Value := QTYISSUED ;
                        XLSHEET.Range('L'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('L'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('L'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('M'+FORMAT(i)).Value := Negadj ;
                        XLSHEET.Range('M'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('M'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('M'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                
                
                        XLSHEET.Range('N'+FORMAT(i)).Value := QTYISSUED1 ;
                        XLSHEET.Range('N'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('N'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('N'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('O'+FORMAT(i)).Value := PurchRetQty ;
                        XLSHEET.Range('O'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('O'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('O'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('O'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('P'+FORMAT(i)).Value := TotalIssued ;
                        XLSHEET.Range('P'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('P'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('P'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('P'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('Q'+FORMAT(i)).Value := CLOSEBAL ;
                        XLSHEET.Range('Q'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('Q'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('Q'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                
                        //RSPL
                        XLSHEET.Range('R'+FORMAT(i)).Value := itemDim ;
                        XLSHEET.Range('R'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('R'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('R'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('R'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('R'+FORMAT(i)).EntireColumn.AutoFit;
                        //RSPL
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.Color :=65535;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 5;
                
                   i += 1;
                 END;
                 *///Commented Automation 14Mar2017
                   //<<2


                //>>3

                //Item, GroupFooter (4) - OnPreSection()
                /* IF (GroupOpBal=0) AND (groupClosBal=0) AND (GroupProduced=0) AND(GroupQtyReceive=0)
                 AND(GroupQctIssuid=0)AND(GroupQcyIssuid1=0)AND(GroupNegAdj=0) AND (GroupPosAdj=0)
                 AND (GroupRecep1=0) AND (GroupPurchretQty=0) AND (GroupIssud1=0)
                 AND (GroupPerQty=0)  THEN
                    reportshow := FALSE
                 ELSE reportshow := TRUE;
                 *///Commented 15Mar2017



                /*
                IF PrinttoExcel AND reportshow=TRUE THEN BEGIN
                        XLSHEET.Range('F'+FORMAT(i)).Value := GroupOpBal;
                        XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('F'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('G'+FORMAT(i)).Value := GroupProduced ;
                        XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('G'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                
                        XLSHEET.Range('H'+FORMAT(i)).Value := GroupPerQty ;
                        XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('H'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('I'+FORMAT(i)).Value := GroupPosAdj ;
                        XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('I'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                
                        XLSHEET.Range('J'+FORMAT(i)).Value := GroupQtyReceive ;
                        XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                
                        XLSHEET.Range('K'+FORMAT(i)).Value := GroupRecep1;
                        XLSHEET.Range('K'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('K'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('K'+FORMAT(i)).HorizontalAlignment := -4152;
                
                
                        XLSHEET.Range('L'+FORMAT(i)).Value := GroupQctIssuid;
                        XLSHEET.Range('L'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('L'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('L'+FORMAT(i)).HorizontalAlignment := -4152;
                
                
                        XLSHEET.Range('M'+FORMAT(i)).Value := GroupNegAdj;
                        XLSHEET.Range('M'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('M'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('M'+FORMAT(i)).HorizontalAlignment := -4152;
                
                
                        XLSHEET.Range('N'+FORMAT(i)).Value := GroupQcyIssuid1;
                        XLSHEET.Range('N'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('N'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('N'+FORMAT(i)).HorizontalAlignment := -4152;
                
                        XLSHEET.Range('O'+FORMAT(i)).Value := GroupPurchretQty;
                        XLSHEET.Range('O'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('O'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('O'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('O'+FORMAT(i)).HorizontalAlignment := -4152;
                
                        XLSHEET.Range('P'+FORMAT(i)).Value :=  GroupIssud1;
                        XLSHEET.Range('P'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('P'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('P'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('P'+FORMAT(i)).HorizontalAlignment := -4152;
                
                        XLSHEET.Range('Q'+FORMAT(i)).Value := groupClosBal;
                        XLSHEET.Range('Q'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('Q'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('Q'+FORMAT(i)).HorizontalAlignment := -4152;
                
                        //RSPL
                        XLSHEET.Range('R'+FORMAT(i)).Value := itemDim;
                        XLSHEET.Range('R'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('R'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('R'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('R'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('R'+FORMAT(i)).HorizontalAlignment := -4152;
                        //RSPL
                
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 6;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.ColorIndex :=15;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Font.Bold := TRUE;
                        XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('B'+FORMAT(i)).Value :='Group Total';
                
                        i+=1;
                
                //Item, GroupFooter (5) - OnPreSection()
                
                GroupPerQty:=0;
                GroupOpBal:=0;
                GroupPosAdj:=0;
                GroupNegAdj:=0;
                GroupProduced:=0;
                GroupQtyReceive:=0;
                GroupRecep1:=0;
                GroupQctIssuid:=0;
                GroupQcyIssuid1:=0;
                GroupPurchretQty:=0;
                groupClosBal:=0;
                GroupIssud1:=0;
                
                END;
                *///Commented Automation 14Mar2017
                  //<<3


                //>>4


                //<<4

            end;

            trigger OnPostDataItem()
            begin

                //>>15Mar2017 ReportGrandTotal
                IF PrinttoExcel THEN BEGIN
                    EnterCell(NN, 1, '', TRUE, TRUE, '', 0);
                    EnterCell(NN, 2, 'Grand Total', TRUE, TRUE, '@', 1);
                    EnterCell(NN, 3, '', TRUE, TRUE, '@', 1);
                    EnterCell(NN, 4, '', TRUE, TRUE, '@', 1);
                    EnterCell(NN, 5, '', TRUE, TRUE, '@', 1);
                    EnterCell(NN, 6, FORMAT(TotalOPBAL), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 7, FORMAT(TotalProduced), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 8, FORMAT(TotalPurqty), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 9, FORMAT(TotalPosAdj), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 10, FORMAT(TotalQTYRECEIVED), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 11, FORMAT(TotalReceipt1), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 12, FORMAT(TotalQTYISSUED), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 13, FORMAT(TotalNegAdj), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 14, FORMAT(TotalQtyIssued1), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 15, FORMAT(TotalPurchRetQty), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 16, FORMAT(TotalIssued1), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 17, FORMAT(TotalCLOSEBAL), TRUE, TRUE, '#####0.000', 0);
                    EnterCell(NN, 18, '', TRUE, TRUE, '@', 1);

                    NN += 1;

                END;
                //<<15Mar2017 ReportGrandTotal
                //>>1

                //Item, Footer (8) - OnPreSection()
                /*
                IF PrinttoExcel THEN BEGIN
                        XLSHEET.Range('F'+FORMAT(i)).Value := TotalOPBAL;
                        XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('F'+FORMAT(i)).HorizontalAlignment := -4152;
                        XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('G'+FORMAT(i)).Value := TotalProduced ;
                        XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('G'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                
                        XLSHEET.Range('H'+FORMAT(i)).Value := TotalPurqty ;
                        XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('H'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                        XLSHEET.Range('I'+FORMAT(i)).Value := TotalPosAdj ;
                        XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('I'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                
                        XLSHEET.Range('J'+FORMAT(i)).Value := TotalQTYRECEIVED ;
                        XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                
                
                        XLSHEET.Range('K'+FORMAT(i)).Value := TotalReceipt1;
                        XLSHEET.Range('K'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('K'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('K'+FORMAT(i)).HorizontalAlignment := -4152;
                
                
                        XLSHEET.Range('L'+FORMAT(i)).Value := TotalQTYISSUED;
                        XLSHEET.Range('L'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('L'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('L'+FORMAT(i)).HorizontalAlignment := -4152;
                
                
                        XLSHEET.Range('M'+FORMAT(i)).Value := TotalNegAdj;
                        XLSHEET.Range('M'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('M'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('M'+FORMAT(i)).HorizontalAlignment := -4152;
                
                
                        XLSHEET.Range('N'+FORMAT(i)).Value := TotalQtyIssued1;
                        XLSHEET.Range('N'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('N'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('N'+FORMAT(i)).HorizontalAlignment := -4152;
                
                        XLSHEET.Range('O'+FORMAT(i)).Value := TotalPurchRetQty;
                        XLSHEET.Range('O'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('O'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('O'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('O'+FORMAT(i)).HorizontalAlignment := -4152;
                
                        XLSHEET.Range('P'+FORMAT(i)).Value := TotalIssued1;
                        XLSHEET.Range('P'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('P'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('P'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('P'+FORMAT(i)).HorizontalAlignment := -4152;
                
                        XLSHEET.Range('Q'+FORMAT(i)).Value := TotalCLOSEBAL;
                        XLSHEET.Range('Q'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('Q'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('Q'+FORMAT(i)).HorizontalAlignment := -4152;
                
                
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 5;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.ColorIndex :=8;
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Font.Bold := TRUE;
                        XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('B'+FORMAT(i)).Value :='Grand Total';
                 END;
                 *///Commented Automation 14Mar2017
                   //<<1

            end;

            trigger OnPreDataItem()
            begin

                SETCURRENTKEY("Item Category Code", "No.");//15Mar2017

                //>>1

                CompInfo.GET;

                /*
                IF PrinttoExcel THEN BEGIN
                CreateXLSHEET;
                CreateHeader;
                i := 7;
                END;
                *///Commented Automation 14Mar2017
                  //<<1

                Item1403.COPYFILTERS(Item);//15Mar2017
                TCOUNT := 0;//15Mar2017

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print To Excel"; PrinttoExcel)
                {
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

        //>>14Mar2017 RB-N
        IF PrinttoExcel THEN BEGIN
            // ExcBuffer.CreateBook('', 'Stock Statement of RM');
            // ExcBuffer.CreateBookAndOpenExcel('', 'Stock Statement of RM', '', '', USERID);
            // ExcBuffer.GiveUserControl;

            ExcBuffer.CreateNewBook('Stock Statement of RM');
            ExcBuffer.WriteSheet('Stock Statement of RM', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Stock Statement of RM', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();

        END;
        //<<
    end;

    trigger OnPreReport()
    begin

        //>>1

        DateFILTER := Item.GETFILTER("Date Filter");
        LocFilter := Item.GETFILTER(Item."Location Filter");
        Location.SETRANGE(Location.Code, LocFilter);
        IF Location.FINDFIRST THEN;
        Loc := COPYSTR(Location.Name, 7, 50);
        filtertext := Loc;

        ItemCatFilter := Item.GETFILTER(Item."Item Category Code");

        recitemCaterogy.RESET;
        recitemCaterogy.SETRANGE(recitemCaterogy.Code, ItemCatFilter);
        IF recitemCaterogy.FINDFIRST THEN BEGIN
            ItemCatName := recitemCaterogy.Description;
        END;

        datefiltertext := 'Stock Statement of Raw Materials  ' + ItemCatFilter + ' - ' + ItemCatName + ' for the Peroid of  ' + DateFILTER;

        BinFilter := Item.GETFILTER("Bin Filter");
        recBIN.SETRANGE(recBIN.Code, BinFilter);
        IF recBIN.FINDFIRST THEN
            BinName := recBIN.Description;

        BinFilterText := BinFilter + ' - ' + BinName;
        //<<1

        //>>15Mar2017 ReportHeader
        IF PrinttoExcel THEN BEGIN
            EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '@', 1);
            EnterCell(1, 18, FORMAT(TODAY, 0, 4), TRUE, FALSE, '@', 1);

            EnterCell(2, 1, 'Stock Statement of Raw Materials', TRUE, FALSE, '@', 1);
            EnterCell(2, 18, USERID, TRUE, FALSE, '@', 1);

            EnterCell(5, 1, 'Item Category', TRUE, TRUE, '@', 1);
            EnterCell(5, 2, ItemCatFilter + ' - ' + ItemCatName, TRUE, TRUE, '@', 1);

            EnterCell(6, 1, 'Date Filter', TRUE, TRUE, '@', 1);
            EnterCell(6, 2, DateFILTER, TRUE, TRUE, '@', 1);
            EnterCell(6, 3, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 4, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 5, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 6, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 7, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 8, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 9, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 10, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 11, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 12, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 13, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 14, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 15, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 16, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 17, '', TRUE, TRUE, '@', 1);
            EnterCell(6, 18, '', TRUE, TRUE, '@', 1);

            EnterCell(7, 1, 'SrNo', TRUE, TRUE, '@', 1);
            EnterCell(7, 2, 'Item Code', TRUE, TRUE, '@', 1);
            EnterCell(7, 3, 'Item Name', TRUE, TRUE, '@', 1);
            EnterCell(7, 4, 'Item Category', TRUE, TRUE, '@', 1);
            EnterCell(7, 5, 'Purchase UOM', TRUE, TRUE, '@', 1);
            EnterCell(7, 6, 'Opening Stock', TRUE, TRUE, '@', 1);
            EnterCell(7, 7, 'Produced', TRUE, TRUE, '@', 1);
            EnterCell(7, 8, 'Purchase', TRUE, TRUE, '@', 1);
            EnterCell(7, 9, 'Positive Adj', TRUE, TRUE, '@', 1);
            EnterCell(7, 10, 'Transfer Receipts', TRUE, TRUE, '@', 1);
            EnterCell(7, 11, 'Total Receipt', TRUE, TRUE, '@', 1);
            EnterCell(7, 12, 'Consumption', TRUE, TRUE, '@', 1);
            EnterCell(7, 13, 'Negative Adj', TRUE, TRUE, '@', 1);
            EnterCell(7, 14, 'Issue For Branch TRF', TRUE, TRUE, '@', 1);
            EnterCell(7, 15, 'Purchase Return', TRUE, TRUE, '@', 1);
            EnterCell(7, 16, 'Total Issued', TRUE, TRUE, '@', 1);
            EnterCell(7, 17, 'Closing Stock', TRUE, TRUE, '@', 1);
            EnterCell(7, 18, 'Default Dimension', TRUE, TRUE, '@', 1);

            NN := 8;

        END;

        //<<15Mar2017 ReportHeader
    end;

    var
        //  ExcBuffer: Record 370 temporary;
        // Item: Record Item;
        OPBAL: Decimal;
        QTYMAN: Decimal;
        QTYRECEIVED: Decimal;
        QTYPRODUCED: Decimal;
        QTYISSUED: Decimal;
        QTYISSUED1: Decimal;
        CLOSEBAL: Decimal;
        ItemMaster: Record 27;
        DateFILTER: Text[30];
        LocFilter: Text[30];
        OpenFilter: Text[30];
        SrNo: Integer;
        PosDate: Date;
        CompInfo: Record 79;
        "<For Excel Generation>": Integer;
        MaxCol: Text[3];
        i: Integer;
        "Grand Total": Text[30];
        respcentname: Text[30];
        filtertext: Text[100];
        datefiltertext: Text[250];
        TotalOPBAL: Decimal;
        TotalQTYMAN: Decimal;
        TotalQTYRECEIVED: Decimal;
        TotalQTYISSUED: Decimal;
        TotalCLOSEBAL: Decimal;
        Location: Record 14;
        Loc: Text[30];
        Usersetup: Record 91;
        CompName: Text[100];
        ItemCatFilter: Text[100];
        PrinttoExcel: Boolean;
        ILE: Record 32;
        ManualNegAdj: Decimal;
        QtyIssued2: Decimal;
        PurQty: Decimal;
        TotalPurqty: Decimal;
        TotalReceipt: Decimal;
        PurchRetQty: Decimal;
        TotalIssued: Decimal;
        TotalQTYRECEIVED1: Decimal;
        TotalReceipt1: Decimal;
        TotalIssued1: Decimal;
        TotalQtyIssued1: Decimal;
        TotalPurchRetQty: Decimal;
        recitemCaterogy: Record 5722;
        ItemCatName: Text[50];
        BinFilter: Text[30];
        recBIN: Record 7354;
        BinName: Text[30];
        BinFilterText: Text[30];
        ItemCaterogy: Record 5722;
        ItemCaterogyName: Text[30];
        Others: Decimal;
        Posadj: Decimal;
        Negadj: Decimal;
        TotalPosAdj: Decimal;
        TotalNegAdj: Decimal;
        TotalProduced: Decimal;
        GroupOpBal: Decimal;
        GroupPerQty: Decimal;
        GroupPosAdj: Decimal;
        GroupNegAdj: Decimal;
        GroupProduced: Decimal;
        GroupQtyReceive: Decimal;
        GroupRecep1: Decimal;
        GroupQctIssuid: Decimal;
        GroupQcyIssuid1: Decimal;
        GroupPurchretQty: Decimal;
        groupClosBal: Decimal;
        GroupIssud1: Decimal;
        reportshow: Boolean;
        itemDim: Code[20];
        recDefualtDim: Record 352;
        "----14Mar2017": Integer;
        NN: Integer;
        ExcBuffer: Record 370 temporary;
        Item1403: Record 27;
        TCOUNT: Integer;
        ICOUNT: Integer;
        NAH: Boolean;
        NDate: Date;
        TempItem: Code[20];

    // [Scope('Internal')]
    procedure CreateXLSHEET()
    begin
        /*
        CLEAR(XLAPP);
        CLEAR(XLWBOOK);
        CLEAR(XLSHEET);
        CREATE(XLAPP);
        XLWBOOK := XLAPP.Workbooks.Add;
        XLSHEET := XLWBOOK.Worksheets.Add;
        XLSHEET.Name := 'Stock Statement of RM';
        XLAPP.Visible(TRUE);
        MaxCol:= 'R';
        */

    end;

    //  [Scope('Internal')]
    procedure CreateHeader()
    begin
        /*
        XLSHEET.Activate;
        XLSHEET.Range('A1',MaxCol+'1').Merge := TRUE;
        XLSHEET.Range('A1',MaxCol+'1').Font.Bold := TRUE;
        XLSHEET.Range('A1',MaxCol+'1').Value := CompInfo.Name;
        XLSHEET.Range('A1',MaxCol+'1').HorizontalAlignment := 3;
        XLSHEET.Range('A1',MaxCol+'1').Interior.ColorIndex := 8;
        XLSHEET.Range('A1',MaxCol+'1').Borders.ColorIndex := 5;
        
        XLSHEET.Range('A2',MaxCol+'2').Merge := TRUE;
        
        XLSHEET.Range('A2',MaxCol+'2').Merge := TRUE;
        XLSHEET.Range('A2',MaxCol+'2').Font.Bold := TRUE;
        XLSHEET.Range('A2',MaxCol+'2').Value := filtertext;
        XLSHEET.Range('A2',MaxCol+'2').HorizontalAlignment := 3;
        XLSHEET.Range('A2',MaxCol+'2').Interior.ColorIndex := 8;
        XLSHEET.Range('A2',MaxCol+'2').Borders.ColorIndex := 5;
        
        XLSHEET.Range('B2',MaxCol+'2').Merge := TRUE;
        
        XLSHEET.Range('A4',MaxCol+'4').Merge := TRUE;
        XLSHEET.Range('A4',MaxCol+'4').Font.Bold := TRUE;
        XLSHEET.Range('A4',MaxCol+'4').Value :=datefiltertext;
        
        XLSHEET.Range('A5',MaxCol+'5').Merge := TRUE;
        XLSHEET.Range('A5',MaxCol+'5').Font.Bold := TRUE;
        XLSHEET.Range('A5',MaxCol+'5').Value := BinFilterText;
        
        
        XLSHEET.Range('A6').Value := 'SrNo';
        XLSHEET.Range('A6').Font.Bold := TRUE;
        XLSHEET.Range('A6',MaxCol+'6').Interior.ColorIndex := 45;
        
        XLSHEET.Range('A6',MaxCol+'6').Borders.ColorIndex := 5;
        
        XLSHEET.Range('B6').Value := 'Item Code';
        XLSHEET.Range('B6').Font.Bold := TRUE;
        
        XLSHEET.Range('C6').Value := 'Item Name';
        XLSHEET.Range('C6').Font.Bold := TRUE;
        
        XLSHEET.Range('D6').Value := 'Item Category';
        XLSHEET.Range('D6').Font.Bold := TRUE;
        
        XLSHEET.Range('E6').Value := 'Purchase UOM';
        XLSHEET.Range('E6').Font.Bold := TRUE;
        
        XLSHEET.Range('F6').Value := 'Opening Stock';
        XLSHEET.Range('F6').Font.Bold := TRUE;
        
        XLSHEET.Range('G6').Value := 'Produced';
        XLSHEET.Range('G6').Font.Bold := TRUE;
        
        
        
        XLSHEET.Range('H6').Value := 'Purchase';
        XLSHEET.Range('H6').Font.Bold := TRUE;
        
        XLSHEET.Range('I6').Value := 'Positive Adj';
        XLSHEET.Range('I6').Font.Bold := TRUE;
        
        
        XLSHEET.Range('J6').Value := 'Transfer Receipts';
        XLSHEET.Range('J6').Font.Bold := TRUE;
        
        XLSHEET.Range('K6').Value := 'Total Receipt';
        XLSHEET.Range('K6').Font.Bold := TRUE;
        
        XLSHEET.Range('L6').Value := 'Consumption';
        XLSHEET.Range('L6').Font.Bold := TRUE;
        
        XLSHEET.Range('M6').Value := 'Negative Adj';
        XLSHEET.Range('M6').Font.Bold := TRUE;
        
        
        XLSHEET.Range('N6').Value := 'Issue For Branch TRF';
        XLSHEET.Range('N6').Font.Bold := TRUE;
        
        XLSHEET.Range('O6').Value := 'Purchase Return';
        XLSHEET.Range('O6').Font.Bold := TRUE;
        
        XLSHEET.Range('P6').Value := 'Total Issued';
        XLSHEET.Range('P6').Font.Bold := TRUE;
        
        XLSHEET.Range('Q6').Value := 'Closing Stock';
        XLSHEET.Range('Q6').Font.Bold := TRUE;
        
        //RSPL
        XLSHEET.Range('R6').Value := 'Default Dimension';
        XLSHEET.Range('R6').Font.Bold := TRUE;
        
        //RSPL
        
        XLSHEET.Range('A6:R6').Borders.ColorIndex := 5;
        */

        //>>15Mar2017 ReportHeader

    end;

    // [Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        IF PrinttoExcel THEN BEGIN
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

    local procedure GroupDataBody()
    begin
        //>>15Mar2017 ReportGroupBody
        //IF PrinttoExcel AND reportshow=TRUE AND (TCOUNT = ICOUNT) THEN
        //IF PrinttoExcel AND (TCOUNT = ICOUNT) THEN
        IF PrinttoExcel THEN BEGIN
            EnterCell(NN, 1, '', TRUE, TRUE, '', 0);
            EnterCell(NN, 2, 'Group Total', TRUE, TRUE, '@', 1);
            EnterCell(NN, 3, '', TRUE, TRUE, '@', 1);
            EnterCell(NN, 4, '', TRUE, TRUE, '@', 1);
            EnterCell(NN, 5, '', TRUE, TRUE, '@', 1);
            EnterCell(NN, 6, FORMAT(GroupOpBal), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 7, FORMAT(GroupProduced), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 8, FORMAT(GroupPerQty), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 9, FORMAT(GroupPosAdj), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 10, FORMAT(GroupQtyReceive), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 11, FORMAT(GroupRecep1), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 12, FORMAT(GroupQctIssuid), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 13, FORMAT(GroupNegAdj), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 14, FORMAT(GroupQcyIssuid1), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 15, FORMAT(GroupPurchretQty), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 16, FORMAT(GroupIssud1), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 17, FORMAT(groupClosBal), TRUE, TRUE, '#####0.000', 0);
            EnterCell(NN, 18, itemDim, TRUE, TRUE, '@', 1);
            //EnterCell(NN,19,FORMAT(TCOUNT)+' - '+FORMAT(ICOUNT)+' - '+Item."No."+' - '+Item1403."No."+Item1403."Item Category Code",TRUE,TRUE,'@',1);//test
            //MESSAGE('Srno %1 \ TCount %2 \ ICount %3 \\ No. %4 \\ F Item No. %5',SrNo,TCOUNT,ICOUNT,Item."No.",Item1403."No.");

            NN += 1;

            TCOUNT := 0;


            GroupOpBal := 0;//6
            GroupProduced := 0;//7
            GroupPerQty := 0;//8
            GroupPosAdj := 0;//9
            GroupQtyReceive := 0;//10
            GroupRecep1 := 0;//11
            GroupQctIssuid := 0;//12
            GroupNegAdj := 0;//13
            GroupQcyIssuid1 := 0;//14
            GroupPurchretQty := 0;//15
            GroupIssud1 := 0;//16
            groupClosBal := 0;//17



        END;
        //<<15Mar2017 ReportGroupBody
    end;

    local procedure DataBody()
    begin
        //>>15Mar2017 ReportDataBody


        IF NOT NAH THEN BEGIN

            SrNo := SrNo + 1;

            EnterCell(NN, 1, FORMAT(SrNo), FALSE, FALSE, '', 0);
            EnterCell(NN, 2, Item."No.", FALSE, FALSE, '@', 1);
            EnterCell(NN, 3, Item.Description, FALSE, FALSE, '@', 1);

            ItemCaterogy.RESET;
            ItemCaterogy.SETRANGE(ItemCaterogy.Code, Item."Item Category Code");
            IF ItemCaterogy.FINDFIRST THEN BEGIN
                ItemCaterogyName := ItemCaterogy.Description
            END;

            EnterCell(NN, 4, Item."Item Category Code" + ' - ' + ItemCaterogyName, FALSE, FALSE, '@', 1);
            EnterCell(NN, 5, Item."Base Unit of Measure", FALSE, FALSE, '@', 1);
            EnterCell(NN, 6, FORMAT(OPBAL), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 7, FORMAT(QTYPRODUCED), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 8, FORMAT(PurQty), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 9, FORMAT(Posadj), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 10, FORMAT(QTYRECEIVED), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 11, FORMAT(TotalReceipt), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 12, FORMAT(QTYISSUED), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 13, FORMAT(Negadj), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 14, FORMAT(QTYISSUED1), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 15, FORMAT(PurchRetQty), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 16, FORMAT(TotalIssued), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 17, FORMAT(CLOSEBAL), FALSE, FALSE, '#####0.000', 0);
            EnterCell(NN, 18, itemDim, FALSE, FALSE, '@', 1);

            NN += 1;



        END;
        //<<15Mar2017 ReportDataBody
    end;
}


report 50016 "Location Wise Finished Goods"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/LocationWiseFinishedGoods.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Item Ledger Entry"; "Item Ledger Entry")
        {
            DataItemTableView = SORTING("Location Code", "Item No.", "Lot No.", "Posting Date", "Document No.", "Document Line No.")
                                ORDER(Ascending)
                                WHERE("Item Category Code" = FILTER('CAT02 | CAT03 | CAT11 | CAT12 | CAT13 | CAT15 | CAT18 | CAT20 | CAT22'));
            RequestFilterFields = "Item No.", "Lot No.", "Item Category Code", "Location Code";

            trigger OnAfterGetRecord()
            begin
                ILE += 1;//--17Feb2017 RB-N
                LCOUNT := COUNT;//--17Feb2017 RB-N

                TempGroup += 1;//--17Feb2017 RB-N,Group No.of Item Increment
                TempLoc += 1;//--17Feb2017 RB-N,Group No.of Item with Location Increment
                             //
                             //CurrReport.CREATETOTALS(Qty);
                             //CurrReport.CREATETOTALS(BARNo);

                //EBT STIVAN ---(09062012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
                /*
                User.GET(USERID);
                Memberof.RESET;
                Memberof.SETRANGE(Memberof."User ID",User."User ID");
                Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
                IF NOT(Memberof.FINDFIRST) THEN
                BEGIN
                 CLEAR(LocResC);
                 recLocation.RESET;
                 recLocation.SETRANGE(recLocation.Code,"Item Ledger Entry"."Location Code");
                 IF recLocation.FINDFIRST THEN
                 BEGIN
                  LocResC := recLocation."Global Dimension 2 Code";
                 END;
                
                 CSOMapping1.RESET;
                 CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
                 CSOMapping1.SETRANGE(Type,CSOMapping1.Type::"1");
                 CSOMapping1.SETRANGE(CSOMapping1.Value,"Item Ledger Entry"."Location Code");
                 IF NOT(CSOMapping1.FINDFIRST) THEN
                 BEGIN
                  CSOMapping1.RESET;
                  CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
                  CSOMapping1.SETRANGE(CSOMapping1.Type,CSOMapping1.Type::"0");
                  CSOMapping1.SETRANGE(CSOMapping1.Value,LocResC);
                  IF NOT(CSOMapping1.FINDFIRST) THEN
                  ERROR ('You are not allowed to run this report other than your location');
                 END;
                END;
                //EBT STIVAN ---(09062012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
                
                *///Commented Because User_Table & Member_Table not present in NAV2016

                //>>2

                //IF CurrReport.TOTALSCAUSEDBY = "Item Ledger Entry".FIELDNO("Lot No.") THEN
                //BEGIN
                //CurrReport.SHOWOUTPUT(TRUE);
                Itemrec.RESET;//17Feb2017
                Itemrec.SETRANGE(Itemrec."No.", "Item No.");
                IF Itemrec.FINDFIRST THEN
                    ItmDesc := Itemrec.Description;

                Mfgdte := 20120331D;
                ILErec3.RESET;
                ILErec3.SETCURRENTKEY("Item No.", "Lot No.");//26Apr2019
                ILErec3.SETRANGE(ILErec3."Item No.", "Item Ledger Entry"."Item No.");
                ILErec3.SETRANGE(ILErec3."Lot No.", "Item Ledger Entry"."Lot No.");
                ILErec3.SETRANGE(ILErec3."Posting Date", 0D, filterdate);
                ILErec3.SETRANGE(ILErec3."Entry Type", ILErec3."Entry Type"::Output);
                IF ILErec3.FIND('-') THEN
                    Mfgdte := ILErec3."Posting Date";


                CLEAR(UOM1);
                recItem1.SETRANGE(recItem1."No.", "Item Ledger Entry"."Item No.");
                IF recItem1.FINDFIRST THEN BEGIN
                    UOM1 := recItem1."Sales Unit of Measure";

                    CLEAR(ItemUOM1);
                    recItemUOM1.RESET;
                    recItemUOM1.SETRANGE(recItemUOM1."Item No.", "Item Ledger Entry"."Item No.");
                    recItemUOM1.SETRANGE(recItemUOM1.Code, UOM1);
                    IF recItemUOM1.FINDFIRST THEN
                        ItemUOM1 := recItemUOM1."Qty. per Unit of Measure";
                END;


                CLEAR(MRPprice);
                CLEAR(SalesPrice);
                CLEAR(MRPMasterPostingDate);
                CLEAR(MRPDATE);
                CLEAR(NDPD);
                recMRPMaster.RESET;
                recMRPMaster.SETRANGE(recMRPMaster."Item No.", "Item Ledger Entry"."Item No.");
                recMRPMaster.SETRANGE(recMRPMaster."Lot No.", "Item Ledger Entry"."Lot No.");
                IF recMRPMaster.FINDFIRST THEN BEGIN
                    MRPprice := recMRPMaster."MRP Price";
                    SalesPrice := recMRPMaster."Sales price" / ItemUOM1;
                    MRPMasterPostingDate := recMRPMaster."Posting Date";


                    IF MRPMasterPostingDate <> 20120331D THEN BEGIN
                        recSP.RESET;
                        recSP.SETRANGE(recSP."Item No.", "Item Ledger Entry"."Item No.");
                        //recSP.SETRANGE(recSP."MRP Price", MRPprice);
                        recSP.SETFILTER(recSP."Starting Date", '<=%1', MRPMasterPostingDate);
                        IF recSP.FINDLAST THEN
                            MRPDATE := recSP."Starting Date";
                    END;
                END;

                IF RecMRPMaster1.GET("Item Ledger Entry"."Item No.", "Item Ledger Entry"."Lot No.") THEN
                    NDPD := RecMRPMaster1."National Discount"; //RSPLAM

                Qty := 0;
                ILErec.RESET;
                ILErec.SETCURRENTKEY("Item No.", "Location Code", "Posting Date");//26Apr2019
                ILErec.SETRANGE(ILErec."Item No.", "Item Ledger Entry"."Item No.");
                ILErec.SETRANGE(ILErec."Location Code", "Item Ledger Entry"."Location Code");
                ILErec.SETRANGE(ILErec."Lot No.", "Item Ledger Entry"."Lot No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', 0D, filterdate);
                /*
                IF ILErec.FIND('-') THEN
                REPEAT
                 Qty += ILErec.Quantity;
                UNTIL ILErec.NEXT=0;
                */
                //>>26Apr2019
                IF ILErec.FINDFIRST THEN BEGIN
                    ILErec.CALCSUMS(Quantity);
                    Qty := ILErec.Quantity;
                END;
                //<<26Apr2019
                remarks := '';
                recItemTable.RESET;
                recItemTable.SETRANGE(recItemTable."No.", "Item Ledger Entry"."Item No.");
                recItemTable.SETRANGE(recItemTable.Blocked, TRUE);
                IF recItemTable.FINDFIRST THEN BEGIN
                    IF Qty <> 0 THEN BEGIN
                        remarks := 'Blocked';
                    END;
                END;

                IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Sales Return Receipt" THEN
                    remarks := 'Rejection';

                /*
                ReceiptDate := 0D;
                recILE.RESET;
                recILE.SETRANGE(recILE."Item No.","Item Ledger Entry"."Item No.");
                recILE.SETRANGE(recILE."Location Code","Item Ledger Entry"."Location Code");
                recILE.SETRANGE(recILE."Lot No.","Item Ledger Entry"."Lot No.");
                recILE.SETRANGE(recILE.Open,TRUE);
                IF recILE.FINDFIRST THEN
                BEGIN
                 ReceiptDate := recILE."Posting Date";
                END;
                */
                //RSPL-Sourav
                ReceiptDate := 0D;
                IF ReceiptDate = 0D THEN BEGIN
                    recILE.RESET;
                    recILE.SETCURRENTKEY("Item No.", "Location Code", "Posting Date");//26Apr2019
                    recILE.SETRANGE(recILE."Item No.", "Item Ledger Entry"."Item No.");
                    recILE.SETRANGE(recILE."Location Code", "Item Ledger Entry"."Location Code");
                    recILE.SETRANGE(recILE."Lot No.", "Item Ledger Entry"."Lot No.");
                    recILE.SETFILTER(recILE."Posting Date", '%1..%2', 0D, filterdate);
                    recILE.SETFILTER(recILE.Quantity, '>%1', 0);
                    IF recILE.FINDLAST THEN BEGIN
                        ReceiptDate := recILE."Posting Date";
                    END;
                END;
                //RSPL-Sourav


                //EBT STIVAN---(12112012)---To capture BRT No and Date as well as Excise Details ----------------------START
                CLEAR("TRONo.");
                CLEAR(TSNo);
                CLEAR(TSHNo);
                CLEAR(TSDate);
                CLEAR(ExciseRate);
                CLEAR(BEDAmount);
                CLEAR(EcessAmount);
                CLEAR(SheCessAmount);

                recILE.RESET;
                recILE.SETCURRENTKEY("Item No.", "Location Code", "Posting Date");//26Apr2019
                recILE.SETRANGE(recILE."Item No.", "Item Ledger Entry"."Item No.");
                recILE.SETRANGE(recILE."Location Code", "Item Ledger Entry"."Location Code");
                recILE.SETRANGE(recILE."Lot No.", "Item Ledger Entry"."Lot No.");
                recILE.SETRANGE(recILE.Open, TRUE);
                IF recILE.FINDFIRST THEN BEGIN
                    //"TRONo." := recILE."Transfer Order No.";
                    "TRONo." := recILE."Order No.";//26Apr2019
                    TSNo := recILE."Document No.";
                    TSDate := recILE."Posting Date";

                    IF recILE."Document Type" = recILE."Document Type"::"Transfer Receipt" THEN BEGIN
                        recTSH.RESET;
                        recTSH.SETRANGE(recTSH."Transfer Order No.", "TRONo.");
                        IF recTSH.FINDFIRST THEN BEGIN
                            TSNo := recTSH."No.";
                            TSDate := recTSH."Posting Date";
                            TSHNo := recTSH."No.";
                        END;

                        recTSL.RESET;
                        recTSL.SETRANGE(recTSL."Document No.", TSHNo);
                        recTSL.SETRANGE(recTSL."Item No.", recILE."Item No.");
                        recTSL.SETRANGE(recTSL."Line No.", recILE."Document Line No.");
                        IF recTSL.FIND('-') THEN BEGIN
                            ExciseRate := 0;//recTSL."Excise Effective Rate";
                            BEDAmount := 0;//recTSL."BED Amount";
                            EcessAmount := 0;//recTSL."eCess Amount";
                            SheCessAmount := 0;// recTSL."SHE Cess Amount";
                        END;
                    END;
                END;
                //EBT STIVAN---(12112012)---To capture BRT No and Date as well as Excise Details ------------------------END


                ItemUOM.GET(Itemrec."No.", Itemrec."Sales Unit of Measure");

                BARNo := Qty / ItemUOM."Qty. per Unit of Measure";
                TotalPackQty += BARNo;
                /*
                ExportToExcel := TRUE;
                IF "Item Ledger Entry".Quantity = 0 THEN
                IF ShowQty THEN
                  ExportToExcel := TRUE
                ELSE
                  ExportToExcel := FALSE;
                
                *///oldCodeCommented For Show qty with 0.
                  //IF ShowQty = FALSE THEN
                  //BEGIN
                  //IF BARNo = 0 THEN
                  //CurrReport.SHOWOUTPUT(FALSE);
                  //END;
                  //IF CurrReport.SHOWOUTPUT THEN//RSPL-Sourav
                vGTotalQty += Quantity;//RSPL-Sourav
                                       //IF ExportToExcel AND CurrReport.SHOWOUTPUT THEN BEGIN

                //--17Feb2017 RB-N Grouping with ItemNo,Location,LotNo.
                GItemCode := "Item No.";
                GLCode := "Location Code";
                GLotNo := "Lot No.";

                //IF ( GItemCode <> "Item No.") AND (GLCode <> "Location Code") AND (GLotNo <> "Lot No.") THEN
                //--- 17Feb2017 RB-N, Group No.of Item
                ILE32.RESET;
                ILE32.SETCURRENTKEY("Item No.", "Location Code", "Posting Date");
                ILE32.SETRANGE("Location Code", "Location Code");
                ILE32.SETRANGE("Item No.", "Item No.");
                ILE32.SETRANGE("Lot No.", "Lot No.");
                ILE32.SETFILTER("Posting Date", '%1..%2', 0D, filterdate);
                IF ILE32.FINDFIRST THEN BEGIN
                    CountGroup := ILE32.COUNT;//Group No.of Item
                END;
                //---

                //----ShowQty with 0 18Feb2017
                IF ShowQty THEN BEGIN
                    IF Qty = 0 THEN BEGIN
                        ZeroQty := TRUE;
                    END;

                    IF Qty <> 0 THEN BEGIN
                        ZeroQty := TRUE;
                    END;

                END;

                IF ShowQty = FALSE THEN BEGIN
                    IF Qty = 0 THEN BEGIN
                        ZeroQty := FALSE;
                        TempGroup := 0;
                    END;

                    IF Qty <> 0 THEN BEGIN
                        ZeroQty := TRUE;
                    END;

                END;

                //RSPLSUM BEGIN>>
                CLEAR(Zone);
                IF "Item Ledger Entry"."Entry Type" = "Item Ledger Entry"."Entry Type"::Sale THEN BEGIN
                    IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Sales Shipment" THEN BEGIN
                        RecSSH.RESET;
                        IF RecSSH.GET("Item Ledger Entry"."Document No.") THEN BEGIN
                            RecSalespersonPurchser.RESET;
                            IF RecSalespersonPurchser.GET(RecSSH."Salesperson Code") THEN
                                Zone := RecSalespersonPurchser."Zone Code";
                        END;
                    END;
                    IF "Item Ledger Entry"."Document Type" = "Item Ledger Entry"."Document Type"::"Sales Return Receipt" THEN BEGIN
                        RecRRH.RESET;
                        IF RecRRH.GET("Item Ledger Entry"."Document No.") THEN BEGIN
                            RecSalespersonPurchser.RESET;
                            IF RecSalespersonPurchser.GET(RecRRH."Salesperson Code") THEN
                                Zone := RecSalespersonPurchser."Zone Code";
                        END;
                    END;
                END;
                //RSPLSUM END<<

                //MESSAGE('LotNo. %1\ Item No. %2 \ Qty %3 \ZeroQty %4\Temp Group %5\CountGroup %6',"Lot No.","Item No.",Qty,ZeroQty,TempGroup,CountGroup);
                //----ShowQty

                //--17Feb2017 RB-N
                //IF TempGroup = CountGroup THEN //
                IF (TempGroup = CountGroup) AND ZeroQty THEN //
                BEGIN

                    ExportToExcel := TRUE;//17Feb2017
                    IF ExportToExcel THEN BEGIN
                        SrNo += 1;
                        EnterCell(k + 1, 1, FORMAT(SrNo), TRUE, FALSE, '', 0);
                        EnterCell(k + 1, 2, FORMAT("Item No."), FALSE, FALSE, '@', 1);
                        EnterCell(k + 1, 3, FORMAT(ItmDesc), FALSE, FALSE, '@', 1);
                        EnterCell(k + 1, 4, FORMAT(MRPprice), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(k + 1, 5, FORMAT(MRPprice * "Qty. per Unit of Measure"), FALSE, FALSE, '#,####0.00', 0);
                        EnterCell(k + 1, 6, FORMAT(MRPDATE), FALSE, FALSE, '', 2);
                        EnterCell(k + 1, 7, FORMAT(SalesPrice), FALSE, FALSE, '#,####0', 0);
                        EnterCell(k + 1, 8, FORMAT(NDPD), FALSE, FALSE, '#,####0', 0);
                        EnterCell(k + 1, 9, FORMAT(Mfgdte), FALSE, FALSE, '', 2);
                        EnterCell(k + 1, 10, FORMAT("Item Category Code"), FALSE, FALSE, '@', 1);
                        EnterCell(k + 1, 11, FORMAT("Lot No."), FALSE, FALSE, '@', 1);
                        EnterCell(k + 1, 12, FORMAT("Unit of Measure Code"), FALSE, FALSE, '@', 1);
                        EnterCell(k + 1, 13, FORMAT("Qty. per Unit of Measure"), FALSE, FALSE, '', 0);
                        EnterCell(k + 1, 14, FORMAT(BARNo), FALSE, FALSE, '#,####0', 0);
                        EnterCell(k + 1, 15, FORMAT(Qty), FALSE, FALSE, '#,####0', 0);
                        EnterCell(k + 1, 16, FORMAT("Item Ledger Entry"."Location Code"), FALSE, FALSE, '@', 1);

                        recloc.RESET;
                        recloc.SETRANGE(recloc.Code, "Item Ledger Entry"."Location Code");
                        IF recloc.FINDFIRST THEN
                            EnterCell(k + 1, 17, FORMAT(recloc.Name), FALSE, FALSE, '@', 1);

                        IF ReceiptDate = 0D THEN BEGIN
                            EnterCell(k + 1, 18, FORMAT(''), FALSE, FALSE, '', 0);
                            //   EnterCell(k+1,18,FORMAT(''),FALSE,FALSE);
                        END ELSE BEGIN
                            EnterCell(k + 1, 18, FORMAT(ReceiptDate), FALSE, FALSE, '', 2);
                            //  EnterCell(k+1,18,FORMAT(-(ReceiptDate-filterdate)),FALSE,FALSE);
                        END;
                        EnterCell(k + 1, 19, FORMAT(remarks), FALSE, FALSE, '@', 1);
                        EnterCell(k + 1, 20, FORMAT(Zone), FALSE, FALSE, '@', 1);//RSPLSUM


                        k += 1;
                        GroupQty := GroupQty + Qty;//Group of Item
                        LocQty := LocQty + Qty;  //Group of Item in a Location
                        TootalQty := TootalQty + Qty; //Group of Total Qty
                        GroupBarNo := GroupBarNo + BARNo; //Group of ItemBarNo
                        LocBarNo := LocBarNo + BARNo;//Group of ItemBarNo in Location
                        TotalBARNo := TotalBARNo + BARNo;//Group Of Total BarNo
                                                         //NAH := NAH + TempGroup;
                                                         //MESSAGE('NAH %1 \ SrNo %2',NAH,SrNo);//
                        TempGroup := 0;
                    END;

                END;
                //END;
                //END
                //ELSE
                //CurrReport.SHOWOUTPUT(FALSE);
                //<<2


                //>>3

                //IF CurrReport.TOTALSCAUSEDBY = "Item Ledger Entry".FIELDNO("Location Code") THEN
                //BEGIN
                //CurrReport.SHOWOUTPUT :=TRUE;

                //--- 17Feb2017 RB-N, Group No.of Item in a Location
                //ILE32.RESET;
                //ILE32.SETCURRENTKEY("Location Code","Item No.","Lot No.","Posting Date","Document No.","Document Line No.");
                //ILE32.SETRANGE("Location Code","Location Code");
                //ILE32.SETRANGE("Item No.","Item No.");//Comm
                //ILE32.SETRANGE("Lot No.","Lot No.");//Coomm
                //ILE32.SETFILTER("Item Category Code",'%1|%2|%3|%4|%5|%6','CAT02','CAT03','CAT12','CAT13','CAT15');
                //ILE32.SETFILTER("Posting Date",'%1..%2',0D,filterdate);
                //IF ILE32.FINDFIRST THEN
                //BEGIN
                //CountLoc := ILE32.COUNT;//Group No.of Item in a Location
                //END;
                //---
                //MESSAGE('Templocation %1 \CountLocation %2 \ TotalQty %3',TempLoc,CountLoc,LCOUNT);
                //-----Location Total

                //-----Location Total 18Feb2017
                //ILE3218.RESET;
                ILE3218.SETCURRENTKEY("Location Code", "Item No.", "Lot No.", "Posting Date", "Document No.", "Document Line No.");
                ILE3218.SETRANGE("Location Code", "Location Code");
                IF ILE3218.FINDFIRST THEN BEGIN
                    CountLoc := ILE3218.COUNT;
                    //MESSAGE('Templocation %1 \CountLocation %2 \ TotalQty %3',TempLoc,CountLoc,LCOUNT);
                END;
                //-----Location Total 18Feb2017
                //----ShowSubTotal with 0 18Feb2017
                /*
                IF ShowSubtot THEN
                BEGIN
                
                  IF ShowQty THEN
                  BEGIN
                    IF LocQty =  0 THEN
                    BEGIN
                      ZeroQty := TRUE;
                    END;
                
                    IF LocQty <> 0 THEN
                    BEGIN
                      ZeroQty := TRUE;
                    END;
                
                  END;
                
                  IF ShowQty = FALSE THEN
                  BEGIN
                
                    IF LocQty =  0 THEN
                    BEGIN
                      ZeroQty := FALSE;
                      TempLoc := 0;
                    END;
                
                    IF LocQty <> 0 THEN
                    BEGIN
                      ZeroQty := TRUE;
                    END;
                
                  END;
                
                END;
                *///commented
                  //MESSAGE('LotNo. %1\ Item No. %2 \ Qty %3 \ZeroQty %4\Temp Group %5\CountGroup %6',"Lot No.","Item No.",Qty,ZeroQty,TempGroup,CountGroup);
                  //----ShowQty

                IF ShowSubtot THEN
                //IF ShowSubtot AND ZeroQty THEN
                BEGIN
                    IF TempLoc = CountLoc THEN
                    //IF TempLoc = 486 THEN //Testing
                    //IF TempLoc = NAH THEN
                    BEGIN
                        ExportToExcel := TRUE;
                        IF ExportToExcel THEN BEGIN
                            EnterCell(k + 1, 1, 'Total', TRUE, FALSE, '@', 1);
                            EnterCell(k + 1, 2, FORMAT("Location Code"), TRUE, FALSE, '@', 1);
                            //EnterCell(k+1,13,FORMAT(TotalPackQty),TRUE,FALSE);
                            EnterCell(k + 1, 13, FORMAT(LocBarNo), TRUE, FALSE, '#,####0', 0);//17Feb2017
                                                                                              //EnterCell(k+1,14,FORMAT(vGTotalQty),TRUE,FALSE); //RSPL-SOURAV
                            EnterCell(k + 1, 14, FORMAT(LocQty), TRUE, FALSE, '#,####0', 0); //17Feb2017
                            k += 1;
                            SrNo := 0;
                            LocQty := 0;//17Feb2017
                            LocBarNo := 0;//17Feb2017
                            TempLoc := 0;//17Feb2017
                            TotalQty += TotalPackQty;//17Feb2017 RB-N
                            vFTotalQty += vGTotalQty;//17Feb2017 RB-N
                                                     //NAH := 0;//17Feb2017
                        END;
                    END;
                END;
                //------
                //SubTotal Commented
                //ELSE BEGIN
                //   ExportToExcel := FALSE;
                //   SrNo := 0;
                //  END
                //END
                //ELSE
                //CurrReport.SHOWOUTPUT :=FALSE;

                //IF CurrReport.SHOWOUTPUT THEN BEGIN
                //TotalQty+=TotalPackQty;
                //vFTotalQty+=vGTotalQty;
                //END;
                //>>3

                //>>4
                //-----Grand Total
                ExportToExcel := TRUE;//--17Feb2017 RB-N
                IF ILE = LCOUNT THEN BEGIN
                    IF ExportToExcel THEN BEGIN
                        EnterCell(k + 1, 1, 'Grand Total', TRUE, FALSE, '@', 1);
                        //EnterCell(k+1,13,FORMAT(TotalQty),TRUE,FALSE);
                        EnterCell(k + 1, 13, FORMAT(TotalBARNo), TRUE, FALSE, '#,####0', 0);//17Feb2017
                                                                                            //EnterCell(k+1,14,FORMAT(vFTotalQty),TRUE,FALSE); //RSPL-Sourav
                        EnterCell(k + 1, 14, FORMAT(TootalQty), TRUE, FALSE, '#,####0', 0); //17Feb2017
                        k += 1;

                    END;
                END;//--
                //------
                //<<4

            end;

            trigger OnPreDataItem()
            begin
                //>>1

                IF filterdate <> 0D THEN BEGIN
                    "Item Ledger Entry".SETFILTER("Item Ledger Entry"."Posting Date", '%1..%2', 0D, filterdate);
                END
                ELSE
                    ERROR('Please Enter Date');
                //<<1

                ILE := 0;//--17Feb2017 RB-N
                SrNo := 0;//--17Feb2017 RB-N
                TempGroup := 0;//--17Feb2017 RB-N,Group No.of Item
                TempLoc := 0;//--17Feb2017 RB-N,Group No.of Item with Location

                //Testing Code 18Feb2017
                ILE3218.RESET;
                ILE3218.COPYFILTERS("Item Ledger Entry");
                //ILE32.COPYFILTER("Location Code","Location Code");
                //ILE32.COPYFILTER("Posting Date","Posting Date");
                //MESSAGE('No. of Count %1',ILE3218.COUNT);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Enter Date"; filterdate)
                {
                }
                field("Show Sub Total"; ShowSubtot)
                {
                }
                field("Show Qty with 0"; ShowQty)
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
        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN
            // ExcBuffer.CreateBook('', 'Location Wise Finished Goods');
            // ExcBuffer.CreateBookAndOpenExcel('', 'Location Wise Finished Goods', '', '', USERID);
            // ExcBuffer.GiveUserControl;

            ExcBuffer.CreateNewBook('Location Wise Finished Goods');
            ExcBuffer.WriteSheet('Location Wise Finished Goods', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Location Wise Finished Goods', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();
        END;
    end;

    trigger OnPreReport()
    begin

        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN

            ExcBuffer.DELETEALL;
            CLEAR(ExcBuffer);
            //EnterCell(1,1,FORMAT(COMPANYNAME+'-'+LocCode),TRUE,FALSE);

            //EnterCell(1,1,FORMAT(COMPANYNAME),TRUE,FALSE);
            EnterCell(1, 1, FORMAT(COMPANYNAME), TRUE, FALSE, '@', 1);
            EnterCell(2, 1, 'Location Wise Finished Goods', TRUE, FALSE, '@', 1);
            EnterCell(1, 18, FORMAT(FORMAT(TODAY, 0, 4)), TRUE, FALSE, '@', 1);
            EnterCell(2, 18, FORMAT(USERID), TRUE, FALSE, '@', 1);
            EnterCell(4, 1, ' FILTER : ', TRUE, FALSE, '@', 1);
            EnterCell(4, 2, (FORMAT(filterdate)), TRUE, FALSE, '@', 1);
            EnterCell(6, 1, 'Sr No.', TRUE, FALSE, '@', 1);
            EnterCell(6, 2, 'FG Code', TRUE, FALSE, '@', 1);
            EnterCell(6, 3, 'Item Description', TRUE, FALSE, '@', 1);
            EnterCell(6, 4, 'MRP PRICE', TRUE, FALSE, '@', 1);
            EnterCell(6, 5, 'MRP PRICE per Ltr', TRUE, FALSE, '@', 1);
            EnterCell(6, 6, 'MRP Date', TRUE, FALSE, '@', 1);
            EnterCell(6, 7, 'SALES PRICE', TRUE, FALSE, '@', 1);
            EnterCell(6, 8, 'ND/PD', TRUE, FALSE, '@', 1);//RSPLAM29072021
            EnterCell(6, 9, 'Mfg. Date', TRUE, FALSE, '@', 1);
            EnterCell(6, 10, 'Item Category Code', TRUE, FALSE, '@', 1);
            EnterCell(6, 11, 'Batch No.', TRUE, FALSE, '@', 1);
            EnterCell(6, 12, 'Pack Type', TRUE, FALSE, '@', 1);
            EnterCell(6, 13, 'Pack UOM', TRUE, FALSE, '@', 1);
            EnterCell(6, 14, 'No. of BAR/PAILS', TRUE, FALSE, '@', 1);
            EnterCell(6, 15, 'Remaining Quantity in LTRS/KGS', TRUE, FALSE, '@', 1);
            EnterCell(6, 16, 'Location Code', TRUE, FALSE, '@', 1);
            EnterCell(6, 17, 'Location Name', TRUE, FALSE, '@', 1);
            EnterCell(6, 18, 'Receipt Date', TRUE, FALSE, '@', 1);
            //   EnterCell(6,18,'No. of Days',TRUE,FALSE);
            EnterCell(6, 19, 'Remarks', TRUE, FALSE, '@', 1);
            EnterCell(6, 20, 'Zone', TRUE, FALSE, '@', 1);//RSPLSUM

            k := +6;
        END;
    end;

    var
        filterdate: Date;
        SrNo: Integer;
        Itemrec: Record 27;
        ItmDesc: Text[50];
        ILErec: Record 32;
        Mfgdte: Date;
        BARNo: Decimal;
        TBARNO: Decimal;
        Qty: Decimal;
        TQty: Decimal;
        TotnoOfBar: Decimal;
        ExportToExcel: Boolean;
        ExcBuffer: Record 370 temporary;
        k: Integer;
        ShowSubtot: Boolean;
        Locrec: Record 14;
        LocCode: Code[10];
        ilerec1: Record 32;
        RefDate: Date;
        Locrec1: Record 14;
        ShowQty: Boolean;
        ILErec3: Record 32;
        ItemUOM: Record 5404;
        TotalPackQty: Decimal;
        TotalQty: Decimal;
        CSOMapping: Record 50006;
        CSOMapping2: Record 50006;
        CSOMapping1: Record 50006;
        recLocation: Record 14;
        LocResC: Code[10];
        recMRPMaster: Record 50013;
        MRPprice: Decimal;
        SalesPrice: Decimal;
        UOM1: Code[10];
        recItem1: Record 27;
        ItemUOM1: Decimal;
        recItemUOM1: Record 5404;
        MRPMasterPostingDate: Date;
        recSP: Record 7002;
        MRPDATE: Date;
        recloc: Record 14;
        recILE: Record 32;
        "TRONo.": Code[20];
        recTSH: Record 5744;
        TSNo: Code[20];
        TSHNo: Code[20];
        TSDate: Date;
        ExciseDetails: Boolean;
        recTSL: Record 5745;
        ExciseRate: Decimal;
        BEDAmount: Decimal;
        EcessAmount: Decimal;
        SheCessAmount: Decimal;
        recItemTable: Record 27;
        remarks: Text[30];
        ReceiptDate: Date;
        vGTotalQty: Decimal;
        vFTotalQty: Decimal;
        recItemLedgrEntry: Record 32;
        Remrks: Text[30];
        "--17Feb2017": Integer;
        ILE: Integer;
        LCOUNT: Integer;
        GItemCode: Code[20];
        GLCode: Code[20];
        GLotNo: Code[20];
        ILE32: Record 32;
        TempGroup: Integer;
        CountGroup: Integer;
        TempLoc: Integer;
        CountLoc: Integer;
        GroupQty: Decimal;
        LocQty: Decimal;
        TootalQty: Decimal;
        GroupBarNo: Decimal;
        LocBarNo: Decimal;
        TotalBARNo: Decimal;
        NAH: Integer;
        ILE3218: Record 32;
        ZeroQty: Boolean;
        RecSalespersonPurchser: Record 13;
        Zone: Code[10];
        RecSSH: Record 110;
        RecRRH: Record 6660;
        NDPD: Decimal;
        RecMRPMaster1: Record 50013;

    // //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN
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


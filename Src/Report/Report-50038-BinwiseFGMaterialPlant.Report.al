report 50038 "Bin-wise FG Material - Plant"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 03Apr2018   RB-N         Bulk Code & Bulk Description
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BinwiseFGMaterialPlant.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem(Item; Item)
        {
            RequestFilterFields = "No.";
            dataitem("Warehouse Entry"; "Warehouse Entry")
            {
                DataItemLink = "Item No." = FIELD("No.");
                DataItemTableView = SORTING("Item No.", "Lot No.", "Bin Code")
                                    ORDER(Ascending);

                trigger OnAfterGetRecord()
                begin
                    //MESSAGE('in whse ');
                    CLEAR(BulkDesc);//RSPLSid 02Jan2020

                    ICOUNT += 1;//07Apr2017

                    GQtyBase += "Warehouse Entry"."Qty. (Base)";//07Apr2017
                    TQtyBase += "Warehouse Entry"."Qty. (Base)";//07Apr2017

                    ItmCat := '';
                    Itemrec1.RESET;
                    Itemrec1.SETRANGE(Itemrec1."No.", "Warehouse Entry"."Item No.");
                    Itemrec1.SETRANGE(Itemrec1."Location Filter", "Warehouse Entry"."Location Code");
                    Itemrec1.SETRANGE(Itemrec1."Date Filter", "Warehouse Entry"."Registering Date");

                    IF Itemrec1.FINDFIRST THEN BEGIN
                        itmDesc := Itemrec1.Description;
                        ItmCat := Itemrec1."Item Category Code";
                        SalesUOM := Itemrec1."Sales Unit of Measure";
                    END;

                    CLEAR(EBTBinCode);
                    IF bincode <> '' THEN BEGIN
                        recBinContent.RESET;
                        recBinContent.SETCURRENTKEY("Item No.");//07Apr2017
                                                                //>>Stranger"Location Code","Bin Code","Item No.","Variant Code","Unit of Measure Code"
                        recBinContent.SETRANGE(recBinContent."Item No.", "Warehouse Entry"."Item No.");
                        recBinContent.SETRANGE(recBinContent."Location Code", "Warehouse Entry"."Location Code");
                        recBinContent.SETFILTER(recBinContent."Bin Code", bincode);

                        IF recBinContent.FINDFIRST THEN BEGIN
                            EBTBinCode := recBinContent."Bin Code";
                        END;
                    END;

                    IF bincode = '' THEN BEGIN
                        recBinContent.RESET;
                        recBinContent.SETCURRENTKEY("Item No.");//07Apr2017
                        recBinContent.SETRANGE(recBinContent."Item No.", "Warehouse Entry"."Item No.");
                        recBinContent.SETRANGE(recBinContent."Location Code", "Warehouse Entry"."Location Code");
                        recBinContent.SETFILTER(recBinContent.Quantity, '<>%1', 0);
                        IF recBinContent.FINDFIRST THEN BEGIN
                            IF recBinContent.Quantity = "Warehouse Entry"."Qty. (Base)" THEN BEGIN
                                EBTBinCode := recBinContent."Bin Code";
                            END;
                        END;
                    END;

                    BinName := '';
                    recBin.RESET;
                    recBin.SETCURRENTKEY(Code);//07Apr2017
                    recBin.SETRANGE(recBin.Code, "Warehouse Entry"."Bin Code");
                    IF recBin.FINDFIRST THEN BEGIN
                        BinName := recBin.Description;
                    END;

                    // EBT MILAN (141013)...To Show Blocked Item.................................START
                    // remarks:='';
                    recItemTable.RESET;
                    recItemTable.SETRANGE(recItemTable."No.", "Warehouse Entry"."Item No.");
                    recItemTable.SETRANGE(recItemTable.Blocked, TRUE);
                    IF recItemTable.FINDFIRST THEN
                        // BEGIN
                        //  IF Qty <> 0 THEN
                        //  BEGIN
                        remarks := 'Blocked'
                    ELSE
                        remarks := '';
                    //  END;
                    // END;
                    // EBT MILAN (141013)...To Show Blocked Item...................................END

                    //>>07Apr2017 NewGrouping
                    WHEntry07APR.SETCURRENTKEY("Item No.", "Lot No.", "Bin Code");
                    WHEntry07APR.SETRANGE("Item No.", "Item No.");
                    WHEntry07APR.SETRANGE("Lot No.", "Lot No.");
                    WHEntry07APR.SETRANGE("Bin Code", "Bin Code");
                    IF WHEntry07APR.FINDLAST THEN BEGIN

                        GCOUNT := WHEntry07APR.COUNT;

                    END;
                    //<<07Apr2017 NewGrouping


                    //>>2

                    //Warehouse Entry, GroupFooter (1) - OnPreSection()
                    //IF CurrReport.TOTALSCAUSEDBY="Warehouse Entry".FIELDNO("Warehouse Entry"."Lot No.") THEN
                    IF ICOUNT = GCOUNT THEN BEGIN
                        //CurrReport.SHOWOUTPUT(TRUE);

                        Qty += "Warehouse Entry"."Qty. (Base)";
                        Qtytotal := "Warehouse Entry"."Qty. (Base)";
                        // TotalQtyinUOM := TotalQtyinUOM + (Qty / ItemUOM); EBT MILAN 040413
                        // TotalQtyinUOM := TotalQtyinUOM + (Qtytotal / ItemUOM);

                        Mfgdte := 0D;
                        BlendorderNo := '';
                        ILErec.RESET;
                        ILErec.SETCURRENTKEY("Item No.", "Entry Type", "Variant Code", "Lot No.", "Drop Shipment", "Location Code", "Posting Date");
                        ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                        ILErec.SETRANGE(ILErec."Lot No.", "Warehouse Entry"."Lot No.");
                        ILErec.SETRANGE(ILErec."Location Code", LocCode);
                        ILErec.SETRANGE(ILErec."Posting Date", 0D, filterdate);
                        ILErec.SETRANGE(ILErec."Entry Type", ILErec."Entry Type"::Output);
                        IF ILErec.FIND('-') THEN BEGIN
                            Mfgdte := ILErec."Posting Date";
                            BlendorderNo := ILErec."Document No.";
                            //>>03Apr2018
                            BulkDesc := '';
                            VE03.RESET;
                            VE03.SETCURRENTKEY("Document No.");
                            VE03.SETRANGE("Document No.", BlendorderNo);
                            VE03.SETRANGE("Source No.", "Item No.");
                            VE03.SETRANGE("Item Ledger Entry Type", VE03."Item Ledger Entry Type"::Consumption);
                            VE03.SETRANGE("Gen. Prod. Posting Group", 'BULK');
                            IF VE03.FINDFIRST THEN BEGIN
                                itm03.RESET;
                                IF itm03.GET(VE03."Item No.") THEN
                                    BulkDesc := itm03.Description;
                            END;
                            //<<03Apr2018
                        END;

                        VersionCode := '';
                        VersionName := '';
                        recProdOrderLine.RESET;
                        recProdOrderLine.SETCURRENTKEY("Inventory Posting Group", "Description 2");
                        recProdOrderLine.SETRANGE(recProdOrderLine."Inventory Posting Group", 'BULK');
                        recProdOrderLine.SETRANGE(recProdOrderLine."Description 2", "Warehouse Entry"."Lot No.");
                        IF recProdOrderLine.FINDFIRST THEN BEGIN
                            VersionCode := recProdOrderLine."Production BOM Version Code";

                            recversion.RESET;
                            recversion.SETRANGE(recversion."Production BOM No.", recProdOrderLine."Item No.");
                            recversion.SETRANGE(recversion."Version Code", VersionCode);
                            IF recversion.FINDFIRST THEN BEGIN
                                VersionName := recversion.Description
                            END;
                        END;

                        MRP := 0;
                        recMRPMaster.RESET;
                        recMRPMaster.SETRANGE(recMRPMaster."Item No.", "Warehouse Entry"."Item No.");
                        recMRPMaster.SETRANGE(recMRPMaster."Lot No.", "Warehouse Entry"."Lot No.");
                        IF recMRPMaster.FINDFIRST THEN BEGIN
                            MRP := recMRPMaster."MRP Price" * recMRPMaster."Qty. per Unit of Measure";
                        END;

                        Price := 0;
                        recSalesPrice.RESET;
                        recSalesPrice.SETRANGE(recSalesPrice."Item No.", "Warehouse Entry"."Item No.");
                        IF recSalesPrice.FINDLAST THEN BEGIN
                            Price := recSalesPrice."Basic Price";
                        END;

                        SalesReturnNo := '';
                        SalesReturnDate := 0D;
                        SalesInvoiceNo := '';
                        SalesInvoiceDate := 0D;
                        TRRno := '';
                        TRRdate := 0D;
                        BRTno := '';
                        BRTdate := 0D;
                        BRTprice := 0;
                        BRTValue := 0;
                        SalesReturnPrice := 0;
                        SalesReturnValue := 0;

                        IF "Warehouse Entry"."Bin Code" = 'V-S-REJ' THEN BEGIN

                            recWarehouseEntry.RESET;
                            recWarehouseEntry.SETCURRENTKEY("Item No.", "Lot No.", "Bin Code");//07Apr2017
                                                                                               //recWarehouseEntry.SETCURRENTKEY("Location Code","Item No.","Variant Code","Zone Code","Bin Code","Lot No.");
                            recWarehouseEntry.SETRANGE(recWarehouseEntry."Source Type", 37);
                            recWarehouseEntry.SETRANGE(recWarehouseEntry."Item No.", "Warehouse Entry"."Item No.");
                            recWarehouseEntry.SETFILTER(recWarehouseEntry."Registering Date", '%1..%2', 0D, filterdate);
                            recWarehouseEntry.SETFILTER(recWarehouseEntry."Location Code", "Warehouse Entry"."Location Code");
                            recWarehouseEntry.SETFILTER(recWarehouseEntry."Bin Code", "Warehouse Entry"."Bin Code");
                            recWarehouseEntry.SETRANGE(recWarehouseEntry."Lot No.", "Warehouse Entry"."Lot No.");
                            IF recWarehouseEntry.FINDFIRST THEN BEGIN

                                recSCH.RESET;
                                recSCH.SETCURRENTKEY("Return Order No.");//)07Apr2017
                                recSCH.SETRANGE(recSCH."Return Order No.", recWarehouseEntry."Source No.");
                                IF recSCH.FINDFIRST THEN BEGIN
                                    SalesReturnNo := recSCH."No.";
                                    SalesReturnDate := recSCH."Posting Date";
                                    SalesInvoiceNo := recSCH."Applies-to Doc. No.";
                                    SalesReturnCustomerCode := recSCH."Sell-to Customer No.";
                                END;

                                recPSIL.RESET;
                                //recPSIL.SETCURRENTKEY();
                                recPSIL.SETRANGE(recPSIL."Document No.", recSCH."Applies-to Doc. No.");
                                IF recPSIL.FIND('-') THEN BEGIN
                                    SalesInvoiceDate := recPSIL."Posting Date";
                                END;

                                recCustomer.RESET;
                                recCustomer.SETRANGE(recCustomer."No.", SalesReturnCustomerCode);
                                IF recCustomer.FINDFIRST THEN BEGIN
                                    SalesReturnCustomerName := recCustomer."Full Name";
                                END;

                                recSCL.RESET;
                                recSCL.SETRANGE(recSCL."Document No.", SalesReturnNo);
                                recSCL.SETRANGE(recSCL."Line No.", recWarehouseEntry."Source Line No.");
                                recSCL.SETRANGE(recSCL."No.", "Warehouse Entry"."Item No.");
                                recSCL.SETRANGE(recSCL."Location Code", "Warehouse Entry"."Location Code");
                                recSCL.SETRANGE(recSCL."Lot No.", "Warehouse Entry"."Lot No.");
                                IF recSCL.FINDFIRST THEN BEGIN
                                    SalesReturnValue := recSCL.Amount;
                                END;

                                Itemrec.RESET;
                                Itemrec.SETRANGE(Itemrec."No.", "Warehouse Entry"."Item No.");
                                Itemrec.SETFILTER(Itemrec."Item Category Code", '%1', 'CAT03');
                                IF Itemrec.FINDFIRST THEN BEGIN

                                    recSCL.RESET;
                                    recSCL.SETRANGE(recSCL."Document No.", SalesReturnNo);
                                    recSCL.SETRANGE(recSCL."Line No.", recWarehouseEntry."Source Line No.");
                                    recSCL.SETRANGE(recSCL."No.", "Warehouse Entry"."Item No.");
                                    recSCL.SETRANGE(recSCL."Location Code", "Warehouse Entry"."Location Code");
                                    recSCL.SETRANGE(recSCL."Lot No.", "Warehouse Entry"."Lot No.");
                                    IF recSCL.FINDFIRST THEN BEGIN
                                        Price := recSCL."Unit Price" / recSCL."Qty. per Unit of Measure";
                                    END;
                                END;
                            END;

                            /*
                            recWarehouseEntry.RESET;
                            recWarehouseEntry.SETRANGE(recWarehouseEntry."Source Type",5741);
                            recWarehouseEntry.SETRANGE(recWarehouseEntry."Item No.","Warehouse Entry"."Item No.");
                            recWarehouseEntry.SETFILTER(recWarehouseEntry."Registering Date",'%1..%2',0D,filterdate);
                            recWarehouseEntry.SETFILTER(recWarehouseEntry."Location Code","Warehouse Entry"."Location Code");
                            recWarehouseEntry.SETFILTER(recWarehouseEntry."Bin Code","Warehouse Entry"."Bin Code");
                            recWarehouseEntry.SETRANGE(recWarehouseEntry."Lot No.","Warehouse Entry"."Lot No.");
                            IF recWarehouseEntry.FINDFIRST THEN
                            */

                            recILE.RESET;
                            recILE.SETCURRENTKEY("Item No.");//07APr2017
                            recILE.SETRANGE(recILE."Item No.", "Warehouse Entry"."Item No.");
                            recILE.SETFILTER(recILE."Posting Date", '%1..%2', 0D, filterdate);
                            recILE.SETRANGE(recILE."Location Code", "Warehouse Entry"."Location Code");
                            recILE.SETRANGE(recILE."Lot No.", "Warehouse Entry"."Lot No.");
                            recILE.SETRANGE(recILE.Open, TRUE);
                            IF recILE.FINDFIRST THEN BEGIN

                                recTRH.RESET;
                                recTRH.SETCURRENTKEY("Transfer Order No.");
                                //recTRH.SETRANGE(recTRH."Transfer Order No.",recWarehouseEntry."Source No.");
                                recTRH.SETRANGE(recTRH."Transfer Order No.", recILE."Order No.");
                                IF recTRH.FINDFIRST THEN BEGIN
                                    TRRno := recTRH."No.";
                                    TRRdate := recTRH."Posting Date";
                                END;

                                recTSH.RESET;
                                recTSH.SETCURRENTKEY("Transfer Order No.");
                                //recTSH.SETRANGE(recTSH."Transfer Order No.",recWarehouseEntry."Source No.");
                                recTSH.SETRANGE(recTSH."Transfer Order No.", recILE."Order No.");
                                IF recTSH.FINDFIRST THEN BEGIN
                                    BRTno := recTSH."No.";
                                    BRTdate := recTSH."Posting Date";
                                    BRTLocationCode := recTSH."Transfer-from Code";
                                END;

                                recLcoation.RESET;
                                recLcoation.SETRANGE(recLcoation.Code, BRTLocationCode);
                                IF recLcoation.FINDFIRST THEN BEGIN
                                    BRTLocationName := recLcoation.Name;
                                END;

                                /*   recRG23D.RESET;
                                  recRG23D.SETCURRENTKEY(recRG23D."Document No.", "Item No.");//07Apr2017
                                  recRG23D.SETRANGE(recRG23D."Document No.", BRTno);
                                  recRG23D.SETRANGE(recRG23D."Item No.", "Warehouse Entry"."Item No.");
                                  recRG23D.SETRANGE(recRG23D."Lot No.", "Warehouse Entry"."Lot No.");
                                  IF recRG23D.FINDFIRST THEN BEGIN
                                      BRTValue := recRG23D."Transfer Price" * "Warehouse Entry"."Qty. (Base)";
                                  END; */

                                Itemrec.RESET;
                                Itemrec.SETRANGE(Itemrec."No.", "Warehouse Entry"."Item No.");
                                Itemrec.SETFILTER(Itemrec."Item Category Code", '%1', 'CAT03');
                                IF Itemrec.FINDFIRST THEN BEGIN
                                    /* 
                                                                        recRG23D.RESET;
                                                                        recRG23D.SETCURRENTKEY(recRG23D."Document No.", "Item No.");//07Apr2017
                                                                        recRG23D.SETRANGE(recRG23D."Document No.", BRTno);
                                                                        recRG23D.SETRANGE(recRG23D."Item No.", "Warehouse Entry"."Item No.");
                                                                        recRG23D.SETRANGE(recRG23D."Lot No.", "Warehouse Entry"."Lot No.");
                                                                        IF recRG23D.FINDFIRST THEN BEGIN
                                                                            Price := recRG23D."Transfer Price";
                                                                        END; */
                                END;

                            END;

                        END;

                        EBTRecIle.RESET;
                        EBTRecIle.SETCURRENTKEY("Item No.");//07Apr2017
                        EBTRecIle.SETRANGE(EBTRecIle."Item No.", "Warehouse Entry"."Item No.");
                        EBTRecIle.SETRANGE(EBTRecIle."Entry Type", EBTRecIle."Entry Type"::Output);
                        EBTRecIle.SETRANGE(EBTRecIle."Lot No.", "Warehouse Entry"."Lot No.");
                        IF EBTRecIle.FINDFIRST THEN BEGIN
                            Density := EBTRecIle."Density Factor";
                        END;


                        //>>07Apr2017
                        ExportToExcel := TRUE;
                        IF GQtyBase = 0 THEN
                            ExportToExcel := FALSE;
                        //<<07Apr2017
                        //IF "Warehouse Entry"."Qty. (Base)" = 0 THEN
                        //ExportToExcel := FALSE;

                        //GroupData;
                        IF ExportToExcel THEN BEGIN
                            SrNo += 1;
                            EnterCell(k, 1, FORMAT(SrNo), FALSE, FALSE, '@');//07Apr2017
                                                                             //EnterCell(k+1,1,FORMAT(SrNo),FALSE,FALSE,'@');
                            EnterCell(k, 2, FORMAT("Item No."), FALSE, FALSE, '@');
                            //ExcBuffer.NewRow;
                            // ExcBuffer.AddColumn("Item No.",FALSE,'',TRUE,FALSE,FALSE,'',2);
                            EnterCell(k, 3, FORMAT(itmDesc), FALSE, FALSE, '@');
                            EnterCell(k, 4, FORMAT(ItmCat), FALSE, FALSE, '@');
                            EnterCell(k, 5, FORMAT("Bin Code"), FALSE, FALSE, '@');
                            EnterCell(k, 6, FORMAT(BinName), FALSE, FALSE, '@');
                            EnterCell(k, 7, FORMAT(BulkDesc), FALSE, FALSE, '@');//03Apr2018

                            EnterCell(k, 8, FORMAT(SalesUOM), FALSE, FALSE, '@');
                            //EnterCell(k+1,8,FORMAT("Qty. (Base)"),FALSE,FALSE,'#,###0.00');
                            EnterCell(k, 9, FORMAT(GQtyBase), FALSE, FALSE, '0.000');//07Apr2017

                            //EnterCell(k+1,9,FORMAT("Qty. (Base)"/ItemUOM),FALSE,FALSE,'#,###0.00');
                            EnterCell(k, 10, FORMAT(GQtyBase / ItemUOM), FALSE, FALSE, '0.000');//07Apr2017
                            TQtySales += GQtyBase / ItemUOM;//07Apr2017
                            GQtyBase := 0;//07Apr2017

                            EnterCell(k, 11, FORMAT("Lot No."), FALSE, FALSE, '@');
                            EnterCell(k, 12, FORMAT(Mfgdte), FALSE, FALSE, '@');
                            EnterCell(k, 13, FORMAT(VersionCode), FALSE, FALSE, '@');
                            EnterCell(k, 14, FORMAT(VersionName), FALSE, FALSE, '@');
                            IF "Warehouse Entry"."Source Type" = 37 THEN BEGIN
                                EnterCell(k, 15, FORMAT(SalesReturnNo), FALSE, FALSE, '@');
                                EnterCell(k, 16, FORMAT(SalesReturnDate), FALSE, FALSE, '@');
                                EnterCell(k, 17, FORMAT(SalesInvoiceNo), FALSE, FALSE, '@');
                                EnterCell(k, 18, FORMAT(SalesInvoiceDate), FALSE, FALSE, '@');
                                EnterCell(k, 19, FORMAT(SalesReturnCustomerCode), FALSE, FALSE, '@');
                                EnterCell(k, 20, FORMAT(SalesReturnCustomerName), FALSE, FALSE, '@');
                                EnterCell(k, 21, FORMAT(SalesReturnValue), FALSE, FALSE, '@');
                            END;
                            IF "Warehouse Entry"."Source Type" = 5741 THEN BEGIN
                                EnterCell(k, 15, FORMAT(TRRno), FALSE, FALSE, '@');
                                EnterCell(k, 16, FORMAT(TRRdate), FALSE, FALSE, '@');
                                EnterCell(k, 17, FORMAT(BRTno), FALSE, FALSE, '@');
                                EnterCell(k, 18, FORMAT(BRTdate), FALSE, FALSE, '@');
                                EnterCell(k, 19, FORMAT(BRTLocationCode), FALSE, FALSE, '@');
                                EnterCell(k, 20, FORMAT(BRTLocationName), FALSE, FALSE, '@');
                                EnterCell(k, 21, FORMAT(BRTValue), FALSE, FALSE, '@');
                            END;
                            EnterCell(k, 22, FORMAT(MRP), FALSE, FALSE, '#,###0.00');
                            EnterCell(k, 23, FORMAT(Price), FALSE, FALSE, '#,###0.00');
                            EnterCell(k, 24, FORMAT(remarks), FALSE, FALSE, '@');
                            EnterCell(k, 25, FORMAT(TransferPrice), FALSE, FALSE, '#,###0.00');
                            EnterCell(k, 26, FORMAT(ROUND(Density, 0.001)), FALSE, FALSE, '@');
                            //EnterCell(k+1,26,FORMAT(ROUND("Warehouse Entry"."Qty. (Base)" *ROUND(Density,0.001),0.01)),FALSE,FALSE,'#,###0.00');
                            EnterCell(k, 27, FORMAT(ROUND("Warehouse Entry"."Qty. (Base)" * ROUND(Density, 0.001), 0.01)), FALSE, FALSE, '0.000');
                            k += 1;
                            //k+=1;
                            Qty += "Warehouse Entry"."Qty. (Base)";
                            Qtytotal += ("Qty. (Base)" / ItemUOM);
                        END;

                        ICOUNT := 0;

                    END;
                    //IF "Warehouse Entry"."Qty. (Base)" = 0 THEN
                    //CurrReport.SHOWOUTPUT(FALSE);

                    /*
                    //>>RSPL/Migration/Rahul
                    
                     CLEAR(QtyBase);
                     WhseEntry.RESET;
                     WhseEntry.SETCURRENTKEY("Item No.","Lot No.","Bin Code");
                     //WhseEntry.SETRANGE(WhseEntry."Source Type",37);
                     WhseEntry.SETRANGE(WhseEntry."Item No.","Warehouse Entry"."Item No.");
                     WhseEntry.SETFILTER(WhseEntry."Registering Date",'%1..%2',0D,filterdate);
                     WhseEntry.SETFILTER(WhseEntry."Location Code","Warehouse Entry"."Location Code");
                     WhseEntry.SETFILTER(WhseEntry."Bin Code","Warehouse Entry"."Bin Code");
                     //WhseEntry.SETRANGE(WhseEntry."Lot No.","Warehouse Entry"."Lot No.");
                     IF WhseEntry.FINDSET THEN REPEAT
                       QtyBase+= WhseEntry."Qty. (Base)";
                    UNTIL WhseEntry.NEXT =0;
                    
                    */

                    //>>2


                    //IF QtyBase =0 THEN
                    //CurrReport.SKIP;
                    //<<

                    //Rahul
                    //IF QtyBase <>0 THEN

                end;

                trigger OnPreDataItem()
                begin

                    "Warehouse Entry".SETFILTER("Warehouse Entry"."Registering Date", '%1..%2', 0D, filterdate);
                    IF LocCode <> '' THEN
                        "Warehouse Entry".SETFILTER("Warehouse Entry"."Location Code", LocCode);
                    IF bincode <> '' THEN
                        "Warehouse Entry".SETFILTER("Warehouse Entry"."Bin Code", bincode);

                    WHEntry07APR.COPYFILTERS("Warehouse Entry");//07Apr2017 NewGrouping
                end;
            }

            trigger OnAfterGetRecord()
            begin

                TransferPrice := 0;
                CLEAR(UOM);
                Itemrec.GET(Item."No.");
                UOM := Itemrec."Sales Unit of Measure";

                CLEAR(ItemUOM);
                recItemUOM.RESET;
                recItemUOM.SETRANGE(recItemUOM."Item No.", Item."No.");
                recItemUOM.SETRANGE(recItemUOM.Code, UOM);
                IF recItemUOM.FINDFIRST THEN BEGIN
                    ItemUOM := recItemUOM."Qty. per Unit of Measure";
                END;

                IF NOT ((Itemrec."Item Category Code" = 'CAT02') OR (Itemrec."Item Category Code" = 'CAT03') OR
                (Itemrec."Item Category Code" = 'CAT11') OR (Itemrec."Item Category Code" = 'CAT12') OR (Itemrec."Item Category Code" = 'CAT13')
                OR (Itemrec."Item Category Code" = 'CAT15')) THEN
                    CurrReport.SKIP;

                //Rahul
                RecSalesPrice1.RESET;
                RecSalesPrice1.SETRANGE(RecSalesPrice1."Item No.", Item."No.");
                IF RecSalesPrice1.FINDLAST THEN BEGIN
                    ItemUnitMesure.RESET;
                    ItemUnitMesure.SETRANGE(ItemUnitMesure."Item No.", Item."No.");
                    ItemUnitMesure.SETRANGE(ItemUnitMesure.Code, RecSalesPrice1."Unit of Measure Code");
                    IF ItemUnitMesure.FINDFIRST THEN
                        TransferPrice := RecSalesPrice1."Transfer Price" / ItemUnitMesure."Qty. per Unit of Measure";
                    TransferPrice := ROUND(TransferPrice, 0.01, '=');
                END
                ELSE
                    TransferPrice := 0;
            end;

            trigger OnPreDataItem()
            begin
                SrNo := 0;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(LocCode; LocCode)
                {
                    Caption = 'Location Code';

                    trigger OnLookup(var Text: Text): Boolean
                    begin

                        Locrec.RESET;
                        //Locrec.SETRANGE(Locrec.Code,"Item Ledger Entry"."Location Code");
                        //Locrec.GET("Item Ledger Entry"."Location Code");
                        IF PAGE.RUNMODAL(0, Locrec) = ACTION::LookupOK THEN
                            LocCode := Locrec.Code;
                    end;
                }
                field(bincode; bincode)
                {
                    Caption = 'Bin Code';

                    trigger OnLookup(var Text: Text): Boolean
                    begin

                        BinCont.RESET;
                        BinCont.SETRANGE(BinCont."Location Code", LocCode);
                        IF BinCont.FINDFIRST THEN BEGIN
                            IF PAGE.RUNMODAL(0, BinCont) = ACTION::LookupOK THEN
                                bincode := BinCont.Code;
                        END
                        ELSE
                            MESSAGE('Location %1 does not contain Bin', LocCode);
                    end;
                }
                field(filterdate; filterdate)
                {
                    Caption = 'Enter Date';
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
            k += 1;//07APr2017
            //EnterCell(k+1,1,'Grand Total',TRUE,FALSE,'@');
            EnterCell(k, 1, 'Grand Total', TRUE, FALSE, '@');//07Apr2017
            //EnterCell(k+1,8,FORMAT(Qty),TRUE,FALSE,'#,###0.00');

            EnterCell(k, 9, FORMAT(TQtyBase), TRUE, FALSE, '0.000'); //07Apr2017
                                                                     //EnterCell(k+1,8,FORMAT(Qty),TRUE,FALSE,'0.000'); //27Mar2017

            // EnterCell(k+1,9,FORMAT(TotalQtyinUOM),Qtytotal,FALSE,FALSE);
            //EnterCell(k+1,9,FORMAT(Qtytotal),TRUE,FALSE,'#,###0.00');

            EnterCell(k, 10, FORMAT(TQtySales), TRUE, FALSE, '0.000'); //07Apr2017
                                                                       //EnterCell(k+1,9,FORMAT(Qtytotal),TRUE,FALSE,'0.000'); //27Mar2017

            k += 1;
        END;
        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN
            //ExcBuffer.CreateBook;
            //ExcBuffer.CreateSheet('Raw Materials Details','',COMPANYNAME,USERID);
            // ExcBuffer.CreateBookAndOpenExcel('', 'Raw Materials Details', '', '', '');
            // ExcBuffer.GiveUserControl;

            ExcBuffer.CreateNewBook('Raw Materials Details');
            ExcBuffer.WriteSheet('Raw Materials Details', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Raw Materials Details', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();
        END;
    end;

    trigger OnPreReport()
    begin

        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN
            ExcBuffer.DELETEALL;
            CLEAR(ExcBuffer);
            EnterCell(1, 1, FORMAT(COMPANYNAME + '' + Item."Location Filter"), TRUE, FALSE, '@');
            EnterCell(2, 1, 'Finished Goods Stock', TRUE, FALSE, '@');
            EnterCell(1, 26, FORMAT(FORMAT(TODAY, 0, 4)), TRUE, FALSE, '@');
            EnterCell(2, 26, FORMAT(USERID), TRUE, FALSE, '@');
            //>>07Apr2017
            EnterCell(4, 1, 'As on  ' + FORMAT(filterdate), TRUE, FALSE, '@');

            LocCode07 := LocCode;
            recLoc07.RESET;
            recLoc07.SETRANGE(Code, LocCode07);
            IF recLoc07.FINDFIRST THEN BEGIN
                LocName07 := recLoc07.Name;

            END;

            EnterCell(5, 1, 'Location :  ' + LocName07, TRUE, FALSE, '@');
            //<<07Apr2017

            EnterCell(6, 1, 'Sr No.', TRUE, TRUE, '@');
            EnterCell(6, 2, 'RM Code', TRUE, TRUE, '@');
            EnterCell(6, 3, 'Item Description', TRUE, TRUE, '@');
            EnterCell(6, 4, 'Item Category Code', TRUE, TRUE, '@');
            EnterCell(6, 5, 'Bin Code', TRUE, TRUE, '@');
            EnterCell(6, 6, 'Bin Name', TRUE, TRUE, '@');
            EnterCell(6, 7, 'Bulk Product', TRUE, TRUE, '@'); //03Apr2018
            EnterCell(6, 8, 'Sales UOM', TRUE, TRUE, '@');
            EnterCell(6, 9, 'Remaining Qty in LTRS/KGS', TRUE, TRUE, '@');
            EnterCell(6, 10, 'Remaining Qty in Sales UOM', TRUE, TRUE, '@');
            EnterCell(6, 11, 'Batch No.', TRUE, TRUE, '@');
            EnterCell(6, 12, 'MFG Date', TRUE, TRUE, '@');
            EnterCell(6, 13, 'Version Code', TRUE, TRUE, '@');
            EnterCell(6, 14, 'Version Name', TRUE, TRUE, '@');
            EnterCell(6, 15, 'Document No.', TRUE, TRUE, '@');
            EnterCell(6, 16, 'Date', TRUE, TRUE, '@');
            EnterCell(6, 17, 'BRT/Invoice No.', TRUE, TRUE, '@');
            EnterCell(6, 18, 'Date', TRUE, TRUE, '@');
            EnterCell(6, 19, 'Location/Customer Code', TRUE, TRUE, '@');
            EnterCell(6, 20, 'Lcoation/Customer Name', TRUE, TRUE, '@');
            EnterCell(6, 21, 'Value', TRUE, TRUE, '@');
            EnterCell(6, 22, 'MRP', TRUE, TRUE, '@');
            EnterCell(6, 23, 'Price', TRUE, TRUE, '@');
            EnterCell(6, 24, 'Remarks', TRUE, TRUE, '@');

            EnterCell(6, 25, 'Current Transfer Price', TRUE, TRUE, '@');
            EnterCell(6, 26, 'Density', TRUE, TRUE, '@');
            EnterCell(6, 27, 'Qty in KGS', TRUE, TRUE, '@');
            //k:=+6;
            k := 7; //07APr2017
        END;
    end;

    var
        SrNo: Integer;
        Itemrec: Record 27;
        Itemrec1: Record 27;
        itmDesc: Text[50];
        ItmCat: Code[10];
        Qty: Decimal;
        ExportToExcel: Boolean;
        ExcBuffer: Record 370 temporary;
        k: Integer;
        LocCode: Code[10];
        filterdate: Date;
        Locrec: Record 14;
        bincode: Code[20];
        BinCont: Record 7354;
        SalesUOM: Text[30];
        recItemUOM: Record 5404;
        ItemUOM: Decimal;
        UOM: Code[10];
        TotalQtyinUOM: Decimal;
        recBin: Record 7354;
        BinName: Text[50];
        ILErec: Record 32;
        Mfgdte: Date;
        recWarehouseEntry: Record 7312;
        EBTBinCode: Code[10];
        BlendorderNo: Code[20];
        recProdOrderLine: Record 5406;
        VersionCode: Code[10];
        recBinContent: Record 7302;
        recversion: Record 99000779;
        VersionName: Text;
        recSCH: Record 114;
        SalesReturnNo: Code[20];
        SalesReturnDate: Date;
        SalesInvoiceNo: Code[20];
        SalesInvoiceDate: Date;
        recPSIL: Record 113;
        TRRno: Code[20];
        TRRdate: Date;
        BRTno: Code[20];
        BRTdate: Date;
        recTRH: Record 5746;
        recTSH: Record 5744;
        BRTLocationCode: Code[10];
        BRTLocationName: Text[50];
        recLcoation: Record 14;
        SalesReturnCustomerCode: Code[10];
        SalesReturnCustomerName: Text[50];
        recCustomer: Record 18;
        BRTprice: Decimal;
        BRTValue: Decimal;
        SalesReturnPrice: Decimal;
        SalesReturnValue: Decimal;
        recSCL: Record 115;
        // recRG23D: Record 16537;
        recMRPMaster: Record 50013;
        MRP: Decimal;
        recSalesPrice: Record 7002;
        Price: Decimal;
        recILE: Record 32;
        remarks: Text;
        recItemTable: Record 27;
        RecSalesPrice1: Record 7002;
        TransferPrice: Decimal;
        ItemUnitMesure: Record 5404;
        EBTRecIle: Record 32;
        Density: Decimal;
        Qtytotal: Decimal;
        "---RSPL-Migration--": Integer;
        WhseEntry: Record 7312;
        QtyBase: Decimal;
        "--07Apr2017": Integer;
        WHEntry07APR: Record 7312;
        GCOUNT: Integer;
        ICOUNT: Integer;
        GQtyBase: Decimal;
        TQtyBase: Decimal;
        TQtySales: Decimal;
        LocCode07: Code[20];
        recLoc07: Record 14;
        LocName07: Text[100];
        "---------------------03Apr2018": Integer;
        BulkDesc: Text;
        itm03: Record 27;
        VE03: Record 5802;

    // //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NumbFormat: Text)
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
            ExcBuffer.NumberFormat := NumbFormat;
            ExcBuffer.INSERT;
        END;
    end;

    local procedure GroupData()
    begin
    end;
}


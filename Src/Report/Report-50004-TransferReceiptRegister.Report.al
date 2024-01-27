report 50004 "Transfer Receipt Register"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TransferReceiptRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Transfer Receipt Header"; "Transfer Receipt Header")
        {
            DataItemTableView = SORTING("Posting Date");
            RequestFilterFields = "Posting Date", "Transfer-from Code", "Transfer-to Code";
            dataitem("Transfer Receipt Line"; "Transfer Receipt Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Quantity = FILTER(<> 0));
                RequestFilterFields = "Item No.", "Shortcut Dimension 1 Code";

                trigger OnAfterGetRecord()
                begin

                    LCOUNT += 1;//08Mar2017
                                //MESSAGE('LC %1 \ HC %2 \ Doc No. %3',LCOUNT,HCOUNT,"Document No.");
                                //>>1


                    CLEAR(PackedUOM);
                    retItem.RESET;
                    retItem.SETRANGE(retItem."No.", "Item No.");
                    IF retItem.FINDFIRST THEN BEGIN
                        PackedUOM := retItem."Sales Unit of Measure";
                    END;

                    //IF "Transfer Receipt Line".Quantity = 0 THEN
                    //CurrReport.SKIP;

                    CLEAR(BatchNo);
                    reILE.RESET;
                    reILE.SETRANGE(reILE."Document No.", "Transfer Receipt Line"."Document No.");
                    reILE.SETRANGE(reILE."Location Code", "Transfer Receipt Line"."Transfer-from Code");
                    reILE.SETRANGE(reILE."Item No.", "Transfer Receipt Line"."Item No.");
                    reILE.SETRANGE(reILE."Document Line No.", "Transfer Receipt Line"."Line No.");
                    IF reILE.FINDFIRST THEN BEGIN
                        BatchNo := BatchNo + reILE."Lot No.";
                    END;

                    /*
                    TotalNetValue:=TotalNetValue+"Transfer Receipt Line".Amount;
                    TotalBedAmt:=TotalBedAmt+"Transfer Receipt Line"."BED Amount";
                    TotaleCessAmt:=TotaleCessAmt+"Transfer Receipt Line"."eCess Amount";
                    TotalSHECess:=TotalSHECess+"Transfer Receipt Line"."SHE Cess Amount";
                    TotAddDuty := TotAddDuty+"Transfer Receipt Line"."ADE Amount";
                    AMT := AMT1 + AMT2;
                    Grpfooterqty := Grpfooterqty +"Transfer Receipt Line".Quantity;
                    Grpamt:= Grpamt+"Transfer Receipt Line".Amount;
                    Grpbed:= Grpbed+"Transfer Receipt Line"."BED Amount";
                    GrpeCess:= GrpeCess+"Transfer Receipt Line"."eCess Amount";
                    GrpSHE:= GrpSHE+"Transfer Receipt Line"."SHE Cess Amount";
                    */
                    IF Item.GET("Transfer Receipt Line"."Item No.") THEN;//Item Description

                    //<<1


                    //>>2


                    //IF RecTraShipHeader.GET("Transfer Receipt Line"."Document No.") THEN

                    //ReceiptTotalQty+= Quantity*"Transfer Receipt Line"."Qty. per Unit of Measure";

                    TotalBRTQty := TotalBRTQty + "Transfer Receipt Line"."Quantity (Base)";//Total Qty

                    //IF NOT Detail THEN
                    //CurrReport.SHOWOUTPUT(FALSE);

                    /*
                    IF printtoexcel AND Detail THEN BEGIN
                    
                            XLSHEET.Range('B'+FORMAT(i)).Value := Item.Description ;
                            XLSHEET.Range('B'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('D'+FORMAT(i)).Value := PackedUOM;
                            XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('D'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('E'+FORMAT(i)).Value := Quantity * "Transfer Receipt Line"."Qty. per Unit of Measure";
                            XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('E'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            AMT1:= ("Transfer Receipt Line".Amount+"Transfer Receipt Line"."BED Amount");
                            AMT2:= ("Transfer Receipt Line"."eCess Amount"+"Transfer Receipt Line"."SHE Cess Amount");
                            AMT:= AMT1+AMT2+Freight+"Transfer Receipt Line"."ADE Amount";
                    
                            XLSHEET.Range('I'+FORMAT(i)).Value := BatchNo;
                            XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('I'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('N'+FORMAT(i)).Value :="Transfer Receipt Line"."Item No.";
                            XLSHEET.Range('N'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('N'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('N'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Borders.ColorIndex := 5;
                    
                       i += 1;
                    END;
                    *///Commented 08Mar2017
                      //<<2

                    //>>08Mar2017 LineDetails
                    IF printtoexcel AND Detail THEN BEGIN

                        EnterCell(NN, 2, Item.Description, FALSE, FALSE, '', 1);
                        EnterCell(NN, 4, PackedUOM, FALSE, FALSE, '', 1);
                        EnterCell(NN, 5, FORMAT("Transfer Receipt Line"."Quantity (Base)"), FALSE, FALSE, '####0.000', 0);
                        //>>
                        GQty += "Transfer Receipt Line"."Quantity (Base)";
                        //<<
                        EnterCell(NN, 9, BatchNo, FALSE, FALSE, '', 1);
                        EnterCell(NN, 14, "Transfer Receipt Line"."Item No.", FALSE, FALSE, '', 1);

                        NN += 1;
                    END;
                    //<<LineDetails


                    //>>3

                    //Transfer Receipt Line, GroupFooter (2) - OnPreSection()
                    Grpfooterqty := 0;
                    /*
                    "Transfer Receipt Line".RESET;
                    "Transfer Receipt Line".SETRANGE("Transfer Receipt Line"."Document No.","Transfer Receipt Header"."No.");
                    "Transfer Receipt Line".SETRANGE("Item No.","Transfer Receipt Line"."Item No.");
                    */
                    RecItemCat.RESET;
                    RecItemCat.SETRANGE(RecItemCat.Code, "Item Category Code");
                    IF RecItemCat.FINDFIRST THEN
                        ItemCatName := RecItemCat.Description;

                    /*
                    IF printtoexcel AND ( CurrReport.SHOWOUTPUT("Transfer Receipt Line".Quantity <> 0) ) THEN BEGIN
                    
                            XLSHEET.Range('A'+FORMAT(i)).Value := "Transfer Receipt Header"."Posting Date" ;
                            XLSHEET.Range('A'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('A'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('A'+FORMAT(i)).NumberFormat := 'dd/mm/yyyy';
                            XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('B'+FORMAT(i)).Value :="Transfer Receipt Header"."Transfer-to Name";
                            XLSHEET.Range('B'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('C'+FORMAT(i)).Value := "Transfer Receipt Header"."No.";
                            XLSHEET.Range('C'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('C'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('C'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
                    
                    
                            XLSHEET.Range('E'+FORMAT(i)).Value :="Transfer Receipt Line".Quantity;
                            XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('E'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            AMT1:= ("Transfer Receipt Line".Amount+"Transfer Receipt Line"."BED Amount");
                            AMT2:= ("Transfer Receipt Line"."eCess Amount"+"Transfer Receipt Line"."SHE Cess Amount");
                            AMT:= AMT1+AMT2+Freight+"Transfer Receipt Line"."ADE Amount";
                    
                            XLSHEET.Range('F'+FORMAT(i)).Value :=LrNo;
                            XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('F'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('G'+FORMAT(i)).Value := VehicleNo;
                            XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('G'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('H'+FORMAT(i)).Value := TtrasporterName;
                            XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('H'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('J'+FORMAT(i)).Value :="Transfer Receipt Header"."Transfer-from Code" ;
                            XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('K'+FORMAT(i)).Value :="Transfer Receipt Header"."Transfer-from Name";
                            XLSHEET.Range('K'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('K'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                            XLSHEET.Range('K'+FORMAT(i)).HorizontalAlignment :=-4152;
                    
                            XLSHEET.Range('L'+FORMAT(i)).Value := ReceiptNo;
                            XLSHEET.Range('L'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('L'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('M'+FORMAT(i)).Value := ItemCatName;
                            XLSHEET.Range('M'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('M'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                    
                            XLSHEET.Range('Q'+FORMAT(i)).Value := "Transfer Receipt Header"."Transfer Order No.";
                            XLSHEET.Range('Q'+FORMAT(i)).RowHeight := 18;
                            XLSHEET.Range('Q'+FORMAT(i)).Font.Size := 10;
                            XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
                            XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                    
                    
                    
                           XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.Color :=65535;
                    
                       i += 1;
                    END;
                    *///Commented 08Mar2017
                      //<<3

                    //>>08Mar2017 LineGrouping
                    IF printtoexcel THEN BEGIN
                        IF LCOUNT = HCOUNT THEN BEGIN

                            EnterCell(NN, 1, FORMAT("Transfer Receipt Header"."Posting Date"), TRUE, FALSE, 'dd/mm/yyyy', 2);
                            EnterCell(NN, 2, "Transfer Receipt Header"."Transfer-to Name", TRUE, FALSE, '', 1);
                            EnterCell(NN, 3, "Transfer Receipt Header"."No.", TRUE, FALSE, '', 1);
                            EnterCell(NN, 5, FORMAT(GQty), TRUE, FALSE, '####0.000', 0);
                            EnterCell(NN, 6, LrNo, TRUE, FALSE, '', 1);
                            EnterCell(NN, 7, VehicleNo, TRUE, FALSE, '', 1);
                            EnterCell(NN, 8, TtrasporterName, TRUE, FALSE, '', 1);
                            EnterCell(NN, 10, "Transfer Receipt Header"."Transfer-from Code", TRUE, FALSE, '', 1);
                            EnterCell(NN, 11, "Transfer Receipt Header"."Transfer-from Name", TRUE, FALSE, '', 1);
                            EnterCell(NN, 12, ReceiptNo, TRUE, FALSE, '', 1);
                            EnterCell(NN, 13, ItemCatName, TRUE, FALSE, '', 1);
                            EnterCell(NN, 17, "Transfer Receipt Header"."Transfer Order No.", TRUE, FALSE, '', 1);

                            NN += 1;
                            LCOUNT := 0;
                            GQty := 0;
                        END;


                    END;
                    //<< LineGrouping

                end;

                trigger OnPreDataItem()
                begin

                    LCOUNT := 0;//08Mar2017
                    HCOUNT := COUNT; //08Mar2017
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                TransShpHeader.RESET;
                TransShpHeader.SETRANGE("Transfer Order No.", "Transfer Order No.");
                IF TransShpHeader.FINDFIRST THEN BEGIN
                    ReceiptNo := TransShpHeader."No.";
                    LrNo := TransShpHeader."LR/RR No.";
                    VehicleNo := TransShpHeader."Vehicle No.";
                    TtrasporterName := TransShpHeader."Transporter Name";
                END ELSE
                    ReceiptNo := 'IN-TRANSIT';

                CLEAR(TIN);
                CLEAR(CST);
                reclocation.RESET;
                reclocation.SETRANGE(reclocation.Code, "Transfer Receipt Header"."Transfer-to Code");
                IF reclocation.FINDFIRST THEN BEGIN
                    TIN := 'reclocation."T.I.N. No."';
                    CST := 'reclocation."C.S.T No."';
                END;

                /*
                CLEAR(Freight);
                CLEAR(AddlDutyAmount);
                PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.","Transfer Receipt Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type",PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group",'FREIGHT');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                Freight:=PostedStrOrdrDetails."Calculation Value";
                
                PostedStrOrdrDetails.RESET;
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."No.","Transfer Receipt Header"."No.");
                PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group",'ADDL.DUTY');
                IF PostedStrOrdrDetails.FINDFIRST THEN
                AddlDutyAmount:=PostedStrOrdrDetails."Calculation Value";
                
                TotalFreight:=TotalFreight+Freight;
                TotalAddlDutyAmount:=TotalAddlDutyAmount+AddlDutyAmount;
                
                ShipmentTotalQty:=0;
                */
                //<<1


                //>>2

                //Transfer Receipt Header, Header (1) - OnPreSection()
                //CompInfo.GET;
                //UserSetup.GET(USERID);
                //<<2

            end;

            trigger OnPostDataItem()
            begin

                //>>1
                /*
                IF printtoexcel THEN BEGIN
                        XLSHEET.Range('A'+FORMAT(i)).Value := 'TOTAL';
                        XLSHEET.Range('A'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('A'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('A'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('E'+FORMAT(i)).Value := TotalBRTQty ;
                        XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('E'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('E'+FORMAT(i)).Interior.ColorIndex := 45;
                
                        XLSHEET.Range('F'+FORMAT(i)).Value := '';
                        XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('G'+FORMAT(i)).Value := '';
                        XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
                
                        XLSHEET.Range('H'+FORMAT(i)).Value := '';
                        XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('H'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('H'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('H'+FORMAT(i)).Interior.ColorIndex := 45;
                
                        XLSHEET.Range('I'+FORMAT(i)).Value := '';
                        XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('I'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('I'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('I'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('J'+FORMAT(i)).Value := '' ;
                        XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('J'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('J'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('J'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('K'+FORMAT(i)).Value := '';
                        XLSHEET.Range('K'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('K'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('K'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('K'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('K'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('K'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('L'+FORMAT(i)).Value := '';
                        XLSHEET.Range('L'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('L'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('L'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('L'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('L'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('L'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('M'+FORMAT(i)).Value := '';
                        XLSHEET.Range('M'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('M'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('M'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('M'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('M'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('M'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('N'+FORMAT(i)).Value := '';
                        XLSHEET.Range('N'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('N'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('N'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('N'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('N'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('N'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('O'+FORMAT(i)).Value := '';
                        XLSHEET.Range('O'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('O'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('O'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('O'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('O'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('O'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('P'+FORMAT(i)).Value := '';
                        XLSHEET.Range('P'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('P'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('P'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('P'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('P'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('P'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('Q'+FORMAT(i)).Value := '';
                        XLSHEET.Range('Q'+FORMAT(i)).RowHeight := 18;
                        XLSHEET.Range('Q'+FORMAT(i)).Font.Size := 10;
                        XLSHEET.Range('Q'+FORMAT(i)).NumberFormat := '@';
                        XLSHEET.Range('Q'+FORMAT(i)).EntireColumn.AutoFit;
                        XLSHEET.Range('Q'+FORMAT(i)).HorizontalAlignment :=-4152;
                        XLSHEET.Range('Q'+FORMAT(i)).Borders.ColorIndex := 5;
                
                        XLSHEET.Range('A'+ FORMAT(i),MaxCol+FORMAT(i)).Interior.ColorIndex :=8;
                
                        XLSHEET.Range('A:Q').Borders.ColorIndex := 5;
                        XLAPP.Visible(TRUE);
                END;
                *///Commented 08Mar2017
                  //<<1

                //>>08Mar2017 FTotal
                IF printtoexcel THEN BEGIN
                    EnterCell(NN, 1, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 2, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 3, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 4, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 5, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 6, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 7, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 8, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 9, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 10, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 11, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 12, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 13, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 14, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 15, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 16, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 17, '', TRUE, TRUE, '', 1);
                    NN += 1;

                    EnterCell(NN, 1, 'TOTAL', TRUE, TRUE, '', 1);
                    EnterCell(NN, 2, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 3, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 4, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 5, FORMAT(TotalBRTQty), TRUE, TRUE, '####0.000', 0);
                    EnterCell(NN, 6, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 7, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 8, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 9, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 10, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 11, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 12, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 13, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 14, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 15, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 16, '', TRUE, TRUE, '', 1);
                    EnterCell(NN, 17, '', TRUE, TRUE, '', 1);

                END;
                //<<FTotal

            end;

            trigger OnPreDataItem()
            begin

                //>>1


                CompInfo.GET;
                UserSetup.GET(USERID);
                Location.SETFILTER(Location.Code, UserSetup."Sales Resp. Ctr. Filter");
                IF Location.FINDFIRST THEN
                    Loc := COPYSTR(Location.Name, 7, 50);
                filtertext := Loc;

                /*
                IF printtoexcel THEN BEGIN
                CreateXLSHEET;
                CreateHeader;
                i := 7;
                END;
                *///Commented 08Mar2017
                  //CurrReport.CREATETOTALS(AddlDutyAmount);

                //<<1

                //>>08Mar2017
                IF printtoexcel THEN BEGIN

                    ExcBuffer.DELETEALL;
                    CLEAR(ExcBuffer);

                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '@', 1);
                    EnterCell(1, 17, FORMAT(FORMAT(TODAY, 0, 4)), TRUE, FALSE, '', 1);

                    EnterCell(2, 1, 'Transfer Receipt Register', TRUE, FALSE, '@', 1);
                    EnterCell(2, 17, FORMAT(USERID), TRUE, FALSE, '', 1);
                    //EnterCell(5,1,datefiltertext,TRUE,FALSE,'@',1);

                    //EnterCell(5,1,filtertext,TRUE,FALSE,'@',1);
                    //EnterCell(6,1,'Date : '+ DateFilter,TRUE,FALSE,'@',1);

                    EnterCell(5, 1, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 2, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 3, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 4, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 5, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 6, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 7, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 8, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 9, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 10, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 11, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 12, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 13, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 14, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 15, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 16, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 17, '', TRUE, TRUE, '@', 1);


                    EnterCell(6, 1, 'Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 2, 'Grade/Branch', TRUE, TRUE, '@', 1);
                    EnterCell(6, 3, 'Transfer Receipt No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 4, 'Pack UOM', TRUE, TRUE, '@', 1);
                    EnterCell(6, 5, 'Quantity', TRUE, TRUE, '@', 1);
                    EnterCell(6, 6, 'LR No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 7, 'Vehicle No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 8, 'Transporter', TRUE, TRUE, '@', 1);
                    EnterCell(6, 9, 'Batch No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 10, 'Location Code', TRUE, TRUE, '@', 1);
                    EnterCell(6, 11, 'Location Name', TRUE, TRUE, '@', 1);
                    EnterCell(6, 12, 'Transfer Shipment No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 13, 'Item Caterogy', TRUE, TRUE, '@', 1);
                    EnterCell(6, 14, 'Item Code', TRUE, TRUE, '@', 1);
                    EnterCell(6, 15, 'Drivers Mobile No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 16, 'Status', TRUE, TRUE, '@', 1);
                    EnterCell(6, 17, 'Transfer Order No.', TRUE, TRUE, '@', 1);

                    NN := 7;

                END;
                //>>08Mar2017

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print To Excel"; printtoexcel)
                {
                }
                field(Details; Detail)
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
        //>>08Mar2017 RB-N
        IF printtoexcel THEN BEGIN
            // ExcBuffer.CreateBook('', 'Transfer Receipt Register');
            // ExcBuffer.CreateBookAndOpenExcel('', 'Transfer Receipt Register', '', '', USERID);
            // ExcBuffer.GiveUserControl;

            ExcBuffer.CreateNewBook('Transfer Receipt Register');
            ExcBuffer.WriteSheet('Transfer Receipt Register', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Transfer Receipt Register', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();

        END;
        //<<
    end;

    var
        PPF: Page "Job Resource Manager RC";
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
        User: Record 2000000120;
        reclocation: Record 14;
        TIN: Text[30];
        CST: Text[30];
        ReceiptTotalQty: Decimal;
        TransShpHeader: Record 5744;
        "---008Mar2017": Integer;
        ExcBuffer: Record 370 temporary;
        NN: Integer;
        LCOUNT: Integer;
        HCOUNT: Integer;
        GQty: Decimal;

    // //[Scope('Internal')]
    procedure CreateXLSHEET()
    begin

        /*
        CLEAR(XLAPP);
        CLEAR(XLWBOOK);
        CLEAR(XLSHEET);
        CREATE(XLAPP);
        XLWBOOK := XLAPP.Workbooks.Add;
        XLSHEET := XLWBOOK.Worksheets.Add;
        XLAPP.Visible(FALSE);
        MaxCol:= 'Q';
        *///Commented 08Mar2017

    end;

    // //[Scope('Internal')]
    procedure CreateHeader()
    begin

        /*
        XLSHEET.Activate;
        
        
        XLSHEET.Range('A1',MaxCol+'1').Merge := TRUE;
        XLSHEET.Range('A1',MaxCol+'1').Font.Bold := TRUE;
        XLSHEET.Range('A1',MaxCol+'1').Value :=CompInfo.Name;
        XLSHEET.Range('A1',MaxCol+'1').HorizontalAlignment := 3;
        XLSHEET.Range('A1',MaxCol+'1').Interior.ColorIndex := 8;
        XLSHEET.Range('A1',MaxCol+'1').Borders.ColorIndex := 5;
        XLSHEET.Range('A1',MaxCol+'1').Merge := TRUE;
        XLSHEET.Range('A1','A2').Merge := TRUE;
        
        
        XLSHEET.Range('A2',MaxCol+'2').Value := filtertext;
        XLSHEET.Range('A2',MaxCol+'2').Font.Bold := TRUE;
        XLSHEET.Range('A2',MaxCol+'2').HorizontalAlignment := 3;
        XLSHEET.Range('A2',MaxCol+'2').Interior.ColorIndex := 8;
        XLSHEET.Range('A2',MaxCol+'2').Borders.ColorIndex := 5;
        XLSHEET.Range('A2',MaxCol+'2').Merge := TRUE;
        
        XLSHEET.Range('A3',MaxCol+'3').Merge := TRUE;
        
        XLSHEET.Range('A4',MaxCol+'4').Merge := TRUE;
        XLSHEET.Range('A4',MaxCol+'4').Font.Bold := TRUE;
        XLSHEET.Range('A4',MaxCol+'4').Value := datefiltertext;
        XLSHEET.Range('A4',MaxCol+'4').Interior.ColorIndex := 8;
        XLSHEET.Range('A4',MaxCol+'4').Borders.ColorIndex := 5;
        XLSHEET.Range('A4',MaxCol+'4').Merge := TRUE;
        
        XLSHEET.Range('A5',MaxCol+'2').Merge := TRUE;
        
        
        XLSHEET.Range('A6').Value :='Date';
        XLSHEET.Range('A6').Font.Bold := TRUE;
        XLSHEET.Range('A6',MaxCol+'6').Interior.ColorIndex := 45;
        XLSHEET.Range('A6',MaxCol+'6').Borders.ColorIndex := 5;
        
        XLSHEET.Range('B6').Value := ' Grade/Branch';
        XLSHEET.Range('B6').Font.Bold := TRUE;
        
        XLSHEET.Range('C6').Value := 'Transfer Receipt No.';
        XLSHEET.Range('C6').Font.Bold := TRUE;
        
        XLSHEET.Range('D6').Value := 'Pack UOM';
        XLSHEET.Range('D6').Font.Bold := TRUE;
        
        XLSHEET.Range('E6').Value := 'Quantity';
        XLSHEET.Range('E6').Font.Bold := TRUE;
        
        XLSHEET.Range('F6').Value := 'LR No.';
        XLSHEET.Range('F6').Font.Bold := TRUE;
        
        XLSHEET.Range('G6').Value := 'Vehicle No.';
        XLSHEET.Range('G6').Font.Bold := TRUE;
        
        XLSHEET.Range('H6').Value := 'Transporter';
        XLSHEET.Range('H6').Font.Bold := TRUE;
        
        XLSHEET.Range('I6').Value := 'Batch No.';
        XLSHEET.Range('I6').Font.Bold := TRUE;
        
        XLSHEET.Range('J6').Value := 'Location Code';
        XLSHEET.Range('J6').Font.Bold := TRUE;
        
        XLSHEET.Range('K6').Value := 'Location Name';
        XLSHEET.Range('K6').Font.Bold := TRUE;
        
        XLSHEET.Range('L6').Value := 'Transfer Shipment No.';
        XLSHEET.Range('L6').Font.Bold := TRUE;
        
        XLSHEET.Range('M6').Value := 'Item Caterogy';
        XLSHEET.Range('M6').Font.Bold := TRUE;
        
        XLSHEET.Range('N6').Value := 'Item Code';
        XLSHEET.Range('N6').Font.Bold := TRUE;
        
        XLSHEET.Range('O6').Value := 'Drivers Mobile No.';
        XLSHEET.Range('O6').Font.Bold := TRUE;
        
        XLSHEET.Range('P6').Value := 'Status';
        XLSHEET.Range('P6').Font.Bold := TRUE;
        
        XLSHEET.Range('Q6').Value := 'Transfer Order No.';
        XLSHEET.Range('Q6').Font.Bold := TRUE;
        
        XLSHEET.Range('A6:Q6').Borders.ColorIndex := 5;
        XLAPP.Visible(TRUE);
        *///Commented 08Mar2017

    end;

    // //[Scope('Internal')]
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


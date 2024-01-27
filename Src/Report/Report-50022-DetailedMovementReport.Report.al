// report 50022 "RM Detailed Movement Report"
// {
//     DefaultLayout = RDLC;
//     RDLCLayout = './RMDetailedMovementReport.rdlc';

//     dataset
//     {
//         dataitem("Item Ledger Entry"; "Item Ledger Entry")
//         {
//             DataItemTableView = SORTING("Item No.")
//                                 ORDER(Ascending)
//                                 WHERE("Item Category Code" = FILTER(<> CAT02 & <> CAT03));
//             RequestFilterFields = "Location Code", "Posting Date", "Item No.";

//             trigger OnAfterGetRecord()
//             begin
//                 TempItem += 1;//>>20Feb2017 RB-N

//                 //>>09Mar2017 LocName1
//                 CLEAR(LocName1);
//                 recLOC.RESET;
//                 recLOC.SETRANGE(Code, "Item Ledger Entry"."Location Code");
//                 IF recLOC.FINDFIRST THEN BEGIN
//                     LocName1 := recLOC.Name;
//                 END;
//                 //<< LocName1

//                 //>>20Feb2017 RB-N
//                 ILE20.SETCURRENTKEY("Item No.");
//                 ILE20.SETRANGE("Item No.", "Item No.");
//                 IF ILE20.FINDFIRST THEN BEGIN
//                     ILECOUNT := ILE20.COUNT;
//                     //MESSAGE('Item No. %1 \ TempItem %2 \ Count %3',ILE20."Item No.",TempItem,ILECOUNT);
//                 END;
//                 //<<


//                 //>>2

//                 CLEAR(OpenQty);
//                 CLEAR(QtyTransfered);
//                 CLEAR(QtyConsumed);
//                 CLEAR(ProdOrdrNo);
//                 CLEAR(BRTNo);

//                 //>>RSPL-Rahul
//                 IF cnt = 0 THEN BEGIN
//                     ILE.RESET;
//                     ILE.SETCURRENTKEY("Item No.", "Entry Type", "Variant Code", "Lot No.", "Drop Shipment", "Location Code", "Posting Date");
//                     ILE.SETRANGE("Item No.", "Item No.");
//                     ILE.SETRANGE("Location Code", "Location Code");
//                     ILE.SETRANGE("Entry No.", 1, "Entry No." - 1);
//                     IF ILE.FINDSET THEN BEGIN
//                         REPEAT
//                             OpenQty2 += ILE.Quantity;
//                         UNTIL ILE.NEXT = 0;
//                     END ELSE BEGIN
//                         OpenQty2 := 0;
//                     END;
//                 END
//                 ELSE
//                     OpenQty2 := VClosingQuantity;
//                 cnt += 1;
//                 //<<RSPL-Rahul

//                 recILE.RESET;
//                 recILE.GET("Item Ledger Entry"."Entry No.");
//                 IF ((recILE."Entry Type" = recILE."Entry Type"::Purchase) AND (recILE."Document Type" = recILE."Document Type"::"Purchase Receipt")) OR
//                   ((recILE."Entry Type" = recILE."Entry Type"::Transfer) AND (recILE."Document Type" = recILE."Document Type"::"Transfer Receipt")) OR
//                   ((recILE."Entry Type" = recILE."Entry Type"::Output) AND (recILE."Document Type" = recILE."Document Type"::" ")) THEN
//                     OpenQty := recILE.Quantity
//                 ELSE
//                     IF (recILE."Entry Type" = recILE."Entry Type"::Transfer) AND
//                       (recILE."Document Type" = recILE."Document Type"::"Transfer Shipment") THEN
//                         QtyTransfered := recILE.Quantity
//                     ELSE
//                         IF (recILE."Entry Type" = recILE."Entry Type"::Consumption) OR (recILE."Entry Type" = recILE."Entry Type"::Sale) THEN
//                             QtyConsumed := recILE.Quantity
//                         ELSE
//                             IF (recILE."Entry Type" = recILE."Entry Type"::"Positive Adjmt.") THEN
//                                 OpenQty := recILE.Quantity;

//                 IF (recILE."Entry Type" = recILE."Entry Type"::Consumption) OR (recILE."Entry Type" = recILE."Entry Type"::Output) THEN
//                     ProdOrdrNo := recILE."Document No."
//                 ELSE
//                     BRTNo := recILE."Document No.";

//                 ItemRec.GET("Item Ledger Entry"."Item No.");

//                 RateofItem := 0;
//                 TotalCost := 0;
//                 TotalQty := 0;
//                 ChargeofItem := 0;
//                 ValueEntry.RESET;
//                 ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
//                 ValueEntry.SETRANGE(ValueEntry."Item Charge No.", '');
//                 IF ValueEntry.FINDSET THEN
//                     REPEAT
//                         TotalCost += ValueEntry."Cost Amount (Expected)" + ValueEntry."Cost Amount (Actual)";
//                         TotalQty += ValueEntry."Item Ledger Entry Quantity";
//                     UNTIL ValueEntry.NEXT = 0;
//                 RateofItem := TotalCost / TotalQty;
//                 ValueEntry.RESET;
//                 ValueEntry.SETRANGE(ValueEntry."Item Ledger Entry No.", "Item Ledger Entry"."Entry No.");
//                 ValueEntry.SETFILTER(ValueEntry."Item Charge No.", '<>%1', '');
//                 IF ValueEntry.FINDFIRST THEN
//                     REPEAT
//                         ChargeofItem += ValueEntry."Cost per Unit";
//                     UNTIL ValueEntry.NEXT = 0;
//                 IF ChargeofItem < 0 THEN BEGIN
//                     RateofItem := RateofItem + ChargeofItem;
//                     ChargeofItem := 0;
//                 END;
//                 //IPOL/001 >>>>
//                 vQtyRecied := 0;
//                 IF ("Document Type" = "Document Type"::"Transfer Receipt") OR ("Document Type" = "Document Type"::"Purchase Receipt") OR
//                    ("Document Type" = "Document Type"::"Sales Return Receipt") OR ("Entry Type" = "Entry Type"::"Positive Adjmt.") THEN
//                     vQtyRecied := ABS(Quantity);
//                 //IPOL/001 <<<<


//                 VClosingQuantity := OpenQty2 + OpenQty + QtyConsumed + QtyTransfered;

//                 //<<2

//                 MakeExcelDataBody;

//                 //>>20Feb2017 RB-N
//                 IF TempItem = ILECOUNT THEN BEGIN
//                     cnt := 0;
//                     TempItem := 0;

//                 END;
//                 //<<
//             end;

//             trigger OnPreDataItem()
//             begin

//                 //>>09Mar2017 LocName
//                 LocCode := "Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code");
//                 recLOC.RESET;
//                 recLOC.SETRANGE(Code, LocCode);
//                 IF recLOC.FINDFIRST THEN BEGIN
//                     LocName := recLOC.Name;
//                 END;
//                 //<< LocName


//                 MakeExcelDataHeader;
//                 //>>1
//                 ClosingQty := 0;
//                 IF GETFILTER("Location Code") = '' THEN
//                     ERROR('Please Specify Location Code');
//                 //<<1


//                 //>>20Feb2017 RB-N
//                 ILE20.RESET;
//                 ILE20.COPYFILTERS("Item Ledger Entry");

//                 //<<
//                 TempItem := 0;//>>20Feb2017 RB-N
//             end;
//         }
//     }

//     requestpage
//     {

//         layout
//         {
//         }

//         actions
//         {
//         }
//     }

//     labels
//     {
//     }

//     trigger OnPostReport()
//     begin
//         CreateExcelbook;
//     end;

//     trigger OnPreReport()
//     begin
//         //MakeExcelInfo;
//     end;

//     var
//         OpenQty: Decimal;
//         QtyTransfered: Decimal;
//         ProdOrdrNo: Code[20];
//         BRTNo: Code[20];
//         QtyConsumed: Decimal;
//         recILE: Record 32;
//         ItemRec: Record 27;
//         ClosingQty: Decimal;
//         ExcelBuf: Record 370 temporary;
//         ValueEntry: Record 5802;
//         RateofItem: Decimal;
//         TotalCost: Decimal;
//         TotalQty: Decimal;
//         ChargeofItem: Decimal;
//         "----": Integer;
//         ILE: Record 32;
//         OpenQty2: Decimal;
//         GetDate: Date;
//         MinDate: Date;
//         VClosingQuantity: Decimal;
//         cnt: Integer;
//         "------IPOL/001---": Integer;
//         vQtyRecied: Decimal;
//         Text001: Label 'As of %1';
//         Text000: Label 'Data';
//         Text0001: Label 'Item by Location ';
//         Text0004: Label 'Company Name';
//         Text0005: Label 'Report No.';
//         Text0006: Label 'Report Name';
//         Text0007: Label 'USER Id';
//         Text0008: Label 'Date';
//         "--20Feb2017": Integer;
//         ILE20: Record 32;
//         ILECOUNT: Integer;
//         TempItem: Integer;
//         "----09Mar2017": Integer;
//         LocCode: Code[100];
//         recLOC: Record 14;
//         LocName: Text[250];
//         LocName1: Text[100];

//     ////[Scope('Internal')]
//     procedure MakeExcelInfo()
//     begin
//     end;

//     ////[Scope('Internal')]
//     procedure MakeExcelDataHeader()
//     begin
//         //>>20Feb2016
//         ExcelBuf.AddColumn(Text001, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//2
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//3
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//4
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//5
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//6
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//7
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//8
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//9
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//10
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//11
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//12
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//13
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//14
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//15
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//16
//         ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
//         ExcelBuf.NewRow;
//         ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//2
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//3
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//4
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//5
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//6
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//7
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//8
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//9
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//10
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//11
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//12
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//13
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//14
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//15
//         ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 0);//16
//         ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
//         ExcelBuf.NewRow;
//         //<<20Feb2017

//         ExcelBuf.NewRow;
//         ExcelBuf.AddColumn('Report Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
//         ExcelBuf.AddColumn('RM Detailed Movement Report', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
//         ExcelBuf.NewRow;
//         ExcelBuf.AddColumn('Data As On', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
//         ExcelBuf.AddColumn("Item Ledger Entry".GETFILTER("Item Ledger Entry"."Posting Date"), FALSE, '', TRUE, FALSE, TRUE, '@', 1);
//         ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
//         ExcelBuf.AddColumn(LocName, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//09Mar2017
//         //ExcelBuf.AddColumn("Item Ledger Entry".GETFILTER("Item Ledger Entry"."Location Code"),FALSE,'',TRUE,FALSE,TRUE,'@',1);

//         ExcelBuf.NewRow;
//         ExcelBuf.NewRow;

//         ExcelBuf.NewRow;
//         ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1); //1
//         ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
//         ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
//         ExcelBuf.AddColumn('Item Category Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
//         ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
//         ExcelBuf.AddColumn('Opening Qty. (in KG)', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
//         //IPOL/001 >>>>
//         ExcelBuf.AddColumn('Quantity Received', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
//         //IPOL/001 <<<<
//         ExcelBuf.AddColumn('Quantity Consumed', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
//         ExcelBuf.AddColumn('Production Order No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
//         ExcelBuf.AddColumn('Batch No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
//         ExcelBuf.AddColumn('Qty Transfered', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
//         ExcelBuf.AddColumn('BRT No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
//         ExcelBuf.AddColumn('Closing Qty.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
//         ExcelBuf.AddColumn('Invoice Rate', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
//         ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
//         ExcelBuf.AddColumn('Rate per LT/Kg', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
//         ExcelBuf.AddColumn('Closing Value', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
//     end;

//     ////[Scope('Internal')]
//     procedure MakeExcelDataBody()
//     begin
//         ExcelBuf.NewRow;
//         ExcelBuf.AddColumn(recILE."Posting Date", FALSE, '', FALSE, FALSE, FALSE, 'dd/mm/yyyy', 1); //Date2Test
//         ExcelBuf.AddColumn(recILE."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
//         ExcelBuf.AddColumn(ItemRec.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);
//         ExcelBuf.AddColumn(recILE."Item Category Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);
//         ExcelBuf.AddColumn(LocName1, FALSE, '', FALSE, FALSE, FALSE, '', 1);//09Mar2017
//         //ExcelBuf.AddColumn(recILE."Location Code",FALSE,'',FALSE,FALSE,FALSE,'',1);
//         //ExcelBuf.AddColumn(OpenQty,FALSE,'',FALSE,FALSE,FALSE,'');

//         ExcelBuf.AddColumn(OpenQty2, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
//         //IPOL/001 >>>>
//         ExcelBuf.AddColumn(vQtyRecied, FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
//         //IPOL/001 <<<<
//         ExcelBuf.AddColumn(ABS(QtyConsumed), FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
//         ExcelBuf.AddColumn(ProdOrdrNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);
//         ExcelBuf.AddColumn(recILE."Lot No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);
//         ExcelBuf.AddColumn(ABS(QtyTransfered), FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
//         ExcelBuf.AddColumn(BRTNo, FALSE, '', FALSE, FALSE, FALSE, '', 1);
//         //ExcelBuf.AddColumn(ClosingQty,FALSE,'',FALSE,FALSE,FALSE,'');
//         ExcelBuf.AddColumn(VClosingQuantity, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);
//         ExcelBuf.AddColumn(ROUND(RateofItem, 0.00001), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00000', 0);
//         ExcelBuf.AddColumn(ChargeofItem, FALSE, '', FALSE, FALSE, FALSE, '', 0);
//         ExcelBuf.AddColumn(ROUND((RateofItem + ChargeofItem), 0.00001), FALSE, '', FALSE, FALSE, FALSE, '#,####0.00000', 0);
//         ExcelBuf.AddColumn(ROUND(ClosingQty * (RateofItem + ChargeofItem), 0.00001), FALSE, '', FALSE, FALSE, FALSE, '#,####0', 0);
//     end;

//     // //[Scope('Internal')]
//     procedure CreateExcelbook()
//     begin
//         //ExcelBuf.CreateBook;
//         //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
//         //ExcelBuf.GiveUserControl;
//         //ERROR('');

//         //>>18Feb2017 RB-N
//         ExcelBuf.CreateBook('', 'RM Detailed Movemenent Report');
//         //ExcelBuf.CreateBookAndOpenExcel('',Text000,Text001,COMPANYNAME,USERID);
//         ExcelBuf.CreateBookAndOpenExcel('', 'RM Detailed Movemenent Report', '', '', USERID);
//         ExcelBuf.GiveUserControl;
//         //<<
//     end;
// }


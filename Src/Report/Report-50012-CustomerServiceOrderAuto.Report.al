// report 50012 "Customer Service Order - Auto"
// {
//     DefaultLayout = RDLC;
//     RDLCLayout = './CustomerServiceOrderAuto.rdlc';

//     dataset
//     {
//         dataitem("Sales Header"; "Sales Header")
//         {
//             RequestFilterFields = "No.";
//             column(CompName; CompanyInfo1.Name)
//             {
//             }
//             column(Approved; Approved)
//             {
//             }
//             column(DocNo; "Sales Header"."No.")
//             {
//             }
//             column(PostingDate; "Sales Header"."Posting Date")
//             {
//             }
//             column(SalesName; SalesPerson.Name)
//             {
//             }
//             column(RespName; RespCentre.Name)
//             {
//             }
//             column(DocDate; "Sales Header"."Document Date")
//             {
//             }
//             column(BillName; "Sales Header"."Sell-to Customer No." + ' - ' + CustName)
//             {
//             }
//             column(BillAdd1; "Bill-to Address")
//             {
//             }
//             column(BillAdd2; "Bill-to Address 2")
//             {
//             }
//             column(BillAdd3; "Bill-to City" + ' - ' + "Bill-to Post Code")
//             {
//             }
//             column(BillAdd4; RecState.Description + ', ' + RecCountry.Name)
//             {
//             }
//             column(BillCSTNo; RecCust."C.S.T. No.")
//             {
//             }
//             column(BillTINNo; RecCust."T.I.N. No.")
//             {
//             }
//             column(BillLSTNo; RecCust."L.S.T. No.")
//             {
//             }
//             column(BillECCNo; RecCust."E.C.C. No.")
//             {
//             }
//             column(ShipToName; ShipToName)
//             {
//             }
//             column(ShipAdd1; "Ship-to Address")
//             {
//             }
//             column(ShipAdd2; "Ship-to Address 2")
//             {
//             }
//             column(ShipAdd3; "Ship-to City" + ' - ' + "Ship-to Post Code")
//             {
//             }
//             column(ShipAdd4; ShiptoState + ' , ' + ShiptoCountry)
//             {
//             }
//             column(ShiptoCST; ShiptoCST)
//             {
//             }
//             column(ShiptoTIN; ShiptoTIN)
//             {
//             }
//             column(ShiptoLST; ShiptoLST)
//             {
//             }
//             column(ShipECCNo; ECCNo)
//             {
//             }
//             column(SupplierCode; RecCust."Our Account No.")
//             {
//             }
//             column(PONo; "External Document No.")
//             {
//             }
//             column(PODate; "Order Date")
//             {
//             }
//             column(ReqDeliDate; "Requested Delivery Date")
//             {
//             }
//             column(NotAfterDate; "Promised Delivery Date")
//             {
//             }
//             column(SupplyName; Loc.Name)
//             {
//             }
//             column(OutStanding; OutStanding)
//             {
//             }
//             column(CustBal5; CustBalanceDueLCY[5])
//             {
//             }
//             column(CustBal4; CustBalanceDueLCY[4])
//             {
//             }
//             column(CustBal3; CustBalanceDueLCY[3])
//             {
//             }
//             column(CustBal2; CustBalanceDueLCY[2])
//             {
//             }
//             column(CustBal1; CustBalanceDueLCY[1])
//             {
//             }
//             column(PreparedBy; UserName)
//             {
//             }
//             column(ApprovalName; ApprovalName)
//             {
//             }
//             column(ApprovalDate; recSalesApprovalEntry."Approval Date")
//             {
//             }
//             column(ApprovalTime; recSalesApprovalEntry."Approval Time")
//             {
//             }
//             column(TaxRate; TaxType + ' @ ' + FORMAT(TaxRate) + '% ' + 'Additional Tax/Cess' + ' @ ' + FORMAT(TaxRate1) + '% ' + AddlTaxType + StateForm)
//             {
//             }
//             column(ExciseRate; 'BED @ ' + FORMAT(ExciseRate) + '%  eCess @ ' + FORMAT(ExceCessRate) + '% SHeCess @' + FORMAT(ExcSHeCessRate) + '%')
//             {
//             }
//             column(PaytermDesc; recPayterm.Description)
//             {
//             }
//             column(PayMethodDesc; recPayMethod.Description)
//             {
//             }
//             column(ShipMethodDesc; recShipMethod.Description)
//             {
//             }
//             column(ShipAgentName; recShipAgent.Name)
//             {
//             }
//             column(FreightCharge; "Freight Charges")
//             {
//             }
//             column(FreightType; "Freight Type")
//             {
//             }
//             column(CreditLimit; RecCust."Credit Limit (LCY)")
//             {
//             }
//             column(ApprovalDesc; "Approval Description")
//             {
//             }
//             dataitem(DataItem1000000001; Table37)
//             {
//                 DataItemLink = Document Type=FIELD(Document Type),
//                                Document No.=FIELD(No.);
//                 DataItemTableView = SORTING(Document Type,Document No.,Line No.);
//                 column(LineNo;"Sales Line"."Line No.")
//                 {
//                 }
//                 column(ItmNo;"Sales Line"."No.")
//                 {
//                 }
//                 column(DimensionName;DimensionName)
//                 {
//                 }
//                 column(UOM;UOM)
//                 {
//                 }
//                 column(SalesUOM;SalesUOM)
//                 {
//                 }
//                 column(OrdQty;"Sales Line".Quantity)
//                 {
//                 }
//                 column(QtyPerUOM;QtyPerUOM)
//                 {
//                 }
//                 column(TotalQty;"Quantity (Base)")
//                 {
//                 }
//                 column(MRP;"Sales Line"."MRP Price")
//                 {
//                 }
//                 column(BillingPrice;"List Price")
//                 {
//                 }
//                 column(BasicPrice;"Sales Line"."Basic Price")
//                 {
//                 }
//                 column(ChgsPrimary;"Sales Line"."Freight/Other Chgs. Primary")
//                 {
//                 }
//                 column(ChgsSecondary;"Sales Line"."Freight/Other Chgs. Secondary")
//                 {
//                 }
//                 column(BillingPriceOld;"MRP Price")
//                 {
//                 }
//                 column(LineAMt;"Line Amount"+"Excise Amount"+ProdDisc)
//                 {
//                 }
//                 column(LineCharge;"Charges To Customer"-ProdDisc)
//                 {
//                 }
//                 column(LineTax;"Tax Amount")
//                 {
//                 }
//                 column(ExRate;"Excise Effective Rate")
//                 {
//                 }
//                 column(ExciseAmount;"Excise Amount")
//                 {
//                 }
//                 column(TotalTaxValue;TotalTaxValue)
//                 {
//                 }
//                 column(TotalAmt;"Line Amount"+"Excise Amount"+"Charges To Customer"+"Tax Amount"-"Line Discount Amount")
//                 {
//                 }
//                 column(ItmCrossRefNo;'('+ "Cross-Reference No."+')')
//                 {
//                 }
//                 column(CrossRefNo;"Cross-Reference No.")
//                 {
//                 }

//                 trigger OnAfterGetRecord()
//                 begin

//                     //>>1

//                     IF RecItem.GET("Sales Line"."No.") THEN
//                       UOM := RecItem."Sales Unit of Measure";

//                     CLEAR(SalesUOM);
//                     recItemUOM.RESET;
//                     recItemUOM.SETRANGE(recItemUOM."Item No.","Sales Line"."No.");
//                     recItemUOM.SETRANGE(recItemUOM.Code,UOM);
//                     IF recItemUOM.FINDFIRST THEN
//                     BEGIN
//                     SalesUOM := recItemUOM."Qty. per Unit of Measure";
//                     END;

//                     //Pratyusha
//                     recExcisePostingSetup.RESET;
//                     recExcisePostingSetup.SETRANGE("Excise Bus. Posting Group","Sales Line"."Excise Bus. Posting Group");
//                     recExcisePostingSetup.SETRANGE("Excise Prod. Posting Group","Sales Line"."Excise Prod. Posting Group");
//                     IF recExcisePostingSetup.FINDFIRST THEN BEGIN
//                       ExciseRate := recExcisePostingSetup."BED %";
//                       ExceCessRate := recExcisePostingSetup."eCess %";
//                       ExcSHeCessRate := recExcisePostingSetup."SHE Cess %";
//                     END;
//                     //Pratyusha

//                     CLEAR(AddlTaxType);
//                     StructureOrderDet.RESET;
//                     StructureOrderDet.SETRANGE(StructureOrderDet."Document No.","Sales Header"."No.");
//                     StructureOrderDet.SETRANGE(StructureOrderDet."Tax/Charge Type",StructureOrderDet."Tax/Charge Type"::"Other Taxes");
//                     IF StructureOrderDet.FINDFIRST THEN
//                     AddlTaxType := StructureOrderDet."Tax/Charge Code";

//                     CLEAR(AddTaxRate);
//                     StructureOrderDet.RESET;
//                     StructureOrderDet.SETRANGE(StructureOrderDet."Document No.","Sales Header"."No.");
//                     StructureOrderDet.SETRANGE(StructureOrderDet."Tax/Charge Type",StructureOrderDet."Tax/Charge Type"::"Other Taxes");
//                     StructureOrderDet.SETRANGE(StructureOrderDet."Tax/Charge Code",'ADDL TAX');
//                     IF StructureOrderDet.FINDFIRST THEN
//                     AddTaxRate := StructureOrderDet."Calculation Value";

//                     SalesHdr.RESET;
//                     SalesHdr.SETRANGE(SalesHdr."No.","Sales Line"."Document No.");
//                     SalesHdr.FINDFIRST;

//                     IF (SalesHdr."Ship-to Code" = '') AND (SalesHdr."Gen. Bus. Posting Group" <> 'FOREIGN')THEN BEGIN
//                     State1.SETRANGE(State1.Code,SalesHdr.State);
//                     IF State1.FINDFIRST THEN;
//                     TaxAreaLocation.RESET;
//                     TaxAreaLocation.SETRANGE(TaxAreaLocation.Type,TaxAreaLocation.Type::Customer);
//                     TaxAreaLocation.SETRANGE(TaxAreaLocation."Dispatch / Receiving Location","Sales Line"."Location Code");
//                     TaxAreaLocation.SETRANGE(TaxAreaLocation."Customer / Vendor Location",State1.Code);
//                     TaxAreaLocation.FINDFIRST;
//                     END
//                     ELSE
//                     IF (SalesHdr."Ship-to Code" <> '') AND (SalesHdr."Gen. Bus. Posting Group" <> 'FOREIGN') THEN
//                     BEGIN
//                     ShipToCode.RESET;
//                     ShipToCode.SETRANGE(ShipToCode.Code,SalesHdr."Ship-to Code");
//                     ShipToCode.SETRANGE(ShipToCode."Customer No.",SalesHdr."Sell-to Customer No.");
//                     ShipToCode.FINDFIRST;
//                     State2.RESET;
//                     State2.SETRANGE(State2.Code,ShipToCode.State);
//                     State2.FINDFIRST;
//                     TaxAreaLocation.RESET;
//                     TaxAreaLocation.SETRANGE(TaxAreaLocation.Type,TaxAreaLocation.Type::Customer);
//                     TaxAreaLocation.SETRANGE(TaxAreaLocation."Dispatch / Receiving Location","Sales Line"."Location Code");
//                     TaxAreaLocation.SETRANGE(TaxAreaLocation."Customer / Vendor Location",State2.Code);
//                     TaxAreaLocation.FINDFIRST;
//                     END;

//                     IF NOT "Sales Line"."Free of Cost" THEN //29Apr2017
//                     BEGIN //29Apr2017
//                       TaxAreaLine.RESET;
//                       TaxAreaLine.SETRANGE(TaxAreaLine."Tax Area",TaxAreaLocation."Tax Area Code");
//                       TaxAreaLine.SETRANGE(TaxAreaLine."Calculation Order",1);
//                       IF TaxAreaLine.FINDFIRST THEN
//                        BEGIN
//                        TaxDetails.RESET;
//                        TaxDetails.SETRANGE(TaxDetails."Tax Jurisdiction Code",TaxAreaLine."Tax Jurisdiction Code" );
//                        TaxDetails.SETRANGE(TaxDetails."Tax Group Code","Sales Line"."Tax Group Code");
//                        TaxDetails.SETRANGE(TaxDetails."Form Code",SalesHdr."Form Code");
//                        TaxDetails.SETFILTER(TaxDetails."Effective Date",'<= %1',SalesHdr."Posting Date");
//                        TaxDetails.FINDLAST;
//                        TaxType:=COPYSTR(TaxAreaLocation."Tax Area Code",1,3);
//                        TaxRate :=  TaxDetails."Tax Below Maximum";
//                        END;



//                       TaxAreaLine.RESET;
//                       TaxAreaLine.SETRANGE(TaxAreaLine."Tax Area",TaxAreaLocation."Tax Area Code");
//                       TaxAreaLine.SETRANGE(TaxAreaLine."Calculation Order",2);
//                       IF TaxAreaLine.FINDFIRST THEN
//                        BEGIN
//                        TaxDetails.RESET;
//                        TaxDetails.SETRANGE(TaxDetails."Tax Jurisdiction Code",TaxAreaLine."Tax Jurisdiction Code" );
//                        TaxDetails.SETRANGE(TaxDetails."Tax Group Code","Sales Line"."Tax Group Code");
//                        TaxDetails.SETRANGE(TaxDetails."Form Code",SalesHdr."Form Code");
//                        TaxDetails.SETFILTER(TaxDetails."Effective Date",'<= %1',SalesHdr."Posting Date");
//                        TaxDetails.FINDLAST;
//                        TaxType1 := LOWERCASE('Additional Tax/Cess');
//                        TaxRate1 :=  TaxDetails."Tax Below Maximum";
//                        END;

//                     END; //29Apr2017


//                     TotalTaxValue:="Sales Line"."Tax Amount";
//                     TotalTaxforOrder:=TotalTaxforOrder+TotalTaxValue;

//                     ProdDiscPerlt := 0;
//                     ProdDisc:=0;
//                     MRpMstr.RESET;
//                     MRpMstr.SETRANGE(MRpMstr."Item No.","Sales Line"."No.");
//                     MRpMstr.SETRANGE(MRpMstr."Lot No.","Sales Line"."Lot No.");
//                     IF MRpMstr.FINDFIRST THEN
//                     BEGIN
//                       ProdDiscPerlt := MRpMstr."National Discount";
//                       ProdDisc := ProdDiscPerlt * "Sales Line"."Quantity (Base)";
//                     END;

//                     //RSPL
//                     /*
//                     DimensionName:='';

//                     IF "Inventory Posting Group"='MERCH' THEN BEGIN
//                      recPosteddocDiemension.RESET;
//                      recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Table ID",37);
//                      recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Document No.","Sales Line"."Document No.");
//                      recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Line No.","Sales Line"."Line No.");
//                      recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Dimension Code",'MERCH');
//                       IF recPosteddocDiemension.FINDFIRST THEN
//                          dimensionCode:=recPosteddocDiemension."Dimension Value Code";

//                        recDimensionValue.RESET;
//                        recDimensionValue.SETRANGE(recDimensionValue.Code,dimensionCode);
//                        IF recDimensionValue.FINDFIRST THEN
//                          DimensionName:=recDimensionValue.Name;
//                          DimensionName2:=DimensionName;

//                     END;
//                     {IF DimensionName=''THEN
//                       DimensionName:='IPOL '+Description;  }

//                     //RSPL
//                     *///Commented DimensionTable#357---23Mar2017

//                     //>>NewCode for DimensionValue 23Mar2017
//                     DimensionName:='';

//                     IF "Inventory Posting Group"='MERCH' THEN
//                     BEGIN

//                       recDimSet.RESET;
//                       recDimSet.SETRANGE("Dimension Set ID","Sales Line"."Dimension Set ID");
//                       recDimSet.SETRANGE("Dimension Code",'MERCH');
//                       IF recDimSet.FINDFIRST THEN
//                       BEGIN
//                         recDimSet.CALCFIELDS(recDimSet."Dimension Value Name");

//                         DimensionName  := recDimSet."Dimension Value Name";
//                         DimensionName2 := DimensionName;

//                       END;
//                     END;
//                     //<<NewCode for DimensionValue 23Mar2017

//                     //RSPL For Category show only item description
//                     IF (DimensionName='') AND ("Sales Line"."Item Category Code"<>'CAT01')
//                     AND ("Sales Line"."Item Category Code"<>'CAT04') AND("Sales Line"."Item Category Code"<>'CAT05')
//                     AND ("Sales Line"."Item Category Code"<>'CAT06') AND("Sales Line"."Item Category Code"<>'CAT07')
//                     AND("Sales Line"."Item Category Code"<>'CAT08') AND("Sales Line"."Item Category Code"<>'CAT09')
//                     AND("Sales Line"."Item Category Code"<>'CAT10') AND ("Sales Line"."Item Category Code"<>'CAT15') THEN
//                       DimensionName:='IPOL '+Description                                                         //RSPL008 CAT15 ADDED
//                     //ELSE
//                     //  DimensionName:=Description;

//                     ELSE IF DimensionName<>'' THEN
//                       DimensionName:=DimensionName2
//                     ELSE
//                         DimensionName:=Description;
//                     //RSPL
//                     //<<1

//                 end;

//                 trigger OnPreDataItem()
//                 begin

//                     NAH := COUNT; //23Mar2017
//                 end;
//             }
//             dataitem(DataItem1000000002;Table44)
//             {
//                 DataItemLink = Document Type=FIELD(Document Type),
//                                No.=FIELD(No.);
//                 DataItemTableView = SORTING(Document Type,No.,Line No.);
//                 column(Comment;"Sales Comment Line".Comment)
//                 {
//                 }
//                 column(NAH1;NAH1)
//                 {
//                 }

//                 trigger OnPreDataItem()
//                 begin

//                     NAH1 := NAH + COUNT;//23Mar2017

//                     //MESSAGE('Line Count %1 \ Com Count %2 ',NAH,NAH1);
//                 end;
//             }

//             trigger OnAfterGetRecord()
//             begin

//                 //>>1

//                 CLEAR(TotalQty);
//                 CLEAR(SumOfItemQty);
//                 IF "Sales Header"."Form Code" <> ''
//                   THEN StateForm := 'Against Form ' + FORMAT("Sales Header"."Form Code")
//                   ELSE StateForm := '';
//                 IF "Sales Header"."Form Code" = 'INDINPUT'
//                   THEN StateForm := 'Against Indusrial Input Declaration';

//                 IF "Sales Header"."External Document No."='' THEN
//                 ERROR('External Document No. is Blank for %1, Plz Define Ext.Doc No',"Sales Header"."No.");

//                 PeriodStartDate[5] := TODAY;
//                 //PeriodStartDate[6] := 12319998D;
//                 PeriodStartDate[6] :=99981231D //31129998D; date formate commented by pcpl-064 6dec2023 //23Mar2017
//                 FOR i := 4 DOWNTO 2 DO
//                  PeriodStartDate[i]
//                  := CALCDATE('<-30D>',PeriodStartDate[i + 1]);


//                 EndDate:=TODAY;
//                  //to get payment terms name
//                  Customer.GET("Sales Header"."Sell-to Customer No.");
//                  paymnettermsrec.SETRANGE(paymnettermsrec.Code,Customer."Payment Terms Code");
//                  IF paymnettermsrec.FINDFIRST THEN
//                  creditvalue:=paymnettermsrec.Description;

//                  PeriodStartDate[1]:=010101D ;//20010101D;
//                  DateExp:=('-' + FORMAT(paymnettermsrec."Due Date Calculation"));

//                  testdate:=CALCDATE(DateExp,EndDate);

//                  //set filter on posting date and then loop through detailed cust. ledger entry table based on cust. ledger table
//                  //then split the balance based on the aging
//                  CustLedgEntry.RESET;
//                  CustLedgEntry.SETCURRENTKEY("Customer No.","Posting Date");
//                  CustLedgEntry.SETRANGE("Customer No.","Sales Header"."Sell-to Customer No.");
//                  CustLedgEntry.SETRANGE("Posting Date",0D,TODAY);
//                  IF CustLedgEntry.FIND('-') THEN
//                   REPEAT
//                      DtldCustLedgEntry.RESET;
//                      DtldCustLedgEntry.SETCURRENTKEY("Cust. Ledger Entry No.","Posting Date");
//                      DtldCustLedgEntry.SETRANGE("Cust. Ledger Entry No.",CustLedgEntry."Entry No.");

//                      IF DtldCustLedgEntry.FIND('-') THEN
//                       REPEAT
//                        IF DtldCustLedgEntry."Posting Date"<=EndDate THEN BEGIN
//                          IF (testdate-CustLedgEntry."Posting Date")<=0 THEN
//                             CustBalanceDueLCY[5]:=CustBalanceDueLCY[5]+DtldCustLedgEntry."Amount (LCY)"
//                          ELSE IF ((testdate-CustLedgEntry."Posting Date")>=1) AND ((testdate-CustLedgEntry."Posting Date")<=30)THEN
//                             CustBalanceDueLCY[4]:=CustBalanceDueLCY[4]+DtldCustLedgEntry."Amount (LCY)"
//                          ELSE IF ((testdate-CustLedgEntry."Posting Date")>=31) AND ((testdate-CustLedgEntry."Posting Date")<=60) THEN
//                             CustBalanceDueLCY[3]:=CustBalanceDueLCY[3]+DtldCustLedgEntry."Amount (LCY)"
//                          ELSE IF ((testdate-CustLedgEntry."Posting Date")>=61) AND ((testdate-CustLedgEntry."Posting Date")<=90) THEN
//                             CustBalanceDueLCY[2]:=CustBalanceDueLCY[2]+DtldCustLedgEntry."Amount (LCY)"
//                          ELSE IF ((testdate-CustLedgEntry."Posting Date")>=91) THEN
//                             CustBalanceDueLCY[1]:=CustBalanceDueLCY[1]+DtldCustLedgEntry."Amount (LCY)";
//                        END;
//                       UNTIL DtldCustLedgEntry.NEXT=0;
//                   UNTIL CustLedgEntry.NEXT=0;


//                 OutStanding:=CustBalanceDueLCY[1]+CustBalanceDueLCY[2] +CustBalanceDueLCY[3] + CustBalanceDueLCY[4] +CustBalanceDueLCY[5];

//                 UserName   :='';
//                 recuser.RESET;
//                 //recuser.SETRANGE(recuser."User ID","Sales Header"."Created By");
//                 recuser.SETRANGE(recuser."User Name","Sales Header"."Created By");//23Mar2017
//                 IF recuser.FINDFIRST THEN
//                   UserName := recuser."Full Name";//23Mar2017
//                   //UserName := recuser.Name;

//                 IF "Sales Header"."Campaign No." = 'APPROVED' THEN
//                 BEGIN
//                 ApprovalName:='';
//                 recSalesAppEntry.RESET;
//                 recSalesAppEntry.SETRANGE("Document No.","Sales Header"."No.");
//                 IF recSalesAppEntry.FINDFIRST THEN
//                   ApprovalName := recSalesAppEntry."Approver Name";
//                 END;
//                 //<<1

//                 //>>2

//                 //Sales Header, Header (1) - OnPreSection()
//                 recSalesApprovalEntry.RESET;
//                 recSalesApprovalEntry.SETRANGE(recSalesApprovalEntry."Document No.","Sales Header"."No.");
//                 recSalesApprovalEntry.SETRANGE(recSalesApprovalEntry.Approved,TRUE);
//                 IF recSalesApprovalEntry.FINDFIRST THEN
//                 Approved := 'APPROVED'
//                 ELSE
//                 Approved := '';
//                 //<<2

//                 //>>3

//                 //Sales Header, Header (2) - OnPreSection()
//                 IF "Sales Header"."Form Code" <> 'CT-1&H' THEN BEGIN
//                 SalesPerson.GET("Sales Header"."Salesperson Code");
//                 END;
//                 RespCentre.GET("Sales Header"."Responsibility Center");

//                 IF "Bill-to Country/Region Code" <> ''
//                   THEN RecCountry.GET("Bill-to Country/Region Code")
//                   ELSE RecCountry.Name := '';

//                 IF RecState.GET(State) THEN
//                 Loc.GET("Sales Header"."Location Code");

//                 RecCust.GET("Sales Header"."Sell-to Customer No.");
//                 IF RecCust."Full Name" <> ''
//                   THEN CustName := RecCust."Full Name"
//                   ELSE CustName := "Sales Header"."Sell-to Customer Name";

//                 IF "Sales Header"."Ship-to Code" = '' THEN
//                 BEGIN
//                 ShiptoState := RecState.Description;
//                 END ELSE
//                 recShip2Add.RESET;
//                 recShip2Add.SETRANGE(recShip2Add."Customer No.","Sales Header"."Sell-to Customer No.");
//                 recShip2Add.SETRANGE(recShip2Add.Code,"Sales Header"."Ship-to Code");
//                 IF recShip2Add.FINDFIRST THEN
//                 IF RecState1.GET(recShip2Add.State) THEN
//                 BEGIN
//                 ShiptoState := RecState1.Description;
//                 END;

//                 IF "Sales Header"."Ship-to Code" = '' THEN
//                 BEGIN
//                 ShiptoCountry := RecCountry.Name;
//                 END ELSE
//                 recShip2Add.RESET;
//                 recShip2Add.SETRANGE(recShip2Add."Customer No.","Sales Header"."Sell-to Customer No.");
//                 recShip2Add.SETRANGE(recShip2Add.Code,"Sales Header"."Ship-to Code");
//                 IF recShip2Add.FINDFIRST THEN
//                 IF RecCountry1.GET(recShip2Add."Country/Region Code") THEN
//                 BEGIN
//                 ShiptoCountry := RecCountry1.Name;;
//                 END;


//                 IF NOT RecCountry1.GET("Sales Header"."Ship-to Country/Region Code")
//                   THEN RecCountry1.GET("Bill-to Country/Region Code");

//                 IF ShipToCode.GET("Sell-to Customer No.","Ship-to Code")
//                   THEN
//                     BEGIN
//                       ECCNo := ShipToCode."E.C.C. No.";
//                       ShipToName := ShipToCode.Name;
//                       ShiptoCST := ShipToCode."C.S.T. No.";
//                       ShiptoLST := ShipToCode."L.S.T. No.";
//                       ShiptoTIN := ShipToCode."T.I.N. No.";
//                     END
//                   ELSE
//                     BEGIN
//                       ECCNo := RecCust."E.C.C. No.";
//                       ShipToName := CustName;
//                       ShiptoCST := RecCust."C.S.T. No.";
//                       ShiptoLST := RecCust."L.S.T. No.";
//                       ShiptoTIN := RecCust."T.I.N. No.";
//                     END;
//                 //<<3


//                 //>>4

//                 //Sales Header, Footer (4) - OnPreSection()
//                 IF "Sales Line"."Excise Amount"=0 THEN BEGIN
//                    ExciseRate := 0;
//                    ExceCessRate := 0;
//                    ExcSHeCessRate := 0;
//                 END;
//                 //<<4

//                 //>>23Mar2017 recPayterm
//                 IF recPayterm.GET("Payment Terms Code") THEN;
//                 //<<23Mar2017 recPayterm

//                 //>>23Mar2017 recPayMethod
//                 IF recPayMethod.GET("Payment Method Code") THEN;
//                 //<<23Mar2017 recPayMethod

//                 //>>23Mar2017 recShipMethod
//                 IF recShipMethod.GET("Shipment Method Code") THEN;
//                 //<<23Mar2017 recShipMethod

//                 //>>23Mar2017 recShipAgent
//                 IF recShipAgent.GET("Shipping Agent Code") THEN;
//                 //<<23Mar2017 recShipAgent
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

//     trigger OnPreReport()
//     begin

//         //>>1
//         CompanyInfo1.GET;//23Mar2017
//         CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
//         CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
//         CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
//         //<<1
//     end;

//     var
//         Loc: Record 14;
//         SalesPerson: Record 13;
//         RespCentre: Record 5714;
//         RecState: Record 13762;
//         RecState1: Record 13762;
//         RecCountry: Record 9;
//         RecCountry1: Record 9;
//         CompanyInfo1: Record 79;
//         RecCust: Record 18;
//         RecItem: Record 27;
//         ShipToCode: Record 222;
//         CustName: Text[100];
//         ShipToName: Text[100];
//         UOM: Code[25];
//         Qty: Decimal;
//         QtyPerUOM: Decimal;
//         TaxType: Code[10];
//         TaxRate: Decimal;
//         ExciseRate: Decimal;
//         ExceCessRate: Decimal;
//         ExcSHeCessRate: Decimal;
//         StateForm: Text[40];
//         ECCNo: Code[20];
//         ShiptoCST: Code[50];
//         ShiptoLST: Code[50];
//         ShiptoTIN: Code[50];
//         Notax: Text[100];
//         ExcisePostingsetup: Record 13711;
//         TotalExciseValue: Decimal;
//         ExciseValue: Decimal;
//         SalesHdr: Record 36;
//         TaxAreaLocation: Record 13761;
//         TaxDetails: Record 322;
//         State1: Record 13762;
//         State2: Record 13762;
//         TotalTaxValue: Decimal;
//         TotalExciseforOrder: Decimal;
//         TotalTaxforOrder: Decimal;
//         TaxAreaLine: Record 319;
//         LineAmtforAL: Decimal;
//         Item: Record 27;
//         TotalQty: Decimal;
//         SumOfItemQty: Decimal;
//         DetailedCustLedgerENtry: Record 21;
//         StartDate: array [6] of Date;
//         EndDate: Date;
//         BalanceOfCust: array [6] of Decimal;
//         Approved: Text[30];
//         StructureOrderDet: Record 13794;
//         paymnettermsrec: Record 3;
//         PeriodStartDate: array [6] of Date;
//         DateExp: Text[40];
//         CustLedgEntry: Record 21;
//         TempCustLedgEntry: Record 21 temporary;
//         tempDtldCustLedgEntry: Record 379 temporary;
//         test: Text[30];
//         testdate: Date;
//         ExcelBuf: Record 370 temporary;
//         PrintToExcel: Boolean;
//         i: Integer;
//         Customer: Record 18;
//         creditvalue: Text[50];
//         EndDate2: Date;
//         DtldCustLedgEntry: Record 379;
//         CustBalanceDueLCY: array [5] of Decimal;
//         OutStanding: Decimal;
//         UserName: Text[30];
//         ApprovalName: Text[30];
//         User: Record 2000000120;
//         recSalesAppEntry: Record 50009;
//         recExcisePostingSetup: Record 13711;
//         AddTaxRate: Decimal;
//         AddlTaxType: Code[10];
//         BillingPrice: Decimal;
//         TotalBillingPrice: Decimal;
//         recItemUOM: Record 5404;
//         SalesUOM: Decimal;
//         recSalesApprovalEntry: Record 50009;
//         recuser: Record 2000000120;
//         recShip2Add: Record "222";
//         ShiptoState: Text[30];
//         ShiptoCountry: Text[30];
//         TaxType1: Code[50];
//         TaxRate1: Decimal;
//         ProdDisc: Decimal;
//         ProdDiscPerlt: Decimal;
//         MRpMstr: Record 50013;
//         Salesprice: Record 7002;
//         "-----------------------RSPL": Integer;
//         recDimensionValue: Record 349;
//         DimensionName: Text[100];
//         dimensionCode: Code[20];
//         DimensionName2: Text[100];
//         "----23Mar2017": Integer;
//         recPayterm: Record 3;
//         recPayMethod: Record 289;
//         recShipMethod: Record 10;
//         recShipAgent: Record 291;
//         NAH: Integer;
//         NAH1: Integer;
//         recDimSet: Record 480;
//         "----------07July2017": Integer;
//         ShiptoGST: Code[15];
//         DetailGST: Record 16412;
//         CGST07: Decimal;
//         CGST07Per: Decimal;
//         SGST07: Decimal;
//         SGST07Per: Decimal;
//         IGST07: Decimal;
//         IGST07Per: Decimal;
// }


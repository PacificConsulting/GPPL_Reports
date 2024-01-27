// report 50021 "Purchase Order"
// {
//     DefaultLayout = RDLC;
//     RDLCLayout = './PurchaseOrder.rdlc';

//     dataset
//     {
//         dataitem(PH; "Purchase Header")
//         {
//             DataItemTableView = SORTING("Document Type", "No.");
//             RequestFilterFields = "No.";
//             column(DocNo; "No.")
//             {
//             }
//             column(PostingDate; PH."Posting Date")
//             {
//             }
//             column(OrderDate; PH."Order Date")
//             {
//             }
//             column(VendorPayName; PH."Pay-to Name")
//             {
//             }
//             column(VendorBillLoc; PH."Pay-to Address" + ' ' + PH."Pay-to Address 2" + ' ' + PH."Pay-to City" + ' ' + PH."Pay-to Post Code" + ' ' + RecState.Description + ',  ' + RecCountry.Name)
//             {
//             }
//             column(VendorSupplyLoc; SupplyLoc)
//             {
//             }
//             column(VendorReference; PH."Vendor Order No.")
//             {
//             }
//             column(SPLLoc; recLoc.Address + ' ' + recLoc."Address 2" + ' ' + recLoc.City + ' ' + recLoc."Post Code" + ' ' + RecCountryLoc.Name + ' ' + RecStateLoc.Description)
//             {
//             }
//             column(VendorContact; PH."Pay-to Contact")
//             {
//             }
//             column(LOCTINNo; recLoc."T.I.N. No.")
//             {
//             }
//             column(CompName; CompanyInfo1.Name)
//             {
//             }
//             column(CompCSTNo; CompanyInfo1."C.S.T No.")
//             {
//             }
//             column(CompLBTNo; CompanyInfo1."LBT No.")
//             {
//             }
//             column(ECCNo; recECC.Description)
//             {
//             }
//             column(ECCRange; recECC."C.E. Range")
//             {
//             }
//             column(ECCDiv; recECC."C.E. Division")
//             {
//             }
//             column(ECCComm; recECC."C.E. Commissionerate")
//             {
//             }
//             column(PromisedDate; "Promised Receipt Date")
//             {
//             }
//             column(FreightCharge; PH."Freight Charges")
//             {
//             }
//             column(CurrCode; PH."Currency Code")
//             {
//             }
//             column(PayTerms; recPayTerm.Description)
//             {
//             }
//             column(Transport; recShip.Description)
//             {
//             }
//             column(ShipAgent; recShipAgent.Name)
//             {
//             }
//             column(VenAmt; VenAmt)
//             {
//             }
//             column(VenChqNo; VenChqNo)
//             {
//             }
//             column(VenPostingDate; VenPostingDate)
//             {
//             }
//             column(TaxType; TaxType)
//             {
//             }
//             column(TaxRate; TaxRate)
//             {
//             }
//             column(StateForm; StateForm)
//             {
//             }
//             column(ClosingDate; PH."Closing Date")
//             {
//             }
//             column(TCCaption; TCCaption)
//             {
//             }
//             dataitem("Purchase Line"; "Purchase Line")
//             {
//                 DataItemLink = "Document Type" = FIELD("Document Type"),
//                                "Document No." = FIELD("No.");
//                 DataItemTableView = SORTING("Document Type", "Document No.", "Line No.");
//                 column(PLDocNo; "Purchase Line"."Document No.")
//                 {
//                 }
//                 column(ItmNo; "Purchase Line"."No.")
//                 {
//                 }
//                 column(VendorItmNo; "Vendor Item No.")
//                 {
//                 }
//                 column(Packging; Packging)
//                 {
//                 }
//                 column(ItmQty; Quantity)
//                 {
//                 }
//                 column(ItmUOM; "Unit of Measure Code")
//                 {
//                 }
//                 column(ItmUnitCost; "Purchase Line"."Direct Unit Cost")
//                 {
//                 }
//                 column(ItmLineAmt; "Purchase Line"."Line Amount")
//                 {
//                 }
//                 column(i; i)
//                 {
//                 }
//                 column(itmSheCess; "SHE Cess Amount")
//                 {
//                 }
//                 column(ItmeCess; "eCess Amount")
//                 {
//                 }
//                 column(ItmBED; "BED Amount")
//                 {
//                 }
//                 column(ItmAED; "ADE Amount")
//                 {
//                 }
//                 column(ItmFinalAmt; "Amount To Vendor")
//                 {
//                 }
//                 column(TaxAmountEBT; TaxAmountEBT)
//                 {
//                 }
//                 column(ServiceTaxRate; ServiceTaxRate)
//                 {
//                 }
//                 column(ServiceTaxEcessRate; ServiceTaxEcessRate)
//                 {
//                 }
//                 column(ServiceTaxSheRate; ServiceTaxSheRate)
//                 {
//                 }
//                 column(ExciseRate; ExciseRate)
//                 {
//                 }
//                 column(ExceCessRate; ExceCessRate)
//                 {
//                 }
//                 column(ExceSheCessRate; ExceSheCessRate)
//                 {
//                 }
//                 column(NAH; NAH)
//                 {
//                 }

//                 trigger OnAfterGetRecord()
//                 begin
//                     i += 1;//28Feb2017
//                     //>>1

//                     //TaxType := "Tax Area Code";
//                     //TaxRate := "Tax %";





//                     TaxAmountEBT := "Purchase Line"."Service Tax Amount" + "Purchase Line"."Service Tax eCess Amount" +
//                                      "Purchase Line"."Service Tax SHE Cess Amount"
//                                     + "Purchase Line"."Tax Amount";

//                     recServiceTax.RESET;
//                     recServiceTax.SETRANGE(recServiceTax.Code, "Purchase Line"."Service Tax Group");
//                     recServiceTax.SETFILTER(recServiceTax."From Date", '<=%1', PH."Document Date");
//                     IF recServiceTax.FINDLAST THEN
//                         ServiceTaxRate := recServiceTax."Service Tax %";
//                     ServiceTaxEcessRate := recServiceTax."eCess %";
//                     ServiceTaxSheRate := recServiceTax."SHE Cess %";

//                     //IF ("Purchase Line".Type = "Purchase Line".Type :: Item) AND ("Purchase Line"."Excise Base Amount" <> 0) THEN // To Show BED Ecess
//                     //SHE Cess in G/L Account Also

//                     IF "Purchase Line"."Excise Amount" <> 0 THEN BEGIN
//                         recExcise.RESET;
//                         recExcise.SETRANGE(recExcise."Excise Bus. Posting Group", "Purchase Line"."Excise Bus. Posting Group");
//                         recExcise.SETRANGE(recExcise."Excise Prod. Posting Group", "Purchase Line"."Excise Prod. Posting Group");
//                         IF recExcise.FINDFIRST THEN
//                             ExciseRate := recExcise."BED %";
//                         ExceCessRate := recExcise."eCess %";
//                         ExceSheCessRate := recExcise."SHE Cess %";
//                     END;


//                     ///Rspl Sachin
//                     RecTaxDeatils.RESET;
//                     RecTaxDeatils.SETRANGE(RecTaxDeatils."Tax Jurisdiction Code", "Purchase Line"."Tax Area Code");
//                     RecTaxDeatils.SETRANGE(RecTaxDeatils."Tax Group Code", "Purchase Line"."Tax Group Code");
//                     RecTaxDeatils.SETFILTER(RecTaxDeatils."Effective Date", '>%1', "Purchase Line"."Posting Date");
//                     IF "Purchase Line"."Form Code" <> '' THEN
//                         RecTaxDeatils.SETRANGE(RecTaxDeatils."Form Code", "Purchase Line"."Form Code");
//                     IF RecTaxDeatils.FINDFIRST THEN BEGIN
//                         TaxRate := RecTaxDeatils."Tax Below Maximum";
//                     END;
//                     //<<1

//                     //>>2

//                     //Purchase Line, Body (2) - OnPreSection()
//                     IF "Vendor Item No." = '' THEN
//                         "Vendor Item No." := "Purchase Line"."No.";

//                     //IF Type = Type::"G/L Account" THEN
//                     //"Vendor Item No." := COPYSTR("Purchase Line".Description,1,29);

//                     IF "Purchase Line"."Bulk Packing" THEN BEGIN
//                         Packging := 'BULK';
//                     END ELSE
//                         Packging := '';

//                     //<<2
//                 end;

//                 trigger OnPreDataItem()
//                 begin
//                     i := 0;//28Feb2017

//                     NAH := COUNT; //05Apr2017
//                 end;
//             }
//             dataitem("Purch. Comment Line"; "Purch. Comment Line")
//             {
//                 DataItemLink = "Document Type" = FIELD("Document Type"),
//                                "No." = FIELD("No.");
//                 DataItemLinkReference = PH;
//                 DataItemTableView = SORTING("Document Type", "No.", "Line No.");
//                 column(CommDocNo; "Purch. Comment Line"."No.")
//                 {
//                 }
//                 column(Comment_PurchCommentLine; "Purch. Comment Line".Comment)
//                 {
//                 }
//                 column(NAH1; NAH1)
//                 {
//                 }

//                 trigger OnPreDataItem()
//                 begin

//                     NAH1 := COUNT;
//                 end;
//             }

//             trigger OnAfterGetRecord()
//             begin
//                 //>>1

//                 IF "Location Code" = '' THEN
//                     ERROR('Please define the Recieving Location');

//                 IF "Form Code" <> '' THEN
//                     StateForm := 'Against Form ' + FORMAT("Form Code")
//                 ELSE
//                     StateForm := '';

//                 //<<1

//                 //>>2

//                 //PH, Header (2) - OnPreSection()
//                 IF "Pay-to Country/Region Code" <> '' THEN
//                     RecCountry.GET("Pay-to Country/Region Code")
//                 ELSE
//                     RecCountry.Name := '';

//                 IF State <> '' THEN
//                     RecState.GET(State)
//                 ELSE
//                     IF PH."Pay-to Country/Region Code" = 'IN' THEN
//                         ERROR('Please give State Code');

//                 //<<2

//                 //>>28Feb2017

//                 IF recOrderAdd.GET(PH."Buy-from Vendor No.", PH."Order Address Code") THEN;

//                 IF recLoc.GET(PH."Location Code") THEN;

//                 IF recECC.GET(recLoc."E.C.C. No.") THEN;

//                 IF recPayTerm.GET(PH."Payment Terms Code") THEN;

//                 IF recShip.GET(PH."Shipment Method Code") THEN;

//                 IF recShipAgent.GET(PH."Shipping Agent Code") THEN;
//                 //<<

//                 //>>3

//                 //Location, Footer (2) - OnPreSection()
//                 //IF "Order Address"."Country/Region Code" <> ''
//                 //THEN RecCountry.GET("Order Address"."Country/Region Code")
//                 //ELSE RecCountry.Name := '';

//                 IF recOrderAdd."Country/Region Code" <> '' THEN
//                     RecCountry.GET(recOrderAdd."Country/Region Code")
//                 ELSE
//                     RecCountry.Name := '';

//                 //IF "Order Address".State <> ''
//                 //THEN RecState.GET("Order Address".State)
//                 //ELSE RecState.Description := '';

//                 IF recOrderAdd.State <> '' THEN
//                     RecState.GET(recOrderAdd.State)
//                 ELSE
//                     RecState.Description := '';

//                 //IF PH."Order Address Code" = ''
//                 //THEN SupplyLoc := PH."Pay-to Address"+' '+PH."Pay-to Address 2"+' '+PH."Pay-to City"+' '+PH."Pay-to Post Code"+
//                 //                   ' '+ RecState.Description+',  '+RecCountry.Name
//                 //ELSE SupplyLoc := "Order Address".Name+' '+"Order Address".Address+' '+"Order Address"."Address 2"+' '+"Order Address".City
//                 //                  +' '+"Order Address"."Post Code"+' '+RecState.Description+',  '+RecCountry.Name;

//                 IF PH."Order Address Code" = '' THEN
//                     SupplyLoc := PH."Pay-to Address" + ' ' + PH."Pay-to Address 2" + ' ' + PH."Pay-to City" + ' ' + PH."Pay-to Post Code" + ' ' + RecState.Description + ',  ' + RecCountry.Name
//                 ELSE
//                     SupplyLoc := recOrderAdd.Name + ' ' + recOrderAdd.Address + ' ' + recOrderAdd."Address 2" + ' ' + recOrderAdd.City + ' ' + recOrderAdd."Post Code" + ' ' + RecState.Description + ',  ' + RecCountry.Name;

//                 //IF "Country/Region Code" <> ''
//                 //THEN RecCountryLoc.GET("Country/Region Code")
//                 //ELSE RecCountryLoc.Name := '';

//                 IF recLoc."Country/Region Code" <> '' THEN
//                     RecCountryLoc.GET(recLoc."Country/Region Code")
//                 ELSE
//                     RecCountryLoc.Name := '';

//                 //IF Location."State Code" <> ''
//                 //THEN RecStateLoc.GET(Location."State Code")
//                 //ELSE RecStateLoc.Description := '';

//                 IF recLoc."State Code" <> '' THEN
//                     RecStateLoc.GET(recLoc."State Code")
//                 ELSE
//                     RecStateLoc.Description := '';


//                 //<<3

//                 //>>4

//                 //PH, Footer (4) - OnPreSection()
//                 /*
//                 IF "Vendor Ledger Entry".Amount = 0 THEN BEGIN
//                   "Vendor Ledger Entry".Amount := PH."Paid Amount";
//                   "Vendor Ledger Entry"."External Document No." := PH."Paid Cheque No.";
//                   "Vendor Ledger Entry"."Posting Date" := PH."Paid Cheque Date";
//                 END;
//                 Vendor No.=FIELD(Buy-from Vendor No.),Document Type=FIELD(Applies-to Doc. Type),Document No.=FIELD(Applies-to Doc. No.)


//                 IF "Vendor Ledger Entry".Amount = 0 THEN BEGIN
//                   "Vendor Ledger Entry".Amount := PH."Paid Amount";
//                   "Vendor Ledger Entry"."External Document No." := PH."Paid Cheque No.";
//                   "Vendor Ledger Entry"."Posting Date" := PH."Paid Cheque Date";
//                 END;


//                 TaxType:=COPYSTR(TaxType,1,3);
//                 *///Commented vendorLedgerEntry
//                   //<<4

//                 //>> 01Mar2017
//                 recVLE.RESET;
//                 recVLE.SETRANGE("Vendor No.", "Buy-from Vendor No.");
//                 recVLE.SETRANGE("Document Type", "Applies-to Doc. Type");
//                 recVLE.SETRANGE("Document No.", "Applies-to Doc. No.");
//                 IF recVLE.FINDFIRST THEN BEGIN
//                     IF recVLE.Amount = 0 THEN BEGIN
//                         VenAmt := PH."Paid Amount";
//                         VenChqNo := PH."Paid Cheque No.";
//                         VenPostingDate := PH."Paid Cheque Date";
//                     END ELSE BEGIN
//                         VenAmt := recVLE.Amount;
//                         VenChqNo := recVLE."External Document No.";
//                         VenPostingDate := recVLE."Posting Date";
//                     END;

//                 END;
//                 //<<


//                 //>>

//                 "Purchase Line".SETRANGE("Purchase Line"."Document No.", PH."No.");
//                 IF "Purchase Line".FINDFIRST THEN BEGIN
//                     //TaxRate := ROUND("Purchase Line"."Tax %",1,'=');
//                     TaxRate := "Purchase Line"."Tax %";
//                     TaxType := "Purchase Line"."Tax Area Code";
//                 END;
//                 //<<

//                 TaxType := COPYSTR(TaxType, 1, 3);

//             end;

//             trigger OnPreDataItem()
//             begin
//                 //>>1

//                 CompanyInfo1.GET;
//                 CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
//                 CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
//                 CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");

//                 //<<1
//             end;
//         }
//     }

//     requestpage
//     {

//         layout
//         {
//             area(content)
//             {
//                 field("Check for TC Caption"; TCCaption)
//                 {
//                 }
//             }
//         }

//         actions
//         {
//         }
//     }

//     labels
//     {
//     }

//     var
//         RecState: Record State;
//         RecCountry: Record 9;
//         RecStateLoc: Record State;
//         RecCountryLoc: Record 9;
//         CompanyInfo1: Record 79;
//         SupplyLoc: Text[180];
//         TaxType: Code[10];
//         TaxRate: Decimal;
//         ExciseRate: Decimal;
//         ExceCessRate: Decimal;
//         ExceSheCessRate: Decimal;
//         StateForm: Text[30];
//         TCCaption: Boolean;
//         TaxAmountEBT: Decimal;
//         ServiceTaxRate: Decimal;
//         ServiceTaxEcessRate: Decimal;
//         ServiceTaxSheRate: Decimal;
//         CompInfo: Record 79;
//         recServiceTax: Record 16472;
//         recExcise: Record 13711;
//         Packging: Code[30];
//         "-------Rspl Sachin------": Integer;
//         RecTaxDeatils: Record 322;
//         "-----28Feb2017": Integer;
//         recOrderAdd: Record 224;
//         recLoc: Record 14;
//         recECC: Record 13708;
//         recPayTerm: Record 3;
//         recShip: Record 10;
//         recShipAgent: Record 291;
//         i: Integer;
//         recVLE: Record 25;
//         VenAmt: Decimal;
//         VenChqNo: Code[20];
//         VenPostingDate: Date;
//         "--05Apr2017": Integer;
//         NAH: Integer;
//         NAH1: Integer;
// }


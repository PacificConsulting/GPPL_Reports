// report 50019 "Vendor Purchase"
// {
//     DefaultLayout = RDLC;
//     RDLCLayout = './VendorPurchase.rdlc';

//     dataset
//     {
//         dataitem("Vendor Ledger Entry"; "Vendor Ledger Entry")
//         {
//             DataItemTableView = SORTING("Vendor No.", "Posting Date")
//                                 WHERE("Document Type" = CONST(Invoice));
//             RequestFilterFields = "Vendor No.";

//             trigger OnAfterGetRecord()
//             begin
//                 //>>RSPL-rahul
//                 i += 1;
//                 CLEAR(vCount);
//                 vVendorLedgerEntry.RESET;
//                 vVendorLedgerEntry.SETRANGE("Vendor No.", "Vendor No.");
//                 vVendorLedgerEntry.SETRANGE("Document Type", vVendorLedgerEntry."Document Type"::Invoice);
//                 vVendorLedgerEntry.SETRANGE("Posting Date", StartDate, EndDate);
//                 vVendorLedgerEntry.SETFILTER("Credit Amount", '>%1', AmountFilter);
//                 IF vVendorLedgerEntry.FINDSET THEN
//                     vCount := vVendorLedgerEntry.COUNT;
//                 //<<RSPL-Rahul

//                 recVendor.GET("Vendor Ledger Entry"."Vendor No.");

//                 recState.RESET;
//                 recState.SETRANGE(recState.Code, recVendor."State Code");
//                 IF recState.FINDFIRST THEN;

//                 recCountry.RESET;
//                 recCountry.SETRANGE(recCountry.Code, recVendor."Country/Region Code");
//                 IF recCountry.FINDFIRST THEN;

//                 recResCentre.RESET;
//                 recResCentre.SETRANGE(recResCentre.Code, recVendor."Responsibility Center");
//                 IF recResCentre.FINDFIRST THEN;

//                 TotalAmount += ABS("Vendor Ledger Entry"."Original Amount");

//                 "Vendor Ledger Entry".CALCFIELDS("Credit Amount");
//                 ;
//                 GrpTotAmt += "Vendor Ledger Entry"."Credit Amount";

//                 IF vCount = i THEN
//                     IF PrintToExcel THEN
//                         MakeExcelDataGroupFooter;
//             end;

//             trigger OnPreDataItem()
//             begin

//                 "Vendor Ledger Entry".SETFILTER("Vendor Ledger Entry"."Posting Date", '%1..%2', StartDate, EndDate);
//                 "Vendor Ledger Entry".SETFILTER("Vendor Ledger Entry"."Credit Amount", '>%1', AmountFilter);
//             end;
//         }
//     }

//     requestpage
//     {

//         layout
//         {
//             area(content)
//             {
//                 field(StartDate; StartDate)
//                 {
//                     Caption = 'Start Date';
//                 }
//                 field(EndDate; EndDate)
//                 {
//                     Caption = 'End Date';
//                 }
//                 field(AmountFilter; AmountFilter)
//                 {
//                     Caption = 'Amount';
//                 }
//                 field(PrintToExcel; PrintToExcel)
//                 {
//                     Caption = 'Print To Excel';
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

//     trigger OnPostReport()
//     begin

//         IF PrintToExcel THEN
//             CreateExcelBook;
//     end;

//     trigger OnPreReport()
//     begin

//         companyinfo.GET;

//         IF (StartDate = 0D) AND (EndDate = 0D) THEN
//             ERROR('Please Specify the Date');
//         IF StartDate = 0D THEN
//             ERROR('Please Specify the Start Date');
//         IF EndDate = 0D THEN
//             ERROR('Please Specify the END Date');

//         IF PrintToExcel THEN
//             MakeExcelInfo;
//     end;

//     var
//         recVendor: Record 23;
//         recState: Record State;
//         recCountry: Record 9;
//         recResCentre: Record 5714;
//         companyinfo: Record 79;
//         AmountFilter: Decimal;
//         StartDate: Date;
//         EndDate: Date;
//         ExcelBuffer: Record 370 temporary;
//         PrintToExcel: Boolean;
//         TotalAmount: Decimal;
//         Text002: Label 'Data';
//         Text003: Label 'Vendor Purchase';
//         Text004: Label 'Company Name';
//         Text005: Label 'Report No.';
//         Text006: Label 'Report Name';
//         Text007: Label 'User ID';
//         Text008: Label 'Date';
//         "--r": Integer;
//         vVendorLedgerEntry: Record 25;
//         vCount: Integer;
//         i: Integer;
//         GrpTotAmt: Decimal;

//     // //[Scope('Internal')]
//     procedure MakeExcelInfo()
//     begin
//         //ExcelBuffer.SetUseInfoSheed;
//         /*
//         ExcelBuffer.AddInfoColumn(FORMAT(Text004),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddInfoColumn(COMPANYNAME,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddInfoColumn(FORMAT(Text006),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddInfoColumn(FORMAT(Text003),FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddInfoColumn(FORMAT(Text005),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddInfoColumn(REPORT::"Vendor Purchase",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddInfoColumn(FORMAT(Text007),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddInfoColumn(USERID,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddInfoColumn(FORMAT(Text008),FALSE,'',TRUE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddInfoColumn(TODAY,FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.NewRow;
//         ExcelBuffer.ClearNewRow;
//         */
//         MakeExcelDataHeader;

//     end;

//     // //[Scope('Internal')]
//     procedure CreateExcelBook()
//     begin
//         //ExcelBuffer.CreateBook;
//         //ExcelBuffer.CreateSheet(Text002,Text003,companyinfo.Name,USERID);
//         //ExcelBuffer.CreateBookAndOpenExcel('','',Text003,'','');
//         ExcelBuffer.CreateBookAndOpenExcel('', Text003, '', '', USERID);
//         //ExcelBuffer.GiveUserControl;
//         ERROR('');
//     end;

//     // //[Scope('Internal')]
//     procedure MakeExcelDataHeader()
//     begin
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddColumn(companyinfo.Name, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         MakeUserIDDetail(TODAY, '');
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddColumn('Vendor - Purchases Over Rs. :', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(AmountFilter, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         MakeUserIDDetail(0D, USERID);
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddColumn('For the Period :', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(FORMAT(StartDate) + '..' + FORMAT(EndDate), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddColumn('Vendor No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('Vendor Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('Vendor Address', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('Vendor City/Post Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('Vendor State', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('Vendor Country', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('Res. Centre Code', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('Res. Centre Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('E.C.C. No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('C.S.T. No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('L.S.T. No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('T.I.N. No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('PAN No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuffer."Cell Type"::Text);
//     end;

//     ////[Scope('Internal')]
//     procedure MakeExcelDataGroupFooter()
//     begin
//         ExcelBuffer.NewRow;
//         ExcelBuffer.AddColumn("Vendor Ledger Entry"."Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor."Full Name", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor.Address + ', ' + recVendor."Address 2", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor.City + ' - ' + recVendor."Post Code", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recState.Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recCountry.Name, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor."Responsibility Center", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recResCentre.Name, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor."E.C.C. No.", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor."C.S.T. No.", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor."L.S.T. No.", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor."T.I.N. No.", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn(recVendor."P.A.N. No.", FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         //"Vendor Ledger Entry".CALCFIELDS("Credit Amount");
//         ExcelBuffer.AddColumn(GrpTotAmt, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', ExcelBuffer."Cell Type"::Number);
//         CLEAR(GrpTotAmt);
//         CLEAR(i);
//     end;

//     local procedure MakeUserIDDetail(PostDate: Date; vUserId: Code[50])
//     begin
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//         ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);

//         IF vUserId = '' THEN BEGIN
//             ExcelBuffer.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', ExcelBuffer."Cell Type"::Text);
//             ExcelBuffer.AddColumn(PostDate, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuffer."Cell Type"::Number);
//         END ELSE
//             ExcelBuffer.AddColumn(vUserId, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', ExcelBuffer."Cell Type"::Number);
//     end;
// }


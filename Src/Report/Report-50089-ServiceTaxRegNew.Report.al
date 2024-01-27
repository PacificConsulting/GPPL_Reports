report 50089 "Service Tax Reg.(New)"
{
    //     // VendorInvoiceNo OLDVALUE--20,,NEWVALUE--35
    //     DefaultLayout = RDLC;
    //     RDLCLayout = './ServiceTaxRegNew.rdlc';


    //     dataset
    //     {
    //         dataitem(DataItem1000000000;service regi)
    //         {
    //             DataItemTableView = SORTING (Document No., Posting Date);

    //             trigger OnAfterGetRecord()
    //             begin

    //                 //>>1

    //                 CLEAR(InvTotal);
    //                 SNo := SNo + 1;


    //                 IF "Document Type" = "Document Type"::Invoice THEN BEGIN
    //                     CLEAR(VendorInvoiceNo);
    //                     CLEAR(VendorInvoiceDate);
    //                     recPIH.RESET;
    //                     recPIH.SETRANGE(recPIH."No.", "Document No.");
    //                     IF recPIH.FINDFIRST THEN BEGIN
    //                         VendorInvoiceNo := recPIH."Vendor Invoice No.";
    //                         VendorInvoiceDate := recPIH."Document Date";
    //                     END;
    //                 END;

    //                 IF "Document Type" = "Document Type"::"Credit Memo" THEN BEGIN
    //                     CLEAR(VendorInvoiceNo);
    //                     CLEAR(VendorInvoiceDate);
    //                     recPCH.RESET;
    //                     recPCH.SETRANGE(recPCH."No.", "Document No.");
    //                     IF recPCH.FINDFIRST THEN BEGIN
    //                         recPIH.RESET;
    //                         recPIH.SETRANGE(recPIH."No.", recPCH."Applies-to Doc. No.");
    //                         IF recPIH.FINDFIRST THEN BEGIN
    //                             VendorInvoiceNo := recPIH."Vendor Invoice No.";
    //                             VendorInvoiceDate := recPIH."Document Date";
    //                         END;
    //                         // RSPL-CAS-05341-Y9Z5N1 Start
    //                         IF VendorInvoiceNo = '' THEN
    //                             VendorInvoiceNo := recPCH."Vendor Cr. Memo No.";
    //                         // RSPL-CAS-05341-Y9Z5N1 Stop
    //                     END;
    //                 END;


    //                 IF Type = Type::Purchase THEN BEGIN
    //                     RecGLentry.RESET;
    //                     RecGLentry.SETRANGE(RecGLentry."Document No.", "Document No.");
    //                     RecGLentry.SETRANGE(RecGLentry."Source Type", RecGLentry."Source Type"::Vendor);
    //                     RecGLentry.SETFILTER(RecGLentry."Source No.", '<>%1', '');
    //                     IF RecGLentry.FINDFIRST THEN BEGIN
    //                         Vendor.GET(RecGLentry."Source No.");
    //                         Name := Vendor.Name;
    //                     END;
    //                 END
    //                 ELSE
    //                     IF Type = Type::Sale THEN BEGIN
    //                         RecGLentry.RESET;
    //                         RecGLentry.SETRANGE(RecGLentry."Document No.", "Document No.");
    //                         RecGLentry.SETRANGE(RecGLentry."Source Type", RecGLentry."Source Type"::Customer);
    //                         RecGLentry.SETFILTER(RecGLentry."Source No.", '<>%1', '');
    //                         IF RecGLentry.FINDFIRST THEN BEGIN
    //                             Customer.GET(RecGLentry."Source No.");
    //                             Name := Customer.Name;
    //                         END;
    //                     END;

    //                 InvTotal := "Base Amount" + "Service Tax Amount" + "Service Tax eCess Amount" + "Service Tax SHE Cess Amount";
    //                 InvoiceTotal := InvoiceTotal + InvTotal;

    //                 RvdPaymentTotal := "Base Amount" + "Service Tax Amount" + "Service Tax eCess Amount" + "Service Tax SHE Cess Amount";
    //                 PaymentTotal := PaymentTotal + RvdPaymentTotal;
    //                 TotalServices := TotalServices + "Service Tax Entry Details"."Base Amount";
    //                 TotalTax := TotalTax + "Service Tax Amount" + "Service Tax eCess Amount" + "Service Tax SHE Cess Amount";
    //                 //<<1


    //                 //>>10May2017 ReportBody

    //                 IF PrintToExcel THEN BEGIN

    //                     EnterCell(NNN, 1, FORMAT(SNo), FALSE, FALSE, '', 0);
    //                     EnterCell(NNN, 2, "Document No.", FALSE, FALSE, '', 1);
    //                     EnterCell(NNN, 3, FORMAT("Posting Date"), FALSE, FALSE, '', 2);
    //                     EnterCell(NNN, 4, Name, FALSE, FALSE, '', 1);
    //                     EnterCell(NNN, 5, FORMAT("Base Amount"), FALSE, FALSE, '#,###0.00', 0);
    //                     TBaseAmt += "Base Amount";//09May2017

    //                     EnterCell(NNN, 6, FORMAT("Service Tax Amount" + "Service Tax eCess Amount" + "Service Tax SHE Cess Amount"), FALSE, FALSE, '#,###0.00', 0);
    //                     TBaseAmt2 += "Service Tax Amount" + "Service Tax eCess Amount" + "Service Tax SHE Cess Amount";

    //                     EnterCell(NNN, 7, FORMAT(InvTotal), FALSE, FALSE, '#,###0.00', 0);
    //                     EnterCell(NNN, 8, FORMAT("Service Tax Entry Details"."Service Type (Rev. Chrg.)"), FALSE, FALSE, '', 1);
    //                     //EnterCell(NNN,8,FORMAT("GTA Service Type"),FALSE,FALSE,'',1);
    //                     EnterCell(NNN, 9, VendorInvoiceNo, FALSE, FALSE, '', 1);
    //                     EnterCell(NNN, 10, FORMAT(VendorInvoiceDate), FALSE, FALSE, '', 2);


    //                     NNN += 1;
    //                 END;
    //                 //<<10May2017 ReportBody

    //                 //>>2

    //                 //Service Tax Entry Details, Body (3) - OnPreSection()
    //                 /*
    //                 IF PrintToExcel THEN
    //                 BEGIN

    //                  XLSHEET.Range('A'+FORMAT(i)).Value := SNo;
    //                  XLSHEET.Range('A'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('A'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('B'+FORMAT(i)).Value := "Document No.";
    //                  XLSHEET.Range('B'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('C'+FORMAT(i)).Value := FORMAT("Posting Date");
    //                  XLSHEET.Range('C'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('C'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('D'+FORMAT(i)).Value := Name;
    //                  XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('E'+FORMAT(i)).Value := "Base Amount";
    //                  XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
    //                  XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('F'+FORMAT(i)).Value := "Service Tax Amount" + "Service Tax eCess Amount" + "Service Tax SHE Cess Amount";
    //                  XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
    //                  XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('G'+FORMAT(i)).Value := InvTotal;
    //                  XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
    //                  XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('H'+FORMAT(i)).Value := FORMAT("GTA Service Type");
    //                  XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('I'+FORMAT(i)).Value := FORMAT(VendorInvoiceNo);
    //                  XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;

    //                  XLSHEET.Range('J'+FORMAT(i)).Value := FORMAT(VendorInvoiceDate);
    //                  XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;


    //                 i += 1;

    //                 END;
    //                 *///Commented 09May2017
    //                   //<<2

    //             end;

    //             trigger OnPostDataItem()
    //             begin


    //                 //>>10May2017 ReportFooter

    //                 IF PrintToExcel THEN BEGIN

    //                     EnterCell(NNN, 1, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 2, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 3, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 4, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 5, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 6, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 7, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 8, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 9, '', FALSE, TRUE, '', 1);
    //                     EnterCell(NNN, 10, '', FALSE, TRUE, '', 1);


    //                     NNN += 1;

    //                     EnterCell(NNN, 1, '', TRUE, TRUE, '', 1);
    //                     EnterCell(NNN, 2, 'TOTAL', TRUE, TRUE, '', 1);
    //                     EnterCell(NNN, 3, '', TRUE, TRUE, '', 1);
    //                     EnterCell(NNN, 4, '', TRUE, TRUE, '', 1);
    //                     EnterCell(NNN, 5, FORMAT(TBaseAmt), TRUE, TRUE, '#,###0.00', 0);
    //                     EnterCell(NNN, 6, FORMAT(TBaseAmt2), TRUE, TRUE, '#,###0.00', 0);
    //                     EnterCell(NNN, 7, FORMAT(InvoiceTotal), TRUE, TRUE, '#,###0.00', 0);
    //                     EnterCell(NNN, 8, '', TRUE, TRUE, '', 1);
    //                     EnterCell(NNN, 9, '', TRUE, TRUE, '', 1);
    //                     EnterCell(NNN, 10, '', TRUE, TRUE, '', 1);


    //                     NNN += 1;
    //                 END;
    //                 //<<10May2017 ReportFooter

    //                 //>>1
    //                 /*
    //                 IF PrintToExcel THEN
    //                 BEGIN

    //                  XLSHEET.Range('A'+FORMAT(i)).Value := '';
    //                  XLSHEET.Range('A'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('A'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('A'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('A'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('B'+FORMAT(i)).Value := 'TOTAL';
    //                  XLSHEET.Range('B'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('B'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('B'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('B'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('C'+FORMAT(i)).Value := '';
    //                  XLSHEET.Range('C'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('C'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('C'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('D'+FORMAT(i)).Value := '';
    //                  XLSHEET.Range('D'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('D'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('D'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('D'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('E'+FORMAT(i)).Value := "Base Amount";
    //                  XLSHEET.Range('E'+FORMAT(i)).NumberFormat := '@';
    //                  XLSHEET.Range('E'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('E'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('E'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('E'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('F'+FORMAT(i)).Value := "Service Tax Amount" + "Service Tax eCess Amount" + "Service Tax SHE Cess Amount";
    //                  XLSHEET.Range('F'+FORMAT(i)).NumberFormat := '@';
    //                  XLSHEET.Range('F'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('F'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('F'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('F'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('G'+FORMAT(i)).Value := InvoiceTotal;
    //                  XLSHEET.Range('G'+FORMAT(i)).NumberFormat := '@';
    //                  XLSHEET.Range('G'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('G'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('G'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('G'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('H'+FORMAT(i)).Value := '';
    //                  XLSHEET.Range('H'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('H'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('H'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('H'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('I'+FORMAT(i)).Value := '';
    //                  XLSHEET.Range('I'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('I'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('I'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('I'+FORMAT(i)).Borders.ColorIndex := 5;

    //                  XLSHEET.Range('J'+FORMAT(i)).Value := '';
    //                  XLSHEET.Range('J'+FORMAT(i)).RowHeight := 18;
    //                  XLSHEET.Range('J'+FORMAT(i)).Font.Size := 10;
    //                  XLSHEET.Range('J'+FORMAT(i)).EntireColumn.AutoFit;
    //                  XLSHEET.Range('J'+FORMAT(i)).Borders.ColorIndex := 5;


    //                 END;
    //                 *///Commented 09May2017
    //                   //<<1

    //             end;

    //             trigger OnPreDataItem()
    //             begin

    //                 //>>1

    //                 IF STGroups.GET(STGroup) THEN
    //                     EndDate := CALCDATE('CM', DMY2DATE(1, (ToMonth + 1), ToYear));
    //                 "Service Tax Entry Details".SETRANGE("Posting Date", DMY2DATE(1, (FromMonth + 1), FromYear), EndDate);
    //                 SETRANGE(Type, "Purch/Sale");
    //                 SETRANGE("Service Tax Group Code", STGroup);
    //                 SETRANGE("Service Tax Registration No.", "ST No.");

    //                 //IF PrintToExcel THEN
    //                 //BEGIN
    //                 // CreateXLSHEET;
    //                 //CreateHeader;
    //                 //i := 11;
    //                 //END;
    //                 //<<1

    //                 //>>09May2017 ReportHeader

    //                 IF PrintToExcel THEN BEGIN

    //                     EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '', 1);
    //                     EnterCell(1, 10, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

    //                     EnterCell(2, 1, 'SERVICE TAX REGISTER', TRUE, FALSE, '', 1);
    //                     EnterCell(2, 10, USERID, TRUE, FALSE, '', 1);

    //                     EnterCell(4, 1, '1. Name of the assessee :', TRUE, FALSE, '', 1);
    //                     EnterCell(4, 2, CompanyInfo.Name, TRUE, FALSE, '', 1);

    //                     EnterCell(5, 1, '2. Category of services', TRUE, FALSE, '', 1);
    //                     EnterCell(5, 2, STGroups.Description, TRUE, FALSE, '', 1);

    //                     EnterCell(6, 1, '3. Purchase/Sale', TRUE, FALSE, '', 1);
    //                     EnterCell(6, 2, FORMAT("Purch/Sale"), TRUE, FALSE, '', 1);

    //                     EnterCell(7, 1, '4. Service tax registration No.', TRUE, FALSE, '', 1);
    //                     EnterCell(7, 2, "ST No.", TRUE, FALSE, '', 1);

    //                     EnterCell(8, 1, '5. Period', TRUE, FALSE, '', 1);
    //                     EnterCell(8, 2, FORMAT(DMY2DATE(1, (FromMonth + 1), FromYear)) + '..' + FORMAT(EndDate), TRUE, FALSE, '', 1);

    //                     EnterCell(10, 1, 'Details of Bills', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 2, '', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 3, '', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 4, '', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 5, '', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 6, '', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 7, '', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 8, '', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 9, '', TRUE, TRUE, '', 1);
    //                     EnterCell(10, 10, '', TRUE, TRUE, '', 1);

    //                     EnterCell(11, 1, 'Sr. No.', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 2, 'Document No.', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 3, 'Document Date', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 4, 'Name of Client', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 5, 'Taxable Services (a)', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 6, 'Service Tax (b)', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 7, 'TOTAL [(a) + (b)]', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 8, 'Status', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 9, 'Vendor Invoice No.', TRUE, TRUE, '', 1);
    //                     EnterCell(11, 10, 'Vendor Invoice Date', TRUE, TRUE, '', 1);

    //                     NNN := 12;

    //                 END;

    //                 //<<09May2017 ReportHeader
    //             end;
    //         }
    //     }

    //     requestpage
    //     {

    //         layout
    //         {
    //             area(content)
    //             {
    //                 field("From Month"; FromMonth)
    //                 {
    //                 }
    //                 field("From Year"; FromYear)
    //                 {
    //                 }
    //                 field("To Month"; ToMonth)
    //                 {
    //                 }
    //                 field("To Year"; ToYear)
    //                 {
    //                 }
    //                 field("Purchase/Sale"; "Purch/Sale")
    //                 {
    //                 }
    //                 field("Service Tax Group"; STGroup)
    //                 {
    //                     TableRelation = "Service Tax Groups";
    //                 }
    //                 field("ST No"; "ST No.")
    //                 {
    //                     TableRelation = "Service Tax Registration Nos.";
    //                 }
    //                 field("Print To Excel"; PrintToExcel)
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

    //     trigger OnInitReport()
    //     begin
    //         //>>1

    //         FromYear := DATE2DMY(TODAY, 3);
    //         ToYear := DATE2DMY(TODAY, 3);
    //         //<<1
    //     end;

    //     trigger OnPostReport()
    //     begin
    //         //>>10May2017
    //         IF PrintToExcel THEN BEGIN
    //             ExcBuffer.CreateBook('', 'SERVICE TAX REGISTER');
    //             ExcBuffer.CreateBookAndOpenExcel('', 'SERVICE TAX REGISTER', '', '', USERID);
    //             ExcBuffer.GiveUserControl;
    //         END;
    //         //<<1
    //     end;

    //     trigger OnPreReport()
    //     begin
    //         //>>1

    //         CLEAR(InvoiceTotal);
    //         CLEAR(PaymentTotal);
    //         CompanyInfo.GET;
    //         CheckDataEntry;
    //         //<<1
    //     end;

    //     var
    //         CompanyInfo: Record "79";
    //         FromMonth: Option January,February,March,April,May,June,July,August,September,October,November,December;
    //         FromYear: Integer;
    //         ToMonth: Option January,February,March,April,May,June,July,August,September,October,November,December;
    //         ToYear: Integer;
    //         EndDate: Date;
    //         "Purch/Sale": Option Sale,Purchase;
    //         STGroup: Code[20];
    //         "ST No.": Code[20];
    //         STGroups: Record "16471";
    //         SNo: Integer;
    //         recPIH: Record "122";
    //         recPCH: Record "124";
    //         VendorInvoiceNo: Code[35];
    //         VendorInvoiceDate: Date;
    //         RecGLentry: Record "17";
    //         Vendor: Record "23";
    //         Customer: Record "18";
    //         Name: Text[100];
    //         InvTotal: Decimal;
    //         InvoiceTotal: Decimal;
    //         RvdPaymentTotal: Decimal;
    //         PaymentTotal: Decimal;
    //         TotalServices: Decimal;
    //         TotalTax: Decimal;
    //         PrintToExcel: Boolean;
    //         MaxCol: Text[3];
    //         i: Integer;
    //         R: Integer;
    //         Text001: Label 'SERVICE TAX REGISTER';
    //         ErrText003: Label 'Please enter correct date period';
    //         Text002: Label 'From';
    //         Text003: Label 'To';
    //         "---09May2017---": Integer;
    //         ExcBuffer: Record "370" temporary;
    //         NNN: Integer;
    //         TBaseAmt: Decimal;
    //         TServiceTaxAmt: Decimal;
    //         TserviceEAmt: Decimal;
    //         TserviceSAmt: Decimal;
    //         TBaseAmt2: Decimal;

    //     [Scope('Internal')]
    //     procedure CheckDataEntry()
    //     begin

    //         IF (((FromYear < 1950) OR (FromYear > 2050)) OR ((ToYear < 1950) OR (ToYear > 2050))) THEN
    //             ERROR(ErrText003);

    //         IF ((ToYear < FromYear) OR ((ToYear - FromYear) > 1)) THEN
    //             ERROR(ErrText003);

    //         IF (ToYear = FromYear) THEN
    //             IF (ToMonth < FromMonth) THEN
    //                 ERROR(ErrText003);

    //         IF (ToYear > FromYear) THEN
    //             IF (ToMonth > FromMonth) THEN
    //                 ERROR(ErrText003);
    //     end;

    //     [Scope('Internal')]
    //     procedure CreateXLSHEET()
    //     begin
    //         /*
    //         CLEAR(XLAPP);
    //         CLEAR(XLWBOOK);
    //         CLEAR(XLSHEET);
    //         CREATE(XLAPP);
    //         XLWBOOK := XLAPP.Workbooks.Add;
    //         XLSHEET := XLWBOOK.Worksheets.Add;
    //         XLAPP.Visible(FALSE);
    //         *///Commented 09May2017
    //         MaxCol := 'J';

    //     end;

    //     [Scope('Internal')]
    //     procedure CreateHeader()
    //     begin
    //         /*
    //         XLSHEET.Activate;

    //         XLSHEET.Range('A1',MaxCol+'1').Merge := TRUE;
    //         XLSHEET.Range('A1',MaxCol+'1').Font.Bold := TRUE;
    //         XLSHEET.Range('A1',MaxCol+'1').Value := 'SERVICE TAX REIGISTER';
    //         XLSHEET.Range('A1',MaxCol+'1').HorizontalAlignment := 3;
    //         XLSHEET.Range('A1',MaxCol+'1').Interior.ColorIndex := 8;
    //         XLSHEET.Range('A1',MaxCol+'1').Borders.ColorIndex := 5;
    //         XLSHEET.Range('A1',MaxCol+'1').Merge := TRUE;
    //         XLSHEET.Range('A1','A2').Merge := TRUE;

    //         XLSHEET.Range('A3').Value := '1. Name of the assessee :';
    //         XLSHEET.Range('A3').Font.Bold := TRUE;
    //         XLSHEET.Range('A3',MaxCol+'3').Borders.ColorIndex := 5;
    //         XLSHEET.Range('A3','D3').Merge := TRUE;

    //         XLSHEET.Range('E3').Value := CompanyInfo.Name;
    //         XLSHEET.Range('E3').Font.Bold := TRUE;
    //         XLSHEET.Range('E3','J3').Merge := TRUE;

    //         XLSHEET.Range('A4').Value := '2. Category of services';
    //         XLSHEET.Range('A4').Font.Bold := TRUE;
    //         XLSHEET.Range('A4',MaxCol+'4').Borders.ColorIndex := 5;
    //         XLSHEET.Range('A4','D4').Merge := TRUE;

    //         XLSHEET.Range('E4').Value := STGroups.Description;
    //         XLSHEET.Range('E4').Font.Bold := TRUE;
    //         XLSHEET.Range('E4','J4').Merge := TRUE;

    //         XLSHEET.Range('A5').Value := '3. Purchase/Sale';
    //         XLSHEET.Range('A5').Font.Bold := TRUE;
    //         XLSHEET.Range('A5',MaxCol+'5').Borders.ColorIndex := 5;
    //         XLSHEET.Range('A5','D5').Merge := TRUE;

    //         XLSHEET.Range('E5').Value := FORMAT("Purch/Sale");
    //         XLSHEET.Range('E5').Font.Bold := TRUE;
    //         XLSHEET.Range('E5','J5').Merge := TRUE;

    //         XLSHEET.Range('A6').Value := '4. Service tax registration No.';
    //         XLSHEET.Range('A6').Font.Bold := TRUE;
    //         XLSHEET.Range('A6',MaxCol+'6').Borders.ColorIndex := 5;
    //         XLSHEET.Range('A6','D6').Merge := TRUE;

    //         XLSHEET.Range('E6').Value := "ST No.";
    //         XLSHEET.Range('E6').Font.Bold := TRUE;
    //         XLSHEET.Range('E6','J6').Merge := TRUE;

    //         XLSHEET.Range('A7').Value := '5. Period';
    //         XLSHEET.Range('A7').Font.Bold := TRUE;
    //         XLSHEET.Range('A7',MaxCol+'7').Borders.ColorIndex := 5;
    //         XLSHEET.Range('A7','D7').Merge := TRUE;

    //         XLSHEET.Range('E7').Value := FORMAT(DMY2DATE(1,(FromMonth+1),FromYear)) + '..' + FORMAT(EndDate);
    //         XLSHEET.Range('E7').Font.Bold := TRUE;
    //         XLSHEET.Range('E7','J7').Merge := TRUE;

    //         XLSHEET.Range('A8','J8').Merge := TRUE;

    //         XLSHEET.Range('A9').Value := 'Details of Bills';
    //         XLSHEET.Range('A9').Font.Bold := TRUE;
    //         XLSHEET.Range('A9').HorizontalAlignment := -4108;
    //         XLSHEET.Range('A9','J9').Borders.ColorIndex := 5;
    //         XLSHEET.Range('A9','J9').Interior.ColorIndex := 8;
    //         XLSHEET.Range('A9','J9').Merge := TRUE;

    //         XLSHEET.Range('A10').Value := 'Sr. No.';
    //         XLSHEET.Range('A10').Font.Bold := TRUE;
    //         XLSHEET.Range('A10').EntireColumn.AutoFit;

    //         XLSHEET.Range('B10').Value := 'Document No.';
    //         XLSHEET.Range('B10').Font.Bold := TRUE;
    //         XLSHEET.Range('B10').EntireColumn.AutoFit;

    //         XLSHEET.Range('C10').Value := 'Document Date';
    //         XLSHEET.Range('C10').Font.Bold := TRUE;
    //         XLSHEET.Range('C10').EntireColumn.AutoFit;

    //         XLSHEET.Range('D10').Value := 'Name of Client';
    //         XLSHEET.Range('D10').Font.Bold := TRUE;
    //         XLSHEET.Range('D10').EntireColumn.AutoFit;

    //         XLSHEET.Range('E10').Value := 'Taxable Services (a)';
    //         XLSHEET.Range('E10').Font.Bold := TRUE;
    //         XLSHEET.Range('E10').EntireColumn.AutoFit;

    //         XLSHEET.Range('F10').Value := 'Service Tax (b)';
    //         XLSHEET.Range('F10').Font.Bold := TRUE;
    //         XLSHEET.Range('F10').EntireColumn.AutoFit;

    //         XLSHEET.Range('G10').Value := 'TOTAL [(a) + (b)]';
    //         XLSHEET.Range('G10').Font.Bold := TRUE;
    //         XLSHEET.Range('G10').EntireColumn.AutoFit;

    //         XLSHEET.Range('H10').Value := 'Status';
    //         XLSHEET.Range('H10').Font.Bold := TRUE;
    //         XLSHEET.Range('H10').EntireColumn.AutoFit;

    //         XLSHEET.Range('I10').Value := 'Vendor Invoice No.';
    //         XLSHEET.Range('I10').Font.Bold := TRUE;
    //         XLSHEET.Range('I10').EntireColumn.AutoFit;

    //         XLSHEET.Range('J10').Value := 'Vendor Invoice Date';
    //         XLSHEET.Range('J10').Font.Bold := TRUE;
    //         XLSHEET.Range('J10').EntireColumn.AutoFit;



    //         XLSHEET.Range('A10',MaxCol+'10').Borders.ColorIndex := 5;

    //         XLAPP.Visible(TRUE);
    //         *///Commented 09May2017

    //     end;

    //     [Scope('Internal')]
    //     procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    //     begin

    //         IF PrintToExcel THEN BEGIN
    //             ExcBuffer.INIT;
    //             ExcBuffer.VALIDATE("Row No.", Rowno);
    //             ExcBuffer.VALIDATE("Column No.", columnno);
    //             ExcBuffer."Cell Value as Text" := Cellvalue;
    //             ExcBuffer.Formula := '';
    //             ExcBuffer.Bold := Bold;
    //             ExcBuffer.Underline := Underline;
    //             ExcBuffer.NumberFormat := NoFormat;
    //             ExcBuffer."Cell Type" := CType;
    //             ExcBuffer.INSERT;
    //         END;

    //         //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    //     end;
}


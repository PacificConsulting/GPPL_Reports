report 50175 "Transport Detail Report"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 14Nov2017   RB-N         CustomerName
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TransportDetailReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Transport Details"; "Transport Details")
        {
            RequestFilterFields = "From Location Code";

            trigger OnAfterGetRecord()
            begin

                //>>1

                //CAS-05340-S2H4K5
                vStatus := 'OPEN';
                cAppliedDocNo := '';
                dPostedOn := 0D;
                dAppliedDate := 0D;
                CLEAR(RatePerKL); // Fahim 01-08-2022
                IF ("Applied Date" >= vStartDate) AND ("Applied Date" <= vEndDate) THEN BEGIN
                    //IF "Applied Date" <= StatusDate THEN BEGIN//RSPLSUM 09Dec19
                    IF NOT vDetailed THEN
                        CurrReport.SKIP;
                    IF "Applied Date" <= StatusDate THEN//RSPLSUM 02Jan20
                        vStatus := 'CLOSED';
                    cAppliedDocNo := "Transport Details"."Applied Document No.";
                    dAppliedDate := "Transport Details"."Applied Date";
                END;
                IF "Applied Date" > vEndDate THEN
                    dPostedOn := "Applied Date";
                //CAS-05340-S2H4K5

                IF "Transport Details"."TPT Invoice Date" = 0D THEN BEGIN
                    TPTINvDate := '';
                END
                ELSE BEGIN
                    TPTINvDate := FORMAT("Transport Details"."TPT Invoice Date");
                END;

                IF "Transport Details"."Local LR Date" = 0D THEN BEGIN
                    LocalLrDate := '';
                END
                ELSE BEGIN
                    LocalLrDate := FORMAT("Transport Details"."Local LR Date");
                END;
                IF "Transport Details"."LR Date" = 0D THEN BEGIN
                    LrDate := ''
                END
                ELSE BEGIN
                    LrDate := FORMAT("Transport Details"."LR Date");
                END;

                Destinatn := '';
                IF "Transport Details".Destination <> '' THEN
                    Destinatn := "Transport Details".Destination;
                IF "Transport Details".Destination = '' THEN
                    Destinatn := "Transport Details"."To Location Name";

                CLEAR(CustName14);//14Nov2017
                invheader.RESET;
                invheader.SETRANGE(invheader."No.", "Transport Details"."Invoice No.");
                IF invheader.FINDFIRST THEN BEGIN
                    CustName14 := invheader."Sell-to Customer Name";
                    IF invheader."Currency Code" <> '' THEN
                        Destinatn := 'EXPORT';
                END;

                IF "Transport Details"."Cancelled Invoice" = TRUE THEN
                    Status := 'Cancelled'
                ELSE
                    Status := 'Active';

                //
                ItemCatCode := '';
                categoryDesc := '';
                CLEAR(cdRegion);//RSPLSUM27May21
                IF "Transport Details".Type = "Transport Details".Type::Transfer THEN BEGIN
                    RecTransferLine.RESET;
                    RecTransferLine.SETRANGE(RecTransferLine."Document No.", "Transport Details"."Invoice No.");
                    IF RecTransferLine.FINDFIRST THEN
                        ItemCatCode := RecTransferLine."Item Category Code";
                    IF RecItemCatgory.GET(ItemCatCode) THEN
                        categoryDesc := RecItemCatgory.Description;
                END;

                IF "Transport Details".Type = "Transport Details".Type::Invoice THEN BEGIN
                    RecSalesInvLine.RESET;
                    RecSalesInvLine.SETRANGE(RecSalesInvLine."Document No.", "Transport Details"."Invoice No.");
                    IF RecSalesInvLine.FINDFIRST THEN
                        ItemCatCode := RecSalesInvLine."Item Category Code";
                    IF RecItemCatgory.GET(ItemCatCode) THEN
                        categoryDesc := RecItemCatgory.Description;
                    //RSPLSUM27May21>>
                    RecSalesInvHeader.RESET;
                    IF RecSalesInvHeader.GET("Transport Details"."Invoice No.") THEN BEGIN
                        IF RecSalesInvHeader."Ship-to Code" <> '' THEN BEGIN
                            RecShipToAddr.RESET;
                            IF RecShipToAddr.GET(RecSalesInvHeader."Sell-to Customer No.", RecSalesInvHeader."Ship-to Code") THEN
                                cdRegion := RecShipToAddr.State;
                        END ELSE BEGIN
                            RecCust.RESET;
                            IF RecCust.GET(RecSalesInvHeader."Sell-to Customer No.") THEN
                                cdRegion := RecCust."State Code";
                        END;
                    END;
                    //RSPLSUM27May21<<
                END;

                IF "Transport Details".Type = "Transport Details".Type::Purchase THEN BEGIN
                    RecPurchaseLine.RESET;
                    RecPurchaseLine.SETRANGE(RecPurchaseLine."Document No.", "Transport Details"."Invoice No.");
                    IF RecPurchaseLine.FINDFIRST THEN
                        ItemCatCode := RecPurchaseLine."Item Category Code";
                    IF RecItemCatgory.GET(ItemCatCode) THEN
                        categoryDesc := RecItemCatgory.Description;
                END;
                //RSPL-DHANANJAY
                invheader.RESET;
                invheader.SETRANGE(invheader."No.", "Transport Details"."Invoice No.");
                IF invheader.FINDFIRST THEN
                    RecCurRecDate := invheader."Customer Receipt Date";
                //MESSAGE('%1',RecCurRecDate);
                //RSPL-DHANANJAY
                //<<1
                IF "Transport Details"."Expected TPT Cost" <> 0 THEN
                    IF "Transport Details".Quantity <> 0 THEN //Fahim 29-04-23
                        RatePerKL := "Transport Details"."Expected TPT Cost" / "Transport Details".Quantity;
                //Fahim 05-03-2022
                CLEAR(InvoiceRcptbyFin);
                RecPurchaseHeader.RESET;
                RecPurchaseHeader.SETRANGE(RecPurchaseHeader."Vendor Invoice No.", "Transport Details"."TPT Invoice No.");
                // RecPurchaseHeader.SETRANGE(RecPurchaseHeader."Document Type",RecPurchaseHeader."Document Type"::Invoice); // Fahim it was from purchase header 38 now its from 122 pur. inv. header. 13-09-23
                IF RecPurchaseHeader.FINDFIRST THEN
                    InvoiceRcptbyFin := RecPurchaseHeader."Invoice Received By Finance";
                //MESSAGE('Test');
                //MESSAGE('%1',RecPurchaseHeader."No.");
                //End Fahim 05-03-2022


                //Fahim 07-03-2022
                /*
                //>>1
                //Vendor Ledger Entry, Body (2) - OnPreSection()
                //Row+=2;
                CLEAR(ExternalDocNo);
                CLEAR(Amount);
                //recVend.GET("Vendor No.");
                recVendorLedger.RESET;
                recVendorLedger.SETRANGE(recVendorLedger."Document No.","Transport Details"."Applied Document No.");
                IF recVendorLedger.FINDFIRST THEN
                  ExternalDocNo:=recVendorLedger."External Document No.";
                  Amount := recVendorLedger."Amount (LCY)";
                
                CLEAR(vCheckNo);
                CLEAR(tNarr);
                CLEAR(dCheckDate);
                //recVend.GET("Vendor No.");
                recCheckLedger.RESET;
                recCheckLedger.SETRANGE(recCheckLedger."Document No.",recVendorLedger."Document No.");
                IF recCheckLedger.FINDFIRST THEN
                BEGIN
                  vCheckNo:=recCheckLedger."Check No.";
                  dCheckDate:=recCheckLedger."Check Date";
                END;
                
                recPostedNarration.RESET;
                recPostedNarration.SETRANGE(recPostedNarration."Transaction No.",recVendorLedger."Transaction No.");
                recPostedNarration.SETRANGE(recPostedNarration."Entry No.",0);
                IF recPostedNarration.FINDFIRST THEN
                  tNarr:=recPostedNarration.Narration;
                  */
                //End Fahim 07-03-2022


                //>>20Apr2017 DataBody

                IF PrintToExcel THEN BEGIN

                    EnterCell(NNN, 1, "Transport Details"."From Location Name", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 2, Destinatn, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 3, "Transport Details"."Invoice No.", FALSE, FALSE, '', 1);
                    //EnterCell(NNN,4,invheader."Sell-to Customer Name",FALSE,FALSE,'',1);
                    EnterCell(NNN, 4, CustName14, FALSE, FALSE, '', 1);//RB-N 14Nov2017
                    EnterCell(NNN, 5, FORMAT("Transport Details"."Invoice Date"), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 6, FORMAT("Transport Details".Quantity), FALSE, FALSE, '0.000', 0);
                    EnterCell(NNN, 7, FORMAT("Transport Details"."Freight Type"), FALSE, FALSE, '', 1);
                    EnterCell(NNN, 8, "Transport Details"."Shipping Agent Name", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 9, "Transport Details"."Vehicle No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 10, "Transport Details"."LR No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 11, FORMAT(LrDate), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 12, "Transport Details"."TPT Invoice No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 13, FORMAT(TPTINvDate), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 14, FORMAT("Transport Details"."Expected TPT Cost"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 15, FORMAT("Transport Details"."Expetced Unloading"), FALSE, FALSE, '#,###0.00', 0);//Fahim 03-01-22
                    EnterCell(NNN, 16, FORMAT("Transport Details"."Bill Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 17, FORMAT("Transport Details"."Passed Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 18, "Transport Details"."Local Shipment Agent Name", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 19, "Transport Details"."Local Vehicle No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 20, "Transport Details"."Local LR No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 21, FORMAT(LocalLrDate), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 22, "Transport Details"."Local TPT Invoice No.", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 23, FORMAT("Transport Details"."Local TPT Invoice Date"), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 24, '', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 25, "Transport Details"."Shortcut Dimension 1 Code", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 26, "Transport Details".UOM, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 27, FORMAT("Transport Details".Type), FALSE, FALSE, '', 1);
                    EnterCell(NNN, 28, "Transport Details"."Vehicle Capacity", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 29, "Transport Details"."Local Vehicle Capacity", FALSE, FALSE, '', 1);
                    EnterCell(NNN, 30, FORMAT("Transport Details"."Other Charges"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 31, FORMAT("Transport Details"."Other Charges-Local"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 32, FORMAT("Transport Details"."Total Amount"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 33, FORMAT("Transport Details"."Total Amount-Local"), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NNN, 34, Status, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 35, cAppliedDocNo, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 36, FORMAT(dAppliedDate), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 37, '', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 38, '', FALSE, FALSE, '', 1);
                    EnterCell(NNN, 39, '', FALSE, FALSE, '', 1);
                    //EnterCell(NNN,40,'',FALSE,FALSE,'',1);//Fahim 01-08-2022
                    //EnterCell(NNN,40,FORMAT(RatePerKL),FALSE,FALSE,'',1);//Fahim 01-08-2022
                    EnterCell(NNN, 40, FORMAT(RatePerKL), FALSE, FALSE, '#,###0.00', 0);//Fahim 01-08-2022
                    EnterCell(NNN, 41, FORMAT("Transport Details"."Type of Vehicle"), FALSE, FALSE, '', 1);
                    EnterCell(NNN, 42, FORMAT("Transport Details"."Category of Tanker 1")
                   + ' ' + FORMAT("Transport Details"."Type of Truck"), FALSE, FALSE, '', 1);
                    EnterCell(NNN, 43, FORMAT("Transport Details"."Category of Tanker 2"), FALSE, FALSE, '', 1);
                    EnterCell(NNN, 44, FORMAT("Transport Details"."Category of Tanker 3"), FALSE, FALSE, '', 1);
                    EnterCell(NNN, 45, FORMAT("Transport Details"."Category of Tanker 4"), FALSE, FALSE, '', 1);
                    EnterCell(NNN, 46, vStatus, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 47, Reason, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 48, FORMAT(dPostedOn), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 49, categoryDesc, FALSE, FALSE, '', 1);
                    EnterCell(NNN, 50, FORMAT(RecCurRecDate), FALSE, FALSE, '', 2);
                    EnterCell(NNN, 51, FORMAT("Transport Details"."TPT Invoice Date"), FALSE, FALSE, '', 2); //RK-Add
                    EnterCell(NNN, 52, FORMAT(cdRegion), FALSE, FALSE, '', 1);//RSPLSUM27May21
                    EnterCell(NNN, 53, FORMAT("Transport Details"."Entry Date"), FALSE, FALSE, '', 2);//Fahim 05-03-2022
                    EnterCell(NNN, 54, FORMAT(InvoiceRcptbyFin), FALSE, FALSE, '', 2);//Fahim 05-03-2022

                    EnterCell(NNN, 55, ExternalDocNo, FALSE, FALSE, '', 1);//Fahim 07-03-2022
                    EnterCell(NNN, 56, vCheckNo, FALSE, FALSE, '', 1);//Fahim 07-03-2022
                    EnterCell(NNN, 57, FORMAT(dCheckDate), FALSE, FALSE, '', 2);//Fahim 07-03-2022
                    EnterCell(NNN, 58, FORMAT(Amount), FALSE, FALSE, '', 1);//Fahim 07-03-2022
                    EnterCell(NNN, 59, tNarr, FALSE, FALSE, '', 1);//Fahim 07-03-2022



                    NNN += 1;

                END;

                //<<20Apr2017 DataBody

            end;

            trigger OnPreDataItem()
            begin

                //>>1

                //CAS-05340-S2H4K5
                //RSPLSUM 06Dec19--IF vDetailed THEN
                //RSPLSUM 06Dec19--vStartDate:=010414D;
                SETRANGE("Transport Details"."Invoice Date", vStartDate, vEndDate);
                //CAS-05340-S2H4K5
                //<<1

                IF PrintToExcel THEN BEGIN

                    EnterCell(1, 1, Companyinfo.Name, TRUE, FALSE, '', 1);
                    EnterCell(1, 51, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

                    EnterCell(2, 1, 'Transport From ' + "Transport Details"."From Location Name", TRUE, FALSE, '', 1);
                    EnterCell(2, 51, USERID, TRUE, FALSE, '', 1);

                    EnterCell(5, 1, 'FROM LOCATION', TRUE, TRUE, '', 1);
                    EnterCell(5, 2, 'DESTINATION', TRUE, TRUE, '', 1);
                    EnterCell(5, 3, 'INVOICE NO', TRUE, TRUE, '', 1);
                    EnterCell(5, 4, 'Customer Name', TRUE, TRUE, '', 1);
                    EnterCell(5, 5, 'INVOICE DATE', TRUE, TRUE, '', 1);
                    EnterCell(5, 6, 'QTY(IN LTR)', TRUE, TRUE, '', 1);
                    EnterCell(5, 7, 'FREIGHT TYPE', TRUE, TRUE, '', 1);
                    EnterCell(5, 8, 'TRANSPORTER NAME', TRUE, TRUE, '', 1);
                    EnterCell(5, 9, 'VEHICLE No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 10, 'LR No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 11, 'LR DATE', TRUE, TRUE, '', 1);
                    EnterCell(5, 12, 'TPT BILL NO.', TRUE, TRUE, '', 1);
                    EnterCell(5, 13, 'TPT BILL DATE', TRUE, TRUE, '', 1);
                    EnterCell(5, 14, 'EXP. TPT COST', TRUE, TRUE, '', 1);
                    EnterCell(5, 15, 'Exp. Unloading Charges', TRUE, TRUE, '', 1);//Fahim 03-01-22
                    EnterCell(5, 16, 'BILL AMOUNT', TRUE, TRUE, '', 1);
                    EnterCell(5, 17, 'PASSED AMOUNT', TRUE, TRUE, '', 1);
                    EnterCell(5, 18, 'LOCAL TPT NAME', TRUE, TRUE, '', 1);
                    EnterCell(5, 19, 'LOCAL VEHICAL No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 20, 'LOCAL LR. No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 21, 'LOCAL LR. DATE', TRUE, TRUE, '', 1);
                    EnterCell(5, 22, 'LOCAL TPT INV No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 23, 'LOCAL TPT INV  DATE', TRUE, TRUE, '', 1);
                    EnterCell(5, 24, 'LOCAL FREIGHT', TRUE, TRUE, '', 1);
                    EnterCell(5, 25, 'DIVISION', TRUE, TRUE, '', 1);
                    EnterCell(5, 26, 'UOM', TRUE, TRUE, '', 1);
                    EnterCell(5, 27, 'TYPE', TRUE, TRUE, '', 1);
                    EnterCell(5, 28, 'Vehicle Capacity', TRUE, TRUE, '', 1);
                    EnterCell(5, 29, 'Local Vehicle Capacity', TRUE, TRUE, '', 1);
                    EnterCell(5, 30, 'Other Charges', TRUE, TRUE, '', 1);
                    EnterCell(5, 31, 'Other Charges-Local', TRUE, TRUE, '', 1);
                    EnterCell(5, 32, 'Total Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 33, 'Total Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 34, 'Status', TRUE, TRUE, '', 1);
                    EnterCell(5, 35, 'Transport Invoice No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 36, 'Transport Invoice Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 37, 'Average Detention Days', TRUE, TRUE, '', 1);
                    EnterCell(5, 38, 'Detention Claimed', TRUE, TRUE, '', 1);
                    EnterCell(5, 39, 'Detention Approved', TRUE, TRUE, '', 1);
                    EnterCell(5, 40, 'Rate per KL', TRUE, TRUE, '', 1);
                    EnterCell(5, 41, 'Type of Vehicle', TRUE, TRUE, '', 1);
                    EnterCell(5, 42, 'Category 1', TRUE, TRUE, '', 1);
                    EnterCell(5, 43, 'Category 2', TRUE, TRUE, '', 1);
                    EnterCell(5, 44, 'Category 3', TRUE, TRUE, '', 1);
                    EnterCell(5, 45, 'Category 4', TRUE, TRUE, '', 1);
                    EnterCell(5, 46, 'Status', TRUE, TRUE, '', 1);
                    EnterCell(5, 47, 'Reason', TRUE, TRUE, '', 1);
                    EnterCell(5, 48, 'Posted On', TRUE, TRUE, '', 1);
                    EnterCell(5, 49, 'Item Category Value', TRUE, TRUE, '', 1);
                    EnterCell(5, 50, 'Customer Receipt Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 51, 'Transport Invoice Receipt Date', TRUE, TRUE, '', 1); //RK-ADD
                    EnterCell(5, 52, 'Region', TRUE, TRUE, '', 1);//RSPLSUM27May21
                    EnterCell(5, 53, 'Entry Date', TRUE, TRUE, '', 1);//Fahim 05-03-2022
                    EnterCell(5, 54, 'Invoice Received by Finance', TRUE, TRUE, '', 1);//Fahim 05-03-2022

                    EnterCell(5, 55, 'External Doc No.', TRUE, TRUE, '', 1);//Fahim 07-03-2022
                    EnterCell(5, 56, 'Cheque No.', TRUE, TRUE, '', 1);//Fahim 07-03-2022
                    EnterCell(5, 57, 'Cheque Date', TRUE, TRUE, '', 1);//Fahim 07-03-2022
                    EnterCell(5, 58, 'Amount', TRUE, TRUE, '', 1);//Fahim 07-03-2022
                    EnterCell(5, 59, 'Narration', TRUE, TRUE, '', 1);//Fahim 07-03-2022


                    NNN := 6;

                END;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(Detailed; vDetailed)
                {

                    trigger OnValidate()
                    begin
                        /*
                        //CAS-05340-S2H4K5
                        IF vDetailed THEN BEGIN
                          RequestOptionsForm.StartDateVal.VISIBLE(FALSE);
                          RequestOptionsForm."Start Date".VISIBLE(FALSE);
                          RequestOptionsForm.EndDateVal.VISIBLE(FALSE);
                          RequestOptionsForm."End Date".VISIBLE(FALSE);
                          RequestOptionsForm.AsVal.VISIBLE(TRUE);
                          RequestOptionsForm.AsText.VISIBLE(TRUE);
                        
                        END
                        ELSE BEGIN
                          RequestOptionsForm.StartDateVal.VISIBLE(TRUE);
                          RequestOptionsForm."Start Date".VISIBLE(TRUE);
                          RequestOptionsForm.EndDateVal.VISIBLE(TRUE);
                          RequestOptionsForm."End Date".VISIBLE(TRUE);
                          RequestOptionsForm.AsVal.VISIBLE(FALSE);
                          RequestOptionsForm.AsText.VISIBLE(FALSE);
                        END;
                        //CAS-05340-S2H4K5
                        *///Commented NAV2009_Code


                        //>>20Apr2017 RequestPage
                        N1 := TRUE;
                        N2 := TRUE;

                        IF vDetailed THEN BEGIN

                            //RSPLSUM 06Dec19--N1 := FALSE;
                            //RSPLSUM 06Dec19--N2 := FALSE;
                            N1 := TRUE;//RSPLSUM 06Dec19
                            N2 := TRUE;//RSPLSUM 06Dec19
                            N3 := TRUE;
                            //RequestOptionsPage."Start Date";
                            //RequestOptionsPage."End Date".VISIBLE(FALSE);
                            //RequestOptionsPage."As on Date".VISIBLE(TRUE);

                        END ELSE BEGIN
                            N1 := TRUE;
                            N2 := TRUE;
                            N3 := FALSE;
                            //RequestOptionsPage."Start Date".VISIBLE(TRUE);
                            //RequestOptionsPage."End Date".VISIBLE(TRUE);
                            //RequestOptionsPage."As on Date";

                        END;
                        //<<20Apr2017 RequestPage

                    end;
                }
                field("Start Date"; vStartDate)
                {
                    ApplicationArea = all;
                    Enabled = N1;
                }
                field("End Date"; vEndDate)
                {
                    ApplicationArea = all;
                    Enabled = N2;
                }
                field("Status Date"; StatusDate)
                {
                    ApplicationArea = all;
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = all;
                }
            }
        }

        actions
        {
        }

        trigger OnOpenPage()
        begin
            //>>20Apr2017 RequestPage
            N1 := TRUE;
            N2 := TRUE;
            //<<
        end;
    }

    labels
    {
    }

    trigger OnPostReport()
    begin

        //>>1 20Apr2017
        IF PrintToExcel THEN BEGIN

            IF vDetailed THEN BEGIN
                // ExcelBuf.CreateBook('', 'Detailed-DATA');
                // ExcelBuf.CreateBookAndOpenExcel('', 'Detailed-DATA', '', '', USERID);
                // ExcelBuf.GiveUserControl;

                ExcelBuf.CreateNewBook('Detailed-DATA');
                ExcelBuf.WriteSheet('Detailed-DATA', CompanyName, UserId);
                ExcelBuf.CloseBook();
                ExcelBuf.SetFriendlyFilename(StrSubstNo('Detailed-DATA', CurrentDateTime, UserId));
                ExcelBuf.OpenExcel();
            END;

            IF NOT vDetailed THEN BEGIN
                // ExcelBuf.CreateBook('', 'Summery-DATA');
                // ExcelBuf.CreateBookAndOpenExcel('', 'Summery-DATA', '', '', USERID);
                // ExcelBuf.GiveUserControl;
                ExcelBuf.CreateNewBook('Summery-DATA');
                ExcelBuf.WriteSheet('Summery-DATA', CompanyName, UserId);
                ExcelBuf.CloseBook();
                ExcelBuf.SetFriendlyFilename(StrSubstNo('Summery-DATA', CurrentDateTime, UserId));
                ExcelBuf.OpenExcel();
            END;

            ExcelBuf.CreateNewBook('Summery-DATA');
            ExcelBuf.WriteSheet('Summery-DATA', CompanyName, UserId);
            ExcelBuf.CloseBook();
            ExcelBuf.SetFriendlyFilename(StrSubstNo('Detailed-DATA', CurrentDateTime, UserId));
            ExcelBuf.OpenExcel();


            /*
            ExSheet.Columns.AutoFit;
            ExApp.Visible := TRUE;
            *///Commented Automation20Apr2017
        END;

        //<<1

    end;

    trigger OnPreReport()
    begin
        //>>1

        //EBT MILAN ---(31072013)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        LocCode := "Transport Details".GETFILTER("Transport Details"."From Location Code");
        //user.GET(USERID);

        /*
        memberof.RESET;
        memberof.SETRANGE(memberof."User ID",user."User ID");
        memberof.SETFILTER(memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.RESET;
         recLocation.SETRANGE(recLocation.Code,LocCode);
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOmapping1.RESET;
         CSOmapping1.SETRANGE(CSOmapping1."User Id",UPPERCASE(USERID));
         CSOmapping1.SETRANGE(Type,CSOmapping1.Type::Location);
         CSOmapping1.SETRANGE(CSOmapping1.Value,LocCode);
         IF NOT(CSOmapping1.FINDFIRST) THEN
         BEGIN
          CSOmapping1.RESET;
          CSOmapping1.SETRANGE(CSOmapping1."User Id",UPPERCASE(USERID));
          CSOmapping1.SETRANGE(CSOmapping1.Type,CSOmapping1.Type::"Responsibility Center");
          CSOmapping1.SETRANGE(CSOmapping1.Value,LocResC);
          IF NOT(CSOmapping1.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        //EBT MILAN ---(31072013)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        *///Commented Automation20Apr2017

        Companyinfo.GET;

        IF PrintToExcel THEN BEGIN

            /*
            CREATE(ExApp);
            ExBook := ExApp.Workbooks.Add;
            ExSheet := ExBook.Worksheets.Add;
            //CAS-05340-S2H4K5
            IF vDetailed THEN
             ExSheet.Name := 'Detailed-DATA'
            ELSE
             ExSheet.Name := 'Summery-DATA';
            //CAS-05340-S2H4K5
            ExApp.Visible := FALSE;
            *///Commented Automation20Apr2017

        END;
        //<<1

    end;

    var
        Companyinfo: Record 79;
        PrintToExcel: Boolean;
        i: Integer;
        TPTINvDate: Text[30];
        LocalLrDate: Text[30];
        LrDate: Text[30];
        user: Record 2000000120;
        recLocation: Record 14;
        CSOmapping1: Record 50006;
        LocCode: Code[10];
        LocResC: Code[10];
        Destinatn: Text[50];
        invheader: Record 112;
        Status: Text[50];
        vStartDate: Date;
        vEndDate: Date;
        cAppliedDocNo: Code[20];
        dAppliedDate: Date;
        vStatus: Text[30];
        vDetailed: Boolean;
        dPostedOn: Date;
        "--------------------": Integer;
        RecSalesInvLine: Record 113;
        RecTransferLine: Record 5745;
        ItemCatCode: Code[10];
        RecPurchaseLine: Record 39;
        RecItemCatgory: Record 5722;
        categoryDesc: Text[50];
        RecCurRecDate: Date;
        "---20Apr2017": Integer;
        ExcelBuf: Record 370 temporary;
        NNN: Integer;
        //[InDataSet]
        N1: Boolean;
        //[InDataSet]
        N2: Boolean;
        //[InDataSet]
        N3: Boolean;
        "-----------14Nov2017": Integer;
        CustName14: Text[100];
        StatusDate: Date;
        RecSalesInvHeader: Record 112;
        RecShipToAddr: Record 222;
        cdRegion: Code[10];
        RecCust: Record 18;
        InvoiceRcptbyFin: Date;
        RecPurchaseHeader: Record 122;
        vCheckNo: Code[20];
        dCheckDate: Date;
        recVend: Record 23;
        recCheckLedger: Record 272;
        recPostedNarration: Record "Posted Narration";
        tNarr: Text[250];
        ExternalDocNo: Text[20];
        recVendorLedger: Record 25;
        Amount: Decimal;
        RatePerKL: Decimal;

    ////[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        IF PrintToExcel THEN BEGIN
            ExcelBuf.INIT;
            ExcelBuf.VALIDATE("Row No.", Rowno);
            ExcelBuf.VALIDATE("Column No.", columnno);
            ExcelBuf."Cell Value as Text" := Cellvalue;
            ExcelBuf.Formula := '';
            ExcelBuf.Bold := Bold;
            ExcelBuf.Underline := Underline;
            ExcelBuf.NumberFormat := NoFormat;
            ExcelBuf."Cell Type" := CType;
            ExcelBuf.INSERT;
        END;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


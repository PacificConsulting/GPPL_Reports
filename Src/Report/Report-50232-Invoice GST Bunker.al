report 50232 "Invoice GST Bunker"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 22Aug2017   RB-N         New GST tax Invoice Report for GP Energy
    // 08Nov2017   RB-N         Sales Order Date
    // 19Dec2017   RB-N         GST Compensation CESS
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/InvoiceGSTBunker.rdl';
    PreviewMode = PrintLayout;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "Posting Date", "No.";
            column(EWayBillNo; EWayBillNo)
            {
            }
            column(EWayBillDate; GeneratedDate)
            {
            }
            column(Text1; Text1)
            {
            }
            column(Text2; Text2)
            {
            }
            column(Text3; Text3)
            {
            }
            column(Text4; Text4)
            {
            }
            column(PortDescription; RecPSAddInfo."Port Description")
            {
            }
            column(ShippingBillNo; RecPSAddInfo."Shipping Bill No")
            {
            }
            column(PlaceofSupply; "Sales Invoice Header"."Ship-to City")
            {
            }
            column(DocNo; "Sales Invoice Header"."No.")
            {
            }
            column(OrderDate; FORMAT(OrderDate08, 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(PostingDate; FORMAT("Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(DocumentDate; FORMAT("Document Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(DueDate; FORMAT("Due Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(BDNDate; FORMAT(RecPSAddInfo."BDN Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(AdvanceReceiptNo; RecPSAddInfo."Advance Receipt No")
            {
            }
            column(BuyersOrderNo; RecPSAddInfo."Buyer's Order No")
            {
            }
            column(BuyersOrderDate; FORMAT(RecPSAddInfo."Buyer's Order Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(ExDocNo; "Sales Invoice Header"."External Document No.")
            {
            }
            column(VesselCode; RecPSAddInfo."Vessel Code")
            {
            }
            column(VesselName; RecPSAddInfo."Vessel Name")
            {
            }
            column(BDNNo; RecPSAddInfo."BDN No.")
            {
            }
            column(SalesOrderNo; "Sales Invoice Header"."EWB Transaction Type")
            {
            }
            column(TermsOfDelivery; RecPSAddInfo."Terms Of Delivery")
            {
            }
            column(BLNo_SalesInvoiceHeader; RecPSAddInfo."B/L No.")
            {
            }
            column(BLDate_SalesInvoiceHeader; RecPSAddInfo."B/L Date")
            {
            }
            column(LRRRNo; "Sales Invoice Header"."LR/RR No.")
            {
            }
            column(LRRRDate; FORMAT("LR/RR Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(VehicleNo; "Sales Invoice Header"."Vehicle No.")
            {
            }
            column(BillOfExportNo; "Sales Invoice Header"."Bill Of Export No.")
            {
            }
            column(BillOfExportDate; FORMAT("Bill Of Export Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(VehicleCapacity; "Sales Invoice Header"."Vehicle Capacity")
            {
            }
            column(DriversName_SalesInvoiceHeader; '')
            {
            }
            column(BankAccountNo_SalesInvoiceHeader; "Sales Invoice Header"."Bank Account No.")
            {
            }
            column(ShippingBillNo_SalesInvoiceHeader; RecPSAddInfo."Shipping Bill No")
            {
            }
            column(ShippingDate_SalesInvoiceHeader; RecPSAddInfo."Shipping Bill  Date")
            {
            }
            column(BilltoCustomerNo_SalesInvoiceHeader; "Sales Invoice Header"."Bill-to Customer No.")
            {
            }
            column(BilltoName; "Sales Invoice Header"."Bill-to Name")
            {
            }
            column(BillAdd1; "Bill-to Address")
            {
            }
            column(BillAdd2; "Sales Invoice Header"."Bill-to Address 2")
            {
            }
            column(BillAdd3; "Bill-to City" + ' - ' + "Bill-to Post Code")
            {
            }
            column(ShiptoName; "Sales Invoice Header"."Ship-to Name")
            {
            }
            column(ShipAdd1; "Sales Invoice Header"."Ship-to Address")
            {
            }
            column(ShipAdd2; "Sales Invoice Header"."Ship-to Address 2")
            {
            }
            column(ShipAdd3; "Ship-to City" + ' - ' + "Ship-to Post Code")
            {
            }
            column(LocAdd1; Loc22.Address)
            {
            }
            column(LocAdd2; Loc22."Address 2")
            {
            }
            column(LocAdd3; Loc22.City + ' - ' + Loc22."Post Code")
            {
            }
            column(LocGSTRegNo; Loc22."GST Registration No.")
            {
            }
            column(CompPANNo; CompInfo."P.A.N. No.")
            {
            }
            column(CompName; CompInfo.Name)
            {
            }
            column(CompGSTRegNo; CompInfo."GST Registration No.")
            {
            }
            column(CompCINNo; CompInfo."Registration No.")
            {
            }
            column(CompPicture; CompInfo.Picture)
            {
            }
            column(InsurNo; CompInfo."Insurance No.")
            {
            }
            column(InsurProv; CompInfo."Insurance Provider")
            {
            }
            column(CompAdd1; CompInfo.Address)
            {
            }
            column(CompAdd2; CompInfo."Address 2")
            {
            }
            column(CompAdd3; CompInfo.City + ' - ' + CompInfo."Post Code")
            {
            }
            column(CompPhoneNo; CompInfo."Phone No.")
            {
            }
            column(InsurExpDate; FORMAT(CompInfo."Policy Expiration Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(ShipGSTStateCode; ShipGSTStateCode)
            {
            }
            column(ShipState; ShipState)
            {
            }
            column(ShipGSTNo; ShipGSTNo)
            {
            }
            column(BillGSTStateCode; BillGSTStateCode)
            {
            }
            column(BillState; BillState)
            {
            }
            column(BillGSTNo; BillGSTNo)
            {
            }
            column(ShipPANNo; ShipPANNo)
            {
            }
            column(BillPANNo; BillPANNo)
            {
            }
            column(LocGSTStateCode; LocGSTStateCode)
            {
            }
            column(LocState; LocState)
            {
            }
            column(TotalCGST; TotalCGST)
            {
            }
            column(TotalSGST; TotalSGST)
            {
            }
            column(TotalIGST; TotalIGST)
            {
            }
            column(TotalCess; TotalCess)
            {
            }
            column(CGSTPer; CGSTPer)
            {
            }
            column(SGSTPer; SGSTPer)
            {
            }
            column(IGSTPer; IGSTPer)
            {
            }
            column(FinalAmtinWord; AmtinWord15[1])
            {
            }
            column(CurCode; CurCode)
            {
            }
            column(GSTinWord; CGSTinWord15[1])
            {
            }
            column(BankAccName; BankAcc."Full Name")
            {
            }
            column(BankAccNo; BankAcc."Bank Account No.")
            {
            }
            column(BankSwiftCode; BankAcc."SWIFT Code")
            {
            }
            column(BankCity; BankAcc.City)
            {
            }
            column(CorrespondingBankSwiftCode; BankAcc."Swift Code Number")
            {
            }
            column(CoorespondingBankName; BankAcc."Corresponding Bank")
            {
            }
            column(CoorespondingBankAccNo; BankAcc."Corresponding Bank Account No.")
            {
            }
            column(BankAdd1; BankAcc.Address)
            {
            }
            column(BankAdd2; BankAcc."Address 2" + ', ' + BankAcc.City + ', ' + BankAcc."Post Code")
            {
            }
            column(RoundOffAmnt; RoundOffAmnt)
            {
            }
            column(Paydesc; PayTerm.Code)
            {
            }
            column(CT1; CommText[1])
            {
            }
            column(CT2; CommText[2])
            {
            }
            column(CT3; CommText[3])
            {
            }
            column(CT4; CommText[4])
            {
            }
            column(CT5; CommText[5])
            {
            }
            column(CT6; CommText[6])
            {
            }
            column(TransporterName; ShipAgent.Name)
            {
            }
            column(CurrencyCode_SalesInvoiceHeader; "Sales Invoice Header"."Currency Code")
            {
            }
            column(QRCode; "Sales Invoice Header"."QR code")
            {
            }
            column(IRN_SalesInvoiceHeader; IRN)
            {
            }
            column(TCSPercent; TCSPercent)
            {
            }
            column(TCSAmount; TCSAmount)
            {
            }
            dataitem(CopyLoop; Integer)
            {
                DataItemTableView = SORTING(Number);
                dataitem(DataItem1000000035; Integer)
                {
                    DataItemTableView = SORTING(Number)
                                        WHERE(Number = FILTER(1));
                    column(InvType; InvType)
                    {
                    }
                    column(number; Number)
                    {
                    }
                    column(outPutNos; OutPutNo)
                    {
                    }
                    dataitem("Sales Invoice Line"; "Sales Invoice Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Sales Invoice Header";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            WHERE(Type = FILTER(Item | "G/L Account"),
                                                  "No." = FILTER(<> 74012210),
                                                  Quantity = FILTER(<> 0));
                        column(NAH; NAH)
                        {
                        }
                        column(SNo; SNo)
                        {
                        }
                        column(LineDocNo; "Sales Invoice Line"."Document No.")
                        {
                        }
                        column(LineNo; "Sales Invoice Line"."Line No.")
                        {
                        }
                        column(Type_SalesInvoiceLine; "Sales Invoice Line".Type)
                        {
                        }
                        column(ItmNo; "Sales Invoice Line"."No.")
                        {
                        }
                        column(ItmDesc; ItmDesc)
                        {
                        }
                        column(UnitofMeasure; "Sales Invoice Line"."Unit of Measure")
                        {
                        }
                        column(HSNCode; "Sales Invoice Line"."HSN/SAC Code")
                        {
                        }
                        column(GSTBaseAmt; GSTBaseAmt)
                        {
                        }
                        column(GSTPer; 0)//"Sales Invoice Line"."GST %")
                        {
                        }
                        column(TotalGSTAmt; 0)//"Sales Invoice Line"."Total GST Amount")
                        {
                        }
                        column(LineAmt; "Sales Invoice Line"."Line Amount")
                        {
                        }
                        column(AmountToCustomer; 0)//"Sales Invoice Line"."Amount To Customer")
                        {
                        }
                        column(QtyBase; "Sales Invoice Line"."Quantity (Base)")
                        {
                        }
                        column(Qty; "Sales Invoice Line".Quantity)
                        {
                        }
                        column(UnitPrice; "Sales Invoice Line"."Unit Price")
                        {
                        }
                        column(UoMCode; "Sales Invoice Line"."Unit of Measure Code")
                        {
                        }
                        column(TDSTCSAmount; 0)// "Sales Invoice Line"."TDS/TCS Amount")
                        {
                        }
                        column(TDSTCSPer; 0)// "Sales Invoice Line"."TDS/TCS %")
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin

                            NAH := COUNT;
                            SNo += 1;

                            //LineAmount
                            CLEAR(GSTBaseAmt);
                            //IF "GST Base Amount" = 0 THEN
                            GSTBaseAmt := "Sales Invoice Line"."Line Amount";
                            // ELSE
                            GSTBaseAmt := 0;// "Sales Invoice Line"."GST Base Amount";//RSPLSUM
                                            //GSTBaseAmt := "Sales Invoice Line"."GST Base Amount" - "Sales Invoice Line"."TDS/TCS Amount";

                            CLEAR(ItmDesc);
                            Itm.RESET;
                            IF "Sales Invoice Line".Type = "Sales Invoice Line".Type::Item THEN
                                IF Itm.GET("No.") THEN
                                    ItmDesc := Itm.Description;

                            IF "Sales Invoice Line".Type <> "Sales Invoice Line".Type::Item THEN
                                ItmDesc := Description;

                            IF "Item Category Code" = 'CAT21' THEN//RSPLSUM 19Jan21
                                ItmDesc := Description;//RSPLSUM 19Jan21
                        end;

                        trigger OnPreDataItem()
                        begin

                            SNo := 0;
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin



                        i += 1;
                        IF i = 1 THEN
                            InvType := 'ORIGINAL FOR RECIPIENT'
                        ELSE
                            IF i = 2 THEN
                                InvType := 'DUPLICATE FOR TRANSPORTER'
                            ELSE
                                IF i = 3 THEN
                                    InvType := 'TRIPLICATE FOR SUPPLIER'
                                ELSE
                                    IF i = 4 THEN
                                        InvType := 'EXTRA COPY'
                                    ELSE
                                        InvType := '';

                        /*
                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'GPUAE\UNNIKRISHNAN.VS') THEN //18July2017
                        BEGIN
                          IF i = 1 THEN
                           InvType := 'ORIGINAL FOR RECIPIENT'
                           //InvType := 'ORIGINAL FOR BUYER'
                           //InvType := 'ORIGINAL FOR BUYER'
                          ELSE IF i = 2 THEN
                           InvType := 'DUPLICATE FOR TRANSPORTER'
                          ELSE IF i = 3 THEN
                            InvType := 'TRIPLICATE FOR SUPPLIER'
                            //InvType := 'TRIPLICATE FOR ASSESSEE'
                          ELSE IF i = 4 THEN BEGIN
                           InvType := 'QUADRAPLICATE FOR H.O';
                          END ELSE IF i=5 THEN BEGIN
                           InvType := 'EXTRA COPY NOT FOR CENVAT';
                           InvType1 := '';
                        END
                        ELSE
                        InvType:='' ;
                        END;
                        
                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'GPUAE\UNNIKRISHNAN.VS') THEN //18July2017
                        BEGIN
                        IF i = 1 THEN
                           InvType := 'EXTRA COPY';
                           InvType1 := ' NOT FOR COMMERCIAL USE';
                        END;
                        
                         */

                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    OutPutNo += 1;
                    IF Number > 1 THEN;
                    //copytext := Text003;
                    //CurrReport.PAGENO := 1;
                end;

                trigger OnPreDataItem()
                begin


                    NoOfCopies := 4;
                    NoOfLoops := ABS(NoOfCopies);
                    IF NoOfLoops <= 0 THEN
                        NoOfLoops := 1;

                    //copytext := '';
                    SETRANGE(Number, 1, NoOfLoops);
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //Location Details
                Loc22.RESET;
                IF Loc22.GET("Location Code") THEN;

                RecState.RESET;
                IF RecState.GET(Loc22."State Code") THEN BEGIN
                    LocGSTStateCode := RecState."State Code (GST Reg. No.)";
                    LocState := RecState.Description;

                END;

                //RSPLSUM-TCS>>
                CLEAR(TCSAmount);
                SL14.RESET;
                SL14.SETCURRENTKEY("Document No.", Type, Quantity);
                SL14.SETRANGE("Document No.", "No.");
                SL14.SETFILTER(Quantity, '<>%1', 0);
                IF SL14.FINDSET THEN
                    REPEAT
                        TCSAmount += 0;// SL14."TDS/TCS Amount";
                    UNTIL SL14.NEXT = 0;
                //RSPLSUM-TCS<<

                //RoundOff Amount
                CLEAR(RoundOffAmnt);
                SL14.RESET;
                SL14.SETRANGE("Document No.", "No.");
                SL14.SETRANGE("No.", '74012210');
                IF SL14.FINDFIRST THEN
                    RoundOffAmnt := 0;// SL14."Amount To Customer";

                CLEAR(FinalAmt);
                SL14.RESET;
                SL14.SETRANGE("Document No.", "No.");
                //SL14.SETFILTER("No.",'<>%1','45012013');
                IF SL14.FINDSET THEN
                    REPEAT
                        FinalAmt += 0;//SL14."Amount To Customer";
                    UNTIL SL14.NEXT = 0;


                //>>AmountinWords
                RCheck.InitTextVariable;
                RCheck.FormatNoText(AmtinWord15, ROUND((FinalAmt), 0.01), "Currency Code");
                AmtinWord15[1] := DELCHR(AmtinWord15[1], '=', '*');

                //<<AmountinWords

                //>>23Aug2017TotalGST Amount
                CLEAR(TotalCGST);
                CLEAR(TotalSGST);
                CLEAR(TotalIGST);
                CLEAR(TotalCess);
                DGST04.RESET;
                DGST04.SETRANGE("Document No.", "No.");
                IF DGST04.FINDSET THEN
                    REPEAT
                        IF DGST04."GST Component Code" = 'CGST' THEN BEGIN
                            //TotalCGST += ABS(DGST04."GST Amount") ;
                            IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                TotalCGST += ABS(DGST04."GST Amount" * "Sales Invoice Header"."Currency Factor")
                            ELSE
                                TotalCGST += ABS(DGST04."GST Amount");
                            CGSTPer := DGST04."GST %";
                        END;

                        IF DGST04."GST Component Code" = 'SGST' THEN BEGIN
                            //TotalSGST += ABS(DGST04."GST Amount") ;
                            SGSTPer := DGST04."GST %";
                            IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                TotalSGST += ABS(DGST04."GST Amount" * "Sales Invoice Header"."Currency Factor")
                            ELSE
                                TotalSGST += ABS(DGST04."GST Amount");
                        END;

                        IF DGST04."GST Component Code" = 'UTGST' THEN BEGIN
                            //TotalSGST += ABS(DGST04."GST Amount") ;
                            SGSTPer := DGST04."GST %";
                            IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                TotalSGST += ABS(DGST04."GST Amount" * "Sales Invoice Header"."Currency Factor")
                            ELSE
                                TotalSGST += ABS(DGST04."GST Amount");

                        END;

                        IF DGST04."GST Component Code" = 'IGST' THEN BEGIN
                            //TotalIGST += ABS(DGST04."GST Amount") ;
                            IGSTPer := DGST04."GST %";
                            IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                TotalIGST += ABS(DGST04."GST Amount" * "Sales Invoice Header"."Currency Factor")
                            ELSE
                                TotalIGST += ABS(DGST04."GST Amount");
                        END;

                    /*
               //RK
                  IF DGST04."GST Component Code" ='IGST' THEN
                    BEGIN
                       IF "Sales Invoice Header"."Currency Factor" = 0 THEN
                       BEGIN
                        TotalIGST := ABS(DGST04."GST Amount");
                       END;

                       IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                       BEGIN
                        TotalIGST += ABS(DGST04."GST Amount" * "Sales Invoice Header"."Currency Factor" );
                       END;
                        IGSTPer := DGST04."GST %";
                    END;
               //RK
                    */

                    /*
                        IF DGST04."GST Component Code" = 'CESS' THEN
                           BEGIN
                              TotalCess += ABS(DGST04."GST Amount") ;
                              CessPer   := DGST04."GST %";
                           END;
                    */

                    UNTIL DGST04.NEXT = 0;

                //>>GST AmountinWords
                RCheck.InitTextVariable;
                //vCheck.FormatNoText(CGSTinWord15,ROUND((TotalCGST),1.0),'');
                RCheck.FormatNoText(CGSTinWord15, ROUND((TotalCGST + TotalSGST + TotalIGST), 0.01), "Currency Code");
                CGSTinWord15[1] := DELCHR(CGSTinWord15[1], '=', '*');

                //>>SGST AmountinWords
                RCheck.InitTextVariable;
                RCheck.FormatNoText(SGSTinWord15, ROUND((TotalSGST), 0.01), "Currency Code");
                SGSTinWord15[1] := DELCHR(SGSTinWord15[1], '=', '*');

                //>>IGST AmountinWords
                RCheck.InitTextVariable;
                RCheck.FormatNoText(IGSTinWord15, ROUND((TotalIGST), 0.01), "Currency Code");
                IGSTinWord15[1] := DELCHR(IGSTinWord15[1], '=', '*');

                //<<23Aug2017TotalGST Amount

                //Bill State Description
                RecState.RESET;
                // IF RecState.GET("Bill to Customer State") THEN BEGIN
                BillState := RecState.Description;
                BillGSTStateCode := RecState."State Code (GST Reg. No.)";
                //  END;


                //Bill GSTNo
                Cus12.RESET;
                IF Cus12.GET("Sales Invoice Header"."Bill-to Customer No.") THEN BEGIN
                    BillGSTNo := Cus12."GST Registration No.";
                    BillPANNo := Cus12."P.A.N. No.";
                END;

                //Bill Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Bill-to Country/Region Code") THEN
                    BillCountry := RecCountry.Name;



                //Ship Country Name
                RecCountry.RESET;
                IF RecCountry.GET("Ship-to Country/Region Code") THEN
                    ShipCountry := RecCountry.Name;



                //Shipping
                IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code") THEN BEGIN

                    RecState.RESET;
                    IF RecState.GET(ShipToCode.State) THEN BEGIN
                        ShipState := RecState.Description;
                        ShipGSTStateCode := RecState."State Code (GST Reg. No.)";
                        ShipGSTNo := ShipToCode."GST Registration No.";
                    END;
                    ShipPANNo := ShipToCode."P.A.N. No.";

                END ELSE BEGIN

                    ShipState := BillState;
                    ShipGSTStateCode := BillGSTStateCode;
                    ShipGSTNo := BillGSTNo;
                    ShipPANNo := BillPANNo;

                END;


                //>>23Aug2017

                //Currency Code
                IF "Currency Code" = '' THEN
                    CurCode := 'INR'
                ELSE
                    CurCode := "Currency Code";


                //BankAccount
                BankAcc.RESET;
                BankAcc.SETRANGE(BankAcc."No.", "Sales Invoice Header"."Bank Account No.");
                IF BankAcc.FINDFIRST THEN;

                //Payment Terms
                PayTerm.RESET;
                IF PayTerm.GET("Sales Invoice Header"."Payment Terms Code") THEN;

                //Comment
                CLEAR(CT);
                CLEAR(CommText);
                SalesComment.RESET;
                SalesComment.SETRANGE(SalesComment."Document Type", SalesComment."Document Type"::"Posted Invoice");
                SalesComment.SETRANGE("No.", "No.");
                IF SalesComment.FINDSET THEN
                    REPEAT
                        CT += 1;
                        CommText[CT] := SalesComment.Comment;

                    UNTIL SalesComment.NEXT = 0;

                //Shipping Agent
                ShipAgent.RESET;
                IF ShipAgent.GET("Sales Invoice Header"."Shipping Agent Code") THEN;

                //RB-N 08Nov2017 Sales Order Date

                IF "Sales Invoice Header"."Sales Order No" <> '' THEN BEGIN

                    SH08.RESET;
                    SH08.SETRANGE("Document Type", SH08."Document Type"::Order);
                    SH08.SETRANGE("No.", "Sales Order No");
                    IF SH08.FINDFIRST THEN BEGIN
                        OrderDate08 := SH08."Order Date";
                    END;

                END ELSE
                    OrderDate08 := "Order Date";
                //RB-N 08Nov2017 Sales Order Date


                //RB-N 08Nov2017 BankAccount Details
                Text1 := '';
                Text2 := '';
                Text3 := '';
                Text4 := '';

                //RSPLSUM 29May2020>>
                IF ("Sales Invoice Header"."Currency Code" = 'INR') OR ("Sales Invoice Header"."Currency Code" = '') THEN BEGIN
                    Text1 := 'IFSC Code';
                    Text3 := BankAcc."IFSC Code";
                END ELSE BEGIN
                    Text1 := 'Swift Code';
                    Text3 := BankAcc."SWIFT Code";
                END;
                //RSPLSUM 29May2020<<

                /*//RSPLSUM 29May2020>>
                IF "Gen. Bus. Posting Group" = 'FOREIGN' THEN
                BEGIN
                
                  Text1 := 'Corresponding Bank';
                  Text2 := 'Swift Code';
                  Text3 := BankAcc."Corresponding Bank";
                  Text4 := BankAcc."SWIFT Code";
                
                END ELSE
                BEGIN
                
                  Text1 := 'IFSC Code';
                  Text2 := 'Account Type';
                  Text3 := BankAcc."IFSC Code";
                  Text4 := BankAcc."Account Type";
                
                END;
                *///RSPLSUM 29May2020<<

                //>>RB-N 19Dec2017  GST CompCess
                /*  PSTL19.RESET;
                 PSTL19.SETRANGE(Type, PSTL19.Type::Sale);
                 PSTL19.SETRANGE("Document Type", PSTL19."Document Type"::Invoice);
                 PSTL19.SETRANGE("Invoice No.", "No.");
                 PSTL19.SETRANGE("Tax/Charge Group", 'CESS');
                 PSTL19.SETRANGE("Tax/Charge Type", PSTL19."Tax/Charge Type"::"Other Taxes");
                 IF PSTL19.FINDSET THEN
                     REPEAT
                         TotalCess += PSTL19.Amount;
                     UNTIL PSTL19.NEXT = 0; */

                //>>RB-N 19Dec2017  GST CompCess

                //RSPLSUM 28May2020>>
                CLEAR(EWayBillNo);
                CLEAR(GeneratedDate);
                GSTLedgerEntry.RESET;
                GSTLedgerEntry.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                GSTLedgerEntry.SETRANGE("Posting Date", "Sales Invoice Header"."Posting Date");
                IF GSTLedgerEntry.FINDFIRST THEN BEGIN
                    EWayBillNo := GSTLedgerEntry."E-Way Bill No.";

                    DetailedsEWayBill1.RESET;
                    DetailedsEWayBill1.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                    DetailedsEWayBill1.SETRANGE("EWB No.", EWayBillNo);
                    IF DetailedsEWayBill1.FINDFIRST THEN
                        GeneratedDate := DetailedsEWayBill1."EWB Updated Date";
                END;
                //RSPLSUM 28May2020<<

                RecPSAddInfo.RESET;
                IF RecPSAddInfo.GET("Sales Invoice Header"."No.") THEN
                    RecPSAddInfo.CALCFIELDS("Port Description", "Vessel Name");

                //RSPLSUM BEGIN>>
                CLEAR(IRN);
                "Sales Invoice Header".CALCFIELDS("Sales Invoice Header".IRN);
                IF "Sales Invoice Header".IRN <> '' THEN
                    IRN := 'IRN: ' + "Sales Invoice Header".IRN;
                //RSPLSUM END>>

                //RSPLSUM-TCS>>
                CLEAR(TCSPercent);
                RecSIL.RESET;
                RecSIL.SETRANGE("Document No.", "No.");
                RecSIL.SETRANGE(Type, RecSIL.Type::Item);
                RecSIL.SETFILTER(Quantity, '<>%1', 0);
                IF RecSIL.FINDFIRST THEN BEGIN
                    // IF RecSIL."TDS/TCS %" <> 0 THEN
                    TCSPercent := '';// FORMAT(RecSIL."TDS/TCS %") + ' %';
                END;
                //RSPLSUM-TCS<<

                "Sales Invoice Header".CALCFIELDS("Sales Invoice Header"."QR code");//ROBOEinv
                //BarcodeForQuarantineLabel;//RSPLSUM // DJ Commented 06/10/2020

            end;
        }
    }

    requestpage
    {

        layout
        {
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPreReport()
    begin

        //Company Information
        CompInfo.GET;
        CompInfo.CALCFIELDS(Picture);
    end;

    var
        Loc22: Record 14;
        CompInfo: Record 79;
        NoOfCopies: Integer;
        NoOfLoops: Integer;
        OutPutNo: Integer;
        i: Integer;
        InvType: Text[50];
        FinalAmt: Decimal;
        SL14: Record 113;
        AmtinWord15: array[2] of Text[150];
        RCheck: Report "Check Report";
        RoundOffAmnt: Decimal;
        CGST15: Text[30];
        SGST15: Text[30];
        IGST15: Text[30];
        DGST04: Record "Detailed GST Ledger Entry";
        TotalCGST: Decimal;
        TotalSGST: Decimal;
        TotalIGST: Decimal;
        TotalCess: Decimal;
        CGSTPer: Decimal;
        SGSTPer: Decimal;
        IGSTPer: Decimal;
        CGSTinWord15: array[2] of Text[150];
        SGSTinWord15: array[2] of Text[150];
        IGSTinWord15: array[2] of Text[150];
        SNo: Integer;
        RecState: Record State;
        ShipAdd1: Text[60];
        ShipAdd2: Text[60];
        ShipAdd3: Text[60];
        ShipGSTStateCode: Code[10];
        ShipGSTNo: Code[15];
        ShiptoName15: Text[100];
        BillState: Text[50];
        BillGSTStateCode: Code[10];
        BillGSTNo: Code[15];
        BillCountry: Text[50];
        ShipCountry: Text[50];
        Ship2Name: Text[60];
        ShipState: Text[50];
        Cus12: Record 18;
        RecCountry: Record 9;
        ShipToCode: Record 222;
        ShipPANNo: Code[20];
        BillPANNo: Code[20];
        LocPANNo: Code[20];
        LocGSTStateCode: Code[10];
        LocState: Text[50];
        NAH: Integer;
        CurCode: Code[10];
        BankAcc: Record 270;
        GSTBaseAmt: Decimal;
        PayTerm: Record 3;
        CessPer: Decimal;
        SalesComment: Record 44;
        CommText: array[10] of Text[150];
        CT: Integer;
        ShipAgent: Record 291;
        "--------08Nov2017": Integer;
        OrderDate08: Date;
        SH08: Record 36;
        Text1: Text;
        Text2: Text;
        Text3: Text;
        Text4: Text;
        "-----------19Dec2017": Integer;
        //  PSTL19: Record 13798;
        ItmDesc: Text;
        Itm: Record 27;
        RecPSAddInfo: Record 50054;
        EWayBillNo: Code[20];
        GSTLedgerEntry: Record "GST Ledger Entry";
        DetailedsEWayBill1: Record 50044;
        GeneratedDate: Text;
        RecGSTLedgEntry: Record "GST Ledger Entry";
        IRN: Text[255];
        TCSPercent: Text[10];
        RecSIL: Record 113;
        TCSAmount: Decimal;

    // //[Scope('Internal')]
    procedure BarcodeForQuarantineLabel()
    var
        Text5000_Ctx: Label '<xpml><page quantity=''0'' pitch=''74.1 mm''></xpml>SIZE 100 mm, 74.1 mm';
        Text5001_Ctx: Label 'DIRECTION 0,0';
        Text5002_Ctx: Label 'REFERENCE 0,0';
        Text5003_Ctx: Label 'OFFSET 0 mm';
        Text5004_Ctx: Label 'SET PEEL OFF';
        Text5005_Ctx: Label 'SET CUTTER OFF';
        Text5006_Ctx: Label 'SET PARTIAL_CUTTER OFF';
        Text5007_Ctx: Label '<xpml></page></xpml><xpml><page quantity=''1'' pitch=''74.1 mm''></xpml>SET TEAR ON';
        Text5008_Ctx: Label 'CLS';
        Text5009_Ctx: Label 'CODEPAGE 1252';
        Text5010_Ctx: Label 'TEXT 806,792,"0",180,17,14,"SCHILLER_"';
        Text5011_Ctx: Label 'TEXT 1093,443,"0",180,14,12,"Item"';
        Text5012_Ctx: Label 'TEXT 1093,303,"0",180,10,12,"Part No."';
        Text5013_Ctx: Label 'TEXT 1093,166,"0",180,12,12,"Sr. No."';
        Text5014_Ctx: Label 'TEXT 933,443,"0",180,14,12,":"';
        Text5015_Ctx: Label 'TEXT 933,303,"0",180,14,12,":"';
        Text5016_Ctx: Label 'TEXT 933,166,"0",180,14,12,":"';
        Text5017_Ctx: Label 'TEXT 896,303,"0",180,10,12,"%1"';
        Text5018_Ctx: Label 'TEXT 896,166,"0",180,10,12,"%1"';
        Text5019_Ctx: Label 'QRCODE 330,300,L,10,A,180,M2,S7,"%1"';
        Text5020_Ctx: Label 'TEXT 896,443,"0",180,10,12,"%1"';
        Text5021_Ctx: Label 'PRINT 1,1';
        Text5022_Ctx: Label '<xpml></page></xpml><xpml><end/></xpml>';
        QRCodeInput: Text;
    //TempBlob: Record 99008535;
    begin
        RecGSTLedgEntry.RESET;
        RecGSTLedgEntry.SETRANGE("Document No.", "Sales Invoice Header"."No.");
        RecGSTLedgEntry.SETRANGE("Document Type", RecGSTLedgEntry."Document Type"::Invoice);
        RecGSTLedgEntry.SETRANGE("Transaction Type", RecGSTLedgEntry."Transaction Type"::Sales);
        IF RecGSTLedgEntry.FINDFIRST THEN BEGIN
            QRCodeInput := STRSUBSTNO('%1', RecGSTLedgEntry."Scan QrCode E-Invoice");
            IF QRCodeInput <> '' THEN BEGIN
                //  CreateQRCode(QRCodeInput, TempBlob);
                // RecGSTLedgEntry."QR Code" := TempBlob.Blob;
                RecGSTLedgEntry.MODIFY;
            END;
        END;
    end;

    /*  local procedure CreateQRCode(QRCodeInput: Text[500]; var TempBLOB: Record 99008535)
     var
         QRCodeFileName: Text[1024];
     begin
         CLEAR(TempBLOB);
         QRCodeFileName := GetQRCode(QRCodeInput);
         UploadFileBLOBImportandDeleteServerFile(TempBLOB, QRCodeFileName);
     end; */

    /* local procedure GetQRCode(QRCodeInput: Text[300]) QRCodeFileName: Text[1024]
    var
        [RunOnClient]
        IBarCodeProvider: DotNet IBarcodeProvider;
    begin
        GetBarCodeProvider(IBarCodeProvider);
        QRCodeFileName := IBarCodeProvider.GetBarcode(QRCodeInput);
    end; */

    // //[Scope('Internal')]
    /*  procedure UploadFileBLOBImportandDeleteServerFile(var TempBlob: Record 99008535; FileName: Text[1024])
     var
         FileManagement: Codeunit 419;
     begin
         FileName := FileManagement.UploadFileSilent(FileName);
         FileManagement.BLOBImportFromServerFile(TempBlob, FileName);
         DeleteServerFile(FileName);
     end; */

    // //[Scope('Internal')]
    /* procedure GetBarCodeProvider(var IBarCodeProvider: DotNet IBarcodeProvider)
    var
        [RunOnClient]
        QRCodeProvider: DotNet QRCodeProvider;
    begin
        IF ISNULL(IBarCodeProvider) THEN
            IBarCodeProvider := QRCodeProvider.QRCodeProvider;
    end; */

    /*  local procedure DeleteServerFile(ServerFileName: Text)
     begin
         IF ERASE(ServerFileName) THEN;
     end; */
}


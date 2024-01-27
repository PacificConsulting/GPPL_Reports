report 50339 "Bond Transfer Invoice_NEW Test"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 14Nov2017   RB-N         Not to allowing SaveasPDF from Preview Mode
    // 18Dec2017   RB-N         LR NO., LR DATE, VehicleNo.
    DefaultLayout = RDLC;

    RDLCLayout = 'Src/Report Layout/BondTransferInvoiceNEWTest.rdl';

    PreviewMode = PrintLayout;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(Name_ComInfo; recCompanyInfo.Name)
            {
            }
            column(LocAddr1; LocForm.Address)
            {
            }
            column(Addr2_Loc; LocForm."Address 2")
            {
            }
            column(City_Loc; LocForm.City)
            {
            }
            column(Name_BillTo; "Bill-to Name")
            {
            }
            column(Addr1_BillTo; "Bill-to Address" + ' ' + "Bill-to Address 2")
            {
            }
            column(City_BilTo; "Bill-to City" + '  ' + "Sales Invoice Header"."Bill-to Post Code" + ' ' + BillState.Description + ' ' + BillCountry.Name)
            {
            }
            column(TIN_BillTo; 'TIN NO.  :')// + BillToCust."T.I.N. No.")
            {
            }
            column(CST_BillTo; 'CST No. :')//+ BillToCust."C.S.T. No.")
            {
            }
            column(PAN_BillTo; 'PAN No. :' + BillToCust."P.A.N. No.")
            {
            }
            column(Name_ShipTo; "Ship-to Name")
            {
            }
            column(Addr_Ship; "Ship-to Address" + ' ' + "Ship-to Address 2")
            {
            }
            column(City_ShipTo; "Ship-to City" + ' ' + "Sales Invoice Header"."Ship-to Post Code" + ' ' + ShipState.Description + ' ' + ShipCountry.Name)
            {
            }
            column(TIN_ShipTo; 'TIN NO.  :' + CustTINNO)
            {
            }
            column(PAN_ShipTo; 'PAN No. :' + CustPanNo)
            {
            }
            column(CST_ShipTo; 'CST No. :' + CustCstNo)
            {
            }
            column(No_SalesInvoiceHeader; "Sales Invoice Header"."No.")
            {
            }
            column(TermOfPayment; PayemntTerm.Description)
            {
            }
            column(CustOrderNo; "Order No.")
            {
            }
            column(OrderDate_SalesInvoiceHeader; "Sales Invoice Header"."Order Date")
            {
            }
            column(ExFactoryDec; ExFactoryDec)
            {
            }
            column(Name_ShippingAgent; ShippingAgent.Name)
            {
            }
            column(PostingDate_SalesInvoiceHeader; "Sales Invoice Header"."Posting Date")
            {
            }
            column(SellToGST; SellToGST)
            {
            }
            column(BillToGST; BillToGST)
            {
            }
            column(LocGST; LocGST)
            {
            }
            column(LRRRNo; "Sales Invoice Header"."LR/RR No.")
            {
            }
            column(LRRRDate; "Sales Invoice Header"."LR/RR Date")
            {
            }
            column(VehicleNo; "Sales Invoice Header"."Vehicle No.")
            {
            }
            column(CompPicture; CompanyInfo1.Picture)
            {
            }
            column(CompName; CompanyInfo1.Name)
            {
            }
            column(CompNamePicture; CompanyInfo1."Name Picture")
            {
            }
            column(FactoryCaption; FactoryCaption)
            {
            }
            column(FactoryLocation; LocFrom.Address + ' ' + LocFrom."Address 2" + ' ' + LocFrom.City + ' - ' + LocFrom."Post Code" + ' ' + LocState + ' ' + ' Phone: ' + LocFrom."Phone No." + ' Email: ' + LocFrom."E-Mail")
            {
            }
            column(CompWeb; CompanyInfo1."Home Page")
            {
            }
            column(CompCINNo; 'CompanyInfo1."Company Registration  No."')
            {
            }
            column(CompPANNo; CompanyInfo1."P.A.N. No.")
            {
            }
            column(LocGSTNo; LocFrom."GST Registration No.")
            {
            }
            column(Name_ShipingAgent; recShippingAgent.Name)
            {
            }
            column(No_Header; "No.")
            {
            }
            column(PostingDate; "Posting Date")
            {
            }
            column(LRRRDate_Header; "LR/RR Date")
            {
            }
            column(LRRRNo_Header; "LR/RR No.")
            {
            }
            column(ExternalDocNo; "External Document No.")
            {
            }
            column(ExternalDocDate2; "Order Date")
            {
            }
            column(PlaceofSupply; "Bill-to City")
            {
            }
            column(FrieghtType_Header; "Freight Type")
            {
            }
            column(BillGSTStateCode; BillGSTStateCode)
            {
            }
            column(VehicleNo_Header; "Vehicle No.")
            {
            }
            column(PaymentDescription; "Payment Terms".Description)
            {
            }
            column(SIH_InvoiceTime; "Sales Invoice Header"."Invoice Print Time")
            {
            }
            column(RoadPermitNo_Header; EWayBillNo)
            {
            }
            column(FullName_SIH; BillToCustomerName)
            {
            }
            column(BillAdd1; "Bill-to Address")
            {
            }
            column(BillAdd2; "Bill-to Address 2")
            {
            }
            column(BillAdd3; "Bill-to City" + ' - ' + "Bill-to Post Code")
            {
            }
            column(SIH_BillAdd3; "Bill-to City" + '-' + "Bill-to Post Code" + ' ' + BillStateNew + ', ' + BillCountryNew)
            {
            }
            column(BillPhoneNo; 'Tel : ' + BillPhoneNo)
            {
            }
            column(BillGSTNo; BillGSTNo)
            {
            }
            column(SuppliersCode; SuppliersCode)
            {
            }
            column(Ship2Name; Ship2Name)
            {
            }
            column(SIH_ShipAdd1; "Sales Invoice Header"."Ship-to Address")
            {
            }
            column(SIH_ShipAdd2; "Sales Invoice Header"."Ship-to Address 2")
            {
            }
            column(SIH_ShipAdd3; "Ship-to City" + '-' + "Ship-to Post Code" + ' ' + ShipStateNew + ', ' + ShipCountryNew)
            {
            }
            column(ShipPhoneNo; 'Tel : ' + ShipPhoneNo)
            {
            }
            column(ShipGSTStateCode; ShipGSTStateCode)
            {
            }
            column(ShipGSTNo; ShipGSTNo)
            {
            }
            dataitem(CopyLoop; Integer)
            {
                DataItemTableView = SORTING(Number);
                dataitem(Integer; Integer)
                {
                    DataItemTableView = SORTING(Number)
                                        WHERE(Number = FILTER(1));
                    column(InvType; InvType)
                    {
                    }
                    column(InvType1; InvType1)
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
                                            ORDER(Ascending)
                                            WHERE(Type = FILTER(Item),
                                                  "No." = FILTER(<> ''),
                                                  Quantity = FILTER(<> 0));
                        column(ItemSearchDescription; ItemSearchDescription)
                        {
                        }
                        column(Amount_SalesInvoiceLine; "Sales Invoice Line".Amount)
                        {
                        }
                        column(UnitPrice_SalesInvoiceLine; "Sales Invoice Line"."Unit Price")
                        {
                        }
                        column(UnitofMeasureCode_SalesInvoiceLine; "Sales Invoice Line"."Unit of Measure Code")
                        {
                        }
                        column(Quantity_SalesInvoiceLine; "Sales Invoice Line".Quantity)
                        {
                        }
                        column(AmountWord; AmountWord)
                        {
                        }
                        column(NoText; DELCHR(NoText[1], '=', '*'))
                        {
                        }
                        column(TIN_CompInfo; 'recCompanyInfo."T.I.N. No."')
                        {
                        }
                        column(CST_CompInfo; 'recCompanyInfo."C.S.T No."')
                        {
                        }
                        column(PAN_ComInfo; RecCompInfo."P.A.N. No.")
                        {
                        }
                        column(text1; text1)
                        {
                        }
                        column(text2; text2)
                        {
                        }
                        column(text3; text3)
                        {
                        }
                        column(LineAmount_SalesInvoiceLine; "Sales Invoice Line"."Line Amount")
                        {
                        }
                        column(Rate; Rate)
                        {
                        }
                        column(Amount1; Amount1)
                        {
                        }
                        column(LineNo_SalesInvoiceLine; "Sales Invoice Line"."Line No.")
                        {
                        }
                        column(No_SalesInvoiceLine; "Sales Invoice Line"."No.")
                        {
                        }
                        column(Description_SalesInvoiceLine; "Sales Invoice Line".Description)
                        {
                        }
                        column(DocumentNo_SalesInvoiceLine; "Sales Invoice Line"."Document No.")
                        {
                        }
                        column(vCount; vCount)
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin
                            vCount := "Sales Invoice Line".COUNT;
                            SrNo := SrNo + 1;
                            cnt := "Sales Invoice Line".COUNT;
                            //RSPL--
                            DimensionName := '';
                            //>>Robosoft/Migration/Rahul***Dim
                            /*
                            IF "Inventory Posting Group"='MERCH' THEN BEGIN
                             recPosteddocDiemension.RESET;
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Table ID",113);
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Document No.","Sales Invoice Line"."Document No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Line No.","Sales Invoice Line"."Line No.");
                             recPosteddocDiemension.SETRANGE(recPosteddocDiemension."Dimension Code",'MERCH');
                              IF recPosteddocDiemension.FINDFIRST THEN
                                 dimensionCode:=recPosteddocDiemension."Dimension Value Code";
                            
                               recDimensionValue.RESET;
                               recDimensionValue.SETRANGE(recDimensionValue.Code,dimensionCode);
                               IF recDimensionValue.FINDFIRST THEN
                                 DimensionName:=recDimensionValue.Name;
                                 DimensionName2:=DimensionName;
                            
                            END;
                            */
                            /*IF DimensionName=''THEN
                              DimensionName:='IPOL '+Description;  */

                            //RSPL

                            //RSPL For Category show only item description
                            IF (DimensionName = '') AND ("Sales Invoice Line"."Item Category Code" <> 'CAT01')
                            AND ("Sales Invoice Line"."Item Category Code" <> 'CAT04') AND ("Sales Invoice Line"."Item Category Code" <> 'CAT05')
                            AND ("Sales Invoice Line"."Item Category Code" <> 'CAT06') AND ("Sales Invoice Line"."Item Category Code" <> 'CAT07')
                            AND ("Sales Invoice Line"."Item Category Code" <> 'CAT08') AND ("Sales Invoice Line"."Item Category Code" <> 'CAT09')
                            AND ("Sales Invoice Line"."Item Category Code" <> 'CAT10') THEN
                                DimensionName := 'IPOL ' + Description
                            //ELSE
                            //  DimensionName:=Description;
                            ELSE
                                IF DimensionName <> '' THEN
                                    DimensionName := DimensionName2
                                ELSE
                                    DimensionName := Description;
                            //RSPL

                            //IF Item.GET("No.") THEN
                            ItemSearchDescription := Description;


                            repcheck.InitTextVariable;
                            repcheck.FormatNoText(NoText, "Sales Invoice Line".Amount, "Sales Invoice Header"."Currency Code");

                        end;

                        trigger OnPreDataItem()
                        begin

                            cnt := 0;
                            SrNo := 0;
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin

                        DocTotal := 0; //06Apr2017
                        FExcise := 0; //06Apr2017
                        FEcess := 0; //06Apr2017
                        FShe := 0; //06Apr2017
                        FInvT := 0; //06Apr2017


                        i += 1;
                        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START
                        IF ("Sales Invoice Header"."Print Invoice" = FALSE) THEN BEGIN
                            IF i = 1 THEN
                                InvType := 'ORIGINAL FOR RECIPIENT'
                            //InvType := 'ORIGINAL FOR BUYER'
                            ELSE
                                IF i = 2 THEN
                                    InvType := 'DUPLICATE FOR TRANSPORTER'
                                ELSE
                                    IF i = 3 THEN
                                        InvType := 'TRIPLICATE FOR SUPPLIER'
                                    //InvType := 'TRIPLICATE FOR ASSESSEE'
                                    ELSE
                                        IF i = 4 THEN BEGIN
                                            InvType := 'QUADRAPLICATE FOR H.O';
                                            //InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                                            //InvType1 := 'NOT FOR CENVAT';
                                        END ELSE
                                            IF i = 5 THEN BEGIN
                                                InvType := 'EXTRA COPY';
                                                //InvType := 'EXTRA COPY NOT FOR CENVAT';
                                                InvType1 := '';
                                            END
                                            ELSE
                                                InvType := '';
                        END;

                        //IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'sa') THEN
                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID = 'GPUAE\UNNIKRISHNAN.VS') THEN //18July2017
                        BEGIN
                            IF i = 1 THEN
                                InvType := 'ORIGINAL FOR RECIPIENT'
                            //InvType := 'ORIGINAL FOR BUYER'
                            //InvType := 'ORIGINAL FOR BUYER'
                            ELSE
                                IF i = 2 THEN
                                    InvType := 'DUPLICATE FOR TRANSPORTER'
                                ELSE
                                    IF i = 3 THEN
                                        InvType := 'TRIPLICATE FOR SUPPLIER'
                                    //InvType := 'TRIPLICATE FOR ASSESSEE'
                                    ELSE
                                        IF i = 4 THEN BEGIN
                                            InvType := 'QUADRAPLICATE FOR H.O';
                                            //InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                                            //InvType1 := 'NOT FOR CENVAT';
                                        END ELSE
                                            IF i = 5 THEN BEGIN
                                                InvType := 'EXTRA COPY NOT FOR CENVAT';
                                                InvType1 := '';
                                            END
                                            ELSE
                                                InvType := '';
                        END;

                        //IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'sa')THEN
                        IF ("Sales Invoice Header"."Print Invoice" = TRUE) AND (USERID <> 'GPUAE\UNNIKRISHNAN.VS') THEN //18July2017
                        BEGIN
                            IF i = 1 THEN
                                InvType := 'EXTRA COPY';
                            InvType1 := ' NOT FOR COMMERCIAL USE';
                        END;

                        //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    OutPutNo += 1;
                    IF Number > 1 THEN
                        copytext := Text003;
                    //CurrReport.PAGENO := 1;
                end;

                trigger OnPreDataItem()
                begin

                    DocTotal := 0; //06Apr2017
                    //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-----START

                    //IF NOT(USERID = 'sa') THEN
                    IF NOT (USERID = 'GPUAE\UNNIKRISHNAN.VS') THEN BEGIN
                        IF "Sales Invoice Header"."Print Invoice" = TRUE THEN BEGIN
                            SETRANGE(Number, 1, 1);
                            //  SETRANGE(Number,1,ABS(NoOfCopies));
                        END ELSE BEGIN

                            NoOfCopies := 4;
                            NoOfLoops := ABS(NoOfCopies);//06July2017
                                                         //NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
                            IF NoOfLoops <= 0 THEN
                                NoOfLoops := 1;
                            copytext := '';
                            SETRANGE(Number, 1, NoOfLoops);
                        END;
                    END ELSE BEGIN

                        NoOfCopies := 4;
                        NoOfLoops := ABS(NoOfCopies);//06July2017
                                                     //NoOfLoops := ABS(NoOfCopies) + cust."Invoice Copies" + 1;
                        IF NoOfLoops <= 0 THEN
                            NoOfLoops := 1;
                        copytext := '';
                        SETRANGE(Number, 1, NoOfLoops);
                    END;

                    //SETRANGE(Number,1,NoOfCopies);
                    //MESSAGE('%1',ABS(NoOfCopies));
                    //EBT STIVAN---(29112012)--- To Print BRT Only Once after First Print out of 5 pages are Taken-------END
                    //i:= 1
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //RSPLSUM 29Aug2020>>
                CLEAR(FactoryCaption);
                IF ("Location Code" = 'PLANT01') OR ("Location Code" = 'PLANT02') THEN
                    FactoryCaption := 'Factory'
                ELSE
                    FactoryCaption := 'Matrl. Sold from';

                LocFrom.RESET;
                IF LocFrom.GET("Sales Invoice Header"."Location Code") THEN BEGIN
                    LocGST := LocFrom."GST Registration No.";
                    RecStateNew.RESET;
                    IF RecStateNew.GET(LocFrom."State Code") THEN
                        LocState := RecStateNew.Description;
                END;

                recShippingAgent.RESET;
                recShippingAgent.SETRANGE(recShippingAgent.Code, "Shipping Agent Code");
                IF recShippingAgent.FINDFIRST THEN;

                RecStateNew1.RESET;
                IF RecStateNew1.GET(State) THEN //09Nov2017
                BEGIN
                    BillStateNew := RecStateNew1.Description;
                    BillGSTStateCode := RecStateNew1."State Code (GST Reg. No.)";
                END;

                "Payment Terms".RESET;
                IF "Payment Terms".GET("Sales Invoice Header"."Payment Terms Code") THEN;

                CLEAR(EWayBillNo);
                "Sales Invoice Header".CALCFIELDS("Sales Invoice Header"."EWB No.");
                IF "Sales Invoice Header"."EWB No." <> '' THEN BEGIN
                    EWayBillNo := "Sales Invoice Header"."EWB No.";
                END ELSE BEGIN
                    EWayBillNo := "Sales Invoice Header"."Road Permit No.";
                END;

                CLEAR(BillToCustomerName);
                IF "Sales Invoice Header"."Sell-to Customer No." <> "Sales Invoice Header"."Bill-to Customer No." THEN BEGIN
                    RecCustNew.RESET;
                    IF RecCustNew.GET("Sales Invoice Header"."Bill-to Customer No.") THEN
                        BillToCustomerName := RecCustNew."Full Name";
                END ELSE
                    BillToCustomerName := "Sales Invoice Header"."Full Name";

                RecCountry.RESET;
                IF RecCountry.GET("Bill-to Country/Region Code") THEN
                    BillCountryNew := RecCountry.Name;

                Cus12.RESET;
                IF Cus12.GET("Sales Invoice Header"."Bill-to Customer No.") THEN BEGIN
                    BillGSTNo := Cus12."GST Registration No.";
                    BillPhoneNo := Cus12."Phone No.";//03Sep2019
                END;

                SuppliersCode := '';
                IF Cust09.GET("Sell-to Customer No.") THEN
                    SuppliersCode := Cust09."Our Account No.";

                IF "Sales Invoice Header"."Ship-to Code" = '' THEN
                    Ship2Name := "Sales Invoice Header"."Full Name"
                ELSE
                    Ship2Name := "Sales Invoice Header"."Ship-to Name";

                RecCountry.RESET;
                IF RecCountry.GET("Ship-to Country/Region Code") THEN
                    ShipCountryNew := RecCountry.Name;

                IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code") THEN BEGIN
                    recState.RESET;
                    IF recState.GET(ShipToCode.State) THEN BEGIN
                        ShipStateNew := recState.Description;
                        ShipGSTStateCode := recState."State Code (GST Reg. No.)";
                        ShipGSTNo := ShipToCode."GST Registration No.";
                        ShipPhoneNo := ShipToCode."Phone No.";//03Sep2019
                    END;
                END;

                IF "Ship-to Code" = '' THEN BEGIN
                    ShipState := BillState;
                    ShipGSTStateCode := BillGSTStateCode;
                    ShipGSTNo := BillGSTNo;
                    ShipPhoneNo := BillPhoneNo;//03S
                END;

                //RSPLSUM 29Aug2020<<

                CLEAR(CustPanNo);
                recCustomer.RESET;
                recCustomer.SETRANGE(recCustomer."No.", "Sales Invoice Header"."Sell-to Customer No.");
                IF recCustomer.FINDSET THEN BEGIN
                    CustTINNO := 'recCustomer."T.I.N. No."';
                    CustCstNo := 'recCustomer."C.S.T. No."';
                    CustPanNo := 'recCustomer."P.A.N. No."';
                    SellToGST := 'recCustomer."GST Registration No."';
                END;

                //>>RSPL-Rahul
                BillToCust.RESET;
                BillToCust.SETRANGE(BillToCust."No.", "Sales Invoice Header"."Bill-to Customer No.");
                IF BillToCust.FINDFIRST THEN;
                BillToGST := BillToCust."GST Registration No.";

                IF "Sales Invoice Header"."Ex-Factory" <> '' THEN BEGIN
                    ExFactoryDec := 'Ex-Bond : ' + "Sales Invoice Header"."Ex-Factory";
                    ExFactoryCaption := 'Ex-Bond ';
                END;
                IF PayemntTerm.GET("Sales Invoice Header"."Payment Terms Code") THEN;
                IF ShippingAgent.GET("Sales Invoice Header"."Shipping Agent Code") THEN;
                //<<
                //RSPL Sachin
                "Sales Invoice Header".RESET;
                "Sales Invoice Header".SETRANGE("Sales Invoice Header"."No.", "No.");
                IF "Sales Invoice Header".FINDFIRST THEN BEGIN
                    CurrCode := "Sales Invoice Header"."Currency Code";
                END;
                AmountWord := 'Amounts In Words' + '  ' + CurrCode;
                Rate := 'Rate' + '  ' + CurrCode;
                Amount1 := 'Amount' + '  ' + CurrCode;
                //Rspl Sachin


                recCompanyInfo.GET;
                recCompanyInfo.CALCFIELDS(recCompanyInfo.Picture);

                txtaddess := recCompanyInfo.Address + recCompanyInfo."Address 2" + recCompanyInfo."Post Code";

                loc.GET("Location Code");


                IF NOT (loc."State Code" = 'MAH') THEN BEGIN
                    IF RespCentre.GET(loc."Global Dimension 2 Code") THEN BEGIN
                        RecState1.GET(loc."State Code");
                        BranchOfficeAddress :=
                        RespCentre.Address + ' ' + RespCentre."Address 2" + ' ' + RespCentre.City + ' ' + RespCentre."Post Code"
                        + ' ' + RecState1.Description + ' Phone: ' + RespCentre."Phone No." + ' ' + ' Email: ' + RespCentre."E-Mail";
                        BranchOffText := 'Branch Office :';
                    END;
                END
                ELSE
                    BranchOfficeAddress := '';
                BranchOffText := '';

                IF recState.GET(State) THEN;
                IF RecCompInfo.GET THEN;


                IF BillState.GET("Sales Invoice Header"."GST Bill-to State Code") THEN;
                IF BillCountry.GET("Sales Invoice Header"."Bill-to Country/Region Code") THEN;

                IF ShipState.GET("Sales Invoice Header".State) THEN;
                IF ShipCountry.GET("Sales Invoice Header"."Ship-to Country/Region Code") THEN;
            end;

            trigger OnPreDataItem()
            begin
                IF NOT CurrReport.PREVIEW THEN;//RB-N 14Nov2017

                //IF ("Sales Invoice Header".GETFILTER("Sales Invoice Header"."Location Code" )= '') THEN
                //  ERROR('Please Specify Location Code');
                /*IF BLocation ='' THEN
                  ERROR('Please Specify B/L Location');
                 //>>RSPL-Rahul
                IF LocForm.GET(BLocation) THEN;
                LocGST:=LocForm."GST Registration No.";
                IF LocState.GET(LocForm."State Code") THEN;
                IF LocCountry.GET(LocForm."Country/Region Code") THEN;
                */

                IF SellToLocation.GET("Sales Invoice Header"."Bill-to Customer No.") THEN;
                IF BillToLocation.GET("Sales Invoice Header"."Ship-to Code") THEN;

                //<<

            end;
        }
    }

    requestpage
    {
        SaveValues = false;

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
        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
    end;

    var
        recCompanyInfo: Record 79;
        txtaddess: Text[200];
        recCustomer: Record 18;
        CustTINNO: Code[20];
        CustCstNo: Code[40];
        NoText: array[2] of Text[100];
        repcheck: Report Check;
        CustOrderNo: Code[20];
        SrNo: Integer;
        cnt: Integer;
        loc: Record 14;
        RespCentre: Record 5714;
        RecState1: Record State;
        BranchOfficeAddress: Text[250];
        BranchOffText: Text[50];
        recState: Record State;
        RecCompInfo: Record 79;
        "---------------": Integer;
        recDimensionValue: Record 349;
        DimensionName: Text[100];
        dimensionCode: Code[20];
        DimensionName2: Text[100];
        "---RSPL-Rahul": Integer;
        BLocation: Code[10];
        Location: Record 14;
        LocForm: Record 14;
        ExFactoryDec: Text[40];
        PayemntTerm: Record 3;
        ShippingAgent: Record 291;
        "--Loc": Integer;
        SellToLocation: Record 14;
        BillToLocation: Record 14;
        BillToCust: Record 18;
        LocCountry: Record 9;
        BillState: Record State;
        BillCountry: Record 9;
        ShipState: Record State;
        ShipCountry: Record 9;
        Item: Record 27;
        ItemSearchDescription: Text[80];
        ExFactoryCaption: Text[30];
        CustPanNo: Code[80];
        CurrCode: Code[10];
        AmountWord: Text[350];
        Rate: Text[30];
        Amount1: Text[30];
        text1: Label 'We declare that this Invoice shows the actual price of the goods described and that all particulars are true and correct. ';
        text2: Label 'Declaration Under Customs Not.102/2007- Cus Dated 14-09-2007:- No. Credit of Additional Customs Duty';
        text3: Label 'Levied Under Sub section (5) of Section 3 of The Customs Tariff ACT 1975 Shall be Addmissible for The Goods Sold Under This Invoice.';
        "---": Integer;
        vCount: Integer;
        SellToGST: Code[20];
        BillToGST: Code[20];
        LocGST: Code[20];
        CompanyInfo1: Record 79;
        FactoryCaption: Text[30];
        LocFrom: Record 14;
        RecStateNew: Record State;
        LocState: Text[80];
        DocTotal: Decimal;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        copytext: Text[30];
        OutPutNo: Integer;
        Text003: Label 'Original for Buyer';
        FExcise: Decimal;
        FEcess: Decimal;
        FShe: Decimal;
        FAdd: Decimal;
        FInvT: Decimal;
        i: Integer;
        InvType: Text[50];
        InvType1: Text[50];
        recShippingAgent: Record 291;
        RecStateNew1: Record State;
        BillStateNew: Text[50];
        BillGSTStateCode: Code[10];
        "Payment Terms": Record 3;
        EWayBillNo: Code[20];
        BillToCustomerName: Text[100];
        RecCustNew: Record 18;
        RecCountry: Record 9;
        BillCountryNew: Text[50];
        Cus12: Record 18;
        BillGSTNo: Code[15];
        BillPhoneNo: Text;
        Cust09: Record 18;
        SuppliersCode: Text[20];
        Ship2Name: Text[60];
        ShipCountryNew: Text[50];
        ShipStateNew: Text[50];
        ShipToCode: Record 222;
        ShipGSTStateCode: Code[10];
        ShipGSTNo: Code[15];
        ShiptoName15: Text[100];
        ShipPhoneNo: Text;
        k: Integer;
}


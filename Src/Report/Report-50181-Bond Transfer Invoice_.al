report 50181 "Bond Transfer Invoice_"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BondTransferInvoice.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
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
            column(CompGSTNo; LocForm."GST Registration No.")
            {
            }
            column(CompPANNo; LocForm."T.C.A.N. No.")
            {
            }
            column(State_Loc; LocState.Description + ' ' + LocCountry.Name)
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
            column(TIN_BillTo; '')// 'TIN NO.  :' + BillToCust."T.I.N. No.")
            {
            }
            column(GST_BillTo; '')//'GST No. :' + BillToCust."GST Registration No.")
            {
            }
            column(CST_BillTo; '')// 'CST No. :' + BillToCust."C.S.T. No.")
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
            column(GSTNo; 'GST No. :' + GSTNo)
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
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = WHERE(Type = FILTER(Item),
                                          "No." = FILTER(<> ''));
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
                column(TIN_CompInfo; recCompanyInfo."GST Registration No.")
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
            dataitem(Integer; Integer)
            {
                column(Number; Number)
                {
                }

                trigger OnPreDataItem()
                begin
                    Integer.SETRANGE(Number, 1, 5 - vCount);//19Apr2017
                    //Integer.SETRANGE(Number,1,22-vCount);
                end;
            }

            trigger OnAfterGetRecord()
            begin

                CLEAR(CustPanNo);
                recCustomer.RESET;
                recCustomer.SETRANGE(recCustomer."No.", "Sales Invoice Header"."Sell-to Customer No.");
                IF recCustomer.FINDSET THEN BEGIN
                    CustTINNO := '';// recCustomer."T.I.N. No.";
                    CustCstNo := '';// recCustomer."C.S.T. No.";
                    CustPanNo := recCustomer."P.A.N. No.";
                    GSTNo := recCustomer."GST Registration No."
                END;

                //>>RSPL-Rahul
                BillToCust.RESET;
                BillToCust.SETRANGE(BillToCust."No.", "Sales Invoice Header"."Bill-to Customer No.");
                IF BillToCust.FINDFIRST THEN;

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


                //IF BillState.GET("Sales Invoice Header"."Bill to Customer State") THEN;
                IF BillCountry.GET("Sales Invoice Header"."Bill-to Country/Region Code") THEN;

                IF ShipState.GET("Sales Invoice Header".State) THEN;
                IF ShipCountry.GET("Sales Invoice Header"."Ship-to Country/Region Code") THEN;
            end;

            trigger OnPreDataItem()
            begin

                //IF ("Sales Invoice Header".GETFILTER("Sales Invoice Header"."Location Code" )= '') THEN
                //  ERROR('Please Specify Location Code');
                IF BLocation = '' THEN
                    ERROR('Please Specify B/L Location');
                //>>RSPL-Rahul
                IF LocForm.GET(BLocation) THEN;
                IF LocState.GET(LocForm."State Code") THEN;
                IF LocCountry.GET(LocForm."Country/Region Code") THEN;


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
            area(content)
            {
                field(BLocation; BLocation)
                {
                    Caption = 'B/L Location';

                    trigger OnLookup(var Text: Text): Boolean
                    begin

                        CLEAR(BLocation);
                        Location.RESET;
                        Location.SETRANGE(Location."Trading Location", FALSE);
                        IF Location.FINDFIRST THEN
                            IF PAGE.RUNMODAL(15, Location) = ACTION::LookupOK THEN
                                BLocation := Location.Code;
                    end;
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

    var
        recCompanyInfo: Record 79;
        txtaddess: Text[200];
        recCustomer: Record 18;
        CustTINNO: Code[20];
        CustCstNo: Code[40];
        NoText: array[2] of Text[100];
        repcheck: Report "Check Report";
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
        LocState: Record State;
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
        GSTNo: Code[20];
        CompGSTNo: Code[20];
        CompPANNo: Code[20];
}


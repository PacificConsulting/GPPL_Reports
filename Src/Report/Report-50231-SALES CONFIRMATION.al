report 50231 "SALES CONFIRMATION"
{
    // 
    // 
    // Date        Version     Remarks
    // --------------------------------------------------------------------
    // 09May2018   RB-N        New Report Development as per NAV2009-GPUAE DataBase
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SALESCONFIRMATION.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Sales Header"; "Sales Header")
        {
            column(CurrCode; CurrCode)
            {
            }
            column(LocAdd1; Loc.Address)
            {
            }
            column(LocAdd2; Loc."Address 2")
            {
            }
            column(LocAdd3; Loc.City + '-' + Loc."Post Code" + ', ' + St.Description + ',' + Country1.Name)
            {
            }
            column(DocumentType_SalesHeader; "Sales Header"."Document Type")
            {
            }
            column(SelltoCustomerNo_SalesHeader; "Sales Header"."Sell-to Customer No.")
            {
            }
            column(No_SalesHeader; "Sales Header"."No.")
            {
            }
            column(BilltoCustomerNo_SalesHeader; "Sales Header"."Bill-to Customer No.")
            {
            }
            column(CompLogo; CI.Picture)
            {
            }
            column(SelltoCustomerName_SalesHeader; "Sales Header"."Sell-to Customer Name")
            {
            }
            column(OrderDate_SalesHeader; FORMAT("Sales Header"."Order Date"))
            {
            }
            column(PostingDate_SalesHeader; "Sales Header"."Posting Date")
            {
            }
            column(CompName; CI.Name)
            {
            }
            column(CompAdd1; CI.Address)
            {
            }
            column(CompAdd2; CI."Address 2")
            {
            }
            column(CompRegNo; CI."Registration No.")
            {
            }
            column(CompVATNo; CI."VAT Registration No.")
            {
            }
            column(CompGSTNo; CI."GST Registration No.")
            {
            }
            column(CountryName; Country1.Name)
            {
            }
            column(SalesPer; SalesPer)
            {
            }
            column(BargeOp; BargeOp)
            {
            }
            column(PaymentTermdesc; PaymentTermdesc)
            {
            }
            column(DropShip; DropShip)
            {
            }
            column(VendorId; VendorId)
            {
            }
            column(ContactPerDt; ContactPerDt)
            {
            }
            column(ContactPhn; ContactPhn)
            {
            }
            column(ContactEmail; ContactEmail)
            {
            }
            column(RemarksText; RemarksText)
            {
            }
            column(VesselDesc; VesselDesc)
            {
            }
            column(DateFilter; DateFilter)
            {
            }
            column(LocationofPort_SalesHeader; RecSHAddInfo."Location (of Port)")
            {
            }
            column(TradeTerms_SalesHeader; RecSHAddInfo."Trade Terms")
            {
            }
            column(AgentCode_SalesHeader; RecSHAddInfo."Agent Code")
            {
            }
            column(DealSpec; DealSpec)
            {
            }
            column(DiscPort; DiscPort)
            {
            }
            column(LoadPort; LoadPort)
            {
            }
            dataitem("Sales Line"; "Sales Line")
            {
                DataItemLink = "Document Type" = FIELD("Document Type"),
                               "Document No." = FIELD("No.");
                column(DocumentNo_SalesLine; "Sales Line"."Document No.")
                {
                }
                column(LineNo_SalesLine; "Sales Line"."Line No.")
                {
                }
                column(Type_SalesLine; "Sales Line".Type)
                {
                }
                column(No_SalesLine; "Sales Line"."No.")
                {
                }
                column(LocationCode_SalesLine; "Sales Line"."Location Code")
                {
                }
                column(Specification_SalesLine; "Sales Line".Specification)
                {
                }
                column(Description_SalesLine; "Sales Line".Description)
                {
                }
                column(UnitofMeasureCode_SalesLine; "Sales Line"."Unit of Measure Code")
                {
                }
                column(Quantity_SalesLine; "Sales Line".Quantity)
                {
                }
                column(UnitPrice_SalesLine; "Sales Line"."Unit Price" + ("Unit Price" * (/*"GST %" */0 / 100)))
                {
                }
                column(MQty; MQty)
                {
                }
                column(AmountToCustomer_SalesLine; "Sales Line"."Amount To Customer")
                {
                }
                column(GST_SalesLine; 0)//"Sales Line"."GST %")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>14May2018
                    CLEAR(MQty);
                    IF "Minimum Quantity" <> 0 THEN
                        MQty := FORMAT("Minimum Quantity") + ' - ' + FORMAT(Quantity)
                    ELSE
                        MQty := FORMAT(Quantity);

                    //<<14May2018
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>SalesPerson Name
                SalesPer := '';
                SalesPerson.RESET;
                SalesPerson.SETRANGE(Code, "Salesperson Code");
                IF SalesPerson.FINDFIRST THEN
                    SalesPer := SalesPerson.Name;
                //>>SalesPerson Name

                //>>Shipping Agent
                BargeOp := '';
                ShippingAgent.RESET;
                ShippingAgent.SETRANGE(Code, "Shipping Agent Code");
                IF ShippingAgent.FINDFIRST THEN
                    BargeOp := ShippingAgent.Name;
                //>>Shipping Agent

                //>>Payment Terms
                PaymentTermdesc := '';
                IF PaymentTerms.GET("Payment Terms Code") THEN;
                PaymentTerms.TranslateDescription(PaymentTerms, "Language Code");
                PaymentTermdesc := (PaymentTerms.Description);
                //<<Payment Terms

                //>>
                DropShip := '';
                VendorId := '';
                DateFilter := '';
                SL.RESET;
                SL.SETRANGE("Document Type", "Document Type");
                SL.SETRANGE("Document No.", "No.");
                IF SL.FINDFIRST THEN BEGIN
                    PH.RESET;
                    PH.SETRANGE("No.", SL."Purchase Order No.");
                    IF PH.FINDFIRST THEN
                        VendorId := PH."Buy-from Vendor No.";

                    Ven.RESET;
                    Ven.SETRANGE("No.", VendorId);
                    IF Ven.FINDFIRST THEN
                        VendorId := Ven.Name;


                    IF SL."Drop Shipment" = TRUE THEN
                        DropShip := VendorId
                    ELSE
                        DropShip := CI.Name;

                    DateFilter := FORMAT(SL."Planned Delivery Date") + ' to ' + FORMAT(SL."Planned Delivery End Date");

                END;

                //<<


                //>>
                ContactPerDt := '';
                ContactPhn := '';
                ContactEmail := '';

                Cust.RESET;
                IF Cust.GET("Sell-to Customer No.") THEN BEGIN
                    ContactPerDt := Cust.Contact;
                    ContactPhn := Cust."Phone No.";
                    ContactEmail := Cust."E-Mail";
                END;
                //<<


                //>>
                RemarksText := '';

                SalesComm.RESET;
                SalesComm.SETRANGE("Document Type", "Document Type");
                SalesComm.SETRANGE("No.", "No.");
                IF SalesComm.FINDSET THEN
                    REPEAT
                        RemarksText := RemarksText + ' , ' + SalesComm.Comment;
                    UNTIL SalesComm.NEXT = 0;
                //<<

                RecSHAddInfo.RESET;
                IF RecSHAddInfo.GET(RecSHAddInfo."Document Type"::Order, "Sales Header"."No.") THEN BEGIN
                    //>>
                    VesselDesc := '';
                    Vessels.RESET;
                    IF Vessels.GET(RecSHAddInfo."Vessel Code") THEN BEGIN
                        VesselDesc := Vessels."Vessel Name" + ' ' + Vessels."IMO No.";
                    END;
                    //<<

                    //>>

                    DiscPort := '';
                    PortMaster.RESET;
                    PortMaster.SETRANGE(Code, RecSHAddInfo."Discharge Port");
                    IF PortMaster.FINDFIRST THEN
                        DiscPort := PortMaster."Port Description";

                    LoadPort := '';
                    PortMaster.RESET;
                    PortMaster.SETRANGE(Code, RecSHAddInfo."Load Port");
                    IF PortMaster.FINDFIRST THEN
                        LoadPort := PortMaster."Port Description";

                    //<<
                END;

                //>>28Aug2018
                Loc.RESET;
                IF Loc.GET("Location Code") THEN BEGIN
                    St.RESET;
                    IF St.GET(Loc."State Code") THEN;

                    Country1.RESET;
                    IF Country1.GET(Loc."Country/Region Code") THEN;
                END;

                GLSetup.GET;
                CurrCode := '';
                IF "Currency Code" <> '' THEN
                    CurrCode := "Currency Code"
                ELSE
                    CurrCode := GLSetup."LCY Code";
                //<<28Aug2018
            end;

            trigger OnPreDataItem()
            begin

                //>>Company Information
                CI.GET;
                CI.CALCFIELDS(Picture);

                Country1.RESET;
                IF Country1.GET(CI."Country/Region Code") THEN;
                //<<
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Deal Specification"; DealSpec)
                {
                    ApplicationArea = all;
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
        CI: Record 79;
        SalesPerson: Record 13;
        ShippingAgent: Record 291;
        PaymentTerms: Record 3;
        SalesPer: Text;
        BargeOp: Text;
        PaymentTermdesc: Text;
        SL: Record 37;
        DropShip: Text;
        VendorId: Text;
        PH: Record 38;
        Ven: Record 23;
        Cust: Record 18;
        ContactPerDt: Text;
        ContactPhn: Text;
        ContactEmail: Text;
        SalesComm: Record 44;
        RemarksText: Text;
        VesselDesc: Text;
        Vessels: Record 50052;
        DateFilter: Text;
        DealSpec: Boolean;
        DiscPort: Text;
        LoadPort: Text;
        PortMaster: Record 50051;
        Country1: Record 9;
        MQty: Text;
        Loc: Record 14;
        St: Record State;
        CurrCode: Code[20];
        GLSetup: Record 98;
        RecSHAddInfo: Record 50053;
}


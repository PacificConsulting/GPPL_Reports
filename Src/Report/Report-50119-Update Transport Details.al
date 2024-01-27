report 50119 "Update Transport Details"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/UpdateTransportDetails.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Integer; Integer)
        {
            MaxIteration = 1;

            trigger OnPreDataItem()
            begin

                //>>1

                IF InvNo <> '' THEN BEGIN
                    RecTD.RESET;
                    IF RecTD.GET(InvNo) THEN
                        ERROR('Transport Details for INV are already Updated');
                END;

                IF ShpInv <> '' THEN BEGIN
                    RecTD.RESET;
                    IF RecTD.GET(ShpInv) THEN
                        ERROR('Transport Details for Ship are already Updated');
                END;

                SalesInvHeader.RESET;
                SalesInvHeader.SETRANGE(SalesInvHeader."No.", InvNo);
                IF SalesInvHeader.FINDFIRST THEN BEGIN

                    recTransportDetails.INIT;
                    recTransportDetails."Invoice No." := SalesInvHeader."No.";
                    recTransportDetails."Invoice Date" := SalesInvHeader."Posting Date";
                    recTransportDetails."Customer Name" := SalesInvHeader."Sell-to Customer Name";
                    recTransportDetails."Shortcut Dimension 1 Code" := SalesInvHeader."Shortcut Dimension 1 Code";

                    EBTSalesInvLine.RESET;
                    EBTSalesInvLine.SETRANGE(EBTSalesInvLine."Document No.", SalesInvHeader."No.");
                    EBTSalesInvLine.SETRANGE(EBTSalesInvLine.Type, EBTSalesInvLine.Type::Item);
                    IF EBTSalesInvLine.FINDSET THEN
                        REPEAT

                            IF (EBTSalesInvLine."Unit of Measure" = 'Litres') OR (EBTSalesInvLine."Unit of Measure" = 'KGS') THEN BEGIN
                                UOMTrans := EBTSalesInvLine."Unit of Measure"
                            END
                            ELSE
                                UOMTrans := '';

                        UNTIL EBTSalesInvLine.NEXT = 0;

                    recTransportDetails.UOM := UOMTrans;
                    recTransportDetails.Type := recTransportDetails.Type::Invoice;

                    IF SalesInvHeader."Transport Type" = SalesInvHeader."Transport Type"::"Local+Intercity" THEN BEGIN
                        recTransportDetails."LR No." := '';
                        recTransportDetails."LR Date" := 0D;
                        recTransportDetails."Vehicle No." := '';
                        recTransportDetails."Local LR No." := SalesInvHeader."LR/RR No.";
                        recTransportDetails."Local LR Date" := SalesInvHeader."LR/RR Date";
                        recTransportDetails."Local Vehicle No." := SalesInvHeader."Vehicle No.";
                        recTransportDetails."Local Vehicle Capacity" := SalesInvHeader."Local Vehicle Capacity";  // EBT MILAN 171213
                        recTransportDetails."Vehicle Capacity" := SalesInvHeader."Vehicle Capacity";

                    END ELSE BEGIN
                        recTransportDetails."LR No." := SalesInvHeader."LR/RR No.";
                        recTransportDetails."LR Date" := SalesInvHeader."LR/RR Date";
                        recTransportDetails."Vehicle No." := SalesInvHeader."Vehicle No.";
                        recTransportDetails."Local LR No." := '';
                        recTransportDetails."Local LR Date" := 0D;
                        recTransportDetails."Local Vehicle No." := '';
                        recTransportDetails."Local Vehicle Capacity" := SalesInvHeader."Local Vehicle Capacity";  // EBT MILAN 171213
                        recTransportDetails."Vehicle Capacity" := SalesInvHeader."Vehicle Capacity";

                    END;

                    recTransportDetails."Shipping Agent Code" := SalesInvHeader."Shipping Agent Code";

                    recShippingAgent.RESET;
                    recShippingAgent.SETRANGE(recShippingAgent.Code, SalesInvHeader."Shipping Agent Code");
                    IF recShippingAgent.FINDFIRST THEN BEGIN
                        recTransportDetails."Shipping Agent Name" := recShippingAgent.Name;
                    END;

                    IF SalesInvHeader."Shipping Agent Code" <> '' THEN BEGIN
                        recvendor.RESET;
                        recvendor.SETRANGE(recvendor."Shipping Agent Code", SalesInvHeader."Shipping Agent Code");
                        IF recvendor.FINDFIRST THEN BEGIN
                            recTransportDetails."Vendor Code" := recvendor."No.";
                            recTransportDetails."Vendor Name" := recvendor."Full Name";
                        END;
                    END;

                    recTransportDetails."From Location Code" := SalesInvHeader."Location Code";

                    recLoc.RESET;
                    recLoc.SETRANGE(recLoc.Code, SalesInvHeader."Location Code");
                    IF recLoc.FINDFIRST THEN BEGIN
                        recTransportDetails."From Location Name" := recLoc.Name;
                    END;

                    recTransportDetails.Destination := SalesInvHeader."Ship-to City";


                    recTransportDetails."Expected TPT Cost" := SalesInvHeader."Expected TPT Cost";

                    recTransportDetails."Local Expected TPT Cost" := SalesInvHeader."Local Expected TPT Cost";

                    recSIL.RESET;
                    recSIL.SETRANGE(recSIL."Document No.", SalesInvHeader."No.");
                    recSIL.SETRANGE(recSIL.Type, recSIL.Type::Item);
                    recSIL.SETFILTER(recSIL.Quantity, '<>%1', 0);
                    IF recSIL.FINDFIRST THEN
                        REPEAT
                            recTransportDetails.Quantity += recSIL."Quantity (Base)";
                        UNTIL recSIL.NEXT = 0;

                    recTransportDetails."Freight Type" := SalesInvHeader."Freight Type";

                    recTransportDetails.Open := TRUE;
                    recTransportDetails.INSERT;
                END;

                recTrShipHdr.RESET;
                recTrShipHdr.SETRANGE(recTrShipHdr."No.", ShpInv);
                IF recTrShipHdr.FINDFIRST THEN BEGIN
                    recTransportDetails.INIT;
                    recTransportDetails."Invoice No." := recTrShipHdr."No.";
                    recTransportDetails."Invoice Date" := recTrShipHdr."Posting Date";
                    //recTransportDetails."Customer Name" := recTrShipHdr."Sell-to Customer Name";
                    recTransportDetails."Shortcut Dimension 1 Code" := recTrShipHdr."Shortcut Dimension 1 Code";
                    recTrShipLn.RESET;
                    recTrShipLn.SETRANGE("Document No.", recTrShipHdr."No.");
                    IF recTrShipLn.FINDSET THEN
                        REPEAT
                            IF (recTrShipLn."Unit of Measure" = 'Litres') OR (recTrShipLn."Unit of Measure" = 'KGS') THEN BEGIN
                                UOMTrans := recTrShipLn."Unit of Measure"
                            END
                            ELSE
                                UOMTrans := '';
                        UNTIL recTrShipLn.NEXT = 0;
                    recTransportDetails.UOM := UOMTrans;
                    recTransportDetails.Type := recTransportDetails.Type::Transfer;

                    IF recTrShipHdr."Transport Type" = recTrShipHdr."Transport Type"::"Local+Intercity" THEN BEGIN
                        recTransportDetails."LR No." := '';
                        recTransportDetails."LR Date" := 0D;
                        recTransportDetails."Vehicle No." := '';
                        recTransportDetails."Local LR No." := recTrShipHdr."LR/RR No.";
                        recTransportDetails."Local LR Date" := recTrShipHdr."LR/RR Date";
                        recTransportDetails."Local Vehicle No." := recTrShipHdr."Vehicle No.";
                        recTransportDetails."Local Vehicle Capacity" := recTrShipHdr."Local Vehicle Capacity";  // EBT MILAN 171213
                        recTransportDetails."Vehicle Capacity" := recTrShipHdr."Vehicle Capacity";

                    END
                    ELSE BEGIN
                        recTransportDetails."LR No." := recTrShipHdr."LR/RR No.";
                        recTransportDetails."LR Date" := recTrShipHdr."LR/RR Date";
                        recTransportDetails."Vehicle No." := recTrShipHdr."Vehicle No.";
                        recTransportDetails."Local LR No." := '';
                        recTransportDetails."Local LR Date" := 0D;
                        recTransportDetails."Local Vehicle No." := '';
                        recTransportDetails."Local Vehicle Capacity" := recTrShipHdr."Local Vehicle Capacity";  // EBT MILAN 171213
                        recTransportDetails."Vehicle Capacity" := recTrShipHdr."Vehicle Capacity";

                    END;

                    recTransportDetails."Shipping Agent Code" := recTrShipHdr."Shipping Agent Code";

                    recShippingAgent.RESET;
                    recShippingAgent.SETRANGE(recShippingAgent.Code, recTrShipHdr."Shipping Agent Code");
                    IF recShippingAgent.FINDFIRST THEN BEGIN
                        recTransportDetails."Shipping Agent Name" := recShippingAgent.Name;
                    END;

                    IF recTrShipHdr."Shipping Agent Code" <> '' THEN BEGIN
                        recvendor.RESET;
                        recvendor.SETRANGE(recvendor."Shipping Agent Code", recTrShipHdr."Shipping Agent Code");
                        IF recvendor.FINDFIRST THEN BEGIN
                            recTransportDetails."Vendor Code" := recvendor."No.";
                            recTransportDetails."Vendor Name" := recvendor."Full Name";
                        END;
                    END;

                    recTransportDetails."From Location Code" := recTrShipHdr."Transfer-from Code";

                    recLoc.RESET;
                    recLoc.SETRANGE(recLoc.Code, SalesInvHeader."Location Code");
                    IF recLoc.FINDFIRST THEN BEGIN
                        recTransportDetails."From Location Name" := recLoc.Name;
                    END;

                    recTransportDetails.Destination := recTrShipHdr."Transfer-to Code";


                    recTransportDetails."Expected TPT Cost" := recTrShipHdr."Expected TPT Cost";

                    recTransportDetails."Local Expected TPT Cost" := recTrShipHdr."Local Expected TPT Cost";

                    recTrShipLn.RESET;
                    recTrShipLn.SETRANGE("Document No.", recTrShipHdr."No.");
                    recTrShipLn.SETFILTER(Quantity, '<>%1', 0);
                    IF recTrShipLn.FINDFIRST THEN
                        REPEAT
                            recTransportDetails.Quantity += recTrShipLn."Quantity (Base)";
                        UNTIL recTrShipLn.NEXT = 0;


                    recTransportDetails.Open := TRUE;
                    recTransportDetails.INSERT;
                END;

                MESSAGE('Transport Details are updted for %1', InvNo);
                //<<1
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Invoice No"; InvNo)
                {
                    TableRelation = "Sales Invoice Header";
                }
                field("Shipment No"; ShpInv)
                {
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
        recTransportDetails: Record 50020;
        SalesInvHeader: Record 112;
        EBTSalesInvLine: Record 113;
        UOMTrans: Text[100];
        recShippingAgent: Record 291;
        recvendor: Record 23;
        recLoc: Record 14;
        recSIL: Record 113;
        InvNo: Code[20];
        RecTD: Record 50020;
        ShpInv: Code[20];
        recTrShipHdr: Record 5744;
        recTrShipLn: Record 5745;
}


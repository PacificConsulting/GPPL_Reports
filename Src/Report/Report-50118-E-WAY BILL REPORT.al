report 50118 "E-WAY BILL REPORT"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/EWAYBILLREPORT.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            //CalcFields = "Amount to Customer";
            column(EWAYBillNo; EWayBillNo)
            {
            }
            column(GeneratedDate; GeneratedDate)
            {
            }
            column(GeneratedBy; RecDetGSTLE."Location  Reg. No.")
            {
            }
            column(ValidUpto; ValidUpto)
            {
            }
            column(Mode; "Sales Invoice Header"."Mode of Transport")
            {
            }
            column(ApproxDistance; "Sales Invoice Header"."Distance in kms")
            {
            }
            column(Type; Type)
            {
            }
            column(DocumentDetails; 'Tax Invoice - ' + "Sales Invoice Header"."No." + ' - ' + FORMAT("Sales Invoice Header"."Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(TransactionType; "Sales Invoice Header"."EWB Transaction Type")
            {
            }
            column(From_GSTIN; DetaiGSTLedEntry."Location  Reg. No.")
            {
            }
            column(From_LocName; FromLocName)
            {
            }
            column(FromLocState; FromLocState)
            {
            }
            column(DispatchFrom1; FromAddr[1])
            {
            }
            column(DispatchFrom2; FromAddr[2])
            {
            }
            column(DispatchFrom3; FromAddr[3])
            {
            }
            column(ToGSTIN; ToGSTIN)
            {
            }
            column(ToLocName; ToLocName)
            {
            }
            column(ToLocState; ToLocStateBillTo)
            {
            }
            column(ShipTo1; ToAddr[1])
            {
            }
            column(ShipTo2; ToAddr[2])
            {
            }
            column(ShipTo3; ToAddr[3])
            {
            }
            column(TotTaxAmt; TotTaxAmt)
            {
            }
            column(TotalInvAmt; TotInvAmt)
            {
            }
            column(CGST; CGST)
            {
            }
            column(SGST; SGST)
            {
            }
            column(IGST; IGST)
            {
            }
            column(Cess; Cess)
            {
            }
            column(CessNonAdvol; CessNonAdvol)
            {
            }
            column(OtherAmt; OtherAmt)
            {
            }
            column(TranspoterDocNoDate; "Sales Invoice Header"."LR/RR No." + ' & ' + FORMAT("Sales Invoice Header"."LR/RR Date"))
            {
            }
            column(TransporterIDName; TransporterIDName)
            {
            }
            column(QRCODE1; 'DetaiGSTLedgEnt."QR Code"')
            {
            }
            column(Barcode2; recTempBlob.Barcode)
            {
            }
            column(ShipToName; "Sales Invoice Header"."Ship-to Name")
            {
            }
            column(ModeShipMethod; ModeShipmentMethod)
            {
            }
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Type = FILTER(Item),
                                          Quantity = FILTER(<> 0));
                column(HSNCode; "Sales Invoice Line"."HSN/SAC Code")
                {
                }
                column(ProdNameDesc; ProdNameDesc)
                {
                }
                column(HSNDesc; HSNSACDIS)
                {
                }
                column(Quantity; "Sales Invoice Line".Quantity)
                {
                }
                column(TaxableAmount; TaxableAmt)
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
                column(CessPer; CessPer)
                {
                }
                column(CessNonAdvolPer; CessNonAdvolPer)
                {
                }
                column(GSTReportingQOC; GSTReportingQOC)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    //RecItem.RESET;
                    //IF RecItem.GET("Sales Invoice Line"."No.") THEN
                    //ProdNameDesc := RecItem.Description + '&' + "Sales Invoice Line".Description;
                    IF "Sales Invoice Line"."Item Category Code" = 'CAT21' THEN BEGIN
                        HSNSACDIS := "Sales Invoice Line".Description;
                    END ELSE BEGIN
                        CLEAR(HSNSACDIS);
                        RecHSNSAC.RESET;
                        RecHSNSAC.SETRANGE(Code, "Sales Invoice Line"."HSN/SAC Code");
                        IF RecHSNSAC.FINDFIRST THEN
                            HSNSACDIS := RecHSNSAC.Description;
                    END;
                    CLEAR(CGSTPer);
                    CLEAR(SGSTPer);
                    CLEAR(IGSTPer);
                    RecDGSTLE.RESET;
                    RecDGSTLE.SETRANGE("Document No.", "Document No.");
                    RecDGSTLE.SETRANGE("Entry Type", RecDGSTLE."Entry Type"::"Initial Entry");
                    //RecDGSTLE.SETRANGE("Original Doc. Type", RecDGSTLE."Original Doc. Type"::Invoice);
                    RecDGSTLE.SETRANGE("Transaction Type", RecDGSTLE."Transaction Type"::Sales);
                    RecDGSTLE.SETRANGE("Document Line No.", "Sales Invoice Line"."Line No.");
                    IF RecDGSTLE.FINDSET THEN
                        REPEAT
                            IF RecDGSTLE."GST Component Code" = 'CGST' THEN
                                CGSTPer := RecDGSTLE."GST %";
                            IF RecDGSTLE."GST Component Code" = 'SGST' THEN
                                SGSTPer := RecDGSTLE."GST %";
                            IF RecDGSTLE."GST Component Code" = 'IGST' THEN
                                IGSTPer := RecDGSTLE."GST %";
                        UNTIL RecDGSTLE.NEXT = 0;

                    CLEAR(GSTReportingQOC);
                    RecUnitOfMeasure.RESET;
                    IF RecUnitOfMeasure.GET("Sales Invoice Line"."Unit of Measure Code") THEN
                        GSTReportingQOC := '';// RecUnitOfMeasure."GST Reporting UQC";

                    //RSPLSUM 08Oct2020>>
                    CLEAR(TaxableAmt);
                    IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                        TaxableAmt := 0;// "Sales Invoice Line"."GST Base Amount" / "Sales Invoice Header"."Currency Factor"
                                        //ELSE
                    TaxableAmt := 0;// "Sales Invoice Line"."GST Base Amount";
                    //RSPLSUM 08Oct2020<<
                end;
            }
            dataitem("GST Ledger Entry"; "GST Ledger Entry")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Entry No.");
                column(Unique_No; "GST Ledger Entry"."E-Way Bill No.")
                {
                }
                column(Entered_Date; EWBUpdateDate)
                {
                }
                column(Entered_By; recCompInfo."GST Registration No." + '  ' + recCompInfo.Name)
                {
                }
                column(LocName; LocName)
                {
                }
                column(VehicleTransDoc; VehicleTransDoc)
                {
                }
                column(ModeVehicle1; Mode)
                {
                }
                column(DistanceInKm; DistanceInKm)
                {
                }
                column(ValidUpto1; ValidUpto)
                {
                }
                column(LocPin; LocPin)
                {
                }
                column(DocNo; "GST Ledger Entry"."Document No.")
                {
                }
                column(SIHDate; SIHDate)
                {
                }
                column(AmtToCut; AmtToCut)
                {
                }
                column(HSNCODE1; HSNCODE)
                {
                }
                column(HSNSACDIS; HSNSACDIS)
                {
                }
                column(ReaTran1; ReaTran1)
                {
                }
                column(QRCODE2; 'DetaiGSTLedgEnt."QR Code"')
                {
                }
                column(Barcode; recTempBlob.Barcode)
                {
                }
                column(GSTINRec; GSTINRec)
                {
                }
                column(TranspoCode; TranspoCode)
                {
                }
                column(TranspoGST; TranspoGST)
                {
                }
                column(Barcode3; recTempBlob.Barcode)
                {
                }
                column(CustName3; CustName3)
                {
                }
                column(PlaceofDelivery; PlaceofDelivery)
                {
                }
                column(TranspoName; TranspoName)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    //recEWB.GET("GST Ledger Entry"."Document No.", "GST Ledger Entry"."E-Way Bill No.");
                    recEWB.SETRANGE(recEWB."Document No.", "GST Ledger Entry"."Document No."); //RSPL-AR AR-EWB
                    recEWB.SETRANGE(recEWB."EWB No.", "GST Ledger Entry"."E-Way Bill No.");    //RSPL-AR AR-EWB
                    recCompInfo.GET;

                    //AKT_EWB 25888 16012020
                    CLEAR(Mode);
                    SalInvHead1.RESET;
                    SalInvHead1.SETRANGE("No.", "GST Ledger Entry"."Document No.");
                    IF SalInvHead1.FINDFIRST THEN BEGIN
                        Mode := SalInvHead1."Mode of Transport";
                        VehicleTransDoc := SalInvHead1."Vehicle No.";
                        DistanceInKm := SalInvHead1."Distance in kms";
                        IF Location1.GET(SalInvHead1."Location Code") THEN
                            //LocName := Location1.Name;
                            SIHDate := SalInvHead1."Document Date";
                    END;

                    //AKT_EWB 25888 16012020


                    //AKT_EWB 25888 18012020
                    SalInvLine.RESET;
                    SalInvLine.SETRANGE("Document No.", "Document No.");
                    IF SalInvLine.FINDSET THEN BEGIN
                        AmtToCut += 0;//SalInvLine."Amount To Customer";
                    END;

                    //AKT_EWB 25888 18012020

                    //AKT_EWB 31012020
                    IF Cust3.GET("GST Ledger Entry"."Source No.") THEN
                        PlaceofDelivery := Cust3.Address;

                    //AKT_EWB 31012020

                    //AKT_EWB 13022020
                    DetaiGSTLedEntry.RESET;
                    DetaiGSTLedEntry.SETRANGE("Document No.", "Document No.");
                    DetaiGSTLedEntry.SETRANGE("Posting Date", "Posting Date");
                    IF DetaiGSTLedEntry.FINDFIRST THEN BEGIN
                        //HSNCODE :=  DetaiGSTLedEntry."HSN/SAC Code";
                        LocRegNo := DetaiGSTLedEntry."Location  Reg. No.";
                        //DetaiGSTLedEntry.CALCFIELDS(DetaiGSTLedEntry."QR Code");
                        GSTINRec := DetaiGSTLedEntry."Buyer/Seller Reg. No.";
                    END;

                    Customer1.RESET;
                    Customer1.SETRANGE("No.", "Source No.");
                    IF Customer1.FINDFIRST THEN BEGIN
                        //GSTINRec := Customer1."GST Registration No.";
                        CustName3 := Customer1.Name;
                    END;
                    /*
                    SalInvHdr.RESET;
                    SalInvHdr.SETRANGE("No.","Document No.");
                    IF SalInvHdr.FINDFIRST THEN BEGIN
                       ReaTran := SalInvHdr."GST Customer Type";
                       IF ReaTran = ReaTran::Export  THEN
                           ReaTran1 := 'Export'
                         ELSE
                           ReaTran1 := 'Supply';
                    END;
                    */
                    SalInvHead2.RESET;
                    SalInvHead2.SETRANGE("No.", "GST Ledger Entry"."Document No.");
                    IF SalInvHead2.FINDFIRST THEN BEGIN
                        LocCode2 := SalInvHead2."Location Code";
                        TranspoCode := SalInvHead2."Transporter Code";

                        IF Location1.GET(LocCode2) THEN
                            //LocName := Location1.Name;
                            LocPin := Location1."Post Code";
                    END;

                    IF Transporter1.GET(TranspoCode) THEN BEGIN
                        TranspoCode := Transporter1.Name;
                        TranspoGST := Transporter1."GSTIN No.";
                        TranspoName := Transporter1.Name;
                    END;

                    /*
                    GSTLedgerEntry.RESET;
                    //GSTLedgerEntry.SETRANGE("Entry No.","Entry No.");
                    GSTLedgerEntry.SETRANGE("Document No.","Document No.");
                    GSTLedgerEntry.SETRANGE("Posting Date","Posting Date");
                    IF GSTLedgerEntry.FINDFIRST THEN BEGIN
                       EWayBillNo := GSTLedgerEntry."E-Way Bill No.";
                       DocNo1 := GSTLedgerEntry."Document No.";
                       EntryNo := GSTLedgerEntry."Entry No.";
                    END;
                    */

                    CLEAR(EWBUpdateDate);
                    DetailedsEWayBill1.RESET;
                    DetailedsEWayBill1.SETRANGE("Document No.", "GST Ledger Entry"."Document No.");
                    DetailedsEWayBill1.SETRANGE("EWB No.", EWayBillNo);
                    IF DetailedsEWayBill1.FINDFIRST THEN BEGIN
                        EWBUpdateDate := DetailedsEWayBill1."EWB Updated Date";
                        ValidUpto := DetailedsEWayBill1."EWB Valid Upto";
                    END;


                    //MESSAGE('%1,%2,%3',EWayBillNo,LocRegNo,EWBUpdateDate);
                    //BarcodeForQuarantineLabel(EWayBillNo,LocRegNo,EWBUpdateDate);


                    //EncodeCode128(FORMAT(EWayBillNo),2,FALSE,recTempBlob);



                    //AKT_EWB 13022020

                end;
            }
            dataitem("Detailed E-Way Bill"; "Detailed E-Way Bill")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "EWB No.", CharA, "Vehicle No.", "Transporter Code")
                                    WHERE(Cancelled = CONST(false));
                column(ModeVehicle; RecShipmentMethod.Description)
                {
                }
                column(VehicleTransDocNoDate; "Detailed E-Way Bill"."Vehicle No." + ' & ' + "Detailed E-Way Bill"."Trans. Doc. No." + ' & ' + "Detailed E-Way Bill"."Trans. Doc. Date")
                {
                }
                column(VehFrom; "Detailed E-Way Bill"."From Place")
                {
                }
                column(EWBUpdateDate; EWBUpdatedDate)
                {
                }
                column(LocationRegNo2; RecDGSTLENew."Location  Reg. No.")
                {
                }
                column(VehicleNo; "Detailed E-Way Bill"."Vehicle No.")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    RecShipmentMethod.RESET;
                    IF RecShipmentMethod.GET("Sales Invoice Header"."Shipment Method Code") THEN;

                    RecDGSTLENew.RESET;
                    RecDGSTLENew.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                    RecDGSTLENew.SETRANGE("Posting Date", "Sales Invoice Header"."Posting Date");
                    IF RecDGSTLENew.FINDFIRST THEN;

                    CLEAR(EWBUpdatedDate);
                    IF "Detailed E-Way Bill"."VH Updated Date" <> '' THEN
                        EWBUpdatedDate := "Detailed E-Way Bill"."VH Updated Date"
                    ELSE
                        EWBUpdatedDate := "Detailed E-Way Bill"."EWB Updated Date";
                end;
            }

            trigger OnAfterGetRecord()
            begin
                recCompInfo.GET;

                CLEAR(EWayBillNo);
                GSTLedgerEntry.RESET;
                GSTLedgerEntry.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                GSTLedgerEntry.SETRANGE("Posting Date", "Sales Invoice Header"."Posting Date");
                IF GSTLedgerEntry.FINDFIRST THEN
                    EWayBillNo := GSTLedgerEntry."E-Way Bill No.";

                CLEAR(EWBUpdateDate);
                DetailedsEWayBill1.RESET;
                DetailedsEWayBill1.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                DetailedsEWayBill1.SETRANGE("EWB No.", EWayBillNo);
                IF DetailedsEWayBill1.FINDFIRST THEN
                    EWBUpdateDate := DetailedsEWayBill1."EWB Updated Date";

                DetailedsEWayBill1.RESET;
                DetailedsEWayBill1.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                DetailedsEWayBill1.SETRANGE("EWB No.", EWayBillNo);
                IF DetailedsEWayBill1.FINDFIRST THEN BEGIN
                    GeneratedDate := DetailedsEWayBill1."EWB Updated Date";
                    ValidUpto := DetailedsEWayBill1."EWB Valid Upto";
                END;

                CLEAR(LocRegNo);
                DetaiGSTLedEntry.RESET;
                DetaiGSTLedEntry.SETRANGE("Document No.", "No.");
                DetaiGSTLedEntry.SETRANGE("Posting Date", "Posting Date");
                IF DetaiGSTLedEntry.FINDFIRST THEN BEGIN
                    //HSNCODE :=  DetaiGSTLedEntry."HSN/SAC Code";
                    LocRegNo := DetaiGSTLedEntry."Location  Reg. No.";
                    // DetaiGSTLedEntry.CALCFIELDS(DetaiGSTLedEntry."QR Code");
                    GSTINRec := DetaiGSTLedEntry."Buyer/Seller Reg. No.";
                END;

                CLEAR(FromLocName);
                CLEAR(FromLocState);
                CLEAR(FromAddr);
                //ss--CLEAR(VehFrom);
                IF Location1.GET("Sales Invoice Header"."Location Code") THEN BEGIN
                    //RSPLSUM 30Jul2020--FromLocName := Location1.Name;
                    FromLocName := recCompInfo.Name;//RSPLSUM 30Jul2020
                                                    //ss--VehFrom := Location1.City;

                    IF "Sales Invoice Header"."Dispatch Code" <> '' THEN BEGIN
                        RecState.RESET;
                        IF RecState.GET(Location1."State Code") THEN BEGIN
                            FromLocState := RecState.Description;
                            IF recDispatchFromLocation.GET("Dispatch Code", "Location Code") THEN BEGIN
                                FromAddr[1] := recDispatchFromLocation."Address 1";
                                FromAddr[2] := recDispatchFromLocation."Address 2";
                                FromAddr[3] := recDispatchFromLocation.City + ',' + FromLocState + '-' + recDispatchFromLocation."Post Code";
                            END;
                        END;
                    END ELSE BEGIN
                        RecState.RESET;
                        IF RecState.GET(Location1."State Code") THEN BEGIN
                            FromLocState := RecState.Description;
                            FromAddr[1] := Location1.Address;
                            FromAddr[2] := Location1."Address 2";
                            FromAddr[3] := Location1.City + ',' + FromLocState + '-' + Location1."Post Code";
                        END;
                    END;
                END;

                CLEAR(ToGSTIN);
                CLEAR(ToLocName);
                CLEAR(ToLocState);
                CLEAR(ToAddr);
                //RSPLSUM 04Aug2020>>
                CLEAR(ToLocStateBillTo);
                RecCust.RESET;
                IF RecCust.GET("Sales Invoice Header"."Bill-to Customer No.") THEN BEGIN
                    ToGSTIN := RecCust."GST Registration No.";
                    ToLocName := RecCust.Name;
                    RecState.RESET;
                    IF RecState.GET(RecCust."State Code") THEN
                        ToLocStateBillTo := RecState.Description;
                END;
                //RSPLSUM 04Aug2020<<

                IF "Sales Invoice Header"."Ship-to Code" <> '' THEN BEGIN
                    RecShipToAddr.RESET;
                    IF RecShipToAddr.GET("Sales Invoice Header"."Sell-to Customer No.", "Sales Invoice Header"."Ship-to Code") THEN BEGIN
                        //ToGSTIN := RecShipToAddr."GST Registration No.";
                        RecState.RESET;
                        IF RecState.GET(RecShipToAddr.State) THEN
                            ToLocState := RecState.Description;
                    END;
                    //ToLocName := "Sales Invoice Header"."Ship-to Name";
                    ToAddr[1] := "Sales Invoice Header"."Ship-to Address";
                    ToAddr[2] := "Sales Invoice Header"."Ship-to Address 2";
                    ToAddr[3] := "Sales Invoice Header"."Ship-to City" + ',' + ToLocState + '-' + "Sales Invoice Header"."Ship-to Post Code";
                END ELSE BEGIN
                    RecCust.RESET;
                    IF RecCust.GET("Sales Invoice Header"."Sell-to Customer No.") THEN BEGIN
                        //ToGSTIN := RecCust."GST Registration No.";
                        //ToLocName := RecCust.Name;
                        RecState.RESET;
                        IF RecState.GET(RecCust."State Code") THEN
                            ToLocState := RecState.Description;
                        ToAddr[1] := RecCust.Address;
                        ToAddr[2] := RecCust."Address 2";
                        ToAddr[3] := RecCust.City + ',' + ToLocState + '-' + RecCust."Post Code";
                    END;
                END;

                CLEAR(TotTaxAmt);
                RecSIL.RESET;
                RecSIL.SETRANGE("Document No.", "No.");
                RecSIL.SETFILTER(Type, '<>%1', RecSIL.Type::" ");
                RecSIL.SETFILTER(Quantity, '<>%1', 0);//RSPLSUM 20Jul2020
                IF RecSIL.FINDSET THEN
                    REPEAT
                        //TotTaxAmt += RecSIL.Amount;
                        IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                            TotTaxAmt += 0// RecSIL."GST Base Amount" / "Sales Invoice Header"."Currency Factor"//RSPLSUM 20Jul2020
                        ELSE
                            TotTaxAmt += 0;// RecSIL."GST Base Amount";
                    UNTIL RecSIL.NEXT = 0;

                CLEAR(CGST);
                CLEAR(SGST);
                CLEAR(IGST);
                RecDGSTLE.RESET;
                RecDGSTLE.SETRANGE("Document No.", "No.");
                RecDGSTLE.SETRANGE("Entry Type", RecDGSTLE."Entry Type"::"Initial Entry");
                //RecDGSTLE.SETRANGE("Original Doc. Type", RecDGSTLE."Original Doc. Type"::Invoice);
                RecDGSTLE.SETRANGE("Transaction Type", RecDGSTLE."Transaction Type"::Sales);
                IF RecDGSTLE.FINDSET THEN
                    REPEAT
                        IF RecDGSTLE."GST Component Code" = 'CGST' THEN BEGIN
                            IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                CGST += ABS(RecDGSTLE."GST Amount" / "Sales Invoice Header"."Currency Factor")
                            ELSE
                                CGST += ABS(RecDGSTLE."GST Amount");
                        END;
                        IF RecDGSTLE."GST Component Code" = 'SGST' THEN BEGIN
                            IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                SGST += ABS(RecDGSTLE."GST Amount" / "Sales Invoice Header"."Currency Factor")
                            ELSE
                                SGST += ABS(RecDGSTLE."GST Amount");
                        END;
                        IF RecDGSTLE."GST Component Code" = 'IGST' THEN BEGIN
                            IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                IGST += ABS(RecDGSTLE."GST Amount" / "Sales Invoice Header"."Currency Factor")
                            ELSE
                                IGST += ABS(RecDGSTLE."GST Amount");
                        END;
                    UNTIL RecDGSTLE.NEXT = 0;

                CLEAR(Type);
                IF "Sales Invoice Header"."GST Customer Type" = "Sales Invoice Header"."GST Customer Type"::Export THEN
                    Type := 'Export'
                ELSE
                    Type := 'Outward - Supply';

                /*
                CLEAR(TransporterIDName);
                RecShippingAgent.RESET;
                IF RecShippingAgent.GET("Sales Invoice Header"."Shipping Agent Code") THEN
                  TransporterIDName := RecShippingAgent."GST Registration No." + ' & ' + RecShippingAgent.Name;
                *///RSPLSUM

                //ss--RecShipmentMethod.RESET;
                //ss--IF RecShipmentMethod.GET("Sales Invoice Header"."Shipment Method Code") THEN;

                /*
                CLEAR(TransporterIDName);
                RecVend.RESET;
                RecVend.SETRANGE("Shipping Agent Code","Sales Invoice Header"."Shipping Agent Code");
                IF RecVend.FINDFIRST THEN
                  TransporterIDName := RecVend."GST Registration No." + ' & ' + RecVend.Name;
                */
                RecDetGSTLE.RESET;
                RecDetGSTLE.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                RecDetGSTLE.SETRANGE("Posting Date", "Sales Invoice Header"."Posting Date");
                IF RecDetGSTLE.FINDFIRST THEN;

                //RSPLSUM BEGIN>>
                CLEAR(TransporterIDName);
                RecDetEwayBillNew.RESET;
                RecDetEwayBillNew.SETRANGE("Document No.", "Sales Invoice Header"."No.");
                IF RecDetEwayBillNew.FINDLAST THEN
                    TransporterIDName := RecDetEwayBillNew."Transporter Code" + ' & ' + RecDetEwayBillNew."Transporter Name";
                //RSPLSUM END<<

                //RSPLSUM-TCS>>
                CLEAR(OtherAmt);
                RecSalesInvLine.RESET;
                RecSalesInvLine.SETCURRENTKEY("Document No.", Type, Quantity);
                RecSalesInvLine.SETRANGE("Document No.", "No.");
                RecSalesInvLine.SETFILTER(Quantity, '<>%1', 0);
                IF RecSalesInvLine.FINDSET THEN
                    REPEAT
                        IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                            OtherAmt += 0;//RecSalesInvLine."TDS/TCS Amount" / "Sales Invoice Header"."Currency Factor"
                                          //ELSE
                        OtherAmt += 0;// RecSalesInvLine."TDS/TCS Amount";
                    UNTIL RecSalesInvLine.NEXT = 0;
                //RSPLSUM-TCS<<

                CLEAR(TotInvAmt);
                //"Sales Invoice Header".CALCFIELDS("Sales Invoice Header"."Amount to Customer");
                IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                    TotInvAmt := 0;//"Sales Invoice Header"."Amount to Customer" / "Sales Invoice Header"."Currency Factor"
                                   //ELSE
                TotInvAmt := 0;//"Sales Invoice Header"."Amount to Customer";

                //RSPLSUM 07Nov2020>>
                CLEAR(ModeShipmentMethod);
                RecShipmentMethod.RESET;
                IF RecShipmentMethod.GET("Sales Invoice Header"."Shipment Method Code") THEN
                    ModeShipmentMethod := FORMAT(RecShipmentMethod."GST Reporting Transport Mode");
                //RSPLSUM 07Nov2020<<

                BarcodeForQuarantineLabel(EWayBillNo, LocRegNo, EWBUpdateDate);

                EncodeCode128(FORMAT(EWayBillNo), 2, FALSE, recTempBlob);

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

    var
        recEWB: Record 50044;
        recCompInfo: Record 79;
        SalInvHead1: Record 112;
        Mode: Code[50];
        VehicleTransDoc: Code[50];
        LocName: Code[50];
        Location1: Record 14;
        DetailedsEWayBill1: Record 50044;
        EWBUpdateDate: Text;
        DetaiGSTLedgEnt: Record "Detailed GST Ledger Entry";
        EWayBillNo: Code[20];
        DispatchLocation: Code[20];
        EwabUpdateDate: Text;
        GSTLedgerEntry: Record "GST Ledger Entry";
        DocNo1: Code[20];
        EntryNo: Integer;
        DistanceInKm: Decimal;
        ValidUpto: Text;
        LocGstRegNo: Code[20];
        LocPin: Code[20];
        SalInvHead2: Record 112;
        LocCode2: Code[20];
        Customer1: Record 18;
        GSTINRec: Code[20];
        SIHDate: Date;
        RecHSNSAC: Record "HSN/SAC";
        HSNSACDIS: Text;
        Transporter1: Record 50046;
        TranspoCode: Code[20];
        TranspoGST: Code[20];
        recTempBlob: Record 50044 temporary;
        "------------------": Integer;
        bxtBarcodeBinary: BigText;
        txcErrorSize: Label 'Valid values for the barcode size are 1, 2, 3, 4 & 5';
        txcErrorLength: Label 'Value to encode should be %1 digits.';
        txcErrorNumber: Label 'Only numbers allowed.';
        Barcode2: Integer;
        SalInvLine: Record 113;
        AmtToCut: Decimal;
        DocNo: Code[20];
        EWBNo: Code[50];
        PlaceofDelivery: Text;
        Cust3: Record 18;
        CustName3: Text;
        ValueofGood: Decimal;
        DetaiGSTLedEntry: Record "Detailed GST Ledger Entry";
        HSNCODE: Code[20];
        SalInvHdr: Record 112;
        ReaTran: Option " ",Registered,Unregistered,Export,"Deemed Export",Exempted,"SEZ Development","SEZ Unit";
        ReaTran1: Text;
        LocRegNo: Code[20];
        QRCODE2: Boolean;
        TranspoName: Text;
        GeneratedDate: Text;
        RecState: Record State;
        FromLocState: Text[50];
        FromLocName: Code[50];
        FromAddr: array[3] of Text;
        i: Integer;
        RecCust: Record 18;
        ToGSTIN: Code[20];
        ToLocName: Text[60];
        ToLocState: Text[50];
        ToAddr: array[3] of Text;
        ProdNameDesc: Text[100];
        RecItem: Record 27;
        RecSIL: Record 113;
        TotTaxAmt: Decimal;
        RecDGSTLE: Record "Detailed GST Ledger Entry";
        CGST: Decimal;
        SGST: Decimal;
        IGST: Decimal;
        Cess: Decimal;
        CessNonAdvol: Decimal;
        OtherAmt: Decimal;
        Type: Text[20];
        RecShippingAgent: Record 291;
        RecShipmentMethod: Record 10;
        CGSTPer: Decimal;
        SGSTPer: Decimal;
        IGSTPer: Decimal;
        CessPer: Decimal;
        CessNonAdvolPer: Decimal;
        RecVend: Record 23;
        TransporterIDName: Text[120];
        RecDetGSTLE: Record "Detailed GST Ledger Entry";
        RecUnitOfMeasure: Record 204;
        GSTReportingQOC: Code[10];
        VehFrom: Text[30];
        RecShipToAddr: Record 222;
        ToLocStateBillTo: Text[50];
        RecDetEwayBillNew: Record 50044;
        RecDGSTLENew: Record "Detailed GST Ledger Entry";
        EWBUpdatedDate: Text;
        RecSalesInvLine: Record 113;
        TaxableAmt: Decimal;
        TotInvAmt: Decimal;
        ModeShipmentMethod: Text[10];
        recDispatchFromLocation: Record 50042;

    // //[Scope('Internal')]
    procedure PrintQRCode(DocNoC: Code[30]; LineNoC: Integer; CustomerOrderNo: Code[10]; ItemNo: Text; Qty: Decimal; DocumentNo: Code[10]; PostingDate: Text; GrossRate: Text; NetRate: Text; VendorCode: Code[20]; CrossRefNo: Code[20]; GstAmount: Text; TaxRateUGST: Text; GSTPer: Text; UGST: Text; Cess: Text; AmountToCustomer: Text; HSNCode: Code[10])
    begin

        //QRMgt_lCdu.BarcodeForQuarantineLabel(DocNo,LineNo,ItmNo,MRNNo,BatchNo,LocCode,MfgName,Desc);
        //BarcodeForQuarantineLabel(DocNo,LineNo,ItmNo,MRNNo,BatchNo,LocCode,MfgName,Desc);
    end;

    /* local procedure CreateQRCode(QRCodeInput: Text[500]; var TempBLOB: Record 99008535)
    var
        QRCodeFileName: Text[1024];
    begin
        CLEAR(TempBLOB);
        QRCodeFileName := GetQRCode(QRCodeInput);
        UploadFileBLOBImportandDeleteServerFile(TempBLOB, QRCodeFileName);
    end;

    // //[Scope('Internal')]
    procedure UploadFileBLOBImportandDeleteServerFile(var TempBlob: Record 99008535; FileName: Text[1024])
    var
        FileManagement: Codeunit 419;
    begin
        FileName := FileManagement.UploadFileSilent(FileName);
        FileManagement.BLOBImportFromServerFile(TempBlob, FileName);
        DeleteServerFile(FileName);
    end;

    local procedure DeleteServerFile(ServerFileName: Text)
    begin
        IF ERASE(ServerFileName) THEN;
    end;

    local procedure GetQRCode(QRCodeInput: Text[300]) QRCodeFileName: Text[1024]
    var
        [RunOnClient]
        IBarCodeProvider: DotNet IBarcodeProvider;
    begin
        GetBarCodeProvider(IBarCodeProvider);
        QRCodeFileName := IBarCodeProvider.GetBarcode(QRCodeInput);
    end;

    // //[Scope('Internal')]
    procedure GetBarCodeProvider(var IBarCodeProvider: DotNet IBarcodeProvider)
    var
        [RunOnClient]
        QRCodeProvider: DotNet QRCodeProvider;
    begin
        IF ISNULL(IBarCodeProvider) THEN
            IBarCodeProvider := QRCodeProvider.QRCodeProvider;
    end; */

    //  //[Scope('Internal')]
    procedure BarcodeForQuarantineLabel(DocNoC: Code[30]; Location: Code[20]; Date: Text)
    var
        recSIL1: Record 113;
        ItemLedgEntry_lRec: Record 32;
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
        QRTempBlob_lRecTmp: Record 50009 temporary;
        SRNo_lInt: Integer;
        QRCodeInput: Text;
        // TempBlob: Record 99008535;
        R50017_lRpt: Report 50017;
        R50018_lRpt: Report 50018;
    begin
        //AKT_EWB 25888 16012020
        DetaiGSTLedgEnt.RESET;
        DetaiGSTLedgEnt.SETRANGE("Document No.", "Sales Invoice Header"."No.");
        DetaiGSTLedgEnt.SETRANGE("Posting Date", "Sales Invoice Header"."Posting Date");
        IF DetaiGSTLedgEnt.FINDFIRST THEN BEGIN
            QRCodeInput := STRSUBSTNO('%1,%2,%3', EWayBillNo, LocRegNo, EWBUpdateDate);
            //CreateQRCode(QRCodeInput, TempBlob);
            // DetaiGSTLedgEnt."QR Code" := TempBlob.Blob;
            DetaiGSTLedgEnt.MODIFY;
        END

        //NB

        //QRTempBlob_lRecTmp.INSERT;

        //UNTIL ItemLedgEntry_lRec.NEXT = 0;
        //END;


        //QRTempBlob_lRecTmp.RESET;
        /*
        IF Barcode100By75 THEN BEGIN
          CLEAR(R50017_lRpt);
          R50017_lRpt.TransfterDate_gFnc(QRTempBlob_lRecTmp);
          //R50017_lRpt.USEREQUESTPAGE(FALSE);
          R50017_lRpt.RUNMODAL;
        END;
        
        IF Barcode100By15 THEN BEGIN
          CLEAR(R50018_lRpt);
          R50018_lRpt.TransfterDate_gFnc(QRTempBlob_lRecTmp);
          //R50017_lRpt.USEREQUESTPAGE(FALSE);
          R50018_lRpt.RUNMODAL;
        END;
        */

    end;

    //  //[Scope('Internal')]
    procedure EncodeCode128(pcodBarcode: Code[1024]; pintSize: Integer; pblnVertical: Boolean; var precTmpTempBlob: Record 50044 temporary)
    var
        lrecTmpCode128: Record 50044 temporary;
        lintCount1: Integer;
        ltxtArray: array[2, 200] of Text[30];
        lcharCurrentCharSet: Char;
        lintWeightSum: Integer;
        lintCount2: Integer;
        lintConvInt: Integer;
        ltxtTerminationBar: Text[2];
        lintCheckDigit: Integer;
        lintConvInt1: Integer;
        lintConvInt2: Integer;
        lblnnumber: Boolean;
        lintLines: Integer;
        lintBars: Integer;
        loutBmpHeader: OutStream;
    begin
        CLEAR(bxtBarcodeBinary);
        CLEAR(precTmpTempBlob);
        CLEAR(lrecTmpCode128);
        lrecTmpCode128.DELETEALL;
        CLEAR(lcharCurrentCharSet);
        ltxtTerminationBar := '11';

        IF NOT (pintSize IN [1, 2, 3, 4, 5]) THEN
            ERROR(txcErrorSize);

        InitCode128(lrecTmpCode128);

        FOR lintCount1 := 1 TO STRLEN(pcodBarcode) DO BEGIN
            lintCount2 += 1;
            lblnnumber := FALSE;
            lrecTmpCode128.RESET;

            IF EVALUATE(lintConvInt1, FORMAT(pcodBarcode[lintCount1])) THEN
                lblnnumber := EVALUATE(lintConvInt2, FORMAT(pcodBarcode[lintCount1 + 1]));

            //A '.' IS EVALUATED AS A 0, EXTRA CHECK NEEDED
            IF FORMAT(pcodBarcode[lintCount1]) = '.' THEN
                lblnnumber := FALSE;

            IF FORMAT(pcodBarcode[lintCount1 + 1]) = '.' THEN
                lblnnumber := FALSE;

            IF lblnnumber AND (lintConvInt1 IN [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]) AND (lintConvInt2 IN [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]) THEN BEGIN
                IF (lcharCurrentCharSet <> 'C') THEN BEGIN
                    IF (lintCount1 = 1) THEN BEGIN
                        //lrecTmpCode128.GET(DocNo1,EWayBillNo,'STARTC'); //RSPL-AR AR-EWB commented DJ 31/08/2020
                        lrecTmpCode128.RESET;
                        lrecTmpCode128.SETRANGE("Document No.", DocNo1);
                        lrecTmpCode128.SETRANGE("EWB No.", EWayBillNo);
                        lrecTmpCode128.SETRANGE(CharA, 'STARTC');
                        IF lrecTmpCode128.FINDFIRST THEN;

                        EVALUATE(lintConvInt, lrecTmpCode128.Value);    //RSPL-AR AR-EWB
                        lintWeightSum := lintConvInt;
                    END ELSE BEGIN
                        //lrecTmpCode128.GET(DocNo1,EWayBillNo,'CODEC');  //RSPL-AR AR-EWB commented DJ 31/08/2020
                        lrecTmpCode128.RESET;
                        lrecTmpCode128.SETRANGE("Document No.", DocNo1);
                        lrecTmpCode128.SETRANGE("EWB No.", EWayBillNo);
                        lrecTmpCode128.SETRANGE(CharA, 'CODEC');
                        IF lrecTmpCode128.FINDFIRST THEN;

                        EVALUATE(lintConvInt, lrecTmpCode128.Value);     //RSPL-AR AR-EWB
                        lintWeightSum += lintConvInt * lintCount2;
                        lintCount2 += 1;
                    END;

                    bxtBarcodeBinary.ADDTEXT(FORMAT(lrecTmpCode128.Encoding));
                    lcharCurrentCharSet := 'C';
                END;
            END ELSE BEGIN
                IF lcharCurrentCharSet <> 'A' THEN BEGIN
                    IF (lintCount1 = 1) THEN BEGIN
                        //lrecTmpCode128.GET(DocNo1,EWayBillNo,'STARTA'); //RSPL-AR AR-EWB commented DJ 31/08/2020
                        lrecTmpCode128.RESET;
                        lrecTmpCode128.SETRANGE("Document No.", DocNo1);
                        lrecTmpCode128.SETRANGE("EWB No.", EWayBillNo);
                        lrecTmpCode128.SETRANGE(CharA, 'STARTA');
                        IF lrecTmpCode128.FINDFIRST THEN;

                        EVALUATE(lintConvInt, lrecTmpCode128.Value);
                        lintWeightSum := lintConvInt;
                    END ELSE BEGIN
                        //CODEA
                        //lrecTmpCode128.GET(DocNo1,EWayBillNo,'FNC4');   //RSPL-AR AR-EWB commented DJ 31/08/2020
                        lrecTmpCode128.RESET;
                        lrecTmpCode128.SETRANGE("Document No.", DocNo1);
                        lrecTmpCode128.SETRANGE("EWB No.", EWayBillNo);
                        lrecTmpCode128.SETRANGE(CharA, 'FNC4');
                        IF lrecTmpCode128.FINDFIRST THEN;

                        EVALUATE(lintConvInt, lrecTmpCode128.Value);
                        lintWeightSum += lintConvInt * lintCount2;
                        lintCount2 += 1;
                    END;

                    bxtBarcodeBinary.ADDTEXT(FORMAT(lrecTmpCode128.Encoding));
                    lcharCurrentCharSet := 'A';
                END;
            END;

            CASE lcharCurrentCharSet OF
                'A':
                    BEGIN
                        lrecTmpCode128.GET(FORMAT(pcodBarcode[lintCount1]));

                        EVALUATE(lintConvInt, lrecTmpCode128.Value);

                        lintWeightSum += lintConvInt * lintCount2;
                        bxtBarcodeBinary.ADDTEXT(FORMAT(lrecTmpCode128.Encoding));
                    END;
                'C':
                    BEGIN
                        lrecTmpCode128.RESET;
                        lrecTmpCode128.SETCURRENTKEY(Value);
                        lrecTmpCode128.SETRANGE(Value, (FORMAT(pcodBarcode[lintCount1]) + FORMAT(pcodBarcode[lintCount1 + 1])));
                        lrecTmpCode128.FINDFIRST;

                        EVALUATE(lintConvInt, lrecTmpCode128.Value);
                        lintWeightSum += lintConvInt * lintCount2;

                        bxtBarcodeBinary.ADDTEXT(FORMAT(lrecTmpCode128.Encoding));
                        lintCount1 += 1;
                    END;
            END;
        END;

        lintCheckDigit := lintWeightSum MOD 103;

        //ADD CHECK DIGIT
        lrecTmpCode128.RESET;
        lrecTmpCode128.SETCURRENTKEY(Value);

        IF lintCheckDigit <= 9 THEN
            lrecTmpCode128.SETRANGE(Value, '0' + FORMAT(lintCheckDigit))
        ELSE
            lrecTmpCode128.SETRANGE(Value, FORMAT(lintCheckDigit));

        lrecTmpCode128.FINDFIRST;
        bxtBarcodeBinary.ADDTEXT(FORMAT(lrecTmpCode128.Encoding));

        //ADD STOP CHARACTER
        //lrecTmpCode128.GET(DocNo1,EWayBillNo,'STOP'); commented DJ 31/08/2020
        lrecTmpCode128.RESET;
        lrecTmpCode128.SETRANGE("Document No.", DocNo1);
        lrecTmpCode128.SETRANGE("EWB No.", EWayBillNo);
        lrecTmpCode128.SETRANGE(CharA, 'STOP');
        IF lrecTmpCode128.FINDFIRST THEN;

        bxtBarcodeBinary.ADDTEXT(FORMAT(lrecTmpCode128.Encoding));

        //ADD TERMINATION BAR
        bxtBarcodeBinary.ADDTEXT(ltxtTerminationBar);

        lintBars := bxtBarcodeBinary.LENGTH;
        lintLines := ROUND(lintBars * 0.25, 1, '>');

        precTmpTempBlob.Barcode.CREATEOUTSTREAM(loutBmpHeader);

        //WRITING HEADER
        CreateBMPHeader(loutBmpHeader, lintBars, lintLines, pintSize, pblnVertical);

        //WRITE BARCODE DETAIL
        CreateBarcodeDetail(lintLines, lintBars, pintSize, pblnVertical, loutBmpHeader);
    end;

    local procedure InitCode128(var precTmpCode128: Record 50044 temporary)
    begin
        precTmpCode128.INIT;
        precTmpCode128.CharA := ' ';
        precTmpCode128.CharB := ' ';
        precTmpCode128.CharC := ' ';
        precTmpCode128.Value := '00';
        precTmpCode128.Encoding := '11011001100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;

        precTmpCode128."Document No." := DocNo1;    //RSPL-AR - AR_EWB
        precTmpCode128."EWB No." := EWayBillNo;     //RSPL-AR - AR_EWB
        precTmpCode128.CharA := '!';
        precTmpCode128.CharB := '!';
        precTmpCode128.CharC := '01';
        precTmpCode128.Value := '01';
        precTmpCode128.Encoding := '11001101100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128."Document No." := DocNo1;  //RSPL-AR - AR_EWB
        precTmpCode128."EWB No." := EWayBillNo;  //RSPL-AR - AR_EWB

        precTmpCode128.CharA := '"';
        precTmpCode128.CharB := '"';
        precTmpCode128.CharC := '02';
        precTmpCode128.Value := '02';
        precTmpCode128.Encoding := '11001100110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '#';
        precTmpCode128.CharB := '#';
        precTmpCode128.CharC := '03';
        precTmpCode128.Value := '03';
        precTmpCode128.Encoding := '10010011000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '$';
        precTmpCode128.CharB := '$';
        precTmpCode128.CharC := '04';
        precTmpCode128.Value := '04';
        precTmpCode128.Encoding := '10010001100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '%';
        precTmpCode128.CharB := '%';
        precTmpCode128.CharC := '05';
        precTmpCode128.Value := '05';
        precTmpCode128.Encoding := '10001001100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '&';
        precTmpCode128.CharB := '&';
        precTmpCode128.CharC := '06';
        precTmpCode128.Value := '06';
        precTmpCode128.Encoding := '10011001000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '''';
        precTmpCode128.CharB := '''';
        precTmpCode128.CharC := '07';
        precTmpCode128.Value := '07';
        precTmpCode128.Encoding := '10011000100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '(';
        precTmpCode128.CharB := '(';
        precTmpCode128.CharC := '08';
        precTmpCode128.Value := '08';
        precTmpCode128.Encoding := '10001100100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := ')';
        precTmpCode128.CharB := ')';
        precTmpCode128.CharC := '09';
        precTmpCode128.Value := '09';
        precTmpCode128.Encoding := '11001001000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '*';
        precTmpCode128.CharB := '*';
        precTmpCode128.CharC := '10';
        precTmpCode128.Value := '10';
        precTmpCode128.Encoding := '11001000100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '+';
        precTmpCode128.CharB := '+';
        precTmpCode128.CharC := '11';
        precTmpCode128.Value := '11';
        precTmpCode128.Encoding := '11000100100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := ',';
        precTmpCode128.CharB := ',';
        precTmpCode128.CharC := '12';
        precTmpCode128.Value := '12';
        precTmpCode128.Encoding := '10110011100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '-';
        precTmpCode128.CharB := '-';
        precTmpCode128.CharC := '13';
        precTmpCode128.Value := '13';
        precTmpCode128.Encoding := '10011011100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '.';
        precTmpCode128.CharB := '.';
        precTmpCode128.CharC := '14';
        precTmpCode128.Value := '14';
        precTmpCode128.Encoding := '10011001110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '/';
        precTmpCode128.CharB := '/';
        precTmpCode128.CharC := '15';
        precTmpCode128.Value := '15';
        precTmpCode128.Encoding := '10111001100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '0';
        precTmpCode128.CharB := '0';
        precTmpCode128.CharC := '16';
        precTmpCode128.Value := '16';
        precTmpCode128.Encoding := '10011101100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '1';
        precTmpCode128.CharB := '1';
        precTmpCode128.CharC := '17';
        precTmpCode128.Value := '17';
        precTmpCode128.Encoding := '10011100110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '2';
        precTmpCode128.CharB := '2';
        precTmpCode128.CharC := '18';
        precTmpCode128.Value := '18';
        precTmpCode128.Encoding := '11001110010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '3';
        precTmpCode128.CharB := '3';
        precTmpCode128.CharC := '19';
        precTmpCode128.Value := '19';
        precTmpCode128.Encoding := '11001011100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '4';
        precTmpCode128.CharB := '4';
        precTmpCode128.CharC := '20';
        precTmpCode128.Value := '20';
        precTmpCode128.Encoding := '11001001110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '5';
        precTmpCode128.CharB := '5';
        precTmpCode128.CharC := '21';
        precTmpCode128.Value := '21';
        precTmpCode128.Encoding := '11011100100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '6';
        precTmpCode128.CharB := '6';
        precTmpCode128.CharC := '22';
        precTmpCode128.Value := '22';
        precTmpCode128.Encoding := '11001110100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '7';
        precTmpCode128.CharB := '7';
        precTmpCode128.CharC := '23';
        precTmpCode128.Value := '23';
        precTmpCode128.Encoding := '11101101110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '8';
        precTmpCode128.CharB := '8';
        precTmpCode128.CharC := '24';
        precTmpCode128.Value := '24';
        precTmpCode128.Encoding := '11101001100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '9';
        precTmpCode128.CharB := '9';
        precTmpCode128.CharC := '25';
        precTmpCode128.Value := '25';
        precTmpCode128.Encoding := '11100101100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := ':';
        precTmpCode128.CharB := ':';
        precTmpCode128.CharC := '26';
        precTmpCode128.Value := '26';
        precTmpCode128.Encoding := '11100100110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := ';';
        precTmpCode128.CharB := ';';
        precTmpCode128.CharC := '27';
        precTmpCode128.Value := '27';
        precTmpCode128.Encoding := '11101100100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '<';
        precTmpCode128.CharB := '<';
        precTmpCode128.CharC := '28';
        precTmpCode128.Value := '28';
        precTmpCode128.Encoding := '11100110100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '=';
        precTmpCode128.CharB := '=';
        precTmpCode128.CharC := '29';
        precTmpCode128.Value := '29';
        precTmpCode128.Encoding := '11100110010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '>';
        precTmpCode128.CharB := '>';
        precTmpCode128.CharC := '30';
        precTmpCode128.Value := '30';
        precTmpCode128.Encoding := '11011011000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '?';
        precTmpCode128.CharB := '?';
        precTmpCode128.CharC := '31';
        precTmpCode128.Value := '31';
        precTmpCode128.Encoding := '11011000110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '@';
        precTmpCode128.CharB := '@';
        precTmpCode128.CharC := '32';
        precTmpCode128.Value := '32';
        precTmpCode128.Encoding := '11000110110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'A';
        precTmpCode128.CharB := 'A';
        precTmpCode128.CharC := '33';
        precTmpCode128.Value := '33';
        precTmpCode128.Encoding := '10100011000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'B';
        precTmpCode128.CharB := 'B';
        precTmpCode128.CharC := '34';
        precTmpCode128.Value := '34';
        precTmpCode128.Encoding := '10001011000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'C';
        precTmpCode128.CharB := 'C';
        precTmpCode128.CharC := '35';
        precTmpCode128.Value := '35';
        precTmpCode128.Encoding := '10001000110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'D';
        precTmpCode128.CharB := 'D';
        precTmpCode128.CharC := '36';
        precTmpCode128.Value := '36';
        precTmpCode128.Encoding := '10110001000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'E';
        precTmpCode128.CharB := 'E';
        precTmpCode128.CharC := '37';
        precTmpCode128.Value := '37';
        precTmpCode128.Encoding := '10001101000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'F';
        precTmpCode128.CharB := 'F';
        precTmpCode128.CharC := '38';
        precTmpCode128.Value := '38';
        precTmpCode128.Encoding := '10001100010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'G';
        precTmpCode128.CharB := 'G';
        precTmpCode128.CharC := '39';
        precTmpCode128.Value := '39';
        precTmpCode128.Encoding := '11010001000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'H';
        precTmpCode128.CharB := 'H';
        precTmpCode128.CharC := '40';
        precTmpCode128.Value := '40';
        precTmpCode128.Encoding := '11000101000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'I';
        precTmpCode128.CharB := 'I';
        precTmpCode128.CharC := '41';
        precTmpCode128.Value := '41';
        precTmpCode128.Encoding := '11000100010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'J';
        precTmpCode128.CharB := 'J';
        precTmpCode128.CharC := '42';
        precTmpCode128.Value := '42';
        precTmpCode128.Encoding := '10110111000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'K';
        precTmpCode128.CharB := 'K';
        precTmpCode128.CharC := '43';
        precTmpCode128.Value := '43';
        precTmpCode128.Encoding := '10110001110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'L';
        precTmpCode128.CharB := 'L';
        precTmpCode128.CharC := '44';
        precTmpCode128.Value := '44';
        precTmpCode128.Encoding := '10001101110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'M';
        precTmpCode128.CharB := 'M';
        precTmpCode128.CharC := '45';
        precTmpCode128.Value := '45';
        precTmpCode128.Encoding := '10111011000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'N';
        precTmpCode128.CharB := 'N';
        precTmpCode128.CharC := '46';
        precTmpCode128.Value := '46';
        precTmpCode128.Encoding := '10111000110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'O';
        precTmpCode128.CharB := 'O';
        precTmpCode128.CharC := '47';
        precTmpCode128.Value := '47';
        precTmpCode128.Encoding := '10001110110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'P';
        precTmpCode128.CharB := 'P';
        precTmpCode128.CharC := '48';
        precTmpCode128.Value := '48';
        precTmpCode128.Encoding := '11101110110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'Q';
        precTmpCode128.CharB := 'Q';
        precTmpCode128.CharC := '49';
        precTmpCode128.Value := '49';
        precTmpCode128.Encoding := '11010001110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'R';
        precTmpCode128.CharB := 'R';
        precTmpCode128.CharC := '50';
        precTmpCode128.Value := '50';
        precTmpCode128.Encoding := '11000101110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'S';
        precTmpCode128.CharB := 'S';
        precTmpCode128.CharC := '51';
        precTmpCode128.Value := '51';
        precTmpCode128.Encoding := '11011101000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'T';
        precTmpCode128.CharB := 'T';
        precTmpCode128.CharC := '52';
        precTmpCode128.Value := '52';
        precTmpCode128.Encoding := '11011100010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'U';
        precTmpCode128.CharB := 'U';
        precTmpCode128.CharC := '53';
        precTmpCode128.Value := '53';
        precTmpCode128.Encoding := '11011101110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'V';
        precTmpCode128.CharB := 'V';
        precTmpCode128.CharC := '54';
        precTmpCode128.Value := '54';
        precTmpCode128.Encoding := '11101011000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'W';
        precTmpCode128.CharB := 'W';
        precTmpCode128.CharC := '55';
        precTmpCode128.Value := '55';
        precTmpCode128.Encoding := '11101000110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'X';
        precTmpCode128.CharB := 'X';
        precTmpCode128.CharC := '56';
        precTmpCode128.Value := '56';
        precTmpCode128.Encoding := '11100010110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'Y';
        precTmpCode128.CharB := 'Y';
        precTmpCode128.CharC := '57';
        precTmpCode128.Value := '57';
        precTmpCode128.Encoding := '11101101000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'Z';
        precTmpCode128.CharB := 'Z';
        precTmpCode128.CharC := '58';
        precTmpCode128.Value := '58';
        precTmpCode128.Encoding := '11101100010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '[';
        precTmpCode128.CharB := '[';
        precTmpCode128.CharC := '59';
        precTmpCode128.Value := '59';
        precTmpCode128.Encoding := '11100011010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '\';
        precTmpCode128.CharB := '\';
        precTmpCode128.CharC := '60';
        precTmpCode128.Value := '60';
        precTmpCode128.Encoding := '11101111010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := ']';
        precTmpCode128.CharB := ']';
        precTmpCode128.CharC := '61';
        precTmpCode128.Value := '61';
        precTmpCode128.Encoding := '11001000010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '^';
        precTmpCode128.CharB := '^';
        precTmpCode128.CharC := '62';
        precTmpCode128.Value := '62';
        precTmpCode128.Encoding := '11110001010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := '_';
        precTmpCode128.CharB := '_';
        precTmpCode128.CharC := '63';
        precTmpCode128.Value := '63';
        precTmpCode128.Encoding := '10100110000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'NUL';
        precTmpCode128.CharB := '`';
        precTmpCode128.CharC := '64';
        precTmpCode128.Value := '64';
        precTmpCode128.Encoding := '10100001100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'SOH';
        precTmpCode128.CharB := 'a';
        precTmpCode128.CharC := '65';
        precTmpCode128.Value := '65';
        precTmpCode128.Encoding := '10010110000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'STX';
        precTmpCode128.CharB := 'b';
        precTmpCode128.CharC := '66';
        precTmpCode128.Value := '66';
        precTmpCode128.Encoding := '10010000110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'ETX';
        precTmpCode128.CharB := 'c';
        precTmpCode128.CharC := '67';
        precTmpCode128.Value := '67';
        precTmpCode128.Encoding := '10000101100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'EOT';
        precTmpCode128.CharB := 'd';
        precTmpCode128.CharC := '68';
        precTmpCode128.Value := '68';
        precTmpCode128.Encoding := '10000100110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'ENQ';
        precTmpCode128.CharB := 'e';
        precTmpCode128.CharC := '69';
        precTmpCode128.Value := '69';
        precTmpCode128.Encoding := '10110010000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'ACK';
        precTmpCode128.CharB := 'f';
        precTmpCode128.CharC := '70';
        precTmpCode128.Value := '70';
        precTmpCode128.Encoding := '10110000100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'BEL';
        precTmpCode128.CharB := 'g';
        precTmpCode128.CharC := '71';
        precTmpCode128.Value := '71';
        precTmpCode128.Encoding := '10011010000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'BS';
        precTmpCode128.CharB := 'h';
        precTmpCode128.CharC := '72';
        precTmpCode128.Value := '72';
        precTmpCode128.Encoding := '10011000010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'HT';
        precTmpCode128.CharB := 'i';
        precTmpCode128.CharC := '73';
        precTmpCode128.Value := '73';
        precTmpCode128.Encoding := '10000110100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'LF';
        precTmpCode128.CharB := 'j';
        precTmpCode128.CharC := '74';
        precTmpCode128.Value := '74';
        precTmpCode128.Encoding := '10000110010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'VT';
        precTmpCode128.CharB := 'k';
        precTmpCode128.CharC := '75';
        precTmpCode128.Value := '75';
        precTmpCode128.Encoding := '11000010010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'FF';
        precTmpCode128.CharB := 'l';
        precTmpCode128.CharC := '76';
        precTmpCode128.Value := '76';
        precTmpCode128.Encoding := '11001010000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'CR';
        precTmpCode128.CharB := 'm';
        precTmpCode128.CharC := '77';
        precTmpCode128.Value := '77';
        precTmpCode128.Encoding := '11110111010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'SO';
        precTmpCode128.CharB := 'n';
        precTmpCode128.CharC := '78';
        precTmpCode128.Value := '78';
        precTmpCode128.Encoding := '11000010100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'SI';
        precTmpCode128.CharB := 'o';
        precTmpCode128.CharC := '79';
        precTmpCode128.Value := '79';
        precTmpCode128.Encoding := '10001111010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'DLE';
        precTmpCode128.CharB := 'p';
        precTmpCode128.CharC := '80';
        precTmpCode128.Value := '80';
        precTmpCode128.Encoding := '10100111100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'DC1';
        precTmpCode128.CharB := 'q';
        precTmpCode128.CharC := '81';
        precTmpCode128.Value := '81';
        precTmpCode128.Encoding := '10010111100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'DC2';
        precTmpCode128.CharB := 'r';
        precTmpCode128.CharC := '82';
        precTmpCode128.Value := '82';
        precTmpCode128.Encoding := '10010011110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'DC3';
        precTmpCode128.CharB := 's';
        precTmpCode128.CharC := '83';
        precTmpCode128.Value := '83';
        precTmpCode128.Encoding := '10111100100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'DC4';
        precTmpCode128.CharB := 't';
        precTmpCode128.CharC := '84';
        precTmpCode128.Value := '84';
        precTmpCode128.Encoding := '10011110100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'NAK';
        precTmpCode128.CharB := 'u';
        precTmpCode128.CharC := '85';
        precTmpCode128.Value := '85';
        precTmpCode128.Encoding := '10011110010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'SYN';
        precTmpCode128.CharB := 'v';
        precTmpCode128.CharC := '86';
        precTmpCode128.Value := '86';
        precTmpCode128.Encoding := '11110100100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'ETB';
        precTmpCode128.CharB := 'w';
        precTmpCode128.CharC := '87';
        precTmpCode128.Value := '87';
        precTmpCode128.Encoding := '11110010100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'CAN';
        precTmpCode128.CharB := 'x';
        precTmpCode128.CharC := '88';
        precTmpCode128.Value := '88';
        precTmpCode128.Encoding := '11110010010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'EM';
        precTmpCode128.CharB := 'y';
        precTmpCode128.CharC := '89';
        precTmpCode128.Value := '89';
        precTmpCode128.Encoding := '11011011110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'SUB';
        precTmpCode128.CharB := 'z';
        precTmpCode128.CharC := '90';
        precTmpCode128.Value := '90';
        precTmpCode128.Encoding := '11011110110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'ESC';
        precTmpCode128.CharB := '{';
        precTmpCode128.CharC := '91';
        precTmpCode128.Value := '91';
        precTmpCode128.Encoding := '11110110110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'FS';
        precTmpCode128.CharB := '|';
        precTmpCode128.CharC := '92';
        precTmpCode128.Value := '92';
        precTmpCode128.Encoding := '10101111000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'GS';
        precTmpCode128.CharB := '}';
        precTmpCode128.CharC := '93';
        precTmpCode128.Value := '93';
        precTmpCode128.Encoding := '10100011110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'RS';
        precTmpCode128.CharB := '~';
        precTmpCode128.CharC := '94';
        precTmpCode128.Value := '94';
        precTmpCode128.Encoding := '10001011110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'US';
        precTmpCode128.CharB := 'DEL';
        precTmpCode128.CharC := '95';
        precTmpCode128.Value := '95';
        precTmpCode128.Encoding := '10111101000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'FNC3';
        precTmpCode128.CharB := 'FNC3';
        precTmpCode128.CharC := '96';
        precTmpCode128.Value := '96';
        precTmpCode128.Encoding := '10111100010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'FNC2';
        precTmpCode128.CharB := 'FNC2';
        precTmpCode128.CharC := '97';
        precTmpCode128.Value := '97';
        precTmpCode128.Encoding := '11110101000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'SHIFT';
        precTmpCode128.CharB := 'SHIFT';
        precTmpCode128.CharC := '98';
        precTmpCode128.Value := '98';
        precTmpCode128.Encoding := '11110100010';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'CODEC';
        precTmpCode128.CharB := 'CODEC';
        precTmpCode128.CharC := '99';
        precTmpCode128.Value := '99';
        precTmpCode128.Encoding := '10111011110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'CODEB';
        precTmpCode128.CharB := 'FNC4';
        precTmpCode128.CharC := 'CODEB';
        precTmpCode128.Value := '100';
        precTmpCode128.Encoding := '10111101110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'FNC4';
        precTmpCode128.CharB := 'CODEA';
        precTmpCode128.CharC := 'CODEA';
        precTmpCode128.Value := '101';
        precTmpCode128.Encoding := '11101011110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'FNC1';
        precTmpCode128.CharB := 'FNC1';
        precTmpCode128.CharC := 'FNC1';
        precTmpCode128.Value := '102';
        precTmpCode128.Encoding := '11110101110';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'STARTA';
        precTmpCode128.CharB := 'STARTA';
        precTmpCode128.CharC := 'STARTA';
        precTmpCode128.Value := '103';
        precTmpCode128.Encoding := '11010000100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'STARTB';
        precTmpCode128.CharB := 'STARTB';
        precTmpCode128.CharC := 'STARTB';
        precTmpCode128.Value := '104';
        precTmpCode128.Encoding := '11010010000';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'STARTC';
        precTmpCode128.CharB := 'STARTC';
        precTmpCode128.CharC := 'STARTC';
        precTmpCode128.Value := '105';
        precTmpCode128.Encoding := '11010011100';
        precTmpCode128.INSERT;

        precTmpCode128.INIT;
        precTmpCode128.CharA := 'STOP';
        precTmpCode128.CharB := 'STOP';
        precTmpCode128.CharC := 'STOP';
        precTmpCode128.Value := '';
        precTmpCode128.Encoding := '11000111010';
        precTmpCode128.INSERT;
    end;

    local procedure CreateBMPHeader(var poutBmpHeader: OutStream; pintCols: Integer; pintRows: Integer; pintSize: Integer; pblnVertical: Boolean)
    var
        charInf: Char;
        lintResolution: Integer;
        lintWidth: Integer;
        lintHeight: Integer;
    begin

        lintResolution := ROUND(2835 / pintSize, 1, '=');

        IF pblnVertical THEN BEGIN
            lintWidth := pintRows * pintSize;
            lintHeight := pintCols;
        END ELSE BEGIN
            lintWidth := pintCols * pintSize;
            lintHeight := pintRows * pintSize;
        END;

        charInf := 'B';
        poutBmpHeader.WRITE(charInf, 1);
        charInf := 'M';
        poutBmpHeader.WRITE(charInf, 1);
        poutBmpHeader.WRITE(54 + pintRows * pintCols * 3, 4); //SIZE BMP
        poutBmpHeader.WRITE(0, 4); //APPLICATION SPECIFIC
        poutBmpHeader.WRITE(54, 4); //OFFSET DATA PIXELS
        poutBmpHeader.WRITE(40, 4); //NUMBER OF BYTES IN HEADER FROM THIS POINT
        poutBmpHeader.WRITE(lintWidth, 4); //WIDTH PIXEL
        poutBmpHeader.WRITE(lintHeight, 4); //HEIGHT PIXEL
        poutBmpHeader.WRITE(65536 * 24 + 1, 4); //COLOR DEPTH
        poutBmpHeader.WRITE(0, 4); //NO. OF COLOR PANES & BITS PER PIXEL
        poutBmpHeader.WRITE(0, 4); //SIZE BMP DATA
        poutBmpHeader.WRITE(lintResolution, 4); //HORIZONTAL RESOLUTION
        poutBmpHeader.WRITE(lintResolution, 4); //VERTICAL RESOLUTION
        poutBmpHeader.WRITE(0, 4); //NO. OF COLORS IN PALETTE
        poutBmpHeader.WRITE(0, 4); //IMPORTANT COLORS
    end;

    local procedure CreateBarcodeDetail(pintLines: Integer; pintBars: Integer; pintSize: Integer; pblnVertical: Boolean; var poutBmpHeader: OutStream)
    var
        lintLineLoop: Integer;
        lintBarLoop: Integer;
        ltxtByte: Text[1];
        lchar: Char;
        lintChainFiller: Integer;
        lintSize: Integer;
        lintCount: Integer;
    begin
        IF pblnVertical THEN BEGIN
            FOR lintBarLoop := 1 TO (bxtBarcodeBinary.LENGTH) DO BEGIN

                FOR lintLineLoop := 1 TO (pintLines * pintSize) DO BEGIN
                    bxtBarcodeBinary.GETSUBTEXT(ltxtByte, lintBarLoop, 1);

                    IF ltxtByte = '1' THEN
                        lchar := 0
                    ELSE
                        lchar := 255;

                    poutBmpHeader.WRITE(lchar, 1);
                    poutBmpHeader.WRITE(lchar, 1);
                    poutBmpHeader.WRITE(lchar, 1);
                END;

                FOR lintChainFiller := 1 TO (lintLineLoop MOD 4) DO BEGIN
                    //Adding 0 bytes if needed - line end
                    lchar := 0;
                    poutBmpHeader.WRITE(lchar, 1);
                END;
            END;
        END ELSE BEGIN
            FOR lintLineLoop := 1 TO pintLines * pintSize DO BEGIN
                FOR lintBarLoop := 1 TO bxtBarcodeBinary.LENGTH DO BEGIN
                    bxtBarcodeBinary.GETSUBTEXT(ltxtByte, lintBarLoop, 1);

                    IF ltxtByte = '1' THEN
                        lchar := 0
                    ELSE
                        lchar := 255;

                    FOR lintSize := 1 TO pintSize DO BEGIN
                        //Putting Pixel: Black or White
                        poutBmpHeader.WRITE(lchar, 1);
                        poutBmpHeader.WRITE(lchar, 1);
                        poutBmpHeader.WRITE(lchar, 1);
                    END
                END;

                FOR lintChainFiller := 1 TO ((lintBarLoop * pintSize) MOD 4) DO BEGIN
                    //Adding 0 bytes if needed - line end
                    lchar := 0;
                    poutBmpHeader.WRITE(lchar, 1);
                END;
            END;
        END
    end;
}


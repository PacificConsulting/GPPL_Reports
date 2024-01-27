report 50065 "Vehicle Entry - Ow/Iw Report"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 08Mar2018   RB-N         New code Added for VehicleTakenDateTime not blank
    // 02May2019   RB-N         GRN No. & GRN Date
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/VehicleEntryOwIwReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Posted Gate Entry Header"; "Posted Gate Entry Header")
        {
            DataItemTableView = SORTING("Entry Type", "No.")
                                ORDER(Ascending)
                                WHERE("Vehicle Out" = CONST(true));
            RequestFilterFields = "Posting Date";
            dataitem("Posted Gate Entry Line"; "Posted Gate Entry Line")
            {
                DataItemLink = "Gate Entry No." = FIELD("No.");
                DataItemTableView = SORTING("Entry Type", "Gate Entry No.", "Line No.")
                                    ORDER(Ascending);

                trigger OnAfterGetRecord()
                begin
                    FinalVehicleCapacity := '';//DJ 22012021
                    LCOUNT += 1;

                    //>>02May2019
                    CLEAR(GRNNo);
                    CLEAR(GRNDate);
                    PRH.RESET;
                    PRH.SETCURRENTKEY("Order No.");
                    PRH.SETRANGE("Order No.", "Posted Gate Entry Line"."Source No.");
                    IF PRH.FINDFIRST THEN BEGIN
                        GRNNo := PRH."No.";
                        GRNDate := PRH."Posting Date";
                    END;
                    //<<02May2019

                    //>>1
                    CLEAR(QtyLoaded);

                    // IF "Source Type" = "Source Type"::"Sales Invoice" THEN BEGIN
                    recSalesShpmt.RESET;
                    recSalesShpmt.SETRANGE("Document No.", "Posted Gate Entry Line"."Source No.");
                    IF recSalesShpmt.FINDSET THEN BEGIN
                        ForLoc := recSalesShpmt."Location Code";
                        recCust.GET(recSalesShpmt."Sell-to Customer No.");
                        PartyName := recCust.Name;
                        BillNo := recSalesShpmt."Document No.";
                        DocPostingDate := recSalesShpmt."Posting Date";
                        salesinvheader.GET(recSalesShpmt."Document No.");
                        Roadpermitno := salesinvheader."Road Permit No.";
                        InvPostTime := salesinvheader."Invoice Print Time";
                        FinalVehicleCapacity := salesinvheader."Vehicle Capacity"; //DJ 22/01/2021
                                                                                   //RSPLSUM 27Jan21>>
                        CLEAR(CityName);
                        IF salesinvheader."Ship-to Code" <> '' THEN BEGIN
                            RecShipToAdd.RESET;
                            IF RecShipToAdd.GET(salesinvheader."Sell-to Customer No.", salesinvheader."Ship-to Code") THEN
                                CityName := RecShipToAdd.City;
                        END ELSE BEGIN
                            RecCustomer.RESET;
                            IF RecCustomer.GET(salesinvheader."Bill-to Customer No.") THEN
                                CityName := RecCustomer.City;
                        END;
                        //RSPLSUM 27Jan21<<
                        REPEAT
                            QtyLoaded += recSalesShpmt."Quantity (Base)";
                            qtytot += recSalesShpmt."Quantity (Base)";
                        UNTIL recSalesShpmt.NEXT = 0;
                    END;
                    //  END
                    //  ELSE
                    IF "Source Type" = "Source Type"::"Transfer Shipment" THEN BEGIN
                        recTransShpmt.RESET;
                        recTransShpmt.SETRANGE("Document No.", "Posted Gate Entry Line"."Source No.");
                        IF recTransShpmt.FINDSET THEN BEGIN
                            ForLoc := recTransShpmt."Transfer-from Code";
                            recLoc.GET(recTransShpmt."Transfer-to Code");
                            PartyName := recLoc.Name;
                            BillNo := recTransShpmt."Document No.";
                            DocPostingDate := recTransShpmt."Shipment Date";
                            transfershipheader.GET(recTransShpmt."Document No.");
                            Roadpermitno := transfershipheader."Road Permit No.";
                            FinalVehicleCapacity := transfershipheader."Vehicle Capacity"; //DJ 22/01/2021
                                                                                           //RSPLSUM 27Jan21>>
                            CLEAR(CityName);
                            RecLocation.RESET;
                            //RSPLSUM02Apr21--IF RecLocation.GET(transfershipheader."Transfer-from Code") THEN
                            IF RecLocation.GET(transfershipheader."Transfer-to Code") THEN//RSPLSUM02Apr21
                                                                                          //RSPLSUM15Mar21--CityName := RecLocation.City;
                                CityName := RecLocation.Name;//RSPLSUM15Mar21
                                                             //RSPLSUM 27Jan21<<
                            REPEAT
                                QtyLoaded += recTransShpmt."Quantity (Base)";
                                qtytot += recTransShpmt."Quantity (Base)";
                            UNTIL recTransShpmt.NEXT = 0;
                        END;
                    END;

                    CLEAR(FromLocaName);
                    recLoc.RESET;
                    recLoc.SETRANGE(recLoc.Code, ForLoc);
                    IF recLoc.FINDFIRST THEN
                        FromLocaName := recLoc.Name
                    ELSE
                        FromLocaName := '';

                    IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Outward THEN BEGIN//RSPLSUM02Apr21
                        QtyLoaded := "Posted Gate Entry Header"."Net Weight";//RSPLSUM02Apr21
                        qtytot := "Posted Gate Entry Header"."Net Weight";//RSPLSUM02Apr21
                    END;//RSPLSUM02Apr21

                    CLEAR(Utilised);
                    Vehiclecapacityrec.RESET;
                    Vehiclecapacityrec.SETRANGE(Vehiclecapacityrec.Code, "Posted Gate Entry Header"."Vehicle Capacity");
                    IF Vehiclecapacityrec.FINDFIRST THEN BEGIN
                        IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Inward THEN
                            Utilised := ROUND(100 - ((Vehiclecapacityrec.Value - inwqty) / Vehiclecapacityrec.Value) * 100, 0.01)
                        ELSE
                            Utilised := ROUND(100 - ((Vehiclecapacityrec.Value - QtyLoaded) / Vehiclecapacityrec.Value) * 100, 0.01);
                    END;

                    CLEAR(Utilisedtot);
                    Vehiclecapacityrec.RESET;
                    Vehiclecapacityrec.SETRANGE(Vehiclecapacityrec.Code, "Posted Gate Entry Header"."Vehicle Capacity");
                    IF Vehiclecapacityrec.FINDFIRST THEN BEGIN
                        IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Inward THEN
                            Utilisedtot := ROUND(100 - ((Vehiclecapacityrec.Value - inwqty) / Vehiclecapacityrec.Value) * 100, 0.01)
                        ELSE
                            Utilisedtot := ROUND(100 - ((Vehiclecapacityrec.Value - qtytot) / Vehiclecapacityrec.Value) * 100, 0.01);
                    END;

                    //<<1

                    //>>21Mar2017 Report_Body
                    IF NOT PrintSummry THEN BEGIN

                        outime := CREATEDATETIME("Posted Gate Entry Header"."Out Date", "Posted Gate Entry Header"."Out Time");

                        IF ("Posted Gate Entry Header"."Vehicle Taken Date" <> 0D) AND ("Posted Gate Entry Header"."Vehicle Taken Time" <> 0T) THEN //08Mar2018
                            inttime := CREATEDATETIME("Posted Gate Entry Header"."Vehicle Taken Date", "Posted Gate Entry Header"."Vehicle Taken Time");

                        IF (outime <> 0DT) AND (inttime <> 0DT) THEN//RSPLSUM 11Jan21
                            InFactoryTime := outime - inttime;

                        EnterCell(NN, 1, "Posted Gate Entry Header"."No.", FALSE, FALSE, '', 1);
                        EnterCell(NN, 2, "Posted Gate Entry Header"."Vehicle No.", FALSE, FALSE, '', 1);
                        EnterCell(NN, 3, "Posted Gate Entry Header"."Transporter Name", FALSE, FALSE, '', 1);
                        EnterCell(NN, 4, FromLocaName, FALSE, FALSE, '', 1);
                        //IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type"::"Sales Invoice" THEN
                        EnterCell(NN, 5, 'Invoice', FALSE, FALSE, '', 1);
                        // ELSE
                        IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type"::"Transfer Shipment" THEN
                            EnterCell(NN, 5, 'Transfer', FALSE, FALSE, '', 1);

                        EnterCell(NN, 6, FORMAT("Posted Gate Entry Header"."Document Date"), FALSE, FALSE, '', 2);
                        //EnterCell(NN,7,FORMAT("Posted Gate Entry Header"."Document Time",0,'<Hours12,2>:<Minutes,2>:<Seconds,2> <AM/PM>'),FALSE,FALSE,'',1);
                        EnterCell(NN, 7, FORMAT("Posted Gate Entry Header"."Document Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        //hh:mm:ss AM/PM
                        EnterCell(NN, 8, FORMAT("Posted Gate Entry Header"."Releasing Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 9, FORMAT("Posted Gate Entry Header"."Releasing Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 10, FORMAT("Posted Gate Entry Header"."Vehicle Taken Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 11, FORMAT("Posted Gate Entry Header"."Vehicle Taken Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 12, FORMAT("Posted Gate Entry Header"."Posting Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 13, FORMAT("Posted Gate Entry Header"."Posting Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 14, FORMAT("Posted Gate Entry Header"."Out Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 15, FORMAT("Posted Gate Entry Header"."Out Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 16, PartyName, FALSE, FALSE, '', 1);
                        EnterCell(NN, 17, "Posted Gate Entry Header"."Vehicle Capacity", FALSE, FALSE, '', 1);
                        //RSPLSUM02Apr21--IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Inward THEN
                        //RSPLSUM02Apr21--EnterCell(NN,18,FORMAT(inwqty),FALSE,FALSE,'0.000',0)
                        //RSPLSUM02Apr21--ELSE
                        //RSPLSUM02Apr21--EnterCell(NN,18,FORMAT(QtyLoaded),FALSE,FALSE,'0.000',0);

                        //EnterCell(NN,19,FORMAT(ROUND(Utilised,0.0))+'%',FALSE,FALSE,'',1);
                        IF Utilised = 0 THEN
                            EnterCell(NN, 18, FORMAT(Utilised), FALSE, FALSE, '0.00 %', 0)
                        ELSE
                            EnterCell(NN, 18, FORMAT(Utilised / 100), FALSE, FALSE, '0.00 %', 0);

                        EnterCell(NN, 19, "Posted Gate Entry Header"."LR/RR No.", FALSE, FALSE, '', 1);
                        EnterCell(NN, 20, FORMAT("Posted Gate Entry Header"."LR/RR Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 21, BillNo, FALSE, FALSE, '', 1);
                        EnterCell(NN, 22, FORMAT(DocPostingDate), FALSE, FALSE, '', 2);
                        EnterCell(NN, 23, FORMAT(Roadpermitno), FALSE, FALSE, '', 1);
                        IF InvPostTime <> 0T THEN
                            EnterCell(NN, 24, FORMAT(InvPostTime), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);

                        EnterCell(NN, 25, FORMAT(InFactoryTime), FALSE, FALSE, '', 1);
                        EnterCell(NN, 26, comment1, FALSE, FALSE, '', 1);
                        EnterCell(NN, 27, FORMAT("Posted Gate Entry Header"."Entry Type"), FALSE, FALSE, '', 1);
                        EnterCell(NN, 28, FORMAT("Posted Gate Entry Header"."Vehicle Body Height/Length"), FALSE, FALSE, '', 1);
                        EnterCell(NN, 29, GRNNo, FALSE, FALSE, '', 1);//02May2019
                        EnterCell(NN, 30, FORMAT(GRNDate), FALSE, FALSE, '', 2);//02May2019
                        EnterCell(NN, 31, FORMAT("Posted Gate Entry Header"."Gross Weight"), FALSE, FALSE, '0.00', 0);//RSPLSUM 28Dec2020
                        EnterCell(NN, 32, FORMAT("Posted Gate Entry Header"."Tare Weight"), FALSE, FALSE, '0.00', 0);//RSPLSUM 28Dec2020
                        EnterCell(NN, 33, FORMAT("Posted Gate Entry Header"."Net Weight"), FALSE, FALSE, '0.00', 0);//RSPLSUM 28Dec2020
                        EnterCell(NN, 34, FinalVehicleCapacity, FALSE, FALSE, '', 1);//DJ 22012021
                        EnterCell(NN, 35, CityName, FALSE, FALSE, '', 1);//RSPLSUM 27Jan21
                        EnterCell(NN, 36, FORMAT("Posted Gate Entry Header"."Type of Vehicle"), FALSE, FALSE, '', 1);//RSPLSUM 04Mar21
                        NN += 1;
                    END;
                    //<<21Mar2017 Report_Body

                    /*
                    //Posted Gate Entry Line, Body (1) - OnPreSection()
                    IF NOT PrintSummry THEN
                    BEGIN
                    //CreateFooter;
                      //  SrNo += 1;
                        RowNo += 1;
                        outime :=  CREATEDATETIME("Posted Gate Entry Header"."Out Date","Posted Gate Entry Header"."Out Time");
                        inttime := CREATEDATETIME("Posted Gate Entry Header"."Vehicle Taken Date","Posted Gate Entry Header"."Vehicle Taken Time");
                        InFactoryTime := outime - inttime;
                    
                        xlSheet.Range('A6:T6').Font.Bold := TRUE;
                       // xlSheet.Range(UpdateColumn(1)+FORMAT(RowNo)).Font.Bold := TRUE;
                        xlSheet.Range(UpdateColumn(1)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."No.";
                        xlSheet.Range(UpdateColumn(2)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Vehicle No." ;
                        xlSheet.Range(UpdateColumn(3)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Transporter Name";
                        xlSheet.Range(UpdateColumn(4)+FORMAT(RowNo)).Value :=  FromLocaName;
                    
                        IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type" :: "Sales Invoice" THEN
                          xlSheet.Range(UpdateColumn(5)+FORMAT(RowNo)).Value :=  'Invoice'
                        ELSE
                        IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type" :: "Transfer Shipment"  THEN
                          xlSheet.Range(UpdateColumn(5)+FORMAT(RowNo)).Value :=  'Transfer';
                        xlSheet.Range(UpdateColumn(6)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Document Date") ;
                        xlSheet.Range(UpdateColumn(7)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Document Time") ;
                        xlSheet.Range(UpdateColumn(8)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Releasing Date") ;
                        xlSheet.Range(UpdateColumn(9)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Releasing Time") ;
                        xlSheet.Range(UpdateColumn(10)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Vehicle Taken Date") ;
                        xlSheet.Range(UpdateColumn(11)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Vehicle Taken Time") ;
                        xlSheet.Range(UpdateColumn(12)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Posting Date") ;
                        xlSheet.Range(UpdateColumn(13)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Posting Time") ;
                        xlSheet.Range(UpdateColumn(14)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Out Date") ;
                        xlSheet.Range(UpdateColumn(15)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Out Time") ;
                    
                        xlSheet.Range(UpdateColumn(16)+FORMAT(RowNo)).Value :=  PartyName ;
                        xlSheet.Range(UpdateColumn(17)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Vehicle Capacity";
                        IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Inward THEN
                        xlSheet.Range(UpdateColumn(18)+FORMAT(RowNo)).Value :=  FORMAT(inwqty)
                        ELSE
                        xlSheet.Range(UpdateColumn(18)+FORMAT(RowNo)).Value :=  FORMAT(QtyLoaded);
                    
                        xlSheet.Range(UpdateColumn(19)+FORMAT(RowNo)).Value :=  FORMAT(Utilised)+'%';
                        xlSheet.Range(UpdateColumn(20)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."LR/RR No.";
                        xlSheet.Range(UpdateColumn(21)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."LR/RR Date");
                        xlSheet.Range(UpdateColumn(22)+FORMAT(RowNo)).Value :=  BillNo;
                        xlSheet.Range(UpdateColumn(23)+FORMAT(RowNo)).Value :=  FORMAT(DocPostingDate);
                        xlSheet.Range(UpdateColumn(24)+FORMAT(RowNo)).Value :=  FORMAT(Roadpermitno);
                        xlSheet.Range(UpdateColumn(24)+FORMAT(RowNo)).NumberFormat := '0';
                        xlSheet.Range(UpdateColumn(25)+FORMAT(RowNo)).Value :=  FORMAT(InvPostTime);
                        xlSheet.Range(UpdateColumn(26)+FORMAT(RowNo)).Value :=  FORMAT(InFactoryTime);
                        xlSheet.Range(UpdateColumn(27)+FORMAT(RowNo)).Value :=  comment1;
                        xlSheet.Range(UpdateColumn(28)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Entry Type");
                        xlSheet.Range(UpdateColumn(29)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Vehicle Body Height/Length");
                    END;
                    */
                    //Automation Commented 20Mar2017

                    //>>3


                    //>>21Mar2017 Report_GroupBody
                    IF NOT PrintSummry AND (HCOUNT = LCOUNT) THEN BEGIN
                        outime := CREATEDATETIME("Posted Gate Entry Header"."Out Date", "Posted Gate Entry Header"."Out Time");
                        IF ("Posted Gate Entry Header"."Vehicle Taken Date" <> 0D) AND ("Posted Gate Entry Header"."Vehicle Taken Time" <> 0T) THEN //08Mar2018
                            inttime := CREATEDATETIME("Posted Gate Entry Header"."Vehicle Taken Date", "Posted Gate Entry Header"."Vehicle Taken Time");

                        IF (outime <> 0DT) AND (inttime <> 0DT) THEN//RSPLSUM 11Jan21
                            InFactoryTime := outime - inttime;

                        EnterCell(NN, 1, "Posted Gate Entry Header"."No.", TRUE, FALSE, '', 1);
                        EnterCell(NN, 2, "Posted Gate Entry Header"."Vehicle No.", TRUE, FALSE, '', 1);
                        EnterCell(NN, 3, "Posted Gate Entry Header"."Transporter Name", TRUE, FALSE, '', 1);
                        EnterCell(NN, 4, FromLocaName, TRUE, FALSE, '', 1);
                        //  IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type"::"Sales Invoice" THEN
                        //RSPLSUM02Apr21--EnterCell(NN,5,'Invoice',TRUE,FALSE,'',1)
                        EnterCell(NN, 5, '', TRUE, FALSE, '', 1);//RSPLSUM02Apr21
                                                                 // ELSE
                        iF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type"::"Transfer Shipment" THEN
                            //RSPLSUM02Apr21--EnterCell(NN,5,'Transfer',TRUE,FALSE,'',1);
                            EnterCell(NN, 5, '', TRUE, FALSE, '', 1);//RSPLSUM02Apr21

                        EnterCell(NN, 6, FORMAT("Posted Gate Entry Header"."Document Date"), TRUE, FALSE, '', 2);
                        EnterCell(NN, 7, FORMAT("Posted Gate Entry Header"."Document Time"), TRUE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 8, FORMAT("Posted Gate Entry Header"."Releasing Date"), TRUE, FALSE, '', 2);
                        EnterCell(NN, 9, FORMAT("Posted Gate Entry Header"."Releasing Time"), TRUE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 10, FORMAT("Posted Gate Entry Header"."Vehicle Taken Date"), TRUE, FALSE, '', 2);
                        EnterCell(NN, 11, FORMAT("Posted Gate Entry Header"."Vehicle Taken Time"), TRUE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 12, FORMAT("Posted Gate Entry Header"."Posting Date"), TRUE, FALSE, '', 2);
                        EnterCell(NN, 13, FORMAT("Posted Gate Entry Header"."Posting Time"), TRUE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 14, FORMAT("Posted Gate Entry Header"."Out Date"), TRUE, FALSE, '', 2);
                        EnterCell(NN, 15, FORMAT("Posted Gate Entry Header"."Out Time"), TRUE, FALSE, 'hh:mm:ss AM/PM', 3);
                        //RSPLSUM02Apr21--EnterCell(NN,16,PartyName,TRUE,FALSE,'',1);
                        EnterCell(NN, 16, '', TRUE, FALSE, '', 1);//RSPLSUM02Apr21
                        EnterCell(NN, 17, "Posted Gate Entry Header"."Vehicle Capacity", TRUE, FALSE, '', 1);
                        //RSPLSUM02Apr21--IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Inward THEN
                        //RSPLSUM02Apr21--EnterCell(NN,18,FORMAT(inwqty),TRUE,FALSE,'0.000',0)
                        //RSPLSUM02Apr21--ELSE
                        //RSPLSUM02Apr21--EnterCell(NN,18,FORMAT(qtytot),TRUE,FALSE,'0.000',0);

                        //EnterCell(NN,19,FORMAT(ROUND(Utilised,0.0))+'%',TRUE,FALSE,'',1);
                        IF Utilisedtot = 0 THEN
                            EnterCell(NN, 18, FORMAT(Utilisedtot), TRUE, FALSE, '0.00 %', 0)
                        ELSE
                            EnterCell(NN, 18, FORMAT(Utilisedtot / 100), TRUE, FALSE, '0.00 %', 0);

                        EnterCell(NN, 19, "Posted Gate Entry Header"."LR/RR No.", TRUE, FALSE, '', 1);
                        EnterCell(NN, 20, FORMAT("Posted Gate Entry Header"."LR/RR Date"), TRUE, FALSE, '', 2);
                        EnterCell(NN, 21, BillNo, TRUE, FALSE, '', 1);
                        EnterCell(NN, 22, FORMAT(DocPostingDate), TRUE, FALSE, '', 2);
                        EnterCell(NN, 23, FORMAT(Roadpermitno), TRUE, FALSE, '', 1);
                        IF InvPostTime <> 0T THEN
                            EnterCell(NN, 24, FORMAT(InvPostTime), TRUE, FALSE, 'hh:mm:ss AM/PM', 3);

                        EnterCell(NN, 25, FORMAT(InFactoryTime), TRUE, FALSE, '', 1);
                        EnterCell(NN, 26, comment1, TRUE, FALSE, '', 1);
                        //RSPLSUM02Apr21--EnterCell(NN,27,FORMAT("Posted Gate Entry Header"."Entry Type"),TRUE,FALSE,'',1);
                        EnterCell(NN, 27, '', TRUE, FALSE, '', 1);//RSPLSUM02Apr21
                        EnterCell(NN, 28, FORMAT("Posted Gate Entry Header"."Vehicle Body Height/Length"), TRUE, FALSE, '', 1);
                        EnterCell(NN, 29, GRNNo, TRUE, FALSE, '', 1);//02May2019
                        EnterCell(NN, 30, FORMAT(GRNDate), TRUE, FALSE, '', 2);//02May2019
                        EnterCell(NN, 31, FORMAT("Posted Gate Entry Header"."Gross Weight"), TRUE, FALSE, '0.00', 0);//RSPLSUM 28Dec2020
                        EnterCell(NN, 32, FORMAT("Posted Gate Entry Header"."Tare Weight"), TRUE, FALSE, '0.00', 0);//RSPLSUM 28Dec2020
                        EnterCell(NN, 33, FORMAT("Posted Gate Entry Header"."Net Weight"), TRUE, FALSE, '0.00', 0);//RSPLSUM 28Dec2020
                        EnterCell(NN, 34, FinalVehicleCapacity, TRUE, FALSE, '', 1);//DJ 22012021
                                                                                    //RSPLSUM02Apr21--EnterCell(NN,36,CityName,TRUE,FALSE,'',1);//RSPLSUM 27Jan21
                        EnterCell(NN, 35, '', TRUE, FALSE, '', 1);//RSPLSUM02Apr21
                        EnterCell(NN, 36, FORMAT("Posted Gate Entry Header"."Type of Vehicle"), TRUE, FALSE, '', 1);//RSPLSUM 04Mar21
                        NN += 1;

                    END;

                    //<<21Mar2017 Report_GroupBody

                    //>>21Mar2017 Report_GroupBody
                    IF PrintSummry AND (HCOUNT = LCOUNT) THEN BEGIN

                        outime := CREATEDATETIME("Posted Gate Entry Header"."Out Date", "Posted Gate Entry Header"."Out Time");
                        IF ("Posted Gate Entry Header"."Vehicle Taken Date" <> 0D) AND ("Posted Gate Entry Header"."Vehicle Taken Time" <> 0T) THEN //08Mar2018
                            inttime := CREATEDATETIME("Posted Gate Entry Header"."Vehicle Taken Date", "Posted Gate Entry Header"."Vehicle Taken Time");

                        IF (outime <> 0DT) AND (inttime <> 0DT) THEN//RSPLSUM 11Jan21
                            InFactoryTime := outime - inttime;

                        EnterCell(NN, 1, "Posted Gate Entry Header"."No.", FALSE, FALSE, '', 1);
                        EnterCell(NN, 2, "Posted Gate Entry Header"."Vehicle No.", FALSE, FALSE, '', 1);
                        EnterCell(NN, 3, "Posted Gate Entry Header"."Transporter Name", FALSE, FALSE, '', 1);
                        EnterCell(NN, 4, FromLocaName, FALSE, FALSE, '', 1);
                        //  IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type"::"Sales Invoice" THEN
                        EnterCell(NN, 5, 'Invoice', FALSE, FALSE, '', 1);
                        //   ELSE
                        iF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type"::"Transfer Shipment" THEN
                            EnterCell(NN, 5, 'Transfer', FALSE, FALSE, '', 1);

                        EnterCell(NN, 6, FORMAT("Posted Gate Entry Header"."Document Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 7, FORMAT("Posted Gate Entry Header"."Document Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 8, FORMAT("Posted Gate Entry Header"."Releasing Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 9, FORMAT("Posted Gate Entry Header"."Releasing Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 10, FORMAT("Posted Gate Entry Header"."Vehicle Taken Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 11, FORMAT("Posted Gate Entry Header"."Vehicle Taken Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 12, FORMAT("Posted Gate Entry Header"."Posting Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 13, FORMAT("Posted Gate Entry Header"."Posting Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 14, FORMAT("Posted Gate Entry Header"."Out Date"), FALSE, FALSE, '', 2);
                        EnterCell(NN, 15, FORMAT("Posted Gate Entry Header"."Out Time"), FALSE, FALSE, 'hh:mm:ss AM/PM', 3);
                        EnterCell(NN, 16, PartyName, FALSE, FALSE, '', 1);
                        EnterCell(NN, 17, "Posted Gate Entry Header"."Vehicle Capacity", FALSE, FALSE, '', 1);
                        //RSPLSUM02Apr21--IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Inward THEN
                        //RSPLSUM02Apr21--EnterCell(NN,18,FORMAT(inwqty),FALSE,FALSE,'0.000',0)
                        //RSPLSUM02Apr21--ELSE
                        //RSPLSUM02Apr21--EnterCell(NN,18,FORMAT(qtytot),FALSE,FALSE,'0.000',0);

                        //EnterCell(NN,19,FORMAT(ROUND(Utilised,0.0))+'%',FALSE,FALSE,'',1);
                        IF Utilisedtot = 0 THEN
                            EnterCell(NN, 18, FORMAT(Utilisedtot), FALSE, FALSE, '0.00 %', 0)
                        ELSE
                            EnterCell(NN, 18, FORMAT(Utilisedtot / 100), FALSE, FALSE, '0.00 %', 0);

                        EnterCell(NN, 19, "Posted Gate Entry Header"."LR/RR No.", FALSE, FALSE, '', 1);
                        EnterCell(NN, 20, FORMAT("Posted Gate Entry Header"."LR/RR Date"), FALSE, FALSE, '', 2);
                        //EnterCell(NN,22,BillNo,FALSE,FALSE,'',1);
                        //EnterCell(NN,23,FORMAT(DocPostingDate),FALSE,FALSE,'',2);
                        EnterCell(NN, 21, FORMAT(Roadpermitno), FALSE, FALSE, '', 1);
                        //IF InvPostTime <> 0T THEN
                        //EnterCell(NN,23,FORMAT(InvPostTime,0,'<Hours12>:<Minutes,2>:<Seconds,2> <AM/PM>'),FALSE,FALSE,'',1);

                        EnterCell(NN, 22, FORMAT(InFactoryTime), FALSE, FALSE, '', 1);
                        EnterCell(NN, 23, comment1, FALSE, FALSE, '', 1);
                        EnterCell(NN, 24, FORMAT("Posted Gate Entry Header"."Entry Type"), FALSE, FALSE, '', 1);
                        EnterCell(NN, 25, FORMAT("Posted Gate Entry Header"."Vehicle Body Height/Length"), FALSE, FALSE, '', 1);
                        EnterCell(NN, 26, GRNNo, FALSE, FALSE, '', 1);//02May2019
                        EnterCell(NN, 27, FORMAT(GRNDate), FALSE, FALSE, '', 2);//02May2019
                        EnterCell(NN, 28, FORMAT("Posted Gate Entry Header"."Gross Weight"), FALSE, FALSE, '0.00', 0);//RSPLSUM 06Jan21
                        EnterCell(NN, 29, FORMAT("Posted Gate Entry Header"."Tare Weight"), FALSE, FALSE, '0.00', 0);//RSPLSUM 06Jan21
                        EnterCell(NN, 30, FORMAT("Posted Gate Entry Header"."Net Weight"), FALSE, FALSE, '0.00', 0);//RSPLSUM 06Jan21
                        EnterCell(NN, 31, FinalVehicleCapacity, TRUE, FALSE, '', 1);//DJ 22012021
                                                                                    //RSPLSUM02Apr21--EnterCell(NN,32,CityName,FALSE,FALSE,'',1);//RSPLSUM 27Jan21
                        EnterCell(NN, 32, '', FALSE, FALSE, '', 1);//RSPLSUM02Apr21
                        EnterCell(NN, 33, FORMAT("Posted Gate Entry Header"."Type of Vehicle"), FALSE, FALSE, '', 1);//RSPLSUM 04Mar21
                        NN += 1;
                    END;
                    //<<21Mar2017 Report_GroupBody
                    /*
                    //Posted Gate Entry Line, GroupFooter (2) - OnPreSection()
                    IF NOT PrintSummry THEN
                    BEGIN
                           RowNo += 1;
                     {
                        outime :=  CREATEDATETIME("Posted Gate Entry Header"."Out Date","Posted Gate Entry Header"."Out Time");
                        inttime := CREATEDATETIME("Posted Gate Entry Header"."Releasing Date","Posted Gate Entry Header"."Releasing Time");
                        InFactoryTime := ABS((inttime - outime)/3600000);
                     }
                        outime :=  CREATEDATETIME("Posted Gate Entry Header"."Out Date","Posted Gate Entry Header"."Out Time");
                        inttime := CREATEDATETIME("Posted Gate Entry Header"."Vehicle Taken Date","Posted Gate Entry Header"."Vehicle Taken Time");
                        InFactoryTime := outime - inttime;
                    
                    
                      //  InFactoryTimetext := FORMAT(InFactoryTime);
                      //  InFactoryTimetext := DELSTR(InFactoryTimetext,19,28);
                      //  IF EVALUATE(VehicleCapacity,"Posted Gate Entry Header"."Vehicle Capacity") THEN;
                    
                        xlSheet.Range('A6:T6').Font.Bold := TRUE;
                     //   xlSheet.Range(UpdateColumn(1)+FORMAT(RowNo)).Font.Bold := TRUE;
                        xlSheet.Range(UpdateColumn(1)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."No.";
                        xlSheet.Range(UpdateColumn(2)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Vehicle No." ;
                        xlSheet.Range(UpdateColumn(3)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Transporter Name";
                        xlSheet.Range(UpdateColumn(4)+FORMAT(RowNo)).Value :=  FromLocaName;
                    
                        IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type" :: "Sales Invoice" THEN
                          xlSheet.Range(UpdateColumn(5)+FORMAT(RowNo)).Value :=  'Invoice'
                        ELSE
                        IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type" :: "Transfer Shipment"  THEN
                          xlSheet.Range(UpdateColumn(5)+FORMAT(RowNo)).Value :=  'Transfer';
                    
                    {
                        xlSheet.Range(UpdateColumn(6)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Document Time") ;
                        xlSheet.Range(UpdateColumn(7)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Releasing Time") ;
                        xlSheet.Range(UpdateColumn(8)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Vehicle Taken Time") ;
                        xlSheet.Range(UpdateColumn(9)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Posting Time") ;
                        xlSheet.Range(UpdateColumn(10)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Out Time") ;
                    }
                        xlSheet.Range(UpdateColumn(6)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Document Date") ;
                        xlSheet.Range(UpdateColumn(7)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Document Time") ;
                        xlSheet.Range(UpdateColumn(8)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Releasing Date") ;
                        xlSheet.Range(UpdateColumn(9)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Releasing Time") ;
                        xlSheet.Range(UpdateColumn(10)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Vehicle Taken Date") ;
                        xlSheet.Range(UpdateColumn(11)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Vehicle Taken Time") ;
                        xlSheet.Range(UpdateColumn(12)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Posting Date") ;
                        xlSheet.Range(UpdateColumn(13)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Posting Time") ;
                        xlSheet.Range(UpdateColumn(14)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Out Date") ;
                        xlSheet.Range(UpdateColumn(15)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Out Time") ;
                    
                    
                        xlSheet.Range(UpdateColumn(16)+FORMAT(RowNo)).Value :=  PartyName ;
                        xlSheet.Range(UpdateColumn(17)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Vehicle Capacity";
                        IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Inward THEN
                        xlSheet.Range(UpdateColumn(18)+FORMAT(RowNo)).Value :=  FORMAT(inwqty)
                        ELSE
                        xlSheet.Range(UpdateColumn(18)+FORMAT(RowNo)).Value :=  FORMAT(qtytot);
                        xlSheet.Range(UpdateColumn(19)+FORMAT(RowNo)).Value :=  FORMAT(Utilisedtot)+'%';
                        xlSheet.Range(UpdateColumn(20)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."LR/RR No.";
                        xlSheet.Range(UpdateColumn(21)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."LR/RR Date");
                        xlSheet.Range(UpdateColumn(22)+FORMAT(RowNo)).Value :=  BillNo;
                        xlSheet.Range(UpdateColumn(23)+FORMAT(RowNo)).Value :=  FORMAT(DocPostingDate);
                        xlSheet.Range(UpdateColumn(24)+FORMAT(RowNo)).Value :=  FORMAT(Roadpermitno);
                        xlSheet.Range(UpdateColumn(24)+FORMAT(RowNo)).NumberFormat := '0';
                        xlSheet.Range(UpdateColumn(25)+FORMAT(RowNo)).Value :=  FORMAT(InvPostTime);
                        xlSheet.Range(UpdateColumn(26)+FORMAT(RowNo)).Value :=  FORMAT(InFactoryTime);
                        xlSheet.Range(UpdateColumn(27)+FORMAT(RowNo)).Value :=  FORMAT(comment1);
                        xlSheet.Range(UpdateColumn(28)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Entry Type");
                        xlSheet.Range(UpdateColumn(29)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Vehicle Body Height/Length");
                     END
                     ELSE
                     BEGIN
                        RowNo += 1;
                        outime :=  CREATEDATETIME("Posted Gate Entry Header"."Out Date","Posted Gate Entry Header"."Out Time");
                        inttime := CREATEDATETIME("Posted Gate Entry Header"."Vehicle Taken Date","Posted Gate Entry Header"."Vehicle Taken Time");
                        InFactoryTime := outime - inttime;
                    
                        xlSheet.Range('A6:T6').Font.Bold := TRUE;
                     //   xlSheet.Range(UpdateColumn(1)+FORMAT(RowNo)).Font.Bold := TRUE;
                        xlSheet.Range(UpdateColumn(1)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."No.";
                        xlSheet.Range(UpdateColumn(2)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Vehicle No." ;
                        xlSheet.Range(UpdateColumn(3)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Transporter Name";
                        xlSheet.Range(UpdateColumn(4)+FORMAT(RowNo)).Value :=  FromLocaName;
                    
                        IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type" :: "Sales Invoice" THEN
                          xlSheet.Range(UpdateColumn(5)+FORMAT(RowNo)).Value :=  'Invoice'
                        ELSE
                        IF "Posted Gate Entry Line"."Source Type" = "Posted Gate Entry Line"."Source Type" :: "Transfer Shipment"  THEN
                          xlSheet.Range(UpdateColumn(5)+FORMAT(RowNo)).Value :=  'Transfer';
                    {
                        xlSheet.Range(UpdateColumn(6)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Document Time") ;
                        xlSheet.Range(UpdateColumn(7)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Releasing Time") ;
                        xlSheet.Range(UpdateColumn(8)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Vehicle Taken Time") ;
                        xlSheet.Range(UpdateColumn(9)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Posting Time") ;
                        xlSheet.Range(UpdateColumn(10)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Out Time") ;
                    }
                        xlSheet.Range(UpdateColumn(6)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Document Date") ;
                        xlSheet.Range(UpdateColumn(7)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Document Time") ;
                        xlSheet.Range(UpdateColumn(8)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Releasing Date") ;
                        xlSheet.Range(UpdateColumn(9)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Releasing Time") ;
                        xlSheet.Range(UpdateColumn(10)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Vehicle Taken Date") ;
                        xlSheet.Range(UpdateColumn(11)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Vehicle Taken Time") ;
                        xlSheet.Range(UpdateColumn(12)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Posting Date") ;
                        xlSheet.Range(UpdateColumn(13)+FORMAT(RowNo)).Value :=   FORMAT("Posted Gate Entry Header"."Posting Time") ;
                        xlSheet.Range(UpdateColumn(14)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Out Date") ;
                        xlSheet.Range(UpdateColumn(15)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Out Time") ;
                    
                    
                        xlSheet.Range(UpdateColumn(16)+FORMAT(RowNo)).Value :=  PartyName ;
                        xlSheet.Range(UpdateColumn(17)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."Vehicle Capacity";
                        IF "Posted Gate Entry Header"."Entry Type" = "Posted Gate Entry Header"."Entry Type"::Inward THEN
                        xlSheet.Range(UpdateColumn(18)+FORMAT(RowNo)).Value :=  FORMAT(inwqty)
                        ELSE
                        xlSheet.Range(UpdateColumn(18)+FORMAT(RowNo)).Value :=  FORMAT(qtytot);
                        xlSheet.Range(UpdateColumn(19)+FORMAT(RowNo)).Value :=  FORMAT(Utilisedtot)+'%';
                        xlSheet.Range(UpdateColumn(20)+FORMAT(RowNo)).Value :=  "Posted Gate Entry Header"."LR/RR No.";
                        xlSheet.Range(UpdateColumn(21)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."LR/RR Date");
                    //    xlSheet.Range(UpdateColumn(17)+FORMAT(RowNo)).Value :=  BillNo;
                    //    xlSheet.Range(UpdateColumn(18)+FORMAT(RowNo)).Value :=  DocPostingDate;
                        xlSheet.Range(UpdateColumn(22)+FORMAT(RowNo)).Value :=  FORMAT(Roadpermitno);
                        xlSheet.Range(UpdateColumn(22)+FORMAT(RowNo)).NumberFormat := '0';
                    //    xlSheet.Range(UpdateColumn(20)+FORMAT(RowNo)).Value :=  FORMAT(InvPostTime);
                        xlSheet.Range(UpdateColumn(23)+FORMAT(RowNo)).Value :=  FORMAT(InFactoryTime);
                        xlSheet.Range(UpdateColumn(24)+FORMAT(RowNo)).Value :=  comment1;
                        xlSheet.Range(UpdateColumn(25)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Entry Type");
                        xlSheet.Range(UpdateColumn(26)+FORMAT(RowNo)).Value :=  FORMAT("Posted Gate Entry Header"."Vehicle Body Height/Length");
                     END;
                    
                    //<<3
                    *///Automation Commented 20Mar2017

                end;

                trigger OnPreDataItem()
                begin

                    HCOUNT := COUNT;//21Mar2017
                    LCOUNT := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                CLEAR(qtytot);
                comment1 := '';
                Gateentrycomentline.RESET;
                Gateentrycomentline.SETRANGE(Gateentrycomentline."Gate Entry Type", "Posted Gate Entry Header"."Entry Type");
                Gateentrycomentline.SETRANGE(Gateentrycomentline."No.", "Posted Gate Entry Header"."No.");
                IF Gateentrycomentline.FINDFIRST THEN
                    comment1 := Gateentrycomentline.Comment
                ELSE
                    comment1 := '';

                CLEAR(inwqty);
                IF "Posted Gate Entry Header"."Total Quantity" <> 0 THEN BEGIN
                    inwqty := "Posted Gate Entry Header"."Total Quantity";
                END;
                IF "Posted Gate Entry Header"."Total Quantity" = 0 THEN BEGIN
                    inwqty := "Posted Gate Entry Header"."Vendor Location Net Weight";
                END;

                //<<1
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                DateFilter := "Posted Gate Entry Header".GETFILTER("Posted Gate Entry Header"."Posting Date");

                /*
                xlSheet.Range('A4').Font.Bold := TRUE;
                xlSheet.Range('A4').Font.Size := 11;
                xlSheet.Range('A4').Value := 'Date Filter :- ' + DateFilter ;
                //<<1
                *///Automation Commented 20Mar2017


                //>>20Mar2017 Report_Header
                IF NOT PrintSummry THEN BEGIN

                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '@', 1);
                    EnterCell(1, 37, FORMAT(TODAY, 0, 4), TRUE, FALSE, '@', 1);

                    EnterCell(2, 1, 'Vehicle Entry - Outward Report', TRUE, FALSE, '@', 1);
                    EnterCell(2, 37, USERID, TRUE, FALSE, '@', 1);

                    //EnterCell(3,1,FORMAT(TIME),TRUE,FALSE,'hh:mm:ss AM/PM',3);
                    //EnterCell(3,2,FORMAT(TIME),TRUE,FALSE,'hh:mm:ss tt',3);

                    EnterCell(5, 1, 'Date Filter :- ' + DateFilter, TRUE, TRUE, '@', 1);
                    EnterCell(5, 2, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 3, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 4, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 5, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 6, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 7, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 8, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 9, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 10, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 11, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 12, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 13, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 14, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 15, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 16, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 17, '', TRUE, TRUE, '@', 1);
                    //RSPLSUM02Apr21--EnterCell(5,18,'',TRUE,TRUE,'@',1);
                    EnterCell(5, 18, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 19, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 20, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 21, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 22, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 23, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 24, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 25, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 26, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 27, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 28, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 29, '', TRUE, TRUE, '@', 1);//02May2019
                    EnterCell(5, 30, '', TRUE, TRUE, '@', 1);//02May2019
                    EnterCell(5, 31, '', TRUE, TRUE, '@', 1);//RSPLSUM 28Dec2020
                    EnterCell(5, 32, '', TRUE, TRUE, '@', 1);//RSPLSUM 28Dec2020
                    EnterCell(5, 33, '', TRUE, TRUE, '@', 1);//RSPLSUM 28Dec2020
                    EnterCell(5, 34, '', TRUE, TRUE, '@', 1);//RSPLSUM 27Jan21
                    EnterCell(5, 35, '', TRUE, TRUE, '@', 1);//RSPLSUM 27Jan21
                    EnterCell(5, 36, '', TRUE, TRUE, '@', 1);//RSPLSUM 04Mar21

                    EnterCell(6, 1, 'No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 2, 'Vehicle No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 3, 'Transporter Name', TRUE, TRUE, '@', 1);
                    EnterCell(6, 4, 'From Location', TRUE, TRUE, '@', 1);
                    EnterCell(6, 5, 'Type', TRUE, TRUE, '@', 1);
                    EnterCell(6, 6, 'Created Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 7, 'Created Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 8, 'Release Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 9, 'Release Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 10, 'Vehicle Taken for Loading Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 11, 'Vehicle Taken for Loading Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 12, 'Posted Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 13, 'Posted Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 14, 'Vehicle Out Date', TRUE, TRUE, '@', 1);//
                    EnterCell(6, 15, 'Vehicle Out Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 16, 'Customer / Location', TRUE, TRUE, '@', 1);
                    EnterCell(6, 17, 'Vehicle Capacity', TRUE, TRUE, '@', 1);
                    //RSPLSUM02Apr21--EnterCell(6,18,'Qty utilized in KG.',TRUE,TRUE,'@',1);
                    EnterCell(6, 18, 'Utilized %', TRUE, TRUE, '@', 1);
                    EnterCell(6, 19, 'LR No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 20, 'LR Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 21, 'Attached Document No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 22, 'Attached Document Posting Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 23, 'Road Permit No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 24, 'Invoicing Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 25, 'Vehicle remain In Factory in Hrs/Mins.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 26, 'Comment', TRUE, TRUE, '@', 1);
                    EnterCell(6, 27, 'Type', TRUE, TRUE, '@', 1);
                    EnterCell(6, 28, 'Vehicle Body Height/Length', TRUE, TRUE, '@', 1);
                    EnterCell(6, 29, 'GRN No.', TRUE, TRUE, '@', 1);//02May2019
                    EnterCell(6, 30, 'GRN Date', TRUE, TRUE, '@', 1);//02May2019
                    EnterCell(6, 31, 'Gross Weight', TRUE, TRUE, '@', 1);//RSPLSUM 28Dec2020
                    EnterCell(6, 32, 'Tare Weight', TRUE, TRUE, '@', 1);//RSPLSUM 28Dec2020
                    EnterCell(6, 33, 'Net Weight', TRUE, TRUE, '@', 1);//RSPLSUM 28Dec2020
                    EnterCell(6, 34, 'Final Vehicle Capacity', TRUE, TRUE, '@', 1);//DJ 22012021
                    EnterCell(6, 35, 'City Name', TRUE, TRUE, '@', 1);//RSPLSUM 27Jan21
                    EnterCell(6, 36, 'Vehicle Type', TRUE, TRUE, '@', 1);//RSPLSUM 04Mar21
                    NN := 7;

                END ELSE BEGIN

                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '@', 1);
                    EnterCell(1, 33, FORMAT(TODAY, 0, 4), TRUE, FALSE, '@', 1);

                    EnterCell(2, 1, 'Vehicle Entry - Outward Report', TRUE, FALSE, '@', 1);
                    EnterCell(2, 33, USERID, TRUE, FALSE, '@', 1);


                    EnterCell(5, 1, 'Date Filter :- ' + DateFilter, TRUE, TRUE, '@', 1);
                    EnterCell(5, 2, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 3, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 4, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 5, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 6, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 7, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 8, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 9, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 10, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 11, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 12, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 13, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 14, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 15, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 16, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 17, '', TRUE, TRUE, '@', 1);
                    //RSPLSUM02Apr21--EnterCell(5,18,'',TRUE,TRUE,'@',1);
                    EnterCell(5, 18, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 19, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 20, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 21, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 22, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 23, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 24, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 25, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 26, '', TRUE, TRUE, '@', 1);//02May2019
                    EnterCell(5, 27, '', TRUE, TRUE, '@', 1);//02May2019
                    EnterCell(5, 28, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Jan21
                    EnterCell(5, 29, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Jan21
                    EnterCell(5, 30, '', TRUE, TRUE, '@', 1);//RSPLSUM 06Jan21
                    EnterCell(5, 31, '', TRUE, TRUE, '@', 1);//RSPLSUM 27Jan21
                    EnterCell(5, 32, '', TRUE, TRUE, '@', 1);//RSPLSUM 27Jan21
                    EnterCell(5, 33, '', TRUE, TRUE, '@', 1);//RSPLSUM 04Mar21

                    EnterCell(6, 1, 'No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 2, 'Vehicle No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 3, 'Transporter Name', TRUE, TRUE, '@', 1);
                    EnterCell(6, 4, 'From Location', TRUE, TRUE, '@', 1);
                    EnterCell(6, 5, 'Type', TRUE, TRUE, '@', 1);
                    EnterCell(6, 6, 'Created Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 7, 'Created Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 8, 'Release Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 9, 'Release Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 10, 'Vehicle Taken for Loading Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 11, 'Vehicle Taken for Loading Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 12, 'Posted Date', TRUE, TRUE, '@', 1);
                    EnterCell(6, 13, 'Posted Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 14, 'Vehicle Out Date', TRUE, TRUE, '@', 1);//
                    EnterCell(6, 15, 'Vehicle Out Time', TRUE, TRUE, '@', 1);
                    EnterCell(6, 16, 'Customer / Location', TRUE, TRUE, '@', 1);
                    EnterCell(6, 17, 'Vehicle Capacity', TRUE, TRUE, '@', 1);
                    //RSPLSUM02Apr21--EnterCell(6,18,'Qty utilized in KG.',TRUE,TRUE,'@',1);
                    EnterCell(6, 18, 'Utilized %', TRUE, TRUE, '@', 1);
                    EnterCell(6, 19, 'LR No.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 20, 'LR Date', TRUE, TRUE, '@', 1);
                    //EnterCell(6,22,'Attached Document No.',TRUE,TRUE,'@',1);
                    //EnterCell(6,23,'Attached Document Posting Date',TRUE,TRUE,'@',1);
                    EnterCell(6, 21, 'Road Permit No.', TRUE, TRUE, '@', 1);
                    //EnterCell(6,23,'Invoicing Time',TRUE,TRUE,'@',1);
                    EnterCell(6, 22, 'Vehicle remain In Factory in Hrs/Mins.', TRUE, TRUE, '@', 1);
                    EnterCell(6, 23, 'Comment', TRUE, TRUE, '@', 1);
                    EnterCell(6, 24, 'Type', TRUE, TRUE, '@', 1);
                    EnterCell(6, 25, 'Vehicle Body Height/Length', TRUE, TRUE, '@', 1);
                    EnterCell(6, 26, 'GRN No.', TRUE, TRUE, '@', 1);//02May2019
                    EnterCell(6, 27, 'GRN Date', TRUE, TRUE, '@', 1);//02May2019
                    EnterCell(6, 28, 'Gross Weight', TRUE, TRUE, '@', 1);//RSPLSUM 06Jan21
                    EnterCell(6, 29, 'Tare Weight', TRUE, TRUE, '@', 1);//RSPLSUM 06Jan21
                    EnterCell(6, 30, 'Net Weight', TRUE, TRUE, '@', 1);//RSPLSUM 06Jan21
                    EnterCell(6, 31, 'Final Vehicle Capacity', TRUE, TRUE, '@', 1);//DJ 22012021
                    EnterCell(6, 32, 'City Name', TRUE, TRUE, '@', 1);//RSPLSUM 27Jan21
                    EnterCell(6, 33, 'Vehicle Type', TRUE, TRUE, '@', 1);//RSPLSUM 04Mar21
                    NN := 7;

                END;
                //<<20Mar2017 Report_Header

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Print Summary"; PrintSummry)
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

    trigger OnPostReport()
    begin
        //>>1
        /*
        xlApp.Visible(TRUE);
        //xlRange.NumberFormat := '0';
        xlSheet.Columns.AutoFit;
        MESSAGE('Report Generation Completed');
        //<<1
        *///Automation Commented 20Mar2017

        //>>20Mar2017 RB-N
        /* ExcBuffer.CreateBook('', 'Vehicle Entry - Outward Report');
        ExcBuffer.CreateBookAndOpenExcel('', 'Vehicle Entry - Outward Report', '', '', USERID);
        ExcBuffer.GiveUserControl; */
        //<<20Mar2017


        ExcBuffer.CreateNewBook('Vehicle Entry - Outward Report');
        ExcBuffer.WriteSheet('Vehicle Entry - Outward Report', CompanyName, UserId);
        ExcBuffer.CloseBook();
        ExcBuffer.SetFriendlyFilename(StrSubstNo('Vehicle Entry - Outward Report', CurrentDateTime, UserId));
        ExcBuffer.OpenExcel();

    end;

    trigger OnPreReport()
    begin
        //>>1

        CompanyInfo.GET;
        //CreateXLSheet; //Automation Commented 20Mar2017
        SrNo := 0;

        //<<1
    end;

    var
        CompInfo: Record 79;
        RowNo: Integer;
        ColNo: Integer;
        i: Integer;
        CompanyInfo: Record 79;
        SrNo: Integer;
        recSalesShpmt: Record 113;
        recTransShpmt: Record 5745;
        ForLoc: Code[10];
        PartyName: Text[60];
        QtyLoaded: Decimal;
        BillNo: Code[20];
        recCust: Record 18;
        recLoc: Record 14;
        InFactoryTime: Duration;
        outime: DateTime;
        inttime: DateTime;
        VehicleCapacity: Decimal;
        DateFilter: Text[30];
        FromLocaName: Text[30];
        Vehiclecapacityrec: Record 50004;
        Utilised: Decimal;
        DocPostingDate: Date;
        Roadpermitno: Text[30];
        salesinvheader: Record 112;
        transfershipheader: Record 5744;
        InvPostTime: Time;
        InFactoryTimetext: Text[100];
        PrintSummry: Boolean;
        qtytot: Decimal;
        Utilisedtot: Decimal;
        Gateentrycomentline: Record "Gate Entry Comment Line";
        comment1: Text[80];
        inwqty: Decimal;
        "-----20Mar2017": Integer;
        ExcBuffer: Record 370 temporary;
        NN: Integer;
        LCOUNT: Integer;
        HCOUNT: Integer;
        PRH: Record 120;
        GRNNo: Code[20];
        GRNDate: Date;
        FinalVehicleCapacity: Text;
        RecShipToAdd: Record 222;
        CityName: Text[30];
        RecCustomer: Record 18;
        RecLocation: Record 14;

    //  [Scope('Internal')]
    procedure CreateXLSheet()
    begin

        /*
        CLEAR(xlApp);
        CLEAR(xlWBook);
        CLEAR(xlSheet);
        CREATE(xlApp);
        xlWBook := xlApp.Workbooks.Add;
        xlSheet := xlWBook.Worksheets.Add;
        CreateHeader;
        *///Automation Commented 20Mar2017

    end;

    // [Scope('Internal')]
    procedure CreateHeader()
    begin

        /*
        IF NOT PrintSummry THEN
        BEGIN
          xlSheet.Activate;
        
          xlSheet.Range('A1').Font.Bold := TRUE;
          xlSheet.Range('A1').Font.Size := 11;
          xlSheet.Range('A1').Value := CompanyInfo.Name;
        
          xlSheet.Range('A3').Font.Bold := TRUE;
          xlSheet.Range('A3').Font.Size := 11;
          xlSheet.Range('A3').Value := 'Vehicle Entry - Outward Report';
        
        
          xlSheet.Range('A3').Font.Bold := TRUE;
          xlSheet.Range('A3').Font.Size := 11;
        
        
        
          xlSheet.Range('A6:AC6').Font.Bold := TRUE;
          xlSheet.Range('A6').Value := 'No.';
          xlSheet.Range('B6').Value := 'Vehicle No.';
          xlSheet.Range('C6').Value := 'Transporter Name';
          xlSheet.Range('D6').Value := 'From Location';
          xlSheet.Range('E6').Value := 'Type';
        {
          xlSheet.Range('F6').Value := 'Created Time';
          xlSheet.Range('G6').Value := 'Release Time';
          xlSheet.Range('H6').Value := 'Vehicle Taken for Loading Time';
          xlSheet.Range('I6').Value := 'Posted Time';
          xlSheet.Range('J6').Value := 'Vehicle Out Time';
        }
        //
          xlSheet.Range('F6').Value := 'Created Date';
          xlSheet.Range('G6').Value := 'Created Time';
          xlSheet.Range('H6').Value := 'Release Date';
          xlSheet.Range('I6').Value := 'Release Time';
          xlSheet.Range('J6').Value := 'Vehicle Taken for Loading Date';
          xlSheet.Range('K6').Value := 'Vehicle Taken for Loading Time';
          xlSheet.Range('L6').Value := 'Posted Date';
          xlSheet.Range('M6').Value := 'Posted Time';
          xlSheet.Range('N6').Value := 'Vehicle Out date';
          xlSheet.Range('O6').Value := 'Vehicle Out Time';
        
        //
        
          xlSheet.Range('P6').Value := 'Customer / Location';
          xlSheet.Range('Q6').Value := 'Vehicle Capacity';
          xlSheet.Range('R6').Value := 'Qty utilized in KG.';
          xlSheet.Range('S6').Value := 'Utilized %';
          xlSheet.Range('T6').Value := 'LR No.';
          xlSheet.Range('U6').Value := 'LR Date';
          xlSheet.Range('V6').Value := 'Attached Document No.';
          xlSheet.Range('W6').Value := 'Attached Document Posting Date';
          xlSheet.Range('X6').Value := 'Road Permit No.';
          xlSheet.Range('Y6').Value := 'Invoicing Time';
          xlSheet.Range('Z6').Value := 'Vehicle remain In Factory in Hrs/Mins.';
          xlSheet.Range('AA6').Value := 'Comment';
          xlSheet.Range('AB6').Value := 'Type';
          xlSheet.Range('AC6').Value := 'Vehicle Body Height/Length';
          //xlSheet.Range('A1:A6').HorizontalAlignment := 3;
          //xlSheet.Range('T1:T6').VerticalAlignment := 3;
        
          RowNo := 7;
        
          xlActiveWindow := xlApp.ActiveWindow;
        END
        ELSE
        BEGIN
          xlSheet.Activate;
        
          xlSheet.Range('A1').Font.Bold := TRUE;
          xlSheet.Range('A1').Font.Size := 11;
          xlSheet.Range('A1').Value := CompanyInfo.Name;
        
          xlSheet.Range('A3').Font.Bold := TRUE;
          xlSheet.Range('A3').Font.Size := 11;
          xlSheet.Range('A3').Value := 'Vehicle Entry - Outward Report';
        
        
          xlSheet.Range('A3').Font.Bold := TRUE;
          xlSheet.Range('A3').Font.Size := 11;
        
        
        
          xlSheet.Range('A6:Z6').Font.Bold := TRUE;
          xlSheet.Range('A6').Value := 'No.';
          xlSheet.Range('B6').Value := 'Vehicle No.';
          xlSheet.Range('C6').Value := 'Transporter Name';
          xlSheet.Range('D6').Value := 'From Location';
          xlSheet.Range('E6').Value := 'Type';
        {
          xlSheet.Range('F6').Value := 'Created Time';
          xlSheet.Range('G6').Value := 'Release Time';
          xlSheet.Range('H6').Value := 'Vehicle Taken for Loading Time';
          xlSheet.Range('I6').Value := 'Posted Time';
          xlSheet.Range('J6').Value := 'Vehicle Out Time';
        }
          xlSheet.Range('F6').Value := 'Created Date';
          xlSheet.Range('G6').Value := 'Created Time';
          xlSheet.Range('H6').Value := 'Release Date';
          xlSheet.Range('I6').Value := 'Release Time';
          xlSheet.Range('J6').Value := 'Vehicle Taken for Loading Date';
          xlSheet.Range('K6').Value := 'Vehicle Taken for Loading Time';
          xlSheet.Range('L6').Value := 'Posted Date';
          xlSheet.Range('M6').Value := 'Posted Time';
          xlSheet.Range('N6').Value := 'Vehicle Out Date';
          xlSheet.Range('O6').Value := 'Vehicle Out Time';
        
        
          xlSheet.Range('P6').Value := 'Customer / Location';
          xlSheet.Range('Q6').Value := 'Vehicle Capacity';
          xlSheet.Range('R6').Value := 'Qty utilized in KG.';
          xlSheet.Range('S6').Value := 'Utilized %';
          xlSheet.Range('T6').Value := 'LR No.';
          xlSheet.Range('U6').Value := 'LR Date';
        //  xlSheet.Range('Q6').Value := 'Attached Document No.';
        //  xlSheet.Range('R6').Value := 'Attached Document Posting Date';
          xlSheet.Range('V6').Value := 'Road Permit No.';
        //  xlSheet.Range('T6').Value := 'Invoicing Time';
          xlSheet.Range('W6').Value := 'Vehicle remain In Factory in Hrs/Mins.';
          xlSheet.Range('X6').Value := 'Comments';
          xlSheet.Range('Y6').Value := 'Type';
          xlSheet.Range('Z6').Value := 'Vehicle Body Height/Length';
          //xlSheet.Range('A1:A6').HorizontalAlignment := 3;
          //xlSheet.Range('T1:T6').VerticalAlignment := 3;
        
          RowNo := 7;
        
          xlActiveWindow := xlApp.ActiveWindow;
        
        END;
        *///Automation Commented 20Mar2017

    end;

    //  [Scope('Internal')]
    procedure UpdateColumn(SourceColNo: Integer): Code[10]
    var
        xlColID: Text[10];
        x: Integer;
        c: Char;
    begin
        IF SourceColNo <> 0 THEN
            ColNo := SourceColNo
        ELSE
            ColNo += 1;


        xlColID := '';
        IF ColNo <> 0 THEN BEGIN
            x := ColNo - 1;
            c := 65 + x MOD 26;
            xlColID[10] := c;
            i := 10;
            WHILE x > 25 DO BEGIN
                x := x DIV 26;
                i := i - 1;
                c := 64 + x MOD 26;
                xlColID[i] := c;
            END;
            FOR x := i TO 10 DO
                xlColID[1 + x - i] := xlColID[x];
        END;
        EXIT(xlColID);
    end;

    //   [Scope('Internal')]
    procedure CreateFooter()
    begin
    end;

    // [Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        ExcBuffer.INIT;
        ExcBuffer.VALIDATE("Row No.", Rowno);
        ExcBuffer.VALIDATE("Column No.", columnno);
        ExcBuffer."Cell Value as Text" := Cellvalue;
        ExcBuffer.Formula := '';
        ExcBuffer.Bold := Bold;
        ExcBuffer.Underline := Underline;
        ExcBuffer.NumberFormat := NoFormat;
        ExcBuffer."Cell Type" := CType;
        ExcBuffer.INSERT;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


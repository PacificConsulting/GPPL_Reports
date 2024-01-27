report 50094 "RM Stock Analysis MonthWise"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/RMStockAnalysisMonthWise.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.")
                                WHERE(Blocked = FILTER(false),
                                      "Item Category Code" = FILTER(<> 'CAT02|CAT03'));
            RequestFilterFields = "No.", "Item Category Code";

            trigger OnAfterGetRecord()
            begin


                //>>1

                OpeningBalance := 0;
                ReceiptQuantity := 0;
                DespQuantity := 0;
                OtherQuantity := 0;
                ClosingBalance := 0;
                LastMonthQuantity[1] := 0;
                LastMonthQuantity[2] := 0;
                LastMonthQuantity[3] := 0;
                LastMonthQuantity[4] := 0;
                LastMonthQuantity[5] := 0;
                LastMonthQuantity[6] := 0;
                LastMonthQuantity[7] := 0;
                LastMonthQuantity[8] := 0;
                LastMonthQuantity[9] := 0;
                LastMonthQuantity[10] := 0;
                LastMonthQuantity[11] := 0;
                LastMonthQuantity[12] := 0;
                LastMonthQuantity[13] := 0;
                LastMonthQuantity[14] := 0;
                LastMonthQuantity[15] := 0;
                LastMonthQuantity[16] := 0;
                LastMonthQuantity[17] := 0;
                LastMonthQuantity[18] := 0;
                LastMonthQuantity[19] := 0;
                LastMonthQuantity[20] := 0;
                LastMonthQuantity[21] := 0;
                LastMonthQuantity[22] := 0;
                LastMonthQuantity[23] := 0;
                LastMonthQuantity[24] := 0;

                recItemCategory.RESET;
                recItemCategory.SETRANGE(recItemCategory.Code, Item."Item Category Code");
                IF recItemCategory.FINDFIRST THEN;

                //RSPL
                itemDim := '';
                recDefualtDim.RESET;
                recDefualtDim.SETRANGE(recDefualtDim."Table ID", 27);
                recDefualtDim.SETRANGE(recDefualtDim."No.", "No.");
                IF recDefualtDim.FINDFIRST THEN
                    itemDim := recDefualtDim."Dimension Value Code";
                //RSPL

                ILE.RESET;
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '<%1', RequestedDate);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        OpeningBalance := OpeningBalance + ILE.Quantity;
                    UNTIL ILE.NEXT = 0;
                END;

                ILE.RESET;
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                ILE.SETFILTER(ILE.Quantity, '>%1', 0);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        ReceiptQuantity := ReceiptQuantity + ILE.Quantity;
                    UNTIL ILE.NEXT = 0;
                END;

                ILE.RESET;
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                ILE.SETRANGE(ILE."Entry Type", 5);
                ILE.SETFILTER(ILE.Quantity, '<%1', 0);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        DespQuantity := DespQuantity + ABS(ILE.Quantity);
                    UNTIL ILE.NEXT = 0;
                END;

                ILE.RESET;
                ILE.SETRANGE("Item No.", Item."No.");
                ILE.SETFILTER(ILE."Posting Date", '%1..%2', RequestedDate, RequestedDate1);
                IF RequestedLocation <> '' THEN
                    ILE.SETRANGE(ILE."Location Code", RequestedLocation);
                ILE.SETFILTER(ILE."Entry Type", '<>5');
                ILE.SETFILTER(ILE.Quantity, '<%1', 0);
                IF ILE.FINDSET THEN BEGIN
                    REPEAT
                        OtherQuantity := OtherQuantity + ABS(ILE.Quantity);
                    UNTIL ILE.NEXT = 0;
                END;


                ClosingBalance := (OpeningBalance + ReceiptQuantity) - (DespQuantity + OtherQuantity);

                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[1], EndDate[1]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[1] := LastMonthQuantity[1] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;


                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[2], EndDate[2]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[2] := LastMonthQuantity[2] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;



                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[3], EndDate[3]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[3] := LastMonthQuantity[3] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;


                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[4], EndDate[4]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[4] := LastMonthQuantity[4] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;


                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[5], EndDate[5]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[5] := LastMonthQuantity[5] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;

                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[6], EndDate[6]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[6] := LastMonthQuantity[6] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;

                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[7], EndDate[7]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[7] := LastMonthQuantity[7] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;

                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[8], EndDate[8]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[8] := LastMonthQuantity[8] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;

                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[9], EndDate[9]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[9] := LastMonthQuantity[9] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;

                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[10], EndDate[10]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[10] := LastMonthQuantity[10] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;

                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[11], EndDate[11]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[11] := LastMonthQuantity[11] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;

                ILErec.RESET;
                ILErec.SETRANGE(ILErec."Item No.", Item."No.");
                ILErec.SETFILTER(ILErec."Posting Date", '%1..%2', StartDate[12], EndDate[12]);
                ILErec.SETFILTER(ILErec."Entry Type", '%1', 5);
                ILErec.SETRANGE(ILErec."Location Code", RequestedLocation);
                IF ILErec.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[12] := LastMonthQuantity[12] + ILErec.Quantity;
                    UNTIL ILErec.NEXT = 0;
                END;


                /*
                SalesInvoiceLine.RESET;
                SalesInvoiceLine.SETRANGE(SalesInvoiceLine."No.",Item."No.");
                SalesInvoiceLine.SETFILTER(SalesInvoiceLine."Shipment Date",'%1..%2',StartDate[1],EndDate[1]);
                IF RequestedLocation <> '' THEN
                SalesInvoiceLine.SETRANGE(SalesInvoiceLine."Location Code",RequestedLocation);
                IF SalesInvoiceLine.FINDSET THEN
                BEGIN
                  REPEAT
                    LastMonthQuantity[1] := LastMonthQuantity[1] + SalesInvoiceLine."Quantity (Base)";
                  UNTIL SalesInvoiceLine.NEXT = 0;
                END;
                */


                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[1], EndDate[1]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[13] := LastMonthQuantity[13] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[2], EndDate[2]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[14] := LastMonthQuantity[14] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[3], EndDate[3]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[15] := LastMonthQuantity[15] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[4], EndDate[4]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[16] := LastMonthQuantity[16] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[5], EndDate[5]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[17] := LastMonthQuantity[17] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[6], EndDate[6]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[18] := LastMonthQuantity[18] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[7], EndDate[7]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[19] := LastMonthQuantity[19] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[8], EndDate[8]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[20] := LastMonthQuantity[20] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[9], EndDate[9]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[21] := LastMonthQuantity[21] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[10], EndDate[10]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[22] := LastMonthQuantity[22] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;


                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[11], EndDate[11]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[23] := LastMonthQuantity[23] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;

                TransferShipmentLine.RESET;
                TransferShipmentLine.SETRANGE(TransferShipmentLine."Item No.", Item."No.");
                TransferShipmentLine.SETFILTER(TransferShipmentLine."Shipment Date", '%1..%2', StartDate[12], EndDate[12]);
                IF RequestedLocation <> '' THEN
                    TransferShipmentLine.SETRANGE(TransferShipmentLine."Transfer-from Code", RequestedLocation);
                IF TransferShipmentLine.FINDSET THEN BEGIN
                    REPEAT
                        LastMonthQuantity[24] := LastMonthQuantity[24] + TransferShipmentLine."Quantity (Base)";
                    UNTIL TransferShipmentLine.NEXT = 0;
                END;



                IF (OpeningBalance = 0) AND (ReceiptQuantity = 0) AND (DespQuantity = 0) AND (OtherQuantity = 0) AND (LastMonthQuantity[1] = 0)
                   AND (LastMonthQuantity[2] = 0) AND (LastMonthQuantity[3] = 0) AND (LastMonthQuantity[4] = 0) AND (LastMonthQuantity[5] = 0)
                   AND (LastMonthQuantity[6] = 0) AND (LastMonthQuantity[7] = 0) AND (LastMonthQuantity[8] = 0) AND (LastMonthQuantity[9] = 0)
                   AND (LastMonthQuantity[10] = 0) AND (LastMonthQuantity[11] = 0) AND (LastMonthQuantity[12] = 0)

                   AND (LastMonthQuantity[13] = 0) AND (LastMonthQuantity[14] = 0) AND (LastMonthQuantity[15] = 0)
                   AND (LastMonthQuantity[16] = 0) AND (LastMonthQuantity[17] = 0) AND (LastMonthQuantity[18] = 0)
                   AND (LastMonthQuantity[19] = 0) AND (LastMonthQuantity[20] = 0) AND (LastMonthQuantity[21] = 0)
                    AND (LastMonthQuantity[22] = 0) AND (LastMonthQuantity[23] = 0) AND (LastMonthQuantity[24] = 0)

                     THEN
                    CurrReport.SKIP;

                //<<1


                //>>2

                //Item, Body (3) - OnPreSection()
                SrNo := SrNo + 1;

                IF PrintToExcel THEN
                    MakeExcelDataBody;
                //<<2

                //>>3

                //Item, GroupFooter (4) - OnPreSection()
                IF NOT FooterPrinted THEN
                    LastFieldNo := CurrReport.TOTALSCAUSEDBY;

                //CurrReport.SHOWOUTPUT := NOT FooterPrinted;

                FooterPrinted := TRUE;

                //<<3

            end;

            trigger OnPostDataItem()
            begin

                //>>4

                //Item, Footer (5) - OnPreSection()
                IF PrintToExcel THEN
                    MakeExcelDataFooter;
                //<<4
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                CurrReport.CREATETOTALS(OpeningBalance, ReceiptQuantity, DespQuantity, OtherQuantity, ClosingBalance);
                CurrReport.CREATETOTALS(LastMonthQuantity[1], LastMonthQuantity[2], LastMonthQuantity[3], LastMonthQuantity[4], LastMonthQuantity[5],
                                        LastMonthQuantity[6], LastMonthQuantity[7], LastMonthQuantity[8], LastMonthQuantity[9], LastMonthQuantity[10]);
                CurrReport.CREATETOTALS(LastMonthQuantity[11], LastMonthQuantity[12]);
                CurrReport.CREATETOTALS(LastMonthQuantity[13], LastMonthQuantity[14], LastMonthQuantity[15], LastMonthQuantity[16],
                                        LastMonthQuantity[17], LastMonthQuantity[18], LastMonthQuantity[19], LastMonthQuantity[20],
                                        LastMonthQuantity[21], LastMonthQuantity[22]);
                CurrReport.CREATETOTALS(LastMonthQuantity[23], LastMonthQuantity[24]);
                SrNo := 0;

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
                field(Month; RequestedMonth)
                {
                }
                field(Year; RequestedYear)
                {
                }
                field(Location; RequestedLocation)
                {
                    TableRelation = Location;
                }
                field("Print To Excel"; PrintToExcel)
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

        IF PrintToExcel THEN
            CreateExcelbook;

        //<<1
    end;

    trigger OnPreReport()
    begin
        //>>1
        CompanyInfo.GET;
        //<<1

        //>>2


        IF Item.GETFILTER(Item."Item Category Code") <> '' THEN BEGIN
            IF ((Item.GETFILTER(Item."Item Category Code") = 'CAT02') OR (Item.GETFILTER(Item."Item Category Code") = 'CAT03')) THEN BEGIN
                ERROR('You can not Select this Category');
            END;
        END;

        IF RequestedLocation = '' THEN BEGIN
            ERROR('Please Select Location');
        END;

        /*
        //EBT STIVAN ---(23072013)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        User.GET(USERID);
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",User."User ID");
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.SETRANGE(recLocation.Code,RequestedLocation);
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping.RESET;
         CSOMapping.SETRANGE(CSOMapping."User Id",UPPERCASE(USERID));
         CSOMapping.SETRANGE(Type,CSOMapping.Type::Location);
         CSOMapping.SETRANGE(CSOMapping.Value,RequestedLocation);
         IF NOT(CSOMapping.FINDFIRST) THEN
         BEGIN
          CSOMapping.RESET;
          CSOMapping.SETRANGE(CSOMapping."User Id",UPPERCASE(USERID));
          CSOMapping.SETRANGE(CSOMapping.Type,CSOMapping.Type::"Responsibility Center");
          CSOMapping.SETRANGE(CSOMapping.Value,LocResC);
          IF NOT(CSOMapping.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        //EBT STIVAN ---(23072013)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        *///Commented 02Mar2017
        RequestedDate := DMY2DATE(1, RequestedMonth, RequestedYear);
        RequestedDate1 := CALCDATE('<+1M', RequestedDate) - 1;

        StartDate[1] := RequestedDate;
        EndDate[1] := CALCDATE('CM', StartDate[1]);

        StartDate[2] := CALCDATE('-1M', RequestedDate);
        EndDate[2] := CALCDATE('CM', StartDate[2]);

        StartDate[3] := CALCDATE('-2M', RequestedDate);
        EndDate[3] := CALCDATE('CM', StartDate[3]);

        StartDate[4] := CALCDATE('-3M', RequestedDate);
        EndDate[4] := CALCDATE('CM', StartDate[4]);

        StartDate[5] := CALCDATE('-4M', RequestedDate);
        EndDate[5] := CALCDATE('CM', StartDate[5]);

        StartDate[6] := CALCDATE('-5M', RequestedDate);
        EndDate[6] := CALCDATE('CM', StartDate[6]);

        StartDate[7] := CALCDATE('-6M', RequestedDate);
        EndDate[7] := CALCDATE('CM', StartDate[7]);

        StartDate[8] := CALCDATE('-7M', RequestedDate);
        EndDate[8] := CALCDATE('CM', StartDate[8]);

        StartDate[9] := CALCDATE('-8M', RequestedDate);
        EndDate[9] := CALCDATE('CM', StartDate[9]);

        StartDate[10] := CALCDATE('-9M', RequestedDate);
        EndDate[10] := CALCDATE('CM', StartDate[10]);

        StartDate[11] := CALCDATE('-10M', RequestedDate);
        EndDate[11] := CALCDATE('CM', StartDate[11]);

        StartDate[12] := CALCDATE('-11M', RequestedDate);
        EndDate[12] := CALCDATE('CM', StartDate[12]);


        Monthly[1] := 'Consumption ' + FORMAT(StartDate[1], 0, '<Month text>,<Year4>');
        Monthly[2] := 'Consumption ' + FORMAT(StartDate[2], 0, '<Month text>,<Year4>');
        Monthly[3] := 'Consumption ' + FORMAT(StartDate[3], 0, '<Month text>,<Year4>');
        Monthly[4] := 'Consumption ' + FORMAT(StartDate[4], 0, '<Month text>,<Year4>');
        Monthly[5] := 'Consumption ' + FORMAT(StartDate[5], 0, '<Month text>,<Year4>');
        Monthly[6] := 'Consumption ' + FORMAT(StartDate[6], 0, '<Month text>,<Year4>');
        Monthly[7] := 'Consumption ' + FORMAT(StartDate[7], 0, '<Month text>,<Year4>');
        Monthly[8] := 'Consumption ' + FORMAT(StartDate[8], 0, '<Month text>,<Year4>');
        Monthly[9] := 'Consumption ' + FORMAT(StartDate[9], 0, '<Month text>,<Year4>');
        Monthly[10] := 'Consumption ' + FORMAT(StartDate[10], 0, '<Month text>,<Year4>');
        Monthly[11] := 'Consumption ' + FORMAT(StartDate[11], 0, '<Month text>,<Year4>');
        Monthly[12] := 'Consumption ' + FORMAT(StartDate[12], 0, '<Month text>,<Year4>');


        Monthly[13] := 'Transfer ' + FORMAT(StartDate[1], 0, '<Month text>,<Year4>');
        Monthly[14] := 'Transfer ' + FORMAT(StartDate[2], 0, '<Month text>,<Year4>');
        Monthly[15] := 'Transfer ' + FORMAT(StartDate[3], 0, '<Month text>,<Year4>');
        Monthly[16] := 'Transfer ' + FORMAT(StartDate[4], 0, '<Month text>,<Year4>');
        Monthly[17] := 'Transfer ' + FORMAT(StartDate[5], 0, '<Month text>,<Year4>');
        Monthly[18] := 'Transfer ' + FORMAT(StartDate[6], 0, '<Month text>,<Year4>');
        Monthly[19] := 'Transfer ' + FORMAT(StartDate[7], 0, '<Month text>,<Year4>');
        Monthly[20] := 'Transfer ' + FORMAT(StartDate[8], 0, '<Month text>,<Year4>');
        Monthly[21] := 'Transfer ' + FORMAT(StartDate[9], 0, '<Month text>,<Year4>');
        Monthly[22] := 'Transfer ' + FORMAT(StartDate[10], 0, '<Month text>,<Year4>');
        Monthly[23] := 'Transfer ' + FORMAT(StartDate[11], 0, '<Month text>,<Year4>');
        Monthly[24] := 'Transfer ' + FORMAT(StartDate[12], 0, '<Month text>,<Year4>');


        //>>09Mar2017 LocName
        LocCode := RequestedLocation;
        recLOC.RESET;
        recLOC.SETRANGE(Code, LocCode);
        IF recLOC.FINDFIRST THEN BEGIN
            LocName := recLOC.Name;
        END;
        //<< LocName


        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<2

    end;

    var
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        ILE: Record 32;
        OpeningBalance: Decimal;
        StartDate: array[12] of Date;
        EndDate: array[12] of Date;
        RequestedMonth: Option " ",January,February,March,April,May,June,July,August,September,October,November,December;
        RequestedYear: Integer;
        RequestedDate: Date;
        RequestedDate1: Date;
        RequestedLocation: Code[20];
        PurchRecptLine: Record 121;
        ReturnRecptLine: Record 6661;
        ReceiptQuantity: Decimal;
        DespQuantity: Decimal;
        OtherQuantity: Decimal;
        ClosingBalance: Decimal;
        SrNo: Integer;
        CompanyInfo: Record 79;
        LastMonthQuantity: array[24] of Decimal;
        SalesInvoiceLine: Record 113;
        Monthly: array[24] of Text[30];
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        recItemCategory: Record 5722;
        User: Record 2000000120;
        LocResC: Code[100];
        recLocation: Record 14;
        CSOMapping: Record 50006;
        TransferShipmentLine: Record 5745;
        ILErec: Record 32;
        itemDim: Code[20];
        recDefualtDim: Record 352;
        Text000: Label 'Data';
        Text001: Label 'RM Stock Analysis MonthWise ';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "----02Mar2017": Integer;
        TOpeningBalance: Decimal;
        TReceiptQuantity: Decimal;
        TDespQuantity: Decimal;
        TClosingBalance: Decimal;
        TLastMonthQuantity: array[24] of Decimal;
        "----09Mar2017": Integer;
        LocCode: Code[100];
        recLOC: Record 14;
        LocName: Text[250];
        LocName1: Text[100];

    // [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//02Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('STOCK ANALYSIS FOR RM', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"RM Stock Analysis MonthWise", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // [Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        /* //>>02Mar2017 RB-N
        ExcelBuf.CreateBook('', Text001);
        ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
        ExcelBuf.GiveUserControl; */
        //<<

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>02Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('RM Stock Analysis MonthWise', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.NewRow;

        //<<

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Stock Analysis From:', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(FORMAT(RequestedDate) + ' .. ' + FORMAT(RequestedDate1), FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        //ExcelBuf.AddColumn('To',FALSE,'',TRUE,FALSE,FALSE,'@',1);
        //ExcelBuf.AddColumn(FORMAT(RequestedDate1),FALSE,'',TRUE,FALSE,FALSE,'@',1);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location', FALSE, '', TRUE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(LocName, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//09Mar2017
        //ExcelBuf.AddColumn(RequestedLocation,FALSE,'',TRUE,FALSE,FALSE,'@',1);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('Sale UOM', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('Default Dimension', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//5 //RSPL
        ExcelBuf.AddColumn('Opening Balance', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('Receipt Qty.', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('Consumption Qty', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//8
        //ExcelBuf.AddColumn('Other Despatches',FALSE,'',TRUE,FALSE,FALSE,'@');
        ExcelBuf.AddColumn('Closing Balance', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn(Monthly[1], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//10
        ExcelBuf.AddColumn(Monthly[13], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn(Monthly[2], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//12
        ExcelBuf.AddColumn(Monthly[14], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//13
        ExcelBuf.AddColumn(Monthly[3], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//14
        ExcelBuf.AddColumn(Monthly[15], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//15
        ExcelBuf.AddColumn(Monthly[4], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.AddColumn(Monthly[16], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//17
        ExcelBuf.AddColumn(Monthly[5], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//18
        ExcelBuf.AddColumn(Monthly[17], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//19
        ExcelBuf.AddColumn(Monthly[6], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//20
        ExcelBuf.AddColumn(Monthly[18], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn(Monthly[7], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn(Monthly[19], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//23
        ExcelBuf.AddColumn(Monthly[8], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//24
        ExcelBuf.AddColumn(Monthly[20], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//25
        ExcelBuf.AddColumn(Monthly[9], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//26
        ExcelBuf.AddColumn(Monthly[21], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//27
        ExcelBuf.AddColumn(Monthly[10], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn(Monthly[22], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn(Monthly[11], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn(Monthly[23], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn(Monthly[12], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//32
        ExcelBuf.AddColumn(Monthly[24], FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Item."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn(Item.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn(Item."Sales Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(recItemCategory.Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(itemDim, FALSE, '', FALSE, FALSE, FALSE, '', 1);//5 //RSPL
        ExcelBuf.AddColumn(OpeningBalance, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//6
                                                                                            //>>03Mar2017
        TOpeningBalance := TOpeningBalance + OpeningBalance;
        //<<

        ExcelBuf.AddColumn(ReceiptQuantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//7
        //>>03mar2017
        TReceiptQuantity := TReceiptQuantity + ReceiptQuantity;
        //<<

        ExcelBuf.AddColumn(DespQuantity, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//8
        //>>03Mar2017
        TDespQuantity := TDespQuantity + DespQuantity;
        //<<

        //ExcelBuf.AddColumn(OtherQuantity,FALSE,'',FALSE,FALSE,FALSE,'');
        ExcelBuf.AddColumn(ClosingBalance, FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//9
        //>>03Mar2017
        TClosingBalance := TClosingBalance + ClosingBalance;
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[1], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//10
        //>>03Mar2017
        TLastMonthQuantity[1] := TLastMonthQuantity[1] + LastMonthQuantity[1];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[13], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//11
        //>>03Mar2017
        TLastMonthQuantity[13] := TLastMonthQuantity[13] + LastMonthQuantity[13];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[2], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//12
        //>>03Mar2017
        TLastMonthQuantity[2] := TLastMonthQuantity[2] + LastMonthQuantity[2];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[14], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//13
        //>>03Mar2017
        TLastMonthQuantity[14] := TLastMonthQuantity[14] + LastMonthQuantity[14];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[3], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//14
        //>>03Mar2017
        TLastMonthQuantity[3] := TLastMonthQuantity[3] + LastMonthQuantity[3];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[15], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//15
        //>>03Mar2017
        TLastMonthQuantity[15] := TLastMonthQuantity[15] + LastMonthQuantity[15];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[4], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//16
        //>>03Mar2017
        TLastMonthQuantity[4] := TLastMonthQuantity[4] + LastMonthQuantity[4];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[16], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//17
        //>>03Mar2017
        TLastMonthQuantity[16] := TLastMonthQuantity[16] + LastMonthQuantity[16];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[5], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//18
        //>>03Mar2017
        TLastMonthQuantity[5] := TLastMonthQuantity[5] + LastMonthQuantity[5];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[17], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//19
        //>>03Mar2017
        TLastMonthQuantity[17] := TLastMonthQuantity[17] + LastMonthQuantity[17];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[6], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//20
        //>>03Mar2017
        TLastMonthQuantity[6] := TLastMonthQuantity[6] + LastMonthQuantity[6];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[18], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//21
        //>>03Mar2017
        TLastMonthQuantity[18] := TLastMonthQuantity[18] + LastMonthQuantity[18];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[7], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//22
        //>>03Mar2017
        TLastMonthQuantity[7] := TLastMonthQuantity[7] + LastMonthQuantity[7];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[19], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//23
        //>>03Mar2017
        TLastMonthQuantity[19] := TLastMonthQuantity[19] + LastMonthQuantity[19];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[8], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//24
        //>>03Mar2017
        TLastMonthQuantity[8] := TLastMonthQuantity[8] + LastMonthQuantity[8];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[20], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//25
        //>>03Mar2017
        TLastMonthQuantity[20] := TLastMonthQuantity[20] + LastMonthQuantity[20];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[9], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//26
        //>>03Mar2017
        TLastMonthQuantity[9] := TLastMonthQuantity[9] + LastMonthQuantity[9];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[21], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//27
        //>>03Mar2017
        TLastMonthQuantity[21] := TLastMonthQuantity[21] + LastMonthQuantity[21];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[10], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//28
        //>>03Mar2017
        TLastMonthQuantity[10] := TLastMonthQuantity[10] + LastMonthQuantity[10];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[22], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//29
        //>>03Mar2017
        TLastMonthQuantity[22] := TLastMonthQuantity[22] + LastMonthQuantity[22];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[11], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//30
        //>>03Mar2017
        TLastMonthQuantity[11] := TLastMonthQuantity[11] + LastMonthQuantity[11];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[23], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//31
        //>>03Mar2017
        TLastMonthQuantity[23] := TLastMonthQuantity[23] + LastMonthQuantity[23];
        //<<

        ExcelBuf.AddColumn(-LastMonthQuantity[12], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//32
        //>>03Mar2017
        TLastMonthQuantity[12] := TLastMonthQuantity[12] + LastMonthQuantity[12];
        //<<

        ExcelBuf.AddColumn(LastMonthQuantity[24], FALSE, '', FALSE, FALSE, FALSE, '#####0.000', 0);//33
        //>>03Mar2017
        TLastMonthQuantity[24] := TLastMonthQuantity[24] + LastMonthQuantity[24];
        //<<
    end;

    // [Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        //ExcelBuf.AddColumn(OpeningBalance,FALSE,'',FALSE,FALSE,FALSE,'@');
        //ExcelBuf.AddColumn(OpeningBalance,FALSE,'',FALSE,FALSE,FALSE,'#####0.000',0);//6
        ExcelBuf.AddColumn(TOpeningBalance, FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//6 03Mar2017
        //ExcelBuf.AddColumn(ReceiptQuantity,FALSE,'',FALSE,FALSE,FALSE,'#####0.000',0);//7
        ExcelBuf.AddColumn(TReceiptQuantity, FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//7 03Mar2017
        //ExcelBuf.AddColumn(DespQuantity,FALSE,'',FALSE,FALSE,FALSE,'#####0.000',0);//8
        ExcelBuf.AddColumn(TDespQuantity, FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//8 03Mar2017
        //ExcelBuf.AddColumn(OtherQuantity,FALSE,'',FALSE,FALSE,FALSE,'@');
        //ExcelBuf.AddColumn(ClosingBalance,FALSE,'',FALSE,FALSE,FALSE,'#####0.000',0);//9
        ExcelBuf.AddColumn(TClosingBalance, FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//9 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[1],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//10
        ExcelBuf.AddColumn(-TLastMonthQuantity[1], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//10 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[13],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//11
        ExcelBuf.AddColumn(TLastMonthQuantity[13], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//11 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[2],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//12
        ExcelBuf.AddColumn(-TLastMonthQuantity[2], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//12 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[14],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//13
        ExcelBuf.AddColumn(TLastMonthQuantity[14], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//13 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[3],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//14
        ExcelBuf.AddColumn(-TLastMonthQuantity[3], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//14 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[15],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//15
        ExcelBuf.AddColumn(TLastMonthQuantity[15], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//15 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[4],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//16
        ExcelBuf.AddColumn(-TLastMonthQuantity[4], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//16 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[16],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//17
        ExcelBuf.AddColumn(TLastMonthQuantity[16], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//17 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[5],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//18
        ExcelBuf.AddColumn(-TLastMonthQuantity[5], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//18 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[17],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//19
        ExcelBuf.AddColumn(TLastMonthQuantity[17], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//19 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[6],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//20
        ExcelBuf.AddColumn(-TLastMonthQuantity[6], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//20 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[18],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//21
        ExcelBuf.AddColumn(TLastMonthQuantity[18], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//21 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[7],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//22
        ExcelBuf.AddColumn(-TLastMonthQuantity[7], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//22 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[19],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//23
        ExcelBuf.AddColumn(TLastMonthQuantity[19], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//23 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[8],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//24
        ExcelBuf.AddColumn(-TLastMonthQuantity[8], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//24 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[20],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//25
        ExcelBuf.AddColumn(TLastMonthQuantity[20], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//25 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[9],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//26
        ExcelBuf.AddColumn(-TLastMonthQuantity[9], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//26 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[21],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//27
        ExcelBuf.AddColumn(TLastMonthQuantity[21], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//27 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[10],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//28
        ExcelBuf.AddColumn(-TLastMonthQuantity[10], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//28 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[22],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//29
        ExcelBuf.AddColumn(TLastMonthQuantity[22], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//29 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[11],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//30
        ExcelBuf.AddColumn(-TLastMonthQuantity[11], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//30 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[23],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//31
        ExcelBuf.AddColumn(TLastMonthQuantity[23], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//31 03Mar2017
        //ExcelBuf.AddColumn(-LastMonthQuantity[12],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//32
        ExcelBuf.AddColumn(-TLastMonthQuantity[12], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//32 03Mar2017
        //ExcelBuf.AddColumn(LastMonthQuantity[24],FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//33
        ExcelBuf.AddColumn(TLastMonthQuantity[24], FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//33 03Mar2017
    end;
}


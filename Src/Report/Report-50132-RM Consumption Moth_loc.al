report 50132 "RM Consumption Moth/loc"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/RMConsumptionMothloc.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.")
                                ORDER(Ascending)
                                WHERE("Item Category Code" = FILTER('CAT01' | 'CAT04' | 'CAT05' | 'CAT06' | 'CAT07' | 'CAT08' | 'CAT09'));
            RequestFilterFields = "Location Filter", "Item Category Code";

            trigger OnAfterGetRecord()
            begin

                //>>1


                CLEAR(Qty[1]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "12stdate", "12nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[1] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[2]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "11stdate", "11nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[2] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[3]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "10stdate", "10nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[3] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[4]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "9stdate", "9nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[4] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[5]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "8stdate", "8nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[5] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[6]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "7stdate", "7nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[6] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[7]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "6stdate", "6nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[7] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[8]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "5stdate", "5nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[8] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[9]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "4stdate", "4nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[9] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[10]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "3stdate", "3nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[10] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;


                CLEAR(Qty[11]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "2stdate", "2nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[11] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(Qty[12]);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Entry Type", RecILE."Entry Type"::Consumption);
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', "1stdate", "1nddate");
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        Qty[12] += ABS(RecILE.Quantity);
                    UNTIL RecILE.NEXT = 0;

                CLEAR(AODQty);
                RecILE.RESET;
                RecILE.SETRANGE(RecILE."Item No.", Item."No.");
                RecILE.SETRANGE(RecILE."Location Code", LocationFilter);
                RecILE.SETRANGE(RecILE."Item Category Code", Item."Item Category Code");
                RecILE.SETFILTER(RecILE."Posting Date", '%1..%2', 0D, TODAY);
                IF RecILE.FINDFIRST THEN
                    REPEAT
                        AODQty += ABS(RecILE."Remaining Quantity");
                    UNTIL RecILE.NEXT = 0;

                TotQty := Qty[1] + Qty[2] + Qty[3] + Qty[4] + Qty[5] + Qty[6] + Qty[7] + Qty[8] + Qty[9] + Qty[10] + Qty[11] + Qty[12];
                MonthlyQty := TotQty / 12;

                //"3MTotQty" := (TotQty/12)*3;
                //"6MTotQty" := (TotQty/12)*6;

                "3MTotQty" := (Qty[1] + Qty[2] + Qty[3]) / 3;
                "6MTotQty" := (Qty[1] + Qty[2] + Qty[3] + Qty[4] + Qty[5] + Qty[6]) / 6;

                IF "3MTotQty" <> 0 THEN BEGIN
                    "3AODQty" := AODQty / MonthlyQty;
                END ELSE
                    "3AODQty" := 0;
                //<<1


                //>>2
                //Item, Body (2) - OnPreSection()
                "SrNo." += 1;

                CLEAR(ReMark);
                IF Item.Blocked THEN
                    ReMark := 'BLocked'
                ELSE
                    ReMark := '';

                //>>16Mar2017 DataBody

                IF PrintToExcel THEN BEGIN
                    EnterCell(NN, 1, FORMAT("SrNo."), FALSE, FALSE, '', 0);
                    EnterCell(NN, 2, Item."No.", FALSE, FALSE, '@', 1);
                    EnterCell(NN, 3, Item.Description, FALSE, FALSE, '@', 1);
                    EnterCell(NN, 4, FORMAT(Qty[1]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 5, FORMAT(Qty[2]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 6, FORMAT(Qty[3]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 7, FORMAT(Qty[4]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 8, FORMAT(Qty[5]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 9, FORMAT(Qty[6]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 10, FORMAT(Qty[7]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 11, FORMAT(Qty[8]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 12, FORMAT(Qty[9]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 13, FORMAT(Qty[10]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 14, FORMAT(Qty[11]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 15, FORMAT(Qty[12]), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 16, FORMAT(TotQty), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 17, FORMAT(MonthlyQty), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 18, FORMAT("3MTotQty"), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 19, FORMAT("6MTotQty"), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 20, FORMAT(AODQty), FALSE, FALSE, '0.000', 0);
                    Value := AODQty * Item."Unit Cost";
                    EnterCell(NN, 21, FORMAT(Value), FALSE, FALSE, '#,###0.00', 0);
                    EnterCell(NN, 22, FORMAT("3AODQty"), FALSE, FALSE, '0.000', 0);
                    EnterCell(NN, 23, ReMark, FALSE, FALSE, '@', 1);

                    NN += 1;

                END;
                //<<16Mar2017 DataBody


                /*
                
                IF PrintToExcel THEN
                BEGIN
                  ExSheet.Range('A'+FORMAT(i)).Value:="SrNo.";
                  ExSheet.Range('B'+FORMAT(i)).Value:=Item."No.";
                  ExSheet.Range('C'+FORMAT(i)).Value:=Item.Description;
                  ExSheet.Range('C'+FORMAT(i)).EntireColumn.AutoFit;
                  ExSheet.Range('D'+FORMAT(i)).Value:=FORMAT(Qty[1]);
                  ExSheet.Range('E'+FORMAT(i)).Value:=FORMAT(Qty[2]);
                  ExSheet.Range('F'+FORMAT(i)).Value:=FORMAT(Qty[3]);
                  ExSheet.Range('G'+FORMAT(i)).Value:=FORMAT(Qty[4]);
                  ExSheet.Range('H'+FORMAT(i)).Value:=FORMAT(Qty[5]);
                  ExSheet.Range('I'+FORMAT(i)).Value:=FORMAT(Qty[6]);
                  ExSheet.Range('J'+FORMAT(i)).Value:=FORMAT(Qty[7]);
                  ExSheet.Range('K'+FORMAT(i)).Value:=FORMAT(Qty[8]);
                  ExSheet.Range('L'+FORMAT(i)).Value:=FORMAT(Qty[9]);
                  ExSheet.Range('M'+FORMAT(i)).Value:=FORMAT(Qty[10]);
                  ExSheet.Range('N'+FORMAT(i)).Value:=FORMAT(Qty[11]);
                  ExSheet.Range('O'+FORMAT(i)).Value:=FORMAT(Qty[12]);
                  ExSheet.Range('P'+FORMAT(i)).Value:= FORMAT(TotQty);
                  ExSheet.Range('Q'+FORMAT(i)).Value:= FORMAT(MonthlyQty);
                  ExSheet.Range('R'+FORMAT(i)).Value:= FORMAT("3MTotQty");
                  ExSheet.Range('S'+FORMAT(i)).Value:= FORMAT("6MTotQty");
                  ExSheet.Range('T'+FORMAT(i)).Value:=FORMAT(AODQty);
                  Value := AODQty * Item."Unit Cost";
                  ExSheet.Range('U'+FORMAT(i)).Value:=FORMAT(Value);
                  ExSheet.Range('V'+FORMAT(i)).Value:=FORMAT(ROUND("3AODQty",1));
                  ExSheet.Range('W'+FORMAT(i)).Value:=ReMark;
                  i+=1;
                END;
                //<<2
                *///Commented Automation 16Mar2017

            end;

            trigger OnPreDataItem()
            begin

                //>>16Mar2017 ReportHeader
                IF PrintToExcel THEN BEGIN
                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '@', 1);
                    EnterCell(1, 23, FORMAT(TODAY, 0, 4), TRUE, FALSE, '@', 1);

                    EnterCell(2, 1, 'RM Consumption Month/LOC', TRUE, FALSE, '@', 1);
                    EnterCell(2, 23, USERID, TRUE, FALSE, '@', 1);



                    EnterCell(5, 1, 'Date Filter', TRUE, TRUE, '@', 1);
                    EnterCell(5, 2, FORMAT("1stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(5, 3, ' To ' + FORMAT("12stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
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
                    EnterCell(5, 18, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 19, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 20, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 21, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 22, '', TRUE, TRUE, '@', 1);
                    EnterCell(5, 23, '', TRUE, TRUE, '@', 1);

                    EnterCell(6, 1, 'SrNo', TRUE, TRUE, '@', 1);
                    EnterCell(6, 2, 'Item Code', TRUE, TRUE, '@', 1);
                    EnterCell(6, 3, 'Item Name', TRUE, TRUE, '@', 1);
                    EnterCell(6, 4, FORMAT("12stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 5, FORMAT("11stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 6, FORMAT("10stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 7, FORMAT("9stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 8, FORMAT("8stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 9, FORMAT("7stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 10, FORMAT("6stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 11, FORMAT("5stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 12, FORMAT("4stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 13, FORMAT("3stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 14, FORMAT("2stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 15, FORMAT("1stdate", 0, '<Month Text> <Year4>'), TRUE, TRUE, '@', 1);
                    EnterCell(6, 16, 'Total consump', TRUE, TRUE, '@', 1);
                    EnterCell(6, 17, 'Monthly Average Consptn', TRUE, TRUE, '@', 1);
                    EnterCell(6, 18, 'Average QTY for 3 month Consumption', TRUE, TRUE, '@', 1);
                    EnterCell(6, 19, 'Average QTY for 6 month Consumption', TRUE, TRUE, '@', 1);
                    EnterCell(6, 20, 'Stock   as on ' + FORMAT(TODAY), TRUE, TRUE, '@', 1);
                    EnterCell(6, 21, 'Value', TRUE, TRUE, '@', 1);
                    EnterCell(6, 22, 'Holding Stock  Levels (in months)', TRUE, TRUE, '@', 1);
                    EnterCell(6, 23, 'Remarks', TRUE, TRUE, '@', 1);

                    NN := 7;

                END;

                //<<15Mar2017 ReportHeader

                //>>1
                /*
                //Item, Header (1) - OnPreSection()
                i:=4;
                
                IF PrintToExcel THEN
                BEGIN
                  ExSheet.Range('A1:V1').Cells.Merge;
                  ExSheet.Range('A1').Value:='R.M. Planning Sheet for the Period of '+FORMAT("1stdate",0,'<Month Text> <Year4>')+' To '+
                                                                                    FORMAT("12stdate",0,'<Month Text> <Year4>') ;
                  ExSheet.Range('A1').Font.Bold:=TRUE;
                  ExSheet.Range('A1').Font.Size:=14;
                  ExSheet.Range('A1').HorizontalAlignment:=-4108;
                  ExSheet.Range('A3:W3').Font.Bold:=TRUE;
                  ExSheet.Range('A3').Value:='Sr No.';
                  ExSheet.Range('B3').Value:='Item No.';
                  ExSheet.Range('C3').Value:='Item Name';
                  ExSheet.Range('C3').EntireColumn.AutoFit;
                  ExSheet.Range('D3').Value:=FORMAT("12stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('E3').Value:=FORMAT("11stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('F3').Value:=FORMAT("10stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('G3').Value:=FORMAT("9stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('H3').Value:=FORMAT("8stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('I3').Value:=FORMAT("7stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('J3').Value:=FORMAT("6stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('K3').Value:=FORMAT("5stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('L3').Value:=FORMAT("4stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('M3').Value:=FORMAT("3stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('N3').Value:=FORMAT("2stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('O3').Value:=FORMAT("1stdate",0,'<Month Text> <Year4>');
                  ExSheet.Range('P3').Value:='Total consump';
                  ExSheet.Range('P3').EntireColumn.AutoFit;
                  ExSheet.Range('Q3').Value:='Monthly Average Consptn';
                  ExSheet.Range('Q3').EntireColumn.AutoFit;
                  ExSheet.Range('R3').Value:='Average QTY for 3 month Consumption';
                  ExSheet.Range('R3').EntireColumn.AutoFit;
                  ExSheet.Range('S3').Value:='Average QTY for 6 month Consumption';
                  ExSheet.Range('S3').EntireColumn.AutoFit;
                  ExSheet.Range('T3').Value:='Stock   as on '+FORMAT(TODAY);
                  ExSheet.Range('T3').EntireColumn.AutoFit;
                  ExSheet.Range('U3').Value:='Value';
                  ExSheet.Range('U3').EntireColumn.AutoFit;
                  ExSheet.Range('V3').Value:='Holding Stock  Levels (in months)';
                  ExSheet.Range('V3').EntireColumn.AutoFit;
                  ExSheet.Range('W3').Value:='Remark';
                END;
                //<<1
                *///Commented Automation 16Mar2017

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(Month; FromMonth)
                {
                }
                field(Year; FromYear)
                {
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

        /*
        //>>1
        
        IF PrintToExcel THEN
        BEGIN
          ExSheet.Columns.AutoFit;
          ExApp.Visible := TRUE;
        END;
        
        //<<1
        *///Commented Automation 16Mar2017

        //>>16Mar2017 RB-N
        IF PrintToExcel THEN BEGIN
            /*  ExcBuffer.CreateBook('', 'RM Monthwise Consumption');
             ExcBuffer.CreateBookAndOpenExcel('', 'RM Monthwise Consumption', '', '', USERID);
             ExcBuffer.GiveUserControl; */


            ExcBuffer.CreateNewBook('RM Monthwise Consumption');
            ExcBuffer.WriteSheet('RM Monthwise Consumption', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('RM Monthwise Consumption', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();
        END;
        //<<

    end;

    trigger OnPreReport()
    begin

        //>>1

        LocationFilter := Item.GETFILTER(Item."Location Filter");

        /*
        User.GET(USERID);
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",User."User ID");
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.RESET;
         recLocation.SETRANGE(recLocation.Code,LocationFilter);
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping1.RESET;
         CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
         CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
         CSOMapping1.SETRANGE(CSOMapping1.Value,LocationFilter);
         IF NOT(CSOMapping1.FINDFIRST) THEN
         BEGIN
          CSOMapping1.RESET;
          CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
          CSOMapping1.SETRANGE(CSOMapping1.Type,CSOMapping1.Type::"Responsibility Center");
          CSOMapping1.SETRANGE(CSOMapping1.Value,LocResC);
          IF NOT(CSOMapping1.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        *///Commented 16Mar2017

        IF (Item.GETFILTER(Item."Item Category Code") = 'CAT02') OR (Item.GETFILTER(Item."Item Category Code") = 'CAT10') THEN BEGIN
            ERROR('You can not Select This Category');
        END;


        Compnyinfo.GET();

        LocationFilter := Item.GETFILTER(Item."Location Filter");

        StartDate := DMY2DATE(1, FromMonth + 1, FromYear);
        StartDate1 := CALCDATE('<+1M>', StartDate) - 1;

        EndDate := CALCDATE('<-1Y>', StartDate);
        EndDate1 := CALCDATE('<+1M>', EndDate);

        temp := EndDate1;
        EndDate1 := StartDate1;
        StartDate1 := temp;

        "1stdate" := StartDate1;
        "1nddate" := CALCDATE('<+1M>', StartDate1) - 1;

        "2stdate" := "1nddate" + 1;
        "2nddate" := CALCDATE('<+1M>', "2stdate") - 1;

        "3stdate" := "2nddate" + 1;
        "3nddate" := CALCDATE('<+1M>', "3stdate") - 1;

        "4stdate" := "3nddate" + 1;
        "4nddate" := CALCDATE('<+1M>', "4stdate") - 1;

        "5stdate" := "4nddate" + 1;
        "5nddate" := CALCDATE('<+1M>', "5stdate") - 1;

        "6stdate" := "5nddate" + 1;
        "6nddate" := CALCDATE('<+1M>', "6stdate") - 1;

        "7stdate" := "6nddate" + 1;
        "7nddate" := CALCDATE('<+1M>', "7stdate") - 1;

        "8stdate" := "7nddate" + 1;
        "8nddate" := CALCDATE('<+1M>', "8stdate") - 1;

        "9stdate" := "8nddate" + 1;
        "9nddate" := CALCDATE('<+1M>', "9stdate") - 1;

        "10stdate" := "9nddate" + 1;
        "10nddate" := CALCDATE('<+1M>', "10stdate") - 1;

        "11stdate" := "10nddate" + 1;
        "11nddate" := CALCDATE('<+1M>', "11stdate") - 1;

        "12stdate" := "11nddate" + 1;
        "12nddate" := CALCDATE('<+1M>', "12stdate") - 1;

        /*
        IF PrintToExcel THEN
        BEGIN
          CLEAR(ExApp);
          CLEAR(ExSheet);
          CLEAR(ExBook);
          CREATE(ExApp);
          ExBook := ExApp.Workbooks.Add;
          ExSheet := ExBook.Worksheets.Add;
          ExSheet.Name := 'RM Monthwise Consumption';
          ExApp.Visible := FALSE;
        END;
        //<<1
        *///Commented Automation 16Mar2017

    end;

    var
        FromMonth: Option January,February,March,April,May,June,July,August,September,October,November,December;
        FromYear: Integer;
        StartDate: Date;
        EndDate: Date;
        "1stdate": Date;
        "1nddate": Date;
        "2stdate": Date;
        "2nddate": Date;
        "3stdate": Date;
        "3nddate": Date;
        Qty: array[12] of Decimal;
        RecILE: Record 32;
        i: Integer;
        "4stdate": Date;
        "4nddate": Date;
        "5stdate": Date;
        "5nddate": Date;
        "6stdate": Date;
        "6nddate": Date;
        "7stdate": Date;
        "7nddate": Date;
        "8stdate": Date;
        "8nddate": Date;
        "9stdate": Date;
        "9nddate": Date;
        "10stdate": Date;
        "10nddate": Date;
        "11stdate": Date;
        "11nddate": Date;
        "12stdate": Date;
        "12nddate": Date;
        TotQty: Decimal;
        "3MTotQty": Decimal;
        "6MTotQty": Decimal;
        Compnyinfo: Record 79;
        PrintToExcel: Boolean;
        "SrNo.": Integer;
        ReMark: Text[30];
        LocationFilter: Text[30];
        CSOMapping1: Record 50006;
        recLocation: Record 14;
        LocResC: Code[10];
        User: Record 2000000120;
        AODQty: Decimal;
        "3AODQty": Decimal;
        "6AODQty": Decimal;
        StartDate1: Date;
        EndDate1: Date;
        temp: Date;
        MonthlyQty: Decimal;
        Value: Decimal;
        "---16Mar2017": Integer;
        ExcBuffer: Record 370 temporary;
        NN: Integer;

    // //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NoFormat: Text[30]; CType: Option Number,Text,Date,Time)
    begin

        IF PrintToExcel THEN BEGIN
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
        END;

        //ExcBuffer.AddColumn(Value,IsFormula,CommentText,IsBold,IsItalics,IsUnderline,NumFormat,CellType)
    end;
}


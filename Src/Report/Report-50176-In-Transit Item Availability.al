report 50176 "In-Transit Item Availability"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/InTransitItemAvailability.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Transfer Line"; "Transfer Line")
        {
            DataItemTableView = SORTING("Transfer-to Code", "Document No.", "Line No.");
            RequestFilterFields = "Transfer-to Code", "Item No.";

            trigger OnAfterGetRecord()
            begin
                IF RecLoc.GET("Transfer Line"."Transfer-to Code") THEN;
                CLEAR(BRTNO);

                TransferHeader.RESET;
                TransferHeader.SETRANGE("No.", "Document No.");
                IF TransferHeader.FIND('-') THEN BEGIN

                    Location.RESET;
                    Location.SETRANGE(Location.Code, TransferHeader."Transfer-from Code");
                    IF Location.FIND('-') THEN
                        TransferFrom := Location.Name;

                    Location.RESET;
                    Location.SETRANGE(Location.Code, TransferHeader."Transfer-to Code");
                    IF Location.FIND('-') THEN
                        TransferTo := Location.Name;
                END;

                Item.RESET;
                Item.SETRANGE("No.", "Item No.");
                IF Item.FIND('-') THEN BEGIN
                    Description := Item.Description;
                    Description2 := Item.Description;
                END;


                RecTransShipHeader.RESET;
                RecTransShipHeader.SETRANGE("Transfer Order No.", "Document No.");
                IF RecTransShipHeader.FINDFIRST THEN
                    BRTNO := RecTransShipHeader."No.";

                IF Recitem.GET("Item No.") THEN;

                IF PrintToExcel THEN
                    MakeExcelDataBody;
            end;

            trigger OnPreDataItem()
            begin

                LastFieldNo := FIELDNO("Transfer-to Code");
                SETFILTER("Transfer Line"."Shipment Date", '<=%1', TillDate);
                SETRANGE("Transfer Line"."Transfer-from Code", 'IN-TRANS');

                //EBT STIVAN ---(23102012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
                CSOMapping2.RESET;
                CSOMapping2.SETRANGE(CSOMapping2."User Id", UPPERCASE(USERID));
                CSOMapping2.SETRANGE(Type, CSOMapping2.Type::Location);
                CSOMapping2.SETRANGE(CSOMapping2.Value, LocCode);
                IF CSOMapping2.FINDFIRST THEN
                    MARK := TRUE;
                BEGIN
                    IF LocCode <> '' THEN
                        "Transfer Line".SETRANGE("Transfer Line"."Transfer-to Code", LocCode);
                END;
                //EBT STIVAN ---(23102012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
            end;
        }
        dataitem("Transfer Receipt Header"; "Transfer Receipt Header")
        {
            DataItemTableView = SORTING("No.");
            dataitem("Transfer Receipt Line"; "Transfer Receipt Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));

                trigger OnAfterGetRecord()
                begin

                    Location.RESET;
                    Location.SETRANGE(Location.Code, "Transfer Receipt Header"."Transfer-from Code");
                    IF Location.FIND('-') THEN BEGIN
                        TransferFrom := Location.Name;
                    END;

                    Location.RESET;
                    Location.SETRANGE(Location.Code, "Transfer Receipt Header"."Transfer-to Code");
                    IF Location.FIND('-') THEN BEGIN
                        TransferTo := Location.Name;
                    END;

                    Item.RESET;
                    Item.SETRANGE("No.", "Transfer Receipt Line"."Item No.");
                    IF Item.FIND('-') THEN BEGIN
                        Description := Item.Description;
                        Description2 := Item.Description;
                    END;

                    CLEAR(BRTNO);
                    RecTransShipHeader.RESET;
                    RecTransShipHeader.SETRANGE("Transfer Order No.", "Transfer Receipt Line"."Transfer Order No.");
                    IF RecTransShipHeader.FINDFIRST THEN BEGIN
                        BRTNO := RecTransShipHeader."No.";
                    END;

                    IF Recitem.GET("Item No.") THEN;

                    IF PrintToExcel THEN
                        MakeExcelTrnsfReceioptBody;
                end;
            }

            trigger OnPreDataItem()
            begin

                CalcReciptDate := CALCDATE('+1D', TillDate);
                //"Transfer Receipt Header".SETFILTER("Transfer Receipt Header"."Receipt Date",'%1..%2',CalcReciptDate,TODAY);
                "Transfer Receipt Header".SETFILTER("Transfer Receipt Header"."Shipment Date", '<=%1', TillDate);
                "Transfer Receipt Header".SETFILTER("Transfer Receipt Header"."Posting Date", '%1..%2', CalcReciptDate, TODAY);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(TillDate; TillDate)
                {
                    ApplicationArea = all;
                    Caption = 'Till Date';
                }
                field(PrintToExcel; PrintToExcel)
                {
                    ApplicationArea = all;
                    Caption = 'Print to Excel';
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

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin

        fromlocfilt := "Transfer Line".GETFILTER("Transfer Line"."Transfer-from Code");
        tolocfilt := "Transfer Line".GETFILTER("Transfer Line"."Transfer-to Code");

        RecLoc1.RESET;
        IF RecLoc1.GET(fromlocfilt) THEN
            FromLoc := RecLoc1.Name;

        RecLoc1.RESET;
        IF RecLoc1.GET(tolocfilt) THEN
            ToLoc := RecLoc1.Name;
        /*
        
        //EBT STIVAN ---(23102012)--- To Filter Report as per the User Setup in CSO Mapping Table ------START
        LocCode := "Transfer Line".GETFILTER("Transfer Line"."Transfer-to Code");
        User.GET(USERID);
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",User."User ID");
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
         CLEAR(LocResC);
         recLocation.RESET;
         recLocation.SETRANGE(recLocation.Code,LocCode);
         IF recLocation.FINDFIRST THEN
         BEGIN
          LocResC := recLocation."Global Dimension 2 Code";
         END;
        
         CSOMapping1.RESET;
         CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
         CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
         CSOMapping1.SETRANGE(CSOMapping1.Value,LocCode);
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
        //EBT STIVAN ---(23102012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        */
        //>>27Mar2017 LocName
        LOCC := "Transfer Line".GETFILTER("Transfer-to Code");

        recLLL.RESET;
        recLLL.SETRANGE(Code, LOCC);
        IF recLLL.FINDFIRST THEN
            LocName := recLLL.Name;
        //<<27Mar2017 LocName



        IF PrintToExcel THEN
            MakeExcelInfo;

    end;

    var
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        TransferHeader: Record 5740;
        TransferFrom: Text[30];
        TransferTo: Text[30];
        Description: Text[100];
        Location: Record 14;
        Item: Record 27;
        ExcelBuf: Record 370 temporary;
        PrintToExcel: Boolean;
        Description2: Text[100];
        usersetup: Record 91;
        LocFilter: Text[30];
        RecLoc: Record 14;
        RecTransHeader: Record 5740;
        RecTransShipHeader: Record 5744;
        BRTNO: Code[30];
        fromlocfilt: Text[30];
        tolocfilt: Text[30];
        RecLoc1: Record 14;
        FromLoc: Text[50];
        ToLoc: Text[50];
        datefilt: Text[30];
        Recitem: Record 27;
        LocCode: Code[100];
        LocResC: Code[20];
        recLocation: Record 14;
        CSOMapping1: Record 50006;
        CSOMapping2: Record 50006;
        TillDate: Date;
        CalcReciptDate: Date;
        Text001: Label 'In-Transit Statement';
        Text000: Label 'Data';
        Text0001: Label 'Khan';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        User: Record 2000000120;
        "----24Mar2017": Integer;
        LOCC: Code[50];
        recLLL: Record 14;
        LocName: Text[100];

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('In-Transit Statement', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"In-Transit Item Availability", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        //>>27Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('In-Transit Statement', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        //<<27Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Till Date : ' + FORMAT(TillDate), FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location Name : ' + LocName, FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Transfer Order No', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item N0.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Base UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty in Sales UOM', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Qty in LT/KG', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('From Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('To Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BRT No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Transfer Price Per Unit', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('BED Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('eCess Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('SHE Cess Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Excise Amount', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Total Amount Incl. Excise', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Transfer Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line"."Shipment Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn("Transfer Line"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Description2, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Recitem."Base Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', ExcelBuf."Cell Type"::Number);//27Mar2017
        ExcelBuf.AddColumn("Transfer Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', ExcelBuf."Cell Type"::Number);//27Mar2017
        ExcelBuf.AddColumn(TransferFrom, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TransferTo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BRTNO, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Line"."Transfer Price of Base Unit", FALSE, '', FALSE, FALSE, FALSE, '0.000', ExcelBuf."Cell Type"::Number);//27Mar2017
        ExcelBuf.AddColumn("Transfer Line"."Transfer Price of Base Unit" * "Transfer Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn(/*"Transfer Line"."BED Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn(/*"Transfer Line"."eCess Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn(/*"Transfer Line"."SHE Cess Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn(/*"Transfer Line"."Excise Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn("Transfer Line".Amount +/*"Transfer Line"."BED Amount" +
                           "Transfer Line"."eCess Amount" + "Transfer Line"."SHE Cess Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelTrnsfReceioptBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Transfer Receipt Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Header"."Shipment Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.AddColumn("Transfer Receipt Line"."Item No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Description2, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(Recitem."Base Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Line"."Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn("Transfer Receipt Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn(TransferFrom, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(TransferTo, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(BRTNO, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Transfer Receipt Line"."Unit Price" /
        "Transfer Receipt Line"."Qty. per Unit of Measure", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn(("Transfer Receipt Line"."Unit Price" / "Transfer Receipt Line"."Qty. per Unit of Measure")
        * "Transfer Receipt Line"."Quantity (Base)", FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//27Mar2017
        ExcelBuf.AddColumn(/*"Transfer Receipt Line"."BED Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn(/*"Transfer Receipt Line"."eCess Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn(/*"Transfer Receipt Line"."SHE Cess Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn(/*"Transfer Receipt Line"."Excise Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
        ExcelBuf.AddColumn("Transfer Receipt Line".Amount + /*"Transfer Receipt Line"."BED Amount" +
                           "Transfer Receipt Line"."eCess Amount" + "Transfer Receipt Line"."SHE Cess Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27Mar2017
    end;

    //  //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000 ,Text001,COMPANYNAME,USERID);
        /*  ExcelBuf.CreateBookAndOpenExcel('', Text000, '', '', '');
         ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook(Text000);
        ExcelBuf.WriteSheet(Text000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        ERROR('');
    end;

    ///  //[Scope('Internal')]
    procedure MakeExcelDataFooter()
    begin
    end;

    // //[Scope('Internal')]
    procedure MakeExcelGroupdata()
    begin
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelGroupFooter()
    begin
    end;
}


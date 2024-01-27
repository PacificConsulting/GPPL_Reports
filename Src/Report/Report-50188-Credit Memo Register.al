report 50188 "Credit Memo Register"
{
    // 07Aug2017 :: RB-N, GST Related Fields added in the Reports
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 11Sep2017   RB-N         Filtering of Report on the basis of Location GSTNo.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CreditMemoRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
        {
            DataItemTableView = WHERE("Source Code" = FILTER(= 'SALES'),
                                      "Return Order No." = FILTER(''));
            RequestFilterFields = "Posting Date", "Location Code";
            dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Type = FILTER("G/L Account"),
                                          "No." = FILTER(<> 74012210 | ''));
                dataitem("Sales Comment Line"; "Sales Comment Line")
                {
                    DataItemLink = "No." = FIELD("No.");
                    DataItemLinkReference = "Sales Cr.Memo Header";
                    DataItemTableView = SORTING("Document Type", "No.", "Document Line No.", "Line No.")
                                        WHERE("Document Type" = FILTER("Posted Credit Memo"));
                }

                trigger OnAfterGetRecord()
                begin

                    LCount += 1; //07Mar2017 RB-N
                    //>>1
                    //Fahim 24-11-21 for Sale person name
                    CLEAR(SalPerName);
                    Cust.RESET;
                    Cust.SETRANGE("No.", "Sales Cr.Memo Header"."Sell-to Customer No.");
                    IF Cust.FINDFIRST THEN;
                    SalPerCode := Cust."Salesperson Code";

                    RecSalPeople.RESET;
                    RecSalPeople.SETRANGE(Code, SalPerCode);
                    IF RecSalPeople.FINDFIRST THEN;
                    SalPerName := RecSalPeople.Name;
                    SalPerZone := RecSalPeople."Zone Code";

                    recResCenter.RESET;
                    recResCenter.SETRANGE(recResCenter.Code, "Sales Cr.Memo Line"."Responsibility Center");
                    IF recResCenter.FINDFIRST THEN;

                    recDimVal.RESET;
                    recDimVal.SETRANGE(recDimVal."Dimension Code", 'DIVISION');
                    recDimVal.SETRANGE(recDimVal.Code, "Sales Cr.Memo Line"."Shortcut Dimension 1 Code");
                    IF recDimVal.FINDFIRST THEN;

                    AccName := '';
                    GLacc.RESET;
                    GLacc.SETRANGE(GLacc."No.", "Sales Cr.Memo Line"."No.");
                    IF GLacc.FINDFIRST THEN
                        AccName := GLacc.Name;

                    CLEAR(Cess);
                    /* recPostedStrOrdDetails.RESET;
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Document Type", recPostedStrOrdDetails."Document Type"::"Credit Memo");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Line"."Document No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Item No.", "Sales Cr.Memo Line"."No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Sales Cr.Memo Line"."Line No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                    recPostedStrOrdDetails.SETFILTER(recPostedStrOrdDetails."Tax/Charge Code", '%1|%2', 'CESS', 'C1');
                    IF recPostedStrOrdDetails.FINDFIRST THEN
                        Cess := recPostedStrOrdDetails.Amount; */

                    //RSPLSUM19Apr21>>
                    /* recPostedStrOrdDetails.RESET;
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Document Type", recPostedStrOrdDetails."Document Type"::"Credit Memo");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Line"."Document No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Item No.", "Sales Cr.Memo Line"."No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Sales Cr.Memo Line"."Line No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                    recPostedStrOrdDetails.SETFILTER(recPostedStrOrdDetails."Tax/Charge Group", '%1', 'CESS');
                    IF recPostedStrOrdDetails.FINDSET THEN
                        REPEAT
                            Cess += recPostedStrOrdDetails.Amount;
                        UNTIL recPostedStrOrdDetails.NEXT = 0; */
                    //RSPLSUM19Apr21<<

                    CLEAR(AddlDuty);
                    /*  recPostedStrOrdDetails.RESET;
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Document Type", recPostedStrOrdDetails."Document Type"::"Credit Memo");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Line"."Document No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Item No.", "Sales Cr.Memo Line"."No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Sales Cr.Memo Line"."Line No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                     recPostedStrOrdDetails.SETFILTER(recPostedStrOrdDetails."Tax/Charge Code", '%1', 'ADDL TAX');
                     IF recPostedStrOrdDetails.FINDFIRST THEN
                         AddlDuty := recPostedStrOrdDetails.Amount; */

                    TCess := TCess + Cess;
                    TAddlDuty := TAddlDuty + AddlDuty;

                    //<<1


                    //>>2

                    //Sales Cr.Memo Line, Body (1) - OnPreSection()
                    TQty := TQty + "Sales Cr.Memo Line".Quantity;
                    TAmt := TAmt + "Sales Cr.Memo Line"."Line Amount";
                    TVAT := TVAT + 0;//"Sales Cr.Memo Line"."Tax Amount";

                    //>>07Aug2017 GST
                    CLEAR(vCGST);
                    CLEAR(vSGST);
                    CLEAR(vIGST);
                    CLEAR(vUTGST);
                    CLEAR(vCGSTRate);
                    CLEAR(vSGSTRate);
                    CLEAR(vIGSTRate);
                    CLEAR(vUTGSTRate);
                    CLEAR(GSTPer);
                    DetailGSTEntry.RESET;
                    DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                    //DetailGSTEntry.SETRANGE(Type, Type);
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT

                            GSTPer += DetailGSTEntry."GST %";
                            IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                vCGST += ABS(DetailGSTEntry."GST Amount");
                                vCGSTRate := DetailGSTEntry."GST %";
                            END;

                            IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                vSGST += ABS(DetailGSTEntry."GST Amount");
                                vSGSTRate := DetailGSTEntry."GST %";

                            END;

                            IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                vUTGST += ABS(DetailGSTEntry."GST Amount");
                                vUTGSTRate := DetailGSTEntry."GST %";

                            END;


                            IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                vIGST += ABS(DetailGSTEntry."GST Amount");
                                vIGSTRate := DetailGSTEntry."GST %";
                                //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");
                            END;

                        UNTIL DetailGSTEntry.NEXT = 0;

                    TCGST += vCGST;
                    TTCGST += vCGST;

                    TSGST += vSGST;
                    TTSGST += vSGST;

                    TIGST += vIGST;
                    TTIGST += vIGST;

                    TUTGST += vUTGST;
                    TTUTGST += vUTGST;
                    //<<07Aug2017 GST


                    IF PrintToExcel AND PrintDetail THEN
                        MakeEcelDataBody;
                    //<<2



                    //>>3

                    //>>18Apr2017 GroupingData
                    IF NOT PrintDetail THEN BEGIN
                        GQty += "Sales Cr.Memo Line".Quantity;
                        GLineAmt += "Sales Cr.Memo Line"."Line Amount";
                        GVatAmt += 0;// "Sales Cr.Memo Line"."Tax Amount";
                        GAddDuty += AddlDuty;
                        GCess += Cess;
                        GFAmt += "Line Amount" + 0;//"Tax Amount" + AddlDuty + Cess + "Total GST Amount";
                    END;
                    //<<18Apr2017 GroupingData

                    //Sales Cr.Memo Line, GroupFooter (2) - OnPreSection()
                    //>>07Mar2017 Document Grouping
                    IF LCount = CrCount THEN BEGIN
                        IF PrintToExcel THEN
                            MakeEcelDataGroupFooter;
                        LCount := 0;

                        GQty := 0;
                        GLineAmt := 0;
                        GVatAmt := 0;
                        GAddDuty := 0;
                        GCess := 0;
                        GFAmt := 0;

                    END;
                    //<<3
                end;

                trigger OnPreDataItem()
                begin
                    LCount := 0; //07Mar2017 RB-N
                    CrCount := COUNT; //07Mar2017 RB-N
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //>>1

                Cust.GET("Sales Cr.Memo Header"."Sell-to Customer No.");

                recPSIH.RESET;
                recPSIH.SETRANGE(recPSIH."No.", "Applies-to Doc. No.");
                IF recPSIH.FIND('-') THEN
                    AppliesDocDate := recPSIH."Posting Date"
                ELSE
                    AppliesDocDate := 0D;


                CLEAR(TotalSalesComment);
                CLEAR(TotalSalesCommentNew);
                vi := 1;
                recSalesCommentLine.RESET;
                recSalesCommentLine.SETRANGE(recSalesCommentLine."Document Type", recSalesCommentLine."Document Type"::"Posted Credit Memo");
                recSalesCommentLine.SETRANGE(recSalesCommentLine."No.", "Sales Cr.Memo Header"."No.");
                IF recSalesCommentLine.FINDFIRST THEN
                    REPEAT
                        TotalSalesComment[vi] := recSalesCommentLine.Comment;
                        TotalSalesCommentNew := TotalSalesCommentNew + ' ' + recSalesCommentLine.Comment;
                        vi += 1;
                    UNTIL recSalesCommentLine.NEXT = 0;
                //<<1

                LocGSTNo := '';//RB-N 11Sep2017
                recLoc.RESET;
                recLoc.SETRANGE(recLoc.Code, "Location Code");
                IF recLoc.FINDFIRST THEN BEGIN
                    LocGSTNo := recLoc."GST Registration No.";

                END;

                //>>Location GST 07Aug2017
                State07.RESET;
                IF State07.GET(recLoc."State Code") THEN;

                //>>RB-N 11Sep2017 GSTNo. Filter

                IF TempGSTNo <> '' THEN BEGIN
                    IF TempGSTNo <> LocGSTNo THEN
                        CurrReport.SKIP;

                END;
                //<<RB-N 11Sep2017 GSTNo. Filter
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                //Sales Cr.Memo Header, Footer (2) - OnPreSection()
                IF PrintToExcel THEN
                    MakeEcelDataFooter;
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
                field("GST Register No."; TempGSTNo)
                {
                    ApplicationArea = all;
                    TableRelation = "GST Registration Nos.";
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = all;
                }
                field("Print Details"; PrintDetail)
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

        CompInfo.GET;

        IF PrintToExcel THEN
            MakeExcelInfo;
        //<<1
    end;

    var
        Cust: Record 18;
        recPSIH: Record 112;
        AppliesDocDate: Date;
        ExcelBuf: Record 370 temporary;
        CompInfo: Record 79;
        PrintToExcel: Boolean;
        PrintDetail: Boolean;
        recLoc: Record 14;
        recResCenter: Record 5714;
        recDimVal: Record 349;
        //recPostedStrOrdDetails: Record 13798;
        Cess: Decimal;
        TCess: Decimal;
        AddlDuty: Decimal;
        TAddlDuty: Decimal;
        TQty: Decimal;
        TAmt: Decimal;
        TVAT: Decimal;
        AccName: Text[60];
        GLacc: Record 15;
        recGLentry: Record 17;
        recSalesCommentLine: Record 44;
        TotalSalesComment: array[5] of Text[250];
        TotalSalesCommentNew: Text[250];
        NOofDays: Integer;
        Text0000: Label 'Data';
        Text0001: Label 'Credit Memo Register';
        Text0002: Label 'Company Name';
        Text0003: Label 'Report No.';
        Text0004: Label 'Report Name';
        Text0005: Label 'USER Id';
        Text0006: Label 'Date';
        "-------07Mar2017": Integer;
        CrCount: Integer;
        LCount: Integer;
        GQty: Decimal;
        GLineAmt: Decimal;
        GVatAmt: Decimal;
        GAddDuty: Decimal;
        GCess: Decimal;
        GFAmt: Decimal;
        vi: Integer;
        "---07Aug2017---GST": Integer;
        State07: Record State;
        GSTPer: Decimal;
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        vUTGST: Decimal;
        vUTGSTRate: Decimal;
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        TCGST: Decimal;
        TTCGST: Decimal;
        TSGST: Decimal;
        TTSGST: Decimal;
        TIGST: Decimal;
        TTIGST: Decimal;
        TUTGST: Decimal;
        TTUTGST: Decimal;
        "------11Sep2017": Integer;
        TempGSTNo: Code[15];
        LocGSTNo: Code[15];
        RecSalPeople: Record 13;
        SalPerName: Text;
        SalPerCode: Text;
        SalPerZone: Text;

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet; //07Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0002), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Credit Memo Register', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0003), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Credit Memo Register", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    //  //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text0000,CompInfo.Name,Text0001,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>07Mar2017 RB-N
        /* ExcelBuf.CreateBook('', Text0001);
        ExcelBuf.CreateBookAndOpenExcel('', Text0001, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text0001);
        ExcelBuf.WriteSheet(Text0001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text0001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>07Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//28 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//30 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//31 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 21-04-23 for Product HSN Code
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//32

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Credit Memo Register', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//28 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//30 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//31 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 21-04-23 for Product HSN Code
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//32

        ExcelBuf.NewRow;

        //<<
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//21
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//23
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31 07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 21-04-23 for Product HSN Code

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Credit Note No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Party Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Name of the Party', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Our Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Our Invoice Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Location State Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9 07Aug2017
        ExcelBuf.AddColumn('Location GSTIN', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10 07Aug2017
        ExcelBuf.AddColumn('Res. Center Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('Res. Center Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('G/L Account No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('G/L Account Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Line Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('VAT Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('Addl. Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('GST %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23 07Aug2017
        ExcelBuf.AddColumn('CGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24 07Aug2017
        ExcelBuf.AddColumn('SGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25 07Aug2017
        ExcelBuf.AddColumn('IGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26 07Aug2017
        ExcelBuf.AddColumn('UTGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27 07Aug2017
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30
        ExcelBuf.AddColumn('GST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31 07Aug2017
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('IRN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//33 //AR 08122020
        ExcelBuf.AddColumn('Sales Person', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn('Sales Person Zone', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn('HSN/SAC Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 21-04-23 for HSN/SAC Code
    end;

    // //[Scope('Internal')]
    procedure MakeEcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Cust."Full Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(AppliesDocDate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//6
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(recLoc.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(State07."State Code (GST Reg. No.)", FALSE, '', FALSE, FALSE, FALSE, '', 1);//9 //07Aug2017
        ExcelBuf.AddColumn(recLoc."GST Registration No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//10 //07Aug2017
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Responsibility Center", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn(recResCenter.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn(recDimVal.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn(AccName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn("Sales Cr.Memo Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//16
        ExcelBuf.AddColumn("Sales Cr.Memo Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '0.000', 0);//17
        //>>07Mar2017
        GQty += "Sales Cr.Memo Line".Quantity;
        //<<

        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Line Amount", FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//18
        //>>07Mar2017
        GLineAmt += "Sales Cr.Memo Line"."Line Amount";
        //<<

        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Tax %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//19
        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Tax Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//20
        //>>07Mar2017
        GVatAmt += 0;// "Sales Cr.Memo Line"."Tax Amount";
        //<<

        ExcelBuf.AddColumn(AddlDuty, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//21
        //>>07Mar2017
        GAddDuty += AddlDuty;
        //<<

        ExcelBuf.AddColumn(Cess, FALSE, '', FALSE, FALSE, FALSE, '#,####0.00', 0);//22
        //>>07Mar2017
        GCess += Cess;
        //<<

        ExcelBuf.AddColumn(GSTPer, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//23 07Aug2017
        ExcelBuf.AddColumn(vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//24 07Aug2017
        ExcelBuf.AddColumn(vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//25 07Aug2017
        ExcelBuf.AddColumn(vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//26 07Aug2017
        ExcelBuf.AddColumn(vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//27 07Aug2017
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Line Amount" + /*"Sales Cr.Memo Line"."Tax Amount"*/0 + AddlDuty
        + Cess +/* "Sales Cr.Memo Line"."Total GST Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,#0.00', 0);//28
        //>>07Mar2017
        GFAmt += "Sales Cr.Memo Line"."Line Amount" + 0;// "Sales Cr.Memo Line"."Tax Amount" + AddlDuty + Cess + "Sales Cr.Memo Line"."Total GST Amount";
        //<<

        ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/'', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/'', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn(FORMAT(Cust."GST Registration No."), FALSE, '', FALSE, FALSE, FALSE, '@', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//32

        "Sales Cr.Memo Header".CALCFIELDS("Sales Cr.Memo Header".IRN);//AR 081220
        ExcelBuf.AddColumn("Sales Cr.Memo Header".IRN, FALSE, '', TRUE, FALSE, FALSE, '', 1);//33 //AR 081220
        ExcelBuf.AddColumn(SalPerName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn(SalPerZone, FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."HSN/SAC Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 21-04-23 for HSN/SAC Code
    end;

    // //[Scope('Internal')]
    procedure MakeEcelDataGroupFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Cust."Full Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Cr.Memo Header"."Applies-to Doc. No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(AppliesDocDate, FALSE, '', TRUE, FALSE, FALSE, '', 2);//6
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(recLoc.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn(State07."State Code (GST Reg. No.)", FALSE, '', TRUE, FALSE, FALSE, '', 1);//9 07Aug2017
        ExcelBuf.AddColumn(recLoc."GST Registration No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//10 07Aug2017
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."Responsibility Center", FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn(recResCenter.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn(recDimVal.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//16
        //ExcelBuf.AddColumn("Sales Cr.Memo Line".Quantity,FALSE,'',TRUE,FALSE,FALSE,'#####0.000',0);//15
        ExcelBuf.AddColumn(GQty, FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//17 //07Mar2017
        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Line Amount",FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//16
        ExcelBuf.AddColumn(GLineAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//18 //07Mar2017
        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."Tax %"*/0, FALSE, '', TRUE, FALSE, FALSE, '0.00', 0);//19
        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Tax Amount",FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//18
        ExcelBuf.AddColumn(GVatAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//20 //07Mar2017
        //ExcelBuf.AddColumn(AddlDuty,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//19
        ExcelBuf.AddColumn(GAddDuty, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//21 //07Mar2017
        //ExcelBuf.AddColumn(Cess,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//20
        ExcelBuf.AddColumn(GCess, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//22 //07Mar2017

        ExcelBuf.AddColumn(/*"Sales Cr.Memo Line"."GST %"*/0, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//23 07Aug2017
        ExcelBuf.AddColumn(TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//24 07Aug2017
        ExcelBuf.AddColumn(TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//25 07Aug2017
        ExcelBuf.AddColumn(TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//26 07Aug2017
        ExcelBuf.AddColumn(TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);//27 07Aug2017
        TCGST := 0; //07Aug2017
        TSGST := 0; //07Aug2017
        TIGST := 0; //07Aug2017
        TUTGST := 0; //07Aug2017

        //ExcelBuf.AddColumn("Sales Cr.Memo Line"."Line Amount" + "Sales Cr.Memo Line"."Tax Amount" + AddlDuty
        //+ Cess,FALSE,'',TRUE,FALSE,FALSE,'#,####0.00',0);//21

        ExcelBuf.AddColumn(GFAmt, FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//28 //07Mar2017
        ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/0, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/0, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30
        ExcelBuf.AddColumn(FORMAT(Cust."GST Registration No."), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//31 //07Aug2017
        ExcelBuf.AddColumn(TotalSalesCommentNew, FALSE, '', TRUE, FALSE, FALSE, '', 1);//32
                                                                                       /*ExcelBuf.AddColumn(TotalSalesComment[1],FALSE,'',TRUE,FALSE,FALSE,'',1);//32
                                                                                       IF TotalSalesComment[2] <> '' THEN
                                                                                         NewColumn(2);
                                                                                       IF TotalSalesComment[3] <> '' THEN
                                                                                         NewColumn(3);
                                                                                       IF TotalSalesComment[4] <> '' THEN
                                                                                         NewColumn(4); */

        "Sales Cr.Memo Header".CALCFIELDS("Sales Cr.Memo Header".IRN);//AR 081220
        ExcelBuf.AddColumn("Sales Cr.Memo Header".IRN, FALSE, '', TRUE, FALSE, FALSE, '', 1);//33 //AR 081220
        ExcelBuf.AddColumn(SalPerName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn(SalPerZone, FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn("Sales Cr.Memo Line"."HSN/SAC Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 23-04-23 for Sale person zone

    end;

    // //[Scope('Internal')]
    procedure MakeEcelDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 21-04-23 for HSN/SAC Code


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//16
        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, TRUE, '0.000', 0);//17
        ExcelBuf.AddColumn(TAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//19
        ExcelBuf.AddColumn(TVAT, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//20
        ExcelBuf.AddColumn(TAddlDuty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//21
        ExcelBuf.AddColumn(TCess, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//22
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//23
        ExcelBuf.AddColumn(TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//24
        ExcelBuf.AddColumn(TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//25
        ExcelBuf.AddColumn(TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//26
        ExcelBuf.AddColumn(TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//27
        ExcelBuf.AddColumn(TAmt + TVAT + TAddlDuty + TCess + TTCGST + TTSGST + TTIGST + TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,####0.00', 0);//28
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 21-04-23 for HSN/SAC Code
    end;

    local procedure NewColumn(i: Integer)
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 2);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#####0.000', 0);//15 //07Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//16 //07Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#####0.00', 0);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//18 //07Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//19 //07Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//20 //07Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,####0.00', 0);//21 //07Mar2017
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
        ExcelBuf.AddColumn(TotalSalesComment[i], FALSE, '', TRUE, FALSE, FALSE, '', 1);//32
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//33
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 24-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 21-04-23 for HSN/SAC Code
    end;
}


report 50189 "Debit Memo Register"
{
    // 07Aug2017 :: RB-N, GST Related Fields added in the Reports
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 11Sep2017   RB-N         Filtering of Report on the basis of Location GSTNo.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/DebitMemoRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;



    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = WHERE("Debit Memo" = FILTER(true));
            RequestFilterFields = "Posting Date", "Location Code";
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Type = FILTER("G/L Account"),
                                          "No." = FILTER(<> 74012210 | ''));

                trigger OnAfterGetRecord()
                begin
                    DLCount += 1;

                    //>>1

                    Cust.RESET;
                    Cust.SETRANGE(Cust."No.", "Sales Invoice Header"."Sell-to Customer No.");
                    IF Cust.FINDFIRST THEN;



                    ResCentre.RESET;
                    ResCentre.SETRANGE(ResCentre.Code, "Sales Invoice Header"."Responsibility Center");
                    IF ResCentre.FINDFIRST THEN;

                    AccName := '';
                    GLacc.RESET;
                    GLacc.SETRANGE(GLacc."No.", "Sales Invoice Line"."No.");
                    IF GLacc.FINDFIRST THEN BEGIN
                        AccName := GLacc.Name;
                    END;


                    CLEAR(Cess);
                    /*  recPostedStrOrdDetails.RESET;
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Document Type", recPostedStrOrdDetails."Document Type"::Invoice);
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Invoice Line"."Document No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Item No.", "Sales Invoice Line"."No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Sales Invoice Line"."Line No.");
                     recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                     recPostedStrOrdDetails.SETFILTER(recPostedStrOrdDetails."Tax/Charge Code", '%1|%2', 'CESS', 'C1');
                     IF recPostedStrOrdDetails.FINDFIRST THEN
                         Cess := recPostedStrOrdDetails.Amount; */

                    CLEAR(AddlDuty);
                    /* recPostedStrOrdDetails.RESET;
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Document Type", recPostedStrOrdDetails."Document Type"::Invoice);
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Invoice Line"."Document No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Item No.", "Sales Invoice Line"."No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Line No.", "Sales Invoice Line"."Line No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                    recPostedStrOrdDetails.SETFILTER(recPostedStrOrdDetails."Tax/Charge Code", '%1', 'ADDL TAX');
                    IF recPostedStrOrdDetails.FINDFIRST THEN
                        AddlDuty := recPostedStrOrdDetails.Amount; */

                    TCess := TCess + Cess;
                    TAddlDuty := TAddlDuty + AddlDuty;
                    //<<1

                    //>>2

                    //Sales Invoice Line, Body (1) - OnPreSection()
                    TQty := TQty + "Sales Invoice Line".Quantity;
                    TAmt := TAmt + "Sales Invoice Line"."Line Amount";
                    TVAT := TVAT + 0;//"Sales Invoice Line"."Tax Amount";

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

                    //>>18Apr2017 GroupingData
                    IF NOT PrintDetail THEN BEGIN
                        GQty += "Sales Invoice Line".Quantity;
                        GLineAmt += "Sales Invoice Line"."Line Amount";
                        GVatAmt += 0;//"Sales Invoice Line"."Tax Amount";
                        GAddDuty += AddlDuty;
                        GCess += Cess;
                        GFAmt += "Line Amount" + 0;//"Tax Amount" + AddlDuty + Cess + "Total GST Amount";
                    END;

                    //<<18Apr2017 GroupingData

                    //>>3

                    //Sales Invoice Line, GroupFooter (2) - OnPreSection()
                    //>>08Mar2017
                    IF DLCount = LCount THEN BEGIN
                        IF PrintToExcel THEN
                            MakeEcelDataGroupFooter;
                        DLCount := 0;

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
                    LCount := COUNT; //08Mar2017
                    DLCount := 0; //08Mar2017
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                CLEAR(TotalSalesComment);
                recSalesCommentLine.RESET;
                recSalesCommentLine.SETRANGE(recSalesCommentLine."Document Type", recSalesCommentLine."Document Type"::"Posted Invoice");
                recSalesCommentLine.SETRANGE(recSalesCommentLine."No.", "Sales Invoice Header"."No.");
                IF recSalesCommentLine.FINDFIRST THEN
                    REPEAT
                        TotalSalesComment := TotalSalesComment + recSalesCommentLine.Comment;
                    UNTIL recSalesCommentLine.NEXT = 0;

                //Fahim 29-11-21 for Sale person name
                CLEAR(SalPerName);
                Cust.RESET;
                Cust.SETRANGE("No.", "Sales Invoice Header"."Sell-to Customer No.");
                IF Cust.FINDFIRST THEN;
                SalPerCode := Cust."Salesperson Code";

                RecSalPeople.RESET;
                RecSalPeople.SETRANGE(Code, SalPerCode);
                IF RecSalPeople.FINDFIRST THEN;
                SalPerName := RecSalPeople.Name;
                SalPerZone := RecSalPeople."Zone Code";

                //<<1

                LocGSTNo := '';//RB-N 11Sep2017
                Loc.RESET;
                Loc.SETRANGE(Loc.Code, "Sales Invoice Header"."Location Code");
                IF Loc.FINDFIRST THEN BEGIN
                    LocGSTNo := Loc."GST Registration No.";//RB-N 11Sep2017
                END;

                //>>Location GST 07Aug2017
                State07.RESET;
                IF State07.GET(Loc."State Code") THEN;


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
                //Sales Invoice Header, Footer (2) - OnPreSection()
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
        Loc: Record 14;
        ResCentre: Record 5714;
        ExcelBuf: Record 370 temporary;
        CompInfo: Record 79;
        PrintToExcel: Boolean;
        PrintDetail: Boolean;
        AccName: Text[60];
        GLacc: Record 15;
        // recPostedStrOrdDetails: Record 13798;
        Cess: Decimal;
        TCess: Decimal;
        AddlDuty: Decimal;
        TAddlDuty: Decimal;
        recSalesCommentLine: Record 44;
        TotalSalesComment: Text[500];
        TQty: Decimal;
        TAmt: Decimal;
        TVAT: Decimal;
        Text0000: Label 'Data';
        Text0001: Label 'Debit Memo Register';
        Text0002: Label 'Company Name';
        Text0003: Label 'Report No.';
        Text0004: Label 'Report Name';
        Text0005: Label 'USER Id';
        Text0006: Label 'Date';
        "-----08Mar2017": Integer;
        DLCount: Integer;
        LCount: Integer;
        GQty: Decimal;
        GLineAmt: Decimal;
        GVatAmt: Decimal;
        GAddDuty: Decimal;
        GCess: Decimal;
        GFAmt: Decimal;
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
        "--": Integer;
        RecSalPeople: Record 13;
        SalPerName: Text;
        SalPerCode: Text;
        SalPerZone: Text;

    ////[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//08Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0002), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Debit Memo Register', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0003), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Debit Memo Register", FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 1);
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

        //>>08Mar2017 RB-N
        /*  ExcelBuf.CreateBook('', Text0001);
         ExcelBuf.CreateBookAndOpenExcel('', Text0001, '', '', USERID);
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text0001);
        ExcelBuf.WriteSheet(Text0001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text0001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>08Mar2017
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//28 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//30 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 29-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 29-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 21-04-23 for HSN/SAC Code
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//31

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Debit Memo Register', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//23 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//24 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//25 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//26 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//27 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//28 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//29 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//30 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 29-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 29-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 21-04-23 for HSN/SAC Code
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//31
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//24 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//25 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//26 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//27 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//28 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//29 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//30 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//31 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//32 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 29-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 29-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 21-04-23 for HSN/SAC Code
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Debit Memo No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Party Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Name of the Party', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Location Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Location Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Location StateCode', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7 //07Aug2017
        ExcelBuf.AddColumn('Location GSTIN', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8 //07Aug2017
        ExcelBuf.AddColumn('Res. Center Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('Res. Center Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('Division', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('External Doc. No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('G/L Account No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('G/L Account Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Quantity', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
        ExcelBuf.AddColumn('Line Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//17
        ExcelBuf.AddColumn('Tax Type', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//18
        ExcelBuf.AddColumn('VAT Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//19
        ExcelBuf.AddColumn('Addl. Duty', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//20
        ExcelBuf.AddColumn('Cess', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//21
        ExcelBuf.AddColumn('GST %', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//22
        ExcelBuf.AddColumn('CGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//23
        ExcelBuf.AddColumn('SGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//24
        ExcelBuf.AddColumn('IGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//25
        ExcelBuf.AddColumn('UTGST', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//26
        ExcelBuf.AddColumn('Gross Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//27
        ExcelBuf.AddColumn('TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//28
        ExcelBuf.AddColumn('CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//29
        ExcelBuf.AddColumn('GST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//30 //07Aug2017
        ExcelBuf.AddColumn('Comment', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//31
        ExcelBuf.AddColumn('IRN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//32
        ExcelBuf.AddColumn('Sales Person', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 29-11-21 for Sale person name
        ExcelBuf.AddColumn('Sales Person Zone', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 29-11-21 for Sale person zone
        ExcelBuf.AddColumn('HSN/SAC Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//Fahim 21-04-23 for HSN/SAC Code
    end;

    //  //[Scope('Internal')]
    procedure MakeEcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Sales Invoice Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Cust."Full Name", FALSE, '', FALSE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Header"."Location Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(Loc.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(State07."State Code (GST Reg. No.)", FALSE, '', FALSE, FALSE, FALSE, '', 1);//7 //07Aug2017
        ExcelBuf.AddColumn(Loc."GST Registration No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//8 //07Aug2017
        ExcelBuf.AddColumn("Sales Invoice Header"."Responsibility Center", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn(ResCentre.Name, FALSE, '', FALSE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn("Sales Invoice Header"."Shortcut Dimension 1 Code", FALSE, '', FALSE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn("Sales Invoice Header"."External Document No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn("Sales Invoice Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//13
        ExcelBuf.AddColumn(AccName, FALSE, '', FALSE, FALSE, FALSE, '', 1);//14
        ExcelBuf.AddColumn("Sales Invoice Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//15
        ExcelBuf.AddColumn("Sales Invoice Line".Quantity, FALSE, '', FALSE, FALSE, FALSE, '####0.000', 0);//16
        //>>08Mar2017
        GQty += "Sales Invoice Line".Quantity;
        //<<

        ExcelBuf.AddColumn("Sales Invoice Line"."Line Amount", FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//17
        //>>08Mar2017
        GLineAmt += "Sales Invoice Line"."Line Amount";
        //<<

        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Tax %"*/0, FALSE, '', FALSE, FALSE, FALSE, '####0.00', 0);//18
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Tax Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//19
        //>>08Mar2017
        GVatAmt += 0;// "Sales Invoice Line"."Tax Amount";
        //<<

        ExcelBuf.AddColumn(AddlDuty, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//20
        //>>08Mar2017
        GAddDuty += AddlDuty;
        //<<

        ExcelBuf.AddColumn(Cess, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//21
        //>>08Mar2017
        GCess += Cess;
        //<<

        ExcelBuf.AddColumn(GSTPer, FALSE, '', FALSE, FALSE, FALSE, '0.00', 0);//22
        ExcelBuf.AddColumn(vCGST, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//23
        ExcelBuf.AddColumn(vSGST, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//24
        ExcelBuf.AddColumn(vIGST, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//25
        ExcelBuf.AddColumn(vUTGST, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//26
        ExcelBuf.AddColumn("Sales Invoice Line"."Line Amount" + /*"Sales Invoice Line"."Tax Amount"*/0 + AddlDuty + Cess
        + /*"Sales Invoice Line"."Total GST Amount"*/0, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//27
        //>>08Mar2017
        GFAmt += "Sales Invoice Line"."Line Amount" + /*"Sales Invoice Line"."Tax Amount"*/0 + AddlDuty + Cess + /*"Sales Invoice Line"."Total GST Amount"*/0;
        //<<

        ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/'', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//28
        ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/'', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//29
        ExcelBuf.AddColumn(FORMAT(Cust."GST Registration No."), FALSE, '', FALSE, FALSE, FALSE, '@', 1);//30 //07Aug2017
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//31


        "Sales Invoice Header".CALCFIELDS("Sales Invoice Header".IRN);//32 //AR  08-12-2020
        ExcelBuf.AddColumn("Sales Invoice Header".IRN, FALSE, '', TRUE, FALSE, FALSE, '', 1);//32 //AR  08-12-2020
    end;

    // //[Scope('Internal')]
    procedure MakeEcelDataGroupFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Sales Invoice Header"."No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn("Sales Invoice Header"."Posting Date", FALSE, '', TRUE, FALSE, FALSE, '', 2);//2
        ExcelBuf.AddColumn("Sales Invoice Header"."Sell-to Customer No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(Cust."Full Name", FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn("Sales Invoice Header"."Location Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//5
        ExcelBuf.AddColumn(Loc.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//6
        ExcelBuf.AddColumn(State07."State Code (GST Reg. No.)", FALSE, '', TRUE, FALSE, FALSE, '', 1);//7
        ExcelBuf.AddColumn(Loc."GST Registration No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//8

        ExcelBuf.AddColumn("Sales Invoice Header"."Responsibility Center", FALSE, '', TRUE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn(ResCentre.Name, FALSE, '', TRUE, FALSE, FALSE, '', 1);//8
        ExcelBuf.AddColumn("Sales Invoice Header"."Shortcut Dimension 1 Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//9
        ExcelBuf.AddColumn("Sales Invoice Header"."External Document No.", FALSE, '', TRUE, FALSE, FALSE, '', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//13
        //ExcelBuf.AddColumn("Sales Invoice Line".Quantity,FALSE,'',TRUE,FALSE,FALSE,'####0.000',0);//14
        ExcelBuf.AddColumn(GQty, FALSE, '', TRUE, FALSE, FALSE, '####0.000', 0);//14 //08Mar2017
        //ExcelBuf.AddColumn("Sales Invoice Line"."Line Amount",FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//15
        ExcelBuf.AddColumn(GLineAmt, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//15 //08Mar2017
        ExcelBuf.AddColumn(/*"Sales Invoice Line"."Tax %"*/0, FALSE, '', TRUE, FALSE, FALSE, '####0.00', 0);//16
        //ExcelBuf.AddColumn("Sales Invoice Line"."Tax Amount",FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//17
        ExcelBuf.AddColumn(GVatAmt, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//17 //08Mar2017

        //ExcelBuf.AddColumn(AddlDuty,FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//18
        ExcelBuf.AddColumn(GAddDuty, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//18 //08Mar2017
        //ExcelBuf.AddColumn(Cess,FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//19
        ExcelBuf.AddColumn(GCess, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//19 //08Mar2017

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//22 //07Aug2017
        ExcelBuf.AddColumn(TCGST, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//23 //07Aug2017
        ExcelBuf.AddColumn(TSGST, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//24 //07Aug2017
        ExcelBuf.AddColumn(TIGST, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//25 //07Aug2017
        ExcelBuf.AddColumn(TUTGST, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//26 //07Aug2017

        TCGST := 0;//07Aug2017
        TSGST := 0;//07Aug2017
        TIGST := 0;//07Aug2017
        TUTGST := 0;//07Aug2017
        //ExcelBuf.AddColumn("Sales Invoice Line"."Line Amount" + "Sales Invoice Line"."Tax Amount" + AddlDuty + Cess
        //,FALSE,'',TRUE,FALSE,FALSE,'#,###0.00',0);//20
        ExcelBuf.AddColumn(GFAmt, FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//20 //08mar2017
        ExcelBuf.AddColumn(/*FORMAT(Cust."T.I.N. No.")*/'', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//21
        ExcelBuf.AddColumn(/*FORMAT(Cust."C.S.T. No.")*/'', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//22
        ExcelBuf.AddColumn(FORMAT(Cust."GST Registration No."), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//30 //07Aug2017
        ExcelBuf.AddColumn(TotalSalesComment, FALSE, '', TRUE, FALSE, FALSE, '', 1);//23
        "Sales Invoice Header".CALCFIELDS("Sales Invoice Header".IRN);//32 //AR  08-12-2020
        ExcelBuf.AddColumn("Sales Invoice Header".IRN, FALSE, '', TRUE, FALSE, FALSE, '', 1);//32 //AR  08-12-2020
        ExcelBuf.AddColumn(SalPerName, FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 29-11-21 for Sale person name
        ExcelBuf.AddColumn(SalPerZone, FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 29-11-21 for Sale person zone
        ExcelBuf.AddColumn("Sales Invoice Line"."HSN/SAC Code", FALSE, '', TRUE, FALSE, FALSE, '', 1);//Fahim 21-04-23 for HSN/SAC Code
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
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 29-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 29-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 21-04-23 for HSN/SAC Code

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
        ExcelBuf.AddColumn(TQty, FALSE, '', TRUE, FALSE, TRUE, '####0.000', 0);
        ExcelBuf.AddColumn(TAmt, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(TVAT, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);
        ExcelBuf.AddColumn(TAddlDuty, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);
        ExcelBuf.AddColumn(TCess, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn(TTCGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//07Aug2017
        ExcelBuf.AddColumn(TTSGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//07Aug2017
        ExcelBuf.AddColumn(TTIGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//07Aug2017
        ExcelBuf.AddColumn(TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//07Aug2017
        ExcelBuf.AddColumn(TAmt + TVAT + TAddlDuty + TCess + TTCGST + TTSGST + TTIGST + TTUTGST, FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 29-11-21 for Sale person name
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 29-11-21 for Sale person zone
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//Fahim 21-04-23 for HSN/SAC Code
    end;
}


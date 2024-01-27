report 50057 "Collection Report"
{
    // 
    // Date          Version       Remarks
    // ---------------------------------------------------------------------------
    // 09Apr2019     RB-N          InvoiceDate,DueDate & No.of Days
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CollectionReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Bank Account Ledger Entry"; "Bank Account Ledger Entry")
        {
            DataItemTableView = SORTING("Posting Date");
            RequestFilterFields = "Posting Date", "Global Dimension 1 Code";
            dataitem("Cust. Ledger Entry"; "Cust. Ledger Entry")
            {
                DataItemLink = "Document No." = FIELD("Document No.");
                DataItemTableView = SORTING("Entry No.");
                dataitem("<Cust. Ledger Entry1>"; "Cust. Ledger Entry")
                {
                    DataItemLink = "Closed by Entry No." = FIELD("Entry No.");
                    DataItemTableView = SORTING("Entry No.");

                    trigger OnAfterGetRecord()
                    begin

                        //>>1

                        //<Cust. Ledger Entry1>, Body (1) - OnPreSection()
                        //IF PrintToExcel THEN
                        // MakeExcelDataBody2;
                        //<<1
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    "Cust. Ledger Entry".CALCFIELDS("Cust. Ledger Entry"."Amount (LCY)");//RSPLSUM 23Dec2020
                    //RSPLSUM 22Dec2020>>
                    IF ("Bank Account Ledger Entry"."Bal. Account Type" <> "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
                        ("Bank Account Ledger Entry"."Bal. Account No." = '') THEN BEGIN//RSPLSUM 23Dec2020
                        IF PrintToExcel THEN
                            MakeExcelDataBody;
                    END;//RSPLSUM 23Dec2020
                    //RSPLSUM 22Dec2020<<

                    DtldCustLedgEntry1.SETCURRENTKEY("Cust. Ledger Entry No.");
                    DtldCustLedgEntry1.SETRANGE("Cust. Ledger Entry No.", "Entry No.");
                    DtldCustLedgEntry1.SETRANGE("Entry Type", DtldCustLedgEntry1."Entry Type"::Application);
                    DtldCustLedgEntry1.SETRANGE(Unapplied, FALSE);
                    IF DtldCustLedgEntry1.FIND('-') THEN BEGIN
                        //REPEAT
                        IF DtldCustLedgEntry1."Cust. Ledger Entry No." =
                           DtldCustLedgEntry1."Applied Cust. Ledger Entry No."
                        THEN BEGIN
                            DtldCustLedgEntry2.INIT;
                            DtldCustLedgEntry2.SETCURRENTKEY("Applied Cust. Ledger Entry No.", "Entry Type");
                            DtldCustLedgEntry2.SETRANGE(
                              "Applied Cust. Ledger Entry No.", DtldCustLedgEntry1."Applied Cust. Ledger Entry No.");
                            DtldCustLedgEntry2.SETRANGE("Entry Type", DtldCustLedgEntry2."Entry Type"::Application);
                            DtldCustLedgEntry2.SETRANGE(Unapplied, FALSE);
                            IF DtldCustLedgEntry2.FIND('-') THEN
                                REPEAT
                                    IF DtldCustLedgEntry2."Cust. Ledger Entry No." <>
                                       DtldCustLedgEntry2."Applied Cust. Ledger Entry No."
                                    THEN BEGIN
                                        RecCLE.SETCURRENTKEY("Entry No.");
                                        RecCLE.SETRANGE("Entry No.", DtldCustLedgEntry2."Cust. Ledger Entry No.");
                                        IF RecCLE.FIND('-') THEN BEGIN
                                            //MESSAGE('%1',RecCLE."Document No.");
                                            //MARK(TRUE);
                                            IF PrintToExcel THEN
                                                MakeExcelDataBody2(RecCLE);
                                        END;
                                    END;
                                UNTIL DtldCustLedgEntry2.NEXT = 0;
                        END ELSE BEGIN
                            RecCLE.SETCURRENTKEY("Entry No.");
                            RecCLE.SETRANGE("Entry No.", DtldCustLedgEntry1."Applied Cust. Ledger Entry No.");
                            IF RecCLE.FIND('-') THEN BEGIN
                                //MARK(TRUE);
                                //MESSAGE('%1',RecCLE."Document No.");
                                IF PrintToExcel THEN
                                    MakeExcelDataBody2(RecCLE);
                            END;
                        END;
                        //UNTIL DtldCustLedgEntry1.NEXT = 0;
                    END;
                end;

                trigger OnPostDataItem()
                begin
                    //RSPLSUM 23Dec2020>>
                    IF ("Bank Account Ledger Entry"."Bal. Account Type" <> "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
                        ("Bank Account Ledger Entry"."Bal. Account No." = '') THEN BEGIN//RSPLSUM 23Dec2020
                        IF PrintToExcel THEN
                            MakeExcelForFooter;
                    END;//RSPLSUM 23Dec2020
                    //RSPLSUM 23Dec2020<<
                end;

                trigger OnPreDataItem()
                begin
                    TAmt := 0;//RSPLSUM 23Dec2020

                    IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
                        ("Bank Account Ledger Entry"."Bal. Account No." <> '') THEN BEGIN//RSPLSUM 23Dec2020
                        SETRANGE("Cust. Ledger Entry"."Customer No.", "Bank Account Ledger Entry"."Bal. Account No.");//RSPLSUM 23Dec2020
                        SETRANGE("Cust. Ledger Entry"."Transaction No.", "Bank Account Ledger Entry"."Transaction No.");//RSPLSUM 23Dec2020
                    END;//RSPLSUM 23Dec2020
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1
                SalesPersonCode := '';
                SalesPersonName := '';
                CLEAR(ZoneCode);//RSPLSUM 28Aug2020
                IF "Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer THEN BEGIN
                    recCustomer.RESET;
                    recCustomer.SETRANGE(recCustomer."No.", "Bank Account Ledger Entry"."Bal. Account No.");
                    IF recCustomer.FINDFIRST THEN BEGIN
                        SalesPersonCode := recCustomer."Salesperson Code";
                    END;

                    recSalespersonPurchaser.RESET;
                    recSalespersonPurchaser.SETRANGE(recSalespersonPurchaser.Code, SalesPersonCode);
                    IF recSalespersonPurchaser.FINDFIRST THEN BEGIN
                        SalesPersonName := recSalespersonPurchaser.Name;
                        ZoneCode := recSalespersonPurchaser."Zone Code";//RSPLSUM 28Aug2020
                    END;
                END;


                IF CustSelection <> '' THEN BEGIN
                    CustLedger.RESET;
                    CustLedger.SETRANGE(CustLedger."Document No.", "Bank Account Ledger Entry"."Document No.");
                    IF CustLedger.FINDFIRST THEN
                        IF CustLedger."Customer No." <> CustSelection THEN
                            CurrReport.SKIP;
                END;


                CustLedger.RESET;
                CustLedger.SETRANGE(CustLedger."Document No.", "Bank Account Ledger Entry"."Document No.");
                IF NOT CustLedger.FINDFIRST THEN
                    CurrReport.SKIP;

                IF RespnCenter <> '' THEN BEGIN
                    IF ("Bank Account Ledger Entry"."Bal. Account No." <> '') THEN BEGIN
                        Customer.RESET;
                        IF Customer.GET("Bank Account Ledger Entry"."Bal. Account No.") THEN BEGIN
                            IF Customer."Responsibility Center" <> RespnCenter THEN
                                CurrReport.SKIP;
                        END;
                    END
                    ELSE BEGIN
                        IF ("Bank Account Ledger Entry"."Bal. Account No." = '') THEN BEGIN
                            CustLedger.RESET;
                            CustLedger.SETRANGE(CustLedger."Document No.", "Bank Account Ledger Entry"."Document No.");
                            IF CustLedger.FINDFIRST THEN BEGIN
                                IF CustLedger."Global Dimension 2 Code" <> RespnCenter THEN
                                    CurrReport.SKIP;
                            END;
                        END
                        ELSE BEGIN
                            IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::"G/L Account")
                               AND ("Bank Account Ledger Entry"."Bal. Account No." = '') THEN BEGIN
                                CustLedger.RESET;
                                CustLedger.SETRANGE(CustLedger."Document No.", "Bank Account Ledger Entry"."Document No.");
                                IF CustLedger.FINDFIRST THEN BEGIN
                                    IF CustLedger."Global Dimension 2 Code" <> RespnCenter THEN
                                        CurrReport.SKIP;
                                END;
                            END;
                        END;
                    END;
                END;

                CustPostingGroup := '';
                IF "Bank Account Ledger Entry"."Bal. Account No." <> '' THEN BEGIN
                    Customer.SETFILTER(Customer."No.", "Bank Account Ledger Entry"."Bal. Account No.");
                    IF Customer.FIND('-') THEN
                        CustNo := Customer."No.";
                    custName := Customer."Full Name";
                    CustPostingGroup := CustLedger."Customer Posting Group";
                END
                ELSE
                    IF "Bank Account Ledger Entry"."Bal. Account No." = '' THEN BEGIN
                        CustLedger.RESET;
                        CustLedger.SETRANGE(CustLedger."Document No.", "Bank Account Ledger Entry"."Document No.");
                        IF CustLedger.FINDFIRST THEN BEGIN
                            Customer.SETFILTER(Customer."No.", CustLedger."Customer No.");
                            IF Customer.FIND('-') THEN
                                CustNo := Customer."No.";
                            custName := Customer."Full Name";
                            CustPostingGroup := CustLedger."Customer Posting Group";
                        END;
                    END;

                responsiblityrec.SETRANGE(responsiblityrec.Code, Customer."Responsibility Center");
                IF responsiblityrec.FINDFIRST THEN
                    "respo.ctr.name" := responsiblityrec.Name;

                CustLedger.RESET;
                CustLedger.SETRANGE(CustLedger."Document No.", "Bank Account Ledger Entry"."Document No.");
                IF CustLedger.FINDFIRST THEN
                    recDimValue.RESET;
                recDimValue.SETRANGE(recDimValue."Dimension Code", 'DIVISION');
                recDimValue.SETRANGE(recDimValue.Code, CustLedger."Global Dimension 1 Code");
                IF recDimValue.FINDFIRST THEN;

                CLEAR(Region);
                CLEAR(RegionCode);
                recCust.RESET;
                recCust.SETRANGE(recCust."No.", CustNo);
                IF recCust.FINDFIRST THEN BEGIN
                    recState.RESET;
                    recState.SETRANGE(recState.Code, recCust."State Code");
                    IF recState.FINDFIRST THEN BEGIN
                        Region := 'Region' + ': ' + '';// recState."Region Code";
                        RegionCode := '';// recState."Region Code";
                    END;
                END;
                //<<1

                //>>2

                //Bank Account Ledger Entry, Body (3) - OnPreSection()
                IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
                    ("Bank Account Ledger Entry"."Bal. Account No." <> '') THEN BEGIN//RSPLSUM 23Dec2020
                    IF PrintToExcel THEN
                        MakeExcelDataBody;
                END;//RSPLSUM 23Dec2020
                //<<2
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                //Bank Account Ledger Entry, Footer (4) - OnPreSection()
                IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
                    ("Bank Account Ledger Entry"."Bal. Account No." <> '') THEN BEGIN//RSPLSUM 23Dec2020
                    IF PrintToExcel THEN
                        MakeExcelForFooter;
                END;//RSPLSUM 23Dec2020
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
                field("Responsibility Centre"; RespnCenter)
                {
                    TableRelation = "Responsibility Center";
                }
                field(Customer; CustSelection)
                {
                    TableRelation = Customer;
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

        reqFilter := "Bank Account Ledger Entry".GETFILTERS;

        IF RespnCenter <> '' THEN BEGIN
            RespCentreFilter := RespnCenter;
            recRespCentr.GET(RespCentreFilter);
            RespCentreLocationFilter := recRespCentr."Location Code";
        END;

        IF CustSelection <> '' THEN BEGIN
            recCust.GET(CustSelection);
            CustomerCodeResCentreFilter := recCust."Responsibility Center";
            recRespCentr.GET(CustomerCodeResCentreFilter);
            RespCentreLocationFilter := recRespCentr."Location Code";
        END;

        /*
        Memberof.RESET;
        Memberof.SETRANGE(Memberof."User ID",USERID);
        Memberof.SETFILTER(Memberof."Role ID",'%1|%2','SUPER','REPORT VIEW');
        IF NOT(Memberof.FINDFIRST) THEN
        BEGIN
        
         CSOMapping1.RESET;
         CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
         CSOMapping1.SETRANGE(Type,CSOMapping1.Type::Location);
         CSOMapping1.SETRANGE(CSOMapping1.Value,RespCentreLocationFilter);
         IF NOT(CSOMapping1.FINDFIRST) THEN
         BEGIN
          CSOMapping1.RESET;
          CSOMapping1.SETRANGE(CSOMapping1."User Id",UPPERCASE(USERID));
          CSOMapping1.SETRANGE(CSOMapping1.Type,CSOMapping1.Type::"Responsibility Center");
          CSOMapping1.SETFILTER(CSOMapping1.Value,'%1|%2',RespCentreFilter,CustomerCodeResCentreFilter);
          IF NOT(CSOMapping1.FINDFIRST) THEN
          ERROR ('You are not allowed to run this report other than your location');
         END;
        END;
        //EBT STIVAN ---(03122012)--- To Filter Report as per the User Setup in CSO Mapping Table --------END
        *///Commented MemberofFunction 30Mar2017

        IF PrintToExcel THEN
            MakeExcelInfo;

        //<<1

    end;

    var
        Customer: Record 18;
        CustNo: Code[10];
        custName: Text[70];
        CustPostingGroup: Code[20];
        reqFilter: Text[150];
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        responsiblityrec: Record 5714;
        "respo.ctr.name": Text[50];
        postingdate: Text[100];
        chquedate: Text[100];
        ResponsibilityCenter: Record 5714;
        RespnCenter: Code[10];
        Customer1: Record 18;
        CustLedger: Record 21;
        CustSelection: Code[20];
        recCustomer: Record 18;
        SalesPersonName: Text[50];
        SalesPersonCode: Code[10];
        recSalespersonPurchaser: Record 13;
        RespCentreFilter: Code[100];
        CustomerCodeResCentreFilter: Code[100];
        CSOMapping1: Record 50006;
        recCust: Record 18;
        recCustPostingGroup: Record 92;
        recRespCentr: Record 5714;
        RespCentreLocationFilter: Code[100];
        Region: Text[30];
        RegionCode: Text[30];
        recState: Record State;
        recDimValue: Record 349;
        Text000: Label 'Data';
        Text001: Label 'Collection Report';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "---30Mar2017": Integer;
        DateText: Text[50];
        TAmt: Decimal;
        ZoneCode: Code[10];
        TCSEntry_Rec: Record "TCS Entry";
        TCSAmount: Decimal;
        EntryNo_: Integer;
        CustLedgerEntryNew: Record 21;
        DetailedCustLedgEntryNew: Record 379;
        AmountLCY: Decimal;
        DtldCustLedgEntry1: Record 379;
        DtldCustLedgEntry2: Record 379;
        RecCLE: Record 21;
        RecCustNew: Record 18;

    //  [Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//30Mar2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Collection Report', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"Collection Report", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelHeader;
    end;

    // [Scope('Internal')]
    procedure CreateExcelbook()
    begin
        /*
        ExcelBuf.CreateBook;
        ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        ExcelBuf.GiveUserControl;
        ERROR('');
        *///30Mar2017

        //>>30Mar2017 RB-N
        /* ExcelBuf.CreateBook('', Text001);
        ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

        //<<

    end;

    // [Scope('Internal')]
    procedure MakeExcelHeader()
    begin
        //ExcelBuf.NewRow;
        //ExcelBuf.AddColumn(reqFilter,FALSE,'',TRUE,FALSE,FALSE,'@',1);
        //>>30Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
                                                                       //
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//12 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//13 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//14 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//15 RSPLSUM 28Aug2020
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('COLLECTION REPORT', FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//10
                                                                       //
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//12 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//13 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//14 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);//15 RSPLSUM 28Aug2020
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//16
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        //<<30Mar2017

        ExcelBuf.NewRow;
        DateText := "Bank Account Ledger Entry".GETFILTER("Posting Date");//30Mar2017

        ExcelBuf.AddColumn('DateFilter : ' + DateText, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//10
                                                                      //
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//12 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//13 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//14 09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, TRUE, '@', 1);//16 RSPLSUM 28Aug2020
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Date', FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('Voucher No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('Customer Code', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('Customer Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('Customer Posting Group', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('Res. Center', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('Region', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('Bill Details', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('Invoice Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9 09Apr2019
        ExcelBuf.AddColumn('Due Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10 09Apr2019
                                                                             //
        ExcelBuf.AddColumn('Invoice Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('TCS Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Applied Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        //
        ExcelBuf.AddColumn('No. of Days', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11 09Apr2019
        ExcelBuf.AddColumn('Cheque No.', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('Cheque Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14
        ExcelBuf.AddColumn('Salesperson Name', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('Zone', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16
    end;

    //   [Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(FORMAT("Bank Account Ledger Entry"."Posting Date"), FALSE, '', FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.AddColumn("Bank Account Ledger Entry"."Document No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
            ("Bank Account Ledger Entry"."Bal. Account No." <> '') THEN BEGIN//RSPLSUM 23Dec2020
            ExcelBuf.AddColumn(CustNo, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(custName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(CustPostingGroup, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        END;//RSPLSUM 23Dec2020

        IF ("Bank Account Ledger Entry"."Bal. Account Type" <> "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
            ("Bank Account Ledger Entry"."Bal. Account No." = '') THEN BEGIN//RSPLSUM 23Dec2020
            ExcelBuf.AddColumn("Cust. Ledger Entry"."Customer No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);//RSPLSUM 22Dec2020
            //RSPLSUM 22Dec2020>>
            RecCustNew.RESET;
            IF RecCustNew.GET("Cust. Ledger Entry"."Customer No.") THEN;
            ExcelBuf.AddColumn(RecCustNew.Name, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
            ExcelBuf.AddColumn(RecCustNew."Customer Posting Group", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
            //RSPLSUM 22Dec2020<<
        END;//RSPLSUM 23Dec2020


        ExcelBuf.AddColumn("respo.ctr.name", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(RegionCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        //
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1); //09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1); //09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1); //09Apr2019
        ExcelBuf.AddColumn("Bank Account Ledger Entry"."Cheque No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(FORMAT("Bank Account Ledger Entry"."Cheque Date"), FALSE, '', FALSE, FALSE, FALSE, '', 2);

        IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
            ("Bank Account Ledger Entry"."Bal. Account No." <> '') THEN BEGIN//RSPLSUM 23Dec2020
            ExcelBuf.AddColumn("Bank Account Ledger Entry".Amount, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);
            //>>30Mar2017 TotalAmount
            TAmt += "Bank Account Ledger Entry".Amount;
            //<<30Mar2017 TotalAmount
        END;//RSPLSUM 23Dec2020

        IF ("Bank Account Ledger Entry"."Bal. Account Type" <> "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
            ("Bank Account Ledger Entry"."Bal. Account No." = '') THEN BEGIN//RSPLSUM 23Dec2020
            ExcelBuf.AddColumn(ABS("Cust. Ledger Entry"."Amount (LCY)"), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//RSPLSUM 23Dec2020
            //>>30Mar2017 TotalAmount
            TAmt += "Cust. Ledger Entry"."Amount (LCY)";//RSPLSUM 23dec2020
            //<<30Mar2017 TotalAmount
        END;//RSPLSUM 23Dec2020

        ExcelBuf.AddColumn(SalesPersonName, FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(ZoneCode, FALSE, '', FALSE, FALSE, FALSE, '@', 1);//RSPLSUM 28Aug2020
    end;

    //  [Scope('Internal')]
    procedure MakeExcelDataBody2(CustLedgEnt: Record "21")
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(CustLedgEnt."Document No.", FALSE, '', FALSE, FALSE, FALSE, '@', 1);
        ExcelBuf.AddColumn(CustLedgEnt."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//09Apr2019
        ExcelBuf.AddColumn(CustLedgEnt."Due Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//09Apr2019
                                                                                          //
        CustLedgEnt.CALCFIELDS(CustLedgEnt."Amount (LCY)");
        CLEAR(TCSAmount);
        TCSEntry_Rec.RESET;
        TCSEntry_Rec.SETRANGE("Document No.", CustLedgEnt."Document No.");
        IF TCSEntry_Rec.FINDFIRST THEN
            REPEAT
                TCSAmount += TCSEntry_Rec."TCS Amount";
            UNTIL TCSEntry_Rec.NEXT = 0;

        CustLedgerEntryNew.RESET;
        CustLedgerEntryNew.SETRANGE("Document No.", "Bank Account Ledger Entry"."Document No.");
        IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
            ("Bank Account Ledger Entry"."Bal. Account No." <> '') THEN//RSPLSUM 23Dec2020
            CustLedgerEntryNew.SETRANGE("Transaction No.", "Bank Account Ledger Entry"."Transaction No.");//RSPLSUM 23Dec2020
        IF CustLedgerEntryNew.FINDFIRST THEN BEGIN
            EntryNo_ := CustLedgerEntryNew."Entry No.";
        END;

        IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
            ("Bank Account Ledger Entry"."Bal. Account No." <> '') THEN BEGIN//RSPLSUM 23Dec2020
            CLEAR(AmountLCY);
            DetailedCustLedgEntryNew.RESET;
            DetailedCustLedgEntryNew.SETRANGE("Applied Cust. Ledger Entry No.", EntryNo_);
            DetailedCustLedgEntryNew.SETRANGE("Cust. Ledger Entry No.", CustLedgEnt."Entry No.");
            IF DetailedCustLedgEntryNew.FINDFIRST THEN
                AmountLCY := DetailedCustLedgEntryNew."Amount (LCY)";
        END;//RSPLSUM 23Dec2020

        //RSPLSUM 23Dec2020>>
        IF ("Bank Account Ledger Entry"."Bal. Account Type" <> "Bank Account Ledger Entry"."Bal. Account Type"::Customer) AND//RSPLSUM 23Dec2020
            ("Bank Account Ledger Entry"."Bal. Account No." = '') THEN BEGIN//RSPLSUM 23Dec2020
            CLEAR(AmountLCY);
            DetailedCustLedgEntryNew.RESET;
            DetailedCustLedgEntryNew.SETRANGE("Applied Cust. Ledger Entry No.", "Cust. Ledger Entry"."Entry No.");
            DetailedCustLedgEntryNew.SETRANGE("Cust. Ledger Entry No.", CustLedgEnt."Entry No.");
            IF DetailedCustLedgEntryNew.FINDFIRST THEN
                AmountLCY := DetailedCustLedgEntryNew."Amount (LCY)";
        END;//RSPLSUM 23Dec2020<<

        ExcelBuf.AddColumn(CustLedgEnt."Amount (LCY)", FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);
        ExcelBuf.AddColumn(TCSAmount, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);
        ExcelBuf.AddColumn(AmountLCY, FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);
        //
        IF "Bank Account Ledger Entry"."Cheque Date" <> 0D THEN
            ExcelBuf.AddColumn(FORMAT(CustLedgEnt."Due Date" - "Bank Account Ledger Entry"."Cheque Date"), FALSE, '', FALSE, FALSE, FALSE, '@', 1)
        ELSE
            ExcelBuf.AddColumn(FORMAT(CustLedgEnt."Due Date"), FALSE, '', FALSE, FALSE, FALSE, '@', 1);
    end;

    // [Scope('Internal')]
    procedure MakeExcelForFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
                                                                     //
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        //
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13 //09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14 //09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15 //09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16 RSPLSUM 28Aug2020
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10 //09Apr2019
                                                                     //
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        //
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11 //09Apr2019
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12 //09Apr2019
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        //ExcelBuf.AddColumn("Bank Account Ledger Entry".Amount,FALSE,'',TRUE,FALSE,TRUE,'#,###0.00',0);//11
        ExcelBuf.AddColumn(ABS(TAmt), FALSE, '', TRUE, FALSE, TRUE, '#,###0.00', 0);//14 //30Mar2017
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//16 RSPLSUM 28Aug2020
    end;
}


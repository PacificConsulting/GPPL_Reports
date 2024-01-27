report 50182 "Trial Balance GST"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 13Sep2017   RB-N         New Report Development from Existing Report#6--TrailBalance
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TrialBalanceGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Trial Balance GST';

    dataset
    {
        dataitem("G/L Account"; "G/L Account")
        {
            CalcFields = "Net Change", "Balance at Date";
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.", "Account Type", "Global Dimension 1 Filter", "Global Dimension 2 Filter";
            column(FORMAT_TODAY_0_4_; FORMAT(TODAY, 0, 4))
            {
            }
            column(STRSUBSTNO_Text000_PeriodText_; STRSUBSTNO(Text000, PeriodText))
            {
            }
            column(CurrReport_PAGENO; CurrReport.PAGENO)
            {
            }
            column(COMPANYNAME; COMPANYNAME)
            {
            }
            column(USERID; USERID)
            {
            }
            column(G_L_Account__TABLECAPTION__________GLFilter; TABLECAPTION + ': ' + GLFilter)
            {
            }
            column(GLFilter; GLFilter)
            {
            }
            column(PrintToExcel; PrintToExcel)
            {
            }
            column(EmptyString; '')
            {
            }
            column(EmptyString_Control12; '')
            {
            }
            column(TotalDebitNetChange; TotalDebitNetChange)
            {
            }
            column(TotalCreditNetChange; TotalCreditNetChange)
            {
            }
            column(TotalDebitBalanceAtDate; TotalDebitBalanceAtDate)
            {
            }
            column(TotalCreditBalanceAtDate; TotalCreditBalanceAtDate)
            {
            }
            column(G_L_Account_No_; "No.")
            {
            }
            column(Trial_BalanceCaption; Trial_BalanceCaptionLbl)
            {
            }
            column(CurrReport_PAGENOCaption; CurrReport_PAGENOCaptionLbl)
            {
            }
            column(Net_ChangeCaption; Net_ChangeCaptionLbl)
            {
            }
            column(BalanceCaption; BalanceCaptionLbl)
            {
            }
            column(G_L_Account___No__Caption; FIELDCAPTION("No."))
            {
            }
            column(PADSTR_____G_L_Account__Indentation___2___G_L_Account__NameCaption; PADSTR_____G_L_Account__Indentation___2___G_L_Account__NameCaptionLbl)
            {
            }
            column(G_L_Account___Net_Change_Caption; G_L_Account___Net_Change_CaptionLbl)
            {
            }
            column(G_L_Account___Net_Change__Control22Caption; G_L_Account___Net_Change__Control22CaptionLbl)
            {
            }
            column(G_L_Account___Balance_at_Date_Caption; G_L_Account___Balance_at_Date_CaptionLbl)
            {
            }
            column(G_L_Account___Balance_at_Date__Control24Caption; G_L_Account___Balance_at_Date__Control24CaptionLbl)
            {
            }
            column(TOTALSCaption; TOTALSCaptionLbl)
            {
            }
            dataitem(DataItem5444; Integer)
            {
                DataItemTableView = SORTING(Number)
                                    WHERE(Number = CONST(1));
                column(G_L_Account___No__; "G/L Account"."No.")
                {
                }
                column(PADSTR_____G_L_Account__Indentation___2___G_L_Account__Name; PADSTR('', "G/L Account".Indentation * 2) + "G/L Account".Name)
                {
                }
                column(G_L_Account___Net_Change_; "G/L Account"."Net Change")
                {
                }
                column(G_L_Account___Net_Change__Control22; -"G/L Account"."Net Change")
                {
                    AutoFormatType = 1;
                }
                column(G_L_Account___Balance_at_Date_; "G/L Account"."Balance at Date")
                {
                }
                column(G_L_Account___Balance_at_Date__Control24; -"G/L Account"."Balance at Date")
                {
                    AutoFormatType = 1;
                }
                column(NetChangeDr; TempAmt)
                {
                }
                column(NetChangeCr; -TempAmt)
                {
                }
                column(BalanceDr; BalanceAmt)
                {
                }
                column(BalanceCr; -BalanceAmt)
                {
                }
                column(G_L_Account___No___Control25; "G/L Account"."No.")
                {
                }
                column(PADSTR_____G_L_Account__Indentation___2_____G_L_Account__Name; PADSTR('', "G/L Account".Indentation * 2) + "G/L Account".Name)
                {
                }
                column(G_L_Account___Net_Change__Control27; "G/L Account"."Net Change")
                {
                }
                column(G_L_Account___Net_Change__Control28; -"G/L Account"."Net Change")
                {
                    AutoFormatType = 1;
                }
                column(G_L_Account___Balance_at_Date__Control29; "G/L Account"."Balance at Date")
                {
                }
                column(G_L_Account___Balance_at_Date__Control30; -"G/L Account"."Balance at Date")
                {
                    AutoFormatType = 1;
                }
                column(G_L_Account___Account_Type_; FORMAT("G/L Account"."Account Type", 0, 2))
                {
                }
                column(No__of_Blank_Lines; "G/L Account"."No. of Blank Lines")
                {
                }
                column(GroupNum; GroupNum)
                {
                }
                column(Integer_Number; Number)
                {
                }
                dataitem(BlankLineRepeater; Integer)
                {
                    column(BlankLineNo; BlankLineNo)
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin
                        IF BlankLineNo = 0 THEN
                            CurrReport.BREAK;

                        BlankLineNo -= 1;
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    BlankLineNo := "G/L Account"."No. of Blank Lines" + 1;
                end;
            }

            trigger OnAfterGetRecord()
            begin
                /*
                CALCFIELDS("Net Change","Balance at Date");
                IF PrintToExcel THEN
                  MakeExcelDataBody;
                
                IF ("G/L Account"."Account Type" = "G/L Account"."Account Type"::Posting) THEN BEGIN
                  IF "Balance at Date" > 0 THEN
                    TotalDebitBalanceAtDate := TotalDebitBalanceAtDate + "G/L Account"."Balance at Date"
                  ELSE
                    TotalCreditBalanceAtDate := TotalCreditBalanceAtDate -"G/L Account"."Balance at Date";
                  IF "Net Change" > 0 THEN
                    TotalDebitNetChange := TotalDebitNetChange + "G/L Account"."Net Change"
                  ELSE
                    TotalCreditNetChange := TotalCreditNetChange -"G/L Account"."Net Change";
                END;
                IF NOT PrintToExcel THEN BEGIN
                  IF LastNewPage THEN
                    GroupNum := GroupNum + 1;
                  LastNewPage := "New Page";
                END;
                */

                //>>RB-N 13Sep2017 NetChange Amount
                CLEAR(NNNN);
                CLEAR(TempAmt);
                GLE1.RESET;
                GLE1.SETCURRENTKEY("G/L Account No.", "Posting Date");
                GLE1.SETRANGE("G/L Account No.", "No.");
                GLE1.SETRANGE("Posting Date", SDate, EDate);
                GLE1.SETFILTER("Global Dimension 2 Code", '<>%1', '');
                IF GLE1.FINDSET THEN
                    REPEAT

                        //MESSAGE(' Count %1 \\ Entry No. %2',GLE1.COUNT,GLE1."Entry No.");

                        RespCent.RESET;
                        RespCent.SETRANGE(Code, GLE1."Global Dimension 2 Code");
                        RespCent.SETFILTER("Location Code", '<>%1', '');
                        IF RespCent.FINDFIRST THEN BEGIN
                            //MESSAGE(' Responsibility Center %1 \\ Location %2',RespCent.Code,RespCent."Location Code");
                            Loc1.RESET;
                            Loc1.SETRANGE(Code, RespCent."Location Code");
                            Loc1.SETRANGE("GST Registration No.", TempGSTNo);
                            IF Loc1.FINDFIRST THEN BEGIN

                                TempAmt += GLE1.Amount;
                                NNNN += 1;
                            END;

                        END;

                    UNTIL GLE1.NEXT = 0;

                //MESSAGE('Temp Amount %1 \\  Count %2 ',TempAmt,NNNN);

                CLEAR(NCrAmt);
                CLEAR(NDrAmt);

                IF TempAmt > 0 THEN
                    NDrAmt := TempAmt
                ELSE
                    NCrAmt := -TempAmt;
                //>>RB-N 13Sep2017 NetChange Amount


                //>>RB-N 14Sep2017 Balance Amount
                CLEAR(NNNN);
                CLEAR(BalanceAmt);
                GLE1.RESET;
                GLE1.SETCURRENTKEY("G/L Account No.", "Posting Date");
                GLE1.SETRANGE("G/L Account No.", "No.");
                GLE1.SETRANGE("Posting Date", 0D, EDate);
                GLE1.SETFILTER("Global Dimension 2 Code", '<>%1', '');
                IF GLE1.FINDSET THEN
                    REPEAT

                        //MESSAGE(' Count %1 \\ Entry No. %2',GLE1.COUNT,GLE1."Entry No.");

                        RespCent.RESET;
                        RespCent.SETRANGE(Code, GLE1."Global Dimension 2 Code");
                        RespCent.SETFILTER("Location Code", '<>%1', '');
                        IF RespCent.FINDFIRST THEN BEGIN
                            //MESSAGE(' Responsibility Center %1 \\ Location %2',RespCent.Code,RespCent."Location Code");
                            Loc1.RESET;
                            Loc1.SETRANGE(Code, RespCent."Location Code");
                            Loc1.SETRANGE("GST Registration No.", TempGSTNo);
                            IF Loc1.FINDFIRST THEN BEGIN

                                BalanceAmt += GLE1.Amount;
                                NNNN += 1;
                            END;

                        END;

                    UNTIL GLE1.NEXT = 0;

                //MESSAGE('Balance Amount %1 \\  Count %2 ',BalanceAmt,NNNN);

                CLEAR(BCrAmt);
                CLEAR(BDrAmt);

                IF BalanceAmt > 0 THEN
                    BDrAmt := BalanceAmt
                ELSE
                    BCrAmt := -BalanceAmt;
                //>>RB-N 14Sep2017 Balance Amount

                IF PrintToExcel THEN
                    MakeExcelDataBody;

                IF ("G/L Account"."Account Type" = "G/L Account"."Account Type"::Posting) THEN BEGIN
                    IF BalanceAmt > 0 THEN
                        TotalDebitBalanceAtDate := TotalDebitBalanceAtDate + BalanceAmt
                    ELSE
                        TotalCreditBalanceAtDate := TotalCreditBalanceAtDate - BalanceAmt;

                    IF TempAmt > 0 THEN
                        TotalDebitNetChange := TotalDebitNetChange + TempAmt
                    ELSE
                        TotalCreditNetChange := TotalCreditNetChange - TempAmt;
                END;

                IF NOT PrintToExcel THEN BEGIN
                    IF LastNewPage THEN
                        GroupNum := GroupNum + 1;
                    LastNewPage := "New Page";
                END;

            end;

            trigger OnPostDataItem()
            begin
                /*
                IF PrintToExcel THEN BEGIN
                  ExcelBuf.NewRow;
                  ExcelBuf.AddColumn('TOTALS',FALSE,'',TRUE,
                    FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
                  ExcelBuf.AddColumn('',FALSE,'',TRUE,
                    FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
                  ExcelBuf.AddColumn(TotalDebitNetChange,FALSE,'',TRUE,
                    FALSE,FALSE,'',ExcelBuf."Cell Type"::Number);
                  ExcelBuf.AddColumn(TotalCreditNetChange,FALSE,'',TRUE,
                    FALSE,FALSE,'',ExcelBuf."Cell Type"::Number);
                  ExcelBuf.AddColumn(TotalDebitBalanceAtDate,FALSE,'',TRUE,
                    FALSE,FALSE,'',ExcelBuf."Cell Type"::Number);
                  ExcelBuf.AddColumn(TotalCreditBalanceAtDate,FALSE,'',TRUE,
                    FALSE,FALSE,'',ExcelBuf."Cell Type"::Number);
                END;
                */

                //>>RB-N 13Sep2017
                IF PrintToExcel THEN BEGIN
                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn('TOTALS', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);
                    ExcelBuf.AddColumn(TotalDebitNetChange, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);
                    ExcelBuf.AddColumn(TotalCreditNetChange, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);
                    ExcelBuf.AddColumn(TotalDebitBalanceAtDate, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);
                    ExcelBuf.AddColumn(TotalCreditBalanceAtDate, FALSE, '', TRUE, FALSE, FALSE, '#,#0.00', 0);
                END;
                //<<RB-N 13Sep2017

            end;

            trigger OnPreDataItem()
            begin

                SETRANGE("G/L Account"."Date Filter", SDate, EDate);//13Sep2017
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                group(Options)
                {
                    Caption = 'Options';
                    field("GST Register No."; TempGSTNo)
                    {
                        ApplicationArea = all;
                        TableRelation = "GST Registration Nos.";
                    }
                    field("Start Date"; SDate)
                    {
                        ApplicationArea = all;
                        trigger OnValidate()
                        begin

                            IF SDate < 20170701D/*010717D*/ THEN
                                ERROR('Trail Balance GST Applicable from July2017');

                            EDate := 0D;
                        end;
                    }
                    field("End Date"; EDate)
                    {
                        ApplicationArea = all;
                        trigger OnValidate()
                        begin

                            IF EDate < SDate THEN
                                ERROR('End Date Must be Greather than Start Date');
                        end;
                    }
                    field(PrintToExcel; PrintToExcel)
                    {
                        ApplicationArea = all;
                        Caption = 'Print to Excel';
                    }
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
        GLFilter := "G/L Account".GETFILTERS;
        GLFilter := 'No.: ' + "G/L Account".GETFILTER("No.") + ', Date Filter: ' + FORMAT(SDate) + '..' + FORMAT(EDate) + ', GST Registration No.: ' + TempGSTNo;
        //PeriodText := "G/L Account".GETFILTER("Date Filter");
        PeriodText := FORMAT(SDate) + '..' + FORMAT(EDate);//14Sep2017
        IF PrintToExcel THEN
            MakeExcelInfo;


        //>>RB-N 13Sep2017
        IF TempGSTNo = '' THEN
            ERROR('Please Enter GST Registration No. to be search');

        IF SDate = 0D THEN
            ERROR('Please Specify Start Date');

        IF EDate = 0D THEN
            ERROR('Please Specify End Date');
        //>>RB-N 13Sep2017
    end;

    var
        Text000: Label 'Period: %1';
        ExcelBuf: Record 370 temporary;
        GLFilter: Text;
        PeriodText: Text[30];
        PrintToExcel: Boolean;
        Text001: Label 'Trial Balance GST';
        Text003: Label 'Debit';
        Text004: Label 'Credit';
        Text005: Label 'Company Name';
        Text006: Label 'Report No.';
        Text007: Label 'Report Name';
        Text008: Label 'User ID';
        Text009: Label 'Date';
        Text010: Label 'G/L Filter';
        Text011: Label 'Period Filter';
        TotalDebitNetChange: Decimal;
        TotalCreditNetChange: Decimal;
        TotalDebitBalanceAtDate: Decimal;
        TotalCreditBalanceAtDate: Decimal;
        NewTotalDebitNetChange: Decimal;
        NewTotalCreditNetChange: Decimal;
        NewTotalDebitBalanceAtDate: Decimal;
        NewTotalCreditBalanceAtDate: Decimal;
        GroupNum: Integer;
        LastNewPage: Boolean;
        Trial_BalanceCaptionLbl: Label 'Trial Balance GST';
        CurrReport_PAGENOCaptionLbl: Label 'Page';
        Net_ChangeCaptionLbl: Label 'Net Change';
        BalanceCaptionLbl: Label 'Balance';
        PADSTR_____G_L_Account__Indentation___2___G_L_Account__NameCaptionLbl: Label 'Name';
        G_L_Account___Net_Change_CaptionLbl: Label 'Debit';
        G_L_Account___Net_Change__Control22CaptionLbl: Label 'Credit';
        G_L_Account___Balance_at_Date_CaptionLbl: Label 'Debit';
        G_L_Account___Balance_at_Date__Control24CaptionLbl: Label 'Credit';
        TOTALSCaptionLbl: Label 'TOTALS';
        BlankLineNo: Integer;
        "--------13Sep2017--GST": Integer;
        SDate: Date;
        EDate: Date;
        LocGSTNo: Code[15];
        TempGSTNo: Code[15];
        GLE1: Record 17;
        GLE2: Record 17;
        RespCent: Record 5714;
        Loc1: Record 14;
        TempAmt: Decimal;
        NNNN: Integer;
        NCrAmt: Decimal;
        NDrAmt: Decimal;
        BalanceAmt: Decimal;
        BCrAmt: Decimal;
        BDrAmt: Decimal;

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text007), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(FORMAT(Text001), FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(50183, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text008), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text009), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text010), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn("G/L Account".GETFILTER("No."), FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text011), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddInfoColumn("G/L Account".GETFILTER("Date Filter"),FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(FORMAT(SDate) + '..' + FORMAT(EDate), FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    local procedure MakeExcelDataHeader()
    begin
        ExcelBuf.AddColumn("G/L Account".FIELDCAPTION("No."), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Account".FIELDCAPTION(Name), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(
          FORMAT("G/L Account".FIELDCAPTION("Net Change") + ' - ' + Text003), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(
          FORMAT("G/L Account".FIELDCAPTION("Net Change") + ' - ' + Text004), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(
          FORMAT("G/L Account".FIELDCAPTION("Balance at Date") + ' - ' + Text003), FALSE, '', TRUE, FALSE, TRUE, '',
          ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(
          FORMAT("G/L Account".FIELDCAPTION("Balance at Date") + ' - ' + Text004), FALSE, '', TRUE, FALSE, TRUE, '',
          ExcelBuf."Cell Type"::Text);
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelDataBody()
    var
        BlankFiller: Text[250];
    begin
        BlankFiller := PADSTR(' ', MAXSTRLEN(BlankFiller), ' ');
        /*
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(
          "G/L Account"."No.",FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
          ExcelBuf."Cell Type"::Text);
        IF "G/L Account".Indentation = 0 THEN
          ExcelBuf.AddColumn(
            "G/L Account".Name,FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
            ExcelBuf."Cell Type"::Text)
        ELSE
          ExcelBuf.AddColumn(
            COPYSTR(BlankFiller,1,2 * "G/L Account".Indentation) + "G/L Account".Name,
            FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        
        CASE TRUE OF
          "G/L Account"."Net Change" = 0:
            BEGIN
              ExcelBuf.AddColumn(
                '',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                ExcelBuf."Cell Type"::Text);
              ExcelBuf.AddColumn(
                '',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                ExcelBuf."Cell Type"::Text);
            END;
          "G/L Account"."Net Change" > 0:
            BEGIN
              ExcelBuf.AddColumn(
                "G/L Account"."Net Change",FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                FALSE,FALSE,'#,##0.00',ExcelBuf."Cell Type"::Number);
              ExcelBuf.AddColumn(
                '',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                ExcelBuf."Cell Type"::Text);
            END;
          "G/L Account"."Net Change" < 0:
            BEGIN
              ExcelBuf.AddColumn(
                '',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                ExcelBuf."Cell Type"::Text);
              ExcelBuf.AddColumn(
                -"G/L Account"."Net Change",FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                FALSE,FALSE,'#,##0.00',ExcelBuf."Cell Type"::Number);
            END;
        END;
        
        CASE TRUE OF
          "G/L Account"."Balance at Date" = 0:
            BEGIN
              ExcelBuf.AddColumn(
                '',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                ExcelBuf."Cell Type"::Text);
              ExcelBuf.AddColumn(
                '',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                ExcelBuf."Cell Type"::Text);
            END;
          "G/L Account"."Balance at Date" > 0:
            BEGIN
              ExcelBuf.AddColumn(
                "G/L Account"."Balance at Date",FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                FALSE,FALSE,'#,##0.00',ExcelBuf."Cell Type"::Number);
              ExcelBuf.AddColumn(
                '',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                ExcelBuf."Cell Type"::Text);
            END;
          "G/L Account"."Balance at Date" < 0:
            BEGIN
              ExcelBuf.AddColumn(
                '',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                ExcelBuf."Cell Type"::Text);
              ExcelBuf.AddColumn(
                -"G/L Account"."Balance at Date",FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                FALSE,FALSE,'#,##0.00',ExcelBuf."Cell Type"::Number);
            END;
        END;
        */

        //>>RB-N 14Sep2017 New DataBody
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("G/L Account"."No.", FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 0);//1

        IF "G/L Account".Indentation = 0 THEN
            ExcelBuf.AddColumn("G/L Account".Name, FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1)
        ELSE
            ExcelBuf.AddColumn(COPYSTR(BlankFiller, 1, 2 * "G/L Account".Indentation) + "G/L Account".Name,
              FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1);//2

        CASE TRUE OF
            TempAmt = 0:
                BEGIN
                    ExcelBuf.AddColumn('', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1);
                    ExcelBuf.AddColumn('', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1);
                END;
            TempAmt > 0:
                BEGIN
                    ExcelBuf.AddColumn(TempAmt, FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '#,#0.00', 0);
                    ExcelBuf.AddColumn('', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 0);
                END;
            TempAmt < 0:
                BEGIN
                    ExcelBuf.AddColumn('', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1);
                    ExcelBuf.AddColumn(-TempAmt, FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '#,#0.00', 0);
                END;
        END;

        CASE TRUE OF
            BalanceAmt = 0:
                BEGIN
                    ExcelBuf.AddColumn('', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1);
                    ExcelBuf.AddColumn('', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1);
                END;
            BalanceAmt > 0:
                BEGIN
                    ExcelBuf.AddColumn(BalanceAmt, FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '#,#0.00', 0);
                    ExcelBuf.AddColumn('', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1);
                END;
            BalanceAmt < 0:
                BEGIN
                    ExcelBuf.AddColumn('', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', 1);
                    ExcelBuf.AddColumn(-BalanceAmt, FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '#,#0.00', 0);
                END;
        END;

        //>>RB-N 14Sep2017 New DataBody

    end;

    ////[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        // ExcelBuf.CreateBookAndOpenExcel('', Text001, '', COMPANYNAME, USERID);
        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;
}


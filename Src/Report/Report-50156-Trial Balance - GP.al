report 50156 "Trial Balance - GP"
{
    //  Hotfix 38801
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TrialBalanceGP.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Trial Balance';

    dataset
    {
        dataitem("G/L Account"; "G/L Account")
        {
            CalcFields = "Net Change", "Balance at Date";
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.", "Account Type", "Date Filter", "Global Dimension 1 Filter", "Global Dimension 2 Filter";
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
            column(TotalDebitOpening; TotalDebitOpening)
            {
            }
            column(TotalCreditOpening; TotalCreditOpening)
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
                column(OpeningDebit; OpeningAmountDebit)
                {
                }
                column(OpeningCredit; -OpeningAmountCredit)
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
                CALCFIELDS("Net Change", "Balance at Date");

                CLEAR(OpeningAmountDebit);
                CLEAR(OpeningAmountCredit);
                RecGLEntry.RESET;
                RecGLEntry.SETCURRENTKEY("G/L Account No.", "Posting Date");
                RecGLEntry.SETRANGE(RecGLEntry."G/L Account No.", "G/L Account"."No.");
                RecGLEntry.SETRANGE(RecGLEntry."Posting Date", 0D, OpeningDate);
                IF UnitFilter <> '' THEN
                    //RecGLEntry.SETRANGE(RecGLEntry."Global Dimension 1 Code",UnitFilter); //RK-Comment
                    RecGLEntry.SETFILTER(RecGLEntry."Global Dimension 1 Code", UnitFilter); //RK-ADD
                IF EmployeeFilter <> '' THEN
                    RecGLEntry.SETRANGE(RecGLEntry."Global Dimension 2 Code", EmployeeFilter);
                IF RecGLEntry.FINDFIRST THEN BEGIN
                    RecGLEntry.CALCSUMS(RecGLEntry.Amount);
                    IF RecGLEntry.Amount > 0 THEN
                        OpeningAmountDebit := RecGLEntry.Amount;
                    IF RecGLEntry.Amount < 0 THEN
                        OpeningAmountCredit := RecGLEntry.Amount;
                END;

                IF PrintToExcel THEN
                    MakeExcelDataBody;

                IF ("G/L Account"."Account Type" = "G/L Account"."Account Type"::Posting) THEN BEGIN
                    IF "Balance at Date" > 0 THEN
                        TotalDebitBalanceAtDate := TotalDebitBalanceAtDate + "G/L Account"."Balance at Date"
                    ELSE
                        TotalCreditBalanceAtDate := TotalCreditBalanceAtDate - "G/L Account"."Balance at Date";
                    IF "Net Change" > 0 THEN BEGIN
                        TotalDebitNetChange := TotalDebitNetChange + "G/L Account"."Net Change";
                        //RSPLSUM 12Jun2020--TotalDebitOpening := TotalDebitOpening + OpeningAmountDebit;
                    END ELSE BEGIN
                        TotalCreditNetChange := TotalCreditNetChange - "G/L Account"."Net Change";
                        //RSPLSUM 12Jun2020--TotalCreditOpening := TotalCreditOpening - OpeningAmountCredit;
                    END;
                    //RSPLSUM 12Jun2020>>
                    IF (OpeningAmountDebit > 0) AND (OpeningAmountCredit = 0) THEN
                        TotalDebitOpening := TotalDebitOpening + OpeningAmountDebit;
                    IF (OpeningAmountDebit = 0) AND (OpeningAmountCredit < 0) THEN
                        TotalCreditOpening := TotalCreditOpening - OpeningAmountCredit;
                    //RSPLSUM 12Jun2020<<
                END;
                IF NOT PrintToExcel THEN BEGIN
                    IF LastNewPage THEN
                        GroupNum := GroupNum + 1;
                    LastNewPage := "New Page";
                END;
            end;

            trigger OnPostDataItem()
            begin
                IF PrintToExcel THEN BEGIN
                    ExcelBuf.NewRow;
                    ExcelBuf.AddColumn('TOTALS', FALSE, '', TRUE,
                      FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn('', FALSE, '', TRUE,
                      FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(TotalDebitOpening, FALSE, '', TRUE,
                      FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(TotalCreditOpening, FALSE, '', TRUE,
                      FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(TotalDebitNetChange, FALSE, '', TRUE,
                      FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(TotalCreditNetChange, FALSE, '', TRUE,
                      FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(TotalDebitBalanceAtDate, FALSE, '', TRUE,
                      FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(TotalCreditBalanceAtDate, FALSE, '', TRUE,
                      FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
                END;
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
        PeriodText := "G/L Account".GETFILTER("Date Filter");
        OpeningDate := "G/L Account".GETRANGEMIN("Date Filter") - 1;
        UnitFilter := "G/L Account".GETFILTER("G/L Account"."Global Dimension 1 Filter");
        EmployeeFilter := "G/L Account".GETFILTER("G/L Account"."Global Dimension 2 Filter");
        IF PrintToExcel THEN
            MakeExcelInfo;
    end;

    var
        Text000: Label 'Period: %1';
        ExcelBuf: Record 370 temporary;
        GLFilter: Text;
        PeriodText: Text[30];
        PrintToExcel: Boolean;
        Text001: Label 'Trial Balance GP';
        Text002: Label 'Data';
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
        Trial_BalanceCaptionLbl: Label 'Trial Balance';
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
        RecGLEntry: Record 17;
        OpeningDate: Date;
        OpeningAmountDebit: Decimal;
        OpeningAmountCredit: Decimal;
        TotalDebitOpening: Decimal;
        TotalCreditOpening: Decimal;
        UnitFilter: Text;
        EmployeeFilter: Text;

    // //[Scope('Internal')]
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
        ExcelBuf.AddInfoColumn(REPORT::"Trial Balance", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Number);
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
        ExcelBuf.AddInfoColumn("G/L Account".GETFILTER("Date Filter"), FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    local procedure MakeExcelDataHeader()
    begin
        ExcelBuf.AddColumn(FORMAT(Text007), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT(Text001), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(FORMAT(Text011), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Account".GETFILTER("Date Filter"), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(FORMAT(Text005), FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Unit Filter', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(UnitFilter, FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("G/L Account".FIELDCAPTION("No."), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("G/L Account".FIELDCAPTION(Name), FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Opening - Debit', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Opening - Credit', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
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

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    var
        BlankFiller: Text[250];
    begin
        BlankFiller := PADSTR(' ', MAXSTRLEN(BlankFiller), ' ');
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(
          "G/L Account"."No.", FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
          ExcelBuf."Cell Type"::Text);
        IF "G/L Account".Indentation = 0 THEN
            ExcelBuf.AddColumn(
              "G/L Account".Name, FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
              ExcelBuf."Cell Type"::Text)
        ELSE
            ExcelBuf.AddColumn(
              COPYSTR(BlankFiller, 1, 2 * "G/L Account".Indentation) + "G/L Account".Name,
              FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);

        //RSPLSUM 12Jun2020>>
        IF (OpeningAmountDebit = 0) AND (OpeningAmountCredit = 0) THEN BEGIN
            ExcelBuf.AddColumn(
              '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
              ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(
              '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
              ExcelBuf."Cell Type"::Text);
        END;

        IF (OpeningAmountDebit > 0) AND (OpeningAmountCredit = 0) THEN BEGIN
            ExcelBuf.AddColumn(
              OpeningAmountDebit, FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
              FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(
              '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
              ExcelBuf."Cell Type"::Text);
        END;

        IF (OpeningAmountDebit = 0) AND (OpeningAmountCredit < 0) THEN BEGIN
            ExcelBuf.AddColumn(
              '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
              ExcelBuf."Cell Type"::Text);
            ExcelBuf.AddColumn(
              -OpeningAmountCredit, FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
              FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        END;
        //RSPLSUM 12Jun2020<<

        CASE TRUE OF
            "G/L Account"."Net Change" = 0:
                BEGIN
                    //RSPLSUM 12Jun2020--ExcelBuf.AddColumn(
                    //RSPLSUM 12Jun2020--'',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                    //RSPLSUM 12Jun2020--ExcelBuf."Cell Type"::Text);
                    //RSPLSUM 12Jun2020--ExcelBuf.AddColumn(
                    //RSPLSUM 12Jun2020--'',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                    //RSPLSUM 12Jun2020--ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(
                      '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
                      ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(
                      '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
                      ExcelBuf."Cell Type"::Text);
                END;
            "G/L Account"."Net Change" > 0:
                BEGIN
                    //RSPLSUM 12Jun2020--ExcelBuf.AddColumn(
                    //RSPLSUM 12Jun2020--OpeningAmountDebit,FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                    //RSPLSUM 12Jun2020--FALSE,FALSE,'#,##0.00',ExcelBuf."Cell Type"::Number);
                    //RSPLSUM 12Jun2020--ExcelBuf.AddColumn(
                    //RSPLSUM 12Jun2020--'',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                    //RSPLSUM 12Jun2020--ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(
                      "G/L Account"."Net Change", FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                      FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(
                      '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
                      ExcelBuf."Cell Type"::Text);
                END;
            "G/L Account"."Net Change" < 0:
                BEGIN
                    //RSPLSUM 12Jun2020--ExcelBuf.AddColumn(
                    //RSPLSUM 12Jun2020--'',FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,FALSE,FALSE,'',
                    //RSPLSUM 12Jun2020--ExcelBuf."Cell Type"::Text);
                    //RSPLSUM 12Jun2020--ExcelBuf.AddColumn(
                    //RSPLSUM 12Jun2020---OpeningAmountCredit,FALSE,'',"G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                    //RSPLSUM 12Jun2020--FALSE,FALSE,'#,##0.00',ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(
                      '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
                      ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(
                      -"G/L Account"."Net Change", FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                      FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
                END;
        END;

        CASE TRUE OF
            "G/L Account"."Balance at Date" = 0:
                BEGIN
                    ExcelBuf.AddColumn(
                      '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
                      ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(
                      '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
                      ExcelBuf."Cell Type"::Text);
                END;
            "G/L Account"."Balance at Date" > 0:
                BEGIN
                    ExcelBuf.AddColumn(
                      "G/L Account"."Balance at Date", FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                      FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
                    ExcelBuf.AddColumn(
                      '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
                      ExcelBuf."Cell Type"::Text);
                END;
            "G/L Account"."Balance at Date" < 0:
                BEGIN
                    ExcelBuf.AddColumn(
                      '', FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting, FALSE, FALSE, '',
                      ExcelBuf."Cell Type"::Text);
                    ExcelBuf.AddColumn(
                      -"G/L Account"."Balance at Date", FALSE, '', "G/L Account"."Account Type" <> "G/L Account"."Account Type"::Posting,
                      FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
                END;
        END;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBookAndOpenExcel(Text002,Text001,COMPANYNAME,USERID);
        //ERROR('');

        /* ExcelBuf.CreateBook('', Text002);
        //ExcelBuf.CreateBookAndOpenExcel('',Text000,Text001,COMPANYNAME,USERID);
        ExcelBuf.CreateBookAndOpenExcel('', Text002, '', '', USERID);
        ExcelBuf.GiveUserControl; */
        ExcelBuf.CreateNewBook(Text002);
        ExcelBuf.WriteSheet(Text002, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text002, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();

    end;
}


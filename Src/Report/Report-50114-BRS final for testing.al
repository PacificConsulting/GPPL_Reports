report 50114 "BRS final for testing"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BRSfinalfortesting.rdl';
    Caption = 'Bank Reco Statement';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Bank Account Ledger Entry"; "Bank Account Ledger Entry")
        {
            DataItemTableView = SORTING("Bank Account No.", "Positive", "Posting Date", "Open", "Document No.")
                                WHERE("Stale Cheque" = CONST(false),
                                      Reversed = CONST(false),
                                      "Entry No." = FILTER(<> 373494));
            RequestFilterFields = "Bank Account No.";
            column(CompName; compinfo.Name)
            {
            }
            column(TODATE; TODATE)
            {
            }
            column(BankName; Bank.Name)
            {
            }
            column(BalanceAmt; ABS(BalanceAsPerBook))
            {
            }
            column(GroupType; GroupType)
            {
            }
            column(DocDate; "Document Date")
            {
            }
            column(ChequeNo; "Cheque No.")
            {
            }
            column(Desc; Description)
            {
            }
            column(cleareddate; cleareddate)
            {
            }
            column(Amt; ABS(Amount))
            {
            }
            column(BookCRDR; BookCRDR)
            {
            }
            column(GroupDRCR; GroupDRCR)
            {
            }
            column(TOTCRDR; TOTCRDR)
            {
            }
            column(FinalAmt; ABS(BalanceAsPerBook - decTotalAmount))
            {
            }
            column(Positive; Positive)
            {
            }
            column(EntryNo; "Entry No.")
            {
            }
            column(DocNo; "Bank Account Ledger Entry"."Document No.")
            {
            }
            column(DocVis; DocVis)
            {
            }
            column(decAmount; decAmount)
            {
            }
            column(decTotalAmount; decTotalAmount)
            {
            }
            column(LastCRDR; LastCRDR)
            {
            }

            trigger OnAfterGetRecord()
            begin



                NNNNN += 1;
                //>>2
                //Bank Account Ledger Entry, Header (1) - OnPreSection()

                IF NNNNN = 1 THEN BEGIN

                    BalanceAsPerBook := 0;
                    IF Bank.GET("Bank Account Ledger Entry"."Bank Account No.") THEN;//12May2017
                                                                                     //Bank.GET("Bank Account Ledger Entry"."Bank Account No.");


                    IF Printexcel THEN
                        MakeExcelInfo;

                    BankLedgerEntry.RESET;
                    BankLedgerEntry.SETCURRENTKEY("Bank Account No.", Positive, "Posting Date", Open, "Document No.");
                    BankLedgerEntry.SETRANGE("Bank Account No.", "Bank Account No.");
                    BankLedgerEntry.SETFILTER(BankLedgerEntry."Posting Date", '%1..%2', 0D, TODATE);
                    IF BankLedgerEntry.FIND('-') THEN
                        REPEAT
                            BalanceAsPerBook := BalanceAsPerBook + BankLedgerEntry.Amount
                        UNTIL BankLedgerEntry.NEXT <= 0;

                    IF BalanceAsPerBook < 0 THEN
                        BookCRDR := 'Cr.'
                    ELSE
                        BookCRDR := 'Dr.';

                    IF Printexcel THEN
                        MakeExcelheader;

                END;
                //<<2

                TTTT += 1;//12May2017

                //>>12May2017 NewGroupHeader

                BALE12.SETCURRENTKEY("Bank Account No.", Positive, "Posting Date", Open, "Document No.");
                //Bank Account No.,Positive,Posting Date,Open,Document No.
                BALE12.SETRANGE(Positive, "Bank Account Ledger Entry".Positive);
                IF BALE12.FINDLAST THEN BEGIN
                    IIII := BALE12.COUNT;
                END;

                IF TTTT = 1 THEN BEGIN
                    decAmount := 0;
                    //IF "Bank Account Ledger Entry".Positive = "Bank Account Ledger Entry".Positive::"1" THEN BEGIN
                    GroupType := Text02;

                    IF Printexcel THEN
                        MakeExcelGroupHeader;
                    //END ELSE BEGIN
                    GroupType := Text01;

                    IF Printexcel THEN
                        MakeExcelGroupHeader;
                    //END;

                END;
                //>>12May2017 NewGroupHeader



                //>>1
                /*
                Bank Account Ledger Entry, GroupHeader ( - OnPreSection()
                CurrReport.SHOWOUTPUT := CurrReport.TOTALSCAUSEDBY=FIELDNO(Positive);
                IF CurrReport.SHOWOUTPUT = TRUE THEN
                BEGIN
                   decAmount:=0;
                      IF "Bank Account Ledger Entry".Positive="Bank Account Ledger Entry".Positive::"1" THEN
                      BEGIN
                        GroupType:=Text02;
                         IF Printexcel THEN
                          MakeExcelGroupHeader;
                      END
                      ELSE
                      BEGIN
                        GroupType:=Text01;
                
                        IF Printexcel THEN
                           MakeExcelGroupHeader;
                      END;
                END;
                *///Commented 12May2017
                  //<<1

                //>>12May2017 NewDataBody
                CLEAR(DocVis);
                IF "Bank Account Ledger Entry".Open = FALSE THEN BEGIN
                    recBankAccStatementLine.RESET;
                    recBankAccStatementLine.SETCURRENTKEY("Bank Account No.", "Document No.", "Value Date");
                    recBankAccStatementLine.SETRANGE("Document No.", "Document No.");
                    recBankAccStatementLine.SETRANGE("Bank Account No.", "Bank Account No.");
                    recBankAccStatementLine.SETRANGE(recBankAccStatementLine."Statement No.", "Statement No.");
                    recBankAccStatementLine.SETRANGE(recBankAccStatementLine."Statement Line No.", "Statement Line No.");
                    recBankAccStatementLine.SETFILTER("Value Date", '>%1', TODATE);
                    IF recBankAccStatementLine.FIND('-') THEN BEGIN
                        cleareddate := recBankAccStatementLine."Value Date";
                        IF Printexcel THEN
                            MakeExcelDataBody;

                        DocVis := TRUE;
                        //CurrReport.SHOWOUTPUT := TRUE;
                    END ELSE BEGIN
                        DocVis := FALSE;
                        //CurrReport.SHOWOUTPUT := FALSE;
                        cleareddate := 0D;
                    END;
                END ELSE BEGIN
                    recBankAccStatementLine.RESET;
                    recBankAccStatementLine.SETCURRENTKEY("Bank Account No.", "Document No.", "Value Date");
                    recBankAccStatementLine.SETRANGE("Document No.", "Document No.");
                    recBankAccStatementLine.SETRANGE("Bank Account No.", "Bank Account No.");
                    recBankAccStatementLine.SETRANGE(recBankAccStatementLine."Statement No.", "Statement No.");
                    IF recBankAccStatementLine.FIND('-') THEN
                        cleareddate := recBankAccStatementLine."Value Date"
                    ELSE
                        cleareddate := 0D;

                    IF Printexcel THEN
                        MakeExcelDataBody;

                    DocVis := TRUE;
                END;

                IF DocVis THEN
                    decAmount += Amount;//12May2017

                //IF CurrReport.SHOWOUTPUT = TRUE THEN
                //decAmount+=Amount;

                //<<12May2017 NewDataBody

                //>>2
                /*
                Bank Account Ledger Entry, Body (3) - OnPreSection()
                IF "Bank Account Ledger Entry".Open =FALSE THEN
                BEGIN
                  recBankAccStatementLine.RESET;
                  recBankAccStatementLine.SETCURRENTKEY("Bank Account No.","Document No.","Value Date");
                  recBankAccStatementLine.SETRANGE("Document No.","Document No.");
                  recBankAccStatementLine.SETRANGE("Bank Account No.","Bank Account No.");
                  recBankAccStatementLine.SETRANGE(recBankAccStatementLine."Statement No.","Statement No.");
                  recBankAccStatementLine.SETRANGE(recBankAccStatementLine."Statement Line No.","Statement Line No.");
                  recBankAccStatementLine.SETFILTER("Value Date",'>%1',TODATE);
                     IF recBankAccStatementLine.FIND('-') THEN
                     BEGIN
                        cleareddate:=recBankAccStatementLine."Value Date";
                        IF Printexcel THEN
                        MakeExcelDataBody;
                        CurrReport.SHOWOUTPUT := TRUE;
                     END
                     ELSE
                     BEGIN
                       CurrReport.SHOWOUTPUT := FALSE;
                       cleareddate:=0D;
                     END
                END
                ELSE
                BEGIN
                  recBankAccStatementLine.RESET;
                  recBankAccStatementLine.SETCURRENTKEY("Bank Account No.","Document No.","Value Date");
                  recBankAccStatementLine.SETRANGE("Document No.","Document No.");
                  recBankAccStatementLine.SETRANGE("Bank Account No.","Bank Account No.");
                  recBankAccStatementLine.SETRANGE(recBankAccStatementLine."Statement No.","Statement No.");
                  IF recBankAccStatementLine.FIND('-') THEN
                      cleareddate:=recBankAccStatementLine."Value Date"
                  ELSE
                      cleareddate:=0D;
                
                  IF Printexcel THEN
                     MakeExcelDataBody;
                END;
                
                IF CurrReport.SHOWOUTPUT = TRUE THEN
                   decAmount+=Amount;
                   */
                //<<2


                //>>12May2017 NewGroupFooter
                IF TTTT = IIII THEN BEGIN
                    decTotalAmount += decAmount;
                    IF decAmount < 0 THEN
                        GroupDRCR := 'Cr.'
                    ELSE
                        GroupDRCR := 'Dr.';

                    IF Printexcel THEN
                        MakeExcelGroupFooter;

                    TTTT := 0;

                END;

                //<<12May2017 NewGroupFooter

                //>>3
                /*
                Bank Account Ledger Entry, GroupFooter ( - OnPreSection()
                CurrReport.SHOWOUTPUT := CurrReport.TOTALSCAUSEDBY=FIELDNO(Positive);
                IF CurrReport.SHOWOUTPUT = TRUE THEN
                BEGIN
                   decTotalAmount+=decAmount;
                  IF decAmount <0 THEN
                    GroupDRCR:='Cr.'
                  ELSE
                    GroupDRCR:='Dr.';
                
                  IF Printexcel THEN
                     MakeExcelGroupFooter;
                END;
                
                */
                //<<3

                IF DocVis THEN BEGIN
                    IF BalanceAsPerBook - decTotalAmount < 0 THEN
                        LastCRDR := 'Cr.'
                    ELSE
                        LastCRDR := 'Dr.';

                END;

            end;

            trigger OnPostDataItem()
            begin
                //>>4
                //Bank Account Ledger Entry, Footer (5) - OnPreSection()
                IF BalanceAsPerBook - decTotalAmount < 0 THEN
                    TOTCRDR := 'Cr.'
                ELSE
                    TOTCRDR := 'Dr.';

                IF Printexcel THEN
                    MakeExcelfooter;

                //<<4
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                SETFILTER("Posting Date", '<=%1', TODATE);

                decTotalAmount := 0;

                //CurrReport.CREATETOTALS(decAmount);
                //<<1



                BALE12.COPYFILTERS("Bank Account Ledger Entry"); //12May2017

                TTTT := 0; //12May2017

                NNNNN := 0;//12May2017
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("TO DATE"; TODATE)
                {
                }
                field("Print to Excel"; Printexcel)
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
        IF Printexcel THEN
            CreateExcelbook;
        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1
        compinfo.GET;
        //<<1
    end;

    var
        GroupType: Text[1024];
        BankLedgerEntry: Record 271;
        BalanceAsPerBook: Decimal;
        BankLedgerEntry1: Record 271;
        compinfo: Record 79;
        TotalAmount: Decimal;
        ReceiptAmount: Decimal;
        PaymentAmount: Decimal;
        TODATE: Date;
        BookCRDR: Text[30];
        GroupDRCR: Text[30];
        TOTCRDR: Text[30];
        recBankAccStatementLine: Record 276;
        ExcelBuf: Record 370 temporary;
        decAmount: Decimal;
        decTotalAmount: Decimal;
        Printexcel: Boolean;
        Bank: Record 270;
        cleareddate: Date;
        Text01: Label 'Add:       Cheques Issued But not cleared by Bank';
        Text02: Label 'Less:      Cheques Received / Deposited but not yet cleared in Bank';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        Text000: Label 'Data';
        Text001: Label 'Bank Reconciliation Statement';
        "----12May2016": Integer;
        BALE12: Record 271;
        TTTT: Integer;
        IIII: Integer;
        NNNNN: Integer;
        DocVis: Boolean;
        LastCRDR: Text[30];

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;//12May2017
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn('Bank Reconciliation Statement', FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(REPORT::"BRS final for testing", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 1);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 1);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 2);
        ExcelBuf.ClearNewRow;
        MakeExcelDataHeader;
    end;

    // //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>12May2017 RB-N

        /*  ExcelBuf.CreateBook('', Text001);
         ExcelBuf.CreateBookAndOpenExcel('', Text001, '', '', USERID);
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataHeader()
    begin
        //>>Report Header 12May2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(COMPANYNAME, FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '', 1);//5

        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Bank Reconciliation Statement', FALSE, '', TRUE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '', 1);//5

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('As on  ' + FORMAT(TODATE), FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        //<<Report Header 12May2017

        //ExcelBuf.AddColumn(compinfo.Name,FALSE,'',TRUE,FALSE,TRUE,'');
        //ExcelBuf.NewRow;
        //ExcelBuf.AddColumn('Bank Reconciliation Statement as on  '+FORMAT(TODATE) ,FALSE,'',TRUE,FALSE,TRUE,'@');


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(Bank.Name, FALSE, '', TRUE, FALSE, TRUE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 1);//5
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Document date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Cheque No', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Description', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Subs. Clearing Date', FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '', 1);
        ExcelBuf.NewRow;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Bank Account Ledger Entry"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', 2);//1
        ExcelBuf.AddColumn("Bank Account Ledger Entry"."Cheque No.", FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn("Bank Account Ledger Entry".Description, FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn(cleareddate, FALSE, '', FALSE, FALSE, FALSE, '', 2);//4
        ExcelBuf.AddColumn(ABS("Bank Account Ledger Entry".Amount), FALSE, '', FALSE, FALSE, FALSE, '#,###0.00', 0);//5
        //decAmount+="Bank Account Ledger Entry".Amount; //12May2017
    end;

    //  //[Scope('Internal')]
    procedure MakeExcelGroupHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn(GroupType, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//1
        //ExcelBuf.AddColumn(GroupType,FALSE,GroupType,TRUE,FALSE,FALSE,'@');
    end;

    // //[Scope('Internal')]
    procedure MakeExcelGroupFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('Subtotal', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ABS(decAmount), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//5
        ExcelBuf.AddColumn(GroupDRCR, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
    end;

    // //[Scope('Internal')]
    procedure MakeExcelheader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('Balance as per Book', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ABS(BalanceAsPerBook), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//5
        ExcelBuf.AddColumn(BookCRDR, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
    end;

    // //[Scope('Internal')]
    procedure MakeExcelfooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 1);//3
        ExcelBuf.AddColumn('Balance as per Bank Pass Book', FALSE, '', TRUE, FALSE, FALSE, '', 1);//4
        ExcelBuf.AddColumn(ABS(BalanceAsPerBook - decTotalAmount), FALSE, '', TRUE, FALSE, FALSE, '#,###0.00', 0);//5
        ExcelBuf.AddColumn(TOTCRDR, FALSE, '', TRUE, FALSE, FALSE, '@', 1);//6
    end;
}


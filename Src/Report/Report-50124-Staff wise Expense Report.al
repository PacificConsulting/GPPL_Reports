report 50124 "Staff wise Expense Report"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/StaffwiseExpenseReport.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Dimension Value"; "Dimension Value")
        {
            DataItemTableView = SORTING("Dimension Code", Code)
                                ORDER(Ascending)
                                WHERE("Dimension Code" = FILTER('STAFF'));
            dataitem("G/L Account"; "G/L Account")
            {
                DataItemTableView = SORTING("No.")
                                    WHERE("Account Type" = FILTER(Posting));
                RequestFilterFields = "No.";

                trigger OnAfterGetRecord()
                begin


                    //>>1

                    NetChangeforTYear := 0;
                    NetChangeforPYear := 0;

                    GLEntry.RESET;
                    GLEntry.SETCURRENTKEY(GLEntry."G/L Account No.", GLEntry."Posting Date");
                    GLEntry.SETRANGE(GLEntry."G/L Account No.", "G/L Account"."No.");
                    GLEntry.SETRANGE(GLEntry."Posting Date", datStartDate, datEndDate);
                    IF GLEntry.FINDSET THEN
                        REPEAT

                            //>>16May2017
                            recDimSet.RESET;
                            recDimSet.SETRANGE("Dimension Set ID", GLEntry."Dimension Set ID");
                            recDimSet.SETRANGE("Dimension Code", 'STAFF');
                            recDimSet.SETRANGE("Dimension Value Code", "Dimension Value".Code);
                            IF recDimSet.FINDFIRST THEN BEGIN
                                NetChangeforTYear += GLEntry.Amount;
                                totNetChangeforTYear += GLEntry.Amount;
                            END;
                        //<<16May2017
                        /*
                          LedgerEntryDim.RESET;
                          LedgerEntryDim.SETCURRENTKEY(LedgerEntryDim."Table ID",LedgerEntryDim."Entry No.",LedgerEntryDim."Dimension Code");
                          LedgerEntryDim.SETRANGE(LedgerEntryDim."Table ID",17);
                          LedgerEntryDim.SETRANGE(LedgerEntryDim."Entry No.",GLEntry."Entry No.");
                          LedgerEntryDim.SETRANGE(LedgerEntryDim."Dimension Code",'STAFF');
                          LedgerEntryDim.SETRANGE(LedgerEntryDim."Dimension Value Code","Dimension Value".Code);
                          IF LedgerEntryDim.FINDFIRST THEN
                          BEGIN
                            NetChangeforTYear+=GLEntry.Amount;
                            totNetChangeforTYear+=GLEntry.Amount;
                          END;
                        *///Commented 16May2017
                        UNTIL GLEntry.NEXT = 0;

                    IF NetChangeforTYear <> 0 THEN //17May2017
                        IIII += 1; //17May2017
                                   //MESSAGE('I Value %1 \\ Dim code %2',IIII,"Dimension Value".Code);

                    /*
                    GLEntry.RESET;
                    GLEntry.SETCURRENTKEY(GLEntry."G/L Account No.",GLEntry."Posting Date");
                    GLEntry.SETRANGE(GLEntry."G/L Account No.","G/L Account"."No.");
                    GLEntry.SETRANGE(GLEntry."Posting Date",datPStartDate,datPEndDate);
                    IF GLEntry.FINDSET THEN
                    REPEAT
                    
                      //>>16May2017
                      recDimSet.RESET;
                      recDimSet.SETRANGE("Dimension Set ID",GLEntry."Dimension Set ID");
                      recDimSet.SETRANGE("Dimension Code",'STAFF');
                      recDimSet.SETRANGE("Dimension Value Code","Dimension Value".Code);
                      IF recDimSet.FINDFIRST THEN
                      BEGIN
                        NetChangeforPYear+=GLEntry.Amount;
                        totNetChangeforPYear+=GLEntry.Amount;
                      END;
                      //<<16May2017
                    
                    {
                      LedgerEntryDim.RESET;
                      LedgerEntryDim.SETCURRENTKEY(LedgerEntryDim."Table ID",LedgerEntryDim."Entry No.",LedgerEntryDim."Dimension Code");
                      LedgerEntryDim.SETRANGE(LedgerEntryDim."Table ID",17);
                      LedgerEntryDim.SETRANGE(LedgerEntryDim."Entry No.",GLEntry."Entry No.");
                      LedgerEntryDim.SETRANGE(LedgerEntryDim."Dimension Code",'STAFF');
                      LedgerEntryDim.SETRANGE(LedgerEntryDim."Dimension Value Code","Dimension Value".Code);
                      IF LedgerEntryDim.FINDFIRST THEN BEGIN
                        NetChangeforPYear+=GLEntry.Amount;
                        totNetChangeforPYear+=GLEntry.Amount;
                      END;
                    
                    }
                    UNTIL GLEntry.NEXT=0;
                    *///Commented Asper Instruction 17May2017
                      //<<1



                    //>>2
                    /*
                    G/L Account, Body (2) - OnPreSection()
                    IF NetChangeforTYear+NetChangeforPYear=0 THEN
                      CurrReport.SHOWOUTPUT(FALSE);
                    IF CurrReport.SHOWOUTPUT THEN BEGIN
                      Row+=1;
                      XLSHEET.Range('A'+FORMAT(Row)).Value :=FORMAT("No.");
                      XLSHEET.Range('B'+FORMAT(Row)).Value :=FORMAT(Name);
                      XLSHEET.Range('C'+FORMAT(Row)).Value :=FORMAT(NetChangeforTYear);
                      XLSHEET.Range('D'+FORMAT(Row)).Value :=FORMAT(NetChangeforPYear);
                    END;
                    *///Automation Commented
                      //<<2

                    //>>16May2017 GLAccountBody
                    IF NetChangeforTYear <> 0 THEN //17May2017
                    //IF NetChangeforTYear+NetChangeforPYear <> 0 THEN ////Commented Asper Instruction 17May2017
                    BEGIN

                        //>>16May2016 DimensionName
                        IF IIII = 1 THEN BEGIN

                            EnterCell(NNN, 1, 'Code', TRUE, TRUE, '', 1);
                            EnterCell(NNN, 2, 'Name', TRUE, TRUE, '', 1);
                            NNN += 1;

                            EnterCell(NNN, 1, "Dimension Value".Code, FALSE, FALSE, '', 1);
                            EnterCell(NNN, 2, "Dimension Value".Name, FALSE, FALSE, '', 1);
                            NNN += 1;
                        END;
                        //<<16May2016

                        //>>16May2017 GLAccountHeader
                        IF IIII = 1 THEN BEGIN

                            EnterCell(NNN, 1, 'No.', TRUE, TRUE, '', 1);
                            EnterCell(NNN, 2, 'Name', TRUE, TRUE, '', 1);
                            EnterCell(NNN, 3, 'Net Change for Current Year ' + FORMAT(datStartDate) + '-' + FORMAT(datEndDate), TRUE, TRUE, '', 1);
                            //EnterCell(NNN,4,'Net Change for Previous Year',TRUE,FALSE,'',1); //Commented Asper Instruction 17May2017
                            NNN += 1;

                            //EnterCell(NNN,3,FORMAT(datStartDate)+'-'+FORMAT(datEndDate),TRUE,FALSE,'',1);
                            //EnterCell(NNN,4,FORMAT(datPStartDate)+'-'+FORMAT(datPEndDate),TRUE,FALSE,'',1); //Commented Asper Instruction 17May2017
                            //NNN += 1;

                        END;
                        //>>16May2017 GLAccountHeader

                        EnterCell(NNN, 1, "No.", FALSE, FALSE, '', 1);
                        EnterCell(NNN, 2, Name, FALSE, FALSE, '', 1);
                        EnterCell(NNN, 3, FORMAT(NetChangeforTYear), FALSE, FALSE, '#,###0.00', 0);
                        //EnterCell(NNN,4,FORMAT(NetChangeforPYear),FALSE,FALSE,'#,###0.00',0); //Commented Asper Instruction 17May2017
                        NNN += 1;



                    END;

                    //>>16May2017 GLAccountBody

                end;

                trigger OnPostDataItem()
                begin

                    //>>1
                    /*
                    G/L Account, Footer (3) - OnPreSection()
                      Row+=1;
                      XLSHEET.Range('A'+FORMAT(Row)).Value :='Total';
                      XLSHEET.Range('A'+FORMAT(Row)).Font.Bold :=TRUE;
                      XLSHEET.Range('C'+FORMAT(Row)).Value :=FORMAT(totNetChangeforTYear);
                      XLSHEET.Range('D'+FORMAT(Row)).Value :=FORMAT(totNetChangeforPYear);
                      Row+=1;
                    *///Automation Commented
                      //<<1

                    //>>16May2017 GLAccountFooter
                    IF totNetChangeforTYear <> 0 THEN BEGIN
                        EnterCell(NNN, 1, 'Total', TRUE, TRUE, '', 1);
                        EnterCell(NNN, 2, '', TRUE, TRUE, '', 1);
                        EnterCell(NNN, 3, FORMAT(totNetChangeforTYear), TRUE, TRUE, '#,###0.00', 0);
                        //EnterCell(NNN,4,FORMAT(totNetChangeforPYear),TRUE,FALSE,'#,###0.00',0); //Commented Asper Instruction 17May2017

                        NNN += 2;
                    END;
                    //>>16May2017 GLAccountFooter

                end;

                trigger OnPreDataItem()
                begin

                    //>>1
                    /*
                    G/L Account, Header (1) - OnPreSection()
                    Row+=1;
                    XLSHEET.Range('A'+FORMAT(Row)).Value :='No.';
                    XLSHEET.Range('A'+FORMAT(Row)).Font.Bold:=TRUE;
                    
                    XLSHEET.Range('B'+FORMAT(Row)).Value :='Name';
                    XLSHEET.Range('B'+FORMAT(Row)).Font.Bold:=TRUE;
                    
                    XLSHEET.Range('C'+FORMAT(Row)).Value :='Net Change for Current Year';
                    XLSHEET.Range('C'+FORMAT(Row)).Font.Bold:=TRUE;
                    
                    XLSHEET.Range('D'+FORMAT(Row)).Value :='Net Change for Previous Year';
                    XLSHEET.Range('D'+FORMAT(Row)).Font.Bold:=TRUE;
                    
                    Row+=1;
                    
                    
                    XLSHEET.Range('C'+FORMAT(Row)).Value :='('+FORMAT(datStartDate)+'-'+FORMAT(datEndDate)+')';
                    XLSHEET.Range('C'+FORMAT(Row)).Font.Bold:=TRUE;
                    
                    XLSHEET.Range('D'+FORMAT(Row)).Value :='('+FORMAT(datPStartDate)+'-'+FORMAT(datPEndDate)+')';
                    XLSHEET.Range('D'+FORMAT(Row)).Font.Bold:=TRUE;
                    *///Automation Commented
                      //<<1

                    IIII := 0; //16May2017

                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                totNetChangeforTYear := 0;
                totNetChangeforPYear := 0;
                //<<1

                //>>2
                /*
                Dimension Value, Body (2) - OnPreSection()
                Row+=1;
                XLSHEET.Range('A'+FORMAT(Row)).Value :=FORMAT("Dimension Value".Code);
                XLSHEET.Range('B'+FORMAT(Row)).Value :=FORMAT("Dimension Value".Name);
                *///Automation Commented
                  //<<2

            end;

            trigger OnPreDataItem()
            begin

                //>>1
                CreateXLSHEET;
                //<<1

                //>>1
                /*
                Dimension Value, Header (1) - OnPreSection()
                IF CurrReport.PAGENO=1 THEN BEGIN
                  recCompInfo.GET;
                  XLSHEET.Activate;
                  XLSHEET.Range('A1').Value := FORMAT(recCompInfo.Name);
                  XLSHEET.Range('A1').Font.Bold:=TRUE;
                  XLSHEET.Range('D1').Value := FORMAT(USERID);
                  XLSHEET.Range('D1').Font.Bold:=TRUE;
                  XLSHEET.Range('E1').Value := FORMAT(WORKDATE);
                  XLSHEET.Range('E1').Font.Bold:=TRUE;
                
                  XLSHEET.Range('A2').Value := 'Code';
                  XLSHEET.Range('A2').Font.Bold:=TRUE;
                  XLSHEET.Range('B2').Value := 'Name';
                  XLSHEET.Range('B2').Font.Bold:=TRUE;
                  XLAPP.Visible(TRUE);
                  Row:=3;
                END;
                *///Automation Commented
                  //<<1
                  //SETRANGE("Dimension Value".Code,'GT01','GT02');//Test

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Current Start Date"; datStartDate)
                {
                }
                field("Current End Date"; datEndDate)
                {
                }
                field("Previous Start Date"; datPStartDate)
                {
                    Visible = false;
                }
                field("Previous End Date"; datPEndDate)
                {
                    Visible = false;
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
        /* ExcBuffer.CreateBook('', 'Staff wise Expense Report');
        ExcBuffer.CreateBookAndOpenExcel('', 'Staff wise Expense Report', '', '', USERID);
        ExcBuffer.GiveUserControl; */

        ExcBuffer.CreateNewBook('Staff wise Expense Report');
        ExcBuffer.WriteSheet('Staff wise Expense Report', CompanyName, UserId);
        ExcBuffer.CloseBook();
        ExcBuffer.SetFriendlyFilename(StrSubstNo('Staff wise Expense Report', CurrentDateTime, UserId));
        ExcBuffer.OpenExcel();
    end;

    trigger OnPreReport()
    begin

        //>>16May2016 ReportHeader

        EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '', 1);
        EnterCell(1, 3, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

        EnterCell(2, 1, 'Staff wise Expense Report', TRUE, FALSE, '', 1);
        EnterCell(2, 3, USERID, TRUE, FALSE, '', 1);

        NNN := 4;
        //<<16May2016
    end;

    var
        NetChangeforTYear: Decimal;
        NetChangeforPYear: Decimal;
        totNetChangeforTYear: Decimal;
        totNetChangeforPYear: Decimal;
        GLEntry: Record 17;
        datStartDate: Date;
        datEndDate: Date;
        datPStartDate: Date;
        datPEndDate: Date;
        Row: Integer;
        recCompInfo: Record 79;
        "---16May2017": Integer;
        recDimSet: Record 480;
        ExcBuffer: Record 370 temporary;
        NNN: Integer;
        IIII: Integer;

    // //[Scope('Internal')]
    procedure CreateXLSHEET()
    begin
        /*
        CLEAR(XLAPP);
        CLEAR(XLWBOOK);
        CLEAR(XLSHEET);
        CREATE(XLAPP);
        XLWBOOK := XLAPP.Workbooks.Add;
        XLSHEET := XLWBOOK.Worksheets.Add;
        XLAPP.Visible(FALSE);
        */

    end;

    // //[Scope('Internal')]
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


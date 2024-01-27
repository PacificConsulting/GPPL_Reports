report 50203 "JV Register"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 12Mar2018   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/JVRegister.rdl';
    Caption = 'JV Register';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("Document No.", "Posting Date", Amount)
                                ORDER(Descending)
                                WHERE("Source Code" = CONST('JOURNALV'));
            column(DocumentNo_GLEntry; "Document No.")
            {
            }
            column(PostingDateFormatted; FORMAT("Posting Date"))
            {
            }
            column(CreditAmount_GLEntry; "Credit Amount")
            {
            }
            column(DebitAmount_GLEntry; "Debit Amount")
            {
            }
            column(AccountName; AccountName)
            {
            }
            column(EntryNo; "Entry No.")
            {
            }
            column(JVNarration; JVNarration)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>AccountName

                CLEAR(AccountName);
                IF "System-Created Entry" AND ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Vendor) THEN BEGIN
                    Vend.GET("G/L Entry"."Source No.");
                    AccountName := Vend.Name;
                END;

                IF "System-Created Entry" AND ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Customer) THEN BEGIN
                    Cust.GET("G/L Entry"."Source No.");
                    AccountName := Cust.Name;
                END;

                IF "System-Created Entry" AND ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::"Fixed Asset") THEN BEGIN
                    FA.GET("G/L Entry"."Source No.");
                    AccountName := FA.Description;
                END;

                IF "System-Created Entry" AND ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::"Bank Account") THEN BEGIN
                    BankAcc.GET("G/L Entry"."Source No.");
                    AccountName := BankAcc.Name;
                END;

                IF "G/L Entry"."Bal. Account Type" = "G/L Entry"."Bal. Account Type"::"G/L Account" THEN BEGIN
                    GL.GET("G/L Entry"."G/L Account No.");
                    AccountName := GL.Name;
                END;

                //VendorName
                IF ("G/L Account No." = '32120300') OR ("G/L Account No." = '32120400')
                OR ("G/L Account No." = '32120100') OR ("G/L Account No." = '32120200')
                OR ("G/L Account No." = '32120500') OR ("G/L Account No." = '32120600')
                OR ("G/L Account No." = '32120700') THEN BEGIN

                    IF ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Vendor) THEN BEGIN
                        Vend.GET("G/L Entry"."Source No.");
                        AccountName := Vend.Name;
                    END;
                END;


                //CustomerName
                IF ("G/L Account No." = '12710000') OR ("G/L Account No." = '12720000') OR ("G/L Account No." = '12730000')
                OR ("G/L Account No." = '12740000') OR ("G/L Account No." = '12750000') OR ("G/L Account No." = '12760000')
                OR ("G/L Account No." = '12770000') OR ("G/L Account No." = '12780000') OR ("G/L Account No." = '12790000')
                OR ("G/L Account No." = '12791000') OR ("G/L Account No." = '12792000') THEN BEGIN
                    IF ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Customer) THEN BEGIN
                        Cust.GET("G/L Entry"."Source No.");
                        AccountName := Cust.Name;
                    END;
                END;

                IF AccountName = '' THEN BEGIN
                    "G/L Entry".CALCFIELDS("G/L Account Name");//
                    AccountName := "G/L Entry"."G/L Account Name";//
                END;
                //<<AccountName

                //>>
                JVNarration := '';
                PNar.RESET;
                PNar.SETRANGE("Entry No.", 0);
                PNar.SETRANGE("Transaction No.", "Transaction No.");
                PNar.SETRANGE("Document No.", "Document No.");
                IF PNar.FINDFIRST THEN
                    REPEAT
                        JVNarration += PNar.Narration;
                    UNTIL PNar.NEXT = 0;
                //<<
            end;

            trigger OnPreDataItem()
            begin
                SETRANGE("Posting Date", SDate, EDate);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Start Date"; SDate)
                {
                    ApplicationArea = all;
                }
                field("End Date"; EDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin
                        IF EDate < SDate THEN
                            ERROR('End Date: %1 must greather than Start Date: %2', EDate, SDate);
                    end;
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

    trigger OnPreReport()
    begin
        IF SDate = 0D THEN
            ERROR('Please Enter the Start Date');

        IF EDate = 0D THEN
            ERROR('Please Enter the End Date');
    end;

    var
        AccountName: Text[80];
        Vend: Record 23;
        Cust: Record 18;
        FA: Record 5600;
        BankAcc: Record 270;
        GL: Record 15;
        PNar: Record "Posted Narration";
        JVNarration: Text;
        SDate: Date;
        EDate: Date;
}


report 50184 "Customer Payment Register"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 12Sep2017   RB-N         New Report Development for Customer Payment Register
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CustomerPaymentRegister.rdl';
    Caption = 'Customer Payment Register';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Cust. Ledger Entry"; "Cust. Ledger Entry")
        {
            CalcFields = "Amount (LCY)";
            DataItemTableView = SORTING("Entry No.")
                                ORDER(Ascending)
                                WHERE("Document Type" = FILTER(<> Invoice));
            RequestFilterFields = "Customer No.", "Posting Date", "Document No.";
            dataitem("Detailed Cust. Ledg. Entry"; "Detailed Cust. Ledg. Entry")
            {
                DataItemLink = "Document No." = FIELD("Document No.");
                DataItemTableView = SORTING("Entry No.")
                                    ORDER(Ascending)
                                    WHERE("Initial Document Type" = FILTER(Invoice));
                dataitem(CustomerLedger; "Cust. Ledger Entry")
                {
                    DataItemLink = "Entry No." = FIELD("Cust. Ledger Entry No.");
                    DataItemTableView = SORTING("Entry No.")
                                        ORDER(Ascending);
                    dataitem("Value Entry"; "Value Entry")
                    {
                        DataItemLink = "Document No." = FIELD("Document No.");
                        DataItemTableView = SORTING("Entry No.")
                                            ORDER(Ascending);
                        dataitem("Item Ledger Entry"; "Item Ledger Entry")
                        {
                            DataItemLink = "Entry No." = FIELD("Item Ledger Entry No.");
                            DataItemTableView = SORTING("Document No.", "Document Type", "Document Line No.")
                                                ORDER(Ascending);

                            trigger OnAfterGetRecord()
                            begin

                                //>>1
                                //Item Ledger Entry, GroupFooter (2) - OnPreSection()
                                IF recPurchRcptHdr.GET("Document No.") THEN;
                                //Row+=1;
                                /*
                                XLSHEET.Range('L'+FORMAT(Row)).Value :=FORMAT("Item Ledger Entry"."Document No.");
                                XLSHEET.Range('M'+FORMAT(Row)).Value :=FORMAT(recPurchRcptHdr."Order No.");
                                IF recPO.GET(recPO."Document Type"::Order,recPurchRcptHdr."Order No.") THEN ;
                                XLSHEET.Range('N'+FORMAT(Row)).Value :=FORMAT(recPO."Blanket Order No.");
                                */
                                //<<1

                                //>>GroupFooter New

                                //ILE11.RESET;
                                //ILE11.SETCURRENTKEY("Document No.","Document Type","Document Line No.");

                                //>>GroupFooter New
                                /*
                                //>>GroupFooter
                                EnterCell(CCC,1,'',FALSE,FALSE,'',2);
                                EnterCell(CCC,2,'',FALSE,FALSE,'',1);
                                EnterCell(CCC,3,'',FALSE,FALSE,'',1);
                                EnterCell(CCC,4,'',FALSE,FALSE,'',1);
                                EnterCell(CCC,5,'',FALSE,FALSE,'',1);
                                EnterCell(CCC,6,'',FALSE,FALSE,'',1);//F
                                EnterCell(CCC,7,'',FALSE,FALSE,'',1);
                                EnterCell(CCC,8,'',FALSE,FALSE,'',2);
                                EnterCell(CCC,9,'',FALSE,FALSE,'',1);
                                EnterCell(CCC,10,'',FALSE,FALSE,'',1);//Item Category Code
                                EnterCell(CCC,11,'',FALSE,FALSE,'',1);//Invoice No.
                                EnterCell(CCC,12,"Item Ledger Entry"."Document No.",FALSE,FALSE,'',1);//GRN No.
                                EnterCell(CCC,13,recPurchRcptHdr."Order No.",FALSE,FALSE,'',1);//PO No.
                                IF recPO.GET(recPO."Document Type"::Order,recPurchRcptHdr."Order No.") THEN;
                                EnterCell(CCC,14,recPO."Blanket Order No.",FALSE,FALSE,'',1);//Blanket PO No.
                                EnterCell(CCC,15,'',FALSE,FALSE,'',1);
                                EnterCell(CCC,16,'',FALSE,FALSE,'',2);
                                EnterCell(CCC,17,'',FALSE,FALSE,'#,#0.00',0);
                                EnterCell(CCC,18,'',FALSE,FALSE,'',1);
                                
                                CCC += 1;
                                //<<GroupFooter
                                */

                            end;

                            trigger OnPreDataItem()
                            begin

                                TCOUNT := 0;
                            end;
                        }

                        trigger OnAfterGetRecord()
                        begin

                            //>>1
                            /*
                            IF (vDoc="Document No.") AND (vPDoc="Vendor Ledger Entry"."Document No.") THEN
                              CurrReport.SKIP;
                            
                            vPDoc:="Detailed Vendor Ledg. Entry"."Document No.";
                            vDoc:="Document No.";
                            */

                            IF (vDoc = "Document No.") AND (vPDoc = "Cust. Ledger Entry"."Document No.") THEN
                                CurrReport.SKIP;

                            vPDoc := "Detailed Cust. Ledg. Entry"."Document No.";
                            vDoc := "Document No.";
                            //<<1

                        end;
                    }
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1
                    /*
                    recVendorLedger.GET("Detailed Vendor Ledg. Entry"."Vendor Ledger Entry No.");
                    
                    IF vInDoc =recVendorLedger."Document No." THEN
                      CurrReport.SKIP;
                    
                    vInDoc:=recVendorLedger."Document No.";
                    */
                    recCusLedger12.GET("Detailed Cust. Ledg. Entry"."Cust. Ledger Entry No.");

                    IF vInDoc = recCusLedger12."Document No." THEN
                        CurrReport.SKIP;

                    vInDoc := recCusLedger12."Document No.";
                    //<<1


                    //>>2
                    //Detailed Vendor Ledg. Entry, Body (1) - OnPreSection()
                    Row += 1;
                    /*
                    XLSHEET.Range('F'+FORMAT(Row)).Value :=FORMAT(recVendorLedger."External Document No.");
                    
                    XLSHEET.Range('K'+FORMAT(Row)).Value :=FORMAT(recVendorLedger."Document No.");
                    XLSHEET.Range('L'+FORMAT(Row)).Value :=FORMAT('');
                    XLSHEET.Range('M'+FORMAT(Row)).Value :=FORMAT('');
                    */
                    //<<2

                    //>>DataBody2
                    EnterCell(CCC, 1, '', FALSE, FALSE, '', 2);
                    EnterCell(CCC, 2, '', FALSE, FALSE, '', 1);
                    EnterCell(CCC, 3, '', FALSE, FALSE, '', 1);
                    EnterCell(CCC, 4, '', FALSE, FALSE, '', 1);
                    EnterCell(CCC, 5, '', FALSE, FALSE, '', 1);
                    EnterCell(CCC, 6, recCusLedger12."External Document No.", FALSE, FALSE, '', 1);//F
                    EnterCell(CCC, 7, '', FALSE, FALSE, '', 1);
                    EnterCell(CCC, 8, '', FALSE, FALSE, '', 2);
                    EnterCell(CCC, 9, '', FALSE, FALSE, '', 1);
                    EnterCell(CCC, 10, '', FALSE, FALSE, '', 1);//Item Category Code
                    EnterCell(CCC, 11, recCusLedger12."Document No.", FALSE, FALSE, '', 1);//Invoice No.
                    EnterCell(CCC, 12, '', FALSE, FALSE, '', 1);//GRN No.
                    EnterCell(CCC, 13, '', FALSE, FALSE, '', 1);//PO No.
                    EnterCell(CCC, 14, '', FALSE, FALSE, '', 1);//Blanket PO No.
                    EnterCell(CCC, 15, '', FALSE, FALSE, '', 1);
                    EnterCell(CCC, 16, '', FALSE, FALSE, '', 2);
                    EnterCell(CCC, 17, '', FALSE, FALSE, '#,#0.00', 0);
                    EnterCell(CCC, 18, '', FALSE, FALSE, '', 1);

                    CCC += 1;
                    //<<DataBody2

                end;
            }

            trigger OnAfterGetRecord()
            begin

                NNN += 1;

                //>>DataHeader
                IF NNN = 1 THEN BEGIN

                    EnterCell(1, 1, COMPANYNAME, TRUE, FALSE, '', 1);
                    EnterCell(1, 18, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

                    EnterCell(2, 1, 'CUSTOMER PAYMENT REGISTER', TRUE, FALSE, '', 1);
                    EnterCell(2, 18, USERID, TRUE, FALSE, '', 1);


                    EnterCell(5, 1, 'Posting Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 2, 'Document No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 3, 'Customer Code', TRUE, TRUE, '', 1);
                    EnterCell(5, 4, 'Customer Name', TRUE, TRUE, '', 1);
                    EnterCell(5, 5, 'Doc Type', TRUE, TRUE, '', 1);
                    EnterCell(5, 6, 'External Doc No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 7, 'Payment Terms Code', TRUE, TRUE, '', 1);
                    EnterCell(5, 8, 'Due Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 9, 'Customer Posting Group', TRUE, TRUE, '', 1);
                    EnterCell(5, 10, 'Item Category', TRUE, TRUE, '', 1);
                    EnterCell(5, 11, 'Invoice No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 12, 'GRN No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 13, 'PO No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 14, 'Blanket PO No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 15, 'Cheque No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 16, 'Cheque Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 17, 'Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 18, 'Narration', TRUE, TRUE, '', 1);

                    CCC := 6;
                END;
                //>>DataHeader


                //>>1
                //Row+=2;
                vCheckNo := '';
                tNarr := '';
                dCheckDate := 0D;
                Cus12.GET("Customer No.");
                /*
                recCheckLedger.RESET;
                recCheckLedger.SETRANGE(recCheckLedger."Document No.","Document No.");
                IF recCheckLedger.FINDFIRST THEN
                BEGIN
                  vCheckNo:=recCheckLedger."Check No.";
                  dCheckDate:=recCheckLedger."Check Date";
                END;
                */

                recBankAccLedger.RESET;
                recBankAccLedger.SETRANGE("Document No.", "Document No.");
                IF recBankAccLedger.FINDFIRST THEN BEGIN
                    vCheckNo := recBankAccLedger."Cheque No.";
                    dCheckDate := recBankAccLedger."Cheque Date";
                END;

                recPostedNarration.RESET;
                recPostedNarration.SETRANGE(recPostedNarration."Transaction No.", "Transaction No.");
                recPostedNarration.SETRANGE(recPostedNarration."Entry No.", 0);
                IF recPostedNarration.FINDFIRST THEN
                    tNarr := recPostedNarration.Narration;


                //>>DataBody
                EnterCell(CCC, 1, FORMAT("Posting Date"), FALSE, FALSE, '', 2);
                EnterCell(CCC, 2, "Document No.", FALSE, FALSE, '', 1);
                EnterCell(CCC, 3, "Customer No.", FALSE, FALSE, '', 1);
                EnterCell(CCC, 4, Cus12.Name, FALSE, FALSE, '', 1);
                EnterCell(CCC, 5, FORMAT("Document Type"), FALSE, FALSE, '', 1);
                EnterCell(CCC, 6, "External Document No.", FALSE, FALSE, '', 1);
                EnterCell(CCC, 7, Cus12."Payment Terms Code", FALSE, FALSE, '', 1);
                EnterCell(CCC, 8, FORMAT("Due Date"), FALSE, FALSE, '', 2);
                EnterCell(CCC, 9, Cus12."Customer Posting Group", FALSE, FALSE, '', 1);
                EnterCell(CCC, 10, '', FALSE, FALSE, '', 1);//Item Category Code
                EnterCell(CCC, 11, '', FALSE, FALSE, '', 1);//Invoice No.
                EnterCell(CCC, 12, '', FALSE, FALSE, '', 1);//GRN No.
                EnterCell(CCC, 13, '', FALSE, FALSE, '', 1);//PO No.
                EnterCell(CCC, 14, '', FALSE, FALSE, '', 1);//Blanket PO No.
                EnterCell(CCC, 15, vCheckNo, FALSE, FALSE, '', 1);
                EnterCell(CCC, 16, FORMAT(dCheckDate), FALSE, FALSE, '', 2);
                EnterCell(CCC, 17, FORMAT("Amount (LCY)"), FALSE, FALSE, '#,#0.00', 0);
                EnterCell(CCC, 18, tNarr, FALSE, FALSE, '', 1);

                CCC += 1;
                //<<DataBody

                /*
                XLSHEET.Range('A'+FORMAT(Row)).Value :=FORMAT("Posting Date");
                XLSHEET.Range('B'+FORMAT(Row)).Value :=FORMAT("Document No.");
                XLSHEET.Range('C'+FORMAT(Row)).Value :=FORMAT("Vendor No.");
                XLSHEET.Range('D'+FORMAT(Row)).Value :=FORMAT(recVend.Name);
                XLSHEET.Range('E'+FORMAT(Row)).Value :=FORMAT("Document Type");
                XLSHEET.Range('F'+FORMAT(Row)).Value :=FORMAT("Vendor Ledger Entry"."External Document No.");
                XLSHEET.Range('G'+FORMAT(Row)).Value :=FORMAT(recVend."Payment Terms Code");
                XLSHEET.Range('H'+FORMAT(Row)).Value :=FORMAT("Due Date");
                XLSHEET.Range('I'+FORMAT(Row)).Value :=FORMAT(recVend."Vendor Posting Group");
                XLSHEET.Range('J'+FORMAT(Row)).Value :=FORMAT('');
                XLSHEET.Range('K'+FORMAT(Row)).Value :=FORMAT('');
                XLSHEET.Range('L'+FORMAT(Row)).Value :=FORMAT('');
                XLSHEET.Range('M'+FORMAT(Row)).Value :=FORMAT('');
                XLSHEET.Range('N'+FORMAT(Row)).Value :=FORMAT('');
                
                XLSHEET.Range('O'+FORMAT(Row)).Value :=FORMAT(vCheckNo);
                XLSHEET.Range('P'+FORMAT(Row)).Value :=FORMAT(dCheckDate);
                XLSHEET.Range('Q'+FORMAT(Row)).Value :=FORMAT("Vendor Ledger Entry"."Amount (LCY)");
                XLSHEET.Range('R'+FORMAT(Row)).Value :=FORMAT(tNarr);
                */
                //<<1

            end;

            trigger OnPreDataItem()
            begin

                //>>0
                //CreateXLSHEET;
                //CreateHeader;
                Row := 0;
                //<<0

                NNN := 0;
            end;
        }
    }

    requestpage
    {

        layout
        {
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

        //>>Excel Creation
        /*  ExcBuffer.CreateBook('', 'Customer Payment Register');
         ExcBuffer.CreateBookAndOpenExcel('', 'Customer Payment Register', '', '', USERID);
         ExcBuffer.GiveUserControl; */

        ExcBuffer.CreateNewBook('Customer Payment Register');
        ExcBuffer.WriteSheet('Customer Payment Register', CompanyName, UserId);
        ExcBuffer.CloseBook();
        ExcBuffer.SetFriendlyFilename(StrSubstNo('Customer Payment Register', CurrentDateTime, UserId));
        ExcBuffer.OpenExcel();
        //<<Excel Creation
    end;

    var
        Row: Integer;
        recVend: Record 23;
        recCheckLedger: Record 272;
        vCheckNo: Code[20];
        dCheckDate: Date;
        recVendorLedger: Record 25;
        vDoc: Code[20];
        recPurchRcptHdr: Record 120;
        vPDoc: Code[20];
        vInDoc: Code[20];
        recPO: Record 38;
        recPostedNarration: Record "Posted Narration";
        tNarr: Text[250];
        ExcBuffer: Record 370 temporary;
        "---11Sep2017": Integer;
        NNN: Integer;
        CCC: Integer;
        ILE11: Record 32;
        TCOUNT: Integer;
        ICOUNT: Integer;
        "---12Sep2017": Integer;
        Cus12: Record 18;
        recCusLedger12: Record 21;
        recBankAccLedger: Record 271;

    //  //[Scope('Internal')]
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


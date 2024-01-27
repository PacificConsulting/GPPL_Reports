report 50183 "Vendor Payment Register"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 11Sep2017   RB-N         New Report Development as per NAV2009
    // 18Sep2017   RB-N         New Fields(GST Vendor Type, GSTIN)
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/VendorPaymentRegister.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Vendor Payment Register';

    dataset
    {
        dataitem("Vendor Ledger Entry"; "Vendor Ledger Entry")
        {
            CalcFields = "Amount (LCY)";
            DataItemTableView = SORTING("Entry No.")
                                ORDER(Ascending)
                                WHERE("Document Type" = FILTER(<> Invoice));
            RequestFilterFields = "Vendor No.", "Posting Date";
            dataitem("Detailed Vendor Ledg. Entry"; "Detailed Vendor Ledg. Entry")
            {
                DataItemLink = "Document No." = FIELD("Document No.");
                DataItemTableView = SORTING("Entry No.")
                                    ORDER(Ascending)
                                    WHERE("Initial Document Type" = FILTER(Invoice));
                dataitem(VendLedger; "Vendor Ledger Entry")
                {
                    DataItemLink = "Entry No." = FIELD("Vendor Ledger Entry No.");
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

                                //>>GroupFooter
                                EnterCell(CCC, 1, '', FALSE, FALSE, '', 2);
                                EnterCell(CCC, 2, '', FALSE, FALSE, '', 1);
                                EnterCell(CCC, 3, '', FALSE, FALSE, '', 1);
                                EnterCell(CCC, 4, '', FALSE, FALSE, '', 1);
                                EnterCell(CCC, 5, '', FALSE, FALSE, '', 1);
                                EnterCell(CCC, 6, '', FALSE, FALSE, '', 1);//F
                                EnterCell(CCC, 7, '', FALSE, FALSE, '', 1);
                                EnterCell(CCC, 8, '', FALSE, FALSE, '', 2);
                                EnterCell(CCC, 9, '', FALSE, FALSE, '', 1);
                                EnterCell(CCC, 10, '', FALSE, FALSE, '', 1);//Item Category Code
                                EnterCell(CCC, 11, '', FALSE, FALSE, '', 1);//Invoice No.
                                EnterCell(CCC, 14, "Item Ledger Entry"."Document No.", FALSE, FALSE, '', 1);//GRN No.
                                EnterCell(CCC, 15, recPurchRcptHdr."Order No.", FALSE, FALSE, '', 1);//PO No.
                                IF recPO.GET(recPO."Document Type"::Order, recPurchRcptHdr."Order No.") THEN;
                                EnterCell(CCC, 16, recPO."Blanket Order No.", FALSE, FALSE, '', 1);//Blanket PO No.
                                EnterCell(CCC, 17, '', FALSE, FALSE, '', 1);
                                EnterCell(CCC, 18, '', FALSE, FALSE, '', 2);
                                EnterCell(CCC, 19, '', FALSE, FALSE, '#,#0.00', 0);
                                EnterCell(CCC, 20, '', FALSE, FALSE, '', 1);

                                CCC += 1;
                                //<<GroupFooter

                            end;

                            trigger OnPreDataItem()
                            begin

                                TCOUNT := 0;
                            end;
                        }

                        trigger OnAfterGetRecord()
                        begin

                            //>>1

                            IF (vDoc = "Document No.") AND (vPDoc = "Vendor Ledger Entry"."Document No.") THEN
                                CurrReport.SKIP;

                            vPDoc := "Detailed Vendor Ledg. Entry"."Document No.";
                            vDoc := "Document No.";
                            //<<1
                        end;
                    }
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1
                    recVendorLedger.GET("Detailed Vendor Ledg. Entry"."Vendor Ledger Entry No.");

                    IF vInDoc = recVendorLedger."Document No." THEN
                        CurrReport.SKIP;

                    vInDoc := recVendorLedger."Document No.";
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

                    /*
                    EnterCell(CCC,1,'',FALSE,FALSE,'',2);
                    EnterCell(CCC,2,'',FALSE,FALSE,'',1);
                    EnterCell(CCC,3,'',FALSE,FALSE,'',1);
                    EnterCell(CCC,4,'',FALSE,FALSE,'',1);
                    EnterCell(CCC,5,'',FALSE,FALSE,'',1);
                    */
                    EnterCell(CCC, 8, recVendorLedger."External Document No.", FALSE, FALSE, '', 1);//F
                    EnterCell(CCC, 9, FORMAT(recVendorLedger."Document Date"), FALSE, FALSE, '', 2);//RSPL--AR 151218
                    EnterCell(CCC, 14, recVendorLedger."Document No.", FALSE, FALSE, '', 1);//Invoice No.
                    EnterCell(CCC, 15, FORMAT(recVendorLedger."Posting Date"), FALSE, FALSE, '', 2);//RSPL--AR 151218


                    CLEAR(CheNo);
                    CLEAR(CheDate);
                    VLE2.RESET;
                    IF VLE2.GET(recVendorLedger."Closed by Entry No.") THEN BEGIN
                        BnkAccLed.RESET;
                        BnkAccLed.SETCURRENTKEY("Document No.");
                        BnkAccLed.SETRANGE("Document No.", VLE2."Document No.");
                        IF BnkAccLed.FINDFIRST THEN BEGIN
                            CheNo := BnkAccLed."Cheque No.";
                            CheDate := BnkAccLed."Cheque Date";
                        END;
                    END;

                    EnterCell(CCC, 19, CheNo, FALSE, FALSE, '', 1);//RSPL--AR 151218
                    EnterCell(CCC, 20, FORMAT(CheDate), FALSE, FALSE, '', 2);//RSPL--AR 151218


                    /*
                    EnterCell(CCC,7,'',FALSE,FALSE,'',1);
                    EnterCell(CCC,8,'',FALSE,FALSE,'',2);
                    EnterCell(CCC,9,'',FALSE,FALSE,'',1);
                    EnterCell(CCC,10,'',FALSE,FALSE,'',1);//Item Category Code
                    
                    EnterCell(CCC,12,'',FALSE,FALSE,'',1);//GRN No.
                    EnterCell(CCC,13,'',FALSE,FALSE,'',1);//PO No.
                    EnterCell(CCC,14,'',FALSE,FALSE,'',1);//Blanket PO No.
                    EnterCell(CCC,15,'',FALSE,FALSE,'',1);
                    EnterCell(CCC,16,'',FALSE,FALSE,'',2);
                    EnterCell(CCC,17,'',FALSE,FALSE,'#,#0.00',0);
                    EnterCell(CCC,18,'',FALSE,FALSE,'',1);
                    */
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
                    EnterCell(1, 20, FORMAT(TODAY, 0, 4), TRUE, FALSE, '', 1);

                    EnterCell(2, 1, 'VENDOR PAYMENT REGISTER', TRUE, FALSE, '', 1);
                    EnterCell(2, 20, USERID, TRUE, FALSE, '', 1);


                    EnterCell(5, 1, 'Posting Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 2, 'Document No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 3, 'Vendor Code', TRUE, TRUE, '', 1);
                    EnterCell(5, 4, 'Vendor Name', TRUE, TRUE, '', 1);
                    EnterCell(5, 5, 'GST Vendor Type', TRUE, TRUE, '', 1);// RB-N 18Sep2017
                    EnterCell(5, 6, 'GSTIN', TRUE, TRUE, '', 1);// RB-N 18Sep2017
                    EnterCell(5, 7, 'Doc Type', TRUE, TRUE, '', 1);
                    EnterCell(5, 8, 'External Doc No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 9, 'Vendor Invoice Date', TRUE, TRUE, '', 1);  //RS-AR 151218
                    EnterCell(5, 10, 'Payment Terms Code', TRUE, TRUE, '', 1);
                    EnterCell(5, 11, 'Due Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 12, 'Vendor Posting Group', TRUE, TRUE, '', 1);
                    EnterCell(5, 13, 'Item Category', TRUE, TRUE, '', 1);
                    EnterCell(5, 14, 'Document No.', TRUE, TRUE, '', 1);        //RS-AR 151218
                    EnterCell(5, 15, 'Posting Date', TRUE, TRUE, '', 1);        //RS-AR 151218
                    EnterCell(5, 16, 'GRN No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 17, 'PO No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 18, 'Blanket PO No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 19, 'Cheque No.', TRUE, TRUE, '', 1);
                    EnterCell(5, 20, 'Cheque Date', TRUE, TRUE, '', 1);
                    EnterCell(5, 21, 'Amount', TRUE, TRUE, '', 1);
                    EnterCell(5, 22, 'Narration', TRUE, TRUE, '', 1);

                    CCC := 6;
                END;
                //>>DataHeader


                //>>1
                //Vendor Ledger Entry, Body (2) - OnPreSection()
                //Row+=2;
                vCheckNo := '';
                tNarr := '';
                dCheckDate := 0D;
                recVend.GET("Vendor No.");
                recCheckLedger.RESET;
                recCheckLedger.SETRANGE(recCheckLedger."Document No.", "Document No.");
                IF recCheckLedger.FINDFIRST THEN BEGIN
                    vCheckNo := recCheckLedger."Check No.";
                    dCheckDate := recCheckLedger."Check Date";
                END;

                recPostedNarration.RESET;
                recPostedNarration.SETRANGE(recPostedNarration."Transaction No.", "Transaction No.");
                recPostedNarration.SETRANGE(recPostedNarration."Entry No.", 0);
                IF recPostedNarration.FINDFIRST THEN
                    tNarr := recPostedNarration.Narration;


                //>>DataBody
                EnterCell(CCC, 1, FORMAT("Posting Date"), FALSE, FALSE, '', 2);
                EnterCell(CCC, 2, "Document No.", FALSE, FALSE, '', 1);
                EnterCell(CCC, 3, "Vendor No.", FALSE, FALSE, '', 1);
                EnterCell(CCC, 4, recVend.Name, FALSE, FALSE, '', 1);
                EnterCell(CCC, 5, FORMAT(recVend."GST Vendor Type"), FALSE, FALSE, '', 1);// RB-N 18Sep2017
                EnterCell(CCC, 6, recVend."GST Registration No.", FALSE, FALSE, '', 1);// RB-N 18Sep2017
                EnterCell(CCC, 7, FORMAT("Document Type"), FALSE, FALSE, '', 1);
                EnterCell(CCC, 8, "Vendor Ledger Entry"."External Document No.", FALSE, FALSE, '', 1);
                EnterCell(CCC, 9, '', FALSE, FALSE, '', 1);  //RSPL -- AR 151218
                EnterCell(CCC, 10, recVend."Payment Terms Code", FALSE, FALSE, '', 1);
                EnterCell(CCC, 11, FORMAT("Due Date"), FALSE, FALSE, '', 2);
                EnterCell(CCC, 12, recVend."Vendor Posting Group", FALSE, FALSE, '', 1);
                EnterCell(CCC, 13, '', FALSE, FALSE, '', 1);//Item Category Code
                EnterCell(CCC, 14, '', FALSE, FALSE, '', 1);//Invoice No.
                EnterCell(CCC, 15, '', FALSE, FALSE, '', 1);//RSPL -- AR 151218
                EnterCell(CCC, 16, '', FALSE, FALSE, '', 1);//GRN No.
                EnterCell(CCC, 17, '', FALSE, FALSE, '', 1);//PO No.
                EnterCell(CCC, 18, '', FALSE, FALSE, '', 1);//Blanket PO No.
                EnterCell(CCC, 19, vCheckNo, FALSE, FALSE, '', 1);
                EnterCell(CCC, 20, FORMAT(dCheckDate), FALSE, FALSE, '', 2);
                EnterCell(CCC, 21, FORMAT("Vendor Ledger Entry"."Amount (LCY)"), FALSE, FALSE, '#,#0.00', 0);
                EnterCell(CCC, 22, tNarr, FALSE, FALSE, '', 1);

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
        /* ExcBuffer.CreateBook('', 'Vendor Payment Register');
        ExcBuffer.CreateBookAndOpenExcel('', 'Vendor Payment Register', '', '', USERID);
        ExcBuffer.GiveUserControl; */

        ExcBuffer.CreateNewBook('Vendor Payment Register');
        ExcBuffer.WriteSheet('Vendor Payment Register', CompanyName, UserId);
        ExcBuffer.CloseBook();
        ExcBuffer.SetFriendlyFilename(StrSubstNo('Vendor Payment Register', CurrentDateTime, UserId));
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
        BnkAccLed: Record 271;
        CheNo: Code[20];
        CheDate: Date;
        VLE2: Record 25;

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


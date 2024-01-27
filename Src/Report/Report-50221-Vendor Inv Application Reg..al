report 50221 "Vendor Inv Application Reg."
{
    // 
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 05Jan2019   RB-N         New Report Development
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/VendorInvApplicationReg.rdl';
    Caption = 'Vendor Invoice Application Register';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(CLE; "Vendor Ledger Entry")
        {
            CalcFields = "Amount", "Amount (LCY)";
            DataItemTableView = SORTING("Document No.", "Document Type", "Vendor No.")
                                WHERE("Document Type" = FILTER(Invoice));
            RequestFilterFields = "Entry No.", "Document No.";
            column(EntryNo_CLE; CLE."Entry No.")
            {
            }
            column(DueDate_CLE; CLE."Due Date")
            {
            }
            column(CustomerNo_CLE; CLE."Vendor No.")
            {
            }
            column(PostingDate_CLE; CLE."Posting Date")
            {
            }
            column(DocumentType_CLE; CLE."Document Type")
            {
            }
            column(DocumentNo_CLE; CLE."Document No.")
            {
            }
            column(CustName; CustName)
            {
            }
            column(V_GSTNo; Cust."GST Registration No.")
            {
            }
            column(V_GSTType; Cust."GST Vendor Type")
            {
            }
            column(Cust_PMT; Cust."Payment Terms Code")
            {
            }
            column(V_VdrPstggrp; Cust."Vendor Posting Group")
            {
            }
            column(Dim1; CLE."Global Dimension 1 Code")
            {
            }
            column(Dim2; CLE."Global Dimension 2 Code")
            {
            }
            column(DimText; DimText)
            {
            }
            column(AmountLCY_CLE; ABS("Amount (LCY)"))
            {
            }
            column(SalesLCY_CLE; '')
            {
            }
            column(ProfitLCY_CLE; '')
            {
            }
            column(OriginalAmtLCY_CLE; ABS(CLE."Original Amt. (LCY)"))
            {
            }
            column(Cdper; CdPer)
            {
            }
            column(CdAmt; CdAmt)
            {
            }
            column(TxtFilter; TxtFilter)
            {
            }
            column(PayDocNo; PayDocNo)
            {
            }
            column(PayDocDate; PayDocDate)
            {
            }
            column(BankName; BankName)
            {
            }
            column(ChqNo; ChqNo)
            {
            }
            column(ChqDate; ChqDate)
            {
            }
            column(ClosedbyAmountLCY_CLE; CLE."Closed by Amount (LCY)")
            {
            }
            column(tNarr; tNarr)
            {
            }
            dataitem(APPCLE1; "Vendor Ledger Entry")
            {
                CalcFields = "Amount", "Amount (LCY)";
                DataItemLink = "Entry No." = FIELD("Closed by Entry No.");
                DataItemTableView = SORTING("Entry No.");
                column(EntryNo_APPCLE1; APPCLE1."Entry No.")
                {
                }
                column(ExternalDocumentNo_APPCLE1; APPCLE1."External Document No.")
                {
                }
                column(VendorNo_APPCLE1; APPCLE1."Vendor No.")
                {
                }
                column(PostingDate_APPCLE1; FORMAT(APPCLE1."Posting Date"))
                {
                }
                column(DocumentType_APPCLE1; APPCLE1."Document Type")
                {
                }
                column(DocumentNo_APPCLE1; APPCLE1."Document No.")
                {
                }
                column(Description_APPCLE1; APPCLE1.Description)
                {
                }
                column(ChqNo3; ChqNo3)
                {
                }
                column(ChqDate3; FORMAT(ChqDate3))
                {
                }
                column(ClosedbyAmount_APPCLE1; ABS(APPCLE1."Closed by Amount"))
                {
                }
                column(ClosedbyAmountLCY_APPCLE1; ClosedAmt1)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    //MESSAGE('%1 app 1  %2',APPCLE1."Entry No.",APPCLE1."Document No.");

                    CLEAR(ChqNo3);
                    CLEAR(ChqDate3);

                    VLE2.RESET;
                    IF VLE2.GET("Closed by Entry No.") THEN BEGIN
                        BLE2.RESET;
                        BLE2.SETCURRENTKEY("Document No.");
                        BLE2.SETRANGE("Document No.", VLE2."Document No.");
                        IF BLE2.FINDFIRST THEN BEGIN
                            ChqNo3 := BLE2."Cheque No.";
                            ChqDate3 := BLE2."Cheque Date";
                        END;
                    END ELSE BEGIN
                        BLE2.RESET;
                        BLE2.SETCURRENTKEY("Document No.");
                        BLE2.SETRANGE("Document No.", "Document No.");
                        IF BLE2.FINDFIRST THEN BEGIN
                            ChqNo3 := BLE2."Cheque No.";
                            ChqDate3 := BLE2."Cheque Date";
                        END;
                    END;

                    CLEAR(ClosedAmt1);
                    IF APPCLE1."Closed by Amount (LCY)" <> 0 THEN
                        ClosedAmt1 := ABS(APPCLE1."Closed by Amount (LCY)")
                    ELSE
                        ClosedAmt1 := ABS(CLE."Closed by Amount (LCY)");
                end;
            }
            dataitem(APPCLE2; "Vendor Ledger Entry")
            {
                CalcFields = "Amount", "Amount (LCY)";
                DataItemLink = "Closed by Entry No." = FIELD("Entry No.");
                DataItemTableView = SORTING("Entry No.");
                column(EntryNo_APPCLE; "Entry No.")
                {
                }
                column(ExternalDocumentNo_APPCLE2; APPCLE2."External Document No.")
                {
                }
                column(CustomerNo_APPCLE; "Vendor No.")
                {
                }
                column(PostingDate_APPCLE; FORMAT("Posting Date"))
                {
                }
                column(DocumentType_APPCLE; "Document Type")
                {
                }
                column(DocumentNo_APPCLE; "Document No.")
                {
                }
                column(Description_APPCLE; Description)
                {
                }
                column(CurrencyCode_APPCLE; "Currency Code")
                {
                }
                column(Amount_APPCLE; Amount)
                {
                }
                column(RemainingAmount_APPCLE; "Remaining Amount")
                {
                }
                column(PayDocNo2; PayDocNo2)
                {
                }
                column(PayDocDate2; PayDocDate2)
                {
                }
                column(BankName2; BankName2)
                {
                }
                column(ChqNo2; ChqNo2)
                {
                }
                column(ChqDate2; FORMAT(ChqDate2))
                {
                }
                column(ClosebyAmountLCY_APPCLE2; ABS("Closed by Amount (LCY)"))
                {
                }
                column(ClosedbyAmount_APPCLE2; ABS(APPCLE2."Closed by Amount"))
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //MESSAGE('%1 app 2  %2',APPCLE2."Entry No.",APPCLE2."Document No.");

                    PayDocNo2 := '';
                    CLEAR(PayDocDate2);
                    CLEAR(BankName2);
                    CLEAR(ChqNo2);
                    CLEAR(ChqDate2);

                    PayDocNo2 := "Document No.";
                    PayDocDate2 := "Posting Date";



                    VLE2.RESET;
                    IF VLE2.GET("Closed by Entry No.") THEN BEGIN
                        BLE2.RESET;
                        BLE2.SETCURRENTKEY("Document No.");
                        IF CLE."Document No." <> VLE2."Document No." THEN
                            BLE2.SETRANGE("Document No.", VLE2."Document No.")
                        ELSE
                            BLE2.SETRANGE("Document No.", "Document No.");
                        IF BLE2.FINDFIRST THEN BEGIN
                            ChqNo2 := BLE2."Cheque No.";
                            ChqDate2 := BLE2."Cheque Date";
                        END;
                    END ELSE BEGIN
                        BLE2.RESET;
                        BLE2.SETCURRENTKEY("Document No.");
                        BLE2.SETRANGE("Document No.", "Document No.");
                        IF BLE2.FINDFIRST THEN BEGIN
                            ChqNo2 := BLE2."Cheque No.";
                            ChqDate2 := BLE2."Cheque Date";
                        END;
                    END;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                CustName := '';
                Cust.RESET;
                IF Cust.GET("Vendor No.") THEN
                    CustName := Cust.Name;

                DimText := '';
                Dim.RESET;
                IF Dim.GET('DIVISION', CLE."Global Dimension 1 Code") THEN
                    DimText := Dim.Name;




                PayDocNo := '';
                CLEAR(PayDocDate);
                CLEAR(BankName);
                CLEAR(ChqNo);
                CLEAR(ChqDate);

                VLE2.RESET;
                IF VLE2.GET("Closed by Entry No.") THEN BEGIN
                    BLE2.RESET;
                    BLE2.SETCURRENTKEY("Document No.");
                    BLE2.SETRANGE("Document No.", VLE2."Document No.");
                    IF BLE2.FINDFIRST THEN BEGIN
                        ChqNo := BLE2."Cheque No.";
                        ChqDate := BLE2."Cheque Date";
                    END;
                END;

                CLEAR(tNarr);
                recPostedNarration.RESET;
                recPostedNarration.SETRANGE(recPostedNarration."Transaction No.", "Transaction No.");
                recPostedNarration.SETRANGE(recPostedNarration."Entry No.", 0);
                IF recPostedNarration.FINDFIRST THEN
                    tNarr := recPostedNarration.Narration;
            end;

            trigger OnPreDataItem()
            begin

                SETRANGE("Posting Date", SDate, EDate);

                IF LocFilter <> '' THEN
                    SETFILTER("Location Code", LocFilter);

                IF CusFilter <> '' THEN
                    SETFILTER("Vendor No.", CusFilter);
            end;
        }
    }

    requestpage
    {
        SaveValues = true;

        layout
        {
            area(content)
            {
                field("Start Date"; SDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin

                        EDate := SDate;
                    end;
                }
                field("End Date"; EDate)
                {
                    ApplicationArea = all;
                    trigger OnValidate()
                    begin

                        IF EDate < SDate THEN
                            ERROR('EndDate must be greater than StartDate');
                    end;
                }
                field("Location Filter"; LocFilter)
                {
                    ApplicationArea = all;
                    TableRelation = Location;
                }
                field("Vendor Filter"; CusFilter)
                {
                    ApplicationArea = all;
                    TableRelation = Vendor;
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
            ERROR('Please Specify Start Date');

        IF EDate = 0D THEN
            ERROR('Please Specify End Date');

        TxtFilter := '';
        TxtFilter := 'Date Filter : ' + FORMAT(SDate) + '..' + FORMAT(EDate);

        IF LocFilter <> '' THEN
            TxtFilter += ', Location : ' + LocFilter;

        IF CusFilter <> '' THEN
            TxtFilter += ', Vendor : ' + CusFilter;
    end;

    var
        SDate: Date;
        EDate: Date;
        LocFilter: Code[100];
        CusFilter: Code[100];
        Cust: Record 23;
        CustName: Text;
        Dim: Record 349;
        DimText: Text;
        //PstrLine: Record 13798;
        CdPer: Decimal;
        CdAmt: Decimal;
        TxtFilter: Text;
        AppCLE: Record 25;
        BLE: Record 271;
        PayDocNo: Code[20];
        PayDocDate: Date;
        BankName: Text;
        ChqNo: Code[20];
        ChqDate: Date;
        BankAcc: Record 270;
        "-----------------------RSPL-----------------": Integer;
        AppCLEPay: Record 25;
        BLE2: Record 271;
        PayDocNo2: Code[20];
        PayDocDate2: Date;
        BankName2: Text;
        ChqNo2: Code[20];
        ChqDate2: Date;
        BankAcc2: Record 270;
        ChqNo3: Code[20];
        ChqDate3: Date;
        recPostedNarration: Record "Posted Narration";
        tNarr: Text[250];
        VLE2: Record 25;
        ClosedAmt1: Decimal;
}


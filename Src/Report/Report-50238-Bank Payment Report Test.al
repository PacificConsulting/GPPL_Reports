report 50238 "Bank Payment Report Test"
{
    DefaultLayout = RDLC;

    RDLCLayout = 'Src/Report Layout/BankPaymentReportTest.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Bank Account Ledger Entry"; "Bank Account Ledger Entry")
        {
            DataItemTableView = WHERE(Reversed = FILTER(false),
                                      "Source Code" = FILTER('BANKPYMTV'));
            column(BankReport; BankReport)
            {
            }
            column(BPV_No; "Bank Account Ledger Entry"."Document No.")
            {
            }
            column(BPV_Date; "Bank Account Ledger Entry"."Posting Date")
            {
            }
            column(TransactionType; "Bank Account Ledger Entry"."Document Type")
            {
            }
            column(BeneficiaryCode; "Bank Account Ledger Entry"."Bal. Account No.")
            {
            }
            column(BeneficiaryAccNo; BeneficiaryAccNo)
            {
            }
            column(TransactionAmount; ABS("Bank Account Ledger Entry".Amount))
            {
            }
            column(BeneficiaryName; BeneficiaryName)
            {
            }
            column(CustomerRefNo; "Bank Account Ledger Entry"."Cheque No.")
            {
            }
            column(TransactionDate; TransactionDate)
            {
            }
            column(IFSC; IFSC)
            {
            }
            column(BeneficiaryBankName; BeneficiaryBankName)
            {
            }
            column(BeneficiaryEmailID; BeneficiaryEmailID)
            {
            }
            column(AX_BPV_No; "Bank Account Ledger Entry"."Document No.")
            {
            }
            column(AX_BPV_Date; "Bank Account Ledger Entry"."Posting Date")
            {
            }
            column(AX_Amount; ABS("Bank Account Ledger Entry".Amount))
            {
            }
            column(AX_ValueDate; TODAY)
            {
            }
            column(AX_BeneficiaryName; BeneficiaryName)
            {
            }
            column(AX_BeneficiaryAccNo; BeneficiaryAccNo)
            {
            }
            column(AX_BeneficiaryEmailID; BeneficiaryEmailID)
            {
            }
            column(AX_EmailBody; "Bank Account Ledger Entry"."External Document No.")
            {
            }
            column(AX_DebitAccNo; "Bank Account Ledger Entry"."Bank Account No.")
            {
            }
            column(AX_CRN; "Bank Account Ledger Entry"."Cheque No.")
            {
            }
            column(AX_ReceiverIFSC; IFSC)
            {
            }
            column(AX_MobNo; MobNo)
            {
            }
            dataitem("Vendor Ledger Entry"; "Vendor Ledger Entry")
            {
                DataItemLink = "Document No." = FIELD("Document No.");
                column(EntryNo_VendorLedgerEntry; "Vendor Ledger Entry"."Entry No.")
                {
                }
                column(DocumentNo_VendorLedgerEntry; "Vendor Ledger Entry"."Document No.")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    IF "Bank Account Ledger Entry"."Bal. Account No." <> '' THEN
                        CurrReport.SKIP;
                end;
            }
            dataitem("Cust. Ledger Entry"; "Cust. Ledger Entry")
            {
                DataItemLink = "Document No." = FIELD("Document No.");
                column(EntryNo_CustLedgerEntry; "Cust. Ledger Entry"."Entry No.")
                {
                }
                column(DocumentNo_CustLedgerEntry; "Cust. Ledger Entry"."Document No.")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    IF "Bank Account Ledger Entry"."Bal. Account No." <> '' THEN
                        CurrReport.SKIP;
                end;
            }

            trigger OnAfterGetRecord()
            begin
                CLEAR(BeneficiaryAccNo);
                CLEAR(BeneficiaryBankName);
                CLEAR(IFSC);

                IF "Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Vendor THEN BEGIN//RSPLSUM 30Jun2020
                    RecVendBankAcc.RESET;
                    RecVendBankAcc.SETRANGE("Vendor No.", "Bank Account Ledger Entry"."Bal. Account No.");
                    RecVendBankAcc.SETRANGE(Blocked, FALSE);//RSPLSUM 23Jul2020
                    IF RecVendBankAcc.FINDFIRST THEN BEGIN
                        BeneficiaryAccNo := RecVendBankAcc."Bank Account No.";
                        BeneficiaryBankName := RecVendBankAcc.Name;
                        IFSC := RecVendBankAcc."SWIFT Code";
                    END;
                END;//RSPLSUM 30Jun2020

                //RSPLSUM 30Jun2020>>
                IF "Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer THEN BEGIN
                    RecCustBankAcc.RESET;
                    RecCustBankAcc.SETRANGE("Customer No.", "Bank Account Ledger Entry"."Bal. Account No.");
                    IF RecCustBankAcc.FINDFIRST THEN BEGIN
                        BeneficiaryAccNo := RecCustBankAcc."Bank Account No.";
                        BeneficiaryBankName := RecCustBankAcc.Name;
                        IFSC := RecCustBankAcc."SWIFT Code";
                    END;
                END;
                //RSPLSUM 30Jun2020<<

                CLEAR(BeneficiaryName);
                CLEAR(BeneficiaryEmailID);
                CLEAR(MobNo);
                IF "Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Vendor THEN BEGIN
                    RecVend.RESET;
                    IF RecVend.GET("Bank Account Ledger Entry"."Bal. Account No.") THEN BEGIN
                        BeneficiaryName := RecVend.Name;
                        BeneficiaryEmailID := RecVend."E-Mail";
                        MobNo := RecVend."Phone No.";
                    END;
                END;

                IF "Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::Customer THEN BEGIN
                    RecCust.RESET;
                    IF RecCust.GET("Bank Account Ledger Entry"."Bal. Account No.") THEN BEGIN
                        BeneficiaryName := RecCust.Name;
                        BeneficiaryEmailID := RecCust."E-Mail";
                        MobNo := RecCust."Phone No.";
                    END;
                END;

                IF "Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::"G/L Account" THEN BEGIN
                    RecGLAcc.RESET;
                    IF RecGLAcc.GET("Bank Account Ledger Entry"."Bal. Account No.") THEN BEGIN
                        BeneficiaryName := RecGLAcc.Name;
                    END;
                END;

                IF ("Bank Account Ledger Entry"."Bal. Account Type" = "Bank Account Ledger Entry"."Bal. Account Type"::"G/L Account") AND
                    ("Bank Account Ledger Entry"."Bal. Account No." = '') THEN BEGIN
                    RecGLE.RESET;
                    RecGLE.SETRANGE("Document No.", "Document No.");
                    RecGLE.SETRANGE("Posting Date", "Posting Date");
                    RecGLE.SETFILTER("Source Type", '<>%1', RecGLE."Source Type"::"Bank Account");
                    IF RecGLE.FINDFIRST THEN BEGIN
                        RecGLAcc.RESET;
                        IF RecGLAcc.GET(RecGLE."G/L Account No.") THEN BEGIN
                            BeneficiaryName := RecGLAcc.Name;
                        END;
                    END;
                END;

                CLEAR(TransactionDate);
                IF "Bank Account Ledger Entry"."Cheque Date" <> 0D THEN
                    TransactionDate := "Bank Account Ledger Entry"."Cheque Date"
                ELSE
                    TransactionDate := "Bank Account Ledger Entry"."Posting Date";
            end;

            trigger OnPreDataItem()
            begin
                "Bank Account Ledger Entry".SETRANGE("Bank Account Ledger Entry"."Posting Date", StartDate, EndDate);
                "Bank Account Ledger Entry".SETFILTER("Bank Account Ledger Entry"."Bank Account No.", '%1', BankAccNo);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(BankReport; BankReport)
                {
                    ApplicationArea = all;
                    Caption = 'Bank';
                }
                field(StartDate; StartDate)
                {
                    ApplicationArea = all;
                    Caption = 'Start Date';
                }
                field(EndDate; EndDate)
                {
                    ApplicationArea = all;
                    Caption = 'End Date';
                }
                field(BankAccNo; BankAccNo)
                {
                    ApplicationArea = all;
                    Caption = 'Bank Account No.';
                    TableRelation = "Bank Account";
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
        IF (StartDate = 0D) OR (EndDate = 0D) THEN
            ERROR('Please enter Start Date and End Date');

        IF EndDate < StartDate THEN
            ERROR('End Date should be greater than or equal to Start Date');

        IF BankAccNo = '' THEN
            ERROR('Please enter Bank Account No.');

        PostingDateDF := "Bank Account Ledger Entry".GETFILTER("Bank Account Ledger Entry"."Posting Date");
        BankAccNoDF := "Bank Account Ledger Entry".GETFILTER("Bank Account Ledger Entry"."Bank Account No.");
        IF PostingDateDF <> '' THEN
            ERROR('Posting Date already selected, Please remove from Bank Account Ledger Entry filter');
        IF BankAccNoDF <> '' THEN
            ERROR('Bank Account No. already selected, Please remove from Bank Account Ledger Entry filter');
    end;

    var
        RecVendBankAcc: Record 288;
        RecVend: Record 23;
        TransactionDate: Date;
        BeneficiaryAccNo: Text[30];
        BeneficiaryBankName: Text[50];
        IFSC: Code[20];
        BankReport: Option HDFC,AXIS;
        StartDate: Date;
        EndDate: Date;
        BankAccNo: Code[20];
        PostingDateDF: Text[50];
        BankAccNoDF: Text[100];
        BeneficiaryName: Text[100];
        RecCust: Record 18;
        BeneficiaryEmailID: Text[80];
        MobNo: Text[30];
        RecGLAcc: Record 15;
        RecGLE: Record 17;
        RecCustBankAcc: Record 287;
}


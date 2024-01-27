report 50080 "Payment Voucher GST"
{
    // 17July2017 :: RB-N, New Report format for GST Payment Voucher.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PaymentVoucherGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("Entry No.")
                                WHERE("Source Code" = FILTER('BANKPYMTV'),
                                      "Bal. Account Type" = FILTER("Bank Account"),
                                      "Source Type" = FILTER(Vendor));
            RequestFilterFields = "Document No.";
            column(Quantity_GLEntry; "G/L Entry".Quantity)
            {
            }
            column(BalAccountNo_GLEntry; "G/L Entry"."Bal. Account No.")
            {
            }
            column(Amount_GLEntry; "G/L Entry".Amount)
            {
            }
            column(GlobalDimension1Code_GLEntry; "G/L Entry"."Global Dimension 1 Code")
            {
            }
            column(GlobalDimension2Code_GLEntry; "G/L Entry"."Global Dimension 2 Code")
            {
            }
            column(UserID_GLEntry; "G/L Entry"."User ID")
            {
            }
            column(EntryNo_GLEntry; "G/L Entry"."Entry No.")
            {
            }
            column(GLAccountNo_GLEntry; "G/L Entry"."G/L Account No.")
            {
            }
            column(PostingDate_GLEntry; "G/L Entry"."Posting Date")
            {
            }
            column(DocumentType_GLEntry; "G/L Entry"."Document Type")
            {
            }
            column(DocumentNo_GLEntry; "G/L Entry"."Document No.")
            {
            }
            column(Description_GLEntry; "G/L Entry".Description)
            {
            }
            column(Description26; Description26)
            {
            }
            column(CompName; CompInfo.Name)
            {
            }
            column(CompPicture; CompInfo.Picture)
            {
            }
            column(CompNamePicture; CompInfo."Name Picture")
            {
            }
            column(CompShadedBoxPicture; CompInfo."Shaded Box")
            {
            }
            column(CompGSTNo; CompInfo."GST Registration No.")
            {
            }
            column(nnn; nnn)
            {
            }
            column(VendorName; Ven.Name)
            {
            }
            column(VenAdd1; Ven.Address)
            {
            }
            column(VenAdd2; Ven."Address 2")
            {
            }
            column(VenAdd3; Ven.City + ' - ' + Ven."Post Code")
            {
            }
            column(VenGSTNo; Ven."GST Registration No.")
            {
            }
            column(VenState; State1.Description)
            {
            }
            column(VenStateCode; State1."State Code (GST Reg. No.)")
            {
            }
            column(Amtinwords; Amtinwords[1])
            {
            }
            column(VenCity; Ven.City)
            {
            }
            column(vCGST; vCGST)
            {
            }
            column(vCGSTRate; vCGSTRate)
            {
            }
            column(vSGST; vSGST)
            {
            }
            column(vSGSTRate; vSGSTRate)
            {
            }
            column(vIGST; vIGST)
            {
            }
            column(vIGSTRate; vIGSTRate)
            {
            }
            column(HSNCode15; HSNCode15)
            {
            }
            column(CGST15; CGST15)
            {
            }
            column(SGST15; SGST15)
            {
            }
            column(IGST15; IGST15)
            {
            }
            column(VenGSTRegNo; VenGSTRegNo)
            {
            }
            column(UserName; UserSetup."Full Name")
            {
            }

            trigger OnAfterGetRecord()
            begin

                IF "G/L Entry"."Source Code" <> 'BANKPYMTV' THEN
                    ERROR('This report is applicable for Bank Payment Voucher.');

                nnn += 1;

                //>>26July2017

                CLEAR(Description26);
                IF "G/L Entry"."Document No." = 'BP/793/0717/0147' THEN
                    Description26 := 'Temporary Monsoon Shed'
                ELSE
                    Description26 := "G/L Entry".Description;

                //<<26July2017

                //>>UserDetails
                UserSetup.RESET;
                UserSetup.SETRANGE("User Name", "G/L Entry"."User ID");
                IF UserSetup.FINDFIRST THEN;


                //>> Vendor Details

                Ven.RESET;
                IF Ven.GET("G/L Entry"."Source No.") THEN BEGIN

                    CLEAR(VenGSTRegNo);
                    VenGSTRegNo := Ven."GST Registration No.";

                    IF Ven."GST Vendor Type" = Ven."GST Vendor Type"::Unregistered THEN
                        VenGSTRegNo := FORMAT(Ven."GST Vendor Type");

                END;
                State1.RESET;
                IF State1.GET(Ven."State Code") THEN;

                //>>18July2017 GST
                CLEAR(vCGST);
                CLEAR(vSGST);
                CLEAR(vIGST);
                CLEAR(vCGSTRate);
                CLEAR(vSGSTRate);
                CLEAR(vIGSTRate);
                CLEAR(CGST15);
                CLEAR(SGST15);
                CLEAR(IGST15);
                CLEAR(HSNCode15);

                DetailGSTEntry.RESET;
                DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                IF DetailGSTEntry.FINDSET THEN
                    REPEAT

                        HSNCode15 := DetailGSTEntry."HSN/SAC Code";

                        IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                            vCGST := ABS(DetailGSTEntry."GST Amount");
                            vCGSTRate := DetailGSTEntry."GST %";

                            IF DetailGSTEntry."GST Amount" = 0 THEN
                                CGST15 := ''
                            ELSE
                                CGST15 := ' @ ' + FORMAT(vCGSTRate) + ' %';
                        END;

                        IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                            vSGST := ABS(DetailGSTEntry."GST Amount");
                            vSGSTRate := DetailGSTEntry."GST %";

                            IF DetailGSTEntry."GST Amount" = 0 THEN
                                SGST15 := ''
                            ELSE
                                SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                        END;

                        IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                            vIGST := ABS(DetailGSTEntry."GST Amount");
                            vIGSTRate := DetailGSTEntry."GST %";
                            //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");

                            IF DetailGSTEntry."GST Amount" = 0 THEN
                                IGST15 := ''
                            ELSE
                                IGST15 := ' @ ' + FORMAT(vIGSTRate) + ' %';

                        END;

                    UNTIL DetailGSTEntry.NEXT = 0;
                //<<18July2017 GST


                //>>Amount in Words
                RCheck.InitTextVariable;
                RCheck.FormatNoText(Amtinwords, ("G/L Entry".Amount + vCGST + vSGST + vIGST), '');
                Amtinwords[1] := DELCHR(Amtinwords[1], '=', '*');
            end;

            trigger OnPreDataItem()
            begin

                nnn := 0;
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

    trigger OnPreReport()
    begin

        CompInfo.GET;
        CompInfo.CALCFIELDS(Picture, "Name Picture");//26July2017
    end;

    var
        CompInfo: Record 79;
        nnn: Integer;
        Ven: Record 23;
        State1: Record State;
        RCheck: Report "Check Report";
        Amtinwords: array[2] of Text[150];
        "-----18July2017---GST": Integer;
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        CGST15: Text[50];
        SGST15: Text[50];
        IGST15: Text[50];
        HSNCode15: Code[10];
        VenGSTRegNo: Text[20];
        UserSetup: Record 2000000120;
        Description26: Text[50];
}


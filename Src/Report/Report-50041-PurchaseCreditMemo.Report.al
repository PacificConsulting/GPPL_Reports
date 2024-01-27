report 50041 "Purchase Credit Memo"
{
    // RSPL-CAS-05223-D4F9L3     Sourav Dey    Changes made for avoiding service tax amount incase of TRANSPORT
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseCreditMemo.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(CurrencyCode_SalesCrMemoHeader; "Purch. Cr. Memo Hdr."."Currency Code")
            {
            }
            column(No_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."No.")
            {
            }
            column(Name_Compinfo; CompanyInfo1.Name)
            {
            }
            column(Comp_Location; recLocation.Address + ', ' + recLocation."Address 2" + ', ' + recLocation.City + ' - ' + recLocation."Post Code" + ', ' + recState.Description + ', ' + recCountry.Name + '.')
            {
            }
            column(Comp_Contact; 'Phone: ' + recLocation."Phone No." + ' Email: ' + recLocation."E-Mail")
            {
            }
            column(BuyfromVendorName_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Buy-from Vendor Name")
            {
            }
            column(BuyfromAddress_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Buy-from Address")
            {
            }
            column(BuyfromAddress2_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Buy-from Address 2")
            {
            }
            column(BuyfromCity_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Buy-from City")
            {
            }
            column(TINNo_Vend; 'recVend1."T.I.N. No."')
            {
            }
            column(VendResCenter; VendResCenter + ' - ' + ResCenterName)
            {
            }
            column(VendorCrMemoNo_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Vendor Cr. Memo No.")
            {
            }
            column(AppliestoDocNo_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Applies-to Doc. No.")
            {
            }
            column(LocationName; LocResCenter + ' - ' + LocationName)
            {
            }
            column(LocTinNo; LocTinNo)
            {
            }
            column(AppliesDocDate; AppliesDocDate)
            {
            }
            column(PostingDate_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Posting Date")
            {
            }
            column(Name_UserSetup; UserSetup.Name)
            {
            }
            column(UserName; UserName)
            {
            }
            column(Commments_1; Commments[1])
            {
            }
            column(Commments_2; Commments[2])
            {
            }
            column(Commments_3; Commments[3])
            {
            }
            column(Commments_4; Commments[4])
            {
            }
            column(Commments_5; Commments[5])
            {
            }
            dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE("No." = FILTER(<> 74012210),
                                          Quantity = FILTER(<> 0));
                column(Quantity_SalesCrMemoLine; "Purch. Cr. Memo Line".Quantity)
                {
                }
                column(SrNo; SrNo)
                {
                }
                column(Description_PurchCrMemoLine; "Purch. Cr. Memo Line".Description)
                {
                }
                column(Quantity_PurchCrMemoLine; "Purch. Cr. Memo Line".Quantity)
                {
                }
                column(UnitCost; UnitCost)
                {
                }
                column(LineAmount; LineAmount)
                {
                }
                column(ExciseAmt; ExciseAmt)
                {
                }
                column(TaxAmont; TaxAmont)
                {
                }
                column(ServiceTaxAmt; ServiceTaxAmt)
                {
                }
                column(vTotalAmt; vTotalAmt)
                {
                }
                column(AmountinWords_1; AmountinWords[1])
                {
                }
                column(LineNo; "Line No.")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    vCount := "Purch. Cr. Memo Line".COUNT;
                    SrNo := SrNo + 1;

                    CLEAR(UnitCost);
                    CLEAR(LineAmount);
                    CLEAR(TaxAmont);
                    CLEAR(ExciseAmt);
                    CLEAR(ServiceTaxAmt);


                    IF "Purch. Cr. Memo Hdr."."Currency Code" = '' THEN BEGIN
                        UnitCost := UnitCost + "Purch. Cr. Memo Line"."Direct Unit Cost";
                        LineAmount := LineAmount + "Purch. Cr. Memo Line"."Line Amount";
                        TaxAmont := TaxAmont + 0;// "Purch. Cr. Memo Line"."Tax Amount";

                        //For Service Tax Amount
                        ServiceTaxAmt := 0;
                        /*  recPostedStrOrdDetails.RESET;
                         recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Cr. Memo Hdr."."No.");
                         recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Service Tax");
                         recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code", 'SERVICETAX');
                         IF recPostedStrOrdDetails.FINDFIRST THEN
                             REPEAT
                                 ServiceTaxAmt := ServiceTaxAmt + recPostedStrOrdDetails.Amount;
                             UNTIL recPostedStrOrdDetails.NEXT = 0; */
                        //For Service Tax Amount

                        //For Excise Amount
                        ExciseAmt := 0;
                        /*  PostedStrOrdrDetails1.RESET;
                         PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Invoice No.", "Purch. Cr. Memo Hdr."."No.");
                         PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Tax/Charge Type", PostedStrOrdrDetails1."Tax/Charge Type"::Excise);
                         PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Tax/Charge Code", 'EXCISE');
                         IF PostedStrOrdrDetails1.FINDFIRST THEN
                             REPEAT
                                 ExciseAmt := ExciseAmt + PostedStrOrdrDetails1.Amount;
                             UNTIL PostedStrOrdrDetails1.NEXT = 0; */
                        //For Excise Amount

                    END ELSE BEGIN

                        UnitCost := UnitCost + "Purch. Cr. Memo Line"."Direct Unit Cost" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        LineAmount := (LineAmount + "Purch. Cr. Memo Line"."Line Amount") / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TaxAmont := (TaxAmont + 0/*"Purch. Cr. Memo Line"."Tax Amount"*/) / "Purch. Cr. Memo Hdr."."Currency Factor";

                        //For Service Tax Amount
                        ServiceTaxAmt := 0;
                        /* recPostedStrOrdDetails.RESET;
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Purch. Cr. Memo Hdr."."No.");
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Service Tax");
                        recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code", 'SERVICETAX');
                        IF recPostedStrOrdDetails.FINDFIRST THEN
                            REPEAT
                                ServiceTaxAmt := (ServiceTaxAmt + recPostedStrOrdDetails.Amount) / "Purch. Cr. Memo Hdr."."Currency Factor";
                            UNTIL recPostedStrOrdDetails.NEXT = 0; */
                        //For Service Tax Amount

                        //For Excise Amount
                        ExciseAmt := 0;
                        /*  PostedStrOrdrDetails1.RESET;
                         PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Invoice No.", "Purch. Cr. Memo Hdr."."No.");
                         PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Tax/Charge Type", PostedStrOrdrDetails1."Tax/Charge Type"::Excise);
                         PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Tax/Charge Code", 'EXCISE');
                         IF PostedStrOrdrDetails1.FINDFIRST THEN
                             REPEAT*/
                        ExciseAmt := (ExciseAmt + 0/* PostedStrOrdrDetails1.Amount*/) / "Purch. Cr. Memo Hdr."."Currency Factor";
                        //  UNTIL PostedStrOrdrDetails1.NEXT = 0; 
                        //For Excise Amount
                    END;

                    //>>RSPL/Migration/Rahul****Section
                    recVend1.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
                    IF recVend1."Vendor Posting Group" = 'TRANSPORT' THEN BEGIN
                        //CurrReport.SHOWOUTPUT(FALSE); //19May2017
                        vTotalAmt += LineAmount + TaxAmont + ExciseAmt; //23May2017 TotalIssue
                                                                        //vTotalAmt:=LineAmount+TaxAmont+ExciseAmt;
                    END ELSE
                        vTotalAmt += LineAmount + TaxAmont + ServiceTaxAmt + ExciseAmt; //23May2017 TotalIssue
                                                                                        //vTotalAmt:=LineAmount+TaxAmont+ServiceTaxAmt+ExciseAmt;

                    FormatNoText(AmountinWords, ROUND((vTotalAmt), 0.01), '');//RSPL-CAS-05223-D4F9L3
                end;
            }

            trigger OnAfterGetRecord()
            begin

                recVend1.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.");

                recPPIH.RESET;
                recPPIH.SETRANGE(recPPIH."No.", "Applies-to Doc. No.");
                IF recPPIH.FIND('-') THEN
                    AppliesDocDate := recPPIH."Posting Date"
                ELSE
                    AppliesDocDate := 0D;


                IF "Location Code" <> '' THEN BEGIN
                    recLocation.RESET;
                    recLocation.SETRANGE(recLocation.Code, "Purch. Cr. Memo Hdr."."Location Code");
                    IF recLocation.FINDFIRST THEN;

                    recState.RESET;
                    recState.SETRANGE(recState.Code, recLocation."State Code");
                    IF recState.FINDFIRST THEN;

                    recCountry.RESET;
                    recCountry.SETRANGE(recCountry.Code, recLocation."Country/Region Code");
                    IF recCountry.FINDFIRST THEN;
                END;

                IF "Purch. Cr. Memo Hdr."."Location Code" <> '' THEN BEGIN
                    Location.RESET;
                    Location.SETFILTER(Location.Code, "Purch. Cr. Memo Hdr."."Location Code");
                    IF Location.FINDFIRST THEN
                        LocationName := Location.Name;
                    LocResCenter := Location."Global Dimension 2 Code";
                    LocTinNo := '';// Location."T.I.N. No.";
                END;

                VendResCenter := '';
                ResCenterName := '';
                recVend.RESET;
                recVend.SETRANGE(recVend."No.", "Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
                IF recVend.FINDFIRST THEN BEGIN
                    VendResCenter := recVend."Responsibility Center";
                    recResCenter.RESET;
                    recResCenter.SETRANGE(recResCenter.Code, VendResCenter);
                    IF recResCenter.FINDFIRST THEN
                        ResCenterName := recResCenter.Name;
                END;
                //IF UserSetup.GET(USERID) THEN; //19May2017


                //>>NAV2009
                //Purch. Cr. Memo Hdr., Footer (3) - OnPreSection()
                UserRec.RESET;
                UserRec.SETRANGE("User Name", "Purch. Cr. Memo Hdr."."User ID");//19May2017
                //UserRec.SETRANGE(UserRec."User ID","Purch. Cr. Memo Hdr."."User ID");
                IF UserRec.FINDFIRST THEN
                    UserName := UserRec."Full Name";//19May2017
                                                    //UserName:=UserRec.Name;

                //<<NAV2009

                //>>Robosoft/Rahul***Comments
                i := 0;
                PurchCommentLine.RESET;
                PurchCommentLine.SETRANGE("No.", "No.");
                IF PurchCommentLine.FINDSET THEN
                    REPEAT
                        i += 1;
                        Commments[i] := PurchCommentLine.Comment;
                    UNTIL PurchCommentLine.NEXT = 0;
                //<<
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

    trigger OnInitReport()
    begin

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
    end;

    trigger OnPreReport()
    begin
        InitTextVariable;
    end;

    var
        recLoc: Record 14;
        recVend1: Record 23;
        CompanyInfo1: Record 79;
        Location: Record 14;
        LocationName: Text[50];
        LocResCenter: Code[20];
        LocTinNo: Code[20];
        VendResCenter: Code[20];
        ResCenterName: Text[50];
        recVend: Record 23;
        recResCenter: Record 5714;
        recPPIH: Record 122;
        AppliesDocDate: Date;
        SrNo: Integer;
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        AmountinWords: array[2] of Text[132];
        TotalAmount: Decimal;
        UserName: Text[30];
        //  recPostedStrOrdDetails: Record 13798;
        ServiceTaxAmt: Decimal;
        /// PostedStrOrdrDetails1: Record 13798;
        ExciseAmt: Decimal;
        LineAmount: Decimal;
        TaxAmont: Decimal;
        UnitCost: Decimal;
        recLocation: Record 14;
        recState: Record "State";
        recCountry: Record 9;
        vTotalAmt: Decimal;
        Text026: Label 'Zero';
        Text027: Label 'Hundred';
        Text028: Label '&';
        Text029: Label '%1 results in a written number that is too long.';
        Text032: Label 'One';
        Text033: Label 'Two';
        Text034: Label 'Three';
        Text035: Label 'Four';
        Text036: Label 'Five';
        Text037: Label 'Six';
        Text038: Label 'Seven';
        Text039: Label 'Eight';
        Text040: Label 'Nine';
        Text041: Label 'Ten';
        Text042: Label 'Eleven';
        Text043: Label 'Twelve';
        Text044: Label 'Thirteen';
        Text045: Label 'Fourteen';
        Text046: Label 'Fifteen';
        Text047: Label 'Sixteen';
        Text048: Label 'Seventeen';
        Text049: Label 'Eighteen';
        Text050: Label 'Nineteen';
        Text051: Label 'Twenty';
        Text052: Label 'Thirty';
        Text053: Label 'Forty';
        Text054: Label 'Fifty';
        Text055: Label 'Sixty';
        Text056: Label 'Seventy';
        Text057: Label 'Eighty';
        Text058: Label 'Ninety';
        Text059: Label 'Thousand';
        Text060: Label 'Million';
        Text061: Label 'Billion';
        Text1280000: Label 'Lakh';
        Text1280001: Label 'Crore';
        "---": Integer;
        vCount: Integer;
        UserSetup: Record 91;
        Commments: array[10] of Text[500];
        i: Integer;
        PurchCommentLine: Record 43;
        "----19May2017": Integer;
        UserRec: Record 2000000120;

    // //[Scope('Internal')]
    procedure FormatNoText(var NoText: array[2] of Text[80]; No: Decimal; CurrencyCode: Code[10])
    var
        PrintExponent: Boolean;
        Ones: Integer;
        Tens: Integer;
        Hundreds: Integer;
        Exponent: Integer;
        NoTextIndex: Integer;
        Currency: Record 4;
        TensDec: Integer;
        OnesDec: Integer;
    begin
        CLEAR(NoText);
        NoTextIndex := 1;
        NoText[1] := '';

        IF No < 1 THEN
            AddToNoText(NoText, NoTextIndex, PrintExponent, Text026)
        ELSE BEGIN
            FOR Exponent := 4 DOWNTO 1 DO BEGIN
                PrintExponent := FALSE;
                IF No > 99999 THEN BEGIN
                    Ones := No DIV (POWER(100, Exponent - 1) * 10);
                    Hundreds := 0;
                END ELSE BEGIN
                    Ones := No DIV POWER(1000, Exponent - 1);
                    Hundreds := Ones DIV 100;
                END;
                Tens := (Ones MOD 100) DIV 10;
                Ones := Ones MOD 10;
                IF Hundreds > 0 THEN BEGIN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Hundreds]);
                    AddToNoText(NoText, NoTextIndex, PrintExponent, Text027);
                END;
                IF Tens >= 2 THEN BEGIN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, TensText[Tens]);
                    IF Ones > 0 THEN
                        AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Ones]);
                END ELSE
                    IF (Tens * 10 + Ones) > 0 THEN
                        AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[Tens * 10 + Ones]);
                IF PrintExponent AND (Exponent > 1) THEN
                    AddToNoText(NoText, NoTextIndex, PrintExponent, ExponentText[Exponent]);
                IF No > 99999 THEN
                    No := No - (Hundreds * 100 + Tens * 10 + Ones) * POWER(100, Exponent - 1) * 10
                ELSE
                    No := No - (Hundreds * 100 + Tens * 10 + Ones) * POWER(1000, Exponent - 1);
            END;
        END;

        IF CurrencyCode <> '' THEN BEGIN
            Currency.GET(CurrencyCode);
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + '' /*Currency."Currency Numeric Description"*/);
        END ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, 'Rupees');

        AddToNoText(NoText, NoTextIndex, PrintExponent, Text028);

        TensDec := ((No * 100) MOD 100) DIV 10;
        OnesDec := (No * 100) MOD 10;
        IF TensDec >= 2 THEN BEGIN
            AddToNoText(NoText, NoTextIndex, PrintExponent, TensText[TensDec]);
            IF OnesDec > 0 THEN
                AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[OnesDec]);
        END ELSE
            IF (TensDec * 10 + OnesDec) > 0 THEN
                AddToNoText(NoText, NoTextIndex, PrintExponent, OnesText[TensDec * 10 + OnesDec])
            ELSE
                AddToNoText(NoText, NoTextIndex, PrintExponent, Text026);
        IF (CurrencyCode <> '') THEN
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + ''/* Currency."Currency Decimal Description"*/ + ' ONLY')
        ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, 'Paisa Only');
    end;

    local procedure AddToNoText(var NoText: array[2] of Text[80]; var NoTextIndex: Integer; var PrintExponent: Boolean; AddText: Text[30])
    begin
        PrintExponent := TRUE;

        WHILE STRLEN(NoText[NoTextIndex] + ' ' + AddText) > MAXSTRLEN(NoText[1]) DO BEGIN
            NoTextIndex := NoTextIndex + 1;
            IF NoTextIndex > ARRAYLEN(NoText) THEN
                ERROR(Text029, AddText);
        END;

        NoText[NoTextIndex] := DELCHR(NoText[NoTextIndex] + ' ' + AddText, '<');
    end;

    // //[Scope('Internal')]
    procedure InitTextVariable()
    begin
        OnesText[1] := Text032;
        OnesText[2] := Text033;
        OnesText[3] := Text034;
        OnesText[4] := Text035;
        OnesText[5] := Text036;
        OnesText[6] := Text037;
        OnesText[7] := Text038;
        OnesText[8] := Text039;
        OnesText[9] := Text040;
        OnesText[10] := Text041;
        OnesText[11] := Text042;
        OnesText[12] := Text043;
        OnesText[13] := Text044;
        OnesText[14] := Text045;
        OnesText[15] := Text046;
        OnesText[16] := Text047;
        OnesText[17] := Text048;
        OnesText[18] := Text049;
        OnesText[19] := Text050;

        TensText[1] := '';
        TensText[2] := Text051;
        TensText[3] := Text052;
        TensText[4] := Text053;
        TensText[5] := Text054;
        TensText[6] := Text055;
        TensText[7] := Text056;
        TensText[8] := Text057;
        TensText[9] := Text058;

        ExponentText[1] := '';
        ExponentText[2] := Text059;
        ExponentText[3] := Text1280000;
        ExponentText[4] := Text1280001;
    end;
}


report 50037 "Sales Credit Memo"
{
    // 
    // //17Apr2017 LineNo. Added in the TableBody & UserNamePreparedBy
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesCreditMemo.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;


    dataset
    {
        dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(CompName; CompanyInfo.Name)
            {
            }
            column(DocNo; "Sales Cr.Memo Header"."No.")
            {
            }
            column(Loc_Location; recLoc.Address + ' ' + recLoc."Address 2" + ' ' + recLoc.City + ' ' + recLoc."Post Code" + ' ' + RecState1.Description + ' Phone: ' + recLoc."Phone No." + ' Email: ' + recLoc."E-Mail")
            {
            }
            column(BilltoName; "Sales Cr.Memo Header"."Bill-to Name")
            {
            }
            column(BillAdd1; "Sales Cr.Memo Header"."Bill-to Address")
            {
            }
            column(BillAdd2; "Sales Cr.Memo Header"."Bill-to Address 2")
            {
            }
            column(BillAdd3; "Bill-to City" + ' - ' + "Bill-to Post Code")
            {
            }
            column(CustTINNo; 'recCust1."T.I.N. No."')
            {
            }
            column(CustResp; CustResCenter + ' - ' + ResCenterName)
            {
            }
            column(AppliestoDocNo; "Sales Cr.Memo Header"."Applies-to Doc. No.")
            {
            }
            column(RefNo; "Sales Cr.Memo Header"."External Document No.")
            {
            }
            column(PostingDate; FORMAT("Sales Cr.Memo Header"."Posting Date", 0, '<Day,2>-<Month,2>-<Year4>'))
            {
            }
            column(LocResCenter; LocResCenter + ' -  ' + LocationName)
            {
            }
            column(LocTinNo; LocTinNo)
            {
            }
            column(AppliesDocDate; FORMAT(AppliesDocDate, 0, '<Day,2>-<Month,2>-<Year4>'))
            {
            }
            column(UserName; UserName)
            {
            }
            column(CurrencyCode_SalesCrMemoHeader; "Sales Cr.Memo Header"."Currency Code")
            {
            }
            column(CurrencyFactor_SalesCrMemoHeader; "Sales Cr.Memo Header"."Currency Factor")
            {
            }
            column(EntryTaxDesc; EntryTaxDesc)
            {
            }
            column(TaxDescription; TaxDescription)
            {
            }
            column(AddlTax_CessText; "AddlTax&CessText")
            {
            }
            column(CessText; CessText)
            {
            }
            column(ADDLtext; ADDLtext)
            {
            }
            column(vComment; vComment)
            {
            }
            column(UName; UserSetup.Name)
            {
            }
            dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));
                column(LineNo; "Line No.")
                {
                }
                column(SrNo; SrNo)
                {
                }
                column(SaleCrMemo_Desc; "Sales Cr.Memo Line".Description)
                {
                }
                column(DimValue; recDimVale.Name)
                {
                }
                column(Quantity_SalesCrMemoLine; "Sales Cr.Memo Line".Quantity)
                {
                }
                column(UnitPrice_SalesCrMemoLine; "Sales Cr.Memo Line"."Unit Price")
                {
                }
                column(LineAmount_SalesCrMemoLine; "Sales Cr.Memo Line"."Line Amount")
                {
                }
                column(EntryTax; EntryTax)
                {
                }
                column(TaxDesc; 0/* "Sales Cr.Memo Line"."Tax Amount"*/ - "AddlTax&Cess")
                {
                }
                column(AddlTaxCess; "AddlTax&Cess")
                {
                }
                column(AddlDutyAmount; AddlDutyAmount)
                {
                }
                column(Cess; Cess)
                {
                }
                column(TotAmt; "Sales Cr.Memo Line"."Line Amount" + 0/* "Sales Cr.Memo Line"."Tax Amount"*/ + AddlDutyAmount + Cess + EntryTax)
                {
                }
                column(AmountinWords_1; AmountinWords[1])
                {
                }

                trigger OnAfterGetRecord()
                begin
                    vCount := "Sales Cr.Memo Line".COUNT;
                    //>>1
                    SalesInvoiceDate := 0D;

                    SalesInvoiceHeader1.SETFILTER(SalesInvoiceHeader1."No.", "Sales Cr.Memo Header"."Applies-to Doc. No.");
                    IF SalesInvoiceHeader1.FIND('-') THEN
                        SalesInvoiceDate := SalesInvoiceHeader1."Posting Date";

                    IF "Sales Cr.Memo Header"."Applies-to Doc. No." = '' THEN
                        SalesInvoiceDate := 0D;
                    /*
                    recPostedDocumentDimension.RESET;
                    recPostedDocumentDimension.SETRANGE(recPostedDocumentDimension."Table ID",115);
                    recPostedDocumentDimension.SETRANGE(recPostedDocumentDimension."Document No.","Sales Cr.Memo Line"."Document No.");
                    recPostedDocumentDimension.SETRANGE(recPostedDocumentDimension."Line No.","Sales Cr.Memo Line"."Line No.");
                    recPostedDocumentDimension.SETFILTER(recPostedDocumentDimension."Dimension Code",'%1|%2|%3|%4',
                                                         'DISCOUNT','IND-ADJ','AUTO-ADJ','DISTRIBUTOR');
                    IF recPostedDocumentDimension.FINDFIRST THEN
                     recDimVale.RESET;
                     recDimVale.SETRANGE(recDimVale.Code,recPostedDocumentDimension."Dimension Value Code");
                     IF recDimVale.FINDFIRST THEN;
                    */
                    SrNo := SrNo + 1;

                    //For Additional Duty & Cess Amount
                    CLEAR("AddlTax&Cess");
                    /*  recDetailedTaxEntry.RESET;
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Cr.Memo Line"."Document No.");
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                     recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Cr.Memo Line"."Tax Area Code");
                     IF recDetailedTaxEntry.FINDFIRST THEN
                         REPEAT
                             "AddlTax&Cess" := "AddlTax&Cess" + recDetailedTaxEntry."Tax Amount";
                         UNTIL recDetailedTaxEntry.NEXT = 0; */

                    "AddlTax&CessText" := '';
                    /*  recDetailedTaxEntry.RESET;
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Cr.Memo Line"."Document No.");
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                     recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Cr.Memo Line"."Tax Area Code");
                     IF recDetailedTaxEntry.FINDFIRST THEN BEGIN
                         "AddlTax&CessText" := 'Addl. Tax/Cess @ ' + FORMAT(recDetailedTaxEntry."Tax %") + ' ' + '%';
                     END; */
                    //For Additional Duty & Cess Amount

                    //<<1


                    //>>2

                    //Sales Cr.Memo Line, Body (2) - OnPreSection()
                    //CurrReport.SHOWOUTPUT("Sales Cr.Memo Header"."Currency Code" = '');

                    IF ReturnReasons.GET("Sales Cr.Memo Line"."Return Reason Code") THEN
                        "Return Reason" := ReturnReasons.Description;
                    //<<2

                    CLEAR(TaxDescription);
                    //IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
                    //BEGIN
                    /* DetailedTaxEntry.RESET;
                    DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Sales Cr.Memo Header"."No.");
                    IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                        IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN
                            FormCode := 'with Form ' + DetailedTaxEntry."Form Code";
                        END;
                        TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + ' @ ' + FORMAT(DetailedTaxEntry."Tax %") + ' %' + FormCode;
                    END; */
                    //END;

                    IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
                        vTotAmt += "Sales Cr.Memo Line"."Line Amount" + 0/*"Sales Cr.Memo Line"."Tax Amount"*/ + AddlDutyAmount + Cess + EntryTax
                    ELSE
                        vTotAmt += "Sales Cr.Memo Line"."Line Amount" + 0/* "Sales Cr.Memo Line"."Tax Amount"*/ + AddlDutyAmount + Cess + EntryTax / "Sales Cr.Memo Header"."Currency Factor";

                    /*
                    CLEAR(vTotAmt);
                    vSalesCrLine.RESET;
                    vSalesCrLine.SETRANGE("Document No.","Sales Cr.Memo Header"."No.");
                    IF vSalesCrLine.FINDSET THEN REPEAT
                     vTotAmt+= "Sales Cr.Memo Line"."Line Amount"+"Sales Cr.Memo Line"."Tax Amount";
                    UNTIL vSalesCrLine.NEXT =0;
                    IF "Sales Cr.Memo Header"."Currency Code" ='' THEN
                      vTotAmt := vTotAmt+AddlDutyAmount+Cess+EntryTax
                    ELSE
                      vTotAmt := vTotAmt+AddlDutyAmount+Cess+EntryTax/"Sales Cr.Memo Header"."Currency Factor";
                    */

                end;
            }
            dataitem("Sales Comment Line"; "Sales Comment Line")
            {
                DataItemLink = "No." = FIELD("No.");
                DataItemTableView = SORTING("Document Type", "No.", "Line No.");
                column(Comment; Comment)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    /*
                    UserRec.RESET;
                    UserRec.SETRANGE(UserRec."User ID","Sales Cr.Memo Header"."User ID");
                    IF UserRec.FINDFIRST THEN
                     UserName:=UserRec.Name;
                    */

                end;
            }
            dataitem(Integer; Integer)
            {
                column(number; Number)
                {
                }
                column(TotText; TotText[1])
                {
                }

                trigger OnAfterGetRecord()
                begin
                    vCheck.InitTextVariable;
                    vCheck.FormatNoText(TotText, vTotAmt, "Sales Cr.Memo Header"."Currency Code");
                    TotText[1] := DELCHR(TotText[1], '=', '*');
                end;

                trigger OnPreDataItem()
                begin
                    Integer.SETRANGE(Number, 1, 20 - vCount);
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1
                recLoc.GET("Sales Cr.Memo Header"."Location Code");
                recCust1.GET("Sales Cr.Memo Header"."Bill-to Customer No.");

                recPSIH.RESET;
                recPSIH.SETRANGE(recPSIH."No.", "Applies-to Doc. No.");
                IF recPSIH.FIND('-') THEN
                    AppliesDocDate := recPSIH."Posting Date"
                ELSE
                    AppliesDocDate := 0D;

                //For Additional Duty Amount
                /*   recPostedStrOrdDetails.RESET;
                  recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Header"."No.");
                  recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                  recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code", 'ADDL TAX');
                  IF recPostedStrOrdDetails.FINDFIRST THEN
                      ADDLtext := 'Addl Tax @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%';
                  REPEAT
                      AddlDutyAmount := AddlDutyAmount + recPostedStrOrdDetails.Amount;
                  UNTIL recPostedStrOrdDetails.NEXT = 0; */
                //For Additional Duty Amount

                //For Cess
                /* PostedStrOrdrDetails1.RESET;
                PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Invoice No.", "Sales Cr.Memo Header"."No.");
                PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Tax/Charge Type", PostedStrOrdrDetails1."Tax/Charge Type"::"Other Taxes");
                PostedStrOrdrDetails1.SETFILTER(PostedStrOrdrDetails1."Tax/Charge Code", '%1|%2', 'Cess', 'C1');
                IF PostedStrOrdrDetails1.FINDFIRST THEN
                    REPEAT
                        Cess := Cess + PostedStrOrdrDetails1.Amount;
                    UNTIL PostedStrOrdrDetails1.NEXT = 0; */

                IF Cess = 0 THEN
                    CessText := ''
                ELSE
                    //CessText :='Cess @'+FORMAT(PostedStrOrdrDetails1."Calculation Value")+' '+'%';
                    CessText := 'Cess';
                //For Cess

                //For Entry Tax
                CLEAR(EntryTax);
                /*  PostedStrOrdrLineDetails.RESET;
                 PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Sales Cr.Memo Header"."No.");
                 PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                 PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Group", 'ENTRYTAX');
                 IF PostedStrOrdrLineDetails.FINDFIRST THEN
                     REPEAT
                         EntryTax := EntryTax + PostedStrOrdrLineDetails.Amount;
                     UNTIL PostedStrOrdrLineDetails.NEXT = 0;
  */
                IF EntryTax = 0 THEN
                    EntryTaxDesc := ''
                ELSE
                    EntryTaxDesc := 'Entry Tax';
                //For Entry Tax
                //<<1


                //>>2
                //Sales Cr.Memo Header, Header (2) - OnPreSection()
                Location.RESET;
                Location.SETFILTER(Location.Code, "Sales Cr.Memo Header"."Location Code");
                IF Location.FINDFIRST THEN
                    LocationName := Location.Name;
                LocResCenter := Location."Global Dimension 2 Code";
                LocTinNo := '';// Location."T.I.N. No.";

                CustResCenter := '';
                ResCenterName := '';
                recCust.RESET;
                recCust.SETRANGE(recCust."No.", "Sales Cr.Memo Header"."Sell-to Customer No.");
                IF recCust.FINDFIRST THEN BEGIN
                    CustResCenter := recCust."Responsibility Center";
                    recResCenter.RESET;
                    recResCenter.SETRANGE(recResCenter.Code, CustResCenter);
                    IF recResCenter.FINDFIRST THEN
                        ResCenterName := recResCenter.Name;
                END;
                //<<2


                //>>3

                //Sales Cr.Memo Header, Footer (4) - OnPreSection()
                UserRec.RESET;
                //UserRec.SETRANGE(UserRec."User ID","Sales Cr.Memo Header"."User ID");//
                UserRec.SETRANGE("User Name", "Sales Cr.Memo Header"."User ID");//24Mar2017
                IF UserRec.FINDFIRST THEN
                    UserName := UserRec."Full Name";//24Mar2017
                                                    //UserName:=UserRec.Name;
                                                    //<<3
                                                    //MESSAGE('%1',"Sales Cr.Memo Header"."Currency Factor");

                //>>Robosoft/Migration/Rahul
                CLEAR(vComment);
                vSalesCommment.RESET;
                vSalesCommment.SETRANGE("No.", "No.");
                IF vSalesCommment.FINDSET THEN
                    REPEAT
                        vComment := vComment + ' ' + vSalesCommment.Comment;
                    UNTIL vSalesCommment.NEXT = 0;

                IF UserSetup.GET(USERID) THEN;
                //<<
            end;

            trigger OnPreDataItem()
            begin

                //>>1
                CompanyInfo.GET;

                UserSetup.GET(USERID);

                Location.RESET;
                Location.SETFILTER(Location.Code, UserSetup."Location Filter");
                IF Location.FIND('-') THEN
                    LocCode := COPYSTR(Location.Name, 7, 30);
                //<<1
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
        //>>1
        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
        //<<1
    end;

    trigger OnPreReport()
    begin
        //InitTextVariable;
    end;

    var
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
        UserSetup: Record 91;
        CompanyInfo: Record 79;
        Location: Record 14;
        LocCode: Text[50];
        InvoiceDate: Date;
        SalesInvHeader: Record 112;
        SalesInvLine: Record 113;
        ReturnReasons: Record 6635;
        "Return Reason": Text[50];
        Item: Record 27;
        ItemName: Text[50];
        NetAmt: Decimal;
        Total: Decimal;
        TotInvQty: Integer;
        TotRejQty: Integer;
        TotQtyRecd: Integer;
        CompanyInfo1: Record 79;
        RecState1: Record State;
        RespCentre: Record 5714;
        SalesInvoiceHeader1: Record 112;
        SalesInvoiceDate: Date;
        QtyPackUOM: Decimal;
        ItemRec: Record 27;
        UserRec: Record 2000000120;
        UserName: Text[50];
        AddlDutyAmount: Decimal;
        //  PostedStrOrdrDetails: Record 13798;
        FOCAMOUNT: Decimal;
        "Trans.Subsidy": Decimal;
        "Trade Discount": Decimal;
        "Cash Discount": Decimal;
        "Invoice Discount": Decimal;
        "InvoiceNo.": Code[20];
        recGLentry: Record 17;
        Cess: Decimal;
        FreightCharges: Decimal;
        LocationName: Text[30];
        reciTEMUOM: Record 5404;
        SrNo: Integer;
        recCust: Record 18;
        recResCenter: Record 5714;
        CustResCenter: Code[10];
        ResCenterName: Text[30];
        LocResCenter: Code[10];
        LocTinNo: Code[20];
        recLoc: Record 14;
        recCust1: Record 18;
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        AmountinWords: array[2] of Text[132];
        TotalAmount: Decimal;
        recPSIH: Record 112;
        AppliesDocDate: Date;
        // recPostedStrOrdDetails: Record 13798;
        ADDLtext: Text[30];
        //  PostedStrOrdrDetails1: Record 13798;
        CessText: Text[30];
        recDimVale: Record 349;
        //  DetailedTaxEntry: Record 16522;
        FormCode: Code[20];
        TaxDescription: Text[50];
        "AddlTax&Cess": Decimal;
        // recDetailedTaxEntry: Record 16522;
        "AddlTax&CessText": Text[30];
        EntryTaxDesc: Text[30];
        EntryTax: Decimal;
        // PostedStrOrdrLineDetails: Record 13798;
        "--": Integer;
        vCount: Integer;
        vCheck: Report "Check Report";
        TotText: array[5] of Text;
        vTotAmt: Decimal;
        vSalesCrLine: Record 115;
        vSalesCommment: Record 44;
        vComment: Text;
        vUser: Record 2000000120;
        "---17Apr2017": Integer;
        User17: Record 2000000120;
        UserNamee: Text[80];

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
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + ''/*Currency."Currency Decimal Description" */+ ' ONLY')
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

    //[Scope('Internal')]
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


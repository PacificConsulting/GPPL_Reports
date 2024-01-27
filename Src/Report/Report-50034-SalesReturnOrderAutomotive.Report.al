report 50034 "Sales Return Order-Automotive"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SalesReturnOrderAutomotive.rdl';
    Caption = 'Sales Return Order-Automotive';
    PreviewMode = PrintLayout;
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
            column(LocAdd; Location.Address + ' ' + Location."Address 2" + ' ' + Location.City + ' ' + Location."Post Code" + ' ' + RecState1.Description + ' Phone: ' + Location."Phone No." + ' Email: ' + Location."E-Mail")
            {
            }
            column(Title; Title)
            {
            }
            column(BillName; "Bill-to Name")
            {
            }
            column(BillAdd; "Bill-to Address")
            {
            }
            column(BillAdd2; "Bill-to Address 2")
            {
            }
            column(BillAdd3; "Bill-to City" + ' - ' + "Bill-to Post Code" + ', ' + recCountry.Name + '.')
            {
            }
            column(CustTINNo; 'recCust1."T.I.N. No."')
            {
            }
            column(CustResp; CustResCenter + ' - ' + ResCenterName)
            {
            }
            column(PartyRejInvNo; "External Document No.")
            {
            }
            column(OurInvNo; "Applies-to Doc. No.")
            {
            }
            column(DocNo; "No.")
            {
            }
            column(DocDate; FORMAT("Sales Cr.Memo Header"."Posting Date", 0, '<Day,2>-<Month,2>-<Year4>'))
            {
            }
            column(LocName; LocResCenter + ' - ' + Location.Name)
            {
            }
            column(LocTINNo; 'Location."T.I.N. No."')
            {
            }
            column(OurInvDate; FORMAT(AppliedInvoiceDate, 0, '<Day,2>-<Month,2>-<Year4>'))
            {
            }
            column(UserName; UserName)
            {
            }
            dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = WHERE(Type = FILTER(Item | 'Charge (Item)'), Quantity = FILTER(<> 0));
                column(LineDocNo;
                "Sales Cr.Memo Line"."Document No.")
                {
                }
                column(ItemName; ItemName)
                {
                }
                column(UOM; "Unit of Measure")
                {
                }
                column(QtyPack; "Sales Cr.Memo Line"."Qty. per Unit of Measure")
                {
                }
                column(LineExcise; '"Sales Cr.Memo Line"."Excise Prod. Posting Group"')
                {
                }
                column(LineLot; "Sales Cr.Memo Line"."Lot No.")
                {
                }
                column(LineQty; "Sales Cr.Memo Line".Quantity)
                {
                }
                column(LineQtyBase; "Sales Cr.Memo Line"."Quantity (Base)")
                {
                }
                column(MRPPrice; MRPPrice)
                {
                }
                column(BillingPrice; BillingPrice + ProdDisc)
                {
                }
                column(ChgDisc; ((chargeamttotal) / "Quantity (Base)") - ProdDisc)
                {
                }
                column(NetBillingPrice; (BillingPrice + (chargeamttotal) / "Quantity (Base)"))
                {
                }
                column(LineTotalAmt; "Sales Cr.Memo Line"."Line Amount" + 0/*"Sales Cr.Memo Line"."Excise Amount"*/ + 0/*"Sales Cr.Memo Line"."Charges To Customer"*/)
                {
                }
                column(LineBasicRate; "Sales Cr.Memo Line"."Unit Price" / "Sales Cr.Memo Line"."Qty. per Unit of Measure")
                {
                }
                column(LineBEDAmt; 0)// "BED Amount")
                {
                }
                column(LineSheAmt; 0)//"SHE Cess Amount")
                {
                }
                column(LineECessAmt; 0)// "eCess Amount")
                {
                }
                column(LineAmt; "Sales Cr.Memo Line"."Line Amount")
                {
                }
                column(OtherDisc; TotalDisc + "Trans.Subsidy")
                {
                }
                column(CashDisc; "Cash Discount")
                {
                }
                column(InvDisc; "Invoice Discount")
                {
                }
                column(FOCAMOUNT; FOCAMOUNT)
                {
                }
                column(EntryTaxDesc; EntryTaxDesc)
                {
                }
                column(EntryTaxAmt; EntryTax)
                {
                }
                column(CashDiscText; CashDiscText)
                {
                }
                column(CashDiscAmt; ABS(CashDiscountAmt))
                {
                }
                column(TaxDescription; TaxDescription)
                {
                }
                column(TaxDescAmt; 0)//"Sales Cr.Memo Line"."Tax Amount")
                {
                }
                column(AddlTaxCText; "AddlTax&CessText")
                {
                }
                column(AddlTaxCAmt; "AddlTax&Cess")
                {
                }
                column(ADDTAXText; ADDTAXText)
                {
                }
                column(AddlDutyAmount; AddlDutyAmount)
                {
                }
                column(CessText; CessText)
                {
                }
                column(CessAmt; Cess)
                {
                }
                column(FreightCharges; FreightCharges)
                {
                }
                column(RoundOffAmnt; RoundOffAmnt)
                {
                }
                column(TotalAmtinWords; DescriptionLineTot[1])
                {
                }
                column(LineTaxAmt; 0)///"Sales Cr.Memo Line"."Tax Amount")
                {
                }
                column(Total; Total)
                {
                }
                column(NAH; NAH)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    SalesInvoiceDate := 0D;

                    SalesInvoiceHeader1.SETFILTER(SalesInvoiceHeader1."No.", "Sales Cr.Memo Header"."Applies-to Doc. No.");
                    IF SalesInvoiceHeader1.FIND('-') THEN
                        SalesInvoiceDate := SalesInvoiceHeader1."Posting Date";

                    IF "Sales Cr.Memo Header"."Applies-to Doc. No." = '' THEN
                        SalesInvoiceDate := 0D;

                    // IF "Sales Cr.Memo Line"."Price Inclusive of Tax" THEN
                    //     IF "Sales Cr.Memo Line"."Abatement %" = 0 THEN
                    MRPPrice := 0;
                    // ELSE
                    //     MRPPrice := "Sales Cr.Memo Line"."MRP Price";

                    BillingPrice := ("Sales Cr.Memo Line"."Unit Price" / "Sales Cr.Memo Line"."Qty. per Unit of Measure")
                                    + 0/* ("Sales Cr.Memo Line"."Excise Amount" *// ("Sales Cr.Memo Line"."Qty. per Unit of Measure" *
                                    "Sales Cr.Memo Line".Quantity);

                    chargeamttotal := 0;
                    /*  PostedStrOrdDetails.RESET;
                     PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Header"."No.");
                     PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Line No.", "Line No.");
                     PostedStrOrdDetails.SETRANGE(PostedStrOrdDetails."Tax/Charge Type", PostedStrOrdDetails."Tax/Charge Type"::Charges);
                     PostedStrOrdDetails.SETFILTER(PostedStrOrdDetails."Tax/Charge Group", '<>%1&<>%2&<>%3', 'CASHDISC', 'CCFC', 'TPTSUBSIDY');//RSPL
                     IF PostedStrOrdDetails.FINDFIRST THEN
                         REPEAT
                             chargeamttotal += PostedStrOrdDetails.Amount;
                         UNTIL PostedStrOrdDetails.NEXT = 0; */

                    //EBT STIVAN ---(31/08/2011)----------------------------------------------START

                    CLEAR(AddlDutyAmount);
                    CLEAR(FOCAMOUNT);
                    CLEAR("Trans.Subsidy");
                    CLEAR("Trade Discount");
                    CLEAR("Cash Discount");
                    CLEAR("Invoice Discount");
                    CLEAR(Cess);
                    CLEAR(EntryTax);
                    CLEAR(FreightCharges);
                    CLEAR(TotalDisc);

                    //For Additional Duty Amount
                    /*
                    recPostedStrOrdDetails.RESET;
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.","Sales Cr.Memo Line"."Document No.");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type",recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                    recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code",'ADDL TAX');
                    IF recPostedStrOrdDetails.FINDFIRST THEN
                    ADDTAXText:='Addl Tax @ '+FORMAT(recPostedStrOrdDetails."Calculation Value")+' '+'%';
                    REPEAT
                    AddlDutyAmount := AddlDutyAmount + recPostedStrOrdDetails.Amount;
                    UNTIL recPostedStrOrdDetails.NEXT = 0;
                    */
                    //For Additional Duty Amount

                    //For Cess
                    /*
                    PostedStrOrdrDetails1.RESET;
                    PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Invoice No.","Sales Cr.Memo Line"."Document No.");
                    PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Tax/Charge Type",PostedStrOrdrDetails1."Tax/Charge Type"::"Other Taxes");
                    PostedStrOrdrDetails1.SETFILTER(PostedStrOrdrDetails1."Tax/Charge Code",'%1|%2','Cess','C1');
                    IF PostedStrOrdrDetails1.FINDFIRST THEN
                    REPEAT
                    Cess := Cess + PostedStrOrdrDetails1.Amount;
                    UNTIL PostedStrOrdrDetails1.NEXT = 0;

                    IF Cess = 0 THEN
                       CessText := ''
                    ELSE
                       CessText:='Cess @'+FORMAT(PostedStrOrdrDetails1."Calculation Value")+' '+'%';
                       */
                    //For Cess

                    //For Additional Duty & Cess Amount
                    CLEAR("AddlTax&Cess");
                    /* recDetailedTaxEntry.RESET;
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                    recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Cr.Memo Line"."Tax Area Code");
                    IF recDetailedTaxEntry.FINDFIRST THEN
                        REPEAT
                            "AddlTax&Cess" := "AddlTax&Cess" + recDetailedTaxEntry."Tax Amount";
                        UNTIL recDetailedTaxEntry.NEXT = 0; */

                    "AddlTax&CessText" := '';
                    /* recDetailedTaxEntry.RESET;
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                    recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Cr.Memo Line"."Tax Area Code");
                    IF recDetailedTaxEntry.FINDFIRST THEN BEGIN
                        "AddlTax&CessText" := 'Addl. Tax/Cess @' + FORMAT(recDetailedTaxEntry."Tax %") + ' ' + '%';
                    END; */
                    //For Additional Duty & Cess Amount

                    //For FOCAMOUNT
                    recGLentry.RESET;
                    recGLentry.SETFILTER("G/L Account No.", '%1|%2', '75002035', '51010030');
                    recGLentry.SETFILTER("Document Type", 'Credit Memo');
                    recGLentry.SETRANGE(recGLentry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    IF recGLentry.FINDFIRST THEN
                        REPEAT
                            FOCAMOUNT := FOCAMOUNT + recGLentry.Amount;
                        UNTIL recGLentry.NEXT = 0;
                    //For FOCAMOUNT

                    //For Transport Subsidy
                    /*   PostedStrOrdrDetails.RESET;
                      PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Invoice No.", "Sales Cr.Memo Line"."Document No.");
                      PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Type", PostedStrOrdrDetails."Tax/Charge Type"::Charges);
                      PostedStrOrdrDetails.SETRANGE(PostedStrOrdrDetails."Tax/Charge Group", 'TPTSUBSIDY');
                      IF PostedStrOrdrDetails.FINDFIRST THEN
                          REPEAT
                              "Trans.Subsidy" := "Trans.Subsidy" + PostedStrOrdrDetails.Amount;
                          UNTIL PostedStrOrdrDetails.NEXT = 0; */
                    //For Transport Subsidy

                    //For TOTAL Discount
                    recGLentry.RESET;
                    recGLentry.SETFILTER("G/L Account No.", '75003040');
                    recGLentry.SETFILTER("Document Type", 'Credit Memo');
                    recGLentry.SETRANGE(recGLentry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    IF recGLentry.FINDFIRST THEN
                        REPEAT
                            TotalDisc := TotalDisc + recGLentry.Amount;
                        UNTIL recGLentry.NEXT = 0;
                    //For TOTAL Discount

                    //For CASH Discount
                    recGLentry.RESET;
                    recGLentry.SETFILTER("G/L Account No.", '75003050');
                    recGLentry.SETFILTER("Document Type", 'Credit Memo');
                    recGLentry.SETRANGE(recGLentry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    IF recGLentry.FINDFIRST THEN
                        "Cash Discount" := recGLentry.Amount;
                    //For CASH Discount


                    //For Freight Charges
                    recGLentry.RESET;
                    recGLentry.SETFILTER("G/L Account No.", '75001010');
                    recGLentry.SETFILTER("Document Type", 'Credit Memo');
                    recGLentry.SETRANGE(recGLentry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    IF recGLentry.FINDFIRST THEN
                        FreightCharges := recGLentry.Amount;
                    //For Freight Charges

                    //For Entry Tax
                    CLEAR(EntryTaxDesc);
                    /*   recPostedStrOrdDetails1.RESET;
                      recPostedStrOrdDetails1.SETRANGE("Invoice No.", "Sales Cr.Memo Header"."No.");
                      recPostedStrOrdDetails1.SETRANGE("Tax/Charge Type", recPostedStrOrdDetails1."Tax/Charge Type"::Charges);
                      recPostedStrOrdDetails1.SETRANGE("Tax/Charge Group", 'ENTRYTAX');
                      IF recPostedStrOrdDetails1.FINDSET THEN BEGIN
                          EntryTaxDesc := 'EntryTax @ ' + FORMAT(recPostedStrOrdDetails1."Calculation Value") + ' ' + '%';
                      END; */


                    recGLentry.RESET;
                    recGLentry.SETFILTER("G/L Account No.", '%1|%2|%3', '32170600', '74012050', '75001030');
                    recGLentry.SETFILTER("Document Type", 'Credit Memo');
                    recGLentry.SETRANGE(recGLentry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    IF recGLentry.FINDFIRST THEN
                        EntryTax := recGLentry.Amount;
                    //For Entry Tax

                    //EBT STIVAN ---(31/08/2011)------------------------------------------------END

                    SalesCreditMemoLine.RESET;
                    SalesCreditMemoLine.SETRANGE(SalesCreditMemoLine."Document No.", "Sales Cr.Memo Line"."Document No.");
                    SalesCreditMemoLine.SETRANGE(SalesCreditMemoLine.Type, SalesCreditMemoLine.Type::"G/L Account");
                    SalesCreditMemoLine.SETRANGE(SalesCreditMemoLine."No.", RoundOffAcc);
                    IF SalesCreditMemoLine.FINDFIRST THEN BEGIN
                        RoundOffAmnt := SalesCreditMemoLine."Line Amount";
                    END;

                    //RSPL
                    ProdDisc := 0;
                    //IF "Posting Date" > 071813D THEN
                    IF "Posting Date" > 20180713D THEN //23Mar2017
                    BEGIN
                        CLEAR(ProdDisc);
                        MRPMaster.RESET;
                        MRPMaster.SETRANGE(MRPMaster."Item No.", "No.");
                        MRPMaster.SETRANGE(MRPMaster."Lot No.", "Lot No.");
                        IF MRPMaster.FINDFIRST THEN BEGIN
                            ProdDisc := MRPMaster."National Discount";
                        END;
                    END;
                    /*
                    cashDisc:=0;
                    recPosredStrOrdeLine.RESET;
                    recPosredStrOrdeLine.SETRANGE(recPosredStrOrdeLine."Invoice No.","Document No.");
                    recPosredStrOrdeLine.SETRANGE(recPosredStrOrdeLine."Line No.","Line No.");
                    recPosredStrOrdeLine.SETRANGE(recPosredStrOrdeLine."Tax/Charge Type",recPosredStrOrdeLine."Tax/Charge Type"::Charges);
                    recPosredStrOrdeLine.SETRANGE(
                    IF
                    */
                    //RSPL
                    //<<1

                    //>>2

                    //Sales Cr.Memo Line, Body (2) - OnPreSection()
                    IF ReturnReasons.GET("Sales Cr.Memo Line"."Return Reason Code") THEN
                        "Return Reason" := ReturnReasons.Description;

                    IF Item.GET("Sales Cr.Memo Line"."No.") THEN
                        ItemName := Item.Description;
                    //<<2

                    //>>3

                    //Sales Cr.Memo Line, Footer (3) - OnPreSection()
                    TaxDescr := 'Tax Descr';
                    /* DetailedTaxEntry.RESET;
                    DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Sales Cr.Memo Header"."No.");
                    IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                        IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN
                            FormCode := 'with Form ' + DetailedTaxEntry."Form Code";
                        END;
                        TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                    END; */

                    NetAmt := "Sales Cr.Memo Line"."Line Amount" + "Sales Cr.Memo Line"."Line Discount Amount" + 0/*"Sales Cr.Memo Line"."Excise Amount"*/
                    + 0/* "Sales Cr.Memo Line"."Tax Amount"*/ + 0/* "Sales Cr.Memo Line"."Charges To Customer" */+ RoundOffAmnt;
                    Total := NetAmt;


                    //FormatNoText(DescriptionLineTot,ROUND((Total),0.01),'');

                    //<<3

                    //>>AmountinWords 24Mar2017

                    /*  RCheck.InitTextVariable;
                     RCheck.FormatNoText(DescriptionLineTot, ROUND((Total), 0.01), ''); */
                    DescriptionLineTot[1] := DELCHR(DescriptionLineTot[1], '=', '*');

                    //<<AmountinWords 24Mar2017

                end;

                trigger OnPreDataItem()
                begin

                    NAH := COUNT;//23Mar2017
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                recCust1.GET("Sales Cr.Memo Header"."Bill-to Customer No.");

                recSIH.RESET;
                recSIH.SETRANGE(recSIH."No.", "Sales Cr.Memo Header"."Applies-to Doc. No.");
                IF recSIH.FINDFIRST THEN
                    AppliedInvoiceDate := recSIH."Posting Date";

                CustPostSetup.RESET;
                CustPostSetup.FINDFIRST;
                RoundOffAcc := CustPostSetup."Invoice Rounding Account";


                //InitTextVariable;
                //RSPL
                CashDiscountAmt := 0;
                /*  recPostedStrOrdDetails.RESET;
                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "No.");
                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::Charges);
                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Group", 'CASHDISC');
                 IF recPostedStrOrdDetails.FINDFIRST THEN
                     REPEAT
                         CashDiscountAmt += recPostedStrOrdDetails.Amount;
                     UNTIL recPostedStrOrdDetails.NEXT = 0; */

                IF CashDiscountAmt <> 0 THEN
                    CashDiscText := 'Cash Discount @ ' + ''/* FORMAT(recPostedStrOrdDetails."Calculation Value")*/ + ' ' + '%'
                ELSE
                    CashDiscText := '';
                //RSPL
                //<<1

                //>>2

                //Sales Cr.Memo Header, Header (1) - OnPreSection()
                Location.RESET;
                Location.GET("Sales Cr.Memo Header"."Location Code");
                //<<2

                //>>3

                //Sales Cr.Memo Header, Header (2) - OnPreSection()
                IF ("Cancelled Invoice" = TRUE) THEN
                    Title := 'Cancel Invoice'
                ELSE
                    IF ("Cancelled Invoice" = FALSE) THEN
                        Title := 'Sales Return Rejection Memo';



                //EBT STIVAN ---- (071211) --------------------------------------------START
                Location.RESET;
                Location.SETFILTER(Location.Code, "Sales Cr.Memo Header"."Location Code");
                IF Location.FINDFIRST THEN
                    LocationName := Location.Name;
                LocResCenter := Location."Global Dimension 2 Code";
                //EBT STIVAN ---- (071211) ----------------------------------------------END


                CustResCenter := '';
                ResCenterName := '';
                recCust.RESET;
                recCust.SETRANGE(recCust."No.", "Sales Cr.Memo Header"."Bill-to Customer No.");
                IF recCust.FINDFIRST THEN BEGIN
                    CustResCenter := recCust."Responsibility Center";
                    recResCenter.RESET;
                    recResCenter.SETRANGE(recResCenter.Code, CustResCenter);
                    IF recResCenter.FINDFIRST THEN
                        ResCenterName := recResCenter.Name;
                END;

                recCountry.RESET;
                recCountry.SETRANGE(recCountry.Code, "Sales Cr.Memo Header"."Bill-to Country/Region Code");
                IF recCountry.FINDFIRST THEN;

                //<<3

                //>>4

                //Sales Cr.Memo Header, Footer (4) - OnPreSection()
                UserRec.RESET;
                //UserRec.SETRANGE(UserRec."User ID","Sales Cr.Memo Header"."User ID");
                UserRec.SETRANGE("User Name", "Sales Cr.Memo Header"."User ID");//23Mar2017
                IF UserRec.FINDFIRST THEN
                    UserName := UserRec."Full Name";//23Mar2017
                                                    //UserName:=UserRec.Name;
                                                    //<<4
            end;

            trigger OnPreDataItem()
            begin

                //>>1

                CompanyInfo.GET;
                //UserSetup.GET(USERID);
                IF UserSetup.GET(USERID) THEN;//23Mar2017

                recRespCenter.RESET;
                recRespCenter.SETRANGE(recRespCenter.Code, UserSetup."Sales Resp. Ctr. Filter");
                IF recRespCenter.FINDFIRST THEN
                    LocCode := recRespCenter."Location Code";
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

        TotInvQty := 0;
        TotRejQty := 0;
        TotQtyRecd := 0;

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
        //<<1
    end;

    var
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
        // PostedStrOrdrDetails: Record 13798;
        FOCAMOUNT: Decimal;
        "Trans.Subsidy": Decimal;
        "Trade Discount": Decimal;
        "Cash Discount": Decimal;
        "Invoice Discount": Decimal;
        "InvoiceNo.": Code[20];
        recGLentry: Record 17;
        Cess: Decimal;
        EntryTax: Decimal;
        FreightCharges: Decimal;
        LocationName: Text;
        reciTEMUOM: Record 5404;
        recRespCenter: Record 5714;
        TotalDisc: Decimal;
        recSIH: Record 112;
        AppliedInvoiceDate: Date;
        MRPPrice: Decimal;
        BillingPrice: Decimal;
        chargeamttotal: Decimal;
        // PostedStrOrdDetails: Record 13798;
        EntryTaxDesc: Text[50];
        // recPostedStrOrdDetails: Record 13798;
        // recPostedStrOrdDetails1: Record 13798;
        TaxDescr: Code[10];
        TaxDescription: Text[50];
        FormCode: Code[30];
        // DetailedTaxEntry: Record 16522;
        ADDTAXText: Text[30];
        CessText: Text[30];
        // PostedStrOrdrDetails1: Record 13798;
        RoundOffAmnt: Decimal;
        CustPostSetup: Record 92;
        RoundOffAcc: Text[30];
        SalesCreditMemoLine: Record 115;
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        DescriptionLineTot: array[2] of Text[132];
        CustResCenter: Code[10];
        ResCenterName: Text;
        recCust: Record 18;
        recResCenter: Record 5714;
        recCust1: Record 18;
        recCountry: Record 9;
        LocResCenter: Code[10];
        Title: Text[50];
        "AddlTax&Cess": Decimal;
        //recDetailedTaxEntry: Record 16522;
        "AddlTax&CessText": Text[30];
        ProdDisc: Decimal;
        MRPMaster: Record 50013;
        CashDiscText: Text[50];
        CashDiscountAmt: Decimal;
        cashDisc: Decimal;
        //recPosredStrOrdeLine: Record 13798;
        "---23Mar2107": Integer;
        NAH: Integer;
        RCheck: Report "Check Report";
}


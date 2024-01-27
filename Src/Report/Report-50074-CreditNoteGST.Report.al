report 50074 "Credit Note GST"
{
    // 
    // //17Apr2017 LineNo. Added in the TableBody & UserNamePreparedBy
    // Date        Version      Remarks
    // .....................................................................................
    // 14Sep2017   RB-N         Sales Comment Displayed
    // 05Feb2018   RB-N         Dispalying RespolLogo Or IpolLogo Depending upon ItemCategory
    // 03Apr2018   RB-N         GST Comp Cess
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/CreditNoteGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Credit Note GST';
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Sales Cr.Memo Header"; "Sales Cr.Memo Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(Logo21; Logo21)
            {
            }
            column(Cat21; Cat21)
            {
            }
            column(RepsolLogo; CompanyInfo1."Repsol Logo")
            {
            }
            column(LocState; LocState)
            {
            }
            column(LocStateCode; LocStateCode)
            {
            }
            column(CompName; CompanyInfo1.Name)
            {
            }
            column(CompPicture; CompanyInfo1.Picture)
            {
            }
            column(CompNamePicture; CompanyInfo1."Name Picture")
            {
            }
            column(CompShadedBoxPicture; CompanyInfo1."Shaded Box")
            {
            }
            column(DocNo; "Sales Cr.Memo Header"."No.")
            {
            }
            column(Loc_Location; recLoc.Address + ' ' + recLoc."Address 2" + ' ' + recLoc.City + ' ' + recLoc."Post Code" + ' ' + RecState1.Description + ' Phone: ' + recLoc."Phone No." + ' Email: ' + recLoc."E-Mail")
            {
            }
            column(LocGSTRegNo; recLoc."GST Registration No.")
            {
            }
            column(LocPANNo; recLoc."T.C.A.N. No.")
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
            column(Ship2Name; Ship2Name)
            {
            }
            column(ShipAdd1; ShipAdd1)
            {
            }
            column(ShipAdd2; ShipAdd2)
            {
            }
            column(ShipAdd3; ShipAdd3)
            {
            }
            column(ShipGSTNo; ShipGSTNo)
            {
            }
            column(ShipPANNo; ShipPANNo)
            {
            }
            column(ShipState; ShipState)
            {
            }
            column(ShipStateCode; ShipStateCode)
            {
            }
            column(BillGSTNo; BillGSTNo)
            {
            }
            column(BillPANNo; BillPANNo)
            {
            }
            column(BillState; BillState)
            {
            }
            column(BillStateCode; BillStateCode)
            {
            }
            column(CT1; CommText[1])
            {
            }
            column(CT2; CommText[2])
            {
            }
            column(CT3; CommText[3])
            {
            }
            column(RoundAmt; RoundAmt)
            {
            }
            column(TCSPercent; TCSPercent)
            {
            }
            column(TCSAmount; TCSAmount)
            {
            }
            column(QRCode; "Sales Cr.Memo Header"."QR code")
            {
            }
            column(IRN_SalesCrMemoHeader; IRN)
            {
            }
            dataitem("Sales Cr.Memo Line"; "Sales Cr.Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0),
                                          "No." = FILTER(<> 74012210));
                column(CessAmt; Cess)
                {
                }
                column(GSTBaseAmount; GSTBaseAmtLine)
                {
                }
                column(UnitofMeasure_SalesCrMemoLine; "Sales Cr.Memo Line"."Unit of Measure")
                {
                }
                column(UnitCost_SalesCrMemoLine; "Sales Cr.Memo Line"."Unit Cost")
                {
                }
                column(AmountToCustomer; AmountToCustomerLine)
                {
                }
                column(Amount_SalesCrMemoLine; AmountLine)
                {
                }
                column(LineDiscountAmount_SalesCrMemoLine; DiscAmtLine)
                {
                }
                column(GSTPer; 0)// "Sales Cr.Memo Line"."GST %")
                {
                }
                column(TotalGSTAmount; TotalGSTLine)
                {
                }
                column(HSNCode; "Sales Cr.Memo Line"."HSN/SAC Code")
                {
                }
                column(LinDocNo; "Document No.")
                {
                }
                column(LineNo; "Line No.")
                {
                }
                column(ItmNo; "No.")
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
                column(QtyperUnitofMeasure_SalesCrMemoLine; "Qty. per Unit of Measure")
                {
                }
                column(QuantityBase_SalesCrMemoLine; "Quantity (Base)")
                {
                }
                column(UnitPrice_SalesCrMemoLine; UnitPriceLine)
                {
                }
                column(LineAmount_SalesCrMemoLine; "Sales Cr.Memo Line"."Line Amount")
                {
                }
                column(EntryTax; EntryTax)
                {
                }
                column(TaxDesc; 0)// "Sales Cr.Memo Line"."Tax Amount" - "AddlTax&Cess")
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
                column(TotAmt; "Sales Cr.Memo Line"."Line Amount" + /*"Sales Cr.Memo Line"."Tax Amount*/0 + AddlDutyAmount + Cess + EntryTax)
                {
                }
                column(AmountinWords_1; AmountinWords[1])
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
                column(CGST15; CGST15)
                {
                }
                column(SGST15; SGST15)
                {
                }
                column(IGST15; IGST15)
                {
                }
                column(AmtinWords15; AmtinWords15[1])
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>20July2017
                    IF "Sales Cr.Memo Line"."No." = '' THEN
                        CurrReport.SKIP;

                    SrNo := SrNo + 1;


                    //<<20July2017


                    vCount := "Sales Cr.Memo Line".COUNT;
                    //>>1
                    SalesInvoiceDate := 0D;

                    SalesInvoiceHeader1.SETFILTER(SalesInvoiceHeader1."No.", "Sales Cr.Memo Header"."Applies-to Doc. No.");
                    IF SalesInvoiceHeader1.FIND('-') THEN
                        SalesInvoiceDate := SalesInvoiceHeader1."Posting Date";

                    IF "Sales Cr.Memo Header"."Applies-to Doc. No." = '' THEN
                        SalesInvoiceDate := 0D;

                    //For Additional Duty & Cess Amount
                    CLEAR("AddlTax&Cess");
                    /* recDetailedTaxEntry.RESET;
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                    recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Cr.Memo Line"."Tax Area Code");
                    IF recDetailedTaxEntry.FINDFIRST THEN
                        REPEAT
                            "AddlTax&Cess" := "AddlTax&Cess" + recDetailedTaxEntry."Tax Amount";
                        UNTIL recDetailedTaxEntry.NEXT = 0;

                    "AddlTax&CessText" := '';
                    recDetailedTaxEntry.RESET;
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Sales Cr.Memo Line"."Document No.");
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Sales Cr.Memo Line"."Tax Area Code");
                    recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Sales Cr.Memo Line"."Tax Area Code");
                    IF recDetailedTaxEntry.FINDFIRST THEN BEGIN
                        "AddlTax&CessText" := 'Addl. Tax/Cess @ ' + FORMAT(recDetailedTaxEntry."Tax %") + ' ' + '%';
                    END; */
                    //For Additional Duty & Cess Amount

                    //<<1


                    //>>2


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
                        vTotAmt += "Sales Cr.Memo Line"."Line Amount" + /*"Sales Cr.Memo Line"."Tax Amount"*/0 + AddlDutyAmount + Cess + EntryTax
                    ELSE
                        vTotAmt += "Sales Cr.Memo Line"."Line Amount" + /*"Sales Cr.Memo Line"."Tax Amount"*/0 + AddlDutyAmount + Cess + EntryTax / "Sales Cr.Memo Header"."Currency Factor";


                    //>>20July2017 GST
                    CLEAR(vCGST);
                    CLEAR(vSGST);
                    CLEAR(vIGST);
                    CLEAR(vCGSTRate);
                    CLEAR(vSGSTRate);
                    CLEAR(vIGSTRate);
                    CLEAR(CGST15);
                    CLEAR(SGST15);
                    CLEAR(IGST15);
                    DetailGSTEntry.RESET;
                    DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                    //DetailGSTEntry.SETRANGE(Type,Type);
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT

                            IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                /*//RSPLSUM 21Jul2020
                                IF "Sales Cr.Memo Header"."Currency Factor" = 0 THEN
                                BEGIN
                                 vCGST := ABS(DetailGSTEntry."GST Amount");
                                END;

                                IF "Sales Cr.Memo Header"."Currency Factor" <> 0 THEN
                                BEGIN
                                 vCGST := ABS(DetailGSTEntry."GST Amount" / "Sales Cr.Memo Header"."Currency Factor");
                                END;
                                *///RSPLSUM 21Jul2020
                                  //RSPLSUM 21Jul2020>>
                                vCGST := ABS(DetailGSTEntry."GST Amount");
                                IF PrintInFCY THEN
                                    vCGST := ABS(DetailGSTEntry."GST Amount" * "Sales Cr.Memo Header"."Currency Factor");
                                //RSPLSUM 21Jul2020<<

                                vCGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    CGST15 := ''
                                ELSE
                                    CGST15 := ' @ ' + FORMAT(vCGSTRate) + ' %';
                            END;

                            IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                /*//RSPLSUM 21Jul2020
                                IF "Sales Cr.Memo Header"."Currency Factor" = 0 THEN
                                BEGIN
                                 vSGST := ABS(DetailGSTEntry."GST Amount");
                                END;

                                IF "Sales Cr.Memo Header"."Currency Factor" <> 0 THEN
                                BEGIN
                                 vSGST := ABS(DetailGSTEntry."GST Amount" / "Sales Cr.Memo Header"."Currency Factor");
                                END;
                                *///RSPLSUM 21Jul2020
                                vSGST := ABS(DetailGSTEntry."GST Amount");//RSPLSUM 21Jul2020
                                                                          //RSPLSUM 21Jul2020>>
                                IF PrintInFCY THEN
                                    vSGST := ABS(DetailGSTEntry."GST Amount" * "Sales Cr.Memo Header"."Currency Factor");
                                //RSPLSUM 21Jul2020<<

                                vSGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    SGST15 := ''
                                ELSE
                                    SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                            END;

                            //>>20July2017 UTGST
                            IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                /*//RSPLSUM 21Jul2020
                                IF "Sales Cr.Memo Header"."Currency Factor" = 0 THEN
                                BEGIN
                                 vSGST := ABS(DetailGSTEntry."GST Amount");
                                END;

                                IF "Sales Cr.Memo Header"."Currency Factor" <> 0 THEN
                                BEGIN
                                 vSGST := ABS(DetailGSTEntry."GST Amount" / "Sales Cr.Memo Header"."Currency Factor");
                                END;
                                *///RSPLSUM 21Jul2020
                                vSGST := ABS(DetailGSTEntry."GST Amount");//RSPLSUM 21Jul2020
                                                                          //RSPLSUM 21Jul2020>>
                                IF PrintInFCY THEN
                                    vSGST := ABS(DetailGSTEntry."GST Amount" * "Sales Cr.Memo Header"."Currency Factor");
                                //RSPLSUM 21Jul2020<<

                                vSGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    SGST15 := ''
                                ELSE
                                    SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                            END;
                            //<<20July2017 UTGST

                            IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                /*//RSPLSUM 21Jul2020
                                IF "Sales Cr.Memo Header"."Currency Factor" = 0 THEN
                                BEGIN
                                 vIGST := ABS(DetailGSTEntry."GST Amount");
                                END;

                                IF "Sales Cr.Memo Header"."Currency Factor" <> 0 THEN
                                BEGIN
                                 vIGST := ABS(DetailGSTEntry."GST Amount" / "Sales Cr.Memo Header"."Currency Factor" );
                                END;
                                *///RSPLSUM 21Jul2020
                                vIGST := ABS(DetailGSTEntry."GST Amount");//RSPLSUM 21Jul2020
                                                                          //RSPLSUM 21Jul2020>>
                                IF PrintInFCY THEN
                                    vIGST := ABS(DetailGSTEntry."GST Amount" * "Sales Cr.Memo Header"."Currency Factor");
                                //RSPLSUM 21Jul2020<<

                                vIGSTRate := DetailGSTEntry."GST %";
                                //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    IGST15 := ''
                                ELSE
                                    IGST15 := ' @ ' + FORMAT(vIGSTRate) + ' %';

                            END;

                        UNTIL DetailGSTEntry.NEXT = 0;
                    //<<20July2017 GST

                    //>>03Apr2018
                    CLEAR(Cess);
                    /* PostedStrOrdrDetails1.RESET;
                    PostedStrOrdrDetails1.SETRANGE(Type, PostedStrOrdrDetails1.Type::Sale);
                    PostedStrOrdrDetails1.SETRANGE("Invoice No.", "Document No.");
                    PostedStrOrdrDetails1.SETRANGE("Line No.", "Line No.");
                    PostedStrOrdrDetails1.SETRANGE("Tax/Charge Type", PostedStrOrdrDetails1."Tax/Charge Type"::Charges);
                    PostedStrOrdrDetails1.SETFILTER("Tax/Charge Group", 'CESS');
                    IF PostedStrOrdrDetails1.FINDFIRST THEN BEGIN
                        PostedStrOrdrDetails1.CALCSUMS(Amount);
                        Cess := PostedStrOrdrDetails1.Amount;
                    END; */
                    //>>03Apr2018

                    CLEAR(UnitPriceLine);
                    CLEAR(AmountLine);
                    CLEAR(DiscAmtLine);
                    CLEAR(GSTBaseAmtLine);
                    CLEAR(AmountToCustomerLine);
                    CLEAR(TotalGSTLine);

                    IF "Sales Cr.Memo Header"."Currency Factor" = 0 THEN BEGIN

                        UnitPriceLine := "Sales Cr.Memo Line"."Unit Price";
                        //RSPLSUM 10Nov2020--AmountLine    := "Sales Cr.Memo Line".Amount;
                        AmountLine := "Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount" + 0;// "Sales Cr.Memo Line"."Charges To Customer";//RSPLSUM 10Nov2020
                        DiscAmtLine := "Sales Cr.Memo Line"."Line Discount Amount";
                        GSTBaseAmtLine := 0;//"Sales Cr.Memo Line"."GST Base Amount";
                        AmountToCustomerLine := 0;// "Sales Cr.Memo Line"."Amount To Customer";
                        TotalGSTLine := 0;//"Sales Cr.Memo Line"."Total GST Amount";

                    END;

                    IF "Sales Cr.Memo Header"."Currency Factor" <> 0 THEN BEGIN

                        UnitPriceLine := "Sales Cr.Memo Line"."Unit Price" / "Sales Cr.Memo Header"."Currency Factor";
                        //RSPLSUM 10Nov2020--AmountLine    := "Sales Cr.Memo Line".Amount / "Sales Cr.Memo Header"."Currency Factor";
                        AmountLine := ("Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount" + 0 /*"Sales Cr.Memo Line"."Charges To Customer"*/) / "Sales Cr.Memo Header"."Currency Factor";//RSPLSUM 10Nov2020
                        DiscAmtLine := "Sales Cr.Memo Line"."Line Discount Amount" / "Sales Cr.Memo Header"."Currency Factor";
                        GSTBaseAmtLine := 0/*"Sales Cr.Memo Line"."GST Base Amount"*/ / "Sales Cr.Memo Header"."Currency Factor";
                        AmountToCustomerLine := 0/* "Sales Cr.Memo Line"."Amount To Customer"*/ / "Sales Cr.Memo Header"."Currency Factor";
                        TotalGSTLine := 0/*"Sales Cr.Memo Line"."Total GST Amount"*/ / "Sales Cr.Memo Header"."Currency Factor";

                    END;

                    //RSPLSUM 21Jul2020>>
                    IF PrintInFCY THEN BEGIN
                        UnitPriceLine := "Sales Cr.Memo Line"."Unit Price";
                        //RSPLSUM 10Nov2020--AmountLine    := "Sales Cr.Memo Line".Amount;
                        AmountLine := "Sales Cr.Memo Line".Amount + "Sales Cr.Memo Line"."Line Discount Amount" + 0;//"Sales Cr.Memo Line"."Charges To Customer";//RSPLSUM 10Nov2020
                        DiscAmtLine := "Sales Cr.Memo Line"."Line Discount Amount";
                        GSTBaseAmtLine := 0;// "Sales Cr.Memo Line"."GST Base Amount";
                        AmountToCustomerLine := 0;// "Sales Cr.Memo Line"."Amount To Customer";
                        TotalGSTLine := 0;// "Sales Cr.Memo Line"."Total GST Amount";
                    END;
                    //RSPLSUM 21Jul2020<<

                end;

                trigger OnPreDataItem()
                begin

                    SrNo := 0;//20July2017
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

            trigger OnAfterGetRecord()
            begin

                //>>05Feb2018 Logo Baseon ItemCategory
                CLEAR(Cat21);
                CLEAR(Logo21);
                SL21.RESET;
                SL21.SETRANGE("Document No.", "No.");
                SL21.SETRANGE(Type, SL21.Type::Item);
                SL21.SETFILTER(Quantity, '<>%1', 0);
                IF SL21.FINDFIRST THEN BEGIN
                    Cat21 := SL21."Item Category Code";

                END;

                IF Cat21 = 'CAT15' THEN
                    Logo21 := 1
                ELSE
                    Logo21 := 3;
                //>>05Feb2018 Logo Baseon ItemCategory

                //>>1
                recLoc.GET("Sales Cr.Memo Header"."Location Code");

                //>>20July2017
                CLEAR(LocState);
                CLEAR(LocStateCode);

                St20.RESET;
                IF St20.GET(recLoc."State Code") THEN BEGIN
                    LocState := St20.Description;
                    LocStateCode := St20."State Code (GST Reg. No.)";

                END;

                //<<20July2017

                recCust1.GET("Sales Cr.Memo Header"."Bill-to Customer No.");

                recPSIH.RESET;
                recPSIH.SETRANGE(recPSIH."No.", "Applies-to Doc. No.");
                IF recPSIH.FIND('-') THEN
                    AppliesDocDate := recPSIH."Posting Date"
                ELSE
                    AppliesDocDate := 0D;

                //For Additional Duty Amount
                /* recPostedStrOrdDetails.RESET;
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "Sales Cr.Memo Header"."No.");
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code", 'ADDL TAX');
                IF recPostedStrOrdDetails.FINDFIRST THEN
                    ADDLtext := 'Addl Tax @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%';
                REPEAT
                    AddlDutyAmount := AddlDutyAmount + recPostedStrOrdDetails.Amount;
                UNTIL recPostedStrOrdDetails.NEXT = 0; */
                //For Additional Duty Amount



                /*
                IF Cess = 0 THEN
                   CessText := ''
                ELSE
                   //CessText :='Cess @'+FORMAT(PostedStrOrdrDetails1."Calculation Value")+' '+'%';
                CessText :='Cess';
                */
                //For Cess

                //For Entry Tax
                CLEAR(EntryTax);
                /* PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Sales Cr.Memo Header"."No.");
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Group", 'ENTRYTAX');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                    REPEAT
                        EntryTax := EntryTax + PostedStrOrdrLineDetails.Amount;
                    UNTIL PostedStrOrdrLineDetails.NEXT = 0; */

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
                                                    /*
                                                    //>>Robosoft/Migration/Rahul
                                                    CLEAR(vComment);
                                                    vSalesCommment.RESET;
                                                    vSalesCommment.SETRANGE("No.","No.");
                                                    IF vSalesCommment.FINDSET THEN REPEAT
                                                      vComment := vComment + ' '+vSalesCommment.Comment;
                                                    UNTIL vSalesCommment.NEXT=0;

                                                    IF UserSetup.GET(USERID) THEN;
                                                    //<<
                                                    */

                //>>20July2017

                CLEAR(Ship2Name);
                CLEAR(ShipAdd1);
                CLEAR(ShipAdd2);
                CLEAR(ShipAdd3);
                CLEAR(ShipGSTNo);
                CLEAR(ShipPANNo);
                CLEAR(ShipState);
                CLEAR(ShipStateCode);
                CLEAR(BillGSTNo);
                CLEAR(BillPANNo);
                CLEAR(BillState);
                CLEAR(BillStateCode);

                IF "Sales Cr.Memo Header"."Ship-to Code" <> '' THEN BEGIN

                    SH20.RESET;
                    IF SH20.GET("Sales Cr.Memo Header"."Bill-to Customer No.", "Sales Cr.Memo Header"."Ship-to Code") THEN BEGIN

                        Ship2Name := SH20.Name;
                        ShipAdd1 := SH20.Address;
                        ShipAdd3 := SH20."Address 2";
                        ShipAdd3 := SH20.City + ' - ' + SH20."Post Code";
                        ShipGSTNo := SH20."GST Registration No.";
                        ShipPANNo := SH20."P.A.N. No.";

                        St20.RESET;
                        IF St20.GET(SH20.State) THEN BEGIN

                            ShipState := St20.Description;
                            ShipStateCode := St20."State Code (GST Reg. No.)";
                        END;
                    END;
                END;

                BillGSTNo := recCust1."GST Registration No.";
                BillPANNo := recCust1."P.A.N. No.";

                //>>24May2019
                DGST04.RESET;
                DGST04.SETCURRENTKEY("Document No.");
                DGST04.SETRANGE("Document No.", "No.");
                IF DGST04.FINDFIRST THEN
                    BillGSTNo := DGST04."Buyer/Seller Reg. No.";
                //<<24May2019

                St20.RESET;
                IF St20.GET(recCust1."State Code") THEN BEGIN

                    BillState := St20.Description;
                    BillStateCode := St20."State Code (GST Reg. No.)";

                END;

                IF "Sales Cr.Memo Header"."Ship-to Code" = '' THEN BEGIN
                    Ship2Name := "Sales Cr.Memo Header"."Bill-to Name";
                    ShipAdd1 := "Sales Cr.Memo Header"."Bill-to Address";
                    ShipAdd2 := "Sales Cr.Memo Header"."Bill-to Address 2";
                    ShipAdd3 := "Sales Cr.Memo Header"."Bill-to City" + ' - ' + "Sales Cr.Memo Header"."Bill-to Post Code";
                    ShipGSTNo := BillGSTNo;
                    ShipPANNo := BillPANNo;
                    ShipState := BillState;
                    ShipStateCode := BillStateCode;

                END;

                //<<20July2017

                //RSPLSUM-TCS>>
                CLEAR(TCSAmount);
                SCL20.RESET;
                SCL20.SETCURRENTKEY("Document No.", Type, Quantity);
                SCL20.SETRANGE("Document No.", "No.");
                SCL20.SETFILTER(Quantity, '<>%1', 0);
                IF SCL20.FINDSET THEN
                    REPEAT
                        TCSAmount += 0;//SCL20."TDS/TCS Amount";
                    UNTIL SCL20.NEXT = 0;
                //RSPLSUM-TCS<<

                //<<20July2017 Amtinwords
                FinalAmt := 0; //20July2017
                SCL20.RESET;
                SCL20.SETRANGE("Document No.", "No.");
                IF SCL20.FINDSET THEN
                    REPEAT
                        //RSPLSUM 21Jul2020>>
                        IF PrintInFCY THEN BEGIN
                            FinalAmt += 0;// SCL20."Amount To Customer";
                        END ELSE BEGIN
                            //RSPLSUM 21Jul2020<<
                            IF "Sales Cr.Memo Header"."Currency Factor" = 0 THEN BEGIN
                                FinalAmt += 0;// SCL20."Amount To Customer";
                            END;

                            IF "Sales Cr.Memo Header"."Currency Factor" <> 0 THEN BEGIN
                                FinalAmt += 0;// SCL20."Amount To Customer" / "Sales Cr.Memo Header"."Currency Factor";
                            END;
                        END;//RSPLSUM 21Jul2020
                    UNTIL SCL20.NEXT = 0;



                vCheck.InitTextVariable;
                IF PrintInFCY THEN//RSPLSUM 21Jul2020
                    vCheck.FormatNoText(AmtinWords15, ROUND(FinalAmt, 0.01), "Sales Cr.Memo Header"."Currency Code")//RSPLSUM 21Jul2020
                ELSE//RSPLSUM 21Jul2020
                    vCheck.FormatNoText(AmtinWords15, ROUND(FinalAmt, 0.01), '');
                AmtinWords15[1] := DELCHR(AmtinWords15[1], '=', '*');
                //<<20July2017 Amtinwords

                //>>26Oct2017 Roundoff Amount
                CLEAR(RoundAmt);
                SCL26.RESET;
                SCL26.SETRANGE("Document No.", "No.");
                SCL26.SETRANGE(Type, SCL26.Type::"G/L Account");
                SCL26.SETRANGE("No.", '74012210');
                IF SCL26.FINDFIRST THEN BEGIN
                    RoundAmt := SCL26.Amount;
                END;
                //<<26Oct2017 Roundoff Amount

                //>>RB-N 14Sep2017 Comment
                CLEAR(CT);
                CLEAR(CommText);
                SalesComment.RESET;
                SalesComment.SETRANGE("Document Type", SalesComment."Document Type"::"Posted Credit Memo");
                SalesComment.SETRANGE("No.", "No.");
                IF SalesComment.FINDSET THEN
                    REPEAT
                        CT += 1;
                        CommText[CT] := SalesComment.Comment;

                    UNTIL SalesComment.NEXT = 0;

                //RSPLSUM-TCS>>
                CLEAR(TCSPercent);
                RecSCML.RESET;
                RecSCML.SETRANGE("Document No.", "No.");
                RecSCML.SETRANGE(Type, RecSCML.Type::Item);
                RecSCML.SETFILTER(Quantity, '<>%1', 0);
                IF RecSCML.FINDFIRST THEN BEGIN
                    // IF RecSCML."TDS/TCS %" <> 0 THEN
                    TCSPercent := '';//FORMAT(RecSCML."TDS/TCS %") + ' %';
                END;
                //RSPLSUM-TCS<<

                //RSPLSUM 31Oct2020 BEGIN>>
                CLEAR(IRN);
                "Sales Cr.Memo Header".CALCFIELDS("Sales Cr.Memo Header".IRN);
                IF "Sales Cr.Memo Header".IRN <> '' THEN
                    IRN := 'IRN: ' + "Sales Cr.Memo Header".IRN;
                //RSPLSUM 31Oct2020 END>>

                "Sales Cr.Memo Header".CALCFIELDS("Sales Cr.Memo Header"."QR code");//ROBOEinv

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
            area(content)
            {
                field("Print in FCY"; PrintInFCY)
                {
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

    trigger OnInitReport()
    begin
        //>>1
        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
        CompanyInfo1.CALCFIELDS("Repsol Logo");//05Feb2018
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
        // PostedStrOrdrDetails: Record 13798;
        FOCAMOUNT: Decimal;
        "Trans.Subsidy": Decimal;
        "Trade Discount": Decimal;
        "Cash Discount": Decimal;
        "Invoice Discount": Decimal;
        "InvoiceNo.": Code[20];
        recGLentry: Record 17;
        Cess: Decimal;
        FreightCharges: Decimal;
        LocationName: Text;
        reciTEMUOM: Record 5404;
        SrNo: Integer;
        recCust: Record 18;
        recResCenter: Record 5714;
        CustResCenter: Code[10];
        ResCenterName: Text;
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
        // PostedStrOrdrDetails1: Record 13798;
        CessText: Text[30];
        recDimVale: Record 349;
        //DetailedTaxEntry: Record 16522;
        FormCode: Code[20];
        TaxDescription: Text[50];
        "AddlTax&Cess": Decimal;
        //recDetailedTaxEntry: Record 16522;
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
        "------20July2017": Integer;
        Ship2Name: Text[100];
        ShipAdd1: Text[100];
        ShipAdd2: Text[100];
        ShipAdd3: Text[100];
        ShipGSTNo: Code[15];
        ShipState: Text[50];
        ShipStateCode: Code[5];
        BillGSTNo: Code[15];
        BillState: Text[50];
        BillStateCode: Code[5];
        SH20: Record 222;
        St20: Record State;
        "--------------GST": Integer;
        vCGST: Decimal;
        vCGSTRate: Decimal;
        vSGST: Decimal;
        vSGSTRate: Decimal;
        vIGST: Decimal;
        vIGSTRate: Decimal;
        DetailGSTEntry: Record "Detailed GST Ledger Entry";
        CGST15: Text[30];
        SGST15: Text[30];
        IGST15: Text[30];
        AmtinWords15: array[2] of Text[120];
        FinalAmt: Decimal;
        LocState: Text[50];
        LocStateCode: Code[5];
        SCL20: Record 115;
        UnitPriceLine: Decimal;
        AmountLine: Decimal;
        DiscAmtLine: Decimal;
        GSTBaseAmtLine: Decimal;
        AmountToCustomerLine: Decimal;
        TotalGSTLine: Decimal;
        "------14Sep2017": Integer;
        SalesComment: Record 44;
        CommText: array[10] of Text[150];
        CT: Integer;
        "------26Oct2017": Integer;
        SCL26: Record 115;
        RoundAmt: Decimal;
        "----------05Feb2018": Integer;
        SL21: Record 115;
        Cat21: Code[20];
        Logo21: Integer;
        DGST04: Record "Detailed GST Ledger Entry";
        PrintInFCY: Boolean;
        RecSCML: Record 115;
        TCSPercent: Text[10];
        TCSAmount: Decimal;
        IRN: Text[255];
        BillPANNo: Code[10];
        ShipPANNo: Code[10];
        LocPANNo: Code[10];
}


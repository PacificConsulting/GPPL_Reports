report 50066 "Purchase Debit Memo - GST"
{
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 07Oct2017   RB-N         New Report Development [CAS-18281-L2B7T6]
    // 25Oct2017   RB-N         GST Calulation for RCM Purchase Structure
    // 06Nov2017   RB-N         GST Amount In Words Display based on RCM Purchase Structure
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseDebitMemoGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Purchase Debit Memo - GST';

    dataset
    {
        dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(GSTText; GSTText)
            {
            }
            column(GSTAmtWords; GSTAmtWord06[1])
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
            column(DocNo; "No.")
            {
            }
            column(Loc_Location; recLoc.Address + ' ' + recLoc."Address 2" + ' ' + recLoc.City + ' ' + recLoc."Post Code" + ' ' + RecState1.Description + ' Phone: ' + recLoc."Phone No." + ' Email: ' + recLoc."E-Mail")
            {
            }
            column(LocGSTRegNo; recLoc."GST Registration No.")
            {
            }
            column(BilltoName; "Pay-to Name")
            {
            }
            column(BillAdd1; "Pay-to Address")
            {
            }
            column(BillAdd2; "Pay-to Address 2")
            {
            }
            column(BillAdd3; "Pay-to City" + ' - ' + "Pay-to Post Code")
            {
            }
            column(CustTINNo; 'Ven07."T.I.N. No."')
            {
            }
            column(CustResp; CustResCenter + ' - ' + ResCenterName)
            {
            }
            column(AppliestoDocNo; "Applies-to Doc. No.")
            {
            }
            column(RefNo; "Vendor Cr. Memo No.")
            {
            }
            column(PostingDate; FORMAT("Posting Date", 0, '<Day,2>-<Month,2>-<Year4>'))
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
            column(CurrencyCode_SalesCrMemoHeader; "Currency Code")
            {
            }
            column(CurrencyFactor_SalesCrMemoHeader; "Currency Factor")
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
            column(ShipState; ShipState)
            {
            }
            column(ShipStateCode; ShipStateCode)
            {
            }
            column(BillGSTNo; BillGSTNo)
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
            column(QRCode; "Purch. Cr. Memo Hdr."."QR Code")
            {
            }
            column(IRN_PurchCrMemoHeader; IRN)
            {
            }
            dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));
                column(GSTBaseAmount; GSTBaseAmtLine)
                {
                }
                column(UnitofMeasure_Line; "Unit of Measure")
                {
                }
                column(UnitCost_Line; "Unit Cost")
                {
                }
                column(AmountToCustomer; AmountToCustomerLine)
                {
                }
                column(Amount_Line; AmountLine)
                {
                }
                column(LineDiscountAmount_Line; DiscAmtLine)
                {
                }
                column(GSTPer; 0)// "GST %")
                {
                }
                column(TotalGSTAmount; TotalGSTLine)
                {
                }
                column(HSNCode; "HSN/SAC Code")
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
                column(Line_Desc; Description)
                {
                }
                column(DimValue; recDimVale.Name)
                {
                }
                column(Quantity_Line; Quantity)
                {
                }
                column(QtyperUnitofMeasure_Line; "Qty. per Unit of Measure")
                {
                }
                column(QuantityBase_Line; "Quantity (Base)")
                {
                }
                column(UnitPrice_Line; UnitPriceLine)
                {
                }
                column(LineAmount_Line; "Line Amount")
                {
                }
                column(EntryTax; EntryTax)
                {
                }
                column(TaxDesc; 0)// "Tax Amount" - "AddlTax&Cess")
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
                column(TotAmt; "Line Amount" + 0/*"Tax Amount"*/ + AddlDutyAmount + Cess + EntryTax)
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

                    //>>
                    IF "No." = '' THEN
                        CurrReport.SKIP;

                    SrNo := SrNo + 1;


                    //<<


                    vCount := COUNT;
                    //>>1




                    //For Additional Duty & Cess Amount
                    CLEAR("AddlTax&Cess");
                    /*  recDetailedTaxEntry.RESET;
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Document No.");
                     recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Tax Area Code");
                     recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Tax Area Code");
                     IF recDetailedTaxEntry.FINDFIRST THEN
                         REPEAT */
                    "AddlTax&Cess" := "AddlTax&Cess" + 0;//recDetailedTaxEntry."Tax Amount";
                                                         // UNTIL recDetailedTaxEntry.NEXT = 0;

                    "AddlTax&CessText" := '';
                    /* recDetailedTaxEntry.RESET;
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Document No.");
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Tax Area Code");
                    recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Tax Area Code");
                    IF recDetailedTaxEntry.FINDFIRST THEN BEGIN */
                    "AddlTax&CessText" := 'Addl. Tax/Cess @ ' + /*FORMAT(recDetailedTaxEntry."Tax %")*/'' + ' ' + '%';
                    //  END;
                    //For Additional Duty & Cess Amount
                    //<<1

                    iF ReturnReasons.GET("Return Reason Code") THEN
                        "Return Reason" := ReturnReasons.Description;

                    CLEAR(TaxDescription);
                    /* DetailedTaxEntry.RESET;
                    DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Document No.");
                    IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                        IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN */
                    FormCode := 'with Form ' + '';//etailedTaxEntry."Form Code";
                                                  // END;
                    TaxDescription := /*FORMAT(DetailedTaxEntry."Tax Type")*/'' + ' @ ' + /*FORMAT(DetailedTaxEntry."Tax %")*/'' + ' %' + FormCode;
                    //  END;


                    //>>GST
                    CLEAR(vCGST);
                    CLEAR(vSGST);
                    CLEAR(vIGST);
                    CLEAR(vCGSTRate);
                    CLEAR(vSGSTRate);
                    CLEAR(vIGSTRate);
                    CLEAR(CGST15);
                    CLEAR(SGST15);
                    CLEAR(IGST15);

                    // IF "Purch. Cr. Memo Hdr.".Structure <> 'GST0007' THEN //RB-N 25Oct2017 RCM Structure
                    BEGIN //RB-N 25Oct2017 RCM Structure
                        DetailGSTEntry.RESET;
                        DetailGSTEntry.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document No.", "Document No.");
                        DetailGSTEntry.SETRANGE(DetailGSTEntry."Document Line No.", "Line No.");
                        //DetailGSTEntry.SETRANGE(Type,Type);
                        //DetailGSTEntry.SETRANGE(DetailGSTEntry."No.","No.");
                        IF DetailGSTEntry.FINDSET THEN
                            REPEAT

                                IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                    IF "Purch. Cr. Memo Hdr."."Currency Factor" = 0 THEN BEGIN
                                        vCGST += ABS(DetailGSTEntry."GST Amount");
                                    END;

                                    IF "Purch. Cr. Memo Hdr."."Currency Factor" <> 0 THEN BEGIN
                                        vCGST += ABS(DetailGSTEntry."GST Amount" / "Purch. Cr. Memo Hdr."."Currency Factor");
                                    END;

                                    vCGSTRate := DetailGSTEntry."GST %";

                                    IF DetailGSTEntry."GST Amount" = 0 THEN
                                        CGST15 := ''
                                    ELSE
                                        CGST15 := ' @ ' + FORMAT(vCGSTRate) + ' %';
                                END;

                                IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                    IF "Purch. Cr. Memo Hdr."."Currency Factor" = 0 THEN BEGIN
                                        vSGST += ABS(DetailGSTEntry."GST Amount");
                                    END;

                                    IF "Purch. Cr. Memo Hdr."."Currency Factor" <> 0 THEN BEGIN
                                        vSGST += ABS(DetailGSTEntry."GST Amount" / "Purch. Cr. Memo Hdr."."Currency Factor");
                                    END;

                                    vSGSTRate := DetailGSTEntry."GST %";

                                    IF DetailGSTEntry."GST Amount" = 0 THEN
                                        SGST15 := ''
                                    ELSE
                                        SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                                END;

                                //>> UTGST
                                IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                    IF "Purch. Cr. Memo Hdr."."Currency Factor" = 0 THEN BEGIN
                                        vSGST += ABS(DetailGSTEntry."GST Amount");
                                    END;

                                    IF "Purch. Cr. Memo Hdr."."Currency Factor" <> 0 THEN BEGIN
                                        vSGST += ABS(DetailGSTEntry."GST Amount" / "Purch. Cr. Memo Hdr."."Currency Factor");
                                    END;

                                    vSGSTRate := DetailGSTEntry."GST %";

                                    IF DetailGSTEntry."GST Amount" = 0 THEN
                                        SGST15 := ''
                                    ELSE
                                        SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                                END;
                                //<< UTGST

                                IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                    IF "Purch. Cr. Memo Hdr."."Currency Factor" = 0 THEN BEGIN
                                        vIGST += ABS(DetailGSTEntry."GST Amount");
                                    END;

                                    IF "Purch. Cr. Memo Hdr."."Currency Factor" <> 0 THEN BEGIN
                                        vIGST += ABS(DetailGSTEntry."GST Amount" / "Purch. Cr. Memo Hdr."."Currency Factor");
                                    END;

                                    vIGSTRate := DetailGSTEntry."GST %";
                                    //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");

                                    IF DetailGSTEntry."GST Amount" = 0 THEN
                                        IGST15 := ''
                                    ELSE
                                        IGST15 := ' @ ' + FORMAT(vIGSTRate) + ' %';

                                END;

                            UNTIL DetailGSTEntry.NEXT = 0;
                    END;//>>RB-N 25Oct2017 RCM Structure
                    //<<GST

                    CLEAR(UnitPriceLine);
                    CLEAR(AmountLine);
                    CLEAR(DiscAmtLine);
                    CLEAR(GSTBaseAmtLine);
                    CLEAR(AmountToCustomerLine);
                    CLEAR(TotalGSTLine);

                    IF "Purch. Cr. Memo Hdr."."Currency Factor" = 0 THEN BEGIN

                        UnitPriceLine := "Purch. Cr. Memo Line"."Unit Cost";
                        AmountLine := Amount;
                        DiscAmtLine := "Line Discount Amount";
                        GSTBaseAmtLine := 0;//"GST Base Amount";
                        AmountToCustomerLine := 0;//"Amount To Vendor";
                        TotalGSTLine := 0;// "Total GST Amount";

                    END;

                    IF "Purch. Cr. Memo Hdr."."Currency Factor" <> 0 THEN BEGIN

                        UnitPriceLine := "Unit Cost" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        AmountLine := Amount / "Purch. Cr. Memo Hdr."."Currency Factor";
                        DiscAmtLine := "Line Discount Amount" / "Purch. Cr. Memo Hdr."."Currency Factor";
                        GSTBaseAmtLine := 0 /*"GST Base Amount"*// "Purch. Cr. Memo Hdr."."Currency Factor";
                        AmountToCustomerLine :=/* "Amount To Vendor"*/0 / "Purch. Cr. Memo Hdr."."Currency Factor";
                        TotalGSTLine := /*"Total GST Amount"*/0 / "Purch. Cr. Memo Hdr."."Currency Factor";

                    END;

                    //>>RB-N 25Oct2017 RCM Structure
                    // IF "Purch. Cr. Memo Hdr.".Structure = 'GST0007' THEN BEGIN
                    AmountToCustomerLine -= 0;//"Total GST Amount";
                    TotalGSTLine -= 0;// "Total GST Amount";
                                      // END;
                                      //<<RB-N 25Oct2017 RCM Structure
                end;
                // end;

                trigger OnPreDataItem()
                begin

                    SrNo := 0;//
                end;
            }
            dataitem("Purch. Comment Line"; "Purch. Comment Line")
            {
                DataItemLink = "No." = FIELD("No.");
                DataItemTableView = SORTING("Document Type", "No.", "Line No.");
                column(Comment; Comment)
                {
                }
            }

            trigger OnAfterGetRecord()
            begin

                //>>1
                recLoc.GET("Location Code");

                //>>
                CLEAR(LocState);
                CLEAR(LocStateCode);

                St20.RESET;
                IF St20.GET(recLoc."State Code") THEN BEGIN
                    LocState := St20.Description;
                    LocStateCode := St20."State Code (GST Reg. No.)";

                END;

                //<<

                Ven07.GET("Buy-from Vendor No.");

                PIH07.RESET;
                PIH07.SETRANGE("No.", "Applies-to Doc. No.");
                IF PIH07.FINDFIRST THEN
                    AppliesDocDate := PIH07."Posting Date"
                ELSE
                    AppliesDocDate := 0D;


                SalesInvoiceDate := AppliesDocDate;


                //For Additional Duty Amount
                /* recPostedStrOrdDetails.RESET;
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "No.");
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code", 'ADDL TAX');
                IF recPostedStrOrdDetails.FINDFIRST THEN */
                ADDLtext := 'Addl Tax @ ' + /*FORMAT(recPostedStrOrdDetails."Calculation Value")*/'' + ' ' + '%';
                // REPEAT
                AddlDutyAmount := AddlDutyAmount + 0;// recPostedStrOrdDetails.Amount;
                                                     // UNTIL recPostedStrOrdDetails.NEXT = 0;
                                                     //For Additional Duty Amount

                //For Cess
                /* PostedStrOrdrDetails1.RESET;
                PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Invoice No.", "No.");
                PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Tax/Charge Type", PostedStrOrdrDetails1."Tax/Charge Type"::"Other Taxes");
                PostedStrOrdrDetails1.SETFILTER(PostedStrOrdrDetails1."Tax/Charge Code", '%1|%2', 'Cess', 'C1');
                IF PostedStrOrdrDetails1.FINDFIRST THEN
                    REPEAT */
                Cess := Cess + 0;// PostedStrOrdrDetails1.Amount;
                                 // UNTIL PostedStrOrdrDetails1.NEXT = 0;

                IF Cess = 0 THEN
                    CessText := ''
                ELSE
                    //CessText :='Cess @'+FORMAT(PostedStrOrdrDetails1."Calculation Value")+' '+'%';
                    CessText := 'Cess';
                //For Cess

                //For Entry Tax
                CLEAR(EntryTax);
                /* PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "No.");
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Group", 'ENTRYTAX');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                    REPEAT */
                EntryTax := EntryTax + 0;// PostedStrOrdrLineDetails.Amount;
                                         // UNTIL PostedStrOrdrLineDetails.NEXT = 0;

                IF EntryTax = 0 THEN
                    EntryTaxDesc := ''
                ELSE
                    EntryTaxDesc := 'Entry Tax';
                //For Entry Tax
                //<<1


                //>>2
                Location.RESET;
                Location.SETFILTER(Location.Code, "Location Code");
                IF Location.FINDFIRST THEN
                    LocationName := Location.Name;
                LocResCenter := Location."Global Dimension 2 Code";
                LocTinNo := '';// Location."T.I.N. No.";

                CustResCenter := '';
                ResCenterName := '';
                recVen.RESET;
                recVen.SETRANGE("No.", "Buy-from Vendor No.");
                IF recVen.FINDFIRST THEN BEGIN
                    CustResCenter := recVen."Responsibility Center";
                    recResCenter.RESET;
                    recResCenter.SETRANGE(recResCenter.Code, CustResCenter);
                    IF recResCenter.FINDFIRST THEN
                        ResCenterName := recResCenter.Name;
                END;
                //<<2


                //>>3

                UserRec.RESET;
                UserRec.SETRANGE("User Name", "User ID");
                IF UserRec.FINDFIRST THEN
                    UserName := UserRec."Full Name";
                //<<3


                //>>

                CLEAR(Ship2Name);
                CLEAR(ShipAdd1);
                CLEAR(ShipAdd2);
                CLEAR(ShipAdd3);
                CLEAR(ShipGSTNo);
                CLEAR(ShipState);
                CLEAR(ShipStateCode);
                CLEAR(BillGSTNo);
                CLEAR(BillState);
                CLEAR(BillStateCode);

                IF "Order Address Code" <> '' THEN BEGIN

                    SH20.RESET;
                    IF SH20.GET("Buy-from Vendor No.", "Order Address Code") THEN BEGIN

                        Ship2Name := SH20.Name;
                        ShipAdd1 := SH20.Address;
                        ShipAdd3 := SH20."Address 2";
                        ShipAdd3 := SH20.City + ' - ' + SH20."Post Code";
                        ShipGSTNo := SH20."GST Registration No.";

                        St20.RESET;
                        IF St20.GET(SH20.State) THEN BEGIN

                            ShipState := St20.Description;
                            ShipStateCode := St20."State Code (GST Reg. No.)";
                        END;
                    END;
                END;

                BillGSTNo := Ven07."GST Registration No.";

                St20.RESET;
                IF St20.GET(Ven07."State Code") THEN BEGIN

                    BillState := St20.Description;
                    BillStateCode := St20."State Code (GST Reg. No.)";

                END;

                IF "Order Address Code" = '' THEN BEGIN
                    Ship2Name := "Purch. Cr. Memo Hdr."."Buy-from Vendor Name";
                    ShipAdd1 := "Purch. Cr. Memo Hdr."."Buy-from Address";
                    ShipAdd2 := "Purch. Cr. Memo Hdr."."Buy-from Address 2";
                    ShipAdd3 := "Purch. Cr. Memo Hdr."."Buy-from City" + ' - ' + "Purch. Cr. Memo Hdr."."Buy-from Post Code";
                    ShipGSTNo := BillGSTNo;
                    ShipState := BillState;
                    ShipStateCode := BillStateCode;

                END;

                //<<

                //<< Amtinwords
                FinalAmt := 0; //
                PCL07.RESET;
                PCL07.SETRANGE("Document No.", "No.");
                IF PCL07.FINDSET THEN
                    REPEAT
                        IF "Currency Factor" = 0 THEN BEGIN
                            FinalAmt += 0;// PCL07."Amount To Vendor";
                        END;

                        IF "Currency Factor" <> 0 THEN BEGIN
                            FinalAmt += 0;// PCL07."Amount To Vendor" / "Currency Factor";
                        END;
                    UNTIL PCL07.NEXT = 0;


                //>>RB-N 25Oct2017 RCM Structure
                //  IF "Purch. Cr. Memo Hdr.".Structure = 'GST0007' THEN BEGIN
                PCL07.RESET;
                PCL07.SETRANGE("Document No.", "No.");
                IF PCL07.FINDSET THEN
                    REPEAT
                        FinalAmt -= 0;// PCL07."Total GST Amount";

                    UNTIL PCL07.NEXT = 0;
                //  END;

                //>>RB-N 25Oct2017 RCM Structure


                vCheck.InitTextVariable;
                vCheck.FormatNoText(AmtinWords15, ROUND(FinalAmt, 0.01), '');
                AmtinWords15[1] := DELCHR(AmtinWords15[1], '=', '*');
                //<< Amtinwords


                //>>RB-N 06Nov2017 GSTAmountinWord
                CLEAR(GSTAmt06);
                CLEAR(GSTAmtWord06);
                CLEAR(GSTText);

                // IF "Purch. Cr. Memo Hdr.".Structure = 'GST0007' THEN BEGIN
                PCL07.RESET;
                PCL07.SETRANGE("Document No.", "No.");
                IF PCL07.FINDSET THEN
                    REPEAT

                        GSTAmt06 += 0;//PCL07."Total GST Amount";

                    UNTIL PCL07.NEXT = 0;
                //   END;

                IF GSTAmt06 <> 0 THEN BEGIN

                    GSTAmtWord06[1] := 'Rs. ' + FORMAT(GSTAmt06);

                    GSTText := 'Total GST Amount';
                END;

                IF GSTAmt06 = 0 THEN BEGIN

                    GSTAmtWord06[1] := '';
                    GSTText := '';
                END;
                //<<RB-N 06Nov2017 GSTAmountinWord

                //>>RB-N Comment
                CLEAR(CT);
                CLEAR(CommText);
                PCom07.RESET;
                PCom07.SETRANGE("Document Type", PCom07."Document Type"::"Posted Credit Memo");
                PCom07.SETRANGE("No.", "No.");
                IF PCom07.FINDSET THEN
                    REPEAT
                        CT += 1;
                        CommText[CT] := PCom07.Comment;

                    UNTIL PCom07.NEXT = 0;

                //RSPLSUM 01Feb21 BEGIN>>
                CLEAR(IRN);
                "Purch. Cr. Memo Hdr.".CALCFIELDS("Purch. Cr. Memo Hdr."."E-Inv Irn");
                IF "Purch. Cr. Memo Hdr."."E-Inv Irn" <> '' THEN
                    IRN := 'IRN: ' + "Purch. Cr. Memo Hdr."."E-Inv Irn";
                //RSPLSUM 01Feb21 END>>

                "Purch. Cr. Memo Hdr.".CALCFIELDS("Purch. Cr. Memo Hdr."."QR Code");//ROBOEinv //RSPLSUM 01Feb21
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
        UserSetup: Record "91";
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
        // PostedStrOrdrDetails: Record "13798";
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
        recLoc: Record "14";
        recCust1: Record "18";
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        AmountinWords: array[2] of Text[132];
        TotalAmount: Decimal;
        recPSIH: Record 112;
        AppliesDocDate: Date;
        // recPostedStrOrdDetails: Record "13798";
        ADDLtext: Text[30];
        //PostedStrOrdrDetails1: Record "13798";
        CessText: Text[30];
        recDimVale: Record 349;
        //  DetailedTaxEntry: Record "16522";
        FormCode: Code[20];
        TaxDescription: Text[50];
        "AddlTax&Cess": Decimal;
        //  recDetailedTaxEntry: Record "16522";
        "AddlTax&CessText": Text[30];
        EntryTaxDesc: Text[30];
        EntryTax: Decimal;
        //  PostedStrOrdrLineDetails: Record "13798";
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
        SH20: Record 224;
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
        SCL20: Record "115";
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
        "-------07Oct2017": Integer;
        Ven07: Record 23;
        PIH07: Record 122;
        recVen: Record 23;
        PCL07: Record 125;
        PCom07: Record 43;
        "---------06Nov2017": Integer;
        GSTAmt06: Decimal;
        GSTAmtWord06: array[2] of Text[120];
        GSTText: Text[50];
        IRN: Text[255];
}


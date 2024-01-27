report 50084 "Self Invoice GST"
{
    // 
    // //17Apr2017 LineNo. Added in the TableBody & UserNamePreparedBy
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/SelfInvoiceGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Self Invoice GST';

    dataset
    {
        dataitem("Purch. Inv. Header"; "Purch. Inv. Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
            column(PaymentTerms; Payterm27.Description)
            {
            }
            column(LocState; LocState)
            {
            }
            column(VendorInvoiceNo; "Vendor Invoice No.")
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
            column(PlaceofSupply; "Pay-to City")
            {
            }
            column(CustTINNo; 'recCust1."T.I.N. No."')
            {
            }
            column(CustResp; CustResCenter + ' - ' + ResCenterName)
            {
            }
            column(AppliestoDocNo; "Applies-to Doc. No.")
            {
            }
            column(PostingDate; FORMAT("Posting Date", 0, '<Day,2>-<Month,2>-<Year4>'))
            {
            }
            column(DocumentDate; FORMAT("Document Date", 0, '<Day,2>-<Month,2>-<Year4>'))
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
            column(CurrencyCode; "Currency Code")
            {
            }
            column(CurrencyFactor; "Currency Factor")
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
            dataitem("Purch. Inv. Line"; "Purch. Inv. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));
                column(GSTBaseAmount; GSTBaseAmtLine)
                {
                }
                column(UnitofMeasure; "Unit of Measure")
                {
                }
                column(UnitCost; "Unit Cost")
                {
                }
                column(AmountToCustomer; AmountToCustomerLine)
                {
                }
                column(AmountLine; AmountLine)
                {
                }
                column(LineDiscountAmount; DiscAmtLine)
                {
                }
                column(GSTPer; 0)//"GST %")
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
                column(SaleCrMemo_Desc; Description)
                {
                }
                column(DimValue; recDimVale.Name)
                {
                }
                column(Quantity; Quantity)
                {
                }
                column(UnitPrice; UnitPriceLine)
                {
                }
                column(LineAmount; "Line Amount")
                {
                }
                column(EntryTax; EntryTax)
                {
                }
                column(TaxDesc; 0)//"Tax Amount" - "AddlTax&Cess")
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
                column(TotAmt; "Line Amount" + /*"Tax Amount"*/0 + AddlDutyAmount + Cess + EntryTax)
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
                    IF "No." = '' THEN
                        CurrReport.SKIP;

                    SrNo := SrNo + 1;


                    //<<20July2017


                    vCount := COUNT;
                    //>>1
                    /*
                    SalesInvoiceDate:=0D;
                    
                    SalesInvoiceHeader1.SETFILTER(SalesInvoiceHeader1."No.","Sales Cr.Memo Header"."Applies-to Doc. No.");
                    IF SalesInvoiceHeader1.FIND('-') THEN
                      SalesInvoiceDate:=SalesInvoiceHeader1."Posting Date";
                    
                    IF "Sales Cr.Memo Header"."Applies-to Doc. No." ='' THEN
                    SalesInvoiceDate:=0D;
                    */
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


                    //For Additional Duty & Cess Amount
                    CLEAR("AddlTax&Cess");
                    /* recDetailedTaxEntry.RESET;
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Document No.");
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Tax Area Code");
                    recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Tax Area Code");
                    IF recDetailedTaxEntry.FINDFIRST THEN
                        REPEAT
                            "AddlTax&Cess" := "AddlTax&Cess" + recDetailedTaxEntry."Tax Amount";
                        UNTIL recDetailedTaxEntry.NEXT = 0;

                    "AddlTax&CessText" := '';
                    recDetailedTaxEntry.RESET;
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Document No.", "Document No.");
                    recDetailedTaxEntry.SETRANGE(recDetailedTaxEntry."Tax Area Code", "Tax Area Code");
                    recDetailedTaxEntry.SETFILTER(recDetailedTaxEntry."Tax Jurisdiction Code", '<>%1', "Tax Area Code");
                    IF recDetailedTaxEntry.FINDFIRST THEN BEGIN
                        "AddlTax&CessText" := 'Addl. Tax/Cess @ ' + FORMAT(recDetailedTaxEntry."Tax %") + ' ' + '%';
                    END; */
                    //For Additional Duty & Cess Amount

                    //<<1


                    //>>2


                    IF ReturnReasons.GET("Return Reason Code") THEN
                        "Return Reason" := ReturnReasons.Description;
                    //<<2

                    CLEAR(TaxDescription);
                    /* DetailedTaxEntry.RESET;
                    DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "No.");
                    IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                        IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN
                            FormCode := 'with Form ' + DetailedTaxEntry."Form Code";
                        END;
                        TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + ' @ ' + FORMAT(DetailedTaxEntry."Tax %") + ' %' + FormCode;
                    END; */

                    /*
                    IF "Currency Code" ='' THEN
                      vTotAmt+= "Sales Cr.Memo Line"."Line Amount"+"Sales Cr.Memo Line"."Tax Amount"+AddlDutyAmount+Cess+EntryTax
                    ELSE
                      vTotAmt+= "Sales Cr.Memo Line"."Line Amount"+"Sales Cr.Memo Line"."Tax Amount"+AddlDutyAmount+Cess+EntryTax/"Sales Cr.Memo Header"."Currency Factor";
                      */
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
                    /*
                    vCheck.InitTextVariable;
                    vCheck.FormatNoText(TotText,vTotAmt,"Sales Cr.Memo Header"."Currency Code");
                    TotText[1] :=  DELCHR(TotText[1],'=','*');
                    */

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
                    //DetailGSTEntry.SETRANGE(Type, Type);
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT

                            IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                IF "Purch. Inv. Header"."Currency Factor" = 0 THEN BEGIN
                                    vCGST := ABS(DetailGSTEntry."GST Amount");
                                END;

                                IF "Purch. Inv. Header"."Currency Factor" <> 0 THEN BEGIN
                                    vCGST := ABS(DetailGSTEntry."GST Amount" / "Purch. Inv. Header"."Currency Factor");
                                END;

                                vCGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    CGST15 := ''
                                ELSE
                                    CGST15 := ' @ ' + FORMAT(vCGSTRate) + ' %';
                            END;

                            IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                IF "Purch. Inv. Header"."Currency Factor" = 0 THEN BEGIN
                                    vSGST := ABS(DetailGSTEntry."GST Amount");
                                END;

                                IF "Purch. Inv. Header"."Currency Factor" <> 0 THEN BEGIN
                                    vSGST := ABS(DetailGSTEntry."GST Amount" / "Purch. Inv. Header"."Currency Factor");
                                END;

                                vSGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    SGST15 := ''
                                ELSE
                                    SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                            END;

                            //>>20July2017 UTGST
                            IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                IF "Purch. Inv. Header"."Currency Factor" = 0 THEN BEGIN
                                    vSGST := ABS(DetailGSTEntry."GST Amount");
                                END;

                                IF "Purch. Inv. Header"."Currency Factor" <> 0 THEN BEGIN
                                    vSGST := ABS(DetailGSTEntry."GST Amount" / "Purch. Inv. Header"."Currency Factor");
                                END;

                                vSGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    SGST15 := ''
                                ELSE
                                    SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                            END;
                            //<<20July2017 UTGST

                            IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                IF "Purch. Inv. Header"."Currency Factor" = 0 THEN BEGIN
                                    vIGST := ABS(DetailGSTEntry."GST Amount");
                                END;

                                IF "Purch. Inv. Header"."Currency Factor" <> 0 THEN BEGIN
                                    vIGST := ABS(DetailGSTEntry."GST Amount" / "Purch. Inv. Header"."Currency Factor");
                                END;

                                vIGSTRate := DetailGSTEntry."GST %";
                                //MESSAGE('%1 Doc No \\ %2 Item No',DetailGSTEntry."Document Line No.",DetailGSTEntry."GST %");

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    IGST15 := ''
                                ELSE
                                    IGST15 := ' @ ' + FORMAT(vIGSTRate) + ' %';

                            END;

                        UNTIL DetailGSTEntry.NEXT = 0;
                    //<<20July2017 GST

                    CLEAR(UnitPriceLine);
                    CLEAR(AmountLine);
                    CLEAR(DiscAmtLine);
                    CLEAR(GSTBaseAmtLine);
                    CLEAR(AmountToCustomerLine);
                    CLEAR(TotalGSTLine);

                    IF "Purch. Inv. Header"."Currency Factor" = 0 THEN BEGIN

                        UnitPriceLine := "Purch. Inv. Line"."Direct Unit Cost";
                        AmountLine := Amount;
                        DiscAmtLine := "Line Discount Amount";
                        GSTBaseAmtLine := 0;// "GST Base Amount";
                        AmountToCustomerLine := "Amount To Vendor";
                        TotalGSTLine := 0;// "Total GST Amount";

                    END;

                    IF "Purch. Inv. Header"."Currency Factor" <> 0 THEN BEGIN

                        UnitPriceLine := "Direct Unit Cost" / "Purch. Inv. Header"."Currency Factor";
                        AmountLine := Amount / "Purch. Inv. Header"."Currency Factor";
                        DiscAmtLine := "Line Discount Amount" / "Purch. Inv. Header"."Currency Factor";
                        GSTBaseAmtLine := 0;//"GST Base Amount" / "Purch. Inv. Header"."Currency Factor";
                        AmountToCustomerLine := "Amount To Vendor" / "Purch. Inv. Header"."Currency Factor";
                        TotalGSTLine := 0;//"Total GST Amount" / "Purch. Inv. Header"."Currency Factor";

                    END;

                end;

                trigger OnPreDataItem()
                begin

                    SrNo := 0;//27July2017
                end;
            }
            dataitem("Purch. Comment Line"; "Purch. Comment Line")
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

                //>>28July2017 Validation
                //Ven28.RESET;
                //IF Ven28.GET("Purch. Inv. Header"."Buy-from Vendor No.") THEN
                //BEGIN

                IF "GST Vendor Type" <> "GST Vendor Type"::Unregistered THEN
                    ERROR('This report is applicable for only Unregistered Vendor, \ Vendor No. :  %1 \ Type : %2', "Buy-from Vendor No.", "GST Vendor Type");
                //END;

                //<<28July2017 Validation

                //>>1
                recLoc.GET("Location Code");

                //>>27July2017
                CLEAR(LocState);
                CLEAR(LocStateCode);

                St20.RESET;
                IF St20.GET(recLoc."State Code") THEN BEGIN
                    LocState := St20.Description;
                    LocStateCode := St20."State Code (GST Reg. No.)";

                END;

                //<<27July2017
                /*
                recPSIH.RESET;
                recPSIH.SETRANGE(recPSIH."No.","Applies-to Doc. No.");
                IF recPSIH.FIND('-') THEN
                AppliesDocDate := recPSIH."Posting Date"
                ELSE
                AppliesDocDate := 0D;
                *///111

                //For Additional Duty Amount
                /*  recPostedStrOrdDetails.RESET;
                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Invoice No.", "No.");
                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Type", recPostedStrOrdDetails."Tax/Charge Type"::"Other Taxes");
                 recPostedStrOrdDetails.SETRANGE(recPostedStrOrdDetails."Tax/Charge Code", 'ADDL TAX');
                 IF recPostedStrOrdDetails.FINDFIRST THEN
                     ADDLtext := 'Addl Tax @ ' + FORMAT(recPostedStrOrdDetails."Calculation Value") + ' ' + '%';
                 REPEAT
                     AddlDutyAmount := AddlDutyAmount + recPostedStrOrdDetails.Amount;
                 UNTIL recPostedStrOrdDetails.NEXT = 0;
                 //For Additional Duty Amount

                 //For Cess
                 PostedStrOrdrDetails1.RESET;
                 PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Invoice No.", "No.");
                 PostedStrOrdrDetails1.SETRANGE(PostedStrOrdrDetails1."Tax/Charge Type", PostedStrOrdrDetails1."Tax/Charge Type"::"Other Taxes");
                 PostedStrOrdrDetails1.SETFILTER(PostedStrOrdrDetails1."Tax/Charge Code", '%1|%2', 'Cess', 'C1');
                 IF PostedStrOrdrDetails1.FINDFIRST THEN
                     REPEAT
                         Cess := Cess + PostedStrOrdrDetails1.Amount;
                     UNTIL PostedStrOrdrDetails1.NEXT = 0;

                 IF Cess = 0 THEN
                     CessText := ''
                 ELSE
                     //CessText :='Cess @'+FORMAT(PostedStrOrdrDetails1."Calculation Value")+' '+'%';
                     CessText := 'Cess';
                 //For Cess

                 //For Entry Tax
                 CLEAR(EntryTax);
                 PostedStrOrdrLineDetails.RESET;
                 PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "No.");
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

                /*
                //>>2
                //Sales Cr.Memo Header, Header (2) - OnPreSection()
                Location.RESET;
                Location.SETFILTER(Location.Code,"Location Code");
                IF Location.FINDFIRST THEN
                LocationName := Location.Name;
                LocResCenter := Location."Global Dimension 2 Code";
                LocTinNo := Location."T.I.N. No.";
                
                CustResCenter := '';
                ResCenterName := '';
                recCust.RESET;
                recCust.SETRANGE(recCust."No.","Sell-to Customer No.");
                IF recCust.FINDFIRST THEN
                BEGIN
                 CustResCenter := recCust."Responsibility Center";
                 recResCenter.RESET;
                 recResCenter.SETRANGE(recResCenter.Code,CustResCenter);
                 IF recResCenter.FINDFIRST THEN
                 ResCenterName := recResCenter.Name;
                END;
                //<<2
                */

                //>>3

                //Sales Cr.Memo Header, Footer (4) - OnPreSection()
                UserRec.RESET;
                //UserRec.SETRANGE(UserRec."User ID","Sales Cr.Memo Header"."User ID");//
                UserRec.SETRANGE("User Name", "User ID");//24Mar2017
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

                //>>27July2017

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

                    Ord27.RESET;
                    IF Ord27.GET("Buy-from Vendor No.", "Order Address Code") THEN BEGIN

                        Ship2Name := Ord27.Name;
                        ShipAdd1 := Ord27.Address;
                        ShipAdd3 := Ord27."Address 2";
                        ShipAdd3 := Ord27.City + ' - ' + Ord27."Post Code";
                        ShipGSTNo := Ord27."GST Registration No.";

                        St20.RESET;
                        IF St20.GET(Ord27.State) THEN BEGIN

                            ShipState := St20.Description;
                            ShipStateCode := St20."State Code (GST Reg. No.)";
                        END;
                    END;
                END;

                IF Ven27.GET("Buy-from Vendor No.") THEN;
                BillGSTNo := Ven27."GST Registration No.";
                St20.RESET;
                IF St20.GET(Ven27."State Code") THEN BEGIN

                    BillState := St20.Description;
                    BillStateCode := St20."State Code (GST Reg. No.)";

                END;

                IF "Order Address Code" = '' THEN BEGIN
                    Ship2Name := "Pay-to Name";
                    ShipAdd1 := "Pay-to Address";
                    ShipAdd2 := "Pay-to Address 2";
                    ShipAdd3 := "Pay-to City" + ' - ' + "Pay-to Post Code";
                    ShipGSTNo := BillGSTNo;
                    ShipState := BillState;
                    ShipStateCode := BillStateCode;

                END;

                IF "GST Vendor Type" = "GST Vendor Type"::Unregistered THEN
                    BillGSTNo := 'URD';
                //<<27July2017

                //<<27July2017 Amtinwords
                FinalAmt := 0; //27July2017
                PIL27.RESET;
                PIL27.SETRANGE("Document No.", "No.");
                IF PIL27.FINDSET THEN
                    REPEAT
                        IF "Currency Factor" = 0 THEN BEGIN
                            FinalAmt += PIL27."Amount To Vendor";
                        END;

                        IF "Currency Factor" <> 0 THEN BEGIN
                            FinalAmt += PIL27."Amount To Vendor" / "Currency Factor";
                        END;

                    UNTIL PIL27.NEXT = 0;



                vCheck.InitTextVariable;
                vCheck.FormatNoText(AmtinWords15, ROUND(FinalAmt, 0.01), '');
                AmtinWords15[1] := DELCHR(AmtinWords15[1], '=', '*');
                //<<27July2017 Amtinwords

                IF Payterm27.GET("Payment Terms Code") THEN; //27July2017

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
        RecState1: Record "state";
        RespCentre: Record 5714;
        SalesInvoiceHeader1: Record 112;
        SalesInvoiceDate: Date;
        QtyPackUOM: Decimal;
        ItemRec: Record 27;
        UserRec: Record 2000000120;
        UserName: Text[50];
        AddlDutyAmount: Decimal;
        //PostedStrOrdrDetails: Record "13798";
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
        //recPostedStrOrdDetails: Record "13798";
        ADDLtext: Text[30];
        // PostedStrOrdrDetails1: Record "13798";
        CessText: Text[30];
        recDimVale: Record 349;
        //DetailedTaxEntry: Record "16522";
        FormCode: Code[20];
        TaxDescription: Text[50];
        "AddlTax&Cess": Decimal;
        // recDetailedTaxEntry: Record "16522";
        "AddlTax&CessText": Text[30];
        EntryTaxDesc: Text[30];
        EntryTax: Decimal;
        //PostedStrOrdrLineDetails: Record "13798";
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
        SCL20: Record 113;
        UnitPriceLine: Decimal;
        AmountLine: Decimal;
        DiscAmtLine: Decimal;
        GSTBaseAmtLine: Decimal;
        AmountToCustomerLine: Decimal;
        TotalGSTLine: Decimal;
        "-----27July2017": Integer;
        Ven27: Record 23;
        Ord27: Record 224;
        PIL27: Record 123;
        Payterm27: Record 3;
        Ven28: Record 23;
}


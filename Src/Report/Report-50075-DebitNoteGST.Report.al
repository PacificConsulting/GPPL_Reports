report 50075 "Debit Note GST"
{
    // 
    // //17Apr2017 LineNo. Added in the TableBody & UserNamePreparedBy
    // 
    // Date        Version      Remarks
    // .....................................................................................
    // 18Dec2017   RB-N         Displaying Comments & Description
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/DebitNoteGST.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Debit Note GST';
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.";
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
            column(Ext_Doc_No; "Sales Invoice Header"."External Document No.")
            {
            }
            column(LocGSTRegNo; recLoc."GST Registration No.")
            {
            }
            column(LocPANNo; recLoc."T.C.A.N. No.")
            {
            }
            column(BilltoName; "Bill-to Name")
            {
            }
            column(BillAdd1; "Bill-to Address")
            {
            }
            column(BillAdd2; "Bill-to Address 2")
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
            column(AppliestoDocNo; "Applies-to Doc. No.")
            {
            }
            column(RefNo; "External Document No.")
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
            column(QRCode; "Sales Invoice Header"."QR code")
            {
            }
            column(IRN_SalesInvoiceHeader; IRN)
            {
            }
            column(TCSPercent; TCSPercent)
            {
            }
            column(TCSAmount; TCSAmount)
            {
            }
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));
                column(TDSTCSAmount; 0)//"TDS/TCS Amount")
                {
                }
                column(GSTBaseAmount; GSTBaseAmtLine)
                {
                }
                column(UnitofMeasure_SalesCrMemoLine; "Unit of Measure")
                {
                }
                column(UnitCost_SalesCrMemoLine; "Unit Cost")
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
                column(SaleCrMemo_Desc; Description)
                {
                }
                column(DimValue; recDimVale.Name)
                {
                }
                column(Quantity; Quantity)
                {
                }
                column(QtyperUnitofMeasure; "Sales Invoice Line"."Qty. per Unit of Measure")
                {
                }
                column(QuantityBase; "Sales Invoice Line"."Quantity (Base)")
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
                column(TotAmt; "Line Amount" + 0)// "Tax Amount" + AddlDutyAmount + Cess + EntryTax)
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
                    /*  recDetailedTaxEntry.RESET;
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

                    //Sales Cr.Memo Line, Body (2) - OnPreSection()
                    //CurrReport.SHOWOUTPUT("Sales Cr.Memo Header"."Currency Code" = '');

                    IF ReturnReasons.GET("Return Reason Code") THEN
                        "Return Reason" := ReturnReasons.Description;
                    //<<2

                    CLEAR(TaxDescription);
                    //IF "Sales Cr.Memo Header"."Currency Code" = '' THEN
                    //BEGIN
                    /* DetailedTaxEntry.RESET;
                    DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "No.");
                    IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                        IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN
                            FormCode := 'with Form ' + DetailedTaxEntry."Form Code";
                        END;
                        TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + ' @ ' + FORMAT(DetailedTaxEntry."Tax %") + ' %' + FormCode;
                    END; */
                    //END;

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
                    // DetailGSTEntry.SETRANGE(Type, Type);
                    DetailGSTEntry.SETRANGE(DetailGSTEntry."No.", "No.");
                    IF DetailGSTEntry.FINDSET THEN
                        REPEAT

                            IF DetailGSTEntry."GST Component Code" = 'CGST' THEN BEGIN
                                //IF "Sales Invoice Header"."Currency Factor" = 0 THEN
                                //BEGIN
                                vCGST := ABS(DetailGSTEntry."GST Amount");
                                //END;

                                //IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                //BEGIN
                                //vCGST := ABS(DetailGSTEntry."GST Amount" / "Sales Invoice Header"."Currency Factor");
                                //END;

                                vCGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    CGST15 := ''
                                ELSE
                                    CGST15 := ' @ ' + FORMAT(vCGSTRate) + ' %';
                            END;

                            IF DetailGSTEntry."GST Component Code" = 'SGST' THEN BEGIN
                                //IF "Sales Invoice Header"."Currency Factor" = 0 THEN
                                //BEGIN
                                vSGST := ABS(DetailGSTEntry."GST Amount");
                                //END;

                                //IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                //BEGIN
                                //vSGST := ABS(DetailGSTEntry."GST Amount" / "Sales Invoice Header"."Currency Factor");
                                //END;

                                vSGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    SGST15 := ''
                                ELSE
                                    SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                            END;

                            //>>20July2017 UTGST
                            IF DetailGSTEntry."GST Component Code" = 'UTGST' THEN BEGIN
                                //IF "Sales Invoice Header"."Currency Factor" = 0 THEN
                                //BEGIN
                                vSGST := ABS(DetailGSTEntry."GST Amount");
                                //END;

                                //IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                //BEGIN
                                //vSGST := ABS(DetailGSTEntry."GST Amount" / "Sales Invoice Header"."Currency Factor");
                                //END;

                                vSGSTRate := DetailGSTEntry."GST %";

                                IF DetailGSTEntry."GST Amount" = 0 THEN
                                    SGST15 := ''
                                ELSE
                                    SGST15 := ' @ ' + FORMAT(vSGSTRate) + ' %';

                            END;
                            //<<20July2017 UTGST

                            IF DetailGSTEntry."GST Component Code" = 'IGST' THEN BEGIN
                                //IF "Sales Invoice Header"."Currency Factor" = 0 THEN
                                //BEGIN
                                vIGST := ABS(DetailGSTEntry."GST Amount");
                                //END;

                                //IF "Sales Invoice Header"."Currency Factor" <> 0 THEN
                                //BEGIN
                                //vIGST := ABS(DetailGSTEntry."GST Amount" / "Sales Invoice Header"."Currency Factor" );
                                //END;

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

                    IF "Sales Invoice Header"."Currency Factor" = 0 THEN BEGIN

                        UnitPriceLine := "Unit Price";
                        AmountLine := Amount;
                        DiscAmtLine := "Line Discount Amount";
                        GSTBaseAmtLine := 0;// "GST Base Amount";
                        AmountToCustomerLine := 0;// "Amount To Customer";
                        TotalGSTLine := 0;// "Total GST Amount";

                    END;

                    IF "Sales Invoice Header"."Currency Factor" <> 0 THEN BEGIN

                        UnitPriceLine := "Unit Price" / "Sales Invoice Header"."Currency Factor";
                        AmountLine := Amount / "Sales Invoice Header"."Currency Factor";
                        DiscAmtLine := "Line Discount Amount" / "Sales Invoice Header"."Currency Factor";
                        GSTBaseAmtLine := 0;//"GST Base Amount" / "Sales Invoice Header"."Currency Factor";
                        AmountToCustomerLine := 0;//"Amount To Customer" / "Sales Invoice Header"."Currency Factor";
                        TotalGSTLine := 0;//"Total GST Amount" / "Sales Invoice Header"."Currency Factor";

                    END;

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

                //>>1
                recLoc.GET("Location Code");

                //>>20July2017
                CLEAR(LocState);
                CLEAR(LocStateCode);

                St20.RESET;
                IF St20.GET(recLoc."State Code") THEN BEGIN
                    LocState := St20.Description;
                    LocStateCode := St20."State Code (GST Reg. No.)";

                END;

                //<<20July2017

                recCust1.GET("Bill-to Customer No.");

                recPSIH.RESET;
                recPSIH.SETRANGE(recPSIH."No.", "Applies-to Doc. No.");
                IF recPSIH.FIND('-') THEN
                    AppliesDocDate := recPSIH."Posting Date"
                ELSE
                    AppliesDocDate := 0D;


                //For Additional Duty Amount
                /* recPostedStrOrdDetails.RESET;
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
                    UNTIL PostedStrOrdrDetails1.NEXT = 0; */

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
                Location.SETFILTER(Location.Code, "Location Code");
                IF Location.FINDFIRST THEN
                    LocationName := Location.Name;
                LocResCenter := Location."Global Dimension 2 Code";
                LocTinNo := '';//Location."T.I.N. No.";

                CustResCenter := '';
                ResCenterName := '';
                recCust.RESET;
                recCust.SETRANGE(recCust."No.", "Sell-to Customer No.");
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

                IF "Ship-to Code" <> '' THEN BEGIN

                    SH20.RESET;
                    IF SH20.GET("Bill-to Customer No.", "Ship-to Code") THEN BEGIN

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

                IF "Ship-to Code" = '' THEN BEGIN
                    Ship2Name := "Bill-to Name";
                    ShipAdd1 := "Bill-to Address";
                    ShipAdd2 := "Bill-to Address 2";
                    ShipAdd3 := "Bill-to City" + ' - ' + "Bill-to Post Code";
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
                        TCSAmount += 0;// SCL20."TDS/TCS Amount";
                    UNTIL SCL20.NEXT = 0;
                //RSPLSUM-TCS<<

                //<<20July2017 Amtinwords
                FinalAmt := 0; //20July2017
                SCL20.RESET;
                SCL20.SETRANGE("Document No.", "No.");
                IF SCL20.FINDSET THEN
                    REPEAT
                        IF "Currency Factor" = 0 THEN BEGIN
                            FinalAmt += 0;// SCL20."Amount To Customer";
                        END;

                        IF "Currency Factor" <> 0 THEN BEGIN
                            FinalAmt += 0;// SCL20."Amount To Customer" / "Currency Factor";
                        END;
                    UNTIL SCL20.NEXT = 0;



                vCheck.InitTextVariable;
                vCheck.FormatNoText(AmtinWords15, ROUND(FinalAmt, 0.01), '');
                AmtinWords15[1] := DELCHR(AmtinWords15[1], '=', '*');
                //<<20July2017 Amtinwords

                //>>RB-N Comment
                CLEAR(CT);
                CLEAR(CommText);
                SalesComment.RESET;
                SalesComment.SETRANGE("Document Type", SalesComment."Document Type"::"Posted Invoice");
                SalesComment.SETRANGE("No.", "No.");
                IF SalesComment.FINDSET THEN
                    REPEAT
                        CT += 1;
                        CommText[CT] := SalesComment.Comment;

                    UNTIL SalesComment.NEXT = 0;

                //RSPLSUM BEGIN>>
                CLEAR(IRN);
                "Sales Invoice Header".CALCFIELDS("Sales Invoice Header".IRN);
                IF "Sales Invoice Header".IRN <> '' THEN
                    IRN := 'IRN: ' + "Sales Invoice Header".IRN;
                //RSPLSUM END>>

                //RSPLSUM-TCS>>
                CLEAR(TCSPercent);
                RecSIL.RESET;
                RecSIL.SETRANGE("Document No.", "No.");
                RecSIL.SETRANGE(Type, RecSIL.Type::Item);
                RecSIL.SETFILTER(Quantity, '<>%1', 0);
                IF RecSIL.FINDFIRST THEN BEGIN
                    //IF RecSIL."TDS/TCS %" <> 0 THEN
                    TCSPercent := '';// FORMAT(RecSIL."TDS/TCS %") + ' %';
                                     //  END;
                                     //RSPLSUM-TCS<<

                    "Sales Invoice Header".CALCFIELDS("Sales Invoice Header"."QR code");//ROBOEinv
                                                                                        //BarcodeForQuarantineLabel;//RSPLSUM // DJ Commented 06/10/2020

                end;
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
        //recPostedStrOrdDetails: Record 13798;
        ADDLtext: Text[30];
        // PostedStrOrdrDetails1: Record 13798;
        CessText: Text[30];
        recDimVale: Record 349;
        //DetailedTaxEntry: Record 16522;
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
        "---------18Dec2017": Integer;
        SalesComment: Record 44;
        CommText: array[10] of Text[150];
        CT: Integer;
        DGST04: Record "Detailed GST Ledger Entry";
        RecGSTLedgEntry: Record "GST Ledger Entry";
        IRN: Text[255];
        TCSPercent: Text[10];
        RecSIL: Record 113;
        TCSAmount: Decimal;
        BillPANNo: Code[10];
        ShipPANNo: Code[10];
        LocPANNo: Code[10];

    //  [Scope('Internal')]
    procedure BarcodeForQuarantineLabel()
    var
        Text5000_Ctx: Label '<xpml><page quantity=''0'' pitch=''74.1 mm''></xpml>SIZE 100 mm, 74.1 mm';
        Text5001_Ctx: Label 'DIRECTION 0,0';
        Text5002_Ctx: Label 'REFERENCE 0,0';
        Text5003_Ctx: Label 'OFFSET 0 mm';
        Text5004_Ctx: Label 'SET PEEL OFF';
        Text5005_Ctx: Label 'SET CUTTER OFF';
        Text5006_Ctx: Label 'SET PARTIAL_CUTTER OFF';
        Text5007_Ctx: Label '<xpml></page></xpml><xpml><page quantity=''1'' pitch=''74.1 mm''></xpml>SET TEAR ON';
        Text5008_Ctx: Label 'CLS';
        Text5009_Ctx: Label 'CODEPAGE 1252';
        Text5010_Ctx: Label 'TEXT 806,792,"0",180,17,14,"SCHILLER_"';
        Text5011_Ctx: Label 'TEXT 1093,443,"0",180,14,12,"Item"';
        Text5012_Ctx: Label 'TEXT 1093,303,"0",180,10,12,"Part No."';
        Text5013_Ctx: Label 'TEXT 1093,166,"0",180,12,12,"Sr. No."';
        Text5014_Ctx: Label 'TEXT 933,443,"0",180,14,12,":"';
        Text5015_Ctx: Label 'TEXT 933,303,"0",180,14,12,":"';
        Text5016_Ctx: Label 'TEXT 933,166,"0",180,14,12,":"';
        Text5017_Ctx: Label 'TEXT 896,303,"0",180,10,12,"%1"';
        Text5018_Ctx: Label 'TEXT 896,166,"0",180,10,12,"%1"';
        Text5019_Ctx: Label 'QRCODE 330,300,L,10,A,180,M2,S7,"%1"';
        Text5020_Ctx: Label 'TEXT 896,443,"0",180,10,12,"%1"';
        Text5021_Ctx: Label 'PRINT 1,1';
        Text5022_Ctx: Label '<xpml></page></xpml><xpml><end/></xpml>';
        QRCodeInput: Text;
    // TempBlob: Record 99008535;
    begin
        RecGSTLedgEntry.RESET;
        RecGSTLedgEntry.SETRANGE("Document No.", "Sales Invoice Header"."No.");
        RecGSTLedgEntry.SETRANGE("Document Type", RecGSTLedgEntry."Document Type"::Invoice);
        RecGSTLedgEntry.SETRANGE("Transaction Type", RecGSTLedgEntry."Transaction Type"::Sales);
        IF RecGSTLedgEntry.FINDFIRST THEN BEGIN
            QRCodeInput := STRSUBSTNO('%1', RecGSTLedgEntry."Scan QrCode E-Invoice");
            IF QRCodeInput <> '' THEN BEGIN
                // CreateQRCode(QRCodeInput, TempBlob);
                // RecGSTLedgEntry."QR Code" := TempBlob.Blob;
                RecGSTLedgEntry.MODIFY;
            END;
        END;
    end;

    /* local procedure CreateQRCode(QRCodeInput: Text[500]; var TempBLOB: Record 99008535)
    var
        QRCodeFileName: Text[1024];
    begin
        CLEAR(TempBLOB);
        QRCodeFileName := GetQRCode(QRCodeInput);
        UploadFileBLOBImportandDeleteServerFile(TempBLOB, QRCodeFileName);
    end;
 */
    /*  local procedure GetQRCode(QRCodeInput: Text[300]) QRCodeFileName: Text[1024]
     var
         [RunOnClient]
         IBarCodeProvider: DotNet IBarcodeProvider;
     begin
         GetBarCodeProvider(IBarCodeProvider);
         QRCodeFileName := IBarCodeProvider.GetBarcode(QRCodeInput);
     end;

     // [Scope('Internal')]
     procedure UploadFileBLOBImportandDeleteServerFile(var TempBlob: Record 99008535; FileName: Text[1024])
     var
         FileManagement: Codeunit 419;
     begin
         FileName := FileManagement.UploadFileSilent(FileName);
         FileManagement.BLOBImportFromServerFile(TempBlob, FileName);
         DeleteServerFile(FileName);
     end;

     //  [Scope('Internal')]
     procedure GetBarCodeProvider(var IBarCodeProvider: DotNet IBarcodeProvider)
     var
         [RunOnClient]
         QRCodeProvider: DotNet QRCodeProvider;
     begin
         IF ISNULL(IBarCodeProvider) THEN
             IBarCodeProvider := QRCodeProvider.QRCodeProvider;
     end;

     local procedure DeleteServerFile(ServerFileName: Text)
     begin
         IF ERASE(ServerFileName) THEN;
     end; */
}


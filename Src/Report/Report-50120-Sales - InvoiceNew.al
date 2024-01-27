report 50120 "Sales - InvoiceNew"
{
    DefaultLayout = RDLC;

    Caption = 'Sales - Invoice';
    Permissions = TableData 7190 = rimd;
    RDLCLayout = 'Src/Report Layout/SalesInvoiceNew.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            RequestFilterFields = "No.", "Sell-to Customer No.", "No. Printed";
            RequestFilterHeading = 'Posted Sales Invoice';
            column(No_SalesInvoiceHeader; "No.")
            {
            }
            column(InvDiscountAmountCaption; InvDiscountAmountCaptionLbl)
            {
            }
            column(VATPercentageCaption; VATPercentageCaptionLbl)
            {
            }
            column(VATAmountCaption; VATAmountCaptionLbl)
            {
            }
            column(VATIdentifierCaption; VATIdentifierCaptionLbl)
            {
            }
            column(TotalCaption; TotalCaptionLbl)
            {
            }
            column(VATBaseCaption; VATBaseCaptionLbl)
            {
            }
            column(PaymentTermsCaption; PaymentTermsCaptionLbl)
            {
            }
            column(ShipmentMethodCaption; ShipmentMethodCaptionLbl)
            {
            }
            column(EMailCaption; EMailCaptionLbl)
            {
            }
            column(DocumentDateCaption; DocumentDateCaptionLbl)
            {
            }
            column(QRCODE; "Sales Invoice Header"."QR code")
            {
            }
            dataitem(CopyLoop; Integer)
            {
                DataItemTableView = SORTING(Number);
                dataitem(PageLoop; Integer)
                {
                    DataItemTableView = SORTING(Number)
                                        WHERE(Number = CONST(1));
                    column(CompanyInfo2Picture; CompanyInfo2.Picture)
                    {
                    }
                    column(CompanyInfo1Picture; CompanyInfo1.Picture)
                    {
                    }
                    column(CompanyInfo3Picture; CompanyInfo3.Picture)
                    {
                    }
                    column(DocumentCaptionCopyText; STRSUBSTNO(DocumentCaption, CopyText))
                    {
                    }
                    column(CompanyRegistrationLbl; CompanyRegistrationLbl)
                    {
                    }
                    column(CompanyInfo_GST_RegistrationNo; CompanyInfo."GST Registration No.")
                    {
                    }
                    column(CustomerRegistrationLbl; CustomerRegistrationLbl)
                    {
                    }
                    column(Customer_GST_RegistrationNo; Customer."GST Registration No.")
                    {
                    }
                    column(GSTComponentCode1; GSTComponentCode[1] + ' Amount')
                    {
                    }
                    column(GSTComponentCode2; GSTComponentCode[2] + ' Amount')
                    {
                    }
                    column(GSTComponentCode3; GSTComponentCode[3] + ' Amount')
                    {
                    }
                    column(GSTComponentCode4; GSTComponentCode[4] + 'Amount')
                    {
                    }
                    column(GSTCompAmount1; ABS(GSTCompAmount[1]))
                    {
                    }
                    column(GSTCompAmount2; ABS(GSTCompAmount[2]))
                    {
                    }
                    column(GSTCompAmount3; ABS(GSTCompAmount[3]))
                    {
                    }
                    column(GSTCompAmount4; ABS(GSTCompAmount[4]))
                    {
                    }
                    column(IsGSTApplicable; IsGSTApplicable)
                    {
                    }
                    column(CustAddr1; CustAddr[1])
                    {
                    }
                    column(CompanyAddr1; CompanyAddr[1])
                    {
                    }
                    column(CustAddr2; CustAddr[2])
                    {
                    }
                    column(CompanyAddr2; CompanyAddr[2])
                    {
                    }
                    column(CustAddr3; CustAddr[3])
                    {
                    }
                    column(CompanyAddr3; CompanyAddr[3])
                    {
                    }
                    column(CustAddr4; CustAddr[4])
                    {
                    }
                    column(CompanyAddr4; CompanyAddr[4])
                    {
                    }
                    column(CustAddr5; CustAddr[5])
                    {
                    }
                    column(CompanyInfoPhoneNo; CompanyInfo."Phone No.")
                    {
                    }
                    column(CustAddr6; CustAddr[6])
                    {
                    }
                    column(PaymentTermsDescription; PaymentTerms.Description)
                    {
                    }
                    column(ShipmentMethodDescription; ShipmentMethod.Description)
                    {
                    }
                    column(CompanyInfoHomePage; CompanyInfo."Home Page")
                    {
                    }
                    column(CompanyInfoEMail; CompanyInfo."E-Mail")
                    {
                    }
                    column(CompanyInfoVATRegNo; CompanyInfo."VAT Registration No.")
                    {
                    }
                    column(CompanyInfoGiroNo; CompanyInfo."Giro No.")
                    {
                    }
                    column(CompanyInfoBankName; CompanyInfo."Bank Name")
                    {
                    }
                    column(CompanyInfoBankAccountNo; CompanyInfo."Bank Account No.")
                    {
                    }
                    column(BillToCustNo_SalesInvHdr; "Sales Invoice Header"."Bill-to Customer No.")
                    {
                    }
                    column(PostingDate_SalesInvHdr; FORMAT("Sales Invoice Header"."Posting Date", 0, 4))
                    {
                    }
                    column(VATNoText; VATNoText)
                    {
                    }
                    column(VATRegNo_SalesInvHdr; "Sales Invoice Header"."VAT Registration No.")
                    {
                    }
                    column(DueDate_SalesInvoiceHdr; FORMAT("Sales Invoice Header"."Due Date", 0, 4))
                    {
                    }
                    column(SalesPersonText; SalesPersonText)
                    {
                    }
                    column(SalesPurchPersonName; SalesPurchPerson.Name)
                    {
                    }
                    column(ReferenceText; ReferenceText)
                    {
                    }
                    column(YourReference_SalesInvHdr; "Sales Invoice Header"."Your Reference")
                    {
                    }
                    column(OrderNoText; OrderNoText)
                    {
                    }
                    column(OrderNo_SalesInvoiceHdr; "Sales Invoice Header"."Order No.")
                    {
                    }
                    column(CustAddr7; CustAddr[7])
                    {
                    }
                    column(CustAddr8; CustAddr[8])
                    {
                    }
                    column(CompanyAddr5; CompanyAddr[5])
                    {
                    }
                    column(CompanyAddr6; CompanyAddr[6])
                    {
                    }
                    column(DocDate_SalesInvHeader; FORMAT("Sales Invoice Header"."Document Date", 0, 4))
                    {
                    }
                    column(PricesInclVAT_SalesInvHdr; "Sales Invoice Header"."Prices Including VAT")
                    {
                    }
                    column(OutputNo; OutputNo)
                    {
                    }
                    column(PricesInclVATYesNo; FORMAT("Sales Invoice Header"."Prices Including VAT"))
                    {
                    }
                    column(PageCaption; PageCaptionCap)
                    {
                    }
                    column(PLAEntryNo_SalesInvHdr; '"Sales Invoice Header"."PLA Entry No."')
                    {
                    }
                    column(SupplementaryText; SupplementaryText)
                    {
                    }
                    column(RG23AEntryNo_SalesInvHdr; '"Sales Invoice Header"."RG 23 A Entry No."')
                    {
                    }
                    column(RG23CEntryNo_SalesInvHdr; '"Sales Invoice Header"."RG 23 C Entry No."')
                    {
                    }
                    column(PhoneNoCaption; PhoneNoCaptionLbl)
                    {
                    }
                    column(HomePageCaption; HomePageCaptionCap)
                    {
                    }
                    column(VATRegNoCaption; VATRegNoCaptionLbl)
                    {
                    }
                    column(GiroNoCaption; GiroNoCaptionLbl)
                    {
                    }
                    column(BankNameCaption; BankNameCaptionLbl)
                    {
                    }
                    column(BankAccNoCaption; BankAccNoCaptionLbl)
                    {
                    }
                    column(DueDateCaption; DueDateCaptionLbl)
                    {
                    }
                    column(InvoiceNoCaption; InvoiceNoCaptionLbl)
                    {
                    }
                    column(PostingDateCaption; PostingDateCaptionLbl)
                    {
                    }
                    column(PLAEntryNoCaption; PLAEntryNoCaptionLbl)
                    {
                    }
                    column(RG23AEntryNoCaption; RG23AEntryNoCaptionLbl)
                    {
                    }
                    column(RG23CEntryNoCaption; RG23CEntryNoCaptionLbl)
                    {
                    }
                    column(ServiceTaxRegistrationNoCaption; ServiceTaxRegistrationNoLbl)
                    {
                    }
                    column(ServiceTaxRegistrationNo; ServiceTaxRegistrationNo)
                    {
                    }
                    column(BillToCustNo_SalesInvHdrCaption; "Sales Invoice Header".FIELDCAPTION("Bill-to Customer No."))
                    {
                    }
                    column(PricesInclVAT_SalesInvHdrCaption; "Sales Invoice Header".FIELDCAPTION("Prices Including VAT"))
                    {
                    }
                    dataitem(DimensionLoop1; Integer)
                    {
                        DataItemLinkReference = "Sales Invoice Header";
                        DataItemTableView = SORTING(Number)
                                            WHERE(Number = FILTER(1 ..));
                        column(DimText_DimensionLoop1; DimText)
                        {
                        }
                        column(Number_Integer; Number)
                        {
                        }
                        column(HeaderDimensionsCaption; HeaderDimensionsCaptionLbl)
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin
                            IF Number = 1 THEN BEGIN
                                IF NOT DimSetEntry1.FINDSET THEN
                                    CurrReport.BREAK;
                            END ELSE
                                IF NOT Continue THEN
                                    CurrReport.BREAK;

                            CLEAR(DimText);
                            Continue := FALSE;
                            REPEAT
                                OldDimText := DimText;
                                IF DimText = '' THEN
                                    DimText := STRSUBSTNO('%1 %2', DimSetEntry1."Dimension Code", DimSetEntry1."Dimension Value Code")
                                ELSE
                                    DimText :=
                                      STRSUBSTNO(
                                        '%1, %2 %3', DimText,
                                        DimSetEntry1."Dimension Code", DimSetEntry1."Dimension Value Code");
                                IF STRLEN(DimText) > MAXSTRLEN(OldDimText) THEN BEGIN
                                    DimText := OldDimText;
                                    Continue := TRUE;
                                    EXIT;
                                END;
                            UNTIL DimSetEntry1.NEXT = 0;
                        end;

                        trigger OnPreDataItem()
                        begin
                            IF NOT ShowInternalInfo THEN
                                CurrReport.BREAK;
                        end;
                    }
                    dataitem("Sales Invoice Line"; "Sales Invoice Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Sales Invoice Header";
                        DataItemTableView = SORTING("Document No.", "Line No.");
                        column(LineAmount_SalesInvLine; "Line Amount")
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(Description_SalesInvLine; Description)
                        {
                        }
                        column(No_SalesInvoiceLine; "No.")
                        {
                        }
                        column(Quantity_SalesInvoiceLine; Quantity)
                        {
                        }
                        column(UOM_SalesInvoiceLine; "Unit of Measure")
                        {
                        }
                        column(UnitPrice_SalesInvLine; "Unit Price")
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 2;
                        }
                        column(LineDiscount_SalesInvLine; "Line Discount %")
                        {
                        }
                        column(LineDiscount_SalesInvLineAmount; "Line Discount Amount")
                        {
                        }
                        column(PostedShipmentDate; FORMAT(PostedShipmentDate))
                        {
                        }
                        column(SalesLineType; FORMAT("Sales Invoice Line".Type))
                        {
                        }
                        column(DirectDebitPLARG_SalesInvLine; '"Direct Debit To PLA / RG"')
                        {
                        }
                        column(SourceDocNo_SalesInvLine; '"Source Document No."')
                        {
                        }
                        column(Supplementary; 'Supplementary')
                        {
                        }
                        column(InvDiscountAmount; -"Inv. Discount Amount")
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(TotalSubTotal; TotalSubTotal)
                        {
                            AutoFormatExpression = "Sales Invoice Header"."Currency Code";
                            AutoFormatType = 1;
                        }
                        column(TotalInvoiceDiscAmount; TotalInvoiceDiscountAmount)
                        {
                            AutoFormatExpression = "Sales Invoice Header"."Currency Code";
                            AutoFormatType = 1;
                        }
                        column(TotalText; TotalText)
                        {
                        }
                        column(SalesInvoiceLineAmount; Amount)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(TotalAmount; TotalAmount)
                        {
                            AutoFormatExpression = "Sales Invoice Header"."Currency Code";
                            AutoFormatType = 1;
                        }
                        column(AmtInclVAT_SalesInvLine; 0)// "Amount To Customer")
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(TotalInclVATText; TotalInclVATText)
                        {
                        }
                        column(TotalAmountInclVAT; TotalAmountInclVAT)
                        {
                            AutoFormatExpression = "Sales Invoice Header"."Currency Code";
                            AutoFormatType = 1;
                        }
                        column(TaxAmount_SalesInvLine; 0)//"Tax Amount")
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(ServiceTaxAmt; ServiceTaxAmt)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(ChargesAmount; ChargesAmount)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(OtherTaxesAmount; OtherTaxesAmount)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(ServiceTaxECessAmt; ServiceTaxECessAmt)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(SalesInvLineTotalTDSTCSInclSHECESS; TotalTCSAmount)
                        {
                        }
                        column(AppliedServiceTaxAmt; AppliedServiceTaxAmt)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(AppliedServiceTaxECessAmt; AppliedServiceTaxECessAmt)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(ServiceTaxSHECessAmt; ServiceTaxSHECessAmt)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(AppliedServTaxSHECessAmt; AppliedServiceTaxSHECessAmt)
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(TotalTaxAmt; TotalTaxAmt)
                        {
                        }
                        column(TotalExciseAmt; TotalExciseAmt)
                        {
                        }
                        column(VATBaseDisc_SalesInvHdr; "Sales Invoice Header"."VAT Base Discount %")
                        {
                            AutoFormatType = 1;
                        }
                        column(TotalPaymentDiscountOnVAT; TotalPaymentDiscountOnVAT)
                        {
                            AutoFormatType = 1;
                        }
                        column(VATAmtLineVATAmtText; VATAmountLine.VATAmountText)
                        {
                        }
                        column(TotalExclVATText; TotalExclVATText)
                        {
                        }
                        column(TotalAmountVAT; TotalAmountVAT)
                        {
                        }
                        column(SalesInvoiceLineLineNo; "Line No.")
                        {
                        }
                        column(UnitPriceCaption; UnitPriceCaptionLbl)
                        {
                        }
                        column(DiscountCaption; DiscountCaptionLbl)
                        {
                        }
                        column(AmountCaption; AmountCaptionLbl)
                        {
                        }
                        column(LineDiscountCaption; LineDiscountCaptionLbl)
                        {
                        }
                        column(PostedShipmentDateCaption; PostedShipmentDateCaptionLbl)
                        {
                        }
                        column(SubtotalCaption; SubtotalCaptionLbl)
                        {
                        }
                        column(ExciseAmountCaption; ExciseAmountCaptionLbl)
                        {
                        }
                        column(TaxAmountCaption; TaxAmountCaptionLbl)
                        {
                        }
                        column(ServiceTaxAmountCaption; ServiceTaxAmountCaptionLbl)
                        {
                        }
                        column(ChargesAmountCaption; ChargesAmountCaptionLbl)
                        {
                        }
                        column(OtherTaxesAmountCaption; OtherTaxesAmountCaptionLbl)
                        {
                        }
                        column(ServTaxeCessAmtCaption; ServTaxeCessAmtCaptionLbl)
                        {
                        }
                        column(TCSAmountCaption; TCSAmountCaptionLbl)
                        {
                        }
                        column(SvcTaxAmtAppliedCaption; SvcTaxAmtAppliedCaptionLbl)
                        {
                        }
                        column(SvcTaxeCessAmtAppliedCaption; SvcTaxeCessAmtAppliedCaptionLbl)
                        {
                        }
                        column(ServTaxSHECessAmtCaption; ServTaxSHECessAmtCaptionLbl)
                        {
                        }
                        column(SvcTaxSHECessAmtAppliedCaption; SvcTaxSHECessAmtAppliedCaptionLbl)
                        {
                        }
                        column(PaymentDiscVATCaption; PaymentDiscVATCaptionLbl)
                        {
                        }
                        column(Description_SalesInvLineCaption; FIELDCAPTION(Description))
                        {
                        }
                        column(No_SalesInvoiceLineCaption; FIELDCAPTION("No."))
                        {
                        }
                        column(Quantity_SalesInvoiceLineCaption; FIELDCAPTION(Quantity))
                        {
                        }
                        column(UOM_SalesInvoiceLineCaption; FIELDCAPTION("Unit of Measure"))
                        {
                        }
                        /* column(DirectDebitPLARG_SalesInvLineCaption; FIELDCAPTION("Direct Debit To PLA / RG"))
                        {
                        } */
                        column(ServiceTaxSBCAmt; ServiceTaxSBCAmt)
                        {
                        }
                        column(AppliedServiceTaxSBCAmt; AppliedServiceTaxSBCAmt)
                        {
                        }
                        column(ServTaxSBCAmtCaption; ServTaxSBCAmtCaptionLbl)
                        {
                        }
                        column(SvcTaxSBCAmtAppliedCaption; SvcTaxSBCAmtAppliedCaptionLbl)
                        {
                        }
                        column(KKCessAmt; KKCessAmt)
                        {
                        }
                        column(AppliedKKCessAmt; AppliedKKCessAmt)
                        {
                        }
                        column(KKCessAmtCaption; KKCessAmtCaptionLbl)
                        {
                        }
                        column(KKCessAmtAppliedCaption; KKCessAmtAppliedCaptionLbl)
                        {
                        }
                        dataitem("Sales Shipment Buffer"; Integer)
                        {
                            DataItemTableView = SORTING(Number);
                            column(SalesShpBufferPostingDate; FORMAT(SalesShipmentBuffer."Posting Date"))
                            {
                            }
                            column(SalesShipmentBufferQty; SalesShipmentBuffer.Quantity)
                            {
                                DecimalPlaces = 0 : 5;
                            }
                            column(ShipmentCaption; ShipmentCaptionLbl)
                            {
                            }

                            trigger OnAfterGetRecord()
                            begin
                                IF Number = 1 THEN
                                    SalesShipmentBuffer.FIND('-')
                                ELSE
                                    SalesShipmentBuffer.NEXT;
                            end;

                            trigger OnPreDataItem()
                            begin
                                SalesShipmentBuffer.SETRANGE("Document No.", "Sales Invoice Line"."Document No.");
                                SalesShipmentBuffer.SETRANGE("Line No.", "Sales Invoice Line"."Line No.");

                                SETRANGE(Number, 1, SalesShipmentBuffer.COUNT);
                            end;
                        }
                        dataitem(DimensionLoop2; Integer)
                        {
                            DataItemTableView = SORTING(Number)
                                                WHERE(Number = FILTER(1 ..));
                            column(DimText_DimensionLoop2; DimText)
                            {
                            }
                            column(LineDimensionsCaption; LineDimensionsCaptionLbl)
                            {
                            }

                            trigger OnAfterGetRecord()
                            begin
                                IF Number = 1 THEN BEGIN
                                    IF NOT DimSetEntry2.FINDSET THEN
                                        CurrReport.BREAK;
                                END ELSE
                                    IF NOT Continue THEN
                                        CurrReport.BREAK;

                                CLEAR(DimText);
                                Continue := FALSE;
                                REPEAT
                                    OldDimText := DimText;
                                    IF DimText = '' THEN
                                        DimText := STRSUBSTNO('%1 %2', DimSetEntry2."Dimension Code", DimSetEntry2."Dimension Value Code")
                                    ELSE
                                        DimText :=
                                          STRSUBSTNO(
                                            '%1, %2 %3', DimText,
                                            DimSetEntry2."Dimension Code", DimSetEntry2."Dimension Value Code");
                                    IF STRLEN(DimText) > MAXSTRLEN(OldDimText) THEN BEGIN
                                        DimText := OldDimText;
                                        Continue := TRUE;
                                        EXIT;
                                    END;
                                UNTIL DimSetEntry2.NEXT = 0;
                            end;

                            trigger OnPreDataItem()
                            begin
                                IF NOT ShowInternalInfo THEN
                                    CurrReport.BREAK;

                                DimSetEntry2.SETRANGE("Dimension Set ID", "Sales Invoice Line"."Dimension Set ID");
                            end;
                        }
                        dataitem(AsmLoop; Integer)
                        {
                            DataItemTableView = SORTING(Number);
                            column(TempPostedAsmLineNo; BlanksForIndent + TempPostedAsmLine."No.")
                            {
                            }
                            column(TempPostedAsmLineDescription; BlanksForIndent + TempPostedAsmLine.Description)
                            {
                            }
                            column(TempPostedAsmLineVariantCode; BlanksForIndent + TempPostedAsmLine."Variant Code")
                            {
                                //DecimalPlaces = 0 : 5;
                            }
                            column(TempPostedAsmLineQuantity; TempPostedAsmLine.Quantity)
                            {
                                DecimalPlaces = 0 : 5;
                            }
                            column(TempPostedAsmLineUOMCode; GetUOMText(TempPostedAsmLine."Unit of Measure Code"))
                            {
                                //DecimalPlaces = 0 : 5;


                            }

                            trigger OnAfterGetRecord()
                            begin
                                IF Number = 1 THEN
                                    TempPostedAsmLine.FINDSET
                                ELSE
                                    TempPostedAsmLine.NEXT;
                            end;

                            trigger OnPreDataItem()
                            begin
                                CLEAR(TempPostedAsmLine);
                                IF NOT DisplayAssemblyInformation THEN
                                    CurrReport.BREAK;
                                CollectAsmInformation;
                                CLEAR(TempPostedAsmLine);
                                SETRANGE(Number, 1, TempPostedAsmLine.COUNT);
                            end;
                        }

                        trigger OnAfterGetRecord()
                        var
                        //StructureLineDetails: Record 13798;
                        begin
                            PostedShipmentDate := 0D;
                            IF Quantity <> 0 THEN
                                PostedShipmentDate := FindPostedShipmentDate;

                            IF (Type = Type::"G/L Account") AND (NOT ShowInternalInfo) THEN
                                "No." := '';

                            VATAmountLine.INIT;
                            VATAmountLine."VAT Identifier" := "VAT Identifier";
                            VATAmountLine."VAT Calculation Type" := "VAT Calculation Type";
                            VATAmountLine."Tax Group Code" := "Tax Group Code";
                            VATAmountLine."VAT %" := "VAT %";
                            VATAmountLine."VAT Base" := Amount;
                            VATAmountLine."Amount Including VAT" := "Amount Including VAT";
                            VATAmountLine."Line Amount" := "Line Amount";
                            IF "Allow Invoice Disc." THEN
                                VATAmountLine."Inv. Disc. Base Amount" := "Line Amount";
                            VATAmountLine."Invoice Discount Amount" := "Inv. Discount Amount";
                            VATAmountLine.InsertLine;

                            TotalTCSAmount += 0;//"Total TDS/TCS Incl. SHE CESS";

                            TotalSubTotal += "Line Amount";
                            TotalInvoiceDiscountAmount -= "Inv. Discount Amount";
                            TotalAmount += Amount;
                            TotalAmountVAT += "Amount Including VAT" - Amount;
                            // TotalAmountInclVAT += "Amount Including VAT";
                            TotalPaymentDiscountOnVAT += -("Line Amount" - "Inv. Discount Amount" - "Amount Including VAT");
                            TotalAmountInclVAT += 0;//"Amount To Customer";
                            TotalExciseAmt += 0;// "Excise Amount";
                            TotalTaxAmt += 0;// "Tax Amount";
                            ServiceTaxAmount += 0;//"Service Tax Amount";
                            ServiceTaxeCessAmount += 0;//"Service Tax eCess Amount";
                            ServiceTaxSHECessAmount += 0;// "Service Tax SHE Cess Amount";
                            ServiceTaxSBCAmount += 0;//"Service Tax SBC Amount";
                            KKCessAmount += 0;// "KK Cess Amount";

                            IF IsGSTApplicable AND (Type <> Type::" ") THEN BEGIN
                                J := 1;
                                //GSTComponent.RESET;
                                //GSTComponent.SETRANGE("GST Jurisdiction Type", "GST Jurisdiction Type");
                                //IF GSTComponent.FINDSET THEN
                                // REPEAT
                                GSTComponentCode[J] := '';//GSTComponent.Code;
                                DetailedGSTLedgerEntry.RESET;
                                DetailedGSTLedgerEntry.SETCURRENTKEY("Transaction Type", "Document Type", "Document No.", "Document Line No.");
                                DetailedGSTLedgerEntry.SETRANGE("Transaction Type", DetailedGSTLedgerEntry."Transaction Type"::Sales);
                                DetailedGSTLedgerEntry.SETRANGE("Document No.", "Document No.");
                                DetailedGSTLedgerEntry.SETRANGE("Document Line No.", "Line No.");
                                DetailedGSTLedgerEntry.SETRANGE("GST Component Code", GSTComponentCode[J]);
                                IF DetailedGSTLedgerEntry.FINDSET THEN BEGIN
                                    REPEAT
                                        GSTCompAmount[J] += DetailedGSTLedgerEntry."GST Amount";
                                    UNTIL DetailedGSTLedgerEntry.NEXT = 0;
                                    J += 1;
                                END;
                                //UNTIL GSTComponent.NEXT = 0;
                                //END;

                                /* StructureLineDetails.RESET;
                                StructureLineDetails.SETRANGE(Type, StructureLineDetails.Type::Sale);
                                StructureLineDetails.SETRANGE("Document Type", StructureLineDetails."Document Type"::Invoice);
                                StructureLineDetails.SETRANGE("Invoice No.", "Document No.");
                                StructureLineDetails.SETRANGE("Item No.", "No.");
                                StructureLineDetails.SETRANGE("Line No.", "Line No.");
                                IF StructureLineDetails.FIND('-') THEN
                                    REPEAT
                                        IF NOT StructureLineDetails."Payable to Third Party" THEN BEGIN
                                            IF StructureLineDetails."Tax/Charge Type" = StructureLineDetails."Tax/Charge Type"::Charges THEN
                                                ChargesAmount := ChargesAmount + ABS(StructureLineDetails.Amount);
                                            IF StructureLineDetails."Tax/Charge Type" = StructureLineDetails."Tax/Charge Type"::"Other Taxes" THEN
                                                OtherTaxesAmount := OtherTaxesAmount + ABS(StructureLineDetails.Amount);
                                        END;
                                    UNTIL StructureLineDetails.NEXT = 0;
                                IF "Sales Invoice Header"."Transaction No. Serv. Tax" <> 0 THEN BEGIN
                                    ServiceTaxEntry.RESET;
                                    ServiceTaxEntry.SETRANGE(Type, ServiceTaxEntry.Type::Sale);
                                    ServiceTaxEntry.SETRANGE("Document Type", ServiceTaxEntry."Document Type"::Invoice);
                                    ServiceTaxEntry.SETRANGE("Document No.", "Document No.");
                                    IF ServiceTaxEntry.FINDFIRST THEN BEGIN

                                        IF "Sales Invoice Header"."Currency Code" <> '' THEN BEGIN
                                            ServiceTaxEntry."Service Tax Amount" :=
                                              ROUND(CurrExchRate.ExchangeAmtLCYToFCY(
                                              "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                              ServiceTaxEntry."Service Tax Amount", "Sales Invoice Header"."Currency Factor"));
                                            ServiceTaxEntry."eCess Amount" :=
                                              ROUND(CurrExchRate.ExchangeAmtLCYToFCY(
                                              "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                              ServiceTaxEntry."eCess Amount", "Sales Invoice Header"."Currency Factor"));
                                            ServiceTaxEntry."SHE Cess Amount" :=
                                              ROUND(CurrExchRate.ExchangeAmtLCYToFCY(
                                              "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                              ServiceTaxEntry."SHE Cess Amount", "Sales Invoice Header"."Currency Factor"));
                                            ServiceTaxEntry."Service Tax SBC Amount" :=
                                              ROUND(CurrExchRate.ExchangeAmtLCYToFCY(
                                              "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                              ServiceTaxEntry."Service Tax SBC Amount", "Sales Invoice Header"."Currency Factor"));
                                            ServiceTaxEntry."KK Cess Amount" :=
                                              ROUND(CurrExchRate.ExchangeAmtLCYToFCY(
                                              "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                              ServiceTaxEntry."KK Cess Amount", "Sales Invoice Header"."Currency Factor"));
                                        END; */

                                ServiceTaxAmt := 0;// ABS(ServiceTaxEntry."Service Tax Amount");
                                ServiceTaxECessAmt := 0;//ABS(ServiceTaxEntry."eCess Amount");
                                ServiceTaxSHECessAmt := 0;//ABS(ServiceTaxEntry."SHE Cess Amount");
                                ServiceTaxSBCAmt := 0;// ABS(ServiceTaxEntry."Service Tax SBC Amount");
                                KKCessAmt := 0;// ABS(ServiceTaxEntry."KK Cess Amount");
                                AppliedServiceTaxAmt := ServiceTaxAmount - 0;//ABS(ServiceTaxEntry."Service Tax Amount");
                                AppliedServiceTaxECessAmt := ServiceTaxeCessAmount - 0;// ABS(ServiceTaxEntry."eCess Amount");
                                AppliedServiceTaxSHECessAmt := ServiceTaxSHECessAmount - 0;// ABS(ServiceTaxEntry."SHE Cess Amount");
                                AppliedServiceTaxSBCAmt := ServiceTaxSBCAmount - 0;//ABS(ServiceTaxEntry."Service Tax SBC Amount");
                                AppliedKKCessAmt := KKCessAmount - 0;// ABS(ServiceTaxEntry."KK Cess Amount");
                                                                     // END ELSE BEGIN
                                AppliedServiceTaxAmt := ServiceTaxAmount;
                                AppliedServiceTaxECessAmt := ServiceTaxeCessAmount;
                                AppliedServiceTaxSHECessAmt := ServiceTaxSHECessAmount;
                                AppliedServiceTaxSBCAmt := ServiceTaxSBCAmount;
                                AppliedKKCessAmt := KKCessAmount;
                                // END;
                                //END ELSE BEGIN
                                ServiceTaxAmt := ServiceTaxAmount;
                                ServiceTaxECessAmt := ServiceTaxeCessAmount;
                                ServiceTaxSHECessAmt := ServiceTaxSHECessAmount;
                                ServiceTaxSBCAmt := ServiceTaxSBCAmount;
                                KKCessAmt := KKCessAmount
                            END;
                        end;


                        trigger OnPreDataItem()
                        begin
                            VATAmountLine.DELETEALL;
                            SalesShipmentBuffer.RESET;
                            SalesShipmentBuffer.DELETEALL;
                            FirstValueEntryNo := 0;
                            MoreLines := FIND('+');
                            WHILE MoreLines AND (Description = '') AND ("No." = '') AND (Quantity = 0) AND (Amount = 0) DO
                                MoreLines := NEXT(-1) <> 0;
                            IF NOT MoreLines THEN
                                CurrReport.BREAK;
                            SETRANGE("Line No.", 0, "Line No.");
                            CurrReport.CREATETOTALS("Line Amount", Amount, "Amount Including VAT", "Inv. Discount Amount");// "Excise Amount", "Tax Amount",
                                                                                                                           //"Service Tax Amount", "Service Tax eCess Amount", "Amount To Customer", "Service Tax SBC Amount");
                                                                                                                           //CurrReport.CREATETOTALS("KK Cess Amount");
                                                                                                                           //CurrReport.CREATETOTALS("Service Tax SHE Cess Amount");
                        end;
                    }
                    dataitem(VATCounter; Integer)
                    {
                        DataItemTableView = SORTING(Number);
                        column(VATAmountLineVATBase; VATAmountLine."VAT Base")
                        {
                            AutoFormatExpression = "Sales Invoice Line".GetCurrencyCode;
                            AutoFormatType = 1;
                        }
                        column(VATAmountLineVATAmount; VATAmountLine."VAT Amount")
                        {
                            AutoFormatExpression = "Sales Invoice Header"."Currency Code";
                            AutoFormatType = 1;
                        }
                        column(VATAmountLineLineAmount; VATAmountLine."Line Amount")
                        {
                            AutoFormatExpression = "Sales Invoice Header"."Currency Code";
                            AutoFormatType = 1;
                        }
                        column(VATAmtLineInvDiscBaseAmt; VATAmountLine."Inv. Disc. Base Amount")
                        {
                            AutoFormatExpression = "Sales Invoice Header"."Currency Code";
                            AutoFormatType = 1;
                        }
                        column(VATAmtLineInvDiscAmt; VATAmountLine."Invoice Discount Amount")
                        {
                            AutoFormatExpression = "Sales Invoice Header"."Currency Code";
                            AutoFormatType = 1;
                        }
                        column(VATAmtLineVAT_VATCounter; VATAmountLine."VAT %")
                        {
                            DecimalPlaces = 0 : 5;
                        }
                        column(VATAmtLineVATIdentifier_VATCounter; VATAmountLine."VAT Identifier")
                        {
                        }
                        column(VATAmountSpecificationCaption; VATAmountSpecificationCaptionLbl)
                        {
                        }
                        column(InvDiscBaseAmtCaption; InvDiscBaseAmtCaptionLbl)
                        {
                        }
                        column(LineAmountCaption; LineAmountCaptionLbl)
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin
                            VATAmountLine.GetLine(Number);
                        end;

                        trigger OnPreDataItem()
                        begin
                            SETRANGE(Number, 1, VATAmountLine.COUNT);
                            CurrReport.CREATETOTALS(
                              VATAmountLine."Line Amount", VATAmountLine."Inv. Disc. Base Amount",
                              VATAmountLine."Invoice Discount Amount", VATAmountLine."VAT Base", VATAmountLine."VAT Amount");
                        end;
                    }
                    dataitem(VatCounterLCY; Integer)
                    {
                        DataItemTableView = SORTING(Number);
                        column(VALSpecLCYHeader; VALSpecLCYHeader)
                        {
                        }
                        column(VALExchRate; VALExchRate)
                        {
                        }
                        column(VALVATBaseLCY; VALVATBaseLCY)
                        {
                            AutoFormatType = 1;
                        }
                        column(VALVATAmountLCY; VALVATAmountLCY)
                        {
                            AutoFormatType = 1;
                        }
                        column(VATAmtLineVAT_VatCounterLCY; VATAmountLine."VAT %")
                        {
                            DecimalPlaces = 0 : 5;
                        }
                        column(VATAmtLineVATIdentifier_VatCounterLCY; VATAmountLine."VAT Identifier")
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin
                            VATAmountLine.GetLine(Number);
                            VALVATBaseLCY :=
                              VATAmountLine.GetBaseLCY(
                                "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                "Sales Invoice Header"."Currency Factor");
                            VALVATAmountLCY :=
                              VATAmountLine.GetAmountLCY(
                                "Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code",
                                "Sales Invoice Header"."Currency Factor");
                        end;

                        trigger OnPreDataItem()
                        begin
                            IF (NOT GLSetup."Print VAT specification in LCY") OR
                               ("Sales Invoice Header"."Currency Code" = '')
                            THEN
                                CurrReport.BREAK;

                            SETRANGE(Number, 1, VATAmountLine.COUNT);
                            CurrReport.CREATETOTALS(VALVATBaseLCY, VALVATAmountLCY);

                            IF GLSetup."LCY Code" = '' THEN
                                VALSpecLCYHeader := Text007 + Text008
                            ELSE
                                VALSpecLCYHeader := Text007 + FORMAT(GLSetup."LCY Code");

                            CurrExchRate.FindCurrency("Sales Invoice Header"."Posting Date", "Sales Invoice Header"."Currency Code", 1);
                            CalculatedExchRate := ROUND(1 / "Sales Invoice Header"."Currency Factor" * CurrExchRate."Exchange Rate Amount", 0.000001);
                            VALExchRate := STRSUBSTNO(Text009, CalculatedExchRate, CurrExchRate."Exchange Rate Amount");
                        end;
                    }
                    dataitem(Total; Integer)
                    {
                        DataItemTableView = SORTING(Number)
                                            WHERE(Number = CONST(1));
                    }
                    dataitem(Total2; Integer)
                    {
                        DataItemTableView = SORTING(Number)
                                            WHERE(Number = CONST(1));
                        column(SellToCustNo_SalesInvHdr; "Sales Invoice Header"."Sell-to Customer No.")
                        {
                        }
                        column(ShipToAddr1; ShipToAddr[1])
                        {
                        }
                        column(ShipToAddr2; ShipToAddr[2])
                        {
                        }
                        column(ShipToAddr3; ShipToAddr[3])
                        {
                        }
                        column(ShipToAddr4; ShipToAddr[4])
                        {
                        }
                        column(ShipToAddr5; ShipToAddr[5])
                        {
                        }
                        column(ShipToAddr6; ShipToAddr[6])
                        {
                        }
                        column(ShipToAddr7; ShipToAddr[7])
                        {
                        }
                        column(ShipToAddr8; ShipToAddr[8])
                        {
                        }
                        column(ShipToAddressCaption; ShipToAddressCaptionLbl)
                        {
                        }
                        column(SellToCustNo_SalesInvHdrCaption; "Sales Invoice Header".FIELDCAPTION("Sell-to Customer No."))
                        {
                        }

                        trigger OnPreDataItem()
                        begin
                            IF NOT ShowShippingAddr THEN
                                CurrReport.BREAK;
                        end;
                    }
                }

                trigger OnAfterGetRecord()
                begin
                    IF Number > 1 THEN BEGIN
                        CopyText := Text003;
                        OutputNo += 1;
                    END;
                    CurrReport.PAGENO := 1;

                    TotalSubTotal := 0;
                    TotalInvoiceDiscountAmount := 0;
                    TotalAmount := 0;
                    TotalAmountVAT := 0;
                    TotalAmountInclVAT := 0;
                    TotalPaymentDiscountOnVAT := 0;
                    TotalExciseAmt := 0;
                    TotalTaxAmt := 0;
                    ServiceTaxAmount := 0;
                    ServiceTaxeCessAmount := 0;
                    ServiceTaxSHECessAmount := 0;
                    ServiceTaxSBCAmount := 0;
                    KKCessAmount := 0;

                    OtherTaxesAmount := 0;
                    ChargesAmount := 0;
                    AppliedServiceTaxSHECessAmt := 0;
                    AppliedServiceTaxECessAmt := 0;
                    AppliedServiceTaxAmt := 0;
                    AppliedServiceTaxSBCAmt := 0;
                    AppliedKKCessAmt := 0;
                    ServiceTaxSHECessAmt := 0;
                    ServiceTaxECessAmt := 0;
                    ServiceTaxAmt := 0;
                    ServiceTaxSBCAmt := 0;
                    KKCessAmt := 0;
                    TotalTCSAmount := 0;
                end;

                trigger OnPostDataItem()
                begin
                    IF NOT CurrReport.PREVIEW THEN
                        SalesInvCountPrinted.RUN("Sales Invoice Header");
                end;

                trigger OnPreDataItem()
                begin
                    NoOfLoops := ABS(NoOfCopies) + Cust."Invoice Copies" + 1;
                    IF NoOfLoops <= 0 THEN
                        NoOfLoops := 1;
                    CopyText := '';
                    SETRANGE(Number, 1, NoOfLoops);
                    OutputNo := 1;
                end;
            }

            trigger OnAfterGetRecord()
            var
                SalesInvLine: Record 113;
                Location: Record 14;
            begin
                //CurrReport.LANGUAGE := Language.GetLanguageID("Language Code");
                //IsGSTApplicable := GSTManagement.IsGSTApplicable(Structure);
                Customer.GET("Bill-to Customer No.");
                IF RespCenter.GET("Responsibility Center") THEN BEGIN
                    FormatAddr.RespCenter(CompanyAddr, RespCenter);
                    CompanyInfo."Phone No." := RespCenter."Phone No.";
                    CompanyInfo."Fax No." := RespCenter."Fax No.";
                END ELSE
                    FormatAddr.Company(CompanyAddr, CompanyInfo);

                DimSetEntry1.SETRANGE("Dimension Set ID", "Dimension Set ID");

                IF "Order No." = '' THEN
                    OrderNoText := ''
                ELSE
                    OrderNoText := FIELDCAPTION("Order No.");
                IF "Salesperson Code" = '' THEN BEGIN
                    SalesPurchPerson.INIT;
                    SalesPersonText := '';
                END ELSE BEGIN
                    SalesPurchPerson.GET("Salesperson Code");
                    SalesPersonText := Text000;
                END;
                IF "Your Reference" = '' THEN
                    ReferenceText := ''
                ELSE
                    ReferenceText := FIELDCAPTION("Your Reference");
                IF "VAT Registration No." = '' THEN
                    VATNoText := ''
                ELSE
                    VATNoText := FIELDCAPTION("VAT Registration No.");
                IF "Currency Code" = '' THEN BEGIN
                    GLSetup.TESTFIELD("LCY Code");
                    TotalText := STRSUBSTNO(Text001, GLSetup."LCY Code");
                    TotalInclVATText := STRSUBSTNO(Text13700, GLSetup."LCY Code");
                    TotalExclVATText := STRSUBSTNO(Text13701, GLSetup."LCY Code");
                END ELSE BEGIN
                    TotalText := STRSUBSTNO(Text001, "Currency Code");
                    TotalInclVATText := STRSUBSTNO(Text13700, "Currency Code");
                    TotalExclVATText := STRSUBSTNO(Text13701, "Currency Code");
                END;
                FormatAddr.SalesInvBillTo(CustAddr, "Sales Invoice Header");
                IF NOT Cust.GET("Bill-to Customer No.") THEN
                    CLEAR(Cust);

                IF "Payment Terms Code" = '' THEN
                    PaymentTerms.INIT
                ELSE BEGIN
                    PaymentTerms.GET("Payment Terms Code");
                    PaymentTerms.TranslateDescription(PaymentTerms, "Language Code");
                END;
                IF "Shipment Method Code" = '' THEN
                    ShipmentMethod.INIT
                ELSE BEGIN
                    ShipmentMethod.GET("Shipment Method Code");
                    ShipmentMethod.TranslateDescription(ShipmentMethod, "Language Code");
                END;
                //FormatAddr.SalesInvShipTo(ShipToAddr, "Sales Invoice Header");
                ShowShippingAddr := "Sell-to Customer No." <> "Bill-to Customer No.";
                FOR i := 1 TO ARRAYLEN(ShipToAddr) DO
                    IF ShipToAddr[i] <> CustAddr[i] THEN
                        ShowShippingAddr := TRUE;

                IF LogInteraction THEN
                    IF NOT CurrReport.PREVIEW THEN BEGIN
                        IF "Bill-to Contact No." <> '' THEN
                            SegManagement.LogDocument(
                              4, "No.", 0, 0, DATABASE::Contact, "Bill-to Contact No.", "Salesperson Code",
                              "Campaign No.", "Posting Description", '')
                        ELSE
                            SegManagement.LogDocument(
                              4, "No.", 0, 0, DATABASE::Customer, "Bill-to Customer No.", "Salesperson Code",
                              "Campaign No.", "Posting Description", '');
                    END;

                SupplementaryText := '';
                SalesInvLine.SETRANGE("Document No.", "No.");
                //SalesInvLine.SETRANGE(Supplementary, TRUE);
                IF SalesInvLine.FINDFIRST THEN
                    SupplementaryText := Text16500;

                IF Location.GET("Location Code") THEN
                    ServiceTaxRegistrationNo := ''// Location."Service Tax Registration No."
                ELSE
                    ServiceTaxRegistrationNo := '';//CompanyInfo."Service Tax Registration No.";

                SalesInvoice.RESET;
                SalesInvoice.SETRANGE("No.", "Sales Invoice Header"."No.");
                IF SalesInvoice.FINDFIRST THEN
                    Doc := SalesInvoice."No.";
                //   MESSAGE('%1',Doc);


                "Sales Invoice Header".CALCFIELDS("Sales Invoice Header"."QR code");

                BarcodeForQuarantineLabel(Doc);
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
                group(Options)
                {
                    Caption = 'Options';
                    field(NoOfCopies; NoOfCopies)
                    {
                        Caption = 'No. of Copies';
                    }
                    field(ShowInternalInfo; ShowInternalInfo)
                    {
                        Caption = 'Show Internal Information';
                    }
                    field(LogInteraction; LogInteraction)
                    {
                        Caption = 'Log Interaction';
                        Enabled = LogInteractionEnable;
                    }
                    field(DisplayAsmInformation; DisplayAssemblyInformation)
                    {
                        Caption = 'Show Assembly Components';
                    }
                }
            }
        }

        actions
        {
        }

        trigger OnInit()
        begin
            LogInteractionEnable := TRUE;
        end;

        trigger OnOpenPage()
        begin
            InitLogInteraction;
            LogInteractionEnable := LogInteraction;
        end;
    }

    labels
    {
    }

    trigger OnInitReport()
    begin
        GLSetup.GET;
        CompanyInfo.GET;
        SalesSetup.GET;
        CompanyInfo.VerifyAndSetPaymentInfo;
        CASE SalesSetup."Logo Position on Documents" OF
            SalesSetup."Logo Position on Documents"::Left:
                BEGIN
                    CompanyInfo3.GET;
                    CompanyInfo3.CALCFIELDS(Picture);
                END;
            SalesSetup."Logo Position on Documents"::Center:
                BEGIN
                    CompanyInfo1.GET;
                    CompanyInfo1.CALCFIELDS(Picture);
                END;
            SalesSetup."Logo Position on Documents"::Right:
                BEGIN
                    CompanyInfo2.GET;
                    CompanyInfo2.CALCFIELDS(Picture);
                END;
        END;
    end;

    trigger OnPreReport()
    begin
        IF NOT CurrReport.USEREQUESTPAGE THEN
            InitLogInteraction;
    end;

    var
        Text000: Label 'Salesperson';
        Text001: Label 'Total %1';
        Text003: Label ' COPY';
        Text004: Label 'Sales - Invoice%1';
        PageCaptionCap: Label 'Page %1 of %2';
        GLSetup: Record 98;
        ShipmentMethod: Record 10;
        PaymentTerms: Record 3;
        SalesPurchPerson: Record 13;
        CompanyInfo: Record 79;
        CompanyInfo1: Record 79;
        CompanyInfo2: Record 79;
        CompanyInfo3: Record 79;
        //GSTComponent: Record "GST Component";
        Customer: Record 18;
        DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry";
        SalesSetup: Record 311;
        Cust: Record 18;
        VATAmountLine: Record 290 temporary;
        DimSetEntry1: Record 480;
        DimSetEntry2: Record 480;
        RespCenter: Record 5714;
        Language: Record 8;
        CurrExchRate: Record 330;
        TempPostedAsmLine: Record 911 temporary;
        //GSTManagement: Codeunit "GST Management";
        SalesInvCountPrinted: Codeunit "Sales Inv.-Printed";
        FormatAddr: Codeunit 365;
        SegManagement: Codeunit SegManagement;
        SalesShipmentBuffer: Record 7190 temporary;
        GSTCompAmount: array[20] of Decimal;
        GSTComponentCode: array[20] of Code[10];
        PostedShipmentDate: Date;
        CustAddr: array[8] of Text[50];
        ShipToAddr: array[8] of Text[50];
        CompanyAddr: array[8] of Text[50];
        OrderNoText: Text[80];
        SalesPersonText: Text[30];
        VATNoText: Text[80];
        ReferenceText: Text[80];
        TotalText: Text[50];
        TotalExclVATText: Text[50];
        TotalInclVATText: Text[50];
        MoreLines: Boolean;
        NoOfCopies: Integer;
        NoOfLoops: Integer;
        CopyText: Text[30];
        ShowShippingAddr: Boolean;
        i: Integer;
        NextEntryNo: Integer;
        FirstValueEntryNo: Integer;
        DimText: Text[120];
        OldDimText: Text[75];
        ShowInternalInfo: Boolean;
        Continue: Boolean;
        LogInteraction: Boolean;
        VALVATBaseLCY: Decimal;
        VALVATAmountLCY: Decimal;
        VALSpecLCYHeader: Text[80];
        Text007: Label 'VAT Amount Specification in ';
        Text008: Label 'Local Currency';
        VALExchRate: Text[50];
        Text009: Label 'Exchange rate: %1/%2';
        CalculatedExchRate: Decimal;
        Text010: Label 'Sales - Prepayment Invoice %1';
        OutputNo: Integer;
        TotalSubTotal: Decimal;
        TotalAmount: Decimal;
        TotalAmountInclVAT: Decimal;
        TotalAmountVAT: Decimal;
        TotalInvoiceDiscountAmount: Decimal;
        TotalPaymentDiscountOnVAT: Decimal;
        ChargesAmount: Decimal;
        OtherTaxesAmount: Decimal;
        Text13700: Label 'Total %1 Incl. Taxes';
        Text13701: Label 'Total %1 Excl. Taxes';
        SupplementaryText: Text[30];
        Text16500: Label 'Supplementary Invoice';
        //ServiceTaxEntry: Record 16473;
        ServiceTaxAmt: Decimal;
        ServiceTaxECessAmt: Decimal;
        AppliedServiceTaxAmt: Decimal;
        AppliedServiceTaxECessAmt: Decimal;
        ServiceTaxSHECessAmt: Decimal;
        AppliedServiceTaxSHECessAmt: Decimal;
        TotalTaxAmt: Decimal;
        TotalExciseAmt: Decimal;
        TotalTCSAmount: Decimal;
        ServiceTaxAmount: Decimal;
        ServiceTaxeCessAmount: Decimal;
        ServiceTaxSHECessAmount: Decimal;
        [InDataSet]
        LogInteractionEnable: Boolean;
        DisplayAssemblyInformation: Boolean;
        PhoneNoCaptionLbl: Label 'Phone No.';
        HomePageCaptionCap: Label 'Home Page';
        VATRegNoCaptionLbl: Label 'VAT Registration No.';
        GiroNoCaptionLbl: Label 'Giro No.';
        BankNameCaptionLbl: Label 'Bank';
        BankAccNoCaptionLbl: Label 'Account No.';
        DueDateCaptionLbl: Label 'Due Date';
        InvoiceNoCaptionLbl: Label 'Invoice No.';
        PostingDateCaptionLbl: Label 'Posting Date';
        PLAEntryNoCaptionLbl: Label 'PLA Entry No.';
        RG23AEntryNoCaptionLbl: Label 'RG23A Entry No.';
        RG23CEntryNoCaptionLbl: Label 'RG23C Entry No.';
        HeaderDimensionsCaptionLbl: Label 'Header Dimensions';
        UnitPriceCaptionLbl: Label 'Unit Price';
        DiscountCaptionLbl: Label 'Discount %';
        AmountCaptionLbl: Label 'Amount';
        LineDiscountCaptionLbl: Label 'Line Discount Amount';
        PostedShipmentDateCaptionLbl: Label 'Posted Shipment Date';
        SubtotalCaptionLbl: Label 'Subtotal';
        ExciseAmountCaptionLbl: Label 'Excise Amount';
        TaxAmountCaptionLbl: Label 'Tax Amount';
        ServiceTaxAmountCaptionLbl: Label 'Service Tax Amount';
        ChargesAmountCaptionLbl: Label 'Charges Amount';
        OtherTaxesAmountCaptionLbl: Label 'Other Taxes Amount';
        ServTaxeCessAmtCaptionLbl: Label 'Service Tax eCess Amount';
        TCSAmountCaptionLbl: Label 'TCS Amount';
        SvcTaxAmtAppliedCaptionLbl: Label 'Svc Tax Amt (Applied)';
        SvcTaxeCessAmtAppliedCaptionLbl: Label 'Svc Tax eCess Amt (Applied)';
        ServTaxSHECessAmtCaptionLbl: Label 'Service Tax SHE Cess Amount';
        SvcTaxSHECessAmtAppliedCaptionLbl: Label 'Svc Tax SHECess Amt(Applied)';
        PaymentDiscVATCaptionLbl: Label 'Payment Discount on VAT';
        ShipmentCaptionLbl: Label 'Shipment';
        LineDimensionsCaptionLbl: Label 'Line Dimensions';
        VATAmountSpecificationCaptionLbl: Label 'VAT Amount Specification';
        InvDiscBaseAmtCaptionLbl: Label 'Invoice Discount Base Amount';
        LineAmountCaptionLbl: Label 'Line Amount';
        ShipToAddressCaptionLbl: Label 'Ship-to Address';
        ServiceTaxRegistrationNo: Code[20];
        ServiceTaxRegistrationNoLbl: Label 'Service Tax Registration No.';
        InvDiscountAmountCaptionLbl: Label 'Invoice Discount Amount';
        VATPercentageCaptionLbl: Label 'VAT %';
        VATAmountCaptionLbl: Label 'VAT Amount';
        VATIdentifierCaptionLbl: Label 'VAT Identifier';
        TotalCaptionLbl: Label 'Total';
        VATBaseCaptionLbl: Label 'VAT Base';
        PaymentTermsCaptionLbl: Label 'Payment Terms';
        ShipmentMethodCaptionLbl: Label 'Shipment Method';
        EMailCaptionLbl: Label 'E-Mail';
        DocumentDateCaptionLbl: Label 'Document Date';
        ServTaxSBCAmtCaptionLbl: Label 'SBC Amount';
        SvcTaxSBCAmtAppliedCaptionLbl: Label 'SBC Amt (Applied)';
        ServiceTaxSBCAmount: Decimal;
        ServiceTaxSBCAmt: Decimal;
        AppliedServiceTaxSBCAmt: Decimal;
        KKCessAmount: Decimal;
        KKCessAmt: Decimal;
        AppliedKKCessAmt: Decimal;
        KKCessAmtCaptionLbl: Label 'KKC Amount';
        KKCessAmtAppliedCaptionLbl: Label 'KKC Amt (Applied)';
        IsGSTApplicable: Boolean;
        J: Integer;
        CompanyRegistrationLbl: Label 'Company Registration No.';
        CustomerRegistrationLbl: Label 'Customer GST Reg No.';
        SalesInvoice: Record 112;
        Doc: Text;

    // //[Scope('Internal')]
    procedure InitLogInteraction()
    begin
        LogInteraction := SegManagement.FindInteractTmplCode(4) <> '';
    end;

    // //[Scope('Internal')]
    procedure FindPostedShipmentDate(): Date
    var
        SalesShipmentHeader: Record 110;
        SalesShipmentBuffer2: Record 7190 temporary;
    begin
        NextEntryNo := 1;
        IF "Sales Invoice Line"."Shipment No." <> '' THEN
            IF SalesShipmentHeader.GET("Sales Invoice Line"."Shipment No.") THEN
                EXIT(SalesShipmentHeader."Posting Date");

        IF "Sales Invoice Header"."Order No." = '' THEN
            EXIT("Sales Invoice Header"."Posting Date");

        CASE "Sales Invoice Line".Type OF
            "Sales Invoice Line".Type::Item:
                GenerateBufferFromValueEntry("Sales Invoice Line");
            "Sales Invoice Line".Type::"G/L Account", "Sales Invoice Line".Type::Resource,
          "Sales Invoice Line".Type::"Charge (Item)", "Sales Invoice Line".Type::"Fixed Asset":
                GenerateBufferFromShipment("Sales Invoice Line");
            "Sales Invoice Line".Type::" ":
                EXIT(0D);
        END;

        SalesShipmentBuffer.RESET;
        SalesShipmentBuffer.SETRANGE("Document No.", "Sales Invoice Line"."Document No.");
        SalesShipmentBuffer.SETRANGE("Line No.", "Sales Invoice Line"."Line No.");
        IF SalesShipmentBuffer.FIND('-') THEN BEGIN
            SalesShipmentBuffer2 := SalesShipmentBuffer;
            IF SalesShipmentBuffer.NEXT = 0 THEN BEGIN
                SalesShipmentBuffer.GET(
                  SalesShipmentBuffer2."Document No.", SalesShipmentBuffer2."Line No.", SalesShipmentBuffer2."Entry No.");
                SalesShipmentBuffer.DELETE;
                EXIT(SalesShipmentBuffer2."Posting Date");
            END;
            SalesShipmentBuffer.CALCSUMS(Quantity);
            IF SalesShipmentBuffer.Quantity <> "Sales Invoice Line".Quantity THEN BEGIN
                SalesShipmentBuffer.DELETEALL;
                EXIT("Sales Invoice Header"."Posting Date");
            END;
        END ELSE
            EXIT("Sales Invoice Header"."Posting Date");
    end;

    // //[Scope('Internal')]
    procedure GenerateBufferFromValueEntry(SalesInvoiceLine2: Record 113)
    var
        ValueEntry: Record 5802;
        ItemLedgerEntry: Record 32;
        TotalQuantity: Decimal;
        Quantity: Decimal;
    begin
        TotalQuantity := SalesInvoiceLine2."Quantity (Base)";
        ValueEntry.SETCURRENTKEY("Document No.");
        ValueEntry.SETRANGE("Document No.", SalesInvoiceLine2."Document No.");
        ValueEntry.SETRANGE("Posting Date", "Sales Invoice Header"."Posting Date");
        ValueEntry.SETRANGE("Item Charge No.", '');
        ValueEntry.SETFILTER("Entry No.", '%1..', FirstValueEntryNo);
        IF ValueEntry.FIND('-') THEN
            REPEAT
                IF ItemLedgerEntry.GET(ValueEntry."Item Ledger Entry No.") THEN BEGIN
                    IF SalesInvoiceLine2."Qty. per Unit of Measure" <> 0 THEN
                        Quantity := ValueEntry."Invoiced Quantity" / SalesInvoiceLine2."Qty. per Unit of Measure"
                    ELSE
                        Quantity := ValueEntry."Invoiced Quantity";
                    AddBufferEntry(
                      SalesInvoiceLine2,
                      -Quantity,
                      ItemLedgerEntry."Posting Date");
                    TotalQuantity := TotalQuantity + ValueEntry."Invoiced Quantity";
                END;
                FirstValueEntryNo := ValueEntry."Entry No." + 1;
            UNTIL (ValueEntry.NEXT = 0) OR (TotalQuantity = 0);
    end;

    // //[Scope('Internal')]
    procedure GenerateBufferFromShipment(SalesInvoiceLine: Record 113)
    var
        SalesInvoiceHeader: Record 112;
        SalesInvoiceLine2: Record 113;
        SalesShipmentHeader: Record 110;
        SalesShipmentLine: Record 111;
        TotalQuantity: Decimal;
        Quantity: Decimal;
    begin
        TotalQuantity := 0;
        SalesInvoiceHeader.SETCURRENTKEY("Order No.");
        SalesInvoiceHeader.SETFILTER("No.", '..%1', "Sales Invoice Header"."No.");
        SalesInvoiceHeader.SETRANGE("Order No.", "Sales Invoice Header"."Order No.");
        IF SalesInvoiceHeader.FIND('-') THEN
            REPEAT
                SalesInvoiceLine2.SETRANGE("Document No.", SalesInvoiceHeader."No.");
                SalesInvoiceLine2.SETRANGE("Line No.", SalesInvoiceLine."Line No.");
                SalesInvoiceLine2.SETRANGE(Type, SalesInvoiceLine.Type);
                SalesInvoiceLine2.SETRANGE("No.", SalesInvoiceLine."No.");
                SalesInvoiceLine2.SETRANGE("Unit of Measure Code", SalesInvoiceLine."Unit of Measure Code");
                IF SalesInvoiceLine2.FIND('-') THEN
                    REPEAT
                        TotalQuantity := TotalQuantity + SalesInvoiceLine2.Quantity;
                    UNTIL SalesInvoiceLine2.NEXT = 0;
            UNTIL SalesInvoiceHeader.NEXT = 0;

        SalesShipmentLine.SETCURRENTKEY("Order No.", "Order Line No.");
        SalesShipmentLine.SETRANGE("Order No.", "Sales Invoice Header"."Order No.");
        SalesShipmentLine.SETRANGE("Order Line No.", SalesInvoiceLine."Line No.");
        SalesShipmentLine.SETRANGE("Line No.", SalesInvoiceLine."Line No.");
        SalesShipmentLine.SETRANGE(Type, SalesInvoiceLine.Type);
        SalesShipmentLine.SETRANGE("No.", SalesInvoiceLine."No.");
        SalesShipmentLine.SETRANGE("Unit of Measure Code", SalesInvoiceLine."Unit of Measure Code");
        SalesShipmentLine.SETFILTER(Quantity, '<>%1', 0);

        IF SalesShipmentLine.FIND('-') THEN
            REPEAT
                IF "Sales Invoice Header"."Get Shipment Used" THEN
                    CorrectShipment(SalesShipmentLine);
                IF ABS(SalesShipmentLine.Quantity) <= ABS(TotalQuantity - SalesInvoiceLine.Quantity) THEN
                    TotalQuantity := TotalQuantity - SalesShipmentLine.Quantity
                ELSE BEGIN
                    IF ABS(SalesShipmentLine.Quantity) > ABS(TotalQuantity) THEN
                        SalesShipmentLine.Quantity := TotalQuantity;
                    Quantity :=
                      SalesShipmentLine.Quantity - (TotalQuantity - SalesInvoiceLine.Quantity);

                    TotalQuantity := TotalQuantity - SalesShipmentLine.Quantity;
                    SalesInvoiceLine.Quantity := SalesInvoiceLine.Quantity - Quantity;

                    IF SalesShipmentHeader.GET(SalesShipmentLine."Document No.") THEN
                        AddBufferEntry(
                          SalesInvoiceLine,
                          Quantity,
                          SalesShipmentHeader."Posting Date");
                END;
            UNTIL (SalesShipmentLine.NEXT = 0) OR (TotalQuantity = 0);
    end;

    // //[Scope('Internal')]
    procedure CorrectShipment(var SalesShipmentLine: Record 111)
    var
        SalesInvoiceLine: Record 113;
    begin
        SalesInvoiceLine.SETCURRENTKEY("Shipment No.", "Shipment Line No.");
        SalesInvoiceLine.SETRANGE("Shipment No.", SalesShipmentLine."Document No.");
        SalesInvoiceLine.SETRANGE("Shipment Line No.", SalesShipmentLine."Line No.");
        IF SalesInvoiceLine.FIND('-') THEN
            REPEAT
                SalesShipmentLine.Quantity := SalesShipmentLine.Quantity - SalesInvoiceLine.Quantity;
            UNTIL SalesInvoiceLine.NEXT = 0;
    end;

    // //[Scope('Internal')]
    procedure AddBufferEntry(SalesInvoiceLine: Record 113; QtyOnShipment: Decimal; PostingDate: Date)
    begin
        SalesShipmentBuffer.SETRANGE("Document No.", SalesInvoiceLine."Document No.");
        SalesShipmentBuffer.SETRANGE("Line No.", SalesInvoiceLine."Line No.");
        SalesShipmentBuffer.SETRANGE("Posting Date", PostingDate);
        IF SalesShipmentBuffer.FIND('-') THEN BEGIN
            SalesShipmentBuffer.Quantity := SalesShipmentBuffer.Quantity + QtyOnShipment;
            SalesShipmentBuffer.MODIFY;
            EXIT;
        END;

        WITH SalesShipmentBuffer DO BEGIN
            "Document No." := SalesInvoiceLine."Document No.";
            "Line No." := SalesInvoiceLine."Line No.";
            "Entry No." := NextEntryNo;
            Type := SalesInvoiceLine.Type;
            "No." := SalesInvoiceLine."No.";
            Quantity := QtyOnShipment;
            "Posting Date" := PostingDate;
            INSERT;
            NextEntryNo := NextEntryNo + 1
        END;
    end;

    local procedure DocumentCaption(): Text[250]
    begin
        IF "Sales Invoice Header"."Prepayment Invoice" THEN
            EXIT(Text010);
        EXIT(Text004);
    end;

    //  //[Scope('Internal')]
    procedure InitializeRequest(NewNoOfCopies: Integer; NewShowInternalInfo: Boolean; NewLogInteraction: Boolean; DisplayAsmInfo: Boolean)
    begin
        NoOfCopies := NewNoOfCopies;
        ShowInternalInfo := NewShowInternalInfo;
        LogInteraction := NewLogInteraction;
        DisplayAssemblyInformation := DisplayAsmInfo;
    end;

    //   //[Scope('Internal')]
    procedure CollectAsmInformation()
    var
        ValueEntry: Record 5802;
        ItemLedgerEntry: Record 32;
        PostedAsmHeader: Record 910;
        PostedAsmLine: Record 911;
        SalesShipmentLine: Record 111;
    begin
        TempPostedAsmLine.DELETEALL;
        IF "Sales Invoice Line".Type <> "Sales Invoice Line".Type::Item THEN
            EXIT;
        WITH ValueEntry DO BEGIN
            SETCURRENTKEY("Document No.");
            SETRANGE("Document No.", "Sales Invoice Line"."Document No.");
            SETRANGE("Document Type", "Document Type"::"Sales Invoice");
            SETRANGE("Document Line No.", "Sales Invoice Line"."Line No.");
            IF NOT FINDSET THEN
                EXIT;
        END;
        REPEAT
            IF ItemLedgerEntry.GET(ValueEntry."Item Ledger Entry No.") THEN
                IF ItemLedgerEntry."Document Type" = ItemLedgerEntry."Document Type"::"Sales Shipment" THEN BEGIN
                    SalesShipmentLine.GET(ItemLedgerEntry."Document No.", ItemLedgerEntry."Document Line No.");
                    IF SalesShipmentLine.AsmToShipmentExists(PostedAsmHeader) THEN BEGIN
                        PostedAsmLine.SETRANGE("Document No.", PostedAsmHeader."No.");
                        IF PostedAsmLine.FINDSET THEN
                            REPEAT
                                TreatAsmLineBuffer(PostedAsmLine);
                            UNTIL PostedAsmLine.NEXT = 0;
                    END;
                END;
        UNTIL ValueEntry.NEXT = 0;
    end;

    // //[Scope('Internal')]
    procedure TreatAsmLineBuffer(PostedAsmLine: Record 911)
    begin
        CLEAR(TempPostedAsmLine);
        TempPostedAsmLine.SETRANGE(Type, PostedAsmLine.Type);
        TempPostedAsmLine.SETRANGE("No.", PostedAsmLine."No.");
        TempPostedAsmLine.SETRANGE("Variant Code", PostedAsmLine."Variant Code");
        TempPostedAsmLine.SETRANGE(Description, PostedAsmLine.Description);
        TempPostedAsmLine.SETRANGE("Unit of Measure Code", PostedAsmLine."Unit of Measure Code");
        IF TempPostedAsmLine.FINDFIRST THEN BEGIN
            TempPostedAsmLine.Quantity += PostedAsmLine.Quantity;
            TempPostedAsmLine.MODIFY;
        END ELSE BEGIN
            CLEAR(TempPostedAsmLine);
            TempPostedAsmLine := PostedAsmLine;
            TempPostedAsmLine.INSERT;
        END;
    end;

    //  //[Scope('Internal')]
    procedure GetUOMText(UOMCode: Code[10]): Text[10]
    var
        UnitOfMeasure: Record 204;
    begin
        IF NOT UnitOfMeasure.GET(UOMCode) THEN
            EXIT(UOMCode);
        EXIT(UnitOfMeasure.Description);
    end;

    // //[Scope('Internal')]
    procedure BlanksForIndent(): Text[10]
    begin
        EXIT(PADSTR('', 2, ' '));
    end;

    // //[Scope('Internal')]
    procedure PrintQRCode(DocNoC: Code[30]; LineNoC: Integer; CustomerOrderNo: Code[10]; ItemNo: Text; Qty: Decimal; DocumentNo: Code[10]; PostingDate: Text; GrossRate: Text; NetRate: Text; VendorCode: Code[20]; CrossRefNo: Code[20]; GstAmount: Text; TaxRateUGST: Text; GSTPer: Text; UGST: Text; Cess: Text; AmountToCustomer: Text; HSNCode: Code[10])
    begin

        //QRMgt_lCdu.BarcodeForQuarantineLabel(DocNo,LineNo,ItmNo,MRNNo,BatchNo,LocCode,MfgName,Desc);
        //BarcodeForQuarantineLabel(DocNo,LineNo,ItmNo,MRNNo,BatchNo,LocCode,MfgName,Desc);
    end;

    /* local procedure CreateQRCode(QRCodeInput: Text; var TempBLOB: Record 99008535)
    var
        QRCodeFileName: Text[1024];
    begin
        CLEAR(TempBLOB);
        QRCodeFileName := GetQRCode(QRCodeInput);
        UploadFileBLOBImportandDeleteServerFile(TempBLOB, QRCodeFileName);
    end;

    //  //[Scope('Internal')]
    procedure UploadFileBLOBImportandDeleteServerFile(var TempBlob: Record 99008535; FileName: Text[1024])
    var
        FileManagement: Codeunit 419;
    begin
        FileName := FileManagement.UploadFileSilent(FileName);
        FileManagement.BLOBImportFromServerFile(TempBlob, FileName);
        DeleteServerFile(FileName);
    end;

    local procedure DeleteServerFile(ServerFileName: Text)
    begin
        IF ERASE(ServerFileName) THEN;
    end;

    local procedure GetQRCode(QRCodeInput: Text[300]) QRCodeFileName: Text[1024]
    var
        [RunOnClient]
        IBarCodeProvider: DotNet IBarcodeProvider;
    begin
        GetBarCodeProvider(IBarCodeProvider);
        QRCodeFileName := IBarCodeProvider.GetBarcode(QRCodeInput);
    end;

    ////[Scope('Internal')]
    procedure GetBarCodeProvider(var IBarCodeProvider: DotNet IBarcodeProvider)
    var
        [RunOnClient]
        QRCodeProvider: DotNet QRCodeProvider;
    begin
        IF ISNULL(IBarCodeProvider) THEN
            IBarCodeProvider := QRCodeProvider.QRCodeProvider;
    end; */

    ////[Scope('Internal')]
    procedure BarcodeForQuarantineLabel(DocNoC: Code[30])
    var
        recSIL1: Record 113;
        ItemLedgEntry_lRec: Record 32;
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
        QRTempBlob_lRecTmp: Record 50009 temporary;
        SRNo_lInt: Integer;
        QRCodeInput: Text;
        //TempBlob: Record 99008535;
        R50017_lRpt: Report 50017;
        R50018_lRpt: Report 50018;
        GSTLedgerEntry: Record "GST Ledger Entry";
    begin
        GSTLedgerEntry.SETRANGE(GSTLedgerEntry."Document No.", "Sales Invoice Header"."No.");
        GSTLedgerEntry.SETRANGE(GSTLedgerEntry."Transaction Type", GSTLedgerEntry."Transaction Type"::Sales);
        IF GSTLedgerEntry.FINDFIRST THEN BEGIN
            Doc := GSTLedgerEntry."Scan QrCode E-Invoice";//YSr  //SalesInvoice."No.";
            QRCodeInput := STRSUBSTNO('%1', Doc);
            //CreateQRCode(QRCodeInput, TempBlob);
            //SalesInvoice."QR code" := TempBlob.Blob;
            SalesInvoice.MODIFY;
        END
    end;
}


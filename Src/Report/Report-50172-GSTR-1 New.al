report 50172 "GSTR-1 New"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 15Sep2017   RB-N         New Report Development from Existing Report#16401--GSTR-1
    // 22Sep2017   RB-N         Grouping w.r.t Invoice No., Place of Supply
    //                          Grouping w.r.t HSN Code for Multiple Invoice No.
    // 
    // 07Nov2017   RB-N         Showing transfer Details
    //                          HSN total Issue, Invoice Total Issue
    // 
    // 08Nov2017   RB-N         Ignoring Transfer Receipt
    // 09Nov2017   RB-N         Number Series Details from SalesInvoice & TransferShipment
    // 13Nov2017   RB-N         Taxable Amount Issue, GST % Breakup for Invoice, Cancel Invoice
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/GSTR1New.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'GSTR-1 New';

    dataset
    {
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number)
                                ORDER(Ascending)
                                WHERE(Number = FILTER(1));
            column(InvoiceDetail; InvoiceDetail)
            {
            }
            column(TitleLbl; TitleLbl)
            {
            }
            column(Title2Lbl; Title2Lbl)
            {
            }
            column(Title3Lbl; Title3Lbl)
            {
            }
            column(Title4Lbl; Title4Lbl)
            {
            }
            column(HeadingLbl; HeadingLbl)
            {
            }
            column(GSTLbl; GSTLbl)
            {
            }
            column(PersonNameLbl; PersonNameLbl)
            {
            }
            column(PersonSubLbl; PersonSubLbl)
            {
            }
            column(AggregateTurnOverLbl; AggregateTurnOverLbl)
            {
            }
            column(AggregateSubLbl; AggregateSubLbl)
            {
            }
            column(PeriodLbl; PeriodLbl)
            {
            }
            column(TaxableHeadingLbl; TaxableHeadingLbl)
            {
            }
            column(FiguresLbl; FiguresLbl)
            {
            }
            column(GSTINUINLbl; GSTINUINLbl)
            {
            }
            column(InvoiceLbl; InvoiceLbl)
            {
            }
            column(NoLbl; NoLbl)
            {
            }
            column(DateLbl; DateLbl)
            {
            }
            column(ValueLbl; ValueLbl)
            {
            }
            column(GoodSerLbl; GoodSerLbl)
            {
            }
            column(HSNSACLbl; HSNSACLbl)
            {
            }
            column(TaxableLbls; TaxableLbl)
            {
            }
            column(IGSTLbl; IGSTLbl)
            {
            }
            column(CGSTLbl; CGSTLbl)
            {
            }
            column(SGSTLbl; SGSTLbl)
            {
            }
            column(RateLbl; RateLbl)
            {
            }
            column(AmtLbl; AmtLbl)
            {
            }
            column(POSLbl; POSLbl)
            {
            }
            column(RevChargeLbl; RevChargeLbl)
            {
            }
            column(AssessmentLbl; AssessmentLbl)
            {
            }
            column(ecommerceLbl; ecommerceLbl)
            {
            }
            column(FCLbl; FCLbl)
            {
            }
            column(NotesLbl; NotesLbl)
            {
            }
            column(Note1Lbl; Note1Lbl)
            {
            }
            column(Note2Lbl; Note2Lbl)
            {
            }
            column(Note3Lbl; Note3Lbl)
            {
            }
            column(AmendmentHeadingLbl; AmendmentHeadingLbl)
            {
            }
            column(OriginalInvLbl; OriginalInvLbl)
            {
            }
            column(RevisedOriginalLbl; RevisedOriginalLbl)
            {
            }
            column(Taxable2HeadingLbl; Taxable2HeadingLbl)
            {
            }
            column(RecipientLbl; RecipientLbl)
            {
            }
            column(RecNameLbl; RecNameLbl)
            {
            }
            column(Amendment2HeadingLbl; Amendment2HeadingLbl)
            {
            }
            column(Taxable3HeadlingLbl; Taxable3HeadlingLbl)
            {
            }
            column(StateCodeLbl; StateCodeLbl)
            {
            }
            column(AggregateLbl; AggregateLbl)
            {
            }
            column(SupplyLbl; SupplyLbl)
            {
            }
            column(Note4Lbl; Note4Lbl)
            {
            }
            column(Ammendment3HeadingLbl; Ammendment3HeadingLbl)
            {
            }
            column(OriginalDetLbl; OriginalDetLbl)
            {
            }
            column(MonthPeriodLbl; MonthPeriodLbl)
            {
            }
            column(StatLbl; StatLbl)
            {
            }
            column(ReverseDetLbl; ReverseDetLbl)
            {
            }
            column(POSStatLbl; POSStatLbl)
            {
            }
            column(DetailsHeadingLbl; DetailsHeadingLbl)
            {
            }
            column(GSTINRecNameLbl; GSTINRecNameLbl)
            {
            }
            column(OtherChargeLbl; OtherChargeLbl)
            {
            }
            column(ReverseChargeLbl; ReverseChargeLbl)
            {
            }
            column(DebitCreditLbl; DebitCreditLbl)
            {
            }
            column(DifferentialValLbl; DifferentialValLbl)
            {
            }
            column(DifferentialTaxLbl; DifferentialTaxLbl)
            {
            }
            column(Note5Lbl; Note5Lbl)
            {
            }
            column(Ammendment4HeadingLbl; Ammendment4HeadingLbl)
            {
            }
            column(OriginalLbl; OriginalLbl)
            {
            }
            column(RevisedLbl; RevisedLbl)
            {
            }
            column(OriginalInvDetLbl; OriginalInvDetLbl)
            {
            }
            column(NilrateHeadlingLbl; NilrateHeadlingLbl)
            {
            }
            column(InterSupplyPerLbl; InterSupplyPerLbl)
            {
            }
            column(IntraSupplyPerLbl; IntraSupplyPerLbl)
            {
            }
            column(InterSupplyConsLbl; InterSupplyConsLbl)
            {
            }
            column(IntraSupplyConsLbl; IntraSupplyConsLbl)
            {
            }
            column(NilRateLbl; NilRateLbl)
            {
            }
            column(ExemptLbl; ExemptLbl)
            {
            }
            column(NonGSTLbl; NonGSTLbl)
            {
            }
            column(Note6Lbl; Note6Lbl)
            {
            }
            column(SuppliesLbl; SuppliesLbl)
            {
            }
            column(DescriptionLbl; DescriptionLbl)
            {
            }
            column(WithoutPaymentGSTLbl; WithoutPaymentGSTLbl)
            {
            }
            column(WithPaymentGSTLbl; WithPaymentGSTLbl)
            {
            }
            column(ShippingLbl; ShippingLbl)
            {
            }
            column(Ammendment5HeadingLbl; Ammendment5HeadingLbl)
            {
            }
            column(RevisedInvLbl; RevisedInvLbl)
            {
            }
            column(TaxHeadingLbl; TaxHeadingLbl)
            {
            }
            column(GSTCustNameLbl; GSTCustNameLbl)
            {
            }
            column(DocNoLbl; DocNoLbl)
            {
            }
            column(HSNSACSupplyLbl; HSNSACSupplyLbl)
            {
            }
            column(AmountofAdvLbl; AmountofAdvLbl)
            {
            }
            column(TaxLbl; TaxLbl)
            {
            }
            column(Note7Lbl; Note7Lbl)
            {
            }
            column(Ammendment6HeadingLbl; Ammendment6HeadingLbl)
            {
            }
            column(DocNumLbl; DocNumLbl)
            {
            }
            column(HSCPOSLbl; HSCPOSLbl)
            {
            }
            column(Tax_Lbl; Tax_Lbl)
            {
            }
            column(Tax2HeadingLbl; Tax2HeadingLbl)
            {
            }
            column(TransactionLbl; TransactionLbl)
            {
            }
            column(TaxPaidLbl; TaxPaidLbl)
            {
            }
            column(Note8Lbl; Note8Lbl)
            {
            }
            column(SuppliesHeadlingLbl; SuppliesHeadlingLbl)
            {
            }
            column(Part1HeadlingLbl; Part1HeadlingLbl)
            {
            }
            column(MarchantLbl; MarchantLbl)
            {
            }
            column(GSTINecommerceLbl; GSTINecommerceLbl)
            {
            }
            column(GrossLbl; GrossLbl)
            {
            }
            column(Note9Lbl; Note9Lbl)
            {
            }
            column(SNoLbl; SNoLbl)
            {
            }
            column(Note10Lbl; Note10Lbl)
            {
            }
            column(Part2HeadlingLbl; Part2HeadlingLbl)
            {
            }
            column(Part2AHeadingLbl; Part2AHeadingLbl)
            {
            }
            column(TaxPeriodLbl; TaxPeriodLbl)
            {
            }
            column(PlaceofSuppStateLbl; PlaceofSuppStateLbl)
            {
            }
            column(InvoicesHeadlingLbl; InvoicesHeadlingLbl)
            {
            }
            column(SeriesLbl; SeriesLbl)
            {
            }
            column(FromLbl; FromLbl)
            {
            }
            column(ToLbl; ToLbl)
            {
            }
            column(TotalInvLbl; TotalInvLbl)
            {
            }
            column(NoCancelInvLbl; NoCancelInvLbl)
            {
            }
            column(NetNoInvLbl; NetNoInvLbl)
            {
            }
            column(Declare3Lbl; Declare3Lbl)
            {
            }
            column(PlaceLbl; PlaceLbl)
            {
            }
            column(Date_Lbl; Date_Lbl)
            {
            }
            column(SignatureLbl; SignatureLbl)
            {
            }
            column(InstructionLbl; InstructionLbl)
            {
            }
            column(TermsLbl; TermsLbl)
            {
            }
            column(TermsGSTLbl; TermsGSTLbl)
            {
            }
            column(TermsUINLbl; TermsUINLbl)
            {
            }
            column(TermsHSNLbl; TermsHSNLbl)
            {
            }
            column(TermsSACLbl; TermsSACLbl)
            {
            }
            column(TermsPOSLbl; TermsPOSLbl)
            {
            }
            column(Terms2Lbl; Terms2Lbl)
            {
            }
            column(Terms3Lbl; Terms3Lbl)
            {
            }
            column(Terms4Lbl; Terms4Lbl)
            {
            }
            column(no; no)
            {
            }
            column(GSTIN; GSTIN)
            {
            }
            column(TaxablePerson; TaxablePerson)
            {
            }
            column(Year; Year)
            {
            }
            column(Month; FORMAT(Period))
            {
            }
            column(Period; Period)
            {
            }
            column(GSTComponent_1; GSTComp[1])
            {
            }
            column(GSTComponent_2; GSTComp[2])
            {
            }
            column(GSTComponent_3; GSTComp[3])
            {
            }
            column(GSTComponent_4; GSTComp[4])
            {
            }
            column(GSTComponent_5; GSTComp[5])
            {
            }
            column(Declare2Lbl; Declare2Lbl)
            {
            }
            column(PersonName; PersonName)
            {
            }
            column(AggregateTurnover; AggregateTurnover)
            {
            }
            column(Place; Place)
            {
            }
            column(Date; FORMAT(Date))
            {
            }

            trigger OnPreDataItem()
            begin
                IF GSTIN = '' THEN
                    ERROR(GSTINErr);
                IF TaxablePerson = '' THEN
                    ERROR(PersonNameErr);
                IF AggregateTurnover = 0 THEN
                    ERROR(AggregateTurnoverErr);
                IF PersonName = '' THEN
                    ERROR(AuthorizedPersonErr);
                IF Place = '' THEN
                    ERROR(PlaceErr);
                IF Date = 0D THEN
                    ERROR(DateErr);
                i := 1;
                /* GSTComponent.SETCURRENTKEY("Report View");
                GSTComponent.SETFILTER("Report View", '<>%1', 0);
                GSTComponent.ASCENDING(TRUE);
                IF GSTComponent.FINDSET THEN BEGIN
                    REPEAT
                        GSTComp[i] := GSTComponent.Code;
                        i += 1;
                    UNTIL GSTComponent.NEXT = 0;
                END ELSE
                    ERROR(ReportViewErr); */
            end;
        }
        dataitem("Detailed GST Ledger Entry"; "Detailed GST Ledger Entry")
        {
            column(DocumentLineNo; "Document Line No.")
            {
            }
            column(BuyerSellerRegNo_DetailedGSTLedgerEntry; BuyerSellerRegNo[1])
            {
            }
            column(DocumentNo_DetailedGSTLedgerEntry; DocNo[1])
            {
            }
            column(PostingDate_DetailedGSTLedgerEntry; FORMAT(PostingDate[1]))
            {
            }
            column(InvoiceValue; ABS(DummyInvoiceValue[1]))
            {
            }
            column(Goods_Service_Value; GoodService[1])
            {
            }
            column(HSNSACCode_DetailedGSTLedgerEntry; HSNCode[1])
            {
            }
            column(GSTBaseAmount_DetailedGSTLedgerEntry; ABS(GSTBaseAmt[1]))
            {
            }
            column(GSTCompRate_1; GSTCompRate[1])
            {
            }
            column(GSTCompRate_2; GSTCompRate[2])
            {
            }
            column(GSTCompRate_3; GSTCompRate[3])
            {
            }
            column(GSTCompAmt_1; ABS(GSTCompAmt[1]))
            {
            }
            column(GSTCompAmt_2; ABS(GSTCompAmt[2]))
            {
            }
            column(GSTCompAmt_3; ABS(GSTCompAmt[3]))
            {
            }
            column(Pos_1; Pos[1])
            {
            }
            column(Ecom_1; Ecom[1])
            {
            }
            column(EcomMerchant_1; EcomMerchant[1])
            {
            }
            column(GSTCompRate_4; GSTCompRate[4])
            {
            }
            column(GSTCompRate_5; GSTCompRate[5])
            {
            }
            column(GSTCompRate_6; GSTCompRate[6])
            {
            }
            column(GSTCompAmt_4; ABS(GSTCompAmt[4]))
            {
            }
            column(GSTCompAmt_5; ABS(GSTCompAmt[5]))
            {
            }
            column(GSTCompAmt_6; ABS(GSTCompAmt[6]))
            {
            }
            column(Pos_2; Pos[2])
            {
            }
            column(BuyerSellerRegNo_2; BuyerSellerRegNo[2])
            {
            }
            column(DocNo_2; DocNo[2])
            {
            }
            column(PostingDate_2; FORMAT(PostingDate[2]))
            {
            }
            column(HSNCode_2; HSNCode[2])
            {
            }
            column(GSTBaseAmt_2; ABS(GSTBaseAmt[2]))
            {
            }
            column(InvoiceValue_2; ABS(DummyInvoiceValue[2]))
            {
            }
            column(GoodService_2; GoodService[2])
            {
            }
            column(BuyerSellerStateCode_2; BuyerSellerStateCode[2])
            {
            }
            column(CustomerName; CustomerName)
            {
            }
            column(Invoice_Amt; ABS(DummyInvoiceAmt))
            {
            }
            column(Amt; Amt)
            {
            }
            column(HSNCode_3; HSNCode[3])
            {
            }
            column(GoodService_3; GoodService[3])
            {
            }
            column(GSTBaseAmt_3; ABS(GSTBaseAmt[3]))
            {
            }
            column(DocNo_8; DocNo[8])
            {
            }
            column(Pos_3; Pos[3])
            {
            }
            column(GSTCompRate_7; GSTCompRate[7])
            {
            }
            column(GSTCompRate_8; GSTCompRate[8])
            {
            }
            column(GSTCompRate_9; GSTCompRate[9])
            {
            }
            column(GSTCompAmt_7; ABS(GSTCompAmt[7]))
            {
            }
            column(GSTCompAmt_8; ABS(GSTCompAmt[8]))
            {
            }
            column(GSTCompAmt_9; ABS(GSTCompAmt[9]))
            {
            }
            column(GST_Amount; ABS(DummyGSTAmount))
            {
            }
            column(GoodSerDescription_1; GoodSerDescription[1])
            {
            }
            column(NilRated_1; DummyNilRated[1])
            {
            }
            column(ExcemptedAmt_1; ABS(ExcemptedAmt[1]))
            {
            }
            column(NonGSTSupplies_1; NonGSTSupplies[1])
            {
            }
            column(GoodSerDescription_2; GoodSerDescription[2])
            {
            }
            column(NilRated_2; DummyNilRated[2])
            {
            }
            column(ExcemptedAmt_2; ABS(ExcemptedAmt[2]))
            {
            }
            column(NonGSTSupplies_2; NonGSTSupplies[2])
            {
            }
            column(GoodSerDescription_3; GoodSerDescription[3])
            {
            }
            column(NilRated_3; DummyNilRated[3])
            {
            }
            column(ExcemptedAmt_3; ABS(ExcemptedAmt[3]))
            {
            }
            column(NonGSTSupplies_3; NonGSTSupplies[3])
            {
            }
            column(GoodSerDescription_4; GoodSerDescription[4])
            {
            }
            column(NilRated_4; DummyNilRated[4])
            {
            }
            column(ItemNo_1; ItemNo[1])
            {
            }
            column(ItemNo_2; ItemNo[2])
            {
            }
            column(ItemNo_3; ItemNo[3])
            {
            }
            column(ItemNo_4; ItemNo[4])
            {
            }
            column(ExcemptedAmt_4; ABS(ExcemptedAmt[4]))
            {
            }
            column(NonGSTSupplies_4; NonGSTSupplies[4])
            {
            }
            column(GSTCompRate_10; GSTCompRate[10])
            {
            }
            column(GSTCompRate_11; GSTCompRate[11])
            {
            }
            column(GSTCompRate_12; GSTCompRate[12])
            {
            }
            column(GSTCompAmt_10; ABS(GSTCompAmt[10]))
            {
            }
            column(GSTCompAmt_11; ABS(GSTCompAmt[11]))
            {
            }
            column(GSTCompAmt_12; ABS(GSTCompAmt[12]))
            {
            }
            column(BuyerSellerRegNo_3; BuyerSellerRegNo[3])
            {
            }
            column(OriginalInvNo; OriginalInvNo)
            {
            }
            column(OriginalInvPostDate; OriginalInvPostDate)
            {
            }
            column(DocNo_3; DocNo[3])
            {
            }
            column(GoodService_4; GoodService[4])
            {
            }
            column(PostingDate_3; FORMAT(PostingDate[3]))
            {
            }
            column(DocumentType_1; DocumentType[1])
            {
            }
            column(GSTBaseAmt_4; ABS(GSTBaseAmt[4]))
            {
            }
            column(DocNo_4; DocNo[4])
            {
            }
            column(ExportNo_1; ExportNo[1])
            {
            }
            column(ExportDate_1; ExportDate[1])
            {
            }
            column(PostingDate_4; FORMAT(PostingDate[4]))
            {
            }
            column(GSTBaseAmt_5; ABS(GSTBaseAmt[5]))
            {
            }
            column(HSNCode_4; HSNCode[4])
            {
            }
            column(GSTCompRate_13; GSTCompRate[13])
            {
            }
            column(GSTCompRate_14; GSTCompRate[14])
            {
            }
            column(GSTCompRate_15; GSTCompRate[15])
            {
            }
            column(GSTCompAmt_13; ABS(GSTCompAmt[13]))
            {
            }
            column(GSTCompAmt_14; ABS(GSTCompAmt[14]))
            {
            }
            column(GSTCompAmt_15; ABS(GSTCompAmt[15]))
            {
            }
            column(BuyerSellerRegNo_4; BuyerSellerRegNo[4])
            {
            }
            column(DocNo_6; DocNo[6])
            {
            }
            column(ExportNo_2; ExportNo[2])
            {
            }
            column(ExportDate_2; ExportDate[2])
            {
            }
            column(PostingDate_6; FORMAT(PostingDate[6]))
            {
            }
            column(GSTBaseAmt_7; ABS(GSTBaseAmt[7]))
            {
            }
            column(HSNCode_6; HSNCode[6])
            {
            }
            column(GoodService_5; GoodService[5])
            {
            }
            column(GSTCompRate_19; GSTCompRate[19])
            {
            }
            column(GSTCompRate_20; GSTCompRate[20])
            {
            }
            column(GSTCompRate_21; GSTCompRate[21])
            {
            }
            column(GSTCompAmt_19; ABS(GSTCompAmt[19]))
            {
            }
            column(GSTCompAmt_20; ABS(GSTCompAmt[20]))
            {
            }
            column(GSTCompAmt_21; ABS(GSTCompAmt[21]))
            {
            }
            column(Pos_4; Pos[4])
            {
            }
            column(DocNo_7; DocNo[7])
            {
            }
            column(GSTCompRate_22; GSTCompRate[22])
            {
            }
            column(GSTCompRate_23; GSTCompRate[23])
            {
            }
            column(GSTCompRate_24; GSTCompRate[24])
            {
            }
            column(GSTCompAmt_22; ABS(GSTCompAmt[22]))
            {
            }
            column(GSTCompAmt_23; ABS(GSTCompAmt[23]))
            {
            }
            column(GSTCompAmt_24; ABS(GSTCompAmt[24]))
            {
            }
            column(TransactionNo; TransactionNo)
            {
            }
            column(NoSeries; NoSeries)
            {
            }
            column(StartingNo; StartingNo)
            {
            }
            column(EndingNo; EndingNo)
            {
            }
            column(RecordCount; RecordCount)
            {
            }
            column(GSTCompRate_16; GSTCompRate[16])
            {
            }
            column(GSTCompRate_17; GSTCompRate[17])
            {
            }
            column(GSTCompRate_18; GSTCompRate[18])
            {
            }
            column(GSTCompAmt_16; ABS(GSTCompAmt[16]))
            {
            }
            column(GSTCompAmt_17; ABS(GSTCompAmt[17]))
            {
            }
            column(GSTCompAmt_18; ABS(GSTCompAmt[18]))
            {
            }
            column(NilRatedText_1; NilRatedText[1])
            {
            }
            column(NilRatedText_2; NilRatedText[2])
            {
            }
            column(NilRatedText_3; NilRatedText[3])
            {
            }
            column(NilRatedText_4; NilRatedText[4])
            {
            }
            column(WithoutPaymentText; WithoutPaymentText)
            {
            }
            column(WithPaymentText; WithPaymentText)
            {
            }
            column(BuyerSellerRegNo_5; BuyerSellerRegNo[5])
            {
            }
            column(DocNo_9; DocNo[9])
            {
            }
            column(DocNo_10; DocNo[10])
            {
            }
            column(DocDate; DocDate)
            {
            }
            column(Pos_5; Pos[5])
            {
            }
            column(PostingDate_7; PostingDate[7])
            {
            }
            column(GoodService_6; GoodService[6])
            {
            }
            column(HSNCode_7; HSNCode[7])
            {
            }
            column(GSTBaseAmt_9; ABS(GSTBaseAmt[9]))
            {
            }
            column(GSTCompRate_25; GSTCompRate[25])
            {
            }
            column(GSTCompRate_26; GSTCompRate[26])
            {
            }
            column(GSTCompRate_27; GSTCompRate[27])
            {
            }
            column(GSTCompAmt_25; ABS(GSTCompAmt[25]))
            {
            }
            column(GSTCompAmt_26; ABS(GSTCompAmt[26]))
            {
            }
            column(GSTCompAmt_27; ABS(GSTCompAmt[27]))
            {
            }
            column(CrossValAmt; CrossValAmt)
            {
            }
            column(SNo; SNo)
            {
            }
            column(Ecom_2; Ecom[2])
            {
            }
            column(EcomMerchant_2; EcomMerchant[2])
            {
            }
            column(Pos_6; Pos[6])
            {
            }
            column(GSTBaseAmt_10; GSTBaseAmt[10])
            {
            }
            column(GSTCompRate_28; GSTCompRate[28])
            {
            }
            column(GSTCompRate_29; GSTCompRate[29])
            {
            }
            column(GSTCompRate_30; GSTCompRate[30])
            {
            }
            column(GSTCompAmt_28; ABS(GSTCompAmt[28]))
            {
            }
            column(GSTCompAmt_29; ABS(GSTCompAmt[29]))
            {
            }
            column(GSTCompAmt_30; ABS(GSTCompAmt[30]))
            {
            }
            column(Component_1; Component[1])
            {
            }
            column(DocNo_11; DocNo[11])
            {
            }
            column(PostingDate_11; PostingDate[11])
            {
            }
            column(HSNCode_11; HSNCode[11])
            {
            }
            column(GSTBaseAmt_11; ABS(GSTBaseAmt[11]))
            {
            }
            column(HSNCode_12; HSNCode[12])
            {
            }
            column(HSNCode_13; HSNCode[13])
            {
            }
            column(HSNCode_14; HSNCode[14])
            {
            }
            column(StartingNo1; StartingNo1)
            {
            }
            column(EndingNo1; EndingNo1)
            {
            }
            column(NoSeries1; NoSeries1)
            {
            }
            column(RecordCount1; RecordCount1)
            {
            }

            trigger OnAfterGetRecord()
            begin
                ClearData;
                TaxableB2B(
                  "Detailed GST Ledger Entry", GSTCompRate[1], GSTCompRate[2], GSTCompRate[3], GSTCompAmt[1], GSTCompAmt[2],
                  GSTCompAmt[3], GSTComp[1], GSTComp[2], GSTComp[3]);
                TaxableB2C(
                  "Detailed GST Ledger Entry", GSTCompRate[4], GSTCompRate[5],
                  GSTCompRate[6], GSTCompAmt[4], GSTCompAmt[5], GSTCompAmt[6], GSTComp[1], GSTComp[2], GSTComp[3]);
                TaxableB2CWithLimit(
                  "Detailed GST Ledger Entry", GSTCompRate[7], GSTCompRate[8], GSTCompRate[9],
                  GSTCompAmt[7], GSTCompAmt[8], GSTCompAmt[9], GSTComp[1], GSTComp[2], GSTComp[3]);
                //IF "Nature of Supply" = "Nature of Supply"::B2B THEN
                IF "GST Jurisdiction Type" = "GST Jurisdiction Type"::Interstate THEN BEGIN
                    NilRatedText[1] := InterSupplyPerLbl;
                    NilRatedExempted(
                      "Detailed GST Ledger Entry", GoodSerDescription[1], ExcemptedAmt[1], ItemNo[1]);
                    NilRatedNonGST(
                      "Detailed GST Ledger Entry", GoodSerDescription[1], NonGSTSupplies[1], ItemNo[1]);
                END ELSE
                    IF "GST Jurisdiction Type" = "GST Jurisdiction Type"::Intrastate THEN BEGIN
                        NilRatedText[2] := IntraSupplyPerLbl;
                        NilRatedExempted(
                          "Detailed GST Ledger Entry", GoodSerDescription[2], ExcemptedAmt[2], ItemNo[2]);
                        NilRatedNonGST(
                          "Detailed GST Ledger Entry", GoodSerDescription[2], NonGSTSupplies[2], ItemNo[2]);
                    END;
                //IF "Nature of Supply" = "Nature of Supply"::B2C THEN
                IF "GST Jurisdiction Type" = "GST Jurisdiction Type"::Interstate THEN BEGIN
                    NilRatedText[3] := InterSupplyConsLbl;
                    NilRatedExempted(
                      "Detailed GST Ledger Entry", GoodSerDescription[3], ExcemptedAmt[3], ItemNo[3]);
                    NilRatedNonGST(
                      "Detailed GST Ledger Entry", GoodSerDescription[3], NonGSTSupplies[3], ItemNo[3]);
                END ELSE
                    IF "GST Jurisdiction Type" = "GST Jurisdiction Type"::Intrastate THEN BEGIN
                        NilRatedText[4] := IntraSupplyConsLbl;
                        NilRatedExempted(
                          "Detailed GST Ledger Entry", GoodSerDescription[4], ExcemptedAmt[4], ItemNo[4]);
                        NilRatedNonGST(
                          "Detailed GST Ledger Entry", GoodSerDescription[4], NonGSTSupplies[4], ItemNo[4]);
                    END;
                IF ("Document Type" = "Document Type"::"Credit Memo")/* AND
                   ("Invoice Type" IN ["Invoice Type"::Supplementary, "Invoice Type"::"Debit Note"])*/
                THEN
                    DetailsCreditAndDebit(
                      "Detailed GST Ledger Entry", GSTCompRate[10], GSTCompRate[11], GSTCompRate[12],
                      GSTCompAmt[10], GSTCompAmt[11], GSTCompAmt[12], GSTComp[1], GSTComp[2], GSTComp[3]);
                IF "GST Without Payment of Duty" THEN BEGIN
                    WithoutPaymentText := WithoutPaymentGSTLbl;
                    SuppliesExported(
                      "Detailed GST Ledger Entry", GSTCompRate[13], GSTCompRate[14], GSTCompRate[15],
                      GSTCompAmt[13], GSTCompAmt[14], GSTCompAmt[15],
                      DocNo[4], PostingDate[4], HSNCode[4], GSTComp[1], GSTComp[2], GSTComp[3], ExportNo[1], ExportDate[1])
                END ELSE BEGIN
                    WithPaymentText := WithPaymentGSTLbl;
                    SuppliesExported(
                      "Detailed GST Ledger Entry", GSTCompRate[16], GSTCompRate[17],
                      GSTCompRate[18], GSTCompAmt[16], GSTCompAmt[17], GSTCompAmt[18],
                      DocNo[6], PostingDate[6], HSNCode[6], GSTComp[1], GSTComp[2], GSTComp[3], ExportNo[2], ExportDate[2]);
                END;
                IF ("Entry Type" = "Entry Type"::"Initial Entry") AND
                   ("Document Type" = "Document Type"::Payment) AND "GST on Advance Payment" AND NOT Paid
                THEN BEGIN
                    //IF "Nature of Supply" = "Nature of Supply"::B2C THEN BEGIN
                    DummyCustomer.GET("Source No.");
                    BuyerSellerRegNo[4] := COPYSTR(DummyCustomer.Name, 1, 15);
                    //END ELSE
                    BuyerSellerRegNo[4] := FORMAT("Buyer/Seller Reg. No.");
                    AdvancePaymentAfterPosting(
                      "Detailed GST Ledger Entry", GSTCompRate[19], GSTCompRate[20], GSTCompRate[21],
                      GSTCompAmt[19], GSTCompAmt[20], GSTCompAmt[21], GSTComp[1], GSTComp[2], GSTComp[3]);
                END;
                IF ("Entry Type" = "Entry Type"::Application) AND ("Transaction Type" = "Transaction Type"::Sales) AND
                   ("Document Type" = "Document Type"::Invoice) AND Paid
                THEN
                    AdvancePaymentBeforePosting(
                      "Detailed GST Ledger Entry", GSTCompRate[22], GSTCompRate[23], GSTCompRate[24],
                      GSTCompAmt[22], GSTCompAmt[23], GSTCompAmt[24], GSTComp[1], GSTComp[2], GSTComp[3]);
                AdvancePaymentAdjustmentJnl(
                  "Detailed GST Ledger Entry", GSTCompRate[25], GSTCompRate[26], GSTCompRate[27],
                  GSTCompAmt[25], GSTCompAmt[26], GSTCompAmt[27], GSTComp[1], GSTComp[2], GSTComp[3]);
                SuppliesUnregCustThroughECom(
                  "Detailed GST Ledger Entry", GSTCompRate[28], GSTCompRate[29], GSTCompRate[30],
                  GSTCompAmt[28], GSTCompAmt[29], GSTCompAmt[30], GSTComp[1], GSTComp[2], GSTComp[3]);
            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("Location  Reg. No.", "Source Type");
                SETRANGE("Location  Reg. No.", GSTIN);
                //SETRANGE("Source Type","Source Type"::Customer);
                SETFILTER("Source Type", '%1|%2', "Source Type"::Customer, "Source Type"::" ");//07Nov2017
                SETFILTER("HSN/SAC Code", '<>%1', '');//07Nov2017
                                                      // SETFILTER("Original Doc. Type", '<>%1', "Original Doc. Type"::"Transfer Receipt");//08Nov2017
                CASE Period OF
                    Period::Jan:
                        SETRANGE("Posting Date", DMY2DATE(1, 1, Year), DMY2DATE(31, 1, Year));
                    Period::Feb:
                        SETRANGE("Posting Date", DMY2DATE(1, 2, Year), DMY2DATE(28, 2, Year));
                    Period::March:
                        SETRANGE("Posting Date", DMY2DATE(1, 3, Year), DMY2DATE(31, 3, Year));
                    Period::Apr:
                        SETRANGE("Posting Date", DMY2DATE(1, 4, Year), DMY2DATE(30, 4, Year));
                    Period::May:
                        SETRANGE("Posting Date", DMY2DATE(1, 5, Year), DMY2DATE(31, 5, Year));
                    Period::June:
                        SETRANGE("Posting Date", DMY2DATE(1, 6, Year), DMY2DATE(30, 6, Year));
                    Period::July:
                        SETRANGE("Posting Date", DMY2DATE(1, 7, Year), DMY2DATE(31, 7, Year));
                    Period::Agu:
                        SETRANGE("Posting Date", DMY2DATE(1, 8, Year), DMY2DATE(31, 8, Year));
                    Period::Sept:
                        SETRANGE("Posting Date", DMY2DATE(1, 9, Year), DMY2DATE(30, 9, Year));
                    Period::Oct:
                        SETRANGE("Posting Date", DMY2DATE(1, 10, Year), DMY2DATE(31, 10, Year));
                    Period::Nov:
                        //SETRANGE("Posting Date",DMY2DATE(1,11,Year),DMY2DATE(30,3,Year));
                        SETRANGE("Posting Date", DMY2DATE(1, 11, Year), DMY2DATE(30, 11, Year));//Bug in GSTR-1 08Jan2018
                    Period::Dec:
                        SETRANGE("Posting Date", DMY2DATE(1, 12, Year), DMY2DATE(31, 12, Year));
                END;
                GetNoSeries;
                GetPurchNoSeries;
            end;
        }
        dataitem(GSTInvoice; "Detailed GST Ledger Entry")
        {
            //The property 'DataItemTableView' shouldn't have an empty value.
            //DataItemTableView = '';
            column(BuyerSellerRegNo_GSTInvoice; "Buyer/Seller Reg. No.")
            {
            }
            column(DocumentLineNo_GSTInvoice; GSTInvoice."Document Line No.")
            {
            }
            column(GSTBaseAmount_GSTInvoice; GSTInvoice."GST Base Amount")
            {
            }
            column(GST_GSTInvoice; GSTInvoice."GST %")
            {
            }
            column(GSTAmount_GSTInvoice; GSTInvoice."GST Amount")
            {
            }
            column(DocumentNo_GSTInvoice; GSTInvoice."Document No.")
            {
            }
            column(PostingDate_GSTInvoice; GSTInvoice."Posting Date")
            {
            }
            column(No_GSTInvoice; GSTInvoice."No.")
            {
            }
            column(HSNCode_GSTInvoice; GSTInvoice."HSN/SAC Code")
            {
            }
            column(ItmDescription; ItmDesc)
            {
            }
            column(CGST17; CGST17)
            {
            }
            column(SGST17; SGST17)
            {
            }
            column(IGST17; IGST17)
            {
            }
            column(UTGST17; UTGST17)
            {
            }
            column(CESS17; CESS17)
            {
            }
            column(CGSTP; CGSTP)
            {
            }
            column(SGSTP; SGSTP)
            {
            }
            column(IGSTP; IGSTP)
            {
            }
            column(UTGSTP; UTGSTP)
            {
            }
            column(InvDD; InvDD)
            {
            }
            column(PSupply; PSupply)
            {
            }
            column(GSTBaseAmt22; GSTBaseAmt22)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //15Sep2017
                IF InvoiceDetail THEN
                    InvDD := 3
                ELSE
                    InvDD := 1;

                IF NOT InvoiceDetail THEN
                    CurrReport.SKIP;


                //>>RB-N 14Sep2017 Item Description
                CLEAR(ItmDesc);
                CLEAR(ItmUom);//RB-N 22Sep2017
                CLEAR(ItmQtyUOM);
                IF Type = Type::Item THEN BEGIN
                    Itm14.RESET;
                    IF Itm14.GET("No.") THEN BEGIN
                        ItmDesc := Itm14.Description;
                        ItmUom := Itm14."Base Unit of Measure";//RB-N 22Sep2017

                        ItmUOM22.RESET;
                        ItmUOM22.SETRANGE("Item No.", Itm14."No.");
                        ItmUOM22.SETRANGE(Code, Itm14."Sales Unit of Measure");
                        IF ItmUOM22.FINDFIRST THEN BEGIN

                            ItmQtyUOM := ItmUOM22."Qty. per Unit of Measure";

                        END;
                    END;

                END;

                //>>DetailedGST
                CLEAR(CGST17);
                CLEAR(SGST17);
                CLEAR(IGST17);
                CLEAR(UTGST17);
                CLEAR(CESS17);
                CLEAR(CGSTP);
                CLEAR(SGSTP);
                CLEAR(IGSTP);
                CLEAR(UTGSTP);

                DGST.RESET;
                DGST.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DGST.SETRANGE("Document No.", "Document No.");
                DGST.SETRANGE("GST %", "GST %");//13Nov2017
                //DGST.SETRANGE("Document Line No.","Document Line No."); //RB-N 22Sep2017
                IF DGST.FINDSET THEN
                    REPEAT

                        IF DGST."GST Component Code" = 'CGST' THEN BEGIN
                            CGST17 += DGST."GST Amount";
                            CGSTP := DGST."GST %";
                        END;

                        IF DGST."GST Component Code" = 'SGST' THEN BEGIN
                            SGST17 += DGST."GST Amount";
                            SGSTP := DGST."GST %";
                        END;

                        IF DGST."GST Component Code" = 'IGST' THEN BEGIN
                            IGST17 += DGST."GST Amount";
                            IGSTP := DGST."GST %";
                        END;

                        IF DGST."GST Component Code" = 'UTGST' THEN BEGIN
                            UTGST17 += DGST."GST Amount";
                            UTGSTP := DGST."GST %";
                        END;

                        IF DGST."GST Component Code" = 'CESS' THEN BEGIN
                            CESS17 += DGST."GST Amount";
                        END;

                    UNTIL DGST.NEXT = 0;


                //>> RB-N 22Sep2017 PlaceofSupply
                CLEAR(PSupply);
                State22.RESET;
                /*  IF State22.GET(GSTInvoice."Buyer/Seller State Code") THEN BEGIN
                     PSupply := State22."State Code (GST Reg. No.)" + ' - ' + State22.Description;

                 END;
  */
                //>>RB-N 22Sep2017 PlaceofSupply


                //>>RB-N 22Sep2017 TaxBaseAmt
                /*
                CLEAR(GSTBaseAmt22);
                
                SIL22.RESET;
                SIL22.SETRANGE("Document No.","Document No.");
                SIL22.SETRANGE(Type,SIL22.Type::Item);
                IF SIL22.FINDSET THEN
                REPEAT
                  GSTBaseAmt22 += SIL22."GST Base Amount";
                
                UNTIL SIL22.NEXT = 0;
                */
                //<<RB-N

                GSTBaseAmt22 := 0; //13Nov2017
                NNN111 += 1;
                //>>RB-N 07Nov2017 GSTBaseAmount

                DGST18.RESET;
                DGST18.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DGST18.SETRANGE("Document No.", "Document No.");
                DGST18.SETRANGE("Document Line No.", "Document Line No.");
                IF DGST18.FINDFIRST THEN BEGIN
                    NNN222 := DGST18.COUNT;
                END;

                IF NNN111 = NNN222 THEN BEGIN

                    GSTBaseAmt22 := ABS("GST Base Amount");
                    NNN111 := 0;

                END;
                //>>RB-N 07Nov2017 GSTBaseAmount

            end;

            trigger OnPreDataItem()
            begin

                //>>RB-N 14Sep2017
                //SETCURRENTKEY("Location  Reg. No.","Source Type");
                SETCURRENTKEY("Location  Reg. No.", "Source Type", "Document No.", "GST %");
                SETRANGE("Location  Reg. No.", GSTIN);
                //SETRANGE("Source Type","Source Type"::Customer);
                SETFILTER("Source Type", '%1|%2', "Source Type"::Customer, "Source Type"::" ");//07Nov2017
                SETFILTER("HSN/SAC Code", '<>%1', '');//07Nov2017
                //SETFILTER("Original Doc. Type", '<>%1', "Original Doc. Type"::"Transfer Receipt");//08Nov2017
                SETRANGE("Posting Date", SDate22, EDate22);//07Nov2017
                //>>RB-N 14Sep2017
            end;
        }
        dataitem(GSTHSN; "Detailed GST Ledger Entry")
        {
            column(EntryNo_GSTHSN; GSTHSN."Entry No.")
            {
            }
            column(EntryType_GSTHSN; GSTHSN."Entry Type")
            {
            }
            column(TransactionType_GSTHSN; GSTHSN."Transaction Type")
            {
            }
            column(DocumentType_GSTHSN; GSTHSN."Document Type")
            {
            }
            column(DocumentNo_GSTHSN; GSTHSN."Document No.")
            {
            }
            column(PostingDate_GSTHSN; GSTHSN."Posting Date")
            {
            }
            column(HSNCode_GSTHSN; "HSN/SAC Code")
            {
            }
            column(Type_GSTHSN; GSTHSN.Type)
            {
            }
            column(No_GSTHSN; GSTHSN."No.")
            {
            }
            column(GSTBaseAmt2211; GSTBaseAmt2211)
            {
            }
            column(QtyBase2211; QtyBase2211)
            {
            }
            column(ItmUom; ItmUom)
            {
            }
            column(HSNDesc; HSNDesc)
            {
            }
            column(SGST1711; SGST1711)
            {
            }
            column(CGST1711; CGST1711)
            {
            }
            column(IGST1711; IGST1711)
            {
            }

            trigger OnAfterGetRecord()
            begin

                CGST1711 := 0;
                SGST1711 := 0;
                IGST1711 := 0;
                QtyBase2211 := 0;
                GSTBaseAmt2211 := 0;

                IF "GST Component Code" = 'SGST' THEN BEGIN
                    SGST1711 := ABS("GST Amount");
                END;

                IF "GST Component Code" = 'CGST' THEN BEGIN
                    CGST1711 := ABS("GST Amount");
                END;

                IF "GST Component Code" = 'IGST' THEN BEGIN
                    IGST1711 := ABS("GST Amount");
                END;

                IF "GST Component Code" = 'UTGST' THEN BEGIN
                    SGST1711 := ABS("GST Amount");
                END;

                //>>RB-N 22sep2017 HSNDescription

                CLEAR(HSNDesc);
                HSN22.RESET;
                HSN22.SETRANGE(Code, "HSN/SAC Code");
                IF HSN22.FINDFIRST THEN BEGIN
                    HSNDesc := HSN22.Description;

                END;
                //<<


                //>>RB-N 22Sep2017
                CLEAR(ItmQtyUOM);
                CLEAR(ItmUom);
                IF Type = Type::Item THEN BEGIN
                    Itm14.RESET;
                    IF Itm14.GET("No.") THEN BEGIN

                        ItmUom := Itm14."Base Unit of Measure";

                        ItmUOM22.RESET;
                        ItmUOM22.SETRANGE("Item No.", Itm14."No.");
                        ItmUOM22.SETRANGE(Code, Itm14."Sales Unit of Measure");
                        IF ItmUOM22.FINDFIRST THEN BEGIN

                            ItmQtyUOM := ItmUOM22."Qty. per Unit of Measure";

                        END;
                    END;

                END;
                //<<

                //>>RB-N 07Nov2017
                IF Type <> Type::Item THEN BEGIN
                    ItmQtyUOM := 1;
                    ItmUom := 'NOS';

                END;
                //>>RB-N 07Nov2017

                TTT += 1;
                //>>Document No.

                DGST18.RESET;
                DGST18.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DGST18.SETRANGE("Document No.", "Document No.");
                DGST18.SETRANGE("Document Line No.", "Document Line No.");
                IF DGST18.FINDFIRST THEN BEGIN
                    III := DGST18.COUNT;
                END;

                IF TTT = III THEN BEGIN

                    QtyBase2211 := ABS(Quantity) * ItmQtyUOM;
                    GSTBaseAmt2211 := ABS("GST Base Amount");

                    TTT := 0;

                END;
                //>>Document No.

                //>>RB-N 08Nov2017 ExportInvoice
                ExpInvoice := FALSE;
                IF TempHSN <> "HSN/SAC Code" THEN BEGIN
                    TempHSN := "HSN/SAC Code";

                    Ex08.SETRANGE("Cell Value as Text", TempHSN);
                    IF NOT Ex08.FINDFIRST THEN BEGIN
                        H08 += 1;
                        Ex08.INIT;
                        Ex08."Row No." := H08;
                        Ex08."Column No." := 1;
                        Ex08."Cell Value as Text" := "HSN/SAC Code";
                        Ex08.INSERT;

                        ExpInvoice := TRUE;

                    END;
                END;

                IF ExpInvoice THEN BEGIN

                    SIL08.RESET;
                    SIL08.SETRANGE("HSN/SAC Code", "HSN/SAC Code");
                    SIL08.SETRANGE("Posting Date", SDate22, EDate22);
                    SIL08.SETRANGE("Gen. Bus. Posting Group", 'FOREIGN');
                    SIL08.SETRANGE("Invoice Type", SIL08."Invoice Type"::Export);
                    IF SIL08.FINDSET THEN
                        REPEAT

                            Loc08.RESET;
                            Loc08.SETRANGE(Code, SIL08."Location Code");
                            IF Loc08.FINDFIRST THEN BEGIN

                                IF Loc08."GST Registration No." <> GSTIN THEN
                                    CurrReport.SKIP;

                            END;

                            QtyBase2211 += SIL08."Quantity (Base)";

                            SIH08.RESET;
                            SIH08.SETRANGE("No.", SIL08."Document No.");
                            IF SIH08.FINDFIRST THEN BEGIN
                                IF SIH08."Currency Factor" <> 0 THEN
                                    GSTBaseAmt2211 += 0// SIL08."Amount To Customer" / SIH08."Currency Factor"
                                ELSE
                                    GSTBaseAmt2211 += 0;// SIL08."Amount To Customer";
                            END;
                        //MESSAGE('%1 Hsn No. \ %2 Doc No \ %3 Item No. ',SIL08."HSN/SAC Code",SIL08."Document No.",SIL08."No.");

                        UNTIL SIL08.NEXT = 0;

                END;
                //>>RB-N 08Nov2017 ExportInvoice

                /*
                //>>HSN Grouping
                TTT111 += 1;
                DGST18.RESET;
                DGST18.SETCURRENTKEY("Location  Reg. No.","Source Type","HSN/SAC Code");
                DGST18.SETRANGE("Location  Reg. No.",GSTIN);
                DGST18.SETRANGE("Source Type",DGST18."Source Type"::Customer);
                DGST18.SETRANGE("Posting Date",SDate22,EDate22);
                DGST18.SETRANGE("HSN/SAC Code","HSN/SAC Code");
                IF DGST18.FINDFIRST THEN
                BEGIN
                
                  III111 := DGST18.COUNT;
                
                END;
                
                
                IF TTT111 = III111 THEN
                BEGIN
                  {
                  //MESSAGE(' %1 \ %2 \ %3  \ %4',III111,GSTBaseAmt2211,QtyBase2211,"HSN/SAC Code");
                  CGST1711 := 0;
                  SGST1711 := 0;
                  IGST1711 := 0;
                  QtyBase2211 := 0;
                  GSTBaseAmt2211 :=0;
                  }
                
                  TTT111 := 0;
                END;
                //<<HSN Grouping
                */

            end;

            trigger OnPreDataItem()
            begin

                //>>RB-N 22Sep2017
                SETCURRENTKEY("Location  Reg. No.", "Source Type", "HSN/SAC Code");
                SETRANGE("Location  Reg. No.", GSTIN);
                //SETRANGE("Source Type","Source Type"::Customer);
                SETFILTER("Source Type", '%1|%2', "Source Type"::Customer, "Source Type"::" ");//07Nov2017
                SETFILTER("HSN/SAC Code", '<>%1', '');//07Nov2017
                //SETFILTER("Original Doc. Type", '<>%1', "Original Doc. Type"::"Transfer Receipt");//08Nov2017
                SETRANGE("Posting Date", SDate22, EDate22);

                //>>RB-N 22Sep2017
            end;
        }
        dataitem("Sales Invoice Line"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.");
            column(DocumentNo_SIL; "Sales Invoice Line"."Document No.")
            {
            }
            column(PostingDate_SIL; "Posting Date")
            {
            }
            column(LineNo_SIL; "Sales Invoice Line"."Line No.")
            {
            }
            column(Type_SIL; "Sales Invoice Line".Type)
            {
            }
            column(No_SIL; "Sales Invoice Line"."No.")
            {
            }
            column(LocationCode_SIL; "Sales Invoice Line"."Location Code")
            {
            }
            column(CGST1708; CGST1708)
            {
            }
            column(SGST1708; SGST1708)
            {
            }
            column(IGST1708; IGST1708)
            {
            }
            column(UTGST1708; UTGST1708)
            {
            }
            column(CESS1708; CESS1708)
            {
            }
            column(CGSTP08; CGSTP08)
            {
            }
            column(SGSTP08; SGSTP08)
            {
            }
            column(IGSTP08; IGSTP08)
            {
            }
            column(UTGSTP08; UTGSTP08)
            {
            }
            column(GSTBaseAmt08; GSTBaseAmt08)
            {
            }

            trigger OnAfterGetRecord()
            begin

                Loc08.RESET;
                Loc08.SETRANGE(Code, "Location Code");
                IF Loc08.FINDFIRST THEN BEGIN

                    IF Loc08."GST Registration No." <> GSTIN THEN
                        CurrReport.SKIP;
                END;

                CLEAR(GSTBaseAmt08);
                SIH08.RESET;
                SIH08.SETRANGE("No.", "Document No.");
                IF SIH08.FINDFIRST THEN BEGIN
                    IF SIH08."Currency Factor" <> 0 THEN
                        GSTBaseAmt08 := 0// "Amount To Customer" / SIH08."Currency Factor"
                    ELSE
                        GSTBaseAmt08 := 0;// "Amount To Customer";
                END;


                //>>DetailedGST
                CLEAR(CGST1708);
                CLEAR(SGST1708);
                CLEAR(IGST1708);
                CLEAR(UTGST1708);
                CLEAR(CESS1708);
                CLEAR(CGSTP08);
                CLEAR(SGSTP08);
                CLEAR(IGSTP08);
                CLEAR(UTGSTP08);

                DGST.RESET;
                DGST.SETCURRENTKEY("Document No.", "Document Line No.", "GST Component Code");
                DGST.SETRANGE("Document No.", "Document No.");
                //DGST.SETRANGE("Document Line No.","Line No.");
                IF DGST.FINDSET THEN
                    REPEAT

                        IF DGST."GST Component Code" = 'CGST' THEN BEGIN
                            CGST1708 += ABS(DGST."GST Amount");
                            CGSTP08 := DGST."GST %";
                        END;

                        IF DGST."GST Component Code" = 'SGST' THEN BEGIN
                            SGST1708 += ABS(DGST."GST Amount");
                            SGSTP08 := DGST."GST %";
                        END;

                        IF DGST."GST Component Code" = 'IGST' THEN BEGIN
                            IGST1708 += ABS(DGST."GST Amount");
                            IGSTP08 := DGST."GST %";
                        END;

                        IF DGST."GST Component Code" = 'UTGST' THEN BEGIN
                            UTGST1708 += ABS(DGST."GST Amount");
                            UTGSTP08 := DGST."GST %";
                        END;

                        IF DGST."GST Component Code" = 'CESS' THEN BEGIN
                            CESS1708 += DGST."GST Amount";
                        END;

                    UNTIL DGST.NEXT = 0;
                //<<DetailGST
            end;

            trigger OnPreDataItem()
            begin

                //>>08Nov2017

                SETRANGE("Posting Date", SDate22, EDate22);
                SETRANGE("Gen. Bus. Posting Group", 'FOREIGN');
                SETRANGE("Invoice Type", "Invoice Type"::Export);
                //<<08Nov2017
            end;
        }
        dataitem(SIHNoS; "Sales Invoice Header")
        {
            DataItemTableView = SORTING("No.");
            column(No_SIH; "No.")
            {
            }
            column(NoSeries_SIH; SIHNoS."No. Series")
            {
            }
            column(H09; H09)
            {
            }
            column(InvStartNo09; InvStartNo09)
            {
            }
            column(InvEndNo09; InvEndNo09)
            {
            }
            column(InvCount09; InvCount09)
            {
            }
            column(CanInv13; CanInv13)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>RB-N 09Nov2017 Invoice NoSeries
                InvNoSeries := FALSE;
                IF TempInvNo <> "No. Series" THEN BEGIN
                    TempInvNo := "No. Series";

                    Ex09.SETRANGE("Cell Value as Text", TempInvNo);
                    IF NOT Ex09.FINDFIRST THEN BEGIN
                        H09 += 1;
                        Ex09.INIT;
                        Ex09."Row No." := H09;
                        Ex09."Column No." := 1;
                        Ex09."Cell Value as Text" := "No. Series";
                        Ex09.INSERT;

                        InvNoSeries := TRUE;

                    END;
                END;

                CLEAR(InvStartNo09);
                CLEAR(InvEndNo09);
                CLEAR(InvCount09);
                CLEAR(CanInv13);//13Nov2017

                IF InvNoSeries THEN BEGIN

                    SIH08.RESET;
                    SIH08.SETCURRENTKEY("Posting Date");//13Nov2017
                    SIH08.SETRANGE("Posting Date", SDate22, EDate22);
                    SIH08.SETRANGE("No. Series", "No. Series");
                    IF SIH08.FINDFIRST THEN BEGIN
                        InvStartNo09 := SIH08."No.";
                        InvCount09 := SIH08.COUNT;
                    END;

                    SIH08.RESET;
                    SIH08.SETCURRENTKEY("Posting Date");//13Nov2017
                    SIH08.SETRANGE("Posting Date", SDate22, EDate22);
                    SIH08.SETRANGE("No. Series", "No. Series");
                    IF SIH08.FINDLAST THEN BEGIN
                        InvEndNo09 := SIH08."No.";
                    END;

                    //13Nov2017 Cancelled Invoice
                    SIH08.RESET;
                    SIH08.SETCURRENTKEY("Posting Date");
                    SIH08.SETRANGE("Posting Date", SDate22, EDate22);
                    SIH08.SETRANGE("No. Series", "No. Series");
                    SIH08.SETRANGE("Cancelled Invoice", TRUE);
                    IF SIH08.FINDFIRST THEN BEGIN

                        CanInv13 := SIH08.COUNT;
                    END;

                END;

                IF NOT InvNoSeries THEN
                    CurrReport.SKIP;

                //>>RB-N 09Nov2017 Invoice NoSeries
            end;

            trigger OnPreDataItem()
            begin

                SETRANGE("Posting Date", SDate22, EDate22);
                IF LocState09 <> '' THEN
                    SETRANGE("Location State Code", LocState09);
            end;
        }
        dataitem(TSHNoS; 5744)
        {
            DataItemTableView = SORTING("No.");
            column(No_TSH; "No.")
            {
            }
            column(NoSeries_TSH; "No. Series")
            {
            }
            column(H09_TSH; H09)
            {
            }
            column(TshStartNo09; TshStartNo09)
            {
            }
            column(TshEndNo09; TshEndNo09)
            {
            }
            column(TshCount09; TshCount09)
            {
            }
            column(CanBrt13; CanBrt13)
            {
            }

            trigger OnAfterGetRecord()
            begin

                //>>RB-N 09Nov2017 Location GSTNo
                Loc09.RESET;
                Loc09.SETRANGE(Code, TSHNoS."Transfer-from Code");
                IF Loc09.FINDFIRST THEN BEGIN
                    IF Loc09."GST Registration No." <> GSTIN THEN
                        CurrReport.SKIP;

                END;
                //>>RB-N 09Nov2017 Location GSTNo


                //>>RB-N 09Nov2017 Transfer Shipment NoSeries
                TSHNoSeries := FALSE;
                IF TempTshNo <> "No. Series" THEN BEGIN
                    TempTshNo := "No. Series";

                    Ex09.SETRANGE("Cell Value as Text", TempTshNo);
                    IF NOT Ex09.FINDFIRST THEN BEGIN
                        H09 += 1;
                        Ex09.INIT;
                        Ex09."Row No." := H09;
                        Ex09."Column No." := 1;
                        Ex09."Cell Value as Text" := "No. Series";
                        Ex09.INSERT;

                        TSHNoSeries := TRUE;

                    END;
                END;

                CLEAR(TshStartNo09);
                CLEAR(TshEndNo09);
                CLEAR(TshCount09);
                CLEAR(CanBrt13);//13Nov2017

                IF TSHNoSeries THEN BEGIN

                    TSH08.RESET;
                    TSH08.SETCURRENTKEY("Posting Date");//13Nov2017
                    TSH08.SETRANGE("Posting Date", SDate22, EDate22);
                    TSH08.SETRANGE("No. Series", "No. Series");
                    IF TSH08.FINDFIRST THEN BEGIN
                        TshStartNo09 := TSH08."No.";
                        TshCount09 := TSH08.COUNT;
                    END;

                    TSH08.RESET;
                    TSH08.SETCURRENTKEY("Posting Date");
                    TSH08.SETRANGE("Posting Date", SDate22, EDate22);
                    TSH08.SETRANGE("No. Series", "No. Series");
                    IF TSH08.FINDLAST THEN BEGIN
                        TshEndNo09 := TSH08."No.";
                    END;

                    //13Nov2017 Cancelled BRT
                    TSH08.RESET;
                    TSH08.SETCURRENTKEY("Posting Date");
                    TSH08.SETRANGE("Posting Date", SDate22, EDate22);
                    TSH08.SETRANGE("No. Series", "No. Series");
                    TSH08.SETRANGE("BRT Cancelled", TRUE);
                    IF TSH08.FINDFIRST THEN BEGIN

                        CanBrt13 := TSH08.COUNT;
                    END;

                END;

                IF NOT TSHNoSeries THEN
                    CurrReport.SKIP;

                //>>RB-N 09Nov2017 Invoice NoSeries
            end;

            trigger OnPreDataItem()
            begin

                SETRANGE("Posting Date", SDate22, EDate22);
                SETFILTER("No. Series", '%1', 'BRT*');
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                group(Options)
                {
                    Caption = 'Options';
                    field(GSTIN; GSTIN)
                    {
                        Caption = 'GSTIN of the location';
                        TableRelation = "GST Registration Nos.".Code;

                        trigger OnValidate()
                        begin
                            /*
                            Location.SETRANGE("GST Registration No.",GSTIN);
                            IF Location.FINDFIRST THEN
                              TaxablePerson := Location.Name
                            ELSE
                              ERROR(LocationErr);
                            
                            */

                            //>> RB-N 07Nov2017
                            LocState09 := '';//09Nov2017
                            Location.SETRANGE("GST Registration No.", GSTIN);
                            IF Location.FINDFIRST THEN BEGIN
                                LocState09 := Location."State Code";
                                State07.RESET;
                                State07.SETRANGE(Code, Location."State Code");
                                IF State07.FINDFIRST THEN
                                    TaxablePerson := COMPANYNAME + ' - ' + State07.Description;
                            END ELSE
                                ERROR(LocationErr);

                        end;
                    }
                    field(TaxablePerson; TaxablePerson)
                    {
                        ApplicationArea = all;
                        Caption = 'Name of the Taxable Person';
                        TableRelation = Location.Name;
                    }
                    field(AggregateTurnover; AggregateTurnover)
                    {
                        ApplicationArea = all;
                        Caption = 'Aggregate Turnover of the Taxable person in the previous financial year';
                    }
                    field(Period; Period)
                    {
                        ApplicationArea = all;
                        Caption = 'Period';
                    }
                    field(PersonName; PersonName)
                    {
                        ApplicationArea = all;
                        Caption = 'Name of the Authorized Person';
                    }
                    field(Place; Place)
                    {
                        ApplicationArea = all;
                        Caption = 'Place';
                    }
                    field(Date; Date)
                    {
                        Caption = 'Date';
                        ApplicationArea = all;
                        trigger OnValidate()
                        begin
                            Year := DATE2DMY(Date, 3);
                        end;
                    }
                    field("Print Invoice Details"; InvoiceDetail)
                    {
                        ApplicationArea = all;
                    }
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
        Amt := 250000; // Defined Amount to check the Sales Invoice Amount limit according to GSTR-1 Rule
    end;

    trigger OnPreReport()
    begin

        //>>RB-N 22Sep2017 Date
        CASE Period OF
            Period::Jan:
                BEGIN
                    SDate22 := DMY2DATE(1, 1, Year);
                    EDate22 := DMY2DATE(31, 1, Year);
                END;
            Period::Feb:
                BEGIN
                    SDate22 := DMY2DATE(1, 2, Year);
                    EDate22 := DMY2DATE(28, 2, Year);
                END;
            Period::March:
                BEGIN
                    SDate22 := DMY2DATE(1, 3, Year);
                    EDate22 := DMY2DATE(31, 3, Year);
                END;
            Period::Apr:
                BEGIN
                    SDate22 := DMY2DATE(1, 4, Year);
                    EDate22 := DMY2DATE(30, 4, Year);
                END;
            Period::May:
                BEGIN
                    SDate22 := DMY2DATE(1, 5, Year);
                    EDate22 := DMY2DATE(31, 5, Year);
                END;
            Period::June:
                BEGIN
                    SDate22 := DMY2DATE(1, 6, Year);
                    EDate22 := DMY2DATE(30, 6, Year);
                END;
            Period::July:
                BEGIN
                    SDate22 := DMY2DATE(1, 7, Year);
                    EDate22 := DMY2DATE(31, 7, Year);
                END;
            Period::Agu:
                BEGIN
                    SDate22 := DMY2DATE(1, 8, Year);
                    EDate22 := DMY2DATE(31, 8, Year);
                END;
            Period::Sept:
                BEGIN
                    SDate22 := DMY2DATE(1, 9, Year);
                    EDate22 := DMY2DATE(30, 9, Year);
                END;
            Period::Oct:
                BEGIN
                    SDate22 := DMY2DATE(1, 10, Year);
                    EDate22 := DMY2DATE(31, 10, Year);
                END;
            Period::Nov:
                BEGIN
                    SDate22 := DMY2DATE(1, 11, Year);
                    EDate22 := DMY2DATE(30, 11, Year);
                END;
            Period::Dec:
                BEGIN
                    SDate22 := DMY2DATE(1, 12, Year);
                    EDate22 := DMY2DATE(31, 12, Year);
                END;
        END;
        //>>RB-N 22Sep2017 Date

        H08 := 0;//08Nov2017
        Ex08.DELETEALL;//08Nov2017

        H09 := 0;//09Nov2017
        Ex09.DELETEALL;//09Nov2017
    end;

    var
        //GSTComponent: Record 16405;
        DummyCustomer: Record 18;
        Location: Record 14;
        GSTComp: array[50] of Code[10];
        GSTIN: Code[20];
        BuyerSellerRegNo: array[50] of Code[15];
        BuyerSellerStateCode: array[50] of Code[10];
        DocNo: array[50] of Code[20];
        TransactionNo: Code[20];
        NoSeries: Code[20];
        NoSeries1: Code[20];
        StartingNo: Code[20];
        StartingNo1: Code[20];
        EndingNo: Code[20];
        EndingNo1: Code[20];
        ItemNo: array[4] of Code[20];
        HSNCode: array[50] of Code[10];
        Pos: array[50] of Code[10];
        Ecom: array[50] of Code[15];
        EcomMerchant: array[50] of Code[30];
        ExportNo: array[5] of Code[10];
        OriginalInvNo: Code[20];
        GSTCompRate: array[50] of Decimal;
        GSTCompAmt: array[50] of Decimal;
        GSTBaseAmt: array[50] of Decimal;
        DummyInvoiceAmt: Decimal;
        Amt: Decimal;
        DummyInvoiceValue: array[4] of Decimal;
        DummyGSTAmount: Decimal;
        DummyNilRated: array[4] of Decimal;
        ExcemptedAmt: array[4] of Decimal;
        NonGSTSupplies: array[4] of Decimal;
        AggregateTurnover: Decimal;
        CrossValAmt: Decimal;
        CustomerName: Text;
        TaxablePerson: Text;
        PersonName: Text;
        Place: Text;
        GoodService: array[20] of Text;
        GoodSerDescription: array[4] of Text;
        NilRatedText: array[4] of Text;
        WithoutPaymentText: Text;
        WithPaymentText: Text;
        Period: Option Jan,Feb,March,Apr,May,June,July,Agu,Sept,Oct,Nov,Dec;
        DocumentType: array[50] of Option;
        Year: Integer;
        TitleLbl: Label 'Government of India/State';
        Title2Lbl: Label 'Department of...............';
        Title3Lbl: Label 'Form GSTR-1';
        Title4Lbl: Label '[See Rule…..]';
        HeadingLbl: Label 'DETAILS OF OUTWARD SUPPLIES';
        GSTLbl: Label '1.  GSTIN:';
        PersonNameLbl: Label '2.  Name of the Taxable Person:';
        PersonSubLbl: Label ' (S. No. 1 and 2 will be auto-populated on logging)';
        AggregateTurnOverLbl: Label '3. Aggregate Turnover of the Taxable Person in the Previous FY';
        AggregateSubLbl: Label ' (To be submitted only in first year. To be auto populated in subsequent year)';
        PeriodLbl: Label '4. Period:';
        TaxableHeadingLbl: Label '5. Taxable outward supplies to a registered person';
        FiguresLbl: Label '(figures in Rs)';
        GSTINUINLbl: Label 'GSTIN/UIN';
        InvoiceLbl: Label 'Invoice';
        NoLbl: Label 'No.';
        DateLbl: Label 'Date';
        ValueLbl: Label 'Value';
        GoodSerLbl: Label 'Goods/Services';
        HSNSACLbl: Label 'HSN/SAC';
        TaxableLbl: Label 'Taxable Value';
        IGSTLbl: Label 'Igst';
        CGSTLbl: Label 'Cgst';
        SGSTLbl: Label 'Sgst';
        RateLbl: Label 'Rate';
        AmtLbl: Label 'Amt';
        POSLbl: Label 'POS(Only if different from the location of recipient)';
        RevChargeLbl: Label 'Indicate if supply attracts reverse charge $';
        AssessmentLbl: Label 'Tax on this Invoice is paid under provisional assessment (Checkbox)';
        ecommerceLbl: Label 'GSTIN of e-commerce operator (if applicable)';
        FCLbl: Label ' $    To be filled only if a supply attracts reverse charge';
        NotesLbl: Label 'Notes:-';
        Note1Lbl: Label '1. Taxable Person has the option to furnish the details of nil rate and exempted supplies in this Table';
        Note2Lbl: Label '2. In case of inter-state supplies, only IGST would be filled';
        Note3Lbl: Label '3. In case of intra-state supplies, CGST & SGST would be filled.';
        AmendmentHeadingLbl: Label '5A. Amendments to details of Outward Supplies to a registered person of earlier tax periods';
        OriginalInvLbl: Label 'Original Invoice';
        RevisedOriginalLbl: Label 'Revised/Original Invoice';
        Taxable2HeadingLbl: Label '6. Taxable outward supplies to a consumer where Place of Supply (State Code) is other than the State where supplier is located (Inter-state supplies) and Invoice value is more than Rs 2.5 lakh ';
        RecipientLbl: Label 'Recipient’s State code';
        RecNameLbl: Label 'Name of the recipient';
        Amendment2HeadingLbl: Label '6A. Amendment to taxable outward supplies to a consumer of earlier tax periods where Place of Supply (State Code) is other than the State where supplier is located (Inter-state supplies) and Invoice value is more than Rs 2.5 lakh';
        Taxable3HeadlingLbl: Label '7.  Taxable outward supplies to consumer (Other than 6 above)';
        StateCodeLbl: Label 'State code  (Place of Supply)';
        AggregateLbl: Label 'Aggregate Taxable Value';
        SupplyLbl: Label 'Tax on this supply is paid under provisional assessment (Checkbox)';
        Note4Lbl: Label '2. Table includes both inter-state supplies (invoice value below 2.5 lakhs) and intra-state supplies.';
        Ammendment3HeadingLbl: Label '7A.    Amendment to Taxable outward supplies to consumer of earlier tax periods (original supplies covered under 7 above in earlier tax period (s))';
        OriginalDetLbl: Label 'Original Details';
        MonthPeriodLbl: Label 'Month (Tax Period)';
        StatLbl: Label 'State Code';
        ReverseDetLbl: Label 'Revised Details';
        POSStatLbl: Label 'State code  (Place of Supply (State Code))';
        DetailsHeadingLbl: Label '8. Details of Credit/Debit Notes';
        GSTINRecNameLbl: Label 'GSTIN /UIN/ Name of recipient';
        OtherChargeLbl: Label 'Other than reverse charge';
        ReverseChargeLbl: Label 'Reverse charge';
        DebitCreditLbl: Label 'Debit Note/credit note';
        DifferentialValLbl: Label 'Differential Value  (Plus or Minus)';
        DifferentialTaxLbl: Label 'Differential Tax';
        Note5Lbl: Label 'Note: Information about Credit Note / Debit Note to be submitted only if issued as a supplier.';
        Ammendment4HeadingLbl: Label '8A.   Amendment to Details of Credit/Debit Notes of earlier tax periods';
        OriginalLbl: Label 'Original';
        RevisedLbl: Label 'Revised';
        OriginalInvDetLbl: Label 'Original Invoice details';
        NilrateHeadlingLbl: Label '9. Nil rated,  Exempted  and Non GST outward supplies*';
        InterSupplyPerLbl: Label 'Interstate supplies to registered person';
        IntraSupplyPerLbl: Label 'Intrastate supplies to registered person';
        InterSupplyConsLbl: Label 'Interstate supplies to consumer';
        IntraSupplyConsLbl: Label 'Intrastate supplies to consumer';
        NilRateLbl: Label 'Nil Rated (Amount)';
        ExemptLbl: Label 'Exempted (Amount)';
        NonGSTLbl: Label 'Non GST supplies (Amount)';
        Note6Lbl: Label '• If the details of “nil”” rated and “exempt” supplies have been provided in Table 5, 6 and 7, then info in column (4) may only be furnished.';
        SuppliesLbl: Label '10. Supplies Exported (including deemed exports)';
        DescriptionLbl: Label 'Description';
        WithoutPaymentGSTLbl: Label 'Without payment of GST';
        WithPaymentGSTLbl: Label 'With payment of GST';
        ShippingLbl: Label 'Shipping bill/ bill of export';
        Ammendment5HeadingLbl: Label '10A. Amendment to Supplies Exported (including deemed exports)';
        RevisedInvLbl: Label 'Revised Invoice';
        TaxHeadingLbl: Label '11. Tax liability arising on account of Time of Supply without issuance of Invoice in the same period.';
        GSTCustNameLbl: Label 'GSTIN/UIN/ Name of customer';
        DocNoLbl: Label 'Document No.';
        HSNSACSupplyLbl: Label 'HSN/SAC  of supply';
        AmountofAdvLbl: Label 'Amount of advance received/ Value of Supply provided without raising a bill';
        TaxLbl: Label 'TAX';
        Note7Lbl: Label 'Note: A transaction id would be generated by system for each transaction on which tax is paid in advance/on account of time of supply';
        Ammendment6HeadingLbl: Label '11A. Amendment to Tax liability arising on account of Time of Supply without issuance of Invoice in the same tax period.';
        DocNumLbl: Label 'Document Number';
        HSCPOSLbl: Label 'HSN/SAC of supply to be made';
        Tax_Lbl: Label 'Tax';
        Tax2HeadingLbl: Label '12. Tax already paid (on advance receipt/ on account of time of supply) on invoices issued in the current period';
        TransactionLbl: Label 'Transaction id (A number assigned by the system when tax was paid)';
        TaxPaidLbl: Label 'TAX Paid on receipt of advance/on account of time of supply';
        Note8Lbl: Label 'Note: Tax liability in respect of invoices issued in this period shall be net of tax already paid on advance receipt/on occurrence of time of supply';
        SuppliesHeadlingLbl: Label '13.   Supplies made through e-commerce portals of other companies';
        Part1HeadlingLbl: Label 'Part 1- Supplies made through e-commerce portals of other companies to Registered Taxable Persons';
        MarchantLbl: Label 'Marchant ID issued by ecommerce operator';
        GSTINecommerceLbl: Label 'GSTIN of e-commerce portal';
        GrossLbl: Label 'Gross Value of supplies';
        Note9Lbl: Label 'Note: Details of supplies made through e-commerce portal to registered Taxable Persons shall be reported in Table 5 of this return, which shall be prepopulated in this table based on the flag provided in the respective table at the time of creation of Return.';
        Part2HeadlingLbl: Label 'Part 2- Supplies made through e-commerce portals of other companies to Unregistered Persons';
        SNoLbl: Label 'Sr No.';
        Note10Lbl: Label 'Note: Details of supplies made through e-commerce portal to unregistered Taxable Persons shall be reported in the table by the Taxable Person in addition to the details which are already provided in Table 6 & 7 of this return, this shall not be included in the turnover again.';
        Part2AHeadingLbl: Label 'Part- 2A Amendment to Supplies made through e-commerce portals of other companies to Unregistered Taxable Persons';
        TaxPeriodLbl: Label 'Tax period of supplies';
        PlaceofSuppStateLbl: Label 'Place of Supply (State Code)';
        InvoicesHeadlingLbl: Label '14. Invoices issued during the tax period including invoices issued in case of inward supplies received from unregistered persons liable for reverse charge';
        SeriesLbl: Label 'Series number of invoices';
        FromLbl: Label 'From';
        ToLbl: Label 'To';
        TotalInvLbl: Label 'Total number of invoices';
        NoCancelInvLbl: Label 'Number of cancelled invoices';
        NetNoInvLbl: Label 'Net Number of invoices issued';
        Declare2Lbl: Label 'hereby declare that the information given in this statement is true, correct and complete in every respect. I further';
        Declare3Lbl: Label 'declare that I have the legal authority to submit this statement.';
        PlaceLbl: Label 'Place:';
        Date_Lbl: Label 'Date:';
        SignatureLbl: Label '(Signature of Authorized Person)';
        InstructionLbl: Label 'INSTRUCTIONS for furnishing the information';
        TermsLbl: Label '1. Terms used:';
        TermsGSTLbl: Label '  GSTIN: Goods and Services Taxable Person Identification Number';
        TermsUINLbl: Label '  UIN:     Unique Identity Number for embassies';
        TermsHSNLbl: Label '  HSN:    Harmonized System of Nomenclature for goods';
        TermsSACLbl: Label '  SAC:     Service Accounting Code';
        TermsPOSLbl: Label '  POS:      Place of Supply (State Code) of goods or services – State Code to be mentioned';
        Terms2Lbl: Label '2. To be furnished by the 10th of the month succeeding the tax period. Not to be furnished by compounding Taxable Person/ISD ';
        Terms3Lbl: Label '3. Aggregate Turnover means as defined under the Goods and Services Tax Act, 20….. ';
        Terms4Lbl: Label '4. HSN/SAC is not mandatory for taxable person whose aggregate turnover is less than 1.5 crores. HSN shall be restricted to maximum 8 digits. If gross turnover in previous financial year is greater than Rs 5 crore, HSN should be minimum of 4 digits. If gross turnover in previous financial year is equal to or greater than Rs 1.5 crore and less than 5 crore, HSN should be minimum of 2 digit and would be mandatory from the second year of GST implementation. In case of Exports HSN should be 8 digits.';
        Printed: Boolean;
        PurchPrint: Boolean;
        no: Integer;
        i: Integer;
        RecordCount: Integer;
        RecordCount1: Integer;
        SNo: Integer;
        PostingDate: array[50] of Date;
        Date: Date;
        GSTINErr: Label 'GSTIN of the location can not be blank.';
        PersonNameErr: Label 'Name of the Taxable Person can not be blank.';
        AggregateTurnoverErr: Label 'Aggregate Turnover of the Taxable Person can not be blank.';
        AuthorizedPersonErr: Label 'Name of the Authorized Person can not be empty.';
        PlaceErr: Label 'Place can not be empty.';
        DateErr: Label 'Date can not be empty.';
        LocationErr: Label 'The given registration is not found with any location.';
        ExportDate: array[5] of Date;
        DocDate: Date;
        OriginalInvPostDate: Date;
        ReportViewErr: Label 'GST Component setup must have value in the Report View field.';
        Component: array[50] of Code[10];
        "----------14Sep2017": Integer;
        Itm14: Record 27;
        ItmDesc: Text[100];
        CGST17: Decimal;
        SGST17: Decimal;
        IGST17: Decimal;
        UTGST17: Decimal;
        CESS17: Decimal;
        DGST: Record "Detailed GST Ledger Entry";
        CGSTP: Decimal;
        SGSTP: Decimal;
        IGSTP: Decimal;
        UTGSTP: Decimal;
        InvoiceDetail: Boolean;
        InvDD: Integer;
        "--------------22Sep2017": Integer;
        State22: Record State;
        PSupply: Text[80];
        GSTBaseAmt22: Decimal;
        TDocNo: Code[20];
        TLinNo: Integer;
        SIL22: Record 113;
        HSN22: Record "HSN/SAC";
        HSNDesc: Text[50];
        ItmUom: Code[20];
        QtyBase2211: Decimal;
        GSTBaseAmt2211: Decimal;
        Loc2211: Record 14;
        SDate22: Date;
        EDate22: Date;
        ItmQtyUOM: Decimal;
        ItmUOM22: Record 5404;
        DGST18: Record "Detailed GST Ledger Entry";
        III: Integer;
        TTT: Integer;
        CGST1711: Decimal;
        SGST1711: Decimal;
        IGST1711: Decimal;
        III111: Integer;
        TTT111: Integer;
        "--------07Nov2017": Integer;
        NNN111: Integer;
        NNN222: Integer;
        Loc07: Record 14;
        State07: Record State;
        "-----08Nov2017--ExportInvoice": Integer;
        SIL08: Record 113;
        SIH08: Record 112;
        ExpInvoice: Boolean;
        TempHSN: Code[10];
        Ex08: Record 370 temporary;
        H08: Integer;
        Loc08: Record 14;
        CGST1708: Decimal;
        SGST1708: Decimal;
        IGST1708: Decimal;
        UTGST1708: Decimal;
        CESS1708: Decimal;
        CGSTP08: Decimal;
        SGSTP08: Decimal;
        IGSTP08: Decimal;
        UTGSTP08: Decimal;
        GSTBaseAmt08: Decimal;
        "-----09Nov2017--Noseries": Integer;
        LocState09: Code[20];
        InvNoSeries: Boolean;
        TempInvNo: Code[20];
        Ex09: Record 370 temporary;
        H09: Integer;
        InvStartNo09: Code[20];
        InvEndNo09: Code[20];
        InvCount09: Integer;
        TSHNoSeries: Boolean;
        TempTshNo: Code[20];
        TshStartNo09: Code[20];
        TshEndNo09: Code[20];
        TshCount09: Integer;
        TSH08: Record 5744;
        Loc09: Record 14;
        CanInv13: Integer;
        CanBrt13: Integer;

    local procedure AssignGSTAmt(DetailedGSTLedgerEntry1: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; var GSTBaseAmt: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    var
        DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry";
        DetailedGSTLedgerEntry2: Record "Detailed GST Ledger Entry";
    begin
        WITH DetailedGSTLedgerEntry DO BEGIN
            SETCURRENTKEY(
              "Document Type", "Document No.", "Document Line No.",/* "Nature of Supply",*/ "GST Jurisdiction Type");
            SETRANGE("Document Type", DetailedGSTLedgerEntry1."Document Type");
            SETRANGE("Document No.", DetailedGSTLedgerEntry1."Document No.");
            // SETRANGE("Nature of Supply", DetailedGSTLedgerEntry1."Nature of Supply");
            SETRANGE("GST Jurisdiction Type", DetailedGSTLedgerEntry1."GST Jurisdiction Type");
            SETRANGE("HSN/SAC Code", DetailedGSTLedgerEntry1."HSN/SAC Code");
            IF FINDSET THEN
                REPEAT
                    IF "GST Component Code" = GSTComp1 THEN BEGIN
                        GSTCompRate1 := "GST %";   //RK
                        GSTCompAmt1 += "GST Amount";
                    END;
                    IF "GST Component Code" = GSTComp2 THEN BEGIN
                        GSTCompRate2 := "GST %"; //RK
                        GSTCompAmt2 += "GST Amount";
                    END;
                    IF "GST Component Code" = GSTComp3 THEN BEGIN
                        GSTCompRate3 := "GST %"; //RK
                        GSTCompAmt3 += "GST Amount";
                    END;
                UNTIL NEXT = 0;
        END;

        DetailedGSTLedgerEntry2.SETRANGE("Document Type", DetailedGSTLedgerEntry1."Document Type");
        DetailedGSTLedgerEntry2.SETRANGE("Document No.", DetailedGSTLedgerEntry1."Document No.");
        DetailedGSTLedgerEntry2.SETRANGE("GST Component Code", DetailedGSTLedgerEntry1."GST Component Code");
        DetailedGSTLedgerEntry2.SETRANGE("No.", DetailedGSTLedgerEntry1."No.");
        IF DetailedGSTLedgerEntry2.FINDFIRST THEN
            IF DetailedGSTLedgerEntry2."GST Jurisdiction Type" = DetailedGSTLedgerEntry2."GST Jurisdiction Type"::Intrastate THEN
                GSTBaseAmt += DetailedGSTLedgerEntry2."GST Base Amount" / 2
            ELSE
                GSTBaseAmt += DetailedGSTLedgerEntry2."GST Base Amount";
    end;

    local procedure TaxableB2B(DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    var
        SalesInvoiceLine: Record 113;
    begin
        /* WITH DetailedGSTLedgerEntry DO
            IF "Nature of Supply" = "Nature of Supply"::B2B THEN BEGIN
                AssignGSTAmt(
                  DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2, GSTCompRate3, GSTCompAmt1,
                  GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[1], GSTComp1, GSTComp2, GSTComp3);
                BuyerSellerRegNo[1] := "Buyer/Seller Reg. No.";
                PostingDate[1] := "Posting Date";
                DocNo[1] := "Document No.";
                HSNCode[1] := "HSN/SAC Code";
                GetGoodsAndService(DetailedGSTLedgerEntry, GoodService[1]);
                IF "Buyer/Seller State Code" <> "Shipping Address State Code" THEN
                    Pos[1] := "Shipping Address State Code";
                Ecom[1] := "e-Comm. Operator GST Reg. No.";
                EcomMerchant[1] := "e-Comm. Merchant Id";
                SalesInvoiceLine.SETRANGE("Document No.", "Document No.");
                IF SalesInvoiceLine.FINDFIRST THEN
                    CrossValAmt += SalesInvoiceLine."Unit Price";
                Component[1] := "GST Component Code";
            END; */
    end;

    local procedure TaxableB2C(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    begin
        WITH DetailedGSTLedgerEntry DO
            //IF ("Nature of Supply" = "Nature of Supply"::B2C) AND ("GST Jurisdiction Type" = "GST Jurisdiction Type"::Interstate) THEN
            IF CheckInvoiceAmount("Document No.") THEN BEGIN
                AssignGSTAmt(
                  DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2, GSTCompRate3,
                  GSTCompAmt1, GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[2], GSTComp1, GSTComp2, GSTComp3);
                BuyerSellerRegNo[2] := "Buyer/Seller Reg. No.";
                PostingDate[2] := "Posting Date";
                DocNo[2] := "Document No.";
                HSNCode[2] := "HSN/SAC Code";
                BuyerSellerStateCode[2] := '';//"Buyer/Seller State Code";
                GetGoodsAndService(DetailedGSTLedgerEntry, GoodService[2]);
                // IF "Buyer/Seller State Code" <> "Shipping Address State Code" THEN
                Pos[2] := '';//"Shipping Address State Code";
                IF DummyCustomer.GET("Source No.") THEN
                    CustomerName := DummyCustomer.Name;
            END;
    end;

    local procedure CheckInvoiceAmount(DocumentNo: Code[20]): Boolean
    var
        SalesInvoiceHeader: Record 112;
    begin
        IF SalesInvoiceHeader.GET(DocumentNo) THEN BEGIN
            SalesInvoiceHeader.CALCFIELDS(Amount);
            EXIT(SalesInvoiceHeader.Amount > Amt);
        END;
    end;

    local procedure CheckInvoiceAmountValue(DocumentNo: Code[20]): Boolean
    var
        SalesInvoiceHeader: Record 112;
    begin
        IF SalesInvoiceHeader.GET(DocumentNo) THEN BEGIN
            SalesInvoiceHeader.CALCFIELDS(Amount);
            EXIT(SalesInvoiceHeader.Amount <= Amt);
        END;
    end;

    local procedure TaxableB2CWithLimit(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    begin
        /*  WITH DetailedGSTLedgerEntry DO
              IF ("Nature of Supply" = "Nature of Supply"::B2C) AND CheckInvoiceAmountValue("Document No.") THEN BEGIN
             AssignGSTAmt(
               DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2, GSTCompRate3,
               GSTCompAmt1, GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[3], GSTComp1, GSTComp2, GSTComp3);
         HSNCode[3] :="HSN/SAC Code";
         GetGoodsAndService(DetailedGSTLedgerEntry, GoodService[3]);
         DocNo[8] := "Document No.";
         IF "Buyer/Seller State Code" <> "Shipping Address State Code" THEN
         Pos[3] :=  "Shipping Address State Code"
                        END; */
    end;

    local procedure NilRatedExempted(DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GoodService: Text; var Exempted: Decimal; ItemNo: Code[20])
    var
        DetailedGSTLedgerEntry1: Record "Detailed GST Ledger Entry";
    begin
        GetGoodsAndService(DetailedGSTLedgerEntry, GoodService);
        Exempted := 0;
        DetailedGSTLedgerEntry1.SETRANGE("No.", DetailedGSTLedgerEntry."No.");
        // DetailedGSTLedgerEntry1.SETRANGE("Nature of Supply", DetailedGSTLedgerEntry."Nature of Supply");
        DetailedGSTLedgerEntry1.SETRANGE("GST Jurisdiction Type", DetailedGSTLedgerEntry."GST Jurisdiction Type");
        DetailedGSTLedgerEntry1.SETRANGE("GST Customer Type", DetailedGSTLedgerEntry1."GST Customer Type"::Exempted);
        // DetailedGSTLedgerEntry1.SETRANGE("Invoice Type", DetailedGSTLedgerEntry1."Invoice Type"::"Bill of Supply");
        IF DetailedGSTLedgerEntry1.FINDSET THEN
            REPEAT
                Exempted += DetailedGSTLedgerEntry1."GST Amount";
                ItemNo := DetailedGSTLedgerEntry1."No.";
            UNTIL DetailedGSTLedgerEntry1.NEXT = 0;
    end;

    local procedure NilRatedNonGST(DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GoodService: Text; var NonGST: Decimal; ItemNo: Code[20])
    var
        DetailedGSTLedgerEntry1: Record "Detailed GST Ledger Entry";
    begin
        GetGoodsAndService(DetailedGSTLedgerEntry, GoodService);
        NonGST := 0;
        DetailedGSTLedgerEntry1.SETRANGE("No.", DetailedGSTLedgerEntry."No.");
        //DetailedGSTLedgerEntry1.SETRANGE("Nature of Supply", DetailedGSTLedgerEntry."Nature of Supply");
        DetailedGSTLedgerEntry1.SETRANGE("GST Jurisdiction Type", DetailedGSTLedgerEntry."GST Jurisdiction Type");
        //DetailedGSTLedgerEntry1.SETRANGE("Invoice Type", DetailedGSTLedgerEntry1."Invoice Type"::"Non-GST");
        IF DetailedGSTLedgerEntry1.FINDSET THEN
            REPEAT
                NonGST += DetailedGSTLedgerEntry1."GST Amount";
                ItemNo := DetailedGSTLedgerEntry1."No.";
            UNTIL DetailedGSTLedgerEntry1.NEXT = 0;
    end;

    local procedure DetailsCreditAndDebit(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    begin
        AssignGSTAmt(
          DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2, GSTCompRate3,
          GSTCompAmt1, GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[4], GSTComp1, GSTComp2, GSTComp3);
        HSNCode[12] := DetailedGSTLedgerEntry."HSN/SAC Code";
        BuyerSellerRegNo[3] := DetailedGSTLedgerEntry."Buyer/Seller Reg. No.";
        PostingDate[3] := DetailedGSTLedgerEntry."Posting Date";
        DocNo[3] := DetailedGSTLedgerEntry."Document No.";
        //DocumentType[1] :=DetailedGSTLedgerEntry."Invoice Type";
        OriginalInvNo := DetailedGSTLedgerEntry."Original Invoice No.";
        OriginalInvPostDate := 0D;// DetailedGSTLedgerEntry."Original Invoice Date";
    end;

    local procedure SuppliesExported(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; var DocNo: Code[20]; var PostingDate: Date; var HSNCode: Code[10]; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10]; ExpNo: Code[20]; ExpDate: Date)
    begin
        IF DetailedGSTLedgerEntry."GST Customer Type" IN [
                                                          DetailedGSTLedgerEntry."GST Customer Type"::Export,
                                                          DetailedGSTLedgerEntry."GST Customer Type"::"Deemed Export"]
        THEN BEGIN
            IF DetailedGSTLedgerEntry."GST Without Payment of Duty" THEN
                AssignGSTAmt(
                  DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2, GSTCompRate3,
                  GSTCompAmt1, GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[5], GSTComp1, GSTComp2, GSTComp3)
            ELSE
                AssignGSTAmt(
                  DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2, GSTCompRate3,
                  GSTCompAmt1, GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[7], GSTComp1, GSTComp2, GSTComp3);
            DocNo := DetailedGSTLedgerEntry."Document No.";
            PostingDate := DetailedGSTLedgerEntry."Posting Date";
            HSNCode := DetailedGSTLedgerEntry."HSN/SAC Code";
            ExpDate := 0D;// DetailedGSTLedgerEntry."Bill Of Export Date";
            ExpNo := ''; //DetailedGSTLedgerEntry."Bill Of Export No.";
            GetGoodsAndService(DetailedGSTLedgerEntry, GoodService[4]);
        END;
    end;

    local procedure AdvancePaymentAfterPosting(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    begin
        WITH DetailedGSTLedgerEntry DO BEGIN
            AssignGSTAmt(
              DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2, GSTCompRate3, GSTCompAmt1,
              GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[11], GSTComp1, GSTComp2, GSTComp3);
            DocNo[11] := "Document No.";
            PostingDate[11] := "Posting Date";
            HSNCode[11] := "HSN/SAC Code";
            /*  IF "Buyer/Seller State Code" <> "Shipping Address State Code" THEN
                 Pos[4] := "Shipping Address State Code"; */
            GetGoodsAndService(DetailedGSTLedgerEntry, GoodService[5]);
        END;
    end;

    local procedure AdvancePaymentBeforePosting(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    begin
        AssignGSTAmt(
          DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2,
          GSTCompRate3, GSTCompAmt1, GSTCompAmt2, GSTBaseAmt[8], GSTCompAmt3, GSTComp1, GSTComp2, GSTComp3);
        HSNCode[13] := DetailedGSTLedgerEntry."HSN/SAC Code";
        DocNo[7] := DetailedGSTLedgerEntry."Document No.";
        TransactionNo := DetailedGSTLedgerEntry."Payment Document No.";
    end;

    local procedure AdvancePaymentAdjustmentJnl(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    var
        Customer: Record 18;
    begin
        /*WITH DetailedGSTLedgerEntry DO
            IF "Adv. Pmt. Adjustment" THEN BEGIN
                AssignGSTAmt(
                  DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2,
                  GSTCompRate3, GSTCompAmt1, GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[9], GSTComp1, GSTComp2, GSTComp3);
                Customer.GET("Source No.");
                IF "Nature of Supply" = "Nature of Supply"::B2C THEN
                    BuyerSellerRegNo[5] := COPYSTR(Customer.Name, 1, 15)
                ELSE
                    BuyerSellerRegNo[5] := FORMAT("Buyer/Seller Reg. No.");
                DocNo[9] := "Original Adv. Pmt Doc. No.";
                DocDate := "Original Adv. Pmt Doc. Date";
                IF "Buyer/Seller State Code" <> "Shipping Address State Code" THEN
                    Pos[5] := "Shipping Address State Code";
                DocNo[10] := "Document No.";
                PostingDate[7] := "Posting Date";
                GetGoodsAndService(DetailedGSTLedgerEntry, GoodService[6]);
                HSNCode[7] := "HSN/SAC Code";
            END; */
    end;

    local procedure SuppliesUnregCustThroughECom(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GSTCompRate1: Decimal; var GSTCompRate2: Decimal; var GSTCompRate3: Decimal; var GSTCompAmt1: Decimal; var GSTCompAmt2: Decimal; var GSTCompAmt3: Decimal; GSTComp1: Code[10]; GSTComp2: Code[10]; GSTComp3: Code[10])
    begin
        WITH DetailedGSTLedgerEntry DO BEGIN
            SNo += 1;
            AssignGSTAmt(
              DetailedGSTLedgerEntry, GSTCompRate1, GSTCompRate2,
              GSTCompRate3, GSTCompAmt1, GSTCompAmt2, GSTCompAmt3, GSTBaseAmt[10], GSTComp1, GSTComp2, GSTComp3);
            Ecom[2] := '';// "e-Comm. Operator GST Reg. No.";
            EcomMerchant[2] := '';// "e-Comm. Merchant Id";
            HSNCode[14] := "HSN/SAC Code";
            /* IF "Buyer/Seller State Code" <> "Shipping Address State Code" THEN
                Pos[6] := "Shipping Address State Code"; */
        END;
    end;

    local procedure GetNoSeries()
    var
        NoSeriesLine: Record 309;
        DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry";
        DetailedGSTLedgerEntry2: Record "Detailed GST Ledger Entry";
        DetailedGSTLedgerEntry3: Record "Detailed GST Ledger Entry";
    begin
        IF NOT Printed THEN BEGIN
            DetailedGSTLedgerEntry.SETRANGE("Transaction Type", DetailedGSTLedgerEntry."Transaction Type"::Sales);
            FilterPostingDate(DetailedGSTLedgerEntry);
            IF DetailedGSTLedgerEntry.FINDFIRST THEN
                StartingNo := DetailedGSTLedgerEntry."Document No.";
            DetailedGSTLedgerEntry2.SETRANGE("Transaction Type", DetailedGSTLedgerEntry2."Transaction Type"::Sales);
            FilterPostingDate(DetailedGSTLedgerEntry2);
            IF DetailedGSTLedgerEntry2.FINDLAST THEN BEGIN
                EndingNo := DetailedGSTLedgerEntry2."Document No.";
                //MESSAGE('Start No %1 \ Ending No. %2',StartingNo,EndingNo);

                NoSeriesLine.SETRANGE("Last No. Used", EndingNo);
                IF NoSeriesLine.FINDFIRST THEN
                    NoSeries := NoSeriesLine."Starting No.";
                //MESSAGE('NoSeries %1 ',NoSeries);
            END;
            DetailedGSTLedgerEntry3.SETRANGE("Transaction Type", DetailedGSTLedgerEntry3."Transaction Type"::Sales);
            DetailedGSTLedgerEntry3.SETRANGE("Document No.", StartingNo, EndingNo);
            IF DetailedGSTLedgerEntry3.FINDFIRST THEN
                RecordCount := DetailedGSTLedgerEntry3.COUNT;
            //MESSAGE('Record Count %1 ',RecordCount);
            Printed := TRUE;
        END;
    end;

    local procedure GetPurchNoSeries()
    var
        NoSeriesLine: Record 309;
        DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry";
        DetailedGSTLedgerEntry2: Record "Detailed GST Ledger Entry";
        DetailedGSTLedgerEntry3: Record "Detailed GST Ledger Entry";
    begin
        IF NOT PurchPrint THEN BEGIN
            DetailedGSTLedgerEntry.SETRANGE("GST Vendor Type", DetailedGSTLedgerEntry."GST Vendor Type"::Unregistered);
            DetailedGSTLedgerEntry.SETRANGE("Transaction Type", DetailedGSTLedgerEntry."Transaction Type"::Purchase);
            FilterPostingDate(DetailedGSTLedgerEntry);
            IF DetailedGSTLedgerEntry.FINDFIRST THEN
                StartingNo1 := DetailedGSTLedgerEntry."Document No.";
            DetailedGSTLedgerEntry2.SETRANGE("GST Vendor Type", DetailedGSTLedgerEntry2."GST Vendor Type"::Unregistered);
            DetailedGSTLedgerEntry2.SETRANGE("Transaction Type", DetailedGSTLedgerEntry2."Transaction Type"::Purchase);
            FilterPostingDate(DetailedGSTLedgerEntry2);
            IF DetailedGSTLedgerEntry2.FINDLAST THEN BEGIN
                EndingNo1 := DetailedGSTLedgerEntry2."Document No.";
                NoSeriesLine.SETRANGE("Last No. Used", EndingNo1);
                IF NoSeriesLine.FINDFIRST THEN
                    NoSeries1 := NoSeriesLine."Starting No.";
            END;
            DetailedGSTLedgerEntry3.SETRANGE("GST Vendor Type", DetailedGSTLedgerEntry3."GST Vendor Type"::Unregistered);
            DetailedGSTLedgerEntry3.SETRANGE("Transaction Type", DetailedGSTLedgerEntry3."Transaction Type"::Purchase);
            DetailedGSTLedgerEntry3.SETRANGE("Document No.", StartingNo1, EndingNo1);
            IF DetailedGSTLedgerEntry3.FINDFIRST THEN
                RecordCount1 := DetailedGSTLedgerEntry3.COUNT;
            PurchPrint := TRUE;
        END;
    end;

    local procedure FilterPostingDate(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry")
    begin
        WITH DetailedGSTLedgerEntry DO
            CASE Period OF
                Period::Jan:
                    SETRANGE("Posting Date", DMY2DATE(1, 1, Year), DMY2DATE(31, 1, Year));
                Period::Feb:
                    SETRANGE("Posting Date", DMY2DATE(1, 2, Year), DMY2DATE(28, 2, Year));
                Period::March:
                    SETRANGE("Posting Date", DMY2DATE(1, 3, Year), DMY2DATE(31, 3, Year));
                Period::Apr:
                    SETRANGE("Posting Date", DMY2DATE(1, 4, Year), DMY2DATE(30, 4, Year));
                Period::May:
                    SETRANGE("Posting Date", DMY2DATE(1, 5, Year), DMY2DATE(31, 5, Year));
                Period::June:
                    SETRANGE("Posting Date", DMY2DATE(1, 6, Year), DMY2DATE(30, 6, Year));
                Period::July:
                    SETRANGE("Posting Date", DMY2DATE(1, 7, Year), DMY2DATE(31, 7, Year));
                Period::Agu:
                    SETRANGE("Posting Date", DMY2DATE(1, 8, Year), DMY2DATE(31, 8, Year));
                Period::Sept:
                    SETRANGE("Posting Date", DMY2DATE(1, 9, Year), DMY2DATE(30, 9, Year));
                Period::Oct:
                    SETRANGE("Posting Date", DMY2DATE(1, 10, Year), DMY2DATE(31, 10, Year));
                Period::Nov:
                    //SETRANGE("Posting Date",DMY2DATE(1,11,Year),DMY2DATE(30,3,Year));
                    SETRANGE("Posting Date", DMY2DATE(1, 11, Year), DMY2DATE(30, 11, Year));//Bug in GSTR-1 08Jan2018
                Period::Dec:
                    SETRANGE("Posting Date", DMY2DATE(1, 12, Year), DMY2DATE(31, 12, Year));
            END;
    end;

    local procedure GetGoodsAndService(var DetailedGSTLedgerEntry: Record "Detailed GST Ledger Entry"; var GoodService: Text)
    var
        Item: Record 27;
        GLAccount: Record 15;
        FixedAsset: Record 5600;
        Resource: Record 156;
    begin
        IF DetailedGSTLedgerEntry.Type = DetailedGSTLedgerEntry.Type::Item THEN BEGIN
            Item.GET(DetailedGSTLedgerEntry."No.");
            GoodService := Item.Description;
        END;
        IF DetailedGSTLedgerEntry.Type = DetailedGSTLedgerEntry.Type::"G/L Account" THEN BEGIN
            GLAccount.GET(DetailedGSTLedgerEntry."No.");
            GoodService := GLAccount.Name;
        END;
        IF DetailedGSTLedgerEntry.Type = DetailedGSTLedgerEntry.Type::"Fixed Asset" THEN BEGIN
            FixedAsset.GET(DetailedGSTLedgerEntry."No.");
            GoodService := FixedAsset.Description;
        END;
        IF DetailedGSTLedgerEntry.Type = DetailedGSTLedgerEntry.Type::Resource THEN BEGIN
            Resource.GET(DetailedGSTLedgerEntry."No.");
            GoodService := Resource.Name;
        END;
    end;

    local procedure ClearData()
    var
        J: Integer;
    begin
        FOR J := 1 TO 30 DO BEGIN
            GSTCompRate[J] := 0;
            GSTCompAmt[J] := 0;
            GSTBaseAmt[J] := 0;
        END;
    end;
}


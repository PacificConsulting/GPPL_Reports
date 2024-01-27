report 50106 "Posted Voucher IPOL"
{
    // //12Apr2017 RB-N Report Migartion as per NAV2009
    // //07Jun2017 RB-N ResCode07-->=Fields!ResDiv9.Value
    // //10July2017 :: RB-N, AccName Value increased to 80 Length (Oldvalue to 50 Length)
    // //03Nov2017  :: RB-N, Journal Voucher with multiple no Post & Print.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PostedVoucherIPOL.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Posted Voucher';

    dataset
    {
        dataitem("G/L Entry"; "G/L Entry")
        {
            DataItemTableView = SORTING("Document No.", "Posting Date", Amount)
                                ORDER(Descending);
            RequestFilterFields = "Posting Date", "Document No.";
            column(UserName; UserName)
            {
            }
            column(VoucherSourceDesc; SourceDesc + ' Voucher')
            {
            }
            column(DocumentNo_GLEntry; "Document No.")
            {
            }
            column(PostingDateFormatted; 'Date: ' + FORMAT("Posting Date"))
            {
            }
            column(CompanyInformationAddress; CompanyInformation.Address + ' ' + CompanyInformation."Address 2" + '  ' + CompanyInformation.City)
            {
            }
            column(CompanyInformationName; CompanyInformation.Name)
            {
            }
            column(CreditAmount_GLEntry; "Credit Amount")
            {
            }
            column(DebitAmount_GLEntry; "Debit Amount")
            {
            }
            column(DrText; DrText)
            {
            }
            column(GLAccName; GLAccName)
            {
            }
            column(AccountName26; AccountName26)
            {
            }
            column(CrText; CrText)
            {
            }
            column(DebitAmountTotal; DebitAmountTotal)
            {
            }
            column(CreditAmountTotal; CreditAmountTotal)
            {
            }
            column(ChequeDetail; 'Cheque No: ' + ChequeNo + '  Dated: ' + FORMAT(ChequeDate))
            {
            }
            column(ChequeNo; ChequeNo)
            {
            }
            column(ChequeDate; ChequeDate)
            {
            }
            column(RsNumberText1NumberText2; 'Rupees  ' + NumberText[1] + ' ' + NumberText[2])
            {
            }
            column(EntryNo_GLEntry; "Entry No.")
            {
            }
            column(PostingDate_GLEntry; "Posting Date")
            {
            }
            column(TransactionNo_GLEntry; "Transaction No.")
            {
            }
            column(VoucherNoCaption; VoucherNoCaptionLbl)
            {
            }
            column(CreditAmountCaption; CreditAmountCaptionLbl)
            {
            }
            column(DebitAmountCaption; DebitAmountCaptionLbl)
            {
            }
            column(ParticularsCaption; ParticularsCaptionLbl)
            {
            }
            column(AmountInWordsCaption; AmountInWordsCaptionLbl)
            {
            }
            column(PreparedByCaption; PreparedByCaptionLbl)
            {
            }
            column(CheckedByCaption; CheckedByCaptionLbl)
            {
            }
            column(ApprovedByCaption; ApprovedByCaptionLbl)
            {
            }
            column(VendorIvnNo; 'Vendor Invoice No :- ' + VendorInvoNo)
            {
            }
            column(DisNav9; "Disc&IndAdj&AutoAdjDim1")
            {
            }
            column(StafflNav9; DimensionValue)
            {
            }
            column(DivNav9; vDimesion)
            {
            }
            column(ResDiv9; "Global Dimension 2 Code")
            {
            }
            column(ResCode07; ResCode07)
            {
            }
            dataitem("Dimension Set Entry"; "Dimension Set Entry")
            {
                DataItemLink = "Dimension Set ID" = FIELD("Dimension Set ID");
                DataItemTableView = SORTING("Dimension Set ID", "Dimension Code");
                column(DimensionValueName_DimensionSetEntry; "Dimension Set Entry"."Dimension Value Name")
                {
                }
                column(DimensionSetID_DimensionSetEntry; "Dimension Set Entry"."Dimension Set ID")
                {
                }
                column(DimensionCode_DimensionSetEntry; "Dimension Set Entry"."Dimension Code")
                {
                }
                column(DimensionValueCode_DimensionSetEntry; "Dimension Set Entry"."Dimension Value Code")
                {
                }
                column(DimensionValueID_DimensionSetEntry; "Dimension Set Entry"."Dimension Value ID")
                {
                }
                column(DimensionName_DimensionSetEntry; "Dimension Set Entry"."Dimension Name")
                {
                }
                column(DepositEMD; DepositEMDcode + ' - ' + DepositEMDName)
                {
                }
                column(DepositDealer; DepositDealerCode + ' - ' + DepositDealerName)
                {
                }
                column(DepositOther; DepositOthersCode + ' - ' + DepositOthersName)
                {
                }
                column(DepositRent; DepositRentCode + ' - ' + DepositRentName)
                {
                }
                column(DepositReceived; DepositReceivedCode + ' - ' + DepositReceivedName)
                {
                }
                column(DepositSales; DepostiSalesTaxCode + ' - ' + DepostiSalesTaxName)
                {
                }
                column(PhoneDesc; PhoneNo + ' - ' + PhoneNoDesc)
                {
                }
                column(PhoneNo; PhoneNo)
                {
                }
                column(VehicleDesc; VehicleCode + '   ' + VehicaleName)
                {
                }
                column(SalesPromoName; SalesPromotionCode + ' - ' + SalesPromotionName)
                {
                }
                column(StaffName; StaffDimensionCode + '   ' + StaffDimensionName)
                {
                }
                column(ResDescc; "Res.Code" + '  ' + "Res.CenterName")
                {
                }
                column(DivisionName; DivisionName)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>NAV2009
                    /*
                    IF ("G/L Entry 1"."Source Code" = 'JOURNALV') OR ("G/L Entry 1"."Source Code" = 'CASHPYMTV')
                       OR ("G/L Entry 1"."Source Code" = 'BANKRCPTV') THEN
                    BEGIN
                     //----------------------- FOR Division------------START
                     RecDimensionValue.RESET;
                     RecDimensionValue.SETRANGE(RecDimensionValue."Dimension Code",'DIVISION');
                     RecDimensionValue.SETRANGE(RecDimensionValue.Code,"G/L Entry 1"."Global Dimension 1 Code");
                     IF RecDimensionValue.FINDFIRST THEN
                     BEGIN
                      DivisionName := RecDimensionValue.Name;
                     END;
                     //----------------------- FOR Division---------------END
                     //----------------------- FOR Res. Center------------START
                     "Res.CenterCode" :="G/L Entry 1"."Global Dimension 2 Code";
                     RecDimensionValue.RESET;
                     RecDimensionValue.SETRANGE(RecDimensionValue."Dimension Code",'RESPCENTER');
                     RecDimensionValue.SETRANGE(RecDimensionValue.Code,"G/L Entry 1"."Global Dimension 2 Code");
                     IF RecDimensionValue.FINDFIRST THEN
                     BEGIN
                      "Res.CenterName" := RecDimensionValue.Name;
                     END;
                     //----------------------- FOR Res. Center---------------END
                     //----------------------- FOR STAFF------------START
                     EBTuser.GET(USERID);
                     IF EBTuser."User ID" <> 'SA' THEN
                     BEGIN
                      IF ("G/L Entry 1"."G/L Account No." = '73310010') OR ("G/L Entry 1"."G/L Account No." = '73310020') OR
                         ("G/L Entry 1"."G/L Account No." = '73310030') OR ("G/L Entry 1"."G/L Account No." = '73310040') OR
                         ("G/L Entry 1"."G/L Account No." = '73310050') OR ("G/L Entry 1"."G/L Account No." = '73310060') OR
                         ("G/L Entry 1"."G/L Account No." = '73310070') OR ("G/L Entry 1"."G/L Account No." = '32160400') OR
                         ("G/L Entry 1"."G/L Account No." = '12913010') OR ("G/L Entry 1"."G/L Account No." = '32160100') OR
                         ("G/L Entry 1"."G/L Account No." = '32160600') OR ("G/L Entry 1"."G/L Account No." = '32160900') THEN
                       BEGIN
                        StaffDimensionCode := '';
                        StaffDimensionName := '';
                       END
                       ELSE
                       BEGIN
                        StaffDimensionCode := '';
                        DocumentDimension.RESET;
                        DocumentDimension.SETRANGE(DocumentDimension."Table ID",17);
                        DocumentDimension.SETRANGE(DocumentDimension."Entry No.","G/L Entry 1"."Entry No.");
                        DocumentDimension.SETRANGE(DocumentDimension."Dimension Code",'STAFF');
                        IF DocumentDimension.FINDFIRST THEN
                         StaffDimensionCode := DocumentDimension."Dimension Value Code";
                         REPEAT
                          StaffDimensionName := '';
                          RecDimensionValue.RESET;
                          RecDimensionValue.SETRANGE(RecDimensionValue."Dimension Code",'STAFF');
                          RecDimensionValue.SETRANGE(RecDimensionValue.Code,StaffDimensionCode);
                          IF RecDimensionValue.FINDFIRST THEN
                           StaffDimensionName := RecDimensionValue.Name;
                         UNTIL DocumentDimension.NEXT = 0;
                       END;
                     END
                     ELSE
                     BEGIN
                      StaffDimensionCode := '';
                      DocumentDimension.RESET;
                      DocumentDimension.SETRANGE(DocumentDimension."Table ID",17);
                      DocumentDimension.SETRANGE(DocumentDimension."Entry No.","G/L Entry 1"."Entry No.");
                      DocumentDimension.SETRANGE(DocumentDimension."Dimension Code",'STAFF');
                      IF DocumentDimension.FINDFIRST THEN
                       StaffDimensionCode := DocumentDimension."Dimension Value Code";
                       REPEAT
                        StaffDimensionName := '';
                        RecDimensionValue.RESET;
                        RecDimensionValue.SETRANGE(RecDimensionValue."Dimension Code",'STAFF');
                        RecDimensionValue.SETRANGE(RecDimensionValue.Code,StaffDimensionCode);
                        IF RecDimensionValue.FINDFIRST THEN
                         StaffDimensionName := RecDimensionValue.Name;
                       UNTIL DocumentDimension.NEXT = 0;
                     END;
                     //END;
                     //----------------------- FOR STAFF--------------END
                     //----------------------- FOR DEPOSIT EMD------------START
                     DepositEMDcode := '';
                     DocumentDimension1.RESET;
                     DocumentDimension1.SETRANGE(DocumentDimension1."Table ID",17);
                     DocumentDimension1.SETRANGE(DocumentDimension1."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension1.SETRANGE(DocumentDimension1."Dimension Code",'DEPOSIT-EMD');
                     IF DocumentDimension1.FINDFIRST THEN
                      DepositEMDcode := DocumentDimension1."Dimension Value Code";
                      REPEAT
                       DepositEMDName := '';
                       RecDimensionValue1.RESET;
                       RecDimensionValue1.SETRANGE(RecDimensionValue1."Dimension Code",'DEPOSIT-EMD');
                       RecDimensionValue1.SETRANGE(RecDimensionValue1.Code,DepositEMDcode);
                       IF RecDimensionValue1.FINDFIRST THEN
                        DepositEMDName := RecDimensionValue1.Name;
                      UNTIL DocumentDimension1.NEXT = 0;
                     //----------------------- FOR DEPOSIT EMD--------------END
                     //----------------------- FOR DEPOSIT DEALER------------START
                     DepositDealerCode := '';
                     DocumentDimension2.RESET;
                     DocumentDimension2.SETRANGE(DocumentDimension2."Table ID",17);
                     DocumentDimension2.SETRANGE(DocumentDimension2."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension2.SETRANGE(DocumentDimension2."Dimension Code",'DEPOSIT-DEALER');
                     IF DocumentDimension2.FINDFIRST THEN
                      DepositDealerCode := DocumentDimension2."Dimension Value Code";
                      REPEAT
                       DepositDealerName := '';
                       RecDimensionValue2.RESET;
                       RecDimensionValue2.SETRANGE(RecDimensionValue2."Dimension Code",'DEPOSIT-DEALER');
                       RecDimensionValue2.SETRANGE(RecDimensionValue2.Code,DepositDealerCode);
                       IF RecDimensionValue2.FINDFIRST THEN
                        DepositDealerName := RecDimensionValue2.Name;
                      UNTIL DocumentDimension2.NEXT = 0;
                     //----------------------- FOR DEPOSIT DEALER--------------END
                     //----------------------- FOR DEPOSIT OTHERS------------START
                     DepositOthersCode := '';
                     DocumentDimension3.RESET;
                     DocumentDimension3.SETRANGE(DocumentDimension3."Table ID",17);
                     DocumentDimension3.SETRANGE(DocumentDimension3."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension3.SETRANGE(DocumentDimension3."Dimension Code",'DEPOSIT-OTHERS');
                     IF DocumentDimension3.FINDFIRST THEN
                      DepositOthersCode := DocumentDimension3."Dimension Value Code";
                      REPEAT
                       DepositOthersName := '';
                       RecDimensionValue3.RESET;
                       RecDimensionValue3.SETRANGE(RecDimensionValue3."Dimension Code",'DEPOSIT-OTHERS');
                       RecDimensionValue3.SETRANGE(RecDimensionValue3.Code,DepositOthersCode);
                       IF RecDimensionValue3.FINDFIRST THEN
                        DepositOthersName := RecDimensionValue3.Name;
                      UNTIL DocumentDimension3.NEXT = 0;
                     //----------------------- FOR DEPOSIT OTHERS--------------END
                     //----------------------- FOR DEPOSIT RENT------------START
                     DepositRentCode := '';
                     DocumentDimension4.RESET;
                     DocumentDimension4.SETRANGE(DocumentDimension4."Table ID",17);
                     DocumentDimension4.SETRANGE(DocumentDimension4."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension4.SETRANGE(DocumentDimension4."Dimension Code",'DEPOSIT-RENT');
                     IF DocumentDimension4.FINDFIRST THEN
                      DepositRentCode := DocumentDimension4."Dimension Value Code";
                      REPEAT
                       DepositRentName := '';
                       RecDimensionValue4.RESET;
                       RecDimensionValue4.SETRANGE(RecDimensionValue4."Dimension Code",'DEPOSIT-RENT');
                       RecDimensionValue4.SETRANGE(RecDimensionValue4.Code,DepositRentCode);
                       IF RecDimensionValue4.FINDFIRST THEN
                        DepositRentName := RecDimensionValue4.Name;
                      UNTIL DocumentDimension4.NEXT = 0;
                     //----------------------- FOR DEPOSIT RENT--------------END
                     //----------------------- FOR PHONE------------START
                     PhoneNo := '';
                     DocumentDimension5.RESET;
                     DocumentDimension5.SETRANGE(DocumentDimension5."Table ID",17);
                     DocumentDimension5.SETRANGE(DocumentDimension5."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension5.SETRANGE(DocumentDimension5."Dimension Code",'PHONE');
                     IF DocumentDimension5.FINDFIRST THEN
                      PhoneNo := DocumentDimension5."Dimension Value Code";
                      REPEAT
                       PhoneNoDesc := '';
                       RecDimensionValue5.RESET;
                       RecDimensionValue5.SETRANGE(RecDimensionValue5."Dimension Code",'PHONE');
                       RecDimensionValue5.SETRANGE(RecDimensionValue5.Code,PhoneNo);
                       IF RecDimensionValue5.FINDFIRST THEN
                        PhoneNoDesc := RecDimensionValue5.Name;
                      UNTIL DocumentDimension5.NEXT = 0;
                     //----------------------- FOR PHONE--------------END
                     //----------------------- FOR VEHICLE------------START
                     VehicleCode := '';
                     DocumentDimension6.RESET;
                     DocumentDimension6.SETRANGE(DocumentDimension6."Table ID",17);
                     DocumentDimension6.SETRANGE(DocumentDimension6."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension6.SETRANGE(DocumentDimension6."Dimension Code",'VEHICLE');
                     IF DocumentDimension6.FINDFIRST THEN
                      VehicleCode := DocumentDimension6."Dimension Value Code";
                      REPEAT
                       VehicaleName := '';
                       RecDimensionValue6.RESET;
                       RecDimensionValue6.SETRANGE(RecDimensionValue6."Dimension Code",'VEHICLE');
                       RecDimensionValue6.SETRANGE(RecDimensionValue6.Code,VehicleCode);
                       IF RecDimensionValue6.FINDFIRST THEN
                        VehicaleName := RecDimensionValue6.Name;
                      UNTIL DocumentDimension6.NEXT = 0;
                      //----------------------- FOR VEHICLE--------------END
                     //----------------------- FOR DEPOSIT RECEIVED------------START
                     DepositReceivedCode := '';
                     DocumentDimension9.RESET;
                     DocumentDimension9.SETRANGE(DocumentDimension9."Table ID",17);
                     DocumentDimension9.SETRANGE(DocumentDimension9."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension9.SETRANGE(DocumentDimension9."Dimension Code",'DEPOSIT-RECEIVED');
                     IF DocumentDimension9.FINDFIRST THEN
                      DepositReceivedCode := DocumentDimension9."Dimension Value Code";
                      REPEAT
                       DepositReceivedName := '';
                       RecDimensionValue9.RESET;
                       RecDimensionValue9.SETRANGE(RecDimensionValue9."Dimension Code",'DEPOSIT-RECEIVED');
                       RecDimensionValue9.SETRANGE(RecDimensionValue9.Code,DepositReceivedCode);
                       IF RecDimensionValue9.FINDFIRST THEN
                        DepositReceivedName := RecDimensionValue9.Name;
                      UNTIL DocumentDimension9.NEXT = 0;
                     //----------------------- FOR DEPOSIT RECEIVED--------------END
                     //----------------------- FOR DEPOSIT SALES TAX------------START
                     DepostiSalesTaxCode := '';
                     DocumentDimension10.RESET;
                     DocumentDimension10.SETRANGE(DocumentDimension10."Table ID",17);
                     DocumentDimension10.SETRANGE(DocumentDimension10."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension10.SETRANGE(DocumentDimension10."Dimension Code",'DEPOSIT-SALESTAX');
                     IF DocumentDimension10.FINDFIRST THEN
                      DepostiSalesTaxCode := DocumentDimension10."Dimension Value Code";
                      REPEAT
                       DepostiSalesTaxName := '';
                       RecDimensionValue10.RESET;
                       RecDimensionValue10.SETRANGE(RecDimensionValue10."Dimension Code",'DEPOSIT-SALESTAX');
                       RecDimensionValue10.SETRANGE(RecDimensionValue10.Code,DepostiSalesTaxCode);
                       IF RecDimensionValue10.FINDFIRST THEN
                        DepostiSalesTaxName := RecDimensionValue10.Name;
                      UNTIL DocumentDimension10.NEXT = 0;
                     //----------------------- FOR DEPOSIT SALES TAX--------------END
                     //----------------------- FOR SALES PROMOTION------------START
                     SalesPromotionCode := '';
                     DocumentDimension11.RESET;
                     DocumentDimension11.SETRANGE(DocumentDimension11."Table ID",17);
                     DocumentDimension11.SETRANGE(DocumentDimension11."Entry No.","G/L Entry 1"."Entry No.");
                     DocumentDimension11.SETRANGE(DocumentDimension11."Dimension Code",'SALESPROMO');
                     IF DocumentDimension11.FINDFIRST THEN
                      SalesPromotionCode := DocumentDimension11."Dimension Value Code";
                      REPEAT
                       SalesPromotionName := '';
                       RecDimensionValue11.RESET;
                       RecDimensionValue11.SETRANGE(RecDimensionValue11."Dimension Code",'SALESPROMO');
                       RecDimensionValue11.SETRANGE(RecDimensionValue11.Code,SalesPromotionCode);
                       IF RecDimensionValue11.FINDFIRST THEN
                        SalesPromotionName := RecDimensionValue11.Name;
                      UNTIL DocumentDimension10.NEXT = 0;
                     //----------------------- FOR SALES PROMOTION--------------END
                    
                    
                    END;
                    */
                    //<<NAV2009


                    //>>05Apr2017 DimensionSetEntry

                    /*
                    IF ("G/L Entry 1"."Source Code" = 'JOURNALV') OR ("G/L Entry 1"."Source Code" = 'CASHPYMTV')
                    OR ("G/L Entry 1"."Source Code" = 'BANKRCPTV') THEN
                    BEGIN
                    */

                    IF ("G/L Entry"."Source Code" = 'JOURNALV') OR ("G/L Entry"."Source Code" = 'CASHPYMTV')
                    OR ("G/L Entry"."Source Code" = 'BANKRCPTV') THEN BEGIN

                        //----------------------- FOR Division------------START
                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'DIVISION');
                        //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 1 Code");
                        recDimSet.SETRANGE("Dimension Value Code", "G/L Entry"."Global Dimension 1 Code");//02Nov2017
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DivisionName := recDimSet."Dimension Value Name";

                        END;
                        //----------------------- FOR Division---------------END
                        //"Res.Code" := "G/L Entry 1"."Global Dimension 2 Code";
                        "Res.Code" := "G/L Entry"."Global Dimension 2 Code";//02Nov2017

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'RESPCENTER');
                        //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                        recDimSet.SETRANGE("Dimension Value Code", "G/L Entry"."Global Dimension 2 Code");//02Nov2017
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            "Res.CenterName" := recDimSet."Dimension Value Name";

                        END;
                        //----------------------- FOR Res. Center------------START


                        //----------------------- FOR STAFF------------START
                        //EBTuser.GET(USERID);
                        //IF EBTuser."User ID" <> 'SA' THEN
                        IF EBTuser."User Name" <> 'SA' THEN BEGIN
                            IF ("G/L Entry"."G/L Account No." = '73310010') OR ("G/L Entry"."G/L Account No." = '73310020') OR
                               ("G/L Entry"."G/L Account No." = '73310030') OR ("G/L Entry"."G/L Account No." = '73310040') OR
                               ("G/L Entry"."G/L Account No." = '73310050') OR ("G/L Entry"."G/L Account No." = '73310060') OR
                               ("G/L Entry"."G/L Account No." = '73310070') OR ("G/L Entry"."G/L Account No." = '32160400') OR
                               ("G/L Entry"."G/L Account No." = '12913010') OR ("G/L Entry"."G/L Account No." = '32160100') OR
                               ("G/L Entry"."G/L Account No." = '32160600') OR ("G/L Entry"."G/L Account No." = '32160900') THEN BEGIN
                                StaffDimensionCode := '';
                                StaffDimensionName := '';
                            END
                            ELSE BEGIN
                                StaffDimensionCode := '';

                                recDimSet.RESET;
                                recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                recDimSet.SETRANGE("Dimension Code", 'STAFF');
                                IF recDimSet.FINDFIRST THEN BEGIN
                                    StaffDimensionCode := recDimSet."Dimension Value Code";
                                    recDimSet.CALCFIELDS("Dimension Value Name");
                                    StaffDimensionName := recDimSet."Dimension Value Name";
                                END;
                            END;

                        END ELSE BEGIN


                            StaffDimensionCode := '';


                            recDimSet.RESET;
                            recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                            recDimSet.SETRANGE("Dimension Code", 'STAFF');
                            IF recDimSet.FINDFIRST THEN BEGIN
                                StaffDimensionCode := recDimSet."Dimension Value Code";
                                recDimSet.CALCFIELDS("Dimension Value Name");
                                StaffDimensionName := recDimSet."Dimension Value Name";
                            END;
                        END;
                        //----------------------- FOR STAFF--------------END

                        //----------------------- FOR DEPOSIT EMD------------START
                        DepositEMDcode := '';
                        DepositEMDName := '';
                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-EMD');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            DepositEMDcode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DepositEMDName := recDimSet."Dimension Value Name";

                        END;
                        //----------------------- FOR DEPOSIT EMD--------------END



                        //----------------------- FOR DEPOSIT DEALER------------START
                        DepositDealerCode := '';
                        DepositDealerName := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-DEALER');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            DepositDealerCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DepositDealerName := recDimSet."Dimension Value Name";

                        END;
                        //----------------------- FOR DEPOSIT DEALER--------------END


                        //----------------------- FOR DEPOSIT OTHERS------------START
                        DepositOthersCode := '';
                        DepositOthersName := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-OTHERS');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            DepositOthersCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DepositOthersName := recDimSet."Dimension Value Name";

                        END;
                        //----------------------- FOR DEPOSIT OTHERS--------------END


                        //----------------------- FOR DEPOSIT RENT------------START
                        DepositRentCode := '';
                        DepositRentName := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-RENT');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            DepositRentCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DepositRentName := recDimSet."Dimension Value Name";
                            //DepositRentCode := recDimSet."Dimension Value Name";

                        END;
                        //----------------------- FOR DEPOSIT RENT--------------END


                        //----------------------- FOR PHONE------------START
                        PhoneNo := '';
                        PhoneNoDesc := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'PHONE');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            PhoneNo := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            PhoneNoDesc := recDimSet."Dimension Value Name";
                        END;

                        //----------------------- FOR PHONE--------------END


                        //----------------------- FOR VEHICLE------------START
                        VehicleCode := '';
                        VehicaleName := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'VEHICLE');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            VehicleCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            VehicaleName := recDimSet."Dimension Value Name";
                        END;
                        //----------------------- FOR VEHICLE--------------END


                        //----------------------- FOR DEPOSIT RECEIVED------------START
                        DepositReceivedCode := '';
                        DepositReceivedName := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-RECEIVED');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            DepositReceivedCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DepositReceivedName := recDimSet."Dimension Value Name";
                        END;

                        //----------------------- FOR DEPOSIT RECEIVED--------------END


                        //----------------------- FOR DEPOSIT SALES TAX------------START
                        DepostiSalesTaxCode := '';
                        DepostiSalesTaxName := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-SALESTAX');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            DepostiSalesTaxCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DepostiSalesTaxName := recDimSet."Dimension Value Name";
                        END;
                        //----------------------- FOR DEPOSIT SALES TAX--------------END


                        //----------------------- FOR SALES PROMOTION------------START
                        SalesPromotionCode := '';
                        SalesPromotionName := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALESPROMO');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            SalesPromotionCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            SalesPromotionName := recDimSet."Dimension Value Name";
                        END;
                        //----------------------- FOR SALES PROMOTION--------------END

                    END;


                    /*
                    IF ("G/L Entry 1"."System-Created Entry") AND ("G/L Entry 1"."Source Code" = 'PURCHASES')
                    OR ("G/L Entry 1"."Source Code" = '' )THEN
                    BEGIN
                      IF "G/L Entry 1"."Source Type" = "G/L Entry 1"."Source Type" :: Vendor THEN
                      BEGIN
                      */
                    IF ("G/L Entry"."System-Created Entry") AND ("G/L Entry"."Source Code" = 'PURCHASES')
                    OR ("G/L Entry"."Source Code" = '') THEN BEGIN
                        IF "G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Vendor THEN BEGIN
                            recVendor.RESET;
                            recVendor.SETRANGE(recVendor."No.", "G/L Entry"."Source No.");
                            IF recVendor.FINDFIRST THEN
                                IF (recVendor."Vendor Posting Group" = 'EXPENSE') OR (recVendor."Vendor Posting Group" = 'CAPGOODS') OR
                                 (recVendor."Vendor Posting Group" = 'TRANSPORT') THEN BEGIN

                                    //----------------------- FOR Division------------START
                                    DivisionName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'DIVISION');
                                    //recDimSet.SETRANGE(recDimSet."Dimension Value Code","G/L Entry 1"."Global Dimension 1 Code");
                                    recDimSet.SETRANGE(recDimSet."Dimension Value Code", "G/L Entry"."Global Dimension 1 Code");//02Nov2017
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        DivisionName := recDimSet."Dimension Value Name";
                                    END;

                                    //----------------------- FOR Division---------------END


                                    //----------------------- FOR Res. Center------------START
                                    "Res.Code" := '';
                                    "Res.CenterName" := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'RESPCENTER');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    recDimSet.SETRANGE("Dimension Value Code", "G/L Entry"."Global Dimension 2 Code");//02Nov2017
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        "Res.Code" := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        "Res.CenterName" := recDimSet."Dimension Value Name";
                                    END;
                                    //----------------------- FOR Res. Center---------------END


                                    //----------------------- FOR STAFF------------START
                                    //EBTuser.GET(USERID);
                                    //IF EBTuser."User ID" <> 'SA' THEN
                                    IF EBTuser."User Name" <> 'SA' THEN BEGIN
                                        IF ("G/L Entry"."G/L Account No." = '73310010') OR ("G/L Entry"."G/L Account No." = '73310020') OR
                                          ("G/L Entry"."G/L Account No." = '73310030') OR ("G/L Entry"."G/L Account No." = '73310040') OR
                                          ("G/L Entry"."G/L Account No." = '73310050') OR ("G/L Entry"."G/L Account No." = '73310060') OR
                                          ("G/L Entry"."G/L Account No." = '73310070') OR ("G/L Entry"."G/L Account No." = '32160400') OR
                                          ("G/L Entry"."G/L Account No." = '12913010') OR ("G/L Entry"."G/L Account No." = '32160100') OR
                                          ("G/L Entry"."G/L Account No." = '32160600') OR ("G/L Entry"."G/L Account No." = '32160900') THEN BEGIN

                                            StaffDimensionCode := '';
                                            StaffDimensionName := '';
                                        END ELSE BEGIN

                                            StaffDimensionCode := '';
                                            StaffDimensionName := '';

                                            recDimSet.RESET;
                                            recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                            recDimSet.SETRANGE("Dimension Code", 'STAFF');
                                            //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                            IF recDimSet.FINDFIRST THEN BEGIN
                                                StaffDimensionCode := recDimSet."Dimension Value Code";
                                                recDimSet.CALCFIELDS("Dimension Value Name");
                                                StaffDimensionName := recDimSet."Dimension Value Name";
                                            END;

                                        END;

                                    END ELSE BEGIN

                                        StaffDimensionCode := '';
                                        StaffDimensionName := '';
                                        recDimSet.RESET;
                                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                        recDimSet.SETRANGE("Dimension Code", 'STAFF');
                                        //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                        IF recDimSet.FINDFIRST THEN BEGIN
                                            StaffDimensionCode := recDimSet."Dimension Value Code";
                                            recDimSet.CALCFIELDS("Dimension Value Name");
                                            StaffDimensionName := recDimSet."Dimension Value Name";
                                        END;
                                    END;

                                    //----------------------- FOR STAFF--------------END


                                    //----------------------- FOR DEPOSIT EMD------------START
                                    DepositEMDcode := '';
                                    DepositEMDName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-EMD');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        DepositEMDcode := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        DepositEMDName := recDimSet."Dimension Value Name";
                                    END;

                                    //----------------------- FOR DEPOSIT EMD--------------END


                                    //----------------------- FOR DEPOSIT DEALER------------START
                                    DepositDealerCode := '';
                                    DepositDealerName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-DEALER');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        DepositDealerCode := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        DepositDealerName := recDimSet."Dimension Value Name";
                                    END;

                                    //----------------------- FOR DEPOSIT DEALER--------------END


                                    //----------------------- FOR DEPOSIT OTHERS------------START
                                    DepositOthersCode := '';
                                    DepositOthersName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-OTHERS');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        DepositOthersCode := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        DepositOthersName := recDimSet."Dimension Value Name";
                                    END;

                                    //----------------------- FOR DEPOSIT OTHERS--------------END


                                    //----------------------- FOR DEPOSIT RENT------------START
                                    DepositRentCode := '';
                                    DepositRentName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-RENT');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        DepositRentCode := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        DepositRentName := recDimSet."Dimension Value Name";
                                    END;
                                    //----------------------- FOR DEPOSIT RENT--------------END


                                    //----------------------- FOR PHONE------------START
                                    PhoneNo := '';
                                    PhoneNoDesc := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'PHONE');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        PhoneNo := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        PhoneNoDesc := recDimSet."Dimension Value Name";
                                    END;

                                    //----------------------- FOR PHONE--------------END


                                    //----------------------- FOR VEHICLE------------START
                                    VehicleCode := '';
                                    VehicaleName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'VEHICLE');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        VehicleCode := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        VehicaleName := recDimSet."Dimension Value Name";
                                    END;

                                    //----------------------- FOR VEHICLE--------------END


                                    //----------------------- FOR DEPOSIT RECEIVED------------START
                                    DepositReceivedCode := '';
                                    DepositReceivedName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-RECEIVED');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        DepositReceivedCode := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        DepositReceivedName := recDimSet."Dimension Value Name";
                                    END;
                                    //----------------------- FOR DEPOSIT RECEIVED--------------END


                                    //----------------------- FOR DEPOSIT SALES TAX------------START
                                    DepostiSalesTaxCode := '';
                                    DepostiSalesTaxName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'DEPOSIT-SALESTAX');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        DepostiSalesTaxCode := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        DepostiSalesTaxName := recDimSet."Dimension Value Name";
                                    END;
                                    //----------------------- FOR DEPOSIT SALES TAX--------------END


                                    //----------------------- FOR SALES PROMOTION------------START
                                    SalesPromotionCode := '';
                                    SalesPromotionName := '';

                                    recDimSet.RESET;
                                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                                    recDimSet.SETRANGE("Dimension Code", 'SALESPROMO');
                                    //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                                    IF recDimSet.FINDFIRST THEN BEGIN
                                        SalesPromotionCode := recDimSet."Dimension Value Code";
                                        recDimSet.CALCFIELDS("Dimension Value Name");
                                        SalesPromotionName := recDimSet."Dimension Value Name";
                                    END;
                                    //----------------------- FOR SALES PROMOTION--------------END



                                END ELSE
                                    CurrReport.SKIP;
                        END;
                    END;



                    // EBT MILAN (25112013)....For Division and Rec. Center in case of credit memo and g/l acc. of Sales Promotion Expenses ....START
                    /*
                    IF ("G/L Entry 1"."Document Type" = "G/L Entry 1"."Document Type"::"Credit Memo") AND
                       ("G/L Entry 1"."G/L Account No." = '75002040') AND ("G/L Entry 1"."Source Code" ='SALES') THEN
                    BEGIN
                    */
                    IF ("G/L Entry"."Document Type" = "G/L Entry"."Document Type"::"Credit Memo") AND
                       ("G/L Entry"."G/L Account No." = '75002040') AND ("G/L Entry"."Source Code" = 'SALES') THEN BEGIN

                        //----------------------- FOR Division------------START
                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'DIVISION');
                        //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 1 Code");
                        recDimSet.SETRANGE("Dimension Value Code", "G/L Entry"."Global Dimension 1 Code");//02Nov2017
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DivisionName := recDimSet."Dimension Value Name";
                        END;

                        //----------------------- FOR Division---------------END


                        //----------------------- FOR Res. Center------------START
                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'RESPCENTER');
                        //recDimSet.SETRANGE("Dimension Value Code","G/L Entry 1"."Global Dimension 2 Code");
                        recDimSet.SETRANGE("Dimension Value Code", "G/L Entry"."Global Dimension 2 Code");//02Nov2017
                        IF recDimSet.FINDFIRST THEN BEGIN
                            "Res.Code" := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            "Res.CenterName" := recDimSet."Dimension Value Name";
                        END;
                        //----------------------- FOR Res. Center---------------END


                        SalesPromotionCode := '';
                        SalesPromotionName := '';

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'SALESPROMO');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            SalesPromotionCode := recDimSet."Dimension Value Code";
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            SalesPromotionName := recDimSet."Dimension Value Name";
                        END;



                    END;
                    // EBT MILAN (25112013)....For Division and Rec. Center in case of credit memo and g/l acc. of Sales Promotion Expenses ......END

                    //<<05Apr2017 DimensionSetEntry

                end;
            }
            dataitem(LineNarration; "Posted Narration")
            {
                DataItemLink = "Transaction No." = FIELD("Transaction No."),
                               "Entry No." = FIELD("Entry No.");
                DataItemTableView = SORTING("Entry No.", "Transaction No.", "Line No.");
                column(Narration_LineNarration; Narration)
                {
                }
                column(PrintLineNarration; PrintLineNarration)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    IF PrintLineNarration THEN BEGIN
                        PageLoop := PageLoop - 1;
                        LinesPrinted := LinesPrinted + 1;
                    END;
                end;
            }
            dataitem(Integer; Integer)
            {
                DataItemTableView = SORTING(Number);
                column(IntegerOccurcesCaption; IntegerOccurcesCaptionLbl)
                {
                }

                trigger OnAfterGetRecord()
                begin
                    PageLoop := PageLoop - 1;
                end;

                trigger OnPreDataItem()
                begin
                    GLEntry.SETCURRENTKEY("Document No.", "Posting Date", Amount);
                    GLEntry.ASCENDING(FALSE);
                    GLEntry.SETRANGE("Posting Date", "G/L Entry"."Posting Date");
                    GLEntry.SETRANGE("Document No.", "G/L Entry"."Document No.");
                    GLEntry.FINDLAST;
                    IF NOT (GLEntry."Entry No." = "G/L Entry"."Entry No.") THEN
                        CurrReport.BREAK;

                    SETRANGE(Number, 1, PageLoop)
                end;
            }
            dataitem(PostedNarration1; "Posted Narration")
            {
                DataItemLink = "Transaction No." = FIELD("Transaction No.");
                DataItemTableView = SORTING("Entry No.", "Transaction No.", "Line No.")
                                    WHERE("Entry No." = FILTER(0));
                column(Narration_PostedNarration1; Narration)
                {
                }
                column(NarrationCaption; NarrationCaptionLbl)
                {
                }

                trigger OnPreDataItem()
                begin
                    GLEntry.SETCURRENTKEY("Document No.", "Posting Date", Amount);
                    GLEntry.ASCENDING(FALSE);
                    GLEntry.SETRANGE("Posting Date", "G/L Entry"."Posting Date");
                    GLEntry.SETRANGE("Document No.", "G/L Entry"."Document No.");
                    GLEntry.FINDLAST;
                    IF NOT (GLEntry."Entry No." = "G/L Entry"."Entry No.") THEN
                        CurrReport.BREAK;
                end;
            }
            dataitem("Purch. Comment Line"; "Purch. Comment Line")
            {
                DataItemLink = "No." = FIELD("Document No.");
                DataItemTableView = SORTING("Document Type", "No.", "Line No.");
                column(PurComDocNo; "Purch. Comment Line"."No.")
                {
                }
                column(PurLineNo; "Line No.")
                {
                }
                column(PurchComm; "Purch. Comment Line".Comment)
                {
                }
            }
            dataitem("Sales Comment Line"; "Sales Comment Line")
            {
                DataItemLink = "No." = FIELD("Document No.");
                DataItemTableView = SORTING("Document Type", "No.", "Line No.");
                column(SalComDocNo; "Sales Comment Line"."No.")
                {
                }
                column(SalLinNo; "Line No.")
                {
                }
                column(SalesComm; "Sales Comment Line".Comment)
                {
                }
            }

            trigger OnAfterGetRecord()
            begin
                /*
                //>>RSPL/Rahul/PostMigration***Code added to block interim entries
                IF ("G/L Entry"."G/L Account No." = '12912050') OR ("G/L Entry"."G/L Account No." = '12912060') THEN
                  //("G/L Entry"."G/L Account No." = '75001090') OR ("G/L Entry"."G/L Account No." = '75001090')THEN    //To be view in future
                   CurrReport.SKIP;
                //<<
                *///commented 15May2017


                GLAccName := FindGLAccName("Source Type", "Entry No.", "Source No.", "G/L Account No.");

                //>>26July2017 AccountName

                CLEAR(AccountName26);
                IF "System-Created Entry" AND ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Vendor) THEN BEGIN
                    Vend.GET("G/L Entry"."Source No.");
                    AccountName26 := Vend.Name;
                END;

                IF "System-Created Entry" AND ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Customer) THEN BEGIN
                    Cust.GET("G/L Entry"."Source No.");
                    AccountName26 := Cust.Name;
                END;

                IF "System-Created Entry" AND ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::"Fixed Asset") THEN BEGIN
                    FA.GET("G/L Entry"."Source No.");
                    AccountName26 := FA.Description;
                END;

                IF "System-Created Entry" AND ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::"Bank Account") THEN BEGIN
                    BankAcc.GET("G/L Entry"."Source No.");
                    AccountName26 := BankAcc.Name;
                END;

                IF "G/L Entry"."Bal. Account Type" = "G/L Entry"."Bal. Account Type"::"G/L Account" THEN BEGIN
                    GL26.GET("G/L Entry"."G/L Account No.");
                    AccountName26 := GL26.Name;
                END;

                //VendorName
                IF ("G/L Account No." = '32120300') OR ("G/L Account No." = '32120400')
                OR ("G/L Account No." = '32120100') OR ("G/L Account No." = '32120200')
                OR ("G/L Account No." = '32120500') OR ("G/L Account No." = '32120600')
                OR ("G/L Account No." = '32120700') THEN BEGIN

                    IF ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Vendor) THEN BEGIN
                        //MESSAGE('%1 --',"G/L Entry"."Source Type"::Vendor);//07Sep
                        Vend.GET("G/L Entry"."Source No.");
                        AccountName26 := Vend.Name;
                    END;
                END;


                //CustomerName
                IF ("G/L Account No." = '12710000') OR ("G/L Account No." = '12720000') OR ("G/L Account No." = '12730000')
                OR ("G/L Account No." = '12740000') OR ("G/L Account No." = '12750000') OR ("G/L Account No." = '12760000')
                OR ("G/L Account No." = '12770000') OR ("G/L Account No." = '12780000') OR ("G/L Account No." = '12790000')
                OR ("G/L Account No." = '12791000') OR ("G/L Account No." = '12792000') THEN BEGIN
                    IF ("G/L Entry"."Source Type" = "G/L Entry"."Source Type"::Customer) THEN BEGIN
                        Cust.GET("G/L Entry"."Source No.");
                        AccountName26 := Cust.Name;
                    END;
                END;

                IF AccountName26 = '' THEN BEGIN
                    //MESSAGE('%1 --',"G/L Entry"."Source Type"::Vendor);//07Sep2017
                    "G/L Entry".CALCFIELDS("G/L Account Name");//07Sep2017
                    AccountName26 := "G/L Entry"."G/L Account Name";//07Sep2017
                                                                    //AccountName26 := "G/L Entry".Description;//01Aug2017
                END;
                //<<26July2017 AccountName

                IF Amount < 0 THEN BEGIN
                    CrText := 'To';
                    DrText := '';
                END ELSE BEGIN
                    CrText := '';
                    DrText := 'Dr';
                END;

                SourceDesc := '';
                IF "Source Code" <> '' THEN BEGIN
                    SourceCode.GET("Source Code");
                    SourceDesc := SourceCode.Description;
                END;


                //>>15May2017 Block only for BankPayment & BankReceipt
                IF ("Source Code" = 'BANKPYMTV') OR ("Source Code" = 'BANKRCPTV') THEN BEGIN
                    IF ("G/L Entry"."G/L Account No." = '12912050') OR ("G/L Entry"."G/L Account No." = '12912060') THEN
                        //("G/L Entry"."G/L Account No." = '75001090') OR ("G/L Entry"."G/L Account No." = '75001090')THEN    //To be view in future
                        CurrReport.SKIP;

                END;
                //<< 15May2017

                //>>11Apr2017 recUser
                recUser.RESET;
                recUser.SETRANGE("User Name", "User ID");
                IF recUser.FINDFIRST THEN BEGIN
                    UserName := recUser."Full Name";

                END;
                //<<11Apr2017 recUser

                //>>NAV2009 05Apr2017

                //EBT STIVAN ---(27072012)--- To Capture the heading of Voucher when its Credit Memo ------START
                IF ("G/L Entry"."Source Code" = 'SALES') AND ("G/L Entry"."Document Type" = "G/L Entry"."Document Type"::"Credit Memo") THEN BEGIN
                    SourceCode.GET("Source Code");
                    SourceDesc := SourceCode.Description + ' ' + 'Credit Memo';
                END;
                //EBT STIVAN ---(27072012)--- To Capture the heading of Voucher when its Credit Memo --------END

                //EBT STIVAN ---(27072012)--- To Capture the heading of Voucher when its Debit Memo ------START
                IF ("G/L Entry"."Source Code" = 'SALES') AND ("G/L Entry"."Document Type" = "G/L Entry"."Document Type"::Invoice) THEN BEGIN
                    recSIH.RESET;
                    recSIH.SETRANGE(recSIH."No.", "G/L Entry"."Document No.");
                    recSIH.SETRANGE(recSIH."Order No.", '');
                    recSIH.SETRANGE(recSIH."Supplimentary Invoice", FALSE);
                    IF recSIH.FINDFIRST THEN BEGIN
                        SourceDesc := 'Sales Debit Memo';
                    END;
                END;
                //EBT STIVAN ---(27072012)--- To Capture the heading of Voucher when its Debit Memo --------END

                //EBT STIVAN ---(30072012)--- To Capture the heading of Voucher when its Reversed ------START
                IF ("G/L Entry".Reversed = TRUE) THEN BEGIN
                    SourceDesc := 'Reversal Entry';
                END;
                //EBT STIVAN ---(30072012)--- To Capture the heading of Voucher when its Reversed --------END

                IF "G/L Entry"."Source Code" = 'PURCHASES' THEN BEGIN
                    IF VendorSource.GET("G/L Entry"."Source No.") THEN
                        SourceDesc := VendorSource."Vendor Posting Group";
                    /*
                      IF PostedPInv.GET("G/L Entry"."Document No.") THEN
                         VendorInvoNo := PostedPInv."Vendor Invoice No." +',     Date :- ' + FORMAT(PostedPInv."Document Date")
                         ELSE
                        VendorInvoNo := '';
                     */
                END;

                GenJournalbatch.RESET;
                GenJournalbatch.SETRANGE(GenJournalbatch.Name, "G/L Entry"."Journal Batch Name");
                IF GenJournalbatch.FINDFIRST THEN BEGIN
                    IF GenJournalbatch."Journal Template Name" = 'CR NOTE' THEN
                        SourceDesc := 'Credit Note';
                    IF GenJournalbatch."Journal Template Name" = 'DR NOTE' THEN
                        SourceDesc := 'Debit Note';
                END;


                CLEAR(TaxJnlDescription);
                IF "G/L Entry"."Source Code" = 'TAXJNL' THEN
                    TaxJnlDescription := 'Narration: ' + "G/L Entry".Description;

                //<<NAV2009

                //>>nav2009

                //EBTuser.GET(USERID);
                //IF EBTuser."User ID" <> 'SA' THEN
                IF EBTuser."User Name" <> 'SA' THEN BEGIN
                    IF ("G/L Entry"."G/L Account No." = '73310010') OR ("G/L Entry"."G/L Account No." = '73310020') OR
                       ("G/L Entry"."G/L Account No." = '73310030') OR ("G/L Entry"."G/L Account No." = '73310040') OR
                       ("G/L Entry"."G/L Account No." = '73310050') OR ("G/L Entry"."G/L Account No." = '73310060') OR
                       ("G/L Entry"."G/L Account No." = '73310070') OR ("G/L Entry"."G/L Account No." = '32160400') OR
                       ("G/L Entry"."G/L Account No." = '12913010') OR ("G/L Entry"."G/L Account No." = '32160100') OR
                       ("G/L Entry"."G/L Account No." = '32160600') OR ("G/L Entry"."G/L Account No." = '32160900') THEN BEGIN
                        DimensionValue := '';
                    END ELSE BEGIN
                        DimensionValue := '';
                        //nah

                        recDimSet.RESET;
                        recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        recDimSet.SETRANGE("Dimension Code", 'STAFF');
                        IF recDimSet.FINDFIRST THEN BEGIN
                            recDimSet.CALCFIELDS("Dimension Value Name");
                            DimensionValue := recDimSet."Dimension Value Code";
                        END;
                    END;

                END ELSE BEGIN

                    DimensionValue := '';

                    recDimSet.RESET;
                    recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                    recDimSet.SETRANGE("Dimension Code", 'STAFF');
                    IF recDimSet.FINDFIRST THEN BEGIN
                        recDimSet.CALCFIELDS("Dimension Value Name");
                        DimensionValue := recDimSet."Dimension Value Code";
                    END;

                END;


                "Disc&IndAdj&AutoAdjDim" := '';
                recDimSet.RESET;
                recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                recDimSet.SETRANGE("Dimension Code", 'DISCOUNT');
                IF recDimSet.FINDFIRST THEN BEGIN
                    recDimSet.CALCFIELDS("Dimension Value Name");
                    "Disc&IndAdj&AutoAdjDim1" := recDimSet."Dimension Value Name";
                END;

                recDimSet.RESET;
                recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                recDimSet.SETRANGE("Dimension Code", 'IND-ADJ');
                IF recDimSet.FINDFIRST THEN BEGIN
                    recDimSet.CALCFIELDS("Dimension Value Name");
                    "Disc&IndAdj&AutoAdjDim1" := recDimSet."Dimension Value Name";
                END;


                //   "Disc&IndAdj&AutoAdjDim" := DocumentDimension."Dimension Value Code";

                recDimSet.RESET;
                recDimSet.SETRANGE("Dimension Set ID", "Dimension Set ID");
                recDimSet.SETRANGE("Dimension Code", 'AUTO-ADJ');
                IF recDimSet.FINDFIRST THEN BEGIN
                    recDimSet.CALCFIELDS("Dimension Value Name");
                    "Disc&IndAdj&AutoAdjDim1" := recDimSet."Dimension Value Name";
                END;

                //      "Disc&IndAdj&AutoAdjDim" := DocumentDimension."Dimension Value Code";
                //<<nav2009

                PageLoop := PageLoop - 1;
                LinesPrinted := LinesPrinted + 1;

                ChequeNo := '';
                ChequeDate := 0D;
                IF ("Source No." <> '') AND ("Source Type" = "Source Type"::"Bank Account") THEN BEGIN
                    IF BankAccLedgEntry.GET("Entry No.") THEN BEGIN
                        ChequeNo := BankAccLedgEntry."Cheque No.";
                        ChequeDate := BankAccLedgEntry."Cheque Date";
                    END;
                END;

                IF (ChequeNo <> '') AND (ChequeDate <> 0D) THEN BEGIN
                    PageLoop := PageLoop - 1;
                    LinesPrinted := LinesPrinted + 1;
                END;
                IF PostingDate <> "Posting Date" THEN BEGIN
                    PostingDate := "Posting Date";
                    TotalDebitAmt := 0;
                END;
                IF DocumentNo <> "Document No." THEN BEGIN
                    DocumentNo := "Document No.";
                    TotalDebitAmt := 0;
                END;

                IF PostingDate = "Posting Date" THEN BEGIN
                    InitTextVariable;
                    TotalDebitAmt += "Debit Amount";
                    FormatNoText(NumberText, ABS(TotalDebitAmt), '');
                    PageLoop := NUMLines;
                    LinesPrinted := 0;
                END;
                IF (PrePostingDate <> "Posting Date") OR (PreDocumentNo <> "Document No.") THEN BEGIN
                    DebitAmountTotal := 0;
                    CreditAmountTotal := 0;
                    PrePostingDate := "Posting Date";
                    PreDocumentNo := "Document No.";
                END;

                DebitAmountTotal := DebitAmountTotal + "Debit Amount";
                CreditAmountTotal := CreditAmountTotal + "Credit Amount";

                //>>2

                //G/L Entry, Header (1) - OnPreSection()
                IF PostedPInv.GET("G/L Entry"."Document No.") THEN
                    VendorInvoNo := PostedPInv."Vendor Invoice No." + ',     Date :- ' + FORMAT(PostedPInv."Document Date")
                ELSE
                    VendorInvoNo := '';
                //<<2
                //G/L Entry, Body (3) - OnPostSection()
                RespCtr.RESET;
                IF "G/L Entry"."Global Dimension 2 Code" <> '' THEN
                    RespCtr.GET("G/L Entry"."Global Dimension 2 Code");

                //>>***Dimesions
                CLEAR(vDimesion);
                CLEAR(ResCode07);//07Jun2017
                IF ("G/L Entry"."G/L Account No." = '12915080') OR ("G/L Entry"."G/L Account No." = '51010030') OR ("G/L Entry"."G/L Account No." = '51035703') OR
                   ("G/L Entry"."G/L Account No." = '51035704') OR ("G/L Entry"."G/L Account No." = '51035705') OR ("G/L Entry"."G/L Account No." = '51035706') OR
                   ("G/L Entry"."G/L Account No." = '71010050') OR ("G/L Entry"."G/L Account No." = '71020045') OR ("G/L Entry"."G/L Account No." = '71030045') OR
                   ("G/L Entry"."G/L Account No." = '71055010') OR ("G/L Entry"."G/L Account No." = '71055015') OR ("G/L Entry"."G/L Account No." = '71055040') OR
                   ("G/L Entry"."G/L Account No." = '73110050') OR ("G/L Entry"."G/L Account No." = '74012080') OR ("G/L Entry"."G/L Account No." = '75001010') OR
                   ("G/L Entry"."G/L Account No." = '75001020') OR ("G/L Entry"."G/L Account No." = '75001025') OR ("G/L Entry"."G/L Account No." = '75001030') OR
                   ("G/L Entry"."G/L Account No." = '75001050') OR ("G/L Entry"."G/L Account No." = '75001110') OR ("G/L Entry"."G/L Account No." = '75003020') OR
                   ("G/L Entry"."G/L Account No." = '75003040') OR ("G/L Entry"."G/L Account No." = '75003050') THEN BEGIN
                    IF "G/L Entry"."Source Code" = 'PURCHASES' THEN //17May2017
                    BEGIN //17May2017

                        //>>05Jun2017
                        vDimSetEntry.RESET;
                        vDimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");//05Jun2017
                        vDimSetEntry.SETRANGE("Dimension Code", 'DIVISION');
                        IF vDimSetEntry.FINDFIRST THEN
                            vDimesion := vDimSetEntry."Dimension Value Code"
                        ELSE
                            vDimesion := "G/L Entry"."Global Dimension 1 Code";
                        //>>05Jun2017

                        //>>07Jun2017
                        vDimSetEntry.RESET;
                        vDimSetEntry.SETRANGE("Dimension Set ID", "Dimension Set ID");
                        vDimSetEntry.SETRANGE("Dimension Code", 'RESPCENTER');
                        IF vDimSetEntry.FINDFIRST THEN
                            ResCode07 := vDimSetEntry."Dimension Value Code"
                        ELSE
                            ResCode07 := "G/L Entry"."Global Dimension 2 Code";
                        //>>07Jun2017

                        //>>14July2017
                        /*  PosStr14.RESET;
                         PosStr14.SETRANGE(Type, PosStr14.Type::Purchase);
                         PosStr14.SETRANGE("Document Type", PosStr14."Document Type"::Invoice);
                         PosStr14.SETRANGE("Invoice No.", "Document No.");
                         PosStr14.SETRANGE("Account No.", "G/L Account No.");
                         PosStr14.SETRANGE("Tax/Charge Group", 'LOADING');
                         PosStr14.SETRANGE(Amount, Amount);
                         IF PosStr14.FINDFIRST THEN BEGIN */

                        PILLineNo += 10000;
                        //MESSAGE(' %1 Line No',PILLineNo);

                        PIL14.RESET;
                        // PIL14.SETRANGE("Document No.", PosStr14."Invoice No.");
                        PIL14.SETRANGE("Line No.", PILLineNo);
                        IF PIL14.FINDFIRST THEN BEGIN
                            ResCode07 := PIL14."Shortcut Dimension 2 Code";
                            vDimesion := PIL14."Shortcut Dimension 1 Code";
                        END;

                        //END;
                        //<<14July2017
                        /*
                        vPurchLine.RESET;
                        vPurchLine.SETRANGE("Document No.","G/L Entry"."Document No.");//17may2017
                        vPurchLine.SETRANGE("Posting Date","G/L Entry"."Posting Date");//17May2017
                        //vPurchLine.SETRANGE("No.","G/L Entry"."G/L Account No."); //Commented 05May2017
                        IF vPurchLine.FINDFIRST THEN
                        BEGIN
                          vDimSetEntry.RESET;
                          vDimSetEntry.SETRANGE("Dimension Set ID","Dimension Set ID");//05Jun2017
                          //vDimSetEntry.SETRANGE("Dimension Set ID",vPurchLine."Dimension Set ID");
                          vDimSetEntry.SETRANGE("Dimension Code",'DIVISION');
                          IF vDimSetEntry.FINDFIRST THEN
                            vDimesion := vDimSetEntry."Dimension Value Code"
                          ELSE
                            vDimesion := "G/L Entry"."Global Dimension 1 Code";
                        END;
                        */

                    END ELSE //17May2017
                    BEGIN
                        vDimesion := "G/L Entry"."Global Dimension 1 Code";//17May2017
                        ResCode07 := "G/L Entry"."Global Dimension 2 Code";//07Jun2017
                    END;
                END ELSE BEGIN
                    vDimesion := "G/L Entry"."Global Dimension 1 Code";
                    ResCode07 := "G/L Entry"."Global Dimension 2 Code";//07Jun2017
                END;
                //<<

            end;

            trigger OnPreDataItem()
            begin
                NUMLines := 13;
                PageLoop := NUMLines;
                LinesPrinted := 0;
                DebitAmountTotal := 0;
                CreditAmountTotal := 0;
            end;
        }
        dataitem("G/L Entry 1"; "G/L Entry")
        {
            DataItemTableView = SORTING("Document No.", "Posting Date", Amount)
                                ORDER(Descending);
            column(TransportEntriesDocNo; TransportEntriesDocNo)
            {
            }
            column(GL1ExpInvNo; "Exp/Purchase Invoice Doc. No.")
            {
            }
            column(GL1DocNo; "Document No.")
            {
            }
            column(GL1CrAmt; "G/L Entry 1"."Credit Amount")
            {
            }
            column(GL1EntryNo; "Entry No.")
            {
            }
            dataitem("G/L Entry 2"; "G/L Entry")
            {
                DataItemLink = "Exp/Purchase Invoice Doc. No." = FIELD("Document No."),
                               "G/L Account No." = FIELD("G/L Account No.");
                DataItemTableView = SORTING("G/L Account No.", "Posting Date")
                                    ORDER(Ascending);
                column(GL2DocNo; "Document No.")
                {
                }
                column(GL2CrAmt; "G/L Entry 2"."Credit Amount")
                {
                }
                column(GL2ExpInvNo; "G/L Entry 2"."Exp/Purchase Invoice Doc. No.")
                {
                }
            }

            trigger OnAfterGetRecord()
            begin

                //>>NAV2009
                //EBT STIVAN---(22/07/2013)---To Capture Transport Details Invocie No.----------START
                TransportEntriesDocNo := '';
                recTransportDetails.RESET;
                recTransportDetails.SETRANGE(recTransportDetails.Open, FALSE);
                recTransportDetails.SETRANGE(recTransportDetails.Applied, TRUE);
                recTransportDetails.SETRANGE(recTransportDetails.Posted, TRUE);
                recTransportDetails.SETRANGE(recTransportDetails."Applied Document No.", "G/L Entry 1"."Document No.");
                IF recTransportDetails.FINDFIRST THEN
                    REPEAT
                        TransportEntriesDocNo := recTransportDetails."Invoice No." + ', ';
                    UNTIL recTransportDetails.NEXT = 0;
                //EBT STIVAN---(22/07/2013)---To Capture Transport Details Invocie No.------------END
                //<<NAV2009
                //MESSAGE('%1 DocNo \\ %2 Entry No.',"G/L Entry 1"."Document No.", "G/L Entry 1"."Entry No.");
            end;

            trigger OnPreDataItem()
            begin
                "G/L Entry 1".SETRANGE("G/L Entry 1"."Document No.", "G/L Entry"."Document No.");//NAV2009
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
                    field(PrintLineNarration; PrintLineNarration)
                    {
                        ApplicationArea = ALL;
                        Caption = 'PrintLineNarration';
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

    trigger OnPreReport()
    begin
        CompanyInformation.GET;

        VendorInvoNo := '';   //05Apr2017
        PILLineNo := 0; //14July2017
    end;

    var
        CompanyInformation: Record 79;
        SourceCode: Record 230;
        GLEntry: Record 17;
        BankAccLedgEntry: Record 271;
        GLAccName: Text[80];
        SourceDesc: Text[80];
        CrText: Text[2];
        DrText: Text[2];
        NumberText: array[2] of Text[80];
        PageLoop: Integer;
        LinesPrinted: Integer;
        NUMLines: Integer;
        ChequeNo: Code[60];
        ChequeDate: Date;
        Text16526: Label 'ZERO';
        Text16527: Label 'HUNDRED';
        Text16528: Label 'AND';
        Text16529: Label '%1 results in a written number that is too long.';
        Text16532: Label 'ONE';
        Text16533: Label 'TWO';
        Text16534: Label 'THREE';
        Text16535: Label 'FOUR';
        Text16536: Label 'FIVE';
        Text16537: Label 'SIX';
        Text16538: Label 'SEVEN';
        Text16539: Label 'EIGHT';
        Text16540: Label 'NINE';
        Text16541: Label 'TEN';
        Text16542: Label 'ELEVEN';
        Text16543: Label 'TWELVE';
        Text16544: Label 'THIRTEEN';
        Text16545: Label 'FOURTEEN';
        Text16546: Label 'FIFTEEN';
        Text16547: Label 'SIXTEEN';
        Text16548: Label 'SEVENTEEN';
        Text16549: Label 'EIGHTEEN';
        Text16550: Label 'NINETEEN';
        Text16551: Label 'TWENTY';
        Text16552: Label 'THIRTY';
        Text16553: Label 'FORTY';
        Text16554: Label 'FIFTY';
        Text16555: Label 'SIXTY';
        Text16556: Label 'SEVENTY';
        Text16557: Label 'EIGHTY';
        Text16558: Label 'NINETY';
        Text16559: Label 'THOUSAND';
        Text16560: Label 'MILLION';
        Text16561: Label 'BILLION';
        Text16562: Label 'LAKH';
        Text16563: Label 'CRORE';
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        PrintLineNarration: Boolean;
        PostingDate: Date;
        TotalDebitAmt: Decimal;
        DocumentNo: Code[20];
        DebitAmountTotal: Decimal;
        CreditAmountTotal: Decimal;
        PrePostingDate: Date;
        PreDocumentNo: Code[60];
        VoucherNoCaptionLbl: Label 'Voucher No. :';
        CreditAmountCaptionLbl: Label 'Credit Amount';
        DebitAmountCaptionLbl: Label 'Debit Amount';
        ParticularsCaptionLbl: Label 'Particulars';
        AmountInWordsCaptionLbl: Label 'Amount (in words):';
        PreparedByCaptionLbl: Label 'Prepared by:';
        CheckedByCaptionLbl: Label 'Checked by:';
        ApprovedByCaptionLbl: Label 'Approved by:';
        IntegerOccurcesCaptionLbl: Label 'IntegerOccurces';
        NarrationCaptionLbl: Label 'Narration :';
        "----05Apr2017": Integer;
        VendorInvoNo: Code[80];
        recSIH: Record 112;
        VendorSource: Record 23;
        GenJournalbatch: Record 232;
        TaxJnlDescription: Text[100];
        TransportEntriesDocNo: Code[1024];
        recTransportDetails: Record 50020;
        recDimSet: Record 480;
        DivisionName: Text[30];
        "Res.CenterName": Text[50];
        "Res.Code": Code[20];
        EBTuser: Record 2000000120;
        StaffDimensionCode: Code[10];
        StaffDimensionName: Text[60];
        DepositEMDcode: Code[10];
        DepositEMDName: Text[60];
        DepositDealerCode: Code[10];
        DepositDealerName: Text[60];
        DepositOthersCode: Code[10];
        DepositOthersName: Text[60];
        DepositRentCode: Code[10];
        DepositRentName: Text[60];
        PhoneNo: Code[20];
        PhoneNoDesc: Text[60];
        VehicleCode: Code[20];
        VehicaleName: Text[60];
        DivisionCode: Code[10];
        "Disc&IndAdj&AutoAdjDim": Code[10];
        DepositReceivedCode: Code[20];
        DepositReceivedName: Text[60];
        DepostiSalesTaxCode: Code[20];
        DepostiSalesTaxName: Text[60];
        SalesPromotionCode: Code[20];
        SalesPromotionName: Text[60];
        SalesPromotion: Text[30];
        "Disc&IndAdj&AutoAdjDim1": Text[30];
        recVendor: Record 23;
        DimensionValue: Text[80];
        PostedPInv: Record 122;
        RespCtr: Record 5714;
        "---11APr2017": Integer;
        recUser: Record 2000000120;
        UserName: Text[80];
        vDimesion: Code[20];
        vPurchLine: Record 123;
        vDimSetEntry: Record 480;
        "---07Jun2017": Integer;
        ResCode07: Code[20];
        "---14July2017": Integer;
        PIH14: Record 122;
        PIL14: Record 123;
        PILLineNo: Integer;
        //PosStr14: Record 13798;
        "--------26July2017": Integer;
        AccountName26: Text[80];
        Vend: Record 23;
        Cust: Record 18;
        FA: Record 5600;
        BankAcc: Record 270;
        GL26: Record 15;

    // //[Scope('Internal')]
    procedure FindGLAccName("Source Type": Option " ",Customer,Vendor,"Bank Account","Fixed Asset"; "Entry No.": Integer; "Source No.": Code[20]; "G/L Account No.": Code[20]): Text[80]
    var
        AccName: Text[80];
        VendLedgerEntry: Record 25;
        Vend: Record 23;
        CustLedgerEntry: Record 21;
        Cust: Record 18;
        BankLedgerEntry: Record 271;
        Bank: Record 270;
        FALedgerEntry: Record 5601;
        FA: Record 5600;
        GLAccount: Record 15;
    begin
        IF "Source Type" = "Source Type"::Vendor THEN
            IF VendLedgerEntry.GET("Entry No.") THEN BEGIN
                Vend.GET("Source No.");
                AccName := Vend.Name;
            END ELSE BEGIN
                GLAccount.GET("G/L Account No.");
                AccName := GLAccount.Name;
            END
        ELSE
            IF "Source Type" = "Source Type"::Customer THEN
                IF CustLedgerEntry.GET("Entry No.") THEN BEGIN
                    Cust.GET("Source No.");
                    AccName := Cust.Name;
                END ELSE BEGIN
                    GLAccount.GET("G/L Account No.");
                    AccName := GLAccount.Name;
                END
            ELSE
                IF "Source Type" = "Source Type"::"Bank Account" THEN
                    IF BankLedgerEntry.GET("Entry No.") THEN BEGIN
                        Bank.GET("Source No.");
                        AccName := Bank.Name;
                    END ELSE BEGIN
                        GLAccount.GET("G/L Account No.");
                        AccName := GLAccount.Name;
                    END
                ELSE BEGIN
                    GLAccount.GET("G/L Account No.");
                    AccName := GLAccount.Name;
                END;

        IF "Source Type" = "Source Type"::" " THEN BEGIN
            GLAccount.GET("G/L Account No.");
            AccName := GLAccount.Name;
        END;

        EXIT(AccName);
    end;

    ////[Scope('Internal')]
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
            AddToNoText(NoText, NoTextIndex, PrintExponent, Text16526)
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
                    AddToNoText(NoText, NoTextIndex, PrintExponent, Text16527);
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
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + '') //Currency."Currency Numeric Description");
        END ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, 'RUPEES');

        AddToNoText(NoText, NoTextIndex, PrintExponent, Text16528);

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
                AddToNoText(NoText, NoTextIndex, PrintExponent, Text16526);
        IF (CurrencyCode <> '') THEN
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' ' + ''/*Currency."Currency Decimal Description"*/ + ' ONLY')
        ELSE
            AddToNoText(NoText, NoTextIndex, PrintExponent, ' PAISA ONLY');
    end;

    local procedure AddToNoText(var NoText: array[2] of Text[80]; var NoTextIndex: Integer; var PrintExponent: Boolean; AddText: Text[30])
    begin
        PrintExponent := TRUE;

        WHILE STRLEN(NoText[NoTextIndex] + ' ' + AddText) > MAXSTRLEN(NoText[1]) DO BEGIN
            NoTextIndex := NoTextIndex + 1;
            IF NoTextIndex > ARRAYLEN(NoText) THEN
                ERROR(Text16529, AddText);
        END;

        NoText[NoTextIndex] := DELCHR(NoText[NoTextIndex] + ' ' + AddText, '<');
    end;

    // //[Scope('Internal')]
    procedure InitTextVariable()
    begin
        OnesText[1] := Text16532;
        OnesText[2] := Text16533;
        OnesText[3] := Text16534;
        OnesText[4] := Text16535;
        OnesText[5] := Text16536;
        OnesText[6] := Text16537;
        OnesText[7] := Text16538;
        OnesText[8] := Text16539;
        OnesText[9] := Text16540;
        OnesText[10] := Text16541;
        OnesText[11] := Text16542;
        OnesText[12] := Text16543;
        OnesText[13] := Text16544;
        OnesText[14] := Text16545;
        OnesText[15] := Text16546;
        OnesText[16] := Text16547;
        OnesText[17] := Text16548;
        OnesText[18] := Text16549;
        OnesText[19] := Text16550;

        TensText[1] := '';
        TensText[2] := Text16551;
        TensText[3] := Text16552;
        TensText[4] := Text16553;
        TensText[5] := Text16554;
        TensText[6] := Text16555;
        TensText[7] := Text16556;
        TensText[8] := Text16557;
        TensText[9] := Text16558;

        ExponentText[1] := '';
        ExponentText[2] := Text16559;
        ExponentText[3] := Text16562;
        ExponentText[4] := Text16563;
    end;
}


report 50180 "Purchase Return Invoice"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseReturnInvoice.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
        {
            RequestFilterFields = "No.";
            column(Location; Loc.Address + ' ' + Loc."Address 2" + ' ' + Loc.City + ' ' + Loc."Post Code" + ' ' + RecState2.Description + ' Phone: ' + Loc."Phone No." + ' ' + ' Email: ' + Loc."E-Mail")
            {
            }
            column(Loc_City; Loc.City)
            {
            }
            column(ECCDesc; ECCDesc)
            {
            }
            column(InvNo; InvNo)
            {
            }
            column(ECC_Range; '')//"E.C.C. Nos."."C.E. Range" + '&' + "E.C.C. Nos."."C.E. Division" + '& ' + "E.C.C. Nos."."C.E. Commissionerate")
            {
            }
            column(CE_Addr1; '"E.C.C. Nos."."C.E Address1"')
            {
            }
            column(CE_Addr2; '"E.C.C. Nos."."C.E Address2"')
            {
            }
            column(CE_Addr3; '')// "E.C.C. Nos."."C.E City" + '-' + "E.C.C. Nos."."C.E Post Code")
            {
            }
            column(TIN_CstNo; '')// Loc."T.I.N. No." + ' & ' + Loc."C.S.T No.")
            {
            }
            column(PostingDate_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Posting Date")
            {
            }
            column(CurrDatetime; CurrDatetime)
            {
            }
            column(No_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."No.")
            {
            }
            column(VarPONo; VarPONo)
            {
            }
            column(VarPODate; VarPODate)
            {
            }
            column(VehOwnersName; VehOwnersName)
            {
            }
            column(a; a)
            {
            }
            column(Desc_PaymentTerm; "Payment Terms".Description)
            {
            }
            column(ResponsibilityCenter_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Responsibility Center")
            {
            }
            column(CustName; CustName)
            {
            }
            column(BuyfromAddress_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Buy-from Address")
            {
            }
            column(BuyfromAddress2_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Buy-from Address 2")
            {
            }
            column(Buyer_City; "Buy-from City" + ' - ' + "Buy-from Post Code")
            {
            }
            column(Buter_Country; RecState3.Description + ', ' + RecCountry.Name)
            {
            }
            column(ShiptoName; ShiptoName)
            {
            }
            column(SuppliersCode; SuppliersCode)
            {
            }
            column(Buyer_CST; 'RecCust."C.S.T. No."')
            {
            }
            column(Buyer_ECC; 'RecCust."E.C.C. No."')
            {
            }
            column(Buyer_TIN; 'RecCust."T.I.N. No."')
            {
            }
            column(ShiptoLST; ShiptoLST)
            {
            }
            column(ShipToCST; ShipToCST)
            {
            }
            column(ECCNo; ECCNo)
            {
            }
            column(ShiptoTIN; ShiptoTIN)
            {
            }
            column(Remarks_PurchCrMemoHdr; "Purch. Cr. Memo Hdr.".Remarks)
            {
            }
            column(InsuranceNo_CompInfo; CompanyInfo1."Insurance No." + ' of the ' + CompanyInfo1."Insurance Provider")
            {
            }
            column(CompInfo_Valid; 'valid up to ' + FORMAT(CompanyInfo1."Policy Expiration Date"))
            {
            }
            column(State1; State1)
            {
            }
            dataitem(CopyLoop; Integer)
            {
                DataItemTableView = SORTING(Number);
                dataitem(Integer; Integer)
                {
                    DataItemTableView = SORTING(Number)
                                        WHERE(Number = FILTER(1));
                    column(InvType; InvType)
                    {
                    }
                    column(InvType1; InvType1)
                    {
                    }
                    column(number; Number)
                    {
                    }
                    column(outPutNos; OutPutNo)
                    {
                    }
                    dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemLinkReference = "Purch. Cr. Memo Hdr.";
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Quantity = FILTER(<> 0));
                        //"Excise Prod. Posting Group" = FILTER(<> ''));
                        column(ExciseProdPostingGroup_PurchCrMemoLine; '"Purch. Cr. Memo Line"."Excise Prod. Posting Group"')
                        {
                        }
                        column(Description_PurchCrMemoLine; "Purch. Cr. Memo Line".Description)
                        {
                        }
                        column(PackDescription; PackDescription)
                        {
                        }
                        column(QuantityBase_PurchCrMemoLine; "Purch. Cr. Memo Line"."Quantity (Base)")
                        {
                        }
                        column(BasicRate; "Direct Unit Cost" / "Qty. per Unit of Measure")
                        {
                        }
                        column(BEDprcnt; BEDprcnt)
                        {
                        }
                        column(DutyAmt; (("Direct Unit Cost" / "Qty. per Unit of Measure") * BEDprcnt) / 100)
                        {
                        }
                        column(BEDAmt; 0)// ABS("BED Amount"))
                        {
                        }
                        column(SheCess_Ecess; 0)// ABS("SHE Cess Amount") + ABS("eCess Amount"))
                        {
                        }
                        column(totamount; totamount)
                        {
                        }
                        column(TaxAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."Tax Amount")
                        {
                        }
                        column(ADEAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."ADE Amount")
                        {
                        }
                        column(SHECessAmount_PurchCrMemoLine; 0)//"Purch. Cr. Memo Line"."SHE Cess Amount")
                        {
                        }
                        column(eCessAmount_PurchCrMemoLine; 0)// "Purch. Cr. Memo Line"."eCess Amount")
                        {
                        }
                        column(Freight; Freight)
                        {
                        }
                        column(RoundOffAmnt; RoundOffAmnt)
                        {
                        }
                        column(TotAmt; totallineamount + 0)// ABS("BED Amount") + ABS("eCess Amount") + ABS("SHE Cess Amount") + "Tax Amount" + Freight + RoundOffAmnt + ABS("Purch. Cr. Memo Line"."ADE Amount"))
                        {
                        }
                        column(DescriptionLineDuty; DescriptionLineDuty[1])
                        {
                        }
                        column(DescriptionLineeCess; DescriptionLineeCess[1])
                        {
                        }
                        column(DescriptionLineSHeCess; DescriptionLineSHeCess[1])
                        {
                        }
                        column(DescriptionLineTot; DescriptionLineTot[1])
                        {
                        }
                        column(EcessPer; FORMAT(ECESSprcnt) + '%')
                        {
                        }
                        column(SheCessPer; FORMAT(SHECESSprcnt) + '%')
                        {
                        }
                        column(TaxDescription; TaxDescription)
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin

                            Ctr := Ctr + 1;

                            IF RecItem.GET("No.") THEN
                                /* recExcisePostnSetup.RESET;
                            recExcisePostnSetup.SETRANGE("Excise Bus. Posting Group", "Purch. Cr. Memo Line"."Excise Bus. Posting Group");
                            recExcisePostnSetup.SETRANGE("Excise Prod. Posting Group", "Purch. Cr. Memo Line"."Excise Prod. Posting Group");
                            IF recExcisePostnSetup.FINDFIRST THEN BEGIN
                                BEDprcnt := recExcisePostnSetup."BED %";
                                ECESSprcnt := recExcisePostnSetup."eCess %";
                                SHECESSprcnt := recExcisePostnSetup."SHE Cess %";
                            END; */

                            PurchCrMemoLine.RESET;
                            PurchCrMemoLine.SETRANGE(PurchCrMemoLine."Document No.", "Purch. Cr. Memo Line"."Document No.");
                            PurchCrMemoLine.SETRANGE(PurchCrMemoLine.Type, PurchCrMemoLine.Type::"G/L Account");
                            PurchCrMemoLine.SETRANGE(PurchCrMemoLine."No.", RoundOffAcc);
                            IF PurchCrMemoLine.FINDFIRST THEN BEGIN
                                RoundOffAmnt := PurchCrMemoLine."Direct Unit Cost";
                            END;


                            CLEAR(totamount);
                            Qty := Quantity;
                            QtyPerUOM := "Qty. per Unit of Measure";

                            IF "Purch. Cr. Memo Line"."Unit of Measure Code" = 'BRL' THEN BEGIN
                                PackDescription := FORMAT(Qty) + '*' +
                                FORMAT("Purch. Cr. Memo Line"."Qty. per Unit of Measure") + '' + "Purch. Cr. Memo Line"."Unit of Measure Code";
                            END
                            ELSE BEGIN
                                PackDescription := FORMAT(Qty) + '*' + "Purch. Cr. Memo Line"."Unit of Measure Code";
                            END;

                            totamount := ("Purch. Cr. Memo Line"."Quantity (Base)") *
                                       ("Purch. Cr. Memo Line"."Direct Unit Cost" /
                                       "Purch. Cr. Memo Line"."Qty. per Unit of Measure");

                            totallineamount += totamount;


                            TaxDescr := 'Tax Descr';

                            /*  DetailedTaxEntry.RESET;
                             DetailedTaxEntry.SETRANGE(DetailedTaxEntry."Document No.", "Purch. Cr. Memo Hdr."."No.");
                             IF DetailedTaxEntry.FINDFIRST THEN BEGIN
                                 IF DetailedTaxEntry."Form Code" <> '' THEN BEGIN
                                     FormCode := 'with Form ' + DetailedTaxEntry."Form Code";
                                 END;
                                 TaxDescription := FORMAT(DetailedTaxEntry."Tax Type") + FORMAT(DetailedTaxEntry."Tax %") + '%' + FormCode;
                             END
                             ELSE BEGIN
                                 TaxDescription := 'NoTax Ag.Form' + "Purch. Cr. Memo Hdr."."Form Code";
                             END;

                             IF (RecCust."Vendor Posting Group" = 'FOREIGN') AND ("Purch. Cr. Memo Hdr."."Form Code" = '') THEN BEGIN
                                 TaxDescription := 'NoTax Ag.Export';
                             END; */

                            vCheck.InitTextVariable;
                            vCheck.FormatNoText(DescriptionLineDuty, ROUND(ABS(/*"BED Amount"*/0), 0.01), '');
                            vCheck.FormatNoText(DescriptionLineeCess, ROUND(ABS(/*"eCess Amount"*/0), 0.01), '');
                            vCheck.FormatNoText(DescriptionLineSHeCess, ROUND(ABS(/*"SHE Cess Amount"*/0), 0.01), '');
                            vCheck.FormatNoText(DescriptionLineTot,
                                         ROUND((totallineamount + ABS(/*"Purch. Cr. Memo Line"."BED Amount") + ABS("Purch. Cr. Memo Line"."eCess Amount") +
                                         ABS("Purch. Cr. Memo Line"."SHE Cess Amount") + ABS("Purch. Cr. Memo Line"."ADE Amount") +
                                         "Purch. Cr. Memo Line"."Tax Amount"*/0 + Freight + RoundOffAmnt)), 0.01), '');

                            DescriptionLineDuty[1] := DELCHR(DescriptionLineDuty[1], '=', '*');
                            DescriptionLineeCess[1] := DELCHR(DescriptionLineeCess[1], '=', '*');
                            DescriptionLineSHeCess[1] := DELCHR(DescriptionLineSHeCess[1], '=', '*');
                            DescriptionLineTot[1] := DELCHR(DescriptionLineTot[1], '=', '*');
                        end;

                        trigger OnPreDataItem()
                        begin
                            Ctr := 1;
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin

                        i += 1;

                        IF i = 1 THEN
                            InvType := 'ORIGINAL FOR BUYER'
                        ELSE
                            IF i = 2 THEN
                                InvType := 'DUPLICATE FOR TRANSPORTER'
                            ELSE
                                IF i = 3 THEN
                                    InvType := 'TRIPLICATE FOR ASSESSEE'
                                ELSE
                                    IF i = 4 THEN BEGIN
                                        InvType := 'QUADRAPLICATE FOR ACCOUNTS';
                                        InvType1 := 'NOT FOR CENVAT';
                                    END ELSE
                                        IF i = 5 THEN BEGIN
                                            InvType := 'EXTRA COPY NOT FOR CENVAT';
                                            InvType1 := '';
                                        END ELSE
                                            InvType := '';
                    end;
                }

                trigger OnAfterGetRecord()
                begin
                    OutPutNo += 1;
                end;

                trigger OnPreDataItem()
                begin
                    NoOfCopies := 4;
                    NoOfLoops := ABS(NoOfCopies) + 1;
                    copytext := '';
                    SETRANGE(Number, 1, NoOfLoops);
                end;
            }

            trigger OnAfterGetRecord()
            begin

                InvNoLen := STRLEN("Purch. Cr. Memo Hdr."."No.");
                InvNo := COPYSTR("Purch. Cr. Memo Hdr."."No.", (InvNoLen - 3), 4);

                CustPostSetup.RESET;
                CustPostSetup.FINDFIRST;
                RoundOffAcc := CustPostSetup."Invoice Rounding Account";

                CLEAR(AddlDutyAmount);
                /* PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Purch. Cr. Memo Hdr."."No.");
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type", PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrLineDetails.SETFILTER(PostedStrOrdrLineDetails."Tax/Charge Group", '%1|%2', 'FREIGHT', 'TRANS');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                    REPEAT
                        Freight := +Freight + PostedStrOrdrLineDetails.Amount;
                    UNTIL PostedStrOrdrLineDetails.NEXT = 0; */

                /* PostedStrOrdrLineDetails.RESET;
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Invoice No.", "Purch. Cr. Memo Hdr."."No.");
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Type",
                                                  PostedStrOrdrLineDetails."Tax/Charge Type"::Charges);
                PostedStrOrdrLineDetails.SETRANGE(PostedStrOrdrLineDetails."Tax/Charge Group", 'ADDL.DUTY');
                IF PostedStrOrdrLineDetails.FINDFIRST THEN
                    REPEAT
                        AddlDutyAmount := AddlDutyAmount + PostedStrOrdrLineDetails.Amount;
                    UNTIL PostedStrOrdrLineDetails.NEXT = 0; */


                VarPONo := '';
                VarPODate := 0D;
                PurchInvLine.SETFILTER(PurchInvLine."Document No.", "Purch. Cr. Memo Hdr."."Applies-to Doc. No.");
                IF PurchInvLine.FINDLAST THEN BEGIN
                    //PurchRcptHeader.SETFILTER(PurchRcptHeader."No.", PurchInvLine."Receipt Document No.");
                    IF PurchRcptHeader.FIND('-') THEN BEGIN
                        VarPONo := PurchRcptHeader."Order No.";
                        VarPODate := PurchRcptHeader."Order Date";
                    END;
                END;


                RespCentre.GET("Responsibility Center");
                Loc.GET("Location Code");

                IF RecState2.GET(Loc."State Code") THEN;

                RecState.RESET;
                IF RecState.GET(State) THEN;


                Loc.GET("Purch. Cr. Memo Hdr."."Location Code");

                IF "Buy-from Country/Region Code" <> ''
                  THEN
                    RecCountry.GET("Ship-to Country/Region Code")
                ELSE
                    RecCountry.Name := '';

                RecState3.RESET;
                IF RecState3.GET(State) THEN;

                RecCust.GET("Pay-to Vendor No.");
                SuppliersCode := RecCust."Our Account No.";
                IF RecCust."Full Name" <> ''
                  THEN
                    CustName := RecCust."Full Name"
                ELSE
                    CustName := "Buy-from Vendor Name";

                IF NOT RecCountry1.GET("Ship-to Country/Region Code")
                  THEN
                    RecCountry1.GET("Ship-to Country/Region Code");

                /* "E.C.C. Nos.".SETRANGE("E.C.C. Nos.".Code, Loc."E.C.C. No.");
                "E.C.C. Nos.".FINDFIRST;
                ECCDesc := "E.C.C. Nos.".Description; */

                IF ShipToCode.GET("Sell-to Customer No.", "Ship-to Code")
                  THEN BEGIN
                    ECCNo := ShipToCode."E.C.C. No.";
                    ShiptoName := ShipToCode.Name;
                    ShipToCST := ShipToCode."C.S.T. No.";
                    ShiptoLST := ShipToCode."L.S.T. No.";
                    ShiptoTIN := ShipToCode."T.I.N. No.";
                END
                ELSE BEGIN
                    ECCNo := '';// RecCust."E.C.C. No.";
                    ShiptoName := '';// CustName;
                    ShipToCST := '';//RecCust."C.S.T. No.";
                    ShiptoLST := '';//RecCust."L.S.T. No.";
                    ShiptoTIN := '';// RecCust."T.I.N. No.";
                END;

                CurrDatetime := TIME;

                IF "Purch. Cr. Memo Hdr."."Ship-to Code" <> '' THEN BEGIN
                    ShipToCode.SETRANGE(ShipToCode.Code, "Ship-to Code");
                    IF ShipToCode.FINDFIRST THEN;
                END
                ELSE BEGIN
                    ShipToCode.SETRANGE(ShipToCode."Customer No.", "Sell-to Customer No.");
                    IF ShipToCode.FINDFIRST THEN;
                END;


                Loc.GET("Location Code");

                RecState.RESET;
                RecState.SETRANGE(RecState.Code, Loc."State Code");
                IF RecState.FINDFIRST THEN
                    State1 := RecState.Description;
            end;

            trigger OnPreDataItem()
            begin

                CompanyInfo1.GET;
                CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
                CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
                CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
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

    var
        Loc: Record 14;
        SalesPerson: Record 13;
        RespCentre: Record 5714;
        RecState: Record State;
        RecState1: Record State;
        RecState2: Record State;
        RecState3: Record State;
        RecCountry: Record 9;
        RecCountry1: Record 9;
        CompanyInfo1: Record 79;
        RecCust: Record 23;
        RecCust1: Record 23;
        RecItem: Record 27;
        ShipToCode: Record 222;
        "Payment Terms": Record 3;
        "Shipment Method": Record 10;
        "Shipping Agent": Record 291;
        ILE: Record 32;
        //"E.C.C. Nos.": Record 13708;
        CustPostSetup: Record 92;
        CustName: Text[100];
        UOM: Code[10];
        Qty: Decimal;
        QtyPerUOM: Decimal;
        ShiptoName: Text[80];
        ShipToCST: Code[50];
        ShiptoTIN: Code[50];
        ShiptoLST: Code[50];
        ECCNo: Code[50];
        OnesText: array[20] of Text[30];
        TensText: array[10] of Text[30];
        ExponentText: array[5] of Text[30];
        DescriptionLineDuty: array[2] of Text[132];
        DescriptionLineeCess: array[2] of Text[132];
        DescriptionLineSHeCess: array[2] of Text[132];
        DescriptionLineTot: array[2] of Text[132];
        PackDescription: Text[30];
        SrNo: Integer;
        InvNo: Code[4];
        InvNoLen: Integer;
        RoundOffAcc: Code[20];
        RoundOffAmnt: Decimal;
        "<EBT>": Integer;
        ItemVarParmResFinal: Record 50000;
        TaxDescr: Code[10];
        ECCDesc: Code[30];
        PackDescriptionforILE: Text[30];
        TaxJurisd: Record 320;
        TaxDescription: Text[50];
        Freight: Decimal;
        Discount: Decimal;
        TotalAmttoCustomer: Decimal;
        InvTotal: Decimal;
        CurrDatetime: Time;
        State: Code[30];
        Ctr: Integer;
        // SegManagement: Codeunit SegManagement;
        LogInteraction: Boolean;
        NoOfLoops: Integer;
        NoOfCopies: Integer;
        cust: Record 18;
        copytext: Text[30];
        SalesInvCountPrinted: Codeunit "Sales Inv.-Printed";
        ShowInternalInfo: Boolean;
        InvType: Text[50];
        InvType1: Text[30];
        i: Integer;
        Value: Decimal;
        Intermvalue: Decimal;
        State1: Text[30];
        // DetailedTaxEntry: Record 16522;
        FormCode: Code[30];
        SuppliersCode: Text[30];
        VehOwnersName: Text[100];
        st: Code[20];
        AddlDutyAmount: Decimal;
        // PostedStrOrdrDetails: Record 13760;
        // PostedStrOrdrLineDetails: Record 13798;
        FrieghtAmount: Decimal;
        PurchCrMemoLine: Record 6651;
        a: Text[30];
        PurchInvLine: Record 123;
        PurchRcptHeader: Record 120;
        VarPONo: Code[20];
        VarPODate: Date;
        totamount: Decimal;
        totallineamount: Decimal;
        //   recExcisePostnSetup: Record 13711;
        BEDprcnt: Decimal;
        ECESSprcnt: Decimal;
        SHECESSprcnt: Decimal;
        Text003: Label 'Original for Buyer';
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
        OutPutNo: Integer;
        "--": Integer;
        vCheck: Report "Check Report";
}


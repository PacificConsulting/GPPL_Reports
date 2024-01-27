report 50141 "PENDING FORM (Purch. Invoice)"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PENDINGFORMPurchInvoice.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Purch. Inv. Header"; "Purch. Inv. Header")
        {
            DataItemTableView = SORTING("Pay-to Name", "No.", "Buy-from Vendor No.");
            /* WHERE("Form Code" = FILTER(<> ''),
                  "Form No." = FILTER('')); */
            RequestFilterFields = "Posting Date";
            column(PIH_DocNo; "Purch. Inv. Header"."No.")
            {
            }
            column(respname; respname)
            {
            }
            column(CompName; CompInfo.Name)
            {
            }
            column(Filtercriteria; Filtercriteria)
            {
            }
            column(Daterange; Daterange)
            {
            }
            column(PIH_VendorNo; "Purch. Inv. Header"."Buy-from Vendor No.")
            {
            }
            column(PIH_VendorName; "Purch. Inv. Header"."Pay-to Name")
            {
            }
            column(PIH_PostingDate; "Purch. Inv. Header"."Posting Date")
            {
            }
            column(PIH_FormCode; '"Purch. Inv. Header"."Form Code"')
            {
            }
            column(LineTaxBaseAmt; 0)// recPIL."Tax Base Amount")
            {
            }
            column(LineTaxAmt; 0)// recPIL."Tax Amount")
            {
            }
            column(LineChargeAmt; 0)// recPIL."Charges To Vendor")
            {
            }
            column(LineAmt; recPIL."Amount To Vendor")
            {
            }
            dataitem("Purch. Inv. Line"; "Purch. Inv. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING(Type, "Document No.")
                                    WHERE(Type = FILTER(Item));
                column(PIL_No; "Purch. Inv. Line"."No.")
                {
                }
                column(LineTaxBaseAmt1; 0)// "Tax Base Amount")
                {
                }
                column(LineTaxAmt1; 0)//"Tax Amount")
                {
                }
                column(LineChargeAmt1; 0)// "Charges To Vendor")
                {
                }
                column(LineAmt1; "Amount To Vendor")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    IF PrintToExcel AND PrintDetail THEN BEGIN
                        MakeExcelDataPILBody;//PurchaseInvoiceLineBody
                    END;
                end;

                trigger OnPreDataItem()
                begin
                    SETCURRENTKEY(Type, "Document No."); //14Feb2017 RB-N
                end;
            }

            trigger OnAfterGetRecord()
            begin
                //>>14Feb2017
                //IF PrintToExcel THEN
                //MakeExcelPurchaseInvDataHeader;//PurchaseInvoiceDataHeader
                //<<

                InCount := COUNT; //No of Count
                I += 1; //increment
                //>>1


                recPIL.RESET;
                recPIL.SETCURRENTKEY(Type, "Document No.");
                recPIL.SETRANGE(recPIL.Type, recPIL.Type::Item);
                recPIL.SETRANGE("Document No.", "Purch. Inv. Header"."No.");
                //recPIL.CALCSUMS("Tax Amount", "Charges To Vendor", "Tax Base Amount", "Amount To Vendor");

                Vendor.GET("Purch. Inv. Header"."Buy-from Vendor No.");
                CSTNo := '';// Vendor."C.S.T. No.";
                TinNo := '';// Vendor."T.I.N. No.";

                "DocNo." := '';
                recGlentry.RESET;
                recGlentry.SETRANGE(recGlentry."Source Type", recGlentry."Source Type"::Vendor);
                recGlentry.SETRANGE(recGlentry."Source No.", "Purch. Inv. Header"."Buy-from Vendor No.");
                recGlentry.SETRANGE(recGlentry."Exp/Purchase Invoice Doc. No.", "Purch. Inv. Header"."No.");
                IF recGlentry.FINDFIRST THEN
                    REPEAT
                        "DocNo." := "DocNo." + recGlentry."Document No." + ', ';
                    UNTIL recGlentry.NEXT = 0;

                IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
                    OrdAddCity := '';
                    OrdAddTIN := '';
                    OrdAddCST := '';
                    recOrderAddress.RESET;
                    recOrderAddress.SETRANGE(recOrderAddress."Vendor No.", "Purch. Inv. Header"."Buy-from Vendor No.");
                    recOrderAddress.SETRANGE(recOrderAddress.Code, "Purch. Inv. Header"."Order Address Code");
                    IF recOrderAddress.FINDFIRST THEN BEGIN
                        OrdAddCity := recOrderAddress.City;
                        OrdAddTIN := recOrderAddress."T.I.N. No.";
                        OrdAddCST := recOrderAddress."C.S.T. No.";
                    END;
                END;
                //<<1

                //>>2
                IF "Purch. Inv. Header"."Responsibility Center" <> '' THEN BEGIN
                    rescenter1.RESET;
                    rescenter1.SETRANGE(rescenter1.Code, "Responsibility Center");
                    IF rescenter1.FIND('-') THEN
                        respname := rescenter1.Name;
                END;
                //<<2

                //>>3 Group Total
                netAmountSales := netAmountSales + recPIL."Amount To Vendor";
                ChargesSales := ChargesSales + 0;// recPIL."Charges To Vendor";
                TaxbaseSales := TaxbaseSales + 0;// recPIL."Tax Base Amount";
                taxamountSales := taxamountSales + 0;// recPIL."Tax Amount";
                //<<3

                IF PrintToExcel THEN BEGIN
                    MakeExcelDataPIHBody;//PurchaseInvoiceHeaderBody
                END;
            end;

            trigger OnPostDataItem()
            begin

                //>>15Feb2017 PIH GroupTotal
                netAmount := netAmountSales - netAmountSalesReturn;
                Charges := ChargesSales - ChargesSalesReturn;
                Taxbase := TaxbaseSales - TaxbaseSalesReturn;
                taxamount := taxamountSales - taxamountSalesReturn;

                IF InCount = I THEN BEGIN
                    IF PrintToExcel THEN
                        MakeExcelPurchaseInvDataFooter;
                END;

                //<<
            end;

            trigger OnPreDataItem()
            begin
                SETCURRENTKEY("Pay-to Name", "No.", "Buy-from Vendor No."); //14Feb2017 RB-N

                //>>1

                IF (cdeselltocustno <> '') THEN BEGIN
                    SETRANGE("Buy-from Vendor No.", cdeselltocustno);
                END;

                IF (cdeformcode <> '') THEN BEGIN
                    //SETRANGE("Form Code", cdeformcode);
                END;

                IF (cderespcentre <> '') THEN BEGIN
                    SETRANGE("Responsibility Center", cderespcentre);
                    respcenter.GET(cderespcentre);
                END;

                IF (salespersoncode <> '') THEN BEGIN
                    SETRANGE("Purchaser Code", salespersoncode);
                    RecSalesperson.GET(salespersoncode);
                END;

                IF (locationcode <> '') THEN BEGIN
                    SETRANGE("Location Code", locationcode);
                    RecLoc.GET(locationcode);
                END;

                Daterange := GETFILTER("Posting Date");
                locfilter := GETFILTER("Location Code");


                IF locfilter <> '' THEN
                    Filtercriteria := 'Pending Form Location wise :- ' + RecLoc.Name
                ELSE
                    IF cdeformcode <> '' THEN
                        Filtercriteria := 'Pending ' + cdeformcode + ' Forms'
                    ELSE
                        IF cderespcentre <> '' THEN
                            Filtercriteria := 'Pending Form Responsibility wise :- ' + respcenter.Name
                        ELSE
                            IF salespersoncode <> '' THEN
                                Filtercriteria := 'Pending Form salesperson wise :- ' + RecSalesperson.Name
                            ELSE BEGIN
                                IF Vendor.GET(cdeselltocustno) THEN;
                                Filtercriteria := 'Pending Form Party wise :- ' + Vendor.Name;
                            END;
                //<<1

                //>>14Feb2017
                IF PrintToExcel THEN
                    MakeExcelInfo; //InformationSheet
                //<<

                I := 0;
            end;
        }
        dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Pre-Assigned No." = FILTER(<> ''));
            /* "Form Code" = FILTER(<> ''),
            "Form No." = FILTER('')); */
            column(PCM_DocNo; "Purch. Cr. Memo Hdr."."No.")
            {
            }
            column(PCM_VendorNo; "Purch. Cr. Memo Hdr."."Buy-from Vendor No.")
            {
            }
            column(PCM_VendorName; "Purch. Cr. Memo Hdr."."Pay-to Name")
            {
            }
            column(PCM_PostingDate; "Purch. Cr. Memo Hdr."."Posting Date")
            {
            }
            column(PCM_FormCode; '"Purch. Cr. Memo Hdr."."Form Code"')
            {
            }
            column(PCM_LineTaxBaseAmt; 0)// recCrMemLine."Tax Base Amount")
            {
            }
            column(PCM_LineTaxAmt; 0)//recCrMemLine."Tax Amount")
            {
            }
            column(PCM_LineChargeAmt; 0)// recCrMemLine."Charges To Vendor")
            {
            }
            column(PCM_LineAmt; 0)// recCrMemLine."Amount To Vendor")
            {
            }
            column(respname1; respname1)
            {
            }
            dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));
                column(PCL_No; "Purch. Cr. Memo Line"."No.")
                {
                }
                column(PCM_LineTaxBaseAmt1; 0)// "Tax Base Amount")
                {
                }
                column(PCM_LineTaxAmt1; 0)//"Tax Amount")
                {
                }
                column(PCM_LineChargeAmt1; 0)// "Charges To Vendor")
                {
                }
                column(PCM_LineAmt1; 0)// "Amount To Vendor")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    //>>1

                    IF (("Purch. Cr. Memo Line".Type = "Purch. Cr. Memo Line".Type::"G/L Account") AND
                       ("Purch. Cr. Memo Line"."No." = '74012210')) THEN BEGIN
                        CurrReport.SKIP;
                    END;

                    //<<1

                    IF PrintToExcel AND PrintDetail THEN BEGIN
                        MakeExcelDataPCLBody;//PurchaseCreditMemoLineBody
                    END;
                end;
            }

            trigger OnAfterGetRecord()
            begin
                Cm += 1; //Increment
                CmCount := COUNT;//No. of Records

                //>>1

                recCrMemLine.RESET;
                recCrMemLine.SETRANGE("Document No.", "Purch. Cr. Memo Hdr."."No.");
                //recCrMemLine.CALCSUMS("Tax Amount", "Charges To Vendor", "Tax Base Amount", "Amount To Vendor");

                Vendor.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
                CSTNo := '';// Vendor."C.S.T. No.";
                TinNo := '';// Vendor."T.I.N. No.";

                AppliedInvoiceDate := 0D;
                recPIH.RESET;
                recPIH.SETRANGE(recPIH."No.", "Purch. Cr. Memo Hdr."."Applies-to Doc. No.");
                IF recPIH.FINDFIRST THEN BEGIN
                    AppliedInvoiceDate := recPIH."Posting Date";
                END;

                IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
                    OrdAddCity := '';
                    OrdAddTIN := '';
                    OrdAddCST := '';
                    recOrderAddress.RESET;
                    recOrderAddress.SETRANGE(recOrderAddress."Vendor No.", "Purch. Cr. Memo Hdr."."Buy-from Vendor No.");
                    recOrderAddress.SETRANGE(recOrderAddress.Code, "Purch. Cr. Memo Hdr."."Order Address Code");
                    IF recOrderAddress.FINDFIRST THEN BEGIN
                        OrdAddCity := recOrderAddress.City;
                        OrdAddTIN := recOrderAddress."T.I.N. No.";
                        OrdAddCST := recOrderAddress."C.S.T. No.";
                    END;
                END;
                //<<1

                //>>2
                IF "Purch. Cr. Memo Hdr."."Responsibility Center" <> '' THEN BEGIN
                    rescenter1.RESET;
                    rescenter1.SETRANGE(rescenter1.Code, "Responsibility Center");
                    IF rescenter1.FIND('-') THEN
                        respname1 := rescenter1.Name; //15Feb2017
                END;
                //<<2

                //>>3GroupTotal
                netAmountSalesCM := netAmountSalesCM + 0;// recCrMemLine."Amount To Vendor";
                ChargesSalesCM := ChargesSalesCM + 0;// recCrMemLine."Charges To Vendor";
                TaxbaseSalesCM := TaxbaseSalesCM + 0;// recCrMemLine."Tax Base Amount";
                taxamountSalesCM := taxamountSalesCM + 0;//recCrMemLine."Tax Amount";
                //<<3

                IF PrintToExcel THEN
                    MakeExcelDataPCHBody;//PurchaseCreditMemoHeaderBody
            end;

            trigger OnPostDataItem()
            begin
                //>>15Feb2017 CreditMemo GroupTotal


                IF CmCount = Cm THEN BEGIN
                    IF PrintToExcel THEN
                        MakeExcelPurchCrMemoDataFooter;
                END;

                //<<
            end;

            trigger OnPreDataItem()
            begin
                //>>1

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Posting Date") <> '' THEN BEGIN
                    "Purch. Cr. Memo Hdr.".SETFILTER("Purch. Cr. Memo Hdr."."Posting Date",
                    "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Posting Date"));
                END;

                IF (cdeselltocustno <> '') THEN BEGIN
                    SETRANGE("Buy-from Vendor No.", cdeselltocustno);
                END;

                IF (cdeformcode <> '') THEN BEGIN
                    //SETRANGE("Form Code", cdeformcode);
                END;

                IF (cderespcentre <> '') THEN BEGIN
                    SETRANGE("Responsibility Center", cderespcentre);
                    respcenter.GET(cderespcentre);
                END;

                IF (salespersoncode <> '') THEN BEGIN
                    SETRANGE("Purchaser Code", salespersoncode);
                    RecSalesperson.GET(salespersoncode);
                END;

                IF (locationcode <> '') THEN BEGIN
                    SETRANGE("Location Code", locationcode);
                    RecLoc.GET(locationcode);
                END;
                //<<1

                Cm := 0; //15Feb2017 RB-N

                IF PrintToExcel THEN
                    MakeExcelPurchCrMemoDataHeader;//PurchaseCreditMemoHeader
            end;
        }
        dataitem("Purch. Cr. Memo Hdr.-ForReturn"; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Pre-Assigned No." = FILTER(= ''));
            /*  "Form Code" = FILTER(<> ''),
             "Form No." = FILTER('')); */
            column(PCR_No; "Purch. Cr. Memo Hdr.-ForReturn"."No.")
            {
            }
            column(PCR_VendorNo; "Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.")
            {
            }
            column(PCR_VendorName; "Purch. Cr. Memo Hdr.-ForReturn"."Pay-to Name")
            {
            }
            column(PCR_PostingDate; "Purch. Cr. Memo Hdr.-ForReturn"."Posting Date")
            {
            }
            column(PCR_FormCode; '"Purch. Cr. Memo Hdr.-ForReturn"."Form Code"')
            {
            }
            column(PCR_LineTaxBaseAmt; 0)// recCrMemLine1."Tax Base Amount")
            {
            }
            column(PCR_LineTaxAmt; 0)// recCrMemLine1."Tax Amount")
            {
            }
            column(PCR_LineChargeAmt; 0)// recCrMemLine1."Charges To Vendor")
            {
            }
            column(PCR_LineAmt; 0)// recCrMemLine1."Amount To Vendor")
            {
            }
            column(respname2; respname2)
            {
            }
            dataitem("Purch. Cr. Memo Line-ForReturn"; "Purch. Cr. Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Quantity = FILTER(<> 0));
                column(PCRL_No; "Purch. Cr. Memo Line-ForReturn"."No.")
                {
                }
                column(PCR_LineTaxBaseAmt1; 0)// "Tax Base Amount")
                {
                }
                column(PCR_LineTaxAmt1; 0)// "Tax Amount")
                {
                }
                column(PCR_LineChargeAmt1; 0)//"Charges To Vendor")
                {
                }
                column(PCR_LineAmt1; 0)// "Amount To Vendor")
                {
                }

                trigger OnAfterGetRecord()
                begin
                    //>>1

                    IF (("Purch. Cr. Memo Line-ForReturn".Type = "Purch. Cr. Memo Line-ForReturn".Type::"G/L Account") AND
                       ("Purch. Cr. Memo Line-ForReturn"."No." = '74012210')) THEN BEGIN
                        CurrReport.SKIP;
                    END;
                    //<<1

                    IF PrintToExcel AND PrintDetail THEN BEGIN
                        MakeExcelDataPRLBody;//PurchaseReturnLineBody
                    END;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                Cr += 1;//Increment
                CrCount := COUNT;//No.of Records

                //>>1


                recCrMemLine1.RESET;
                recCrMemLine1.SETRANGE("Document No.", "Purch. Cr. Memo Hdr.-ForReturn"."No.");
                //recCrMemLine1.CALCSUMS("Tax Amount", "Charges To Vendor", "Tax Base Amount", "Amount To Vendor");

                Vendor.GET("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.");
                CSTNo := '';// Vendor."C.S.T. No.";
                TinNo := '';// Vendor."T.I.N. No.";

                AppliedInvoiceDate := 0D;
                recPIH.RESET;
                recPIH.SETRANGE(recPIH."No.", "Purch. Cr. Memo Hdr.-ForReturn"."Applies-to Doc. No.");
                IF recPIH.FINDFIRST THEN BEGIN
                    AppliedInvoiceDate := recPIH."Posting Date";
                END;

                IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
                    OrdAddCity := '';
                    OrdAddTIN := '';
                    OrdAddCST := '';
                    recOrderAddress.RESET;
                    recOrderAddress.SETRANGE(recOrderAddress."Vendor No.", "Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.");
                    recOrderAddress.SETRANGE(recOrderAddress.Code, "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code");
                    IF recOrderAddress.FINDFIRST THEN BEGIN
                        OrdAddCity := recOrderAddress.City;
                        OrdAddTIN := recOrderAddress."T.I.N. No.";
                        OrdAddCST := recOrderAddress."C.S.T. No.";
                    END;
                END;
                //<<1

                //>>2
                IF "Purch. Cr. Memo Hdr.-ForReturn"."Responsibility Center" <> '' THEN BEGIN
                    rescenter1.RESET;
                    rescenter1.SETRANGE(rescenter1.Code, "Responsibility Center");
                    IF rescenter1.FIND('-') THEN
                        respname2 := rescenter1.Name;//15Feb2017
                END;
                //<<2

                /*
                //>>3GroupTotal Old Code Commented
                netAmountSalesReturn:=netAmountSalesReturn + recCrMemLine."Amount To Vendor";
                ChargesSalesReturn := ChargesSalesReturn + recCrMemLine."Charges To Vendor";
                TaxbaseSalesReturn :=TaxbaseSalesReturn + recCrMemLine."Tax Base Amount";
                taxamountSalesReturn :=taxamountSalesReturn + recCrMemLine."Tax Amount";
                //<<3
                */

                //>>3 15Feb17 RB-N
                //NewCode for PurchaseReturn Group Total
                netAmountSalesReturn := netAmountSalesReturn + 0;// recCrMemLine1."Amount To Vendor";
                ChargesSalesReturn := ChargesSalesReturn + 0;// recCrMemLine1."Charges To Vendor";
                TaxbaseSalesReturn := TaxbaseSalesReturn + 0;//recCrMemLine1."Tax Base Amount";
                taxamountSalesReturn := taxamountSalesReturn + 0;// recCrMemLine1."Tax Amount";

                //<<

                IF PrintToExcel THEN BEGIN
                    MakeExcelDataPRHBody;//PurchaseReturnHeaderBody
                END;

            end;

            trigger OnPostDataItem()
            begin

                //>>15Feb2017 PurchaseReturn GroupTotal


                IF CrCount = Cr THEN BEGIN
                    IF PrintToExcel THEN
                        MakeExcelPurchaseRetDataFooter;
                END;

                //<<
            end;

            trigger OnPreDataItem()
            begin
                Cr := 0;

                //>>1

                IF "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Posting Date") <> '' THEN BEGIN
                    "Purch. Cr. Memo Hdr.-ForReturn".SETFILTER("Purch. Cr. Memo Hdr.-ForReturn"."Posting Date",
                    "Purch. Inv. Header".GETFILTER("Purch. Inv. Header"."Posting Date"));
                END;

                IF (cdeselltocustno <> '') THEN BEGIN
                    SETRANGE("Buy-from Vendor No.", cdeselltocustno);
                END;

                IF (cdeformcode <> '') THEN BEGIN
                    //SETRANGE("Form Code", cdeformcode);
                END;

                IF (cderespcentre <> '') THEN BEGIN
                    SETRANGE("Responsibility Center", cderespcentre);
                    respcenter.GET(cderespcentre);
                END;

                IF (salespersoncode <> '') THEN BEGIN
                    SETRANGE("Purchaser Code", salespersoncode);
                    RecSalesperson.GET(salespersoncode);
                END;

                IF (locationcode <> '') THEN BEGIN
                    SETRANGE("Location Code", locationcode);
                    RecLoc.GET(locationcode);
                END;
                //<<1

                IF PrintToExcel THEN
                    MakeExcelPurchaseRetDataHeader;//PurchaseReturnDataHeader
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Customer No."; cdeselltocustno)
                {
                    ApplicationArea = all;
                    TableRelation = Customer;
                }
                field("Responsiblity Center"; cderespcentre)
                {
                    ApplicationArea = all;
                    TableRelation = "Responsibility Center";
                }
                field("Form Code"; cdeformcode)
                {
                    ApplicationArea = all;
                    //TableRelation = "Form Codes";
                }
                field("SalesPerson Code"; salespersoncode)
                {
                    ApplicationArea = all;
                    TableRelation = "Salesperson/Purchaser";
                }
                field(Location; locationcode)
                {
                    ApplicationArea = all;
                    TableRelation = Location;
                }
                field("Print To Excel"; PrintToExcel)
                {
                    ApplicationArea = all;
                }
                field("Print Details"; PrintDetail)
                {
                    ApplicationArea = all;
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

    trigger OnPostReport()
    begin

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin
        //>>1

        CompInfo.GET;

        netAmount := 0;
        Charges := 0;
        Taxbase := 0;
        taxamount := 0;

        //IF PrintToExcel THEN
        // MakeExcelInfo;
        //<<
    end;

    var
        cdeselltocustno: Code[10];
        cderespcentre: Code[10];
        cdeformcode: Code[10];
        salespersoncode: Code[10];
        recPIH: Record 122;
        recPIL: Record 123;
        netAmount: Decimal;
        Charges: Decimal;
        Partyname: Text[60];
        Vendor: Record 23;
        display1: Boolean;
        Daterange: Text[60];
        CompInfo: Record 79;
        PrintToExcel: Boolean;
        ExcelBuf: Record 370 temporary;
        Check: Report "Check Report";
        SIH: Record 112;
        respcenter: Record 5714;
        Taxbase: Decimal;
        taxamount: Decimal;
        respname: Text[60];
        rescenter1: Record 5714;
        locfilter: Text[60];
        RecLoc: Record 14;
        locationcode: Text[40];
        LocationName: Text[100];
        recCrMemLine: Record 125;
        netAmountSales: Decimal;
        ChargesSales: Decimal;
        TaxbaseSales: Decimal;
        taxamountSales: Decimal;
        netAmountSalesCM: Decimal;
        ChargesSalesCM: Decimal;
        TaxbaseSalesCM: Decimal;
        taxamountSalesCM: Decimal;
        netAmountSalesReturn: Decimal;
        ChargesSalesReturn: Decimal;
        TaxbaseSalesReturn: Decimal;
        taxamountSalesReturn: Decimal;
        Filtercriteria: Text[200];
        RecSalesperson: Record 13;
        CSTNo: Code[100];
        TinNo: Code[100];
        recVend: Record 18;
        VendResCen: Code[10];
        PrintDetail: Boolean;
        recPCH: Record 124;
        AppliedInvoiceDate: Date;
        recCrMemLine1: Record 125;
        recGlentry: Record 17;
        "DocNo.": Code[100];
        recOrderAddress: Record 224;
        OrdAddCity: Code[20];
        OrdAddTIN: Code[20];
        OrdAddCST: Code[20];
        RecOAState: Record State;
        RecOrderAdd: Record 224;
        OAState: Text[1024];
        Text000: Label 'Data';
        Text001: Label 'Pending Form';
        Text0004: Label 'Company Name';
        Text0005: Label 'Report No.';
        Text0006: Label 'Report Name';
        Text0007: Label 'USER Id';
        Text0008: Label 'Date';
        "--14Feb17": Integer;
        CrCount: Integer;
        InCount: Integer;
        I: Integer;
        respname1: Text[60];
        respname2: Text[60];
        Cm: Integer;
        CmCount: Integer;
        Cr: Integer;

    //  //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.SetUseInfoSheet;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn('PENDING FORM (Purchase Invoice)', FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(REPORT::"PENDING FORM (Purch. Invoice)", FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0007), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0008), FALSE, TRUE, FALSE, FALSE, '', 0);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.ClearNewRow;

        //14Feb2017
        MakeExcelPurchaseInvDataHeader;//PurchaseInvoiceDataHeader
    end;

    //  //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text000,Text001,COMPANYNAME,USERID);
        //ExcelBuf.GiveUserControl;
        //ERROR('');

        //>>14Feb2017
        /*  ExcelBuf.CreateBook('', Text001);
         //ExcelBuf.CreateBookAndOpenExcel('','Pending Form Report','Pending Form Report',COMPANYNAME,USERID);
         ExcelBuf.CreateBookAndOpenExcel('', 'Pending Form Report', '', '', USERID); //17Feb2017 RB-N
         ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text001);
        ExcelBuf.WriteSheet(Text001, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text001, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        //<<
    end;

    //   //[Scope('Internal')]
    procedure MakeExcelPurchaseInvDataHeader()
    begin
        ExcelBuf.NewRow;
        //>>13Feb2017 To Display UserId, Filter ,Date
        ExcelBuf.AddColumn(CompInfo.Name, FALSE, '', TRUE, FALSE, FALSE, '@', 0);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//20
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', 0);//21
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn(Filtercriteria, FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//20
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', 0);//21
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('Date Range :' + Daterange, FALSE, '', TRUE, FALSE, FALSE, '@', 0);

        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        //<<

        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//14
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//15
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//16
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//17
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//18
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//19
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//20
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);//21
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Invoice Details', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);//15Feb2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Vendor No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Name Of Party', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Responsibility Center', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Bill No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Vendor Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Taxable Amt', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('JV No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address State', FALSE, '', TRUE, FALSE, TRUE, '@', 0);   //mayuri
    end;

    //[Scope('Internal')]
    procedure MakeExcelDataPIHBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Inv. Header"."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Pay-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(CSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(TinNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(respname, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Vendor Invoice No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*"Purch. Inv. Header"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recPIL."Tax Base Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recPIL."Tax Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recPIL."Charges To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(recPIL."Amount To Vendor", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Order Address Code", FALSE, '', FALSE, FALSE, FALSE, '', 0);

        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '', 0);

        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);

        END;

        ExcelBuf.AddColumn("DocNo.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        //mayuri
        RecOrderAdd.RESET;
        RecOAState.RESET;
        OAState := '';
        IF "Purch. Inv. Header"."Order Address Code" <> '' THEN
            IF RecOrderAdd.GET("Purch. Inv. Header"."Buy-from Vendor No.", "Purch. Inv. Header"."Order Address Code") THEN
                IF RecOAState.GET(RecOrderAdd.State) THEN
                    OAState := RecOAState.Description;

        //mayuri

        ExcelBuf.AddColumn(OAState, FALSE, '', FALSE, FALSE, FALSE, '', 0);  //mayuri
    end;

    //[Scope('Internal')]
    procedure MakeExcelDataPILBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Header"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Inv. Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);//mayuri
    end;

    //[Scope('Internal')]
    procedure MakeExcelPurchaseInvDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);//15Feb2017 RB-N,
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ABS(TaxbaseSales), FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ABS(taxamountSales), FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ChargesSales, FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(netAmountSales, FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);//mayuri
    end;

    // //[Scope('Internal')]
    procedure MakeExcelPurchCrMemoDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Credit Memo Details', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);//15Feb2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);//15Feb2017 RB-N
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Vendor No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Name Of Party', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Responsibility Center', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Cr. Memo No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Vendor Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Applied Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Applied Invoice Date', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Taxable Amt', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address State', FALSE, '', TRUE, FALSE, TRUE, '@', 0);   //mayuri
    end;

    //[Scope('Internal')]
    procedure MakeExcelDataPCHBody()
    begin
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Pay-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(CSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(TinNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(respname, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Vendor Cr. Memo No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(AppliedInvoiceDate, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Hdr."."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recCrMemLine."Tax Base Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recCrMemLine."Tax Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recCrMemLine."Charges To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recCrMemLine."Amount To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Order Address Code", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '', 0);

        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);

        END;
        //mayuri
        RecOrderAdd.RESET;
        RecOAState.RESET;
        OAState := '';
        IF "Purch. Cr. Memo Hdr."."Order Address Code" <> '' THEN
            IF RecOrderAdd.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.", "Purch. Cr. Memo Hdr."."Order Address Code") THEN
                IF RecOAState.GET(RecOrderAdd.State) THEN
                    OAState := RecOAState.Description;
        //mayuri
        ExcelBuf.AddColumn(OAState, FALSE, '', FALSE, FALSE, FALSE, '', 0);  //mayuri
    end;

    //[Scope('Internal')]
    procedure MakeExcelDataPCLBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr."."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line".Description, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);//mayuri
    end;

    //[Scope('Internal')]
    procedure MakeExcelPurchCrMemoDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);//15Feb2017 RB-N
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ABS(TaxbaseSalesCM), FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ABS(taxamountSalesCM), FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ChargesSalesCM, FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(netAmountSalesCM, FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);//mayuri
    end;

    // //[Scope('Internal')]
    procedure MakeExcelPurchaseRetDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', 0);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Return Details', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);//15Feb2017 RB-N
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);//15Feb2017 RB-N
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Vendor No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Name Of Party', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Responsibility Center', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Cr. Memo No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Posting Date', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Vendor Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Applied Invoice No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Applied Invoice Date', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Item No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Form Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Taxable Amt', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Tax Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Charges', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Amount', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address Code', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address City', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address TIN No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address CST No.', FALSE, '', TRUE, FALSE, TRUE, '@', 0);
        ExcelBuf.AddColumn('Order Address State', FALSE, '', TRUE, FALSE, TRUE, '@', 0);   //mayuri
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataPRHBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Pay-to Name", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(CSTNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(TinNo, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(respname, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Vendor Cr. Memo No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Applies-to Doc. No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(AppliedInvoiceDate, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*"Purch. Cr. Memo Hdr.-ForReturn"."Form Code"*/'', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recCrMemLine1."Tax Base Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*ABS(recCrMemLine1."Tax Amount")*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recCrMemLine1."Charges To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn(/*recCrMemLine1."Amount To Vendor"*/0, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN BEGIN
            ExcelBuf.AddColumn(OrdAddCity, FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn(OrdAddTIN, FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn(OrdAddCST, FALSE, '', FALSE, FALSE, FALSE, '', 0);

            //mayuri
            RecOrderAdd.RESET;
            RecOAState.RESET;
            OAState := '';
            IF "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code" <> '' THEN
                IF RecOrderAdd.GET("Purch. Cr. Memo Hdr.-ForReturn"."Buy-from Vendor No.",
                "Purch. Cr. Memo Hdr.-ForReturn"."Order Address Code") THEN
                    IF RecOAState.GET(RecOrderAdd.State) THEN
                        OAState := RecOAState.Description;
            //mayuri
            ExcelBuf.AddColumn(OAState, FALSE, '', FALSE, FALSE, FALSE, '', 0);  //mayuri
        END
        ELSE BEGIN
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
            ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);//mayuri
        END;
    end;

    // //[Scope('Internal')]
    procedure MakeExcelDataPRLBody()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Hdr.-ForReturn"."Posting Date", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn"."No.", FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn("Purch. Cr. Memo Line-ForReturn".Description, FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', FALSE, FALSE, FALSE, '', 0);//mayuri
    end;

    // //[Scope('Internal')]
    procedure MakeExcelPurchaseRetDataFooter()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 0);//15Feb2017 RB-N
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ABS(TaxbaseSalesReturn), FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ABS(taxamountSalesReturn), FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(ChargesSalesReturn, FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn(netAmountSalesReturn, FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '', 0);//mayuri
    end;
}


report 50008 "QC Test Report for Sale--"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 25Oct2017   RB-N         Dispalying RespolLogo Or IpolLogo Depending upon ItemCategory
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/QCTestReportforSale.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Permissions = TableData 112 = rim,
                  TableData 50015 = rim;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            RequestFilterFields = "No.";
            column(Logo21; Logo21)
            {
            }
            column(Cat21; Cat21)
            {
            }
            column(CompName; CompanyInfo1.Name)
            {
            }
            column(CompPicture; CompanyInfo1.Picture)
            {
            }
            column(IpolLogo; CompanyInfo1."Name Picture")
            {
            }
            column(CompShadedBoxPicture; CompanyInfo1."Shaded Box")
            {
            }
            column(RepsolLogo; CompanyInfo1."Repsol Logo")
            {
            }
            column(Comp_GSTRegNo; CompanyInfo1."GST Registration No.")
            {
            }
            column(CustName; "Sell-to Customer Name")
            {
            }
            column(InvNo; "No.")
            {
            }
            column(InvDate; FORMAT("Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(LocAdd1; Loc1.Address + ',' + Loc1."Address 2" + ',' + Loc1.City + '-' + Loc1."Post Code" + ',' + RecState.Description + ',' + ' Tel: ' + Loc1."Phone No." + ',' + ' Fax: ' + Loc1."Fax No." + ',' + ' E-mail: ' + Loc1."E-Mail")
            {
            }
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE(Type = FILTER(Item),
                                          Quantity = FILTER(<> 0));
                column(LineNo; "Line No.")
                {
                }
                column(ItmNo; "No.")
                {
                }
                column(LotNo; "Sales Invoice Line"."Lot No.")
                {
                }
                column(TestDate; FORMAT(TestDate, 0, '<Day,2>/<Month,2>/<Year4>'))
                {
                }
                column(PackDescription; PackDescription)
                {
                }
                column(ProductCode; '"Sales Invoice Line"."Cross-Reference No."')
                {
                }
                column(vItemDesc; vItemDesc)
                {
                }
                dataitem("Item Version Parameters-Result"; "Item Version Parameters-Result")
                {
                    DataItemLink = "Item No." = FIELD("No."),
                                   "Batch No./DC No" = FIELD("Lot No.");
                    DataItemTableView = SORTING("Item No.", "Version Code", "Line No.", Parameter, "Blend Order No")
                                        WHERE(Result = FILTER(<> ''));
                    column(CompWWW; CompanyInfo1."Home Page")
                    {
                    }
                    column(ItemNo_QC; "Item No.")
                    {
                    }
                    column(LineNo_QC; "Item Version Parameters-Result"."Line No.")
                    {
                    }
                    column(CompCINNo; 'CompanyInfo1."Company Registration  No."')
                    {
                    }
                    column(CompPANNo; CompanyInfo1."P.A.N. No.")
                    {
                    }
                    column(BatchNo; "Item Version Parameters-Result"."Batch No./DC No")
                    {
                    }
                    column(BatchDate; BatchDate)
                    {
                    }
                    column(Grade; 'IPOL  ' + "Item Version Parameters-Result"."Item Description")
                    {
                    }
                    column(Ctr; ctr)
                    {
                    }
                    column(SNo; SNo)
                    {
                    }
                    column(Parameter; "Item Version Parameters-Result".Parameter)
                    {
                    }
                    column(TestingMethod; "Testing Method")
                    {
                    }
                    column(TypicalValue; "Item Version Parameters-Result"."Typical Value")
                    {
                    }
                    column(MinV; "Min Value")
                    {
                    }
                    column(MaxV; "Max Value")
                    {
                    }
                    column(Result; Result)
                    {
                    }
                    column(Remark1; "Item Version Parameters-Result"."Remarks 1")
                    {
                    }
                    column(Remark; "Item Version Parameters-Result".Remarks)
                    {
                    }
                    column(ApprovedBy; "Item Version Parameters-Result"."Approved By")
                    {
                    }
                    column(LocAdd; Loc.Address + ' ' + Loc."Address 2" + ' ' + Loc.City + ' ' + Loc."Post Code" + ' ' + Loc."Phone No." + 'E-mail : ' + ' ' + Loc."E-Mail")
                    {
                    }
                    column(NAH; NAH)
                    {
                    }
                    column(MfgDate; FORMAT(BatchDate, 0, '<Day,2>/<Month,2>/<Year4>'))
                    {
                    }

                    trigger OnAfterGetRecord()
                    begin

                        SNo += 1;//10Mar2017

                        //>>1

                        //EBT STIVAN (30052012)-----------------------------------------------------------START
                        IVP.RESET;
                        IVP.SETRANGE("Item No.", "Item Version Parameters-Result"."Item No.");
                        IVP.SETRANGE("Batch No./DC No", "Item Version Parameters-Result"."Batch No./DC No");
                        IF IVP.FINDFIRST THEN BEGIN
                            "BlendOrderNo." := IVP."Blend Order No";
                        END;
                        //EBT STIVAN (30052012)-------------------------------------------------------------END


                        CLEAR(BatchDate);//25Oct2017
                        ItemVarParameters.RESET;
                        ItemVarParameters.SETRANGE("Item No.", "Item Version Parameters-Result"."Item No.");
                        ItemVarParameters.SETRANGE("Batch No./DC No", "Item Version Parameters-Result"."Batch No./DC No");
                        /*IF ItemVarParameters.FINDFIRST THEN
                        BEGIN
                          ILE1.RESET;
                          ILE1.SETRANGE(ILE1."Document No.",ItemVarParameters."Blend Order No");
                          IF ILE1.FINDFIRST THEN
                             BatchDate := ILE1."Posting Date";
                        END;*/

                        //Fahim MFG date to be captured from Finish prod. Due date as discussed with Kiran Rathore email confirmation
                        IF ItemVarParameters.FINDFIRST THEN BEGIN
                            BlendHdr.RESET;
                            BlendHdr.SETRANGE(BlendHdr."No.", ItemVarParameters."Blend Order No");
                            IF BlendHdr.FINDFIRST THEN
                                BatchDate := BlendHdr."Due Date";
                            //MESSAGE('%1',ItemVarParameters."Blend Order No");
                            //MESSAGE('%1',BlendHdr."Due Date");

                        END;
                        //End Fahim
                        //<<1

                        //>>2

                        //Item Version Parameters-Result, Body (2) - OnPreSection()
                        ctr += 1;
                        //MESSAGE('%1 Ctr',ctr);

                        //EBT STIVAN (30052012)-----------------------------------------------------------START
                        "Item Version Parameters-Result".SETRANGE("Item Version Parameters-Result"."Blend Order No", "BlendOrderNo.");
                        //EBT STIVAN (30052012)-------------------------------------------------------------END
                        //<<2

                    end;

                    trigger OnPreDataItem()
                    begin

                        SNo := 0;//10Mar2017

                        NAH := COUNT;//10Mar2017
                    end;
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    //>>RSPL
                    IF "Item Category Code" = 'CAT13' THEN
                        ERROR('QC Test Result can not be printed for Transformer Oils and White Oils');
                    //<<RSPL
                    //RSPL008 >>>> Item Category wise IPOL Description has print
                    vItemDesc := '';
                    IF ("Sales Invoice Line"."Item Category Code" = 'CAT02')
                      OR ("Sales Invoice Line"."Item Category Code" = 'CAT03') OR ("Sales Invoice Line"."Item Category Code" = 'CAT11')
                      OR ("Sales Invoice Line"."Item Category Code" = 'CAT12') OR ("Sales Invoice Line"."Item Category Code" = 'CAT13') THEN
                        vItemDesc := 'IPOL ' + Description
                    ELSE
                        vItemDesc := Description;
                    //RSPL008 <<<<

                    Qty1 := Quantity;
                    QtyPerUOM := "Qty. per Unit of Measure";

                    IF "Sales Invoice Line"."Unit of Measure Code" = 'BRL' THEN BEGIN
                        PackDescription := FORMAT(Qty1) + '*' +
                        FORMAT("Sales Invoice Line"."Qty. per Unit of Measure") + '' + "Sales Invoice Line"."Unit of Measure Code";
                    END ELSE BEGIN
                        PackDescription := FORMAT(Qty1) + '*' + "Sales Invoice Line"."Unit of Measure Code";
                    END;

                    Item.GET("Sales Invoice Line"."No.");
                    //<<1


                    //>>2

                    //Sales Invoice Line, Body (1) - OnPreSection()
                    MinDate := MinDate;
                    MaxDate := MaxDate;
                    ILE.RESET;
                    ILE.SETRANGE(ILE."Entry Type", ILE."Entry Type"::"Positive Adjmt.", ILE."Entry Type"::Output);
                    IF ILE.FINDLAST THEN BEGIN
                        //TestDate:=ILE."Posting Date";
                        Loc.RESET;
                        Loc.GET(ILE."Location Code");
                    END;

                    recIVP.RESET;
                    recIVP.SETRANGE(recIVP."Item No.", "Sales Invoice Line"."No.");
                    recIVP.SETRANGE(recIVP."Batch No./DC No", "Sales Invoice Line"."Lot No.");
                    IF recIVP.FINDFIRST THEN BEGIN
                        TestDate := recIVP."Posting Date";
                    END;

                    /*
                    "Item Version Parameters-Result".RESET;
                    "Item Version Parameters-Result".SETRANGE("Item Version Parameters-Result"."Batch No./DC No","Sales Invoice Line"."Lot No.");
                    "Item Version Parameters-Result".SETRANGE("Item Version Parameters-Result"."QC Test Report Generated",FALSE);
                    IF "Item Version Parameters-Result".FINDSET THEN
                    BEGIN
                      REPEAT
                       "Item Version Parameters-Result"."QC Test Report Generated":=TRUE;
                       "Item Version Parameters-Result".MODIFY;
                      UNTIL "Item Version Parameters-Result".NEXT=0;
                    END;
                    *///Commented 23May2017

                    //>>23May2017  RB-N NewCode for QCReportGeneration
                    ItemVarPar23.RESET;
                    ItemVarPar23.SETRANGE("Batch No./DC No", "Sales Invoice Line"."Lot No.");
                    ItemVarPar23.SETRANGE("QC Test Report Generated", FALSE);
                    IF ItemVarPar23.FINDSET THEN BEGIN
                        REPEAT

                            //ItemVarPar23.VALIDATE("QC Test Report Generated",TRUE);
                            //ItemVarPar23."QC Test Report Generated" := TRUE;
                            //ItemVarPar23.MODIFY;
                            ItemVarPar23.MODIFYALL("QC Test Report Generated", TRUE);
                        UNTIL ItemVarPar23.NEXT = 0;

                    END;
                    //<<23May2017 RB-N NewCode for QCReportGeneration

                    //<<2

                end;

                trigger OnPreDataItem()
                begin

                    //>>1

                    IF "LineNo." <> 0 THEN
                        "Sales Invoice Line".SETRANGE("Sales Invoice Line"."Line No.", "LineNo.");
                    //<<1

                    SNo := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>RB-N 25Oct2017 Logo Baseon ItemCategory
                CLEAR(Cat21);
                CLEAR(Logo21);
                SL25.RESET;
                SL25.SETRANGE("Document No.", "No.");
                SL25.SETRANGE(Type, SL25.Type::Item);
                IF SL25.FINDFIRST THEN BEGIN
                    Cat21 := SL25."Item Category Code";

                END;

                IF Cat21 = 'CAT15' THEN
                    Logo21 := 1
                ELSE
                    Logo21 := 3;
                //>>RB-N 25Oct2017 Logo Baseon ItemCategory
                //>>1
                /*
                //EBT STIVAN---(29012013)---Error Message POP UP,if ECC No of Location Code is not Specified----START
                recLocation.RESET;
                recLocation.SETRANGE(recLocation.Code,"Sales Invoice Header"."Location Code");
                IF recLocation.FINDFIRST THEN
                BEGIN
                 recECC.RESET;
                 recECC.SETRANGE(recECC.Code,recLocation."E.C.C. No.");
                 IF  recECC.FINDFIRST THEN
                 BEGIN
                  IF recECC.Description = '' THEN
                  ERROR('Your Location is not Excise Registered');
                 END;
                END;
                //EBT STIVAN---(29012013)---Error Message POP UP,if ECC No of Location Code is not Specified------END
                */
                //<<1


                //>>2

                //Sales Invoice Header, Header (1) - OnPreSection()
                /*
                recSIL.RESET;
                recSIL.SETRANGE(recSIL."Document No.","Sales Invoice Header".GETFILTER("Sales Invoice Header"."No."));
                IF recSIL.FINDFIRST THEN
                BEGIN
                 ProductionOrder.RESET;
                 ProductionOrder.SETRANGE(ProductionOrder."Description 2",recSIL."Lot No.");
                 IF ProductionOrder.FINDFIRST THEN
                  Loc1.RESET;
                  Loc1.SETRANGE(Loc1.Code,ProductionOrder."Location Code");
                  IF Loc1.FINDFIRST THEN;
                   RecState.RESET;
                   IF RecState.GET(Loc."State Code") THEN;
                END;
                */

                //RB-N  10Nov2017
                IF Loc1.GET("Location Code") THEN BEGIN
                    RecState.RESET;
                    IF RecState.GET(Loc1."State Code") THEN;
                END;
                ctr := 0;


                IF RECCOMPINFO.GET THEN;

                //<<2

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

    trigger OnPostReport()
    begin

        //>>1

        //EBT STIVAN---(17042013)---To Update the Field,if QC TEST REPORT Generated is FALSE in Sales Invoice Header----START
        IF "Sales Invoice Header"."QC Test Report Generated" = FALSE THEN BEGIN
            "Sales Invoice Header"."QC Test Report Generated" := TRUE;
            "Sales Invoice Header".MODIFY;
        END;
        //EBT STIVAN---(17042013)---To Update the Field,if QC TEST REPORT Generated is FALSE in Sales Invoice Header------END

        //<<1
    end;

    trigger OnPreReport()
    begin

        //>>1
        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
        CompanyInfo1.CALCFIELDS("Repsol Logo");//25Oct2017
        //<<1
    end;

    var
        ItemNo: Code[20];
        "Variant Code": Code[20];
        Result: array[5] of Code[30];
        ItemVarParameters: Record 50015;
        ctr: Integer;
        ItemVarParametersResult: Record 50015;
        "Batch No": Code[30];
        Invno: Code[20];
        salinvlin: Record 113;
        Qty: Integer;
        Date: Date;
        salinvhdr: Record 112;
        grade: Code[50];
        CompanyInfo1: Record 79;
        Loc: Record 14;
        SalesPerson: Record 13;
        RecState: Record state;
        RecState1: Record State;
        RecCountry: Record 9;
        RecCountry1: Record 9;
        TestDate: Date;
        Remarks: Text[500];
        UserSetup: Record 91;
        UserName: Text[100];
        User: Record 2000000120;
        Remarks1: Text[250];
        Qty1: Decimal;
        QtyPerUOM: Decimal;
        PackDescription: Text[30];
        BlendHdr: Record 5405;
        ILE: Record 32;
        MaxDate: Date;
        MinDate: Date;
        Item: Record 27;
        Location: Record 14;
        batno: Code[20];
        Loc1: Record 14;
        ProductionOrder: Record 5405;
        BatchNo: Code[10];
        ILE1: Record 32;
        BatchDate: Date;
        recSalesline: Record 113;
        itemversion: Record 50025;
        "LineNo.": Integer;
        IVP: Record 50015;
        "BlendOrderNo.": Code[20];
        recSIL: Record 113;
        recLocation: Record 14;
        //recECC: Record 13708;
        recIVP: Record 50015;
        RECCOMPINFO: Record 79;
        "------RSPL------": Integer;
        vItemDesc: Text[250];
        "--10Mar2017": Integer;
        SNo: Integer;
        NAH: Integer;
        vItemVerParameter: Record 50015;
        "-----23May2017": Integer;
        ItemVarPar23: Record 50015;
        "-------25Oct2017": Integer;
        SL25: Record 113;
        Cat21: Code[20];
        Logo21: Integer;
}


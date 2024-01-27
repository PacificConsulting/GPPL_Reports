report 50090 "QC Test Report for Plant"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/QCTestReportforPlant.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Item Version Parameters-Result"; "Item Version Parameters-Result")
        {
            DataItemTableView = SORTING("Item No.", "Version Code", "Line No.", Parameter, "Batch No./DC No")
                                WHERE(Result = FILTER(<> ''),
                                      "Test Result Approved" = FILTER(true));
            RequestFilterFields = "Item No.", "Batch No./DC No", "Posting Date";
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
            column(IsoText; IsoText)
            {
            }
            column(BatchNo; "Item Version Parameters-Result"."Batch No./DC No")
            {
            }
            column(LineNo; "Line No.")
            {
            }
            column(BatchDate; BatchDate)
            {
            }
            column(Grade; 'IPOL  ' + "Item Version Parameters-Result"."Item Description")
            {
            }
            column(GradeText; GradeText)
            {
            }
            column(VersionCode; recProdBomVersion.Description)
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
            column(ApprovedBy; "Approved By" + ' - ' + ApprovedUser)
            {
            }
            column(LocAdd; Loc.Address + ' ' + Loc."Address 2" + ' ' + Loc.City + ' ' + Loc."Post Code" + ' ' + Loc."Phone No." + 'E-mail : ' + ' ' + Loc."E-Mail")
            {
            }
            column(NAH; NAH)
            {
            }

            trigger OnAfterGetRecord()
            begin
                SNo += 1;//02Mar2017

                //>>19Sep2017 Logo Baseon ItemCategory
                CLEAR(Cat21);
                CLEAR(Logo21);

                Itm21.RESET;
                Itm21.SETRANGE("No.", "Item No.");
                IF Itm21.FINDFIRST THEN BEGIN
                    Cat21 := Itm21."Item Category Code";
                END;

                IF Cat21 = 'CAT15' THEN
                    Logo21 := 1
                ELSE
                    Logo21 := 3;

                GradeText := '';
                IF Cat21 = 'CAT15' THEN
                    GradeText := "Item Description"
                ELSE
                    GradeText := 'IPOL ' + "Item Description";

                //>>19Sep2017 Logo Baseon ItemCategory
                //>>1

                IVP.RESET;
                IVP.SETRANGE("Item No.", "Item Version Parameters-Result"."Item No.");
                IVP.SETRANGE("Batch No./DC No", "Item Version Parameters-Result"."Batch No./DC No");
                IF IVP.FINDFIRST THEN BEGIN
                    "BlendOrderNo." := IVP."Blend Order No";
                END;

                recProdOrderLine.RESET;
                recProdOrderLine.SETRANGE(recProdOrderLine."Inventory Posting Group", 'BULK');
                recProdOrderLine.SETRANGE(recProdOrderLine."Description 2", "Item Version Parameters-Result"."Batch No./DC No");
                IF recProdOrderLine.FINDFIRST THEN BEGIN
                    recProdBomVersion.RESET;
                    recProdBomVersion.SETRANGE(recProdBomVersion."Production BOM No.", recProdOrderLine."Item No.");
                    recProdBomVersion.SETRANGE(recProdBomVersion."Version Code", recProdOrderLine."Production BOM Version Code");
                    IF recProdBomVersion.FINDFIRST THEN;
                END;

                recProdBomVersion.RESET;
                recProdBomVersion.SETRANGE(recProdBomVersion."Production BOM No.", "Item Version Parameters-Result"."Item No.");
                recProdBomVersion.SETRANGE(recProdBomVersion."Version Code", "Item Version Parameters-Result"."Version Code");
                IF recProdBomVersion.FINDFIRST THEN;

                ILE1.RESET;
                ILE1.SETRANGE(ILE1."Document No.", "Item Version Parameters-Result"."Blend Order No");
                IF ILE1.FINDFIRST THEN
                    BatchDate := ILE1."Posting Date";

                CLEAR(TestedUser);
                UserSetup.RESET;
                UserSetup.SETRANGE(UserSetup."User ID", "Item Version Parameters-Result"."Tested By");
                IF UserSetup.FINDFIRST THEN
                    TestedUser := UserSetup.Name;

                CLEAR(ApprovedUser);
                UserSetup.RESET;
                UserSetup.SETRANGE(UserSetup."User ID", "Item Version Parameters-Result"."Approved By");
                IF UserSetup.FINDFIRST THEN
                    ApprovedUser := UserSetup.Name;
                //<<1


                //>>2

                //Item Version Parameters-Result, Header ( - OnPreSection()
                Item.RESET;
                Item.GET("Item Version Parameters-Result"."Item No.");


                MinDate := "Item Version Parameters-Result".GETRANGEMIN("Item Version Parameters-Result"."Posting Date");
                MaxDate := "Item Version Parameters-Result".GETRANGEMAX("Item Version Parameters-Result"."Posting Date");

                BlendHdr.RESET;
                BlendHdr.RESET;
                BlendHdr.SETRANGE(BlendHdr."Description 2", "Item Version Parameters-Result"."Batch No./DC No");
                BlendHdr.SETFILTER(BlendHdr."Starting Date", '>=%1', MinDate);
                BlendHdr.SETFILTER(BlendHdr."Starting Date", '<=%1', MaxDate);
                IF BlendHdr.FINDFIRST THEN BEGIN
                    IF Loc.GET(BlendHdr."Location Code") THEN;
                    IF BlendHdr."Location Code" = 'PLANT01' THEN
                        IsoText := 'ISO 9001 & EMS 14001 CERTIFIED COMPANY';
                    ItemName := BlendHdr.Description;
                END;


                ILE.RESET;
                ILE.SETRANGE(ILE."Entry Type", ILE."Entry Type"::"Positive Adjmt.");
                ILE.SETRANGE(ILE."Lot No.", "Item Version Parameters-Result"."Batch No./DC No");
                ILE.SETRANGE(ILE."Item No.", "Item Version Parameters-Result"."Item No.");
                ILE.SETFILTER(ILE."Posting Date", '>=%1', MinDate);
                ILE.SETFILTER(ILE."Posting Date", '<=%1', MaxDate);

                IF ILE.FINDLAST THEN BEGIN
                    BatchDate := ILE."Posting Date";
                    Loc.RESET;
                    Loc.GET(ILE."Location Code");
                END;

                //<<2

                //>>3

                ctr += 1;

                //"Item Version Parameters-Result".SETRANGE("Item Version Parameters-Result"."Blend Order No","BlendOrderNo.");

                ItemVarParametersResult.RESET;
                ItemVarParametersResult.SETRANGE(ItemVarParametersResult."Batch No./DC No", "Item Version Parameters-Result"."Batch No./DC No");
                IF ItemVarParametersResult.FINDFIRST THEN
                    REPEAT
                        ItemVarParametersResult.Remarks := Remakrs;
                        ItemVarParametersResult."Remarks 1" := Remarks1;
                        ItemVarParametersResult.MODIFY;
                    UNTIL ItemVarParametersResult.NEXT = 0;

                //<<3
            end;

            trigger OnPostDataItem()
            begin

                //MESSAGE(' NAH %1 SSS %2',NAH,SNo);
            end;

            trigger OnPreDataItem()
            begin

                SNo := 0;//02Mar2017

                NAH := COUNT;//02Mar2017
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(Remarks; Remakrs)
                {
                }
                field("Remarks 1"; Remarks1)
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

    trigger OnPreReport()
    begin
        //>>1

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
        CompanyInfo1.CALCFIELDS(CompanyInfo1."Shaded Box");
        CompanyInfo1.CALCFIELDS("Repsol Logo");//19Sep2017
        //<<1
    end;

    var
        ItemNo: Code[20];
        "Variant Code": Code[20];
        Result: array[5] of Code[30];
        ItemVarParameters: Record 50011;
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
        RecState: Record "state";
        RecState1: Record "state";
        RecCountry: Record 9;
        RecCountry1: Record 9;
        TestDate: Date;
        Remakrs: Text[500];
        UserSetup: Record 91;
        UserName: Text[100];
        Item: Record 27;
        Remarks1: Text[200];
        BlendHdr: Record 5405;
        ILE: Record 32;
        BatchDate: Date;
        MaxDate: Date;
        MinDate: Date;
        Location: Record 14;
        IsoText: Text[50];
        ILE1: Record 32;
        TestedUser: Text[30];
        ApprovedUser: Text[30];
        IVP: Record 50015;
        "BlendOrderNo.": Code[20];
        ItemName: Text[50];
        recProdBomVersion: Record 99000779;
        recProdOrderLine: Record 5406;
        "--02Mar17": Integer;
        SNo: Integer;
        NAH: Integer;
        "------19Sep2017": Integer;
        Cat21: Code[20];
        Logo21: Integer;
        SL21: Record 113;
        Itm21: Record 27;
        GradeText: Text[100];
}


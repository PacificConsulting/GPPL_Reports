report 50168 "Item Specification"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/ItemSpecification.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Item Version Parameters"; "Item Version Parameters")
        {
            DataItemTableView = SORTING("Item No.", "Version Code", "Line No.", "Version Descrption")
                                ORDER(Ascending);
            RequestFilterFields = "Item No.", "Version Code";
            column(COMPANYNAME; COMPANYNAME)
            {
            }
            column(IsoText; IsoText)
            {
            }
            column(ItemDesc; "Item Version Parameters"."Item No." + ' - ' + vItemDesc)
            {
            }
            column(VersionCode; versioncode + ' - ' + recVersion.Description)
            {
            }
            column(SrNo; SrNo)
            {
            }
            column(Parameter; Parameter)
            {
            }
            column(TypicalValue_ItemVersionParameters; "Item Version Parameters"."Typical Value")
            {
            }
            column(MinValue_ItemVersionParameters; "Item Version Parameters"."Min Value")
            {
            }
            column(MaxValue_ItemVersionParameters; "Item Version Parameters"."Max Value")
            {
            }
            column(TestMethod_ItemVersionParameters; "Item Version Parameters"."Test Method")
            {
            }

            trigger OnAfterGetRecord()
            begin

                //RSPL008 >>>>
                vCountItemDesc := COPYSTR(item.Description, 1, 6);
                IF vCountItemDesc = 'REPSOL' THEN
                    vItemDesc := item.Description
                ELSE
                    vItemDesc := 'IPOL ' + item.Description;
                //RSPL008 <<<<
                SrNo := SrNo + 1;

                recVersion.RESET;
                recVersion.SETRANGE(recVersion."Production BOM No.", "Item Version Parameters"."Item No.");
                recVersion.SETRANGE(recVersion."Version Code", versioncode);
                IF recVersion.FINDFIRST THEN;
            end;

            trigger OnPreDataItem()
            begin

                IF item.GET("Item Version Parameters".GETFILTER("Item Version Parameters"."Item No.")) THEN;
                versioncode := "Item Version Parameters".GETFILTER("Item Version Parameters"."Version Code");

                companyinfo1.GET;
                companyinfo1.CALCFIELDS(companyinfo1.Picture);
                companyinfo1.CALCFIELDS(companyinfo1."Name Picture");
                companyinfo1.CALCFIELDS(companyinfo1."Shaded Box");
                LastFieldNo := FIELDNO("Version Code");
                IsoText := 'ISO 9001 & EMS 14001 CERTIFIED COMPANY';
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
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        SrNo: Integer;
        item: Record 27;
        versioncode: Code[30];
        companyinfo1: Record 79;
        Loc: Record 14;
        ILE: Record 32;
        IsoText: Text[50];
        recVersion: Record 99000779;
        "-----RSPL-----": Integer;
        vCountItemDesc: Text[30];
        vItemDesc: Text[250];
}


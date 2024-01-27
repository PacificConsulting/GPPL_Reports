report 50209 "Update VendInvNo"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/UpdateVendInvNo.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Permissions = TableData 17 = rm,
                  TableData 25 = rm,
                  TableData 122 = rm,
                  TableData 5802 = rm,
                  TableData "GST Ledger Entry" = rm,
                  TableData "Detailed GST Ledger Entry" = rm;

    dataset
    {
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number)
                                WHERE(Number = CONST(1));

            trigger OnAfterGetRecord()
            begin
                RecPIH.RESET;
                RecPIH.SETCURRENTKEY("No.");
                IF RecPIH.GET(DocNo) THEN BEGIN
                    RecPIH."Vendor Invoice No." := VendInvNo;
                    RecPIH.MODIFY;

                    RecGLE.RESET;
                    RecGLE.SETCURRENTKEY("Document No.");
                    RecGLE.SETRANGE("Document No.", RecPIH."No.");
                    IF RecGLE.FINDFIRST THEN BEGIN
                        RecGLE.MODIFYALL("External Document No.", VendInvNo);
                    END;

                    RecVLE.RESET;
                    RecVLE.SETCURRENTKEY("Document No.");
                    RecVLE.SETRANGE("Document No.", RecPIH."No.");
                    IF RecVLE.FINDFIRST THEN BEGIN
                        RecVLE.MODIFYALL("External Document No.", VendInvNo);
                    END;

                    RecVE.RESET;
                    RecVE.SETCURRENTKEY("Document No.");
                    RecVE.SETRANGE("Document No.", RecPIH."No.");
                    IF RecVE.FINDFIRST THEN BEGIN
                        RecVE.MODIFYALL("External Document No.", VendInvNo);
                    END;

                    RecGSTLE.RESET;
                    RecGSTLE.SETCURRENTKEY("Document No.");
                    RecGSTLE.SETRANGE("Document No.", RecPIH."No.");
                    IF RecGSTLE.FINDFIRST THEN BEGIN
                        RecGSTLE.MODIFYALL("External Document No.", VendInvNo);
                    END;

                    RecDetGSTLE.RESET;
                    RecDetGSTLE.SETCURRENTKEY("Document No.");
                    RecDetGSTLE.SETRANGE("Document No.", RecPIH."No.");
                    IF RecDetGSTLE.FINDFIRST THEN BEGIN
                        RecDetGSTLE.MODIFYALL("External Document No.", VendInvNo);
                    END;

                END;
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("Document No."; DocNo)
                {
                    ApplicationArea = all;
                    TableRelation = "Purch. Inv. Header"."No.";
                }
                field("Vendor Invoice No."; VendInvNo)
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
        MESSAGE('done');
    end;

    trigger OnPreReport()
    begin
        IF DocNo = '' THEN
            ERROR('Please enter Document No.');
        IF VendInvNo = '' THEN
            ERROR('Please enter Vendor Invoice No.');
    end;

    var
        RecPIH: Record 122;
        RecGLE: Record 17;
        RecVLE: Record 25;
        RecVE: Record 5802;
        RecGSTLE: Record "GST Ledger Entry";
        RecDetGSTLE: Record "Detailed GST Ledger Entry";
        DocNo: Code[20];
        VendInvNo: Code[35];
}


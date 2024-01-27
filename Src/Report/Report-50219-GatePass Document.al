report 50219 "GatePass Document"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/GatePassDocument.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.")
                                WHERE("Gate Pass" = CONST(true));
            RequestFilterFields = "No.";
            column(BuyfromVendorNo_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Buy-from Vendor No.")
            {
            }
            column(No_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."No.")
            {
            }
            column(GatePass_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Gate Pass")
            {
            }
            column(GatePassStatus_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Gate Pass Status")
            {
            }
            column(ModeofTransport_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Mode of Transport")
            {
            }
            column(RequestorsDept_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Requestor's Dept")
            {
            }
            column(ReturnMaterialDate_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Return Material Date")
            {
            }
            column(PurposeofGatePass_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Purpose of GatePass")
            {
            }
            column(GatePassDesc; GatePassDesc)
            {
            }
            column(LocAdd1; Loc.Address)
            {
            }
            column(LocAdd2; Loc."Address 2")
            {
            }
            column(LocAdd3; Loc."Post Code" + ' - ' + Loc.City)
            {
            }
            column(CompName; CI.Name)
            {
            }
            column(PostingDate_PurchCrMemoHdr; "Purch. Cr. Memo Hdr."."Posting Date")
            {
            }
            column(VendorCode; "Purch. Cr. Memo Hdr."."Buy-from Vendor No.")
            {
            }
            column(VendorName; "Purch. Cr. Memo Hdr."."Buy-from Vendor Name")
            {
            }
            column(VenAdd1; "Purch. Cr. Memo Hdr."."Buy-from Address")
            {
            }
            column(VenAdd2; "Purch. Cr. Memo Hdr."."Buy-from Address 2")
            {
            }
            column(VenAdd3; "Buy-from City" + ' - ' + "Buy-from Post Code")
            {
            }
            column(VenPhNo; VenPhNo)
            {
            }
            dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.");
                column(DirectUnitCost_PurchCrMemoLine; "Purch. Cr. Memo Line"."Direct Unit Cost")
                {
                }
                column(Amount_PurchCrMemoLine; "Purch. Cr. Memo Line".Amount)
                {
                }
                column(UnitPriceLCY_PurchCrMemoLine; "Purch. Cr. Memo Line"."Unit Price (LCY)")
                {
                }
                column(LineAmount_PurchCrMemoLine; "Purch. Cr. Memo Line"."Line Amount")
                {
                }
                column(UnitofMeasureCode_PurchCrMemoLine; "Purch. Cr. Memo Line"."Unit of Measure Code")
                {
                }
                column(GrossWeight_PurchCrMemoLine; "Purch. Cr. Memo Line"."Gross Weight")
                {
                }
                column(NetWeight_PurchCrMemoLine; "Purch. Cr. Memo Line"."Net Weight")
                {
                }
                column(Quantity_PurchCrMemoLine; "Purch. Cr. Memo Line".Quantity)
                {
                }
                column(No_PurchCrMemoLine; "Purch. Cr. Memo Line"."No.")
                {
                }
                column(Description_PurchCrMemoLine; "Purch. Cr. Memo Line".Description)
                {
                }
                column(Description2_PurchCrMemoLine; "Purch. Cr. Memo Line"."Description 2")
                {
                }
                column(DocumentNo_PurchCrMemoLine; "Purch. Cr. Memo Line"."Document No.")
                {
                }
                column(LineNo_PurchCrMemoLine; "Purch. Cr. Memo Line"."Line No.")
                {
                }
                column(SrNo; SrNo)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    SrNo += 1;
                end;

                trigger OnPreDataItem()
                begin
                    SrNo := 0;
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>
                GatePassDesc := '';
                IF "Gate Pass Status" = "Gate Pass Status"::Returnable THEN
                    GatePassDesc := 'RETURNABLE GATE PASS'
                ELSE
                    GatePassDesc := 'NON-RETURNABLE GATE PASS';

                Loc.RESET;
                IF Loc.GET("Location Code") THEN;

                VenPhNo := '';
                Ven.RESET;
                IF Ven.GET("Purch. Cr. Memo Hdr."."Buy-from Vendor No.") THEN
                    VenPhNo := Ven."Phone No.";
                //<<
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

    trigger OnPreReport()
    begin

        CI.GET;
        CI.CALCFIELDS(Picture);
    end;

    var
        CI: Record 79;
        GatePassDesc: Text;
        Loc: Record 14;
        SrNo: Integer;
        Ven: Record 23;
        VenPhNo: Text;
}


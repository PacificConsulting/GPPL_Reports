report 50058 "Loaded Vehicle Posted O/W"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/LoadedVehiclePostedOW.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Posted Gate Entry Header"; "Posted Gate Entry Header")
        {
            DataItemTableView = SORTING("Entry Type", "No.")
                                ORDER(Ascending);
            RequestFilterFields = "Entry Type", "No.";
            column(CompName; Companyinfo.Name)
            {
            }
            column(HeaderDate; FORMAT(TODAY, 0, 4))
            {
            }
            column(GateEntryNo; "Posted Gate Entry Header"."No.")
            {
            }
            column(DocDate; FORMAT("Posted Gate Entry Header"."Document Date", 0, 4))
            {
            }
            column(LocName; LocName)
            {
            }
            column(TransporterName; "Posted Gate Entry Header"."Transporter Name")
            {
            }
            column(VehicleNo; "Posted Gate Entry Header"."Vehicle No.")
            {
            }
            column(DriverName; "Posted Gate Entry Header"."Driver's Name")
            {
            }
            column(DriverMobNo; "Posted Gate Entry Header"."Driver's Mobile No.")
            {
            }
            column(LRNo; "Posted Gate Entry Header"."LR/RR No.")
            {
            }
            column(LRDate; "Posted Gate Entry Header"."LR/RR Date")
            {
            }
            column(VehicleCapacity; "Posted Gate Entry Header"."Vehicle Capacity")
            {
            }
            column(Purchlineqty; Purchlineqty)
            {
            }
            column(Utilized; FORMAT(Utilised) + '%')
            {
            }
            column(VehicleHeight; "Posted Gate Entry Header"."Vehicle Body Height/Length")
            {
            }
            column(CreatedDate; "Posted Gate Entry Header"."Document Date")
            {
            }
            column(CreatedTime; "Posted Gate Entry Header"."Document Time")
            {
            }
            column(ReleaseDate; "Posted Gate Entry Header"."Releasing Date")
            {
            }
            column(ReleaseTime; "Posted Gate Entry Header"."Releasing Time")
            {
            }
            column(UnloadingDate; "Posted Gate Entry Header"."Vehicle Taken Date")
            {
            }
            column(UnloadingTime; "Posted Gate Entry Header"."Vehicle Taken Time")
            {
            }
            column(PostingDate; "Posted Gate Entry Header"."Posting Date")
            {
            }
            column(PostingTime; "Posted Gate Entry Header"."Posting Time")
            {
            }
            column(CreatedUser; CreatedUser)
            {
            }
            column(ReleasedUser; ReleasedUser)
            {
            }
            column(Vehicletakenuser; Vehicletakenuser)
            {
            }
            column(Postuser; Postuser)
            {
            }
            dataitem("Posted Gate Entry Line"; "Posted Gate Entry Line")
            {
                DataItemLink = "Gate Entry No." = FIELD("No.");
                DataItemTableView = SORTING("Entry Type", "Gate Entry No.", "Line No.");
                dataitem("Transfer Shipment Header"; "Transfer Shipment Header")
                {
                    DataItemLink = "No." = FIELD("Source No.");
                    DataItemTableView = SORTING("Transfer-from Code")
                                        ORDER(Ascending);
                    column(TShipPostingDate; "Transfer Shipment Header"."Posting Date")
                    {
                    }
                    column(TShipName; "Transfer Shipment Header"."Transfer-to Name")
                    {
                    }
                    dataitem("Transfer Shipment Line"; "Transfer Shipment Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending);
                        column(TShipDocNo; "Document No.")
                        {
                        }
                        column(TQtyBase; "Quantity (Base)")
                        {
                        }
                    }
                }
                dataitem("Sales Shipment Header"; "Sales Shipment Header")
                {
                    DataItemLink = "No." = FIELD("Source No.");
                    DataItemTableView = SORTING("Location Code")
                                        ORDER(Ascending);
                    column(SShipName; "Bill-to Name" + ' - ' + "Bill-to City")
                    {
                    }
                    column(SShipPostingDate; "Posting Date")
                    {
                    }
                    dataitem("Sales Shipment Line"; "Sales Shipment Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Type = FILTER(Item));
                        column(SShipNo; "Document No.")
                        {
                        }
                        column(ShipBaseQty; "Quantity (Base)")
                        {
                        }
                    }
                }
                dataitem("Sales Invoice Header"; "Sales Invoice Header")
                {
                    DataItemLink = "No." = FIELD("Source No.");
                    DataItemTableView = SORTING("No.")
                                        ORDER(Ascending);
                    column(SInvName; "Bill-to Name" + ' - ' + "Bill-to City")
                    {
                    }
                    column(SInvPostingDate; "Posting Date")
                    {
                    }
                    dataitem("Sales Invoice Line"; "Sales Invoice Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemTableView = SORTING("Document No.", "Line No.")
                                            ORDER(Ascending)
                                            WHERE(Type = FILTER(Item));
                        column(SInvNo; "Document No.")
                        {
                        }
                        column(SInvBaseQty; "Quantity (Base)")
                        {
                        }
                    }
                }
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                //Posted Gate Entry Header, Body (2) - OnPreSection()
                RecLoca.RESET;
                RecLoca.SETRANGE(RecLoca.Code, "Posted Gate Entry Header"."Location Code");
                IF RecLoca.FINDFIRST THEN BEGIN
                    LocName := RecLoca.Name;
                END;
                //CurrReport.SHOWOUTPUT(NOT Gateentybody);
                //Gateentybody := TRUE;
                //<<1
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
        //>>1
        Companyinfo.GET;
        //<<1
    end;

    var
        TQty: Decimal;
        TQtyBase: Decimal;
        GroupTotQty: Decimal;
        GroupTotQtyBase: Decimal;
        ShipQty: Decimal;
        ShipBaseQty: Decimal;
        GroupShipQty: Decimal;
        GroupShiptBaseQty: Decimal;
        ILE: Record 32;
        LotNo: Code[20];
        LocCode: Code[100];
        PostDate: Text[100];
        PstGteEnt: Record "Posted Gate Entry Header";
        TOTQuantityBase: Decimal;
        TGroupShiptBaseQty: Decimal;
        TGroupTotQtyBase: Decimal;
        Utilised: Decimal;
        Companyinfo: Record 79;
        LocName: Text[30];
        Gateentybody: Boolean;
        RecLoca: Record 14;
        Shipheader: Boolean;
        Shipgpheader: Boolean;
        ShipBody: Boolean;
        Vehiclecapacity: Record 50004;
        RecUser: Record 2000000120;
        ReleasedUser: Text[80];
        CreatedUser: Text[80];
        Vehicletakenuser: Text[80];
        Postuser: Text[80];
        Purchlineqty: Decimal;
}


report 50056 "Unloaded Vehicle Posted I/W"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/UnloadedVehiclePostedIW.rdl';
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

                        trigger OnAfterGetRecord()
                        begin

                            //>>1

                            //Transfer Shipment Line, Body (2) - OnPreSection()
                            TQty += "Transfer Shipment Line".Quantity;
                            TQtyBase += "Transfer Shipment Line"."Quantity (Base)";
                            //<<1


                            //>>2

                            //Transfer Shipment Line, Footer (3) - OnPreSection()
                            GroupTotQty += TQty;
                            GroupTotQtyBase += TQtyBase;
                            //<<2
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin

                        //>>1

                        //Transfer Shipment Header, Body (4) - OnPreSection()
                        //CurrReport.SHOWOUTPUT(NOT ShipBody);
                        //ShipBody := TRUE;
                        //<<1

                        //>>2

                        //Transfer Shipment Header, GroupFooter (5 - OnPreSection()
                        TGroupTotQtyBase += GroupTotQtyBase;
                        //<<2
                    end;
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

                    trigger OnAfterGetRecord()
                    begin

                        //>>1

                        //Sales Shipment Header, GroupFooter (4) - OnPreSection()
                        TGroupShiptBaseQty += GroupShiptBaseQty;
                        //<<1
                    end;
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

                        trigger OnAfterGetRecord()
                        begin

                            //>>1

                            //Sales Invoice Line, Body (2) - OnPreSection()
                            ShipQty += "Sales Invoice Line".Quantity;
                            ShipBaseQty += "Sales Invoice Line"."Quantity (Base)";
                            //<<1

                            //>>2

                            //Sales Invoice Line, Footer (3) - OnPreSection()
                            GroupShipQty += Quantity;
                            GroupShiptBaseQty += "Quantity (Base)";
                            //<<2
                        end;
                    }

                    trigger OnAfterGetRecord()
                    begin

                        //>>1

                        //Sales Invoice Header, GroupHeader (2) - OnPreSection()
                        GroupShipQty := 0;
                        GroupShiptBaseQty := 0;
                        //<<1
                    end;
                }
                dataitem("Purchase Header"; "Purchase Header")
                {
                    DataItemLink = "No." = FIELD("Source No.");
                    DataItemTableView = SORTING("Document Type", "No.")
                                        ORDER(Ascending);
                    column(PurName; "Purchase Header"."Pay-to Name" + ' - ' + "Purchase Header"."Pay-to City")
                    {
                    }
                    column(PurPostingDate; "Posting Date")
                    {
                    }
                    dataitem("Purchase Line"; "Purchase Line")
                    {
                        DataItemLink = "Document No." = FIELD("No.");
                        DataItemTableView = SORTING("Document Type", "Document No.", "Line No.")
                                            ORDER(Ascending);
                        column(PurNo; "Document No.")
                        {
                        }
                        column(PurBaseQty; "Quantity (Base)")
                        {
                        }

                        trigger OnAfterGetRecord()
                        begin

                            //>>1

                            //Purchase Line, Footer (2) - OnPreSection()
                            Purchlineqty += "Purchase Line"."Quantity (Base)";
                            //<<1
                        end;
                    }
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    //Posted Gate Entry Line, Body (1) - OnPreSection()
                    //CurrReport.SHOWOUTPUT(NOT Shipheader);
                    //Shipheader := TRUE;
                    //<<1
                end;
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


                //>>2

                //Transfer Shipment Header, GroupHeader (3 - OnPreSection()
                GroupTotQty := 0;
                GroupTotQtyBase := 0;
                //CurrReport.SHOWOUTPUT(NOT Shipgpheader);
                //Shipgpheader := TRUE;
                //<<2


                //>>3

                //Posted Gate Entry Header, Footer (4) - OnPreSection()
                Vehiclecapacity.RESET;
                Vehiclecapacity.SETRANGE(Vehiclecapacity.Code, "Posted Gate Entry Header"."Vehicle Capacity");
                IF Vehiclecapacity.FINDFIRST THEN BEGIN
                    Utilised := ROUND(100 - ((Vehiclecapacity.Value - TOTQuantityBase) / Vehiclecapacity.Value) * 100, 0.01);
                END;

                RecUser.RESET;
                //RecUser.SETRANGE(RecUser."User ID","Posted Gate Entry Header"."Release User ID");
                RecUser.SETRANGE(RecUser."User Name", "Posted Gate Entry Header"."Release User ID");//28Mar2017
                IF RecUser.FINDFIRST THEN
                    ReleasedUser := RecUser."Full Name";//28Mar2017
                //ReleasedUser := RecUser.Name;

                RecUser.RESET;
                //RecUser.SETRANGE(RecUser."User ID","Posted Gate Entry Header"."Created User ID");
                RecUser.SETRANGE(RecUser."User Name", "Posted Gate Entry Header"."Created User ID");//28Mar2017
                IF RecUser.FINDFIRST THEN
                    CreatedUser := RecUser."Full Name";//28Mar2017
                //CreatedUser := RecUser.Name;

                RecUser.RESET;
                //RecUser.SETRANGE(RecUser."User ID","Posted Gate Entry Header"."Vehicle Taken user ID");
                RecUser.SETRANGE(RecUser."User Name", "Posted Gate Entry Header"."Vehicle Taken user ID");//28Mar2017
                IF RecUser.FINDFIRST THEN
                    Vehicletakenuser := RecUser."Full Name";//28Mar2017
                //Vehicletakenuser := RecUser.Name;

                RecUser.RESET;
                //RecUser.SETRANGE(RecUser."User ID","Posted Gate Entry Header"."Posted User ID");
                RecUser.SETRANGE(RecUser."User Name", "Posted Gate Entry Header"."Posted User ID");//28Mar2017
                IF RecUser.FINDFIRST THEN
                    Postuser := RecUser."Full Name";//28Mar2017
                //Postuser := RecUser.Name;
                //<<3
            end;

            trigger OnPostDataItem()
            begin

                //>>1

                //Posted Gate Entry Header, Footer (3) - OnPreSection()
                TOTQuantityBase := TGroupTotQtyBase + TGroupShiptBaseQty;
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


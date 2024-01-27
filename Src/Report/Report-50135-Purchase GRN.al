report 50135 "Purchase GRN"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseGRN.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Purch. Rcpt. Header"; "Purch. Rcpt. Header")
        {
            RequestFilterFields = "No.";
            column(Name_CompInfo; CompanyInfo1.Name)
            {
            }
            column(Location_3; Loc.Address + ' ' + Loc."Address 2" + ' ' + Loc.City + ' ' + Loc."Post Code" + ' ' + RecState1.Description + ' Phone: ' + Loc."Phone No." + ' Email: ' + Loc."E-Mail")
            {
            }
            column(No_PurchRcptHeader; "Purch. Rcpt. Header"."No.")
            {
            }
            column(BuyfromVendorName_PurchRcptHeader; "Purch. Rcpt. Header"."Buy-from Vendor Name")
            {
            }
            column(BuyfromAddress_PurchRcptHeader; "Purch. Rcpt. Header"."Buy-from Address")
            {
            }
            column(BuyfromAddress2_PurchRcptHeader; "Purch. Rcpt. Header"."Buy-from Address 2")
            {
            }
            column(BuyfromCity_PurchRcptHeader; "Purch. Rcpt. Header"."Buy-from City")
            {
            }
            column(BuyfromPostCode_PurchRcptHeader; "Purch. Rcpt. Header"."Buy-from Post Code")
            {
            }
            column(DivisionName; DivisionName)
            {
            }
            column(PostingDate_PurchRcptHeader; "Purch. Rcpt. Header"."Posting Date")
            {
            }
            column(BlanketOrderNo_PurchRcptHeader; "Purch. Rcpt. Header"."Blanket Order No.")
            {
            }
            column(OrderNo_PurchRcptHeader; "Purch. Rcpt. Header"."Order No.")
            {
            }
            column(VehicleNo; PostedGateEntryHeader."Vehicle No.")
            {
            }
            column(LRNo; PostedGateEntryHeader."LR/RR No.")
            {
            }
            column(LRDate; PostedGateEntryHeader."LR/RR Date")
            {
            }
            column(VendorShipmentNo_PurchRcptHeader; "Purch. Rcpt. Header"."Vendor Shipment No.")
            {
            }
            column(UserName; vUser."User Name")
            {
            }
            column(LocationCode_PurchRcptHeader; "Purch. Rcpt. Header"."Location Code")
            {
            }
            column(PostedUserIDName; PostedUserIDName)
            {
            }
            column(PostedUser; PostedUser)
            {
            }
            dataitem("Purch. Rcpt. Line"; "Purch. Rcpt. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Type = FILTER(= Item | "G/L Account"),
                                          Quantity = FILTER(<> 0));
                column(ItmNo; "No.")
                {
                }
                column(ItemName; ItemName)
                {
                }
                column(Quantity; Quantity)
                {
                }
                column(UnitofMeasureCode_PurchRcptLine; "Purch. Rcpt. Line"."Unit of Measure Code")
                {
                }
                column(DensityFactor_PurchRcptLine; "Purch. Rcpt. Line"."Density Factor")
                {
                }
                column(QuantityBase_PurchRcptLine; ROUND("Purch. Rcpt. Line"."Quantity (Base)", 1.0, '='))
                {
                }
                column(VendorQuantity_PurchRcptLine; "Purch. Rcpt. Line"."Vendor Quantity")
                {
                }
                column(VendorUnitofMeasure_PurchRcptLine; "Purch. Rcpt. Line"."Vendor Unit of Measure")
                {
                }
                column(VendorQty_DestinyFactor; ROUND(("Purch. Rcpt. Line"."Vendor Quantity" * "Purch. Rcpt. Line"."Density Factor"), 1.0, '='))
                {
                }
                column(UOM; UOM)
                {
                }
                column(ShortEcess; -((ROUND(("Purch. Rcpt. Line"."Vendor Quantity" * "Purch. Rcpt. Line"."Density Factor"), 1.0, '=')) - (ROUND("Purch. Rcpt. Line"."Quantity (Base)", 1.0, '='))))
                {
                }
                column(Description_PurchRcptLine; "Purch. Rcpt. Line".Description)
                {
                }
                column(Correction_PurchRcptLine; "Purch. Rcpt. Line".Correction)
                {
                }
                column(ExpectedQty; ROUND("Purch. Rcpt. Line"."Vendor Quantity" * "Purch. Rcpt. Line"."Density Factor", 1.0, '='))
                {
                }
                column(Type_PurchRcptLine; "Purch. Rcpt. Line".Type)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    SrNo := SrNo + 1;

                    IF "Purch. Rcpt. Line"."Vendor Item No." <> '' THEN BEGIN
                        ItemName := "Purch. Rcpt. Line"."Vendor Item No.";
                    END ELSE BEGIN
                        ItemName := "Purch. Rcpt. Line".Description;
                    END;

                    UOM := '';
                    Item.RESET;
                    Item.SETRANGE(Item."No.", "Purch. Rcpt. Line"."No.");
                    IF Item.FINDFIRST THEN BEGIN
                        UOM := Item."Base Unit of Measure";
                    END;
                end;
            }
            dataitem("Posted Whse. Receipt Line"; "Posted Whse. Receipt Line")
            {
                DataItemLink = "Posted Source No." = FIELD("No.");
                DataItemTableView = SORTING("No.", "Line No.")
                                    WHERE("Posted Source Document" = FILTER("Posted Receipt"));
                column(WHItmNo; "No.")
                {
                }
                column(UnitofMeasureCode_PostedWhseReceiptLine; "Posted Whse. Receipt Line"."Unit of Measure Code")
                {
                }
                column(Quantity_PostedWhseReceiptLine; "Posted Whse. Receipt Line".Quantity)
                {
                }
                column(Description_PostedWhseReceiptLine; "Posted Whse. Receipt Line".Description)
                {
                }
                column(DesityFactor_Whserec; "recWhse.RcptHeader"."Density Factor")
                {
                }
                column(Qty_Density_Qty; -((Quantity * "recWhse.RcptHeader"."Density Factor") - (Quantity)))
                {
                }
                column(VendorQty_WhseRec; "recWhse.RcptHeader"."Vendor Quantity")
                {
                }
                column(VendorQty_Density; "recWhse.RcptHeader"."Vendor Quantity" * "recWhse.RcptHeader"."Density Factor")
                {
                }
                column(ShortEcess_WhseRec; -(("recWhse.RcptHeader"."Vendor Quantity" * "recWhse.RcptHeader"."Density Factor") - ("Posted Whse. Receipt Line".Quantity)))
                {
                }
                column(ShortEcess_WhseRec2; -((Quantity * "recWhse.RcptHeader"."Density Factor") - (Quantity)))
                {
                }
                column(ShortEcess_WhseRec1; -((Quantity * "recWhse.RcptHeader"."Density Factor") - (Quantity)))
                {
                }
                column(Qty_recWhse; "recWhse.RcptHeader"."Vendor Quantity")
                {
                }
                column(ExpectedQty2_Whse; (Quantity * "recWhse.RcptHeader"."Density Factor"))
                {
                }
                column(VendorQty_Whse; "recWhse.RcptHeader"."Vendor Quantity")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    "recWhse.RcptHeader".RESET;
                    "recWhse.RcptHeader".SETRANGE("recWhse.RcptHeader"."No.", "Posted Whse. Receipt Line"."No.");
                    IF "recWhse.RcptHeader".FINDFIRST THEN;


                    Item.GET("Posted Whse. Receipt Line"."Item No.");
                    ItemName := Item.Description;
                    BaseUOM := Item."Base Unit of Measure";

                    "recWhse.RcptHeader".GET("Posted Whse. Receipt Line"."No.");

                    Item.GET("Posted Whse. Receipt Line"."Item No.");
                    ItemName := Item.Description;
                    BaseUOM := Item."Base Unit of Measure";

                    "recWhse.RcptHeader".GET("Posted Whse. Receipt Line"."No.");

                    Item.GET("Posted Whse. Receipt Line"."Item No.");
                    ItemName := Item.Description;
                    BaseUOM := Item."Base Unit of Measure";

                    "recWhse.RcptHeader".GET("Posted Whse. Receipt Line"."No.");
                    //MESSAGE('Qty=%1..desnsity=%2',"Posted Whse. Receipt Line".Quantity,"recWhse.RcptHeader"."Density Factor");
                end;
            }

            trigger OnAfterGetRecord()
            begin
                RespCentre.GET("Responsibility Center");
                Loc.GET("Location Code");
                IF RecState1.GET(Loc."State Code") THEN;

                // IF RecState.GET(state) THEN;

                //IF vUser.GET(USERID) THEN; //GUID Error
                // EBT MILAN 190214--- TO Show Vehicle and LR Detail from Gate Entry------------------------------------START
                PostedGateEntryLine.RESET;
                PostedGateEntryLine.SETRANGE(PostedGateEntryLine."Source No.", "Purch. Rcpt. Header"."Order No.");
                IF PostedGateEntryLine.FINDFIRST THEN BEGIN
                    PostedGateEntryHeader.RESET;
                    PostedGateEntryHeader.SETRANGE(PostedGateEntryHeader."No.", PostedGateEntryLine."Gate Entry No.");
                    IF PostedGateEntryHeader.FINDFIRST THEN;
                END;
                // EBT MILAN 190214--- TO Show Vehicle and LR Detail from Gate Entry--------------------------------------END
                recPstdGateEntryAttch.RESET;
                recPstdGateEntryAttch.SETRANGE("Receipt No.", "Purch. Rcpt. Header"."No.");
                IF recPstdGateEntryAttch.FINDFIRST THEN
                    recPstdGateEntryHdr.RESET;
                recPstdGateEntryHdr.SETRANGE("No.", recPstdGateEntryAttch."Gate Entry No.");
                IF recPstdGateEntryHdr.FINDFIRST THEN BEGIN
                    LRNo := recPstdGateEntryHdr."LR/RR No.";
                    LRDate := recPstdGateEntryHdr."LR/RR Date";
                    VehicleNo := recPstdGateEntryHdr."Vehicle No.";
                END;

                // RSPL-PARAG 03.03.2016
                RecDimension.RESET;
                RecDimension.SETRANGE(RecDimension.Code, "Purch. Rcpt. Header"."Shortcut Dimension 1 Code");
                IF RecDimension.FINDFIRST THEN
                    DivisionName := RecDimension.Name;
                // RSPL-PARAG 03.03.2016


                //>>06APr2017

                //Purch. Rcpt. Header, Footer (5) - OnPreSection()
                //EBT STIVAN ---(240112)--- To capture posted User ID and USer ID Name ----------------START
                User.RESET;
                User.SETRANGE(User."User Name", "Purch. Rcpt. Header"."User ID");//06Apr2017
                //User.SETRANGE(User."User ID","Purch. Rcpt. Header"."User ID");
                IF User.FINDFIRST THEN BEGIN
                    PostedUserIDName := User."Full Name";//06Apr2017
                    PostedUser := "Purch. Rcpt. Header"."User ID";//06Apr2017
                END;
                //PostedUserIDName := User.Name;
                //EBT STIVAN ---(240112)--- To capture posted User ID and USer ID Name ------------------END

                //<<06APr2017
            end;

            trigger OnPreDataItem()
            begin
                CompanyInfo1.CALCFIELDS(CompanyInfo1."Name Picture");
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

        CompanyInfo1.GET;
        CompanyInfo1.CALCFIELDS(CompanyInfo1.Picture);
    end;

    var
        CompanyInfo1: Record 79;
        RespCentre: Record 5714;
        Loc: Record 14;
        RecState: Record State;
        RecState1: Record State;
        SrNo: Integer;
        Item: Record 27;
        BaseUOM: Code[20];
        ItemName: Code[60];
        PostedUserIDName: Text[80];
        recPstdGateEntryAttch: Record "Posted Gate Entry Attachment";
        recPstdGateEntryHdr: Record "Posted Gate Entry Header";
        LRNo: Code[20];
        LRDate: Date;
        VehicleNo: Code[20];
        "recWhse.RcptHeader": Record 7318;
        UOM: Code[10];
        PostedGateEntryLine: Record "Posted Gate Entry Line";
        PostedGateEntryHeader: Record "Posted Gate Entry Header";
        DivisionName: Text[30];
        Dimension: Record 349;
        RecDimension: Record 349;
        "--": Integer;
        vUser: Record 2000000120;
        "---06Apr2017": Integer;
        User: Record 2000000120;
        PostedUser: Text[80];
}


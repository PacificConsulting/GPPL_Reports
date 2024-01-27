report 50244 "Purchase GRN - 1 test"
{
    // Date        Version      Remarks
    // .....................................................................................
    // 09Jan2018   RB-N         Report Detail Condition.
    // 02Apr2018   RB-N         Removed Rounding qty for the G/L Account
    // 19Jul2018   RB-N         Service Received Note Details for Work Order Values
    // 02May2019   RB-N         GateEntry No.
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/PurchaseGRN1test.rdl';
    Caption = 'Purchase GRN - 1';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    PreviewMode = PrintLayout;

    dataset
    {
        dataitem("Purch. Rcpt. Header"; "Purch. Rcpt. Header")
        {
            RequestFilterFields = "No.";
            column(GateEntryNo; GateEntryNo)
            {
            }
            column(GateEntryDate; FORMAT(GateEntryDate, 0, '<Day,2>/<Month,2>/<Year4>'))
            {
            }
            column(Text1; Text1)
            {
            }
            column(Text2; Text2)
            {
            }
            column(Text3; Text3)
            {
            }
            column(Text4; Text4)
            {
            }
            column(Text5; Text5)
            {
            }
            column(Text6; Text6)
            {
            }
            column(Text7; Text7)
            {
            }
            column(GRNText; GRNText)
            {
            }
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
            column(PostingDate_PurchRcptHeader; FORMAT("Posting Date", 0, '<Day,2>/<Month,2>/<Year4>'))
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
            column(LRDate; FORMAT(PostedGateEntryHeader."LR/RR Date", 0, '<Day,2>/<Month,2>/<Year4>'))
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
            column(txtPurchGRNRemark; txtPurchGRNRemark)
            {
            }
            dataitem("Purch. Rcpt. Line"; "Purch. Rcpt. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    ORDER(Ascending)
                                    WHERE(Type = FILTER(= Item | "G/L Account"),
                                          Quantity = FILTER(<> 0));
                column(DocumentNo_PurchRcptLine; "Purch. Rcpt. Line"."Document No.")
                {
                }
                column(LineNo_PurchRcptLine; "Purch. Rcpt. Line"."Line No.")
                {
                }
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
                column(BodyLine1; BodyLine1)
                {
                }
                column(BodyLine2; BodyLine2)
                {
                }
                column(BodyLine3; BodyLine3)
                {
                }
                column(PRLLine; PRLLine)
                {
                }
                column(ItmName09; ItmName09)
                {
                }
                column(VenQty09; VenQty09)
                {
                }
                column(VenUOM09; VenUOM09)
                {
                }
                column(DenFac09; DenFac09)
                {
                }
                column(ExpQty09; ExpQty09)
                {
                }
                column(RecdQty09; RecdQty09)
                {
                }
                column(RecdUOM09; RecdUOM09)
                {
                }
                column(ShortQty09; ShortQty09)
                {
                }

                trigger OnAfterGetRecord()
                begin

                    CLEAR(ItmName09);
                    CLEAR(VenQty09);
                    CLEAR(VenUOM09);
                    CLEAR(DenFac09);
                    CLEAR(ExpQty09);
                    CLEAR(RecdQty09);
                    CLEAR(RecdUOM09);
                    CLEAR(ShortQty09);
                    //>>1st Condition
                    CLEAR(BodyLine1);
                    IF ("Location Code" = 'PLANT01') AND (Type = Type::"G/L Account") THEN BEGIN
                        BodyLine1 := TRUE;

                        IF "Vendor Item No." <> '' THEN
                            ItmName09 := "Vendor Item No."
                        ELSE
                            ItmName09 := "Purch. Rcpt. Line".Description;

                        VenQty09 := Quantity;
                        VenUOM09 := "Unit of Measure Code";
                        DenFac09 := "Density Factor";
                        ExpQty09 := 0;
                        //RecdQty09 := ROUND("Purch. Rcpt. Line"."Quantity (Base)",1.0,'=');
                        RecdQty09 := "Purch. Rcpt. Line"."Quantity (Base)";//02Apr2018
                        RecdUOM09 := '';
                        ShortQty09 := 0;

                    END ELSE
                        BodyLine1 := FALSE;
                    //<<1st Condition


                    //>>BodyLine2
                    CLEAR(BodyLine2);
                    IF ("Location Code" <> 'PLANT01') AND ("Vendor Quantity" = 0) THEN BEGIN
                        BodyLine2 := TRUE;

                        IF "Vendor Item No." <> '' THEN
                            ItmName09 := "Vendor Item No."
                        ELSE
                            ItmName09 := "Purch. Rcpt. Line".Description;

                        VenQty09 := Quantity;
                        VenUOM09 := "Unit of Measure Code";
                        DenFac09 := "Density Factor";
                        ExpQty09 := 0;
                        IF Type <> Type::"G/L Account" THEN
                            RecdQty09 := ROUND("Purch. Rcpt. Line"."Quantity (Base)", 1.0, '=')
                        ELSE
                            RecdQty09 := "Purch. Rcpt. Line"."Quantity (Base)";//22Jun2018
                        RecdUOM09 := '';
                        ShortQty09 := 0;

                    END ELSE
                        BodyLine2 := FALSE;
                    //MESSAGE('%1 BodyLine2',BodyLine2);
                    //<<BodyLine2

                    //>>BodyLine3
                    CLEAR(BodyLine3);
                    IF ("Location Code" <> 'PLANT01') AND ("Vendor Quantity" <> 0) THEN BEGIN
                        BodyLine3 := TRUE;

                        IF "Vendor Item No." <> '' THEN
                            ItmName09 := "Vendor Item No."
                        ELSE
                            ItmName09 := "Purch. Rcpt. Line".Description;

                        UOM := '';
                        Item.RESET;
                        IF Item.GET("Purch. Rcpt. Line"."No.") THEN BEGIN
                            UOM := Item."Base Unit of Measure";
                        END;

                        VenQty09 := "Purch. Rcpt. Line"."Vendor Quantity";
                        VenUOM09 := "Purch. Rcpt. Line"."Vendor Unit of Measure";
                        DenFac09 := "Density Factor";
                        ExpQty09 := ROUND(("Purch. Rcpt. Line"."Vendor Quantity" * "Purch. Rcpt. Line"."Density Factor"), 1.0, '=');
                        IF Type <> Type::"G/L Account" THEN
                            RecdQty09 := ROUND("Purch. Rcpt. Line"."Quantity (Base)", 1.0, '=')
                        ELSE
                            RecdQty09 := "Purch. Rcpt. Line"."Quantity (Base)";//22Jun2018
                        RecdUOM09 := UOM;
                        ShortQty09 := -((ROUND(("Purch. Rcpt. Line"."Vendor Quantity" * "Purch. Rcpt. Line"."Density Factor"), 1.0, '='))
                                      - (ROUND("Purch. Rcpt. Line"."Quantity (Base)", 1.0, '=')));

                    END ELSE
                        BodyLine3 := FALSE;
                    //MESSAGE('%1 BodyLine3',BodyLine3);
                    //<<BodyLine3


                    //>>BodyLine4
                    CLEAR(BodyLine4);
                    IF ("Purch. Rcpt. Line".Correction = TRUE) THEN BEGIN
                        PWRL09.RESET;
                        PWRL09.SETCURRENTKEY("Posted Source No.");
                        PWRL09.SETRANGE("Posted Source Document", PWRL09."Posted Source Document"::"Posted Receipt");
                        PWRL09.SETRANGE("Posted Source No.", "Document No.");
                        PWRL09.SETRANGE("Line No.", "Line No.");
                        PWRL09.SETRANGE("Item No.", "No.");//29Jan2018
                        PWRL09.SETRANGE(Quantity, Quantity);//29Jan2018
                        IF PWRL09.FINDFIRST THEN BEGIN
                            PWRH09.RESET;
                            IF PWRH09.GET(PWRL09."No.") THEN
                                IF PWRH09."Vendor Quantity" <> 0 THEN BEGIN
                                    BodyLine4 := TRUE;
                                    ItmName09 := PWRL09.Description;

                                    VenQty09 := PWRL09.Quantity;
                                    VenUOM09 := PWRL09."Unit of Measure Code";
                                    DenFac09 := PWRH09."Density Factor";
                                    ExpQty09 := PWRL09.Quantity * PWRH09."Density Factor";
                                    RecdQty09 := PWRL09.Quantity;
                                    RecdUOM09 := '';
                                    ShortQty09 := -((PWRL09.Quantity * PWRH09."Density Factor") - (PWRL09.Quantity));
                                END;
                        END;
                    END ELSE
                        BodyLine4 := FALSE;
                    //MESSAGE('%1 BodyLine4',BodyLine4);
                    //<<BodyLine4

                    //>>BodyLine5
                    CLEAR(BodyLine5);
                    IF ("Purch. Rcpt. Line".Correction = FALSE) THEN BEGIN
                        PWRL09.RESET;
                        PWRL09.SETCURRENTKEY("Posted Source No.");
                        PWRL09.SETRANGE("Posted Source Document", PWRL09."Posted Source Document"::"Posted Receipt");
                        PWRL09.SETRANGE("Posted Source No.", "Document No.");
                        PWRL09.SETRANGE("Source Line No.", "Line No.");
                        IF PWRL09.FINDFIRST THEN BEGIN
                            PWRH09.RESET;
                            IF PWRH09.GET(PWRL09."No.") THEN
                                IF PWRH09."Vendor Quantity" <> 0 THEN BEGIN
                                    BodyLine5 := TRUE;
                                    ItmName09 := PWRL09.Description;

                                    VenQty09 := PWRH09."Vendor Quantity";
                                    VenUOM09 := PWRL09."Unit of Measure Code";
                                    DenFac09 := PWRH09."Density Factor";
                                    ExpQty09 := PWRH09."Vendor Quantity" * PWRH09."Density Factor";
                                    RecdQty09 := Quantity;
                                    RecdUOM09 := '';
                                    ShortQty09 := -((PWRH09."Vendor Quantity" * PWRH09."Density Factor") - (PWRL09.Quantity));

                                END;
                        END;
                    END ELSE
                        BodyLine5 := FALSE;
                    //MESSAGE('%1 BodyLine5',BodyLine5);
                    //<<BodyLine5

                    //>>BodyLine6
                    CLEAR(BodyLine6);
                    PWRL09.RESET;
                    PWRL09.SETCURRENTKEY("Posted Source No.");
                    PWRL09.SETRANGE("Posted Source Document", PWRL09."Posted Source Document"::"Posted Receipt");
                    PWRL09.SETRANGE("Posted Source No.", "Document No.");
                    PWRL09.SETRANGE("Source Line No.", "Line No.");
                    IF PWRL09.FINDFIRST THEN BEGIN
                        PWRH09.RESET;
                        IF PWRH09.GET(PWRL09."No.") THEN BEGIN
                            IF PWRH09."Vendor Quantity" = 0 THEN BEGIN
                                BodyLine6 := TRUE;
                                ItmName09 := PWRL09.Description;

                                VenQty09 := PWRL09.Quantity;
                                VenUOM09 := PWRL09."Unit of Measure Code";
                                DenFac09 := PWRH09."Density Factor";
                                ExpQty09 := PWRL09.Quantity * PWRH09."Density Factor";
                                RecdQty09 := PWRL09.Quantity;
                                RecdUOM09 := '';
                                ShortQty09 := -((PWRL09.Quantity * PWRH09."Density Factor") - (PWRL09.Quantity));

                            END ELSE
                                BodyLine6 := FALSE;
                        END;
                    END;

                    SrNo := SrNo + 1;

                    //>>19Jul2018
                    IF "Purch. Rcpt. Header"."Work Order" THEN BEGIN

                        PL19.RESET;
                        PL19.SETRANGE("Document Type", PL19."Document Type"::"Blanket Order");
                        PL19.SETRANGE("Document No.", "Blanket Order No.");
                        PL19.SETRANGE("Line No.", "Blanket Order Line No.");
                        IF PL19.FINDFIRST THEN BEGIN
                            VenQty09 := PL19."Line Amount";
                            VenUOM09 := '';
                        END;

                        PRL19.RESET;
                        PRL19.SETCURRENTKEY("Blanket Order No.", "Blanket Order Line No.");
                        PRL19.SETRANGE("Blanket Order No.", "Blanket Order No.");
                        PRL19.SETRANGE("Blanket Order Line No.", "Blanket Order Line No.");
                        PRL19.SETRANGE("Document No.", '', "Document No.");
                        IF PRL19.FINDSET THEN
                            DenFac09 := 0;
                        REPEAT
                            IF (PRL19."Document No." <> "Document No.") THEN BEGIN
                                DenFac09 += PRL19."Unit Cost (LCY)" * PRL19.Quantity;
                            END;
                        UNTIL PRL19.NEXT = 0;

                        ExpQty09 := "Unit Cost (LCY)" * Quantity;

                        RecdQty09 := VenQty09 - DenFac09 - ExpQty09;
                    END;
                    //<<19Jul2018
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
                column(LineNo_PostedWhseReceiptLine; "Posted Whse. Receipt Line"."Line No.")
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
                column(BodyLine4; BodyLine4)
                {
                }
                column(BodyLine5; BodyLine5)
                {
                }
                column(BodyLine6; BodyLine6)
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

                    /*
                    Item.GET("Posted Whse. Receipt Line"."Item No.");
                    ItemName := Item.Description;
                    BaseUOM:=Item."Base Unit of Measure";
                    
                    "recWhse.RcptHeader".GET("Posted Whse. Receipt Line"."No.");
                    
                    Item.GET("Posted Whse. Receipt Line"."Item No.");
                    ItemName := Item.Description;
                    BaseUOM:=Item."Base Unit of Measure";
                    
                    "recWhse.RcptHeader".GET("Posted Whse. Receipt Line"."No.");
                    */
                    //MESSAGE('Qty=%1..desnsity=%2',"Posted Whse. Receipt Line".Quantity,"recWhse.RcptHeader"."Density Factor");

                end;
            }

            trigger OnAfterGetRecord()
            begin
                //>>02May2019
                CLEAR(GateEntryNo);
                CLEAR(GateEntryDate);
                PGEL.RESET;
                PGEL.SETCURRENTKEY("Entry Type", "Source Type", "Source No.");
                PGEL.SETRANGE("Entry Type", PGEL."Entry Type"::Inward);
                PGEL.SETRANGE("Source Type", PGEL."Source Type"::"Purchase Order");
                PGEL.SETRANGE("Source No.", "Purch. Rcpt. Header"."Order No.");
                IF PGEL.FINDFIRST THEN BEGIN
                    GateEntryNo := PGEL."Gate Entry No.";
                    PGEH.RESET;
                    IF PGEH.GET(PGEL."Entry Type", PGEL."Gate Entry No.") THEN
                        GateEntryDate := PGEH."Document Date";
                END;
                //<<02May2019

                //>>19Jul2018
                CLEAR(Text1);
                CLEAR(Text2);
                CLEAR(Text3);
                CLEAR(Text4);
                CLEAR(Text5);
                CLEAR(Text6);
                CLEAR(Text7);

                IF "Purch. Rcpt. Header"."Work Order" THEN BEGIN
                    GRNText := 'SERVICE RECEIVED NOTE';
                    Text1 := 'Total WO Value';
                    Text2 := 'Completed WO Value';
                    Text3 := 'Received WO Value';
                    Text4 := 'Balance WO Value';
                    Text5 := '';
                    Text6 := 'RECEIVED BY';
                    Text7 := 'VERIFIED BY';
                END ELSE BEGIN
                    GRNText := 'GOODS RECEIVED NOTE';
                    Text1 := 'Vendor Quantity';
                    Text2 := 'Density';
                    Text3 := 'Expected Quantity';
                    Text4 := 'Actual Recd.';
                    Text5 := 'Short/Excess';
                    Text6 := 'STORES';
                    Text7 := 'Q.C';
                END;
                //<<19Jul2018

                RespCentre.GET("Responsibility Center");
                Loc.GET("Location Code");
                IF RecState1.GET(Loc."State Code") THEN;

                IF RecState.GET("Purch. Rcpt. Header"."Location State Code") THEN;

                //IF vUser.GET(USERID) THEN; //GUID Error
                // TO Show Vehicle and LR Detail from Gate Entry------------------------------------START
                PostedGateEntryLine.RESET;
                PostedGateEntryLine.SETRANGE(PostedGateEntryLine."Source No.", "Purch. Rcpt. Header"."Order No.");
                IF PostedGateEntryLine.FINDFIRST THEN BEGIN
                    PostedGateEntryHeader.RESET;
                    PostedGateEntryHeader.SETRANGE(PostedGateEntryHeader."No.", PostedGateEntryLine."Gate Entry No.");
                    IF PostedGateEntryHeader.FINDFIRST THEN;
                END;
                //  TO Show Vehicle and LR Detail from Gate Entry--------------------------------------END
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

                //RSPLSUM 23Feb21>>
                CLEAR(txtPurchGRNRemark);
                RecPostedWhseRcptLine.RESET;
                RecPostedWhseRcptLine.SETRANGE("Posted Source No.", "No.");
                RecPostedWhseRcptLine.SETRANGE("Posting Date", "Posting Date");
                IF RecPostedWhseRcptLine.FINDFIRST THEN BEGIN
                    RecPostedWhseRcptHeader.RESET;
                    RecPostedWhseRcptHeader.SETRANGE("No.", RecPostedWhseRcptLine."No.");
                    IF RecPostedWhseRcptHeader.FINDFIRST THEN BEGIN
                        RecWhseCommentLine.RESET;
                        RecWhseCommentLine.SETCURRENTKEY("Table Name", Type, "No.");
                        RecWhseCommentLine.SETRANGE("Table Name", RecWhseCommentLine."Table Name"::"Posted Whse. Receipt");
                        RecWhseCommentLine.SETRANGE(Type, RecWhseCommentLine.Type::" ");
                        RecWhseCommentLine.SETRANGE("No.", RecPostedWhseRcptHeader."No.");
                        IF RecWhseCommentLine.FINDSET THEN
                            REPEAT
                                txtPurchGRNRemark += ', ' + RecWhseCommentLine.Comment;
                            UNTIL RecWhseCommentLine.NEXT = 0;
                    END;
                END;

                IF txtPurchGRNRemark <> '' THEN
                    txtPurchGRNRemark := COPYSTR(txtPurchGRNRemark, 2, STRLEN(txtPurchGRNRemark));
                //RSPLSUM 23Feb21<<
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
        "------------09Jan2018": Integer;
        BodyLine1: Boolean;
        BodyLine2: Boolean;
        BodyLine3: Boolean;
        BodyLine4: Boolean;
        BodyLine5: Boolean;
        BodyLine6: Boolean;
        PRLLine: Boolean;
        ItmName09: Text;
        VenQty09: Decimal;
        VenUOM09: Code[10];
        DenFac09: Decimal;
        ExpQty09: Decimal;
        RecdQty09: Decimal;
        RecdUOM09: Code[10];
        ShortQty09: Decimal;
        PWRL09: Record 7319;
        PWRH09: Record 7318;
        GRNText: Code[50];
        "-----------------19Jul2018": Integer;
        Text1: Text;
        Text2: Text;
        Text3: Text;
        Text4: Text;
        Text5: Text;
        Text6: Text;
        Text7: Text;
        PL19: Record 39;
        PRL19: Record 121;
        PGEL: Record "Posted Gate Entry Line";
        GateEntryNo: Code[20];
        PGEH: Record "Posted Gate Entry Header";
        GateEntryDate: Date;
        RecPostedWhseRcptLine: Record 7319;
        RecPostedWhseRcptHeader: Record 7318;
        RecWhseCommentLine: Record 5770;
        txtPurchGRNRemark: Text[512];
}


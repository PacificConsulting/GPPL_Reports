report 50024 "Bin-wise Raw Material - Plant"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BinwiseRawMaterialPlant.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Item; Item)
        {
            RequestFilterFields = "No.", "Item Category Code";
            dataitem("Warehouse Entry"; "Warehouse Entry")
            {
                DataItemLink = "Item No." = FIELD("No.");
                DataItemTableView = SORTING("Item No.", "Bin Code")
                                    ORDER(Ascending)
                                    WHERE("Qty. (Base)" = FILTER(<> 0));

                trigger OnAfterGetRecord()
                begin

                    Itemrec1.RESET;
                    Itemrec1.SETRANGE(Itemrec1."No.", "Warehouse Entry"."Item No.");
                    Itemrec1.SETRANGE(Itemrec1."Location Filter", "Warehouse Entry"."Location Code");
                    Itemrec1.SETRANGE(Itemrec1."Date Filter", "Warehouse Entry"."Registering Date");
                    IF Itemrec1.FINDFIRST THEN BEGIN
                        itmDesc := Itemrec1.Description;
                        ItmCat := Itemrec1."Item Category Code";
                    END;



                    Qty += ROUND("Warehouse Entry"."Qty. (Base)", 0.01);

                    vRemainingQty += "Qty. (Base)";

                    //>>29Apr2017 GroupData
                    TTT += 1;

                    WH29.SETCURRENTKEY("Item No.", "Bin Code");
                    WH29.SETRANGE("Item No.", "Item No.");
                    WH29.SETRANGE("Bin Code", "Bin Code");
                    IF WH29.FINDLAST THEN BEGIN
                        IIII := WH29.COUNT;

                    END;

                    IF IIII = TTT THEN BEGIN

                        IF vRemainingQty <> 0 THEN BEGIN
                            SrNo += 1;

                            EnterCell(k + 1, 1, FORMAT(SrNo), FALSE, FALSE, '', 0);//29Apr2017
                                                                                   //EnterCell(k+1,1,FORMAT(SrNo),FALSE,FALSE,'',ExcBuffer."Cell Type"::Text);
                            EnterCell(k + 1, 2, FORMAT("Item No."), FALSE, FALSE, '', ExcBuffer."Cell Type"::Text);
                            EnterCell(k + 1, 3, FORMAT(itmDesc), FALSE, FALSE, '', ExcBuffer."Cell Type"::Text);
                            EnterCell(k + 1, 4, FORMAT(ItmCat), FALSE, FALSE, '', ExcBuffer."Cell Type"::Text);
                            EnterCell(k + 1, 5, FORMAT("Bin Code"), FALSE, FALSE, '', ExcBuffer."Cell Type"::Text);
                            EnterCell(k + 1, 6, FORMAT("Unit of Measure Code"), FALSE, FALSE, '', ExcBuffer."Cell Type"::Text);
                            EnterCell(k + 1, 7, FORMAT(vRemainingQty), FALSE, FALSE, '0.000', 0);//29Apr2017
                                                                                                 //EnterCell(k+1,7,FORMAT(vRemainingQty),FALSE,FALSE,'#,####0.00',ExcBuffer."Cell Type"::Number);
                            EnterCell(k + 1, 8, FORMAT(itemDim), FALSE, FALSE, '', ExcBuffer."Cell Type"::Text); //RSPL
                                                                                                                 //EnterCell(k+1,9,"Warehouse Entry"."Registering Date",FALSE,FALSE,'',ExcBuffer."Cell Type"::Text); //Fahim 21-05-2022

                            k += 1;
                        END;

                        TTT := 0;
                        vRemainingQty := 0;
                    END;

                    //<<29Apr2017 GroupData
                end;

                trigger OnPostDataItem()
                begin
                    /*
                    IF vRemainingQty <>0 THEN BEGIN
                         SrNo+=1;

                         EnterCell(k+1,1,FORMAT(SrNo),FALSE,FALSE,'',0);//29Apr2017
                         //EnterCell(k+1,1,FORMAT(SrNo),FALSE,FALSE,'',ExcBuffer."Cell Type"::Text);
                         EnterCell(k+1,2,FORMAT("Item No."),FALSE,FALSE,'',ExcBuffer."Cell Type"::Text);
                         EnterCell(k+1,3,FORMAT(itmDesc),FALSE,FALSE,'',ExcBuffer."Cell Type"::Text);
                         EnterCell(k+1,4,FORMAT(ItmCat),FALSE,FALSE,'',ExcBuffer."Cell Type"::Text);
                         EnterCell(k+1,5,FORMAT("Bin Code"),FALSE,FALSE,'',ExcBuffer."Cell Type"::Text);
                         EnterCell(k+1,6,FORMAT("Unit of Measure Code"),FALSE,FALSE,'',ExcBuffer."Cell Type"::Text);
                         EnterCell(k+1,7,FORMAT(vRemainingQty),FALSE,FALSE,'0.000',0);//29Apr2017
                         //EnterCell(k+1,7,FORMAT(vRemainingQty),FALSE,FALSE,'#,####0.00',ExcBuffer."Cell Type"::Number);
                         EnterCell(k+1,8,FORMAT(itemDim),FALSE,FALSE,'',ExcBuffer."Cell Type"::Text); //RSPL
                         k+=1;
                   END;
                   */ //commented 29Apr2017

                end;

                trigger OnPreDataItem()
                begin

                    "Warehouse Entry".SETFILTER("Warehouse Entry"."Registering Date", '%1..%2', 0D, filterdate);
                    "Warehouse Entry".SETFILTER("Warehouse Entry"."Location Code", LocCode);
                    "Warehouse Entry".SETFILTER("Warehouse Entry"."Bin Code", bincode);
                    vRemainingQty := 0;

                    //>>29Apr2017
                    WH29.COPYFILTERS("Warehouse Entry");

                    TTT := 0;
                    //<<29Apr2017
                end;
            }

            trigger OnAfterGetRecord()
            begin

                Itemrec.GET(Item."No.");
                IF ((Itemrec."Item Category Code" = 'CAT02') OR (Itemrec."Item Category Code" = 'CAT03')) THEN
                    CurrReport.SKIP;

                //RSPL
                itemDim := '';
                recDefualtDim.RESET;
                recDefualtDim.SETRANGE(recDefualtDim."Table ID", 27);
                recDefualtDim.SETRANGE(recDefualtDim."No.", Item."No.");
                IF recDefualtDim.FINDFIRST THEN
                    itemDim := recDefualtDim."Dimension Value Code";
                //RSPL
            end;

            trigger OnPreDataItem()
            begin
                SrNo := 0;
            end;
        }
    }

    requestpage
    {
        SaveValues = true;

        layout
        {
            area(content)
            {
                field(LocCode; LocCode)
                {
                    Caption = 'Location Code';

                    trigger OnLookup(var Text: Text): Boolean
                    begin

                        Locrec.RESET;
                        //Locrec.SETRANGE(Locrec.Code,"Item Ledger Entry"."Location Code");
                        //Locrec.GET("Item Ledger Entry"."Location Code");
                        IF PAGE.RUNMODAL(0, Locrec) = ACTION::LookupOK THEN
                            LocCode := Locrec.Code;
                    end;
                }
                field(bincode; bincode)
                {
                    Caption = 'Bin Code';

                    trigger OnLookup(var Text: Text): Boolean
                    begin


                        BinCont.RESET;
                        BinCont.SETRANGE(BinCont."Location Code", LocCode);
                        IF BinCont.FINDFIRST THEN BEGIN
                            IF PAGE.RUNMODAL(0, BinCont) = ACTION::LookupOK THEN
                                bincode := BinCont.Code;
                        END
                        ELSE
                            MESSAGE('Location %1 does not contain Bin', LocCode);
                    end;
                }
                field(filterdate; filterdate)
                {
                    Caption = 'Filter Date';
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
        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN
            EnterCell(k + 1, 1, 'Grand Total', TRUE, FALSE, '@', ExcBuffer."Cell Type"::Text);
            EnterCell(k + 1, 7, FORMAT(Qty), TRUE, FALSE, '#,####0.00', ExcBuffer."Cell Type"::Number);
            k += 1;
        END;



        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN
            // ExcBuffer.CreateBook;
            // ExcBuffer.CreateSheet('Raw Materials Details','',COMPANYNAME,USERID);
            //ExcBuffer.CreateBookAndOpenExcel('', 'Raw Materials Details', '', '', '');
            //ExcBuffer.GiveUserControl;

            ExcBuffer.CreateNewBook('Raw Materials Details');
            ExcBuffer.WriteSheet('Raw Materials Details', CompanyName, UserId);
            ExcBuffer.CloseBook();
            ExcBuffer.SetFriendlyFilename(StrSubstNo('Raw Materials Details', CurrentDateTime, UserId));
            ExcBuffer.OpenExcel();

        END;
    end;

    trigger OnPreReport()
    begin

        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN
            ExcBuffer.DELETEALL;
            CLEAR(ExcBuffer);
            EnterCell(1, 1, FORMAT(COMPANYNAME + '' + Item."Location Filter"), TRUE, FALSE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(2, 1, 'Raw Materials Stock', TRUE, FALSE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(3, 1, 'Date :' + FORMAT(filterdate), TRUE, FALSE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(1, 8, FORMAT(FORMAT(TODAY, 0, 4)), TRUE, FALSE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(2, 8, FORMAT(USERID), TRUE, FALSE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(6, 1, 'Sr No.', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(6, 2, 'RM Code', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(6, 3, 'Item Description', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(6, 4, 'Item Category Code', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(6, 5, 'Bin Code', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(6, 6, 'Unit of Measure Code', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(6, 7, 'Remaining Qty in LTRS/KGS', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text);
            EnterCell(6, 8, 'Default Dimension', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text); //RSPL
            EnterCell(6, 9, 'Manufacturing Date', TRUE, TRUE, '', ExcBuffer."Cell Type"::Text); // Fahim 21-05-2022
            k := +6;
        END;
    end;

    var
        SrNo: Integer;
        Itemrec: Record 27;
        Itemrec1: Record 27;
        itmDesc: Text[50];
        ItmCat: Code[10];
        Qty: Decimal;
        ExportToExcel: Boolean;
        ExcBuffer: Record 370 temporary;
        k: Integer;
        LocCode: Code[10];
        filterdate: Date;
        Locrec: Record 14;
        bincode: Code[20];
        BinCont: Record 7354;
        itemDim: Code[20];
        recDefualtDim: Record 352;
        "--": Integer;
        BinCode2: Code[60];
        ItemNo: Code[30];
        vRemainingQty: Decimal;
        "----29Apr2017": Integer;
        IIII: Integer;
        TTT: Integer;
        WH29: Record 7312;

    // //[Scope('Internal')]
    procedure EnterCell(Rowno: Integer; columnno: Integer; Cellvalue: Text[250]; Bold: Boolean; Underline: Boolean; NumberFormat: Code[20]; vCellType: Option Number,Text,Date,Time)
    begin
        ExportToExcel := TRUE;
        IF ExportToExcel THEN BEGIN
            ExcBuffer.INIT;
            ExcBuffer.VALIDATE("Row No.", Rowno);
            ExcBuffer.VALIDATE("Column No.", columnno);
            ExcBuffer."Cell Value as Text" := Cellvalue;
            ExcBuffer.Formula := '';
            ExcBuffer.Bold := Bold;
            ExcBuffer.Underline := Underline;
            ExcBuffer.NumberFormat := NumberFormat;
            ExcBuffer."Cell Type" := vCellType;
            ExcBuffer.INSERT;
        END;

    end;
}


report 50010 "PO Confirmation for Customer"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/POConfirmationforCustomer.rdl';
    ApplicationArea = all;
    UsageCategory = ReportsAndAnalysis;

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            RequestFilterFields = "No.";
            column(CustName; Cust."Full Name")
            {
            }
            column(CustAdd1; Cust.Address + ', ' + Cust."Address 2" + ', ' + Cust.City + '-' + Cust."Post Code" + ', ' + recstate.Description + ', ' + reccountry.Name)
            {
            }
            column(CompHome; 'Website :' + RECCOMPINFO."Home Page" + '   C.I.N No. :' + 'RECCOMPINFO."Company Registration  No."' + '  PAN No. :' + RECCOMPINFO."P.A.N. No.")
            {
            }
            column(Date; Date)
            {
            }
            column(LocAdd1; Loc.Address)
            {
            }
            column(LocAdd2; Loc."Address 2")
            {
            }
            column(LocAdd3; Loc.City + ' - ' + Loc."Post Code")
            {
            }
            column(LocAdd4; recstate1.Description + ', ' + reccountry1.Name)
            {
            }
            column(DocNo; "Sales Invoice Header"."No.")
            {
            }
            column(DocDate; "Sales Invoice Header"."Posting Date")
            {
            }
            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "No.")
                                    WHERE(Type = FILTER(Item),
                                          Quantity = FILTER(<> 0));
                column(SrNo; SrNo)
                {
                }
                column(Particulars; Description)
                {
                }
                column(SalesUOM; "Unit of Measure")
                {
                }
                column(QtyPerSales; "Qty. per Unit of Measure")
                {
                }
                column(SalesQty; Quantity)
                {
                }
                column(QtyBase; "Quantity (Base)")
                {
                }
                column(NAH; NAH)
                {
                }
                column(LinDocNo; "Sales Invoice Line"."Document No.")
                {
                }
                column(LinNo; "Sales Invoice Line"."Line No.")
                {
                }
                column(ItmNo; "Sales Invoice Line"."No.")
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1
                    SrNo := SrNo + 1;
                    //<<1
                end;

                trigger OnPreDataItem()
                begin
                    NAH := COUNT; //21Mar2017
                end;
            }

            trigger OnAfterGetRecord()
            begin

                //>>1

                IF Cust.GET("Sales Invoice Header"."Sell-to Customer No.") THEN;

                recstate.RESET;
                recstate.SETRANGE(recstate.Code, Cust."State Code");
                IF recstate.FINDFIRST THEN
                    reccountry.RESET;
                reccountry.SETRANGE(reccountry.Code, Cust."Country/Region Code");
                IF reccountry.FINDFIRST THEN
                    Loc.GET("Sales Invoice Header"."Location Code");

                recstate1.RESET;
                recstate1.SETRANGE(recstate1.Code, Loc."State Code");
                IF recstate1.FINDFIRST THEN
                    reccountry1.RESET;
                reccountry1.SETRANGE(reccountry1.Code, Loc."Country/Region Code");
                IF reccountry1.FINDFIRST THEN
                    Date := CALCDATE('-3D', "Sales Invoice Header"."Posting Date");
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
        SrNo := 0;
        //<<1
    end;

    var
        Cust: Record 18;
        Loc: Record 14;
        ResCen: Record 5714;
        SrNo: Integer;
        recstate: Record State;
        reccountry: Record 9;
        recstate1: Record State;
        reccountry1: Record 9;
        recCust: Record 18;
        Date: Date;
        "-----------------PARAG": Integer;
        RECCOMPINFO: Record 79;
        "---------21Mar2017": Integer;
        NAH: Integer;
}


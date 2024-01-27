report 50169 "Statutory Form Remider"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'src/GPPL/Report Layout/StatutoryFormRemider.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Customer; Customer)
        {
            CalcFields = Balance;
            DataItemTableView = SORTING("No.")
                                WHERE(Blocked = FILTER(' '));
            RequestFilterFields = "No.", Type;
            column(CompanyName; CompnayInfo.Name)
            {
            }
            column(CompanyAdd; CompnayInfo.Address + ' ' + CompnayInfo."Address 2" + ' ' + CompnayInfo.City + ' ' + CompnayInfo."Post Code" + ' Tel: 66301911, 22873097 Fax: 91 22 22875751  E-mail: ipol@gulfpetrochem.com')
            {
            }
            column(CustNo; "No.")
            {
            }
            column(CustName; Name)
            {
            }
            column(CustAdd1; Customer.Address)
            {
            }
            column(CustAdd2; Customer."Address 2")
            {
            }
            column(CustAdd3; Customer.City + ' - ' + Customer."Post Code")
            {
            }
            column(CustAdd4; recState.Description + ' , ' + recCountry.Name)
            {
            }
            column(vAsonDate; vAsonDate)
            {
            }
            dataitem("Sales Invoice Header"; "Sales Invoice Header")
            {
                DataItemLink = "Sell-to Customer No." = FIELD("No.");
                DataItemTableView = SORTING("Posting Date")
                                    ORDER(Ascending);
                /*  WHERE("Form Code" = FILTER(<> ''),
                       "Form No." = FILTER('')); */
                column(finyear;
                finyear)
                {
                }
                column(PostingDate; "Sales Invoice Header"."Posting Date")
                {
                }
                column(DocNo; "Sales Invoice Header"."No.")
                {
                }
                column(dec_assesble; dec_assesble)
                {
                }
                column(dec_taxAmt; dec_taxAmt)
                {
                }
                column(InvAmt; 0)//"Sales Invoice Header"."Amount to Customer")
                {
                }
                column(FormCode; '"Sales Invoice Header"."Form Code"')
                {
                }

                trigger OnAfterGetRecord()
                begin

                    //>>1

                    dec_assesble := 0;
                    dec_taxAmt := 0;

                    finyear := FORMAT(DATE2DMY(("Sales Invoice Header"."Posting Date"), 3));

                    recSalesInvLine.RESET;
                    recSalesInvLine.SETRANGE(recSalesInvLine."Document No.", "Sales Invoice Header"."No.");
                    IF recSalesInvLine.FINDSET THEN BEGIN
                        REPEAT
                            dec_assesble += 0;// recSalesInvLine."Tax Base Amount";
                            dec_taxAmt += 0;// recSalesInvLine."Tax Amount";
                        UNTIL recSalesInvLine.NEXT = 0;
                    END;
                    //<<1
                end;

                trigger OnPreDataItem()
                begin

                    //>>1
                    SETRANGE("Posting Date", 0D, vAsonDate);
                    IF vFormCode <> '' THEN;
                    //SETRANGE("Form Code", vFormCode);
                    //<<1
                end;
            }

            trigger OnPreDataItem()
            begin

                //>>1
                CompnayInfo.GET;
                //<<1

                //>>2

                /*
                //Customer, Header (2) - OnPreSection()
                pgno+=1;
                IF vCustomerNo<> Customer."No." THEN
                    pgno:=1;
                //<<2

                //>>3
                //Customer, Header (3) - OnPreSection()
                IF CurrReport.PAGENO>1 THEN
                   CurrReport.SHOWOUTPUT(FALSE);
                *///Commented 11May2017
                  //<<3

            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("As on Date"; vAsonDate)
                {
                    ApplicationArea = all;
                }
                field("Form Code"; vFormCode)
                {
                    ApplicationArea = all;
                    // TableRelation = "Form Codes".Code;
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

    var
        CompnayInfo: Record 79;
        pgno: Integer;
        LastFieldNo: Integer;
        FooterPrinted: Boolean;
        recSalesInvLine: Record 113;
        vBillalreadydue: Decimal;
        vBillsbecomingdue: Decimal;
        vBillbecomenextcycle: Decimal;
        vDay: Integer;
        vMonth: Integer;
        vYear: Integer;
        vDate: Integer;
        vStratDate: Date;
        vEndDate: Date;
        vLastEnd: Date;
        vMonthDes: Text[50];
        vStartMonth: Option Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec;
        vStartYear: Integer;
        vMontSrat: Date;
        vMontSrat1: Date;
        vMontSrat2: Integer;
        vMontSrat3: Integer;
        vMontEnd: Date;
        vEndDate1: Integer;
        vEndDate2: Integer;
        vEndDate3: Date;
        recCountry: Record 9;
        recState: Record State;
        vCustomerNo: Code[20];
        vMonthDescription: Code[20];
        vAsonDate: Date;
        vAsDate: Integer;
        vAsMonth: Integer;
        vAsMonthDescription: Text[30];
        vAsYear: Integer;
        vDays: Integer;
        vEndDateofason: Date;
        dec_assesble: Decimal;
        dec_taxAmt: Decimal;
        finyear: Code[10];
        vFormCode: Code[10];

    //  [Scope('Internal')]
    procedure Month()
    begin
    end;
}


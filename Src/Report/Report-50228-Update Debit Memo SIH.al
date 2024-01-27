report 50228 "Update Debit Memo SIH"
{
    DefaultLayout = RDLC;

    RDLCLayout = 'Src/Report Layout/UpdateDebitMemoSIH.rdl';
    Permissions = TableData 112 = rm;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number)
                                WHERE(Number = CONST(1));

            trigger OnAfterGetRecord()
            begin
                RecSIH.RESET;
                IF RecSIH.FINDSET THEN
                    REPEAT
                        NoSeriesRelationship.RESET;
                        IF NoSeriesRelationship.GET('SALEINV', RecSIH."No. Series") THEN BEGIN
                            IF NoSeriesRelationship."Document Type" = NoSeriesRelationship."Document Type"::"Sale Debit Memo" THEN BEGIN
                                RecSIH."Debit Memo" := TRUE;
                                RecSIH.MODIFY;
                            END;
                        END;
                    UNTIL RecSIH.NEXT = 0;
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

    trigger OnPostReport()
    begin
        MESSAGE('Done');
    end;

    var
        RecSIH: Record 112;
        NoSeries: Record 308;
        NoSeriesRelationship: Record 310;
}


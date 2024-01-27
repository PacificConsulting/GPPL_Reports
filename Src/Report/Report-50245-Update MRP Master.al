report 50245 "Update MRP Master"
{
    DefaultLayout = RDLC;

    RDLCLayout = 'Src/Report Layout/UpdateMRPMaster.rdl';
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

                recMRPMaster.RESET;
                recMRPMaster.SETFILTER("Item No.", '<>%1', '');
                IF recMRPMaster.FINDSET THEN
                    REPEAT
                        IF recMRPMaster."National Discount" <> 0 THEN BEGIN
                            recMRPMaster."National Discount" := 0;
                            recMRPMaster.MODIFY;
                        END;
                    UNTIL recMRPMaster.NEXT = 0;
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
        recMRPMaster: Record "50013";
}


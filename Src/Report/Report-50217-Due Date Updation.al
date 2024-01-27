report 50217 "Due Date Updation"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/DueDateUpdation.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Permissions = TableData 21 = rm,
                  TableData 112 = rm,
                  TableData 114 = rm,
                  TableData 379 = rm;

    dataset
    {
        dataitem(DataItem1000000000; 50042)
        {

            trigger OnAfterGetRecord()
            begin

                //SIH.RESET;
                //IF SIH.GET("Doc No") THEN
                //BEGIN
                //SIH."Due Date" := "Temp SIH"."New Due Date";
                //SIH.MODIFY;
                I += 1;

                /*
                  CLE.RESET;
                  CLE.SETCURRENTKEY("Document No.");
                  CLE.SETRANGE("Document No.","Doc No");
                  IF CLE.FINDFIRST THEN
                  BEGIN
                    CLE."Due Date" := "Temp SIH"."New Due Date";
                    CLE.MODIFY;
                  END;
                
                  DCLE.RESET;
                  DCLE.SETCURRENTKEY("Document No.");
                  DCLE.SETRANGE("Document No.","Doc No");
                  DCLE.SETRANGE("Entry Type",DCLE."Entry Type"::"Initial Entry");
                  IF DCLE.FINDFIRST THEN
                  BEGIN
                    DCLE."Initial Entry Due Date" := "Temp SIH"."New Due Date";
                    DCLE.MODIFY;
                  END;
                
                */

                Cus.RESET;
                Cus.SETRANGE("No.", "Location Code");
                IF Cus.FINDFIRST THEN BEGIN
                    // Cus.Blocked :="Dispatch From Location"."Address 2";
                    Cus.MODIFY;
                END;
                //END;

            end;

            trigger OnPostDataItem()
            begin
                MESSAGE('%1 Document has been Upadted', I);
            end;

            trigger OnPreDataItem()
            begin
                I := 0;
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

    var
        SIH: Record 114;
        CLE: Record 21;
        DCLE: Record 379;
        I: Integer;
        Cus: Record 18;
}


report 50204 "Test Email2"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/TestEmail2.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    Caption = 'Test Email2';
    Permissions = TableData 112 = rm;

    dataset
    {
        dataitem(Integer; Integer)
        {
            MaxIteration = 1;

            trigger OnAfterGetRecord()
            begin
                /*
                SMSetup.GET;
                SM.CreateMessage('QWERT','kaustubh.parab@gpglobal.com','kaustubh.parab@gpglobal.com','KKKKKKKKK'+' '+FORMAT(Number),'qwerhjgu',TRUE);
                //SM.AddRecipients('nahnabeel@yahoo.com');
                SM.AddRecipients('nabeel.hajiameen@robo-soft.net');
                //SM.AddRecipients('khambalravi@gmail.com');
                SM.AddRecipients('abhishek.ishwalkar@gmail.com');
                SM.AddRecipients('nahnabeel@yahoo.com');
                SM.Send;
                //MESSAGE(USERID);
                
                PrOrd.RESET;
                PrOrd.SETRANGE(Status,PrOrd.Status::Released);
                PrOrd.SETRANGE("No.",'V-F/BLND/1819/7314');
                IF PrOrd.FINDFIRST THEN
                BEGIN
                  PrOrd."Refresh ProdOrder Qty" :=  FALSE;
                  PrOrd.MODIFY;
                END;
                */

                /*
                Res.RESET;
                Res.SETRANGE("Entry No.",1016670);
                IF Res.FINDFIRST THEN
                BEGIN
                  //Res."Created By" := 'GPUAE\ANTHONY.KURIAPPAN';
                  //Res."Created By" := 'GPUAE\KAPIL.PAWAR';
                  //Res."Qty. per Unit of Measure" := 210;
                  //Res."Item Tracking" := Res."Item Tracking"::"Lot No.";
                  //Res.MODIFY;
                  Res.RENAME(1016670,TRUE);
                END;
                */
                /*
                PH.RESET;
                PH.SETRANGE("Document Type",PH."Document Type"::"Blanket Order");
                PH.SETRANGE(Status,PH.Status::Released);
                PH.SETRANGE("Approved by Finance",FALSE);
                IF PH.FINDFIRST THEN
                  PH.MODIFYALL("Approved by Finance",TRUE);
                
                */

                /*
                Itm.RESET;
                Itm.SETRANGE(Blocked,TRUE);
                Itm.SETFILTER("Production BOM No.",'<>%1','');
                IF Itm.FINDSET THEN
                REPEAT
                  PBOMH.RESET;
                  IF PBOMH.GET(Itm."Production BOM No.") THEN
                  BEGIN
                    PBOMH.Blocked := Itm.Blocked;
                    PBOMH.MODIFY;
                  END;
                UNTIL Itm.NEXT = 0;
                */
                /*
                SH.RESET;
                IF SH.GET(SH."Document Type"::Order,'CSO/02/1920/1256') THEN
                BEGIN
                  SH.Status := SH.Status::Released;
                  SH."Campaign No." := 'APPROVED';
                  SH."Credit Limit Approval" := SH."Credit Limit Approval"::Approved;
                  SH.MODIFY;
                END;
                */

                SIH.RESET;
                IF SIH.GET('I/05/I/1920/0047') THEN BEGIN
                    SIH."Full Name" := 'Hiten Fasteners Private Limited';
                    SIH.MODIFY;
                END;

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
        // SM: Codeunit 400;
        //  SMSetup: Record 409;
        PWRH: Record 7318;
        PrOrd: Record 5405;
        SH: Record 36;
        Ven: Record 23;
        Res: Record 337;
        PH: Record 38;
        Itm: Record 27;
        PBOMH: Record 99000771;
        SIH: Record 112;
}


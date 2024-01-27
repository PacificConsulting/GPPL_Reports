report 60001 "Generate Leaves"
{
    // 10-Jan-06
    // 17-Dec-05

    ProcessingOnly = true;
    UseRequestPage = false;


    dataset
    {
        dataitem(Employee_; Employee_)
        {
            DataItemTableView = SORTING("No.")
                                ORDER(Ascending);

            trigger OnAfterGetRecord()
            begin
                Employee_.TESTFIELD("Pay Cadre");
                IF NOT Employee_.Blocked THEN BEGIN
                    IF Employee_."Employment Date" <= Startdate THEN BEGIN
                        IF Employee_.Probation THEN BEGIN
                            Window.UPDATE(1, Employee_."No.");
                            InsertLeaveEntitlement;
                            Employee_."Leaves Not Generated" := FALSE;
                            Employee_.MODIFY
                        END ELSE BEGIN
                            Window.UPDATE(1, Employee_."No.");
                            RegularEmployeeLeaves;
                            Employee_."Leaves Not Generated" := FALSE;
                            Employee_.MODIFY;
                        END;
                    END;
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

    trigger OnPostReport()
    begin
        IF Check THEN
            MESSAGE(Text000)
        ELSE
            MESSAGE(Text001);

        Window.CLOSE;
    end;

    trigger OnPreReport()
    begin
        Payyear.SETRANGE("Year Type", 'LEAVE YEAR');
        Payyear.SETRANGE(Closed, FALSE);
        IF Payyear.FIND('-') THEN BEGIN
            Startdate := Payyear."Year Start Date";
            Enddate := Payyear."Year End Date";
        END;

        Window.OPEN(Text002);
    end;

    var
        HRsetup: Record 60016;
        Payyear: Record 60020;
        LeaveMaster: Record 60030;
        LeaveEntitlement: Record 60031;
        LeaveEntitlement2: Record 60031;
        LeaveEntitlement3: Record 60031;
        Num: Integer;
        Text000: Label 'Leaves created for the employees';
        Check: Boolean;
        Text001: Label 'Leaves are already created';
        Enddate: Date;
        Startdate: Date;
        Window: Dialog;
        Text002: Label 'Employee......#1######################\  ';

    //  //[Scope('Internal')]
    procedure InsertLeaveEntitlement()
    var
        CreditingLeaves: Decimal;
        CreditingMonth: Integer;
        CreditingYear: Integer;
        CreditingDate: Date;
    begin
        LeaveMaster.SETRANGE("Applicable During Probation", TRUE);
        IF LeaveMaster.FIND('-') THEN BEGIN
            REPEAT
                IF LeaveEntitlement2.FIND('+') THEN
                    Num := LeaveEntitlement2."Entry No."
                ELSE
                    Num := 0;

                LeaveEntitlement.INIT;
                LeaveEntitlement."Entry No." := Num + 1;
                LeaveEntitlement."Employee No." := Employee_."No.";
                LeaveEntitlement."Employee Name" := Employee_."First Name";
                LeaveEntitlement.Probation := Employee_.Probation;
                LeaveEntitlement."Leave Code" := LeaveMaster."Leave Code";
                CreditingLeaves := ROUND((LeaveMaster."No.of Leaves During Probation") / 12, 0.5, '>');
                CreditingDate := CALCDATE(LeaveMaster."Crediting Interval", Startdate);
                CreditingMonth := DATE2DMY(CreditingDate, 2);
                CreditingYear := DATE2DMY(CreditingDate, 3);

                LeaveEntitlement."Leave Year Closing Period" := Enddate;

                IF HRsetup.FIND('-') THEN BEGIN
                    LeaveEntitlement.Month := HRsetup."Salary Processing month";
                    LeaveEntitlement.Year := HRsetup."Salary Processing Year";
                END;

                IF LeaveMaster."Crediting Type" = LeaveMaster."Crediting Type"::"After the Period" THEN BEGIN
                    IF (CreditingMonth = LeaveEntitlement.Month) AND (CreditingYear = LeaveEntitlement.Year) THEN
                        LeaveEntitlement."No.of Leaves" := CreditingLeaves;
                END ELSE
                    LeaveEntitlement."No.of Leaves" := CreditingLeaves;

                LeaveEntitlement."Total Leaves" := LeaveEntitlement."No.of Leaves";
                LeaveEntitlement3.SETRANGE("Employee No.", LeaveEntitlement."Employee No.");
                LeaveEntitlement3.SETRANGE("Leave Code", LeaveEntitlement."Leave Code");
                LeaveEntitlement3.SETRANGE(Year, LeaveEntitlement.Year);
                LeaveEntitlement3.SETRANGE(Month, LeaveEntitlement.Month);
                IF NOT LeaveEntitlement3.FIND('-') THEN BEGIN
                    LeaveEntitlement.INSERT;
                    Check := TRUE;
                END ELSE
                    Check := FALSE;
            UNTIL LeaveMaster.NEXT = 0;
        END;
    end;

    //  //[Scope('Internal')]
    procedure RegularEmployeeLeaves()
    var
        CreditingLeaves: Decimal;
        CreditingMonth: Integer;
        CreditingYear: Integer;
        CreditingDate: Date;
    begin
        LeaveMaster.RESET;
        IF LeaveMaster.FIND('-') THEN BEGIN
            REPEAT
                IF LeaveEntitlement2.FIND('+') THEN
                    Num := LeaveEntitlement2."Entry No."
                ELSE
                    Num := 0;

                LeaveEntitlement.INIT;
                LeaveEntitlement."Entry No." := Num + 1;
                LeaveEntitlement."Employee No." := Employee_."No.";
                LeaveEntitlement."Employee Name" := Employee_."First Name";
                LeaveEntitlement.Probation := Employee_.Probation;
                LeaveEntitlement."Leave Code" := LeaveMaster."Leave Code";
                LeaveEntitlement."No.of Leaves" := LeaveMaster."No. of Leaves in Year";
                LeaveEntitlement."Leave Year Closing Period" := Enddate;

                IF HRsetup.FIND('-') THEN BEGIN
                    LeaveEntitlement.Month := HRsetup."Salary Processing month";
                    LeaveEntitlement.Year := HRsetup."Salary Processing Year";
                END;

                LeaveEntitlement."Total Leaves" := LeaveEntitlement."No.of Leaves";

                LeaveEntitlement3.SETRANGE("Employee No.", LeaveEntitlement."Employee No.");
                LeaveEntitlement3.SETRANGE("Leave Code", LeaveEntitlement."Leave Code");
                LeaveEntitlement3.SETRANGE(Year, LeaveEntitlement.Year);
                LeaveEntitlement3.SETRANGE(Month, LeaveEntitlement.Month);
                IF NOT LeaveEntitlement3.FIND('-') THEN BEGIN
                    LeaveEntitlement.INSERT;
                    Check := TRUE
                END ELSE
                    Check := FALSE;
            UNTIL LeaveMaster.NEXT = 0;
        END;
    end;
}


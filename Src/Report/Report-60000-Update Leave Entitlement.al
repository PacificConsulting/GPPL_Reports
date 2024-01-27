report 60000 "Update Leave Entitlement"
{
    // 18-Jan-06

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
                IF NOT Employee_.Blocked THEN BEGIN

                    Employee_.TESTFIELD("Employment Date");
                    IF Employee_.Probation THEN
                        LeaveUpdation
                    ELSE
                        RegularEmpLeaveUpdation;
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
    end;

    trigger OnPreReport()
    begin
        PayYear.SETRANGE("Year Type", 'LEAVE YEAR');
        PayYear.SETRANGE(Closed, FALSE);
        IF PayYear.FIND('-') THEN BEGIN
            Enddate := PayYear."Year End Date";
            PayMonth := DATE2DMY(Enddate, 2);
            Year := DATE2DMY(Enddate, 3);
        END;
    end;

    var
        HRSetup: Record 60016;
        LeaveMaster: Record 60030;
        LeaveEntitle: Record 60031;
        LeaveEntitle2: Record 60031;
        LeaveEntitle3: Record 60031;
        PayYear: Record 60020;
        Text000: Label 'Leaves updated for the employees';
        Text001: Label 'Leaves are already updated';
        LeavesCarry: Decimal;
        LeavesExp: Decimal;
        Num: Integer;
        PayMonth: Integer;
        Year: Integer;
        CountMonth: Integer;
        Check: Boolean;
        Enddate: Date;
        Window: Dialog;

    //  //[Scope('Internal')]
    procedure LeaveUpdation()
    var
        EmpMonth: Integer;
        EmpYear: Integer;
    begin
        EmpMonth := DATE2DMY(Employee_."Employment Date", 2);
        EmpYear := DATE2DMY(Employee_."Employment Date", 3);
        IF EmpYear = Year THEN BEGIN
            IF HRSetup.FIND('-') THEN BEGIN
                IF EmpMonth <= HRSetup."Salary Processing month" + 1 THEN
                    UpdateLeaveEntitle;
            END;
        END ELSE
            IF EmpYear < Year THEN
                UpdateLeaveEntitle;
    end;

    //  //[Scope('Internal')]
    procedure RegularEmpLeaveUpdation()
    var
        LeaveEntitlement: Record 60031;
        EmpMonth: Integer;
        EmpYear: Integer;
    begin
        EmpMonth := DATE2DMY(Employee_."Employment Date", 2);
        EmpYear := DATE2DMY(Employee_."Employment Date", 3);
        IF EmpYear = Year THEN BEGIN
            IF HRSetup.FIND('-') THEN BEGIN
                IF EmpMonth <= HRSetup."Salary Processing month" + 1 THEN BEGIN
                    CountMonth := (PayMonth - EmpMonth) + 1;
                    RegularEmpLeave;
                    Employee_."Leaves Not Generated" := FALSE;
                    Employee_.MODIFY;
                END;
            END;
        END ELSE
            IF EmpYear < Year THEN BEGIN
                CountMonth := 12;
                RegularEmpLeave;
            END;
    end;

    //  //[Scope('Internal')]
    procedure UpdateLeaveEntitle()
    begin
        LeavesCarry := 0;
        LeaveMaster.SETRANGE("Applicable During Probation", TRUE);
        IF LeaveMaster.FIND('-') THEN BEGIN
            REPEAT
                IF LeaveEntitle2.FIND('+') THEN
                    Num := LeaveEntitle2."Entry No."
                ELSE
                    Num := 0;

                LeaveEntitle.INIT;
                LeaveEntitle."Entry No." := Num + 1;
                LeaveEntitle."Employee No." := Employee_."No.";
                LeaveEntitle."Employee Name" := Employee_."First Name";
                LeaveEntitle.Probation := Employee_.Probation;
                LeaveEntitle."Leave Code" := LeaveMaster."Leave Code";
                LeaveEntitle."No.of Leaves" := (LeaveMaster."No.of Leaves During Probation") / 12;
                LeaveEntitle."Leave Year Closing Period" := Enddate;
                IF HRSetup.FIND('-') THEN BEGIN
                    IF HRSetup."Salary Processing month" = 12 THEN BEGIN
                        LeaveEntitle.Month := 1;
                        LeaveEntitle.Year := HRSetup."Salary Processing Year" + 1;
                        LeaveEntitle."Leave Year Closing Period" := 0D;
                    END ELSE BEGIN
                        LeaveEntitle.Month := HRSetup."Salary Processing month" + 1;
                        LeaveEntitle.Year := HRSetup."Salary Processing Year";
                    END;
                END;

                LeavesCarried(LeaveEntitle, LeaveMaster);
                LeaveEntitle."Leaves Carried" := LeavesCarry;
                IF HRSetup.FIND('-') THEN BEGIN
                    IF HRSetup."Salary Processing month" = 12 THEN
                        LeaveEntitle."Leaves Expired" := LeavesExp;
                END;
                LeaveEntitle."Total Leaves" := LeaveEntitle."No.of Leaves" + LeaveEntitle."Leaves Carried";
                LeaveEntitle3.SETRANGE("Employee No.", LeaveEntitle."Employee No.");
                LeaveEntitle3.SETRANGE("Leave Code", LeaveEntitle."Leave Code");
                LeaveEntitle3.SETRANGE(Year, LeaveEntitle.Year);
                LeaveEntitle3.SETRANGE(Month, LeaveEntitle.Month);
                IF NOT LeaveEntitle3.FIND('-') THEN BEGIN
                    LeaveEntitle.INSERT;
                    Check := TRUE
                END ELSE
                    Check := FALSE;
            UNTIL LeaveMaster.NEXT = 0;
        END;
    end;

    //  //[Scope('Internal')]
    procedure RegularEmpLeave()
    begin
        LeaveMaster.RESET;
        IF LeaveMaster.FIND('-') THEN BEGIN
            REPEAT
                IF LeaveEntitle2.FIND('+') THEN
                    Num := LeaveEntitle2."Entry No."
                ELSE
                    Num := 0;

                LeaveEntitle.INIT;
                LeaveEntitle."Entry No." := Num + 1;
                LeaveEntitle."Employee No." := Employee_."No.";
                LeaveEntitle."Employee Name" := Employee_."First Name";
                LeaveEntitle.Probation := Employee_.Probation;
                LeaveEntitle."Leave Code" := LeaveMaster."Leave Code";
                LeaveEntitle."No.of Leaves" := (LeaveMaster."No. of Leaves in Year" / 12) * CountMonth;
                LeaveEntitle."Leave Year Closing Period" := Enddate;
                IF HRSetup.FIND('-') THEN BEGIN
                    IF HRSetup."Salary Processing month" = 12 THEN BEGIN
                        LeaveEntitle.Month := 1;
                        LeaveEntitle.Year := HRSetup."Salary Processing Year" + 1;
                        LeaveEntitle."Leave Year Closing Period" := 0D;
                    END ELSE BEGIN
                        LeaveEntitle.Month := HRSetup."Salary Processing month" + 1;
                        LeaveEntitle.Year := HRSetup."Salary Processing Year";
                    END;
                END;
                LeavesCarried(LeaveEntitle, LeaveMaster);
                IF Employee_."Leaves Not Generated" = TRUE THEN
                    LeaveEntitle."Total Leaves" := LeaveEntitle."No.of Leaves"
                ELSE
                    LeaveEntitle."Total Leaves" := LeavesCarry;

                IF HRSetup.FIND('-') THEN BEGIN
                    IF HRSetup."Salary Processing month" = 12 THEN BEGIN
                        LeaveEntitle."Leaves Carried" := LeavesCarry;
                        LeaveEntitle."Leaves Expired" := LeavesExp;
                        LeaveEntitle."Total Leaves" := LeaveEntitle."No.of Leaves" + LeaveEntitle."Leaves Carried";
                    END;
                END;
                LeaveEntitle3.SETRANGE("Employee No.", LeaveEntitle."Employee No.");
                LeaveEntitle3.SETRANGE("Leave Code", LeaveEntitle."Leave Code");
                LeaveEntitle3.SETRANGE(Year, LeaveEntitle.Year);
                LeaveEntitle3.SETRANGE(Month, LeaveEntitle.Month);
                IF NOT LeaveEntitle3.FIND('-') THEN BEGIN
                    LeaveEntitle.INSERT;
                    Check := TRUE
                END ELSE
                    Check := FALSE;
            UNTIL LeaveMaster.NEXT = 0;
        END;
    end;

    //  //[Scope('Internal')]
    procedure LeavesCarried(LeaveEntitlement: Record 60031; LeaveMaster: Record 60030)
    var
        LeaveEntitlement2: Record 60031;
        Month: Integer;
        CurrMonth: Integer;
        CurrYear: Integer;
        Year: Integer;
    begin
        CurrMonth := LeaveEntitlement.Month;
        CurrYear := LeaveEntitlement.Year;

        IF CurrMonth = 1 THEN
            Year := CurrYear - 1
        ELSE
            Year := CurrYear;

        IF CurrMonth = 1 THEN
            Month := 12
        ELSE
            Month := CurrMonth - 1;

        LeaveEntitlement2.SETRANGE("Employee No.", LeaveEntitlement."Employee No.");
        LeaveEntitlement2.SETRANGE("Leave Code", LeaveEntitlement."Leave Code");
        LeaveEntitlement2.SETRANGE(Year, Year);
        LeaveEntitlement2.SETRANGE(Month, Month);
        IF LeaveEntitlement2.FIND('-') THEN BEGIN
            IF LeaveEntitlement.Probation THEN BEGIN
                LeavesCarry := LeaveEntitlement2."Leave Bal. at the Month End";
                IF HRSetup.FIND('-') THEN BEGIN
                    IF HRSetup."Salary Processing month" = 12 THEN BEGIN
                        IF LeavesCarry > LeaveMaster."Max.Leaves to Carry Forward" THEN
                            LeavesCarry := LeaveMaster."Max.Leaves to Carry Forward";
                        LeavesExp := LeaveEntitlement2."Leave Bal. at the Month End" - LeavesCarry;
                    END;
                END;
            END ELSE BEGIN
                LeavesCarry := LeaveEntitlement2."Leave Bal. at the Month End";
                IF HRSetup.FIND('-') THEN BEGIN
                    IF HRSetup."Salary Processing month" = 12 THEN BEGIN
                        IF LeaveEntitlement2."Leave Bal. at the Month End" > LeaveMaster."Max.Leaves to Carry Forward" THEN
                            LeavesCarry := LeaveMaster."Max.Leaves to Carry Forward";
                        LeavesExp := LeaveEntitlement2."Leave Bal. at the Month End" - LeavesCarry;
                    END;
                END;
            END;
        END;
    end;
}


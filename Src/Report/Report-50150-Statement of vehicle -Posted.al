report 50150 "Statement of vehicle -Posted"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/StatementofvehiclePosted.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Gate Entry Header - Loading"; "Posted Gate Entry Header")
        {
            DataItemTableView = SORTING("Location Code", "Posting Date", "No.")
                                ORDER(Ascending)
                                WHERE("Entry Type" = FILTER(Outward));
            RequestFilterFields = "Document Date", "Location Code";

            trigger OnAfterGetRecord()
            begin
                LSrNo := LSrNo + 1;

                IF PrintToExcel AND Loading THEN
                    "MakeExcelDataBody-Loading";
            end;

            trigger OnPreDataItem()
            begin

                //>>LocName Filter 23Mar2017

                LocCode := GETFILTER("Location Code");
                DateFil := GETFILTER("Document Date");

                recLoc.RESET;
                recLoc.SETRANGE(Code, LocCode);
                IF recLoc.FINDFIRST THEN
                    LocName := recLoc.Name;

                //<<LocName Filter 23Mar2017

                IF PrintToExcel AND Loading THEN
                    "MakeExcelDataHeader-Loading";
            end;
        }
        dataitem("Gate Entry Header - Unloading"; "Posted Gate Entry Header")
        {
            DataItemTableView = SORTING("Location Code", "Posting Date", "No.")
                                ORDER(Ascending)
                                WHERE("Entry Type" = FILTER(Inward));

            trigger OnAfterGetRecord()
            begin
                USrNo := USrNo + 1;

                IF PrintToExcel AND UnLoading THEN
                    "MakeExcelDataBody-UnLoading";
            end;

            trigger OnPreDataItem()
            begin

                "Gate Entry Header - Unloading".SETRANGE("Gate Entry Header - Unloading"."Document Date",
                                                         "Gate Entry Header - Loading"."Document Date");
                "Gate Entry Header - Unloading".SETRANGE("Gate Entry Header - Unloading"."Location Code",
                                                         "Gate Entry Header - Loading"."Location Code");

                IF PrintToExcel AND UnLoading THEN
                    "MakeExcelDataHeader-UnLoading";
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field(PrintToExcel; PrintToExcel)
                {
                    ApplicationArea = all;
                    Caption = 'Export to Excel';
                }
                field(Loading; Loading)
                {
                    ApplicationArea = all;
                    Caption = 'Loading';
                }
                field(UnLoading; UnLoading)
                {
                    ApplicationArea = all;
                    Caption = 'UnLoading';
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

        IF PrintToExcel THEN
            CreateExcelbook;
    end;

    trigger OnPreReport()
    begin

        LSrNo := 0;
        USrNo := 0;

        IF (Loading = FALSE) AND (UnLoading = FALSE) THEN
            ERROR('You have not selected any option');

        IF PrintToExcel THEN
            MakeExcelInfo;
    end;

    var
        PrintToExcel: Boolean;
        Loading: Boolean;
        UnLoading: Boolean;
        ExcelBuf: Record 370 temporary;
        LSrNo: Integer;
        USrNo: Integer;
        Text0000: Label 'Data';
        Text0001: Label 'Statement of Vehicle Reported';
        Text0002: Label 'Company Name';
        Text0003: Label 'Report No.';
        Text0004: Label 'Report Name';
        Text0005: Label 'USER Id';
        Text0006: Label 'Date';
        "---23Mar2017": Integer;
        LocName: Text[100];
        recLoc: Record 14;
        LocCode: Code[50];
        DateFil: Text[50];

    // //[Scope('Internal')]
    procedure MakeExcelInfo()
    begin
        //ExcelBuf.SetUseInfoSheed;
        ExcelBuf.AddInfoColumn(FORMAT(Text0002), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(COMPANYNAME, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0004), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn('Statement of Vehicle Reported', FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0003), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(REPORT::"Statement of vehicle -Posted", FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0005), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(USERID, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddInfoColumn(FORMAT(Text0006), FALSE, TRUE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddInfoColumn(TODAY, FALSE, FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.ClearNewRow;
    end;

    //  //[Scope('Internal')]
    procedure CreateExcelbook()
    begin
        //ExcelBuf.CreateBook;
        //ExcelBuf.CreateSheet(Text0000,Text0001,COMPANYNAME,USERID);
        /* ExcelBuf.CreateBookAndOpenExcel('', Text0000, '', '', '');
        ExcelBuf.GiveUserControl; */

        ExcelBuf.CreateNewBook(Text0000);
        ExcelBuf.WriteSheet(Text0000, CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.SetFriendlyFilename(StrSubstNo(Text0000, CurrentDateTime, UserId));
        ExcelBuf.OpenExcel();
        ERROR('');
    end;

    // //[Scope('Internal')]
    procedure "MakeExcelDataHeader-Loading"()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('GP Petroleums Limited', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(FORMAT(TODAY, 0, 4), FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);


        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Statement Of Vehicle Reported For Loading', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(USERID, FALSE, '', TRUE, FALSE, FALSE, '@', ExcelBuf."Cell Type"::Text);

        //>>Report Filter 23Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Location : ' + LocName, FALSE, '', TRUE, FALSE, TRUE, '@', 1);
        ExcelBuf.NewRow;

        ExcelBuf.AddColumn('Date Filter : ' + DateFil, FALSE, '', TRUE, FALSE, TRUE, '@', 1);//1
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//2
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//3
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//4
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//5
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//6
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//7
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//8
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//9
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//10
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//11
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//12
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//13
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', 1);//14

        //<<Report Filter 23Mar2017
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Reported Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Reported Time', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('From Party', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vehicle No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Transporter', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('LR No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('LR Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vehicle Capacity', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vehicle For Location', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Drivers Name', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Drivers Mobile No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Drivers License No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    // //[Scope('Internal')]
    procedure "MakeExcelDataBody-Loading"()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn("Gate Entry Header - Loading"."Document Date",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date); //23Mar2017
        //ExcelBuf.AddColumn("Gate Entry Header - Loading"."Document Time",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Document Time", FALSE, '', FALSE, FALSE, FALSE, 'hh:mm:ss AM/PM', ExcelBuf."Cell Type"::Time);//23Mar2017
        ExcelBuf.AddColumn("Gate Entry Header - Loading".Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Vehicle No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Transporter Name", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Item Description", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."LR/RR No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn("Gate Entry Header - Loading"."LR/RR Date",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."LR/RR Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//23Mar2017
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Vehicle Capacity", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Vehicle For Location", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Driver's Name", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Driver's Mobile No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Loading"."Driver's License No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
    end;

    //  //[Scope('Internal')]
    procedure "MakeExcelDataHeader-UnLoading"()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Statement Of Vehicle Reported For UnLoading', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Reported Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Reported Time', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('From Party', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Vehicle No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Transporter', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Description', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('LR No.', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('LR Date', FALSE, '', TRUE, FALSE, TRUE, '@', ExcelBuf."Cell Type"::Text);
    end;

    //  //[Scope('Internal')]
    procedure "MakeExcelDataBody-UnLoading"()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn("Gate Entry Header - Unloading"."No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn("Gate Entry Header - Unloading"."Document Date",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Unloading"."Document Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//23Mar2017
        //ExcelBuf.AddColumn("Gate Entry Header - Unloading"."Document Time",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Unloading"."Document Time", FALSE, '', FALSE, FALSE, FALSE, 'hh:mm:ss AM/PM', ExcelBuf."Cell Type"::Time);//23Mar2017
        ExcelBuf.AddColumn("Gate Entry Header - Unloading".Description, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Unloading"."Vehicle No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Unloading".Transporter, FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Unloading"."Item Description", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Unloading"."LR/RR No.", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Text);
        //ExcelBuf.AddColumn("Gate Entry Header - Unloading"."LR/RR Date",FALSE,'',FALSE,FALSE,FALSE,'',ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn("Gate Entry Header - Unloading"."LR/RR Date", FALSE, '', FALSE, FALSE, FALSE, '', ExcelBuf."Cell Type"::Date);//23Mar2017
    end;
}


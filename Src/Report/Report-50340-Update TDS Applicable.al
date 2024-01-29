report 50340 "Update TDS Applicable"
{
    //// //robotdsDJ : Batch Report to Update TDS Applicable from 1st April to today for Purchase and Purchase Invoice Lines and Purchase Credit Memo Lines

    Permissions = TableData 123 = rm,
                  TableData 125 = rm;
    ProcessingOnly = true;

    dataset
    {
        dataitem("Purchase Header"; "Purchase Header")
        {
            DataItemTableView = SORTING("Document Type", "No.");
            dataitem("Purchase Line"; "Purchase Line")
            {
                DataItemLink = "Document Type" = FIELD("Document Type"),
                               "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document Type", "Document No.", "Line No.")
                                    WHERE("GST Group Type" = CONST(Goods),
                                          "HSN/SAC Code" = FILTER(<> ''));

                trigger OnAfterGetRecord()
                begin
                    //robotdsDJ_v2
                    HSNSACRec.RESET;
                    IF HSNSACRec.GET("GST Group Code", "HSN/SAC Code") THEN BEGIN
                        IF HSNSACRec.Type <> HSNSACRec.Type::HSN THEN
                            CurrReport.SKIP;
                        //robotdsDJ_v2
                        "TDS Applicable" := TRUE;
                        MODIFY;
                    END;
                end;
            }

            trigger OnPreDataItem()
            begin
                //SETRANGE("Posting Date",010421D,TODAY);
            end;
        }
        dataitem("Purch. Inv. Header"; "Purch. Inv. Header")
        {
            DataItemTableView = SORTING("No.");
            dataitem("Purch. Inv. Line"; "Purch. Inv. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE("GST Group Type" = CONST(Goods),
                                          "HSN/SAC Code" = FILTER(<> ''));

                trigger OnAfterGetRecord()
                begin
                    //robotdsDJ_v2
                    HSNSACRec.RESET;
                    IF HSNSACRec.GET("GST Group Code", "HSN/SAC Code") THEN BEGIN
                        IF HSNSACRec.Type <> HSNSACRec.Type::HSN THEN
                            CurrReport.SKIP;
                        //robotdsDJ_v2
                        "TDS Applicable" := TRUE;
                        MODIFY;
                    END;
                end;
            }

            trigger OnPreDataItem()
            begin
                SETRANGE("Posting Date", 20210401D/*010421D*/, TODAY);
            end;
        }
        dataitem("Purch. Cr. Memo Hdr."; "Purch. Cr. Memo Hdr.")
        {
            DataItemTableView = SORTING("No.");
            dataitem("Purch. Cr. Memo Line"; "Purch. Cr. Memo Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.")
                                    WHERE("GST Group Type" = CONST(Goods),
                                          "HSN/SAC Code" = FILTER(<> ''));

                trigger OnAfterGetRecord()
                begin
                    //robotdsDJ_v2
                    HSNSACRec.RESET;
                    IF HSNSACRec.GET("GST Group Code", "HSN/SAC Code") THEN BEGIN
                        IF HSNSACRec.Type <> HSNSACRec.Type::HSN THEN
                            CurrReport.SKIP;
                        //robotdsDJ_v2
                        "TDS Applicable" := TRUE;
                        MODIFY;
                    END;
                end;
            }

            trigger OnPreDataItem()
            begin
                SETRANGE("Posting Date", 20210401D/*010421D*/, TODAY);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                label("*Kindly check Date format in your system")
                {
                    Caption = '*Kindly check Date format in your system';
                    Style = Strong;
                    StyleExpr = TRUE;
                }
                label("*Data format in report is dd mm yy")
                {
                    Caption = '*Data format in report is dd mm yy';
                    Style = Strong;
                    StyleExpr = TRUE;
                }
                label("------------------------------------")
                {
                    Caption = '------------------------------------';
                }
                label("Click OK to Update TDS application on Purchase documents")
                {
                    Caption = 'Click OK to Update TDS application on Purchase documents';
                    Style = Strong;
                    StyleExpr = TRUE;
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
        MESSAGE('Process is completed');
    end;

    var
        RecGL: Record 15;
        HSNSACRec: Record "HSN/SAC";
}


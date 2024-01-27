report 50248 "Update Assesse Code"
{
    // //robotdsDJ_v3 : Batch Report to Update Assesse Code on Purchase Document and Posted Purchase Receipt document.

    Permissions = TableData 120 = rm,
                  TableData 121 = rm;
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
                DataItemTableView = SORTING("Document Type", "Document No.", "Line No.");

                trigger OnAfterGetRecord()
                begin
                    /*  IF NODNOCHeader.GET(NODNOCHeader.Type::Vendor, "Pay-to Vendor No.") THEN BEGIN
                         IF "Assessee Code" = '' THEN BEGIN
                             "Assessee Code" := NODNOCHeader."Assesse Code"; */
                    MODIFY;
                END;
                // END;
                //end;
            }

            trigger OnAfterGetRecord()
            begin
                /* IF NODNOCHeader.GET(NODNOCHeader.Type::Vendor, "Pay-to Vendor No.") THEN BEGIN
                    IF "Assessee Code" = '' THEN BEGIN
                        "Assessee Code" := NODNOCHeader."Assesse Code";
                        MODIFY;
                    END;
                END; */
            end;

            trigger OnPreDataItem()
            begin
                IF FromDate <> 0D THEN
                    SETRANGE("Posting Date", FromDate, TODAY);
            end;
        }
        dataitem("Purch. Rcpt. Header"; "Purch. Rcpt. Header")
        {
            DataItemTableView = SORTING("No.");
            dataitem("Purch. Rcpt. Line"; "Purch. Rcpt. Line")
            {
                DataItemLink = "Document No." = FIELD("No.");
                DataItemTableView = SORTING("Document No.", "Line No.");
            }

            trigger OnAfterGetRecord()
            begin
                /* IF NODNOCHeader.GET(NODNOCHeader.Type::Vendor, "Pay-to Vendor No.") THEN BEGIN
                    IF "Assessee Code" = '' THEN BEGIN
                        "Assessee Code" := NODNOCHeader."Assesse Code";
                        MODIFY;
                    END;
                END; */
            end;

            trigger OnPreDataItem()
            begin
                IF FromDate <> 0D THEN
                    SETRANGE("Posting Date", FromDate, TODAY);
            end;
        }
    }

    requestpage
    {

        layout
        {
            area(content)
            {
                field("From Date"; FromDate)
                {
                    ApplicationArea = all;
                    Caption = 'From Date';
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
        //NODNOCHeader: Record "NOD/NOC Header";
        FromDate: Date;
}


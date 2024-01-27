report 50197 "BatchTax Entry"
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/BatchTaxEntry.rdl';
    Permissions = TableData 113 = rm;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;
    dataset
    {
        dataitem("Sales Invoice Line"; "Sales Invoice Line")
        {
            DataItemTableView = SORTING("Document No.", "Line No.");
            RequestFilterFields = "Document No.", "Posting Date";

            trigger OnAfterGetRecord()
            begin
                IF "Sales Invoice Line".GETFILTER("Document No.") = '' THEN
                    ERROR('Document No. must not be blank');
                //**Modify Invoice Line
                // IF "Sales Invoice Line"."Tax Amount" <> 0 THEN BEGIN
                /*  PostStrLine.RESET;
                // PostStrLine.SETRANGE("Invoice No.", "Sales Invoice Line"."Document No.");
                 PostStrLine.SETRANGE(Type, PostStrLine.Type::Sale);
                 PostStrLine.SETRANGE("Item No.", "No.");
                 PostStrLine.SETRANGE("Line No.", "Line No.");
                 PostStrLine.SETRANGE("Tax/Charge Code", 'SALES TAX');
                 IF PostStrLine.FINDFIRST THEN BEGIN
                     "Sales Invoice Line"."Tax Base Amount" := PostStrLine."Base Amount";
                     "Sales Invoice Line".MODIFY;
                     vSL += 1;
                 END; */

                //**Modify detail Tax Entry
                /*   DetTaxEntry.RESET;
                  DetTaxEntry.SETFILTER("Document Type", '%1|%2', DetTaxEntry."Document Type"::Invoice, DetTaxEntry."Document Type"::"Credit Memo");
                  DetTaxEntry.SETRANGE("Document No.", "Document No.");
                  DetTaxEntry.SETRANGE("Document Line No.", "Sales Invoice Line"."Line No.");
                  //DetTaxEntry.SETRANGE("Posting Date",010417D,300417D);
                  IF DetTaxEntry.FINDSET THEN
                      REPEAT
                          DetTaxEntry."Tax Base Amount" := -1 * ("Sales Invoice Line"."Tax Base Amount");
                          DetTaxEntry."Tax Amount" := ROUND(-1 * ("Sales Invoice Line"."Tax Base Amount") * (DetTaxEntry."Tax %" / 100), 1);
                          DetTaxEntry."Input Credit/Output Tax Amount" := ROUND(-1 * ("Sales Invoice Line"."Tax Base Amount") * (DetTaxEntry."Tax %" / 100), 1);
                          DetTaxEntry."Remaining Tax Amount" := ROUND(-1 * ("Sales Invoice Line"."Tax Base Amount") * (DetTaxEntry."Tax %" / 100), 1);
                          DetTaxEntry.MODIFY;
                          vCount += 1;
                      UNTIL DetTaxEntry.NEXT = 0; */
                /// END ELSE BEGIN
                //**Modify detail Tax Entry
                /*  DetTaxEntry.RESET;
                 DetTaxEntry.SETFILTER("Document Type", '%1|%2', DetTaxEntry."Document Type"::Invoice, DetTaxEntry."Document Type"::"Credit Memo");
                 DetTaxEntry.SETRANGE("Document No.", "Document No.");
                 DetTaxEntry.SETRANGE("Document Line No.", "Sales Invoice Line"."Line No.");
                 //DetTaxEntry.SETRANGE("Posting Date",010417D,300417D);
                 IF DetTaxEntry.FINDSET THEN
                     REPEAT
                         DetTaxEntry."Tax Base Amount" := 0;
                         DetTaxEntry."Tax Amount" := 0;
                         DetTaxEntry."Input Credit/Output Tax Amount" := 0;
                         DetTaxEntry."Remaining Tax Amount" := 0;
                         DetTaxEntry."Tax %" := 0;
                         DetTaxEntry.MODIFY;
                         vCount += 1;
                     UNTIL DetTaxEntry.NEXT = 0; */

            END;
            //end;

            trigger OnPostDataItem()
            begin
                MESSAGE('No of records modified Successfully in Sales Line =%1 ,Detail Tax Entry =%2', vSL, vCount);
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
        //  DetTaxEntry: Record 16522;
        vCount: Integer;
        // PostStrLine: Record 13798;
        vSL: Integer;
}


report 50334 testAe
{
    DefaultLayout = RDLC;
    RDLCLayout = 'Src/Report Layout/testAe.rdl';
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = all;

    dataset
    {
        dataitem("Sales Approval Entry"; "Sales Approval Entry")
        {
            column(DocumentType_SalesApprovalEntry; "Sales Approval Entry"."Document Type")
            {
            }
            column(DocumentNo_SalesApprovalEntry; "Sales Approval Entry"."Document No.")
            {
            }
            column(UserID_SalesApprovalEntry; "Sales Approval Entry"."User ID")
            {
            }
            column(ApprovarID_SalesApprovalEntry; "Sales Approval Entry"."Approvar ID")
            {
            }
            column(MandatoryID_SalesApprovalEntry; "Sales Approval Entry"."Mandatory ID")
            {
            }
            column(Approved_SalesApprovalEntry; "Sales Approval Entry".Approved)
            {
            }
            column(ApprovalDate_SalesApprovalEntry; "Sales Approval Entry"."Approval Date")
            {
            }
            column(ApprovalTime_SalesApprovalEntry; "Sales Approval Entry"."Approval Time")
            {
            }
            column(UserName_SalesApprovalEntry; "Sales Approval Entry"."User Name")
            {
            }
            column(ApproverName_SalesApprovalEntry; "Sales Approval Entry"."Approver Name")
            {
            }
            column(VersionNo_SalesApprovalEntry; "Sales Approval Entry"."Version No.")
            {
            }
            column(DateSentforApproval_SalesApprovalEntry; "Sales Approval Entry"."Date Sent for Approval")
            {
            }
            column(TimeSentforApproval_SalesApprovalEntry; "Sales Approval Entry"."Time Sent for Approval")
            {
            }
            column(CreditLimitApprovarID_SalesApprovalEntry; "Sales Approval Entry"."Credit Limit Approvar ID")
            {
            }
            column(CreditLimitApprovarName_SalesApprovalEntry; "Sales Approval Entry"."Credit Limit Approvar Name")
            {
            }
            column(CreditLimitApprovalDate_SalesApprovalEntry; "Sales Approval Entry"."Credit Limit Approval Date")
            {
            }
            column(CreditLimitApprovalTime_SalesApprovalEntry; "Sales Approval Entry"."Credit Limit Approval Time")
            {
            }
            column(OverDueApprovarID_SalesApprovalEntry; "Sales Approval Entry"."Over Due Approvar ID")
            {
            }
            column(OverDueApprovarName_SalesApprovalEntry; "Sales Approval Entry"."Over Due Approvar Name")
            {
            }
            column(OverDueApprovalDate_SalesApprovalEntry; "Sales Approval Entry"."Over Due Approval Date")
            {
            }
            column(OverDueApprovalTime_SalesApprovalEntry; "Sales Approval Entry"."Over Due Approval Time")
            {
            }
            column(Level2ApprovarID_SalesApprovalEntry; "Sales Approval Entry"."Level2 Approvar ID")
            {
            }
            column(Level2ApprovarName_SalesApprovalEntry; "Sales Approval Entry"."Level2 Approvar Name")
            {
            }
            column(Level2ApprovarDate_SalesApprovalEntry; "Sales Approval Entry"."Level2 Approvar Date")
            {
            }
            column(Level2ApprovarTime_SalesApprovalEntry; "Sales Approval Entry"."Level2 Approvar Time")
            {
            }
            column(SequenceNo_SalesApprovalEntry; "Sales Approval Entry"."Sequence No.")
            {
            }
            column(DivisionCode_SalesApprovalEntry; "Sales Approval Entry"."Division Code")
            {
            }
            column(Rejected_SalesApprovalEntry; "Sales Approval Entry".Rejected)
            {
            }
            column(RejectedDate_SalesApprovalEntry; "Sales Approval Entry"."Rejected Date")
            {
            }
            column(RejectedTime_SalesApprovalEntry; "Sales Approval Entry"."Rejected Time")
            {
            }
            column(Cancelled_SalesApprovalEntry; "Sales Approval Entry".Cancelled)
            {
            }
            column(CancelledDate_SalesApprovalEntry; "Sales Approval Entry"."Cancelled Date")
            {
            }
            column(CancelledTime_SalesApprovalEntry; "Sales Approval Entry"."Cancelled Time")
            {
            }
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
}


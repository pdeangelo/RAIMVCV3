using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAIMVCV3.Models
{
    public class Loan
    {
        public Loan()
        {

        }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public bool IsChecked { get; set; }
        public int LoanID { get; set; }
        [DisplayName("Status")]
        public int? LoanStatusID { get; set; }
        [DisplayName("Client")]
        public int? ClientID { get; set; }
        [DisplayName("Dwelling Type")]
        public int? DwellingTypeID { get; set; }
        [DisplayName("State")]
        public int? StateID { get; set; }
        [DisplayName("Entity")]
        public int? EntityID { get; set; }
        [DisplayName("Investor")]
        public int? InvestorID { get; set; }

        [DisplayName("Loan Type")]
        public int? LoanTypeID { get; set; }

        [DisplayName("Loan Number")]
        public string LoanNumber { get; set; }
        [Column(TypeName = "date")]
        [DisplayName("Funding Date")]
        public DateTime? LoanFundingDate { get; set; }

        [StringLength(200)]
        [DisplayName("Mortgagee")]
        [DefaultValue("")]
        public string LoanMortgagee { get; set; }

        [DisplayName("Address")]
        [DefaultValue("")]
        public string LoanPropertyAddress { get; set; }

        [DisplayName("Interest Rate")]
        public double? LoanInterestRate { get; set; }

        [DisplayName("Mortgage Amount")]
        
        public double? LoanMortgageAmount { get; set; }

        [DisplayName("Advance Rate")]
        public double? LoanAdvanceRate { get; set; }


        [Column(TypeName = "date")]
        [DisplayName("Entered Date")]
        public DateTime? LoanEnteredDate { get; set; }

        [DisplayName("Updated Date")]
        public DateTime? LoanUpdateDate { get; set; }

        public int? LoanUpdateUserID { get; set; }

        [StringLength(200)]
        [DisplayName("Business")]
        [DefaultValue("")]
        public string LoanMortgageeBusiness { get; set; }

        //UW Fields
        [DisplayName("Loan Application")]
        public bool? LoanUW10031008LoanApplication { get; set; }
        [DisplayName("Allonge Promissory Note")]

        public bool? LoanUWAllongePromissoryNote { get; set; }
        [DisplayName("Appraisal %")]

        public bool? LoanUWAppraisal { get; set; }
        [DisplayName("Appraisal Amount")]

        public double? LoanUWAppraisalAmount { get; set; }
        [DisplayName("Post Repair Appraisal Amount")]

        public double? LoanUWPostRepairAppraisalAmount { get; set; }
        [DisplayName("Assignment of Mortgage")]

        public bool? LoanUWAssignmentOfMortgage { get; set; }
        [DisplayName("Background Check")]

        public bool? LoanUWBackgroundCheck { get; set; }
        [DisplayName("Cert of Good Standing Formation")]

        public bool? LoanUWCertofGoodStandingFormation { get; set; }
        [DisplayName("Closing Protection Letter")]

        public bool? LoanUWClosingProtectionLetter { get; set; }
        [DisplayName("Loan Committee Review")]

        public bool? LoanUWCommitteeReview { get; set; }
        [DisplayName("Credit Report")]

        public bool? LoanUWCreditReport { get; set; }
        [DisplayName("Flood Certificate")]

        public bool? LoanUWFloodCertificate { get; set; }
        [DisplayName("HUD 1 Settlement Statement")]

        public bool? LoanUWHUD1SettlementStatement { get; set; }
        [DisplayName("Homeowners Insurance")]

        public bool? LoanUWHomeownersInsurance { get; set; }
        [DisplayName("Loan Package")]

        public bool? LoanUWLoanPackage { get; set; }
        [DisplayName("Loan Sizer Loan Submission Tape")]

        public bool? LoanUWLoanSizerLoanSubmissionTape { get; set; }
        [DisplayName("Title Commitment")]

        public bool? LoanUWTitleCommitment { get; set; }
        [DisplayName("Clayton Report Approval Email")]

        public bool? LoanUWClaytonReportApprovalEmail { get; set; }
        [DisplayName("Zillow Sq Ft")]

        public double? LoanUWZillowSqFt { get; set; }
        [DisplayName("Is Complete")]

        public bool? LoanUWIsComplete { get; set; }
        
        public int? CompletedBy { get; set; }

        //end UW Fields
        //Fudning Fields 
        [Column(TypeName = "date")]
        [DisplayName("Date Deposited in Escrow")]
        public DateTime? DateDepositedInEscrow { get; set; }

        [Column(TypeName = "date")]
        [DisplayName("Bailee Letter Date")]
        public DateTime? BaileeLetterDate { get; set; }

        [Column(TypeName = "date")]
        [DisplayName("Investor Proceeds Date")]
        public DateTime? InvestorProceedsDate { get; set; }

        [DisplayName("Investor Proceeds")]
        [DefaultValue(0)]
        public double? InvestorProceeds { get; set; }

        [Column(TypeName = "date")]
        [DisplayName("Closing Date")]
        public DateTime? ClosingDate { get; set; }
        //End funding fields
        //Fee Fields

        [DisplayName("Minimum Interest")]
        [DefaultValue(0)]
        public double? MinimumInterest { get; set; }

        [DisplayName("Origination Discount")]
        [DefaultValue(0)]
        public double? OriginationDiscount { get; set; }

        [DisplayName("Origination Discount 2")]
        [DefaultValue(0)]
        public double? OriginationDiscount2 { get; set; }

        [DisplayName("Origination Discount Num Days")]
        [DefaultValue(0)]
        public int? OriginationDiscountNumDays { get; set; }

        [DisplayName("Origination Discount Num Days 2")]
        [DefaultValue(0)]
        public int? OriginationDiscountNumDays2 { get; set; }

        [DisplayName("Interest Based On Advance")]
        public bool? InterestBasedOnAdvance { get; set; }

        [DisplayName("Origination Based On Advance")]
        public bool? OriginationBasedOnAdvance { get; set; }

        [DisplayName("No Interest?")]
        public bool? NoInterest { get; set; }

        [DisplayName("Interest Fee")]
        [DefaultValue(0)]
        public double? InterestFee { get; set; }

        [DisplayName("Origination Fee")]
        [DefaultValue(0)]
        public double? OriginationFee { get; set; }

        [DisplayName("Underwriting Fee")]
        [DefaultValue(0)]
        public double? UnderwritingFee { get; set; }

        [DisplayName("Cust Svc Underwriting Discount")]
        [DefaultValue(0)]
        public double? CustSvcUnderwritingDiscount { get; set; }

        [DisplayName("Cust Svc Interest Discount")]
        [DefaultValue(0)]
        public double? CustSvcInterestDiscount { get; set; }

        [DisplayName("Cust Svc Origination Discount")]
        [DefaultValue(0)]
        public double? CustSvcOriginationDiscount { get; set; }

        [DisplayName("Prime Rate")]
        [DefaultValue(0)]
        public double? ClientPrimeRate { get; set; }

        [DisplayName("Prime Rate Spread")]
        [DefaultValue(0)]
        public double? ClientPrimeRateSpread { get; set; }

        //end fee fields
        //Computed Fields
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public double DaysOutstandingClosed
        {
            get {
                if (InvestorProceedsDate == null || DateDepositedInEscrow == null)
                    return 0;
                else
                    return (double)DateDepositedInEscrow.Value.Subtract(InvestorProceedsDate.Value).TotalDays;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public double DaysOutstandingPending
        {
            get
            {
                if (InvestorProceedsDate == null)
                    return 0;
                else if (OpenMortgageBalance > 1)
                    return (double)DateTime.Now.Subtract(InvestorProceedsDate.Value).TotalDays;
                else
                    return 0;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? LoanAdvanceAmount
        {
            get { return LoanMortgageAmount * LoanAdvanceRate; }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? LoanReserveAmount
        {
            get { return LoanMortgageAmount * (1 - LoanAdvanceRate); }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? OpenMortgageBalance
        {
            get
            {
                if (InvestorProceeds == 0 && this.LoanStatus.LoanStatusName != "1 - Underwriting")
                    return LoanMortgageAmount;
                else
                    return 0;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? OpenAdvanceBalance
        {
            get
            {
                if (InvestorProceeds == 0 && this.LoanStatus.LoanStatusName != "1 - Underwriting")
                    return LoanAdvanceAmount;
                else
                    return 0;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? MortgageOriginatorProceeds
        {
            get
            {
                return InvestorProceeds - TotalFees - LoanAdvanceAmount;
            }
        }
        //[fn_RAIOriginationDiscount]
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? OriginationDiscountFee
        {
            get
            {
                double originationDiscount = 0;
                double daysOutstandingCalc = 0;

                if (DaysOutstandingClosed > 0)
                {
                    daysOutstandingCalc = DaysOutstandingClosed;
                    //Period 1

                    daysOutstandingCalc = daysOutstandingCalc - OriginationDiscountNumDays.Value;

                    //Period 2

                    if (daysOutstandingCalc > 0)
                    {
                        daysOutstandingCalc = daysOutstandingCalc - OriginationDiscountNumDays2.Value;
                        originationDiscount += OriginationDiscount2.Value;
                    }
                    // Period 3

                    if (daysOutstandingCalc > 0)
                    {
                        daysOutstandingCalc = daysOutstandingCalc - OriginationDiscountNumDays2.Value;
                        originationDiscount += OriginationDiscount2.Value;
                    }
                }

                else
                {
                    originationDiscount = 0;
                }
                    return originationDiscount;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? UnderwritingFeeCalc
        {
            get
            {
                if (DaysOutstandingClosed > 0)
                    return 100.0;
                else
                    return 0.0;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? LoanMinimumInterest
        {
            get {
                double minInterest = 0;
                double UseMinRate = 0;
                if (this.Client.MinimumInterest > 0)
                    UseMinRate = this.Client.MinimumInterest;

                else if (this.Client.ClientPrimeRateSpread > 0)
                    UseMinRate = this.Client.ClientPrimeRateSpread + this.Client.ClientPrimeRate;

                else
                    UseMinRate = LoanInterestRate.Value;

                if (UseMinRate < LoanInterestRate)
                    MinimumInterest = LoanInterestRate;

                else
                    MinimumInterest = UseMinRate;

                return minInterest;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? DailyInterestRate
        {
            get
            {
                return MinimumInterest / 360;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? InterestIncome
        {
            get
            { double interestIncome = 0;
                if (InterestBasedOnAdvance.Value)
                    interestIncome = DaysOutstandingClosed * DailyInterestRate.Value * LoanAdvanceAmount.Value;
                else
                    interestIncome = DaysOutstandingClosed * DailyInterestRate.Value * LoanMortgageAmount.Value;
                return interestIncome;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double AnnualizedYield
        {
            get
            {
                if (InvestorProceeds == null || LoanAdvanceAmount == null || DaysOutstandingClosed == null)
                    return 0;
                if (InvestorProceeds.Value > 0 && LoanAdvanceAmount.Value != 0 && DaysOutstandingClosed > 0)
                    return (double)(TotalFees / LoanAdvanceAmount / DaysOutstandingClosed * 360);
                else
                    return 0;
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? TotalTransfer
        {
            get
            {
                return InvestorProceeds -  LoanAdvanceAmount - (TotalFees);
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double? TotalFees
        {
            get
            {
                return(InterestFee +  OriginationFee + UnderwritingFee +
                       CustSvcInterestDiscount +  CustSvcOriginationDiscount + CustSvcUnderwritingDiscount);
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double WireFee
        {
            get
            {
                return (25);
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double AdvanceWithWireFee
        {
            get
            {
                return LoanAdvanceAmount == null ? 0 : LoanAdvanceAmount.Value - WireFee;
                
            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double GreaterofMinandCouponInterest
        {
            get
            {
                if (MinimumInterest != null && LoanInterestRate != null)
                    return MinimumInterest.Value > LoanInterestRate.Value ? MinimumInterest.Value : LoanInterestRate.Value;
                else
                    return 0;

            }
        }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public Double AppraisalPcnt
        {
            get
            {
                return LoanUWAppraisalAmount == null || LoanUWAppraisalAmount.Value == 0  ? 0 : LoanMortgageAmount.Value / LoanUWAppraisalAmount.Value;

            }
        }

        //End Computed Fields
        public virtual LoanStatus LoanStatus { get; set; }
        public virtual LoanType LoanType { get; set; }
        public virtual Client Client { get; set; }
        public virtual DwellingType DwellingType { get; set; }
        public virtual State State { get; set; }
        public virtual Entity Entity { get; set; }
        public virtual Investor Investor { get; set; }

    }
}

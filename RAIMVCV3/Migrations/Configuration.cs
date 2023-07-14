namespace RAIMVCV3.Migrations
{
    using RAIMVCV3.Models;
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;

    internal sealed class Configuration : DbMigrationsConfiguration<RAIMVCV3.Models.ApplicationDbContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;
        }

        protected override void Seed(RAIMVCV3.Models.ApplicationDbContext context)
        {
            //  This method will be called after migrating to the latest version.

            //  You can use the DbSet<T>.AddOrUpdate() helper extension method 
            //  to avoid creating duplicate seed data.
            context.Entity.AddOrUpdate(
    r => r.EntityID,
    new Entity() { EntityName = "AccuCredit", EntityBank = "TAB Bank", ABA = "124384657", Account = "450520543" },
    new Entity() { EntityName = "RAI Funding, LLC", EntityBank = "Dime Community Bank", ABA = "226070306", Account = "10000444084" }
);
            context.LoanType.AddOrUpdate(
                r => r.LoanTypeID,
                new LoanType() { LoanTypeName = "Fix and Flip" },
                new LoanType() { LoanTypeName = "30yr" }
            );
            context.LoanStatus.AddOrUpdate(
                r => r.LoanStatusID,
                new LoanStatus() { LoanStatusName = "1 - Underwriting" },
                new LoanStatus() { LoanStatusName = "2 - Funding" },
                new LoanStatus() { LoanStatusName = "3 - Open" },
                new LoanStatus() { LoanStatusName = "4 - Completed" },
                new LoanStatus() { LoanStatusName = "5 - Cancelled" },
                new LoanStatus() { LoanStatusName = "6 - Postponed" }

            );
            context.DwellingType.AddOrUpdate(
                r => r.DwellingTypeID,
               new DwellingType() { DwellingType1 = "Residential – 1-4 units" },
                new DwellingType() { DwellingType1 = "Residential  - Condo" },
                new DwellingType() { DwellingType1 = "Residential – Multi Family 5+ units" },
                new DwellingType() { DwellingType1 = "Commercial" }
            );

            context.State.AddOrUpdate(
                r => r.StateID,
               new State() { State1 = "Alabama" },
                new State() { State1 = "Alaska" },
                new State() { State1 = "Arizona" },
                new State() { State1 = "Arkansas" },
                new State() { State1 = "California" },
                new State() { State1 = "Colorado" },
                new State() { State1 = "Connecticut" },
                new State() { State1 = "Delaware" },
                new State() { State1 = "Florida" },
                new State() { State1 = "Georgia" },
                new State() { State1 = "Hawaii" },
                new State() { State1 = "Idaho" },
                new State() { State1 = "Illinois" },
                new State() { State1 = "Indiana" },
                new State() { State1 = "Iowa" },
                new State() { State1 = "Kansas" },
                new State() { State1 = "Kentucky" },
                new State() { State1 = "Louisiana" },
                new State() { State1 = "Maine" },
                new State() { State1 = "Maryland" },
                new State() { State1 = "Massachusetts" },
                new State() { State1 = "Michigan" },
                new State() { State1 = "Minnesota" },
                new State() { State1 = "Mississippi" },
                new State() { State1 = "Missouri" },
                new State() { State1 = "Montana" },
                new State() { State1 = "Nebraska" },
                new State() { State1 = "Nevada" },
                new State() { State1 = "New Hampshire" },
                new State() { State1 = "New Jersey" },
                new State() { State1 = "New Mexico" },
                new State() { State1 = "New York" },
                new State() { State1 = "North Carolina" },
                new State() { State1 = "North Dakota" },
                new State() { State1 = "Ohio" },
                new State() { State1 = "Oklahoma" },
                new State() { State1 = "Oregon" },
                new State() { State1 = "Pennsylvania" },
                new State() { State1 = "Rhode Island" },
                new State() { State1 = "South Carolina" },
                new State() { State1 = "South Dakota" },
                new State() { State1 = "Tennessee" },
                new State() { State1 = "Texas" },
                new State() { State1 = "Utah" },
                new State() { State1 = "Vermont" },
                new State() { State1 = "Virginia" },
                new State() { State1 = "Washington" },
                new State() { State1 = "West Virginia" },
                new State() { State1 = "Wisconsin" },
                new State() { State1 = "Wyoming" }

            );

            context.Investor.AddOrUpdate(
                r => r.InvestorID,
               new Investor() { InvestorName = "Toorak", CustodianName = "US Bank" },
               new Investor() { InvestorName = "Toorak", CustodianName = "US Bank" },
                new Investor() { InvestorName = "Smith Graham", CustodianName = "US Bank" },
                new Investor() { InvestorName = "Alpha Flow", CustodianName = "US Bank" },
                new Investor() { InvestorName = "Churchill", CustodianName = "US Bank" },
                new Investor() { InvestorName = "A & D", CustodianName = "US Bank" },
                new Investor() { InvestorName = "Bayview", CustodianName = String.Empty },
                new Investor() { InvestorName = "ViaNova", CustodianName = String.Empty },
                new Investor() { InvestorName = "Saluda Grade", CustodianName = String.Empty },
                new Investor() { InvestorName = "Reigo", CustodianName = String.Empty },
                new Investor() { InvestorName = "Palisades Group", CustodianName = String.Empty },
                new Investor() { InvestorName = "SG Capital Partners", CustodianName = String.Empty },
                new Investor() { InvestorName = "Fidelis", CustodianName = String.Empty },
                new Investor() { InvestorName = "Stormfield Capital", CustodianName = String.Empty },
                new Investor() { InvestorName = "Nextres", CustodianName = String.Empty },
                new Investor() { InvestorName = "TGP Ventures", CustodianName = String.Empty }
                 );



            //context.Loans.AddOrUpdate(
            //    r => r.LoanID,
            //    new Loan()
            //    {
            //        ClientID = 1,
            //        LoanStatusID = 1,
            //        DwellingTypeID = 1,
            //        EntityID = 1,
            //        StateID = 1,
            //        InvestorID = 1,
            //        LoanTypeID = 1,
            //        LoanNumber = "1234",
            //        LoanPropertyAddress = "123 Main St, New York, NY 12345",
            //        LoanAdvanceRate = .85,
            //        LoanMortgageAmount = 100000,
            //        LoanUWAppraisalAmount = 200000,
            //        LoanMortgagee = "John Doe",
            //        LoanInterestRate = .05,
            //        DateDepositedInEscrow = new DateTime(2023, 1, 1),
            //        BaileeLetterDate = new DateTime(2023, 1, 1),
            //        InvestorProceedsDate = new DateTime(2023, 2, 1),
            //        InvestorProceeds = 80000,
            //        ClosingDate = new DateTime(2023, 2, 1),
            //        MinimumInterest = .1,
            //        OriginationDiscount = 1,
            //        OriginationDiscount2 = 2,
            //        OriginationDiscountNumDays = 30,
            //        OriginationDiscountNumDays2 = 15,
            //        InterestBasedOnAdvance = true,
            //        OriginationBasedOnAdvance = false,
            //        NoInterest = false,
            //        InterestFee = 1000,
            //        OriginationFee = 100,
            //        UnderwritingFee = 500,
            //        CustSvcUnderwritingDiscount = 0,
            //        CustSvcInterestDiscount = 0,
            //        CustSvcOriginationDiscount = 0,
            //        ClientPrimeRate = .04,
            //        ClientPrimeRateSpread = .05
            //    }
            //    );

            //context.Loans.AddOrUpdate(
            //    r => r.LoanID,
            //    new Loan()
            //    {
            //        ClientID = 2,
            //        LoanStatusID = 1,
            //        DwellingTypeID = 1,
            //        EntityID = 1,
            //        StateID = 1,
            //        InvestorID = 1,
            //        LoanTypeID = 1,
            //        LoanNumber = "56789",
            //        LoanPropertyAddress = "123 Main St, New York, NY 12345",
            //        LoanAdvanceRate = .85,
            //        LoanMortgageAmount = 100000,
            //        LoanUWAppraisalAmount = 200000,
            //        LoanMortgagee = "John Doe",
            //        LoanInterestRate = .05,
            //        DateDepositedInEscrow = new DateTime(2023, 1, 1),
            //        BaileeLetterDate = new DateTime(2023, 1, 1),
            //        InvestorProceedsDate = new DateTime(2023, 2, 1),
            //        InvestorProceeds = 80000,
            //        ClosingDate = new DateTime(2023, 2, 1),
            //        MinimumInterest = .1,
            //        OriginationDiscount = 1,
            //        OriginationDiscount2 = 2,
            //        OriginationDiscountNumDays = 30,
            //        OriginationDiscountNumDays2 = 15,
            //        InterestBasedOnAdvance = true,
            //        OriginationBasedOnAdvance = false,
            //        NoInterest = false,
            //        InterestFee = 1000,
            //        OriginationFee = 100,
            //        UnderwritingFee = 500,
            //        CustSvcUnderwritingDiscount = 0,
            //        CustSvcInterestDiscount = 0,
            //        CustSvcOriginationDiscount = 0,
            //        ClientPrimeRate = .04,
            //        ClientPrimeRateSpread = .05
            //    }

                //);
        }
    }
}

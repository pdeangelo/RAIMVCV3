using RAIMVCV3.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Linq;
using System.Web;

namespace RAIMVCV3
{
    //public class DatabaseInitializer : DropCreateDatabaseIfModelChanges<ApplicationDbContext>
    //{
    //    protected override void Seed(ApplicationDbContext context)
    //    {
    //        var loanStatus1 = new LoanStatus()
    //        {
    //            LoanStatusName = "1 - Underwriting"
    //        };
    //        var client1 = new Client()
    //        {
    //            ClientName = "Client 1",
    //            AdvanceRate = 0,
    //            MinimumInterest = 0,
    //            ClientPrimeRate = 0,
    //            ClientPrimeRateSpread = 0,
    //            OriginationDiscount = 0,
    //            OriginationDiscount2 = 0,
    //            OriginationDiscountNumDays = 0,
    //            OriginationDiscountNumDays2 = 0,
    //            InterestBasedOnAdvance = false,
    //            NoInterest = false
    //        };
    //        var state = new State()
    //        {
    //            State1 = "NY"
    //        };
    //        var investor = new Investor()
    //        {
    //            InvestorName = "Investor 1"
    //        };
    //        var dwellingType = new DwellingType()
    //        {
    //            DwellingType1 = "Dwelling Type 1"
    //        };
    //        var entity = new Entity()
    //        {
    //            EntityName = "Entity 1"
    //        };
    //        var loanType = new LoanType()
    //        {
    //            LoanTypeName = "Loan Type 1"
    //        };
    //        var loan1 = new Loan()
    //        {
    //            LoanNumber = "Loan 1"
    //        };
    //        loan1.LoanStatus = loanStatus1;
    //        loan1.Investor = investor;
    //        loan1.Entity = entity;
    //        loan1.DwellingType = dwellingType;
    //        loan1.Client = client1;
    //        loan1.State = state;
    //        loan1.LoanType = loanType;
    //        loan1.BaileeLetterDate = DateTime.Now;
    //        loan1.ClosingDate = DateTime.Now;
    //        loan1.LoanMortgageAmount = 100000;
    //        loan1.LoanInterestRate = .08;
    //        loan1.LoanMortgagee = "John Smith";
    //        loan1.LoanPropertyAddress = "123 Main St, Anytown, NY 12345";
    //        context.Loans.Add(loan1);
    //        var loan2 = new Loan()
    //        {
    //            LoanNumber = "Loan 1"
    //        };
    //        loan2.LoanStatus = loanStatus1;
    //        loan2.Investor = investor;
    //        loan2.Entity = entity;
    //        loan2.DwellingType = dwellingType;
    //        loan2.Client = client1;
    //        loan2.State = state;
    //        loan2.LoanType = loanType;
    //        loan2.BaileeLetterDate = DateTime.Now;
    //        loan2.ClosingDate = DateTime.Now;
    //        loan2.LoanMortgageAmount = 100000;
    //        loan2.LoanInterestRate = .08;
    //        loan2.LoanMortgagee = "John Smith";
    //        loan2.LoanPropertyAddress = "123 Main St, Anytown, NY 12345";
    //        context.Loans.Add(loan2);

    //        //var roleView = new Role()
    //        //{
    //        //    RoleName = "View Only"
    //        //};
    //        //var roleAccounting = new Role()
    //        //{
    //        //    RoleName = "Accounting"
    //        //};
    //        //var roleUnderwriting = new Role()
    //        //{
    //        //    RoleName = "Underwriting"
    //        //};
    //        //context.Role.Add(roleAccounting);
    //        //context.Role.Add(roleUnderwriting);
    //        //context.Role.Add(roleView);
    //        //var user = new User()
    //        //{
    //        //    WinUserID = "Paul",
    //        //    UserName = "Paul",
    //        //    //Role = roleView,
    //        //    IsAdmin = false,
    //        //    Password = "password"

    //        //};
    //        //context.User.Add(user);

    //        try
    //        {
    //            context.SaveChanges();
    //        }
    //        catch (DbEntityValidationException e)
    //        {
    //            foreach (var eve in e.EntityValidationErrors)
    //            {
    //                Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
    //                    eve.Entry.Entity.GetType().Name, eve.Entry.State);
    //                foreach (var ve in eve.ValidationErrors)
    //                {
    //                    Console.WriteLine("- Property: \"{0}\", Error: \"{1}\"",
    //                        ve.PropertyName, ve.ErrorMessage);
    //                }
    //            }
    //            throw;
    //        }

    //    }
    //}
}

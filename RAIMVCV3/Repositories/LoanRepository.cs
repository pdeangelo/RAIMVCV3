
using RAIMVCV3.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RAIMVCV3
{
    public class LoanRepository
    {
        public IList<Loan> GetLoans()
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.Loans
                  .Include(l => l.State)
                  .Include(l => l.Client)
                  .Include(l => l.LoanStatus)
                  .Include(l => l.Entity)
                  .Include(l => l.Investor)
                  .Include(l => l.DwellingType)
                  .Include(l => l.LoanType)
                  .ToList();
            }
        }
        /// <summary>
        /// Returns a single loan.
        /// </summary>
        /// <returns>A fully populated Loan entity instance.</returns>
        public Loan GetLoan(int loanID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.Loans
                 .Include(l => l.State)
                 .Include(l => l.Client)
                 .Include(l => l.LoanStatus)
                 .Include(l => l.Entity)
                 .Include(l => l.Investor)
                 .Include(l => l.DwellingType)
                 .Include(l => l.LoanType)
                   .Where(cb => cb.LoanID == loanID)
                   .SingleOrDefault();
            }
        }
        public void AddLoan(Loan loan)
        {
            using (ApplicationDbContext context = GetContext())
            {
                if (loan.LoanMortgagee == null) loan.LoanMortgagee = String.Empty;
                if (loan.LoanMortgageeBusiness == null) loan.LoanMortgageeBusiness = String.Empty;
                if (loan.LoanPropertyAddress == null) loan.LoanPropertyAddress = String.Empty;

                loan.LoanStatusID = 1;
                UpdateLoanFeeInfo(loan);
                context.Loans.Add(loan);

                if (loan.Client != null && loan.Client.ClientID > 0)
                {
                    context.Entry(loan.Client).State = EntityState.Unchanged;
                }

                if (loan.DwellingType != null && loan.DwellingType.DwellingTypeID > 0)
                {
                    context.Entry(loan.DwellingType).State = EntityState.Unchanged;
                }

                if (loan.Entity != null && loan.Entity.EntityID > 0)
                {
                    context.Entry(loan.Entity).State = EntityState.Unchanged;
                }

                if (loan.LoanStatus != null && loan.LoanStatus.LoanStatusID > 0)
                {
                    context.Entry(loan.LoanStatus).State = EntityState.Unchanged;
                }

                if (loan.LoanType != null && loan.LoanType.LoanTypeID > 0)
                {
                    context.Entry(loan.LoanType).State = EntityState.Unchanged;
                }

                if (loan.State != null && loan.State.StateID > 0)
                {
                    context.Entry(loan.State).State = EntityState.Unchanged;
                }

                context.SaveChanges();
            }
        }
        public void UpdateLoanFeeInfo(Loan loan)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var client = context.Client.Where(x => x.ClientID == loan.ClientID).SingleOrDefault();
                
                //Copy client fields used for calcs in case they change after the loans close and we want to calc the same fee
                loan.ClientPrimeRate = client.ClientPrimeRate;
                loan.ClientPrimeRateSpread = client.ClientPrimeRateSpread;
                loan.OriginationDiscount = client.OriginationDiscount;
                loan.OriginationDiscount2 = client.OriginationDiscount2;
                loan.OriginationDiscountNumDays = client.OriginationDiscountNumDays;
                loan.OriginationDiscountNumDays2 = client.OriginationDiscountNumDays2;
                loan.InterestBasedOnAdvance = client.InterestBasedOnAdvance;
                loan.OriginationBasedOnAdvance = client.OriginationBasedOnAdvance;
                loan.NoInterest = client.NoInterest;
            }
           
            loan.InterestFee = loan.InterestIncome;
            loan.OriginationFee = loan.OriginationDiscountFee;
            loan.UnderwritingFee = loan.UnderwritingFeeCalc;
        }
        public void UpdateLoan(Loan loan)
        {
            using (ApplicationDbContext context = GetContext())
            {
                if (!loan.LoanUWIsComplete.Value)
                    loan.LoanStatusID = 1;
                else if (loan.LoanUWIsComplete.Value &&
                         (loan.InvestorProceedsDate == null) && loan.DateDepositedInEscrow == null)
                    loan.LoanStatusID = 2;
                else if (loan.DateDepositedInEscrow != null && loan.InvestorProceedsDate == null)
                    loan.LoanStatusID = 3;
                else
                    loan.LoanStatusID = 4;

                UpdateLoanFeeInfo(loan);
                context.Loans.Attach(loan);

                var loanEntry = context.Entry(loan);
                loanEntry.State = EntityState.Modified;
                //comicBookEntry.Property("IssueNumber").IsModified = false;

                context.SaveChanges();
            }
        }
        public void DeleteLoan(int loanID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var loan = new Loan() { LoanID = loanID };
                context.Entry(loan).State = EntityState.Deleted;

                context.SaveChanges();
            }
        }
        public SelectList GetInvestors()
        {
            using (ApplicationDbContext context = GetContext())
            {
                var investors = context.Investor.ToList();
                List<SelectListItem> list = new List<SelectListItem>();
                foreach (Investor investor in investors)
                {
                    list.Add(new SelectListItem()
                    {
                        Value = investor.InvestorID.ToString(),
                        Text = investor.InvestorName
                    });
                }
                return new SelectList(list, "Value", "Text");
            }


        }
        public SelectList GetTypes()
        {
            using (ApplicationDbContext context = GetContext())
            {
                var loanTypes = context.LoanType.ToList();
                List<SelectListItem> list = new List<SelectListItem>();
                foreach (LoanType loanType in loanTypes)
                {
                    list.Add(new SelectListItem()
                    {
                        Value = loanType.LoanTypeID.ToString(),
                        Text = loanType.LoanTypeName
                    });
                }
                return new SelectList(list, "Value", "Text");
            }
        }
        public SelectList GetEntities()
        {
            using (ApplicationDbContext context = GetContext())
            {
                var entities = context.Entity.ToList();
                List<SelectListItem> list = new List<SelectListItem>();
                foreach (Entity entity in entities)
                {
                    list.Add(new SelectListItem()
                    {
                        Value = entity.EntityID.ToString(),
                        Text = entity.EntityName
                    });
                }
                return new SelectList(list, "Value", "Text");
            }

        }
        public Entity GetEntity(int entityID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.Entity
                   .Where(cb => cb.EntityID == entityID)
                   .SingleOrDefault();
            }
        }
        public Client GetClient(int clientID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.Client
                   .Where(cb => cb.ClientID == clientID)
                   .SingleOrDefault();
            }
        }
        public SelectList GetStatus()
        {
            using (ApplicationDbContext context = GetContext())
            {
                var statuss = context.LoanStatus.ToList();
                List<SelectListItem> list = new List<SelectListItem>();
                foreach (LoanStatus status in statuss)
                {
                    list.Add(new SelectListItem()
                    {
                        Value = status.LoanStatusID.ToString(),
                        Text = status.LoanStatusName
                    });
                }
                return new SelectList(list, "Value", "Text");
            }

        }
        public SelectList GetStates()
        {
            using (ApplicationDbContext context = GetContext())
            {
                var states = context.State.ToList();
                List<SelectListItem> list = new List<SelectListItem>();
                foreach (State state in states)
                {
                    list.Add(new SelectListItem()
                    {
                        Value = state.StateID.ToString(),
                        Text = state.State1
                    });
                }
                return new SelectList(list, "Value", "Text");
            }

        }
        //public SelectList GetRoles()
        //{
        //    using (ApplicationDbContext context = GetContext())
        //    {
        //        var roles = context.Role.ToList();
        //        List<SelectListItem> list = new List<SelectListItem>();
        //        foreach (Role role in roles)
        //        {
        //            list.Add(new SelectListItem()
        //            {
        //                Value = role.RoleID.ToString(),
        //                Text = role.RoleName
        //            });
        //        }
        //        return new SelectList(list, "Value", "Text");
        //    }

        //}
        public SelectList GetUsers()
        {
            using (ApplicationDbContext context = GetContext())
            {
                var users = context.Users.ToList();
                List<SelectListItem> list = new List<SelectListItem>();
                foreach (ApplicationUser user in users)
                {
                    list.Add(new SelectListItem()
                    {
                        Value = user.UserName.ToString(),
                        Text = user.UserName
                    });
                }
                return new SelectList(list, "Value", "Text");
            }

        }

        public SelectList GetClients()
        {
            using (ApplicationDbContext context = GetContext())
            {
                var clients = context.Client.ToList();
                List<SelectListItem> list = new List<SelectListItem>();
                foreach (Client client in clients)
                {
                    list.Add(new SelectListItem()
                    {
                        Value = client.ClientID.ToString(),
                        Text = client.ClientName
                    });
                }
                return new SelectList(list, "Value", "Text");
            }


        }
        public List<SalesReport> GetSalesReport(DateTime fDate, DateTime tDate)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var tDateParameter = new SqlParameter("@FromDate", SqlDbType.DateTime);
                var fDateParameter = new SqlParameter("@ToDate", SqlDbType.DateTime);

                //var result = context.Database
                //    .SqlQuery<SalesReport>("TableFunding_SalesReport @FromDate @ToDate", tDateParameter, fDateParameter)
                //    .ToList();

                var result = context.Database.SqlQuery<SalesReport>(
                    "TableFunding_SalesReport @FromDate, @ToDate",
                    new SqlParameter("@FromDate", fDate),
                    new SqlParameter("@ToDate", tDate)
                    ).ToList();
                return result;
            }
        }
        /// <summary>
        /// Private method that returns a database context.
        /// </summary>
        /// <returns>An instance of the Context class.</returns>
        static ApplicationDbContext GetContext()
        {
            var context = new ApplicationDbContext();
            context.Database.Log = (message) => Debug.WriteLine(message);
            return context;
        }

    }
}
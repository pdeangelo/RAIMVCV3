using RAIMVCV3.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using System.Web.Mvc;
using System.Diagnostics;
using RAIMVCV3.Models;

namespace RAIMVCV3.Repository
{
    public class MiscRepository
    {
        public IList<LoanStatus> GetLoanStatuss()
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.LoanStatus
                  .ToList();
            }
        }
        /// <summary>
        /// Returns a single loan.
        /// </summary>
        /// <returns>A fully populated Loan entity instance.</returns>
        public LoanStatus GetLoanStatus(int loanStatusID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.LoanStatus
                   .Where(cb => cb.LoanStatusID == loanStatusID)
                   .SingleOrDefault();
            }
        }
        public void AddLoanStatus(LoanStatus loanStatus)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.LoanStatus.Add(loanStatus);

                context.SaveChanges();
            }
        }
        public void UpdateLoanStatus(LoanStatus loanStatus)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.LoanStatus.Attach(loanStatus);

                var loanEntry = context.Entry(loanStatus);

                loanEntry.State = EntityState.Modified;
                context.SaveChanges();
            }
        }
        public void DeleteLoanStatus(int loanStatusID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var loanStatus = new LoanStatus() { LoanStatusID = loanStatusID };
                context.Entry(loanStatus).State = EntityState.Deleted;

                context.SaveChanges();
            }
        }

        //
        public IList<DwellingType> GetDwellingTypes()
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.DwellingType
                  .ToList();
            }
        }
        /// <summary>
        /// Returns a single loan.
        /// </summary>
        /// <returns>A fully populated Loan entity instance.</returns>
        public DwellingType GetDwellingType (int dwellingTypeID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.DwellingType
                   .Where(cb => cb.DwellingTypeID == dwellingTypeID)
                   .SingleOrDefault();
            }
        }
        public void AddDwellingType(DwellingType dwellingType)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.DwellingType.Add(dwellingType);

                context.SaveChanges();
            }
        }
        public void UpdateDwellingType(DwellingType dwellingType)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.DwellingType.Attach(dwellingType);

                var loanEntry = context.Entry(dwellingType);

                loanEntry.State = EntityState.Modified;
                context.SaveChanges();
            }
        }
        public void DeleteDwellingType(int dwellingTypeID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var dwellingType = new DwellingType() { DwellingTypeID = dwellingTypeID };
                context.Entry(dwellingType).State = EntityState.Deleted;

                context.SaveChanges();
            }
        }

        //
        public IList<State> GetStates()
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.State
                  .ToList();
            }
        }
        /// <summary>
        /// Returns a single loan.
        /// </summary>
        /// <returns>A fully populated Loan entity instance.</returns>
        public State GetState(int stateID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.State
                   .Where(cb => cb.StateID == stateID)
                   .SingleOrDefault();
            }
        }
        public void AddState(State state)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.State.Add(state);

                context.SaveChanges();
            }
        }
        public void UpdateState(State state)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.State.Attach(state);

                var loanEntry = context.Entry(state);

                loanEntry.State = EntityState.Modified;
                context.SaveChanges();
            }
        }
        public void DeleteState(int stateID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var state = new State() { StateID = stateID };
                context.Entry(state).State = EntityState.Deleted;

                context.SaveChanges();
            }
        }

        //
        public IList<LoanType> GetLoanTypes()
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.LoanType
                  .ToList();
            }
        }
        /// <summary>
        /// Returns a single loan.
        /// </summary>
        /// <returns>A fully populated Loan entity instance.</returns>
        public LoanType GetLoanType(int loanTypeID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.LoanType
                   .Where(cb => cb.LoanTypeID == loanTypeID)
                   .SingleOrDefault();
            }
        }
        public void AddLoanType(LoanType loanType)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.LoanType.Add(loanType);

                context.SaveChanges();
            }
        }
        public void UpdateLoanType(LoanType loanType)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.LoanType.Attach(loanType);

                var loanEntry = context.Entry(loanType);

                loanEntry.State = EntityState.Modified;
                context.SaveChanges();
            }
        }
        public void DeleteLoanType(int loanTypeID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var loanType = new LoanType() { LoanTypeID = loanTypeID };
                context.Entry(loanType).State = EntityState.Deleted;

                context.SaveChanges();
            }
        }

        //
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
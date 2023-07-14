using RAIMVCV3.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using System.Web.Mvc;
using System.Diagnostics;

namespace RAIMVCV3.Repository
{
    public class InvestorRepository
    {
        public IList<Investor> GetInvestors()
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.Investor                  
                  .ToList();
            }
        }
        /// <summary>
        /// Returns a single loan.
        /// </summary>
        /// <returns>A fully populated Loan entity instance.</returns>
        public Investor GetInvestor(int investorID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                return context.Investor
                   .Where(cb => cb.InvestorID == investorID)
                   .SingleOrDefault();
            }
        }
        public void AddInvestor(Investor investor)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.Investor.Add(investor);

                context.SaveChanges();
            }
        }
        public void UpdateInvestor(Investor investor)
        {
            using (ApplicationDbContext context = GetContext())
            {
                context.Investor.Attach(investor);

                var clientEntry = context.Entry(investor);
                clientEntry.State = EntityState.Modified;
                //comicBookEntry.Property("IssueNumber").IsModified = false;

                context.SaveChanges();
            }
        }
        public void DeleteInvestor(int investorID)
        {
            using (ApplicationDbContext context = GetContext())
            {
                var investor = new Investor() { InvestorID = investorID };
                context.Entry(investor).State = EntityState.Deleted;

                context.SaveChanges();
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
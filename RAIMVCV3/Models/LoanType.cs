using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAIMVCV3.Models
{
    public class LoanType
    {
        public LoanType()
        {
            Loans = new HashSet<Loan>();
        }

        public int LoanTypeID { get; set; }
        [DisplayName("Loan Type")]
        public string LoanTypeName { get; set; }
        public virtual ICollection<Loan> Loans { get; set; }
    }
}

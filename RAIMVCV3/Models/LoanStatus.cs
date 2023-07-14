using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAIMVCV3.Models
{
    public class LoanStatus
    {
        public LoanStatus()
        {
            Loans = new List<Loan>();
        }
        public int LoanStatusID { get; set; }
        [DisplayName("Loan Status")]
        public string LoanStatusName { get; set; }
        public ICollection<Loan> Loans { get; set; }
    }
}

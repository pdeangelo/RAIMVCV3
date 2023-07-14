using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAIMVCV3.Models
{
    public class Investor
    {
        public int InvestorID { get; set; }

        [StringLength(200)]
        [DisplayName("Investor Name")]
        public string InvestorName { get; set; }

        [StringLength(200)]
        [DisplayName("Custodian Name")]
        public string CustodianName { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Loan> Loans { get; set; }
    }
}

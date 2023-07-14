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
    public class LoanForLetters
    {
        public LoanForLetters()
        {

        }

        public bool IsChecked { get; set; }
        public int LoanID { get; set; }
       
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

    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAIMVCV3.Models
{
    public class Client
    {
        public Client()
        {
            Loans = new List<Loan>();
        }
        public int ClientID { get; set; }
        [DisplayName("Client Name")]
        public string ClientName { get; set; }

        [DisplayName("Advance Rate")]
        public double AdvanceRate { get; set; }

        [DisplayName("Minimum Interest")]
        public double MinimumInterest { get; set; }

        [DisplayName("Prime Rate")]
        public double ClientPrimeRate { get; set; }

        [DisplayName("Prime Rate Spread")]
        public double ClientPrimeRateSpread { get; set; }

        [DisplayName("Origination Discount")]
        public double OriginationDiscount { get; set; }
        [DisplayName("Origination Discount 2")]
        public double OriginationDiscount2 { get; set; }
        [DisplayName("Origination Discount Num Days")]
        public int OriginationDiscountNumDays { get; set; }

        [DisplayName("Origination Discount Num Days 2")]
        public int OriginationDiscountNumDays2 { get; set; }
        [DisplayName("Interest Based On Advance?")]
        public bool InterestBasedOnAdvance { get; set; }
        [DisplayName("Origination  Based On Advance?")]
        public bool OriginationBasedOnAdvance { get; set; }
        [DisplayName("No Interest?")]
        public bool NoInterest { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Loan> Loans { get; set; }
    }
}

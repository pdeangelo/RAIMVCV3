using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RAIMVCV3.Models
{
    public class Entity
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Entity()
        {
            Loans = new HashSet<Loan>();
        }

        public int EntityID { get; set; }

        [StringLength(200)]
        [DisplayName("Entity")]
        public string EntityName { get; set; }

        [StringLength(100)]
        [DisplayName("Bank")]
        public string EntityBank { get; set; }

        [StringLength(100)]
        public string ABA { get; set; }

        [StringLength(100)]
        [DisplayName("Bank Account")]
        public string Account { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Loan> Loans { get; set; }
    }
}

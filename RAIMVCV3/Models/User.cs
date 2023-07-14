//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.ComponentModel.DataAnnotations;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace RAIMVCV3.Models
//{
//    public class User
//    {
//        public int UserID { get; set; }

//        [Required]
//        [StringLength(100)]
//        [DisplayName("User Name")]
//        public string UserName { get; set; }

//        [Required]
//        [StringLength(50)]
//        [DisplayName("Windows User ID")]
//        public string WinUserID { get; set; }

//        public int RoleID { get; set; }

//        [DisplayName("Is Admin?")]
//        public bool IsAdmin { get; set; }

//        [StringLength(20)]
//        [DisplayName("Password")]
//        public string Password { get; set; }

//        //[DisplayName("Role")]
//        //public virtual Role Role { get; set; }

//    }
//}

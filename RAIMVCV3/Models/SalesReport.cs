using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RAIMVCV3.Models
{
    public class SalesReport
    {
        public string ClientName {get; set;}
        public string EntityName {get; set;}
        public string Investor { get; set; }
        public string LoanNumber {get; set;}
        public string CustomerName {get; set;}
        public string GrossAmount {get; set;}
        public string Advance { get; set; }
        public string DateDepositedInEscrow {get; set;}
        public string State { get; set; }
        public string DwellingType {get; set;}
        public string LoanType {get; set;}
        public string WeekToDate { get; set; }
        public string MonthToDate { get; set; }
        public string RowType { get; set; }
        public string SortOrder { get; set; }
        public string SortField { get; set; }




}
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SignWordDoc.Data
{
    public class Policy
    {
        public string PolicyId { get; set; }
        public List<Customer> Insured { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public double SumInsured { get; set; }
        public double Premium { get; set; }
        public string SpecialConditions { get; set; }
        public string Exclusions { get; set; }
    }
}

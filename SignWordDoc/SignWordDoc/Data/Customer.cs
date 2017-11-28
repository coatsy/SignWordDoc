using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SignWordDoc.Data
{
    public class Customer
    {
        public string LegalName { get; set; }
        public string Address { get; set; }
        public DateTime DoB { get; set; }
        public string PreExistingConditions { get; set; }
    }
}

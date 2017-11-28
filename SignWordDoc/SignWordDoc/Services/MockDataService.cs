using SignWordDoc.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SignWordDoc.Services
{
    public class MockDataService : IDataService
    {
        public Policy GetPolicy(string policyId)
        {
            return new Policy()
            {
                PolicyId = policyId,
                StartDate = DateTime.Now,
                EndDate = DateTime.Now.AddDays(30),
                Insured = new List<Customer>()
                {
                    new Customer()
                    {
                        LegalName = "William H Gates III",
                        Address = "1 Microsoft Way\nREDMOND WA 98052\nUSA",
                        DoB = new DateTime(1955, 10, 28),
                        PreExistingConditions = "Nil"
                    },
                    new Customer()
                    {
                        LegalName = "Melinda Ann Gates",
                        Address = "1 Microsoft Way\nREDMOND WA 98052\nUSA",
                        DoB = new DateTime(1964, 08, 15),
                        PreExistingConditions = "Nil"
                    }
                },
                SumInsured = 25000.00d,
                Premium = 1250.00d,
                Exclusions = "Dental (other than emergency)\nTheraputic Massage",
                SpecialConditions = "None"
            };
        }

        public bool IsValidPolicyId(string policyId)
        {
            // for now, just return true
            return true;
        }
    }
}

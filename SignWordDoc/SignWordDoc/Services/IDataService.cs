using SignWordDoc.Data;

namespace SignWordDoc.Services
{
    public interface IDataService
    {
        Policy GetPolicy(string policyId);
        bool IsValidPolicyId(string policyId);
    }
}
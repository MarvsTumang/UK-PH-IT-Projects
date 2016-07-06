using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArrowGlobal
{
    public class Account : IEquatable<Account>
    {
        public string ArrowKey { get; set; }
        public int AccountId { get; set; }
        public string LoadDate { get; set; }

        public bool Equals(Account acct)
        {
            return acct.ArrowKey.Equals(ArrowKey) && acct.AccountId.Equals(AccountId);
        }
    }
}

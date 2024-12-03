using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryDesignPattern
{
    internal class PLatinum : ICreditCard
    {
        public int GetAnnualCharge()
        {
            return (2000);
        }

        public string GetCardType()
        {
            return ("PLatinum");
        }

        public int GetCreditLimit()
        {
            return (200000);
            
        }
    }
}

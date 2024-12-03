using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryDesignPattern
{
    internal class Titanium : ICreditCard
    {
        public int GetAnnualCharge()
        {
            return (1500);
        }

        public string GetCardType()
        {
            return ("Tinanuim");
        }

        public int GetCreditLimit()
        {
            return (10000);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 음료관리프로그램
{
    class Drink
    {
        public string Name { get; set; }
        public int Price { get; set; }

        public DateTime Dt { get; set; }


        public Drink()
        {

        }
        public Drink(string n, int p)
        {
            Name = n;
            Price = p;
            Dt = DateTime.Now;
        }



    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Parser
{
    public class Ad
    {
        private string name;
        private string price;

        public Ad(string name, string price)
        {
            this.name = name;
            this.price = price;
        }

        public string Name
        {
            get { return name; }
            set { value = this.name;  }
        }
        public string Price
        {
            get { return price; }
            set { value = this.price; }
        }
    }
}

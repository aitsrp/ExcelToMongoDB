using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CovertToFirebase
{
    class cServices
    {
        public string Header;
        public List<string> Service;
        
        public cServices()
        {
            Initialize();
        }

        public void Initialize()
        {
            Header = "";
            Service = new List<string>();
        }
    }
}

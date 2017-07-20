using System.Collections.Generic;

namespace CovertToFirebase
{
    public class cPerson
    {
        public string name;
        public string title;

        public cPerson()
        {
            name = "";
            title = "";
        }

        public cPerson(string name, string title)
        {
            this.name = name;
            this.title = title;
        }
    }
}
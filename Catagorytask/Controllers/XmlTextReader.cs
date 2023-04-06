using Catagorytask.Models;

namespace Catagorytask.Controllers
{
    internal class XmlTextReader : AllCatagory_Table1
    {
        private string fileLocation;

        public XmlTextReader(string fileLocation)
        {
            this.fileLocation = fileLocation;
        }
    }
}
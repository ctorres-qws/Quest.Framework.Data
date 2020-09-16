using Quest.ScanToPrint.Business;
using Quest.ScanToPrint.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace TaskExec
{
    class Program
    {
        static void Main(string[] args)
        {
            FacadeController controller = new FacadeController(new LocalOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ScanToPrint\ScanToPrint.mdb;Persist Security Info=False;"),
                new OnlineOLEDBPersistenceStrategiesFactory(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=V:\Quest.mdb;Persist Security Info=False;"));

            Console.WriteLine(string.Format("Updating local data...{0}",DateTime.Now.TimeOfDay.ToString()));
            controller.UpdateTagData();
            Console.WriteLine(string.Format("Local data update...{0}", DateTime.Now.TimeOfDay.ToString()));
        }
    }
}

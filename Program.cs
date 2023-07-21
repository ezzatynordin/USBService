using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace USBService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
           

            if (!System.Diagnostics.EventLog.SourceExists("USBService"))
            {
                System.Diagnostics.EventLog.CreateEventSource("USBService", "Application");
            }

            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new USBService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}

using System;
using System.ServiceProcess;

namespace USBService
{
    static class Program
    {
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

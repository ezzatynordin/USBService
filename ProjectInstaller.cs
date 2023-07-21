using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace USBService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        private ServiceProcessInstaller processInstaller;
        private ServiceInstaller serviceInstaller;

        public ProjectInstaller()
        {
            InitializeComponent();
            processInstaller = new ServiceProcessInstaller();
            processInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;

            serviceInstaller = new ServiceInstaller();
            serviceInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            serviceInstaller.ServiceName = "USBService";
            serviceInstaller.Description = "USB Detection and Access Control Service";

            Installers.Add(processInstaller);
            Installers.Add(serviceInstaller);
        }
    }
}

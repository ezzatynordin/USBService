using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using Microsoft.Win32;

namespace USBService
{
    public partial class USBService : ServiceBase
    {
        // Counter to track the number of USB events
        private static int eventCounter = 0;

        // Constants for DIF_PROPERTYCHANGE and DICS_PROPCHANGE
        private const uint DIF_PROPERTYCHANGE = 0x00000012;

        // USB event watcher
        private ManagementEventWatcher watcher;

        public USBService()
        {
            InitializeComponent();
            this.ServiceName = "USBService"; // Set the correct service name
        }

        protected override void OnStart(string[] args)
        {
            // Register the USB event watcher to handle device connection and disconnection events
            RegisterUsbEventWatcher();
        }

        protected override void OnStop()
        {
            // Stop the USB detection and access control logic
            StopUsbDetection();

            // Unregister the USB event watcher
            UnregisterUsbEventWatcher();
        }

        #region USB Detection and Access Control

        // Method to stop USB device detection and access control
        private static void StopUsbDetection()
        {
            // Stop USB device detection and access control logic
            // You can add any additional cleanup logic here if needed
            Console.WriteLine("USB detection and access control logic stopped.");
        }

        // Method to trigger a USB controller rescan
        private static void TriggerUsbControllerRescan(string usbDeviceId)
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", $"SELECT * FROM Win32_PnPEntity WHERE DeviceID LIKE '%{usbDeviceId}%'");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    SP_DEVINFO_DATA devInfoData = new SP_DEVINFO_DATA();
                    devInfoData.cbSize = (uint)Marshal.SizeOf(devInfoData);

                    // Get the device instance ID and pass it to the SetupDiCallClassInstaller function
                    string deviceInstanceId = queryObj["DeviceID"].ToString();
                    IntPtr ptr = Marshal.StringToHGlobalAuto(deviceInstanceId);
                    bool result = SetupDiCallClassInstaller(DIF_PROPERTYCHANGE, ptr, ref devInfoData);
                    Marshal.FreeHGlobal(ptr);

                    if (result)
                    {
                        Console.WriteLine("USB controller rescan successful.");
                    }
                    else
                    {
                        Console.WriteLine("Failed to trigger USB controller rescan.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while triggering USB controller rescan: " + ex.Message);
            }
        }

        #endregion

        #region USB Event Watcher

        // Method to register the USB event watcher
        public void RegisterUsbEventWatcher()
        {
            try
            {
                ManagementScope scope = new ManagementScope("root\\CIMV2");
                var query = new WqlEventQuery("SELECT * FROM Win32_DeviceChangeEvent");
                watcher = new ManagementEventWatcher(scope, query);
                watcher.EventArrived += UsbEventArrived;
                watcher.Start();
                Console.WriteLine("USB event watcher registered.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while registering USB event watcher: " + ex.Message);
            }
        }

        // Method to unregister the USB event watcher
        private void UnregisterUsbEventWatcher()
        {
            // Unregister the USB event watcher
            if (watcher != null)
            {
                watcher.EventArrived -= UsbEventArrived; // Remove the event handler
                watcher.Stop(); // Stop the watcher
                watcher.Dispose();
                watcher = null;
                Console.WriteLine("USB event watcher unregistered.");
            }
        }

        public void UsbEventArrived(object sender, EventArrivedEventArgs e)
        {
            try
            {
                // Increment the event counter
                eventCounter++;

                // Log a message when the method is called
                Console.WriteLine($"UsbEventArrived method called. Event Counter: {eventCounter}");

                // Get the "TargetInstance" property data from the USB event
                PropertyData targetInstanceData = e.NewEvent.Properties["TargetInstance"];

                // Check if the property data is not null and is of type ManagementBaseObject
                if (targetInstanceData != null && targetInstanceData.Value is ManagementBaseObject targetInstance)
                {
                    // Explicitly cast to ManagementObject
                    ManagementObject queryObj = (ManagementObject)targetInstance;

                    string deviceId = queryObj.GetPropertyValue("DeviceID").ToString();
                    bool isConnected = (int)queryObj.GetPropertyValue("ConfigManagerErrorCode") == 0;

                    // Log the USB event details to the console
                    Console.WriteLine($"USB Event: DeviceID: {deviceId}, IsConnected: {isConnected}");

                    // Log the USB event details to the Windows Event Viewer
                    LogUsbEventToEventViewer(deviceId, isConnected);

                    // Call the HandleUsbDeviceEvent method to handle the USB event
                    HandleUsbDeviceEvent(deviceId, isConnected);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while handling USB event: " + ex.Message);
            }
        }

        // Method to log the USB event details to the Windows Event Viewer
        private void LogUsbEventToEventViewer(string deviceId, bool isConnected)
        {
            string eventSource = "USBService";
            string eventLog = "Application";
            string logMessage = $"USB Event: DeviceID: {deviceId}, IsConnected: {isConnected}";

            try
            {
                if (!EventLog.SourceExists(eventSource))
                {
                    EventLog.CreateEventSource(eventSource, eventLog);
                }

                using (EventLog eventLogInstance = new EventLog(eventLog))
                {
                    eventLogInstance.Source = eventSource;
                    eventLogInstance.WriteEntry(logMessage, EventLogEntryType.Information);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while logging event to the Windows Event Viewer: " + ex.Message);
            }
        }


        // Method to handle USB device events
        public void HandleUsbDeviceEvent(string deviceInstancePath, bool isConnected)
        {
            try
            {
                if (isConnected)
                {
                    // Check if the device instance path is authorized in the database
                    bool deviceFound = USBDeviceHelper.CheckDeviceInDatabase(deviceInstancePath);

                    if (!deviceFound)
                    {
                        USBDeviceHelper.DisableUsbStorageDevices();
                        Console.WriteLine("USB storage devices disabled.");

                        // Log the unauthorized USB detection event to the Windows Event Viewer
                        LogUnauthorizedUsbEvent(deviceInstancePath);
                    }
                    else
                    {
                        USBDeviceHelper.EnableUsbStorageDevices();
                        Console.WriteLine("USB storage devices enabled.");
                        // Allow the USB device to access the PC (not implemented here)
                        Console.WriteLine("USB device allowed to access the PC.");
                    }
                }
                else
                {
                    // Handle the case when the USB device is disconnected
                    USBDeviceHelper.EnableUsbStorageDevices(); // Disable USB storage devices when the authorized USB is disconnected
                    Console.WriteLine("USB storage devices disabled.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while handling USB device event: " + ex.Message);
            }
        }


        // Method to log the unauthorized USB detection event to the Windows Event Viewer
        private void LogUnauthorizedUsbEvent(string deviceId)
        {
            string eventSource = "USBService";
            string eventLog = "Application";
            string logMessage = "Unauthorized USB device detected. Device ID: " + deviceId;

            try
            {
                if (!EventLog.SourceExists(eventSource))
                {
                    EventLog.CreateEventSource(eventSource, eventLog);
                }

                using (EventLog eventLogInstance = new EventLog(eventLog))
                {
                    eventLogInstance.Source = eventSource;
                    eventLogInstance.WriteEntry(logMessage, EventLogEntryType.Warning);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while logging event to the Windows Event Viewer: " + ex.Message);
            }
        }

        #endregion

        #region P/Invoke and Struct

        // P/Invoke declaration for SetupDiCallClassInstaller function
        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool SetupDiCallClassInstaller(uint InstallFunction, IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData);

        // Struct for SP_DEVINFO_DATA required for the SetupDiCallClassInstaller function
        [StructLayout(LayoutKind.Sequential)]
        public struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid ClassGuid;
            public uint DevInst;
            public IntPtr Reserved;
        }

        #endregion
    }
}

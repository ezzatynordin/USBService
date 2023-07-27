using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using Microsoft.Win32;

namespace USBService
{
    public partial class USBService : ServiceBase
    {
        private const string USBSTORKeyPath = @"SYSTEM\CurrentControlSet\Services\USBSTOR";
        private const string StartValueName = "Start";
        private const int DisableValue = 4; // 4 disables the USBSTOR service
        private const string ConnectionString = "Data Source=DESKTOP-3G61K1D\\SQLEXPRESS;Initial Catalog=USB;Integrated Security=True;";
        private const string AuthorizationQuery = "SELECT COUNT(*) FROM [USB].[dbo].[UsbDevices] WHERE DeviceID = @DeviceID";
        private bool shouldStopMonitoring = false; // Flag to indicate when to stop the USB event monitoring

        private ManagementEventWatcher usbWatcher;
        private int eventCounter = 0; // Added to track USB events
        private bool isUSBSTOREnabled = false; // Flag to track USBSTOR status
        private string lastConnectedDeviceId = string.Empty; // To track the last connected USB device
        private bool isConnected = false; // Class-level variable to track USB connection status
        private HashSet<string> connectedDevices = new HashSet<string>();
        public USBService()
        {
            InitializeComponent();
            ServiceName = "USBService";
        }

        protected override void OnStart(string[] args)
        {
            AlwaysDisableUSBSTOR();
            RegisterUsbEventWatcher();
            // Start monitoring USB events
            MonitorUsbEvents();
        }

        protected override void OnStop()
        {
            StopUsbDetection();
        }

        private void AlwaysDisableUSBSTOR()
        {
            isUSBSTOREnabled = false;
            SetUSBSTORStartValue(DisableValue);
            Console.WriteLine("USB storage devices disabled.");
        }

        private void EnableUSBSTOR()
        {
            isUSBSTOREnabled = true;
            SetUSBSTORStartValue(3); // 3 enables the USBSTOR service
            Console.WriteLine("USB storage devices enabled.");
        }

        private void EnableUSBSTORIfAuthorized(string deviceInstancePathId)
        {
            if (IsUSBAuthorized(deviceInstancePathId))
            {
                if (!isUSBSTOREnabled)
                {
                    EnableUSBSTOR();
                }
            }
            else
            {
                if (isUSBSTOREnabled)
                {
                    AlwaysDisableUSBSTOR();
                }
            }
        }

        private void SetUSBSTORStartValue(int value)
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(USBSTORKeyPath, true))
                {
                    if (key != null)
                    {
                        key.SetValue(StartValueName, value, RegistryValueKind.DWord);
                    }
                    else
                    {
                        Console.WriteLine("Error: Unable to access the USBSTOR registry key.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        private void RegisterUsbEventWatcher()
        {
            try
            {
                ManagementScope scope = new ManagementScope("root\\CIMV2");
                var query = new WqlEventQuery("SELECT * FROM Win32_DeviceChangeEvent");
                usbWatcher = new ManagementEventWatcher(scope, query);
                usbWatcher.EventArrived += UsbEventArrived;
                usbWatcher.Start();
                Console.WriteLine("USB event watcher registered.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while registering USB event watcher: " + ex.Message);
            }
        }

        private void UnregisterUsbEventWatcher()
        {
            if (usbWatcher != null)
            {
                usbWatcher.EventArrived -= UsbEventArrived;
                usbWatcher.Stop();
                usbWatcher.Dispose();
                usbWatcher = null;
                Console.WriteLine("USB event watcher unregistered.");
            }
        }

        private void StopUsbDetection()
        {
            // Stop monitoring USB events and exit the loop
            shouldStopMonitoring = true;
            Console.WriteLine("USB detection and access control logic stopped.");
        }

        private void UsbEventArrived(object sender, EventArrivedEventArgs e)
        {
            try
            {
                // Increment the event counter
                eventCounter++;

                // Log a message when the method is called
                Console.WriteLine($"UsbEventArrived method called. Event Counter: {eventCounter}");

                // Get the "TargetInstance" property data from the USB event
                var targetInstanceData = e.NewEvent.Properties["TargetInstance"];

                // Check if the property data is not null and is of type ManagementBaseObject
                if (targetInstanceData != null && targetInstanceData.Value is ManagementBaseObject targetInstance)
                {
                    // Explicitly cast to ManagementObject
                    var queryObj = (ManagementObject)targetInstance;

                    var deviceId = queryObj.GetPropertyValue("DeviceID").ToString();
                    var isConnected = (int)queryObj.GetPropertyValue("ConfigManagerErrorCode") == 0;

                    Console.WriteLine($"USB Event: DeviceID: {deviceId}, IsConnected: {isConnected}");

                    if (isConnected && deviceId != lastConnectedDeviceId)
                    {
                        connectedDevices.Add(deviceId);
                        EnableUSBSTORIfAuthorized(deviceId);
                        lastConnectedDeviceId = deviceId; // Update the last connected device
                    }
                    else if (!isConnected)
                    {
                        connectedDevices.Remove(lastConnectedDeviceId);
                        AlwaysDisableUSBSTOR();
                        lastConnectedDeviceId = string.Empty; // Reset the last connected device
                    }

                    isConnected = connectedDevices.Count > 0;
                
                RefreshRegistryEditor();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while handling USB event: " + ex.Message);
            }
        }

        private void MonitorUsbEvents()
        {
            while (!shouldStopMonitoring)
            {
                try { 
                // You can add some delay here to avoid excessive CPU usage
                Thread.Sleep(1000); // Wait for 1 second before checking for the next USB event

                // Check if USBSTOR should be enabled
                if (isConnected)
                {
                    EnableUSBSTORIfAuthorized(lastConnectedDeviceId);
                }
                else
                {
                    AlwaysDisableUSBSTOR();
                }
            }
                catch (Exception ex)
                {
                    // Log the exception for debugging purposes
                    Console.WriteLine("Error in MonitorUsbEvents: " + ex.Message);
                }

                // Introduce a delay to avoid excessive CPU usage
                Thread.Sleep(1000);
            }

        }


        private bool IsUSBAuthorized(string deviceInstancePathId)
        {
            bool isAuthorized = false;

            try
            {
                using (var connection = new SqlConnection(ConnectionString))
                {
                    connection.Open();
                    using (var command = new SqlCommand(AuthorizationQuery, connection))
                    {
                        command.Parameters.AddWithValue("@DeviceID", deviceInstancePathId);
                        var count = (int)command.ExecuteScalar();
                        isAuthorized = count > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return isAuthorized;
        }

        private static IntPtr FindRegistryEditorWindow()
        {
            const string className = "RegEdit_RegEdit";
            const string windowName = "Registry Editor";

            return FindWindow(className, windowName);
        }

        private void RefreshRegistryEditor()
        {
            var registryEditorHandle = FindRegistryEditorWindow();
            if (registryEditorHandle != IntPtr.Zero)
            {
                // Send the refresh message to the registry editor window
                SendMessage(registryEditorHandle, WM_COMMAND, (IntPtr)ID_REFRESH, IntPtr.Zero);
            }
            else
            {
                Console.WriteLine("Registry Editor window not found. Unable to refresh.");
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string className, string windowName);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        private const int WM_COMMAND = 0x0111;
        private const int ID_REFRESH = 0x3028;
    }
}

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Management;

namespace USBService
{
    public static class USBDeviceHelper
    {
        // Connection string for the database
        const string ConnectionString = "Data Source=F1-LAPTOP-MPC\\SQLEXPRESS;Initial Catalog=USB;Integrated Security=True;";

        // Method to get the USB mass storage device ID and instance path
        public static List<string> GetConnectedUsbDeviceInstancePaths()
        {
            List<string> deviceInstancePaths = new List<string>();

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_USBHub");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    if (queryObj["DeviceID"] != null)
                    {
                        string instancePath = queryObj["DeviceID"].ToString();
                        deviceInstancePaths.Add(instancePath);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while retrieving USB device instance paths: " + ex.Message);
            }

            return deviceInstancePaths;
        }


        // Method to check if the USB device is authorized in the database
        public static bool CheckDeviceInDatabase(string deviceId)
        {
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    connection.Open();
                    // SQL command to check if the device ID exists in the database
                    string sql = "SELECT COUNT(*) FROM [USB].[dbo].[UsbDevices] WHERE DeviceID = @DeviceID";

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@DeviceID", deviceId);
                        int count = (int)command.ExecuteScalar();

                        return count > 0;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error while checking the database: " + ex.Message);
                    return false; // Return false or handle the error accordingly
                }
            }
        }


        // Method to enable USB storage devices
        public static void EnableUsbStorageDevices()
        {
            try
            {
                // Set the registry key value to enable USB storage devices (set to 3)
                string keyPath = @"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR";
                Registry.SetValue(keyPath, "Start", 3, RegistryValueKind.DWord);
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while enabling USB storage devices: " + ex.Message);
            }
        }

        // Method to disable USB storage devices
        public static void DisableUsbStorageDevices()
        {
            try
            {
                // Set the registry key value to disable USB storage devices (set to 4)
                string keyPath = @"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR";
                Registry.SetValue(keyPath, "Start", 4, RegistryValueKind.DWord);
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while disabling USB storage devices: " + ex.Message);
            }
        }

        // Other methods...
    }
}

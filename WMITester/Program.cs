using System;
using System.Management;

namespace WMITester
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //NOTE: This program need Administrator Rights

                // I ended up using MSSerial_PortName from root\WMI, because Win32_SerialPort from root\CIMV2 won't detect my FTDI Serial Converter 
                // (but it's showed up properly in Device manager) and it's also firing multiple events per device disconnecting/connecting
                // more info here https://stackoverflow.com/questions/19840811/list-of-serialports-queried-using-wmi-differs-from-devicemanager
                ManagementScope managementScope = new ManagementScope(@"root\WMI");

                // Declaring queries for creation and deletion events of MSSerial_PortName instances
                // we want pull query every 1 second so we get small latency between 
                WqlEventQuery creatonQuery = new WqlEventQuery("__InstanceCreationEvent", TimeSpan.FromSeconds(1),
                                                        "TargetInstance ISA 'MSSerial_PortName'");

                WqlEventQuery deletionQuery = new WqlEventQuery("__InstanceDeletionEvent", TimeSpan.FromSeconds(1),
                                                        "TargetInstance ISA 'MSSerial_PortName'");

                // Initalize ManagementEventWatchers each for attach and detach
                ManagementEventWatcher deviceAttachWatcher = new ManagementEventWatcher(managementScope, creatonQuery);
                ManagementEventWatcher deviceDetachWatcher = new ManagementEventWatcher(managementScope, deletionQuery);
                Console.WriteLine("Waiting for an events...");

                // Event handlers
                deviceAttachWatcher.EventArrived += deviceAttachWatcher_EventArrived;
                deviceDetachWatcher.EventArrived += DeviceDetachWatcher_EventArrived;

                // Connect to our scope
                managementScope.Connect();

                // Start listening for events
                deviceAttachWatcher.Start();
                deviceDetachWatcher.Start();
            }
            catch( Exception ex)
            {
                Console.WriteLine(ex);
            }

            while (true) { }
        }
       
        private static void DeviceDetachWatcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
            // Get instance of MSSerial_PortName class 
            // More info about that class http://wutils.com/wmi/root/wmi/msserial_portname/
            var instance = e.NewEvent.GetPropertyValue("TargetInstance") as ManagementBaseObject;

            // Now we are ready to access properties of MSSerial_PortName itself
            Console.WriteLine($"Port: {instance.GetPropertyValue("PortName")} disconnected!");
        }

        private static void deviceAttachWatcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
            // Get instance of MSSerial_PortName
            var instance = e.NewEvent.GetPropertyValue("TargetInstance") as ManagementBaseObject;

            // Now we are ready to access properties of MSSerial_PortName itself
            Console.WriteLine($"Port: {instance.GetPropertyValue("PortName")} connected!");
        }
    }
}

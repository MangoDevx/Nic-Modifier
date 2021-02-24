using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;

namespace SetupNetAdapter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
            {
                Console.WriteLine("You need to run as admin! Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(-1);
            }
            new Program().DefineNic();
        }

        public void DefineNic()
        {
            Console.WriteLine("Hello, this program is to assist you in setting up a non-dhcp adapter.\n" +
                              "Created by Mango.\n");
            Thread.Sleep(1000);
            var networkAdapters = new Dictionary<int, NetworkInterface>();
            var nics = NetworkInterface.GetAllNetworkInterfaces();
            var ignoredNics = new[] { "loopback", "bluetooth", "vpn", "nord", "hamachi", "local", "psuedo" };
            for (var i = 0; i < nics.Length; i++)
            {
                var nic = nics[i];
                if (ignoredNics.Any(x => nic.Name.ToLower().Contains(x))) continue;
                networkAdapters.Add(i, nic);
            }

            Console.WriteLine("Please select the adapter you wish to modify by typing the corresponding number.");
            foreach (var (key, value) in networkAdapters)
            {
                Console.WriteLine($"{key}: {value.Name}");
            }

            NetworkInterface selectedNic;
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(">> ");
                Console.ResetColor();
                if (!int.TryParse(Console.ReadLine(), out var nicInput) || !networkAdapters.ContainsKey(nicInput))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Invalid input!");
                    Console.ResetColor();
                    continue;
                }

                selectedNic = networkAdapters[nicInput];
                break;
            }

            Console.WriteLine($"\nNic selected: {selectedNic.Name}");
            var selectedNicIpProperties = selectedNic.GetIPProperties();
            IPAddress nicIpv4 = default;
            foreach (var ip in selectedNicIpProperties.UnicastAddresses)
            {
                if (ip.Address.AddressFamily != AddressFamily.InterNetwork) continue;
                if (ip.Address.ToString().ToLowerInvariant().Contains(":"))
                {
                    Console.WriteLine("Please disable your IPv6 first. Press any key to continue, or escape to quit.");
                    if (Console.ReadKey().Key == ConsoleKey.Escape)
                        Environment.Exit(-1);
                }
                Console.WriteLine($"Nic IpV4: {ip.Address}");
                nicIpv4 = ip.Address;
            }

            if (nicIpv4 == null || nicIpv4.Equals(default))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Nic ipv4 failed to parse.");
                Console.ResetColor();
                Environment.Exit(-1);
            }

            var nicDefaultGateway = selectedNicIpProperties.GatewayAddresses[0].Address;
            if (nicDefaultGateway.ToString().ToLowerInvariant().Contains(":"))
            {
                Console.WriteLine("Please disable your IPv6 first. Press any key to continue, or escape to quit.");
                if (Console.ReadKey().Key == ConsoleKey.Escape)
                    Environment.Exit(-1);
            }

            Console.WriteLine($"Nic Default Gateway: {nicDefaultGateway}\n");
            GetNewValues(selectedNic, nicIpv4, nicDefaultGateway);
        }

        private void GetNewValues(NetworkInterface nic, IPAddress defaultIpv4, IPAddress defaultGatewayAddress)
        {
            Console.WriteLine("Ipv4 editing, you have 3 options.\n" +
                              "1) Leave blank and hit enter for default\n" +
                              "2) Type \"n\" to use the next available address\n" +
                              "3) Type a new local ipv4 manually");
            IPAddress localIp;
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(">> ");
                Console.ResetColor();
                var input = Console.ReadLine();

                if (string.IsNullOrEmpty(input))
                {
                    localIp = defaultIpv4;
                    break;
                }

                if (input.Contains("n"))
                {
                    var nextAvailableAddress = GetNextAvailableAddress(defaultIpv4);
                    if (nextAvailableAddress.ToString().Equals("0.0.0.0"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Unable to find next available address!");
                        Console.ResetColor();
                        Thread.Sleep(1);
                        continue;
                    }
                    localIp = nextAvailableAddress;
                    break;
                }

                if (!IPAddress.TryParse(input, out var localIpInput))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Invalid input!");
                    Console.ResetColor();
                    Thread.Sleep(1);
                    continue;
                }
                localIp = localIpInput;
                break;
            }
            Console.WriteLine();

            Console.WriteLine("Please enter the default gateway you wish to set, leave blank for original");
            IPAddress defaultGateway;
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(">> ");
                Console.ResetColor();
                var input = Console.ReadLine();
                if (string.IsNullOrEmpty(input))
                {
                    defaultGateway = defaultGatewayAddress;
                    break;
                }
                if (!IPAddress.TryParse(input, out var defaultGatewayInput))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Invalid input!");
                    Console.ResetColor();
                    Thread.Sleep(1);
                    continue;
                }
                defaultGateway = defaultGatewayInput;
                break;
            }
            Console.WriteLine();

            Console.WriteLine("Please enter the dns one you wish to set, leave blank for 8.8.8.8");
            IPAddress dnsOne;
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(">> ");
                Console.ResetColor();
                var input = Console.ReadLine();
                if (string.IsNullOrEmpty(input))
                {
                    dnsOne = IPAddress.Parse("8.8.8.8");
                    break;
                }
                if (!IPAddress.TryParse(input, out var dnsOneInput))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Invalid input!");
                    Console.ResetColor();
                    Thread.Sleep(1);
                    continue;
                }
                dnsOne = dnsOneInput;
                break;
            }
            Console.WriteLine();

            Console.WriteLine("Please enter the dns two you wish to set, leave blank for 8.8.4.4");
            IPAddress dnsTwo;
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(">> ");
                Console.ResetColor();
                var input = Console.ReadLine();
                if (string.IsNullOrEmpty(input))
                {
                    dnsTwo = IPAddress.Parse("8.8.4.4");
                    break;
                }
                if (!IPAddress.TryParse(input, out var dnsTwoInput))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Invalid input!");
                    Console.ResetColor();
                    Thread.Sleep(1);
                    continue;
                }
                dnsTwo = dnsTwoInput;
                break;
            }
            Console.WriteLine();

            Console.WriteLine($"Updating nic: {nic.Name}, IPV4: {localIp}, Default Gateway {defaultGateway}, Dns One: {dnsOne}, Dns Two: {dnsTwo}");
            Console.WriteLine("Press any key to confirm, press escape to cancel.");
            var keyInput = Console.ReadKey();
            if (keyInput.Key.Equals(ConsoleKey.Escape))
            {
                Console.WriteLine("EExiting in 3 seconds");
                Thread.Sleep(3000);
                Environment.Exit(-1);
            }

            RunEditProcesses(nic, localIp, defaultGateway, dnsOne, dnsTwo);
        }

        private IPAddress GetNextAvailableAddress(IPAddress ipv4)
        {
            var ping = new Ping();
            var stringIPv4 = ipv4.ToString();
            var ipv4Split = stringIPv4.Split('.');
            var lastValue = ipv4Split.Last();
            if (!int.TryParse(lastValue, out var editableLastValue))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to parse last ip value to integer.");
                Console.ResetColor();
                return default;
            }

            var pingIps = new List<IPAddress>();
            var newIpString = string.Empty;
            for (var j = 5; j < 255; j++)
            {
                for (var i = 0; i < ipv4Split.Length; i++)
                {
                    if (i == ipv4Split.Length - 1)
                    {
                        newIpString += $"{j}";
                        break;
                    }
                    newIpString += $"{ipv4Split[i]}.";
                }

                var parsedIp = IPAddress.Parse(newIpString);
                pingIps.Add(parsedIp);
                newIpString = string.Empty;
            }

            foreach (var ip in pingIps)
            {
                var response = ping.Send(ip, 2000);
                if (response.Status != IPStatus.Success)
                    return ip;
            }

            return default;
        }

        private void RunEditProcesses(NetworkInterface nic, IPAddress localIp, IPAddress defaultGateway, IPAddress dnsOne, IPAddress dnsTwo)
        {
            try
            {
                Console.WriteLine("Starting nic edits, if you lose connectivity please wait until it finishes.");
                var startInfo = new ProcessStartInfo("netsh", $"interface ip set address {nic.Name} static {localIp} 255.255.255.0 {defaultGateway}");
                var process = new Process { StartInfo = startInfo };
                process.Start();
                Console.WriteLine("netsh updated IPv4, and gateway");
            }
            catch (Exception e)
            {
                if (e.Message.Contains("(2)"))
                {
                    Console.WriteLine("You need to enable automatic ip. Check the guide for this issue.");
                    return;
                }
                Console.WriteLine($"ipv4 {e.Message}\n{e.StackTrace}");
            }

            try
            {
                Console.WriteLine("Starting DNS update");
                var result = SetDNS($"{dnsOne},{dnsTwo}", nic);

                if (result.Equals(0))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Failed to set DNS!");
                    Console.ResetColor();
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    Console.WriteLine("Exiting in 3 seconds.");
                    Thread.Sleep(3000);
                    Environment.Exit(-1);
                }

                Console.WriteLine("Updating nic finished. Press any key to exit.");
                Console.ReadKey();
                Console.WriteLine("Exiting in 3 seconds.");
                Thread.Sleep(3000);
                Environment.Exit(-1);
            }
            catch (Exception e)
            {
                if (e.Message.Contains("(2)"))
                {
                    Console.WriteLine("You need to enable automatic ip. Check the guide for this issue.");
                    return;
                }
                Console.WriteLine($"dns {e.Message}\n{e.StackTrace}");
                Console.ReadKey();
            }
        }

        private int SetDNS(string dns, NetworkInterface nic)
        {
            var objMC = new ManagementClass("Win32_NetworkAdapterConfiguration");
            var objMOC = objMC.GetInstances();

            foreach (ManagementObject objMo in objMOC)
            {
                if (!(bool)objMo["IPEnabled"]) continue;
                //Console.WriteLine(objMo["Caption"].ToString().ToLowerInvariant());
                //Console.WriteLine(nic.Description.ToLowerInvariant());
                if (!objMo["Caption"].ToString().ToLowerInvariant().Contains(nic.Description.ToLowerInvariant())) continue;
                var newDns = objMo.GetMethodParameters("SetDNSServerSearchOrder");
                newDns["DNSServerSearchOrder"] = dns.Split(',');
                var setDns = objMo.InvokeMethod("SetDNSServerSearchOrder", newDns, null);
                return 1;
            }

            return 0;
        }
    }
}
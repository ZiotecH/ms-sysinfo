using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Management;

namespace Get_SystemInformation
{
    class Program
    {
        static void Main(string[] args)
        {
            
            ManagementObjectSearcher mgmtOpSysSearcher = new ManagementObjectSearcher("SELECT CSName, Version, LastBootUpTime, OSArchitecture, Caption, TotalVisibleMemorySize FROM Win32_OperatingSystem");
            ManagementObjectSearcher mgmtObjProcSearcher = new ManagementObjectSearcher("SELECT Name, SocketDesignation, MaxClockSpeed, CurrentClockSpeed, NumberOfCores, NumberOfEnabledCore, NumberOfLogicalProcessors, Caption FROM Win32_Processor");
            ManagementObjectSearcher mgmtObjBBSearcher = new ManagementObjectSearcher("SELECT Manufacturer, Product, Version FROM Win32_BaseBoard");
            ManagementObjectSearcher mgmtObjVCSearcher = new ManagementObjectSearcher("SELECT AdapterCompatibility, VideoProcessor, CurrentHorizontalResolution, CurrentVerticalResolution, MinRefreshRate, MaxRefreshRate, DriverVersion FROM Win32_VideoController");
            bool doVerbose = false;

            if (args.Count() > 0)
            {
                for(int i = 0; i < args.Count(); i++)
                {
                    if (args[i] == "-v")
                    {
                        doVerbose = true;
                    }
                }
            }
            if (doVerbose) { Console.WriteLine("Verbose mode enabled."); }


            //Init Strings
            if (doVerbose) { Console.Write("Init Strings"); }
            string OSMachineName = null;
            string OSCaption = null;
            string OSVersion = null;
            string OSArchitecture = null;
            string OSPhysMemorySize = null;
            string OSLBUT = null;
            string ProcName = null;
            string ProcSocketDesgination = null;
            string ProcMaxClockSpeed = null;
            string ProcCurrentClockSpeed = null;
            string ProcNumberOfCores = null;
            string ProcNumberOfEnabledCores = null;
            string ProcNumberOfLogicalProcessors = null;
            string ProcCaption = null;
            string BBManufacturer = null;
            string BBProduct = null;
            string BBVersion = null;
            string LineDivider = "======================";
            List<string> VCAdapterCompatibility = new List<string>();
            List<string> VCVideoProcessor = new List<string>();
            List<string> VCCurrentHorizontalResolution = new List<string>();
            List<string> VCCurrentVerticalResolution = new List<string>();
            List<string> VCMinRefreshRate = new List<string>();
            List<string> VCMaxRefreshRate = new List<string>();
            List<string> VCDriverVersion = new List<string>();

            //Declare file path:
            if (doVerbose){ Console.Write(" - Success\nDeclare file path"); }
            string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Desktop";
            string FilePath = DesktopPath + @"\System_Info.txt";

            //Manage File Operations
            if (doVerbose) { Console.Write(" - Success\nManage File Operations"); }
            ProcessStartInfo OpenFolder = new ProcessStartInfo
            {
                Arguments = DesktopPath,
                FileName = "explorer.exe"
            };
            ProcessStartInfo OpenTextFile = new ProcessStartInfo
            {
                Arguments = FilePath,
                FileName = "Notepad.exe"
            };

            //CollectInfo
            if (doVerbose) { Console.WriteLine(" - Success\nCollectInfo"); }
            Console.WriteLine("Collecting information.. [1/4]");
            ManagementObjectCollection colOpSys = mgmtOpSysSearcher.Get();
            //CSName, Version, LastBootUpTime, OSArchitecture, Caption
            foreach (ManagementObject tempObj in colOpSys)
            {
                OSMachineName = tempObj["CSName"].ToString();
                OSCaption = tempObj["Caption"].ToString();
                OSVersion = tempObj["Version"].ToString();
                OSArchitecture = tempObj["OSArchitecture"].ToString();
                OSLBUT = tempObj["LastBootUpTime"].ToString();
                OSPhysMemorySize = (Convert.ToDouble(Convert.ToInt64(tempObj["TotalVisibleMemorySize"].ToString())) / (1024 * 1024)).ToString("F2") + "GB";
            }
            Console.WriteLine("Collecting information.. [2/4]");
            ManagementObjectCollection colProc = mgmtObjProcSearcher.Get();
            //SocketDesignation, MaxClockSpeed, CurrentClockSpeed, NumberOfCores, NumberOfEnabledCore, NumberOfLogicalProcessors, Caption
            foreach (ManagementObject tempObj in colProc)
            {
                ProcName = tempObj["Name"].ToString();
                ProcSocketDesgination = tempObj["SocketDesignation"].ToString();
                ProcMaxClockSpeed = tempObj["MaxClockSpeed"].ToString();
                ProcCurrentClockSpeed = tempObj["CurrentClockSpeed"].ToString();
                ProcNumberOfCores = tempObj["NumberOfCores"].ToString();
                ProcNumberOfEnabledCores = tempObj["NumberOfEnabledCore"].ToString();
                ProcNumberOfLogicalProcessors = tempObj["NumberOfLogicalProcessors"].ToString();
                ProcCaption = tempObj["Caption"].ToString();
            }
            Console.WriteLine("Collecting information.. [3/4]");
            ManagementObjectCollection colBB = mgmtObjBBSearcher.Get();
            //Manufacturer, Product, Version
            foreach (ManagementObject tempObj in colBB)
            {
                BBManufacturer = tempObj["Manufacturer"].ToString();
                BBProduct = tempObj["Product"].ToString();
                BBVersion = tempObj["Version"].ToString();
            }
            Console.WriteLine("Collecting information.. [4/4]");
            ManagementObjectCollection colVC = mgmtObjVCSearcher.Get();
            //AdapterCompatibility, VideoProcessor, CurrentHorizontalResolution, CurrentVerticalResolution, MinRefreshRate, MaxRefreshRate, DriverVersion
            if (colVC.Count > 1)
            {
                if (doVerbose) { Console.WriteLine("Number of GPUs detected: {0}", colVC.Count); }
                foreach(ManagementBaseObject x in colVC)
                {
                    foreach(var a in x.Properties)
                    {
                        switch (a.Name.ToString())
                        {
                            case ("AdapterCompatibility"):
                                if (a.Value != null)
                                {
                                    VCAdapterCompatibility.Add(a.Value.ToString());
                                }
                                else
                                {
                                    VCAdapterCompatibility.Add("0");
                                }
                                break;
                            case ("CurrentHorizontalResolution"):
                                if (a.Value != null)
                                {
                                    VCCurrentHorizontalResolution.Add(a.Value.ToString());
                                }
                                else
                                {
                                    VCCurrentHorizontalResolution.Add("0");
                                }
                                break;
                            case ("CurrentVerticalResolution"):
                                if (a.Value != null)
                                {
                                    VCCurrentVerticalResolution.Add(a.Value.ToString());
                                }
                                else
                                {
                                    VCCurrentVerticalResolution.Add("0");
                                }
                                break;
                            case ("DriverVersion"):
                                if (a.Value != null)
                                {
                                    VCDriverVersion.Add(a.Value.ToString());
                                }
                                else
                                {
                                    VCDriverVersion.Add("0");
                                }
                                break;
                            case ("MaxRefreshRate"):
                                if (a.Value != null)
                                {
                                    VCMaxRefreshRate.Add(a.Value.ToString());
                                }
                                else
                                {
                                    VCMaxRefreshRate.Add("0");
                                }
                                break;
                            case ("MinRefreshRate"):
                                if (a.Value != null)
                                {
                                    VCMinRefreshRate.Add(a.Value.ToString());
                                }
                                else
                                {
                                    VCMinRefreshRate.Add("0");
                                }
                                break;
                            case ("VideoProcessor"):
                                if (a.Value != null)
                                {
                                    VCVideoProcessor.Add(a.Value.ToString());
                                }
                                else
                                {
                                    VCVideoProcessor.Add("0");
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            else
            {
                foreach (ManagementObject tempObj in colVC)
                {
                    VCAdapterCompatibility.Add(tempObj["AdapterCompatibility"].ToString());
                    VCVideoProcessor.Add(tempObj["VideoProcessor"].ToString());
                    VCCurrentHorizontalResolution.Add(tempObj["CurrentHorizontalResolution"].ToString());
                    VCCurrentVerticalResolution.Add(tempObj["CurrentVerticalResolution"].ToString());
                    VCMinRefreshRate.Add(tempObj["MinRefreshRate"].ToString());
                    VCMaxRefreshRate.Add(tempObj["MaxRefreshRate"].ToString());
                    VCDriverVersion.Add(tempObj["DriverVersion"].ToString());
                }
            }
            //Manage time info
            if (doVerbose) { Console.Write("CollectInfo - Success\nManage time info"); }

            if (doVerbose) { Console.Write(" - Success\nLBUT_Year"); }
            int LBUT_Year = Convert.ToInt32(OSLBUT.Substring(0, 4));

            if (doVerbose) { Console.Write(" - Success\nLBUT_Month"); }
            int LBUT_Month = Convert.ToInt32(OSLBUT.Substring(4, 2));

            if (doVerbose) { Console.Write(" - Success\nLBUT_Day"); }
            int LBUT_Day = Convert.ToInt32(OSLBUT.Substring(6, 2));

            if (doVerbose) { Console.Write(" - Success\nLBUT_Hour"); }
            int LBUT_Hour = Convert.ToInt32(OSLBUT.Substring(8, 2));

            if (doVerbose) { Console.Write(" - Success\nLBUT_Minute"); }
            int LBUT_Minute = Convert.ToInt32(OSLBUT.Substring(10, 2));

            if (doVerbose) { Console.Write(" - Success\nLBUT_Second"); }
            int LBUT_Second = Convert.ToInt32(OSLBUT.Substring(12, 2));

            if (doVerbose) { Console.Write(" - Success\nLBUT_DateTime"); }
            DateTime LBUT_DateTime = new DateTime(LBUT_Year, LBUT_Month, LBUT_Day, LBUT_Hour, LBUT_Minute, LBUT_Second);

            if (doVerbose) { Console.Write(" - Success\nCurrentDateTime"); }
            DateTime CurrentDateTime = DateTime.Now;

            if (doVerbose) { Console.Write(" - Success\nSystemUptime"); }
            TimeSpan SystemUptime = TimeSpan.FromTicks(CurrentDateTime.Ticks - LBUT_DateTime.Ticks);

            if (doVerbose) { Console.WriteLine(" - Success"); }

            Console.Write("\n");
            //Process info to file.
            if (doVerbose) { Console.WriteLine("Process info to file."); }

            Console.WriteLine("Finished gathering system info, writing to file..");

            //Remove before launch
            //System.Environment.Exit(0);

            //FileCheckStuff
            if (doVerbose) { Console.Write("Checking if file exists"); }
            if (File.Exists(FilePath))
            {
                Console.Write("\n\n\n");
                Console.WriteLine("Warning, {0} already exists.", FilePath);
                Console.Write("Press enter to confirm overwriting.");
                Console.ReadLine();
            }
            if (doVerbose) { Console.Write("\nFile check - Success\nWriting to file"); }
            using (StreamWriter sw = File.CreateText(FilePath))
            {
                sw.WriteLine("MachineName          : {0}", OSMachineName);
                sw.WriteLine("OSEdition            : {0}", OSCaption);
                sw.WriteLine("OSVersion            : {0}", OSVersion);
                sw.WriteLine("OSArchitecture       : {0}", OSArchitecture);
                sw.WriteLine("OSPhysicalMemory     : {0}", OSPhysMemorySize);
                sw.WriteLine(LineDivider);
                sw.WriteLine("ProcessorName        : {0}", ProcName);
                sw.WriteLine("ProcessorSocket      : {0}", ProcSocketDesgination);
                sw.WriteLine("MaxSpeed             : {0}mhz", ProcMaxClockSpeed);
                sw.WriteLine("CurrentSpeed         : {0}mhz", ProcCurrentClockSpeed);
                sw.WriteLine("ProcessorCores       : {0}", ProcNumberOfCores);
                sw.WriteLine("EnabledCores         : {0}", ProcNumberOfEnabledCores);
                sw.WriteLine("ProcessorThreads     : {0}", ProcNumberOfLogicalProcessors);
                sw.WriteLine("ProcessorVersion     : {0}", ProcCaption);
                sw.WriteLine(LineDivider);
                sw.WriteLine("MoboManufacturer     : {0}", BBManufacturer);
                sw.WriteLine("MoboName             : {0}", BBProduct);
                sw.WriteLine("MoboVersion          : {0}", BBVersion);
                sw.WriteLine(LineDivider);
                sw.WriteLine("Number of GPUs       : {0}", colVC.Count);
                for(var i = 0; i < VCAdapterCompatibility.Count; i++)
                {
                    sw.WriteLine(LineDivider);
                    sw.WriteLine("GPUManufacturer[{0}]   : {1}", i, VCAdapterCompatibility[i]);
                    sw.WriteLine("GPUName[{0}]           : {1}", i, VCVideoProcessor[i]);
                    sw.WriteLine("GPUResolution[{0}]     : {1}x{2}", i, VCCurrentHorizontalResolution[i], VCCurrentVerticalResolution[i]);
                    sw.WriteLine("GPURefreshRateMin[{0}] : {1}hz", i, VCMinRefreshRate[i]);
                    sw.WriteLine("GPURefreshRateMax[{0}] : {1}hz", i, VCMaxRefreshRate[i]);
                    sw.WriteLine("GPUDriverVersion[{0}]  : {1}", i, VCDriverVersion[i]);
                }
                sw.WriteLine(LineDivider);
                sw.WriteLine("LastBoot             : {0}", LBUT_DateTime);
                if(SystemUptime.Days > 0)
                {
                    sw.WriteLine("Uptime               : {0:N0} days, {1} hours, {2} minutes, {3} seconds", SystemUptime.Days, SystemUptime.Hours, SystemUptime.Minutes, SystemUptime.Seconds);
                }else if(SystemUptime.Hours > 0)
                {
                    sw.WriteLine("Uptime               : {0} hours, {1} minutes, {2} seconds", SystemUptime.Hours, SystemUptime.Minutes, SystemUptime.Seconds);
                }
                else if(SystemUptime.Minutes > 0)
                {
                    sw.WriteLine("Uptime               : {0} minutes, {1} seconds", SystemUptime.Minutes, SystemUptime.Seconds);
                }
                else
                {
                    sw.WriteLine("Uptime               : {0} seconds", SystemUptime.Seconds);
                }
            }
            if (doVerbose) { Console.WriteLine(" - Success"); } 

            //Finish execution
            Console.WriteLine("\n");
            Console.WriteLine("Finished writing to {0}.", FilePath);
            Console.Write("Press enter to open the file.");
            Console.ReadLine();
            Process.Start(OpenTextFile);
        }
    }
}

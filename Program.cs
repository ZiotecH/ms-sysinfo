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
            
            ManagementObjectSearcher mgmtOpSysSearcher = new ManagementObjectSearcher("SELECT CSName, Version, LastBootUpTime, OSArchitecture, Caption, TotalVisibleMemorySize, EncryptionLevel FROM Win32_OperatingSystem");
            ManagementObjectSearcher mgmtObjProcSearcher = new ManagementObjectSearcher("SELECT Name, SocketDesignation, MaxClockSpeed, CurrentClockSpeed, NumberOfCores, NumberOfEnabledCore, NumberOfLogicalProcessors, Caption FROM Win32_Processor");
            ManagementObjectSearcher mgmtObjBBSearcher = new ManagementObjectSearcher("SELECT Manufacturer, Product, SerialNumber, Version FROM Win32_BaseBoard");
            ManagementObjectSearcher mgmtObjVCSearcher = new ManagementObjectSearcher("SELECT AdapterCompatibility, VideoProcessor, CurrentHorizontalResolution, CurrentVerticalResolution, MinRefreshRate, MaxRefreshRate, DriverVersion FROM Win32_VideoController");
            ManagementObjectSearcher mgmtObjLDisk = new ManagementObjectSearcher("SELECT DriveType, DeviceID, VolumeSerialNumber, Name, SystemName, VolumeName, Description, FileSystem, Size, FreeSpace, Status, StatusInfo, SupportsFileBasedCompression FROM Win32_LogicalDisk");
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
            string OSEncryptionLevel = null;
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
            string BBOther = null;
            string BBSerialNumber = null;
            string LineDivider = "======================";
            List<string> VCAdapterCompatibility = new List<string>();
            List<string> VCVideoProcessor = new List<string>();
            List<string> VCCurrentHorizontalResolution = new List<string>();
            List<string> VCCurrentVerticalResolution = new List<string>();
            List<string> VCMinRefreshRate = new List<string>();
            List<string> VCMaxRefreshRate = new List<string>();
            List<string> VCDriverVersion = new List<string>();
            //List of drives?
            List<string> LDDriveType = new List<string>();
            List<string> LDDeviceID = new List<string>();
            List<string> LDVolumeSN = new List<string>();
            List<string> LDName = new List<string>();
            List<string> LDSName = new List<string>();
            List<string> LDVName = new List<string>();
            List<string> LDDescription = new List<string>();
            List<string> LDFS = new List<string>();
            List<UInt64> LDSize = new List<UInt64>();
            List<UInt64> LDFSpace = new List<UInt64>();
            float LDTotalPerc = 0.0f;
            List<string> LDStatus = new List<string>();
            List<string> LDStatusInfo = new List<string>();
            List<string> LDSFBC = new List<string>();
            //???

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
            Console.WriteLine("Collecting information.. [1/5]");
            ManagementObjectCollection colOpSys = mgmtOpSysSearcher.Get();
            //CSName, Version, LastBootUpTime, OSArchitecture, Caption
            foreach (ManagementObject tempObj in colOpSys)
            {
                OSMachineName = tempObj["CSName"].ToString();
                OSCaption = tempObj["Caption"].ToString();
                OSVersion = tempObj["Version"].ToString();
                OSArchitecture = tempObj["OSArchitecture"].ToString();
                OSEncryptionLevel = tempObj["EncryptionLevel"].ToString();
                OSLBUT = tempObj["LastBootUpTime"].ToString();
                OSPhysMemorySize = displayInBytes((UInt64)tempObj["TotalVisibleMemorySize"]*1000);
            }
            Console.WriteLine("Collecting information.. [2/5]");
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
            Console.WriteLine("Collecting information.. [3/5]");
            ManagementObjectCollection colBB = mgmtObjBBSearcher.Get();
            //Manufacturer, Product, Version, OtherIdentifyingInfo, SerialNumber
            foreach (ManagementObject tempObj in colBB)
            {
                BBManufacturer = tempObj["Manufacturer"].ToString();
                BBProduct = tempObj["Product"].ToString();
                BBVersion = tempObj["Version"].ToString();
                try {
                    if (doVerbose)
                    {
                        Console.WriteLine("BBOther - Success");
                    }
                    BBOther = tempObj["OtherIdentifyingInfo"].ToString();
                }
                catch {
                    if (doVerbose)
                    {
                        Console.WriteLine("BBOther - Failure");
                    }
                    BBOther = "N/A";
                }
                BBSerialNumber = tempObj["SerialNumber"].ToString();
            }
            Console.WriteLine("Collecting information.. [4/5]");
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
            Console.WriteLine("Collecting information.. [5/5]");
            ManagementObjectCollection colLD = mgmtObjLDisk.Get();
            //DriveType, DeviceID, VolumeSerialNumber, Name, SystemName, VolumeName, Description, FileSystem, Size, FreeSpace, Status, StatusInfo, SupportsFileBasedCompression
            if(colLD.Count > 1)
            {
                if(doVerbose) { Console.WriteLine("Number of LogicalDisks detected: {0}", colLD.Count); }
                foreach(ManagementBaseObject x in colLD)
                {
                    foreach(var a in x.Properties)
                    {
                        switch (a.Name.ToString())
                        {
                            case ("DriveType"):
                                if (a.Value != null)
                                {
                                    LDDriveType.Add(driveTypeTable((UInt32)a.Value));
                                }
                                else
                                {
                                    LDDriveType.Add(driveTypeTable(7));
                                }
                                break;
                            case ("DeviceID"):
                                if (a.Value != null)
                                {
                                    LDDeviceID.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDDeviceID.Add("N/A");
                                }
                                break;
                            case ("VolumeSerialNumber"):
                                if (a.Value != null)
                                {
                                    LDVolumeSN.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDVolumeSN.Add("N/A");
                                }
                                break;
                            case ("Name"):
                                if (a.Value != null)
                                {
                                    LDName.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDName.Add("N/A");
                                }
                                break;
                            case ("SystemName"):
                                if (a.Value != null)
                                {
                                    LDSName.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDSName.Add("N/A");
                                }
                                break;
                            case ("VolumeName"):
                                if (a.Value != null)
                                {
                                    LDVName.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDVName.Add("N/A");
                                }
                                break;
                            case ("Description"):
                                if (a.Value != null)
                                {
                                    LDDescription.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDDescription.Add("N/A");
                                }
                                break;
                            case ("FileSystem"):
                                if (a.Value != null)
                                {
                                    LDFS.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDFS.Add("N/A");
                                }
                                break;
                            case ("Size"):
                                if (a.Value != null)
                                {
                                    LDSize.Add((ulong)a.Value);
                                }
                                else
                                {
                                    LDSize.Add(1);
                                }
                                break;
                            case ("FreeSpace"):
                                if (a.Value != null)
                                {
                                    LDFSpace.Add((ulong)a.Value);
                                }
                                else
                                {
                                    LDFSpace.Add(1);
                                }
                                break;
                            case ("Status"):
                                if (a.Value != null)
                                {
                                    LDStatus.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDStatus.Add("N/A");
                                }
                                break;
                            case ("StatusInfo"):
                                if (a.Value != null)
                                {
                                    LDStatusInfo.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDStatusInfo.Add("N/A");
                                }
                                break;
                            case ("SupportsFileBasedCompression"):
                                if (a.Value != null)
                                {
                                    LDSFBC.Add(a.Value.ToString());
                                }
                                else
                                {
                                    LDSFBC.Add("N/A");
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            } else
            {
                foreach (ManagementObject tempObj in colLD)
                {
                LDDriveType.Add(tempObj["DriveType"].ToString());
                LDDeviceID.Add(tempObj["DeviceID"].ToString());
                LDVolumeSN.Add(tempObj["VolumeSerialNumber"].ToString());
                LDName.Add(tempObj["Name"].ToString());
                LDSName.Add(tempObj["SystemName"].ToString());
                LDVName.Add(tempObj["VolumeName"].ToString());
                LDDescription.Add(tempObj["Description"].ToString());
                LDFS.Add(tempObj["FileSystem"].ToString());
                LDSize.Add((ulong)tempObj["Size"]);
                LDFSpace.Add((ulong)tempObj["FreeSpace"]);
                LDStatus.Add(tempObj["Status"].ToString());
                LDStatusInfo.Add(tempObj["StatusInfo"].ToString());
                LDSFBC.Add(tempObj["SupportsFileBasedCompression"].ToString());
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
                sw.WriteLine("EncryptionLevel      : {0}-bit", OSEncryptionLevel);
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
                sw.WriteLine("OtherInfo            : {0}", BBOther);
                sw.WriteLine("SerialNumber         : {0}", BBSerialNumber);
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
                sw.WriteLine("Uptime               : {0}",hmsCalc(((Int64)CurrentDateTime.Ticks - (Int64)LBUT_DateTime.Ticks) / (Int64)TimeSpan.TicksPerSecond));
                sw.WriteLine(LineDivider);
                sw.WriteLine("Number of LogicalDisks: {0}", colLD.Count);
                for (var i = 0; i < LDDriveType.Count; i++)
                {
                    sw.WriteLine(LineDivider);
                    sw.WriteLine("LDDriveType[{0}]       : {1}", i, LDDriveType[i]);
                    sw.WriteLine("LDDeviceID[{0}]        : {1}", i, LDDeviceID[i]);
                    sw.WriteLine("LDVolumeSN[{0}]        : {1}", i, LDVolumeSN[i]);
                    //sw.WriteLine("LDName[{0}]            : {1}", i, LDName[i]);
                    //sw.WriteLine("LDSName[{0}]           : {1}", i, LDSName[i]);
                    sw.WriteLine("LDVName[{0}]           : {1}", i, LDVName[i]);
                    sw.WriteLine("LDDescription[{0}]     : {1}", i, LDDescription[i]);
                    sw.WriteLine("LDFS[{0}]              : {1}", i, LDFS[i]);
                    sw.WriteLine("LDSize[{0}]            : {1}", i, displayInBytes(LDSize[i]));
                    sw.WriteLine("LDFreeSpace[{0}]       : {1}", i, displayInBytes(LDFSpace[i]));
                    LDTotalPerc = (float)(((float)LDFSpace[i] / (float)LDSize[i]));
                    //sw.WriteLine("LDFree%[{0}]          : {1:0%}", i, LDTotalPerc);
                    sw.WriteLine("LDUsed%[{0}]           : {1:0%}", i, (float)(1 - LDTotalPerc));
                    sw.WriteLine("LDStatus[{0}]          : {1}", i, LDStatus[i]);
                    sw.WriteLine("LDStatusInfo[{0}]      : {1}", i, LDStatusInfo[i]);
                    sw.WriteLine("LDSFBC[{0}]            : {1}", i, LDSFBC[i]);
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
        static string displayInBytes(UInt64 bytes)
        {
            string sizeString = null;
            string[] suffix = { "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB" };
            Int32 suffixIndex = 0;
            float byteSafe = (float)bytes;
            while (byteSafe > 1.00f)
            {
                if(byteSafe/1024f < 1)
                {
                    break;
                }
                else
                {
                    byteSafe = byteSafe / 1024f;
                    suffixIndex++;
                }
            }
            

            sizeString = string.Format("{0:N}{1}", byteSafe,suffix[suffixIndex]);

            //Console.WriteLine("{0} => {1:N} | {2}", bytes, byteSafe, sizeString);

            return sizeString;
        }
        static string driveTypeTable(UInt32 type)
        {
            if(type > 6)
            {
                type = 7;
            }
            string[] typeString = {"Unknown", "No Root Directory", "Removable Disk", "Logical Disk", "Network Drive", "Compact Disc", "RAM Disk", "N/A" };
            return typeString[type];
        }
        static string hmsCalc(Int64 seconds)
        {
            string hmsString = null;
            Int64 timeVal = seconds;
            Int64[] hmsTable = { 0, 0, 0, 0, 0 };
            Int64[] divTable = { (3600 * 24 * 365), (3600 * 24), 3600, 60, 1 };
            Int64 tempVal = 0;
            if (timeVal >= divTable[0])
            {
                tempVal = (timeVal / divTable[0]);
                hmsTable[0] = tempVal;
                timeVal -= (tempVal * divTable[0]);

            }

            if (timeVal >= divTable[1])
            {
                tempVal = (timeVal / divTable[1]);
                hmsTable[1] = tempVal;
                timeVal -= (tempVal * divTable[1]);

            }

            if (timeVal >= divTable[2])
            {
                tempVal = (timeVal / divTable[2]);
                hmsTable[2] = tempVal;
                timeVal -= (tempVal * divTable[2]);

            }

            if (timeVal >= divTable[3])
            {
                tempVal = (timeVal / divTable[3]);
                hmsTable[3] = tempVal;
                timeVal -= (tempVal * divTable[3]);

            }

            hmsTable[4] = timeVal;

            if (hmsTable[0] >= 1) { hmsString += String.Format("{0} years, ",hmsTable[0]); }
            if (hmsTable[1] >= 1) { hmsString += String.Format("{0} days, ", hmsTable[1]); }
            if (hmsTable[2] >= 1) { hmsString += String.Format("{0} hours, ", hmsTable[2]); }
            if (hmsTable[3] >= 1) { hmsString += String.Format("{0} minutes, ", hmsTable[3]); }
            hmsString += String.Format("{0} seconds", hmsTable[4]);
            return hmsString;
        }
    }
}

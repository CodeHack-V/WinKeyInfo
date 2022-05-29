using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Management;

namespace WinKeyInfo
{
    internal class Helper
    {
        public static MainWindow mainWindow = null;
        public static LType CL = LType.CN;//当前语言字符串
        public static void MainWindowLoad(MainWindow MThis)
        {
            mainWindow = MThis;
            MThis.Loaded += (object sender, RoutedEventArgs e) =>
            {

            };
            MThis.CN.Click += LClick;
            MThis.EN.Click += LClick;
            Task.Run(() =>
            {
                //获取CPU信息 Win32_processor
                ManagementClass CPUmanagementClass = new ManagementClass("Win32_processor");
                ManagementObjectCollection CPUmanagementBaseObjects = CPUmanagementClass.GetInstances();
                string CPUName = "", ClockSpeed = "", Role = "", NumberOfCores = "", ThreadCount = "";
                foreach (ManagementObject Management in CPUmanagementBaseObjects)
                {
                    foreach (PropertyData Pd in Management.Properties)
                    {
                        if (Pd.Name == "Name")
                        {

                            CPUName = Pd.Value + "".Trim();
                        }
                        if (Pd.Name == "CurrentClockSpeed")
                        {
                            double CurrentClockSpeed = Math.Round(Convert.ToDouble(Pd.Value) / 1000, 2);
                            ClockSpeed = "@" + CurrentClockSpeed.ToString() + "GHz";
                        }
                        if (Pd.Name == "Role")
                        {
                            Role = Pd.Value + "";
                        }
                        if (Pd.Name == "NumberOfCores")
                        {
                            NumberOfCores = Pd.Value + "";
                        }
                        if (Pd.Name == "ThreadCount")
                        {
                            ThreadCount = Pd.Value + "";
                        }
                        if (CPUName != "" && ClockSpeed != "" && Role != "" && ThreadCount != "")
                        {
                            MThis.CPUValue.Dispatcher.Invoke(() =>
                            {
                                MThis.CPUValue.Text = CPUName + "   " + Role + ClockSpeed + "   " + "Cores:" + NumberOfCores + "   " + "Threads:" + ThreadCount;
                            });
                            break;
                        }
                    }
                }
                //获取CPU信息
                //获取显卡信息 Win32_DisplayConfiguration
                ManagementClass GPUmanagementClass = new ManagementClass("Win32_DisplayConfiguration");
                ManagementObjectCollection GPUmanagementBaseObjects = GPUmanagementClass.GetInstances();
                string GPUInfo = "";
                foreach (ManagementObject Management in GPUmanagementBaseObjects)
                {
                    foreach (PropertyData Pd in Management.Properties)
                    {
                        if (Pd.Name == "Caption")
                        {
                            GPUInfo += Pd.Value + "" + "+";
                        }
                    }
                }
                GPUInfo = GPUInfo.Remove(GPUInfo.Length - 1);
                mainWindow.GPUValue.Dispatcher.Invoke(() =>
                {
                    mainWindow.GPUValue.Text = GPUInfo;
                });
                //获取显卡信息
                //获取内存条信息 Win32_PhysicalMemory
                string RAMInfo = "", RAMSpeed = "", RAMManufacturer = "", RAMPartNumber = "", RAMDeviceLocator = "";
                int RAMGB = 0, TRAMGB = 0;
                ManagementClass RAMmanagementClass = new ManagementClass("Win32_PhysicalMemory");
                ManagementObjectCollection RAMmanagementBaseObjects = RAMmanagementClass.GetInstances();
                foreach (ManagementObject Management in RAMmanagementBaseObjects)
                {
                    foreach (PropertyData Pd in Management.Properties)
                    {
                        if (Pd.Name == "Speed")
                        {
                            RAMSpeed = Pd.Value + "";
                        }
                        if (Pd.Name == "Capacity")//Speed Name Capacity Manufacturer PartNumber DeviceLocator
                        {
                            TRAMGB = Convert.ToInt32(Convert.ToInt64(Pd.Value) / 1024 / 1024 / 1024);
                            RAMGB += TRAMGB;
                        }
                        if (Pd.Name == "Manufacturer")
                        {
                            RAMManufacturer = Pd.Value + "";
                        }
                        if (Pd.Name == "PartNumber")
                        {
                            RAMPartNumber = Pd.Value + "";
                        }
                        if (Pd.Name == "DeviceLocator")
                        {
                            RAMDeviceLocator = Pd.Value + "";
                        }
                        if (RAMSpeed != "" && TRAMGB != 0 && RAMManufacturer != "" && RAMPartNumber != "" && RAMDeviceLocator != "")
                        {
                            RAMInfo += RAMManufacturer + "   " + "   " + RAMPartNumber + "   " + RAMDeviceLocator + "   " + TRAMGB + "GB" + "   " + RAMSpeed + "   " + "+" + "   ";
                            RAMSpeed = ""; RAMManufacturer = ""; RAMPartNumber = ""; RAMDeviceLocator = "";
                            TRAMGB = 0;
                        }
                    }
                }
                RAMInfo = RAMInfo.Trim();
                RAMInfo = RAMInfo.Remove(RAMInfo.Length - 1);
                mainWindow.RAMValue.Dispatcher.Invoke(() =>
                {
                    mainWindow.RAMValue.Text = RAMInfo + "=   " + RAMGB + "GB";
                });
                //获取内存条信息
                //获取硬盘 Win32_DiskDrive
                ManagementClass DiskDrivemanagementClass = new ManagementClass("Win32_DiskDrive");
                ManagementObjectCollection DiskDrivemanagementBaseObjects = DiskDrivemanagementClass.GetInstances();
                string DiskDriveInfo = "";
                foreach (ManagementObject Management in DiskDrivemanagementBaseObjects)
                {
                    foreach (PropertyData Pd in Management.Properties)
                    {
                        if (Pd.Name == "Caption")
                        {
                            DiskDriveInfo += Pd.Value + "+";
                        }
                    }
                }
                DiskDriveInfo = DiskDriveInfo.Remove(DiskDriveInfo.Length - 1);
                mainWindow.DiskDriveValue.Dispatcher.Invoke(() =>
                {
                    mainWindow.DiskDriveValue.Text = DiskDriveInfo;
                });
                //获取硬盘
                //获取主板信息 Win32_BaseBoard
                string BaseBoardManufacturer = "", BaseBoardProduct = "";
                ManagementClass BaseBoardmanagementClass = new ManagementClass("Win32_BaseBoard");
                ManagementObjectCollection BaseBoardmanagementBaseObjects = BaseBoardmanagementClass.GetInstances();
                foreach (ManagementObject Management in BaseBoardmanagementBaseObjects)
                {
                    foreach (PropertyData Pd in Management.Properties)
                    {
                        if (Pd.Name == "Manufacturer")//Manufacturer Product
                        {
                            BaseBoardManufacturer = Pd.Value + "";
                        }
                        if (Pd.Name == "Product")//Manufacturer Product
                        {
                            BaseBoardProduct = Pd.Value + "";
                        }
                        if (BaseBoardManufacturer != "" && BaseBoardProduct != "")
                        {
                            mainWindow.BaseBoardValue.Dispatcher.Invoke(() =>
                            {
                                mainWindow.BaseBoardValue.Text = BaseBoardManufacturer + "   " + BaseBoardProduct;
                            });
                        }
                    }
                }
                //获取主板信息
                //获取声卡 Win32_SoundDevice
                ManagementClass SoundDevicemanagementClass = new ManagementClass("Win32_SoundDevice");
                ManagementObjectCollection SoundDevicemanagementBaseObjects = SoundDevicemanagementClass.GetInstances();
                string SoundDeviceInfo = "";
                foreach (ManagementObject Management in SoundDevicemanagementBaseObjects)
                {
                    foreach (PropertyData Pd in Management.Properties)
                    {
                        if (Pd.Name == "Caption")
                        {
                            SoundDeviceInfo += Pd.Value + "+";
                        }
                    }
                }
                SoundDeviceInfo = SoundDeviceInfo.Remove(SoundDeviceInfo.Length - 1);
                mainWindow.SoundDeviceValue.Dispatcher.Invoke(() =>
                {
                    mainWindow.SoundDeviceValue.Text = SoundDeviceInfo;
                });
                //获取声卡
                //获取显示器 Win32_DesktopMonitor
                ManagementClass DesktopMonitormanagementClass = new ManagementClass("Win32_DesktopMonitor");
                ManagementObjectCollection DesktopMonitormanagementBaseObjects = DesktopMonitormanagementClass.GetInstances();
                string DesktopMonitorInfo = "";
                foreach (ManagementObject Management in DesktopMonitormanagementBaseObjects)
                {
                    foreach (PropertyData Pd in Management.Properties)
                    {
                        if (Pd.Name == "PNPDeviceID")
                        {
                            DesktopMonitorInfo += (Pd.Value + "").Split("\\")[1].Trim() + "+";
                        }
                    }
                }
                DesktopMonitorInfo = DesktopMonitorInfo.Trim();
                DesktopMonitorInfo = DesktopMonitorInfo.Remove(DesktopMonitorInfo.Length - 1);
                mainWindow.DisplayValue.Dispatcher.Invoke(() =>
                {
                    mainWindow.DisplayValue.Text = DesktopMonitorInfo;
                });
                //获取显示器
                //获取电源 Win32_Battery
                ManagementClass BatterymanagementClass = new ManagementClass("Win32_Battery");
                ManagementObjectCollection BatterymanagementBaseObjects = BatterymanagementClass.GetInstances();
                string BatteryCaption = "", BatteryDeviceID = "", BatteryEstimatedChargeRemaining = "", BatteryInfo = "";
                foreach (ManagementObject Management in BatterymanagementBaseObjects)
                {
                    foreach (PropertyData Pd in Management.Properties)
                    {
                        if (Pd.Name == "Caption")
                        {
                            BatteryCaption = Pd.Value + "";
                        }
                        if (Pd.Name == "DeviceID")
                        {
                            BatteryDeviceID = Pd.Value + "";
                        }
                        if (Pd.Name == "EstimatedChargeRemaining")
                        {
                            BatteryEstimatedChargeRemaining = Pd.Value + "";
                        }
                    }
                }
                BatteryInfo = BatteryCaption + "   " + BatteryDeviceID + "   (" + BatteryEstimatedChargeRemaining + "%)";
                if (BatteryDeviceID.Trim() == "")
                {
                    BatteryInfo = "Not Found";
                }
                mainWindow.BatteryValue.Dispatcher.Invoke(() =>
                {
                    mainWindow.BatteryValue.Text = BatteryInfo;
                });
                //获取电源
            });
        }
        public static void LClick(object sender, RoutedEventArgs e)
        {
            string CmenuItem = ((MenuItem)sender).Name;
            LType TL = CmenuItem == "CN" ? LType.CN : LType.EN;
            if (CL != TL)
            {
                if (TL == LType.CN)
                {
                    CL = LType.CN;
                    mainWindow.CPUTitle.Content = "处理器:";
                    mainWindow.GPUTitle.Content = "显卡:";
                    mainWindow.RAMTitle.Content = "内存条:";
                    mainWindow.DiskDriveTitle.Content = "硬盘:";
                    mainWindow.BaseBoardTitle.Content = "主板:";
                    mainWindow.SoundDeviceTitle.Content = "声卡:";
                    mainWindow.DisplayTitle.Content = "显示器:";
                    mainWindow.BatteryTitle.Content = "电池:";
                }
                else
                {
                    CL = LType.EN;
                    mainWindow.CPUTitle.Content = "CPU:";
                    mainWindow.GPUTitle.Content = "GPU:";
                    mainWindow.RAMTitle.Content = "RAM:";
                    mainWindow.DiskDriveTitle.Content = "DiskDrive:";
                    mainWindow.BaseBoardTitle.Content = "BaseBoard:";
                    mainWindow.SoundDeviceTitle.Content = "SoundDevice:";
                    mainWindow.DisplayTitle.Content = "Display:";
                    mainWindow.BatteryTitle.Content = "Battery:";
                }
            }
        }
    }
    public enum LType
    {
        CN, EN
    }
}

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Print_Limit.Model;
using System.Windows.Forms;
using System.Printing;
using System.ServiceProcess;
using System.Security.Principal;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.Graph.SecurityNamespace;
using System.Runtime.Remoting.Contexts;
using System.Diagnostics;
using System.Net;

namespace Print_Limit
{
    public class EventWatcherAsync
    {

        Print_LimitEntities db = new Print_LimitEntities();


        // viết hàm trả về số lượng giới hạn bản in trong tháng (int ID_NhanVien)


        public int GetSoLuongBanInTrongThang(int thang, int nam, string bios)
        {
            var ID = db.DM_NhanVien.Where(_ => _.Bios_MayTinh == bios).FirstOrDefault().ID_NhanVien;
            var data = db.NV_BanIn.Where(_ => _.ID_NhanVien == ID && _.ThoiGianPrint.Value.Year == nam && _.ThoiGianPrint.Value.Month == thang);
            if (data.Count() != 0)
            {
                return (int)data.Sum(_ => _.TongSoTrangDaIn);
            }
            return 0;
        }

        public int GioiHanBanInNhanVien(string bios)
        {
            Print_LimitEntities db1 = new Print_LimitEntities();
            var checkLoaiNhanVien = db1.DM_NhanVien.Where(_ => _.Bios_MayTinh == bios).FirstOrDefault();

            if (checkLoaiNhanVien.KeyNhomNhanVien == "DEFALT")
            {
                int SoLuongBanIn = (int)checkLoaiNhanVien.SoLuongBanInTrongThang;
                return SoLuongBanIn;
            }
            else
            {
                var SoLuongBanIn = (int)db1.DM_NhomNhanVien.Where(_ => _.KeyNhomNhanVien == checkLoaiNhanVien.KeyNhomNhanVien).Select(_ => _.SoLuongBanInTrongThang).Single();
                return SoLuongBanIn;
            }
        }

        private int isOnline(string Name)
        {
            ManagementScope scope = new ManagementScope(@"\root\cimv2");
            scope.Connect();

            // Select Printers from WMI Object Collections
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
            string printerName = "";
            try
            {
                foreach (ManagementObject printer in searcher.Get())
                {
                    printerName = printer["Name"].ToString().ToLower();
                    if (printerName.Equals(Name))
                    {
                        var t = Int32.Parse(printer["DetectedErrorState"].ToString());

                        return t;
                    }
                    else
                    {
                        return -1;
                    }
                }
            }
            catch
            {
                return -1;
            }
            return -1;
        }

        void CleanPrinterQueue(string printerName)
        {
            using (var ps = new PrintServer())
            {
                using (var pq = new PrintQueue(ps, printerName, PrintSystemDesiredAccess.UsePrinter))
                {
                    foreach (var job in pq.GetPrintJobInfoCollection())
                        job.Cancel();
                }
            }
        }


        // Fix lỗi xuất hiện 2 event printing job làm cho thống kê 2 lần trên web.
        // Nếu bắt được sự kiện printing job lần đầu thì lần sau sẽ bỏ qua tránh việc lưu xuống DB 2 lần gây nên thống kê sai.
        private bool eventHandled = false;

        private void WmiEventHandler(object sender, EventArrivedEventArgs e)
        {
            string ip4Address = "";
            ManagementObjectSearcher NetworkSearcher = new ManagementObjectSearcher("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'");
            ManagementObjectCollection collectNetWork = NetworkSearcher.Get();
            foreach (ManagementObject obj in collectNetWork)
            {
                string[] arrIPAddress = (string[])(obj["IPAddress"]);
                ip4Address = arrIPAddress[0];
            }

            var Caption = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Caption"];
            var Description = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Description"];
            var InstallDate = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["InstallDate"];
            var Name = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Name"].ToString();
            var Status = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Status"].ToString();
            var ElapsedTime = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["ElapsedTime"];
            var JobStatus = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["JobStatus"].ToString();
            var Notify = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Notify"];
            var Owner = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Owner"];
            var Priority = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Priority"];
            var StartTime = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["StartTime"];
            var TimeSubmitted = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["TimeSubmitted"];
            var UntilTime = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["UntilTime"];
            var Color = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Color"];
            var DataType = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["DataType"];
            var Document = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Document"].ToString();
            var DriverName = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["DriverName"].ToString();
            var HostPrintQueue = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["HostPrintQueue"];
            var JobId = Int32.Parse(((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["JobId"].ToString());
            var PagesPrinted = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["PagesPrinted"];
            var PaperLength = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["PaperLength"];
            var PaperSize = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["PaperSize"].ToString();
            var PaperWidth = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["PaperWidth"];
            var Parameters = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Parameters"];
            var PrintProcessor = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["PrintProcessor"];
            var Size = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["Size"];
            var TotalPages = Int32.Parse(((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["TotalPages"].ToString());
            var StatusMask = ((ManagementBaseObject)e.NewEvent.Properties["TargetInstance"].Value)["StatusMask"];
            var TenMayIn = Name.Split(',')[0].ToString().ToLower();
            var SoMayIn = DriverName.ToString().ToLower();
            var t = isOnline(TenMayIn);
            var statusText = "";
            switch (t)
            {
                case 0:
                    statusText = " Unknown";
                    break;
                case 1:
                    statusText = " Other";
                    break;
                case 2:
                    statusText = " No Error";
                    break;
                case 3:
                    statusText = " Low Paper";
                    break;
                case 4:
                    statusText = " No Paper";
                    break;
                case 5:
                    statusText = " Low Toner";
                    break;
                case 6:
                    statusText = " No Toner";
                    break;
                case 7:
                    statusText = " Door Open";
                    break;
                case 8:
                    statusText = " Jammed";
                    break;
                case 9:
                    statusText = " Offline ";
                    break;
                case 10:
                    statusText = " Service Requested";
                    break;
                case 11:
                    statusText = " Output Bin Full";
                    break;
                default:
                    // code block
                    break;
            }
            var flag = true;

            if (flag)
            {
                if (t != 9)
                {
                    if (TenMayIn.ToLower() != "microsoft print to pdf".ToLower() && 
                        TenMayIn.ToLower() != "fax" && 
                        TenMayIn.ToLower() != "microsoft xps document writer" && 
                        TenMayIn.ToLower() != "onenote for windows 10")
                    {
                        var printjob = new NV_PrintTam();
                        printjob.JobID = JobId;
                        printjob.JobStatus = JobStatus != "" ? JobStatus.ToString() : "";
                        printjob.TenMayIn = Name.Split(',')[0].ToLower();
                        printjob.NgayIn = DateTime.Now;
                        printjob.TongSoTrang = TotalPages;
                        printjob.TenTaiLieu = Document;
                        printjob.Bios_MayTinh = ip4Address;
                        printjob.StatusPrint = Status;
                        printjob.PaperSize = PaperSize;
                        printjob.SoMayIn = DriverName.ToLower();
                        printjob.TrangThaiText = "Đã In Thành Công";
                        if (JobStatus == "Printing" && Status == "OK")
                        {
                            db.NV_PrintTam.Add(printjob);
                            db.SaveChanges();
                        }
                    }
                }
                else
                {
                    var printjob = new NV_PrintTam();
                    printjob.JobID = JobId;
                    printjob.JobStatus = (JobStatus != null) ? JobStatus.ToString() : "";
                    printjob.TenMayIn = Name.Split(',')[0].ToLower();
                    printjob.NgayIn = DateTime.Now;
                    printjob.TongSoTrang = TotalPages;
                    printjob.TenTaiLieu = Document;
                    printjob.Bios_MayTinh = ip4Address;
                    printjob.StatusPrint = Status;
                    printjob.SoMayIn = DriverName.ToLower();
                    printjob.TrangThaiText = statusText;
                    printjob.PaperSize = PaperSize;
                    db.NV_PrintTam.Add(printjob);
                    db.SaveChanges();
                    CancelPrintJob(Name.Split(',')[0], JobId);
                }
            }
        }


        public void CancelPrintJob(string printerName, int printJobID)
        {
            // Variable declarations.
            bool isActionPerformed = false;
            string searchQuery;
            String jobName;
            char[] splitArr;
            int prntJobID;
            ManagementObjectSearcher searchPrintJobs;
            ManagementObjectCollection prntJobCollection;

            try
            {
                ServiceController controller = new ServiceController("Spooler");

                controller.Stop();

            }
            catch (Exception ex)
            {
                var contend = "";
                using (StreamReader readtext = new StreamReader("D:\\test.txt"))
                {
                    string line;
                    // Read and display lines from the file until the end of
                    // the file is reached.
                    while ((line = readtext.ReadLine()) != null)
                    {
                        contend += $"{line}\n";
                    }
                }
                using (StreamWriter writetext = new StreamWriter("D:\\test.txt"))
                {
                    writetext.WriteLine($"{contend}\n Lỗi không chặn được");
                }
            }

            try
            {


                LocalPrintServer localPrintServer = new LocalPrintServer(PrintSystemDesiredAccess.AdministratePrinter);
                PrintQueue printQueue = localPrintServer.GetPrintQueue(printerName);

                if (printQueue.NumberOfJobs > 0)
                {
                    printQueue.Purge();
                }



                CleanPrinterQueue(printerName);
                // Query to get all the queued printer jobs.
                searchQuery = "SELECT * FROM Win32_PrintJob";
                // Create an object using the above query.
                searchPrintJobs = new ManagementObjectSearcher(searchQuery);
                // Fire the query to get the collection of the printer jobs.
                prntJobCollection = searchPrintJobs.Get();

                // Look for the job you want to delete/cancel.
                foreach (ManagementObject prntJob in prntJobCollection)
                {
                    jobName = prntJob.Properties["Name"].Value.ToString();
                    // Job name would be of the format [Printer name], [Job ID]
                    splitArr = new char[1];
                    splitArr[0] = Convert.ToChar(",");
                    // Get the job ID.
                    prntJobID = Convert.ToInt32(jobName.Split(splitArr)[1]);
                    // If the Job Id equals the input job Id, then cancel the job.
                    if (prntJobID == printJobID)
                    {
                        // Performs a action similar to the cancel
                        // operation of windows print console
                        prntJob.Delete();
                        isActionPerformed = true;

                    }
                }
            }
            catch (Exception sysException)
            {
                var contend = "";
                using (StreamReader readtext = new StreamReader("D:\\test.txt"))
                {
                    string line;
                    // Read and display lines from the file until the end of
                    // the file is reached.
                    while ((line = readtext.ReadLine()) != null)
                    {
                        contend += $"{line}\n";
                    }
                }
                using (StreamWriter writetext = new StreamWriter("D:\\test.txt"))
                {
                    writetext.WriteLine($"{contend}\n Lỗi không chặn được");
                }
            }

        }

        public EventWatcherAsync()
        {
            try
            {
                string ComputerName = "localhost";
                string WmiQuery;
                ManagementEventWatcher Watcher;
                ManagementScope Scope;


                if (!ComputerName.Equals("localhost", StringComparison.OrdinalIgnoreCase))
                {
                    ConnectionOptions Conn = new ConnectionOptions();
                    Conn.Username = "";
                    Conn.Password = "";
                    Conn.Authority = "ntlmdomain:DOMAIN";
                    Scope = new ManagementScope(String.Format("\\\\{0}\\root\\CIMV2", ComputerName), Conn);
                }
                else
                    Scope = new ManagementScope(String.Format("\\\\{0}\\root\\CIMV2", ComputerName), null);
                Scope.Connect();

                WmiQuery = "Select * From __InstanceOperationEvent Within 0.1 " +
                "Where TargetInstance ISA 'Win32_PrintJob' ";

                Watcher = new ManagementEventWatcher(Scope, new EventQuery(WmiQuery));
                Watcher.EventArrived += new EventArrivedEventHandler(this.WmiEventHandler);
                Watcher.Start();
                Console.Read();
                Watcher.Stop();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception {0} Trace {1}", e.Message, e.StackTrace);
            }

        }
        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        readonly static IntPtr handle = GetConsoleWindow();
        [DllImport("kernel32.dll")] static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public static void Hide()
        {
            ShowWindow(handle, SW_HIDE); //hide the console
        }
        public static void Show()
        {
            ShowWindow(handle, SW_SHOW); //show the console
        }

        public void enforceAdminPrivilegesWorkaround()
        {
            RegistryKey rk;
            string registryPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\";

            try
            {
                if (Environment.Is64BitOperatingSystem)
                {
                    rk = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
                }
                else
                {
                    rk = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry32);
                }

                rk = rk.OpenSubKey(registryPath, true);
            }
            catch (System.Security.SecurityException ex)
            {
                MessageBox.Show("Please run as administrator");
                System.Environment.Exit(1);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public static void Main(string[] args)
        {

            //Hide();
            //RegistryKey rkApp = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            //rkApp.SetValue("MyAPP", Assembly.GetExecutingAssembly().Location);
            Print_LimitEntities db2 = new Print_LimitEntities();

            var ip4Address = "";
            try
            {

                ManagementObjectSearcher NetworkSearcher = new ManagementObjectSearcher("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'");
                ManagementObjectCollection collectNetWork = NetworkSearcher.Get();
                foreach (ManagementObject obj in collectNetWork)
                {
                    string[] arrIPAddress = (string[])(obj["IPAddress"]);

                    ip4Address = arrIPAddress[0];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            var checkLoaiNhanVien = db2.DM_NhanVien.Where(_ => _.Bios_MayTinh == ip4Address).FirstOrDefault();
            if (checkLoaiNhanVien != null)
            {
                try
                {
                    Process process = new Process();
                    process.StartInfo.FileName = "cmd.exe";
                    process.StartInfo.CreateNoWindow = true;
                    process.StartInfo.RedirectStandardInput = true;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.UseShellExecute = false;
                    process.Start();
                    process.StandardInput.WriteLine("powershell -windowstyle hidden -command \"Start-Process cmd -ArgumentList '/s,/c,net stop spooler & DEL /F /S /Q %systemroot%\\System32\\spool\\PRINTERS\\* & net start spooler' -Verb runAs\"");
                    process.StandardInput.Flush();
                    process.StandardInput.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ex: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("Bạn vui lòng tạo IP máy tính trong Danh mục máy tính");
                Application.Exit();
            }

            EventWatcherAsync eventWatcher = new EventWatcherAsync();
        }
    }
}
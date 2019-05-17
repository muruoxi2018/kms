using System;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace KMS
{
    public partial class Form1 : Form
    {
        public static string kmsserver = "kms.muruoxi.com";
        public Form1()
        {
            InitializeComponent();
        }
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern bool SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);

        bool beginMove = false;//初始化鼠标位置
        int currentXPosition;
        int currentYPosition;

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) => Process.Start("https://www.muruoxi.com");


        private void Button1_Click(object sender, EventArgs e)
        {
            TopMost = false;
            try
            {
                #region 激活Windows
                button1.Text = "正在激活中，请等待。。。";
                button1.Enabled = false;
                RegistryKey localMachine = Registry.LocalMachine;
                string text = localMachine.OpenSubKey("software", true).OpenSubKey("microsoft", true).OpenSubKey("Windows NT", true).OpenSubKey("CurrentVersion", true).GetValue("ProductName").ToString();
                localMachine.Close();
                string key = "";
                switch (text)
                {
                    case "Windows 10 Pro for Workstations":
                        key = "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J";
                        break;
                    case "Windows 10 Enterprise G":
                        key = "YYVX9-NTFWV-6MDM3-9PT4T-4M68B";
                        break;
                    case "Windows 10 Enterprise":
                        key = "NPPR9-FWDCX-D2C8J-H872K-2YT43";
                        break;
                    case "Windows 10 Pro":
                        key = "W269N-WFGWX-YVC9B-4J6C9-T83GX";
                        break;
                    case "Windows 10 Education":
                        key = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2";
                        break;
                    case "Windows 10 Professional Education":
                        key = "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y";
                        break;
                    case "Windows 10 Home":
                        key = "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99";
                        break;
                    case "Windows 10 Home China":
                        key = "PVMJN-6DFY6-9CCP6-7BKTT-D3WVR";
                        break;
                    case "Windows 10 Home Single Language":
                        key = "7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH";
                        break;
                    case "Windows 10 Enterprise 2016 LTSB":
                        key = "DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ";
                        break;
                    case "Windows 10 Enterprise 2015 LTSB":
                        key = "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4";
                        break;
                    case "Windows 8.1 Pro":
                        key = "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9";
                        break;
                    case "Windows 8.1 Enterprise":
                        key = "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7";
                        break;
                    case "Windows 8.1 Pro with Media Center":
                        key = "789NJ-TQK6T-6XTH8-J39CJ-J8D3P";
                        break;
                    case "Windows 8.1":
                        key = "M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK";
                        break;
                    case "Windows 8.1 China":
                        key = "NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3";
                        break;
                    case "Windows 8.1 Single Language":
                        key = "BB6NG-PQ82V-VRDPW-8XVD2-V8P66";
                        break;
                    case "Windows 8 Pro":
                        key = "NG4HW-VH26C-733KW-K6F98-J8CK4";
                        break;
                    case "Windows 8 Enterprise":
                        key = "32JNW-9KQ84-P47T8-D8GGY-CWCK7";
                        break;
                    case "Windows 8":
                        key = "BN3D2-R7TKB-3YPBD-8DRP2-27GG4";
                        break;
                    case "Windows 8 Pro with Media Center":
                        key = "GNBB8-YVD74-QJHX6-27H4K-8QHDG";
                        break;
                    case "Windows 8 China":
                        key = "4K36P-JN4VD-GDC6V-KDT89-DYFKP";
                        break;
                    case "Windows 8 Single Language":
                        key = "2WN2H-YGCQR-KFX6K-CD6TF-84YXQ";
                        break;
                    case "Windows 7 Professional":
                        key = "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4";
                        break;
                    case "Windows 7 Enterprise":
                        key = "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH";
                        break;
                    case "Windows Server 2016 Datacenter":
                        key = "CB7KF-BWN84-R7R2Y-793K2-8XDDG";
                        break;
                    case "Windows Server 2016 Standard":
                        key = "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY";
                        break;
                    case "Windows Server 2016 Essentials":
                        key = "JCKRF-N37P4-C2D82-9YXRT-4M63B";
                        break;
                    case "Windows Server 2012 R2 Datacenter":
                        key = "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9";
                        break;
                    case "Windows Server 2012 R2 Standard":
                        key = "D2N9P-3P6X9-2R39C-7RTCD-MDVJX";
                        break;
                    case "Windows Server 2012 Datacenter":
                        key = "48HP8-DN98B-MYWDG-T2DCC-8W83P";
                        break;
                    case "Windows Server 2012 Standard":
                        key = "XC9B7-NBPP2-83J2H-RHMBY-92BT4";
                        break;
                    case "Windows Server 2008 R2 Datacenter":
                        key = "74YFP-3QFB3-KQT8W-PMXWJ-7M648";
                        break;
                    case "Windows Server 2008 R2 Enterprise":
                        key = "489J6-VHDMP-X63PK-3K798-CPX3Y";
                        break;
                    case "Windows Server 2008 R2 Standard":
                        key = "YC6KT-GKW9T-YTKYR-T4X34-R7VHC";
                        break;
                    case "Windows Server 2008 R2 Web":
                        key = "6TPJF-RBVHG-WBW2R-86QPH-6RTM4";
                        break;
                    case "Windows Server 2008 Datacenter":
                        key = "7M67G-PC374-GR742-YH8V4-TCBY3";
                        break;
                    case "Windows Server 2008 Enterprise":
                        key = "YQGMW-MPWTJ-34KDK-48M3W-X4Q6V";
                        break;
                    case "Windows Server 2008 Standard":
                        key = "TM24T-X9RMF-VWXK6-X8JC9-BFGM2";
                        break;
                    case "Windows Server 2008 Web":
                        key = "WYR28-R7TFJ-3X2YQ-YCY4H-M249D";
                        break;
                    case "Windows Server 2008 HPC":
                        key = "RCTX3-KWVHP-BR6TB-RB6DM-6X7HP";
                        break;
                    case "Windows Server 2008 R2 HPC":
                        key = "TT8MH-CG224-D3D7Q-498W2-9QCTX";
                        break;
                    case "Windows 7 Starter":
                    case "Windows 7 Home Basic":
                    case "Windows 7 Home Premium":
                    case "Windows 7 Ultimate":
                        break;
                    default:
                        if (text.Contains("Windows 7"))
                        {
                        }
                        else if (text.Contains("Preview") || text.Contains("Evaluation"))
                        {
                            MessageBox.Show("你当前的系统为测试评估版本，请更换为正式版后激活", "by：慕若曦");
                        }
                        break;
                }
                Process process = new Process();
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk {0}", key);
                process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms {0}",kmsserver);
                process.StandardInput.WriteLine(@"cscript %windir%\\System32\\slmgr.vbs /ato >{0}",Application.StartupPath+@"\Windows.log");
                process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                process.StandardInput.WriteLine("exit");
                process.WaitForExit();
                process.Close();
                #endregion
                #region 激活Office
                object ProductReleaseIds = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration", "ProductReleaseIds", null);
                object value = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration", "InstallationPath", null);
                object value2 = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration", "ClientVersionToReport", null);
                object value3 = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\16.0\\Common\\InstallRoot", "Path", null);
                object value4 = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\16.0\\Common\\InstallRoot", "Path", null);
                object value5 = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\15.0\\Common\\InstallRoot", "Path", null);
                object value6 = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\15.0\\Common\\InstallRoot", "Path", null);
                object value7 = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\14.0\\Common\\InstallRoot", "Path", null);
                object value8 = Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\14.0\\Common\\InstallRoot", "Path", null);
                string officekey = "";
                string patch = "";
                if (value != null && value2.ToString().Substring(0, 2).ToString() == "16")   
                {
                    string str = value.ToString() + "\\office16";
                    string str2 = value.ToString() + "\\root\\Licenses16";
                    string str3 = "cscript \"" + str + "\\ospp.vbs\" /inslic:";
                    string text1 = "";
                    string text2 = "";
                    string text3 = "";
                    if (value.ToString().Contains("ProPlus2019Retail"))
                    {
                        text1 = str3 + "\"" + str2 + "\\ProPlus2019VL_KMS_Client_AE-ppd.xrm-ms\"";
                        text2 = str3 + "\"" + str2 + "\\ProPlus2019VL_KMS_Client-ul.xrm-ms\"";
                        text3 = str3 + "\"" + str2 + "\\ProPlus2019VL_KMS_Client-ul-oob.xrm-ms\"";
                    }
                    else
                    {
                        text1 = str3 + "\"" + str2 + "\\ProPlusVL_KMS_Client-ppd.xrm-ms\"";
                        text2 = str3 + "\"" + str2 + "\\ProPlusVL_KMS_Client-ul.xrm-ms\"";
                        text3 = str3 + "\"" + str2 + "\\ProPlusVL_KMS_Client-ul-oob.xrm-ms\"";

                    }
                    string text4 = str3 + "\"" + str2 + "\\client-issuance-bridge-office.xrm-ms\"";
                    string text5 = str3 + "\"" + str2 + "\\client-issuance-root.xrm-ms\"";
                    string text6 = str3 + "\"" + str2 + "\\client-issuance-root-bridge-test.xrm-ms\"";
                    string text7 = str3 + "\"" + str2 + "\\client-issuance-stil.xrm-ms\"";
                    string text8 = str3 + "\"" + str2 + "\\client-issuance-ul.xrm-ms\"";
                    string text9 = str3 + "\"" + str2 + "\\client-issuance-ul-oob.xrm-ms\"";
                    string text10 = str3 + "\"" + str2 + "\\pkeyconfig-office.xrm-ms\"";
                    string text11 = "exit";
                    string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.System);
                    if (File.Exists(folderPath + "\\Retail2Vol.cmd"))
                    {
                        File.Delete(folderPath + "\\Retail2Vol.cmd");
                    }
                    if (File.Exists(folderPath + "\\Retail2Vol.vbs"))
                    {
                        File.Delete(folderPath + "\\Retail2Vol.vbs");
                    }
                    StreamWriter streamWriter = new StreamWriter(folderPath + "\\Retail2Vol.cmd", true, Encoding.GetEncoding("GB2312"));
                    streamWriter.WriteLine(string.Concat(new string[]
                    {
                text1,
                "\r\n",
                text2,
                "\r\n",
                text3,
                "\r\n\r\n",
                text4,
                "\r\n",
                text5,
                "\r\n",
                text6,
                "\r\n",
                text7,
                "\r\n",
                text8,
                "\r\n",
                text9,
                "\r\n",
                text10,
                "\r\n",
                text11
                    }));
                    streamWriter.Close();
                    StreamWriter streamWriter2 = new StreamWriter(folderPath + "\\Retail2Vol.vbs", true, Encoding.GetEncoding("GB2312"));
                    streamWriter2.WriteLine("createobject(\"wscript.shell\").run \"" + folderPath + "\\Retail2Vol.cmd\",0");
                    streamWriter2.Close();
                    Process process1 = new Process();
                    process1.StartInfo.FileName = "cmd.exe";
                    process1.StartInfo.UseShellExecute = false;
                    process1.StartInfo.RedirectStandardInput = true;
                    process1.StartInfo.RedirectStandardOutput = true;
                    process1.StartInfo.RedirectStandardError = true;
                    process1.StartInfo.CreateNoWindow = true;
                    process1.Start();
                    process1.StandardInput.WriteLine("start " + folderPath + "\\Retail2Vol.vbs");
                    process1.StandardInput.WriteLine("exit");
                    process1.WaitForExit();
                    process1.Close();

                    patch = "cd /d " + value.ToString() + "\\office16";
                    if (value.ToString().Contains("ProPlus2019Retail"))
                    {
                        officekey = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP";
                    }
                    else
                    {
                        officekey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99";
                    }
                   
                }
                if (value3 != null)
                {
                    patch = "cd /d " + value3.ToString();
                    officekey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99";
                }
                else if (value4 != null)
                {
                    patch = "cd /d " + value4.ToString();
                    officekey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99";
                }
                else if (value5 != null)
                {
                    patch = "cd /d " + value5.ToString();
                    officekey = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT";
                }
                else if (value6 != null)
                {
                    patch = "cd /d " + value6.ToString();
                    officekey = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT";
                }
                else if (value7 != null)
                {
                    patch = "cd /d " + value7.ToString();
                    officekey = "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB";
                }
                else if (value8 != null)
                {
                    patch = "cd /d " + value8.ToString();
                    officekey = "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB";
                }
                
                Process officeprocess = new Process();
                officeprocess.StartInfo.FileName = "cmd.exe";
                officeprocess.StartInfo.UseShellExecute = false;
                officeprocess.StartInfo.RedirectStandardInput = true;
                officeprocess.StartInfo.RedirectStandardOutput = true;
                officeprocess.StartInfo.RedirectStandardError = true;
                officeprocess.StartInfo.CreateNoWindow = true;
                officeprocess.Start();
                officeprocess.StandardInput.WriteLine(patch);
                officeprocess.StandardInput.WriteLine("cscript ospp.vbs /inpkey:{0}",officekey);
                officeprocess.StandardInput.WriteLine("cscript ospp.vbs /sethst:{0}", kmsserver);
                officeprocess.StandardInput.WriteLine("cscript ospp.vbs /act >{0}",Application.StartupPath +@"\Office.log");
                officeprocess.StandardInput.WriteLine("cscript ospp.vbs /remhst");
                officeprocess.StandardInput.WriteLine("exit");
                officeprocess.WaitForExit();
                officeprocess.Close();
                
                MessageBox.Show("您的" + text + "和MS Offie激活完成！", "by:慕若曦", MessageBoxButtons.OK, MessageBoxIcon.Information);
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "by:慕若曦");

            }
            finally
            {
                button1.Text = "永久体验Windows和Office正版";
                button1.Enabled = true;
                this.TopMost = true;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label3.Parent = pictureBox1;
            label4.Parent = pictureBox1;
            label4.MouseDown += PictureBox1_MouseDown;
            label4.MouseMove += PictureBox1_MouseMove;
            label4.MouseUp += PictureBox1_MouseUp;
            string str;
            if (Directory.Exists(Environment.GetEnvironmentVariable("systemroot").Substring(0, 3) + "Program Files (x86)"))
            {
                str = " 64位";
            }
            else
            {
                str = " 32位";
            }
            RegistryKey localMachine = Registry.LocalMachine;
            string text = localMachine.OpenSubKey("software", true).OpenSubKey("microsoft", true).OpenSubKey("Windows NT", true).OpenSubKey("CurrentVersion", true).GetValue("ProductName").ToString();
            localMachine.Close();
            switch (text)
            {
                case "Windows 10 Pro for Workstations":
                    label1.Text += "Windows 10 专业工作站版" + str;
                    break;
                case "Windows 10 Home":
                    label1.Text += "Windows 10 家庭版" + str;
                    break;
                case "Windows 10 Home Single Lanuage":
                    label1.Text += "Windows 10 家庭单语言版" + str;
                    break;
                case "Windows 10 Home China":
                    label1.Text += "Windows 10 家庭中文版" + str;
                    break;
                case "Windows 10 Education":
                    label1.Text += "Windows 10 教育版" + str;
                    break;
                case "Windows 10 Pro":
                    label1.Text += "Windows 10 专业版" + str;
                    break;
                case "Windows 10 Professional Education":
                    label1.Text += "Windows 10 专业教育版" + str;
                    break;
                case "Windows 10 Enterprise G":
                    label1.Text += "Windows 10 企业版G" + str;
                    break;
                case "Windows 10 Enterprise":
                    label1.Text += "Windows 10 企业版" + str;
                    break;
                case "Windows 10 Enterprise 2016 LTSB":
                    label1.Text += "Windows 10 企业版 2016 LTSB" + str;
                    break;
                case "Windows 10 Enterprise 2015 LTSB":
                    label1.Text += "Windows 10 企业版 2015 LTSB" + str;
                    break;
                case "Windows 8.1 Pro":
                    label1.Text += "Windows 8.1 专业版" + str;
                    break;
                case "Windows 8.1 Enterprise":
                    label1.Text += "Windows 8.1 企业版" + str;
                    break;
                case "Windows 8.1 Pro with Media Center":
                    label1.Text += "Windows 8.1 专业版(含WMC)" + str;
                    break;
                case "Windows 8.1":
                    label1.Text += "Windows 8.1 核心版" + str;
                    break;
                case "Windows 8.1 China":
                    label1.Text += "Windows 8.1 中文版" + str;
                    break;
                case "Windows 8.1 Single Language":
                    label1.Text += "Windows 8.1 单语言版" + str;
                    break;
                case "Windows 8 Pro":
                    label1.Text += "Windows 8 专业版" + str;
                    break;
                case "Windows 8 Enterprise":
                    label1.Text += "Windows 8 企业版" + str;
                    break;
                case "Windows 8":
                    label1.Text += "Windows 8 核心版" + str;
                    break;
                case "Windows 8 Pro with Media Center":
                    label1.Text += "Windows 8 专业版(含WMC)" + str;
                    break;
                case "Windows 8 China":
                    label1.Text += "Windows 8 中文版" + str;
                    break;
                case "Windows 8 Single Language":
                    label1.Text += "Windows 8 单语言版" + str;
                    break;
                case "Windows 7 Starter":
                    label1.Text += "Windows 7 初级版" + str;
                    break;
                case "Windows 7 Home Basic":
                    label1.Text += "Windows 7 家庭普通版" + str;
                    break;
                case "Windows 7 Home Premium":
                    label1.Text += "Windows 7 家庭高级版" + str;
                    break;
                case "Windows 7 Professional":
                    label1.Text += "Windows 7 专业版" + str;
                    break;
                case "Windows 7 Enterprise":
                    label1.Text += "Windows 7 企业版" + str;
                    break;
                case "Windows 7 Ultimate":
                    label1.Text += "Windows 7 旗舰版" + str;
                    break;
                default:
                    label1.Text += text + str;
                    break;
            }
            try {
                var ipAddress = System.Net.Dns.GetHostAddresses(kmsserver);
                if (ipAddress[0].ToString() != "58.216.219.36")
                {
                    MessageBox.Show("您可能没有联网或者KMS服务器已经失效！\n如果是后者那我想这个工具可能不再适用，请使用新版本。", "警告：");
                }
            }
            catch
            {
                MessageBox.Show("您可能没有联网！\n这个工具只能在联网状态下使用。", "警告：");
            }
               
        }

        private void Label3_Click(object sender, EventArgs e) => Close();

        private void PictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if(e.Button == MouseButtons.Left)
{
                beginMove = true;
                currentXPosition = MousePosition.X;
                currentYPosition = MousePosition.Y;
            }
        }

        private void PictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (beginMove)
            {
                Left += MousePosition.X - currentXPosition;
                Top += MousePosition.Y - currentYPosition;
                currentXPosition = MousePosition.X;
                currentYPosition = MousePosition.Y;
            }
        }

        private void PictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                currentXPosition = 0; 
                currentYPosition = 0;
                beginMove = false;
            }
        }
    }
}

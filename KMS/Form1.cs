using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace KMS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://www.muruoxi.com");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                button1.Text = "正在激活中，请等待。。。";
                button1.Enabled = false;
                RegistryKey localMachine = Registry.LocalMachine;
                string text = localMachine.OpenSubKey("software", true).OpenSubKey("microsoft", true).OpenSubKey("Windows NT", true).OpenSubKey("CurrentVersion", true).GetValue("ProductName").ToString();
                localMachine.Close();
                Process process = new Process();
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                if (text == "Windows 10 Pro for Workstations")
                {
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com" );
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                }
                else if (text == "Windows 10 Enterprise G")
                {
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                 
                }
                else if (text == "Windows 10 Enterprise")
                {
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                }
                else if (text == "Windows 10 Pro")
                {
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                }
                else if (text == "Windows 10 Education")
                {
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato" );
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    button1.Enabled = true;
                }
                else if (text == "Windows 10 Professional Education")
                {
       
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");

                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 10 Home")
                {

                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");

                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                }
                else if (text == "Windows 10 Home China")
                {

                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                 
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();

                }
                else if (text == "Windows 10 Home Single Language")
                {
        
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                   
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato" );
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 10 Enterprise 2016 LTSB")
                {
               
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                   
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato" );
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 10 Enterprise 2015 LTSB")
                {
                   
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato >");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 8.1 Pro")
                {
                   
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                   
                }
                else if (text == "Windows 8.1 Enterprise")
                {
                  
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                  
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                   
                }
                else if (text == "Windows 8.1 Pro with Media Center")
                {
                  
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk 789NJ-TQK6T-6XTH8-J39CJ-J8D3P");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                   
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 8.1")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 8.1 China")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3 ");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 8.1 Single Language")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk BB6NG-PQ82V-VRDPW-8XVD2-V8P66 ");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                   
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 8 Pro")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk NG4HW-VH26C-733KW-K6F98-J8CK4");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                   
                }
                else if (text == "Windows 8 Enterprise")
                {
                   
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk 32JNW-9KQ84-P47T8-D8GGY-CWCK7");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                   
                }
                else if (text == "Windows 8")
                {
                
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk BN3D2-R7TKB-3YPBD-8DRP2-27GG4");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                   
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 8 Pro with Media Center")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk GNBB8-YVD74-QJHX6-27H4K-8QHDG");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 8 China")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk 4K36P-JN4VD-GDC6V-KDT89-DYFKP ");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 8 Single Language")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk 2WN2H-YGCQR-KFX6K-CD6TF-84YXQ ");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato" );
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                   
                }
                else if (text == "Windows 7 Professional")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text == "Windows 7 Enterprise")
                {
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /skms smk.msdn123.com");
                    
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ato");
                    process.StandardInput.WriteLine("cscript %windir%\\System32\\slmgr.vbs /ckms");
                    process.StandardInput.WriteLine("exit");
                    process.WaitForExit();
                    process.Close();
                    
                }
                else if (text.Contains("Preview") || text.Contains("Evaluation"))
                {
                    MessageBox.Show("你当前的系统为测试评估版本，请更换为正式版后激活", "by：慕若曦");
                    button1.Enabled = true;
                }
                else if (text == "Windows 7 Starter" || text == "Windows 7 Home Basic" || text == "Windows 7 Home Premium" || text == "Windows 7 Ultimate")
                {
                    button1.Enabled = true;
                }
                else if (text.Contains("Windows 7"))
                {
                    button1.Enabled = true;
                }
                else
                {
                    MessageBox.Show("1、你的系统版本不对或系统激活机制遭到了破坏\r2、你安装的版本为测试评估版本", "by:慕若曦");

                    button1.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "by:慕若曦");
                button1.Enabled = true;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
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
            Console.WriteLine(text);
            if (text == "Windows 10 Pro for Workstations")
            {
                label1.Text = label1.Text + "Windows 10 专业工作站版" + str;
            }
            else if (text == "Windows 10 Home")
            {
                label1.Text += "Windows 10 家庭版" + str;
            }
            else if (text == "Windows 10 Home Single Lanuage")
            {
                label1.Text += "Windows 10 家庭单语言版" + str;
            }
            else if (text == "Windows 10 Home China")
            {
                label1.Text += "Windows 10 家庭中文版" + str;
            }
            else if (text == "Windows 10 Education")
            {
                label1.Text += "Windows 10 教育版" + str;
            }
            else if (text == "Windows 10 Pro")
            {
                label1.Text += "Windows 10 专业版" + str;
            }
            else if (text == "Windows 10 Professional Education")
            {
                label1.Text += "Windows 10 专业教育版" + str;
            }
            else if (text == "Windows 10 Enterprise G")
            {
                label1.Text += "Windows 10 企业版G" + str;
            }
            else if (text == "Windows 10 Enterprise")
            {
                label1.Text += "Windows 10 企业版" + str;
            }
            else if (text == "Windows 10 Enterprise 2016 LTSB")
            {
                label1.Text += "Windows 10 企业版 2016 LTSB" + str;
            }
            else if (text == "Windows 10 Enterprise 2015 LTSB")
            {
                label1.Text += "Windows 10 企业版 2015 LTSB" + str;
            }
            else if (text == "Windows 8.1 Pro")
            {
                label1.Text += "Windows 8.1 专业版" + str;
            }
            else if (text == "Windows 8.1 Enterprise")
            {
                label1.Text += "Windows 8.1 企业版" + str;
            }
            else if (text == "Windows 8.1 Pro with Media Center")
            {
                label1.Text += "Windows 8.1 专业版(含WMC)" + str;
            }
            else if (text == "Windows 8.1")
            {
                label1.Text += "Windows 8.1 核心版" + str;
            }
            else if (text == "Windows 8.1 China")
            {
                label1.Text += "Windows 8.1 中文版" + str;
            }
            else if (text == "Windows 8.1 Single Language")
            {
                label1.Text += "Windows 8.1 单语言版" + str;
            }
            else if (text == "Windows 8 Pro")
            {
                label1.Text += "Windows 8 专业版" + str;
            }
            else if (text == "Windows 8 Enterprise")
            {
                label1.Text += "Windows 8 企业版" + str;
            }
            else if (text == "Windows 8")
            {
                label1.Text += "Windows 8 核心版" + str;
            }
            else if (text == "Windows 8 Pro with Media Center")
            {
                label1.Text += "Windows 8 专业版(含WMC)" + str;
            }
            else if (text == "Windows 8 China")
            {
                label1.Text += "Windows 8 中文版" + str;
            }
            else if (text == "Windows 8 Single Language")
            {
                label1.Text += "Windows 8 单语言版" + str;
            }
            else if (text == "Windows 7 Starter")
            {
                label1.Text += "Windows 7 初级版" + str;
            }
            else if (text == "Windows 7 Home Basic")
            {
                label1.Text += "Windows 7 家庭普通版" + str;
            }
            else if (text == "Windows 7 Home Premium")
            {
                label1.Text += "Windows 7 家庭高级版" + str;
            }
            else if (text == "Windows 7 Professional")
            {
                label1.Text += "Windows 7 专业版" + str;
            }
            else if (text == "Windows 7 Enterprise")
            {
                label1.Text += "Windows 7 企业版" + str;
            }
            else if (text == "Windows 7 Ultimate")
            {
                label1.Text += "Windows 7 旗舰版" + str;
            }
            else
            {
                label1.Text += text + str;
            }
        }
    }
}

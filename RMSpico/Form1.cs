using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.Threading;
using Microsoft.VisualBasic;
using System.IO;

namespace RMSpico
{
    public partial class Form1 : Form
    {

        String WinProKey = "W269N-WFGWX-YVC9B-4J6C9-T83GX";
        String WinHomeKey = "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99";
        String WinEnterpriseKey = "NPPR9-FWDCX-D2C8J-H872K-2YT43";

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
            {
                infoLabel.Text = "Please select a Windows Edition before trying to activate";
            }
            else
            {
                infoLabel.Text = "Activating Windows...";
                activateWindows(getSelectedEdition());
            }




        }

        public String getSelectedEdition()
        {
            if (comboBox1.Text == "Windows 10 Pro")
            {
                return WinProKey;
            }
            if (comboBox1.Text == "Windows 10 Home")
            {
                return WinHomeKey;
            }
            if (comboBox1.Text == "Windows 10 Enterprise")
            {
                return WinEnterpriseKey;
            }
            if (comboBox1.Text == "Custom Key")
            {
                String customKey = Interaction.InputBox("Enter Custom Key:","Custom Key","XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" ,500,500);
                return customKey;
            }
            return "no matching key";

        }

        public void activateWindows(String key)
        {
            if (key == "no matching key")
            {
                infoLabel.Text = "ERROR: No Matching Key for this version of Windows";
            }
            else
            {
                try
                {
                    progressBar1.Value = 0;
                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                    startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = "/C slmgr /ipk " + key;
                    startInfo.Verb = "runas";
                    process.StartInfo = startInfo;

                    process.Start();
                    progressBar1.Value = 40;

                    startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = "/C slmgr /skms kms8.msguides.com";
                    startInfo.Verb = "runas";
                    process.StartInfo = startInfo;

                    process.Start();

                    progressBar1.Value = 75;

                    startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = "/C slmgr /ato";
                    startInfo.Verb = "runas";
                    process.StartInfo = startInfo;

                    process.Start();
                    progressBar1.Value = 100;

                    infoLabel.Text = "Windows 10 Key Installed!";
                }
                catch
                {
                    infoLabel.Text = "Error installing Windows 10 Key";
                }
                
            }

        }

        public void activateOffice(String path)
        {
            if (path == "no matching path")
            {
                infoLabel.Text = "ERROR: No Matching Key for this version of Windows";
            }
            else
            {
                try
                {
                    progressBar1.Value = 0;
                    System.Diagnostics.Process officeProcess = new System.Diagnostics.Process();
                    System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                    startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = "/C cd /d " + path;
                    startInfo.Verb = "runas";
                    officeProcess.StartInfo = startInfo;

                    officeProcess.Start();
                    progressBar1.Value = 40;


                    startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = @"/C for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:'..\root\Licenses16\% x'";
                    startInfo.Verb = "runas";
                    officeProcess.StartInfo = startInfo;

                    officeProcess.Start();
                    progressBar1.Value = 70;



                    startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = @"/C cscript ospp.vbs /setprt:1688
                        cscript ospp.vbs /unpkey:6MWKP >nul
                        cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
                        cscript ospp.vbs /sethst:kms8.msguides.com
                        cscript ospp.vbs /act";
                    startInfo.Verb = "runas";
                    officeProcess.StartInfo = startInfo;

                    officeProcess.Start();
                    progressBar1.Value = 100;


                    infoLabel.Text = "Office 2019 Key Installed!";
                }
                catch
                {
                    infoLabel.Text = "Error installing Office 2019 Key";
                }

            }

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string r = "";
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem"))
            {
                ManagementObjectCollection information = searcher.Get();
                if (information != null)
                {
                    foreach (ManagementObject obj in information)
                    {
                        r = obj["Caption"].ToString() + " - " + obj["OSArchitecture"].ToString();
                    }
                }
                r = r.Replace("NT 5.1.2600", "XP");
                r = r.Replace("NT 5.2.3790", "Server 2003");
                MessageBox.Show("You are using Windows Edition: \r\n" + r);
            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://msguides.com/");
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                label6.Text = "Please enter a valid file path";
            }
            else
            {
                label6.Text = "Activating Office 2019...";
                activateOffice(textBox1.Text);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (FolderBrowserDialog openFileDialog = new FolderBrowserDialog())
            {


                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.SelectedPath;


                }
            }
            textBox1.Text = filePath;
        }
    }
}

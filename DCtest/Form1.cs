using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using dxp01sdk;
using System.Management;
using System.Collections;
using System.Drawing.Printing;

namespace DCtest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.AddRange(GetPrinters());
        }

        

        

        private string[] GetPrinters()
        {
            ManagementObjectSearcher searchPrinters =
             new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
            ManagementObjectCollection printerCollection = searchPrinters.Get();

            ArrayList printers = new ArrayList();
            foreach (ManagementObject printer in printerCollection)
            {
                string printerName = printer.Properties["Name"].Value.ToString();
                //				if (printerName.IndexOf("Copy") < 0)
                printers.Add(printerName);
            }

            return (string[])printers.ToArray(typeof(string));
        }

        private void GetPrinterProperties(String printerName)
        {
            BidiSplWrap bidiSpl = null;
            int printerJobID = 0;
            try
            {
              
                bidiSpl = new BidiSplWrap();
                String BindResponse = bidiSpl.BindDevice(printerName);
                textBox2.Text += Environment.NewLine + BindResponse;
                string driverVersionXml = bidiSpl.GetPrinterData(strings.SDK_VERSION);
                textBox2.Text += Environment.NewLine + "driver version: " + Util.ParseDriverVersionXML(driverVersionXml) + Environment.NewLine;

                string printerOptionsXML = bidiSpl.GetPrinterData(strings.PRINTER_OPTIONS2);
                PrinterOptionsValues printerOptionsValues = Util.ParsePrinterOptionsXML(printerOptionsXML);

                
                if ("Ready" != printerOptionsValues._printerStatus && "Busy" != printerOptionsValues._printerStatus)
                {
                    throw new Exception(printerName + " is not ready. status: " + printerOptionsValues._printerStatus);
                }
                else
                {
                    string printerOptionsXMLSupply = bidiSpl.GetPrinterData(strings.SUPPLIES_STATUS3);
                   SuppliesValues printerOptionsSupplyValues = Util.ParseSuppliesXML(printerOptionsXMLSupply);

                    textBox2.Text += Environment.NewLine+"Printer Status: " + printerOptionsValues._printerStatus;
                    textBox2.Text += Environment.NewLine + "Ribbon Type: " + printerOptionsSupplyValues._printRibbonType;
                    textBox2.Text += Environment.NewLine + "Ribbon Serial Number: " + printerOptionsSupplyValues._printRibbonSerialNumber;
                    textBox2.Text += Environment.NewLine + "Ribbon Level: " + printerOptionsSupplyValues._ribbonRemaining + "%";

                }


            }
            catch (BidiException e) {
                textBox2.Text += Environment.NewLine + e.Message;
                Util.CancelJob(bidiSpl,e.PrinterJobID, e.ErrorCode);
            }
            catch (Exception e) {
                textBox2.Text += Environment.NewLine + e.Message;
                
                if (0 != printerJobID)
                {
                    Util.CancelJob(bidiSpl, printerJobID, 0);
                }
            }
            finally
            {
                bidiSpl.UnbindDevice();
            }
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            String printerName = comboBox1.Text;
            GetPrinterProperties(printerName);

            ///printing is activated here

            BidiSplWrap bidiSpl = null;
            int printerJobID = 0;
            try
            {

                bidiSpl = new BidiSplWrap();
                String BindResponse = bidiSpl.BindDevice(printerName);
                textBox2.Text += Environment.NewLine + BindResponse;
                string driverVersionXml = bidiSpl.GetPrinterData(strings.SDK_VERSION);
                textBox2.Text += Environment.NewLine + "driver version: " + Util.ParseDriverVersionXML(driverVersionXml) + Environment.NewLine;

                string printerOptionsXML = bidiSpl.GetPrinterData(strings.PRINTER_OPTIONS2);
                PrinterOptionsValues printerOptionsValues = Util.ParsePrinterOptionsXML(printerOptionsXML);


                if ("Ready" != printerOptionsValues._printerStatus && "Busy" != printerOptionsValues._printerStatus)
                {
                    throw new Exception(printerName + " is not ready. status: " + printerOptionsValues._printerStatus);
                }
                else
                {

                    string hopperID = string.Empty;
                   
                        printerJobID = Util.StartJob(
                            bidiSpl,
                            false ,
                            hopperID);

                    Application.DoEvents();

                    Util.PrintTextAndGraphics(bidiSpl, printerName,textBox1.Text);
                    Application.DoEvents();
                    if (0 != printerJobID)
                    {
                        // wait for the print spooling to finish and then issue an EndJob():
                        Util.WaitForWindowsJobID(bidiSpl, printerName);
                        bidiSpl.SetPrinterData(strings.ENDJOB);
                    }

                    Application.DoEvents();
                    Util.PollForJobCompletion(bidiSpl, printerJobID);
                    
                }


            }
            catch (BidiException ee)
            {
                textBox2.Text += Environment.NewLine + ee.Message;
                Util.CancelJob(bidiSpl, ee.PrinterJobID, ee.ErrorCode);
            }
            catch (Exception ee)
            {
                textBox2.Text += Environment.NewLine + ee.Message;

                if (0 != printerJobID)
                {
                    Util.CancelJob(bidiSpl, printerJobID, 0);
                }
            }
            finally
            {
                bidiSpl.UnbindDevice();
            }

        }

        private void OnPrintPage(object sender, PrintPageEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String printerName = comboBox1.Text;
            GetPrinterProperties(printerName);
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            magForm MF = new magForm();
            MF.ShowDialog();
        }
    }
}

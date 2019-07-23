using dxp01sdk;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DCtest
{
    public partial class magForm : Form
    {
        public magForm()
        {
            InitializeComponent();
        }

        private void GetPrinterProperties(String printerName)
        {
            BidiSplWrap bidiSpl = null;
            int printerJobID = 0;
            try
            {

                bidiSpl = new BidiSplWrap();
                String BindResponse = bidiSpl.BindDevice(printerName);
                textBox4.Text += Environment.NewLine + BindResponse;
                string driverVersionXml = bidiSpl.GetPrinterData(strings.SDK_VERSION);
                textBox4.Text += Environment.NewLine + "driver version: " + Util.ParseDriverVersionXML(driverVersionXml) + Environment.NewLine;

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

                    textBox4.Text += Environment.NewLine + "Printer Status: " + printerOptionsValues._printerStatus;
                    textBox4.Text += Environment.NewLine + "Ribbon Type: " + printerOptionsSupplyValues._printRibbonType;
                    textBox4.Text += Environment.NewLine + "Ribbon Serial Number: " + printerOptionsSupplyValues._printRibbonSerialNumber;
                    textBox4.Text += Environment.NewLine + "Ribbon Level: " + printerOptionsSupplyValues._ribbonRemaining + "%";

                }


            }
            catch (BidiException e)
            {
                textBox4.Text += Environment.NewLine + e.Message;
                Util.CancelJob(bidiSpl, e.PrinterJobID, e.ErrorCode);
            }
            catch (Exception e)
            {
                textBox4.Text += Environment.NewLine + e.Message;

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

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void magForm_Load(object sender, EventArgs e)
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String printerName = comboBox1.Text;
            GetPrinterProperties(printerName);
        }



        static void ResumeJob(BidiSplWrap bidiSpl, int printerJobID, int errorCode)
        {
            string xmlFormat = strings.PRINTER_ACTION_XML;
            string input = string.Format(xmlFormat, (int)Actions.Resume, printerJobID, errorCode);
            bidiSpl.SetPrinterData(strings.PRINTER_ACTION, input);
            Application.DoEvents();
        }



        private void writeBtn_Click(object sender, EventArgs e)
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
                textBox4.Text += Environment.NewLine + BindResponse;
                string driverVersionXml = bidiSpl.GetPrinterData(strings.SDK_VERSION);
                textBox4.Text += Environment.NewLine + "driver version: " + Util.ParseDriverVersionXML(driverVersionXml) + Environment.NewLine;

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
                        false,
                        hopperID);

                    Application.DoEvents();

                    Util.EncodeMagstripe(bidiSpl, false, textBox1.Text, textBox2.Text, textBox3.Text);

                  //   tried to cancel job;
                  //  if () 
                  //  {
                  //      Util.CancelJob(bidiSpl,printerJobID,0);
                  //  }

                }
            }
            catch (Exception ee) { }
        } 

        private void readBtn_Click(object sender, EventArgs e)
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
                textBox4.Text += Environment.NewLine + BindResponse;
                string driverVersionXml = bidiSpl.GetPrinterData(strings.SDK_VERSION);
                textBox4.Text += Environment.NewLine + "driver version: " + Util.ParseDriverVersionXML(driverVersionXml) + Environment.NewLine;

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
                        false,
                        hopperID);

                    Application.DoEvents();

                    String[] returnedArray = Util.ReadMagstripe(bidiSpl, false);


                    Application.DoEvents();

                    textBox1.Text += returnedArray[0];
                    textBox2.Text += returnedArray[1];
                    textBox3.Text += returnedArray[2];


                    textBox4.Text += "Magstripe data read successfully; printer job id: " + printerOptionsValues._printerStatus;
                }
            }
            catch (Exception ee) { }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            Form1 F1 = new Form1();
            F1.ShowDialog();
        }
    }
}
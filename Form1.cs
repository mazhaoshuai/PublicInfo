using System;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Management;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Win32;
using System.Threading;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace testCOM
{
    public partial class Form1 : Form
    {
        //Creating serial objects
        SerialPort serialPort=new SerialPort();
        //Create a delegate
        public delegate void Displaydelegate(byte[] buf);
        //Creating a delegate object
        public Displaydelegate disp_delegate;
        //Defining first start (First data random code  )
        bool isComOne = false;

        public Form1()
        {
            InitMyData();
            InitializeComponent();
        }

        /// <summary>
        /// Initialization of data 
        /// </summary>
        private void InitMyData()
        {
            //initialize serial port settings 
            serialPort = new SerialPort(searchDevicesRegistry(), 115200, Parity.None, 8, StopBits.One);  
            // Executive Commission 
            disp_delegate = new Displaydelegate(DispUI);
            //Data reception processing 
            serialPort.DataReceived += new SerialDataReceivedEventHandler(CommDataReceived); 
        }

        /// <summary>
        /// VID+PID gets the serial port 
        /// </summary>
        /// <returns></returns>
        private string searchDevicesRegistry()
        {
            string[] available_spectrometers = SerialPort.GetPortNames();
            ManagementObjectCollection.ManagementObjectEnumerator enumerator = null;
            string commData = "";
            ManagementObjectSearcher mObjs = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM WIN32_PnPEntity");
            try
            {
                enumerator = mObjs.Get().GetEnumerator();
                while (enumerator.MoveNext())
                {
                    ManagementObject current = (ManagementObject)enumerator.Current;

                    if (Strings.InStr(Conversions.ToString(current["Caption"]), "(COM", CompareMethod.Binary) <= 0)
                    {
                        continue;
                    }
                    //foreach (var property in current.Properties)
                    //{
                    //    Console.WriteLine(property.Name + ":" + property.Value);
                    //}
                    if ((current["ClassGuid"].ToString().Equals("{4d36e978-e325-11ce-bfc1-08002be10318}") & current["DeviceID"].ToString().Equals("FTDIBUS\\VID_0403+PID_6001+FTZ6XM7RA\\0000")))
                    {
                        commData = current["Name"].ToString().Substring(17);
                        commData = commData.Substring(0, checked(commData.Length - 1));
                        break;
                    }
                }
            }
            finally
            {
                if (enumerator != null)
                {
                    ((IDisposable)enumerator).Dispose();
                }
            }
            return commData;
        }

        /// <summary>
        /// Receive and parse the data 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void CommDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            int i = 0;
            StringBuilder sbgData = new StringBuilder();
            try
            {
                if (!isComOne)
                {
                    while (i < 6)
                    {
                        sbgData.Append(serialPort.ReadLine());
                        i++;
                    }
                    isComOne = true;
                }
                if (serialPort.IsOpen & serialPort.RtsEnable == true)
                {
                    byte[] received_bytes = new byte[1];
                    do
                    {
                        Thread.Sleep(70);
                        byte[] numArray = new byte[serialPort.BytesToRead];
                        serialPort.Read(numArray, 0, serialPort.BytesToRead);
                        this.CopyReceiveBytes(ref numArray, ref received_bytes);
                        Thread.Sleep(70);
                    }
                    while (serialPort.BytesToRead != 0);
                    this.Invoke(disp_delegate, received_bytes);
                }
            }
            catch (TimeoutException ex) 
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void DispUI(byte[] buf)
        {
            //Executing parse data 
            //Close Rts
            if (serialPort.BytesToWrite == 0)
            {
                serialPort.RtsEnable = false;
            }
        }

        /// <summary>
        /// Copy array 
        /// </summary>
        /// <param name="source"></param>
        /// <param name="destination"></param>
        private void CopyReceiveBytes(ref byte[] source, ref byte[] destination)
        {
            int length = 0;
            if (checked((int)destination.Length) != 0)
            {
                length = checked(checked((int)destination.Length) - 1);
                byte[] numArray = new byte[0];
                numArray = new byte[checked(checked(checked(checked((int)source.Length) + checked((int)destination.Length)) - 1) + 1)];
                destination.CopyTo(numArray, 0);
                destination = numArray;
            }
            else
            {
                destination = new byte[checked(checked(checked((int)source.Length) - 1) + 1)];
            }
            source.CopyTo(destination, length);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //startCmdTime_ms = DateTime.Now.Ticks / 10000;
            try
            {
                if (!serialPort.IsOpen)
                {
                    serialPort.Open();
                    button1.Text = "发送";
                    serialPort.RtsEnable = true;
                }
                else
                {
                    serialPort.RtsEnable = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "False hints ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        
        //long startCmdTime_ms;
        //int index = 0;
        //void CommDataReceived(object sender, SerialDataReceivedEventArgs e)
        //{
        //    int n = serialPort.BytesToRead;
        //    Byte[] buf = new Byte[n];
        //    int i = 0,j=0;
        //    StringBuilder sbgData = new StringBuilder();
        //    try
        //    {
        //        if (!isComOne)
        //        {
        //            while (j < 6)
        //            {
        //                sbgData.Append(serialPort.ReadLine());
        //                j++;
        //            }
        //            isComOne = true;
        //        }
        //        sbgData.Clear();
        //        while (i < 6)
        //        {
        //            sbgData.Append(serialPort.ReadLine());
        //            i++;
        //        }
        //        buf = Encoding.ASCII.GetBytes(sbgData.ToString());
        //        this.Invoke(disp_delegate, buf);
        //        //if (DateTime.Now.Ticks / 10000 - startCmdTime_ms; >= 1000)
        //        //{
        //        //}
        //        //else
        //        //{
        //        //    serialPort.RtsEnable = true;//无限触发 
        //        //    index++;
        //        //}
        //    }
        //    catch (TimeoutException ex)         //超时处理
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

        private void Form1_Load(object sender, EventArgs e)
        {
            serialPort.Close();
        }
    }
}

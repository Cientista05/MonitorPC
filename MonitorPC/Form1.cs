using System;
using System.IO.Ports;
using System.Diagnostics;
using System.Windows.Forms;
using OpenHardwareMonitor.Hardware;

namespace MonitorPC
{
    public partial class Form1 : Form
    {
        Computer thisComputer;

        public Form1()
        {
            thisComputer = new Computer(); 
            thisComputer.CPUEnabled = true;                                 
            thisComputer.HDDEnabled = true;
            thisComputer.MainboardEnabled = true;
            thisComputer.RAMEnabled = true;
            thisComputer.Open();

            InitializeComponent();
            init();
        }
        private void dataCheck()
        {
            string cpuName = "CPU";
            string cpuTemp = "C";
            string cpuLoad = "c";
            string CpuFanSpeedLoad = "";       

            string ramLoad = "RL";
            string ramUsed = "R";              
            string ramAvailable = "RA"; 

            string gpuName = "GPU";
            string gpuTemp = "G";
            string gpuLoad = "c";

            string gpuCoreClock = "GCC";
            string gpuMemoryClock = "GCM";
            string gpuShaderClock = "GSC";

            string gpuMemTotal = "GMT";        
            string gpuMemUsed = "GMU";         
            string gpuMemLoad = "GML";         

            string gpuFanSpeedLoad = "GFANL";  
            string gpuFanSpeedRPM = "GRPM";  
            string gpuPower = "GPWR";
            string cpuClock = "";
            int highestCPUClock = 0;
            
            // enumerating all the hardware
            foreach (IHardware hw in thisComputer.Hardware)                                                 // for each hardware item thisComputer
            {
                if (hw.HardwareType.ToString().IndexOf("CPU") > -1)
                {
                    cpuName = "CPU:";
                    cpuName += hw.Name;
                }
                else if (hw.HardwareType.ToString().IndexOf("Gpu") > -1)
                {
                    gpuName = "GPU:";
                    gpuName += hw.Name;
                }
                hw.Update();                                                                                //update it

                // searching for all sensors and adding data to listbox
                foreach (ISensor s in hw.Sensors)                                                           //for each sensor in the sensors part of the hardware
                {
                    //-------------------------------------------------------- Sensor CPU GPU Temps -----------------------------------------------------
                    if (s.SensorType == SensorType.Temperature)                                             // if the sensor type is a temperature sensor
                    {
                        if (s.Value != null)                                                                //if the value is not null
                        {
                            int curTemp = (int)s.Value;                                                     //create a new int and set its value to the temperature value
                            switch (s.Name)                                                                 // create a switch based on the sensor name
                            {
                                case "CPU Package":                                                         // if the name is "CPU package"
                                    cpuTemp = curTemp.ToString();                                           // set the string cpuTemp to the int value above converted to a string 
                                    break;                                                                  // break from the switch so it doesnt run the case below
                                case "GPU Core":                                                            //if the name is "GPU Core"
                                    gpuTemp = curTemp.ToString();
                                    break;
                            }
                        }
                    }

                    //-------------------------------------------------------- Sensor Type  Clocks ----------------------------------------------------

                    if (s.SensorType == SensorType.Clock)                                                   // if the sensor type is a temperature sensor
                    {
                        if (s.Value != null)                                                                //if the value is not null
                        {
                            int clockSpeed = (int)s.Value;                                                  //create a new int and set its value to the temperature value
                            switch (s.Name)                                                                 // create a switch based on the sensor name
                            {
                                // break from the switch so it doesnt run the case below
                                case "GPU Core":                                                            //if the name is "GPU Core"
                                    gpuCoreClock = "|GCC" + clockSpeed.ToString();
                                    break;
                                case "GPU Memory":                                                          //if the name is "GPU Memory"
                                    gpuMemoryClock = "|GMC" + clockSpeed.ToString();
                                    break;
                                case "GPU Shader":                                                          //if the name is "GPU Shader"
                                    gpuShaderClock = "|GSC" + clockSpeed.ToString();
                                    break;
                            }
                            if (s.Name.IndexOf("CPU Core") > -1)
                            {
                                if (clockSpeed > highestCPUClock)                                           // run through each iteration of CPU Core and if the speed is higher than the last save it
                                {
                                    highestCPUClock = clockSpeed;
                                    cpuClock = "|CHC" + highestCPUClock.ToString() + "|";
                                }
                            }
                        }
                    }

                    //-------------------------------------------------------- Sensor  Type Loads -----------------------------------------------------

                    if (s.SensorType == SensorType.Load)                                                    // if the sensor type is a load value
                    {
                        if (s.Value != null)                                                                // if the value is not null
                        {
                            int curLoad = (int)s.Value;                                                     // create a new int and set its value to the sensor value
                            switch (s.Name)                                                                 //create a switch based on the name again
                            {
                                case "CPU Total":                                                           //if the name is "CPU Total"
                                    cpuLoad = curLoad.ToString();                                           //set the string cpuLoad to the int value converted to a string
                                    break;
                                case "GPU Core":
                                    gpuLoad = curLoad.ToString();
                                    break;
                                case "Memory":                                                              // Bundle System Ram into into main load sensor
                                    ramLoad = curLoad.ToString();       
                                    break;

                                case "GPU Memory":                                                          //GPU Memory Load   // Bundle System Ram into into main load sensor
                                    gpuMemLoad = curLoad.ToString();
                                    break;
                            }
                        }
                    }

                    //-------------------------------------------------------- Sensor  Type Data -----------------------------------------------------

                    if (s.SensorType == SensorType.Data)                                                    // if the sensor is a data value etc.etc.
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                case "Available Memory":                                                    //if the name is "used memory"
                                    decimal decimalAram = Math.Round((decimal)s.Value, 1);                  // create a new decimal and set the value to the sensor value a rounded to 1 decimal place
                                    ramAvailable = decimalAram.ToString();                                  // set the ramused string to the decimal converted to a string
                                    break;
                            }
                        }
                    }
                    if (s.SensorType == SensorType.Data)                                                    // if the sensor is a data value etc.etc.
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                //case "Memory Used":                                                       // libreHardwareMonitor DLL uses Memory Used      //if the name is "memory used"
                                case "Used Memory":                                                         // OpenHardwareMonitor DLL uses Memory Used      //if the name is "used memory"
                                    decimal decimalUram = Math.Round((decimal)s.Value, 1);                  // create a new decimal and set the value to the sensor value a rounded to 1 decimal place
                                    ramUsed = decimalUram.ToString();                                       // set the ramused string to the decimal converted to a string
                                    break;
                            }
                        }
                    }

                    //-------------------------------------------------------- Sensor  Type Small Data -----------------------------------------------------
                    
                    if (s.SensorType == SensorType.SmallData)                                               // if the sensor is a data value etc.etc.
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                case "GPU Memory Total":                                                    //  see NvidiaGPU.cs "memoryAvail = new Sensor("GPU Memory Total", 3, SensorType.SmallData, this, settings);"
                                    decimal decimalGtram = Math.Round((decimal)s.Value, 0);                 // create a new decimal and set the value to the sensor value a rounded to 1 decimal place
                                    gpuMemTotal = decimalGtram.ToString();                                  // set the gpuMemTotal string to the decimal converted to a string
                                    break;
                            }
                        }
                    }
                    if (s.SensorType == SensorType.SmallData)                                               // if the sensor is a data value etc.etc.
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                case "GPU Memory Used":                                                     // MB & GB      //  see NvidiaGPU.cs memoryUsed = new Sensor("GPU Memory Used", 2, SensorType.SmallData, this, settings);
                                    decimal decimalGuram = Math.Round((decimal)s.Value, 0);                 // create a new decimal and set the value to the sensor value a rounded to 1 decimal place
                                    gpuMemUsed = decimalGuram.ToString();                                   // set the gpuMemUsed string to the decimal converted to a string
                                    break;
                            }
                        }
                    }

                    //-------------------------------------------------------- Sensor Type  FAN RPM -----------------------------------------------------

                    if (s.SensorType == SensorType.Fan)                                                     // if the sensor is a data value etc.etc.
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                case "GPU":                                                                 // GPU Fan Speed RPM  this does not work need more investigating
                                    decimal decimalfanRPM = Math.Round((decimal)s.Value, 1);                // create a new decimal and set the value to the sensor value a rounded to 1 decimal place
                                    gpuFanSpeedRPM = decimalfanRPM.ToString();                              // set the gpuMemTotal string to the decimal converted to a string

                                    break;
                            }
                        }
                    }

                    //-------------------------------------------------------- Sensor Type  Control  %-----------------------------------------------------

                    if (s.SensorType == SensorType.Control)                                                 // if the sensor is a data value etc.etc.
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                case "GPU Fan":                                                             // GPU Fan Speed %
                                    decimal decimalGfan = Math.Round((decimal)s.Value, 1);                  // create a new decimal and set the value to the sensor value a rounded to 1 decimal place
                                    gpuFanSpeedLoad = decimalGfan.ToString();                               // set the gpuMemTotal string to the decimal converted to a string

                                    break;
                            }
                        }
                    }

                    //-------------------------------------------------------- Sensor Type Power -----------------------------------------------------

                    if (s.SensorType == SensorType.Power)                                                   // if the sensor is a data value etc.etc.
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                case "GPU Power":                                                           // GPU Power Watts
                                    decimal decimalGpwr = Math.Round((decimal)s.Value, 1);                  // create a new decimal and set the value to the sensor value a rounded to 1 decimal place
                                    gpuPower = decimalGpwr.ToString();                                      // set the gpuMemTotal string to the decimal converted to a string

                                    break;
                            }
                        }
                    }

                    //-------------------------------------------------------- Sensor CPU Fan Percentage MainBoard  (Needs Work)-----------------------------------------------------
                   
                    if (s.SensorType == SensorType.Control)                                                 // if the sensor is a data value etc.etc.
                                                                                                            //if (s.SensorType == OpenHardwareMonitor.Hardware.Mainboard.Mainboard)           // if the sensor is a data value etc.etc.
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                case "CPU Fan":                                                             // CPU Fan Speed %   needs fixing  SuperIOHardware.cs ??????????????????
                                    decimal decimalCfan = Math.Round((decimal)s.Value, 1);                  // create a new decimal and set the value to the sensor value a rounded to 1 decimal place
                                    CpuFanSpeedLoad = decimalCfan.ToString();                               // set the gpuMemTotal string to the decimal converted to a string

                                    break;
                            }
                        }
                    }
                }

                //-------------------------------------------------------- Sensor CPU Average Temps -----------------------------------------------------

                if (cpuTemp == "")                                                                          // if there is no cpuTemp assigned from earlier functions, get the average cpu temp
                {
                    foreach (ISensor s in hw.Sensors)                                                       //for each sensor in the sensors part of the hardware
                    {
                        int numTemps = 0;
                        int averageTemp = 0;
                        try
                        {
                            if (s.SensorType == SensorType.Temperature)                                     // if the sensor type is a temperature sensor
                            {
                                if (s.Name.IndexOf("CPU Core") > -1)
                                {
                                    averageTemp = averageTemp + (int)s.Value;
                                    numTemps++;
                                }
                            }
                            if (numTemps > 0)
                            {
                                averageTemp = averageTemp / numTemps;
                                cpuTemp = averageTemp.ToString();
                            }
                        }
                        catch { }
                    }
                }
            }

            //-------------------------------------------------------- Serial Arduino Strings -----------------------------------------------------

            string stats = string.Empty;                                                                    //create a new string and instantiate it as empty

            //stats = "C" + cpuTemp + "c " + cpuLoad + "%|G" + gpuTemp + "c " + gpuLoad + "%|R" + ramUsed + "G|"; // HSM v1.1b write the strings to the new string along with separators and denotations the arduino can understand

            //stats = "C" + cpuTemp + "c " + cpuLoad + "%|G" + gpuTemp + "c " + gpuLoad + "%|R" + ramUsed + "|RA" + ramAvailable + "GB|RL" + ramLoad + "%|GMT" + gpuMemTotal + "MB|GML" + gpuMemLoad + "%|GFANL" + gpuFanSpeedLoad + "%|GPWR" + gpuPower + "W|GMU" + gpuMemUsed + "MB|GRPM" + gpuFanSpeedRPM + "RPM|";  //"RPM |CFANL" + CpuFanSpeedLoad + "%|";  //  write the strings to the new string along with separators and denotations the arduino can understand

            stats = "C" + cpuTemp + "c " + cpuLoad + "%|G" + gpuTemp + "c " + gpuLoad + "%|R" + ramUsed + "GB|RA" + ramAvailable + "|RL" + ramLoad +
              "|GMT" + gpuMemTotal + "|GMU" + gpuMemUsed + "|GML" + gpuMemLoad + "|GFANL" + gpuFanSpeedLoad + "|GRPM" + gpuFanSpeedRPM + "|GPWR" + gpuPower + "|";  //"RPM |CFANL" + CpuFanSpeedLoad + "%|";  //  write the strings to the new string along with separators and denotations the arduino can understand


            if (stats != string.Empty)//so long as its not empty
            {
                sendToArduino(stats + cpuName + gpuName + gpuCoreClock + gpuMemoryClock + gpuShaderClock + cpuClock);//send the string to the function                                                                                                                     // sendToArduino(cpuName+gpuName);
            }
        }
        //function to send the data to the arduino over the com port       
        private void sendToArduino(string arduinoData)
        {
            if (serialPort1.IsOpen)                                                                         //if the port is not open
            {
                try
                {
                    serialPort1.WriteLine(arduinoData);                                                         //try write to the manual port
                }
                catch (Exception e)
                {
                    Debug.WriteLine("Error sending Serial Data: " + e.ToString());                              //catch any errors
                    string errorString = e.ToString();
                }
            }
        }
        private void init()
        {
            try
            {
                string[] ports = SerialPort.GetPortNames();
                foreach (string port in ports)
                {
                    comboBox1.Items.Add(port);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            dataCheck();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!serialPort1.IsOpen)
                {
                    serialPort1.PortName = comboBox1.Text;
                    serialPort1.Open();
                    timer1.Enabled = true;
                    timer1.Interval = Convert.ToInt32(comboBox2.Text);
                    toolStripStatusLabel1.Text = "Sending data...";
                    label3.Text = "Connected";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                serialPort1.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            label3.Text = "Disconnected";
            timer1.Enabled = false;
            toolStripStatusLabel1.Text = "Connect to Arduino...";
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                notifyIcon1.Visible = true;
                try
                {
                    notifyIcon1.ShowBalloonTip(500, "Arduino", toolStripStatusLabel1.Text, ToolTipIcon.Info);
                }
                catch (Exception)
                {

                }
                this.Hide();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            serialPort1.Close();
        }
    }         
}
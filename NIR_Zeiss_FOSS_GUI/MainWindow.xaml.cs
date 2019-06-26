using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO.Ports;
using System.ComponentModel;
using System.Windows.Threading;
using System.Threading;
using OxyPlot;
using OxyPlot.Axes;
using System.Management;
using System.Text.RegularExpressions;
using System.IO;
using Squirrel;
using Microsoft.Win32;

namespace NIR_Zeiss_FOSS_GUI
{   //Squirrel --releasify MyApp.1.0.0.nupkg
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //private DispatcherTimer dispatcherTimerFOSSPort;
        private System.Threading.Timer dispatcherTimerFOSSPort;
        private SerialPort FOSSserialPort;
        private SerialPort MOTORserialPort;
        private DateTime lastSerRec = new DateTime();
        private DateTime lastWhiteRef = new DateTime();
        private DateTime lastDarkRef = new DateTime();
        private ManagementObjectCollection queryCollection;
        //string gateway = "192.168.0.1";
        private readonly string subnetMask = "255.255.255.0";
        private readonly string address = "192.168.0.99";
        private CancellationTokenSource cts = new CancellationTokenSource();
        PlotModel ChartPlot1 = new PlotModel();
        PlotModel ChartPlot2 = new PlotModel();
        Regex regex = new Regex("^-?\\d*(\\.\\d+)?(?=.*\n)");
        Dictionary<int, string> autoIntTimeRetVal = new Dictionary<int, string>();
        private SemaphoreSlim mySemaphoreSlim = new SemaphoreSlim(1);
        string dirpath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "NIRdata", DateTime.Now.ToString("yyyy-MM-dd"));
        double VISrefIntTime = 0;
        double NIRrefIntTime = 0;
        double[][] refSpectra = new double[4][];
        List<double[][]> measSpec = new List<double[][]>();
        StringBuilder FOSStext = new StringBuilder();
        string[] FOSStextArray = new string[] { };
        string FOSSdatetime = "";
        string FOSSappModel = "";
        string FOSScrop = "";
        string FOSSoil = "";
        string FOSSprotein = "";
        string FOSSmoisture = "";
        string FOSSstarch = "";
        string FOSStemp = "";
        string FOSStestWeight = "";


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            //try
            //{
            //    using (var mgr = new UpdateManager(@"\\my.files.iastate.edu\engr\research\ABE\matt-darr\Projects\Moisture\Software\NIR_software\published\NIR_Zeiss_FOSS_GUI\Releases"))
            //    {
            //        await mgr.UpdateApp();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}

            commportslist.ItemsSource = SerialPort.GetPortNames();
            dispatcherTimerFOSSPort = new Timer(new System.Threading.TimerCallback(dispatcherTimerFOSSport_Tick),
                                   new AutoResetEvent(true), 0, 1000);

            //baudrates.ItemsSource = new List<int> { 4800, 9600, 19200, 38400, 57600, 115200, 230400 };
            openSerialPort.Background = Brushes.LightPink;
            commportslist_motor.Background = Brushes.LightPink;

            lastSerRec = DateTime.Now;
            DHCPstatus.Content = "";
            //dispatcherTimerFOSSPort = new System.Windows.Threading.DispatcherTimer();
            //dispatcherTimerFOSSPort.Tick += new EventHandler(dispatcherTimerFOSSport_Tick);
            //dispatcherTimerFOSSPort.Interval = new TimeSpan(0, 0, 1);
            //dispatcherTimerFOSSPort.Start();           

            rawORnormalized.IsEnabled = false;

            updateAdapterList();

            if (!Directory.Exists(dirpath)) Directory.CreateDirectory(dirpath);

            autoIntTimeRetVal.Add(0, "error");
            autoIntTimeRetVal.Add(1, "success");
            autoIntTimeRetVal.Add(2, "energy values are too high (at least one value is greater than optimum+tolerance multiplied by saturation value)");
            autoIntTimeRetVal.Add(3, "energy values are too small (all values are less than optimum-tolerance multiplied by saturation value)");
            autoIntTimeRetVal.Add(4, "energy signal is not stable enough");
            autoIntTimeRetVal.Add(5, "timeout");

            runTimeBox.Text = "10";
            spectraNum.Text = "10";

            //test stuff out
            //var _bigDataPoints = new List<DataPoint>();

            // #x1 double #y double

            //_bigDataPoints.Add(new DataPoint(410.8070281, 4943000.0000000));
            //_bigDataPoints.Add(new DataPoint(432.7935746, 5041000.0000000));
            //_bigDataPoints.Add(new DataPoint(436.8319199, 5059000.0000000));
            //_bigDataPoints.Add(new DataPoint(9918.193582, 47320000.0000000));
            //_bigDataPoints.Add(new DataPoint(10099.91912, 48130000.0000000));
            //_bigDataPoints.Add(new DataPoint(10216.58243, 48650000.0000000));
            //_bigDataPoints.Add(new DataPoint(104904.5616, 470700000.0000000));
            //_bigDataPoints.Add(new DataPoint(107282.6983, 481300000.0000000));
            //_bigDataPoints.Add(new DataPoint(108337.1551, 486000000.0000000));
            //_bigDataPoints.Add(new DataPoint(991388.6853, 4422000128.0000000));
            //_bigDataPoints.Add(new DataPoint(1000362.786, 4462000128.0000000));
            //_bigDataPoints.Add(new DataPoint(1006195.923, 4488000000));

            //var logAxisX = new LogarithmicAxis() { Position = AxisPosition.Bottom, Title = "wavelength [nm]", UseSuperExponentialFormat = false, Base = 10 };
            var linearAxisX1 = new LinearAxis() { Position = AxisPosition.Bottom, Title = "wavelength [nm]", UseSuperExponentialFormat = false, IsZoomEnabled = false, IsPanEnabled = false };
            var linearAxisY1 = new LinearAxis() { Position = AxisPosition.Left, Title = "Reflectance (%)", UseSuperExponentialFormat = false, Angle = -90, IsZoomEnabled = false, IsPanEnabled = false };

            ChartPlot1.Axes.Add(linearAxisY1);
            ChartPlot1.Axes.Add(linearAxisX1);
            ChartPlot1.Series.Clear();
            ChartPlot1.Annotations.Clear();
            // This is the line you're missing
            plotView1.Model = ChartPlot1;
            plotView1.Model.PlotMargins = new OxyPlot.OxyThickness(20, 0, 0, 20);

            //var lineSeriesBigData = new OxyPlot.Series.LineSeries();
            //lineSeriesBigData.Points.AddRange(_bigDataPoints);
            //ChartPlot1.Series.Add(lineSeriesBigData);

            //ChartPlot1.Series.Clear();

            var linearAxisX2 = new LinearAxis() { Position = AxisPosition.Bottom, Title = "wavelength [nm]", UseSuperExponentialFormat = false, IsZoomEnabled = false, IsPanEnabled = false };
            var linearAxisY2 = new LinearAxis() { Position = AxisPosition.Left, Title = "Reflectance (%)", UseSuperExponentialFormat = false, Angle = -90, IsZoomEnabled = false, IsPanEnabled = false };

            ChartPlot2.Axes.Add(linearAxisX2);
            ChartPlot2.Axes.Add(linearAxisY2);
            ChartPlot2.Series.Clear();
            ChartPlot2.Annotations.Clear();
            plotView2.Model = ChartPlot2;
            plotView2.Model.PlotMargins = new OxyPlot.OxyThickness(20, 0, 0, 20);
            //plotView2.Model.LegendPlacement = LegendPlacement.Inside;
            //plotView2.Model.LegendPosition = LegendPosition.LeftTop;

        }

        private void Button_Click_FOSS(object sender, RoutedEventArgs e)
        {
            try
            {

                if (FOSSserialPort == null)
                {
                    FOSSserialPort = new SerialPort();
                    FOSSserialPort.BaudRate = 115200;
                    FOSSserialPort.Parity = Parity.None;
                    FOSSserialPort.StopBits = StopBits.One;
                    FOSSserialPort.DataBits = 8;
                    FOSSserialPort.Handshake = Handshake.XOnXOff;
                    FOSSserialPort.RtsEnable = true;
                    FOSSserialPort.DataReceived += new SerialDataReceivedEventHandler(FOSSDataReceivedHandler);

                }

                if (!FOSSserialPort.IsOpen)
                {
                    FOSSserialPort.PortName = commportslist.SelectedValue.ToString();
                    FOSSserialPort.Open();
                    openSerialPort.Background = Brushes.LightGreen;
                }
                else
                {
                    FOSSserialPort.Close();
                    openSerialPort.Background = Brushes.LightPink;
                }

                //openSerialPort.Background = new SolidColorBrush(Color.FromRgb(1,2,3));

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "comm message err", MessageBoxButton.OK, MessageBoxImage.Error);

            }

        }

        private void dispatcherTimerFOSSport_Tick(object sender)
        {
            Application.Current.Dispatcher.BeginInvoke((Action)(async () =>
           {
               if (DateTime.Now.Subtract(lastWhiteRef).TotalHours >= 1 && DateTime.Now.Subtract(lastWhiteRef).TotalDays < 100)
               {
                   ExternRef.Background = Brushes.LightGray;

               }
               if (DateTime.Now.Subtract(lastDarkRef).TotalHours >= 1 && DateTime.Now.Subtract(lastDarkRef).TotalDays < 100)
               {
                   InternalRef.Background = Brushes.LightGray;
               }

               commportslist.ItemsSource = await Task.Run(() => SerialPort.GetPortNames());
               commportslist_motor.ItemsSource = await Task.Run(() => SerialPort.GetPortNames());
               if (OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.IsComplete &&
               ExternRef.Background == Brushes.LightGreen && InternalRef.Background == Brushes.LightGreen)
               {
                   recordBtn.IsEnabled = true;
                   recordBtn.Background = Brushes.LightGreen;
                   rawORnormalized.IsEnabled = true;
               }
               else
               {
                   recordBtn.Background = Brushes.LightGray;
                   recordBtn.IsEnabled = false;
                   rawORnormalized.IsEnabled = false;
               }

           }));

        }

        private void FOSSDataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();

            if (DateTime.Now.Subtract(lastSerRec).TotalMilliseconds > 3000)
            {
                FOSStext.Clear();
                Application.Current.Dispatcher.BeginInvoke((Action)(() =>
                {
                    FOSSreceived.Document.Blocks.Clear();
                }));
            }


            FOSStext.Append(indata);
            FOSStextArray = FOSStext.ToString().Split(new string[] { "\n", "\r", Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            //Regex rx = new Regex(@"\d+-\d+-\d+\s+\d+:\d+.+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            if (FOSStextArray.Length >= 10)
            {
                Application.Current.Dispatcher.BeginInvoke((Action)(() =>
                {
                    FOSSreceived.Background = Brushes.LightSeaGreen;
                }));
                Regex rxdate;
                Regex rxCrop;
                string[] temp;

                FOSSdatetime = "";
                FOSSappModel = "";
                FOSScrop = "";
                FOSSoil = "";
                FOSSprotein = "";
                FOSSmoisture = "";
                FOSSstarch = "";
                FOSStemp = "";
                FOSStestWeight = "";

                foreach (string item in FOSStextArray.Take(11))
                {
                    rxdate = new Regex(@"\d+-\d+-\d+\s+\d+:\d+.+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    rxCrop = new Regex(@"[a-z0-9]{3,}\s+[a-z]{3,}$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    if (rxdate.IsMatch(item) && !(item.IndexOf("date", StringComparison.CurrentCultureIgnoreCase) > -1)) FOSSdatetime = item;
                    else if (rxCrop.IsMatch(item))
                    {
                        temp = item.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        FOSSappModel = temp[0];
                        foreach (string s in temp.Skip(1))
                        {
                            FOSScrop = FOSScrop + s;
                        }
                    }
                    else if (item.IndexOf("oil", StringComparison.CurrentCultureIgnoreCase) > -1)
                    {
                        temp = item.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        FOSSoil = temp[Array.FindIndex(temp, x => x.IndexOf("DM") > -1 || x.ToLower().IndexOf("%") > -1) + 1];
                    }
                    else if (item.IndexOf("prot", StringComparison.CurrentCultureIgnoreCase) > -1)
                    {
                        temp = item.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        FOSSprotein = temp[Array.FindIndex(temp, x => x.IndexOf("DM") > -1 || x.ToLower().IndexOf("%") > -1) + 1];
                    }
                    else if (item.IndexOf("moisture", StringComparison.CurrentCultureIgnoreCase) > -1)
                    {
                        temp = item.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        FOSSmoisture = temp[Array.FindIndex(temp, x => x.ToLower().IndexOf("moisture") > -1) + 1];
                    }
                    else if (item.IndexOf("Starch", StringComparison.CurrentCultureIgnoreCase) > -1)
                    {
                        temp = item.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        FOSSstarch = temp[Array.FindIndex(temp, x => x.IndexOf("DM") > -1 || x.ToLower().IndexOf("%") > -1) + 1];
                    }
                    else if (item.IndexOf("temp", StringComparison.CurrentCultureIgnoreCase) > -1)
                    {
                        temp = item.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        FOSStemp = temp[Array.FindIndex(temp,x => x.ToLower().IndexOf("temp") > -1) + 1];
                    }
                    else if (item.IndexOf("weight", StringComparison.CurrentCultureIgnoreCase) > -1)
                    {
                        temp = item.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        FOSStestWeight = temp[Array.FindIndex(temp, x => x.ToLower().IndexOf("weight") > -1) + 1];
                    }


                }

                Task.Run(() => fossParsed_Checked(null, null));

            }

            lastSerRec = DateTime.Now;


        }

        private void MOTORDataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            //SerialPort sp = (SerialPort)sender;
            //string indata = sp.ReadExisting();

            //if (DateTime.Now.Subtract(lastSerRec).TotalMilliseconds > 3000)
            //{
            //    Application.Current.Dispatcher.BeginInvoke((Action)(() =>
            //    {
            //        FOSSreceived.Document.Blocks.Clear();
            //    }));
            //}

            //Application.Current.Dispatcher.BeginInvoke((Action)(() =>
            //{
            //    FOSSreceived.AppendText(indata);
            //}));

            //lastSerRec = DateTime.Now;

        }

        void Window_Closing(object sender, CancelEventArgs e)
        {
            //Your code to handle the event
            if (dispatcherTimerFOSSPort != null)
            {
                dispatcherTimerFOSSPort.Dispose();
            }
            if (FOSSserialPort != null)
            {
                FOSSserialPort.Close();
            }
            if (cts.Token.CanBeCanceled) cts.Cancel();

            if (OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized) OsisSdkDemo.Program._equipmentControl.Deinit();
        }

        private void Button_Click_motor(object sender, RoutedEventArgs e)
        {
            try
            {

                if (MOTORserialPort == null)
                {
                    MOTORserialPort = new SerialPort();
                    MOTORserialPort.BaudRate = 9600;
                    MOTORserialPort.Parity = Parity.None;
                    MOTORserialPort.StopBits = StopBits.One;
                    MOTORserialPort.DataBits = 8;
                    MOTORserialPort.Handshake = Handshake.None;
                    MOTORserialPort.RtsEnable = false;
                    MOTORserialPort.DataReceived += new SerialDataReceivedEventHandler(MOTORDataReceivedHandler);

                }

                if (!MOTORserialPort.IsOpen)
                {
                    MOTORserialPort.PortName = commportslist_motor.SelectedValue.ToString();
                    MOTORserialPort.Open();
                    openSerialPort_motor.Background = Brushes.LightGreen;
                }
                else
                {
                    MOTORserialPort.Close();
                    openSerialPort_motor.Background = Brushes.LightPink;
                }

                //openSerialPort.Background = new SolidColorBrush(Color.FromRgb(1,2,3));


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "comm message err", MessageBoxButton.OK, MessageBoxImage.Error);

            }

        }

        private async void Button_Click_ZEISSinit(object sender, RoutedEventArgs e)
        {
            if (OsisSdkDemo.Program._SpectralParameter_JJ == null)
            {
                try
                {
                    int ret = await Task.Run(() => OsisSdkDemo.Program.initializeStuff_JJ());
                    if (ret == 0 && OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters.Count == 2)
                    {
                        if (OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
                        {
                            ZEISSinit.Background = Brushes.LightGreen;
                            ZEISSinfo.Document.Blocks.Clear();
                            ZEISSinfo.AppendText(OsisSdkDemo.Program.ShowBasicSpectralDeviceInfo_JJ());

                            NIRslider.IsEnabled = true;
                            VISslider.IsEnabled = true;
                            VISintTime.IsEnabled = true;
                            NIRintTime.IsEnabled = true;

                            VISmaxADC.IsEnabled = true;
                            NIRmaxADC.IsEnabled = true;

                            VISslider.Minimum = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].IntegrationTime.Min;
                            VISslider.Maximum = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].IntegrationTime.Max;
                            VISslider.Value = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].IntegrationTime.Value;

                            NIRslider.Minimum = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].IntegrationTime.Min;
                            NIRslider.Maximum = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].IntegrationTime.Max;
                            NIRslider.Value = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].IntegrationTime.Value;

                            //await Task.Run(() => OsisSdkDemo.Program.RefreshSpectralDeviceParameter_jj());

                            VISintTime.Text = VISslider.Value.ToString();
                            NIRintTime.Text = NIRslider.Value.ToString();
                        }
                        else ZEISSinit.Background = Brushes.LightPink;
                    }
                    else if (OsisSdkDemo.Program.exceptions_JJ != null)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Clear();
                        foreach (Exception ex in OsisSdkDemo.Program.exceptions_JJ)
                        {
                            sb.Append(ex.ToString());
                            sb.Append("\n");
                        }
                        MessageBox.Show(sb.ToString());
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }


            }

        }


        private void NetAdapters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            updateAdapterList();
        }

        private async void DHCPstatus_Click(object sender, RoutedEventArgs e)
        {

            //EnableDHCP
            if (NetAdapters.SelectedIndex > -1)
            {
                foreach (ManagementObject adapter in queryCollection)
                {
                    string description = adapter["Description"] as string;
                    if (string.Compare(description,
                        NetAdapters.SelectedItem.ToString(), StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        try
                        {
                            if ((bool)adapter["DHCPEnabled"])
                            {
                                // Set DefaultGateway
                                //var newGateway = adapter.GetMethodParameters("SetGateways");
                                //newGateway["DefaultIPGateway"] = new string[] { gateway };
                                //newGateway["GatewayCostMetric"] = new int[] { 1 };

                                // Set IPAddress and Subnet Mask
                                var newAddress = adapter.GetMethodParameters("EnableStatic");

                                newAddress["IPAddress"] = new string[] { address };
                                newAddress["SubnetMask"] = new string[] { subnetMask };

                                await Task.Run(() => adapter.InvokeMethod("EnableStatic", newAddress, null));
                                //adapter.InvokeMethod("SetGateways", newGateway, null);
                            }
                            else await Task.Run(() => adapter.InvokeMethod("EnableDHCP", null, null));

                            updateAdapterList();

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                            //Console.WriteLine("Unable to Set IP : " + ex.Message);
                        }
                    }
                }
            }
        }

        private void updateAdapterList()
        {
            //ObjectQuery query = new ObjectQuery("SELECT * FROM win32_networkadapter where Index = 3");
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_NetworkAdapterConfiguration where IPEnabled = True and not Description like '%virtual%'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            queryCollection = searcher.Get();
            List<string> networkAdapters = new List<string>();
            List<List<string>> adapterProps = new List<List<string>>();
            foreach (ManagementObject adapter in queryCollection)
            {
                ObjectQuery query2 = new ObjectQuery("SELECT * FROM win32_networkadapter where Index = " + ((UInt32)adapter["Index"]).ToString());
                ManagementObjectSearcher searcher2 = new ManagementObjectSearcher(query2);
                ManagementObjectCollection queryCollection2 = searcher2.Get();

                bool netenabled = false;
                List<string> adapterProps_sublist2 = new List<string>();
                foreach (ManagementObject adapter2 in queryCollection2)
                {
                    foreach (PropertyData pd in adapter2.Properties)
                    {
                        adapterProps_sublist2.Add(pd.Name + ": " + (pd.Value ?? "N/A"));
                    }
                    netenabled = (bool)adapter2["NetEnabled"];
                }

                //List<string> adapterProps_sublist = new List<string>();
                //foreach (PropertyData pd in adapter.Properties)
                //{
                //    adapterProps_sublist.Add(pd.Name + ": " + (pd.Value ?? "N/A"));
                //}
                //adapterProps.Add(adapterProps_sublist);

                string description = adapter["Description"] as string;
                networkAdapters.Add(description);
                if (null != NetAdapters.SelectedItem && string.Compare(description,
                    NetAdapters.SelectedItem.ToString(), StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    try
                    {
                        DHCPstatus.Content = (bool)adapter["DHCPEnabled"] ? "DHCP" : "Static";
                        DHCPstatus.Background = netenabled ? Brushes.LightGreen : Brushes.LightPink;

                    }
                    catch { }
                }
            }
            NetAdapters.ItemsSource = networkAdapters;
        }

        private async void recordBtn_Click(object sender, RoutedEventArgs e)
        {


            //check that integration times are correct and FOSS data has been received
            if (VISrefIntTime == VISslider.Value && NIRrefIntTime == NIRslider.Value &&
                OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
            {
                //set timer
                int runTimeSec;
                if (!(int.TryParse(runTimeBox.Text, out runTimeSec)))
                {
                    MessageBox.Show("Run Time of {0} is not a valid entry", runTimeBox.Text);
                    return;
                }
                int numSamples;
                if (!(int.TryParse(spectraNum.Text, out numSamples)))
                {
                    MessageBox.Show("{0} is not a valid # of Spectra Samples", spectraNum.Text);
                    return;
                }

                measSpec.Clear();
                progressRuntime.Value = 0;
                progressSpectra.Value = 0;
                FOSSreceived.Background = Brushes.LightSlateGray;
                //turn on motor if serial port for motor is active and box is checked
                if (await mySemaphoreSlim.WaitAsync(5000))
                {
                    try
                    {
                        if ((bool)motorCheckBox.IsChecked && MOTORserialPort != null && MOTORserialPort.IsOpen)
                        {
                            byte[] bytearr = new byte[2];
                            bytearr[0] = 254;
                            bytearr[1] = 8;
                            MOTORserialPort.Write(bytearr, 0, 2);

                            //function with while loop to take measurements from Zeiss (await task.run function)
                            await Task.Run(() => recordMeasure(runTimeSec, numSamples));
                            //turn off motor if it was turned on
                            bytearr[0] = 254;
                            bytearr[1] = 0;
                            MOTORserialPort.Write(bytearr, 0, 2);

                        }
                        else
                        {
                            //function with while loop to take measurements from Zeiss (await task.run function)
                            await Task.Run(() => recordMeasure(runTimeSec, numSamples));

                        }
                    }
                    finally
                    {
                        mySemaphoreSlim.Release();
                    }

                    //write spectral reflectance, foss data, integration times, reference spectra to file
                    string recordDataPath = System.IO.Path.Combine(dirpath, "measuredSpectra_" + DateTime.Now.ToString("yyyy-MM-dd") + ".csv");
                    StringBuilder sb = new StringBuilder();
                    if (!(File.Exists(recordDataPath)) && measSpec.Count > 0 && IsFileLocked(new FileInfo(recordDataPath)))
                    {
                        using (StreamWriter sw = new StreamWriter(recordDataPath))
                        {
                            sb.Append("FOSSdatetime").Append(",").Append("FOSSappModel").Append(",").Append("FOSScrop").Append(",").Append("FOSSoil");
                            sb.Append(",").Append("FOSSprotein").Append(",").Append("FOSSmoisture").Append(",").Append("FOSSstarch").Append(",").Append("FOSStemp");
                            sb.Append(",").Append("FOSStestWeight").Append(",");
                            sb.Append("VStime").Append(",").Append("VISintegrationTime").Append(",").Append("NIRintegrationTime");
                            foreach (double item in measSpec[0][1])
                            {
                                sb.Append(",").Append(item + "_Norm");
                            }
                            foreach (double item in measSpec[0][3])
                            {
                                sb.Append(",").Append(item + "_VISrawMeas");
                            }
                            foreach (double item in measSpec[0][5])
                            {
                                sb.Append(",").Append(item + "_NIRrawMeas");
                            }
                            foreach (double item in measSpec[0][3])
                            {
                                sb.Append(",").Append(item + "_VISrefWhite");
                            }
                            foreach (double item in measSpec[0][5])
                            {
                                sb.Append(",").Append(item + "_NIRrefWhite");
                            }
                            foreach (double item in measSpec[0][3])
                            {
                                sb.Append(",").Append(item + "_VISrefDark");
                            }
                            foreach (double item in measSpec[0][5])
                            {
                                sb.Append(",").Append(item + "_NIRrefDark");
                            }


                            //need FOSS headers

                            sw.WriteLine(sb);
                        }
                    }

                    using (StreamWriter sw = new StreamWriter(recordDataPath, true))
                    {

                        string mtime = DateTime.Now.ToString("HH:mm:ss");
                        foreach (var supItem in measSpec)
                        {
                            sb.Clear();
                            sb.Append(FOSSdatetime).Append(",").Append(FOSSappModel).Append(",").Append(FOSScrop).Append(",").Append(FOSSoil);
                            sb.Append(",").Append(FOSSprotein).Append(",").Append(FOSSmoisture).Append(",").Append(FOSSstarch).Append(",").Append(FOSStemp);
                            sb.Append(",").Append(FOSStestWeight).Append(",");
                            sb.Append(mtime).Append(",").Append(VISrefIntTime).Append(",").Append(NIRrefIntTime);

                            foreach (double item in supItem[0])
                            {
                                sb.Append(",").Append(item);
                            }
                            foreach (double item in supItem[2])
                            {
                                sb.Append(",").Append(item);
                            }
                            foreach (double item in supItem[4])
                            {
                                sb.Append(",").Append(item);
                            }
                            foreach (var item in refSpectra)
                            {
                                foreach (var subitem in item)
                                {
                                    sb.Append(",").Append(subitem);
                                }
                            }
                            //need FOSS data

                            sw.WriteLine(sb);
                        }
                    }

                    //plot results
                    ChartPlot1.Series.Clear();

                    foreach (var item in measSpec)
                    {
                        var ls = new OxyPlot.Series.LineSeries { };// { Title = "asdf" };
                        foreach (var it in item[0].Zip(item[1], (d, w) => new double[] { d, w }))
                        {
                            ls.Points.Add(new DataPoint(it[1], it[0]));
                        }
                        ChartPlot1.Series.Add(ls);
                    }
                    plotView1.InvalidatePlot(true);

                }
            }
        }

        private async Task<int> recordMeasure(int recTime, int sampleNum)
        {

            int cntr = 0;
            DateTime dt = DateTime.Now;

            while (DateTime.Now.Subtract(dt).TotalSeconds < recTime && cntr < sampleNum)
            {
                await Task.Run(() => OsisSdkDemo.Program.MeasurementForSpectralDevice_JJ(1, 1));
                OSIS.Common.Data.SpectralChannelData[] items = OsisSdkDemo.Program.result_inst.GetSpectralChannelDataArray();
                measSpec.Add(new double[6][]
                    {
                                //normailzed data
                                items[7].Data.Data[0],
                                items[7].Data.WavelengthArray.Data[0],
                                //VIS raw ADC measurements
                                items[0].Data.Data[0],
                                items[0].Data.WavelengthArray.Data[0],
                                //NIR raw ADC measurements
                                items[3].Data.Data[0],
                                items[3].Data.WavelengthArray.Data[0],
                    });
                cntr++;
                Application.Current.Dispatcher.BeginInvoke((Action)(async () =>
                {
                    progressRuntime.Value = (1.0 * DateTime.Now.Subtract(dt).TotalSeconds) / (recTime * 1.0);
                    progressSpectra.Value = (1.0 * cntr) / (sampleNum * 1.0);
                }));

            } //while loop

            return 0;
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MyTabItem1.IsSelected)
            {
                cts.Cancel();

            }

            else if (MyTabItem2.IsSelected)
            {

                cts.Dispose();
                cts = new CancellationTokenSource();
                Task.Run(() => rawMeasureContinuous());

            }
        }

        private async void rawMeasureContinuous()
        {
            while (!cts.Token.IsCancellationRequested)
            {
                await Task.Delay(25);
                if (OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
                {
                    await Application.Current.Dispatcher.BeginInvoke((Action)(async () =>
                    {
                        if (await mySemaphoreSlim.WaitAsync(0))
                        {
                            try
                            {
                                if ((bool)rawORnormalized.IsChecked) await Task.Run(() => OsisSdkDemo.Program.MeasurementForSpectralDevice_JJ());
                                else await Task.Run(() => OsisSdkDemo.Program.MeasurementForSpectralDevice_JJ(1));

                                //FOSSreceived.Document.Blocks.Clear();

                                //OSIS.Common.Data.SpectralChannelData[] items = await Task.Run(() => OsisSdkDemo.Program.result_inst.GetSpectralChannelDataArray());
                                OSIS.Common.Data.SpectralChannelData[] items = OsisSdkDemo.Program.result_inst.GetSpectralChannelDataArray();

                                ChartPlot2.Series.Clear();

                                foreach (var item in items)
                                {
                                    var ls = new OxyPlot.Series.LineSeries { };// { Title = "asdf" };
                                    double[] data = item.Data.Data[0];
                                    double[] wavlen = item.Data.WavelengthArray.Data[0];
                                    foreach (var it in data.Zip(wavlen, (d, w) => new double[] { d, w }))
                                    {
                                        ls.Points.Add(new DataPoint(it[1], it[0]));
                                    }
                                    ChartPlot2.Series.Add(ls);

                                }
                                plotView2.InvalidatePlot(true);

                                if (items.Count() == 2)
                                {
                                    VISmaxADC.Text = (items[0].Data.Data[0].Max() / OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].AdcMax).ToString();
                                    NIRmaxADC.Text = (items[1].Data.Data[0].Max() / OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].AdcMax).ToString();
                                }
                            }
                            finally
                            {
                                mySemaphoreSlim.Release();
                            }

                        }
                    }));

                }
            }
        }

        private async void NIRslider_DragCompleted(object sender, RoutedEventArgs e)
        {
            NIRintTime.Text = NIRslider.Value.ToString();
            if (OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
            {
                OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].IntegrationTime.Value = NIRslider.Value;
                if (await mySemaphoreSlim.WaitAsync(5000))
                {
                    try
                    {
                        await Task.Run(() =>
                    OsisSdkDemo.Program._equipmentControl.SetUnitParameter(OsisSdkDemo.Program._SpectralParameter_JJ));
                        InternalRef.Background = Brushes.LightGray;
                        ExternRef.Background = Brushes.LightGray;
                    }
                    finally
                    {
                        mySemaphoreSlim.Release();
                    }
                    NIRintTime.Text = NIRslider.Value.ToString();
                }
                //else await Task.Run(() => OsisSdkDemo.Program.MeasurementForSpectralDevice_JJ(1));
            }
        }

        private async void VISslider_DragCompleted(object sender, RoutedEventArgs e)
        {
            VISintTime.Text = VISslider.Value.ToString();
            if (OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
            {
                OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].IntegrationTime.Value = VISslider.Value;

                if (await mySemaphoreSlim.WaitAsync(5000))
                {
                    try
                    {
                        await Task.Run(() =>
                    OsisSdkDemo.Program._equipmentControl.SetUnitParameter(OsisSdkDemo.Program._SpectralParameter_JJ));
                        InternalRef.Background = Brushes.LightGray;
                        ExternRef.Background = Brushes.LightGray;
                    }
                    finally
                    {
                        mySemaphoreSlim.Release();
                    }

                    //else await Task.Run(() => OsisSdkDemo.Program.MeasurementForSpectralDevice_JJ(1));
                }
            }

        }

        private async void VISintTime_TextChanged(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (await mySemaphoreSlim.WaitAsync(5000))
                {
                    try
                    {
                        if (OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
                        {
                            double number;
                            if (double.TryParse(VISintTime.Text, out number))
                            {
                                if (number < VISslider.Minimum)
                                {
                                    VISintTime.Text = VISslider.Minimum.ToString();
                                }
                                else if (number > VISslider.Maximum)
                                {
                                    VISintTime.Text = VISslider.Maximum.ToString();
                                }
                                else
                                {
                                    OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].IntegrationTime.Value = number;
                                    VISslider.Value = number;

                                    await Task.Run(() =>
                                    OsisSdkDemo.Program._equipmentControl.SetUnitParameter(OsisSdkDemo.Program._SpectralParameter_JJ));

                                    InternalRef.Background = Brushes.LightGray;
                                    ExternRef.Background = Brushes.LightGray;

                                    //await Task.Run(() => OsisSdkDemo.Program.RefreshSpectralDeviceParameter_jj());
                                }
                            }
                            else
                            {
                                VISintTime.Text = VISslider.Value.ToString();
                            }
                        }
                    }
                    finally
                    {
                        mySemaphoreSlim.Release();
                    }
                }
            }

        }

        private async void NIRintTime_TextChanged(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {

                if (await mySemaphoreSlim.WaitAsync(5000))
                {
                    try
                    {
                        if (OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
                        {
                            double number;

                            //Match m = regex.Match(NIRintTime.Text);
                            if (double.TryParse(NIRintTime.Text, out number))
                            {
                                if (number < NIRslider.Minimum)
                                {
                                    NIRintTime.Text = NIRslider.Minimum.ToString();
                                }
                                else if (number > NIRslider.Maximum)
                                {
                                    NIRintTime.Text = NIRslider.Maximum.ToString();
                                }
                                else
                                {
                                    OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].IntegrationTime.Value = number;
                                    NIRslider.Value = number;

                                    await Task.Run(() =>
                                    OsisSdkDemo.Program._equipmentControl.SetUnitParameter(OsisSdkDemo.Program._SpectralParameter_JJ));

                                    InternalRef.Background = Brushes.LightGray;
                                    ExternRef.Background = Brushes.LightGray;

                                    //await Task.Run(() => OsisSdkDemo.Program.RefreshSpectralDeviceParameter_jj());
                                }
                            }
                            else
                            {
                                NIRintTime.Text = NIRslider.Value.ToString();
                            }
                        }
                    }
                    finally
                    {
                        mySemaphoreSlim.Release();
                    }
                }
            }
        }

        private async void autoCal_Click(object sender, RoutedEventArgs e)
        {

            if (OsisSdkDemo.Program._SpectralParameter_JJ == null || !OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
                return;

            if (MessageBox.Show("Prepare external maximum reference (i.e. white standard or mirror) ...",
                "Automatic Integration Time Calculation", MessageBoxButton.OKCancel, MessageBoxImage.None) == MessageBoxResult.OK)
            {
                try
                {
                    if (await mySemaphoreSlim.WaitAsync(5000))
                    {
                        try
                        {
                            int ret = await Task.Run(() => OsisSdkDemo.Program.DetermineIntegrationTimeForSpectralDevice_jj());
                            if (ret == 1)
                            {
                                VISslider.Value = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].IntegrationTime.Value;
                                NIRslider.Value = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].IntegrationTime.Value;
                                VISintTime.Text = VISslider.Value.ToString();
                                NIRintTime.Text = NIRslider.Value.ToString();
                                InternalRef.Background = Brushes.LightGray;
                                ExternRef.Background = Brushes.LightGray;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                        finally
                        {
                            mySemaphoreSlim.Release();
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            }
        }

        private async void InternalRef_Click(object sender, RoutedEventArgs e)
        {
            rawORnormalized.IsChecked = true;
            InternalRef.Background = Brushes.LightGray;
            if (OsisSdkDemo.Program._SpectralParameter_JJ == null || !OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
                return;

            if (MessageBox.Show("Prepare Dark reference (i.e. cover lens) ...",
                "Reference Calibraiton Measurement", MessageBoxButton.OKCancel, MessageBoxImage.None) == MessageBoxResult.OK)
            {
                try
                {
                    if (ExternRef.Background == Brushes.LightGreen)
                    {
                        if (await mySemaphoreSlim.WaitAsync(5000))
                        {
                            try
                            {
                                string ret = await Task.Run(() => OsisSdkDemo.Program.InternalReferenceMeasurementsForSpectralDevice_JJ());

                                if (ret == null)
                                {
                                    InternalRef.Background = Brushes.LightGreen;

                                    string retstr = await Task.Run(() => OsisSdkDemo.Program.LoadSaveReferences_JJ(dirpath));
                                    if (retstr != null)
                                    {
                                        MessageBox.Show(retstr);
                                        InternalRef.Background = Brushes.LightGray;
                                        ExternRef.Background = Brushes.LightGray;
                                    }
                                    else
                                    {
                                        lastDarkRef = DateTime.Now;

                                        //get raw reference spectra
                                        OSIS.Common.Data.Spectrum[,] spectrum = OsisSdkDemo.Program._equipmentControl.GetReferencesAsSpectra(0);

                                        //dark
                                        refSpectra[0] = spectrum[0, 1].Data[0];
                                        refSpectra[1] = spectrum[1, 1].Data[0];
                                        //dark
                                        refSpectra[2] = spectrum[0, 0].Data[0];
                                        refSpectra[3] = spectrum[1, 0].Data[0];

                                        //double[] wavlen = item.Data.WavelengthArray.Data[0];

                                        bool unsaferef = false;

                                        await Task.Run(() => OsisSdkDemo.Program.MeasurementForSpectralDevice_JJ(1));

                                        //FOSSreceived.Document.Blocks.Clear();

                                        //OSIS.Common.Data.SpectralChannelData[] items = await Task.Run(() => OsisSdkDemo.Program.result_inst.GetSpectralChannelDataArray());
                                        OSIS.Common.Data.SpectralChannelData[] items = OsisSdkDemo.Program.result_inst.GetSpectralChannelDataArray();
                                        if (items[0].Data.ExtendedInfo.IsUnsafe)
                                        {
                                            ZEISSinfo.Document.Blocks.Clear();

                                            foreach (var unsafeEx in items[0].Data.ExtendedInfo.UnsafeFlags)
                                            {
                                                if (!(unsafeEx.Value == OSIS.Common.Constants.SpectralStatusFlags.Unsafe_NettoTooLowInInterpolationRange ||
                                                    unsafeEx.Value == OSIS.Common.Constants.SpectralStatusFlags.Unsafe_MissingWhiteCertificate))
                                                {
                                                    unsaferef = true;
                                                    ZEISSinfo.AppendText(unsafeEx.Key.ToString() + ":  \n");
                                                    ZEISSinfo.AppendText(unsafeEx.Value.ToString() + ":  \n");
                                                }
                                            }
                                            //InternalRef.Background = Brushes.LightGray;
                                            //ExternRef.Background = Brushes.LightGray;
                                        }
                                        if (unsaferef)
                                        {
                                            ZEISSinfo.Background = Brushes.LightCoral;
                                        }
                                        else
                                        {
                                            ZEISSinfo.Background = Brushes.LightSteelBlue;
                                            ZEISSinfo.Document.Blocks.Clear();
                                            ZEISSinfo.AppendText(OsisSdkDemo.Program.ShowBasicSpectralDeviceInfo_JJ());
                                        }
                                    }
                                }
                                else
                                {
                                    InternalRef.Background = Brushes.LightGray;
                                    MessageBox.Show(ret);
                                }
                            }
                            finally
                            {
                                mySemaphoreSlim.Release();
                            }
                        }
                    }
                    else MessageBox.Show("Perform White Reference First");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private async void ExternRef_Click(object sender, RoutedEventArgs e)
        {
            rawORnormalized.IsChecked = true;
            InternalRef.Background = Brushes.LightGray;
            ExternRef.Background = Brushes.LightGray;
            if (OsisSdkDemo.Program._SpectralParameter_JJ == null || !OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
                return;

            if (MessageBox.Show("Prepare external maximum reference (i.e. white standard) ...",
                "Automatic Integration Time Calculation", MessageBoxButton.OKCancel, MessageBoxImage.None) == MessageBoxResult.OK)
            {
                try
                {
                    if (await mySemaphoreSlim.WaitAsync(5000))
                    {
                        try
                        {
                            int ret = await Task.Run(() => OsisSdkDemo.Program.DetermineIntegrationTimeForSpectralDevice_jj());
                            if (ret == 1)
                            {

                                VISslider.Value = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].IntegrationTime.Value;
                                NIRslider.Value = OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].IntegrationTime.Value;
                                VISintTime.Text = VISslider.Value.ToString();
                                NIRintTime.Text = NIRslider.Value.ToString();

                                VISrefIntTime = VISslider.Value;
                                NIRrefIntTime = NIRslider.Value;

                                string retval = await Task.Run(() => OsisSdkDemo.Program.ExternalReferenceMeasurementsForSpectralDevice_JJ());

                                if (retval == null)
                                {
                                    ExternRef.Background = Brushes.LightGreen;
                                    lastWhiteRef = DateTime.Now;
                                    //get raw reference spectra

                                }
                                else
                                {
                                    ExternRef.Background = Brushes.LightGray;
                                    MessageBox.Show(retval);
                                }
                            }
                        }
                        finally
                        {
                            mySemaphoreSlim.Release();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void fossParsed_Checked(object sender, RoutedEventArgs e)
        {
            Application.Current.Dispatcher.BeginInvoke((Action)(() =>
            {
                if ((bool)fossParsed.IsChecked)
                {
                    FOSSreceived.Document.Blocks.Clear();

                    FOSSreceived.AppendText(FOSSdatetime + "\n\n");
                    FOSSreceived.AppendText(FOSSappModel + "\n");
                    FOSSreceived.AppendText(FOSScrop + "\n");
                    FOSSreceived.AppendText("Oil%: " + FOSSoil + "\n");
                    FOSSreceived.AppendText("Protein%: " + FOSSprotein + "\n");
                    FOSSreceived.AppendText("Moisture%: " + FOSSmoisture + "\n");
                    FOSSreceived.AppendText("Starch%: " + FOSSstarch + "\n");
                    FOSSreceived.AppendText("Temp (C): " + FOSStemp + "\n");
                    FOSSreceived.AppendText("Test Weight: " + FOSStestWeight + "\n");
                }
                else
                {
                    FOSSreceived.Document.Blocks.Clear();
                    foreach (string s in FOSStextArray.Take(11))
                    {
                        FOSSreceived.AppendText(s + "\n");
                    }
                }
            }));
        }

        private async void cleanout_Click(object sender, RoutedEventArgs e)
        {
            if (MOTORserialPort != null && MOTORserialPort.IsOpen)
            {
                int runTimeSec;
                if (!(int.TryParse(runTimeBox.Text, out runTimeSec)))
                {
                    runTimeSec = 100000;
                }
                else
                {
                    runTimeSec = 1000 * runTimeSec;
                }
                byte[] bytearr = new byte[2];
                bytearr[0] = 254;
                bytearr[1] = 8;
                MOTORserialPort.Write(bytearr, 0, 2);

                //function with while loop to take measurements from Zeiss (await task.run function)
                await Task.Delay(runTimeSec);
                //turn off motor if it was turned on
                bytearr[0] = 254;
                bytearr[1] = 0;
                MOTORserialPort.Write(bytearr, 0, 2);
            }
        }


        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        private async void loadRef_Click(object sender, RoutedEventArgs e)
        {
            if (OsisSdkDemo.Program._SpectralParameter_JJ != null && OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters.Count == 2)
            {
                if (OsisSdkDemo.Program._SpectralParameter_JJ.IsInitialized)
                {
                    // Create OpenFileDialog 
                    Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

                    // Set filter for file extension and default file extension 
                    dlg.DefaultExt = ".oscx";
                    dlg.Filter = "OSCX Files (*.oscx)|*.oscx";


                    // Display OpenFileDialog by calling ShowDialog method 
                    Nullable<bool> result = dlg.ShowDialog();


                    // Get the selected file name and display in a TextBox 
                    if (result == true)
                    {
                        string filename = dlg.FileName;
                        string retstr = null;
                        if (await mySemaphoreSlim.WaitAsync(10000))
                        {
                            try
                            {
                                retstr = await Task.Run(() => OsisSdkDemo.Program.LoadSaveReferences_JJ(dirpath, filename));
                            }
                            finally
                            {
                                mySemaphoreSlim.Release();
                            }

                            // Open document 

                            if (retstr != null)
                            {
                                MessageBox.Show(retstr);
                                InternalRef.Background = Brushes.LightGray;
                                ExternRef.Background = Brushes.LightGray;
                            }
                            else
                            {
                                Regex VISintRG = new Regex(@"(?<=(VIS))\d+(\.\d{1,2})?", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                                Regex NIRintRG = new Regex(@"(?<=(NIR))\d+(\.\d{1,2})?", RegexOptions.Compiled | RegexOptions.IgnoreCase);

                                Match mvis = VISintRG.Match(filename);
                                Match mnir = NIRintRG.Match(filename);


                                if (mvis.Success && mnir.Success)
                                {
                                    //get raw reference spectra
                                    OSIS.Common.Data.Spectrum[,] spectrum = OsisSdkDemo.Program._equipmentControl.GetReferencesAsSpectra(0);

                                    //dark
                                    refSpectra[0] = spectrum[0, 1].Data[0];
                                    refSpectra[1] = spectrum[1, 1].Data[0];
                                    //dark
                                    refSpectra[2] = spectrum[0, 0].Data[0];
                                    refSpectra[3] = spectrum[1, 0].Data[0];

                                    double dnumVIS;
                                    double dnumNIR;

                                    if (!(double.TryParse(mvis.ToString(), out dnumVIS)))
                                    {
                                        dnumVIS = 20;
                                        MessageBox.Show("No Valid Integration Time for VIS parsed");
                                    }
                                    if (!(double.TryParse(mnir.ToString(), out dnumNIR)))
                                    {
                                        dnumNIR = 8;
                                        MessageBox.Show("No Valid Integration Time for NIR parsed");
                                    }

                                    //integration times
                                    VISslider.Value = dnumVIS;
                                    NIRslider.Value = dnumNIR;
                                    VISintTime.Text = dnumVIS.ToString();
                                    NIRintTime.Text = dnumNIR.ToString();

                                    VISrefIntTime = dnumVIS;
                                    NIRrefIntTime = dnumNIR;

                                    //double[] wavlen = item.Data.WavelengthArray.Data[0];

                                    OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[0].IntegrationTime.Value = VISslider.Value;
                                    OsisSdkDemo.Program._SpectralParameter_JJ.ModuleParameters[1].IntegrationTime.Value = NIRslider.Value;

                                    if (await mySemaphoreSlim.WaitAsync(10000))
                                    {
                                        try
                                        {
                                            await Task.Run(() =>
                                        OsisSdkDemo.Program._equipmentControl.SetUnitParameter(OsisSdkDemo.Program._SpectralParameter_JJ));
                                            if (dnumNIR > 8 && dnumVIS > 20)
                                            {
                                                lastDarkRef = DateTime.Now;
                                                lastWhiteRef = DateTime.Now;
                                                InternalRef.Background = Brushes.LightGreen;
                                                ExternRef.Background = Brushes.LightGreen;
                                            }
                                        }
                                        finally
                                        {
                                            mySemaphoreSlim.Release();
                                        }

                                        //else await Task.Run(() => OsisSdkDemo.Program.MeasurementForSpectralDevice_JJ(1));
                                    }

                                }
                            }
                        }
                    }
                }
            }
        }
    }

}
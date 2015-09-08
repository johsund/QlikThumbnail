using Qlik.Engine;
using Qlik.Sense.Client;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace QlikThumbnail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            Rectangle resolution = Screen.PrimaryScreen.Bounds;

            //ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
            
            
        }

        public static void IgnoreBadCertificates()
        {
            ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(AcceptAllCertifications);
        }
        private static bool AcceptAllCertifications(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certification, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        public static ILocation ConnectQlikEngine(string qlikHost)
        {
            IgnoreBadCertificates();
            //Connect to Qlik Engine API using Uri
            Uri uri = new Uri(qlikHost);
            ILocation location = Qlik.Engine.Location.FromUri(uri);
            location.AsNtlmUserViaProxy(proxyUsesSsl: true);

            return location;
        }

        public static string GetFQDN()
        {
            string domainName = IPGlobalProperties.GetIPGlobalProperties().DomainName;
            string hostName = Dns.GetHostName();

            if (!hostName.EndsWith(domainName))  // if hostname does not already include domain name
            {
                hostName += "." + domainName;   // add the domain name part
            }

            return hostName;                    // return the fully qualified name
        }

        public static CookieContainer ConnectQRSApi(string qlikHost)
        {

            //Form1.textBox5.Text = "Connecting to Qlik Repository API" + Environment.NewLine + textBox5.Text;
            //Connect to the Qlik Repository API
            //Create the HTTP Request and add required headers and content in xrfkey
            string xrfkey = "0123456789abcdef";

            // Create Cookie Container
            CookieContainer QRSCookieContainer = new CookieContainer();


            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(@"" + qlikHost + "/QRS/app?xrfkey=" + xrfkey);
            request.Method = "GET";
            request.UserAgent = "Windows";
            request.Accept = "application/json";
            request.Headers.Add("X-Qlik-xrfkey", xrfkey);
            // specify to run as the current Microsoft Windows user
            request.UseDefaultCredentials = true;

            request.CookieContainer = QRSCookieContainer;

            // make the web request and return the content
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream stream = response.GetResponseStream();

            foreach (Cookie myCookie in QRSCookieContainer.GetCookies(request.RequestUri))
            {
                QRSCookieContainer.Add(myCookie);
                //Console.WriteLine(myCookie.Name + myCookie.Value);
            }
            return QRSCookieContainer;
        }

        public static IAppIdentifier[] ListAllApps(ILocation location)
        {

            //Create array for apps
            IAppIdentifier[] array = new IAppIdentifier[location.GetAppIdentifiers().Count()];

            int i = 0;

            foreach (IAppIdentifier appIdentifier in location.GetAppIdentifiers())
            {
                //Store each app into a new array slot
                IAppIdentifier appCurrent = appIdentifier;
                array[i] = appCurrent;
                i++;
            }

            return array;

        }

        public static ISheet[] ListAllSheets(IAppIdentifier inputApp, ILocation location)
        {

            IApp appCurrent = location.App(inputApp);
            //Create a new array for sheets within an app
            ISheet[] array = new ISheet[appCurrent.GetSheets().Count()];

            int i = 0;

            foreach (Sheet appSheet in appCurrent.GetSheets())
            {
                //Store each sheet into a new array slot
                array[i] = appSheet;
                i++;
            }

            return array;
        }

        public static byte[] imageToByteArray(System.Drawing.Image imageIn)
        {
            MemoryStream ms = new MemoryStream();
            imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            return ms.ToArray();
        }
        public static Image CaptureApplication(Process procName)
        {
            //var proc = Process.GetProcessesByName(procName)[0]; //procName
            var proc = procName;
            var rect = new User32.Rect();

            User32.GetWindowRect(proc.MainWindowHandle, ref rect);

            int width = rect.right - rect.left-60;
            int height = rect.bottom - rect.top - 130;

            var bmp = new Bitmap(width, height); //, PixelFormat.Format32bppArgb);
            Graphics graphics = Graphics.FromImage(bmp);
            graphics.CopyFromScreen(rect.left+30, rect.top+100, 0, 0, new System.Drawing.Size(width, height), CopyPixelOperation.SourceCopy);

            var newWidth = 146;
            var newHeight = 96;

            //if (System.IO.File.Exists(textBox2.Text + Filename + ".png"))
             //   System.IO.File.Delete(textBox2.Text + Filename + ".png");

            Bitmap newImage = new Bitmap(newWidth, newHeight);
            using (Graphics gr = Graphics.FromImage(newImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(bmp, new Rectangle(0, 0, newWidth, newHeight));
            }


            //newImage.Save(textBox2.Text + Filename + ".png", ImageFormat.Png);

            //GlobalImage = newImage;
            
            return newImage;

        }

        //private void LoadPage(string URLPath)
        //{
        //    string webAddressString = URLPath;
        //    Uri webAddress;
        //    if (Uri.TryCreate(webAddressString, UriKind.Absolute, out webAddress))
        //    {
        //        webBrowser1.Navigate(webAddress);
        //    }
        //}





        public void UploadAndStore(Image myImageStream, Sheet mySheet, CookieContainer QRSCookieContainer, IAppIdentifier myApp)
        {
            //textBox4.Text = "Fetching list of Apps" + Environment.NewLine + textBox4.Text;
            //string Filename = textBox3.Text + myFile;

            //textBox5.Text = "Uploading " + Filename + Environment.NewLine + textBox5.Text;

            string myFile = myApp.AppName + "_" + mySheet.Id;
            //MessageBox.Show(myFile);

            string xrfkey = "0123456789abcdef";

            HttpWebRequest postImageRequest = (HttpWebRequest)WebRequest.Create(@"" + textBox1.Text + "/QRS/contentlibrary/Default/uploadfile?externalpath=" + myFile + ".png&overwrite=true&xrfkey=" + xrfkey);

            postImageRequest.Method = "POST";
            postImageRequest.Accept = "application/json";
            postImageRequest.UserAgent = "Windows";
            postImageRequest.Headers.Add("X-Qlik-xrfkey", xrfkey);
            postImageRequest.UseDefaultCredentials = true;
            postImageRequest.CookieContainer = QRSCookieContainer;
            //postImageRequest.ContentLength = 0;
            postImageRequest.ContentType = "image/png";

            //MessageBox.Show(Filename);

            System.Drawing.Image myImage = myImageStream;
            byte[] imageByte = imageToByteArray(myImage);
            long length = imageByte.Length;
            postImageRequest.ContentLength = length;



            using (Stream postStream = postImageRequest.GetRequestStream())
            {
                postStream.Write(imageByte, 0, (int)length);
            }

            try
            {
                HttpWebResponse postImageResponse = (HttpWebResponse)postImageRequest.GetResponse();
            }
            catch (WebException ex)
            {
                string message = ((System.Net.HttpWebResponse)(ex.Response)).StatusDescription;
                throw new Exception(message);
            }
        }

        public void AssignThumbnail(Sheet mySheet, IAppIdentifier myApp )
        {
            GenericObjectProperties xyz = mySheet.GetProperties();

            mySheet.Properties.Thumbnail.StaticContentUrlDef.Url = "/content/Default/" + myApp.AppName + "_" + mySheet.Id + ".png";

            mySheet.SetProperties(xyz);
        }

        private class User32
        {
            [StructLayout(LayoutKind.Sequential)]
            public struct Rect
            {
                public int left;
                public int top;
                public int right;
                public int bottom;
            }

            [DllImport("user32.dll")]
            public static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);
        }

        private void SelectAllCheckBoxes(bool CheckThem, CheckedListBox clb)
        {
            for (int i = 0; i <= (clb.Items.Count - 1); i++)
            {
                if (CheckThem)
                {
                    clb.SetItemCheckState(i, CheckState.Checked);
                }
                else
                {
                    clb.SetItemCheckState(i, CheckState.Unchecked);
                }
            }
        }

        private void PopulateListBox(CheckedListBox clb, string Folder, string FileType)
        {
            DirectoryInfo dinfo = new DirectoryInfo(Folder);
            FileInfo[] Files = dinfo.GetFiles(FileType);
            foreach (FileInfo file in Files)
            {
                clb.Items.Add(file.Name);
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            textBox4.Text = "";
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Clear the checkedListBox1
            checkedListBox1.Items.Clear();


            textBox4.Text = "Connecting to Qlik Engine API" + Environment.NewLine + textBox4.Text;
            //Connect to Qlik Sense
            ILocation location = ConnectQlikEngine(textBox1.Text);

            textBox4.Text = "Fetching list of Apps" + Environment.NewLine + textBox4.Text;
            //This section fetches all apps and pushes them to checkedListBox1
            IAppIdentifier[] myAppArray = ListAllApps(location);

            

            foreach (IAppIdentifier myApp in myAppArray)
            {
                
                checkedListBox1.Items.Add(myApp.AppName);

            }



        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox4.Text = "Preparing Thumbnail creation" + Environment.NewLine + textBox4.Text;
            //Connect to Qlik Sense
            ILocation location = ConnectQlikEngine(textBox1.Text);

            //This section fetches all apps & sheets and pushes them to treeView1
            IAppIdentifier[] myAppArray = ListAllApps(location);

            // Create string array of checked apps
            String[] selectedApps = new string[checkedListBox1.CheckedItems.Count];
            Int32 counter = 0;
            foreach (object item in this.checkedListBox1.CheckedItems)
            {
                String temp = Convert.ToString(item);
                selectedApps[counter] = temp;
                counter++;
            }

            CookieContainer myCookies = ConnectQRSApi(textBox1.Text);
            //MessageBox.Show("Connected to QRS API");
            textBox4.Text = "Connected to QRS API" + Environment.NewLine + textBox4.Text;

            foreach (IAppIdentifier myApp in myAppArray)
            {


                if(Array.IndexOf(selectedApps, myApp.AppName)>-1)
                {


                    IApp appCurrent = location.App(myApp);
                    foreach (Sheet appSheet in appCurrent.GetSheets())
                    {
                        textBox4.Text = "Launching " + myApp.AppName +"_" + appSheet.MetaAttributes.Title + Environment.NewLine + textBox4.Text;
                        //Launch each sheet - 1 at a time and capture a screenshot + resize it to thumbnail size.
                        Process proc = Process.Start("iexplore", "-k "+ textBox1.Text +"/sense/app/" + myApp.AppId + "/sheet/" + appSheet.Id + "/state/analysis");

                        

                        string Filename = myApp.AppName + "_" + appSheet.Id;

                        //Pause while IE loads and wait for the sheet to finish rendering
                        Thread.Sleep(25000);

                        textBox4.Text = "Assigning thumbnail" + Environment.NewLine + textBox4.Text;
                        //Capture screenshot
                        if (appSheet.Id.Length > 0)
                        {

                            //LoadPage(textBox1.Text + "/sense/app/" + myApp.AppId + "/sheet/" + appSheet.Id + "/state/analysis");
                            Image myImage = CaptureApplication(proc);

                            //while (webBrowser1.ReadyState == WebBrowserReadyState.Complete)
                            //{ }

                            imageList1.Images.Add(myImage);

                            int picCount = imageList1.Images.Count;
                            pictureBox1.Image = imageList1.Images[picCount-1];
                            label4.Text = " of " + picCount;
                            label5.Text = picCount.ToString();

                            //Image myImage = CaptureApplication2(textBox1.Text + "/sense/app/" + myApp.AppId + "/sheet/" + appSheet.Id + "/state/analysis", webBrowser1);

                            //MessageBox.Show("Thumbnail captured");
                            textBox4.Text = "Thumbnail captured" + Environment.NewLine + textBox4.Text;
                            UploadAndStore(myImage, appSheet, myCookies, myApp);
                            //MessageBox.Show("Thumbnail stored in Content Library");
                            textBox4.Text = "Thumbnail stored in Content Library" + Environment.NewLine + textBox4.Text;
                            AssignThumbnail(appSheet, myApp);
                            //MessageBox.Show("Thumbnail assigned");
                            textBox4.Text = "Thumbnail assigned" + Environment.NewLine + textBox4.Text;
                            
                        }
                        //proc.Kill();
                        //proc.WaitForExit();


                        //Ensure IE is killed between screenshots
                        foreach (var process in Process.GetProcessesByName(Path.GetFileNameWithoutExtension("iexplore.exe")))
                        {
                            if (process.MainWindowTitle.Length > 0)
                            {
                                process.Kill();
                                process.WaitForExit();
                            }
                        }


                        //Wait a few more seconds before launching the next iteration.
                        Thread.Sleep(1000);
                    }
                }

            }

            //textBox5.Text = "COMPLETED: All thumbnails created" + Environment.NewLine + textBox5.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SelectAllCheckBoxes(true, checkedListBox1);

        }



        private void button7_Click(object sender, EventArgs e)
        {
            textBox4.Text = "Testing connection to Qlik Engine API and Qlik Repository API..." + Environment.NewLine;
            try
            {
                ILocation location = ConnectQlikEngine(textBox1.Text);
                location.GetAppIdentifiers();

                textBox4.Text = textBox4.Text + "Connection to Qlik Engine API successful." + Environment.NewLine ;
            }
            catch (Exception Ex)
            {
                textBox4.Text = textBox4.Text + "Connection to Qlik Engine API Failed. Error message: " + Ex.Message + Environment.NewLine;
            }

            try {
                //Connect to the Qlik Repository API
                //Create the HTTP Request and add required headers and content in xrfkey
                string xrfkey = "0123456789abcdef";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(@"" + textBox1.Text + "/QRS/app?xrfkey=" + xrfkey);
                request.Method = "GET";
                request.UserAgent = "Windows";
                request.Accept = "application/json";
                request.Headers.Add("X-Qlik-xrfkey", xrfkey);
                // specify to run as the current Microsoft Windows user
                request.UseDefaultCredentials = true;

                // make the web request and return the content
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream stream = response.GetResponseStream();

                textBox4.Text = textBox4.Text + "Connection to Qlik Repository API successful." + Environment.NewLine;

            }

            catch (Exception Ex)
            {
                textBox4.Text = textBox4.Text + "Connection to Qlik Repository API Failed. Error message: " + Ex.Message + Environment.NewLine;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = "https://" + GetFQDN();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int picCount = imageList1.Images.Count;
            if (Convert.ToInt32(label5.Text) - 1 > 0)
            {
                label5.Text = (Convert.ToInt32(label5.Text) - 1).ToString();
                pictureBox1.Image = imageList1.Images[Convert.ToInt32(label5.Text) - 1];
                
            }
            label4.Text = " of " + picCount;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int picCount = imageList1.Images.Count;

            if (Convert.ToInt32(label5.Text) + 1 <= picCount)
            {
                label5.Text = (Convert.ToInt32(label5.Text) + 1).ToString();
                pictureBox1.Image = imageList1.Images[Convert.ToInt32(label5.Text) - 1];
                
            }
            label4.Text = " of " + picCount;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SelectAllCheckBoxes(false, checkedListBox1);
        }
    }
}

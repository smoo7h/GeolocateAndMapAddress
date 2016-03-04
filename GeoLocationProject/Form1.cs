using DevExpress.Spreadsheet;
using DevExpress.XtraTreeList;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GeoLocationProject
{
    public partial class Form1 : Form
    {
        int runningCount = 0;
        Style styleWarning;
        int totalCount = 0;
        int currentCount = 0;

        string location = "";

        public string apiKey = "";
        int keyIndex = 0;

        List<GeoCode> mapList = new List<GeoCode>();

        ProgressForm progForm; //= new ProgressForm();

        public Form1()
        {
            InitializeComponent();
            styleWarning = this.spreadsheetControl1.Document.Styles[BuiltInStyleId.Accent2_60percent];

            ChangeKey();
           
        }

       

        void bgworker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            runningCount--;
            currentCount++;

            progForm.progressBar1.Value = currentCount - 1;
            GeoCode geo = e.Result as GeoCode;

            if (geo.Latitude != 0 && geo.Longitude != 0)
            {
                mapList.Add(geo);
                this.spreadsheetControl1.ActiveWorksheet.Cells[geo.index, 1].Value = geo.Latitude.ToString();
                this.spreadsheetControl1.ActiveWorksheet.Cells[geo.index, 2].Value = geo.Longitude.ToString();
                this.spreadsheetControl1.ActiveWorksheet.Cells[geo.index, 3].Value = geo.FullAddress.ToString();            
            }
            else
            {

                this.spreadsheetControl1.ActiveWorksheet.Cells[geo.index, 0].Style = styleWarning;

                if (geo.FullAddress == "API Key Full")
                {
                 //   MessageBox.Show("API Key Used UP Try Agin Tomorrow Or Change Key");
                }
            }

            if (currentCount == totalCount + 1)
            {
                progForm.Close();
                CreateMap(mapList);
            }

        }

        void bgworker_DoWork(object sender, DoWorkEventArgs e)
        {
           // currentCount++;

            if (runningCount > 5)
            {
               // Thread.Sleep(3000);
            }

            Address add = e.Argument as Address;
            GeoCode geo = GetLongAndLat(add.AddressName, add.Location, add.index);
            geo.Name = add.AddressName;
            e.Result = geo;
            
        
        }

        public GeoCode GetLongAndLat(string address , string Location, int index)
        {
            //toosmooth key AIzaSyCQz1oqJmJrzVrvMllR95dwhAAoup8IopU       AIzaSyAkDqc5_AK6b2FRh3DcH2QkvGYo1rP1TO0
            //gs key        AIzaSyBG6Rd7tKOlrLSHyAnVrBujwS1BUNhePmY;
            Thread.Sleep(3000);
            //reformat name
            string formattedName = address.Trim().Replace(" ", "+");
            Location = Location.Trim().Replace(" ", "+");
            formattedName = formattedName + "+" + Location;
            
            //get json txt
            GeoCode geocode = new GeoCode();

            string url = @"https://maps.googleapis.com/maps/api/place/textsearch/json?query=" + formattedName + @"&key=" + apiKey; ;

            string jsontxt;//AIzaSyBLAlQLaDjTN9vk12oKetgSW_lI3lEjRx0

            try
            {
                using (WebClient wc = new WebClient())
                {
                    jsontxt = wc.DownloadString(url);
                }

          
                if (jsontxt.Contains("ZERO_RESULTS") == false && jsontxt.Contains("error") == false)
                {
                    JToken token = JObject.Parse(jsontxt);

                    double lat = (double)token.SelectToken("results[0].geometry.location.lat");
                    double lon = (double)token.SelectToken("results[0].geometry.location.lng");
                    string fulladdress = (string)token.SelectToken("results[0].formatted_address");

              
                    geocode.Latitude = lat;
                    geocode.Longitude = lon;
                    geocode.index = index;
                    geocode.FullAddress = fulladdress;

                }
                else
                {
                    if (jsontxt.Contains("error") == true)
                    {
                        geocode.Latitude = 0;
                        geocode.Longitude = 0;
                        geocode.index = index;
                        geocode.FullAddress = "API Key Full";
                    }
                    else
                    {
                        geocode.Latitude = 0;
                        geocode.Longitude = 0;
                        geocode.index = index;
                        geocode.FullAddress = "";
                    
                    }
             
                   
                }
            }
            catch (Exception e)
            {

                geocode.Latitude = 0;
                geocode.Longitude = 0;
                geocode.index = index;
            //    geocode.FullAddress = "API Key Full";
            }

            return geocode;
        }

        public void CreateMap(List<GeoCode> pointList)
        {
            StringBuilder sb = new StringBuilder();           
            sb.Append("[");
            foreach (var item in pointList)
            {
                sb.Append("[\'");
                sb.Append(item.Name);
                sb.Append("\', ");
                sb.Append(item.Latitude);
                sb.Append(", ");
                sb.Append(item.Longitude);
                sb.Append(", ");
                sb.Append(item.index);
                sb.Append("],\r\n ");
            }
            sb.Append("];");
            //temp location
            string temppath = Path.GetTempPath() + "map.html";

            using (StreamReader sr = new StreamReader("map.html"))
            {
                string fileData = sr.ReadToEnd();
                fileData = fileData.Replace("[locationstring]",sb.ToString());

                using (StreamWriter sw = new StreamWriter(temppath))
                {
                    sw.Write(fileData);
                    sw.Flush();
                    sw.Close();
                }
            }

            MessageBox.Show("Where would you like to save the file?");
            SaveFileDialog folderBrowserDialog1 = new SaveFileDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                //
                // The user selected a folder and pressed the OK button.
                // We print the number of files found.
                //
                if (folderBrowserDialog1.FileName.Contains(".html"))
                {
                    File.Copy(temppath, folderBrowserDialog1.FileName, true);
                    Process.Start(folderBrowserDialog1.FileName);
                }
                else
                {
                    File.Copy(temppath, folderBrowserDialog1.FileName + ".html", true);
                    Process.Start(folderBrowserDialog1.FileName + ".html");
                }
            }

            

        }

      
  
        public void ChangeKey()
        {
            IEnumerable<string> lines = File.ReadLines("apiKeys.txt");
            var lineToRead = keyIndex;
            var line = lines.Skip(lineToRead - 1).First();
            apiKey = line;
            keyIndex++;
        }

        private void EditKey_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Process.Start("apiKeys.txt");
        }

        private string CleanUpText(string text)
        {

            text = text.Replace("\'", "");


            return text;
        }

        private void GenerateMap_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapList.Clear();

            progForm = new ProgressForm();

            progForm.progressBar1.Maximum = this.spreadsheetControl1.ActiveWorksheet.Rows.LastUsedIndex;
            progForm.Show();
            progForm.TopMost = true;

            currentCount = 0;
            location = bTxtLocation.EditValue.ToString();
            totalCount = this.spreadsheetControl1.ActiveWorksheet.Rows.LastUsedIndex;

            for (int i = 0; i <= this.spreadsheetControl1.ActiveWorksheet.Rows.LastUsedIndex; i++)
            {

                Address add = new Address();

                add.AddressName = CleanUpText(this.spreadsheetControl1.ActiveWorksheet.Cells[i, 0].Value.ToString());
                add.Location = location;
                add.index = i;



                BackgroundWorker bgworker = new BackgroundWorker();
                bgworker.DoWork += bgworker_DoWork;
                bgworker.RunWorkerCompleted += bgworker_RunWorkerCompleted;
                bgworker.RunWorkerAsync(add);
                runningCount++;
            }
        }

     
      

    }

    public class Address
    {
        public string AddressName { get; set; }
        public string Location { get; set; }
        public int index { get; set; }
        public string FullAddress { get; set; }
    }

    public class GeoCode
    {
        public double Longitude { get; set; }
        public double Latitude { get; set; }
        public int index { get; set; }
        public string Name { get; set; }
        public string FullAddress { get; set; }
    }
}

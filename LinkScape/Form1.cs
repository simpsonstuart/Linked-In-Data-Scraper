using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Diagnostics;


namespace LinkScape
{
    public partial class Form1 : Form
    {
        public CsvFileWriter CurrentFile { get; set; }
        public string URL { get; set; }

        public Form1()
        {
            InitializeComponent();
            //add to other part of installer sets IE emulation mode
            string installkey = @"SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION";
            string entryLabel = "LinkScape.exe";
            System.OperatingSystem osInfo = System.Environment.OSVersion;

            string version = osInfo.Version.Major.ToString() + '.' + osInfo.Version.Minor.ToString();
            uint editFlag = (uint)((version == "6.2") ? 0x2710 : 0x2328); // 6.2 = Windows 8 and therefore IE10

            RegistryKey existingSubKey = Registry.LocalMachine.OpenSubKey(installkey, false); // readonly key

            if (existingSubKey.GetValue(entryLabel) == null)
            {
                existingSubKey = Registry.LocalMachine.OpenSubKey(installkey, true); // writable key
                existingSubKey.SetValue(entryLabel, unchecked((int)editFlag), RegistryValueKind.DWord);
            }
            //hide add to list until we click get data first
            button1.Visible = false;
            //navigate to linked in and spoof user agent
            webBrowser1.Navigate("http://www.linkedin.com", "_self", null, "User-Agent: Mozilla/5.0 (Windows NT 6.3; WOW64; rv:29.0) Gecko/20100101 Firefox/29.0");

        }
        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            // Set text while the page has not yet loaded.
            this.Text = "Loading";
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            //hides real user agent to trick linked in
            webBrowser1.Navigate(txtURL.Text, "_self", null, "User-Agent: Mozilla/5.0 (Windows NT 6.3; WOW64; rv:29.0) Gecko/20100101 Firefox/29.0");
        }

        private void webBrowser1_Navigated(object sender,
            WebBrowserNavigatedEventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void button_GetData_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.HtmlDocument mydoc = webBrowser1.Document;
            //gets the name of the linked in contact
            HtmlElement ename = mydoc.GetElementById("name");
            HtmlElement eCompany = mydoc.GetElementById("title");
            HtmlElement elocation = mydoc.GetElementById("location");
            HtmlElement eposition = mydoc.GetElementById("background-experience-container");

            //set strings from html document
            string name = ename.InnerText;
            string company = eCompany.InnerText;
            string location = elocation.OuterText;
            string URL = webBrowser1.Url.ToString();

            //make the add to list button visible and hide get data
            button_GetData.Visible = false;
            button1.Visible = true;

            string sJobCompany = "";
            if (eposition != null)
            {
                // Get the first h5 element
                HtmlElementCollection h5Collection = eposition.GetElementsByTagName("h5");

                if (h5Collection != null)
                {
                    switch (h5Collection.Count)
                    {
                        case 0:
                            //show not found text and set textbox color to red
                            sJobCompany = "No Company Found!";
                            textBox_CompanyName.BackColor = Color.IndianRed;

                            break;

                        case 1:
                            sJobCompany = h5Collection[0].InnerText;
                            break;

                        case 2:
                            sJobCompany = h5Collection[1].InnerText;
                            break;

                        default:

                            if (h5Collection.Count > 2)
                            {
                                // Greater than two
                                sJobCompany = h5Collection[1].InnerText;
                                if (sJobCompany == null)
                                {
                                    sJobCompany = "No Company Found!";
                                }
                            }
                            else
                            {
                                // negative 
                                //show not found text and set textbox color to red
                                sJobCompany = "No Company Found!";
                                textBox_CompanyName.BackColor = Color.IndianRed;
                            }
                            break;

                    }
                }
                else
                {
                    //show not found text and set textbox color to red
                    sJobCompany = "No Company Found!";
                    textBox_CompanyName.BackColor = Color.IndianRed;
                }
            }
            string slocation = "";
            if (elocation != null)
            {
                if (elocation.GetElementsByTagName("dd") != null)
                {
                    //get the first dd in location
                    foreach (HtmlElement ddElement in elocation.GetElementsByTagName("dd"))
                    {
                        // Get the first dd element
                        slocation = ddElement.InnerText;
                        break;
                    }
                }
                else
                {
                    slocation = "BAD-DATA";
                }
            }
            //do validation on the position
            string clean_current_company = "";
            if (company != null)
            {
                //remove the words experience and current from text
                clean_current_company = company.Replace("Edit experienceCurrent", "").Trim();
            }
            else
            {
                clean_current_company = "No Position Found";
                textBox_Position.BackColor = Color.IndianRed;
            }

            //remove the words experience and current from text
            string clean_current_position = sJobCompany.Replace("Edit experienceCurrent", "").Trim();


            // Newline separator
            string[] saSeparatorNewline = new string[1];
            saSeparatorNewline[0] = "\r\n";

            // Comma separator
            string[] saSeparatorComma = new string[1];
            saSeparatorComma[0] = ",";

            // Space separator
            string[] saSeparatorSpace = new string[1];
            saSeparatorSpace[0] = " ";

            // Truncate anything after first comma for postition
            string sCompanyCleanse = clean_current_company;
            textBox_Position.Text = sCompanyCleanse;

            // Truncate anything after first comma for postition
            string sPositionCleanse = clean_current_position.Split(saSeparatorComma, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
            textBox_CompanyName.Text = sPositionCleanse;
            //set strings to null so we can get the outside of loop
            string clean_state = "";
            string sCity = "";

            //check to see if data is bad from source
            if (slocation != "BAD-DATA")
            {
                //split on return new line
                string[] saLocationStrings = slocation.Split(saSeparatorNewline, StringSplitOptions.RemoveEmptyEntries);

                //do filter logic to seperate city and state and reject "area"
                foreach (string sLocation in saLocationStrings)
                {
                    int iCommaLoc = sLocation.IndexOf(",");
                    if (iCommaLoc > -1)
                    {
                        string[] saParts = sLocation.Split(',');
                        if (saParts.Length >= 2)
                        {
                            // We know we have two part here that look like the city and state
                            sCity = saParts[0];
                            textBox_City.Text = sCity;
                            string sState = saParts[1];
                            clean_state = sState.Replace("Area", "").Trim();
                            textBox_State.Text = clean_state;
                        }
                    }
                }
            }
            else
            {
                clean_state = "No State Found!";
                textBox_State.BackColor = Color.IndianRed;
            }

            //split on space detect
            // Truncate anything after first comma

            string sNameCleanse = name.Split(saSeparatorComma, StringSplitOptions.RemoveEmptyEntries)[0].Trim();

            // Parse out name by spaces
            string[] saNameStrings = sNameCleanse.Split(saSeparatorSpace, StringSplitOptions.RemoveEmptyEntries);

            string sFirst_Name = "";
            string sLast_Name = "";
            if (saNameStrings != null)
            {
                // Declare Name 
                sFirst_Name = "";
                string sMiddle_Name = "";
                sLast_Name = "";

                switch (saNameStrings.Length)
                {
                    case 0:
                        sFirst_Name = "";
                        sMiddle_Name = "";
                        sLast_Name = "";
                        sFirst_Name = "Not Found";
                        textBox_FirstName.BackColor = Color.IndianRed;
                        sLast_Name = "Not Found";
                        textBox_LastName.BackColor = Color.IndianRed;
                        break;

                    case 1:
                        sFirst_Name = saNameStrings[0];
                        textBox_FirstName.Text = saNameStrings[0];
                        sMiddle_Name = "";
                        sLast_Name = "";
                        sLast_Name = "Not Found";
                        textBox_LastName.BackColor = Color.IndianRed;
                        break;

                    case 2:
                        sFirst_Name = saNameStrings[0];
                        textBox_FirstName.Text = saNameStrings[0];
                        sMiddle_Name = "";
                        sLast_Name = saNameStrings[1];
                        textBox_LastName.Text = saNameStrings[1];
                        break;

                    case 3:
                        sFirst_Name = saNameStrings[0];
                        textBox_FirstName.Text = saNameStrings[0];
                        sMiddle_Name = saNameStrings[1];
                        sLast_Name = saNameStrings[2];
                        textBox_LastName.Text = saNameStrings[2];
                        break;

                    default:
                        if (saNameStrings.Length > 3)
                        {
                            sFirst_Name = saNameStrings[0];
                            textBox_FirstName.Text = saNameStrings[0];
                            sMiddle_Name = saNameStrings[1];
                            sLast_Name = saNameStrings[2];
                            textBox_LastName.Text = saNameStrings[2];
                        }
                        else
                        {
                            sFirst_Name = "";
                            sMiddle_Name = "";
                            sLast_Name = "";
                        }

                        break;
                }
            }
            else
            {
                sFirst_Name = "Not Found";
                textBox_FirstName.BackColor = Color.IndianRed;
                sLast_Name = "Not Found";
                textBox_LastName.BackColor = Color.IndianRed;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //display confirmation message if checked
            if (checkBox_confirm.Checked)
            {
                MessageBox.Show(textBox_FirstName.Text + " " + textBox_LastName.Text + " - " + textBox_Position.Text + " " + "at " + textBox_CompanyName.Text + " " + " was added to the list");

            }
            else
            {
                //do nothing
            }
            
            
            //write actual data every time we click add to list
            string d = textBox_FirstName.Text + "," + textBox_LastName.Text + "," + textBox_CompanyName.Text + "," + textBox_Position.Text + "," + textBox_email1.Text + "," + textBox_email2.Text + "," + textBox_City.Text + "," + textBox_State.Text + "," + URL + "," + txtbx_source.Text;
            List<string> data = new List<string>();
            data.AddRange(d.Split(','));
            CurrentFile.WriteRow(data);


            //clear fields and set button back to default
            button1.Visible = false;
            button_GetData.Visible = true;
            textBox_FirstName.Text = "";
            textBox_LastName.Text = "";
            textBox_Position.Text = "";
            textBox_CompanyName.Text = "";
            textBox_email1.Text = "";
            textBox_email2.Text = "";
            textBox_City.Text = "";
            textBox_State.Text = "";
            URL = "";
        }
        private void btn_back_Click(object sender, EventArgs e)
        {
            if (webBrowser1.CanGoBack)
                webBrowser1.GoBack();
        }

        private void btn_forward_Click(object sender, EventArgs e)
        {
            if (webBrowser1.CanGoForward)
                webBrowser1.GoForward();
        }

        private void btn_refresh_Click(object sender, EventArgs e)
        {
            webBrowser1.Refresh();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }


        //begin csv class logic

        private void Form1_Load(object sender, EventArgs e)
        {
            CurrentFile = new CsvFileWriter(txtSaveLocations.Text);

            // write header data
            string s = "First Name,Last Name,Title,Company,Email 1,Email 2,City,State,Link,Source";
            List<string> columns = new List<string>();
            columns.AddRange(s.Split(','));
            CurrentFile.WriteRow(columns);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (CurrentFile != null)
            {
                CurrentFile.Dispose();
            }
        }

        private void txtSaveLocations_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void txtbx_source_TextChanged(object sender, EventArgs e)
        {

        }


    }
}

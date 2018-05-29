﻿using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Windows.Forms;

namespace ResumeEngineV2
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
            //Overlays progress bar ontop of rich text area where results are displayed
            progressBar1.BringToFront();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //User must enter something for search to work
            if (textBox1.Text == "")
            {
                MessageBox.Show("Please enter a valid keyword!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                progressBar1.Visible = true;

                //Fixes weird bug where label is cut off
                label2.Text = "";
                label2.Text = "Results:";

                richTextBox1.Text = "";
                button1.Enabled = false;
                textBox1.Enabled = false;

                string targetSiteURL = @"https://aecon1.sharepoint.com/sites/bd/resume/";

                var login = "JBraham@aecon.com";
                var password = "Winter@99";

                string term = textBox1.Text;

                var securePassword = new SecureString();

                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                }
                SharePointOnlineCredentials onlineCredentials = new SharePointOnlineCredentials(login, securePassword);

                ClientContext ctx = new ClientContext(targetSiteURL);
                ctx.Credentials = onlineCredentials;
                var web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();

                //Gets all folders under Documents
                var list = web.Lists.GetByTitle("Documents");
                ctx.Load(list);
                ctx.Load(list.RootFolder);
                ctx.Load(list.RootFolder.Folders);
                ctx.ExecuteQuery();
                FolderCollection fcol = list.RootFolder.Folders;
                List<string> lstFile = new List<string>();

                List<string> names = new List<string>();

                //Loops through each folder
                foreach (Folder f in fcol)
                {
                    //If folder is text
                    if (f.Name == "Text")
                    {
                        //Get all files under text folder
                        ctx.Load(f.Files);
                        ctx.ExecuteQuery();
                        var listItems = f.Files;

                        List<object> arguments = new List<object>();
                        arguments.Add(listItems);
                        arguments.Add(web);
                        arguments.Add(ctx);
                        arguments.Add(term);
                        arguments.Add(names);
                        arguments.Add(listItems.Count());
                        backgroundWorker1.WorkerReportsProgress = true;
                        backgroundWorker1.RunWorkerAsync(arguments);
                        break;
                    }
                }
            }
        }

        private void backgroundWoker1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> arguments = e.Argument as List<object>;
            int totalCount = (int)arguments[5];
            int count = 0;
            List<string> names = (System.Collections.Generic.List<string>)arguments[4];
            Web web = (Web)arguments[1];
            string postData = "[";
            //Loops through each file
            foreach (var item in (FileCollection)arguments[0])
            {
                count++;
                string fileName = item.Name;
                names.Add(fileName.Replace(".txt",""));
                var filePath = web.ServerRelativeUrl + "/Shared%20Documents/Text/" + fileName;
                //Gets file from SharePoint
                FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect((ClientContext)arguments[2], filePath);

                string ext = System.IO.Path.GetExtension(fileName);
                string convText = "";

                //Reads file into string
                using (StreamReader reader = new StreamReader(fileInformation.Stream))
                {
                    convText = reader.ReadToEnd();
                }

                List<Char> builder = new List<char>();
                //Used to fix if there are multiple newlines in a row
                bool isNewLine = true;

                //Remove special characters which would need to be escaped for JSON and creates new string using builder var
                for (int i = 0; i < convText.Length; i++)
                {
                    if (convText[i] == '\t')
                    {
                        builder.Add(' ');
                    }
                    else if ((convText[i] == '\n' || convText[i] == '\r') && isNewLine == false)
                    {
                        if (convText[i - 1] == '.' || convText[i - 1] == ':' || convText[i - 1] == ',')
                        {
                            builder.Add(' ');
                        }
                        else if (convText[i - 1] != ' ')
                        {
                            builder.Add('.');
                            builder.Add(' ');
                        }
                        isNewLine = true;
                    }
                    else if (convText[i] != '\n' && convText[i] != '\r')
                    {
                        isNewLine = false;
                        //If '"' is already escaped ignore
                        if (convText[i] == '"' && convText[i - 1] != '\\')
                        {
                            //Adds a single '\' before the '"'
                            builder.Add('\\');
                            builder.Add('"');
                        }
                        else
                        {
                            builder.Add(convText[i]);
                        }
                    }
                }
                string newConvText = new string(builder.ToArray());

                //System.IO.File.WriteAllText(@"C:\\Users\\brahamj\\Downloads\\newFormatTempText" + count + ".txt", newConvText);

                //Build JSON request string each loop
                postData += "[{\"term\": \"" + (string)arguments[3] + "\"},{\"text\": \"" + newConvText + "\"}],";

                //Send new progress bar value to backgroundWorker1_ProgressChanged as fields cannot be updated in backgroundWorker thread
                double progressPercent = ((double)count / totalCount) * 100;
                progressPercent = Math.Round(progressPercent, 0);
                backgroundWorker1.ReportProgress((int)progressPercent);
            }

            //Removes trailing ',' and replaces with ']' to close JSON object
            postData = postData.Remove(postData.Length - 1, 1) + "]";

            //System.IO.File.WriteAllText(@"C:\\Users\\brahamj\\Downloads\\jsonPost.txt", postData);

            //API Request to cortical.io to compare text taken from SharePoint with a keyword the user provided
            WebRequest webRequest = WebRequest.Create("http://api.cortical.io:80/rest/compare/bulk?retina_name=en_associative");
            webRequest.Method = "POST";
            webRequest.Headers["api-key"] = "bb355cc0-5873-11e8-9172-3ff24e827f76";
            webRequest.ContentType = "application/json";
            //Send request with postData string as the body
            using (var streamWriter = new StreamWriter(webRequest.GetRequestStream()))
            {
                streamWriter.Write(postData);
                streamWriter.Flush();
                streamWriter.Close();
            }
            string result = "";
            //Recieve response from cortical.io API
            try
            {
                WebResponse webResp = webRequest.GetResponse();
                using (var streamReader = new StreamReader(webResp.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("\nCannot connect to cortical.io API. Aborting!\nError: " + ex);
                Application.Exit();
            }

            //Formats return string as JSON
            dynamic jsonObj = JsonConvert.DeserializeObject<dynamic>(result);

            //Calculates match percent for each return object which correlates to each resume
            List<KeyValuePair<double, int>> percentName = new List<KeyValuePair<double, int>>();
            for (int i = 0; i < jsonObj.Count; i++)
            {
                double matchPercent = Math.Round((double)jsonObj[i].cosineSimilarity, 3);
                if (matchPercent <= 0.1)
                {
                    matchPercent = 0;
                }
                else
                {
                    matchPercent = Math.Round(((Math.Pow(Math.Log10(1 / matchPercent), 3.55) * -1) + 1) * 100, 2);
                }

                percentName.Add(new KeyValuePair<double, int>(matchPercent, i));

                percentName = percentName.OrderByDescending(x => x.Key).ToList();
            }

            string responseFull = "";
            //Generates response to populate rich text area
            for (int i = 0; i < percentName.Count(); i++)
            {
                responseFull += (i + 1 + ". \"" + names[percentName[i].Value] + "\" with " + percentName[i].Key + "%\n");
            }
            
            //Sends finished data to e.Result so when backgroundWorker1 is completed it can access the data and correctly update the fields
            //This has to be done as you cannot update the fields inside backgroundWorker thread
            List<object> returnArgs = new List<object>();
            returnArgs.Add("Results for \"" + (string)arguments[3] + "\":");
            returnArgs.Add(responseFull);
            e.Result = returnArgs;
        }

        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Update progress bar
            progressBar1.Value = e.ProgressPercentage;
        }

        void backgroundWorker1_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            //Update fields
            List<object> arguments = e.Result as List<object>;
            label2.Text = (string)arguments[0];
            richTextBox1.Text = (string)arguments[1];
            button1.Enabled = true;
            textBox1.Enabled = true;
            progressBar1.Value = 0;
            progressBar1.Visible = false;
        }
    }
}

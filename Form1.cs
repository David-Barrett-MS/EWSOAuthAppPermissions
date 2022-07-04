using System;
using System.Xml;
using System.Net.Http;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Identity.Client;

namespace EWSOAuthAppPermissions
{
    public partial class Form1 : Form
    {
        private HttpClient _httpClient = new HttpClient();
        private string _oAuthHeader = "";
        private string _pfXAnchorMailbox = "";
        private string _pfXAnchorMailboxContent = "";
        private string _pfXPublicFolderInformation = "";
        private ExtendedPropertyDefinition PR_REPLICA_LIST = new ExtendedPropertyDefinition(0x6698, MapiPropertyType.Binary);


        public Form1()
        {
            InitializeComponent();
        }

        private void WriteToResults(string data)
        {
            // Add the given data to the results textbox

            Action action = new Action(() => {
                textBoxResults.AppendText($"{data}{Environment.NewLine}");
            });
            if (textBoxResults.InvokeRequired)
                textBoxResults.Invoke(action);
            else
                action();
        }

        private void buttonFindFolders_Click(object sender, EventArgs e)
        {
            Action action = new Action(async () =>
            {
                // Configure the MSAL client to get tokens
                var ewsScopes = new string[] { "https://outlook.office.com/.default" };

                var app = ConfidentialClientApplicationBuilder.Create(textBoxAppId.Text)
                    .WithAuthority(AzureCloudInstance.AzurePublic, textBoxTenantId.Text)
                    .WithClientSecret(textBoxClientSecret.Text)
                    .Build();

                AuthenticationResult result = null;

                try
                {
                    // Acquire the token
                    result = await app.AcquireTokenForClient(ewsScopes)
                        .ExecuteAsync();

                    // Configure the ExchangeService with the access token
                    var ewsClient = new ExchangeService(ExchangeVersion.Exchange2016);
                    ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                    ewsClient.Credentials = new OAuthCredentials(result.AccessToken);
                    ewsClient.TraceListener = new TraceListener("trace.log");
                    ewsClient.TraceFlags = TraceFlags.All;
                    ewsClient.TraceEnabled = true;

                    // Add authorization header to HttpClient
                    _oAuthHeader = $"Bearer {result.AccessToken}";
                    _httpClient.DefaultRequestHeaders.Add("Authorization", _oAuthHeader);

                    //Impersonate the mailbox you'd like to access.
                    ewsClient.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, textBoxMailboxSMTPAddress.Text);

                    // Make an EWS call
                    if (checkBoxGetPublicFolders.Checked)
                    {
                        SetPublicFolderHeirarchyHeaders(ewsClient, textBoxMailboxSMTPAddress.Text);
                        FolderView folderView = new FolderView(500, 0, OffsetBasePoint.Beginning);
                        folderView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, PR_REPLICA_LIST);
                        var folders = ewsClient.FindFolders(WellKnownFolderName.PublicFoldersRoot, folderView);
                        foreach (var folder in folders)
                        {
                            WriteToResults($"Folder: {folder.DisplayName}");
                            ReadItemsFromFolder(folder);
                        }
                    }
                    else
                    {
                        ewsClient.HttpHeaders.Add("X-AnchorMailbox", textBoxMailboxSMTPAddress.Text);
                        var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
                        foreach (var folder in folders)
                        {
                            WriteToResults($"Folder: {folder.DisplayName}");
                        }
                    }
                }
                catch (MsalException ex)
                {
                    WriteToResults($"Error acquiring access token: {ex}");
                }
                catch (Exception ex)
                {
                    WriteToResults($"Error: {ex}");
                }


            });
            System.Threading.Tasks.Task.Run(action);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private AutodiscoverService GetAutodiscoverService(ExchangeService exchangeService)
        {
            AutodiscoverService autodiscover = new AutodiscoverService(new Uri("https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"), ExchangeVersion.Exchange2016);
            autodiscover.Credentials = exchangeService.Credentials;
            autodiscover.TraceListener = exchangeService.TraceListener;
            autodiscover.TraceFlags = TraceFlags.All;
            autodiscover.TraceEnabled = true;
            return autodiscover;
        }

        private void SetPublicFolderHeirarchyHeaders(ExchangeService exchangeService, string mailbox)
        {
            // We need to set specific headers when accessing public folders using EWS
            // https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-route-public-folder-hierarchy-requests
            // For public folders, we first need to get X-AnchorMailbox

            // Create the autodiscover service object
            AutodiscoverService autodiscover = GetAutodiscoverService(exchangeService);

            // Retrieve the autodiscover information for X-AnchorMailbox
            GetUserSettingsResponse userSettings = null;
            try
            {
                userSettings = autodiscover.GetUserSettings(mailbox, UserSettingName.PublicFolderInformation);
            }
            catch (Exception ex)
            {
                WriteToResults($"Error: {ex}");
                return;
            }

            // We have X-AnchorMailbox
            _pfXAnchorMailbox = userSettings.Settings[UserSettingName.PublicFolderInformation].ToString();
            WriteToResults($"Set X-AnchorMailbox to {_pfXAnchorMailbox}");
            exchangeService.HttpHeaders.Add("X-AnchorMailbox", _pfXAnchorMailbox);

            // Retrieve the autodiscover information for X-PublicFolderInformation

            if (checkBoxPOXBasicAuth.Checked)
            {
                // Use basic auth for the POX request
                // We work out the auth header and replace it
                String basicAuth = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{textBoxMailboxSMTPAddress.Text}:{textBoxAutoDiscoverPW.Text}"));
                _httpClient.DefaultRequestHeaders.Remove("Authorization");
                _httpClient.DefaultRequestHeaders.Add("Authorization", "Basic " + basicAuth);
            }

            // As we are using the POX endpoint, we send the Autodiscover request using HttpClient
            string autodiscoverXml = System.IO.File.ReadAllText("Autodiscover.xml").Replace("<!--EMAILADDRESS-->", _pfXAnchorMailbox);
            exchangeService.TraceListener?.Trace("AutodiscoverRequest", autodiscoverXml);
            HttpResponseMessage response = _httpClient.PostAsync("https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml",
                new StringContent(autodiscoverXml, System.Text.Encoding.UTF8, "text/xml")).Result;
            if (!response.IsSuccessStatusCode)
            {
                WriteToResults("Failed to get X-PublicFolderInformation");
                return;
            }

            // We've got an autodiscover response.  Log it, then we need to parse it.
            string autodiscoverResponse = response.Content.ReadAsStringAsync().Result;
            exchangeService.TraceListener?.Trace("AutodiscoverResponse", autodiscoverResponse);

            // The X-PublicFolderInformation value is the Server value from the EXCH Protocol Type of the response
            // <Account><Protocol><Type>EXCH</Type></Protocol><Protocol/></Account>
            XmlDocument autodiscoverXmlResponse = new XmlDocument();
            autodiscoverXmlResponse.LoadXml(autodiscoverResponse);

            XmlNodeList responseNodes = autodiscoverXmlResponse.GetElementsByTagName("Response");
            if (responseNodes.Count == 1)
            {
                foreach (XmlNode xmlAccountNode in responseNodes[0].ChildNodes)
                {
                    if (xmlAccountNode.Name == "Account")
                    {
                        foreach (XmlNode xmlProtocolNode in xmlAccountNode.ChildNodes)
                        {
                            bool isEXCH = false;
                            if (xmlProtocolNode.Name == "Protocol")
                            {
                                // Check if this is EXCH
                                foreach (XmlNode node in xmlProtocolNode.ChildNodes)
                                {
                                    if (node.Name == "Type" && node.InnerText == "EXCH")
                                    {
                                        //  We want the Server entry from this node
                                        isEXCH = true;
                                    }
                                    if (isEXCH && node.Name == "Server")
                                    {
                                        //  We want the Server entry from this node
                                        _pfXPublicFolderInformation = node.InnerText;
                                        exchangeService.HttpHeaders.Add("X-PublicFolderInformation", _pfXPublicFolderInformation);
                                        WriteToResults($"Set X-PublicFolderInformation to {_pfXPublicFolderInformation}");
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //  If we get here, we failed to find X-PublicFolderInformation
            WriteToResults("Failed to get X-PublicFolderInformation");
        }

        public void SetPublicFolderContentHeaders(Folder folder)
        {
            // X- headers for content requests must be set per:
            // https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-route-public-folder-content-requests

            if (String.IsNullOrEmpty(_pfXAnchorMailbox) || String.IsNullOrEmpty(_pfXPublicFolderInformation))
                return;

            // First we need the replica Guid (in string format).  We requested this in the original FindFolder call.
            string replicaGuid = "";
            foreach (ExtendedProperty prop in folder.ExtendedProperties)
                if (prop.PropertyDefinition==PR_REPLICA_LIST)
                {
                    byte[] ByteArr = (byte[])prop.Value;
                    replicaGuid = System.Text.Encoding.ASCII.GetString(ByteArr, 0, 36);
                }

            if (String.IsNullOrEmpty(replicaGuid))
                return;

            // We use the Guid to determine the Autodiscover address
            int domainStart = textBoxMailboxSMTPAddress.Text.IndexOf("@");
            if (domainStart < 0)
                return;
            string domain = textBoxMailboxSMTPAddress.Text.Substring(domainStart);
            string autodiscoverAddress = $"{replicaGuid}{domain}";

            // We send a POX autodiscover request to get X-AnchorMailbox (using whatever auth method has
            // already been configured on the HttpClient)
            string autodiscoverXml = System.IO.File.ReadAllText("Autodiscover.xml").Replace("<!--EMAILADDRESS-->", autodiscoverAddress);
            folder.Service.TraceListener?.Trace("AutodiscoverRequest", autodiscoverXml);
            HttpResponseMessage response = _httpClient.PostAsync("https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml",
                new StringContent(autodiscoverXml, System.Text.Encoding.UTF8, "text/xml")).Result;
            if (!response.IsSuccessStatusCode)
            {
                WriteToResults("Failed to get X-AnchorMailbox");
                return;
            }
            string autodiscoverResponse = response.Content.ReadAsStringAsync().Result;
            folder.Service.TraceListener?.Trace("AutodiscoverResponse", autodiscoverResponse);

            // AutoDiscoverSMTPAddress is the value we need for both headers
            XmlDocument autodiscoverXmlResponse = new XmlDocument();
            autodiscoverXmlResponse.LoadXml(autodiscoverResponse);
            XmlNodeList addressNodes = autodiscoverXmlResponse.GetElementsByTagName("AutoDiscoverSMTPAddress");
            if (addressNodes.Count == 1)
                _pfXAnchorMailboxContent = addressNodes[0].InnerText;
            else
            {
                WriteToResults("Failed to get X-AnchorMailbox");
                return;
            }

            // Set the headers on the service request
            folder.Service.HttpHeaders.Remove("X-AnchorMailbox");
            folder.Service.HttpHeaders.Add("X-AnchorMailbox", _pfXAnchorMailboxContent);
            WriteToResults($"Set X-AnchorMailbox to {_pfXAnchorMailboxContent}");

            folder.Service.HttpHeaders.Remove("X-PublicFolderInformation");
            folder.Service.HttpHeaders.Add("X-PublicFolderInformation", _pfXAnchorMailboxContent);
            WriteToResults($"Set X-PublicFolderInformation to {_pfXAnchorMailboxContent}");
        }

        private void ReadItemsFromFolder(Folder folder)
        {
            ItemView itemView = new ItemView(10, 0, OffsetBasePoint.Beginning);

            if (checkBoxGetPublicFolders.Checked)
                SetPublicFolderContentHeaders(folder);
            FindItemsResults<Item> results = folder.FindItems(itemView);
            foreach (var item in results.Items)
                WriteToResults($"Item: {item.Subject}");
        }

        private void checkBoxGetPublicFolders_CheckedChanged(object sender, EventArgs e)
        {
            textBoxAutoDiscoverPW.Enabled = checkBoxGetPublicFolders.Checked;
        }
    }
}

/*
 * By David Barrett, Microsoft Ltd. 2022. Use at your own risk.  No warranties are given.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * */

using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Net.Http;
using System.Windows.Forms;
using System.Xml;

namespace EWSOAuthAppPermissions
{
    public partial class Form1 : Form
    {
        private HttpClient _httpClient = new HttpClient();
        private string _oAuthHeader = "";
        private string _pfXAnchorMailbox = "";
        private string _internalRpcClientServer = "";
        private string _pfXAnchorMailboxContent = "";
        private string _pfXPublicFolderMailbox = "";
        private ExtendedPropertyDefinition PR_REPLICA_LIST = new ExtendedPropertyDefinition(0x6698, MapiPropertyType.Binary);
        private DateTime _throttledBackoffEnd = DateTime.MinValue;

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
                    ewsClient.UserAgent = textBoxUserAgent.Text;

                    // Add OAuth authorization header to HttpClient
                    _oAuthHeader = $"Bearer {result.AccessToken}";
                    if (_httpClient.DefaultRequestHeaders.Contains("Authorization"))
                        _httpClient.DefaultRequestHeaders.Remove("Authorization");
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
                            if (checkBoxGetItems.Checked)
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
            autodiscover.UserAgent = textBoxUserAgent.Text;
            return autodiscover;
        }

        private void ApplyBasicAuthHeader()
        {
            if (checkBoxPOXBasicAuth.Checked)
            {
                // We work out the basic auth header and then set it (replacing the OAuth header)
                String basicAuth = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{textBoxMailboxSMTPAddress.Text}:{textBoxAutoDiscoverPW.Text}"));
                _httpClient.DefaultRequestHeaders.Remove("Authorization");
                _httpClient.DefaultRequestHeaders.Add("Authorization", "Basic " + basicAuth);
            }
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
                UserSettingName[] settings = new UserSettingName[] { UserSettingName.PublicFolderInformation, UserSettingName.InternalRpcClientServer};
                userSettings = autodiscover.GetUserSettings(mailbox, settings);
            }
            catch (Exception ex)
            {
                WriteToResults($"Error: {ex}");
                return;
            }

            // Set X-AnchorMailbox
            _pfXAnchorMailbox = userSettings.Settings[UserSettingName.PublicFolderInformation].ToString();
            if (checkBoxHeaders.Checked)
            {
                WriteToResults($"Set X-AnchorMailbox to {_pfXAnchorMailbox}");
                exchangeService.HttpHeaders.Add("X-AnchorMailbox", _pfXAnchorMailbox);
            }

            // Retrieve the autodiscover information for X-PublicFolderMailbox
            // The docs state that this should be done by sending a POX request, but we can obtain the same information
            // from a UserSettings request for InternalRpcClientServer against the anchor mailbox.
            try
            {
                userSettings = autodiscover.GetUserSettings(_pfXAnchorMailbox, UserSettingName.InternalRpcClientServer);
                _internalRpcClientServer = userSettings.Settings[UserSettingName.InternalRpcClientServer].ToString();
                WriteToResults($"InternalRpcClientServer: {_internalRpcClientServer}");

                // InternalRpcClientServer is the value we need for X-PublicFolderMailbox
                _pfXPublicFolderMailbox = _internalRpcClientServer;
                if (checkBoxHeaders.Checked)
                {
                    exchangeService.HttpHeaders.Add("X-PublicFolderMailbox", _pfXPublicFolderMailbox);
                    WriteToResults($"Set X-PublicFolderMailbox to {_pfXPublicFolderMailbox}");
                }
                return;
            }
            catch (Exception ex)
            {
                WriteToResults($"Error retrieving InternalRpcClientServer: {ex}");
            }

            // If we get here, we failed to get the user settings.  Likely this is a mailbox issue, but we'll try POX.

            // Retrieve the autodiscover information for X-PublicFolderMailbox using POX

            ApplyBasicAuthHeader();

            // As we are using the POX endpoint, we send the Autodiscover request using HttpClient
            string autodiscoverXml = System.IO.File.ReadAllText("Autodiscover.xml").Replace("<!--EMAILADDRESS-->", _pfXAnchorMailbox);
            exchangeService.TraceListener?.Trace("AutodiscoverRequest", autodiscoverXml);
            HttpResponseMessage response = _httpClient.PostAsync("https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml",
                new StringContent(autodiscoverXml, System.Text.Encoding.UTF8, "text/xml")).Result;
            if (!response.IsSuccessStatusCode)
            {
                WriteToResults("Failed to get X-PublicFolderMailbox");
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
                                        _pfXPublicFolderMailbox = node.InnerText;
                                        if (checkBoxHeaders.Checked)
                                        {
                                            exchangeService.HttpHeaders.Add("X-PublicFolderMailbox", _pfXPublicFolderMailbox);
                                            WriteToResults($"Set X-PublicFolderMailbox to {_pfXPublicFolderMailbox}");
                                        }
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

            if (String.IsNullOrEmpty(_pfXAnchorMailbox) )
                return;

            _pfXAnchorMailboxContent = "";

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
            WriteToResults($"Autodiscover for {autodiscoverAddress} to access folder {folder.DisplayName}");

            // Retrieve the autodiscover information for X-AnchorMailbox
            AutodiscoverService autodiscover = GetAutodiscoverService(folder.Service);
            GetUserSettingsResponse userSettings = null;
            try
            {
                UserSettingName[] settings = new UserSettingName[] { UserSettingName.PublicFolderInformation, UserSettingName.AutoDiscoverSMTPAddress };
                userSettings = autodiscover.GetUserSettings(autodiscoverAddress, settings);
            }
            catch (Exception ex)
            {
                WriteToResults($"Error: {ex}");
                return;
            }
            _pfXAnchorMailboxContent = userSettings.Settings[UserSettingName.AutoDiscoverSMTPAddress].ToString();

            if (String.IsNullOrEmpty(_pfXAnchorMailboxContent))
            {
                // GetUserSettings failed for some reason, we fall back to POX (which will only work with basic auth)

                ApplyBasicAuthHeader();

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
            }

            // Set the headers on the service request
            if (checkBoxHeaders.Checked)
            {
                folder.Service.HttpHeaders.Remove("X-AnchorMailbox");
                folder.Service.HttpHeaders.Add("X-AnchorMailbox", _pfXAnchorMailboxContent);
                WriteToResults($"Set X-AnchorMailbox to {_pfXAnchorMailboxContent}");

                folder.Service.HttpHeaders.Remove("X-PublicFolderMailbox");
                folder.Service.HttpHeaders.Add("X-PublicFolderMailbox", _pfXAnchorMailboxContent);
                WriteToResults($"Set X-PublicFolderMailbox to {_pfXAnchorMailboxContent}");
            }
        }

        private bool IsCatastrophicError(Exception ex)
        {
            //  Check for catastrophic errors
            WriteToResults(ex.Message);
            return true;
        }

        private void ReadItemsFromFolder(Folder folder)
        {
            WriteToResults($"Starting folder read: {folder.DisplayName}");
            ItemView itemView = new ItemView(10, 0, OffsetBasePoint.Beginning);
            itemView.PropertySet = BasePropertySet.IdOnly;

            if (checkBoxGetPublicFolders.Checked)
                SetPublicFolderContentHeaders(folder);
            bool moreItems = true;

            while (moreItems)
            {
                try
                {
                    FindItemsResults<Item> results = folder.FindItems(itemView);
                    moreItems = results.MoreAvailable;
                    itemView.Offset += results.Items.Count;

                    foreach (var item in results.Items)
                    {
                        try
                        {
                            Item item1 = Item.Bind(folder.Service, item.Id, BasePropertySet.FirstClassProperties);
                            WriteToResults($"Item: {item1.Subject}");
                        }
                        catch (Exception ex)
                        {
                            if (IsCatastrophicError(ex))
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (IsCatastrophicError(ex))
                        break;
                }
            }
            WriteToResults($"Completed folder read: {folder.DisplayName}");
        }

        private void checkBoxGetPublicFolders_CheckedChanged(object sender, EventArgs e)
        {
            textBoxAutoDiscoverPW.Enabled = checkBoxGetPublicFolders.Checked;
        }

        private void buttonClearLog_Click(object sender, EventArgs e)
        {
            textBoxResults.Clear();
        }
    }
}

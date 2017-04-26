using CommandLine;
using SpiraAttachmentImport.Classes;
using SpiraAttachmentImport.SpiraService;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SpiraAttachmentImport
{
    class Program
    {
        static Settings _options;

        static void Main(string[] args)
        {
            _options = new Settings();
            if (Parser.Default.ParseArguments(args, _options))
            {

                //Make sure that the Spira string is a URL.
                Uri serverUrl = null;
                string servicePath = _options.SpiraServer +
                    ((_options.SpiraServer.EndsWith("/") ? "" : "/")) +
                    "Services/v5_0/SoapService.svc";
                if (!Uri.TryCreate(servicePath, UriKind.Absolute, out serverUrl))
                {
                    //Throw error: URL given not a URL.
                    return;
                }

                //See if we have a mappings file specified, if so, make sure it exists
                List<ArtifactMapping> artifactMappings = null;
                string customPropertyField = null;
                if (!String.IsNullOrWhiteSpace(_options.SpiraMappingsFile))
                {
                    if (!File.Exists(_options.SpiraMappingsFile))
                    {
                        //Throw error: Bad path.
                        ConsoleLog(LogLevelEnum.Normal, "Cannot access the mapping file, please check the location and try again!");
                        Environment.Exit(-1);
                    }
                    artifactMappings = new List<ArtifactMapping>();

                    //Read in the lines, the first column should contain:
                    //Filename,ArtifactTypeId,Custom_03
                    //where the number in the third column is the name of the custom property that the IDs will be using
                    using (StreamReader streamReader = File.OpenText(_options.SpiraMappingsFile))
                    {
                        string firstLine = streamReader.ReadLine();
                        string[] headings = firstLine.Split(',');

                        //See if we have a match on the custom property number
                        if (headings.Length > 2 && !String.IsNullOrWhiteSpace(headings[2]))
                        {
                            customPropertyField = headings[2].Trim();

                            //Now read in the rows of mappings
                            while (!streamReader.EndOfStream)
                            {
                                string mappingLine = streamReader.ReadLine();
                                ArtifactMapping artifactMapping = new ArtifactMapping();
                                string[] components = mappingLine.Split(',');
                                artifactMapping.Filename = components[0];
                                artifactMapping.ArtifactTypeId = Int32.Parse(components[1]);
                                artifactMapping.ExternalKey = components[2];
                                artifactMappings.Add(artifactMapping);
                            }
                            streamReader.Close();
                        }
                    }
                }

                //Make sure the path given is a real path..
                try
                {
                    Directory.GetCreationTime(_options.ImportPath);
                }
                catch
                {
                    //Throw error: Bad path.
                    ConsoleLog(LogLevelEnum.Normal, "Cannot access the import path, please check the location and try again!");
                    Environment.Exit(-1);
                }

                //Tell user we're operating.
                ConsoleLog(LogLevelEnum.Normal, "Importing files in " + _options.ImportPath);

                //Now run through connecting procedures.
                ConsoleLog(LogLevelEnum.Verbose, "Connecting to Spira server...");
                SoapServiceClient client = CreateClient_Spira5(serverUrl);
                client.Open();

                if (client.Connection_Authenticate2(_options.SpiraLogin, _options.SpiraPass, "DocumentImporter"))
                {
                    ConsoleLog(LogLevelEnum.Verbose, "Selecting Spira project...");
                    var Projects = client.Project_Retrieve();
                    RemoteProject proj = Projects.FirstOrDefault(p => p.ProjectId == _options.SpiraProject);

                    if (proj != null)
                    {
                        //Connect to the project.
                        if (client.Connection_ConnectToProject((int)_options.SpiraProject))
                        {
                            ConsoleLog(LogLevelEnum.Normal, "Uploading files...");

                            //Now let's get a list of all the files..
                            SearchOption opt = ((_options.PathRecursive) ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
                            List<string> files = Directory.EnumerateFiles(_options.ImportPath, _options.ImportFilter, opt).ToList();

                            //Loop through each file and upload it.
                            ConsoleLog(LogLevelEnum.Verbose, "Files:");
                            foreach (string file in files)
                            {
                                string conLog = Path.GetFileName(file);
                                try
                                {
                                    //Get the file details..
                                    FileInfo fileInf = new FileInfo(file);
                                    string size = "";
                                    if (fileInf.Length >= 1000000000)
                                        size = (fileInf.Length / 1000000000).ToString() + "Gb";
                                    else if (fileInf.Length >= 1000000)
                                        size = (fileInf.Length / 1000000).ToString() + "Mb";
                                    else if (fileInf.Length >= 1000)
                                        size = (fileInf.Length / 1000).ToString() + "Kb";
                                    else
                                        size = fileInf.Length.ToString() + "b";
                                    conLog += " (" + size + ")";

                                    //Generate the RemoteDocument object.
                                    RemoteDocument newFile = new RemoteDocument();
                                    newFile.AttachmentTypeId = 1;
                                    newFile.FilenameOrUrl = Path.GetFileName(file);
                                    newFile.Size = (int)fileInf.Length;
                                    newFile.UploadDate = DateTime.UtcNow;

                                    //Now we see if have mapped artifact
                                    ArtifactMapping mappedArtifact = artifactMappings.FirstOrDefault(m => m.Filename.ToLowerInvariant() == newFile.FilenameOrUrl.ToLowerInvariant());
                                    if (mappedArtifact != null && !String.IsNullOrEmpty(customPropertyField))
                                    {
                                        //We have to lookup the artifact, currently only incidents are supported
                                        if (mappedArtifact.ArtifactTypeId == 3)
                                        {
                                            //Retrieve the incident
                                            RemoteSort sort = new RemoteSort();
                                            sort.PropertyName = "IncidentId";
                                            sort.SortAscending = true;
                                            List<RemoteFilter> filters = new List<RemoteFilter>();
                                            RemoteFilter filter = new RemoteFilter();
                                            filter.PropertyName = customPropertyField;
                                            filter.StringValue = mappedArtifact.ExternalKey;
                                            filters.Add(filter);
                                            RemoteIncident remoteIncident = client.Incident_Retrieve(filters, sort, 1, 1).FirstOrDefault();

                                            if (remoteIncident != null)
                                            {
                                                RemoteLinkedArtifact remoteLinkedArtifact = new SpiraService.RemoteLinkedArtifact();
                                                remoteLinkedArtifact.ArtifactTypeId = mappedArtifact.ArtifactTypeId;
                                                remoteLinkedArtifact.ArtifactId = remoteIncident.IncidentId.Value;
                                                newFile.AttachedArtifacts = new List<RemoteLinkedArtifact>();
                                                newFile.AttachedArtifacts.Add(remoteLinkedArtifact);
                                            }
                                        }
                                        else
                                        {
                                            ConsoleLog(LogLevelEnum.Normal, "Warning: Only incident mapped artifacts currently supported, so ignoring the mapped artifacts of type: " + mappedArtifact.ArtifactTypeId);
                                        }
                                    }

                                    //Read the file contents. (Into memory! Beware, large files!)
                                    byte[] fileContents = null;
                                    fileContents = File.ReadAllBytes(file);

                                    ConsoleLog(LogLevelEnum.Verbose, conLog);
                                    if (fileContents != null && fileContents.Length > 1)
                                        newFile.AttachmentId = client.Document_AddFile(newFile, fileContents).AttachmentId.Value;
                                    else
                                        throw new FileEmptyException();
                                }
                                catch (Exception ex)
                                {
                                    conLog += " - Error. (" + ex.GetType().ToString() + ")";
                                    ConsoleLog(LogLevelEnum.Normal, conLog);
                                }
                            }
                        }
                        else
                        {
                            ConsoleLog(LogLevelEnum.Normal, "Cannot connect to project. Verify your Project Role.");
                            Environment.Exit(-1);
                        }
                    }
                    else
                    {
                        ConsoleLog(LogLevelEnum.Normal, "Cannot connect to project. Project #" + _options.SpiraProject.ToString() + " does not exist.");
                        Environment.Exit(-1);
                    }
                }
                else
                {
                    ConsoleLog(LogLevelEnum.Normal, "Cannot log in. Check username and password.");
                    Environment.Exit(-1);
                }
            }
            else
                Environment.Exit(-1);
        }

        #region Private Methods
        /// <summary>Write an output message.</summary>
        /// <param name="type"></param>
        /// <param name="Message"></param>
        private static void ConsoleLog(LogLevelEnum type, string Message)
        {
            if (type < LogLevelEnum.Verbose || _options.Verbose)
            {
                Console.WriteLine(Message);
            }
        }

        /// <summary>Creates the WCF endpoints</summary>
        /// <param name="fullUri">The URI</param>
        /// <returns>The client class</returns>
        /// <remarks>We need to do this in code because the app.config file is not available in VSTO</remarks>
        public static SpiraService.SoapServiceClient CreateClient_Spira5(Uri baseUri)
        {
            //Create the binding, first. Allow cookies, and allow margest message sizes.
            BasicHttpBinding httpBinding = new BasicHttpBinding();
            httpBinding.AllowCookies = true;
            httpBinding.MaxBufferSize = int.MaxValue;
            httpBinding.MaxBufferPoolSize = int.MaxValue;
            httpBinding.MaxReceivedMessageSize = int.MaxValue;
            httpBinding.ReaderQuotas.MaxStringContentLength = int.MaxValue;
            httpBinding.ReaderQuotas.MaxDepth = int.MaxValue;
            httpBinding.ReaderQuotas.MaxBytesPerRead = int.MaxValue;
            httpBinding.ReaderQuotas.MaxNameTableCharCount = int.MaxValue;
            httpBinding.ReaderQuotas.MaxArrayLength = int.MaxValue;
            httpBinding.CloseTimeout = new TimeSpan(0, 2, 0);
            httpBinding.OpenTimeout = new TimeSpan(0, 2, 0);
            httpBinding.ReceiveTimeout = new TimeSpan(0, 5, 0);
            httpBinding.SendTimeout = new TimeSpan(0, 5, 0);
            httpBinding.BypassProxyOnLocal = false;
            httpBinding.HostNameComparisonMode = HostNameComparisonMode.StrongWildcard;
            httpBinding.MessageEncoding = WSMessageEncoding.Text;
            httpBinding.TextEncoding = Encoding.UTF8;
            httpBinding.TransferMode = TransferMode.Buffered;
            httpBinding.UseDefaultWebProxy = true;
            httpBinding.Security.Mode = BasicHttpSecurityMode.None;
            httpBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.None;
            httpBinding.Security.Transport.ProxyCredentialType = HttpProxyCredentialType.None;
            httpBinding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;
            httpBinding.Security.Message.AlgorithmSuite = System.ServiceModel.Security.SecurityAlgorithmSuite.Default;

            //Handle SSL if necessary
            if (baseUri.Scheme == "https")
            {
                httpBinding.Security.Mode = BasicHttpSecurityMode.Transport;
                httpBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.None;

                //Allow self-signed certificates
                PermissiveCertificatePolicy.Enact("");
            }
            else
                httpBinding.Security.Mode = BasicHttpSecurityMode.None;

            //Create the new client with endpoint and HTTP Binding
            EndpointAddress endpointAddress = new EndpointAddress(baseUri.AbsoluteUri);
            SpiraService.SoapServiceClient spiraSoapService = new SpiraService.SoapServiceClient(httpBinding, endpointAddress);

            //Modify the operation behaviors to allow unlimited objects in the graph
            foreach (var operation in spiraSoapService.Endpoint.Contract.Operations)
            {
                var behavior = operation.Behaviors.Find<DataContractSerializerOperationBehavior>() as DataContractSerializerOperationBehavior;
                if (behavior != null)
                {
                    behavior.MaxItemsInObjectGraph = 2147483647;
                }
            }
            return spiraSoapService;
        }
        #endregion

        #region Enumerations
        private enum LogLevelEnum : int
        {
            Normal = 1,
            Verbose = 2
        }
        #endregion
    }
}

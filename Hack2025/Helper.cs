using Azure.AI.OpenAI;
using Azure.Core;
using Azure.Identity;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using static Program;
using AutoGen.OpenAI;
using AutoGen.OpenAI.Extension;
using AutoGen.Core;
using OpenAI.Chat;
using System.Text.Json;
using Microsoft.Identity.Client;
using System;
//using SautinSoft.Document;
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using System;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Hack2025.Models;

namespace Hack2025
{
    public class Helper
    {
        // Recursive function to print the JSON structure
        public static void PrintJsonStructure(JObject jsonObject, int indentLevel)
        {
            foreach (var property in jsonObject.Properties())
            {
                // Indentation based on the level of recursion
                Console.WriteLine(new string(' ', indentLevel * 2) + property.Name);

                // Check if the value is a JObject or JArray, and print the structure recursively
                if (property.Value is JObject)
                {
                    PrintJsonStructure((JObject)property.Value, indentLevel + 1);
                }
                else if (property.Value is JArray)
                {
                    // If it's an array, we print the structure for each item
                    PrintJsonArrayStructure((JArray)property.Value, indentLevel + 1);
                }
                else
                {
                    // Print the property type if it's a simple value
                    Console.WriteLine(new string(' ', (indentLevel + 1) * 2) + property.Value.Type);
                }
            }
        }

        // Function to print the structure of JSON arrays
        public static void PrintJsonArrayStructure(JArray jsonArray, int indentLevel)
        {
            for (int i = 0; i < jsonArray.Count; i++)
            {
                Console.WriteLine(new string(' ', indentLevel * 2) + $"Item {i + 1}:");
                if (jsonArray[i] is JObject)
                {
                    PrintJsonStructure((JObject)jsonArray[i], indentLevel + 1);
                }
                else
                {
                    // Print the item type if it's a simple value
                    Console.WriteLine(new string(' ', (indentLevel + 1) * 2) + jsonArray[i].Type);
                }
            }
        }

        public static string ConvertToVisualStudioUrl(string originalUrl, string prid, string project, string reponame)
        {
            string ans = $"https://microsoft.visualstudio.com/{project}/_git/{reponame}/pullrequest/{prid}";

            return ans;
        }

        // Function to fetch work items associated with a pull request
        public static async Task<List<string>> GetWorkItemsForPullRequest(string url)
        {
            var credential = new DefaultAzureCredential();
            var tokenRequestContext = new TokenRequestContext(new[] { "499b84ac-1321-427f-aa17-267ca6975798/.default" }); // Azure DevOps resource
            var token = await credential.GetTokenAsync(tokenRequestContext);
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
                HttpResponseMessage response = await client.GetAsync(url);
                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"Error: {response.StatusCode}, {response.ReasonPhrase}");
                    return new List<string>();
                }
                string responseBody = await response.Content.ReadAsStringAsync();
                JObject jsonResponse = JObject.Parse(responseBody);
                if (jsonResponse["completionOptions"] != null && jsonResponse["completionOptions"]["mergeCommitMessage"] != null)
                {
                    string inputString = jsonResponse["completionOptions"]["mergeCommitMessage"].ToString();
                    List<string> workItems = ExtractWorkItems(inputString);
                    return workItems;
                }
                else
                {
                    return new List<string>();
                }
            }
        }

        // Method to extract work items from a string
        public static List<string> ExtractWorkItems(String input)
        {
            List<string> workItems = new List<string>();
            // Regex pattern to match work items like #12345678
            string pattern = @"#(\d+)";

            // Find matches using the regular expression
            MatchCollection matches = Regex.Matches(input, pattern);

            // Iterate through the matches and add to the list
            foreach (Match match in matches)
            {
                if (match.Success)
                {
                    //workItems.Add(match.Value);  // Add the work item to the list
                    workItems.Add($"https://microsoft.visualstudio.com/OS/_workitems/edit/{match.Value.Substring(1)}");  // Add the work item to the list
                }
            }

            return workItems;
        }

        //public static void PrintListOfPrs(List<PullRequest> activePrList, List<PullRequest> completedPrList)
        public static void PrintListOfPrs(List<PullRequest> completedPrList)
        {

            foreach (var pr in completedPrList)
            {
                Console.WriteLine($"Title: {pr.title}");
                Console.WriteLine($"ID: {pr.id}");
                Console.WriteLine($"Created by (display name): {pr.createdByName}");
                Console.WriteLine($"Status: {pr.status}");
                Console.WriteLine($"Creation Date: {pr.creationDate}");
                Console.WriteLine($"Closed Date: {pr.closedDate}");
                Console.WriteLine($"Url: {pr.url}");

                // Extract and print work items from description
                Console.WriteLine($"Work Items: {string.Join(", ", pr.listOfTaskUrls)}");

                // Print approvers
                Console.WriteLine($"Approved By: {string.Join(", ", pr.approvedBy)}");
                Console.WriteLine($"AI Summary: {pr.aiPrSummary}");
                Console.WriteLine($"AI Classification: {pr.aiPrClassification}");
                Console.WriteLine("--------------------------------------------");
            }
        }

        // Writes PR details to both console and a txt file
        public static void PrintAndWriteListOfPrs(List<PullRequest> completedPrList, string userAlias = null)
        {
            string aliasPart = "";
            if (!string.IsNullOrEmpty(userAlias))
            {
                // Remove domain if present
                var alias = userAlias.Split('@')[0];
                aliasPart = $"_{alias}";
            }
            // Get the directory where Helper.cs/Program.cs exist (project directory)
            // Option 2: Use a more robust method to find project root
            string projectRoot = AppDomain.CurrentDomain.BaseDirectory;
            while (!string.IsNullOrEmpty(projectRoot) && !File.Exists(Path.Combine(projectRoot, "Hack2025.csproj")))
            {
                projectRoot = Directory.GetParent(projectRoot)?.FullName;
            }
            if (string.IsNullOrEmpty(projectRoot))
            {
                throw new InvalidOperationException("Could not find project root directory");
            }
            string outputDir = Path.Combine(projectRoot, "Outputs");
            Directory.CreateDirectory(outputDir); // Ensure it exists
            string filePath = Path.Combine(outputDir, $"output{aliasPart}.txt");
            Console.WriteLine($"Writing PR details to {filePath}...");
            using (var writer = new StreamWriter(filePath, append: false)) // Overwrite the file
            {
                // Write total number of PRs at the top
                writer.WriteLine($"Total PRs: {completedPrList.Count}\n");
                int serial = 1;
                foreach (var pr in completedPrList)
                {
                    string prInfo = $"PR #{serial}\n" +
                                    $"Title: {pr.title}\n" +
                                    $"ID: {pr.id}\n" +
                                    $"Created by (display name): {pr.createdByName}\n" +
                                    $"Status: {pr.status}\n" +
                                    $"Creation Date: {pr.creationDate}\n" +
                                    $"Closed Date: {pr.closedDate}\n" +
                                    $"Url: {pr.url}\n" +
                                    $"Work Items: {string.Join(", ", pr.listOfTaskUrls)}\n" +
                                    $"Approved By: {string.Join(", ", pr.approvedBy)}\n" +
                                    $"AI Summary: {pr.aiPrSummary}\n" +
                                    $"AI Classification: {pr.aiPrClassification}\n" +
                                    "--------------------------------------------\n";

                    Console.Write(prInfo);
                    writer.Write(prInfo);
                    serial++;
                }
            }
        }

        public static async Task<string> GetPrSummaryFromDescription(PullRequest pullRequest)
        {
            // Create a list of chat messages
            var messages = new List<ChatMessage>
            {
                new SystemChatMessage("You are an assistant that generated a summary of a pull request when given the description of that pull request in not more than 2 lines."),
                new UserChatMessage($"Summarize this pr description in not more than 2 lines. {pullRequest.description}"),
            };

            // Create chat completion options  
            var options = new ChatCompletionOptions
            {
                Temperature = (float)0.7,
                MaxOutputTokenCount = 800,

                TopP = (float)0.95,
                FrequencyPenalty = (float)0,
                PresencePenalty = (float)0
            };

            TokenCredential managedIdentityCredential = new DefaultAzureCredential();
            //var modelId = "gpt-4o";
            var modelId = "gpt-4o-mini";
            var endpoint = new Uri("https://nikakkir-openai.openai.azure.com/");
            var openAIClient = new AzureOpenAIClient(endpoint, managedIdentityCredential);
            var chatClient = openAIClient.GetChatClient(modelId);


            try
            {
                // Create the chat completion request
                ChatCompletion completion = await chatClient.CompleteChatAsync(messages, options);

                // Print the response
                if (completion != null)
                {
                    var answer = completion.Content[0].Text;
                    return answer;
                }
                else
                {
                    Console.WriteLine("No response received.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return string.Empty;
        }

        public static async Task<string> GetPrClassificationFromDescription(PullRequest pullRequest)
        {
            // Create a list of chat messages
            var messages = new List<ChatMessage>
            {
                new SystemChatMessage("You are an assistant that generates a classification of a pull request when given the description of that pull request. The classification can be one of 'Self Work' or 'Engineering Improvement'. One pull request can belong to both the categories as well. The response is a single string which is either 'Self Work' and/ or 'Engineering Improvement' without quotation marks, no other text should be returned as part of response."),
                new UserChatMessage($"Classify this pr description as 'Self Work' and/ or 'Engineering Improvement' without quotation mark, with no other text as part of response. {pullRequest.description}"),
            };
            // Create chat completion options  
            var options = new ChatCompletionOptions
            {
                Temperature = (float)0.7,
                MaxOutputTokenCount = 800,
                TopP = (float)0.95,
                FrequencyPenalty = (float)0,
                PresencePenalty = (float)0
            };
            TokenCredential managedIdentityCredential = new DefaultAzureCredential();
            var modelId = "gpt-4o-mini";
            var endpoint = new Uri("https://nikakkir-openai.openai.azure.com/");
            var openAIClient = new AzureOpenAIClient(endpoint, managedIdentityCredential);
            var chatClient = openAIClient.GetChatClient(modelId);

            try
            {
                // Create the chat completion request
                ChatCompletion completion = await chatClient.CompleteChatAsync(messages, options);

                // Print the response
                if (completion != null)
                {
                    var answer = completion.Content[0].Text;
                    return answer;
                }
                else
                {
                    Console.WriteLine("No response received.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return string.Empty;
        }

        public static async Task<List<string>> GetDocumentsFromSharepoint(string siteUrl, string libraryName)
        {
            string accessToken = await GetAccessTokenForSharePoint();

            // Create an HttpClient instance
            using (HttpClient client = new HttpClient())
            {
                // Set the authorization header using the access token
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
                client.DefaultRequestHeaders.Add("Accept", "application/json");

                // Construct API endpoint to get documents from a SharePoint library and folder
                //string requestUrl = $"{siteUrl}/_api/web/lists/getbytitle('{libraryName}')/items?$filter=FileDirRef eq '/teams/ResearcherEnablement/{libraryName}' and Created ge '2025-02-25'"; // Date filter for last month
                string requestUrl = $"https://graph.microsoft.com/v1.0/sites/microsoft.sharepoint-df.com:/teams/ResearcherEnablement/"; // Date filter for last month

                // Make a GET request to the SharePoint REST API
                HttpResponseMessage response = await client.GetAsync(requestUrl);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    Console.WriteLine("Response from SharePoint: " + content);
                    // Parse the response to get the document details
                }
                else
                {
                    Console.WriteLine("Error while accessing files: " + response.ReasonPhrase);
                }
            }

            return new List<string>();
        }

        public static async Task<string> GetAccessTokenForSharePoint()
        {
            // Get Access Token
            var credential = new DefaultAzureCredential();

            // Define the SharePoint API resource URL
            //string sharepointResource = "https://microsoft.sharepoint-df.com/";
            string sharepointResource = "https://graph.microsoft.com/.default";

            // Get the access token for the resource (SharePoint)
            var tokenRequestContext = new TokenRequestContext(new[] { sharepointResource });
            var token = await credential.GetTokenAsync(tokenRequestContext);

            return token.Token;
        }

        public static async Task<AuthenticationResult> GetAccessTokenAsync()
        {
            string clientId = "e104b683-1ff5-486c-bb34-3ea394a510f1";
            string tenantId = "72f988bf-86f1-41af-91ab-2d7cd011db47";
            string redirectUri = "http://localhost";
            //string[] scopes = new string[] { "https://microsoft.sharepoint-df.com/.default" }; // SharePoint specific scope
            string[] scopes = new string[] { "https://graph.microsoft.com/.default" }; // SharePoint specific scope
            IPublicClientApplication _publicClientApp;
            _publicClientApp = PublicClientApplicationBuilder.Create(clientId)
            .WithRedirectUri(redirectUri)
            .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
            .Build();
            AuthenticationResult result = null;

            var accounts = await _publicClientApp.GetAccountsAsync();

            try
            {
                // Try to get token silently first
                result = await _publicClientApp.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException)
            {
                // Fallback to interactive login if silent authentication fails
                result = await _publicClientApp.AcquireTokenInteractive(scopes)
                    .WithPrompt(Prompt.SelectAccount)
                    .ExecuteAsync();
            }

            return result;
        }

        // Function to make a SharePoint REST API call
        public static async Task MakeSharePointApiCallAsync(string accessToken)
        {
            var sharePointSiteUrl = "https://microsoft.sharepoint-df.com/teams/ResearcherEnablement";
            var apiEndpoint = $"{sharePointSiteUrl}/_api/web/lists/getbytitle('Documents')/items"; // Accessing items in the 'Documents' library

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await client.GetAsync(apiEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    Console.WriteLine("SharePoint API Response: " + content);
                }
                else
                {
                    Console.WriteLine("Error calling SharePoint API: " + response.StatusCode);
                }
            }
        }


        public static async Task<String> GetDocuments(DateTime creationDate)
        {
            // Create a list of chat messages
            var messages = new List<ChatMessage>
            {
                new SystemChatMessage("You are a helper that reads the documents in Storage accounts and answers whatever questions are asked over it"),
                new UserChatMessage($"what materials were used in the pyramids"),
            };
            // Create chat completion options  
            var options = new ChatCompletionOptions
            {
                Temperature = (float)0.7,
                MaxOutputTokenCount = 800,
                TopP = (float)0.95,
                FrequencyPenalty = (float)0,
                PresencePenalty = (float)0
            };
            TokenCredential managedIdentityCredential = new DefaultAzureCredential();
            var modelId = "gpt-4o";
            var endpoint = new Uri("https://vvibhor-test-openai.openai.azure.com/");
            var openAIClient = new AzureOpenAIClient(endpoint, managedIdentityCredential);
            var chatClient = openAIClient.GetChatClient(modelId);

            try
            {
                // Create the chat completion request
                ChatCompletion completion = await chatClient.CompleteChatAsync(messages, options);

                // Print the response
                if (completion != null)
                {
                    var answer = completion.Content[0].Text;
                    Console.WriteLine($"{answer}");
                    return answer;
                }
                else
                {
                    Console.WriteLine("No response received.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return string.Empty;
        }

        public static async Task<List<WordprocessingDocument>> GetLocalDocuments(DateTime creationDate, string createdBy)
        {
            string directoryPath = "D:\\docs\\New folder";
            var docxFiles = Directory.GetFiles(directoryPath, "*.docx", SearchOption.AllDirectories);
            List<WordprocessingDocument> docs = new List<WordprocessingDocument>();

            Console.WriteLine("---------------------------------------------------");
            foreach (var filePath in docxFiles)
            {
                //Console.WriteLine("File: " + filePath);

                try
                {
                    var document = WordprocessingDocument.Open(filePath, false);
                    if (document.PackageProperties.Created >= creationDate && document.PackageProperties.Creator == createdBy)
                    {
                        Console.WriteLine("Document Name: " + Path.GetFileName(filePath));
                        Console.WriteLine("Creator: " + document.PackageProperties.Creator);
                        Console.WriteLine("Created on: " + document.PackageProperties.Created);

                        string summary = await GetDocumentSummaryWithAi(filePath);

                        Console.WriteLine("AI summary: " + summary);

                        Console.WriteLine("---------------------------------------------------");
                        docs.Add(document);
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
            }

            return docs;
        }

        public static async Task<string> GetDocumentSummaryWithAi(string filePath)
        {
            string content = GetWordDocumentText(filePath);
            // Create a list of chat messages
            var messages = new List<ChatMessage>
            {
                new SystemChatMessage("You are an agent which when given the contents of a document generated the summary of it in not more than 2 lines."),
                new UserChatMessage($"Give summary of this doc in not more than 2 lines, explaing what this documents does. Starting the response with this line 'In this document'. {content}"),
            };
            // Create chat completion options  
            var options = new ChatCompletionOptions
            {
                Temperature = (float)0.7,
                MaxOutputTokenCount = 800,
                TopP = (float)0.95,
                FrequencyPenalty = (float)0,
                PresencePenalty = (float)0
            };
            TokenCredential managedIdentityCredential = new DefaultAzureCredential();
            var modelId = "gpt-4o";
            var endpoint = new Uri("https://msechackathon-eastus2.openai.azure.com/");
            var openAIClient = new AzureOpenAIClient(endpoint, managedIdentityCredential);
            var chatClient = openAIClient.GetChatClient(modelId);

            try
            {
                // Create the chat completion request
                ChatCompletion completion = await chatClient.CompleteChatAsync(messages, options);

                // Print the response
                if (completion != null)
                {
                    var answer = completion.Content[0].Text;
                    return answer;
                }
                else
                {
                    Console.WriteLine("No response received.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return string.Empty;
        }

        public static string GetWordDocumentText(string filePath)
        {
            StringBuilder content = new StringBuilder();

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    // Access the main document part
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    // Iterate through all paragraphs and extract text
                    foreach (var paragraph in body.Elements<Paragraph>())
                    {
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            foreach (var text in run.Elements<Text>())
                            {
                                content.Append(text.Text);
                            }
                        }
                        content.AppendLine(); // Add a line break for each paragraph
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error reading the Word document: " + ex.Message);
            }

            return content.ToString();
        }

        public static string GetAuthor(string documentPath)
        {
            using (var document = WordprocessingDocument.Open(documentPath, false))
            {
                return document.PackageProperties.Creator;
            }
        }

        public static string GetOfficeFileAuthor(string filePath)
        {
            try
            {
                // Using Microsoft.Office.Interop to get author from Office documents
                var application = new Microsoft.Office.Interop.Word.Application();
                var document = application.Documents.Open(filePath);
                var properties = document.BuiltInDocumentProperties as Microsoft.Office.Core.DocumentProperties;
                string author = "Unknown Author";
                if (properties != null && properties != null)
                {
                    author = properties["Author"].Value.ToString();
                }
                document.Close();
                application.Quit();

                return author;
            }
            catch (Exception)
            {
                return "Unknown Author";
            }
        }

        public static async Task<List<PullRequest>> GetAllPrs(string organization, List<string> projects, string userAlias, string prStatus, DateTime completionDateTime, DateTime creationDateTime)
        {
            List<PullRequest> PrList = new List<PullRequest>();
            var credential = new DefaultAzureCredential();
            var tokenRequestContext = new TokenRequestContext(new[] { "499b84ac-1321-427f-aa17-267ca6975798/.default" }); // Azure DevOps resource
            var token = await credential.GetTokenAsync(tokenRequestContext);
            foreach (var project in projects)
            {
                Console.WriteLine($"\n==============================\nProcessing project: {project}\n==============================");
                string url = $"https://dev.azure.com/{organization}/{project}/_apis/git/repositories?api-version=7.1-preview.1";
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
                    HttpResponseMessage response;
                    try
                    {
                        response = await client.GetAsync(url);
                    }
                    catch (Exception e)
                    {
                        continue; // Skip to the next project if there's an error
                    }
                    if (response.IsSuccessStatusCode)
                    {
                        var content = await response.Content.ReadAsStringAsync();
                        JObject jsonResponse = JObject.Parse(content);
                        JArray listOfRepos = (JArray)jsonResponse["value"];
                        int totalRepos = listOfRepos.Count;
                        int repoSerialNumber = 1;
                        foreach (var repo in listOfRepos)
                        {
                            string repoId = repo["id"].ToString();
                            string repoName = repo["name"].ToString();
                            Console.Write($"\r \t Processing repository [{repoSerialNumber}/{totalRepos}]: {repoName}                           \n");
                            string prUrlEndpoint = $"https://dev.azure.com/{organization}/{project}/_apis/git/repositories/{repoId}/pullrequests?searchCriteria.status=all&api-version=6.0&$top=1000&$skip=0";
                            HttpResponseMessage repoResponse;
                            try
                            {
                                repoResponse = await client.GetAsync(prUrlEndpoint);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error fetching PRs for repo {repoName}: {ex.Message}");
                                continue; // Skip to next repo
                            }
                            if (repoResponse.IsSuccessStatusCode)
                            {
                                var repoContent = await repoResponse.Content.ReadAsStringAsync();
                                JObject repoJsonResponse = JObject.Parse(repoContent);
                                JArray repos = (JArray)repoJsonResponse["value"];
                                foreach (var pr in repos)
                                {
                                    bool closedInTime = true;
                                    string prCloseTime = "null";
                                    if (pr["closedDate"] != null)
                                    {
                                        DateTime prCloseDateTime = DateTime.Parse(pr["closedDate"].ToString());
                                        closedInTime = prCloseDateTime <= completionDateTime;
                                        prCloseTime = pr["closedDate"].ToString();
                                    }
                                    bool createdInTime = true;
                                    if (pr["creationDate"] != null)
                                    {
                                        DateTime prCreationDateTime = DateTime.Parse(pr["creationDate"].ToString());
                                        createdInTime = prCreationDateTime >= creationDateTime;
                                    }
                                    bool createdByUser = pr["createdBy"]["uniqueName"].ToString() == userAlias;
                                    if (createdByUser && closedInTime && createdInTime && (string.IsNullOrEmpty(prStatus) || pr["status"].ToString() == prStatus))
                                    {
                                        string prTitle = pr["title"].ToString();
                                        string prId = pr["pullRequestId"].ToString();
                                        string createdDate = pr["creationDate"].ToString();
                                        string state = pr["status"].ToString();
                                        string prCreatedBy = pr["createdBy"]["uniqueName"].ToString();
                                        string prCreatedByName = pr["createdBy"]["displayName"].ToString();
                                        string prDetailsUrl = $"https://dev.azure.com/{organization}/{project}/_apis/git/pullrequests/{prId}?api-version=6.0";
                                        HttpResponseMessage prResponse = await client.GetAsync(prDetailsUrl);
                                        string prResponseBody = await prResponse.Content.ReadAsStringAsync();
                                        JObject prJson = JObject.Parse(prResponseBody);
                                        string newPrDescription = prJson["description"]?.ToString() ?? "No description available.";
                                        string prDescription = newPrDescription;
                                        string prTask = pr["workItemRefs"] != null ? pr["workItemRefs"]["url"].ToString() : "no work item";
                                        string prUrl = pr["url"] != null ? Helper.ConvertToVisualStudioUrl(pr["url"].ToString(), prId, project, repoName) : "no url";
                                        JToken prApprovers = pr["reviewers"];
                                        List<string> prApproversList = new List<string>();
                                        foreach (var prR in prApprovers)
                                        {
                                            if (prR["uniqueName"] != null && !prR["uniqueName"].ToString().StartsWith("vstfs") && (prR["vote"].ToString() == "10" || prR["vote"].ToString() == "5"))
                                            {
                                                prApproversList.Add(prR["uniqueName"].ToString());
                                            }
                                        }
                                        string workItemUrl = $"https://dev.azure.com/{organization}/{project}/_apis/git/pullrequests/{prId}?api-version=7.1";
                                        List<string> workItems = await Helper.GetWorkItemsForPullRequest(workItemUrl);
                                        PullRequest pullRequest = new PullRequest
                                        {
                                            description = prDescription,
                                            id = prId,
                                            title = prTitle,
                                            creationDate = pr["creationDate"] != null ? pr["creationDate"].ToString() : null,
                                            closedDate = pr["closedDate"] != null ? pr["closedDate"].ToString() : null,
                                            listOfTaskUrls = workItems,
                                            status = state,
                                            url = prUrl,
                                            approvedBy = prApproversList,
                                            createdByName = prCreatedByName,
                                        };
                                        PrList.Add(pullRequest);
                                    }
                                }
                            }
                            repoSerialNumber++;
                        }
                        Console.WriteLine(); // Newline after all repos in this project
                    }
                }
            }
            return PrList;
        }

        // Get user alias (prompt and return)
        public static string GetUserAlias()
        {
            Console.WriteLine("Enter your alias (do not include '@microsoft.com'):");
            string userInputAlias = Console.ReadLine();
            return userInputAlias + "@microsoft.com";
        }

        // Project selection helper
        public static List<string> SelectProjects()
        {
            var allProjects = new List<string>
            {
                "Windows%20Defender",
                "OS",
                "Safety%20Platform",
                "WDATP"
            };
            Console.WriteLine("Select one or more projects by entering their indexes separated by commas:");
            for (int i = 0; i < allProjects.Count; i++)
            {
                Console.WriteLine($"{i + 1}: {allProjects[i]}");
            }
            Console.Write("Your selection: ");
            string projectIndexesInput = Console.ReadLine();
            var selectedIndexes = projectIndexesInput.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(idx => int.TryParse(idx, out int val) ? val : -1)
                .Where(idx => idx >= 1 && idx <= allProjects.Count)
                .ToList();
            var listOfProjects = allProjects.Where((proj, idx) => selectedIndexes.Contains(idx + 1)).ToList();
            return listOfProjects;
        }

        // Days back helper
        public static int GetDaysBack()
        {
            Console.WriteLine("Enter the number of days in the past to look for PRs (e.g., 90):");
            int daysBack = 90;
            string daysInput = Console.ReadLine();
            if (!int.TryParse(daysInput, out daysBack) || daysBack < 1)
            {
                Console.WriteLine("Invalid input. Defaulting to 90 days.");
                daysBack = 90;
            }
            return daysBack;
        }
    }
}

using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using Hack2025;
using Hack2025.Models;
using Newtonsoft.Json.Linq;
using Azure.AI.OpenAI;
using Azure.Core;
using Azure.Identity;
using Microsoft.Identity.Client;
using static Program;
using DocumentFormat.OpenXml.Packaging;

public class Program
{
    static async Task Main(string[] args)
    {
        var stopwatch = Stopwatch.StartNew(); // Start timing

        // Prompt user for alias
        string userAlias = Helper.GetUserAlias();

        // Project options
        var listOfProjects = Helper.SelectProjects();
        if (listOfProjects.Count == 0)
        {
            Console.WriteLine("No valid projects selected. Exiting.");
            return;
        }

        string organization = "microsoft"; // Replace with your Azure DevOps organization

        // Ask user for time range in days
        int daysBack = Helper.GetDaysBack();

        DateTime creationDateTime = DateTime.Now.AddDays(-daysBack);
        DateTime completionDateTime = DateTime.Now;

        List<PullRequest> prList = await Helper.GetAllPrs(organization, listOfProjects, userAlias, prStatus: string.Empty, completionDateTime, creationDateTime);
        Console.WriteLine("Total number of your PRs: " + prList.Count);

        //for AI review
        int totalPrs = prList.Count;
        for (int i = 0; i < totalPrs; i++)
        {
            var pr = prList[i];
            Console.Write($"Processing PR [{i + 1}/{totalPrs}]......\r");
            pr.aiPrSummary = await Helper.GetPrSummaryFromDescription(pr);
            pr.aiPrClassification = await Helper.GetPrClassificationFromDescription(pr);
        }
        Console.WriteLine(); // Move to next line after progress bar

        Console.WriteLine("\n*****************************************************\n");
        Console.WriteLine("Getting your completed PR details...");
        Console.WriteLine("*****************************************************\n\n");
        Helper.PrintAndWriteListOfPrs(prList, userAlias);
        Console.WriteLine($"output_{userAlias.Split('@')[0]}.txt has been updated with the latest info.");
        Console.WriteLine("*****************************************************\n");

        stopwatch.Stop(); // Stop timing
        Console.WriteLine($"Total script run time: {stopwatch.Elapsed}");
    }
}

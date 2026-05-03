using System.Collections.Generic;

namespace Hack2025.Models
{
    public class PullRequest
    {
        public string description { get; set; }
        public string id { get; set; }
        public string createdByName { get; set; }
        public string title { get; set; }
        public string creationDate { get; set; }
        public string? closedDate { get; set; }
        public List<string>? listOfTaskUrls { get; set; }
        public string status { get; set; }
        public string url { get; set; }
        public string? aiPrSummary { get; set; }
        public string? aiPrClassification { get; set; }
        public List<string> approvedBy { get; set; }
    }
}

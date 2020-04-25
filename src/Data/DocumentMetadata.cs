using System.Text.Json.Serialization;

namespace BlazorOfficeAddIn.Data
{
    public class DocumentMetadata
    {
        [JsonPropertyName("title")]
        public string Title { get; set; }

        [JsonPropertyName("subject")]
        public string Subject { get; set; }
    }
}
using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace BlazorOfficeAddIn.Data
{
    public class DocumentMetadata
    {
        [JsonPropertyName("title")]
        [Required]
        [MaxLength(20)]
        public string Title { get; set; }

        [JsonPropertyName("subject")]
        [MaxLength(50)]
        public string Subject { get; set; }
    }
}
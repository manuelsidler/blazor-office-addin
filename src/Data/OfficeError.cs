using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace BlazorOfficeAddIn.Data
{
    // https://docs.microsoft.com/en-us/javascript/api/office/officeextension.error?view=common-js
    public class OfficeError
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("code")]
        public string Code { get; set; }

        [JsonPropertyName("traceMessages")]
        public List<string> TraceMessages { get; set; }

        [JsonPropertyName("innerError")]
        public string InnerError { get; set; }

        [JsonPropertyName("debugInfo")]
        public DebugInfo DebugInfo { get; set; }
    }

    // https://docs.microsoft.com/en-us/javascript/api/office/officeextension.debuginfo?view=common-js
    public class DebugInfo
    {
        [JsonPropertyName("code")]
        public string Code { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("errorLocation")]
        public string ErrorLocation { get; set; }

        [JsonPropertyName("statement")]
        public string Statement { get; set; }

        [JsonPropertyName("surroundingStatements")]
        public List<string> SurroundingStatements { get; set; }

        [JsonPropertyName("fullStatements")]
        public List<string> FullStatements { get; set; }
    }
}
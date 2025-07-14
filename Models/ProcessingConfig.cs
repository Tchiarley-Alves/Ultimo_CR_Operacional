using System;
using System.Collections.Generic;

namespace ExcelProcessor.Models
{
    public record ProcessingConfig
    {
        public string InputFile { get; init; } = "";
        public string OutputFile { get; init; } = "";
        public string SheetName { get; init; } = "Planilha1";
        public HashSet<string> SuffixesToFilter { get; init; } = new()
        {
            "90600", "90610", "92600", "92610", "92670", "92660",
            "90660", "92640", "24099", "24010", "24024", "23019",
            "24014", "24018", "24009", "20580", "20500", "40900"
        };
    }

    public class ProcessingException : Exception
    {
        public ProcessingException(string message) : base(message) { }
        public ProcessingException(string message, Exception innerException) : base(message, innerException) { }
    }

    public record ColumnMapping
    {
        public int CentroCusto { get; init; }
        public int ValidoDesde { get; init; }
        public int ValidoAte { get; init; }
        public int Equipamento { get; init; }
    }

    public class DataRow
    {
        public List<object?> Values { get; set; } = new();
        public DateTime? ValidoDesde { get; set; }
        public DateTime? ValidoAte { get; set; }
        public string? Equipamento { get; set; }
    }
}
using ExcelProcessor.Models;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelProcessor.Services
{
    public interface IExcelProcessorService
    {
        event EventHandler<string>? LogMessage;
        event EventHandler<int>? ProgressChanged;
        Task<bool> ProcessAsync(ProcessingConfig config, CancellationToken cancellationToken = default);
    }
}
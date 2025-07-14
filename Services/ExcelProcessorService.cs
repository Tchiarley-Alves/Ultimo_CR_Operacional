using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExcelProcessor.Models;

namespace ExcelProcessor.Services
{
    public class ExcelProcessorService : IExcelProcessorService
    {
        public event EventHandler<string>? LogMessage;
        public event EventHandler<int>? ProgressChanged;

        private void OnLogMessage(string message)
        {
            LogMessage?.Invoke(this, message);
        }

        private void OnProgressChanged(int progress)
        {
            ProgressChanged?.Invoke(this, progress);
        }

        public async Task<bool> ProcessAsync(ProcessingConfig config, CancellationToken cancellationToken = default)
        {
            try
            {
                OnLogMessage("Iniciando processamento da planilha...");
                OnProgressChanged(0);

                // Validar arquivos
                if (!File.Exists(config.InputFile))
                {
                    throw new ProcessingException($"Arquivo não encontrado: {config.InputFile}");
                }

                OnProgressChanged(10);
                await Task.Delay(100, cancellationToken);

                // Carregar planilha
                OnLogMessage("Carregando planilha...");
                using var workbook = new XLWorkbook(config.InputFile);
                
                if (!workbook.Worksheets.Contains(config.SheetName))
                {
                    throw new ProcessingException($"Aba '{config.SheetName}' não encontrada");
                }

                var worksheet = workbook.Worksheet(config.SheetName);
                OnProgressChanged(20);
                await Task.Delay(100, cancellationToken);

                // Mapear cabeçalhos
                OnLogMessage("Mapeando cabeçalhos...");
                var columnMapping = MapHeaders(worksheet);
                OnProgressChanged(30);
                await Task.Delay(100, cancellationToken);

                // Filtrar por sufixo
                OnLogMessage("Filtrando linhas por sufixo...");
                int removedBySuffix = FilterBySuffix(worksheet, columnMapping, config.SuffixesToFilter);
                OnLogMessage($"Removidas {removedBySuffix} linhas por filtro de sufixo");
                OnProgressChanged(50);
                await Task.Delay(100, cancellationToken);

                // Processar e ordenar dados
                OnLogMessage("Processando e ordenando dados...");
                ProcessAndSortData(worksheet, columnMapping);
                OnProgressChanged(70);
                await Task.Delay(100, cancellationToken);

                // Remover duplicados
                OnLogMessage("Removendo duplicados...");
                int removedDuplicates = RemoveDuplicates(worksheet, columnMapping);
                OnLogMessage($"Removidas {removedDuplicates} linhas duplicadas");
                OnProgressChanged(90);
                await Task.Delay(100, cancellationToken);

                // Salvar planilha
                OnLogMessage("Salvando planilha...");
                Directory.CreateDirectory(Path.GetDirectoryName(config.OutputFile) ?? "");
                workbook.SaveAs(config.OutputFile);
                OnProgressChanged(100);

                OnLogMessage($"Processamento concluído com sucesso! Arquivo salvo em: {config.OutputFile}");
                return true;
            }
            catch (Exception ex)
            {
                OnLogMessage($"Erro: {ex.Message}");
                throw new ProcessingException($"Falha no processamento: {ex.Message}", ex);
            }
        }

        private ColumnMapping MapHeaders(IXLWorksheet worksheet)
        {
            var headers = new Dictionary<string, int>();
            var headerRow = worksheet.Row(1);
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            
            for (int col = 1; col <= lastColumn; col++)
            {
                var cellValue = headerRow.Cell(col).GetString();
                if (!string.IsNullOrEmpty(cellValue))
                {
                    headers[cellValue] = col;
                }
            }

            try
            {
                return new ColumnMapping
                {
                    CentroCusto = headers["Centro custo"],
                    ValidoDesde = headers["Válido desde"],
                    ValidoAte = headers["Válido até"],
                    Equipamento = headers["Equipamento"]
                };
            }
            catch (KeyNotFoundException ex)
            {
                throw new ProcessingException($"Cabeçalho obrigatório não encontrado: {ex.Message}");
            }
        }

        private int FilterBySuffix(IXLWorksheet worksheet, ColumnMapping columnMapping, HashSet<string> suffixes)
        {
            var rowsToDelete = new List<int>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

            for (int row = 2; row <= lastRow; row++)
            {
                var cellValue = worksheet.Cell(row, columnMapping.CentroCusto).GetString();
                if (cellValue.Length >= 5 && suffixes.Contains(cellValue[^5..]))
                {
                    rowsToDelete.Add(row);
                }
            }

            foreach (var rowNum in rowsToDelete.OrderByDescending(x => x))
            {
                worksheet.Row(rowNum).Delete();
            }

            return rowsToDelete.Count;
        }

        private void ProcessAndSortData(IXLWorksheet worksheet, ColumnMapping columnMapping)
        {
            var rowsData = new List<DataRow>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            var lastCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

            for (int row = 2; row <= lastRow; row++)
            {
                var dataRow = new DataRow();
                
                for (int col = 1; col <= lastCol; col++)
                {
                    var cellValue = worksheet.Cell(row, col).Value;
                    dataRow.Values.Add(cellValue);
                }

                dataRow.ValidoDesde = ParseDate(worksheet.Cell(row, columnMapping.ValidoDesde).Value);
                dataRow.ValidoAte = ParseDate(worksheet.Cell(row, columnMapping.ValidoAte).Value);
                dataRow.Equipamento = worksheet.Cell(row, columnMapping.Equipamento).GetString();

                rowsData.Add(dataRow);
            }

            var sortedData = rowsData
                .OrderBy(x => x.Equipamento ?? string.Empty)
                .ThenByDescending(x => x.ValidoDesde ?? DateTime.MinValue)
                .ThenByDescending(x => x.ValidoAte ?? DateTime.MinValue)
                .ToList();

            WriteDataToSheet(worksheet, sortedData);
        }

        private DateTime? ParseDate(object? dateValue)
        {
            return dateValue switch
            {
                DateTime dt => dt,
                string str when DateTime.TryParseExact(str, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out var parsed) => parsed,
                string str when DateTime.TryParse(str, out var parsed2) => parsed2,
                _ => null
            };
        }

        private void WriteDataToSheet(IXLWorksheet worksheet, List<DataRow> rowsData)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            if (lastRow > 1)
            {
                worksheet.Range($"2:{lastRow}").Delete(XLShiftDeletedCells.ShiftCellsUp);
            }

            for (int rowIdx = 0; rowIdx < rowsData.Count; rowIdx++)
            {
                var xlRow = worksheet.Row(rowIdx + 2);
                for (int colIdx = 0; colIdx < rowsData[rowIdx].Values.Count; colIdx++)
                {
                    var value = rowsData[rowIdx].Values[colIdx];
                    xlRow.Cell(colIdx + 1).Value = XLCellValue.FromObject(value);
                }
            }
        }

        private int RemoveDuplicates(IXLWorksheet worksheet, ColumnMapping columnMapping)
        {
            var seenValues = new HashSet<string>();
            var rowsToDelete = new List<int>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

            for (int row = 2; row <= lastRow; row++)
            {
                var equipmentValue = worksheet.Cell(row, columnMapping.Equipamento).GetString();

                if (!string.IsNullOrEmpty(equipmentValue))
                {
                    if (seenValues.Contains(equipmentValue))
                    {
                        rowsToDelete.Add(row);
                    }
                    else
                    {
                        seenValues.Add(equipmentValue);
                    }
                }
            }

            foreach (var rowNum in rowsToDelete.OrderByDescending(x => x))
            {
                worksheet.Row(rowNum).Delete();
            }

            return rowsToDelete.Count;
        }
    }
}
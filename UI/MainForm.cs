using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelProcessor.Models;
using ExcelProcessor.Services;

namespace ExcelProcessor.UI
{
    public partial class MainForm : Form
    {
        private readonly IExcelProcessorService _processorService;
        private CancellationTokenSource? _cancellationTokenSource;

        public MainForm(IExcelProcessorService processorService)
        {
            _processorService = processorService;
            InitializeComponent();
            SetupEventHandlers();
        }

        private void SetupEventHandlers()
        {
            _processorService.LogMessage += OnLogMessage;
            _processorService.ProgressChanged += OnProgressChanged;
        }

        private void OnLogMessage(object? sender, string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(msg => logTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\r\n")), message);
            }
            else
            {
                logTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
            }
        }

        private void OnProgressChanged(object? sender, int progress)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<int>(p => progressBar.Value = p), progress);
            }
            else
            {
                progressBar.Value = progress;
            }
        }

        private void btnSelectInput_Click(object sender, EventArgs e)
        {
            using var openFileDialog = new OpenFileDialog
            {
                Title = "Selecione o arquivo Excel de entrada",
                Filter = "Arquivos Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*",
                FilterIndex = 1
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtInputFile.Text = openFileDialog.FileName;
                
                // Sugerir nome de saída baseado no arquivo de entrada
                if (string.IsNullOrEmpty(txtOutputFile.Text))
                {
                    var directory = Path.GetDirectoryName(openFileDialog.FileName);
                    var filename = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                    var extension = Path.GetExtension(openFileDialog.FileName);
                    txtOutputFile.Text = Path.Combine(directory ?? "", $"{filename}_processado{extension}");
                }
            }
        }

        private void btnSelectOutput_Click(object sender, EventArgs e)
        {
            using var saveFileDialog = new SaveFileDialog
            {
                Title = "Selecione onde salvar o arquivo processado",
                Filter = "Arquivos Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*",
                FilterIndex = 1
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtOutputFile.Text = saveFileDialog.FileName;
            }
        }

        private async void btnProcess_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtInputFile.Text))
            {
                MessageBox.Show("Por favor, selecione o arquivo de entrada.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtOutputFile.Text))
            {
                MessageBox.Show("Por favor, selecione o arquivo de saída.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                SetProcessingState(true);
                logTextBox.Clear();
                progressBar.Value = 0;

                var suffixes = txtSuffixes.Text
                    .Split(',')
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToHashSet();

                var config = new ProcessingConfig
                {
                    InputFile = txtInputFile.Text,
                    OutputFile = txtOutputFile.Text,
                    SheetName = txtSheetName.Text,
                    SuffixesToFilter = suffixes
                };

                _cancellationTokenSource = new CancellationTokenSource();
                bool success = await _processorService.ProcessAsync(config, _cancellationTokenSource.Token);

                if (success)
                {
                    MessageBox.Show("Processamento concluído com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (ProcessingException ex)
            {
                MessageBox.Show($"Erro no processamento: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Processamento cancelado pelo usuário.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro inesperado: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetProcessingState(false);
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _cancellationTokenSource?.Cancel();
        }

        private void SetProcessingState(bool isProcessing)
        {
            btnProcess.Enabled = !isProcessing;
            btnCancel.Enabled = isProcessing;
            btnSelectInput.Enabled = !isProcessing;
            btnSelectOutput.Enabled = !isProcessing;
            txtInputFile.Enabled = !isProcessing;
            txtOutputFile.Enabled = !isProcessing;
            txtSheetName.Enabled = !isProcessing;
            txtSuffixes.Enabled = !isProcessing;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _processorService.LogMessage -= OnLogMessage;
                _processorService.ProgressChanged -= OnProgressChanged;
                _cancellationTokenSource?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
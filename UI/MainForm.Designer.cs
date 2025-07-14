using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelProcessor.UI
{
    partial class MainForm
    {
        private TextBox txtInputFile;
        private TextBox txtOutputFile;
        private TextBox txtSheetName;
        private TextBox txtSuffixes;
        private TextBox logTextBox;
        private Button btnSelectInput;
        private Button btnSelectOutput;
        private Button btnProcess;
        private Button btnCancel;
        private ProgressBar progressBar;
        private Label lblInputFile;
        private Label lblOutputFile;
        private Label lblSheetName;
        private Label lblSuffixes;
        private Label lblLog;
        private GroupBox groupBoxFiles;
        private GroupBox groupBoxSettings;
        private GroupBox groupBoxLog;

        private void InitializeComponent()
        {
            this.txtInputFile = new TextBox();
            this.txtOutputFile = new TextBox();
            this.txtSheetName = new TextBox();
            this.txtSuffixes = new TextBox();
            this.logTextBox = new TextBox();
            this.btnSelectInput = new Button();
            this.btnSelectOutput = new Button();
            this.btnProcess = new Button();
            this.btnCancel = new Button();
            this.progressBar = new ProgressBar();
            this.lblInputFile = new Label();
            this.lblOutputFile = new Label();
            this.lblSheetName = new Label();
            this.lblSuffixes = new Label();
            this.lblLog = new Label();
            this.groupBoxFiles = new GroupBox();
            this.groupBoxSettings = new GroupBox();
            this.groupBoxLog = new GroupBox();
            this.groupBoxFiles.SuspendLayout();
            this.groupBoxSettings.SuspendLayout();
            this.groupBoxLog.SuspendLayout();
            this.SuspendLayout();

            // Form
            this.Text = "Último CR Operacional VIX Logística";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(600, 500);

            // Group Box Files
            this.groupBoxFiles.Text = "Arquivos";
            this.groupBoxFiles.Location = new Point(12, 12);
            this.groupBoxFiles.Size = new Size(760, 100);
            this.groupBoxFiles.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Input File
            this.lblInputFile.Text = "Arquivo de entrada:";
            this.lblInputFile.Location = new Point(10, 25);
            this.lblInputFile.Size = new Size(120, 20);

            this.txtInputFile.Location = new Point(140, 22);
            this.txtInputFile.Size = new Size(500, 23);
            this.txtInputFile.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtInputFile.ReadOnly = true;

            this.btnSelectInput.Text = "Selecionar";
            this.btnSelectInput.Location = new Point(650, 21);
            this.btnSelectInput.Size = new Size(100, 25);
            this.btnSelectInput.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // Output File
            this.lblOutputFile.Text = "Arquivo de saída:";
            this.lblOutputFile.Location = new Point(10, 60);
            this.lblOutputFile.Size = new Size(120, 20);

            this.txtOutputFile.Location = new Point(140, 57);
            this.txtOutputFile.Size = new Size(500, 23);
            this.txtOutputFile.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtOutputFile.ReadOnly = true;

            this.btnSelectOutput.Text = "Selecionar";
            this.btnSelectOutput.Location = new Point(650, 56);
            this.btnSelectOutput.Size = new Size(100, 25);
            this.btnSelectOutput.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // Group Box Settings
            this.groupBoxSettings.Text = "Configurações";
            this.groupBoxSettings.Location = new Point(12, 120);
            this.groupBoxSettings.Size = new Size(760, 100);
            this.groupBoxSettings.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Sheet Name
            this.lblSheetName.Text = "Nome da aba:";
            this.lblSheetName.Location = new Point(10, 25);
            this.lblSheetName.Size = new Size(120, 20);

            this.txtSheetName.Location = new Point(140, 22);
            this.txtSheetName.Size = new Size(200, 23);
            this.txtSheetName.Text = "Planilha1";

            // Suffixes
            this.lblSuffixes.Text = "Sufixos a filtrar:";
            this.lblSuffixes.Location = new Point(10, 60);
            this.lblSuffixes.Size = new Size(120, 20);

            this.txtSuffixes.Location = new Point(140, 57);
            this.txtSuffixes.Size = new Size(610, 23);
            this.txtSuffixes.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtSuffixes.Text = "90600,90610,92600,92610,92670,92660,90660,92640,24099,24010,24024,23019,24014,24018,24009,20580,20500,40900";

            // Progress Bar
            this.progressBar.Location = new Point(12, 230);
            this.progressBar.Size = new Size(650, 23);
            this.progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Buttons
            this.btnProcess.Text = "Processar";
            this.btnProcess.Location = new Point(670, 230);
            this.btnProcess.Size = new Size(100, 25);
            this.btnProcess.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.btnProcess.BackColor = Color.LightGreen;

            this.btnCancel.Text = "Cancelar";
            this.btnCancel.Location = new Point(670, 260);
            this.btnCancel.Size = new Size(100, 25);
            this.btnCancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.btnCancel.Enabled = false;

            // Group Box Log
            this.groupBoxLog.Text = "Log de Processamento";
            this.groupBoxLog.Location = new Point(12, 290);
            this.groupBoxLog.Size = new Size(760, 260);
            this.groupBoxLog.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

            // Log TextBox
            this.logTextBox.Location = new Point(10, 20);
            this.logTextBox.Size = new Size(740, 230);
            this.logTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            this.logTextBox.Multiline = true;
            this.logTextBox.ReadOnly = true;
            this.logTextBox.ScrollBars = ScrollBars.Vertical;
            this.logTextBox.Font = new Font("Consolas", 9);

            // Add controls to groups
            this.groupBoxFiles.Controls.Add(this.lblInputFile);
            this.groupBoxFiles.Controls.Add(this.txtInputFile);
            this.groupBoxFiles.Controls.Add(this.btnSelectInput);
            this.groupBoxFiles.Controls.Add(this.lblOutputFile);
            this.groupBoxFiles.Controls.Add(this.txtOutputFile);
            this.groupBoxFiles.Controls.Add(this.btnSelectOutput);

            this.groupBoxSettings.Controls.Add(this.lblSheetName);
            this.groupBoxSettings.Controls.Add(this.txtSheetName);
            this.groupBoxSettings.Controls.Add(this.lblSuffixes);
            this.groupBoxSettings.Controls.Add(this.txtSuffixes);

            this.groupBoxLog.Controls.Add(this.logTextBox);

            // Event handlers
            this.btnSelectInput.Click += new EventHandler(this.btnSelectInput_Click);
            this.btnSelectOutput.Click += new EventHandler(this.btnSelectOutput_Click);
            this.btnProcess.Click += new EventHandler(this.btnProcess_Click);
            this.btnCancel.Click += new EventHandler(this.btnCancel_Click);

            // Add controls to form
            this.Controls.Add(this.groupBoxFiles);
            this.Controls.Add(this.groupBoxSettings);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBoxLog);

            this.groupBoxFiles.ResumeLayout(false);
            this.groupBoxSettings.ResumeLayout(false);
            this.groupBoxLog.ResumeLayout(false);
            this.ResumeLayout(false);
        }
    }
}
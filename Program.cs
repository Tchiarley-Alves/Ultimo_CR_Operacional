using System;
using System.Windows.Forms;
using ExcelProcessor.Services;
using ExcelProcessor.UI;

namespace ExcelProcessor
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Configurar serviços
            var processorService = new ExcelProcessorService();
            
            // Criar e executar a aplicação
            using var mainForm = new MainForm(processorService);
            Application.Run(mainForm);
        }
    }
}
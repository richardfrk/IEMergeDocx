using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace ieMergeDocx
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            /// Usando Office.Interop
            /// Definindo "Application" e "Document".
            var App = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document documentSource = null, documentTarget = null;

            try
            {
                string      targetFile = @"C:\Users\Richard Frank\Desktop\IE\TargetFiles\base.docx";
                string[]    appendFile = Directory.GetFiles(@"C:\Users\Richard Frank\Desktop\IE\AppendFiles", "*.docx");
                var         outputFile = @"C:\Users\Richard Frank\Desktop\IE\OutputFile.docx";

                /// Abrindo arquivo "Target".
                var fileNameTarget = targetFile;
                documentTarget = App.Documents.Open(fileNameTarget, Type.Missing, true);

                for (int i = 0; i < appendFile.Length; i++)
                {
                    /// Abrindo arquivo.
                    var fileNameSource = @appendFile[i];
                    documentSource = App.Documents.Open(fileNameSource, Type.Missing, true);

                    /// Copiando conteúdo do arquivo "Append".
                    Range sourceRange = documentSource.Content;
                    sourceRange.Copy();

                    /// Verificando tamanho do "conteúdo" do arquivo "Target".
                    Range rng = documentTarget.Content;

                    /// Definindo "Range" do arquivo "Target".
                    rng.SetRange(documentTarget.Content.End, documentTarget.Content.End);
                    rng.Paste();

                    /// Finalizando arquivo "Source".
                    if (documentSource != null)
                        documentSource.Close(false);

                }

                /// Salvando arquivo gerado.
                documentTarget.SaveAs(outputFile);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                if (documentTarget != null)
                    documentTarget.Close();

                if (App != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(App);

                App = null;
                documentSource = null;
                documentTarget = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Console.WriteLine("Processo Finalizado.");
            }

        }
    }
}

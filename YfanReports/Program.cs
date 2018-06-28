using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using VotacaoEstampas.Model;

namespace YfanReports
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //diretorio local
                string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                DirectoryInfo outputDir = new DirectoryInfo(dir);

                Console.WriteLine("Criando relatório da votação corrente..");
                string output = CriarRelatorio(outputDir);
                Console.WriteLine("Relatório criado: {0}", output);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Falha ao exportar relatório: Error: {0}", ex.Message);
                Console.WriteLine("Press the return key to exit...");
                Console.Read();
            }
        }

        private static string CriarRelatorio(DirectoryInfo outputDir)
        {
            try
            {
                //carrega dados das votacoes/clientes/estampas
                var colecaoXml = File.ReadAllText(outputDir.FullName + @"\colecao.xml");
                var colecao = DeserializeObject(colecaoXml);


                FileInfo newFile = new FileInfo(outputDir.FullName + @"\RelatorioYFAN.xlsx");
                if (newFile.Exists)
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(outputDir.FullName + @"\RelatorioYFAN.xlsx");
                }
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    // add a new worksheet to the empty workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Votações");

                    //Add the headers
                    worksheet.Cells[1, 1].Value = "Nome";
                    worksheet.Cells[1, 2].Value = "Email";
                    worksheet.Cells[1, 3].Value = "Telefone";
                    worksheet.Cells[1, 4].Value = "Data Votação";
                    var estampas = CarregarEstampas(outputDir);
                    int coluna = 5;
                    foreach (Bitmap estampa in estampas)
                    {
                        ExcelPicture pic = worksheet.Drawings.AddPicture("estampa_" + 0 + "_" + coluna, estampa);
                        pic.SetSize(100, 100);
                        pic.SetPosition(0, 0, coluna - 1, 4);
                        coluna++;
                    }

                    int linha = 2;
                    foreach (Votacao votacao in colecao.Votacoes)
                    {
                        worksheet.Cells[linha, 1].Value = votacao.Cliente.Nome;
                        worksheet.Cells[linha, 2].Value = votacao.Cliente.Email;
                        worksheet.Cells[linha, 3].Value = votacao.Cliente.Telefone;
                        worksheet.Cells[linha, 4].Value = votacao.Data.ToString(@"dd\/MM\/yyyy HH:mm");
                        coluna = 5;
                        foreach (bool voto in votacao.Votos)
                        {
                            worksheet.Cells[linha, coluna].Value = voto ? "Sim" : "Não";
                            coluna++;
                        }
                        linha++;
                    }

                    //Ok now format the values;
                    using (var range = worksheet.Cells[1, 1, 1, 4])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.BlueViolet);
                        range.Style.Font.Color.SetColor(Color.White);
                        range.Style.Font.Size = 15f;
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        range.Style.Border.Left.Color.SetColor(Color.White);
                    }

                    worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells
                    coluna = 5;
                    foreach (Bitmap estampa in estampas)
                    {
                        worksheet.Column(coluna).Width = 15; //hack
                        coluna++;
                    }
                    worksheet.Row(1).Height = 76;//nao sei pq teve q botar isso pra bater os 100 px da imagem

                    // lets set the header text 
                    worksheet.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\" Relatório YFAN";
                    // add the page number to the footer plus the total number of pages
                    worksheet.HeaderFooter.OddFooter.RightAlignedText =
                        string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                    // add the sheet name to the footer
                    worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                    // add the file path to the footer
                    worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;

                    worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:2"];
                    worksheet.PrinterSettings.RepeatColumns = worksheet.Cells["A:G"];
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape;

                    // Change the sheet view to show it in page layout mode
                    worksheet.View.PageLayoutView = true;

                    // set some document properties
                    package.Workbook.Properties.Title = "Votações";
                    package.Workbook.Properties.Author = "YFAN";
                    package.Workbook.Properties.Comments = "Relatório de votações";

                    // set some extended property values
                    package.Workbook.Properties.Company = "Voluta Tecnologia.";

                    // set some custom property values
                    package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jefferson Fidencio");
                    package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "YFAN");
                    // save our new workbook and we are done!
                    package.Save();

                }

                return newFile.FullName;
            }
            catch (Exception)
            {

                throw;
            }
            throw new NotImplementedException();
        }

        private static List<Bitmap> CarregarEstampas(DirectoryInfo outputDir)
        {
            DirectoryInfo estampasDir = new DirectoryInfo(outputDir + @"\Estampas");
            var files = estampasDir.GetFiles();
            List<Bitmap> estampas = new List<Bitmap>();
            foreach (var file in files)
            {
                var imagem = (Bitmap)Bitmap.FromFile(file.FullName);
                var resized = new Bitmap(imagem, new Size(imagem.Width / 10, imagem.Height / 10));
                estampas.Add(resized);
            }
            return estampas;
        }

        public static byte[] StreamToBytes(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

        private static Colecao DeserializeObject(string res)
        {
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(Colecao));
                using (TextReader reader = new StringReader(res))
                {
                    return ((Colecao)xmlSerializer.Deserialize(reader));
                }
            }
            catch (Exception)
            {
                return new Colecao();
            }
        }
    }
}

using System.Text.RegularExpressions;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace PDFExtractor;
class Program
{
    public static string CsvPath = "/Users/chenzejun/Downloads/PDF_results.csv";

    public static string ErrorPath = "/Users/chenzejun/Downloads/PDF_errors.csv";

    static void Main(string[] args)
    {
        if (File.Exists(CsvPath))
        {
            File.Delete(CsvPath);
        }
        if (File.Exists(ErrorPath))
        {
            File.Delete(ErrorPath);
        }
        string[] headers = { "海关编号", "出口日期", "合同协议号", "商品名称、规格型号", "数量及单位", "单价", "总价" };
        WriteToCsv(string.Join(",", headers));
        string[] pdfPaths = Directory.GetFiles("/Users/chenzejun/Downloads/PDF_test");
        foreach (string pdfPath in pdfPaths)
        {
            ExtractTextFromPdf(pdfPath);
        }
    }

    public static void ExtractTextFromPdf(string pdfPath)
    {
        try
        {
            using (PdfReader reader = new PdfReader(pdfPath))
            {
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {                  
                    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                    {
                        var strategy = new SimpleTextExtractionStrategy();
                        PdfCanvasProcessor parser = new PdfCanvasProcessor(strategy);
                        parser.ProcessPageContent(pdfDoc.GetPage(i));
                        string text = strategy.GetResultantText();
                        string[] words = text.Split(" ");
                        ExportToCsv(words, string.Empty);
                    }
                    WriteToCsv(string.Empty);
                }
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read the {pdfPath}: {ex}");
            string[] temp = pdfPath.Split('/');
            WriteToCsv(temp[temp.Length - 1], true, ErrorPath);
        }
    }

    public static void ExportToCsv(string[] words, string subHeader)
    {
        string hgbh = words[6].Split("：")[1].Trim();
        string ckrq = words[10].Split('\n')[0].Trim();
        int index = 0;     
        for (int i = 0; i < words.Length; i++)
        {
            if (words[i].Equals("/0/\n合同协议号\n"))
            {
                index = i + 1;
                break;
            }
        }
        string htxyh = words[index].Split('\n')[0].Trim();
        int start = 0;
        int minus = 0;
        for(int i = 0; i < words.Length; i++)
        {
            if (words[i].Equals("币制"))
            {
                start = i + 3;
                break;
            }
        }
        for (int i = 0; i < 4; i++)
        {
            if (!ValidateGoods(words[start - 2 + i * 12 - minus].Split('\n')[1]))
            {
                break;
            }
            string spmc = words[start + i * 12 - minus];
            string sl = words[start + 2 + i * 12 - minus];
            string djzj = words[start + 4 + i * 12 - minus];
            string dj;
            string zj;
            if (djzj.Contains('\n'))
            {
                minus++;
                string[] temp = djzj.Split('\n');
                dj = temp[0] + temp[1];
                zj = temp[2];
            } else
            {
                dj = words[start + 4 + i * 12 - minus];
                zj = words[start + 5 + i * 12 - minus];
            }
            string[] result = { hgbh, ckrq, htxyh, spmc, sl, dj, zj };
            WriteToCsv(string.Join(",", result));
        }       
    }

    private static void WriteToCsv(string content, bool append = true, string csvPath = "/Users/chenzejun/Downloads/PDF_results.csv")
    {
        using (StreamWriter writer = new StreamWriter(csvPath, append))
        {
            writer.WriteLine(content);
        }
    }

    private static bool ValidateGoods(string spxh)
    {
        string pattern = @"\b([1-9]|[1-9][0-9]|100)\b";
        Regex regex = new Regex(pattern);
        Match match = regex.Match(spxh);
        return match.Success;
    }
}


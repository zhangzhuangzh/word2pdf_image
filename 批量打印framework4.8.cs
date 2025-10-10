using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WordToPdfBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            // 获取目标文件夹
            string folderPath = args.Length > 0 ? args[0] : "";
            if (string.IsNullOrEmpty(folderPath))
            {
                Console.Write("请输入目标文件夹路径：");
                folderPath = Console.ReadLine();
            }

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("指定的文件夹不存在！");
                return;
            }

            // 创建临时目录
            string tempDir = Path.Combine(folderPath, "temp_pdf_pages");
            Directory.CreateDirectory(tempDir);

            List<string> tempFiles = new List<string>();
            Application wordApp = null;

            try
            {
                wordApp = new Application();
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                // 递归搜索所有文件并筛选 Word 文件
                string[] allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
                List<string> wordFiles = new List<string>();
                foreach (string file in allFiles)
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (ext == ".doc" || ext == ".docx")
                        wordFiles.Add(file);
                }

                int totalFiles = wordFiles.Count;
                if (totalFiles == 0)
                {
                    Console.WriteLine("未找到任何 Word 文件。");
                    return;
                }

                int processedCount = 0;
                foreach (string file in wordFiles)
                {
                    string pdfFile = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(file) + ".pdf");
                    Console.WriteLine($"[{++processedCount}/{totalFiles}] 转换 Word 文件: {Path.GetFileName(file)} ...");

                    Microsoft.Office.Interop.Word.Document doc = null;
                    try
                    {
                        doc = wordApp.Documents.Open(file, ReadOnly: true, Visible: false);
                        doc.ExportAsFixedFormat(pdfFile, WdExportFormat.wdExportFormatPDF);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Word 转 PDF 失败: {ex.Message}");
                        continue;
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            doc.Close(false);
                            Marshal.ReleaseComObject(doc);
                        }
                    }

                    // 提取前 2 页生成临时 PDF
                    string tempPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(file) + "_first2pages.pdf");
                    try
                    {
                        using (PdfReader reader = new PdfReader(pdfFile))
                        {
                            int pages = Math.Min(2, reader.NumberOfPages);
                            using (iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document())
                            {
                                using (PdfCopy copy = new PdfCopy(pdfDoc, new FileStream(tempPdf, FileMode.Create)))
                                {
                                    pdfDoc.Open();
                                    for (int i = 1; i <= pages; i++)
                                    {
                                        copy.AddPage(copy.GetImportedPage(reader, i));
                                    }
                                }
                            }
                        }
                        tempFiles.Add(tempPdf);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"提取前 2 页失败: {ex.Message}");
                    }
                }

                // 合并所有临时 PDF
                string finalPdf = Path.Combine(folderPath, "打印（内部）.pdf");
                try
                {
                    using (iTextSharp.text.Document finalDoc = new iTextSharp.text.Document())
                    {
                        using (PdfCopy copy = new PdfCopy(finalDoc, new FileStream(finalPdf, FileMode.Create)))
                        {
                            finalDoc.Open();
                            foreach (string temp in tempFiles)
                            {
                                using (PdfReader reader = new PdfReader(temp))
                                {
                                    for (int i = 1; i <= reader.NumberOfPages; i++)
                                    {
                                        copy.AddPage(copy.GetImportedPage(reader, i));
                                    }
                                }
                            }
                        }
                    }
                    Console.WriteLine($"最终 PDF 已生成: {finalPdf}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"合并 PDF 失败: {ex.Message}");
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    wordApp.Quit(false);
                    Marshal.ReleaseComObject(wordApp);
                }

                // 清理临时文件夹
                try
                {
                    Directory.Delete(tempDir, true);
                }
                catch { }
            }

            Console.WriteLine("处理完成，按任意键退出...");
            Console.ReadKey();
        }
    }
}

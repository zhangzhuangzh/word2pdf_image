using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
// 给 System.Windows.Forms 建立别名，避免与其他命名空间冲突
using WinForms = System.Windows.Forms;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;

class Program
{
    // 入口必须是 STA 以便使用 WinForms 对话框
    [STAThread]
    static int Main(string[] args)
    {
        // 必须在创建任何 WinForms 对象之前设置高 DPI 模式（避免 4K 下模糊）
        try
        {
            WinForms.Application.SetHighDpiMode(WinForms.HighDpiMode.PerMonitorV2);
        }
        catch
        {
            // 某些环境可能没有该 API，忽略以保持兼容
        }

        WinForms.Application.EnableVisualStyles();
        WinForms.Application.SetCompatibleTextRenderingDefault(false);

        string targetDir;

        // 如果命令行提供了有效目录则优先使用
        if (args.Length >= 1 && Directory.Exists(args[0]))
        {
            targetDir = args[0];
        }
        else
        {
            // 弹出目录选择对话框
            using (var fbd = new WinForms.FolderBrowserDialog())
            {
                fbd.Description = "请选择要处理的目录（程序将在该目录中查找 .doc/.docx 并将转换后的 PDF 的前2页合并）";
                fbd.ShowNewFolderButton = true;
                var dr = fbd.ShowDialog();
                if (dr != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    Console.WriteLine("未选择目录，程序退出。");
                    return 1;
                }
                targetDir = fbd.SelectedPath;
            }
        }

        if (!Directory.Exists(targetDir))
        {
            Console.WriteLine($"目录不存在: {targetDir}");
            return 1;
        }

        string outputPdf = Path.Combine(targetDir, "打印（内部）.pdf");
        string tempDir = Path.Combine(Path.GetTempPath(), "word2pdf_temp_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        Console.WriteLine($"开始处理目录: {targetDir}");
        Console.WriteLine($"临时目录: {tempDir}");

        var generatedPdfFiles = new List<string>(); // 仅保存由 doc/docx 转换而来的 PDF（临时文件）
        try
        {
            // 1) 遍历 Word 文件并转换为 PDF（生成到临时目录）
            var wordFiles = Directory.EnumerateFiles(targetDir, "*.*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                .ToList();

            Console.WriteLine($"找到 {wordFiles.Count} 个 Word 文件，开始转换（尝试 WPS/Word COM 自动化）...");
            foreach (var word in wordFiles)
            {
                try
                {
                    string pdfPath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(word) + "_" + Guid.NewGuid().ToString("N") + ".pdf");
                    bool ok = TryConvertWordToPdfUsingCom(word, pdfPath);
                    if (ok)
                    {
                        Console.WriteLine($"已转换: {word} -> {pdfPath}");
                        generatedPdfFiles.Add(pdfPath);
                    }
                    else
                    {
                        Console.WriteLine($"转换失败（跳过）: {word}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"转换时异常 (文件: {word}) : {ex.Message}");
                }
            }

            // 现在我们只处理 generatedPdfFiles（也就是仅合并由 doc/docx 转换得到的 PDF）
            // 排除任何与最终输出同路径的条目（以防冲突）
            var pdfsToProcess = generatedPdfFiles
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(p => !string.Equals(Path.GetFullPath(p), Path.GetFullPath(outputPdf), StringComparison.OrdinalIgnoreCase))
                .ToList();

            Console.WriteLine($"将处理 {pdfsToProcess.Count} 个由 Word 转换得到的 PDF（提取每个文件的前 2 页并合并）...");

            // 3) 对每个转换得到的 PDF 提取前 2 页，生成临时 pdf 文件
            var extractedTempPdfs = new List<string>();
            foreach (var pdf in pdfsToProcess)
            {
                try
                {
                    string extracted = Path.Combine(tempDir, "extracted_" + Path.GetFileNameWithoutExtension(pdf) + "_" + Guid.NewGuid().ToString("N") + ".pdf");
                    bool ok = TryExtractFirstNPages(pdf, extracted, 2);
                    if (ok)
                    {
                        extractedTempPdfs.Add(extracted);
                        Console.WriteLine($"已提取前2页: {pdf} -> {extracted}");
                    }
                    else
                    {
                        Console.WriteLine($"提取失败或文件无页数(跳过): {pdf}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"提取时异常 (文件: {pdf}) : {ex.Message}");
                }
            }

            if (extractedTempPdfs.Count == 0)
            {
                Console.WriteLine("没有任何可合并的页（来自转换的 PDF），程序结束。");
                return 0;
            }

            // 4) 合并所有提取的临时 pdf 为最终 PDF
            Console.WriteLine($"开始合并 {extractedTempPdfs.Count} 个临时 PDF -> {outputPdf}");
            MergePdfs(extractedTempPdfs, outputPdf);
            Console.WriteLine($"合并完成: {outputPdf}");

            // 5) 删除中间产生的 PDF（转换出的 PDF 和提取出的临时 PDF）
            Console.WriteLine("开始删除中间文件（仅删除由转换产生的临时文件）...");
            foreach (var f in generatedPdfFiles.Concat(extractedTempPdfs))
            {
                TryDeleteFile(f);
            }
            Console.WriteLine("中间文件删除完成。");
        }
        finally
        {
            TryDeleteDirectory(tempDir);
        }

        Console.WriteLine("全部处理完成。");
        return 0;
    }

    // 尝试删除文件，捕获异常继续
    static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
                Console.WriteLine($"已删除: {path}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"删除失败: {path} , {ex.Message}");
        }
    }

    static void TryDeleteDirectory(string dir)
    {
        try
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, true);
            }
        }
        catch
        {
            // 忽略
        }
    }

    // 使用 COM Automation（尝试多种 ProgID）来将 Word 转为 PDF
    // 返回 true 表示成功生成输出 pdf 文件
    static bool TryConvertWordToPdfUsingCom(string inputDocPath, string outputPdfPath)
    {
        // 尝试的 ProgID 列表：优先 WPS 常见 ProgID，再尝试 Word 的 ProgID
        var progIds = new[] { "Kwps.Application", "wps.Application", "WPS.Application", "Word.Application" };

        foreach (var progId in progIds)
        {
            Type comType = null;
            try
            {
                comType = Type.GetTypeFromProgID(progId);
            }
            catch { comType = null; }

            if (comType == null) continue;

            dynamic app = null;
            dynamic doc = null;
            try
            {
                app = Activator.CreateInstance(comType);
                // 一些 COM 实现需要可见设为 false
                try { app.Visible = false; } catch { }

                // 打开文档（只读）。不同实现位置参数可能差异，但通常可用。
                doc = app.Documents.Open(inputDocPath, ReadOnly: true);

                // 保存为 PDF，Word 的 FileFormat = 17 (wdFormatPDF)
                try
                {
                    doc.SaveAs2(outputPdfPath, 17);
                }
                catch
                {
                    doc.SaveAs(outputPdfPath, 17);
                }

                try { doc.Close(false); } catch { }
                try { app.Quit(); } catch { }

                ReleaseComObject(doc);
                ReleaseComObject(app);

                return File.Exists(outputPdfPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用 {progId} 转换失败: {ex.Message}");
                try { if (doc != null) doc.Close(false); } catch { }
                try { if (app != null) app.Quit(); } catch { }
                ReleaseComObject(doc);
                ReleaseComObject(app);
                // 尝试下一个 progId
                continue;
            }
        }

        Console.WriteLine("未找到可用的 WPS/Word COM 对象或转换失败，无法将 Word 转为 PDF。");
        return false;
    }

    static void ReleaseComObject(object o)
    {
        try
        {
            if (o == null) return;
            if (Marshal.IsComObject(o))
            {
                Marshal.FinalReleaseComObject(o);
            }
        }
        catch { /* ignore */ }
    }

    // 提取 PDF 的前 N 页到目标文件。返回是否成功（且至少有 1 页）。
    static bool TryExtractFirstNPages(string sourcePdf, string targetPdf, int n)
    {
        try
        {
            using (var stream = File.OpenRead(sourcePdf))
            {
                var input = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
                int pagesToTake = Math.Min(n, input.PageCount);
                if (pagesToTake <= 0) return false;

                using (var outDoc = new PdfDocument())
                {
                    for (int i = 0; i < pagesToTake; i++)
                    {
                        var page = input.Pages[i];
                        outDoc.AddPage(page);
                    }
                    using (var outStream = File.Create(targetPdf))
                    {
                        outDoc.Save(outStream);
                    }
                }
            }
            return File.Exists(targetPdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"提取 PDF 页失败 ({sourcePdf}): {ex.Message}");
            return false;
        }
    }

    // 合并多个 PDF（按列表顺序），将所有页面追加到目标文件
    static void MergePdfs(IEnumerable<string> pdfFiles, string outputPath)
    {
        using (var outDoc = new PdfDocument())
        {
            foreach (var f in pdfFiles)
            {
                try
                {
                    using (var fs = File.OpenRead(f))
                    {
                        var tmp = PdfReader.Open(fs, PdfDocumentOpenMode.Import);
                        for (int i = 0; i < tmp.PageCount; i++)
                        {
                            outDoc.AddPage(tmp.Pages[i]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"合并时跳过文件 {f} ，错误: {ex.Message}");
                }
            }

            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var outFs = File.Create(outputPath))
            {
                outDoc.Save(outFs);
            }
        }
    }
}

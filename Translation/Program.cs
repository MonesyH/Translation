using OfficeOpenXml;
using NPOI.XWPF.UserModel;
using Google.Cloud.Translation.V2;

namespace Translation;

public class Program
{
    public static async Task Main(string[] args)
    {
        await ConvertExcelLanguageAsync(new FileInfo("Chinese excel.xlsx"), LanguageCodes.English, "English excel.xlsx",CancellationToken.None);
        
        await ConvertExcelLanguageAsync(new FileInfo("English excel.xlsx"), LanguageCodes.ChineseTraditional, "Chinese excel copy.xlsx",CancellationToken.None);
        
        await ConvertWordLanguageAsync("Chinese word.docx", LanguageCodes.English, "English word.docx",CancellationToken.None).ConfigureAwait(false);
        
        await ConvertWordLanguageAsync("English word.docx", LanguageCodes.ChineseTraditional, "Chinese word copy.docx",CancellationToken.None).ConfigureAwait(false);
    }

    private static async Task ConvertWordLanguageAsync(
        string wordName, string targetLanguage, string outputFileName, CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(wordName, FileMode.Open, FileAccess.Read);

        var doc = new XWPFDocument(stream);

        foreach (var para in doc.Paragraphs)
        {
            var translatedText = await TranslateTextAsync(targetLanguage, para.Text, cancellationToken).ConfigureAwait(false);
            
            if(string.IsNullOrEmpty(translatedText)) continue;
            
            para.ReplaceText(para.Text, translatedText);
        }
 
        await using var outputStream = new FileStream(outputFileName, FileMode.Create);
        doc.Write(outputStream);
    }
    
    private static async Task ConvertExcelLanguageAsync(
        FileInfo excel, string targetLanguage, string outputFileName, CancellationToken cancellationToken)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        using var package = new ExcelPackage(excel);
        var worksheet = package.Workbook.Worksheets[0];
        
        for (var row = 1; row <= worksheet.Dimension.Rows; row++)
        {
            for (var col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                var cell = worksheet.Cells[row, col];
                
                if (cell.Text.Trim() != "")
                {
                    var translatedText = await TranslateTextAsync(targetLanguage, cell.Text, cancellationToken).ConfigureAwait(false);
                    cell.Value = translatedText;
                }
            }
        }

        await package.SaveAsAsync(new FileInfo(outputFileName), cancellationToken).ConfigureAwait(false);
    }
    
    private static async Task<string> TranslateTextAsync(string targetLanguage, string text, CancellationToken cancellationToken)
    {
        var translationClient = TranslationClient.CreateFromApiKey("xxxxx");

        var response = await translationClient
            .TranslateTextAsync(text, targetLanguage, cancellationToken: cancellationToken).ConfigureAwait(false);

        return response.TranslatedText;
    }
}
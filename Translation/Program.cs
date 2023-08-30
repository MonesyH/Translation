using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using Google.Cloud.Translation.V2;

namespace Translation;

public class Program
{
    public static async Task Main(string[] args)
    {
        await ConvertWordLanguageAsync("Chinese word.docx", LanguageCodes.English, "English word.docx",CancellationToken.None).ConfigureAwait(false);
        
        await ConvertWordLanguageAsync("English word.docx", LanguageCodes.ChineseTraditional, "Chinese word copy.docx",CancellationToken.None).ConfigureAwait(false);
        
        await ConvertExcelLanguageAsync("Chinese excel.xlsx", LanguageCodes.English, "English excel.xlsx", CancellationToken.None);
        
        await ConvertExcelLanguageAsync("English excel.xlsx", LanguageCodes.ChineseTraditional, "Chinese excel copy.xlsx", CancellationToken.None);
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
        string excelName, string targetLanguage, string outputFileName, CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(excelName, FileMode.Open, FileAccess.Read);

        var excel = new XSSFWorkbook(stream);
        
        var worksheet = excel.GetSheetAt(0);

        for (var row = 0; row <= worksheet.LastRowNum; row++)
        {
            var currentRow = worksheet.GetRow(row);
            if (currentRow == null) continue;

            for (var col = 0; col < currentRow.LastCellNum; col++)
            {
                var cell = currentRow.GetCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                if (!string.IsNullOrWhiteSpace(cell.ToString()))
                {
                    var translatedText = await TranslateTextAsync(targetLanguage, cell.StringCellValue, cancellationToken).ConfigureAwait(false);
                    cell.SetCellValue(translatedText);
                }
            }
        }
        
        await using var outputStream = new FileStream(outputFileName, FileMode.Create);
        excel.Write(outputStream);
    } 
    
    private static async Task<string> TranslateTextAsync(string targetLanguage, string text, CancellationToken cancellationToken)
    {
        var translationClient = TranslationClient.CreateFromApiKey("xxxx");

        var response = await translationClient
            .TranslateTextAsync(text, targetLanguage, cancellationToken: cancellationToken).ConfigureAwait(false);

        return response.TranslatedText;
    }
}
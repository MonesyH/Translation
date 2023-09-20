using DocumentFormat.OpenXml;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using Google.Cloud.Translation.V2;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace Translation;

public class Program
{
    public static async Task Main(string[] args)
    {
        await ConvertWordLanguageAsync("Chinese word.docx", LanguageCodes.English, "English word.docx", CancellationToken.None).ConfigureAwait(false);

        await ConvertWordLanguageAsync("English word.docx", LanguageCodes.ChineseTraditional, "Chinese word copy.docx", CancellationToken.None).ConfigureAwait(false);

        await ConvertExcelLanguageAsync("Chinese excel.xlsx", LanguageCodes.English, "English excel.xlsx", CancellationToken.None).ConfigureAwait(false);

        await ConvertExcelLanguageAsync("English excel.xlsx", LanguageCodes.ChineseTraditional, "Chinese excel copy.xlsx", CancellationToken.None).ConfigureAwait(false);

        await ConvertPowerPointLanguageAsync("Chinese Powerpoint.pptx", LanguageCodes.English, "English Powerpoint.pptx", CancellationToken.None).ConfigureAwait(false);

        await ConvertPowerPointLanguageAsync("English Powerpoint.pptx", LanguageCodes.ChineseTraditional, "Chinese Powerpoint copy.pptx", CancellationToken.None).ConfigureAwait(false);
    }

    private static async Task ConvertPowerPointLanguageAsync(string inputPath, string targetLanguage, string outputPath, CancellationToken cancellationToken)
    {
        using var presentationDocument = PresentationDocument.Open(inputPath, false);
        
        using var newDocument = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation);

        foreach (var part in presentationDocument.Parts)
        {
            newDocument.AddPart(part.OpenXmlPart, part.RelationshipId);
        }

        foreach (var slidePart in newDocument.PresentationPart.SlideParts)
        {
            foreach (var textElement in slidePart.Slide.Descendants<A.Text>())
            {
                var text = textElement.Text;

                if (text != null)
                {
                    var translatedText = await TranslateTextAsync(targetLanguage, text, cancellationToken)
                        .ConfigureAwait(false);

                    textElement.Text = translatedText;
                }
            }
        }
        newDocument.Save();
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
        var translationClient = TranslationClient.CreateFromApiKey("");

        var response = await translationClient
            .TranslateTextAsync(text, targetLanguage, cancellationToken: cancellationToken).ConfigureAwait(false);

        return response.TranslatedText;
    }
}
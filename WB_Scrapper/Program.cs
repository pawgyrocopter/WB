using System.Net.Http.Json;
using OfficeOpenXml;
using WB_Scrapper;

HttpClient httpClient = new();

var path = @"\Keys.txt"; //keys.txt
var xlFilePath = @"\completedTask.xlsx"; //xlsx file
    
    
IEnumerable<string> content = File.ReadLines(path);

using var package = new ExcelPackage();
{
    foreach (var name in content)
    {
        var items = await GetItems(CreateUrl(name));
        var worksheet = package.Workbook.Worksheets.Add(name);
       
        worksheet.Cells[1, 1].Value = "Tittle"; //A
        worksheet.Cells[1, 2].Value = "Brand"; //B
        worksheet.Cells[1, 3].Value = "Id"; //C
        worksheet.Cells[1, 4].Value = "Feedback"; //D
        worksheet.Cells[1, 5].Value = "Price"; //E
        int i = 2;
        foreach (var item in items.Data.Products)
        {
            worksheet.Cells[$"A{i}"].Value = item.Name;
            worksheet.Cells[$"B{i}"].Value = item.Brand;
            worksheet.Cells[$"C{i}"].Value = item.Id;
            worksheet.Cells[$"D{i}"].Value = item.Feedbacks;
            worksheet.Cells[$"E{i}"].Value = item.PriceU / 100;
            i++;
        }
    }
    var xlFile = new FileInfo(xlFilePath);
    package.SaveAs(xlFile);
}
async Task<JsonData> GetItems(string url)
{
    HttpResponseMessage response = await httpClient.GetAsync(url);
    return await response.Content.ReadFromJsonAsync<JsonData>();
}

string CreateUrl(string s1)
{
    return
        "https://search.wb.ru/exactmatch/ru/common/v4/search?appType=1&couponsGeo=12,3,18,15,21&curr=rub&dest=-1029256,-102269,-2162196,-1257786&emp=0&lang=ru&locale=ru&pricemarginCoeff=1.0&query=" +
        s1 +
        "&reg=0&regions=68,64,83,4,38,80,33,70,82,86,75,30,69,22,66,31,40,1,48,71&resultset=catalog&sort=popular&spp=0&suppressSpellcheck=false";
}

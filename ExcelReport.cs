//Install ClosedXML package﻿
public class ReportsController : Controller
{
	public async Task<IActionResult> BookExcelReport()
	{
		IEnumerable<Book> books = _unitOfWord.Books.GetAll()
			.Include(b => b.Author)
			.Include(b => b.Categories!)
			.ThenInclude(b => b.Category);

		//Start Report
		using var Workbook = new XLWorkbook();//عملت ورك بوك
		var sheet = Workbook.AddWorksheet("Books");// ضفت فيه شيت وإدتله إسم;

		/*//Add Picture
		sheet.AddPicture($"{_webHostEnvironment.WebRootPath}/Mecatronic/Img/Logo.png")
			.MoveTo(sheet.Cell("A1"))
			.Scale(.2);//size*/
		//Fill header
		string[] Header = { "Title", "Author", "Categories", "publisher", "publishing Date", "Hall", "Available for rental", "Status" };
		//Writ this with Extension method
		for (int i = 0; i < Header.Length; i++)
		{
			sheet.Cell(1, i + 1).SetValue(Header[i]);
		}
		//format header                        
		/*var header = sheet.Range("A1", "G1");//(1,1,1,8)(r,c,r,c)
		header.Style.Fill.BackgroundColor = XLColor.DarkGray;
		header.Style.Font.FontColor = XLColor.White;
		header.Style.Font.SetBold();*/

		//Fill body
		var row = 2;
		foreach (var item in books)
		{
			sheet.Cell(row, 1).SetValue(item.Title);
			sheet.Cell(row, 2).SetValue(item.Author!.Name);
			sheet.Cell(row, 3).SetValue(string.Join(", ", item.Categories!.Select(c => c.Category!.Name)));
			sheet.Cell(row, 4).SetValue(item.Publisher);
			sheet.Cell(row, 5).SetValue(item.PublishingDate.ToString("dd MMM yyyy"));
			sheet.Cell(row, 6).SetValue(item.Hall);
			sheet.Cell(row, 7).SetValue(item.IsAvailableForRental ? "Yes" : "NO");
			sheet.Cell(row, 8).SetValue(item.IsDeleted ? "Available" : "Deleted");
			row++;
		}
		//format body
		//Writ this with Extension method
		/*sheet.ColumnsUsed().AdjustToContents();//علي كد اطول كلمه
		sheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;//سنتر الكلام
		sheet.CellsUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
		sheet.CellsUsed().Style.Border.OutsideBorderColor = XLColor.Black;*/

		//Add Table Style (Writ this with Extension method)
		var rang = sheet.Range(1, 1, books.Count() + 1, 8);
		var table = rang.CreateTable();
		table.Theme = XLTableTheme.TableStyleMedium13;
		table.ShowAutoFilter = false;

		await using var stream = new MemoryStream();
		Workbook.SaveAs(stream);
			                            //لازم الإمتداد    //النوع ده شغال اكسل وبي دي اف
		return File(stream.ToArray(), MediaTypeNames.Application.Octet, "Book.xlsx");
		//End Report
	}
		
}


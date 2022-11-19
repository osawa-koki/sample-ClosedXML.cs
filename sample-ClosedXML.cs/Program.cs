using System.IO;
using ClosedXML.Excel;

// 新規のExcelブックを作成する。
XLWorkbook book = new();

// ワークシートを追加する。
IXLWorksheet sheet = book.Worksheets.Add("Sheet1");

// セルに値を設定する。
sheet.Cell("A1").Value = "Hello, World!";
sheet.Cell("A2").Value = "こんにちは、世界！";

// 行の高さを変更する。
sheet.Row(1).Height = 30;
sheet.Row(2).Height = 50;
sheet.Row(3).Height = 70;

// 列の幅を変更する。
sheet.Column(1).Width = 30;
sheet.Column(2).Width = 50;
sheet.Column(3).Width = 70;

// フォント色を変更する。
sheet.Cell("A1").Style.Font.FontColor = XLColor.Red;
sheet.Cell("A2").Style.Font.FontColor = XLColor.FromArgb(0, 0, 255, 1);

// フォントサイズを変更する。
sheet.Cell("A1").Style.Font.FontSize = 20;
sheet.Cell("A2").Style.Font.FontSize = 40;

// フォントを太字にする。
sheet.Cell("A1").Style.Font.Bold = true;

// フォントを斜体にする。
sheet.Cell("A2").Style.Font.Italic = true;

// フォントを下線にする。
sheet.Cell("A1").Style.Font.Underline = XLFontUnderlineValues.Single;

// フォントを取り消し線にする。
sheet.Cell("A2").Style.Font.Strikethrough = true;

// セルの背景色を設定する。
sheet.Cell("A1").Style.Fill.BackgroundColor = XLColor.Yellow;
sheet.Cell("A2").Style.Fill.BackgroundColor = XLColor.FromArgb(0, 255, 0, 1);

// Excelブックを保存保存する。
book.SaveAs(@"sample.xlsx");

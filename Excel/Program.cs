using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.PixelFormats;
using System.Linq;

namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string sFile = @"C:\Temp\workBook.xlsx";
            string imageFileName = @"C:\Temp\test1.jpg";

            try
            {
                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(sFile, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "mySheet"
                };
                sheets.Append(sheet);

                var worksheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Add headers to the worksheet (1st row)
                var headerRow = new Row();
                headerRow.Append(
                    new Cell() { CellValue = new CellValue("Name"), DataType = CellValues.String },
                    new Cell() { CellValue = new CellValue("QrCodeValue"), DataType = CellValues.String }
                );
                worksheetData.AppendChild(headerRow);

                // Add data to the worksheet (2nd row)
                var dataRow = new Row();
                dataRow.Append(
                    new Cell() { CellValue = new CellValue("ARUN"), DataType = CellValues.String }, // Replace "ARUNe" with your actual data
                    new Cell() { CellValue = new CellValue(" "), DataType = CellValues.String } // Leave a space or use an empty string to allow space for the image
                );
                worksheetData.AppendChild(dataRow);

                var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
                {
                    worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                }

                if (drawingsPart.WorksheetDrawing == null)
                {
                    drawingsPart.WorksheetDrawing = new WorksheetDrawing();
                }

                var worksheetDrawing = drawingsPart.WorksheetDrawing;

                var imagePart = drawingsPart.AddImagePart(ImagePartType.Jpeg);

                using (var stream = new FileStream(imageFileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                using (var image = Image.Load(imageFileName))
                {
                    var extentsCx = image.Width * 914400 / image.Width;
                    var extentsCy = image.Height * 914400 / image.Height;

                    var colOffset = 0;
                    var rowOffset = 0;
                    int colNumber = 2;
                    int rowNumber = 7; // Start from the third row (after the header and data)

                    var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
                    var nvpId = nvps.Count() > 0 ?
                        (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1 :
                        1U;

                    var twoCellAnchor = new Xdr.TwoCellAnchor(
                        new Xdr.FromMarker
                        {
                            ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                            RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                            ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                            RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                        },
                        new Xdr.ToMarker
                        {
                            ColumnId = new Xdr.ColumnId(colNumber.ToString()),
                            RowId = new Xdr.RowId(rowNumber.ToString()),
                            ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                            RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                        },
                        new Xdr.Picture(
                            new Xdr.NonVisualPictureProperties(
                                new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imageFileName },
                                new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                            ),
                            new Xdr.BlipFill(
                                new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print }
                                //new A.Stretch(new A.FillRectangle())
                            ),
                            new Xdr.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = extentsCx, Cy = extentsCy }
                                ),
                                new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        ),
                        new Xdr.ClientData()
                    );

                    worksheetDrawing.Append(twoCellAnchor);
                }

                workbookpart.Workbook.Save();
                spreadsheetDocument.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

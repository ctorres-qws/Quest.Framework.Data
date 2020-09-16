using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iText.Barcodes;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Font;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using QRCoder;
using Quest.ScanToPrint.Data.Entities;
using iText.IO.Font;
using iText.Kernel.Colors;
using Quest.ScanToPrint.Business;

namespace Quest.ScantToPrint.Presentation
{
    //public class NoBorder : Border
    //{
    //    public NoBorder() : base(0)
    //    {
    //        this.SetWidth(0);
    //    }
    //    public override 
    //}
    public class PDFcreator
    {
        public void CreatePDF6x4(BarcodeReading barcodeReading, string color)
        {
            PdfWriter writer = new PdfWriter(string.Format(@"..\..\Labels\{0}.pdf", barcodeReading.Barcode));
            PdfDocument pdf = new PdfDocument(writer);
            pdf.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(new iText.Kernel.Geom.Rectangle(576, 384)));
            Document document = new Document(pdf);
            document.SetMargins(0, 0, 0, 0);

            QRCodeGenerator qr = new QRCodeGenerator();
            QRCodeData data = qr.CreateQrCode(barcodeReading.Barcode, QRCodeGenerator.ECCLevel.Q);
            QRCode code = new QRCode(data);

            Image logo = new Image(ImageDataFactory
            .Create(@"..\..\Images\full-logo.png"))
            .SetTextAlignment(TextAlignment.CENTER);

            Image imgQR = new Image(ImageDataFactory
            .Create((System.Drawing.Image)code.GetGraphic(3, System.Drawing.Color.Black, System.Drawing.Color.White, false), null))
            .SetTextAlignment(TextAlignment.CENTER).SetWidth(76);
            UnitValue[] columnWidths = new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 14 ),
                new UnitValue(UnitValue.PERCENT, 18 ),
                new UnitValue(UnitValue.PERCENT, 18 ),
                new UnitValue(UnitValue.PERCENT, 18 ),
                new UnitValue(UnitValue.PERCENT, 18 ),
                new UnitValue(UnitValue.PERCENT, 14 ),
            };


            Table table = new Table(columnWidths);
            table.SetFixedLayout();
            table.SetWidth(new UnitValue(UnitValue.PERCENT, 100));

            table.SetBorder(Border.NO_BORDER);
            Cell cellQRTL = new Cell(2, 1).SetBorder(Border.NO_BORDER);
            cellQRTL.Add(imgQR);

            Cell cellLogo = new Cell(1, 4).SetBorder(Border.NO_BORDER);
            cellLogo.Add(logo.SetHeight(50).SetMarginLeft(100).SetMarginTop(10).SetMarginBottom(10));

            Cell cellQRTR = new Cell(2, 1).SetBorder(Border.NO_BORDER);
            cellQRTR.Add(imgQR);

            Table centralContent = new Table(new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 23 ),
                new UnitValue(UnitValue.PERCENT, 23 ),
                new UnitValue(UnitValue.PERCENT, 6 ),
                new UnitValue(UnitValue.PERCENT, 24 ),
                new UnitValue(UnitValue.PERCENT, 19 ),
                new UnitValue(UnitValue.PERCENT, 5 ) });
            centralContent.SetWidth(new UnitValue(UnitValue.PERCENT, 100));
            centralContent.SetBorder(Border.NO_BORDER);

            //centralContent.AddCell(new Cell(3, 1).SetBorder(Border.NO_BORDER).SetMargin(0).SetPadding(0));
            Paragraph pJob = new Paragraph("JOB\n").SetTextAlignment(TextAlignment.LEFT);
            pJob.Add(new Paragraph(barcodeReading.Job)
                .SetFontSize(36)
                .SetStrokeWidth(1f)
                .SetStrokeColor(DeviceGray.BLACK)
                .SetBorder(new SolidBorder(1))
                .SetPaddingLeft(15)
                .SetPaddingRight(15)
                .SetPaddingTop(5)
                .SetPaddingBottom(5)
                .SetWidth(new UnitValue(UnitValue.PERCENT, 80))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBold()
                );

            Cell jobCell = new Cell(3, 2).SetBorder(Border.NO_BORDER);
            jobCell.Add(pJob);
            centralContent.AddCell(jobCell);
            centralContent.AddCell(new Cell(3, 1).SetBorder(Border.NO_BORDER));
            Paragraph pFloor = new Paragraph("FLOOR\n").SetBorder(Border.NO_BORDER).SetMarginLeft(0).SetPaddingLeft(0);
            pFloor.Add(new Paragraph(barcodeReading.Floor)
                .SetFontSize(36)
                .SetStrokeWidth(.9f)
                .SetStrokeColor(DeviceGray.BLACK)
                .SetBorder(new SolidBorder(1))
                .SetPaddingLeft(15)
                .SetPaddingRight(15)
                .SetPaddingTop(5)
                .SetPaddingBottom(5)
                .SetWidth(new UnitValue(UnitValue.PERCENT, 100))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBold()
                );

            Cell floorCell = new Cell(3, 4).SetBorder(Border.NO_BORDER);
            floorCell.Add(pFloor);
            centralContent.AddCell(floorCell);

            Paragraph pTag = new Paragraph("TAG\n").SetBorder(Border.NO_BORDER);
            pTag.Add(new Paragraph(barcodeReading.Tag)
                .SetFontSize(45)
                .SetStrokeWidth(.9f)
                .SetStrokeColor(DeviceGray.BLACK)
                .SetBorder(new SolidBorder(1))
                .SetPaddingLeft(15)
                .SetPaddingRight(15)
                .SetPaddingTop(5)
                .SetPaddingBottom(5)
                .SetWidth(new UnitValue(UnitValue.PERCENT, 100))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBold()
                );

            Cell tagCell = new Cell(3, 6).SetBorder(Border.NO_BORDER);
            tagCell.Add(pTag);
            centralContent.AddCell(tagCell);



            Cell cellCentralContent = new Cell(2, 4).SetBorder(Border.NO_BORDER);
            cellCentralContent.Add(centralContent);
            table.AddCell(cellQRTL);
            table.AddCell(cellLogo);
            table.AddCell(cellQRTR);
            //table.AddCell(new Cell(2, 1).SetBorder(Border.NO_BORDER).SetBackgroundColor(GetRgb(color)));
            table.AddCell(cellCentralContent);

            Table leftColor = new Table(new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 60 ),
                new UnitValue(UnitValue.PERCENT, 40 )});
            leftColor.SetWidth(new UnitValue(UnitValue.PERCENT, 100));
            leftColor.SetBorder(Border.NO_BORDER);
            leftColor.SetMarginTop(25);
            leftColor.SetMarginBottom(25);
            leftColor.SetHeight(new UnitValue(UnitValue.POINT, 170));

            Paragraph pleftColor = new Paragraph();
            leftColor.AddCell(new Cell(5, 1).Add(pleftColor).SetBorder(Border.NO_BORDER).SetBackgroundColor(GetRgb(color)));

            Cell cellLeftColor = new Cell(1, 1).SetBorder(Border.NO_BORDER);

            cellLeftColor.Add(leftColor);



            table.AddCell(cellLeftColor);


            Table rightColor = new Table(new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 45 ),
                new UnitValue(UnitValue.PERCENT, 15 )});
            rightColor.SetWidth(new UnitValue(UnitValue.POINT, 60));
            rightColor.SetBorder(Border.NO_BORDER);
            rightColor.SetMarginTop(25);
            rightColor.SetMarginLeft(31);
            rightColor.SetMarginBottom(25);
            rightColor.SetHeight(new UnitValue(UnitValue.POINT, 170));

            Paragraph prightColor = new Paragraph();
            rightColor.AddCell(new Cell(5, 1).Add(prightColor).SetBorder(Border.NO_BORDER).SetBackgroundColor(GetRgb(color)));

            Cell cellRightColor = new Cell(1, 1).SetBorder(Border.NO_BORDER);

            cellRightColor.Add(rightColor);


            table.AddCell(cellRightColor);


            Cell cellQRTL2 = new Cell(2, 1).SetBorder(Border.NO_BORDER);
            cellQRTL2.Add(imgQR);

            Table tableBottom = new Table(new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 5),
                new UnitValue(UnitValue.PERCENT, 19),
                new UnitValue(UnitValue.PERCENT, 19),
                new UnitValue(UnitValue.PERCENT, 19),
                new UnitValue(UnitValue.PERCENT, 19),
                new UnitValue(UnitValue.PERCENT, 19) });

            tableBottom.SetWidth(new UnitValue(UnitValue.PERCENT, 90)).SetTextAlignment(TextAlignment.CENTER);
            tableBottom.SetMarginLeft(30);
            tableBottom.SetBorder(Border.NO_BORDER);


            Paragraph pDate = new Paragraph("DATE:").SetTextAlignment(TextAlignment.LEFT).SetFontSize(14);
            Paragraph pDateTimeInfo = new Paragraph(barcodeReading.ScanDate.ToString("MMM-dd-yyyy hh:mm tt").ToUpper()).SetTextAlignment(TextAlignment.LEFT).SetFontSize(14);
            Paragraph pLine = new Paragraph(string.Format("L{0}", barcodeReading.Line.ToString())).SetTextAlignment(TextAlignment.LEFT).SetFontSize(18);


            tableBottom.AddCell(new Cell(1, 2).Add(pDate).SetBorder(Border.NO_BORDER));
            tableBottom.AddCell(new Cell(1, 3).Add(pDateTimeInfo).SetBorder(Border.NO_BORDER));
            tableBottom.AddCell(new Cell(1, 1).Add(pLine).SetBorder(Border.NO_BORDER));
            

            tableBottom.AddCell(new Cell(1, 6).Add(new Paragraph()).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER).SetMarginLeft(10).SetWidth(new UnitValue(UnitValue.PERCENT, 100)).SetBackgroundColor(GetRgb(color)).SetHeight(new UnitValue(UnitValue.POINT, 30)));
            Cell cellQRTR2 = new Cell(2, 1).SetBorder(Border.NO_BORDER);
            cellQRTR2.Add(imgQR);

            table.AddCell(cellQRTL2);
            table.AddCell(new Cell(1, 4).Add(tableBottom).SetBorder(Border.NO_BORDER));
            table.AddCell(cellQRTR2);

            document.Add(table);
            
            document.Add(table);
            document.Close();

            

        }
        public void CreatePDF8x6(BarcodeReading barcodeReading, string color)
        {
            PdfWriter writer = new PdfWriter(string.Format(@"..\..\Labels\{0}.pdf", barcodeReading.Barcode));
            PdfDocument pdf = new PdfDocument(writer);
            pdf.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(new iText.Kernel.Geom.Rectangle(768, 384)));
            Document document = new Document(pdf);
            document.SetMargins(0, 0, 0, 0);

            QRCodeGenerator qr = new QRCodeGenerator();
            QRCodeData data = qr.CreateQrCode(barcodeReading.Barcode, QRCodeGenerator.ECCLevel.Q);
            QRCode code = new QRCode(data);

            Image logo = new Image(ImageDataFactory
            .Create(@"..\..\Images\full-logo.png"))
            .SetTextAlignment(TextAlignment.CENTER);

            Image imgQR = new Image(ImageDataFactory
            .Create((System.Drawing.Image)code.GetGraphic(5, System.Drawing.Color.Black, System.Drawing.Color.White, false), null))
            .SetTextAlignment(TextAlignment.CENTER).SetWidth(new UnitValue(UnitValue.PERCENT, 83));
            UnitValue[] columnWidths = new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 14 ),
                new UnitValue(UnitValue.PERCENT, 18 ),
                new UnitValue(UnitValue.PERCENT, float.Parse((20.4).ToString())),
                new UnitValue(UnitValue.PERCENT, 18 ),
                new UnitValue(UnitValue.PERCENT, 18 ),
                new UnitValue(UnitValue.PERCENT, 14 ),
            };


            Table table = new Table(columnWidths);
            table.SetFixedLayout();
            table.SetWidth(new UnitValue(UnitValue.PERCENT, 100));

            table.SetBorder(Border.NO_BORDER);
            Cell cellQRTL = new Cell(2, 1).SetBorder(Border.NO_BORDER);
            cellQRTL.Add(imgQR);

            Cell cellLogo = new Cell(1, 4).SetBorder(Border.NO_BORDER);
            cellLogo.Add(logo.SetHeight(50).SetMarginLeft(175).SetMarginTop(10).SetMarginBottom(10));

            Cell cellQRTR = new Cell(2, 1).SetBorder(Border.NO_BORDER);
            cellQRTR.Add(imgQR);

            Table centralContent = new Table(new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 23 ),
                new UnitValue(UnitValue.PERCENT, 23 ),
                new UnitValue(UnitValue.PERCENT, 6 ),
                new UnitValue(UnitValue.PERCENT, 24 ),
                new UnitValue(UnitValue.PERCENT, 19 ),
                new UnitValue(UnitValue.PERCENT, 5 ) });
            centralContent.SetWidth(new UnitValue(UnitValue.PERCENT, 100));
            centralContent.SetBorder(Border.NO_BORDER).SetAutoLayout();

            //centralContent.AddCell(new Cell(3, 1).SetBorder(Border.NO_BORDER).SetMargin(0).SetPadding(0));
            Paragraph pJob = new Paragraph("JOB\n").SetTextAlignment(TextAlignment.LEFT);
            pJob.Add(new Paragraph(barcodeReading.Job)
                .SetFontSize(36)
                .SetStrokeWidth(1f)
                .SetStrokeColor(DeviceGray.BLACK)
                .SetBorder(new SolidBorder(1))
                .SetPaddingLeft(15)
                .SetPaddingRight(15)
                .SetPaddingTop(5)
                .SetPaddingBottom(5)
                .SetWidth(new UnitValue(UnitValue.PERCENT, 80))
                .SetTextAlignment(TextAlignment.CENTER)
                );

            Cell jobCell = new Cell(3, 2).SetBorder(Border.NO_BORDER);
            jobCell.Add(pJob);
            centralContent.AddCell(jobCell);
            centralContent.AddCell(new Cell(3, 1).SetBorder(Border.NO_BORDER));
            Paragraph pFloor = new Paragraph("FLOOR\n").SetBorder(Border.NO_BORDER).SetMarginLeft(0).SetPaddingLeft(0);
            pFloor.Add(new Paragraph(barcodeReading.Floor)
                .SetFontSize(36)
                .SetStrokeWidth(.9f)
                .SetStrokeColor(DeviceGray.BLACK)
                .SetBorder(new SolidBorder(1))
                .SetPaddingLeft(15)
                .SetPaddingRight(15)
                .SetPaddingTop(5)
                .SetPaddingBottom(5)
                .SetWidth(new UnitValue(UnitValue.PERCENT, 100))
                .SetTextAlignment(TextAlignment.CENTER)
                );

            Cell floorCell = new Cell(3, 4).SetBorder(Border.NO_BORDER);
            floorCell.Add(pFloor);
            centralContent.AddCell(floorCell);

            Paragraph pTag = new Paragraph("TAG\n").SetBorder(Border.NO_BORDER);
            pTag.Add(new Paragraph(barcodeReading.Tag)
                .SetFontSize(45)
                .SetStrokeWidth(.9f)
                .SetStrokeColor(DeviceGray.BLACK)
                .SetBorder(new SolidBorder(1))
                .SetPaddingLeft(15)
                .SetPaddingRight(15)
                .SetPaddingTop(5)
                .SetPaddingBottom(5)
                .SetWidth(new UnitValue(UnitValue.PERCENT, 100))
                .SetTextAlignment(TextAlignment.CENTER)
                );

            Cell tagCell = new Cell(3, 6).SetBorder(Border.NO_BORDER);
            tagCell.Add(pTag);
            centralContent.AddCell(tagCell);



            Cell cellCentralContent = new Cell(2, 4).SetBorder(Border.NO_BORDER);
            cellCentralContent.Add(centralContent);
            table.AddCell(cellQRTL);
            table.AddCell(cellLogo);
            table.AddCell(cellQRTR);
            //table.AddCell(new Cell(2, 1).SetBorder(Border.NO_BORDER).SetBackgroundColor(GetRgb(color)));
            table.AddCell(cellCentralContent);

            Table leftColor = new Table(new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 60 ),
                new UnitValue(UnitValue.PERCENT, 40 )});
            leftColor.SetWidth(new UnitValue(UnitValue.PERCENT, 100));
            leftColor.SetBorder(Border.NO_BORDER);
            leftColor.SetMarginTop(25);
            leftColor.SetMarginBottom(25);
            leftColor.SetHeight(new UnitValue(UnitValue.POINT, 150));

            Paragraph pleftColor = new Paragraph();
            leftColor.AddCell(new Cell(5, 1).Add(pleftColor).SetBorder(Border.NO_BORDER).SetBackgroundColor(GetRgb(color)));

            Cell cellLeftColor = new Cell(1, 1).SetBorder(Border.NO_BORDER);

            cellLeftColor.Add(leftColor);



            table.AddCell(cellLeftColor);


            Table rightColor = new Table(new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 40 ),
                new UnitValue(UnitValue.PERCENT, 60 )});
            rightColor.SetWidth(new UnitValue(UnitValue.POINT, 106));
            rightColor.SetBorder(Border.NO_BORDER);
            rightColor.SetMarginTop(25);
            rightColor.SetMarginLeft(36);
            rightColor.SetMarginBottom(25);
            rightColor.SetHeight(new UnitValue(UnitValue.POINT, 150));

            Paragraph prightColor = new Paragraph();
            rightColor.AddCell(new Cell(5, 1).Add(prightColor).SetBorder(Border.NO_BORDER).SetBackgroundColor(GetRgb(color)));

            Cell cellRightColor = new Cell(1, 1).SetBorder(Border.NO_BORDER);

            cellRightColor.Add(rightColor);


            table.AddCell(cellRightColor);


            Cell cellQRTL2 = new Cell(2, 1).SetBorder(Border.NO_BORDER);
            cellQRTL2.Add(imgQR);

            Table tableBottom = new Table(new UnitValue[] {
                new UnitValue(UnitValue.PERCENT, 5),
                new UnitValue(UnitValue.PERCENT, 19),
                new UnitValue(UnitValue.PERCENT, 19),
                new UnitValue(UnitValue.PERCENT, 19),
                new UnitValue(UnitValue.PERCENT, 19),
                new UnitValue(UnitValue.PERCENT, 19) });

            tableBottom.SetWidth(new UnitValue(UnitValue.PERCENT, 90)).SetTextAlignment(TextAlignment.CENTER);
            tableBottom.SetMarginLeft(30);
            tableBottom.SetBorder(Border.NO_BORDER);


            Paragraph pDate = new Paragraph("DATE:").SetTextAlignment(TextAlignment.LEFT).SetFontSize(14);
            Paragraph pDateTimeInfo = new Paragraph(barcodeReading.ScanDate.ToString("MMM-dd-yyyy hh:mm tt").ToUpper()).SetTextAlignment(TextAlignment.LEFT).SetFontSize(14);
            Paragraph pLine = new Paragraph(string.Format("L{0}", barcodeReading.Line.ToString())).SetTextAlignment(TextAlignment.LEFT).SetFontSize(18);


            tableBottom.AddCell(new Cell(1, 2).Add(pDate).SetBorder(Border.NO_BORDER));
            tableBottom.AddCell(new Cell(1, 3).Add(pDateTimeInfo).SetBorder(Border.NO_BORDER));
            tableBottom.AddCell(new Cell(1, 1).Add(pLine).SetBorder(Border.NO_BORDER));


            tableBottom.AddCell(new Cell(1, 6).Add(new Paragraph()).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER).SetMarginLeft(10).SetWidth(new UnitValue(UnitValue.PERCENT, 100)).SetBackgroundColor(GetRgb(color)).SetHeight(new UnitValue(UnitValue.POINT, 35)));
            Cell cellQRTR2 = new Cell(2, 1).SetBorder(Border.NO_BORDER);
            cellQRTR2.Add(imgQR);

            table.AddCell(cellQRTL2);
            table.AddCell(new Cell(1, 4).Add(tableBottom).SetBorder(Border.NO_BORDER));
            table.AddCell(cellQRTR2);


            document.Add(table);

            document.Close();


        }
        DeviceRgb GetRgb(string hex)
        {
            string r = "00", g = "00", b = "00";
            if (hex.Length >= 7)
            {
                r = hex.Substring(1, 2);
                g = hex.Substring(3, 2);
                b = hex.Substring(5, 2);
            }
            return new DeviceRgb(int.Parse(r, System.Globalization.NumberStyles.HexNumber), int.Parse(g, System.Globalization.NumberStyles.HexNumber), int.Parse(b, System.Globalization.NumberStyles.HexNumber));
        }
    }
}

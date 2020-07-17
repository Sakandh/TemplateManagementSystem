using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Spire.Barcode;
using Spire.Pdf;
using System.Drawing;
using System.IO;
using Spire.Pdf.Graphics;
using ZXing;

namespace TMS_Framework.Models
{
    public class BarCode
    {
        //Write bar code, produce an image for the bar code
        public void WriteBarCode(string barCodeNumber, string fileName)
        {
            BarcodeSettings barCodeSettings = new BarcodeSettings();

            // set the x dimension
            barCodeSettings.Unit = GraphicsUnit.Point;
            barCodeSettings.X = 1f;

            // set the data
            barCodeSettings.Data = barCodeNumber;
            barCodeSettings.Data2D = barCodeNumber;


            // generate barcode
            barCodeSettings.Type = BarCodeType.Code39Extended;
            BarCodeGenerator bargenerator = new BarCodeGenerator(barCodeSettings);
            bargenerator.GenerateImage().Save(fileName);
        }

        //Read bar code form a full pafh for pdf
        public string ReadBarCode(string filePath)
        {

            if (filePath == null)
                return "";

            PdfDocument pdfDocument = new PdfDocument();

            if (File.Exists(filePath))
            {
                // opens the pdf and returns the barcode value
                pdfDocument.LoadFromFile(filePath);
                Bitmap image = (Bitmap)pdfDocument.SaveAsImage(0);
                string[] barcodeData = BarcodeScanner.Scan(image, BarCodeType.Code39Extended);

                // if there's no barcode then it returns the no barcode message
                if (barcodeData.Length < 1)
                {
                    return "";
                }
                else
                {
                    return barcodeData[0];
                }
            }
            else
            {
                throw new FileNotFoundException("File does not exist: " + filePath);
            }

        }

        public string ReadQRCode(string filePath)
        {
            // save the pdf in a larger size (2245px x 3179px)
            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(filePath);
            PdfDocument newPDF = new PdfDocument();
            foreach (PdfPageBase page in doc.Pages)
            {
                PdfPageBase newPage = newPDF.Pages.Add(PdfPageSize.A1, new PdfMargins(0));
                PdfTextLayout loLayout = new PdfTextLayout();
                loLayout.Layout = PdfLayoutType.OnePage;
                page.CreateTemplate().Draw(newPage, new PointF(0, 0), loLayout);
            }

            // temporary pdf file path
            string fileNameOnly = Path.GetFileNameWithoutExtension(filePath);
            string extension = ".pdf";

            string tempPDF = Path.GetTempPath() + fileNameOnly + extension;
            newPDF.SaveToFile(tempPDF);
            doc.LoadFromFile(tempPDF);

            // converting pdf to image
            Bitmap image = (Bitmap)doc.SaveAsImage(0);

            // reading qr code from the new larger pdf
            var qrcodeReader = new BarcodeReader();
            var qrcodeResult = qrcodeReader.Decode(image);

            if (qrcodeResult == null)
            {
                return "";
            }
            else
            {
                return qrcodeResult.Text;
            }
        }
    }
}
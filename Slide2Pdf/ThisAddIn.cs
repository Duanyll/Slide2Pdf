using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;

namespace Slide2Pdf
{
    public struct Rect
    {
        public double Top { get; set; }
        public double Left { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Bottom { get { return Top + Height; } }
        public double Right { get { return Left + Width; } }
        public Rect(double top, double left, double width, double height)
        {
            Top = top;
            Left = left;
            Width = width;
            Height = height;
        }
    }

    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void ExportCurrentSlideAsPdf(string outPath)
        {
            var presentation = Application.ActivePresentation;
            presentation.ExportAsFixedFormat(
                outPath,
                PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                Intent: PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                RangeType: PowerPoint.PpPrintRangeType.ppPrintCurrent
            );
        }

        public bool GetCurrentSlideContentBoundingRect(out Rect rect)
        {
            var slide = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Visible == Office.MsoTriState.msoTrue)
                {
                    minX = Math.Min(minX, shape.Left);
                    minY = Math.Min(minY, shape.Top);
                    maxX = Math.Max(maxX, shape.Left + shape.Width);
                    maxY = Math.Max(maxY, shape.Top + shape.Height);
                }
            }
            if (minX == double.MaxValue || minY == double.MaxValue || maxX == double.MinValue || maxY == double.MinValue)
            {
                rect = new Rect(0, 0, 0, 0);
                return false;
            }
            // Return the bounding rectangle in relative coordinates
            double slideWidth = slide.Master.Width;
            double slideHeight = slide.Master.Height;
            rect = new Rect(
                minY / slideHeight,
                minX / slideWidth,
                (maxX - minX) / slideWidth,
                (maxY - minY) / slideHeight
            );
            return true;
        }

        public void CropPdf(string pdfPath, Rect rect)
        {
            using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify))
            {
                foreach (var page in document.Pages)
                {
                    double pdfWidth = page.Width.Point;
                    double pdfHeight = page.Height.Point;
                    var pdfRect = new PdfRectangle(
                        new XPoint(rect.Left * pdfWidth, (1 - rect.Bottom) * pdfHeight),
                        new XPoint(rect.Right * pdfWidth, (1 - rect.Top) * pdfHeight)
                    );
                    page.TrimBox = pdfRect;
                    page.CropBox = page.TrimBox;
                }
                string tempPath = System.IO.Path.GetTempFileName();
                document.Save(tempPath);
                document.Close();
                System.IO.File.Delete(pdfPath);
                System.IO.File.Move(tempPath, pdfPath);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

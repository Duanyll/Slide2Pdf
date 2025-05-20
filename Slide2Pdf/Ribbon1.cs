using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Slide2Pdf
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private bool SaveCurrentSlide(out string outPath)
        {
            var addIn = Globals.ThisAddIn;
            var presentation = addIn.Application.ActivePresentation;

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Title = "Save PDF File",
                DefaultExt = "pdf",
            };

            if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                string pptPath = presentation.Path;
                string pptName = Path.GetFileNameWithoutExtension(presentation.Name);
                if (!pptPath.StartsWith("http"))
                {
                    saveFileDialog.InitialDirectory = pptPath;
                }
                saveFileDialog.FileName = pptName + ".pdf";
            }

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                outPath = saveFileDialog.FileName;
                try
                {
                    addIn.ExportCurrentSlideAsPdf(outPath);
                    return true;
                }
                catch (Exception)
                {
                    outPath = null;
                    return false;
                }
            }
            else
            {
                outPath = null;
                return false;
            }
        }


        private void btnExportSlideToPdf_Click(object sender, RibbonControlEventArgs e)
        {
            if (SaveCurrentSlide(out var outPath))
            {
                MessageBox.Show("Slide exported successfully to " + outPath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Failed to export slide.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportContent_Click(object sender, RibbonControlEventArgs e)
        {
            var addIn = Globals.ThisAddIn;
            if (addIn.GetCurrentSlideContentBoundingRect(out Rect rect))
            {
                if (SaveCurrentSlide(out var outPath))
                {
                    addIn.CropPdf(outPath, rect);
                    MessageBox.Show("Slide content exported successfully to " + outPath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Failed to export slide content.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("No visible shapes found on the current slide.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

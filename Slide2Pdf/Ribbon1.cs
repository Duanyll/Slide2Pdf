using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms; // Required for Control.ModifierKeys, Keys, SaveFileDialog, DialogResult, MessageBox

namespace Slide2Pdf
{
    public partial class Ribbon1
    {
        private string currentPresentationFullName = string.Empty;
        private readonly Dictionary<int, string> slideSavePaths = new Dictionary<int, string>();

        // This method is assumed to be in ThisAddIn.cs or a similar helper class
        // public void ExportCurrentSlideAsPdf(string filePath) { /* ... */ }
        // public bool GetCurrentSlideContentBoundingRect(out Rect rect) { /* ... */ }
        // public void CropPdf(string pdfPath, Rect cropRect) { /* ... */ }
        // public struct Rect { public float X1, Y1, X2, Y2; } // Define if not already defined

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Initialization code, if any
        }

        /// <summary>
        /// Checks the current presentation's state. If it's new, unsaved, or changed,
        /// it resets the stored paths.
        /// </summary>
        /// <returns>True if paths can be managed (presentation is saved and identified), false otherwise.</returns>
        private bool UpdatePresentationContextAndPathStatus()
        {
            var app = Globals.ThisAddIn.Application;
            if (app?.ActivePresentation == null)
            {
                currentPresentationFullName = string.Empty;
                slideSavePaths.Clear();
                return false;
            }

            Presentation presentation = app.ActivePresentation;

            // If the presentation isn't saved, we can't reliably track its path or slide IDs persistently.
            if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoFalse || string.IsNullOrEmpty(presentation.FullName))
            {
                // Clear paths if the presentation becomes unsaved, as its identity is now ambiguous
                if (!string.IsNullOrEmpty(currentPresentationFullName))
                {
                    currentPresentationFullName = string.Empty;
                    slideSavePaths.Clear();
                }
                return false;
            }

            // If the presentation's FullName has changed (e.g., saved to a new file, or different presentation opened)
            if (presentation.FullName != currentPresentationFullName)
            {
                currentPresentationFullName = presentation.FullName;
                slideSavePaths.Clear();
            }
            return true;
        }

        /// <summary>
        /// Tries to get the last saved export path for the currently active slide.
        /// </summary>
        /// <returns>The path if found, otherwise null.</returns>
        private string GetSavedPathForCurrentSlide()
        {
            if (!UpdatePresentationContextAndPathStatus())
            {
                return null;
            }

            Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow?.View?.Slide as Slide;
            if (currentSlide != null && slideSavePaths.TryGetValue(currentSlide.SlideID, out string savedPath))
            {
                return savedPath;
            }
            return null;
        }

        /// <summary>
        /// Stores the export path for the currently active slide.
        /// </summary>
        /// <param name="path">The path to save.</param>
        private void StorePathForCurrentSlide(string path)
        {
            if (!UpdatePresentationContextAndPathStatus() || string.IsNullOrEmpty(path))
            {
                return;
            }

            Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow?.View?.Slide as Slide;
            if (currentSlide != null)
            {
                slideSavePaths[currentSlide.SlideID] = path;
            }
        }

        /// <summary>
        /// Handles the core logic of exporting the current slide to PDF.
        /// It prompts for a new path if Shift is pressed or no path is remembered.
        /// </summary>
        /// <param name="forceNewPathSelection">True to force the Save File Dialog, e.g., when Shift is pressed.</param>
        /// <param name="exportedPdfPath">The path where the PDF was saved.</param>
        /// <returns>True if export was successful, false otherwise.</returns>
        private bool ExportCurrentSlideToFile(bool forceNewPathSelection, out string exportedPdfPath)
        {
            exportedPdfPath = null;
            var addIn = Globals.ThisAddIn;
            Presentation presentation = addIn.Application.ActivePresentation;
            Slide currentSlide = addIn.Application.ActiveWindow?.View?.Slide as Slide;

            if (presentation == null || currentSlide == null)
            {
                MessageBox.Show("No active presentation or slide to export.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            string targetPath = forceNewPathSelection ? null : GetSavedPathForCurrentSlide();

            if (string.IsNullOrEmpty(targetPath))
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
                    saveFileDialog.Title = "Save Current Slide as PDF";
                    saveFileDialog.DefaultExt = "pdf";

                    string pptName = Path.GetFileNameWithoutExtension(presentation.Name);
                    // Suggest a filename like "PresentationName_Slide1.pdf"
                    saveFileDialog.FileName = $"{pptName}_Slide{currentSlide.SlideIndex}.pdf";

                    // Set initial directory if presentation is saved locally
                    if (presentation.Saved == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        !string.IsNullOrEmpty(presentation.Path) && // Path is directory
                        !presentation.Path.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    {
                        saveFileDialog.InitialDirectory = presentation.Path;
                    }

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        targetPath = saveFileDialog.FileName;
                        StorePathForCurrentSlide(targetPath);
                    }
                    else
                    {
                        return false; // User cancelled
                    }
                }
            }

            if (string.IsNullOrEmpty(targetPath)) // Should not happen if dialog wasn't cancelled
            {
                return false;
            }

            try
            {
                addIn.ExportCurrentSlideAsPdf(targetPath); 
                exportedPdfPath = targetPath;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to export slide: {ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Optionally, clear the stored path if export failed, to force re-selection next time.
                // if (currentSlide != null) slideSavePaths.Remove(currentSlide.SlideID);
                return false;
            }
        }

        private void btnExportSlideToPdf_Click(object sender, RibbonControlEventArgs e)
        {
            bool forceNewPath = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            if (ExportCurrentSlideToFile(forceNewPath, out string outputPath))
            {
                MessageBox.Show("Slide exported successfully to: " + outputPath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            // If ExportCurrentSlideToFile returns false, it (or its sub-methods) should have already shown an error or handled cancellation.
        }

        private void btnExportContent_Click(object sender, RibbonControlEventArgs e)
        {
            var addIn = Globals.ThisAddIn; // Assuming Rect is defined, e.g. public struct Rect { public float X1, Y1, X2, Y2; }

            if (!addIn.GetCurrentSlideContentBoundingRect(out Rect rect)) // Replace var with your actual Rect type
            {
                MessageBox.Show("No visible shapes found on the current slide, or failed to calculate content bounds.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool forceNewPath = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            if (ExportCurrentSlideToFile(forceNewPath, out string outputPath))
            {
                try
                {
                    addIn.CropPdf(outputPath, rect);
                    MessageBox.Show("Slide content exported and cropped successfully to: " + outputPath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"PDF was exported, but failed to crop: {ex.Message}\nFile: {outputPath}", "Cropping Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

// The controller class
public class ControllerName : Controller
{
    // The action method to display the form
    public IActionResult Index()
    {
        return View();
    }

    // The action method to handle the form submission
    [HttpPost]
    public IActionResult ConvertFiles(IFormFileCollection files, bool defaultNamesCheck, bool pmiCheck, string selectedFolderPath)
    {
        // If no files were selected, return an error message
        if (files.Count == 0)
        {
            ModelState.AddModelError("", "Please select at least one file to convert.");
            return View("Index");
        }

        // If the selected folder path is empty or null, return an error message
        if (string.IsNullOrEmpty(selectedFolderPath))
        {
            ModelState.AddModelError("", "Please select a folder to save the converted files.");
            return View("Index");
        }

        // For each file in the collection
        foreach (var file in files)
        {
            // The file info is obtained
            var fileInfo = new FileInfo(file.FileName);

            // The PDF is read via the stream
            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                var pdfs = new Document(stream);

                // Save options for the Aspose package
                var options = new ExcelSaveOptions();
                options.CloseResponse = true;

                // Get the new filename without extension from the user
                var filenameWithoutExtension = defaultNamesCheck ? fileInfo.Name.Substring(0, fileInfo.Name.Length - 4) : string.Empty;
                if (!defaultNamesCheck && !InputBox("PDFtoXLS", "Enter new Filename", ref filenameWithoutExtension))
                {
                    continue; // Skip this file if the user cancels the filename input
                }

                // The path is combined and the proper extension added
                var newFilePath = Path.Combine(selectedFolderPath, filenameWithoutExtension + ".xlsx");

                // The PDF is converted and saved under the new extension
                pdfs.Save(newFilePath, options);

                // An instance of the ExcelClear class is created
                var clear = new Clearing();

                // The "CleanXls" method is run on the newly saved file
                if (pmiCheck)
                {
                    clear.PMICleanXls(newFilePath);
                }
                else
                {
                    clear.CleanXls(newFilePath);
                }
            }
        }

        // Return a success message
        ViewBag.Message = "Conversion completed successfully.";
        return View("Index");
    }
}

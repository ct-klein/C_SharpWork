using System.IO.Compression;
using System.Xml.Linq;
using JobKeeper.WinForms.Models;

namespace JobKeeper.WinForms.Services;

/// <summary>
/// Service for importing job applications from OpenOffice ODS files
/// </summary>
public class OdsImportService
{
    /// <summary>
    /// Import job applications from an ODS file
    /// </summary>
    public List<JobApplication> ImportFromOds(string filePath)
    {
        var applications = new List<JobApplication>();

        try
        {
            // ODS files are ZIP archives containing XML files
            using var archive = ZipFile.OpenRead(filePath);
            var contentEntry = archive.GetEntry("content.xml");

            if (contentEntry == null)
                throw new InvalidOperationException("Invalid ODS file: content.xml not found");

            using var stream = contentEntry.Open();
            var doc = XDocument.Load(stream);

            // Define namespaces used in ODS files
            XNamespace tableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
            XNamespace textNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            XNamespace officeNs = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

            // Find the first table
            var table = doc.Descendants(tableNs + "table").FirstOrDefault();
            if (table == null)
                throw new InvalidOperationException("No table found in ODS file");

            var rows = table.Elements(tableNs + "table-row").ToList();
            if (rows.Count < 2)
                throw new InvalidOperationException("ODS file must contain at least a header row and one data row");

            // Parse header row to find column indices
            var headerRow = rows[0];
            var headers = ParseRow(headerRow, tableNs, textNs);
            var columnMap = MapColumns(headers);

            // Parse data rows (skip header)
            for (int i = 1; i < rows.Count; i++)
            {
                var row = rows[i];
                var cells = ParseRow(row, tableNs, textNs);

                // Skip empty rows
                if (cells.All(string.IsNullOrWhiteSpace))
                    continue;

                var application = ParseRowToApplication(cells, columnMap);
                if (application != null)
                {
                    applications.Add(application);
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error importing ODS file: {ex.Message}", ex);
        }

        return applications;
    }

    private List<string> ParseRow(XElement row, XNamespace tableNs, XNamespace textNs)
    {
        var cells = new List<string>();

        foreach (var cell in row.Elements(tableNs + "table-cell"))
        {
            // Handle repeated cells
            var repeatAttr = cell.Attribute(tableNs + "number-columns-repeated");
            int repeatCount = repeatAttr != null ? int.Parse(repeatAttr.Value) : 1;

            // Get cell text content
            var textElements = cell.Descendants(textNs + "p");
            var cellText = string.Join(" ", textElements.Select(p => p.Value));

            // Add the cell value (repeated if necessary)
            for (int i = 0; i < repeatCount && cells.Count < 50; i++)
            {
                cells.Add(cellText);
            }
        }

        return cells;
    }

    private Dictionary<string, int> MapColumns(List<string> headers)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < headers.Count; i++)
        {
            var header = headers[i].Trim().ToUpper();

            // Map various possible header names to our property names
            if (header.Contains("COMPANY"))
                map["Company"] = i;
            else if (header.Contains("WEBSITE") || header.Contains("URL") || header.Contains("SITE"))
                map["Website"] = i;
            else if (header.Contains("JOB") && header.Contains("TITLE") || header.Contains("POSITION"))
                map["JobTitle"] = i;
            else if (header.Contains("SUBMIT") || header.Contains("APPLIED") || header.Contains("DATE"))
                map["Submitted"] = i;
            else if (header.Contains("RESUME") || header.Contains("CV"))
                map["Resume"] = i;
            else if (header.Contains("COVER") || header.Contains("LETTER"))
                map["CoverLetter"] = i;
            else if (header.Contains("STATUS"))
                map["Status"] = i;
            else if (header.Contains("INTERVIEW") && (header.Contains("1") || header.Contains("FIRST")))
                map["Interview1"] = i;
            else if (header.Contains("INTERVIEW") && (header.Contains("2") || header.Contains("SECOND")))
                map["Interview2"] = i;
        }

        return map;
    }

    private JobApplication? ParseRowToApplication(List<string> cells, Dictionary<string, int> columnMap)
    {
        try
        {
            // Company is required
            var company = GetCellValue(cells, columnMap, "Company");
            if (string.IsNullOrWhiteSpace(company))
                return null;

            var application = new JobApplication
            {
                Company = company,
                Website = GetCellValue(cells, columnMap, "Website"),
                JobTitle = GetCellValue(cells, columnMap, "JobTitle"),
                Resume = GetCellValue(cells, columnMap, "Resume"),
                CoverLetter = GetCellValue(cells, columnMap, "CoverLetter"),
                Submitted = ParseDate(GetCellValue(cells, columnMap, "Submitted")),
                Interview1 = ParseDate(GetCellValue(cells, columnMap, "Interview1")),
                Interview2 = ParseDate(GetCellValue(cells, columnMap, "Interview2")),
                Status = ParseStatus(GetCellValue(cells, columnMap, "Status"))
            };

            return application;
        }
        catch
        {
            return null;
        }
    }

    private string GetCellValue(List<string> cells, Dictionary<string, int> columnMap, string columnName)
    {
        if (columnMap.TryGetValue(columnName, out int index) && index < cells.Count)
        {
            return cells[index]?.Trim() ?? string.Empty;
        }
        return string.Empty;
    }

    private DateTime? ParseDate(string dateString)
    {
        if (string.IsNullOrWhiteSpace(dateString))
            return null;

        // Try various date formats
        string[] formats = {
            "MM/dd/yyyy", "M/d/yyyy",
            "dd/MM/yyyy", "d/M/yyyy",
            "yyyy-MM-dd", "yyyy/MM/dd",
            "MM-dd-yyyy", "M-d-yyyy"
        };

        foreach (var format in formats)
        {
            if (DateTime.TryParseExact(dateString, format, null, System.Globalization.DateTimeStyles.None, out var date))
            {
                return date;
            }
        }

        // Try general parse
        if (DateTime.TryParse(dateString, out var parsedDate))
        {
            return parsedDate;
        }

        return null;
    }

    private ApplicationStatus ParseStatus(string statusString)
    {
        if (string.IsNullOrWhiteSpace(statusString))
            return ApplicationStatus.Submitted;

        statusString = statusString.Trim().ToUpper();

        if (statusString.Contains("REJECT"))
            return ApplicationStatus.Rejected;
        if (statusString.Contains("INTERVIEW") || statusString.Contains("CHANGED"))
            return ApplicationStatus.InterviewChanged;
        if (statusString.Contains("GHOST"))
            return ApplicationStatus.Ghosted;

        return ApplicationStatus.Submitted;
    }
}

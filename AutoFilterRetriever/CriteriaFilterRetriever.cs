using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace AutoFilterRetriever
{
    /// <summary>
    /// An example class how to get AutoFilter criteria from internal XML representation
    /// </summary>
    public class CriteriaFilterRetriever : IDisposable
    {
        private readonly Excel.Worksheet wks;
        private readonly string filePath;
        private Stream docStream;
        private SpreadsheetDocument openXmlDoc;
        private WorkbookPart wkb;        

        public CriteriaFilterRetriever(Excel.Worksheet sheet)
        {
            wks = sheet;
            filePath = sheet.Application.ActiveWorkbook.FullName;
            if (!filePath.Contains("\\"))
            {
                throw new FileLoadException("Save the file in order to get autofilter criteria");
            }
        }

        /// <summary>
        /// This can be changed to a complex object instead of just list of strings
        /// </summary>
        public List<string> FilterCriteria { get; private set; }

        public void GetFilterCriteria()
        {
            if (!OpenFile()) throw new FileLoadException($"Couldn't open the file - {filePath}");
            if (wks.AutoFilter == null) return;
                        
            // here we get sheet in the workbook.xml (Equals don't work here)
            var sheetInWkb = wkb.Workbook.Descendants<Sheet>().Where(s => s.Name == wks.Name).FirstOrDefault();
            // get a reference to the worksheet part. Imagine part as the folder in the zip structure
            WorksheetPart wsPart = (WorksheetPart)(wkb.GetPartById(sheetInWkb.Id));
            // finally get the xml file e.g. sheet1.xml
            var sheet = wsPart.Worksheet;
            // there should be just one autofilter per sheet
            var filter = sheet.Descendants<AutoFilter>().First();
            if (filter == null) throw new InvalidOperationException($"Couldn't get autofilter data from the {wks.Name} sheet.");
            ManageFilterData(filter);
        }

        private void ManageFilterData(AutoFilter filter)
        {
            FilterCriteria = new List<string>();
            // this is always the first element in AutoFilter
            foreach (FilterColumn filterCol in filter)
            {
                // here we get the filters data
                var filters = filterCol.FirstChild;
                if (filters is Filters)
                {
                    foreach (var item in filters)
                    {
                        if (item is DateGroupItem)
                        {
                            FilterCriteria.Add(GetDateFilterCriteria(item as DateGroupItem));
                        }
                        else if (item is Filter)
                        {
                            FilterCriteria.Add(((Filter)item).Val);
                        }
                        else
                        {
                            throw new Exception("Not sure what to do here");
                        }
                    }
                }
                else if (filters is CustomFilters)
                {
                    // if custom filter is applied (more than one criteria it falls to this category
                    foreach (var item in filters)
                    {
                        if (item is CustomFilter)
                        {
                            var tmp = item as CustomFilter;                            
                            FilterCriteria.Add($"{tmp.Operator}, {tmp.Val}");
                        }                        
                        else
                        {
                            throw new Exception("Not sure what to do here");
                        }
                    }
                }
            }
        }

        private string GetDateFilterCriteria(DateGroupItem criteria)
        {
            if (criteria.DateTimeGrouping == DateTimeGroupingValues.Year)
            {
                return criteria.Year.ToString();
            }
            else if (criteria.DateTimeGrouping == DateTimeGroupingValues.Month)
            {
                return $"{criteria.Year.ToString()}-{criteria.Month.ToString()}";
            }
            else if (criteria.DateTimeGrouping == DateTimeGroupingValues.Day)
            {
                return $"{criteria.Year.ToString()}-{criteria.Month.ToString()}-{criteria.Day.ToString()}";
            }
            else if (criteria.DateTimeGrouping == DateTimeGroupingValues.Hour)
            {
                return $"{criteria.Year.ToString()}-{criteria.Month.ToString()}-{criteria.Day.ToString()} {criteria.Hour.ToString()}:00:00";
            }
            else if (criteria.DateTimeGrouping == DateTimeGroupingValues.Minute)
            {
                return $"{criteria.Year.ToString()}-{criteria.Month.ToString()}-{criteria.Day.ToString()} {criteria.Hour.ToString()}:{criteria.Minute.ToString()}:00";
            }
            else
            {
                return $"{criteria.Year.ToString()}-{criteria.Month.ToString()}-{criteria.Day.ToString()} " +
                       $"{criteria.Hour.ToString()}:{criteria.Minute.ToString()}:{criteria.Second.ToString()}";
            }
        }

        /// <summary> Opens the given file via the DocumentFormat package </summary>
        /// <returns></returns>
        private bool OpenFile()
        {
            try
            {
                docStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                openXmlDoc = SpreadsheetDocument.Open(docStream, false);
                wkb = openXmlDoc.WorkbookPart;                
                return true;
            }
            catch (Exception)
            {                
                return false;
            }
        }

        public void Dispose()
        {
            openXmlDoc?.Close();
            docStream?.Close();
        }

    }
}

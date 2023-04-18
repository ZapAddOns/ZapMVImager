using ScottPlot;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using ZapMVImager.Objects;

namespace ZapMVImager
{
    public static class Export
    {
        static List<string> header = new List<string> { "Planname", "Date", "Time", "TreatmentType", "Beam", "Isocenter", "Node", "ColliSize", "Axial", "Oblique", "Intensity", "FieldsizeInMM", "PlannedMU", "DeliveredMU", "ImagerMU", "DifferenceMU", "DifferencePercent", "CumulativePlannedMU", "CumulativeDeliveredMU", "CumulativeDifferenceMU", "CumulativeDifferencePercent", "IsValid", "IsFlagged" };

        /// <summary>
        /// Save founded entries of one or more log files to a CSV file
        /// </summary>
        /// <param name="filename">Filename to use for the CSV file</param>
        /// <param name="entries">Log entries to save</param>
        public static void ExportCSVData(string filename, List<LogEntry> entries)
        {
            var line = new List<string>();
            var beam = 1;
            var separator = ";";

            // Try to get the separator from apps config file
            separator = ConfigurationManager.AppSettings["CSVSeparator"] != null ? ConfigurationManager.AppSettings["CSVSeparator"] : separator;

            using (var streamWriter = File.CreateText(filename))
            {
                streamWriter.WriteLine(string.Join(separator, header));

                foreach (var entry in entries)
                {
                    line.Clear();

                    line.Add("\"" + entry.PlanName + "\"");
                    line.Add(entry.Time.ToShortDateString());
                    line.Add(entry.Time.ToLongTimeString());
                    line.Add("\"" + entry.TreatmentType + "\"");
                    line.Add(beam.ToString("0"));
                    line.Add(entry.Isocenter.ToString("0"));
                    line.Add(entry.Node.ToString("0"));
                    line.Add(entry.ColliSize.ToString("0.0"));
                    line.Add(entry.Axial.ToString("0.0"));
                    line.Add(entry.Oblique.ToString("0.0"));
                    line.Add(entry.Intensity.ToString("0.000"));
                    line.Add(entry.FieldSizeInMm.ToString("0.0"));
                    line.Add(entry.PlannedMU.ToString("0.000"));
                    line.Add(entry.DeliveredMU.ToString("0.000"));
                    line.Add(entry.ImagerMU.ToString("0.000"));
                    line.Add(entry.DifferenceMU.ToString("0.000"));
                    line.Add(entry.DifferencePercent.ToString("0.000"));
                    line.Add(entry.CumulativeDeliveredMU.ToString("0.000"));
                    line.Add(entry.CumulativeImagerMU.ToString("0.000"));
                    line.Add(entry.CumulativeDifferenceMU.ToString("0.000"));
                    line.Add(entry.CumulativeDifferencePercent.ToString("0.000"));
                    line.Add(entry.IsValid.ToString());
                    line.Add(entry.IsFlagged.ToString());

                    streamWriter.WriteLine(string.Join(separator, line));

                    beam++;
                }

                streamWriter.Close();
            }
        }

        /// <summary>
        /// Save founded entries of one or more log files to a XLSX file
        /// </summary>
        /// <param name="filename">Filename to use for the CSV file</param>
        /// <param name="allEntries">All log entries found in the log files</param>
        /// <param name="planName">Plan name to save data for</param>
        /// <param name="dateToExport">Date of log entries to save. If no date is given, 
        ///                            data for all dates are saved. Each date gets its 
        ///                            own worksheet</param>
        public static void ExportXLSXData(string filename, LogEntries allEntries, string planName, string dateToExport = null)
        {
            var dates = dateToExport != null ? new List<string> { dateToExport } : allEntries.GetDatesForPlan(planName);

            using (SLDocument sl = new SLDocument())
            {
                foreach (var date in dates)
                {
                    var entries = allEntries.GetEntriesForPlanAndDate(planName, DateTime.Parse(date));

                    sl.AddWorksheet(date);

                    for (var i = 0; i < header.Count; i++)
                    {
                        sl.SetCellValue(1, i + 1, header[i]);
                    }

                    var row = 2;

                    foreach (var entry in entries)
                    {
                        var col = 1;

                        sl.SetCellValue(row, col++, entry.PlanName);
                        sl.SetCellValue(row, col++, entry.Time.ToShortDateString());
                        sl.SetCellValue(row, col++, entry.Time.ToLongTimeString());
                        sl.SetCellValue(row, col++, entry.TreatmentType);
                        sl.SetCellValue(row, col++, row - 1);
                        sl.SetCellValue(row, col++, entry.Isocenter);
                        sl.SetCellValue(row, col++, entry.Node);
                        sl.SetCellValue(row, col++, entry.ColliSize);
                        sl.SetCellValue(row, col++, entry.Axial);
                        sl.SetCellValue(row, col++, entry.Oblique);
                        sl.SetCellValue(row, col++, entry.Intensity);
                        sl.SetCellValue(row, col++, entry.FieldSizeInMm);
                        sl.SetCellValue(row, col++, entry.PlannedMU);
                        sl.SetCellValue(row, col++, entry.DeliveredMU);
                        sl.SetCellValue(row, col++, entry.ImagerMU);
                        sl.SetCellValue(row, col++, entry.DifferenceMU);
                        sl.SetCellValue(row, col++, entry.DifferencePercent);
                        sl.SetCellValue(row, col++, entry.CumulativeDeliveredMU);
                        sl.SetCellValue(row, col++, entry.CumulativeImagerMU);
                        sl.SetCellValue(row, col++, entry.CumulativeDifferenceMU);
                        sl.SetCellValue(row, col++, entry.CumulativeDifferencePercent);
                        sl.SetCellValue(row, col++, entry.IsValid.ToString());
                        sl.SetCellValue(row, col++, entry.IsFlagged.ToString());

                        row++;
                    }
                }

                sl.DeleteWorksheet(SLDocument.DefaultFirstSheetName);

                sl.SaveAs(filename);
            }
        }

        public static void ExportChart(string filename, WpfPlot chart, string planName, string date)
        {
            var fontName = "Arial";
            var oldWidth = chart.Plot.Width;
            var oldHeight = chart.Plot.Height;
            var oldAxisFontName = chart.Plot.LeftAxis.AxisLabel.Font.Name;
            var oldAxisFontSize = chart.Plot.LeftAxis.AxisLabel.Font.Size;

            chart.Plot.Resize(1920, 1080);
            chart.Plot.LeftAxis.AxisLabel.Font.Name = fontName;
            chart.Plot.LeftAxis.AxisLabel.Font.Size = 24;
            chart.Plot.BottomAxis.AxisLabel.Font.Name = fontName;
            chart.Plot.BottomAxis.AxisLabel.Font.Size = 24;
            chart.Plot.Layout(left: 80, right: 50, bottom: 70, top: 80);
            chart.Plot.Title($"Comparison of delivered to by MV detected dose for plan \'{planName}\' at {date}", size: 32, fontName: fontName);

            var bitmap = chart.Plot.Render();

            if (Path.GetExtension(filename).ToUpper() == ".JPG")
            {
                Export.ExportJPGData(filename, bitmap);
            }
            else
            {
                Export.ExportPNGData(filename, bitmap);
            }

            chart.Plot.Title(string.Empty);
            chart.Plot.LeftAxis.AxisLabel.Font.Name = oldAxisFontName;
            chart.Plot.LeftAxis.AxisLabel.Font.Size = oldAxisFontSize;
            chart.Plot.BottomAxis.AxisLabel.Font.Name = oldAxisFontName;
            chart.Plot.BottomAxis.AxisLabel.Font.Size = oldAxisFontSize;
            chart.Plot.Resize(oldWidth, oldHeight);
            chart.Plot.ResetLayout();
        }

        public static void ExportPNGData(string filename, Bitmap bitmap)
        {
            bitmap.Save(filename, ImageFormat.Png);
        }

        public static void ExportJPGData(string filename, Bitmap bitmap)
        {
            bitmap.Save(filename, ImageFormat.Jpeg);
        }
    }
}

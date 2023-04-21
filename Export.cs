using DocumentFormat.OpenXml.Spreadsheet;
using ScottPlot;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ZapMVImager.Objects;

namespace ZapMVImager
{
    public static class Export
    {
        static List<string> header = new List<string> { "Planname", "Date", "Time", "TreatmentType", "Beam", "Isocenter", "Node", "ColliSize", "Axial", "Oblique", "Intensity", "FieldsizeInMM", "PlannedMU", "DeliveredMU", "ImagerMU", "DifferenceMU", "DifferencePercent", "CumulativeDeliveredMU", "CumulativeImagerMU", "CumulativeDifferenceMU", "CumulativeDifferencePercent", "CumulativeInside10PercentDifferenceMU", "CumulativeInside10PercentDifferencePercent", "IsValid", "IsFlagged", "IsInside10Percent" };

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

                    var cumulativeInsideDeliveredMU = 0.0;
                    var cumulativeInsideImagerMU = 0.0;

                    if (!entry.IsFlagged && Math.Abs(entry.DifferencePercent) < 10.0)
                    {
                        cumulativeInsideDeliveredMU += entry.DeliveredMU;
                        cumulativeInsideImagerMU += entry.ImagerMU;
                    }

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
                    line.Add((cumulativeInsideImagerMU - cumulativeInsideDeliveredMU).ToString("0.000"));
                    line.Add((cumulativeInsideDeliveredMU == 0 ? 0 : (cumulativeInsideImagerMU - cumulativeInsideDeliveredMU) / cumulativeInsideDeliveredMU * 100.0).ToString("0.000"));
                    line.Add(entry.IsValid.ToString());
                    line.Add(entry.IsFlagged.ToString());
                    line.Add((Math.Abs(entry.DifferencePercent) < 10.0).ToString());

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

                    var cumulativeInsideDeliveredMU = 0.0;
                    var cumulativeInsideImagerMU = 0.0;

                    foreach (var entry in entries)
                    {
                        var col = 1;

                        if (!entry.IsFlagged && Math.Abs(entry.DifferencePercent) < 10.0)
                        {
                            cumulativeInsideDeliveredMU += entry.DeliveredMU;
                            cumulativeInsideImagerMU += entry.ImagerMU;
                        }

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
                        sl.SetCellValue(row, col++, cumulativeInsideImagerMU - cumulativeInsideDeliveredMU);
                        sl.SetCellValue(row, col++, cumulativeInsideDeliveredMU == 0 ? 0 : (cumulativeInsideImagerMU - cumulativeInsideDeliveredMU) / cumulativeInsideDeliveredMU * 100.0);
                        sl.SetCellValue(row, col++, entry.IsValid.ToString());
                        sl.SetCellValue(row, col++, entry.IsFlagged.ToString());
                        sl.SetCellValue(row, col++, (Math.Abs(entry.DifferencePercent) < 10.0).ToString());

                        row++;
                    }
                }

                sl.DeleteWorksheet(SLDocument.DefaultFirstSheetName);

                sl.SaveAs(filename);
            }
        }

        public static void ExportChart(string filename, Plot plot, string planName, string date)
        {
            var bitmap = CreateBitmap(plot, planName, date);

            if (Path.GetExtension(filename).ToUpper() == ".JPG")
            {
                Export.ExportJPGData(filename, bitmap);
            }
            else
            {
                Export.ExportPNGData(filename, bitmap);
            }
        }

        public static void ExportClipboard(Plot plot, string planName, string date)
        {
            var bitmap = CreateBitmap(plot, planName, date);

            Thread thread = new Thread(() =>
            {
                System.Windows.Clipboard.SetImage(CreateBitmapSourceFromBitmap(bitmap));
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public static void ExportPNGData(string filename, Bitmap bitmap)
        {
            bitmap.Save(filename, ImageFormat.Png);
        }

        public static void ExportJPGData(string filename, Bitmap bitmap)
        {
            bitmap.Save(filename, ImageFormat.Jpeg);
        }

        private static Bitmap CreateBitmap(Plot plot, string planName, string date)
        {
            PreparePlot(plot, planName, date);

            return plot.Render();
        }

        public static void PreparePlot(Plot plot, string planName, string date, double factor = 1.0)
        {
            var fontName = "Arial";

            plot.Resize(1920, 1080);
            plot.LeftAxis.AxisLabel.Font.Name = fontName;
            plot.LeftAxis.AxisLabel.Font.Size = (int)(24 / factor);
            plot.BottomAxis.AxisLabel.Font.Name = fontName;
            plot.BottomAxis.AxisLabel.Font.Size = (int)(24 / factor);
            plot.Layout(left: (int)(80 / factor), right: (int)(50 / factor), bottom: (int)(70 / factor), top: (int)(80 / factor));
            plot.Title($"MV error for plan \'{planName}\' at {date}", size: (int)(32 / factor), fontName: fontName);
        }

        // Found at https://www.codeproject.com/Articles/104929/Bitmap-to-BitmapSource
        private static BitmapSource CreateBitmapSourceFromBitmap(Bitmap bitmap)
        {
            if (bitmap == null)
                throw new ArgumentNullException("bitmap");

            if (Application.Current.Dispatcher == null)
                return null; // Is it possible?

            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    // You need to specify the image format to fill the stream. 
                    // I'm assuming it is PNG
                    bitmap.Save(memoryStream, ImageFormat.Png);
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    // Make sure to create the bitmap in the UI thread
                    if (InvokeRequired)
                        return (BitmapSource)Application.Current.Dispatcher.Invoke(
                            new Func<Stream, BitmapSource>(CreateBitmapSourceFromBitmap),
                            DispatcherPriority.Normal,
                            memoryStream);

                    return CreateBitmapSourceFromBitmap(memoryStream);
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool InvokeRequired
        {
            get { return Dispatcher.CurrentDispatcher != Application.Current.Dispatcher; }
        }

        private static BitmapSource CreateBitmapSourceFromBitmap(Stream stream)
        {
            BitmapDecoder bitmapDecoder = BitmapDecoder.Create(
                stream,
                BitmapCreateOptions.PreservePixelFormat,
                BitmapCacheOption.OnLoad);

            // This will disconnect the stream from the image completely...
            WriteableBitmap writable = new WriteableBitmap(bitmapDecoder.Frames.Single());
            writable.Freeze();

            return writable;
        }
    }
}

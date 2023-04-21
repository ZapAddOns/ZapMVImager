using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using ScottPlot;
using ScottPlot.Plottable;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ZapMVImager.Helpers;
using ZapMVImager.Objects;

namespace ZapMVImager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        LogEntries _logEntries;
        List<LogEntry> _activeEntries;
        ObservableCollection<LogFile> _files = new ObservableCollection<LogFile>();
        // Plots
        int _lastCursorPointIndex;
        VLine _cursorPoint;
        MarkerPlot _cursorPointBeam;
        MarkerPlot _cursorPointCumulative;
        VSpan _rangeLevel;
        HLine _lowerLevel;
        HLine _zeroLevel;
        HLine _upperLevel;
        ScatterPlot _plotAllBeams;
        ScatterPlot _plotValidBeams;
        ScatterPlot _plotFlaggedBeams;
        ScatterPlot _plotCumulative;
        ScatterPlot _plotInsideRange;
        ScatterPlot _plotOutsideRange;

        public MainWindow()
        {
            InitializeComponent();

            Title = $"ZapMVImager {System.Reflection.Assembly.GetEntryAssembly().GetName().Version}";

            Loaded += Init;
        }

        private void Init(object sender, RoutedEventArgs e)
        {
            InitChart();
            InitContextMenu();

            Loaded -= Init;
        }

        private void InitContextMenu()
        {
            // unsubscribe from the default right-click menu event
            chart.RightClicked -= chart.DefaultRightClickEvent;

            // add a custom right-click action
            chart.RightClicked += DeployCustomMenu;
        }

        private void DeployCustomMenu(object sender, EventArgs e)
        {
            MenuItem miCopy = new MenuItem() { Header = "Copy" };
            miCopy.Click += CopyToClipboard;

            MenuItem miReopen = new MenuItem() { Header = "Open in new window" };
            miReopen.Click += OpenNewWindow;

            ContextMenu rightClickMenu = new ContextMenu();
            rightClickMenu.Items.Add(miCopy);
            rightClickMenu.Items.Add(new Separator());
            rightClickMenu.Items.Add(miReopen);

            rightClickMenu.IsOpen = true;
        }

        private void OpenNewWindow(object sender, RoutedEventArgs e)
        {
            var newWin = new WpfPlotViewer(chart.Plot.Copy());

            newWin.Title = string.Empty;
            newWin.Width = 1024;
            newWin.Height = 768;

            newWin.wpfPlot1?.ContextMenu?.Items?.Clear();

            Export.PreparePlot(newWin.wpfPlot1.Plot, (string)cbPlans.SelectedItem, (string) cbDates.SelectedItem, 1.5);
            
            newWin.Show();
        }

        private void CopyToClipboard(object sender, RoutedEventArgs e)
        {
            Export.ExportClipboard(chart.Plot, (string)cbPlans.SelectedItem, (string)cbDates.SelectedItem);
        }

        private void UpdateFileProgress(int fileNumber, int maxFileNumber, string filename)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                barProgress.Value = fileNumber;
                lblDetails.Text = filename;
            }));
        }

        private void UpdateLogEntriesProgress(int planNumber, int totalEntries)
        {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    lblPlans.Text = $"Plans: {planNumber}";
                    lblEntries.Text = $"Entries: {totalEntries}";
                }));
        }

        private void UpdateDates()
        {
            var planName = (string)cbPlans.SelectedItem;

            var dates = _logEntries.GetDatesForPlan(planName);

            cbDates.ItemsSource = dates;
            cbDates.SelectedIndex = 0;

            UpdateChart();
        }

        private void BtnFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Log files (*.log)|*.log|Zip files (*.zip)|*.zip|All files (*.*)|*.*";
            var initialDirectory = AppSettings.Get("FilePath", Directory.GetCurrentDirectory());
            if (!Directory.Exists(initialDirectory))
            {
                initialDirectory = Directory.GetCurrentDirectory();
            }
            openFileDialog.InitialDirectory = initialDirectory;
            openFileDialog.FilterIndex = AppSettings.Get("FileFilterIndex", 1);

            if (openFileDialog.ShowDialog() == false)
            {
                return;
            }

            AppSettings.AddOrUpdate("FilePath", Path.GetDirectoryName(openFileDialog.FileNames[0]));
            AppSettings.AddOrUpdate("FileFilterIndex", openFileDialog.FilterIndex.ToString());

            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                var files = new List<LogFile>();

                foreach (var filename in openFileDialog.FileNames)
                {
                    files.AddRange(LogFiles.CreateListOfFiles(filename));
                }

                foreach (var file in files)
                {
                    if (!_files.Contains(file))
                    {
                        _files.Add(file);
                    }
                }

                _files = LogFiles.SortFiles(_files.ToList());

                lblFolderOrFile.Content = "File" + (_files.Count > 1 ? "s" : string.Empty);

                lbFileOrFolder.ItemsSource = _files;

                btnExtract.IsEnabled = _files.Count > 0;
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void BtnFolder_Click(object sender, RoutedEventArgs e)
        {
            var openFolderDialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();

            openFolderDialog.Multiselect = true;

            if (openFolderDialog.ShowDialog() == false)
            {
                return;
            }

            try
            {
                Mouse.OverrideCursor = Cursors.Wait;


                var files = new List<LogFile>();

                foreach (var folder in openFolderDialog.SelectedPaths)
                {
                    files.AddRange(LogFiles.CreateListOfFiles(folder));
                }

                foreach (var file in files)
                {
                    if (!_files.Contains(file))
                    {
                        _files.Add(file);
                    }
                }

                _files = LogFiles.SortFiles(_files.ToList());

                lblFolderOrFile.Content = "File" + (_files.Count > 1 ? "s" : string.Empty);

                lbFileOrFolder.ItemsSource = _files;

                btnExtract.IsEnabled = _files.Count > 0;

            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            _files.Clear();
            _activeEntries = null;

            cbPlans.ItemsSource = null;
            cbDates.ItemsSource = null;

            btnExtract.IsEnabled = false;
            btnExport.IsEnabled = false;
        }

        private async void BtnExtract_Click(object sender, RoutedEventArgs e)
        {
            if (_files == null || _files.Count == 0)
            {
                return;
            }

            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                barProgress.Minimum = 0;
                barProgress.Maximum = _files.Count;

                _logEntries = new LogEntries(_files.ToList(), UpdateFileProgress, UpdateLogEntriesProgress);

                var task = Task.Run(() => _logEntries.CreateLogEntries());

                await task;

                cbPlans.ItemsSource = _logEntries.PlanNames;
                cbPlans.SelectedIndex = 0;
            }
            finally 
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_activeEntries == null)
            {
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.FileName = $"MV Data - {cbPlans.SelectedItem} - {cbDates.SelectedItem}";
            saveFileDialog.Filter = "CSV file (*.csv)|*.csv|Excel file (*.xlsx)|*.xlsx|Excel file with all dates (*.xlsx)|*.xlsx|PNG file (*.png)|*.png|JPG file (*.jpg)|*.jpg|All files (*.*)|*.*";
            var initialDirectory = AppSettings.Get("SaveFilePath", Directory.GetCurrentDirectory());
            if (!Directory.Exists(initialDirectory))
            {
                initialDirectory = Directory.GetCurrentDirectory();
            }
            saveFileDialog.InitialDirectory = initialDirectory;
            saveFileDialog.FilterIndex = AppSettings.Get("SaveFilterIndex", 2);

            if (saveFileDialog.ShowDialog() == false)
            {
                return;
            }

            AppSettings.AddOrUpdate("SaveFilePath", Path.GetDirectoryName(saveFileDialog.FileNames[0]));
            AppSettings.AddOrUpdate("SaveFilterIndex", saveFileDialog.FilterIndex.ToString());

            var filename = saveFileDialog.FileName;

            if (!Path.HasExtension(filename))
            {
                filename = Path.ChangeExtension(filename, ".xlsx");
            }

            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                if (Path.GetExtension(filename).ToUpper() == ".CSV")
                {
                    Export.ExportCSVData(filename, _activeEntries);
                }
                else if (Path.GetExtension(filename).ToUpper() == ".XLSX")
                {
                    if (saveFileDialog.FilterIndex == 3)
                    {
                        Export.ExportXLSXData(filename, _logEntries, (string)cbPlans.SelectedItem);
                    }
                    else
                    {
                        Export.ExportXLSXData(filename, _logEntries, (string)cbPlans.SelectedItem, (string)cbDates.SelectedItem);
                    }
                }
                else if (Path.GetExtension(filename).ToUpper() == ".PNG" || Path.GetExtension(filename).ToUpper() == ".JPG")
                {
                    Export.ExportChart(filename, chart.Plot.Copy(), (string)cbPlans.SelectedItem, (string)cbDates.SelectedItem);
                }
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void Chart_MouseLeave(object sender, MouseEventArgs e)
        {
            _cursorPoint.IsVisible = false;
            _cursorPointBeam.IsVisible = false;
            _cursorPointCumulative.IsVisible = false;
        }

        private void Chart_MouseMove(object sender, MouseEventArgs e)
        {
            double pointAllY = 0;
            double pointCumY = 0;
            int pointIndex = 0;

            if (_activeEntries == null)
            {
                return;
            }

            // determine point nearest the cursor
            (double mouseCoordX, double mouseCoordY) = chart.GetMouseCoordinates();
            double xyRatio = chart.Plot.XAxis.Dims.PxPerUnit / chart.Plot.YAxis.Dims.PxPerUnit;

            try
            {
                (_, _, pointIndex) = _plotAllBeams.GetPointNearestX(mouseCoordX);
                pointAllY = _plotAllBeams.Ys[pointIndex];
                pointCumY = _plotCumulative.Ys[pointIndex];
            }
            catch
            {
                // return;
            }

            _cursorPoint.IsVisible = false;
            _cursorPointBeam.IsVisible = !double.IsNaN(pointAllY);
            _cursorPointCumulative.IsVisible = !double.IsNaN(pointCumY);

            // place the highlight over the point of interest
            _cursorPoint.X = pointIndex + 1;
            _cursorPoint.IsVisible = true;

            if (!double.IsNaN(pointAllY))
            {
                _cursorPointBeam.X = pointIndex + 1;
                _cursorPointBeam.Y = pointAllY;
                _cursorPointBeam.IsVisible = true;
            }

            if (!double.IsNaN(pointCumY))
            {
                _cursorPointCumulative.X = pointIndex + 1;
                _cursorPointCumulative.Y = pointCumY;
                _cursorPointCumulative.IsVisible = true;
            }

            // render if the highlighted point chnaged
            if (_lastCursorPointIndex != pointIndex)
            {
                _lastCursorPointIndex = pointIndex;
                chart.Render();
            }

            var entry = _activeEntries[pointIndex];

            // Update the GUI to describe the cursor point
            lblDetails.Text = $"Time: {entry.Time}";
            lblIsocenter.Text = $"Isocenter: #{ entry.Isocenter}";
            lblColliSize.Text = $"Size [mm]: {entry.ColliSize}";
            lblNode.Text = $"Node: #{entry.Node} (A:{(int)(entry.Axial / 6) * 6}|O:{(int)(entry.Oblique / 6) * 6})";
            lblPlannedMU.Text = $"Planned [MU]: {entry.PlannedMU:0.0}";
            lblDeliveredMU.Text = $"Delivered [MU]: {entry.DeliveredMU:0.0}";
            lblImagerMU.Text = $"MV Imager [MU]: {entry.ImagerMU:0.0}";
            lblDifferencePercent.Text = $"Difference [%]: {entry.DifferencePercent:0.0}";
            lblCumulativeDifferencePercent.Text = $"Cumulative Difference [%]: {entry.CumulativeDifferencePercent:0.0}";
        }

        private void Plans_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            btnExport.IsEnabled = true;

            UpdateDates();
        }

        private void Dates_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateChart();
        }

        private void InitChart()
        {
            var fontName = "Arial";
            var markerSize = AppSettings.Get("MarkerSize", 6);

            chart.Plot.Title(string.Empty, size: 32, fontName: fontName);
            chart.Plot.LeftAxis.Label("Difference [%]", fontName: fontName, size: 20);
            chart.Plot.BottomAxis.Label("Beam", fontName: fontName, size: 20);
            chart.Plot.ResetLayout();

            _rangeLevel = chart.Plot.AddVerticalSpan(-10.0, 10.0, color: System.Drawing.Color.LightGray);
            _upperLevel = chart.Plot.AddHorizontalLine(10.0, color: System.Drawing.Color.DarkBlue);
            _zeroLevel = chart.Plot.AddHorizontalLine(0.0, color: System.Drawing.Color.Black);
            _lowerLevel = chart.Plot.AddHorizontalLine(-10.0, color: System.Drawing.Color.DarkBlue);

            // Add a red circle we can move around later as a highlighted point indicator
            _cursorPoint = chart.Plot.AddVerticalLine(0, System.Drawing.Color.LightPink, 2);
            _cursorPoint.IsVisible = false;
            _cursorPointBeam = chart.Plot.AddPoint(0, 0, size: markerSize + 6, shape: MarkerShape.openCircle, color: System.Drawing.Color.Red);
            _cursorPointBeam.IsVisible = false;
            _cursorPointCumulative = chart.Plot.AddPoint(0, 0, size: markerSize + 6, shape: MarkerShape.openCircle, color: System.Drawing.Color.Red);
            _cursorPointCumulative.IsVisible = false;

            var empty = new double[1] { 0 };

            _plotAllBeams = chart.Plot.AddScatter(empty, empty, markerSize: 0, color: System.Drawing.Color.Gray, label: "All Beams");
            _plotValidBeams = chart.Plot.AddScatter(empty, empty, lineWidth: 0, markerShape: ScottPlot.MarkerShape.filledSquare, markerSize: markerSize, color: System.Drawing.Color.Black, label: "Valid Beams");
            _plotFlaggedBeams = chart.Plot.AddScatter(empty, empty, lineWidth: 0, markerShape: ScottPlot.MarkerShape.eks, markerSize: markerSize, color: System.Drawing.Color.Black, label: "Flagged Beams");
            _plotCumulative = chart.Plot.AddScatter(empty, empty, markerSize: 0, color: System.Drawing.Color.DarkGreen, label: "Cumulative");
            _plotInsideRange = chart.Plot.AddScatter(empty, empty, lineWidth: 0, markerShape: ScottPlot.MarkerShape.filledCircle, markerSize: markerSize, color: System.Drawing.Color.Green, label: "Inside Threshold");
            _plotOutsideRange = chart.Plot.AddScatter(empty, empty, lineWidth: 0, markerShape: ScottPlot.MarkerShape.filledCircle, markerSize: markerSize, color: System.Drawing.Color.Red, label: "Outside Threshold");

            _plotAllBeams.OnNaN = ScatterPlot.NanBehavior.Ignore;
            _plotCumulative.OnNaN = ScatterPlot.NanBehavior.Ignore;

            SetVisibilityOfPlots(false);

            chart.Refresh();
        }

        private void UpdateChart()
        {
            var planName = (string)cbPlans.SelectedItem;

            if (string.IsNullOrEmpty(planName) || string.IsNullOrEmpty((string)cbDates.SelectedItem))
            {
                SetVisibilityOfPlots(false);

                return;
            }

            var date = DateTime.Parse((string)cbDates.SelectedItem);

            var entries = _logEntries.GetEntriesForPlanAndDate(planName, date);

            SetVisibilityOfPlots(false);

            if (entries.Count == 0)
            {
                // Nothing to show, so leave with plots invisible
                return;
            }

            _rangeLevel.IsVisible = AppSettings.Get("Range", false);
            _lowerLevel.IsVisible = true;
            _zeroLevel.IsVisible = true;
            _upperLevel.IsVisible = true;

            _activeEntries = entries;

            var beamNumber = new double[entries.Count()];
            var beamLabel = new string[entries.Count()];

            var allBeamPoints = new double[entries.Count()];
            var cumulativeBeamPoints = new double[entries.Count()];
            var validPointsX = new List<double>();
            var validPointsY = new List<double>();
            var flaggedPointsX = new List<double>();
            var flaggedPointsY = new List<double>();
            var outsidePointsX = new List<double>();
            var outsidePointsY = new List<double>();
            var insidePointsX = new List<double>();
            var insidePointsY = new List<double>();

            var beam = 0;

            foreach (var entry in entries)
            {
                beamNumber[beam] = beam + 1;
                beamLabel[beam] = $"{entry.Isocenter}.{entry.Node}";

                allBeamPoints[beam] = entry.DifferencePercent;

                if (!double.IsNaN(entry.DifferencePercent))
                {
                    if (entry.IsFlagged)
                    {
                        flaggedPointsX.Add(beam + 1);
                        flaggedPointsY.Add(entry.DifferencePercent);
                    }
                    else
                    {
                        validPointsX.Add(beam + 1);
                        validPointsY.Add(entry.DifferencePercent);
                    }
                }

                if (entry.PlannedMU >= 10.0)
                {
                    cumulativeBeamPoints[beam] = entry.CumulativeDifferencePercent;
                }
                else
                {
                    cumulativeBeamPoints[beam] = double.NaN;
                }

                if (!double.IsNaN(entry.CumulativeDifferencePercent) && entry.PlannedMU >= 10.0)
                {
                    if (Math.Abs(entry.CumulativeDifferencePercent) >= 10.0)
                    {
                        outsidePointsX.Add(beam + 1);
                        outsidePointsY.Add(entry.CumulativeDifferencePercent);
                    }
                    else
                    {
                        insidePointsX.Add(beam + 1);
                        insidePointsY.Add(entry.CumulativeDifferencePercent);
                    }
                }

                beam++;
            }

            _plotAllBeams.Update(beamNumber, allBeamPoints);
            _plotAllBeams.IsVisible = true;

            if (validPointsX.Count > 0)
            {
                _plotValidBeams.Update(validPointsX.ToArray(), validPointsY.ToArray());
                _plotValidBeams.IsVisible = true;
            }

            if (flaggedPointsX.Count > 0)
            {
                _plotFlaggedBeams.Update(flaggedPointsX.ToArray(), flaggedPointsY.ToArray());
                _plotFlaggedBeams.IsVisible = true;
            }

            _plotCumulative.Update(beamNumber, cumulativeBeamPoints);
            _plotCumulative.IsVisible = true;

            if (insidePointsX.Count > 0)
            {
                _plotInsideRange.Update(insidePointsX.ToArray(), insidePointsY.ToArray());
                _plotInsideRange.IsVisible = true;
            }

            if (outsidePointsX.Count > 0)
            {
                _plotOutsideRange.Update(outsidePointsX.ToArray(), outsidePointsY.ToArray());
                _plotOutsideRange.IsVisible = true;
            }

            //chart.Plot.XAxis.AutomaticTickPositions(beamNumber, beamLabel);
            //chart.Plot.XAxis.TickLabelFormat(b => (int)b-1 > 0 && (int)b-1 < beamLabel.Length ? beamLabel[(int)b-1] : string.Empty);
            chart.Plot.XAxis.MinimumTickSpacing(10);
            chart.Plot.XAxis.TickDensity(2);
            chart.Plot.AxisAuto();

            chart.Plot.Legend(true, ScottPlot.Alignment.LowerRight);

            chart.Refresh();
        }

        private void SetVisibilityOfPlots(bool flag)
        {
            _rangeLevel.IsVisible = flag;
            _lowerLevel.IsVisible = flag;
            _zeroLevel.IsVisible = flag;
            _upperLevel.IsVisible = flag;
            _plotAllBeams.IsVisible = flag;
            _plotValidBeams.IsVisible = flag;
            _plotFlaggedBeams.IsVisible = flag;
            _plotCumulative.IsVisible = flag;
            _plotInsideRange.IsVisible = flag;
            _plotOutsideRange.IsVisible = flag;

            chart.Refresh();
        }
    }
}

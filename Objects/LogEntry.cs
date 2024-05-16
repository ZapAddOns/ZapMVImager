using System;
using System.Text.RegularExpressions;

namespace ZapMVImager.Objects
{
    public class LogEntry
    {
        // Regexs for getting information from the different log file lines.
        // The information about one beam is distributed over more than one line. 
        public static Regex RegexMeasured = new Regex(@"(\d{2}.\d{2}.\d{4})\s(\d{2}:\d{2}:\d{2}\.\d{3}).*TUIController.OnMVImagerDoseMeasured.*intensity:\s*(\d*\.?\d*)\s*estDose:\s*(\d*\.?\d*).*fieldSizeMm:\s*(\d*\.?\d*).*id:\s*(-?\d*).*isValid:\s*(\w*)", RegexOptions.Compiled);
        public static Regex RegexDoseChecker = new Regex(@"(\d{2}.\d{2}.\d{4})\s(\d{2}:\d{2}:\d{2}\.\d{3}).*MVImageDoseChecker:\s*node:\s*(\d*).*beamDoseLIN:\s*(\d*\.?\d*).*beamDoseIMG:\s*(\d*\.?\d*).*beamPercErr:\s*(-?\d*\.?\d*).*FlaggedErroneous:\s*(\w*).*", RegexOptions.Compiled);
        public static Regex RegexCumulative = new Regex(@"(\d{2}.\d{2}.\d{4})\s(\d{2}:\d{2}:\d{2}\.\d{3}).*MVImageDoseChecker:\s*Cumulative\s*\(LIN IMG DIFF PCENT\)\s*(\d*\.?\d*)\s*(\d*\.?\d*)\s*(-?\d*\.?\d*)\s*(-?\d*\.?\d*)", RegexOptions.Compiled);
        public static Regex RegexSystemData = new Regex(@"(\d{2}.\d{2}.\d{4})\s(\d{2}:\d{2}:\d{2}\.\d{3}).*SystemDeliveryData\s*TreatmentType:>\s*(\w*),\s*isTreatment:>\s*(\w*),\s*MV:>\s*(\d*\.?\d*)\s*MU,\s*Collimator:>\s*(\d*\.?\d*)\s*mm,\s*Axial:>\s*(\d*\.?\d*),\s*Oblique:>\s*(\d*\.?\d*),\s*PlanName:>\s*(.*)", RegexOptions.Compiled);
        public static Regex RegexGantryMoveCompleted = new Regex(@"(\d{2}.\d{2}.\d{4})\s(\d{2}:\d{2}:\d{2}\.\d{3}).*OnGantryMoveCompleted.*", RegexOptions.Compiled);
        public static Regex RegexEnd = new Regex(@"(\d{2}.\d{2}.\d{4})\s(\d{2}:\d{2}:\d{2}\.\d{3}).*PlanRecorder: Saved ongoing plan data in file.*", RegexOptions.Compiled);

        public LogEntry()
        {
        }

        public DateTime Time { get; set; }

        public string PlanName { get; set; }

        public string TreatmentType { get; set; }

        public bool IsTreatment { get; set; }

        public double ColliSize { get; set; }

        public double Axial { get; set; }

        public double Oblique { get; set; }

        public int Isocenter { get; set; }

        public int Node { get; set; }

        public int TotalBeam { get; set; }

        public double Intensity { get; set; }

        public double FieldSizeInMm { get; set; }

        public double PlannedMU { get; set; }

        public double DeliveredMU { get; set; }

        public double ImagerMU { get; set; }

        public double CumulativeDeliveredMU { get; set; }

        public double CumulativeImagerMU { get; set; }

        public double DifferenceMU { get => ImagerMU - DeliveredMU; }

        public double DifferencePercent { get => DeliveredMU == 0.0 ? 0.0 : (ImagerMU - DeliveredMU) / DeliveredMU * 100.0; }

        public double CumulativeDifferenceMU { get => CumulativeImagerMU - CumulativeDeliveredMU; }

        public double CumulativeDifferencePercent { get => CumulativeDeliveredMU == 0.0 ? 0.0 : (CumulativeImagerMU - CumulativeDeliveredMU) / CumulativeDeliveredMU * 100.0; }

        public bool IsValid { get; set; }

        public bool IsFlagged { get; set; }
    }
}

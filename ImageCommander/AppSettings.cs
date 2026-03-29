using System;

namespace ImageCommander
{
    public class AppSettings
    {
        public string DefaultSource { get; set; } = string.Empty;
        public string DefaultDestination { get; set; } = string.Empty;
        public bool WatermarkEnabled { get; set; }
        public string WatermarkFile { get; set; } = string.Empty; // absolute path to watermark png in appdata or elsewhere
    }
}

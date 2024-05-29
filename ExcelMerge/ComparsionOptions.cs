namespace ExcelMerge
{
    public class ComparsionOptions
    {
        public bool IgnoreEmptyCells { get; set; }
        public bool YellowBackground { get; set; }

        public bool GreenBackground { get; set; }
        public bool CopyCellsFormat { get; set; }

        public bool IgnoreCase { get; set; }
        public bool IgnoreSpace { get; set; }

        private ComparsionOptions() { }

        public ComparsionOptions(bool ignoreEmptyCells, bool yellowBackground, bool greenBackground, bool copyCellsFormat, bool ignoreCase, bool ignoreSpace)
        {
            this.IgnoreEmptyCells = ignoreEmptyCells;
            this.YellowBackground = yellowBackground;

            this.GreenBackground = greenBackground;
            this.CopyCellsFormat = copyCellsFormat;

            this.IgnoreCase = ignoreCase;
            this.IgnoreSpace = ignoreSpace;
        }
    }
}

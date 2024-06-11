namespace ExcelMerge
{
    public class ComparisonOptions
    {
        public bool IgnoreEmptyCells { get; set; }
        public bool YellowBackground { get; set; }

        public bool GreenBackground { get; set; }
        public bool CopyCellsFormat { get; set; }
		public bool SkipFurtherMatches { get; set; }

		public bool IgnoreCase { get; set; }
        public bool IgnoreSpace { get; set; }

        private ComparisonOptions() { }

        public ComparisonOptions(bool ignoreEmptyCells, bool yellowBackground, bool greenBackground, bool copyCellsFormat, bool skipFurtherMatches, bool ignoreCase, bool ignoreSpace)
        {
            this.IgnoreEmptyCells = ignoreEmptyCells;
            this.YellowBackground = yellowBackground;

            this.GreenBackground = greenBackground;
            this.CopyCellsFormat = copyCellsFormat;
            this.SkipFurtherMatches = skipFurtherMatches;

			this.IgnoreCase = ignoreCase;
            this.IgnoreSpace = ignoreSpace;
        }
    }
}

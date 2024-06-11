namespace ExcelMerge
{
	public struct ComparisonCounts
	{
		public int MatchCount { get; set; }
		public int CellCount { get; set; }
		public int EmptyCellsCount { get; set; }

		public ComparisonCounts()
		{
			this.MatchCount = 0;
			this.CellCount = 0;
			this.EmptyCellsCount = 0;	
		}

		public override string ToString()
		{
			return $"\nКол-во совпадений: {this.MatchCount}\n" +
				   $"Кол-во пройденных ячеек: {this.CellCount}\n" +
				   $"Кол-во пустых ячеек: {this.EmptyCellsCount}\n";
		}
	}
}

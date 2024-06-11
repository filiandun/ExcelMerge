using System.Windows.Media;

namespace ExcelMerge
{
    public class ProgressEventArgs : EventArgs
    {
        public string Message { get; }
        public double Progress { get; }
		public SolidColorBrush Color { get; }


		public ProgressEventArgs(string message, double progress)
        {
            this.Message = message;
            this.Progress = progress;
            this.Color = Brushes.Black;
        }
		public ProgressEventArgs(string message, double progress, SolidColorBrush color)
		{
			this.Message = message;
			this.Progress = progress;
			this.Color = color;
		}
	}
}

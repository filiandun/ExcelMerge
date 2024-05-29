namespace ExcelMerge
{
    public class ProgressEventArgs : EventArgs
    {
        public string Message { get; }
        public double Progress { get; }

        public ProgressEventArgs(string message, double progress)
        {
            Message = message;
            Progress = progress;
        }
    }
}

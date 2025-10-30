namespace MontfoortIT.Library.Streams.FileConvertors
{
    public class CommaSeperatedToObject<T> : SeperatedFileToObject<T>
    {
        private bool _continueReadBlockStarted = false;

        protected override char Seperator => ',';

        protected override char[] ExcludeChars
        {
            get
            {
                return new[] { '\\' };
            }
        }

        protected override bool ContinueRead
        {
            get
            {
                return _continueReadBlockStarted;
            }
        }

        protected override bool ProcessChar(char ch)
        {
            if (ch == '"') // Do not split in " blocks
            {
                _continueReadBlockStarted = !_continueReadBlockStarted;
                return false;
            }

            return base.ProcessChar(ch);
        }
    }
}

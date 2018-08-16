namespace WordInteroper
{
    public class TokenReplacement : WordModification
    {
        public string Token { get; set; }
        public string Replacement { get; set; }
    }

    public class SetCheckBox : TitleTagModification<bool>, IWordCheckBox
    {
        public bool Checked
        {
            get => Value;
            set => Value = value;
        }
    }

    public abstract class TitleTagModification<T> : WordModification
    {
        public string Title { get; set; }
        public string Tag { get; set; }
        protected T Value { get; set; }
    }

    public abstract class WordModification{}
}
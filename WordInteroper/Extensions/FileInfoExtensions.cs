using System.IO;

namespace WordInteroper.Extensions
{
    public static class FileInfoExtensions
    {
        public static FileInfo ChangeExtension(this FileInfo file, string extension)
        {
            string path = Path.ChangeExtension(file.FullName, extension);
            return new FileInfo(path);
        }
    }
}
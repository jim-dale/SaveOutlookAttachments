
namespace SaveOutlookAttachments
{
    using System;
    using System.IO;

    public static class FileHelpers
    {
        public static string GenerateUniqueFileName(string fileName, string saveToDirectory)
        {
            return GenerateUniqueFileName(fileName, String.Empty, saveToDirectory);
        }

        public static string GenerateUniqueFileName(string fileName, string suffix, string saveToDirectory)
        {
            string result = null;

            if (fileName != null)
            {
                int counter = 1;
                string baseFileName = Path.GetFileNameWithoutExtension(fileName) + suffix;
                string extension = Path.GetExtension(fileName);

                string path = $"{baseFileName}{extension}";
                path = Path.Combine(saveToDirectory, path);

                while (File.Exists(path))
                {
                    path = $"{baseFileName} ({counter}){extension}";
                    path = Path.Combine(saveToDirectory, path);
                    ++counter;
                }
                result = path;
            }
            return result;
        }
    }
}

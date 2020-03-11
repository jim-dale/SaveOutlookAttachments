using System;
using System.IO;

namespace SaveOutlookAttachments
{
    public static class FileHelpers
    {
        public static string GetPathWithEnvVars(string value)
        {
            string result = string.Empty;

            if (string.IsNullOrWhiteSpace(value) == false)
            {
                result = Environment.ExpandEnvironmentVariables(value);
                result = Path.GetFullPath(result);
            }

            return result;
        }

        public static string CleanFileName(string path, char replaceChar)
        {
            string result = default;

            if (string.IsNullOrWhiteSpace(path) == false)
            {
                result = path;

                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    result = result.Replace(c, replaceChar);
                }
            }

            return result;
        }

        public static string GenerateUniqueFileName(string folder, string fileName)
        {
            string result = default;

            if (fileName != default)
            {
                string baseFileName = Path.GetFileNameWithoutExtension(fileName);
                string extension = Path.GetExtension(fileName);

                result = baseFileName + extension;
                result = Path.Combine(folder, result);

                int counter = 1;
                while (File.Exists(result))
                {
                    result = $"{baseFileName} ({counter}){extension}";
                    result = Path.Combine(folder, result);
                    ++counter;
                }
            }

            return result;
        }
    }
}

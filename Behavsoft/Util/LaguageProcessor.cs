using System.Globalization;
using System.Threading;

namespace Behavsoft
{
	public class LanguageProcessor
    {
        public static void SetLanguage(string language)
        {
            if (language.ToLower() == "pt-br" && Thread.CurrentThread.CurrentUICulture.Name.ToLower() != "pt-br")
            {
                Thread.CurrentThread.CurrentUICulture = new CultureInfo(language);
                Thread.CurrentThread.CurrentCulture = new CultureInfo(language);
            }
        }
    }
}

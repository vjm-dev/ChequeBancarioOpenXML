using System.Text;
using System.Text.RegularExpressions;

namespace ChequeBancarioOpenXML
{
    public class Utilidades
    {
        public static string SanitizarTexto(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            // Eliminar caracteres de control, incluyendo el carácter nulo 0x00
            string sanitized = Regex.Replace(input, @"[\p{C}-[\r\n\t]]", "");

            // Normalizar el texto a Form C de Unicode (composición canónica)
            sanitized = sanitized.Normalize(NormalizationForm.FormC);

            return sanitized;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace A2B_App.Client.Services
{
    public class FormatService
    {
        /// <summary>
        /// Replace HTML <p></p> with new line \n
        /// </summary>
        /// <param name="source"></param>
        /// <returns>string</returns>
        public string ReplaceTagHtmlParagraph(string source)
        {
            string output = source.Replace("<p>", string.Empty);
            output = output.Replace("<b />", string.Empty);
            output = output.Replace("</p>", "\n\n");

            //remove other html tag
            output = Regex.Replace(output, "<.*?>", String.Empty);

            return output;
        }

        public string PrettyJson(string unPrettyJson)
        {
            var options = new JsonSerializerOptions()
            {
                WriteIndented = true
            };

            var jsonElement = JsonSerializer.Deserialize<JsonElement>(unPrettyJson);

            return JsonSerializer.Serialize(jsonElement, options);
        }
    }
}

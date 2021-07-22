using A2B_App.Shared.Sox;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    public class FormatService
    {
        public string ReplaceTagHtmlParagraph(string source, bool isNewline)
        {
            string output = string.Empty;
            if (source != string.Empty)
            {
                /* original code commented.
                  output = source.Replace("<p>", string.Empty);
                  output = output.Replace("<br />", string.Empty);
                  output = output.Replace("<br/>", string.Empty);
                  output = output.Replace("</p>", isNewline ? "\n\n" : string.Empty);

                  //remove other html tag
                  output = Regex.Replace(output, "<.*?>", String.Empty);
               */

                output = source.Replace("<p>", string.Empty);
                output = output.Replace("</p>", string.Empty);
                output = output.Replace("<br/>", isNewline ? "\n" : string.Empty);
                output = output.Replace("\"", string.Empty);
                output = output.Replace("<li>", string.Empty);
                output = output.Replace("</li>", isNewline ? "\n" : string.Empty);
                output = output.Replace("<ol>", string.Empty);
                output = output.Replace("</ol>", string.Empty);

            }          
            return output;
        }

        public string ReplaceTagHtmlParagraph2(string source, bool isNewline)
        {
            string output = string.Empty;
            if (source != string.Empty)
            {
                output = source.Replace("<p>", string.Empty);
                output = output.Replace("<br />", isNewline ? "\n" : string.Empty);
                output = output.Replace("<br/>", isNewline ? "\n" : string.Empty);
                output = output.Replace("</p>", isNewline ? "\n" : string.Empty);

                //remove other html tag
                output = Regex.Replace(output, "<.*?>", String.Empty);
            }
            return output;
        }

        public string FormatwithNewLine(string source, bool isNewline)
        {
            string output = string.Empty;
            if (source != string.Empty)
            {
                output = source.Replace("<p>", string.Empty);
                output = output.Replace("</p>", string.Empty);
                output = output.Replace("<br/>", isNewline ? "\n" : string.Empty);
                output = output.Replace("\"", string.Empty);
                output = output.Replace("<li>", string.Empty);
                output = output.Replace("</li>", isNewline ? "\n" : string.Empty);
                output = output.Replace("<ol>", string.Empty);
                output = output.Replace("</ol>", string.Empty);

                //remove other html tag
                // output = Regex.Replace(output, "<.*?>", String.Empty);
            }
            return output;
        }

    }
}

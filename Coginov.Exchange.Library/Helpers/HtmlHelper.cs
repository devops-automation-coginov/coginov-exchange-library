using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace Coginov.Exchange.Library.Helpers
{
    public static class HtmlHelper
    {
        public static string PrependHtmlToBody(string html, string htmlToPrepend)
        {
            var htmlToPrependBody = GetBodyInnerHtml(htmlToPrepend);

            var pattern = new Regex("<body.*?>");
            var match = pattern.Match(html);
            var result = match.Success
                ? html.Insert(match.Index + match.Length, htmlToPrependBody)
                : html.Insert(0, htmlToPrependBody);

            return result;
        }

        public static string AppendHtmlToBody(string html, string htmlHeaderToAppend, string htmlToAppend)
        {
            var cleanHtmlToAppend = GetBodyInnerHtml(htmlToAppend);

            var bodyEndIndex = html.IndexOf("</div>");

            var result = bodyEndIndex == -1
                ? $"{html} {htmlHeaderToAppend} {cleanHtmlToAppend}"
                : html.Insert(bodyEndIndex, $"{htmlHeaderToAppend} {cleanHtmlToAppend}");

            return result;
        }

        private static string GetBodyInnerHtml(string html)
        {
            var pattern = new Regex("<body.*?>");

            var match = pattern.Match(html);
            if (!match.Success)
            {
                return html;
            }

            var htmlStartindex = match.Index + match.Length;
            var htmlEndIndex = html.IndexOf("</body>");

            return html.Substring(htmlStartindex, htmlEndIndex - htmlStartindex);
        }

        public static string HtmlBodyReplaceParams(string html, Dictionary<string, string> parameters)
        {
            var paramTemplate = "<b>{0}</b>";

            foreach (var parameter in parameters)
            {
                var paramToReplace = string.Format(paramTemplate, parameter.Key);
                html = html.Replace(paramToReplace, parameter.Value);
            }

            return html;
        }
    }
}

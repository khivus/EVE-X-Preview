using System.Text;
using System.Text.Json.Nodes;
using EveX.Settings.Bridge;

namespace EveX.Settings.State;

public static class TableClipboard
{
    public static string Write(JsonArray rows) => string.Join("\r\n", rows.OfType<JsonArray>().Select(row => string.Join('\t', row.Select(cell => Quote(JsonExtensions.Display(cell))))));
    private static string Quote(string text) => text.IndexOfAny(['\t', '\r', '\n', '"']) < 0 ? text : '"' + text.Replace("\"", "\"\"") + '"';
    public static JsonArray Read(string text, int columns)
    {
        if (text.Length > Protocol.MaximumBytes) throw new FormatException("Clipboard rows exceed the 4 MiB limit.");
        var rows = new JsonArray(); var row = new JsonArray(); var cell = new StringBuilder(); bool quoted = false, quoteClosed = false;
        void EndCell() { row.Add(cell.ToString()); cell.Clear(); quoteClosed = false; }
        void EndRow()
        {
            EndCell();
            if (row.Count != columns) throw new FormatException($"Pasted row {rows.Count + 1} has {row.Count} columns; this table needs {columns}.");
            rows.Add(row); row = new JsonArray();
            if (rows.Count > 10000) throw new FormatException("Paste at most 10,000 rows at once.");
        }
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (quoted)
            {
                if (c == '"' && i + 1 < text.Length && text[i + 1] == '"') { cell.Append('"'); i++; }
                else if (c == '"') { quoted = false; quoteClosed = true; }
                else cell.Append(c);
            }
            else if (c == '"' && cell.Length == 0 && !quoteClosed) quoted = true;
            else if (c == '\t') EndCell();
            else if (c is '\r' or '\n') { if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++; EndRow(); }
            else if (quoteClosed) throw new FormatException("Unexpected text after a quoted clipboard cell.");
            else cell.Append(c);
        }
        if (quoted) throw new FormatException("The clipboard contains an unfinished quoted cell.");
        if (cell.Length != 0 || row.Count != 0 || quoteClosed) EndRow();
        return rows;
    }
}

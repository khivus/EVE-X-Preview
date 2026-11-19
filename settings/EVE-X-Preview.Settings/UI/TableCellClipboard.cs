using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using EveX.Settings.Bridge;
using EveX.Settings.State;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private static JsonNode? PasteCell(JsonNode? value, JsonObject column) =>
        JsonExtensions.Display(value) == "" ? ClearedTableValue(column) : value?.DeepClone();

    private static JsonNode? ClearedTableValue(JsonObject column) => column["default"]?.DeepClone() ??
        column.Text("kind") switch
        {
            "bool" or "boolean" => JsonValue.Create(false),
            "gesture" => JsonValue.Create(0),
            "color" => JsonValue.Create("000000"),
            _ => JsonValue.Create("")
        };

    private sealed class TableCellClipboard
    {
        private sealed record Cell(Border Control, Func<string> Read, Action<string?>? Write);
        private readonly Dictionary<(int Row, int Column), Cell> cells = new();
        private readonly HashSet<(int Row, int Column)> selected = new();
        private readonly MainWindow owner;
        private readonly Panel table;
        private readonly int generation;
        private (int Row, int Column) anchor;

        public TableCellClipboard(MainWindow owner, Panel table)
        {
            this.owner = owner;
            this.table = table;
            generation = owner.pageGeneration;
            table.AddHandler(InputElement.KeyDownEvent, OnKeyDown, RoutingStrategies.Tunnel);
        }

        public void Add(Grid row, int column, Func<string> read, Action<string?>? write = null)
        {
            var border = row.Children.OfType<Border>().Single(cell => Grid.GetColumn(cell) == column);
            var position = (table.Children.IndexOf(row), column);
            cells[position] = new Cell(border, read, write);
            border.Focusable = true;
            border.AddHandler(InputElement.PointerPressedEvent, (_, e) =>
            {
                if (!e.GetCurrentPoint(border).Properties.IsLeftButtonPressed) return;
                var extend = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
                var toggle = e.KeyModifiers.HasFlag(KeyModifiers.Control);
                if (extend && selected.Count != 0)
                {
                    selected.Clear();
                    foreach (var (r,c) in cells.Keys)
                        if (r >= Math.Min(anchor.Row, position.Item1) && r <= Math.Max(anchor.Row, position.Item1) &&
                            c >= Math.Min(anchor.Column, column) && c <= Math.Max(anchor.Column, column))
                            selected.Add((r,c));
                }
                else
                {
                    if (!toggle) selected.Clear();
                    if (!toggle || !selected.Remove(position)) selected.Add(position);
                    anchor = position;
                }
                foreach (var (key, cell) in cells)
                    cell.Control.Background = selected.Contains(key) ? ThemeController.TableSelection(false) : null;
                if (extend || toggle || write is null)
                {
                    border.Focus();
                    if (extend || toggle) e.Handled = true;
                }
            }, RoutingStrategies.Tunnel, true);
        }

        private async void OnKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyModifiers != KeyModifiers.Control || e.Key is not (Key.C or Key.X or Key.V) ||
                selected.Count == 0 || owner.Clipboard is not { } clipboard) return;
            e.Handled = true;
            var addresses = selected.ToHashSet();
            var top = addresses.Min(p => p.Row);
            var left = addresses.Min(p => p.Column);
            try
            {
                if (e.Key != Key.V)
                {
                    var rows = new JsonArray();
                    var bottom = addresses.Max(p => p.Row);
                    var right = addresses.Max(p => p.Column);
                    for (var row = top; row <= bottom; row++)
                    {
                        var values = new JsonArray();
                        for (var column = left; column <= right; column++)
                            values.Add(addresses.Contains((row,column)) ? cells[(row,column)].Read() : "");
                        rows.Add(values);
                    }
                    await clipboard.SetTextAsync(TableClipboard.Write(rows));
                    if (e.Key == Key.X && generation == owner.pageGeneration)
                    {
                        var writes = addresses.Select(p => cells[p].Write).OfType<Action<string?>>().ToArray();
                        foreach (var write in writes) write(null);
                        if (writes.Length != 0) owner.RenderPage();
                    }
                    return;
                }

                if (await clipboard.TryGetTextAsync() is not { Length: > 0 } text) return;
                if (generation != owner.pageGeneration) return;
                var pasted = TableClipboard.ReadGrid(text, cells.Keys.Max(p => p.Column) - left + 1);
                var changes = new List<(Action<string?> Write, string Value)>();
                if (pasted.Count == 1 && pasted[0] is JsonArray { Count: 1 } single && addresses.Count > 1)
                {
                    foreach (var key in addresses)
                        if (cells[key].Write is { } write) changes.Add((write, JsonExtensions.Display(single[0])));
                }
                else
                {
                    if (top + pasted.Count - 1 > cells.Keys.Max(p => p.Row))
                        throw new FormatException("This table has fixed rows; paste within its existing entries.");
                    for (var r = 0; r < pasted.Count; r++)
                    {
                        var row = (JsonArray)pasted[r]!;
                        for (var c = 0; c < row.Count; c++)
                            if (cells.TryGetValue((top+r,left+c), out var cell) && cell.Write is { } write)
                                changes.Add((write, JsonExtensions.Display(row[c])));
                    }
                }
                foreach (var change in changes) change.Write(change.Value);
                if (changes.Count != 0) owner.RenderPage();
            }
            catch (Exception error) { owner.ShowError(error); }
        }
    }
}

using Avalonia.VisualTree;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Input.Platform;
using EveX.Settings.Bridge;
using EveX.Settings.State;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private async void LoadRowsEditor(StackPanel host, string kind, string group, bool inherited, int generation)
    {
        try
        {
            host.Children.Add(Muted("Loading rows"));
            var data = await session.ReadAsync("rows", new() { ["kind"] = kind, ["group"] = group });
            if (generation != pageGeneration) return;
            BuildRowsEditor(host, kind, group, data, inherited);
            RestoreLoadedScroll(generation);
        }
        catch (Exception e) { if (generation == pageGeneration) { host.Children.Clear(); host.Children.Add(new TextBlock { Text = e.Message, Classes = { "error" } }); } }
    }
    private void BuildRowsEditor(StackPanel host, string kind, string group, JsonObject data, bool inherited)
    {
        host.Children.Clear();
        var columns = data.Array("columns").OfType<JsonObject>().ToArray();
        var key = EditKey("rows|" + kind + "|" + group);
        var rows = (JsonArray)(session.GetPending(key)?.Arguments["rows"] as JsonArray ?? data.Array("rows")).DeepClone();
        var readOnlyRows = columns.Any(column => column.Text("kind") == "readonly");
        var selected = new HashSet<(int row, int column)>();
        (int row, int column) anchor = (0, 0);
        var cells = new Dictionary<(int, int), Border>();
        var selectingRange = false;
        var selectingRows = false;
        var rowAnchor = 0;
        var wholeRows = new HashSet<int>();
        var rowHeaders = new Dictionary<int,Button>();
        (int row,int column) activeCell = (0,0);
        string Definitions() => "24," + string.Join(',',columns.Select(c=>c.Text("kind") is "character" or "exe" ? "1.6*" : c.Text("kind")=="hotkey" ? "2*" : "*"));
        void SelectRows(int first,int last,bool toggle=false)
        {
            if(!toggle) wholeRows.Clear();
            for(var row=Math.Min(first,last);row<=Math.Max(first,last);row++)
                if(!toggle || !wholeRows.Remove(row)) wholeRows.Add(row);
            selected.Clear();
            foreach(var row in wholeRows) for(var column=0;column<columns.Length;column++) selected.Add((row,column));
            activeCell=(last,0); PaintSelection();
        }
        var list = new StackPanel();
        var toolbar = new WrapPanel(); host.Children.Add(toolbar);
        var feedback = new TextBlock { TextWrapping = TextWrapping.Wrap }; host.Children.Add(feedback);
        BindFeedback(key, kind, feedback, host);
        bool editable = !inherited && !data.Bool("inherited");
        JsonArray BlankRow() => new(columns.Select(c => c["default"]?.DeepClone() ?? JsonValue.Create("")).ToArray());
        bool Blank(JsonNode? row) => row is JsonArray values && values.All(v => JsonExtensions.Display(v) == "");
        if (!readOnlyRows && (rows.Count == 0 || !Blank(rows.LastOrDefault()))) rows.Add(BlankRow());
        void StageRows()
        {
            var committed = (JsonArray)rows.DeepClone();
            if (!readOnlyRows && committed.Count != 0 && Blank(committed.LastOrDefault())) committed.RemoveAt(committed.Count - 1);
            Stage(key, group == "" ? kind : group + " members", "rows", new() { ["kind"] = kind, ["group"] = group, ["rows"] = committed });
        }
        void PaintSelection()
        {
            foreach(var (row,header) in rowHeaders) {
                var picked = wholeRows.Contains(row);
                header.Background = picked ? ThemeController.TableSelection(true) : Brushes.Transparent;
                if (picked) header.Foreground = ThemeController.SelectionText;
                else header.ClearValue(Button.ForegroundProperty);
                header.BorderBrush = picked ? ThemeController.SelectionOutline : ThemeController.Line;
            }
            foreach (var (address, border) in cells) {
                var picked = selected.Contains(address);
                border.Background = picked ? ThemeController.TableSelection(false) : Brushes.Transparent;
                border.BorderBrush = picked ? ThemeController.SelectionOutline : ThemeController.Line;
                border.BorderThickness = picked ? new Thickness(2) : new Thickness(0,0,1,1);
                // Keep the editor in the same place when its selection outline appears.
                border.Padding = picked ? new Thickness(2,1,3,2) : new Thickness(4,3);
            }
        }
        void Select(int row, int column, bool extend, bool toggle)
        {
            wholeRows.Clear(); activeCell=(row,column);
            if (!extend && !toggle) { selected.Clear(); anchor = (row, column); }
            if (extend)
            {
                selected.Clear();
                for (var r = Math.Min(anchor.row, row); r <= Math.Max(anchor.row, row); r++)
                    for (var c = Math.Min(anchor.column, column); c <= Math.Max(anchor.column, column); c++) selected.Add((r,c));
            }
            else if (!toggle || !selected.Remove((row,column))) selected.Add((row,column));
            PaintSelection();
        }
        void AppendGridRow(int r)
        {
            var definitions = Definitions();
                var ri=r; var values=(JsonArray)rows[r]!;
                var line=new Grid { ColumnDefinitions = new ColumnDefinitions(definitions) };
                var number=new Button { Width=24,MinWidth=0,Content = (r+1).ToString(), CornerRadius=new CornerRadius(0), Padding=new Thickness(0), HorizontalAlignment=HorizontalAlignment.Stretch,VerticalAlignment=VerticalAlignment.Stretch,MinHeight=42,HorizontalContentAlignment=HorizontalAlignment.Center,VerticalContentAlignment=VerticalAlignment.Center };
                number.Classes.Add("row-number"); ToolTip.SetTip(number,"Select row " + (r+1) + (Blank(values)?" — enter a new item here":""));
                number.Click += (_,_)=> {rowAnchor=ri;SelectRows(ri,ri);};
                number.AddHandler(Avalonia.Input.InputElement.PointerPressedEvent, (_,e)=> {
                    if(!e.GetCurrentPoint(number).Properties.IsLeftButtonPressed)return;
                    selectingRows=true;selectingRange=false;
                    if(!e.KeyModifiers.HasFlag(Avalonia.Input.KeyModifiers.Shift)) rowAnchor=ri;
                    SelectRows(rowAnchor,ri,e.KeyModifiers.HasFlag(Avalonia.Input.KeyModifiers.Control));
                    number.Focus();e.Pointer.Capture(host);e.Handled=true;
                },Avalonia.Interactivity.RoutingStrategies.Tunnel,true);
                rowHeaders[r]=number;line.Children.Add(number);
                for(var c=0;c<columns.Length;c++)
                {
                    var ci=c;
                    while(values.Count<=ci) values.Add("");
                    var control=TypedInput(columns[c],values[c],value=> {
                        values[ci]=value?.DeepClone();
                        // A gesture dropdown assigns one action. Table selection is
                        // reserved for explicit clipboard/clear operations here.
                        if(kind!="gestures")
                            foreach(var address in selected.Where(p=>p.column==ci && p.row!=ri && p.row<rows.Count && !Blank(rows[p.row])))rows[address.row]![ci]=value?.DeepClone();
                        StageRows();
                        if(!readOnlyRows && ri==rows.Count-1 && !Blank(values)) {
                            rows.Add(BlankRow());
                            AppendGridRow(rows.Count-1); PaintSelection();
                        }
                    }, true);
                    control.MinWidth=0; control.MaxWidth=double.PositiveInfinity; control.HorizontalAlignment=columns[c].Text("kind") is "bool" or "boolean" ? HorizontalAlignment.Center : HorizontalAlignment.Stretch; control.VerticalAlignment=VerticalAlignment.Center; control.IsEnabled=editable;
                    var cell=new Border { Child=control, BorderBrush=ThemeController.Line, BorderThickness=new Thickness(0,0,1,1), Padding=new Thickness(4,3), MinHeight=42 };
                    cell.AddHandler(Avalonia.Input.InputElement.PointerPressedEvent, (_,e)=>{ if(!e.GetCurrentPoint(cell).Properties.IsLeftButtonPressed)return; selectingRange = kind!="gestures"; Select(ri,ci,e.KeyModifiers.HasFlag(Avalonia.Input.KeyModifiers.Shift),e.KeyModifiers.HasFlag(Avalonia.Input.KeyModifiers.Control)); },Avalonia.Interactivity.RoutingStrategies.Tunnel);
                    cell.PointerEntered += (_,e)=> { if(selectingRange && e.GetCurrentPoint(cell).Properties.IsLeftButtonPressed) Select(ri,ci,true,false); };
                    cells[(r,c)]=cell; Grid.SetColumn(cell,c+1); line.Children.Add(cell);
                }
                line.HorizontalAlignment=HorizontalAlignment.Stretch; list.Children.Add(line);
        }
        void RenderRows()
        {
            list.Children.Clear(); cells.Clear(); rowHeaders.Clear();
            var definitions = Definitions();
            var header = new Grid { ColumnDefinitions = new ColumnDefinitions(definitions),HorizontalAlignment=HorizontalAlignment.Stretch };
            var allRows=new Button { Width=24,MinWidth=0,HorizontalAlignment=HorizontalAlignment.Stretch,Content="№",Padding=new Thickness(0),CornerRadius=new CornerRadius(0),HorizontalContentAlignment=HorizontalAlignment.Center,VerticalContentAlignment=VerticalAlignment.Center };
            allRows.Click+=(_,_)=>{ if(rows.Count>0)SelectRows(0,rows.Count-1); };header.Children.Add(allRows);
            for (var c = 0; c < columns.Length; c++)
            {
                var column = c;
                var title = new Button { Content = columns[c].Text("label"), HorizontalAlignment = HorizontalAlignment.Stretch, CornerRadius = new CornerRadius(0),HorizontalContentAlignment=columns[c].Text("kind") is "bool" or "boolean" ? HorizontalAlignment.Center : HorizontalAlignment.Left,VerticalContentAlignment=VerticalAlignment.Center };
                title.Click += (_, _) => { wholeRows.Clear(); selected.Clear(); for (var r=0;r<rows.Count;r++) selected.Add((r,column)); PaintSelection(); };
                Grid.SetColumn(title,c+1); header.Children.Add(title);
            }
            list.Children.Add(header);
            for (var r=0;r<rows.Count;r++) AppendGridRow(r);
            PaintSelection();
        }
        void Tool(string label,Func<Task> action,bool enabled=true) {
            var button=ActionButton(label,action,false); button.Margin=new Thickness(0,0,8,8); button.IsEnabled=editable&&enabled; toolbar.Children.Add(button);
        }
        async Task Copy()
        {
            if(Clipboard is not {} clipboard || selected.Count==0) return;
            var result=new JsonArray();
            for(var r=selected.Min(p=>p.row);r<=selected.Max(p=>p.row);r++) {
                var line=new JsonArray();
                for(var c=selected.Min(p=>p.column);c<=selected.Max(p=>p.column);c++) line.Add(selected.Contains((r,c))?rows[r]![c]?.DeepClone():JsonValue.Create(""));
                result.Add(line);
            }
            await clipboard.SetTextAsync(TableClipboard.Write(result));
        }
        async Task Paste()
        {
            if(!editable || Clipboard is not {} clipboard || await clipboard.TryGetTextAsync() is not {} text) return;
            await PasteText(text);
        }
        Task PasteText(string text)
        {
            if(!editable) return Task.CompletedTask;
            var startRow=selected.Count==0?0:selected.Min(p=>p.row); var startCol=selected.Count==0?0:selected.Min(p=>p.column);
            var pasted=TableClipboard.ReadGrid(text,columns.Length-startCol);
            if(startCol==0 && columns[0].Text("kind")=="character" && pasted.OfType<JsonArray>().All(row=>row.Count==1))
                pasted=new JsonArray(pasted.OfType<JsonArray>().Where(row=>!string.IsNullOrWhiteSpace(JsonExtensions.Display(row[0]))).Select(row=>(JsonNode?)new JsonArray(JsonValue.Create(JsonExtensions.Display(row[0]).Trim()))).ToArray());
            if(pasted.Count==1 && ((JsonArray)pasted[0]!).Count==1 && selected.Count>1 && wholeRows.Count==0) {
                foreach(var (r,c) in selected) if(columns[c].Text("kind")!="readonly") rows[r]![c]=pasted[0]![0]?.DeepClone();
            } else for(var r=0;r<pasted.Count;r++) {
                while(rows.Count<=startRow+r) rows.Add(BlankRow());
                for(var c=0;c<((JsonArray)pasted[r]!).Count && startCol+c<columns.Length;c++) if(columns[startCol+c].Text("kind")!="readonly") rows[startRow+r]![startCol+c]=pasted[r]![c]?.DeepClone();
            }
            if(!readOnlyRows && !Blank(rows.LastOrDefault())) rows.Add(BlankRow()); StageRows(); RenderRows();
            return Task.CompletedTask;
        }
        host.Tag = new TableInteractions(PasteText);
        void Clear() { if(!editable)return; foreach(var (r,c) in selected) if(columns[c].Text("kind")!="readonly") rows[r]![c]=""; StageRows(); RenderRows(); }
        Tool("Copy",Copy); Tool("Paste",Paste);
        Tool("Clear cells",()=>{Clear();return Task.CompletedTask;});
        void DeleteRows() {
            if(!editable||readOnlyRows)return;
            foreach(var r in selected.Select(p=>p.row).Distinct().OrderDescending()) if(r<rows.Count)rows.RemoveAt(r);
            if(rows.Count==0||!Blank(rows.LastOrDefault()))rows.Add(BlankRow());selected.Clear();wholeRows.Clear();StageRows();RenderRows();
        }
        Tool("Delete rows",()=>{DeleteRows();return Task.CompletedTask;},!readOnlyRows);
        Tool("Move up",()=>{var r=selected.Count==0?0:selected.Min(p=>p.row);if(r>0){var row=rows[r]!.DeepClone();rows.RemoveAt(r);rows.Insert(r-1,row);StageRows();RenderRows();}return Task.CompletedTask;},!readOnlyRows);
        Tool("Move down",()=>{var r=selected.Count==0?0:selected.Min(p=>p.row);if(r<rows.Count-2){var row=rows[r]!.DeepClone();rows.RemoveAt(r);rows.Insert(r+1,row);StageRows();RenderRows();}return Task.CompletedTask;},!readOnlyRows);
        if(columns.FirstOrDefault()?.Text("kind") is "character" or "exe") Tool("Add running windows",async()=> {
            var applications=columns[0].Text("kind")=="exe";
            var result=await session.ReadAsync("command",new(){["name"]="running-windows",["arguments"]=new JsonObject{["applications"]=applications}});
            foreach(var window in await ChooseRunningRowsAsync(result.Array("rows"),applications)) { var row=BlankRow();row[0]=window.Text(applications?"exe":"character");if(applications&&columns.Length>1)row[1]=window.Text("title");rows.Insert(Math.Max(0,rows.Count-1),row); }
            StageRows();RenderRows();
        },session.Role!="offline");
        host.AddHandler(Avalonia.Input.InputElement.KeyDownEvent,async (_,e)=> {
            if(e.KeyModifiers==Avalonia.Input.KeyModifiers.Control && e.Key==Avalonia.Input.Key.A){selected.Clear();for(var r=0;r<rows.Count;r++)for(var c=0;c<columns.Length;c++)selected.Add((r,c));PaintSelection();e.Handled=true;}
            if(e.KeyModifiers==Avalonia.Input.KeyModifiers.Control && e.Key==Avalonia.Input.Key.C){await Copy();e.Handled=true;}
            if(e.KeyModifiers==Avalonia.Input.KeyModifiers.Control && e.Key==Avalonia.Input.Key.V){await Paste();e.Handled=true;}
            if(e.Key==Avalonia.Input.Key.Delete && (wholeRows.Count>0 || selected.Count>1)) {
                var complete=selected.Count>0 && selected.GroupBy(p=>p.row).All(g=>g.Select(p=>p.column).Distinct().Count()==columns.Length);
                if(!readOnlyRows && (wholeRows.Count>0 || complete))DeleteRows();else Clear();e.Handled=true;
            }
            if(e.Key==Avalonia.Input.Key.Enter && e.KeyModifiers==Avalonia.Input.KeyModifiers.None) {
                var next=Math.Min(activeCell.row+1,rows.Count-1);Select(next,activeCell.column,false,false);
                if(cells.TryGetValue((next,activeCell.column),out var target)) {
                    var editor=target.GetVisualDescendants().OfType<TextBox>().FirstOrDefault();
                    if(editor is not null) editor.Focus();else target.Child?.Focus();target.BringIntoView();
                }
                e.Handled=true;
            }
        },Avalonia.Interactivity.RoutingStrategies.Tunnel);
        host.AddHandler(Avalonia.Input.InputElement.PointerMovedEvent, (_,e)=> {
            if(!e.GetCurrentPoint(host).Properties.IsLeftButtonPressed)return;
            if(selectingRows) {
                foreach(var (row,header) in rowHeaders) {var point=e.GetPosition(header);if(point.Y>=0&&point.Y<header.Bounds.Height){SelectRows(rowAnchor,row);e.Handled=true;break;}}
                return;
            }
            if (!selectingRange) return;
            foreach(var (address,cell) in cells) {
                var point = e.GetPosition(cell);
                if(point.X >= 0 && point.Y >= 0 && point.X < cell.Bounds.Width && point.Y < cell.Bounds.Height && address != anchor) {
                    Select(address.Item1,address.Item2,true,false); e.Handled=true; break;
                }
            }
        },Avalonia.Interactivity.RoutingStrategies.Tunnel,true);
        host.AddHandler(Avalonia.Input.InputElement.PointerReleasedEvent,(_,e)=>{selectingRange=false;if(selectingRows){selectingRows=false;e.Pointer.Capture(null);}},Avalonia.Interactivity.RoutingStrategies.Tunnel,true);
        host.Children.Add(Muted(kind=="gestures"
            ? "Each gesture dropdown changes only its own action. Shift-click selects a range; Ctrl-click adds cells. Ctrl+A/C/V selects, copies and pastes. Paste one value to fill a selected range. Enter moves down."
            : "Shift-click or drag to select a range. Ctrl-click adds cells. Ctrl+A/C/V selects, copies and pastes. Paste one value to fill a range. Drag row numbers to select complete rows; Delete removes them. Enter moves down. Paste newline-separated names into the character column."));
        host.Children.Add(new ScrollViewer { Content=list,MaxHeight=520,HorizontalScrollBarVisibility=Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled });
        RenderRows();
    }
    private async Task<List<JsonNode>> ChooseRunningRowsAsync(JsonArray rows, bool applications)
    {
        if (rows.Count == 0) { await Dialogs.MessageAsync(this, "Running windows", "No matching windows were found. You can enter names manually in the table."); return []; }
        var selected = new HashSet<int>(); var content = new StackPanel { Spacing = 10 };
        content.Children.Add(Muted("Select windows to append. Existing assignments and unfinished edits are kept."));
        for (var i = 0; i < rows.Count; i++)
        {
            var index = i; var item = rows[i];
            var check = new CheckBox { Content = item.Text(applications ? "exe" : "character") + " · " + item.Text("title") };
            check.IsCheckedChanged += (_, _) => { if (check.IsChecked == true) selected.Add(index); else selected.Remove(index); }; content.Children.Add(check);
        }
        return await Dialogs.ChoiceAsync(this, applications ? "Choose running applications" : "Choose open EVE characters", content, [("Append selected", "append"), ("Cancel", "cancel")]) == "append" ? selected.Select(index => rows[index]!).ToList() : [];
    }
}

internal sealed record TableInteractions(Func<string,Task> PasteText);

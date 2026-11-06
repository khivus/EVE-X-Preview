using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using EveX.Settings.Bridge;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;
public sealed partial class MainWindow
{
    private static string SoundName(string path) => string.IsNullOrWhiteSpace(path) ? "None" : Path.GetFileNameWithoutExtension(path);
    private static Border TableCell(Control child,int column) {
        var cell=new Border {Child=child,Padding=new Thickness(4),BorderThickness=new Thickness(0,0,0,1),BorderBrush=new SolidColorBrush(Color.FromArgb(45,128,128,128))};
        Grid.SetColumn(cell,column);return cell;
    }
    private static Grid TableLine(string columns) {
        var row=new Grid {ColumnDefinitions=new ColumnDefinitions(columns),HorizontalAlignment=HorizontalAlignment.Stretch};
        for(var i=0;i<row.ColumnDefinitions.Count;i++)if(row.ColumnDefinitions[i].Width.IsAuto)row.ColumnDefinitions[i].SharedSizeGroup="ActionColumn"+i;
        return row;
    }
    private static void TableHeading(StackPanel table,string columns,params string[] labels) {
        Grid.SetIsSharedSizeScope(table,true);
        var row=TableLine(columns);for(var i=0;i<labels.Length;i++)row.Children.Add(TableCell(new TextBlock {Text=labels[i],FontWeight=FontWeight.SemiBold,TextWrapping=TextWrapping.Wrap},i));table.Children.Add(row);
    }
    private void StageEvent(string id,string label,string field,JsonNode? value) => Stage(EditKey("event|"+id+"|"+field),label+" / "+field,"event",new(){["id"]=id,["values"]=new JsonObject {[field]=value?.DeepClone()}});
    private JsonNode? EventValue(string id,string field,JsonObject data) => session.GetPending(EditKey("event|"+id+"|"+field))?.Arguments.Object("values")[field] ?? data.Object("values")[field];
    private void EventFeedback(StackPanel host,string id,string field,Control input) {
        var feedback=new TextBlock {FontSize=12,TextWrapping=TextWrapping.Wrap};host.Children.Add(feedback);BindFeedback(EditKey("event|"+id+"|"+field),id+":"+field,feedback,input);
    }
    private async void LoadEventOverview(StackPanel table,int generation) {
        try {
            const string columns="1.5*,44,1.2*,1.2*";
            TableHeading(table,columns,"Event","On","Color","Sound");
            var events = new List<(JsonObject descriptor,JsonObject data,CheckBox enabled)>();
            foreach(var descriptor in session.Metadata!.Events) {
                var id=descriptor.Text("id");var data=await session.ReadAsync("event",new(){["id"]=id});if(generation!=pageGeneration)return;
                var row=TableLine(columns);row.Children.Add(TableCell(SearchText(descriptor.Text("label")),0));
                var enabled=new CheckBox {IsChecked=JsonExtensions.IsTrue(EventValue(id,"enabled",data)),VerticalAlignment=VerticalAlignment.Center,IsEnabled=!data.Bool("inherited")};
                Avalonia.Automation.AutomationProperties.SetName(enabled,descriptor.Text("label")+" enabled");
                enabled.IsCheckedChanged+=(_,_)=> {
                    data.Object("values")["enabled"] = enabled.IsChecked == true;
                    StageEvent(id,descriptor.Text("label"),"enabled",JsonValue.Create(enabled.IsChecked==true));
                };
                row.Children.Add(TableCell(enabled,1));
                var colorHost=new StackPanel {Spacing=2};var color=TypedInput(new(){["kind"]="color"},EventValue(id,"color",data),v=>StageEvent(id,descriptor.Text("label"),"color",v),true);color.IsEnabled=!data.Bool("inherited");colorHost.Children.Add(color);EventFeedback(colorHost,id,"color",color);row.Children.Add(TableCell(colorHost,2));
                if(color is Grid colorInput && colorInput.Children.OfType<TextBox>().FirstOrDefault() is {} colorText) {colorText.FontSize=14;colorText.Padding=new Thickness(6,4);}
                var soundPath=data.Object("values").Text("sound");var sound=ActionButton(SoundName(soundPath),async()=>{selections["sound-event"]=id;await NavigateToAsync("sounds");});
                sound.HorizontalAlignment=HorizontalAlignment.Stretch;sound.MinWidth=0;sound.Padding=new Thickness(5);ToolTip.SetTip(sound,"Manage sound assignment in Sound library\n"+soundPath);row.Children.Add(TableCell(sound,3));table.Children.Add(row);
                events.Add((descriptor,data,enabled));
            }
            var options = new StackPanel { Spacing = 16 };
            foreach(var category in events.GroupBy(e=>e.descriptor.Text("category"))) {
                var categoryPanel = new StackPanel { Spacing = 14 };
                categoryPanel.Children.Add(new TextBlock { Text = category.Key, FontWeight = FontWeight.SemiBold, FontSize = 16 });
                foreach(var (descriptor,data,enabled) in category) {
                    var fields = data.Array("fields").OfType<JsonObject>().Where(f=>f.Text("id") is not ("enabled" or "color" or "sound" or "soundVolume")).ToArray();
                    if(fields.Length == 0)continue;
                    var id = descriptor.Text("id");
                    var eventPanel = new StackPanel { Spacing = 10 };
                    eventPanel.Children.Add(new TextBlock { Text = descriptor.Text("label"), FontWeight = FontWeight.SemiBold });
                    var controls = new StackPanel { Spacing = 10, IsEnabled = !data.Bool("inherited") && enabled.IsChecked == true };
                    enabled.IsCheckedChanged += (_,_) => controls.IsEnabled = !data.Bool("inherited") && enabled.IsChecked == true;
                    foreach(var field in fields) {
                        var key = field.Text("id");var fieldPanel = new StackPanel { Spacing = 5 };
                        fieldPanel.Children.Add(SearchText(field.Text("label")));
                        var control = TypedInput(field,EventValue(id,key,data),v=>StageEvent(id,descriptor.Text("label"),key,v));
                        Avalonia.Automation.AutomationProperties.SetName(control,descriptor.Text("label")+" / "+field.Text("label"));
                        fieldPanel.Children.Add(control);EventFeedback(fieldPanel,id,key,control);controls.Children.Add(fieldPanel);
                    }
                    eventPanel.Children.Add(controls);categoryPanel.Children.Add(eventPanel);
                }
                if(categoryPanel.Children.Count > 1)options.Children.Add(Card(categoryPanel));
            }
            if(generation == pageGeneration && options.Children.Count > 0) {
                options.Children.Insert(0,new TextBlock { Text = "Event options", FontSize = 18, FontWeight = FontWeight.SemiBold });
                page.Children.Add(options);
            }
        } catch(Exception e) {if(generation==pageGeneration)table.Children.Add(new TextBlock {Text=e.Message,Classes={"error"}});}
    }
    private Control VolumeInput(JsonNode? value,Action<JsonNode?> changed) {
        var generated=(StackPanel)TypedInput(new(){["kind"]="slider",["min"]=0,["max"]=300,["unit"]="%"},value,changed,true);
        var grid=generated.Children.OfType<Grid>().Single();generated.Children.Clear();
        ToolTip.SetTip(grid,"Volume: 0–300%; above 100% boosts the sound");return grid;
    }
    private async void LoadSoundOverview(StackPanel host,int generation) {
        try {
            var library=await session.ReadAsync("sound",new(){["action"]="list"});if(generation!=pageGeneration)return;
            var events=new List<(JsonObject descriptor,JsonObject data)>();
            foreach(var descriptor in session.Metadata!.Events){var data=await session.ReadAsync("event",new(){["id"]=descriptor.Text("id")});if(generation!=pageGeneration)return;events.Add((descriptor,data));}
            host.Children.Clear();
            var files=new StackPanel {Spacing=4};files.Children.Add(new TextBlock {Text="Library files",FontWeight=FontWeight.SemiBold});
            files.Children.Add(ActionButton("Add sound files",AddSoundFilesAsync));
            files.Children.Add(Muted("Add files once and reuse them below. Removing a reference keeps the original file."));
            const string fileColumns="1.3*,2*,Auto,Auto,Auto";TableHeading(files,fileColumns,"Sound","File gain (%)","","","");
            foreach(var sound in library.Array("rows").OfType<JsonObject>()) {
                var path=sound.Text("path");var row=TableLine(fileColumns);var name=SearchText(SoundName(path));ToolTip.SetTip(name,path);row.Children.Add(TableCell(name,0));
                var key=EditKey("gain|"+path);var group=new StackPanel {Spacing=2};var gain=VolumeInput(session.GetPending(key)?.Arguments["volume"]??sound["volume"],v=>Stage(key,"File gain: "+SoundName(path),"sound",new(){["action"]="gain",["path"]=path,["volume"]=v?.DeepClone()}));group.Children.Add(gain);var feedback=new TextBlock {FontSize=12};group.Children.Add(feedback);BindFeedback(key,"",feedback,gain);row.Children.Add(TableCell(group,1));
                row.Children.Add(TableCell(ActionButton("Assign",()=>AssignSoundAsync(path,sound.Array("assigned"))),2));
                row.Children.Add(TableCell(ActionButton("Play",()=>PlayLibrarySoundAsync(path),false),3));
                row.Children.Add(TableCell(ActionButton("Remove",async()=>{if(await Dialogs.ConfirmAsync(this,"Remove reference?","Keep the original file and remove this library reference?"))await session.MutateAsync("sound",new(){["action"]="remove",["path"]=path});}),4));files.Children.Add(row);
            }
            if(library.Array("rows").Count==0)files.Children.Add(Muted("No files in the library yet. Add WAV or MP3 files above."));host.Children.Add(Card(files));
            var assignments=new StackPanel {Spacing=2};assignments.Children.Add(new TextBlock {Text="Event sound assignments",FontWeight=FontWeight.SemiBold});
            const string assignmentColumns="1.3*,1.4*,2*,Auto";TableHeading(assignments,assignmentColumns,"Event","Sound","Event volume (%)","");
            foreach(var (descriptor,data) in events) {
                var id=descriptor.Text("id");var row=TableLine(assignmentColumns);row.Children.Add(TableCell(SearchText(descriptor.Text("label")),0));
                var current=JsonExtensions.Display(EventValue(id,"sound",data));var paths=library.Array("rows").Select(s=>s.Text("path")).Concat(events.Select(e=>e.data.Object("values").Text("sound"))).Where(p=>p!="").Distinct(StringComparer.OrdinalIgnoreCase).Prepend("").ToArray();
                var choose=new ComboBox {ItemsSource=paths.Select(SoundName).ToArray(),SelectedIndex=Math.Max(0,Array.FindIndex(paths,p=>p.Equals(current,StringComparison.OrdinalIgnoreCase))),HorizontalAlignment=HorizontalAlignment.Stretch,MinWidth=0,IsEnabled=!data.Bool("inherited")};
                Avalonia.Automation.AutomationProperties.SetName(choose,descriptor.Text("label")+" sound");
                choose.SelectionChanged+=(_,_)=>{if(choose.SelectedIndex>=0)StageEvent(id,descriptor.Text("label"),"sound",JsonValue.Create(paths[choose.SelectedIndex]));};ToolTip.SetTip(choose,current);row.Children.Add(TableCell(choose,1));
                var volumeHost=new StackPanel {Spacing=2};var volume=VolumeInput(EventValue(id,"soundVolume",data),v=>StageEvent(id,descriptor.Text("label"),"soundVolume",v));volume.IsEnabled=!data.Bool("inherited");Avalonia.Automation.AutomationProperties.SetName(volume,descriptor.Text("label")+" volume");volumeHost.Children.Add(volume);EventFeedback(volumeHost,id,"soundVolume",volume);row.Children.Add(TableCell(volumeHost,2));
                row.Children.Add(TableCell(ActionButton("Play",async()=>{var preview=session.Metadata!.Commands.First(c=>c.Text("id")=="preview-sound");await RunCommandAsync(preview,new(){["event"]=id});},false),3));
                if(selections.GetValueOrDefault("sound-event")==id)row.Background=new SolidColorBrush(Color.FromArgb(40,110,160,210));assignments.Children.Add(row);
            }
            host.Children.Add(Card(assignments));
        } catch(Exception e){if(generation==pageGeneration)host.Children.Add(new TextBlock {Text=e.Message,Classes={"error"}});}
    }
    private async Task AddSoundFilesAsync() {
        await ProtectModalAsync(true);
        try {
            var files=await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {Title="Add sounds to the shared library",AllowMultiple=true,FileTypeFilter=[new FilePickerFileType("Sounds"){Patterns=["*.wav","*.mp3"]}]});
            foreach(var file in files)if(file.TryGetLocalPath() is {} path)await session.MutateAsync("sound",new(){["action"]="add",["path"]=path});
        } finally {await ProtectModalAsync(false);}
    }
    private Task PlayLibrarySoundAsync(string path)=>RunCommandAsync(session.Metadata!.Commands.First(c=>c.Text("id")=="preview-sound"),new(){["path"]=path});
}

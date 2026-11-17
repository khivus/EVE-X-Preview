using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Presenters;
using Avalonia.Interactivity;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using EveX.Settings.Bridge;
using EveX.Settings.State;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private async Task VerifyPolishAsync(string directory)
    {
        var log=Path.Combine(directory,"polish-checks.log");File.WriteAllText(log,"");
        void Check(bool ok,string label) {if(!ok)throw new InvalidOperationException(label);File.AppendAllText(log,"PASS "+label+Environment.NewLine);}
        uint Rgb(IBrush brush) {var c=((SolidColorBrush)brush).Color;return ((uint)c.R<<16)|((uint)c.G<<8)|c.B;}
        foreach(var mode in new[]{"system","light","dark"}.Concat(session.Metadata!.Value.Array("themes").Select(t=>t.Text("id")))) {
            ThemeController.Apply(mode,session.Metadata);
            var resources=Application.Current!.Resources;
            var tableFill=(SolidColorBrush)ThemeController.TableSelection(false);
            Check(tableFill.Color.A==255 && SelectionContrast.Ratio(Rgb(tableFill),Rgb(ThemeController.SelectionText))>=4.5,
                mode+": opaque table selection keeps row labels readable");
            Check(SelectionContrast.Ratio(Rgb(ThemeController.SelectionOutline),Rgb((IBrush)resources["EveBackgroundBrush"]!))>=3 &&
                SelectionContrast.Ratio(Rgb(ThemeController.SelectionOutline),Rgb((IBrush)resources["EveInputBrush"]!))>=3,
                mode+": table selection outlines contrast with both page and cell backgrounds");
            Check(SelectionContrast.Ratio(Rgb((IBrush)resources["EveActiveSelectionBrush"]!),Rgb((IBrush)resources["EveActiveSelectionTextBrush"]!))>=4.5,
                mode+": selected text and dropdown items remain readable");
        }
        void Click(Button button)=>button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        async Task Shot(string name) {
            await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(UpdateLayout,Avalonia.Threading.DispatcherPriority.Render);
            using var bitmap=new RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height));bitmap.Render(this);bitmap.Save(Path.Combine(directory,name+".png"),PngBitmapEncoderOptions.Default);
        }
        Width=1100;Height=800;
        await NavigateToAsync("about");
        Check(page.GetVisualDescendants().OfType<Button>().Any(b=>b.Content?.ToString()=="Open settings backups" && b.IsEnabled),"About exposes the backup-folder button while editing offline");
        await NavigateToAsync("shortcuts");UpdateLayout();
        Check(page.GetVisualDescendants().OfType<HotkeyEditor>().Count(c=>Avalonia.Automation.AutomationProperties.GetName(c) is "Pause EVE switching hotkeys" or "Pause all hotkeys")==2,"Both independent hotkey pause assignments are visible in Keyboard shortcuts");
        await Shot("pause-hotkeys");
        await NavigateToAsync("appearance");UpdateLayout();
        var delayInput=page.GetVisualDescendants().OfType<Control>().Single(c=>Avalonia.Automation.AutomationProperties.GetName(c)==session.Metadata!.Field("UnfocusedBorderClearDelay").Text("label"));
        Check(delayInput.GetVisualDescendants().OfType<NumericUpDown>().Single().Value==1000 && delayInput.IsEffectivelyEnabled,"The configurable unfocused-border delay is editable with a one-second default");
        delayInput.BringIntoView();await Task.Delay(100);await Shot("focus-border-delay");
        await NavigateToAsync("switching-behavior");
        Check(session.Metadata!.Field("GroupsHoldDelay").Number("min")==30,"The companion exposes the 30 ms group hold-delay minimum");
        await NavigateToAsync("tray");
        Check(page.GetVisualDescendants().OfType<ToggleSwitch>().Count(c=>Avalonia.Automation.AutomationProperties.GetName(c) is "Pause EVE switching hotkeys" or "Pause all hotkeys")==2,"Both hotkey pause commands have independent tray visibility controls");
        await NavigateToAsync("interface");await Task.Delay(150);UpdateLayout();
        var header=(Border)workbench.Children[0];var footer=(Border)workbench.Children.Last();
        var themeChoice=page.GetVisualDescendants().OfType<ComboBox>().Single(c=>Avalonia.Automation.AutomationProperties.GetName(c)==session.Metadata!.Field("SettingsTheme").Text("label"));
        var themeOptions=session.Metadata!.Field("SettingsTheme")!.Array("choices");
        themeChoice.SelectedIndex=Array.FindIndex(themeOptions.Select(o=>o.Text("value")).ToArray(),id=>id=="kitsune");
        Check(((SolidColorBrush)header.Background!).Color==Color.Parse("#"+session.Metadata.Value.Array("themes").Single(p=>p.Text("id")=="kitsune")!.Text("sidebar")),"Clicking the actual theme selector updates the top bar before validation finishes");
        await Task.Delay(350);await NavigateToAsync("appearance");UpdateLayout();
        var font=page.GetVisualDescendants().OfType<ComboBox>().First();
        var sameFont=font;
        foreach(var mode in new[]{"graphite","kitsune","light","cyberpunk","graphite"}) {
            ThemeController.Apply(mode,session.Metadata);
            var palette=session.Metadata!.Value.Array("themes").FirstOrDefault(p=>p.Text("id")==mode);
            var expected=palette is null?Color.Parse("#EFEFEF"):Color.Parse("#"+palette.Text("sidebar"));
            Check(((SolidColorBrush)header.Background!).Color==expected && ((SolidColorBrush)footer.Background!).Color==expected,"Header and footer update immediately to "+mode);
            Check(ReferenceEquals(sameFont,page.GetVisualDescendants().OfType<ComboBox>().First()) && header.BorderThickness.Bottom==1 && footer.BorderThickness.Top==1,"Theme changes preserve controls and the four-zone dividers");
        }
        captureTheme="graphite";ThemeController.Apply(captureTheme,session.Metadata);
        var toggle=page.GetVisualDescendants().OfType<ToggleSwitch>().First();toggle.IsChecked=true;await Task.Delay(350);
        ((IPseudoClasses)toggle.Classes).Set(":pointerover",true);UpdateLayout();
        var onTrack=toggle.GetVisualDescendants().OfType<Border>().Single(b=>b.Name=="SwitchKnobBounds");
        var knob=toggle.GetVisualDescendants().OfType<Avalonia.Controls.Shapes.Ellipse>().Single(b=>b.Name=="SwitchKnobOn");
        Check(onTrack.Background is SolidColorBrush track && knob.Fill is SolidColorBrush fill && track.Color!=fill.Color,"Graphite hover keeps the enabled toggle knob distinct from its track");
        ((IPseudoClasses)toggle.Classes).Set(":pointerover",false);
        search.Text="opacity";await Task.Delay(100);UpdateLayout();
        var icon=(PathIcon)clearSearch.Content!;
        Check(Math.Abs(icon.TranslatePoint(default,clearSearch)!.Value.Y+icon.Bounds.Height/2-clearSearch.Bounds.Height/2)<1 && Math.Abs(icon.TranslatePoint(default,clearSearch)!.Value.X+icon.Bounds.Width/2-clearSearch.Bounds.Width/2)<1,"Search clear icon is vertically centered inside its button");
        await Shot("search-and-chrome");search.Text="";await Task.Delay(100);
        var everDisabled=false;
        void Observe(object? sender,AvaloniaPropertyChangedEventArgs args) {if(args.Property==IsEnabledProperty && (!profiles.IsEnabled||!search.IsEnabled||!page.IsEnabled||!apply.IsEnabled))everDisabled=true;}
        await session.MutateAsync("set",new(){["id"]="ThumbnailTextSize",["value"]=17});UpdateStatus();
        profiles.PropertyChanged+=Observe;search.PropertyChanged+=Observe;page.PropertyChanged+=Observe;apply.PropertyChanged+=Observe;
        foreach(var target in new[]{"sources","layout","monitored-characters","appearance"})await NavigateToAsync(target);
        profiles.PropertyChanged-=Observe;search.PropertyChanged-=Observe;page.PropertyChanged-=Observe;apply.PropertyChanged-=Observe;
        Check(!everDisabled && profiles.IsEnabled && search.IsEnabled && page.IsEnabled,"Category navigation never disables the editor chrome or page");
        await NavigateToAsync("sources");UpdateLayout();
        var folder=page.GetVisualDescendants().OfType<Grid>().First(g=>g.Children.OfType<Button>().Any(b=>b.Content?.ToString()=="Browse"));
        Check(folder.TranslatePoint(default,(Visual)folder.Parent!)!.Value.X==0 && folder.HorizontalAlignment==HorizontalAlignment.Left,"Game and chat folder inputs align with the left edge of their card");
        await NavigateToAsync("monitored-characters");
        Check(session.Metadata!.SectionFields(session.Metadata.Sections.Single(s=>s.Text("id")=="monitored-characters")).Single().Text("id")=="monitorOnlySelectedChars","Monitored characters owns only its character-selection toggle");
        Check(!session.Metadata.SectionFields(session.Metadata.Sections.Single(s=>s.Text("id")=="sources")).Any(f=>f.Text("id")=="monitorOnlySelectedChars"),"Log sources no longer duplicates the character filter");
        await session.MutateAsync("profile",new(){["action"]="add",["name"]="Polish inherited",["copyCurrent"]=true,["inheritance"]=new JsonObject {["appearance"]=true,["events"]=true,["mouse-actions"]=true}});
        await NavigateToAsync("appearance");await Task.Delay(150);UpdateLayout();
        var swatch=page.GetVisualDescendants().OfType<Button>().First(b=>Avalonia.Automation.AutomationProperties.GetName(b)=="Choose color");
        var presenter=swatch.GetVisualDescendants().OfType<ContentPresenter>().First();
        Check(!swatch.IsEffectivelyEnabled && swatch.Background is SolidColorBrush color && presenter.Background is SolidColorBrush visible && color.Color==visible.Color,"Inherited color previews retain their RGB fill while editing is disabled");
        var disabledNumber=page.GetVisualDescendants().OfType<NumericUpDown>().First();
        var innerEditor=disabledNumber.GetVisualDescendants().OfType<TextBox>().Single();
        var innerBorder=innerEditor.GetVisualDescendants().OfType<Border>().Single(b=>b.Name=="PART_BorderElement");
        Check(innerBorder.Background is null || innerBorder.Background is ISolidColorBrush innerFill && innerFill.Color.A==0,"Disabled numeric text uses the outer fill without lighter corners");
        await Shot("inherited-color-preview");
        await NavigateToAsync("mouse-actions");await Task.Delay(180);
        Click(page.GetVisualDescendants().OfType<Button>().Single(b=>b.Content?.ToString()=="Edit Default here"));await Task.Delay(300);
        var gestures=page.GetVisualDescendants().OfType<ComboBox>().Where(c=>c.ItemsSource?.Cast<object>().Any(i=>i?.ToString()=="Disabled")==true).ToArray();
        var rows=await session.ReadAsync("rows",new(){["kind"]="gestures"});
        var original=gestures[0].SelectedIndex;var duplicate=gestures.Skip(1).First(c=>c.SelectedIndex>0).SelectedIndex;
        var secondIndex=Array.FindIndex(gestures,1,c=>c.SelectedIndex>0);var secondOriginal=gestures[secondIndex].SelectedIndex;
        var pointer=new Pointer(9911,PointerType.Mouse,true);var pressed=new PointerPointProperties(RawInputModifiers.LeftMouseButton,PointerUpdateKind.LeftButtonPressed);
        void SelectGesture(int index,bool extend) {
            var cell=(Border)gestures[index].Parent!;
            cell.RaiseEvent(new PointerPressedEventArgs(cell,pointer,this,new Point(0,0),0,pressed,extend?KeyModifiers.Shift:KeyModifiers.None,1));
        }
        var initialChoices=gestures.Select(c=>c.SelectedIndex).ToArray();
        SelectGesture(0,false);SelectGesture(secondIndex,true);gestures[0].SelectedIndex=0;
        Check(gestures.Skip(1).Select(c=>c.SelectedIndex).SequenceEqual(initialChoices.Skip(1)),"A gesture dropdown changes only its own entry even with a selected range");
        var stagedRows=session.Pending.Single(p=>p.Key.Contains("rows|gestures")).Arguments.Array("rows");
        Check(stagedRows.Skip(1).Zip(rows.Array("rows").Skip(1)).All(pair=>JsonNode.DeepEquals(pair.First,pair.Second)),"Independent gesture editing leaves all other staged action assignments unchanged");
        gestures[0].SelectedIndex=original;gestures[secondIndex].SelectedIndex=0;gestures[secondIndex].SelectedIndex=secondOriginal;await Task.Delay(450);
        Check(gestures.Select(c=>c.SelectedIndex).SequenceEqual(initialChoices),"Gesture entries remain independently editable when changed back and forth");
        Check(!session.Pending.Any(e=>e.Key.Contains("gestures")),"Restoring independently edited gesture cells leaves a valid canonical table");
        SelectGesture(0,false);
        gestures[0].SelectedIndex=duplicate;await Task.Delay(400);
        Check(session.Pending.Any(e=>e.Error is not null),"A real duplicate gesture is reported while editing Default in place");
        gestures[0].SelectedIndex=original;await NavigateToAsync("appearance");await Task.Delay(550);await session.ValidateAsync();
        Check(!session.Pending.Any(e=>e.Key.Contains("gestures")) && !CurrentProblems().Any(p=>p.Id=="gestures"),"Restoring a gesture then leaving Default editing clears its former duplicate error");
        await NavigateToAsync("inheritance");await Task.Delay(100);UpdateLayout();
        pageScroll.Offset=new Vector(0,200);await Task.Delay(80);var offset=pageScroll.Offset;
        scrollPositions[ScrollKey]=new Vector(0,20);
        var inherit=page.GetVisualDescendants().OfType<CheckBox>().First(c=>c.TranslatePoint(default,pageScroll)!.Value.Y>0 && c.TranslatePoint(default,pageScroll)!.Value.Y<250);
        inherit.IsChecked=inherit.IsChecked!=true;await Task.Delay(400);
        Check(Math.Abs(pageScroll.Offset.Y-offset.Y)<1,"Changing inheritance retains the current scroll rather than a previous visit's position");
        var creation=CreateProfileAsync();await Task.Delay(250);var dialog=OwnedWindows.Single();UpdateLayout();
        var name=dialog.GetVisualDescendants().OfType<TextBox>().Single(t=>t.PlaceholderText=="Profile name");name.Text="Pinned profile name";
        var create=dialog.GetVisualDescendants().OfType<Button>().Single(b=>b.Content?.ToString()=="Create");
        var choicesScroll=dialog.GetVisualDescendants().OfType<ScrollViewer>().First(v=>v.Content is StackPanel stack && stack.GetVisualDescendants().OfType<CheckBox>().Any(c=>c.Tag is string));
        var nameTop=name.TranslatePoint(default,dialog)!.Value.Y;var buttonTop=create.TranslatePoint(default,dialog)!.Value.Y;
        choicesScroll.Offset=new Vector(0,choicesScroll.Extent.Height);await Task.Delay(100);UpdateLayout();
        Check(name.TranslatePoint(default,dialog)!.Value.Y==nameTop && create.TranslatePoint(default,dialog)!.Value.Y==buttonTop && buttonTop+create.Bounds.Height<=dialog.ClientSize.Height,"Create profile keeps its name and action buttons visible while the categories scroll");
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)dialog.Bounds.Width,(int)dialog.Bounds.Height))) {bitmap.Render(dialog);bitmap.Save(Path.Combine(directory,"pinned-profile.png"),PngBitmapEncoderOptions.Default);}
        Click(dialog.GetVisualDescendants().OfType<Button>().Single(b=>b.Content?.ToString()=="Cancel"));await creation;
        var soundFile=Path.Combine(directory,"Alignment fixture.wav");
        using(var writer=new BinaryWriter(File.Create(soundFile))) {
            writer.Write(System.Text.Encoding.ASCII.GetBytes("RIFF"));writer.Write(836);writer.Write(System.Text.Encoding.ASCII.GetBytes("WAVEfmt "));writer.Write(16);writer.Write((short)1);writer.Write((short)1);writer.Write(8000);writer.Write(8000);writer.Write((short)1);writer.Write((short)8);writer.Write(System.Text.Encoding.ASCII.GetBytes("data"));writer.Write(800);writer.Write(Enumerable.Repeat((byte)128,800).ToArray());
        }
        await session.MutateAsync("sound",new(){["action"]="add",["path"]=soundFile});
        await NavigateToAsync("sounds");await Task.Delay(450);UpdateLayout();
        Check(page.GetVisualDescendants().OfType<TextBlock>().Count(t=>t.Text?.StartsWith("Event sound assignments are inherited")==true)==1,"Sound assignments show one clear inheritance explanation at the top");
        var soundChoice=page.GetVisualDescendants().OfType<ComboBox>().First(c=>Avalonia.Automation.AutomationProperties.GetName(c)?.EndsWith(" sound")==true);
        Check(!soundChoice.IsEnabled,"Inherited event sound assignments remain protected");
        var editSounds=page.GetVisualDescendants().OfType<Button>().Single(b=>b.Content?.ToString()=="Edit Default sound assignments here");Click(editSounds);await Task.Delay(450);
        Check(sectionId=="sounds" && session.Profile=="Default" && defaultEditProfile=="Polish inherited","Sound assignments can edit Default here while retaining the selected profile");
        var volume=page.GetVisualDescendants().OfType<NumericUpDown>().First();
        var slider=page.GetVisualDescendants().OfType<Slider>().First();var thumb=slider.GetVisualDescendants().OfType<Thumb>().Single();UpdateLayout();
        Check(Math.Abs(thumb.TranslatePoint(default,page)!.Value.Y+thumb.Bounds.Height/2-volume.TranslatePoint(default,page)!.Value.Y-volume.Bounds.Height/2)<1,"Sound sliders and their numeric fields share the same visual center");
        Check(page.GetVisualDescendants().OfType<Slider>().All(s=> {
            var n=((Grid)s.Parent!).Children.OfType<NumericUpDown>().Single();var t=s.GetVisualDescendants().OfType<Thumb>().Single();
            return Math.Abs(t.TranslatePoint(default,page)!.Value.Y+t.Bounds.Height/2-n.TranslatePoint(default,page)!.Value.Y-n.Bounds.Height/2)<1;
        }),"Master volume, file gain and all event volume sliders align with their value fields");
        pageScroll.Offset=new Vector(0,Math.Max(0,pageScroll.Extent.Height-pageScroll.Viewport.Height-80));await Task.Delay(100);offset=pageScroll.Offset;
        var stop=page.GetVisualDescendants().OfType<Button>().Single(b=>b.Content?.ToString()=="Stop sound preview");Click(stop);await Task.Delay(120);
        Check(Math.Abs(pageScroll.Offset.Y-offset.Y)<1,"Stop sound preview preserves the current page position");
        await Shot("sound-controls");
        await NavigateToAsync("diagnostics");SetPreviewStatus(new(){["paused"]=true});
        Check(previewPauseButton?.Content?.ToString()=="Resume live previews","Paused diagnostics previews show Resume live previews");SetPreviewStatus(new(){["paused"]=false});
        Check(previewPauseButton?.Content?.ToString()=="Pause live previews","Running diagnostics previews show Pause live previews");
        Check(session.Metadata.Sections.Single(s=>s.Text("id")=="performance").Text("title")=="Advanced","The performance category has a short one-line name");
        await session.DiscardAsync();editHistory.Clear();profileBaselines.Clear();defaultEditProfile=defaultEditSection=defaultEditGroup=null;captureTheme=null;UpdateStatus();
    }
}

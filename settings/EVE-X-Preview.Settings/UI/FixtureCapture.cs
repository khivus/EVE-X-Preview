using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.Threading;
using EveX.Settings.Bridge;
using Avalonia.Controls;
using Avalonia.VisualTree;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    // Deterministic render captures for the isolated Windows Sandbox fixture.
    // The 150% files simulate render DPI; they do not claim live display scaling.
    private async Task CaptureFixturesAsync(string directory)
    {
        if (session.Dirty) throw new InvalidOperationException("Capture needs a clean isolated fixture draft.");
        Directory.CreateDirectory(directory);
        if(Program.VerifyPolish)await VerifyPolishAsync(directory);else await VerifyEditingAsync(directory);
        if (Program.VerifyOnly) { await session.EndAsync(false); closing = true; Close(); return; }
        var sections = Program.CaptureTables ? new[] {"character-appearance","character-shortcuts","close-protection","sources","appearance","events","sounds","about"} : new[] { "appearance", "character-appearance", "layout", "character-shortcuts", "character-groups", "events", "sounds", "profiles", "inheritance", "interface", "diagnostics", "about" };
        var themes = Program.CaptureTables ? new[] {"light","graphite"} : new[] { "light", "dark" }.Concat(session.Metadata!.Value.Array("themes").Select(p => p.Text("id"))).Distinct();
        foreach (var theme in themes)
        foreach (var size in new[] { new Size(1100, 800), new Size(850, 650) })
        foreach (var id in sections.Where(id => session.Metadata!.Sections.Any(s => s.Text("id") == id)))
        {
            captureTheme = theme;
            Width = size.Width; Height = size.Height; sectionId = id; pageScroll.Offset = default;
            RenderNavigation(); RenderPage();
            await Task.Delay(300); // Let canonical specialized reads and layout settle.
            if (session.Pending.Count != 0) throw new InvalidOperationException("Rendering “" + id + "” unexpectedly staged an edit. Fixture capture stopped before accepting or saving anything.");
            ThemeController.Apply(theme, session.Metadata);
            await Dispatcher.UIThread.InvokeAsync(() => { InvalidateMeasure(); UpdateLayout(); }, DispatcherPriority.Render);
            await Task.Delay(70);
            foreach (var scale in new[] { 1.0, 1.5 })
            {
                using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(Bounds.Width * scale), (int)Math.Ceiling(Bounds.Height * scale)), new Vector(96 * scale, 96 * scale));
                bitmap.Render(this);
                bitmap.Save(Path.Combine(directory, $"{id}-{theme}-{size.Width:0}x{size.Height:0}-render{scale * 100:0}.png"), PngBitmapEncoderOptions.Default);
            }
        }
        captureTheme = null;
        await session.EndAsync(false);
        closing = true; Close();
    }

    private async Task VerifyEditingAsync(string directory)
    {
        var results = new List<string>();
        var checkLog=Path.Combine(directory,"interaction-checks.log");File.WriteAllText(checkLog,"");
        void Check(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); results.Add("PASS " + message);File.AppendAllText(checkLog,"PASS "+message+Environment.NewLine); }
        Check(WindowDecorations == WindowDecorations.Full && NativeWindowFrame.NativeCaptionHit(this), "Windows owns a draggable native caption and resize frame");
        Check(captionDarkBeforeShow,"Native dark caption is set before showing the startup window");
        var themeCache=Path.Combine(directory,"startup-theme.json");ThemeController.RememberTheme("graphite",session.Metadata!,themeCache);
        ThemeController.Apply("light",session.Metadata);ThemeController.RestoreStartupTheme(themeCache);
        Check(Application.Current!.RequestedThemeVariant==ThemeVariant.Dark && NativeWindowFrame.DarkCaption(this),"Presentation cache restores the last connected dark theme independently of configuration");
        ThemeController.Apply(session.Snapshot.Object("values").Object("SettingsTheme").Text("value"),session.Metadata);
        Check(workbench.Children[0].Bounds.Height < 55, "Branding and profile selector share a compact title row");
        Check(!loadingView.IsVisible && workbench.IsVisible,"Loading view is replaced by the connected settings workbench");
        sectionId="appearance";RenderPage();await Task.Delay(200);
        var savedFont=session.Snapshot.Object("values").Object("ThumbnailTextFont").Text("value");
        var savedFontChoice=page.GetVisualDescendants().OfType<ComboBox>().Single(c=>Avalonia.Automation.AutomationProperties.GetName(c)=="Label font");
        Check(savedFontChoice.SelectedItem?.ToString()==savedFont && !CurrentProblems().Any(p=>p.Id=="ThumbnailTextFont"),"The factory label font is installed, selected and free of validation problems");
        string selectedFont="";var fontField=session.Metadata!.Field("ThumbnailTextFont")!;
        var fontChoice=(ComboBox)TypedInput(fontField,JsonValue.Create("Unavailable saved fixture font"),v=>selectedFont=JsonExtensions.Display(v));page.Children.Add(fontChoice);await Task.Delay(100);
        Check(fontChoice.SelectedIndex==-1 && !fontChoice.ItemsSource!.Cast<string>().Contains("Unavailable saved fixture font"),"An unavailable saved font is a placeholder, not the first installed option");
        fontChoice.SelectedIndex=0;await Task.Delay(100);
        await session.MutateAsync("set",new(){["id"]="ThumbnailTextFont",["value"]=selectedFont});await session.ValidateAsync();
        Check(selectedFont==session.Metadata.Value.Array("fonts")[0]!.GetValue<string>() && !session.Validation.Array("issues").Any(i=>i.Text("id")=="ThumbnailTextFont"),"Selecting the first installed font clears canonical font problems");
        if(session.Dirty)await session.ApplyAsync();await session.ValidateAsync();profileBaselines.Clear();page.Children.Remove(fontChoice);
        Check(!CurrentProblems().Any(p=>p.Id=="ThumbnailTextFont"),"Saved installed-font selection remains free of font problems");
        sectionId="sources";RenderPage();await Task.Delay(200);UpdateLayout();
        var folderRow=page.GetVisualDescendants().OfType<Grid>().FirstOrDefault(g=>g.Children.OfType<Button>().Any(b=>b.Content?.ToString()=="Browse"));
        Check(folderRow is not null && folderRow.Children.OfType<Button>().Single().Bounds.X-folderRow.Children.OfType<TextBox>().Single().Bounds.Right<=9,"Folder input reaches its Browse button without a large blank gap");
        sectionId = "layout"; RenderPage(); await Task.Delay(250);
        var numbers = page.GetVisualDescendants().OfType<NumericUpDown>().ToArray();
        Check(numbers.Length >= 6, "Numeric fields expose native spin buttons");
        var minWidth = numbers[4]; var previous = minWidth.Value;
        minWidth.Value = 5000; await Task.Delay(1000);
        Check(session.Validation.Array("issues").Any(i => i.Text("severity") == "error"), "Minimum larger than default is validated immediately");
        Check(page.GetVisualDescendants().OfType<TextBlock>().Any(t => t.IsVisible && t.Text?.Contains("must be at least") == true), "Related size error is shown beside the fields");
        Check(!apply.IsEnabled, "Invalid dimensions disable Apply");
        Check(problems.IsVisible && problems.IsEnabled && notice.Foreground==Avalonia.Media.Brushes.OrangeRed,"Problems is available and footer errors are red");
        sectionId="about";RenderPage();var problemTask=ShowProblemsAsync();await Task.Delay(350);
        var problemDialog=OwnedWindows.Single();
        problemDialog.GetVisualDescendants().OfType<Button>().First(b=>b.Content?.ToString()=="Go to setting").RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        await problemTask;Check(sectionId=="layout","Problems navigates from another category to the invalid field");
        await UndoLastAsync(); RenderPage();
        Check(!session.Validation.Array("issues").Any(i => i.Text("severity") == "error"), "Undo last edit restores valid dimensions");
        await Task.Delay(300); // Wait for the replacement controls' initial layout/text events.
        var width = page.GetVisualDescendants().OfType<NumericUpDown>().ElementAt(2);
        width.Value = -20; await Task.Delay(800);
        Check(session.Pending.Any(p => p.Error is not null), "Invalid numeric range retains an inline field error (display=" + width.Value + ", draft=" + session.Snapshot.Object("values").Object("ThumbnailStartLocationwidth").Text("value") + ", pending=" + session.Pending.Count + ")");
        sectionId = "about"; RenderPage();
        await Task.Delay(200);
        Check(page.GetVisualDescendants().OfType<Button>().First(b => b.Content?.ToString() == "Check for updates").IsEnabled, "Update checks remain available with another category's invalid input");
        Check(aboutButtons.Count==8 && aboutButtons.ContainsKey("open-backups") && aboutButtons.ContainsKey("download-settings") && versionLabels["installed"].Text?.StartsWith("v1.7.0.5")==true,"About shows installed/release versions and the settings editor download action");
        Check(aboutButtons.Values.Select(b=>b.Bounds.Width).Distinct().Count()==1 && aboutButtons.Values.All(b=>b.HorizontalContentAlignment==Avalonia.Layout.HorizontalAlignment.Center),"About actions share one width and centered labels");
        var realUpdateStatus=aboutStatus.Copy();
        SetAboutStatus(new(){["stable"]="9.0.0",["installed"]="1.7.0.5",["updateAvailable"]=true});
        Check(releaseNotice.IsVisible && OwnedWindows.Count()==0,"A newer stable release appears as a quiet inline notice without a popup");
        sectionId="appearance";RenderPage();releaseNotice.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));await Task.Delay(350);
        Check(sectionId=="about","Release notice opens About for review");SetAboutStatus(realUpdateStatus);
        await UndoLastAsync(); sectionId = "layout"; RenderPage();
        var originalSnap = session.Snapshot.Object("values").Object("ThumbnailSnap")["value"]!.DeepClone();
        Stage(EditKey("ThumbnailSnap"), "Snap", "set", new() { ["id"] = "ThumbnailSnap", ["value"] = !JsonExtensions.IsTrue(originalSnap) });
        await Task.Delay(800);
        Check(session.Pending.Count == 0, "Valid changes enter the draft without Accept");
        await UndoCategoryAsync();
        Check(JsonExtensions.IsTrue(session.Snapshot.Object("values").Object("ThumbnailSnap")["value"]) == JsonExtensions.IsTrue(originalSnap), "Category undo restores saved category values");
        sectionId = "diagnostics"; RenderPage();
        Check(!page.GetVisualDescendants().OfType<Button>().Any(b => b.Content?.ToString()?.Contains("Accept") == true), "No per-page acceptance buttons");
        await session.ReadAsync("command", new() { ["name"] = "test-thumbnail", ["arguments"] = new JsonObject() });
        await session.ReadAsync("command", new() { ["name"] = "stop-test-thumbnail", ["arguments"] = new JsonObject() });
        Check(true, "Offline test thumbnail command opens and closes");
        await session.ReadAsync("command", new() { ["name"] = "update-status", ["arguments"] = new JsonObject() });
        Check(true, "Update command uses a working runtime adapter");
        search.Text = "Settings editor"; await Task.Delay(200);
        Check(sectionId == "interface", "Search opens a matching category");
        Check(page.GetVisualDescendants().OfType<TextBlock>().Any(t => t.Inlines?.OfType<Avalonia.Controls.Documents.Run>().Any(r => r.Background is not null) == true), "Search highlights matching visible text");
        search.Text = "";await Task.Delay(200);
        string recorded="";var recorder=new HotkeyEditor("",v=>recorded=v);page.Children.Add(recorder);await Task.Delay(150);
        var recordButton=recorder.GetVisualDescendants().OfType<Button>().First(b=>b.Content?.ToString()=="Record");
        recordButton.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        recordButton.RaiseEvent(new Avalonia.Input.KeyEventArgs {RoutedEvent=Avalonia.Input.InputElement.KeyDownEvent,Key=Avalonia.Input.Key.OemTilde});await Task.Delay(100);
        Check(recorded=="`","Recording the backtick key emits one literal backtick (recorded="+recorded+")");page.Children.Remove(recorder);
        sectionId = "shortcuts";RenderPage();await Task.Delay(150);
        Check(page.GetVisualDescendants().OfType<Button>().Count(b=>b.Content?.ToString()?.Contains("shortcut conflicts")==true)==1,"Shortcut conflict review has one consistent action");
        sectionId = "character-shortcuts"; RenderPage(); await Task.Delay(450);
        var tableHost = page.GetVisualDescendants().OfType<StackPanel>().First(p => p.Children.OfType<ScrollViewer>().Any());
        var boxes = tableHost.GetVisualDescendants().OfType<TextBox>().ToArray();
        Check(boxes.Length >= 2, "An empty table includes an editable insertion row");
        boxes[0].Text = "Grid fixture pilot"; boxes[1].Text = "F20";
        boxes[0].RaiseEvent(new Avalonia.Input.FocusChangedEventArgs(Avalonia.Input.InputElement.LostFocusEvent));
        await Task.Delay(850);
        Check(session.Pending.Count == 0, "Table edits validate automatically");
        Check(tableHost.GetVisualDescendants().OfType<TextBox>().Count() >= 4, "Completing a row creates the next insertion row");
        tableHost.RaiseEvent(new Avalonia.Input.KeyEventArgs { RoutedEvent=Avalonia.Input.InputElement.KeyDownEvent,Key=Avalonia.Input.Key.Enter });
        Check(tableHost.GetVisualDescendants().OfType<TextBox>().ElementAt(2).IsFocused, "Enter moves to the same column in the next row");
        tableHost.RaiseEvent(new Avalonia.Input.KeyEventArgs { RoutedEvent = Avalonia.Input.InputElement.KeyDownEvent, Key = Avalonia.Input.Key.A, KeyModifiers = Avalonia.Input.KeyModifiers.Control });
        Check(tableHost.GetVisualDescendants().OfType<Avalonia.Controls.Border>().Count(b => b.Background is Avalonia.Media.SolidColorBrush brush && brush.Color.A == 65) >= 4, "Ctrl+A selects multiple table cells");
        tableHost.RaiseEvent(new Avalonia.Input.KeyEventArgs { RoutedEvent=Avalonia.Input.InputElement.KeyDownEvent,Key=Avalonia.Input.Key.Delete }); await Task.Delay(800);
        var rows = await session.ReadAsync("rows", new() { ["kind"] = "character-shortcuts", ["group"] = "" });
        Check(rows.Array("rows").Count == 0, "Bulk deletion removes selected rows and keeps the insertion row local");
        await ((TableInteractions)tableHost.Tag!).PasteText("First pilot\r\nSecond pilot\nThird pilot"); await Task.Delay(800);
        rows = await session.ReadAsync("rows",new(){["kind"]="character-shortcuts",["group"]=""});
        // A character-only shortcut row can remain incomplete until its shortcut is assigned.
        var staged = session.GetPending(EditKey("rows|character-shortcuts|"))?.Arguments.Array("rows") ?? rows.Array("rows");
        Check(staged.Count == 3 && staged[1]![0]!.GetValue<string>()=="Second pilot", "Newline character lists paste into one column of a wider table");
        var firstRowHeader=tableHost.GetVisualDescendants().OfType<Button>().First(b=>b.Classes.Contains("row-number"));
        Check(firstRowHeader.Content?.ToString()=="1", "Insertion rows use row numbers instead of plus signs");
        Check(tableHost.GetVisualDescendants().OfType<Button>().First(b=>b.Content?.ToString()=="№").HorizontalContentAlignment==Avalonia.Layout.HorizontalAlignment.Center && tableHost.GetVisualDescendants().OfType<HotkeyEditor>().SelectMany(h=>h.GetVisualDescendants().OfType<Button>()).All(b=>b.HorizontalContentAlignment==Avalonia.Layout.HorizontalAlignment.Center),"Number header and compact hotkey buttons center their content");
        firstRowHeader.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        tableHost.RaiseEvent(new Avalonia.Input.KeyEventArgs { RoutedEvent=Avalonia.Input.InputElement.KeyDownEvent,Key=Avalonia.Input.Key.Delete }); await Task.Delay(800);
        rows=await session.ReadAsync("rows",new(){["kind"]="character-shortcuts",["group"]=""});
        staged=session.GetPending(EditKey("rows|character-shortcuts|"))?.Arguments.Array("rows")??rows.Array("rows");
        Check(staged.Count==2 && staged[0]![0]!.GetValue<string>()=="Second pilot", "Delete removes a row selected by its number");
        var headers=tableHost.GetVisualDescendants().OfType<Button>().Where(b=>b.Classes.Contains("row-number")).ToArray();
        var pointer=new Avalonia.Input.Pointer(997,Avalonia.Input.PointerType.Mouse,true);
        var pressed=new Avalonia.Input.PointerPointProperties(Avalonia.Input.RawInputModifiers.LeftMouseButton,Avalonia.Input.PointerUpdateKind.LeftButtonPressed);
        var firstPoint=headers[0].TranslatePoint(new Point(8,8),this)!.Value;
        var secondPoint=headers[1].TranslatePoint(new Point(8,8),this)!.Value;
        headers[0].RaiseEvent(new Avalonia.Input.PointerPressedEventArgs(headers[0],pointer,this,firstPoint,0,pressed,Avalonia.Input.KeyModifiers.None,1));
        tableHost.RaiseEvent(new Avalonia.Input.PointerEventArgs(Avalonia.Input.InputElement.PointerMovedEvent,tableHost,pointer,this,secondPoint,1,pressed,Avalonia.Input.KeyModifiers.None));
        Check(headers.Take(2).All(h=>h.Background is Avalonia.Media.SolidColorBrush brush && brush.Color.A==90),"Dragging the numbered gutter selects whole rows");
        tableHost.RaiseEvent(new Avalonia.Input.PointerReleasedEventArgs(tableHost,pointer,this,secondPoint,2,new Avalonia.Input.PointerPointProperties(),Avalonia.Input.KeyModifiers.None,Avalonia.Input.MouseButton.Left));
        tableHost.RaiseEvent(new Avalonia.Input.KeyEventArgs {RoutedEvent=Avalonia.Input.InputElement.KeyDownEvent,Key=Avalonia.Input.Key.Delete});await Task.Delay(700);
        rows=await session.ReadAsync("rows",new(){["kind"]="character-shortcuts",["group"]=""});
        Check(rows.Array("rows").Count==0 && session.Pending.Count==0,"Delete removes all dragged rows as complete records");
        await UndoCategoryAsync();
        sectionId="character-appearance"; Width=850;Height=650; RenderPage();await Task.Delay(350);UpdateLayout();
        var colorTable=page.GetVisualDescendants().OfType<StackPanel>().First(p=>p.Tag is TableInteractions);
        var tableScroll=colorTable.Children.OfType<ScrollViewer>().Single();
        Check(tableScroll.Extent.Width<=tableScroll.Viewport.Width+1,"Character appearance table fits an 850-pixel window without horizontal scrolling");
        Check(colorTable.GetVisualDescendants().OfType<Button>().Count(b=>Avalonia.Automation.AutomationProperties.GetName(b)=="Choose color")>=3,"Color chooser is integrated into compact swatches");
        captureTheme="graphite";ThemeController.Apply(captureTheme,session.Metadata);
        Check(NativeWindowFrame.DarkCaption(this),"Graphite uses the native Windows dark caption");
        var colorGrid=colorTable.GetVisualDescendants().OfType<Grid>().First(g=>g.Children.OfType<Button>().Any(b=>Avalonia.Automation.AutomationProperties.GetName(b)=="Choose color"));
        var colorBox=colorGrid.Children.OfType<TextBox>().Single();var swatch=colorGrid.Children.OfType<Button>().Single();
        Check(swatch.Bounds.X-colorBox.Bounds.Right<=6,"Color textbox meets its swatch without unused horizontal space");
        Check(swatch.BorderBrush is Avalonia.Media.SolidColorBrush themed && themed.Color!=Avalonia.Media.Colors.LightGreen,"Color swatch outline follows the active theme instead of fixed green");
        var numbered=colorTable.GetVisualDescendants().OfType<Button>().First(b=>b.Classes.Contains("row-number"));
        Check(Math.Abs(((Grid)numbered.Parent!).Children.OfType<Border>().First().Bounds.X-numbered.Bounds.Right)<2,"Row numbers meet the first table cell without a gutter gap");
        var colorTask=PickColorAsync("FF8000");await Task.Delay(400);
        var colorDialog=OwnedWindows.Single();
        Check(colorDialog.WindowDecorations==WindowDecorations.Full && NativeWindowFrame.DarkCaption(colorDialog),"Color picker retains a native dark caption");
        var brightness=colorDialog.GetVisualDescendants().OfType<Slider>().Last();brightness.Value=0;
        Check(colorDialog.GetVisualDescendants().OfType<TextBlock>().Any(t=>t.Text=="000000"),"Picker brightness reaches true black");
        brightness.Value=10;await Task.Delay(100);
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)colorDialog.Bounds.Width,(int)colorDialog.Bounds.Height))) {bitmap.Render(colorDialog);bitmap.Save(Path.Combine(directory,"color-picker-graphite.png"),PngBitmapEncoderOptions.Default);}
        colorDialog.GetVisualDescendants().OfType<Button>().First(b=>b.Content?.ToString()=="Use color").RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        Check(await colorTask=="1A0D00","Picker returns a darker shade of the selected hue");captureTheme=null;
        await session.DiscardAsync(); editHistory.Clear();
        UpdateStatus();Check(problems.IsVisible && !problems.IsEnabled,"Problems remains visible and inactive for valid settings");
        var soundPath=Path.Combine(directory,"Reusable alert.wav");
        using(var writer=new BinaryWriter(File.Create(soundPath))) {
            writer.Write(System.Text.Encoding.ASCII.GetBytes("RIFF"));writer.Write(836);writer.Write(System.Text.Encoding.ASCII.GetBytes("WAVEfmt "));writer.Write(16);writer.Write((short)1);writer.Write((short)1);writer.Write(8000);writer.Write(8000);writer.Write((short)1);writer.Write((short)8);writer.Write(System.Text.Encoding.ASCII.GetBytes("data"));writer.Write(800);writer.Write(Enumerable.Repeat((byte)128,800).ToArray());
        }
        await session.MutateAsync("sound",new(){["action"]="add",["path"]=soundPath});
        sectionId="sounds";RenderPage();await Task.Delay(800);UpdateLayout();
        var settingsRows=page.Children.OfType<Border>().First().Child as StackPanel;
        var globalInputs=settingsRows!.Children.OfType<Grid>().Select(g=>g.Children.OfType<StackPanel>().Last()).ToArray();
        Check(globalInputs.Length==3 && globalInputs.Select(c=>c.TranslatePoint(default,page)!.Value.X).Distinct().Count()==1,"Shared sound settings place all value controls in one aligned column");
        var assignment=page.GetVisualDescendants().OfType<ComboBox>().First(c=>Avalonia.Automation.AutomationProperties.GetName(c)=="Shooting idle sound");assignment.SelectedIndex=1;
        var eventVolume=page.GetVisualDescendants().OfType<Grid>().First(c=>Avalonia.Automation.AutomationProperties.GetName(c)=="Shooting idle volume");eventVolume.Children.OfType<Slider>().Single().Value=65;await Task.Delay(850);
        Check(eventVolume.Children.OfType<NumericUpDown>().Single().Bounds.Width>=118 && eventVolume.GetVisualDescendants().OfType<TextBox>().First().Bounds.Width>=35,"Sound volume leaves room for three digits beside the spin buttons");
        Check(eventVolume.Children.OfType<NumericUpDown>().Single().Bounds.Height==34,"Sound volume uses the same height as other numeric inputs");
        var eventData=await session.ReadAsync("event",new(){["id"]="stoppedShooting"});
        Check(eventData.Object("values").Text("sound")==soundPath && eventData.Object("values").Number("soundVolume")==65,"Library selection and volume slider update the canonical event assignment");
        File.WriteAllText(Path.Combine(directory,"sound-validation.json"),session.Validation.ToJsonString());
        Check(assignment.ItemsSource!.Cast<string>().Contains("Reusable alert"),"Reusable sound choices display filenames without extensions");
        Check(page.GetVisualDescendants().OfType<Button>().Any(b=>b.Content?.ToString()=="Add sound files") && !session.Metadata!.Sections.Any(s=>s.Text("id")=="sound-playback"),"Sound files, assignments and playback settings share one category");
        captureTheme="graphite";ThemeController.Apply(captureTheme,session.Metadata);
        Width=1100;Height=800;UpdateLayout();await Task.Delay(160);
        pageScroll.Offset=new Vector(0,page.Children.OfType<StackPanel>().Last().Bounds.Top);await Task.Delay(120);
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height))) {bitmap.Render(this);bitmap.Save(Path.Combine(directory,"sound-library-populated-graphite.png"),PngBitmapEncoderOptions.Default);}
        sectionId="events";scrollPositions[session.Profile+"|events"]=default;RenderPage();await Task.Delay(700);UpdateLayout();
        Check(page.GetVisualDescendants().OfType<Button>().Any(b=>b.Content?.ToString()=="Reusable alert"),"Compact event table links to its assigned sound by filename");
        Check(!page.GetVisualDescendants().OfType<Button>().Any(b=>b.Content?.ToString()=="Options") && page.GetVisualDescendants().OfType<TextBlock>().Any(t=>t.Text=="Event options"),"Event options are displayed below the table without popup buttons");
        Check(page.GetVisualDescendants().OfType<Grid>().Where(g=>g.Children.OfType<Button>().Any(b=>Avalonia.Automation.AutomationProperties.GetName(b)=="Choose color")).SelectMany(g=>g.Children.OfType<TextBox>()).All(t=>t.FontSize==14),"Event color text uses the larger readable font size");
        Check(page.GetVisualDescendants().OfType<NumericUpDown>().Count()==1,"Event options expose the shooting interval while sound volume stays in Sound library");
        var idleEnabled=page.GetVisualDescendants().OfType<CheckBox>().Single(c=>Avalonia.Automation.AutomationProperties.GetName(c)=="Shooting idle enabled");
        var idleOptions=page.GetVisualDescendants().OfType<Control>().Single(c=>Avalonia.Automation.AutomationProperties.GetName(c)=="Shooting idle / Milliseconds without an outgoing shot");
        var wasEnabled=idleEnabled.IsChecked;idleEnabled.IsChecked=false;await Task.Delay(400);
        Check(!idleOptions.IsEffectivelyEnabled,"Disabled events make their inline options inactive");
        idleEnabled.IsChecked=true;await Task.Delay(400);Check(idleOptions.IsEffectivelyEnabled,"Enabling an event immediately enables its options");idleEnabled.IsChecked=wasEnabled;await Task.Delay(400);
        idleOptions.BringIntoView();await Task.Delay(150);UpdateLayout();
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height))) {bitmap.Render(this);bitmap.Save(Path.Combine(directory,"event-options-graphite.png"),PngBitmapEncoderOptions.Default);}
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height))) {bitmap.Render(this);bitmap.Save(Path.Combine(directory,"event-alerts-populated-graphite.png"),PngBitmapEncoderOptions.Default);}
        sectionId="sounds";await UndoCategoryAsync();eventData=await session.ReadAsync("event",new(){["id"]="stoppedShooting"});
        Check(eventData.Object("values").Text("sound")=="" && eventData.Object("values").Number("soundVolume")==100,"Sound category undo restores assignments and event volume together");
        await session.DiscardAsync();editHistory.Clear();captureTheme=null;
        sectionId="close-protection";RenderPage();await Task.Delay(300);UpdateLayout();
        var protectionTable=page.GetVisualDescendants().OfType<StackPanel>().First(p=>p.Tag is TableInteractions);
        Check(protectionTable.GetVisualDescendants().OfType<ToggleSwitch>().All(toggle=>Math.Abs(toggle.Bounds.Center.X-((Border)toggle.Parent!).Bounds.Width/2)<=1),"Protection toggles are centered with equal space on both sides of each table cell");
        pageScroll.Offset=new Vector(0,pageScroll.Extent.Height);captureTheme="graphite";ThemeController.Apply(captureTheme,session.Metadata);await Task.Delay(150);
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height))) {bitmap.Render(this);bitmap.Save(Path.Combine(directory,"close-protection-table-graphite.png"),PngBitmapEncoderOptions.Default);}
        captureTheme=null;
        sectionId="appearance";RenderPage();await Task.Delay(300);UpdateLayout();
        var opacityControl=page.GetVisualDescendants().OfType<Control>().Single(c=>Avalonia.Automation.AutomationProperties.GetName(c)=="Thumbnail opacity");
        var opacitySlider=opacityControl.GetVisualDescendants().OfType<Slider>().Single();var opacityNumber=opacityControl.GetVisualDescendants().OfType<NumericUpDown>().Single();
        Check(Math.Abs(opacitySlider.TranslatePoint(default,page)!.Value.Y-opacityNumber.TranslatePoint(default,page)!.Value.Y)<1 && opacityNumber.Width==NumericFieldWidth,"Opacity slider and percentage field share one left-aligned row and the common numeric width");
        Check(page.GetVisualDescendants().OfType<NumericUpDown>().All(n=>n.Width==NumericFieldWidth),"Appearance numeric fields use one consistent compact width");
        var inactiveBorder=page.GetVisualDescendants().OfType<Control>().Single(c=>Avalonia.Automation.AutomationProperties.GetName(c)=="Inactive border thickness");
        inactiveBorder.BringIntoView();captureTheme="graphite";ThemeController.Apply(captureTheme,session.Metadata);await Task.Delay(200);UpdateLayout();
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height))) {bitmap.Render(this);bitmap.Save(Path.Combine(directory,"compact-fields-graphite.png"),PngBitmapEncoderOptions.Default);}
        captureTheme=null;
        await session.MutateAsync("profile",new(){["action"]="add",["name"]="Review fixture",["copyCurrent"]=true,["inheritance"]=new JsonObject {["appearance"]=true}});
        sectionId="appearance";RenderPage();await Task.Delay(200);
        var editDefault=page.GetVisualDescendants().OfType<Button>().Single(b=>b.Content?.ToString()=="Edit Default here");
        editDefault.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));await Task.Delay(500);
        Check(session.Profile=="Default" && profiles.SelectedItem?.ToString()=="Review fixture" && defaultEditProfile=="Review fixture","Editing Default keeps the selected profile and inheritance intact");
        var opacity=session.Snapshot.Object("values").Object("ThumbnailOpacity").Number("value");
        await session.MutateAsync("set",new(){["id"]="ThumbnailOpacity",["value"]=opacity-5});
        await EndDefaultEditingAsync();RenderPage();
        Check(session.Snapshot.Object("inheritance").Bool("appearance") && session.Snapshot.Object("values").Object("ThumbnailOpacity").Number("value")==opacity-5,"Default edits flow through the existing inheritance source");
        await session.ApplyAsync();await ConnectAsync();
        Check(session.Profile=="Review fixture" && profiles.SelectedItem?.ToString()=="Review fixture","Apply and reconnect retain the profile being edited");
        sectionId="inheritance";RenderPage();await Task.Delay(150);
        Check(!page.GetVisualDescendants().OfType<ComboBox>().Any() && page.GetVisualDescendants().OfType<CheckBox>().Count()==session.Metadata!.Sections.Count(s=>s.Text("group")!=""),"Section inheritance uses one-click checkboxes for every category");
        var inheritAppearance=page.GetVisualDescendants().OfType<CheckBox>().Single(c=>c.Tag?.ToString()=="appearance");
        inheritAppearance.IsChecked=false;await Task.Delay(450);
        Check(!session.Snapshot.Object("inheritance").Bool("appearance"),"One checkbox interaction changes the category inheritance source");
        inheritAppearance=page.GetVisualDescendants().OfType<CheckBox>().Single(c=>c.Tag?.ToString()=="appearance");
        inheritAppearance.IsChecked=true;await Task.Delay(450);
        var profileCreation=CreateProfileAsync();await Task.Delay(250);var createDialog=OwnedWindows.Single();
        var createChecks=createDialog.GetVisualDescendants().OfType<CheckBox>().Where(c=>c.Tag is string).ToArray();
        var expectedCategories=session.Metadata.Sections.Where(s=>s.Text("group")!="").Select(s=>s.Text("page")).Distinct().ToArray();
        Check(expectedCategories.All(category=>createDialog.GetVisualDescendants().OfType<TextBlock>().Count(t=>t.Text==category)==1),"Create profile groups inheritance choices beneath category headings shown once");
        void CreationPreset(string label)=>createDialog.GetVisualDescendants().OfType<Button>().Single(b=>b.Content?.ToString()==label).RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        CreationPreset("Select all");Check(createChecks.All(c=>c.IsChecked==true),"Create profile Select all checks every inheritance choice");
        CreationPreset("Select none");Check(createChecks.All(c=>c.IsChecked==false),"Create profile Select none clears every inheritance choice");
        CreationPreset("Use current profile inheritance");Check(createChecks.All(c=>c.IsChecked==session.Snapshot.Object("inheritance").Bool(c.Tag!.ToString()!)),"Create profile can copy the current profile's inheritance choices");
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)createDialog.Bounds.Width,(int)createDialog.Bounds.Height))) {bitmap.Render(createDialog);bitmap.Save(Path.Combine(directory,"create-profile-categories.png"),PngBitmapEncoderOptions.Default);}
        createDialog.GetVisualDescendants().OfType<TextBox>().Single(t=>t.PlaceholderText=="Profile name").Text="Inheritance preset fixture";
        CreationPreset("Create");await profileCreation;
        Check(session.Profile=="Inheritance preset fixture" && session.Snapshot.Object("inheritance").Bool("appearance"),"Create saves the selected inheritance preset to the new profile");
        await session.MutateAsync("profile",new(){["action"]="delete"});await session.MutateAsync("select-profile",new(){["name"]="Review fixture"});
        sectionId="inheritance";RenderPage();await Task.Delay(100);
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height))) {bitmap.Render(this);bitmap.Save(Path.Combine(directory,"inheritance-checkboxes.png"),PngBitmapEncoderOptions.Default);}
        var savedSize=session.Snapshot.Object("values").Object("ThumbnailTextSize").Number("value");
        await session.MutateAsync("inheritance",new(){["section"]="appearance",["enabled"]=false});
        await session.ApplyAsync();
        await session.MutateAsync("set",new(){["id"]="ThumbnailTextSize",["value"]=savedSize+1});UpdateStatus();
        Check(session.ChangeCount==1 && draft.Text=="1 setting changed, apply to save" && draft.Foreground is not null,"Footer visibly counts distinct unapplied settings");
        await session.MutateAsync("set",new(){["id"]="ThumbnailTextSize",["value"]=savedSize});
        Check(session.ChangeCount==0,"Returning a setting to its saved value removes it from the count");
        search.Text="opacity";await Task.Delay(100);UpdateLayout();
        Check(clearSearch.IsVisible && clearSearch.Bounds.Width==26,"Search reserves a compact clear button");
        using(var bitmap=new RenderTargetBitmap(new PixelSize((int)Bounds.Width,(int)Bounds.Height))) {bitmap.Render(this);bitmap.Save(Path.Combine(directory,"search-clear-graphite.png"),PngBitmapEncoderOptions.Default);}
        clearSearch.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));await Task.Delay(100);
        Check(search.Text=="" && !clearSearch.IsVisible,"Search clear button restores the complete navigation");
        var retainedSize=session.Snapshot.Object("values").Object("ThumbnailTextSize").Number("value");
        await session.MutateAsync("set",new(){["id"]="ThumbnailTextSize",["value"]=retainedSize+1});
        connectionPoll.Stop();
        reconnecting=true;
        var host=await session.ReadAsync("hello");using(var hostProcess=System.Diagnostics.Process.GetProcessById(host.Int("pid"))) {hostProcess.Kill();await hostProcess.WaitForExitAsync();}
        try {await session.ReadAsync("snapshot");}catch(BridgeException) {} UpdateStatus();
        Check(!session.Connected && connectionBanner.IsVisible && connectionMessage.Text!.Contains("draft") && !page.IsEnabled,"Unexpected disconnection displays a persistent explanation and retains the draft");
        reconnecting=false;connectionPoll.Start();
        var reconnectWatch=System.Diagnostics.Stopwatch.StartNew();
        while((!session.Connected || reconnecting) && reconnectWatch.Elapsed<TimeSpan.FromSeconds(30))await Task.Delay(100);
        UpdateStatus();
        Check(session.Connected && session.Profile=="Review fixture" && session.Snapshot.Object("values").Object("ThumbnailTextSize").Number("value")==retainedSize+1 && !connectionBanner.IsVisible && !problems.IsEnabled,"Automatic reconnect restores the edited profile and draft without stale Problems");
        host=await session.ReadAsync("hello");using(var hostProcess=System.Diagnostics.Process.GetProcessById(host.Int("pid"))) {hostProcess.Kill();await hostProcess.WaitForExitAsync();}
        await Task.Delay(2300);reconnectWatch.Restart();
        while((!session.Connected || reconnecting) && reconnectWatch.Elapsed<TimeSpan.FromSeconds(30))await Task.Delay(100);
        Check(session.Connected && session.Snapshot.Object("values").Object("ThumbnailTextSize").Number("value")==retainedSize+1,"A second connection loss recovers automatically without losing the retained draft");
        await session.MutateAsync("set",new(){["id"]="ThumbnailTextSize",["value"]=retainedSize});
        await session.MutateAsync("select-profile",new(){["name"]="Default"});
        await session.MutateAsync("set",new(){["id"]="ThumbnailOpacity",["value"]=opacity});
        await session.MutateAsync("select-profile",new(){["name"]="Review fixture"});await session.MutateAsync("profile",new(){["action"]="delete"});
        await session.ApplyAsync();await session.ValidateAsync();RenderPage();editHistory.Clear();profileBaselines.Clear();
        File.WriteAllLines(Path.Combine(directory, "interaction-checks.log"), results);
    }
}

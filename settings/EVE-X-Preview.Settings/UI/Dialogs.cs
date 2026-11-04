using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace EveX.Settings.UI;

internal static class Dialogs
{
    public static async Task<bool> ConfirmAsync(Window owner, string title, string message) => await ChoiceAsync(owner, title, new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap }, [("Continue", "yes"), ("Cancel", "cancel")]) == "yes";
    public static async Task MessageAsync(Window owner, string title, string message) => await ChoiceAsync(owner, title, new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap, MaxWidth = 650 }, [("Close", "close")]);
    public static async Task<string> CloseChoiceAsync(Window owner, bool pending) => await ChoiceAsync(owner, "Unsaved settings", new TextBlock { Text = pending ? "Some edits have not been accepted. Keep editing to correct them, or discard the draft and close." : "Apply your settings before closing, discard the draft, or keep editing.", TextWrapping = TextWrapping.Wrap }, pending ? [("Keep editing", "cancel"), ("Discard and close", "discard")] : [("Apply and close", "apply"), ("Discard and close", "discard"), ("Keep editing", "cancel")]);
    public static async Task<string?> NameAsync(Window owner, string title, string initial = "")
    {
        var text = new TextBox { Text = initial, PlaceholderText = "Name", MinWidth = 350 };
        var result = await ChoiceAsync(owner, title, text, [("Save to draft", "save"), ("Cancel", "cancel")]);
        return result == "save" ? text.Text ?? "" : null;
    }
    public static async Task NameActionAsync(Window owner, string title, string initial, Func<string, Task> save)
    {
        var text = new TextBox { Text = initial, PlaceholderText = "Name", MinWidth = 350 };
        await ChoiceAsync(owner, title, text, [("Save to draft", "save"), ("Cancel", "cancel")], async choice => { if (choice == "save") await save(text.Text ?? ""); });
    }
    public static async Task<string> ChoiceAsync(Window owner, string title, Control content, (string Label, string Value)[] choices, Func<string, Task>? onChoice = null)
    {
        if (owner is MainWindow modern) await modern.ProtectModalAsync(true);
        try
        {
        var dialog = new Window { Title = title, Width = 560, MaxHeight = 750, MinWidth = 400, SizeToContent = SizeToContent.Height, WindowStartupLocation = WindowStartupLocation.CenterOwner, CanResize = true, ShowInTaskbar = false };
        var panel = new StackPanel { Spacing = 20, Margin = new Thickness(24) };
        panel.Children.Add(new TextBlock { Text = title, FontSize = 22, FontWeight = FontWeight.SemiBold, TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(content);
        var feedback = new TextBlock { Classes = { "error" }, IsVisible = false }; panel.Children.Add(feedback);
        var buttons = new WrapPanel { HorizontalAlignment = HorizontalAlignment.Right };
        foreach (var choice in choices)
        {
            var button = new Button { Content = choice.Label, Margin = new Thickness(6, 6, 0, 0) };
            button.Click += async (_, _) =>
            {
                buttons.IsEnabled = false;
                try { if (onChoice is not null) await onChoice(choice.Value); dialog.Close(choice.Value); }
                catch (Exception error) { feedback.Text = error.Message; feedback.IsVisible = true; }
                finally { buttons.IsEnabled = true; }
            }; buttons.Children.Add(button);
        }
        panel.Children.Add(buttons); dialog.Content = new ScrollViewer { Content = panel };
        return await dialog.ShowDialog<string?>(owner) ?? "cancel";
        }
        finally { if (owner is MainWindow modernOwner) await modernOwner.ProtectModalAsync(false); }
    }
}

using Avalonia.Controls;
using EveX.Settings.Bridge;
using System.Text.Json.Nodes;

namespace EveX.Settings.UI;

public sealed partial class MainWindow
{
    private Button? previewPauseButton;

    private async Task RefreshPreviewStatusAsync()
    {
        if(!session.Connected || session.Role=="offline")return;
        try {
            var status=await session.ReadAsync("command",new(){["name"]="live-preview-status"});
            if(sectionId=="diagnostics")SetPreviewStatus(status);
        } catch(Exception error) {System.Diagnostics.Trace.WriteLine(error.Message);}
    }
    private void SetPreviewStatus(JsonObject status)
    {
        if(previewPauseButton is null)return;
        previewPauseButton.Content=status.Bool("paused")?"Resume live previews":"Pause live previews";
        ToolTip.SetTip(previewPauseButton,status.Text("message"));
    }
    private void RenderDependencySummary(JsonObject section)
    {
        var summary=Muted("");page.Children.Add(summary);
        void Update() {
            var inactive=session.Metadata!.SectionFields(section).Where(field=>field.Text("dependency")!="" &&
                !session.Snapshot.Object("values").Object(field.Text("id")).Bool("inherited") &&
                !session.Snapshot.Object("values").Object(field.Text("id")).Bool("enabled",true));
            var explanations=inactive.GroupBy(field=>field.Text("dependency")).Select(group=>
                "Enable “"+session.Metadata.Field(group.Key).Text("label",group.Key)+"” to edit "+string.Join(", ",group.Select(f=>f.Text("label")))+".");
            summary.Text=string.Join("\n",explanations);summary.IsVisible=summary.Text.Length>0;
        }
        fieldFeedback["dependencies"]=Update;Update();
    }
}

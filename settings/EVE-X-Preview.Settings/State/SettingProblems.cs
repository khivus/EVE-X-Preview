using System.Text.Json.Nodes;
using EveX.Settings.Bridge;
namespace EveX.Settings.State;
public sealed record SettingProblem(string Severity,string Profile,string Section,string Id,string Label,string Message,bool Shortcut=false);
public static class SettingProblems
{
    public static IReadOnlyList<SettingProblem> Build(IEnumerable<PendingEdit> pending,JsonObject validation,Metadata? metadata) {
        var result=new List<SettingProblem>();var shortcuts=new List<SettingProblem>();
        foreach(var edit in pending.Where(p=>p.Error is not null))result.Add(new("error",edit.Profile,edit.Key.Split('|').ElementAtOrDefault(1)??"",edit.Arguments.Text("id",edit.Arguments.Text("kind")),edit.Label,edit.Error!));
        foreach(var issue in validation.Array("issues")) {
            if(issue.Text("kind")=="shortcut-conflict" || issue.Text("message").StartsWith("Shortcut overlaps:",StringComparison.Ordinal)) {
                var shortcutSection=issue.Text("section",ShortcutSection(issue.Text("id")));
                shortcuts.Add(new("warning",issue.Text("profile"),shortcutSection,issue.Text("id"),"Shortcut conflicts — saving is allowed",issue.Text("message"),true));continue;
            }
            var id=issue.Text("id");var field=metadata?.Field(id);var section=field?.Text("section")??metadata?.Tables.FirstOrDefault(t=>t.Text("kind")==id).Text("section")??"";
            if(section=="" && metadata?.Events.Any(e=>id==e.Text("id") || id.StartsWith(e.Text("id")+":"))==true)section="events";
            var descriptor=metadata?.Events.FirstOrDefault(e=>id==e.Text("id") || id.StartsWith(e.Text("id")+":"));
            if(descriptor is not null && (id.EndsWith(":sound") || id.EndsWith(":soundVolume")))section="sounds";
            var label=field?.Text("label")??(descriptor is null?id:descriptor.Text("label")+" / "+id.Split(':').Last());
            result.Add(new(issue.Text("severity","error"),issue.Text("profile"),section,id,label,issue.Text("message")));
        }
        foreach(var conflict in validation.Array("conflicts"))result.Add(new("error","","profiles","","Concurrent settings change",JsonExtensions.Display(conflict)));
        var ordered=result.DistinctBy(p=>(p.Profile,p.Id,p.Message)).OrderBy(p=>p.Severity=="error"?0:1).ToList();
        foreach(var group in shortcuts.GroupBy(p=>(p.Profile,p.Section)))
            ordered.Add(group.First() with {Message=string.Join("\n",group.Select(p=>p.Message).Distinct())+"\nYou can save these settings as they are. Overlapping shortcuts can trigger different actions depending on context; their behavior is not guaranteed."});
        return ordered;
    }
    private static string ShortcutSection(string id) => id.Split(':')[0] switch {
        "character"=>"character-shortcuts", "group"=>"character-groups", "nonGroup"=>"app-groups", "nonApp"=>"app-shortcuts", "profile"=>"profiles", "quick"=>"quick-group", _=>"shortcuts"
    };
}

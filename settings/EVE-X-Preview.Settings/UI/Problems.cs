using Avalonia.VisualTree;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using EveX.Settings.Bridge;
using EveX.Settings.State;
using System.Text.Json.Nodes;
namespace EveX.Settings.UI;
public sealed partial class MainWindow
{
    private readonly Button problems = new(){Content="Problems"};
    private IReadOnlyList<SettingProblem> CurrentProblems()=>SettingProblems.Build(session.Pending,session.Validation,session.Metadata);
    private async Task ShowProblemsAsync() {
        await session.ValidateAsync();var issues=CurrentProblems();
        if(issues.Count==0)return;
        var panel=new StackPanel {Spacing=12};var target=-1;
        for(var i=0;i<issues.Count;i++) {
            var issue=issues[i];var index=i;var row=new StackPanel {Spacing=5};
            row.Children.Add(new TextBlock {Text=(issue.Profile==""?"":issue.Profile+" / ")+issue.Label,FontWeight=FontWeight.SemiBold,TextWrapping=TextWrapping.Wrap});
            row.Children.Add(new TextBlock {Text=issue.Message,TextWrapping=TextWrapping.Wrap,Foreground=issue.Severity=="error"?Brushes.OrangeRed:Brushes.Goldenrod});
            if(issue.Section!="") {var go=new Button {Content="Go to setting"};go.Click+=(_,_)=>{target=index;OwnedWindows.LastOrDefault()?.Close("go");};row.Children.Add(go);}
            panel.Children.Add(Card(row));
        }
        await Dialogs.ChoiceAsync(this,"Settings problems",panel,[("Close","close")]);
        if(target>=0) {
            var issue=issues[target];if(issue.Profile!="" && issue.Profile!=session.Profile)await session.MutateAsync("select-profile",new(){["name"]=issue.Profile});
            search.Text="";SaveScroll();sectionId=issue.Section;scrollPositions[ScrollKey]=default;RenderNavigation();RenderPage();
            await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(()=>{},Avalonia.Threading.DispatcherPriority.Loaded);
            var key=issue.Id;
            if(key!="") {
                var control=page.GetVisualDescendants().OfType<Control>().FirstOrDefault(c=>Avalonia.Automation.AutomationProperties.GetName(c)==session.Metadata?.Field(key).Text("label"));
                control?.BringIntoView();control?.Focus();
            }
        }
    }
}

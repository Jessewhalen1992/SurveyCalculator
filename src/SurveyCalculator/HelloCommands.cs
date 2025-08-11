using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(SurveyCalculator.HelloCommands))]

namespace SurveyCalculator
{
    public class HelloCommands
    {
        [CommandMethod("HELLOALS")]
        public void HelloAls()
        {
            var ed = Application.DocumentManager.MdiActiveDocument?.Editor;
            ed?.WriteMessage("\nSurveyCalculator: hello! If you can read this, the DLL loaded fine.");
            ed?.WriteMessage("\nNext step: add PLANLINK, PLANEXTRACT, EVILINK, ALSADJ.");
        }
    }
}

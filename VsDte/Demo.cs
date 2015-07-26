using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Codeer.Friendly;
using Codeer.Friendly.Dynamic;
using Codeer.Friendly.Windows;
using EnvDTE;
using VSHTC.Friendly.PinInterface;
using EnvDTE80;
using System.IO;
using Codeer.Friendly.DotNetExecutor;

namespace VsDte
{
    [TestClass]
    public class Demo
    {
        string SolutionDir = Path.GetFullPath("TestDir");
        System.Diagnostics.Process vsProcess;

        [TestInitialize]
        public void TestInitialize()
        {
            //ソリューション作成用のディレクトリを作成
            while (Directory.Exists(SolutionDir))
            {
                try
                {
                    Directory.Delete(SolutionDir, true);
                    break;
                }
                catch { }
                System.Threading.Thread.Sleep(10);
            }
            Directory.CreateDirectory(SolutionDir);

            //VS起動
            //パスを書き換えると2013も操作できるよ。
            var path = @"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\devenv.exe";
            vsProcess = System.Diagnostics.Process.Start(path);
            while (0 == vsProcess.MainWindowTitle.Length) 
            {
                System.Threading.Thread.Sleep(10);
                vsProcess = System.Diagnostics.Process.GetProcessById(vsProcess.Id);
            }
        }

        [TestCleanup]
        public void TestCleanup() 
        {
            vsProcess.Kill();
        }

        static Type DTEType { get { return typeof(_DTE); } }

        [TestMethod]
        public void Test()
        {
            //アタッチ
            WindowsAppFriend app = new WindowsAppFriend(vsProcess);

            //DLLインジェクション
            WindowsAppExpander.LoadAssembly(app, GetType().Assembly);

            //Microsoft.VisualStudio.Shell.Package.GetGlobalServiceの呼び出し
            var dteType = app.Type(GetType()).DTEType;
            AppVar obj = app.Type().Microsoft.VisualStudio.Shell.Package.GetGlobalService(dteType);

            //注目！
            //インターフェイスでプロキシが作成できる！
            var dte = obj.Pin<DTE2>();
            var solution = dte.Solution;

            //ソリューション作成
            solution.Create(SolutionDir, "Test.sln");
            string solutionPath = Path.Combine(SolutionDir, "Test.sln");
            solution.SaveAs(solutionPath);

            //プロジェクト追加
            solution.AddFromTemplate(
                @"C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\ProjectTemplates\CSharp\Windows\1041\WPFApplication\csWPFApplication.vstemplate",
                Path.Combine(SolutionDir, "WPF"), "WPF", true);

            //保存
            solution.SaveAs(solutionPath);

            //閉じる
            solution.Close();

            //再度開く
            solution.Open(solutionPath);

            //クリーン→ビルド
            solution.SolutionBuild.Clean(true);
            solution.SolutionBuild.Build(true);

            //デバッグ
            solution.SolutionBuild.Debug();

            //デバッグ対象プロセスを操作
            var debProcess = System.Diagnostics.Process.GetProcessById(dte.Debugger.DebuggedProcesses.Item(1).ProcessID);
            using (var debApp = new WindowsAppFriend(debProcess)) 
            {
                debApp.Type().System.Windows.Application.Current.MainWindow.Close(new Async());
            }

            //編集モードに戻るまで待つ
            while (dte.Debugger.CurrentMode != dbgDebugMode.dbgDesignMode) 
            {
                System.Threading.Thread.Sleep(10);
            }
        }
    }
}

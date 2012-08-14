using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Web.Application;
using Microsoft.Web.Administration;
using Application = Microsoft.Web.Administration.Application;
using Binding = Microsoft.Web.Administration.Binding;

namespace Whut.AttachTo
{
    //// This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    //// This attribute is used to register the informations needed to show the this package in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    //// This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidAttachToPkgString)]
    [ProvideOptionPage(typeof(GeneralOptionsPage), "Whut.AttachTo", "General", 110, 120, false)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    public sealed class AttachToPackage : Package
    {
        protected override void Initialize()
        {
            base.Initialize();

            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

            AddAttachToCommand(mcs, PkgCmdIDList.cmdidWhutAttachToIIS, gop => gop.ShowAttachToIIS, "w3wp.exe");
            AddAttachToCommand(mcs, PkgCmdIDList.cmdidWhutAttachToIISExpress, gop => gop.ShowAttachToIISExpress, "iisexpress.exe");
            AddAttachToCommand(mcs, PkgCmdIDList.cmdidWhutAttachToNUnit, gop => gop.ShowAttachToNUnit, "nunit-agent.exe", "nunit.exe", "nunit-console.exe", "nunit-agent-x86.exe", "nunit-x86.exe", "nunit-console-x86.exe");
        }

        private void AddAttachToCommand(OleMenuCommandService mcs, uint commandId, Func<GeneralOptionsPage, bool> isVisible, params string[] programsToAttach)
        {
            var menuItemCommand = new OleMenuCommand((sender, args) => Attach(programsToAttach),
                new CommandID(GuidList.guidAttachToCmdSet, (int)commandId));

            menuItemCommand.BeforeQueryStatus += (s, e) => menuItemCommand.Visible = isVisible((GeneralOptionsPage)GetDialogPage(typeof(GeneralOptionsPage)));
            mcs.AddCommand(menuItemCommand);
        }

        private void Attach(string[] programsToAttach)
        {
            if (programsToAttach[0] == "w3wp.exe")
            {
                AttachToIIS();
            } 
            else
            {
                var dte = (DTE)GetService(typeof(DTE));
                foreach (Process process in dte.Debugger.LocalProcesses)
                {
                    if (programsToAttach.Any(p => process.Name.EndsWith(p)))
                    {
                        process.Attach();
                    }
                }
            }
        }

        private void AttachToIIS()
        {
            var dte = (DTE)GetService(typeof(DTE));

            foreach (Project project in dte.Solution.Projects)
            {
                foreach (object item in (Array)project.ExtenderNames)
                {
                    var extend = project.Extender[item.ToString()] as WAProjectExtender;
                    if (extend == null)
                    {
                        continue;
                    }

                    using (var serverManager = new ServerManager())
                    {
                        var uri = new Uri(extend.IISUrl);
                        var site = serverManager.Sites.FirstOrDefault(x => x.Bindings.Any(y =>
                                                                                          EqualHosts(y, uri) &&
                                                                                          y.EndPoint.Port == uri.Port &&
                                                                                          y.Protocol == uri.Scheme));

                        if (site == null)
                        {
                            continue;
                        }

                        if (site.Applications.Count > 1)
                        {
                            MessageBox.Show("Multiple applications in single site are not supported, attaching to first application pool");
                        }

                        Application application = site.Applications.FirstOrDefault();

                        if (application == null)
                        {
                            MessageBox.Show("Wrong IIS configuration (site does not have any applications)");
                            return;
                        }

                        var pool = serverManager.ApplicationPools[application.ApplicationPoolName];
                        foreach (var workerProcess in pool.WorkerProcesses)
                        {
                            foreach (Process process in dte.Debugger.LocalProcesses)
                            {
                                if (process.ProcessID == workerProcess.ProcessId)
                                {
                                    process.Attach();
                                    return;
                                }
                            }
                        }

                        MessageBox.Show("Could not find process to attach to, it is not started?");
                    }
                }
            }
        }

        private static bool EqualHosts(Binding binding, Uri uri)
        {
            return binding.Host == uri.Host || (binding.Host == "" && uri.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase));
        }
    }
}

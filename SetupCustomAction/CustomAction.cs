using System;
using System.Collections;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace SetupCustomAction
{
    [System.ComponentModel.RunInstaller(true)]
    public sealed class CustomAction : Installer
    {
        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            var outlookProcess = Process.GetProcessesByName("OUTLOOK");
            if (outlookProcess.Length > 0)
            {
                throw new InstallException();
            }

        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

        }

        public override void Commit(IDictionary savedState)
        {
        }

        public override void Rollback(IDictionary savedState)
        {
        }
    }
}

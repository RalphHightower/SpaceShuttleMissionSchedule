//-----------------------------------------------------------------------
// 
//  Copyright (C) Microsoft Corporation.  All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY
// KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//-----------------------------------------------------------------------

using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace CustomActions
{
    [RunInstaller(true)]
    [System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public sealed partial class UpdateManifest : Installer
    {
        public UpdateManifest()
        {
            InitializeComponent();
        }

        // Override Install to update the customization location
        // in the application manifest.
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            // Call the base implementation.
            base.Install(stateSaver);

            UpdateApplicationManifest();
        }

        private void UpdateApplicationManifest()
        {
            // Define the parameters passed to the task.
            string targetDir = this.Context.Parameters["targetDir"];
            string documentName = this.Context.Parameters["documentName"];
            string assemblyName = this.Context.Parameters["assemblyName"];

            if (String.IsNullOrEmpty(targetDir))
                throw new InstallException("Cannot update the application manifest. The specified target directory name is not valid.");
            if (String.IsNullOrEmpty(documentName))
                throw new InstallException("Cannot update the application manifest. The specified document name is not valid.");
            if (String.IsNullOrEmpty(assemblyName))
                throw new InstallException("Cannot update the application manifest. The specified assembly name is not valid.");

            // Get the application manifest from the document.
            string documentPath = Path.Combine(targetDir, documentName);
            ServerDocument serverDocument = new ServerDocument(documentPath, FileAccess.ReadWrite);
            try
            {
                AppManifest appManifest = serverDocument.AppManifest;

                string assemblyPath = Path.Combine(targetDir, assemblyName);
                appManifest.Dependency.AssemblyPath = assemblyPath;

                serverDocument.Save();
            }
            finally
            {
                if (serverDocument != null)
                {
                    serverDocument.Close();
                }
            }
        }
    }
}
/*
    NasaTvScheduleImport.  This program reads the NASA TV Schedule in Excel
    Format for the Space Shuttle and transfers the entries into Microsoft
    Outlook Calendar as Appointment items.

    Copyright (C) 2007-2011  Ralph M. Hightower, Jr (Permanent Vacations)

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

    mailto:Ralph.Hightower@gmail.com
 * Ralph Hightower
 * Chapin, SC 29036
*/
/* Change History
 * 20110429 Updated Copyright period; added Technical Article Link
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;

namespace PermanentVacations.Nasa.Sts.OutlookCalendar
{
	partial class AboutBox : Form
	{
		public AboutBox()
		{
			InitializeComponent();

			//  Initialize the AboutBox to display the product information from the assembly information.
			//  Change assembly information settings for your application through either:
			//  - Project->Properties->Application->Assembly Information
			//  - AssemblyInfo.cs
			this.Text = String.Format("About {0}", AssemblyTitle);
			this.labelProductName.Text = AssemblyProduct;
			this.labelVersion.Text = String.Format("Version {0}", AssemblyVersion);
			this.labelCopyright.Text = AssemblyCopyright;
			this.labelCompanyName.Text = AssemblyCompany;
			this.textBoxDescription.Text = AssemblyDescription;
		}

		#region Assembly Attribute Accessors

		public string AssemblyTitle
		{
			get
			{
				// Get all Title attributes on this assembly
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
				// If there is at least one Title attribute
				if (attributes.Length > 0)
				{
					// Select the first one
					AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
					// If it is not an empty string, return it
					if (titleAttribute.Title != "")
						return titleAttribute.Title;
				}
				// If there was no Title attribute, or if the Title attribute was the empty string, return the .exe name
				return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
			}
		}

		public string AssemblyVersion
		{
			get
			{
				return Assembly.GetExecutingAssembly().GetName().Version.ToString();
			}
		}

		public string AssemblyDescription
		{
			get
			{
				// Get all Description attributes on this assembly
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
				// If there aren't any Description attributes, return an empty string
				if (attributes.Length == 0)
					return "";
				// If there is a Description attribute, return its value
				return ((AssemblyDescriptionAttribute)attributes[0]).Description;
			}
		}

		public string AssemblyProduct
		{
			get
			{
				// Get all Product attributes on this assembly
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
				// If there aren't any Product attributes, return an empty string
				if (attributes.Length == 0)
					return "";
				// If there is a Product attribute, return its value
				return ((AssemblyProductAttribute)attributes[0]).Product;
			}
		}

		public string AssemblyCopyright
		{
			get
			{
				// Get all Copyright attributes on this assembly
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
				// If there aren't any Copyright attributes, return an empty string
				if (attributes.Length == 0)
					return "";
				// If there is a Copyright attribute, return its value
				return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
			}
		}

		public string AssemblyCompany
		{
			get
			{
				// Get all Company attributes on this assembly
				object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
				// If there aren't any Company attributes, return an empty string
				if (attributes.Length == 0)
					return "";
				// If there is a Company attribute, return its value
				return ((AssemblyCompanyAttribute)attributes[0]).Company;
			}
		}
		#endregion

		private void okButton_Click(object sender, EventArgs e)
		{

		}

        /// <summary>
        /// Opens web browser to CodePlex page for NASA STS TV Schedule application
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void urlProjectPage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            urlProjectPage.Links[urlProjectPage.Links.IndexOf(e.Link)].Visited = true;

            System.Diagnostics.Process.Start(Properties.Resources.URL_CODEPLEX);
        }

        /// <summary>
        /// Opens web browser to the Technical Article page on Code Project about the development process and reason for development
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void urlTechArticle_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            urlTechArticle.Links[urlTechArticle.Links.IndexOf(e.Link)].Visited = true;

            System.Diagnostics.Process.Start(Properties.Resources.URL_CODEPROJECT);

        }

        /// <summary>
        /// Closes About Box when Escape is hit (haven't entered this path during testing)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AboutBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (e.KeyChar)
            {
                case (char)27:  //  Escape key '\e'
                    Close();
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Closes About Box when Escape is hit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void okButton_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (e.KeyChar)
            {
                case (char)27:  //  Escape key '\e'
                    Close();
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Closes About Box when Escape is hit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxDescription_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (e.KeyChar)
            {
                case (char)27:  //  Escape key '\e'
                    Close();
                    break;
                default:
                    break;
            }

        }
	}
}

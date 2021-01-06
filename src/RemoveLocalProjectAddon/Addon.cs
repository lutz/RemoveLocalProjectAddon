using RemoveLocalProject.Properties;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Settings;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace RemoveLocalProject
{
    public class Addon : CitaviAddOnEx<ProjectDetailsControl>
    {
        public override void OnHostingFormLoaded(ProjectDetailsControl projectDetailsControl)
        {
            var WorkingKnownProject = projectDetailsControl
                                        .GetType()
                                        .GetProperty("WorkingKnownProject", BindingFlags.Instance | BindingFlags.NonPublic)?
                                        .GetValue(projectDetailsControl) as KnownProject;

            if (WorkingKnownProject == null || WorkingKnownProject.ProjectType != ProjectType.DesktopSQLite) return;

            var deleteProjectLabel = projectDetailsControl
                                        .Controls[0]
                                        .Controls.OfType<SwissAcademic.Controls.Label>()
                                        .FirstOrDefault(label => label.Name.Equals("deleteProjectLabel", StringComparison.OrdinalIgnoreCase));

            if (deleteProjectLabel == null) return;

            var deleteProjectLabelForDesktopSQLite = new SwissAcademic.Controls.Label
            {
                Name = "deleteProjectLabelForDesktopSQLite",
                Location = deleteProjectLabel.Location,
                Text = deleteProjectLabel.Text,
                Style = SwissAcademic.Controls.LabelStyle.StandardLink,
                TabIndex = 2,
                TabStop = true
            };

            deleteProjectLabelForDesktopSQLite.LinkClicked += DeleteProjectLabel_LinkClicked;
            projectDetailsControl.Controls[0].Controls.Add(deleteProjectLabelForDesktopSQLite);
        }

        private void DeleteProjectLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (sender is SwissAcademic.Controls.Label label && label.Parent?.Parent is ProjectDetailsControl projectDetailsControl)
            {
                var WorkingKnownProject = projectDetailsControl
                                      .GetType()
                                      .GetProperty("WorkingKnownProject", BindingFlags.Instance | BindingFlags.NonPublic)?
                                      .GetValue(projectDetailsControl) as KnownProject;

                using (var input = new InputBox(projectDetailsControl))
                {
                    input.Title = Resources.DeleteProject_Title;
                    input.Description = string.Format(Resources.DeleteProject, WorkingKnownProject.Name);
                    input.InputLabel = string.Empty;
                    var hnd = new InputBoxValidateEventHandler((s) =>
                    {
                        return s == WorkingKnownProject.Name;
                    });
                    if (input.ShowDialog(projectDetailsControl, null, hnd) != DialogResult.OK)
                    {
                        return;
                    }
                }

                try
                {
                    if (Directory.Exists(WorkingKnownProject.AttachmentsPath))
                    {
                        Directory.Delete(WorkingKnownProject.AttachmentsPath, true);
                    }

                    if (File.Exists(WorkingKnownProject.ConnectionString))
                    {
                        var directory = Path.GetDirectoryName(WorkingKnownProject.ConnectionString);

                        if (Directory.Exists(directory))
                        {
                            Directory.Delete(directory, true);
                        }
                    }
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.ToString());
                }

                if (projectDetailsControl.Owner is StartForm startForm)
                {
                    Program.Engine.Settings.General.KnownProjects.Validate();
                    startForm.GetType().GetMethod("RefreshKnownProjectsPanel", BindingFlags.NonPublic | BindingFlags.Instance)?.Invoke(startForm, new object[] { });
                }

                projectDetailsControl.Close();
            }
        }
    }
}
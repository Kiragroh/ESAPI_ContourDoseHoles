using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Controls;
using System.ComponentModel;

[assembly: AssemblyVersion("1.2.2")]
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        // xChange these IDs to match your clinical conventions
        const string SCRIPT_NAME = "ContourDoseHoles";
        int HoleCount = 0;

        public Script()
        {
        }

        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {

            // get list of structures for loaded plan
            PlanSetup planSetup = context.PlanSetup;
            PlanSum psum = context.PlanSumsInScope.FirstOrDefault();
            StructureSet ss;
            //bool isPlanSum = false;
            string psumName = null;
            double rx95_5;
            Structure Iso95;
            context.Patient.BeginModifications();

            // If there's no selected plan with calculated dose throw an exception
            if (planSetup == null && psum == null)
            {
                throw new ApplicationException("Please open a calculated plan/planSum before using this script.");
            }
            else if (planSetup != null)
            {
                if (planSetup.Dose == null)
                    throw new ApplicationException("Please open a calculated plan before using this script.");

                ss = planSetup.StructureSet;
                Iso95 = context.StructureSet.AddStructure("DOSE_REGION", "zz95_ISO");  // second parameter is structure ID
                Iso95.ConvertDoseLevelToStructure(context.PlanSetup.Dose, new DoseValue(95.5, DoseValue.DoseUnit.Percent));
            }
            // in this simple script the planSum dose is a sum of the included planSetups. 
            else
            {
                if (psum.Dose == null)
                    throw new ApplicationException("Please open a calculated plansum before using this script.");

                //isPlanSum = true;

                if (context.PlanSumsInScope.Count() > 1)
                {
                    List<String> psumList = new List<String>();
                    foreach (var sum in context.PlanSumsInScope)
                    {
                        foreach (var course in context.Patient.Courses)
                        {
                            if (course.PlanSums.FirstOrDefault(ps => ps.Id == sum.Id) != null)
                            {
                                if (sum.Equals(course.PlanSums.First(ps => ps.Id == sum.Id)))
                                {
                                    psumList.Add(String.Format("{0}/{1}", course.Id, sum.Id));
                                }
                            }
                        }
                    }
                    psumName = Contour_95holes.Form5.ShowMiniForm(psumList);
                    ss = context.Patient.Courses.First(c => c.Id == psumName.Split('/')[0]).PlanSums.First(ps => ps.Id == psumName.Split('/')[1]).StructureSet;
                    List<double> rx_doses = new List<double>();
                    foreach (PlanSetup ps in (context.Patient.Courses.First(c => c.Id == psumName.Split('/')[0]).PlanSums.First(ps => ps.Id == psumName.Split('/')[1]) as PlanSum).PlanSetups)
                    {
                        try
                        {
                            rx_doses.Add(ps.TotalDose.Dose);
                        }
                        catch
                        {
                            System.Windows.MessageBox.Show("One of the prescriptions for the plansum is not defined");
                            return;
                        }
                    }
                    double rx = rx_doses.Sum();
                    rx95_5 = rx * 0.955;
                    Iso95 = ss.AddStructure("DOSE_REGION", "zz95_ISO");  // second parameter is structure ID
                    Iso95.ConvertDoseLevelToStructure(context.Patient.Courses.First(c => c.Id == psumName.Split('/')[0]).PlanSums.First(ps => ps.Id == psumName.Split('/')[1]).Dose, new DoseValue(value: rx95_5, unit: DoseValue.DoseUnit.Gy));
                }
                else
                {
                    ss = psum.StructureSet;
                    List<double> rx_doses = new List<double>();
                    foreach (PlanSetup ps in (psum as PlanSum).PlanSetups)
                    {
                        try
                        {
                            rx_doses.Add(ps.TotalDose.Dose);
                        }
                        catch
                        {
                            System.Windows.MessageBox.Show("One of the prescriptions for the plansum is not defined");
                            return;
                        }
                    }
                    double rx = rx_doses.Sum();
                    rx95_5 = rx * 0.955;
                    Iso95 = ss.AddStructure("DOSE_REGION", "zz95_ISO");  // second parameter is structure ID
                    Iso95.ConvertDoseLevelToStructure(psum.Dose, new DoseValue(value: rx95_5, unit: DoseValue.DoseUnit.Gy));
                }
            }

            // Retrieve StructureSet
            if (ss == null)
                throw new ApplicationException("The selected plan does not reference a StructureSet.");

            var listStructures = ss.Structures;

            // find PTV
            var ptv = SelectStructureWindow.SelectStructure(ss);
            if (ptv == null) return;
            

            //============================
            // contour 95% isodose holes
            //============================
            foreach (Structure scan in listStructures)
            {
                if (scan.Id.Contains("95"))
                    HoleCount++;
            }


            Structure hk_hole = ss.AddStructure("DOSE_Region", "zzHK 95_" + HoleCount);
            //hk_hole.SegmentVolume = ptv.Margin(-1.5);
            Structure ptv_minus = ss.AddStructure("PTV", "PTV-1.5mm");

            if (ptv.IsHighResolution == true)
            {
                hk_hole.ConvertToHighResolution();
                ptv_minus.ConvertToHighResolution();
                Iso95.ConvertToHighResolution();
            }

            ptv_minus.SegmentVolume = ptv.Margin(-1.5);
            hk_hole.SegmentVolume = ptv_minus.Sub(Iso95);

            //============================
            // remove structures unneccesary for optimization
            //============================
            ss.RemoveStructure(ptv_minus);
            ss.RemoveStructure(Iso95);

        }
    }



    class SelectStructureWindow : Window
    {
        public static Structure SelectStructure(StructureSet ss)
        {

            m_w = new Window();
            //m_w.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            m_w.WindowStartupLocation = WindowStartupLocation.Manual;
            m_w.Left = 500;
            m_w.Top = 150;
            m_w.Width = 300;
            m_w.Height = 350;
            //m_w.SizeToContent = SizeToContent.Height;
            //m_w.SizeToContent = SizeToContent.Width;
            m_w.Title = "Choose Target:";
            var grid = new Grid();
            m_w.Content = grid;
            var list = new System.Windows.Controls.ListBox();
            foreach (var s in ss.Structures.OrderByDescending(x=> x.Id))
            {
                var tempStruct = s.ToString();
                if (tempStruct.ToUpper().Contains("PTV") || tempStruct.ToUpper().Contains("ZHK") || tempStruct.ToUpper().Contains("SIB") || tempStruct.ToUpper().Contains("CTV") || tempStruct.ToUpper().Contains("GTV") || tempStruct.ToUpper().StartsWith("Z") &!tempStruct.Contains("zz95_ISO"))
                {
                    if (tempStruct.Contains(":"))
                    {
                        int index = tempStruct.IndexOf(":");
                        tempStruct = tempStruct.Substring(0, index);
                    }
                    list.Items.Add(s);
                }
            }
            list.VerticalAlignment = VerticalAlignment.Top;
            list.Margin = new Thickness(10, 10, 10, 55);
            grid.Children.Add(list);
            var button = new System.Windows.Controls.Button();
            button.Content = "OK";
            button.Height = 40;
            button.VerticalAlignment = VerticalAlignment.Bottom;
            button.Margin = new Thickness(10, 10, 10, 10);
            button.Click += button_Click;
            grid.Children.Add(button);
            if (m_w.ShowDialog() == true)
            {
                return (Structure)list.SelectedItem;
            }
            return null;
        }

        static Window m_w = null;

        static void button_Click(object sender, RoutedEventArgs e)
        {
            m_w.DialogResult = true;
            m_w.Close();
        }
    }
}
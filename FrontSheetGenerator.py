''' Version 5
    Implementing second worst case goals
'''

import wx, sys, re, ctypes
from connect import *
import wx.lib.agw.genericmessagedialog as GMD
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.platypus import PageBreak, Image
from datetime import datetime
import textwrap
import time

# getting all 'currents'
try:
    patient = get_current("Patient")
    plan = get_current("Plan")
    case = get_current("Case")
    exam = plan.TreatmentCourse.TotalDose.OnDensity.FromExamination.Name
    beam_set = get_current("BeamSet")
    structure_set = case.PatientModel.StructureSets[exam]
except:
    ctypes.windll.user32.MessageBoxW(0, "Script failed, either no patient or no plan loaded.", "Error!", 1)
    sys.exit(1)


class FSGFrame(wx.Frame):
    def __init__(self, *args, **kw):
        # ensure the parent's __init__ is called
        super(FSGFrame, self).__init__(*args, **kw)
        self.makeMenuBar()
        self.CreateStatusBar()
        self.SetStatusText("Front Sheet Generator v1.0")
        self.count = 0
        self.BuildPage()

    def makeMenuBar(self):
        fileMenu = wx.Menu()
        exitItem = fileMenu.Append(wx.ID_EXIT)
        helpMenu = wx.Menu()
        aboutItem = helpMenu.Append(wx.ID_ABOUT)

        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, "&File")
        menuBar.Append(helpMenu, "&Help")

        # Give the menu bar to the frame
        self.SetMenuBar(menuBar)

        # Finally, associate a handler function with the EVT_MENU event for menu item
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

    def OnExit(self, event):
        """Close the frame, terminating the application."""
        self.Close(True)

    def OnAbout(self, event):
        """Display an About Dialog"""
        Awin = GMD.GenericMessageDialog(None,
                                        message="This is the front sheet generator, select your prescription volumes.\nROIs must be set to type: Target.",
                                        caption="Guidance",
                                        agwStyle=wx.OK | wx.ICON_INFORMATION)
        Awin.ShowModal()
        Awin.Destroy()

    def is_number(self, i):
        try:
            float(i)
            return True
        except ValueError:
            return False

    def UpdateCombo1(self, event):
        self.dyn_txt1.SetValue(str(self.combo1.GetValue()))
        self.combo_rp1.SetValue((self.combo1.GetValue()))
        self.rp_num1.SetValue('95')

        # Target box 1 updates
        if str(self.combo1.GetValue()) != '':
            self.def_txt1.Clear()
            self.def_txt1.AppendText(str(self.combo1.GetValue() + ' = '))

            templist = re.split("gtv|Gtv|GTV|ctv|Ctv|CTV|ptv|Ptv|PTV|itv|Itv|\\+|_|ITV", str(self.combo1.GetValue()))
            print(templist)

            for i in templist:
                if 'mm' in str(i) or 'cm' in str(i):
                    self.txt2.SetValue('')
                    print('found cm')
                elif self.is_number(i):
                    self.txt2.SetValue(i)
                    self.rp_dose1.SetValue(i)
                    break
                else:
                    continue
        else:
            self.def_txt1.Clear()
            self.txt2.SetValue('')
            self.rp_num1.SetValue('')

    def UpdateCombo2(self, event):
        self.dyn_txt2.SetValue(str(self.combo2.GetValue()))
        self.combo_rp2.SetValue(self.combo2.GetValue())
        self.rp_num2.SetValue('95')

        # Target box 2 updates
        if str(self.combo2.GetValue()) != '':
            self.def_txt2.Clear()
            self.def_txt2.AppendText(str(self.combo2.GetValue() + ' = '))
            self.combo5.SetSelection(2)

            templist = re.split("gtv|Gtv|GTV|ctv|Ctv|CTV|ptv|Ptv|PTV|itv|Itv|\\+|_|ITV", str(self.combo2.GetValue()))
            print(templist)

            for i in templist:
                if 'mm' in str(i) or 'cm' in str(i):
                    self.txt3.SetValue('')
                    print('found cm')
                elif self.is_number(i):
                    self.txt3.SetValue(i)
                    self.rp_dose2.SetValue(i)
                    break
                else:
                    continue
        else:
            self.def_txt2.Clear()
            self.txt3.SetValue('')
            self.rp_num2.SetValue('')
            self.combo5.SetSelection(0)

    def UpdateCombo3(self, event):
        self.dyn_txt3.SetValue(str(self.combo3.GetValue()))
        self.combo_rp3.SetValue(self.combo3.GetValue())
        self.rp_num3.SetValue('95')
        self.combo6.SetSelection(2)

        # Target box 3 updats

        if str(self.combo3.GetValue()) != '':
            self.def_txt3.Clear()
            self.def_txt3.AppendText(str(self.combo3.GetValue() + ' = '))

            templist = re.split("gtv|Gtv|GTV|ctv|Ctv|CTV|ptv|Ptv|PTV|itv|Itv|\\+|_|ITV", str(self.combo3.GetValue()))
            print(templist)

            for i in templist:
                if 'mm' in str(i) or 'cm' in str(i):
                    self.txt4.SetValue('')
                    print('found cm')
                elif self.is_number(i):
                    self.txt4.SetValue(i)
                    self.rp_dose3.SetValue(i)
                    break
                else:
                    continue
        else:
            self.def_txt3.Clear()
            self.txt4.SetValue('')
            self.rp_num3.SetValue('')
            self.combo6.SetSelection(0)

    def UpdateComboRP1(self, event):
        if str(self.combo_rp1.GetValue()) == '':
            self.rp_num1.SetValue('')
        else:
            self.rp_num1.SetValue('95')

    def UpdateComboRP2(self, event):
        if str(self.combo_rp2.GetValue()) == '':
            self.rp_num2.SetValue('')
        else:
            self.rp_num2.SetValue('95')

    def UpdateComboRP3(self, event):
        if str(self.combo_rp3.GetValue()) == '':
            self.rp_num3.SetValue('')
        else:
            self.rp_num3.SetValue('95')

    def BuildPage(self):
        # Getting Target List
        target_list = ['']
        for i in structure_set.RoiGeometries:
            if i.OfRoi.Type == 'Ptv' or i.OfRoi.Type == 'Ctv' or i.OfRoi.Type == 'Gtv':
                target_list.append(i.OfRoi.Name)

        # Getting Poi List
        poi_list = ['']
        x = 1
        temp_poi = 1
        for i in structure_set.PoiGeometries:
            poi_list.append(i.OfPoi.Name)
            if str(i.OfPoi.Type) == 'Isocenter':
                temp_poi = x
            x += 1

        self.doctor_name = case.Physician.Name.split('^')

        try:
            print(self.doctor_name[3] + self.doctor_name[1][0] + self.doctor_name[0])
        except:
            self.Close(True)
            ctypes.windll.user32.MessageBoxW(0, "Script failed, please check Dr name entered correctly.", "Error!",
                                             0x40000)
            sys.exit(1)

        # create panel and box to manage the layout of widgets
        pnl = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # Prescription Title
        title_p = wx.StaticText(pnl, label='Prescriptions')
        font = title_p.GetFont()
        font.SetPointSize(15)
        font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_p.SetFont(wx.Font(font.GetPointSize(), font.GetFamily(), font.GetStyle(), wx.FONTWEIGHT_BOLD))
        vbox.Add(title_p, proportion=0, flag=wx.ALL | wx.EXPAND, border=7)

        # Objects for table 1
        cl1 = wx.StaticText(pnl, label="Target:")
        cl2 = wx.StaticText(pnl, label="Prescription method:")
        cl3 = wx.StaticText(pnl, label="Dose:")
        cl4 = wx.StaticText(pnl, label="Fractions:")
        cl5 = wx.StaticText(pnl, label="Plan Type:")
        blank = wx.StaticText(pnl, label='-')
        blank2 = wx.StaticText(pnl, label='-')
        self.combo1 = wx.ComboBox(pnl, choices=target_list, style=wx.CB_READONLY)
        self.combo1.Bind(wx.EVT_COMBOBOX, self.UpdateCombo1)
        self.combo2 = wx.ComboBox(pnl, choices=target_list, style=wx.CB_READONLY)
        self.combo2.Bind(wx.EVT_COMBOBOX, self.UpdateCombo2)
        self.combo3 = wx.ComboBox(pnl, choices=target_list, style=wx.CB_READONLY)
        self.combo3.Bind(wx.EVT_COMBOBOX, self.UpdateCombo3)
        try:
            if beam_set.Prescription.DosePrescriptions[0].OnStructure.Name != '':
                tempx = 0
                for i in target_list:
                    if i == str(beam_set.Prescription.DosePrescriptions[0].OnStructure.Name):
                        break
                    tempx = tempx + 1
                print(beam_set.Prescription.DosePrescriptions[0].OnStructure.Name)
                self.combo1.SetSelection(tempx)
        except:
            self.combo1.SetSelection(0)

        self.combo4 = wx.ComboBox(pnl, choices=['Target Mean', 'Isodose'], style=wx.CB_READONLY | wx.EXPAND)
        self.combo4.SetSelection(0)
        self.combo5 = wx.ComboBox(pnl, choices=['', 'Target Mean', 'Isodose'], style=wx.CB_READONLY | wx.EXPAND)
        self.combo6 = wx.ComboBox(pnl, choices=['', 'Target Mean', 'Isodose'], style=wx.CB_READONLY | wx.EXPAND)
        self.txt2 = wx.TextCtrl(pnl)

        try:
            self.txt2.SetValue(str(beam_set.Prescription.DosePrescriptions[0].DoseValue / 100))
        except:
            self.txt2.SetValue('')
        self.txt3 = wx.TextCtrl(pnl)
        self.txt4 = wx.TextCtrl(pnl)

        self.txt5 = wx.TextCtrl(pnl)
        self.txt5.SetValue(str(beam_set.FractionationPattern.NumberOfFractions))

        self.combo7 = wx.ComboBox(pnl, choices=['IMPT', 'SFUD', 'VMAT'], style=wx.CB_READONLY)
        self.combo7.SetSelection(0)
        blank7 = wx.StaticText(pnl, label='-')
        blank8 = wx.StaticText(pnl, label='-')

        # Page build
        # First table
        fgs = wx.FlexGridSizer(4, 5, 10, 10)
        fgs.AddMany([cl1, cl2, cl3, cl4, cl5])
        fgs.AddMany([self.combo1, (self.combo4, 1, wx.EXPAND), self.txt2, self.txt5, self.combo7])
        fgs.AddMany([self.combo2, (self.combo5, 1, wx.EXPAND), self.txt3, blank, blank7])
        fgs.AddMany([self.combo3, (self.combo6, 1, wx.EXPAND), self.txt4, blank2, blank8])
        vbox.Add(fgs, proportion=0, flag=wx.ALL | wx.EXPAND, border=2)

        # Table 2 : treatment frequency
        self.tfreq = wx.StaticText(pnl, label='Treatment Frequency:')
        self.tfreq_combo = wx.ComboBox(pnl, choices=['Daily', 'Alternate Days', 'Weekly', 'Once'], style=wx.CB_READONLY)
        self.tfreq_combo.SetSelection(0)

        fgs2 = wx.FlexGridSizer(1, 2, 10, 10)
        fgs2.AddMany([self.tfreq, self.tfreq_combo])
        vbox.Add(fgs2, proportion=0, flag=wx.ALL | wx.EXPAND, border=2)

        # Definitions Title
        title_d = wx.StaticText(pnl, label='Target Definitions')
        font = title_d.GetFont()
        font.SetPointSize(15)
        font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_d.SetFont(wx.Font(font.GetPointSize(), font.GetFamily(), font.GetStyle(), wx.FONTWEIGHT_BOLD))
        vbox.Add(title_d, proportion=0, flag=wx.ALL | wx.EXPAND, border=7)

        # objects for table 3
        target_lbl = wx.StaticText(pnl, label='Target:')
        definition_lbl = wx.StaticText(pnl,
                                       label='Definition:                                                                               ',
                                       style=wx.LEFT)
        self.dyn_txt1 = wx.TextCtrl(pnl)
        self.dyn_txt1.SetValue(str(self.combo1.GetValue()))
        self.dyn_txt2 = wx.TextCtrl(pnl)
        self.dyn_txt3 = wx.TextCtrl(pnl)
        self.def_txt1 = wx.TextCtrl(pnl)
        self.def_txt1.AppendText(str(self.combo1.GetValue() + ' = '))
        self.def_txt2 = wx.TextCtrl(pnl)
        self.def_txt3 = wx.TextCtrl(pnl)

        # Third table
        fgs3 = wx.FlexGridSizer(4, 2, 10, 10)
        fgs3.AddMany([target_lbl, definition_lbl,
                      self.dyn_txt1, (self.def_txt1, 2, wx.EXPAND),
                      self.dyn_txt2, (self.def_txt2, 2, wx.EXPAND),
                      self.dyn_txt3, (self.def_txt3, 2, wx.EXPAND)])
        vbox.Add(fgs3, proportion=0, flag=wx.ALL | wx.EXPAND, border=2)

        # Statistics Title
        title_s = wx.StaticText(pnl, label='Plan Statistics')
        font = title_s.GetFont()
        font.SetPointSize(15)
        font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_s.SetFont(wx.Font(font.GetPointSize(), font.GetFamily(), font.GetStyle(), wx.FONTWEIGHT_BOLD))
        vbox.Add(title_s, proportion=0, flag=wx.ALL | wx.EXPAND, border=7)

        # Objects for table 4
        self.lbl_target = wx.StaticText(pnl, label='Target(s) for reporting:')
        self.lbl_dose = wx.StaticText(pnl, label='Dose Gy')
        self.lbl_num = wx.StaticText(pnl, label='reporting %:')
        self.lbl_poi = wx.StaticText(pnl, label='POI(s) for dose:')

        self.combo_rp1 = wx.ComboBox(pnl, choices=target_list, style=wx.CB_READONLY)
        self.combo_rp1.SetSelection(tempx)
        self.combo_rp1.Bind(wx.EVT_COMBOBOX, self.UpdateComboRP1)
        self.combo_rp2 = wx.ComboBox(pnl, choices=target_list, style=wx.CB_READONLY)
        self.combo_rp2.Bind(wx.EVT_COMBOBOX, self.UpdateComboRP2)
        self.combo_rp3 = wx.ComboBox(pnl, choices=target_list, style=wx.CB_READONLY)
        self.combo_rp3.Bind(wx.EVT_COMBOBOX, self.UpdateComboRP3)
        self.rp_num1 = wx.TextCtrl(pnl)
        self.rp_num1.SetValue('95')
        self.rp_num2 = wx.TextCtrl(pnl)
        self.rp_num3 = wx.TextCtrl(pnl)
        self.rp_poi1 = wx.ComboBox(pnl, choices=poi_list, style=wx.CB_READONLY)
        self.rp_poi1.SetSelection(temp_poi)
        self.rp_poi2 = wx.ComboBox(pnl, choices=poi_list, style=wx.CB_READONLY)
        self.rp_poi3 = wx.ComboBox(pnl, choices=poi_list, style=wx.CB_READONLY)
        self.rp_dose1 = wx.TextCtrl(pnl)
        self.rp_dose1.SetValue(self.txt2.GetValue())
        self.rp_dose2 = wx.TextCtrl(pnl)
        self.rp_dose3 = wx.TextCtrl(pnl)
        blank3 = wx.StaticText(pnl, label='                                  ')
        blank4 = wx.StaticText(pnl, label='                                  ')
        blank5 = wx.StaticText(pnl, label='                                  ')
        blank6 = wx.StaticText(pnl, label='                                  ')

        # Fourth table
        fgs4 = wx.FlexGridSizer(4, 5, 10, 10)
        fgs4.AddMany([self.lbl_target, self.lbl_dose, self.lbl_num, blank3, self.lbl_poi,
                      self.combo_rp1, self.rp_dose1, self.rp_num1, blank4, self.rp_poi1,
                      self.combo_rp2, self.rp_dose2, self.rp_num2, blank5, self.rp_poi2,
                      self.combo_rp3, self.rp_dose3, self.rp_num3, blank6, self.rp_poi3])
        vbox.Add(fgs4, proportion=0, flag=wx.ALL | wx.EXPAND, border=2)

        # Objects for table 5
        comments_lbl = wx.StaticText(pnl,
                                     label='Additional Comments:                                                                                             ',
                                     style=wx.LEFT)
        self.comments = wx.TextCtrl(pnl, style=wx.TE_MULTILINE, size=(20, 100))
        blank3 = wx.StaticText(pnl, label='')
        self.btn_pdf = wx.Button(pnl, label='Generate PDF', style=wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM)
        self.Bind(wx.EVT_BUTTON, self.MakePDF, self.btn_pdf)

        # Fifth table
        fgs5 = wx.FlexGridSizer(2, 2, 10, 10)
        fgs5.AddMany([comments_lbl, blank3,
                      (self.comments, 0, wx.EXPAND), self.btn_pdf])
        vbox.Add(fgs5, proportion=0, flag=wx.ALL | wx.EXPAND, border=2)

        pnl.SetSizer(vbox)

    def FindPosition(self):
        # Finding correct RA, returning position or -1 if not found.
        position = 0
        print('got in')
        for x in case.TreatmentDelivery.RadiationSetScenarioGroups:
            print('looping')
            num_beams = 0
            pass_beams = 0
            if x.ReferencedRadiationSet.DicomPlanLabel == beam_set.DicomPlanLabel:
                print(beam_set.DicomPlanLabel)
                print(x.ReferencedRadiationSet.DicomPlanLabel)
                print(x.ReferencedRadiationSet.FractionDose.OnDensity.FromExamination.Name)
                print(exam)
                if x.ReferencedRadiationSet.FractionDose.OnDensity.FromExamination.Name == exam:
                    print('same ct')
                    for beam in x.ReferencedRadiationSet.Beams:
                        print(beam.Name)
                        num_beams += 1
                        print(num_beams)
                        if beam.BeamMU == beam_set.Beams[num_beams - 1].BeamMU:
                            pass_beams += 1
                            print(pass_beams)

                    if num_beams == pass_beams:
                        return (position)
            position += 1
        return -1

    def SecondWorst(self, target, rvalue, position):
        # this is where i find and return second worst RA

        templist = []
        x1 = float('inf')
        x2 = float('inf')

        for x in case.TreatmentDelivery.RadiationSetScenarioGroups[position].DiscreteFractionDoseScenarios:
            z = x.GetRelativeVolumeAtDoseValues(RoiName=target,
                                                DoseValues=[rvalue / int(self.txt5.GetValue())])
            templist.append(z[0] * 100)

        for i in templist:
            print(i)
            if i < x1:
                x2 = x1
                x1 = i
            elif i < x2:
                x2 = i

        self.scenarios = len(templist)
        value = x2
        return (value)

    def MakePDF(self, event):
        start_time = time.time()
        self.btn_pdf.Disable()

        # Fetching position of RA / checking if RA exists
        ra_list = []
        temp_target = []
        temp_rp_dose = []
        position = self.FindPosition()

        # Making list of targets and doses to interpret for RA
        if self.combo_rp1.GetValue() != '' and self.rp_dose1.GetValue() != '' and self.rp_num1 != '':
            temp_target.append(str(self.combo_rp1.GetValue()))
            temp_rp_dose.append(float(self.rp_dose1.GetValue()))
        if self.combo_rp2.GetValue() != '' and self.rp_dose2.GetValue() != '' and self.rp_num2 != '':
            temp_target.append(str(self.combo_rp2.GetValue()))
            temp_rp_dose.append(float(self.rp_dose2.GetValue()))
        if self.combo_rp3.GetValue() != '' and self.rp_dose3.GetValue() != '' and self.rp_num3 != '':
            temp_target.append(str(self.combo_rp3.GetValue()))
            temp_rp_dose.append(float(self.rp_dose3.GetValue()))

        if self.combo7.GetValue() == 'IMPT' or self.combo7.GetValue() == 'SFUD':
            position = self.FindPosition()
            if position == -1:
                self.Close(True)
                ctypes.windll.user32.MessageBoxW(0, "Script failed, can't find relevant robust analysis.", "Error!",
                                                 0x40000)
                sys.exit(1)

            i = 0
            for target in temp_target:
                ra_list.append(self.SecondWorst(target, temp_rp_dose[i], position))
                i += 1

        # Gathering information for header and setting save directory
        pname = patient.Name.split("^")
        pdf_name = (pname[1] + pname[0] + '_' + patient.PatientID + "_frontsheet.PDF")
        save_name = os.path.join(os.path.expanduser("~"), "//ppgbcipmsqdat01/MOSAIQ_DATA/DB/ESCAN/RaystationPrintouts/",
                                 pdf_name)
        logo = ImageReader(r'S:\\Clinical\Radiotherapy\Planning\Raystation\\logo.jpg')

        c = canvas.Canvas(save_name)

        cheight = 18
        cwidth = 525

        # creating header (Name)
        c.translate(2.25 * inch, 10.75 * inch)
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(0, 0, "Rutherford Cancer Centre – Proton Beam Therapy Treatment Plan")
        c.translate(-1.75 * inch, 0.1 * inch)
        c.drawImage(logo, 0, 0, 125, 51)
        c.setFillColor('black')
        c.rect(0, -0.2 * inch, 525, 1, stroke=1, fill=1)

        # Patient
        c.translate(0, -0.5 * inch)
        c.rect(0, -5, cwidth, cheight, stroke=1, fill=0)
        c.drawString(5, 0, "Patient:")
        c.setFont("Helvetica", 12)
        name = str(pname[1]).lower().capitalize() + ' ' + str(pname[0]).upper()
        c.drawString(0.42 * cwidth, 0, name)

        # MRN
        c.translate(0, -0.25 * inch)
        c.rect(0, -5, cwidth, cheight, stroke=1, fill=0)
        c.rect(0.4 * cwidth, (-1.5 * cheight) + 4, 1, 3 * cheight, stroke=1, fill=0)
        c.drawString(0.42 * cwidth, 0, patient.PatientID)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(5, 0, "RCC Number:")

        # Plan/Trial
        c.translate(0, -0.25 * inch)
        c.rect(0, -5, cwidth, cheight, stroke=1, fill=0)
        c.drawString(5, 0, "Plan/Trial:")
        c.setFont("Helvetica", 12)
        plan_lbl = case.CaseName + r' / ' + beam_set.DicomPlanLabel
        c.drawString(0.42 * cwidth, 0, plan_lbl)

        # Prescriptions
        c.translate(0, -0.6 * inch)
        c.rect(0, -5, cwidth, cheight, stroke=1, fill=0)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(5, 0, "Prescriptions")
        c.translate(0, -3 * cheight)
        c.rect(0, -5, cwidth, 3 * cheight, stroke=1, fill=0)
        c.rect(0, -5, 0.4 * cwidth, 3 * cheight, stroke=1, fill=0)
        c.rect(0, -5, 0.2 * cwidth, 3 * cheight, stroke=1, fill=0)

        # Prescription headers and rectangles
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(0.1 * cwidth, 2 * cheight, "Clinical")
        c.drawCentredString(0.1 * cwidth, 1.25 * cheight, "Oncologist")
        c.drawCentredString(0.3 * cwidth, cheight, "Target/Site")
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(0.1 * cwidth, 0, "(IRMER Practitioner)")
        c.rect(0.4 * cwidth, -5, 0.05 * cwidth, 3 * cheight, stroke=1, fill=0)
        c.drawCentredString(0.425 * cwidth, 1.75 * cheight, "Total")
        c.drawCentredString(0.425 * cwidth, cheight, "Dose")
        c.drawCentredString(0.425 * cwidth, 0.25 * cheight, "(Gy)")
        c.rect(0.45 * cwidth, -5, 0.25 * cwidth, 3 * cheight, stroke=1, fill=0)
        c.rect(0.45 * cwidth, -5, 0.05 * cwidth, 1.5 * cheight, stroke=1, fill=0)
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(0.575 * cwidth, 1.75 * cheight, "Fractionation")
        c.drawCentredString(0.475 * cwidth, 0.25 * cheight, "no.")
        c.rect(0.5 * cwidth, -5, 0.13 * cwidth, 1.5 * cheight, stroke=1, fill=0)
        c.drawCentredString(0.565 * cwidth, 0.25 * cheight, "Schedule")
        c.rect(0.63 * cwidth, -5, 0.07 * cwidth, 1.5 * cheight, stroke=1, fill=0)
        c.drawCentredString(0.665 * cwidth, 0.25 * cheight, "dose/#")
        c.rect(0.7 * cwidth, -5, 0.13 * cwidth, 3 * cheight, stroke=1, fill=0)
        c.drawCentredString(0.765 * cwidth, 1.75 * cheight, "Prescription")
        c.drawCentredString(0.765 * cwidth, cheight, "Point /")
        c.drawCentredString(0.765 * cwidth, 0.25 * cheight, "Isodose")
        c.rect(0.83 * cwidth, -5, 0.17 * cwidth, 3 * cheight, stroke=1, fill=0)
        c.drawCentredString(0.915 * cwidth, cheight, "Comments")

        # Prescription Lines
        # Consultant
        c.setFont("Helvetica", 11)
        c.translate(0, -2 * cheight)
        # c.rect(0, -5, 0.4 * cwidth, 2 * cheight, stroke=1, fill=0)
        c.rect(0.2 * cwidth, -5, 0.2 * cwidth, 2 * cheight, stroke=1, fill=0)

        # Prescription data
        total_prescriptions = 1
        dose = float(self.txt2.GetValue())
        fractions = int(self.txt5.GetValue())
        dose_per = dose / fractions

        c.drawCentredString(0.3 * cwidth, 0.5 * cheight, self.combo1.GetValue())
        c.rect(0.4 * cwidth, -5, 0.05 * cwidth, 2 * cheight, stroke=1, fill=0)
        c.drawCentredString(0.425 * cwidth, 0.5 * cheight, self.txt2.GetValue())
        c.drawCentredString(0.665 * cwidth, 0.5 * cheight, str("{:.2f}".format(dose_per)))
        c.rect(0.63 * cwidth, -5, 0.07 * cwidth, 2 * cheight, stroke=1, fill=0)
        c.rect(0.7 * cwidth, -5, 0.13 * cwidth, 2 * cheight, stroke=1, fill=0)
        c.drawCentredString(0.765 * cwidth, 0.5 * cheight, self.combo4.GetValue())

        if self.combo2.GetValue() != '' and self.combo5.GetValue() != '' and self.txt3.GetValue() != '':
            total_prescriptions += 1
            c.translate(0, -2 * cheight)
            c.rect(0.2 * cwidth, -5, 0.2 * cwidth, 2 * cheight, stroke=1, fill=0)

            # Prescription data
            dose = float(self.txt3.GetValue())
            dose_per = dose / fractions

            c.drawCentredString(0.3 * cwidth, 0.5 * cheight, self.combo2.GetValue())
            c.rect(0.4 * cwidth, -5, 0.05 * cwidth, 2 * cheight, stroke=1, fill=0)
            c.drawCentredString(0.425 * cwidth, 0.5 * cheight, str("{:.1f}".format(dose)))
            c.drawCentredString(0.665 * cwidth, 0.5 * cheight, str("{:.2f}".format(dose_per)))
            c.rect(0.63 * cwidth, -5, 0.07 * cwidth, 2 * cheight, stroke=1, fill=0)
            c.rect(0.7 * cwidth, -5, 0.13 * cwidth, 2 * cheight, stroke=1, fill=0)
            c.drawCentredString(0.765 * cwidth, 0.5 * cheight, self.combo5.GetValue())

        if self.combo3.GetValue() != '' and self.combo6.GetValue() != '' and self.txt4.GetValue() != '':
            total_prescriptions += 1
            c.translate(0, -2 * cheight)
            c.rect(0.2 * cwidth, -5, 0.2 * cwidth, 2 * cheight, stroke=1, fill=0)

            # Prescription data
            dose = float(self.txt4.GetValue())
            dose_per = dose / fractions

            c.drawCentredString(0.3 * cwidth, 0.5 * cheight, self.combo3.GetValue())
            c.rect(0.4 * cwidth, -5, 0.05 * cwidth, 2 * cheight, stroke=1, fill=0)
            c.drawCentredString(0.425 * cwidth, 0.5 * cheight, str("{:.1f}".format(dose)))
            c.drawCentredString(0.665 * cwidth, 0.5 * cheight, str("{:.2f}".format(dose_per)))
            c.rect(0.63 * cwidth, -5, 0.07 * cwidth, 2 * cheight, stroke=1, fill=0)
            c.rect(0.7 * cwidth, -5, 0.13 * cwidth, 2 * cheight, stroke=1, fill=0)
            c.drawCentredString(0.765 * cwidth, 0.5 * cheight, self.combo6.GetValue())

        # Checking number of beams
        no_beams = 0
        for i in beam_set.Beams:
            no_beams += 1
        beam_string = str(no_beams) + ' Beam '

        # Setting merge cells using count of prescription
        temp_cellh = total_prescriptions * (2 * cheight)
        if total_prescriptions == 1:
            temp_texth = (0.25 * temp_cellh)
        elif total_prescriptions == 2:
            temp_texth = (0.375 * temp_cellh)
        else:
            temp_texth = (0.4167 * temp_cellh)

        c.drawCentredString(0.1 * cwidth, temp_texth,
                            self.doctor_name[3] + ' ' + self.doctor_name[1][0] + ' ' + self.doctor_name[0])

        c.drawCentredString(0.915 * cwidth, temp_texth, str(beam_string) + str(self.combo7.GetValue()))
        c.drawCentredString(0.475 * cwidth, temp_texth, str(fractions))
        c.drawCentredString(0.565 * cwidth, temp_texth, self.tfreq_combo.GetValue())

        c.rect(0, -5, 0.2 * cwidth, temp_cellh, stroke=1, fill=0)
        c.rect(0.83 * cwidth, -5, 0.17 * cwidth, temp_cellh, stroke=1, fill=0)
        c.rect(0.45 * cwidth, -5, 0.05 * cwidth, temp_cellh, stroke=1, fill=0)
        c.rect(0.5 * cwidth, -5, 0.13 * cwidth, temp_cellh, stroke=1, fill=0)

        # Printing target definitions
        c.translate(0, -0.6 * inch)
        c.rect(0, -5, cwidth, cheight, stroke=1, fill=0)
        c.drawString(5, 0, str(self.def_txt1.GetValue()))

        if self.def_txt2.GetValue() != '':
            c.translate(0, -0.25 * inch)
            c.rect(0, -5, cwidth, cheight, stroke=1, fill=0)
            c.drawString(5, 0, str(self.def_txt2.GetValue()))

        if self.def_txt3.GetValue() != '':
            c.translate(0, -0.25 * inch)
            c.rect(0, -5, cwidth, cheight, stroke=1, fill=0)
            c.drawString(5, 0, str(self.def_txt3.GetValue()))

        # Print statistics
        c.translate(0, -0.6 * inch)
        c.rect(0, -5, cwidth, cheight, stroke=1, fill=0)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(5, 0, 'Treatment Plan Statistics (all percentages are ±1%)')
        c.translate(0, -0.25 * inch)
        c.setFillAlpha(0.25)
        c.setFillColor('lightgrey')
        c.rect(0, -5, cwidth, cheight, stroke=1, fill=1)
        c.setFillColor('black')
        c.drawString(5, 0, "‘ICRU Max’ dose:")

        # Finding and printing the 2cc max dose
        for i in structure_set.RoiGeometries:
            if i.OfRoi.Type == 'External':
                total_volume = plan.TreatmentCourse.TotalDose.GetDoseGridRoi(
                    RoiName=i.OfRoi.Name).RoiVolumeDistribution.TotalVolume
                print(total_volume)
                print(i.OfRoi.Name)
                print(2 / total_volume)
                icru = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName=i.OfRoi.Name,
                                                                               RelativeVolumes=[2 / total_volume])
                icru = icru / float(self.txt2.GetValue())
                icru = str("{:.1f}".format(icru[0])) + '% of prescribed dose received by 2cc of tissue'
                c.setFont("Helvetica", 11)
                c.drawString(0.33 * cwidth, 0, str(icru))

        c.translate(0, -0.25 * inch)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(5, 0, 'Target Coverage:')
        c.setFont("Helvetica", 11)
        total_reported = 1

        # calculating and printing the reporting structure statistics
        rv_1 = plan.TreatmentCourse.TotalDose.GetRelativeVolumeAtDoseValues(RoiName=str(self.combo_rp1.GetValue()),
                                                                            DoseValues=[
                                                                                float(self.rp_dose1.GetValue()) * (
                                                                                    float(
                                                                                        self.rp_num1.GetValue()))]) * 100
        cov_str1 = str("{:.1f}".format(rv_1[0])) + '% of ' + self.combo_rp1.GetValue() + ' receives at least ' + str(
            self.rp_num1.GetValue()) + '% of ' + str(self.rp_dose1.GetValue()) + 'Gy'
        c.drawString(0.33 * cwidth, 0, cov_str1)

        if str(self.combo7.GetValue()) == 'IMPT' or str(self.combo7.GetValue()) == 'SFUD':
            c.setFillColorRGB(34 / 255, 139 / 255, 34 / 255)
            cov_str1_2 = str(self.combo_rp1.GetValue()) + ' robustly optimised, ' + str(
                "{:.1f}".format(ra_list[0])) + '% receives at least ' + str(
                self.rp_num1.GetValue()) + '% of ' + str(self.rp_dose1.GetValue()) + 'Gy in 2nd worst case (see Robust Analysis summary)'

            if len(cov_str1_2) > 70:
                wrap_text = textwrap.wrap(cov_str1_2, width=65)
                for i in wrap_text:
                    c.translate(0, -0.25 * inch)
                    c.drawString(0.33 * cwidth, 0, i)
            else:
                c.translate(0, -0.25 * inch)
                c.drawString(0.33 * cwidth, 0, cov_str1_2)

        if self.combo_rp2.GetValue() != '' and self.rp_dose2.GetValue() != '' and self.rp_num2 != '':
            total_reported += 1
            c.setFillColorRGB(0, 0, 0)
            c.translate(0, -0.25 * inch)
            rv_2 = plan.TreatmentCourse.TotalDose.GetRelativeVolumeAtDoseValues(RoiName=self.combo_rp2.GetValue(),
                                                                                DoseValues=[
                                                                                    float(self.rp_dose2.GetValue()) * (
                                                                                        float(
                                                                                            self.rp_num2.GetValue()))]) * 100
            cov_str2 = str(
                "{:.1f}".format(rv_2[0])) + '% of ' + self.combo_rp2.GetValue() + ' receives at least ' + str(
                self.rp_num2.GetValue()) + '% of ' + str(self.rp_dose2.GetValue()) + 'Gy'
            c.drawString(0.33 * cwidth, 0, cov_str2)

            if str(self.combo7.GetValue()) == 'IMPT' or str(self.combo7.GetValue()) == 'SFUD':
                c.setFillColorRGB(34/255,139/255,34/255)
                cov_str2_2 = str(self.combo_rp2.GetValue()) + ' robustly optimised, ' + str(
                    "{:.1f}".format(ra_list[total_reported-1])) + '% receives at least ' + str(
                    self.rp_num2.GetValue()) + '% of '+ str(self.rp_dose2.GetValue()) +'Gy in 2nd worst case (see Robust Analysis summary)'

                if len(cov_str2_2) > 70:
                    wrap_text = textwrap.wrap(cov_str2_2, width=65)
                    for i in wrap_text:
                        c.translate(0, -0.25 * inch)
                        c.drawString(0.33*cwidth, 0, i)
                else:
                    c.translate(0, -0.25 * inch)
                    c.drawString(0.33*cwidth, 0, cov_str2_2)

        if self.combo_rp3.GetValue() != '' and self.rp_dose3.GetValue() != '' and self.rp_num3 != '':
            total_reported += 1
            c.setFillColorRGB(0, 0, 0)
            c.translate(0, -0.25 * inch)
            rv_3 = plan.TreatmentCourse.TotalDose.GetRelativeVolumeAtDoseValues(RoiName=self.combo_rp3.GetValue(),
                                                                                DoseValues=[
                                                                                    float(self.rp_dose3.GetValue()) * (
                                                                                        float(
                                                                                            self.rp_num3.GetValue()))]) * 100
            cov_str3 = str(
                "{:.1f}".format(rv_3[0])) + '% of ' + self.combo_rp3.GetValue() + ' receives at least ' + str(
                self.rp_num3.GetValue()) + '% of ' + str(self.rp_dose3.GetValue()) + 'Gy'
            c.drawString(0.33 * cwidth, 0, cov_str3)

            if str(self.combo7.GetValue()) == 'IMPT' or str(self.combo7.GetValue()) == 'SFUD':
                c.setFillColorRGB(34/255,139/255,34/255)
                cov_str3_2 = str(self.combo_rp3.GetValue()) + ' robustly optimised, ' + str(
                    "{:.1f}".format(ra_list[total_reported-1])) + '% receives at least ' + str(
                    self.rp_num3.GetValue()) + '% of '+ str(self.rp_dose3.GetValue()) +'Gy in 2nd worst case (see Robust Analysis summary)'

                if len(cov_str3_2) > 70:
                    wrap_text = textwrap.wrap(cov_str3_2, width=65)
                    for i in wrap_text:
                        c.translate(0, -0.25 * inch)
                        c.drawString(0.33*cwidth, 0, i)
                else:
                    c.translate(0, -0.25 * inch)
                    c.drawString(0.33*cwidth, 0, cov_str3_2)

        # Drawing dynamic rectangle, for IMPT there are double the number of stats
        c.setFillColorRGB(0, 0, 0)
        if str(self.combo7.GetValue()) == 'IMPT' or str(self.combo7.GetValue()) == 'SFUD':
            total_reported = total_reported * 3

        c.rect(0, -5, cwidth, total_reported * cheight)
        c.translate(0, -0.25 * inch)

        stat_total = 1
        if self.rp_poi2.GetValue() != '':
            stat_total += 1
        if self.rp_poi3.GetValue() != '':
            stat_total += 1

        c.setFillColor('LightGrey')
        c.rect(0, -5 - ((stat_total - 1) * cheight), cwidth, stat_total * cheight, fill=1)
        c.setFillColor('Black')

        # POI Statistics
        c.setFont("Helvetica-Bold", 11)
        c.drawString(5, 0, 'ICRU Dose Reference Point:')

        point = structure_set.PoiGeometries[self.rp_poi1.GetValue()].Point
        point = {'x': point.x, 'y': point.y, 'z': point.z}
        beamFOR = beam_set.FrameOfReference
        stat1 = plan.TreatmentCourse.TotalDose.InterpolateDoseInPoint(Point=point, PointFrameOfReference=beamFOR) / 100
        c.setFont("Helvetica", 11)
        c.drawString(0.33 * cwidth, 0, str("{:.1f}".format(stat1) + 'Gy at ' + str(self.rp_poi1.GetValue())))

        if self.rp_poi2.GetValue() != '':
            c.translate(0, -0.25 * inch)

            point = structure_set.PoiGeometries[self.rp_poi2.GetValue()].Point
            point = {'x': point.x, 'y': point.y, 'z': point.z}
            beamFOR = beam_set.FrameOfReference
            stat2 = plan.TreatmentCourse.TotalDose.InterpolateDoseInPoint(Point=point,
                                                                          PointFrameOfReference=beamFOR) / 100
            c.setFont("Helvetica", 11)
            c.drawString(0.33 * cwidth, 0, str("{:.1f}".format(stat2) + 'Gy at ' + str(self.rp_poi2.GetValue())))

        if self.rp_poi3.GetValue() != '':
            c.translate(0, -0.25 * inch)

            point = structure_set.PoiGeometries[self.rp_poi3.GetValue()].Point
            point = {'x': point.x, 'y': point.y, 'z': point.z}
            beamFOR = beam_set.FrameOfReference
            stat2 = plan.TreatmentCourse.TotalDose.InterpolateDoseInPoint(Point=point,
                                                                          PointFrameOfReference=beamFOR) / 100
            c.setFont("Helvetica", 11)
            c.drawString(0.33 * cwidth, 0, str("{:.1f}".format(stat2)) + 'Gy at ' + str(self.rp_poi3.GetValue()))

        c.translate(0, -0.5 * inch)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(5, 0, 'Additional Comments:')
        c.translate(0, -0.35 * inch)
        c.setFont("Helvetica", 11)

        # Text wrapped comments section
        if self.comments.GetValue() != '':
            if len(self.comments.GetValue()) > 70:
                wrap_text = textwrap.wrap(self.comments.GetValue(), width=95)
                for i in wrap_text:
                    c.drawString(5, 0, i)
                    c.translate(0, -0.35 * inch)
            else:
                c.drawString(5, 0, self.comments.GetValue())
                c.translate(0, -0.35 * inch)

        c.drawString(5, 0, 'This plan document contains details of each Pencil Beam Scanning Target Volume used.')
        c.translate(0, -0.35 * inch)
        c.drawString(5, 0, 'This plan document contains OAR dose statistics.')
        c.translate(0, -0.35 * inch)
        c.drawString(5, 0, 'The full treatment plan is available for review in Raystation Treatment Planning System.')
        c.translate(0, -0.35 * inch)

        # printing pdf
        c.showPage()
        c.save()

        total_reported = total_reported/3

        # saving time stats
        elapsed_time = time.time() - start_time
        f = open(r"S:\\Clinical\Radiotherapy\Planning\Raystation\\FSG_times.txt", "a")
        f.write(str(patient.PatientID) + ' ' + str(self.combo7.GetValue()) + ' num_targets: ' + str(total_reported) + ' num_scenarios: ' + str(self.scenarios) + ' ' + str(elapsed_time) + 's\n')
        f.close()
        self.Close(True)


if __name__ == '__main__':
    app = wx.App()
    frm = FSGFrame(None, title='Front Sheet Generator', style=wx.STAY_ON_TOP | wx.DEFAULT_FRAME_STYLE,
                   size=wx.Size(600, 650))

    # display centre
    dw, dh = wx.DisplaySize()
    w, h = frm.GetSize()
    x = dw / 2 - w / 2
    y = dh / 2 - h / 2
    frm.SetPosition((x, y))

    frm.Show()
    app.MainLoop()


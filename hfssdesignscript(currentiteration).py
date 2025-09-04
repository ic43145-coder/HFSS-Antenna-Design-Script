# -*- coding: utf-8 -*-
import ScriptEnv
import os

# Initialize HFSS scripting environment
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oProject = oDesktop.NewProject()
oDesign = oProject.InsertDesign("HFSS", "DualMicrostripAntenna_SymmetricStepped", "DrivenModal", "")
oEditor = oDesign.SetActiveEditor("3D Modeler")

# -----------------------------------------------------
# Parameters (all in mm)
# -----------------------------------------------------
sub_x = 25.5
sub_y = 25.5
sub_h = 1.0

trace_thick_w = 1.0     # thick trace width (1 mm)
trace_thin_w = 0.04     # thin trace width (40 μm)
trace_h = 0.035         # copper thickness (35 μm)

L_parallel = 8.0
L_perpendicular = 5.0
L_taper = 3.0
taper_steps = 50        # reduced number of steps for Student version

bottom_clearance = 2.0
miter_size = 1.0        # miter cut size at bend (mm)

# Positions
antenna1_x_start = -sub_x / 2
antenna2_x_start = sub_x / 2 - L_parallel
parallel_y = -sub_y / 2 + bottom_clearance

# -----------------------------------------------------
# Helper function
# -----------------------------------------------------
def mm(val):
    return str(val) + "mm"

# -----------------------------------------------------
# Substrate (solid)
# -----------------------------------------------------
oEditor.CreateBox([
    "NAME:BoxParameters",
    "XPosition:=", mm(-sub_x/2),
    "YPosition:=", mm(-sub_y/2),
    "ZPosition:=", mm(0),
    "XSize:=", mm(sub_x),
    "YSize:=", mm(sub_y),
    "ZSize:=", mm(sub_h)
], [
    "NAME:Attributes",
    "Name:=", "Substrate",
    "MaterialName:=", "Rogers RO4350 (tm)",
    "Color:=", "(255 255 255)",
    "Transparency:=", 0
])

# -----------------------------------------------------
# Ground Plane (sheet)
# -----------------------------------------------------
oEditor.CreateRectangle([
    "NAME:RectangleParameters",
    "IsCovered:=", True,
    "XStart:=", mm(-sub_x/2),
    "YStart:=", mm(-sub_y/2),
    "ZStart:=", mm(0),
    "Width:=", mm(sub_x),
    "Height:=", mm(sub_y),
    "WhichAxis:=", "Z"
], [
    "NAME:Attributes",
    "Name:=", "GroundPlane",
    "Color:=", "(255 165 0)",
    "MaterialName:=", "copper",
    "SolveInside:=", False
])

# -----------------------------------------------------
# Stepped taper helper (symmetric, sheet rectangles)
# -----------------------------------------------------
def create_symmetric_stepped_taper(center_x, y_start, z_start, L_taper, w_start, w_end, h, steps, name_prefix):
    step_length = L_taper / steps
    delta_w = (w_start - w_end) / float(steps)
    names = []
    for i in range(steps):
        w_i = w_start - i * delta_w
        if w_i < w_end:
            w_i = w_end
        x_start = center_x - w_i / 2
        y0 = y_start + i * step_length
        name = name_prefix + "_Step" + str(i+1)
        oEditor.CreateRectangle([
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", mm(x_start),
            "YStart:=", mm(y0),
            "ZStart:=", mm(z_start),
            "Width:=", mm(w_i),
            "Height:=", mm(step_length),
            "WhichAxis:=", "Z"
        ], [
            "NAME:Attributes",
            "Name:=", name,
            "Color:=", "(255 255 0)",
            "MaterialName:=", "gold",
            "SolveInside:=", False
        ])
        names.append(name)
    return names

# -----------------------------------------------------
# Antenna 1 (Left)
# -----------------------------------------------------
# Parallel segment
oEditor.CreateRectangle([
    "NAME:RectangleParameters",
    "IsCovered:=", True,
    "XStart:=", mm(antenna1_x_start),
    "YStart:=", mm(parallel_y),
    "ZStart:=", mm(sub_h),
    "Width:=", mm(L_parallel),
    "Height:=", mm(trace_thick_w),
    "WhichAxis:=", "Z"
], [
    "NAME:Attributes",
    "Name:=", "Antenna1_Parallel",
    "Color:=", "(255 255 0)",
    "MaterialName:=", "gold",
    "SolveInside:=", False
])

# Perpendicular segment
perp_x1 = antenna1_x_start + L_parallel - trace_thick_w
perp_y1 = parallel_y + trace_thick_w
oEditor.CreateRectangle([
    "NAME:RectangleParameters",
    "IsCovered:=", True,
    "XStart:=", mm(perp_x1),
    "YStart:=", mm(perp_y1),
    "ZStart:=", mm(sub_h),
    "Width:=", mm(trace_thick_w),
    "Height:=", mm(L_perpendicular),
    "WhichAxis:=", "Z"
], [
    "NAME:Attributes",
    "Name:=", "Antenna1_Perpendicular",
    "Color:=", "(255 255 0)",
    "MaterialName:=", "gold",
    "SolveInside:=", False
])

# Corrected miter cut
bend_x1 = antenna1_x_start + L_parallel
bend_y1 = parallel_y
oEditor.CreatePolyline([
    "NAME:PolylineParameters",
    "IsPolylineCovered:=", True,
    "IsPolylineClosed:=", True,
    ["NAME:PolylinePoints",
        ["NAME:PLPoint", "X:=", mm(bend_x1), "Y:=", mm(bend_y1), "Z:=", mm(sub_h)],
        ["NAME:PLPoint", "X:=", mm(bend_x1 - miter_size), "Y:=", mm(bend_y1), "Z:=", mm(sub_h)],
        ["NAME:PLPoint", "X:=", mm(bend_x1), "Y:=", mm(bend_y1 + miter_size), "Z:=", mm(sub_h)]
    ],
    ["NAME:PolylineSegments",
        ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 0, "EndIndex:=", 1],
        ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 1, "EndIndex:=", 2],
        ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 2, "EndIndex:=", 0]
    ],
    ["NAME:PolylineXSection", "XSectionType:=", "None"]
], [
    "NAME:Attributes",
    "Name:=", "Antenna1_MiterCut",
    "MaterialName:=", "gold",
    "Color:=", "(255 0 0)",
    "SolveInside:=", False
])
oEditor.Subtract([
    "NAME:Selections",
    "Blank Parts:=", "Antenna1_Parallel",
    "Tool Parts:=", "Antenna1_MiterCut"
], ["NAME:SubtractParameters", "KeepOriginals:=", False])

# Taper
taper1_names = create_symmetric_stepped_taper(perp_x1 + trace_thick_w/2,
                                              perp_y1 + L_perpendicular,
                                              sub_h,
                                              L_taper, trace_thick_w, trace_thin_w,
                                              trace_h, taper_steps, "Antenna1_Taper")

# -----------------------------------------------------
# Antenna 2 (Right)
# -----------------------------------------------------
oEditor.CreateRectangle([
    "NAME:RectangleParameters",
    "IsCovered:=", True,
    "XStart:=", mm(antenna2_x_start),
    "YStart:=", mm(parallel_y),
    "ZStart:=", mm(sub_h),
    "Width:=", mm(L_parallel),
    "Height:=", mm(trace_thick_w),
    "WhichAxis:=", "Z"
], [
    "NAME:Attributes",
    "Name:=", "Antenna2_Parallel",
    "Color:=", "(255 255 0)",
    "MaterialName:=", "gold",
    "SolveInside:=", False
])

# Perpendicular
perp_x2 = antenna2_x_start
perp_y2 = parallel_y + trace_thick_w
oEditor.CreateRectangle([
    "NAME:RectangleParameters",
    "IsCovered:=", True,
    "XStart:=", mm(perp_x2),
    "YStart:=", mm(perp_y2),
    "ZStart:=", mm(sub_h),
    "Width:=", mm(trace_thick_w),
    "Height:=", mm(L_perpendicular),
    "WhichAxis:=", "Z"
], [
    "NAME:Attributes",
    "Name:=", "Antenna2_Perpendicular",
    "Color:=", "(255 255 0)",
    "MaterialName:=", "gold",
    "SolveInside:=", False
])

# Corrected miter cut
bend_x2 = antenna2_x_start
bend_y2 = parallel_y
oEditor.CreatePolyline([
    "NAME:PolylineParameters",
    "IsPolylineCovered:=", True,
    "IsPolylineClosed:=", True,
    ["NAME:PolylinePoints",
        ["NAME:PLPoint", "X:=", mm(bend_x2), "Y:=", mm(bend_y2), "Z:=", mm(sub_h)],
        ["NAME:PLPoint", "X:=", mm(bend_x2 + miter_size), "Y:=", mm(bend_y2), "Z:=", mm(sub_h)],
        ["NAME:PLPoint", "X:=", mm(bend_x2), "Y:=", mm(bend_y2 + miter_size), "Z:=", mm(sub_h)]
    ],
    ["NAME:PolylineSegments",
        ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 0, "EndIndex:=", 1],
        ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 1, "EndIndex:=", 2],
        ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 2, "EndIndex:=", 0]
    ],
    ["NAME:PolylineXSection", "XSectionType:=", "None"]
], [
    "NAME:Attributes",
    "Name:=", "Antenna2_MiterCut",
    "MaterialName:=", "gold",
    "Color:=", "(255 0 0)",
    "SolveInside:=", False
])
oEditor.Subtract([
    "NAME:Selections",
    "Blank Parts:=", "Antenna2_Parallel",
    "Tool Parts:=", "Antenna2_MiterCut"
], ["NAME:SubtractParameters", "KeepOriginals:=", False])

# Taper
taper2_names = create_symmetric_stepped_taper(perp_x2 + trace_thick_w/2,
                                              perp_y2 + L_perpendicular,
                                              sub_h,
                                              L_taper, trace_thick_w, trace_thin_w,
                                              trace_h, taper_steps, "Antenna2_Taper")

# -----------------------------------------------------
# Unite geometries
# -----------------------------------------------------
oEditor.Unite([
    "NAME:Selections",
    "Selections:=", "Antenna1_Parallel,Antenna1_Perpendicular," + ",".join(taper1_names)
], ["NAME:UniteParameters", "KeepOriginals:=", False])
oEditor.ChangeProperty([
    "NAME:AllTabs",
    ["NAME:Geometry3DAttributeTab",
        ["NAME:PropServers", "Antenna1_Parallel"],
        ["NAME:ChangedProps", ["NAME:Name", "Value:=", "Antenna1"]]
    ]
])

oEditor.Unite([
    "NAME:Selections",
    "Selections:=", "Antenna2_Parallel,Antenna2_Perpendicular," + ",".join(taper2_names)
], ["NAME:UniteParameters", "KeepOriginals:=", False])
oEditor.ChangeProperty([
    "NAME:AllTabs",
    ["NAME:Geometry3DAttributeTab",
        ["NAME:PropServers", "Antenna2_Parallel"],
        ["NAME:ChangedProps", ["NAME:Name", "Value:=", "Antenna2"]]
    ]
])

# -----------------------------------------------------
# Lumped Ports
# -----------------------------------------------------
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", mm(12.75),
     "YStart:=", mm(-10.75),
     "ZStart:=", mm(1.0),
     "Width:=", mm(1.0),
     "Height:=", mm(-1.0),
     "WhichAxis:=", "X"],
    ["NAME:Attributes",
     "Name:=", "Port1_Rectangle",
     "MaterialName:=", "vacuum",
     "Color:=", "(0 255 0)",
     "Transparency:=", 0]
)

oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", mm(-12.75),
     "YStart:=", mm(-9.75),
     "ZStart:=", mm(1.0),
     "Width:=", mm(-1.0),
     "Height:=", mm(-1.0),
     "WhichAxis:=", "X"],
    ["NAME:Attributes",
     "Name:=", "Port2_Rectangle",
     "MaterialName:=", "vacuum",
     "Color:=", "(0 255 0)",
     "Transparency:=", 0]
)

# -----------------------------------------------------
# Boundaries
# -----------------------------------------------------
oBoundarySetup = oDesign.GetModule("BoundarySetup")

# Radiation box
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", mm(-100),
     "YPosition:=", mm(-100),
     "ZPosition:=", mm(-100),
     "XSize:=", mm(200),
     "YSize:=", mm(200),
     "ZSize:=", mm(200)],
    ["NAME:Attributes",
     "Name:=", "RadiationBox",
     "MaterialName:=", "air",
     "Color:=", "(128 128 128)",
     "Transparency:=", 1]
)
oBoundarySetup.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["RadiationBox"],
     "IsFssReference:=", False,
     "IsForPML:=", False]
)

# Finite conductivity
oModule = oDesign.GetModule("BoundarySetup")
for trace in ["Antenna1", "Antenna2"]:
    oModule.AssignFiniteCond([
        "NAME:" + trace,
        "Objects:=", [trace],
        "UseMaterial:=", True,
        "Material:=", "gold",
        "Thickness:=", mm(trace_h),
        "IsTwoSided:=", False,
        "InfGroundPlane:=", False
    ])
oModule.AssignFiniteCond([
    "NAME:GroundPlane",
    "Objects:=", ["GroundPlane"],
    "UseMaterial:=", True,
    "Material:=", "copper",
    "Thickness:=", mm(trace_h),
    "IsTwoSided:=", False,
    "InfGroundPlane:=", False
])

# Lumped port assignments
oBoundarySetup.AssignLumpedPort(
    ["NAME:Port1",
     "Objects:=", ["Port1_Rectangle"],
     "RenormalizeAllTerminals:=", True,
     "DoDeembed:=", False,
     ["NAME:Modes",
      ["NAME:Mode1",
       "ModeNum:=", 1,
       "UseIntLine:=", True,
       ["NAME:IntLine",
        "Start:=", ["12.75mm", "-10.75mm", "1.0mm"],
        "End:=", ["12.75mm", "-10.75mm", "-0.0mm"]],
       "CharImp:=", "Zpi",
       "AlignmentGroup:=", 0,
       "RenormImp:=", "50ohm"]],
     "ShowReporterFilter:=", False,
     "ReporterFilter:=", [True]]
)
oBoundarySetup.AssignLumpedPort(
    ["NAME:Port2",
     "Objects:=", ["Port2_Rectangle"],
     "RenormalizeAllTerminals:=", True,
     "DoDeembed:=", False,
     ["NAME:Modes",
      ["NAME:Mode1",
       "ModeNum:=", 1,
       "UseIntLine:=", True,
       ["NAME:IntLine",
        "Start:=", ["-12.75mm", "-9.75mm", "1.0mm"],
        "End:=", ["-12.75mm", "-9.75mm", "-0.0mm"]],
       "CharImp:=", "Zpi",
       "AlignmentGroup:=", 0,
       "RenormImp:=", "50ohm"]],
     "ShowReporterFilter:=", False,
     "ReporterFilter:=", [True]]
)

# -----------------------------------------------------
# Analysis Setup
# -----------------------------------------------------
oModule = oDesign.GetModule("AnalysisSetup")

# Setup: tighten convergence + allow more passes
oModule.InsertSetup("HfssDriven",
    [
        "NAME:Setup1",
        "Frequency:=", "3GHz",
        "MaxDeltaS:=", 0.005,          # was 0.02, now stricter
        "MaximumPasses:=", 25,         # was 10, more adaptive passes
        "MinimumPasses:=", 2,          # at least 2 passes
        "MinimumConvergedPasses:=", 2, # require convergence twice
        "PercentRefinement:=", 20,     # smaller refinement step
        "IsEnabled:=", True
    ]
)

# Sweep
oModule.InsertFrequencySweep("Setup1",
    [
        "NAME:Sweep",
        "IsEnabled:=", True,
        "RangeType:=", "LinearStep",
        "RangeStart:=", "2GHz",
        "RangeEnd:=", "4GHz",
        "RangeStep:=", "10MHz",
        "Type:=", "Discrete",    
        "SaveFields:=", False,
        "SaveRadFields:=", False,
        "InterpUseS:=", True,         
        "InterpUsePortImped:=", True,
        "InterpUsePropConst:=", True
    ]
)


# -----------------------------------------------------
# Run Simulation
# -----------------------------------------------------
print("Running all analysis setups...")
oDesign.AnalyzeAll()   

print("Simulation completed!")

# -----------------------------------------------------
# Extract S-Parameters and Save to CSV in Desktop Folder
# -----------------------------------------------------
# -----------------------------------------------------
# Extract S-Parameters and Save to CSV
# -----------------------------------------------------
save_folder = r"C:\Users\kaust\OneDrive\Desktop\HFSS_S_Params_Export_Data"
if not os.path.exists(save_folder):
    os.makedirs(save_folder)

csv_file = os.path.join(save_folder, "SParameters.csv")

oReportSetup = oDesign.GetModule("ReportSetup")
oReportSetup.CreateReport(
    "SParameters", "Modal Solution Data", "Rectangular Plot", "Setup1 : Sweep",
    ["Domain:=", "Sweep"],
    ["Freq:=", ["All"]],
    [
        "X Component:=", "Freq",
        "Y Component:=", ["dB(S(1,1))", "dB(S(1,2))", "dB(S(2,1))", "dB(S(2,2))"]
    ]
)

oReportSetup.ExportToFile("SParameters", csv_file)
print("S-parameters exported to:", csv_file)

# -----------------------------------------------------
# Finalize
# -----------------------------------------------------
oEditor.FitAll()
print("Dual microstrip antenna geometry united with ground plane sheet, symmetric stepped taper, corrected miter cuts, and S-parameters exported.")


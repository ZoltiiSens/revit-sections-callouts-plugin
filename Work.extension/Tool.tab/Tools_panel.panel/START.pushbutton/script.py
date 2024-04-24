# ------------------------------ IMPORTS ------------------------------
import sys

from Autodesk.Revit.ApplicationServices import *
# from Autodesk.Revit.DB.__init___parts.XYZ import XYZ
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from pyrevit import forms

# ------------------------------ VARIABLES ------------------------------

__title__ = "Sections automation"
__doc__ = """
Plugin which creates 3 sections depending on 
"""
uidoc = __revit__.ActiveUIDocument          # type: UIDocument
doc = __revit__.ActiveUIDocument.Document   # type: Document
app = __revit__.Application                 # type: Application
print('HI!')

# ------------------------------ MAIN ------------------------------

# windows = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()
# dict_windows = {}
# cal = FilteredElementCollector(doc, doc.ActiveView.Id).ToElements()
# print(cal)
# print('------')

selection = uidoc.Selection.GetElementIds()
if len(selection) != 1:
    # Provide a message to the user
    TaskDialog.Show("Selection Error", "Please select only one element. Plugin will now stop.")
    sys.exit()

windowFamilyObject = doc.GetElement(selection[0])  # type: FamilyInstance
print("Selected Element:", windowFamilyObject.Symbol.Family.Name)

# print('childs:')
# for el in windowFamilyObject.GetDependentElements(None):
#     elemet = doc.GetElement(el)
#     print(elemet.GetDependentElements(None))
#     print(elemet)
#     print(el)


transaction = Transaction(doc, 'Generate Window Section')
transaction.Start()

window_origin = windowFamilyObject.Location.Point   # type: XYZ
host_object = windowFamilyObject.Host
curve = host_object.Location.Curve                  # type: Curve
pt_start = curve.GetEndPoint(0)                     # type: XYZ
pt_end = curve.GetEndPoint(1)                       # type: XYZ
vector = pt_end - pt_start                          # type: XYZ

win_height = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.GENERIC_HEIGHT).AsDouble()
win_width = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble()
win_depth = UnitUtils.ConvertToInternalUnits(40, UnitTypeId.Centimeters)
offset = UnitUtils.ConvertToInternalUnits(1, UnitTypeId.Centimeters)
vector = vector.Normalize()
wallDepth = UnitUtils.ConvertToInternalUnits(host_object.Width, UnitTypeId.Feet)

wall_level_id = host_object.LevelId

# Get the actual Level element using the level ID
if wall_level_id:
    wall_level = doc.GetElement(wall_level_id)  # Get the Level element
else:
    print("Wall does not have an associated level.")
# 1. Get the wall's level
# wall_level = host_object.BaseLevel

# 2. Get all floor elements in the document
floor_collector = FilteredElementCollector(doc).OfClass(Floor)

# 3. Filter floors based on level relationship with the wall's level
upper_floor = None
lower_floor = None
for floor in floor_collector:
    floor_level_id = floor.LevelId
    if floor_level_id:
        floor_level = doc.GetElement(floor_level_id)  # Get the Level element
    else:
        print("floor does not have an associated level.")
    if floor_level and wall_level:  # Check for level existence
        print('Elevations:')
        print(floor.Name)
        print(floor_level.Elevation)
        print(wall_level.Elevation)
        print('-----------------------------')
        if floor_level.Elevation > wall_level.Elevation:
            upper_floor = floor
            # break  # Stop searching once an upper floor is found
        elif floor_level.Elevation <= wall_level.Elevation:
            lower_floor = floor

# 4. Print the information about the upper and lower floors (if found)
if upper_floor:
    print("Upper Floor:", upper_floor.Name)
else:
    print("No upper floor found for the wall.")

if lower_floor:
    print("Lower Floor:", lower_floor.Name)
else:
    print("No lower floor found for the wall.")



# wall_location = host_object.Location.Curve
#
# # Get the endpoints of the wall
# wall_start = wall_location.GetEndPoint(0)
# wall_end = wall_location.GetEndPoint(1)
#
# print('startend')
# print(wall_start, wall_end)
#
# # Find the floors above and below the wall
# floor_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType()
#
# upper_floor = None
# lower_floor = None
# print('floors:')
# print(floor_collector)
#
# for floor in floor_collector:
#     # Get the elevation of the floor
#     floor_elevation = doc.GetElement(floor.LevelId).Elevation + floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble()
#
#     # Print floor elevation for debugging
#     print("Floor Elevation:", floor_elevation)
#     print("min", min(wall_start.Z, wall_end.Z), "max", max(wall_start.Z, wall_end.Z))
#     # Check if the floor is above or below the wall
#     if min(wall_start.Y, wall_end.Y) <= floor_elevation <= max(wall_start.Y, wall_end.Y):
#         upper_floor = floor if floor_elevation > max(wall_start.Y, wall_end.Y) else upper_floor
#         lower_floor = floor if floor_elevation < min(wall_start.Y, wall_end.Y) else lower_floor
#
# # Print results after the loop
# print("Upper Floor:", upper_floor)
# print("Lower Floor:", lower_floor)

# for floor in floor_collector:
#     print(floor)
#     # Get the elevation of the floor
#     floor_elevation = doc.GetElement(floor.LevelId).Elevation + floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble()
#
#     # Check if the floor is above or below the wall
#     if min(wall_start.Z, wall_end.Z) <= floor_elevation <= max(wall_start.Z, wall_end.Z):
#         upper_floor = floor if floor_elevation > max(wall_start.Z, wall_end.Z) else upper_floor
#         lower_floor = floor if floor_elevation < min(wall_start.Z, wall_end.Z) else lower_floor
#
# print(upper_floor)
# print(lower_floor)







# ---------------------------------------------------------------------------------------------------------- Front view
# transform = Transform.Identity
# transform.Origin = window_origin
# transform.BasisX = vector
# transform.BasisY = XYZ.BasisZ
# transform.BasisZ = vector.CrossProduct(XYZ.BasisZ)
# section_box = BoundingBoxXYZ()
# # section_box.Min = XYZ(-win_width / 2 - offset, 0 - offset, -win_depth)
# # section_box.Max = XYZ(win_width / 2 + offset, win_height + offset, win_depth)
# # section_box.Min = XYZ(-win_width, 0 - offset, -win_depth)
# # section_box.Max = XYZ(win_width, win_height + offset, win_depth)
# section_box.Min = XYZ(-win_width - (offset * 50), 0 - (offset * 50), -win_depth)
# section_box.Max = XYZ(win_width + (offset * 50), win_height + (offset * 50), win_depth)
# section_box.Transform = transform
# section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
# win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
# # new_name = 'py_{}'.format(windowFamilyObject.Symbol.Family.Name)
# new_name = 'py_MAMAD_Window_Front_View'
#
# for i in range(10):
#     try:
#         win_elevation.Name = new_name
#         print('Created a section: {}'.format(new_name))
#         break
#     except:
#         new_name += '*'


# ----------------------------------------------------------------------------------------------------- From down to up     SHOULD BE CALLOUT
# print(type(host_object))
# wallDepth = UnitUtils.ConvertToInternalUnits(host_object.Width, UnitTypeId.Feet)
#
# transform = Transform.Identity
# transform.Origin = window_origin
# transform.BasisX = vector
# transform.BasisY = XYZ.BasisZ.CrossProduct(vector)
# transform.BasisZ = XYZ.BasisZ
# section_box = BoundingBoxXYZ()
# # section_box.Min = XYZ(-win_width - UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters), -(wallDepth / 2 + UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)), win_height / 2)
# # section_box.Max = XYZ(win_width + UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters), (wallDepth / 2) + UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters), win_height / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters))
# section_box.Min = XYZ(-win_width - UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters), -(wallDepth / 2 + UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)), win_height / 2)
# section_box.Max = XYZ(win_width + UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters), (wallDepth / 2) + UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters), win_height / 2 + UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters))
# section_box.Transform = transform
# section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
# views = FilteredElementCollector(doc).OfClass(View).ToElements()
# viewTemplates = [v for v in views if v.IsTemplate and "Section_Reinforcement" in v.Name.ToString()]
# win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
# win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
# # new_name = 'py_{}'.format(windowFamilyObject.Symbol.Family.Name)
# new_name = 'py_MAMAD_Window_Section_1'
# for i in range(10):
#     try:
#         win_elevation.Name = new_name
#         print('Created a section: {}'.format(new_name))
#         break
#     except:
#         new_name += '*'


# ----------------------------------------------------------------------------------------------------- Perpendicular window
transform = Transform.Identity
transform.Origin = window_origin
vector_perp = vector.CrossProduct(XYZ.BasisZ)
transform.BasisX = vector_perp
transform.BasisY = XYZ.BasisZ
transform.BasisZ = vector_perp.CrossProduct(XYZ.BasisZ)
section_box = BoundingBoxXYZ()
# section_box.Min = XYZ(-win_width, -(win_height * 2), -UnitUtils.ConvertToInternalUnits(80, UnitTypeId.Centimeters))
# section_box.Max = XYZ(win_width, (win_height * 2) + offset, UnitUtils.ConvertToInternalUnits(80, UnitTypeId.Centimeters))
offset_20cm = UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)
section_box.Min = XYZ(
    -wallDepth / 2 - offset_50cm, 0 - offset_20cm, -win_width / 2
)
section_box.Max = XYZ(
    wallDepth / 2 + offset_70cm, win_height + offset_20cm, -win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
)

# glybyna(inner-outer) -> height

section_box.Transform = transform
section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
views = FilteredElementCollector(doc).OfClass(View).ToElements()
viewTemplates = [v for v in views if v.IsTemplate and "Section_Reinforcement" in v.Name.ToString()]
win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
# new_name = 'py_{}'.format(windowFamilyObject.Symbol.Family.Name)
new_name = 'py_MAMAD_Window_Section_1'
for i in range(10):
    try:
        win_elevation.Name = new_name
        print('Created a section: {}'.format(new_name))
        break
    except:
        new_name += '*'


# ----------------------------------------------------------------------------------------------------- Perpendicular shelter
# transform = Transform.Identity
# transform.Origin = window_origin
# vector_perp = vector.CrossProduct(XYZ.BasisZ)
# transform.BasisX = vector_perp
# transform.BasisY = XYZ.BasisZ
# transform.BasisZ = vector_perp.CrossProduct(XYZ.BasisZ)
# section_box = BoundingBoxXYZ()
# # section_box.Min = XYZ(-win_width, -(win_height * 2), -UnitUtils.ConvertToInternalUnits(80, UnitTypeId.Centimeters))
# # section_box.Max = XYZ(win_width, (win_height * 2) + offset, UnitUtils.ConvertToInternalUnits(80, UnitTypeId.Centimeters))
# offset_20cm = UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
# offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
# offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)
# section_box.Min = XYZ(
#     -wallDepth / 2 - offset_50cm, 0 - offset_20cm, win_width / 2
# )
# section_box.Max = XYZ(
#     wallDepth / 2 + offset_70cm, win_height + offset_20cm, win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
# )
#
# # glybyna(inner-outer) -> height
#
# section_box.Transform = transform
# section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
# views = FilteredElementCollector(doc).OfClass(View).ToElements()
# viewTemplates = [v for v in views if v.IsTemplate and "Section_Reinforcement" in v.Name.ToString()]
# win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
# win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
# # new_name = 'py_{}'.format(windowFamilyObject.Symbol.Family.Name)
# new_name = 'py_MAMAD_Window_Section_2'
# for i in range(10):
#     try:
#         win_elevation.Name = new_name
#         print('Created a section: {}'.format(new_name))
#         break
#     except:
#         new_name += '*'


transaction.Commit()
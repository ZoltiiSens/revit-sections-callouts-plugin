# ------------------------------ IMPORTS ------------------------------
import sys

from Autodesk.Revit.ApplicationServices import *
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




# upper_floor_height = upper_floor.Parameters.LookupParameter("Core Thickness").AsDouble()
# print(upper_floor_height)

# print('Elevations:')
# print(floor.Name)
# print(floor_level.Elevation)
# print(wall_level.Elevation)
# print('MINUS:')
# print(floor_level.Elevation - wall_level.Elevation)
# print(type(floor_level.Elevation - wall_level.Elevation))
# print('-----------------------------')


# def get_distance_to_floor(windowFamilyObject, host_object, upper_floor):
#     """
#     Calculates the distance in centimeters from the top of the window to the top of the floor.
#
#     Args:
#       windowFamilyObject (Revit.DB.Element): The window element.
#       host_object (Revit.DB.Element): The wall element hosting the window.
#       upper_floor (Revit.DB.Element): The top floor element.
#
#     Returns:
#       float: The distance in centimeters from the top of the window to the top of the floor,
#              or None if any element is invalid.
#     """
#     # Check if all elements are valid
#     if not windowFamilyObject or not host_object or not upper_floor:
#         return None
#
#     # Get window location
#     # window_location = windowFamilyObject.Location.Curve
#
#     window_location = host_object.Location.Curve
#
#     if isinstance(window_location, LocationCurve):
#         window_top_xyz = window_location.Curve.GetEndPoint(1)
#     else:
#         # Handle LocationPoint case (provide alternative logic here)
#         print("Window location is not a curve. Distance calculation not applicable.")
#         return None
#
#     # Get wall bottom level
#     if isinstance(host_object, RevitLinkInstance):
#         # Linked element - get its level from its properties
#         wall_level_id = host_object.GetLinkInfo().CorrespondingRevitElement.LevelId
#     else:
#         # Regular wall element - get level from base constraint
#         wall_base_constraint = host_object.GetConstraints(XYZ(0, 0, 1))[0]
#         wall_level_id = wall_base_constraint.LevelId
#
#     # Get wall level elevation
#     wall_level = doc.GetElement(wall_level_id)
#     wall_level_elevation = wall_level.Level.Elevation
#
#     # Get floor top level elevation
#     floor_level = upper_floor.Level
#     floor_level_elevation = floor_level.Elevation
#
#     # Calculate window top elevation (assuming window is placed vertically)
#     window_top_xyz = window_location.GetEndPoint(1)
#     window_top_elevation = window_top_xyz.Z
#
#     # Calculate distance from window top to floor top
#     distance_in_meters = window_top_elevation - (wall_level_elevation + floor_level_elevation)
#
#     # Convert distance to centimeters and return
#     distance_in_cm = distance_in_meters * 100
#     return distance_in_cm
#
#
# # Example usage (assuming you have retrieved the elements from Revit)
# distance_to_floor = get_distance_to_floor(windowFamilyObject, host_object, upper_floor)
#
# if distance_to_floor is not None:
#     print("Distance from window top to floor top", distance_to_floor)
# else:
#     print("Failed to calculate distance. Please check element validity.")



# def get_height(element):
#     bounding_box = element.get_BoundingBox(None)
#     if bounding_box is not None:
#         min_point = bounding_box.Min
#         max_point = bounding_box.Max
#         height = max_point.Z - min_point.Z
#         return height
#     else:
#         return 0.0
#
#
# def calculate_distance(window, wall, top_floor):
#     window_height = get_height(window)
#     wall_height = get_height(wall)
#     floor_level_id = top_floor.LevelId
#     floor_level = doc.GetElement(floor_level_id)
#     floor_to_ceiling_height = floor_level.Elevation - wall.Location.Curve.GetEndPoint(1).Z
#     distance = floor_to_ceiling_height - window_height
#     # distance_in_cm = distance * 100 # Convert meters to centimeters
#     distance_in_cm = UnitUtils.ConvertToInternalUnits(distance, UnitTypeId.Feet)
#     return distance_in_cm
#
#
# distance_to_top_floor_cm = calculate_distance(windowFamilyObject, host_object, upper_floor)
# print("Distance from top of the window to top part of the floor:", distance_to_top_floor_cm, "centimeters")


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


# # ----------------------------------------------------------------------------------------------------- Perpendicular window
# # Getting floors
# floor_collector = FilteredElementCollector(doc).OfClass(Floor)
# upper_floor = None
# lower_floor = None
# upper_floor_level_best_height = float('inf')
# lower_floor_level_best_height = float('-inf')
# for floor in floor_collector:
#     floor_level_id = floor.LevelId
#     if floor_level_id:
#         floor_level = doc.GetElement(floor_level_id)
#     else:
#         print("floor does not have an associated level.")
#     if floor_level and wall_level:
#         if floor_level.Elevation > wall_level.Elevation:
#             if floor_level.Elevation < upper_floor_level_best_height:
#                 upper_floor_level_best_height = floor_level.Elevation
#                 upper_floor = floor
#         elif floor_level.Elevation <= wall_level.Elevation:
#             if floor_level.Elevation > lower_floor_level_best_height:
#                 lower_floor_level_best_height = floor_level.Elevation
#                 lower_floor = floor
#
# if upper_floor:
#     print("Upper Floor:", upper_floor.Name)
#     lalala = upper_floor.Parameters
#     upper_floor_height = None
#     for param in lalala:
#         if param.Definition.Name == 'Core Thickness':
#             upper_floor_height = param.AsDouble()
#             break
#     if upper_floor_height is not None:
#         upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(upper_floor_height * 30.48,
#                                                                       UnitTypeId.Centimeters)
#         print('Top floor height in cm', upper_floor_height * 30.48)
#     else:
#         upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#     window_bbox = windowFamilyObject.get_BoundingBox(None)  # False for local coordinate system
#     floor_bbox = upper_floor.get_BoundingBox(None)
#     window_top_z = window_bbox.Max.Z
#     floor_bottom_z = floor_bbox.Min.Z
#     distance_z = floor_bottom_z - window_top_z
#     top_floor_distance_cm = UnitUtils.ConvertToInternalUnits(distance_z * 30.48, UnitTypeId.Centimeters)
#     print('Distance to top floor in cm', distance_z * 30.48)
#
# else:
#     print("No upper floor found for the wall.")
# if lower_floor:
#     print("Lower Floor:", lower_floor.Name)
#     lalala = lower_floor.Parameters
#     for param in lalala:
#         if param.Definition.Name == 'Core Thickness':
#             lower_floor_height = param.AsDouble()
#             break
#     if lower_floor_height is not None:
#         lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(lower_floor_height * 30.48,
#                                                                       UnitTypeId.Centimeters)
#         print(lower_floor_height * 30.48)
#     else:
#         lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#     window_bbox = windowFamilyObject.get_BoundingBox(None)  # False for local coordinate system
#     floor_bbox = lower_floor.get_BoundingBox(None)
#     window_top_z = window_bbox.Min.Z
#     floor_bottom_z = floor_bbox.Max.Z
#     distance_z = window_top_z - floor_bottom_z
#     bottom_floor_distance_cm = UnitUtils.ConvertToInternalUnits(distance_z * 30.48, UnitTypeId.Centimeters)
#     print('Distance to bottom floor in cm', distance_z * 30.48)
# else:
#     print("No lower floor found for the wall.")
#
#
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
# offset_30cm = UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)
# offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
# offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)
# offset_400cm = UnitUtils.ConvertToInternalUnits(400, UnitTypeId.Centimeters)
# # section_box.Min = XYZ(
# #     -wallDepth / 2 - offset_50cm, 0 - offset_20cm, -win_width / 2
# # )
# # section_box.Max = XYZ(
# #     wallDepth / 2 + offset_70cm, win_height + offset_20cm, -win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
# # )
#
# section_box.Min = XYZ(
#     -wallDepth / 2 - offset_50cm, 0 - bottom_floor_distance_cm - lower_floor_in_centimeters - offset_30cm, -win_width / 2
# ) # To bottom
# section_box.Max = XYZ(
#     wallDepth / 2 + offset_70cm, win_height + top_floor_distance_cm + upper_floor_in_centimeters + offset_30cm, -win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
# ) # To top
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
# new_name = 'py_MAMAD_Window_Section_1'
# for i in range(10):
#     try:
#         win_elevation.Name = new_name
#         print('Created a section: {}'.format(new_name))
#         break
#     except:
#         new_name += '*'


# ----------------------------------------------------------------------------------------------------- Perpendicular shelter
# floor_collector = FilteredElementCollector(doc).OfClass(Floor)
# upper_floor = None
# lower_floor = None
# upper_floor_level_best_height = float('inf')
# lower_floor_level_best_height = float('-inf')
# for floor in floor_collector:
#     floor_level_id = floor.LevelId
#     if floor_level_id:
#         floor_level = doc.GetElement(floor_level_id)
#     else:
#         print("floor does not have an associated level.")
#     if floor_level and wall_level:
#         if floor_level.Elevation > wall_level.Elevation:
#             if floor_level.Elevation < upper_floor_level_best_height:
#                 upper_floor_level_best_height = floor_level.Elevation
#                 upper_floor = floor
#         elif floor_level.Elevation <= wall_level.Elevation:
#             if floor_level.Elevation > lower_floor_level_best_height:
#                 lower_floor_level_best_height = floor_level.Elevation
#                 lower_floor = floor
#
# if upper_floor:
#     print("Upper Floor:", upper_floor.Name)
#     lalala = upper_floor.Parameters
#     upper_floor_height = None
#     for param in lalala:
#         if param.Definition.Name == 'Core Thickness':
#             upper_floor_height = param.AsDouble()
#             break
#     if upper_floor_height is not None:
#         upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(upper_floor_height * 30.48,
#                                                                       UnitTypeId.Centimeters)
#         print('Top floor height in cm', upper_floor_height * 30.48)
#     else:
#         upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#     window_bbox = windowFamilyObject.get_BoundingBox(None)  # False for local coordinate system
#     floor_bbox = upper_floor.get_BoundingBox(None)
#     window_top_z = window_bbox.Max.Z
#     floor_bottom_z = floor_bbox.Min.Z
#     distance_z = floor_bottom_z - window_top_z
#     top_floor_distance_cm = UnitUtils.ConvertToInternalUnits(distance_z * 30.48, UnitTypeId.Centimeters)
#     print('Distance to top floor in cm', distance_z * 30.48)
#
# else:
#     print("No upper floor found for the wall.")
# if lower_floor:
#     print("Lower Floor:", lower_floor.Name)
#     lalala = lower_floor.Parameters
#     for param in lalala:
#         if param.Definition.Name == 'Core Thickness':
#             lower_floor_height = param.AsDouble()
#             break
#     if lower_floor_height is not None:
#         lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(lower_floor_height * 30.48,
#                                                                       UnitTypeId.Centimeters)
#         print('Height of lower floor in cm', lower_floor_height * 30.48)
#     else:
#         lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#     window_bbox = windowFamilyObject.get_BoundingBox(None)  # False for local coordinate system
#     floor_bbox = lower_floor.get_BoundingBox(None)
#     window_top_z = window_bbox.Min.Z
#     floor_bottom_z = floor_bbox.Max.Z
#     distance_z = window_top_z - floor_bottom_z
#     bottom_floor_distance_cm = UnitUtils.ConvertToInternalUnits(distance_z * 30.48, UnitTypeId.Centimeters)
#     print('Distance to bottom floor in cm', distance_z * 30.48)
# else:
#     print("No lower floor found for the wall.")
#
#
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
# offset_30cm = UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)
# offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
# offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)
# section_box.Min = XYZ(
#     -wallDepth / 2 - offset_50cm, 0 - bottom_floor_distance_cm - lower_floor_in_centimeters - offset_30cm, win_width / 2
# )
# section_box.Max = XYZ(
#     wallDepth / 2 + offset_70cm, win_height + top_floor_distance_cm + upper_floor_in_centimeters + offset_30cm, win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
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
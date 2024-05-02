# ------------------------------ IMPORTS ------------------------------
import math
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


# ------------------------------ Finctions ------------------------------
def find_upper_window(wall_lvl_id, current_window):
    """
    Finds a window directly above the given window on the level above and calculates distance.
    Args:
        wall_lvl_id (ElementId): The ID of the level the wall belongs to.
        current_window (FamilyInstance): The current window element.
    Returns:
        FamilyInstance: The window directly above, if found; None otherwise.
    """
    wall_level = doc.GetElement(wall_lvl_id)
    upper_level = None
    all_levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).ToElements()
    for level in all_levels:
        if type(level) == LevelType:
            continue
        if level.Elevation > wall_level.Elevation:
            if not level.Name.startswith('Floor'):
                continue
            upper_level = level
            break
    if upper_level:
        elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        current_window_bbox = current_window.get_BoundingBox(None)
        current_window_x, current_window_y = current_window_bbox.Min.X, current_window_bbox.Min.Y
        upperWindow = None
        distance = float('inf')
        for elem in elements:
            if (elem.Category is not None and
                    elem.Category.Id == current_window.Category.Id and
                    elem.LevelId == upper_level.Id):
                if upperWindow is None:
                    upperWindow = elem
                    continue
                elem_bbox = elem.get_BoundingBox(None)
                elem_x, elem_y = elem_bbox.Min.X, elem_bbox.Min.Y
                new_distance = math.sqrt((elem_x - current_window_x) ** 2 + (elem_y - current_window_y) ** 2)
                if new_distance < distance:
                    distance = new_distance
                    upperWindow = elem
        if upperWindow is not None:
            print('Upper window:', upperWindow.Name)
            return UnitUtils.ConvertToInternalUnits(
                (upperWindow.get_BoundingBox(None).Min.Z - current_window.get_BoundingBox(None).Max.Z) * 30.48,
                UnitTypeId.Centimeters)
    return None


def find_floors_offsets(wall_lvl):
    floor_collector = FilteredElementCollector(doc).OfClass(Floor)
    upper_floor, lower_floor = None, None
    upper_floor_level_best_height, lower_floor_level_best_height = float('inf'), float('-inf')
    for floor in floor_collector:
        floor_level_id = floor.LevelId
        floor_level = None
        if floor_level_id:
            floor_level = doc.GetElement(floor_level_id)
        if floor_level and wall_lvl:
            if floor_level.Elevation > wall_lvl.Elevation:
                if floor_level.Elevation < upper_floor_level_best_height:
                    upper_floor_level_best_height = floor_level.Elevation
                    upper_floor = floor
            elif floor_level.Elevation <= wall_lvl.Elevation:
                if floor_level.Elevation > lower_floor_level_best_height:
                    lower_floor_level_best_height = floor_level.Elevation
                    lower_floor = floor
    if upper_floor:
        print("Upper Floor:", upper_floor.Name)
        upper_floor_params = upper_floor.Parameters
        upper_floor_height = None
        for param in upper_floor_params:
            if param.Definition.Name == 'Core Thickness':
                upper_floor_height = param.AsDouble()
                break
        if upper_floor_height is not None:
            upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(upper_floor_height * 30.48,
                                                                          UnitTypeId.Centimeters)
            # print('Top floor height in cm', upper_floor_height * 30.48)
        else:
            upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
        window_bbox = windowFamilyObject.get_BoundingBox(None)  # False for local coordinate system
        floor_bbox = upper_floor.get_BoundingBox(None)
        window_top_z = window_bbox.Max.Z
        floor_bottom_z = floor_bbox.Min.Z
        distance_z = floor_bottom_z - window_top_z
        upper_floor_distance_cm = UnitUtils.ConvertToInternalUnits(distance_z * 30.48, UnitTypeId.Centimeters)
        print('Distance to upper floor in cm', distance_z * 30.48)
    else:
        upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
        upper_floor_distance_cm = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
        print("No upper floor found for the wall")
    if lower_floor:
        print("Lower Floor:", lower_floor.Name)
        lower_floor_params = lower_floor.Parameters
        lower_floor_height = None
        for param in lower_floor_params:
            if param.Definition.Name == 'Core Thickness':
                lower_floor_height = param.AsDouble()
                break
        if lower_floor_height is not None:
            lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(lower_floor_height * 30.48, UnitTypeId.Centimeters)
            # print('Height of lower floor in cm', lower_floor_height * 30.48)
        else:
            lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
        window_bbox = windowFamilyObject.get_BoundingBox(None)  # False for local coordinate system
        floor_bbox = lower_floor.get_BoundingBox(None)
        window_top_z = window_bbox.Min.Z
        floor_bottom_z = floor_bbox.Max.Z
        distance_z = window_top_z - floor_bottom_z
        lower_floor_distance_cm = UnitUtils.ConvertToInternalUnits(distance_z * 30.48, UnitTypeId.Centimeters)
        print('Distance to lower floor in cm', distance_z * 30.48)
    else:
        lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
        lower_floor_distance_cm = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
        print("No lower floor found for the wall")
    return lower_floor_distance_cm, lower_floor_in_centimeters, upper_floor_distance_cm, upper_floor_in_centimeters

# ------------------------------ MAIN ------------------------------

selection = uidoc.Selection.GetElementIds()
if len(selection) != 1:
    TaskDialog.Show("Selection Error", "Please select only one element. Plugin will now stop.")
    sys.exit()
windowFamilyObject = doc.GetElement(selection[0])  # type: FamilyInstance
print("Selected Element:", windowFamilyObject.Symbol.Family.Name)

transaction = Transaction(doc, 'Generate Window Sections')
transaction.Start()

window_origin = windowFamilyObject.Location.Point   # type: XYZ
host_object = windowFamilyObject.Host               # type: Wall
curve = host_object.Location.Curve                  # type: Curve
pt_start = curve.GetEndPoint(0)                     # type: XYZ
pt_end = curve.GetEndPoint(1)                       # type: XYZ
vector = pt_end - pt_start                          # type: XYZ
win_height = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.GENERIC_HEIGHT).AsDouble()
win_width = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble()
win_depth = UnitUtils.ConvertToInternalUnits(40, UnitTypeId.Centimeters)
vector = vector.Normalize()
wallDepth = UnitUtils.ConvertToInternalUnits(host_object.Width, UnitTypeId.Feet)
wall_level_id = host_object.LevelId

if wall_level_id:
    wall_level = doc.GetElement(wall_level_id)  # Get the Level element
else:
    print("Wall does not have an associated level. Plugin will now stop.")
    sys.exit()


# def check_upper_floor_window(wall):
#     """
#     Checks if a floor exists above a wall and if that floor contains another window family object.
#
#     Args:
#       wall: A Wall element.
#
#     Returns:
#       A tuple containing two booleans:
#           - has_upper_floor: True if a floor exists above the wall, False otherwise.
#           - has_upper_floor_window: True if the upper floor contains another window family object, False otherwise.
#     """
#
#     if not isinstance(wall, Wall):
#         print("Error: Invalid element type provided.")
#         return None, None
#
#     # try:
#     # Get the wall's level
#     # wall_level = wall.Level
#     wall_level_id = wall.LevelId
#     if wall_level_id:
#         wall_level = doc.GetElement(wall_level_id)
#     # print(wall_level)
#     # Check if there's a level above the wall
#
#     all_levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).ToElements()
#     upper_level = None
#     for level in all_levels:
#         if type(level) == LevelType:
#             continue
#         # print('hahahahahhahahhahahahahahahhahah')
#         # print(level.Elevation)
#         # print(wall_level.Elevation)
#         if level.Elevation > wall_level.Elevation:
#             upper_level = level
#             break  # Stop after finding the first higher level
#
#     # print(upper_level)
#     has_upper_floor = upper_level is not None
#
#     if not has_upper_floor:
#         return has_upper_floor, False  # No upper floor, so no window
#
#     # Check for windows in the upper floor
#     upper_floor_elements = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsElementType().ToElements()
#
#     has_upper_floor_window = False
#     for element in upper_floor_elements:
#         if isinstance(element, FamilyInstance):
#             print(element)
#             # if element.SymbolName.lower() == "window":  # Adjust window family name comparison as needed
#             #     has_upper_floor_window = True
#             #     break  # Stop after finding one window
#     return has_upper_floor, has_upper_floor_window
#
#     # except Exception as e:
#     #     print("An error occurred", e)
#     #     return None, None
#
#
# # Example usage (assuming you have 'wall' defined)
# has_upper_floor, has_upper_floor_window = check_upper_floor_window(host_object)
#
# if has_upper_floor:
#     if has_upper_floor_window:
#         print("There is a floor above the wall and it contains another window.")
#     else:
#         print("There is a floor above the wall, but it does not contain another window.")
# else:
#     print("There is no floor above the wall.")


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
# Getting floors

upper_distance = find_upper_window(wall_level_id, windowFamilyObject)
if upper_distance is None:
    top_offset = None
else:
    top_offset = upper_distance - UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)

(lower_floor_distance_cm,
 lower_floor_in_centimeters,
 upper_floor_distance_cm,
 upper_floor_in_centimeters) = find_floors_offsets(wall_level)

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
offset_30cm = UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)
offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)
offset_400cm = UnitUtils.ConvertToInternalUnits(400, UnitTypeId.Centimeters)

section_box.Min = XYZ(
    -wallDepth / 2 - offset_50cm, 0 - lower_floor_distance_cm - lower_floor_in_centimeters - offset_30cm, -win_width / 2
)

if top_offset is None:
    section_box.Max = XYZ(
        wallDepth / 2 + offset_70cm, win_height + upper_floor_distance_cm + upper_floor_in_centimeters + offset_30cm,
        -win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
    )
else:
    section_box.Max = XYZ(
        wallDepth / 2 + offset_70cm, win_height + top_offset,
        -win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
    )


# section_box.Min = XYZ(
#     -wallDepth / 2 - offset_50cm, 0 - lower_floor_distance_cm - lower_floor_in_centimeters - offset_30cm, -win_width / 2
# ) # To bottom
# section_box.Max = XYZ(
#     wallDepth / 2 + offset_70cm, win_height + upper_floor_distance_cm + upper_floor_in_centimeters + offset_30cm, -win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
# ) # To top

# glybyna(inner-outer) -> height

section_box.Transform = transform
section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
views = FilteredElementCollector(doc).OfClass(View).ToElements()
viewTemplates = [v for v in views if v.IsTemplate and "Section_Reinforcement" in v.Name.ToString()]
win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
# new_name = 'py_{}'.format(windowFamilyObject.Symbol.Family.Name)
new_name = 'MAMAD_Window_Section_1'
for i in range(10):
    try:
        win_elevation.Name = new_name
        print('Created a section: {}'.format(new_name))
        break
    except:
        new_name += '*'


# ----------------------------------------------------------------------------------------------- Perpendicular shelter
# Getting upper offset if upper floor exists
upper_distance = find_upper_window(wall_level_id, windowFamilyObject)
if upper_distance is None:
    top_offset = None
else:
    top_offset = upper_distance - UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)

(lower_floor_distance_cm,
 lower_floor_in_centimeters,
 upper_floor_distance_cm,
 upper_floor_in_centimeters) = find_floors_offsets(wall_level)

transform = Transform.Identity
transform.Origin = window_origin
vector_perp = vector.CrossProduct(XYZ.BasisZ)
transform.BasisX = vector_perp
transform.BasisY = XYZ.BasisZ
transform.BasisZ = vector_perp.CrossProduct(XYZ.BasisZ)
section_box = BoundingBoxXYZ()
offset_30cm = UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)
offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)
section_box.Min = XYZ(
    -wallDepth / 2 - offset_50cm, 0 - lower_floor_distance_cm - lower_floor_in_centimeters - offset_30cm, win_width / 2
)
if top_offset is None:
    section_box.Max = XYZ(
        wallDepth / 2 + offset_70cm, win_height + upper_floor_distance_cm + upper_floor_in_centimeters + offset_30cm,
        win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
    )
else:
    section_box.Max = XYZ(
        wallDepth / 2 + offset_70cm, win_height + top_offset,
        win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
    )

section_box.Transform = transform
section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
views = FilteredElementCollector(doc).OfClass(View).ToElements()
viewTemplates = [v for v in views if v.IsTemplate and "Section_Reinforcement" in v.Name.ToString()]
win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
new_name = 'MAMAD_Window_Section_2'
for i in range(10):
    try:
        win_elevation.Name = new_name
        print('Created a section: {}'.format(new_name))
        break
    except:
        new_name += '*'

transaction.Commit()

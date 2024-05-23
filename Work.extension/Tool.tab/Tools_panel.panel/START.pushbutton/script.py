# ------------------------------ IMPORTS ------------------------------
import math
import sys

from Autodesk.Revit.ApplicationServices import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
# from pyrevit import forms


# ------------------------------ VARIABLES ------------------------------
__title__ = "Sections automation"
__doc__ = """
Plugin which creates 3 sections depending on 
"""
uidoc = __revit__.ActiveUIDocument          # type: UIDocument
doc = __revit__.ActiveUIDocument.Document   # type: Document
app = __revit__.Application                 # type: Application
print('HI!')


# ------------------------------ FUnctions ------------------------------
# Old algorythm
# def find_upper_window(wall_lvl_id, current_window):
#     """
#     Finds a window directly above the given window on the level above and calculates distance.
#     Args:
#         wall_lvl_id (ElementId): The ID of the level the wall belongs to.
#         current_window (FamilyInstance): The current window element.
#     Returns:
#         FamilyInstance: The window directly above, if found; None otherwise.
#     """
#     wall_level = doc.GetElement(wall_lvl_id)
#     upper_level = None
#     all_levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).ToElements()
#     for level in all_levels:
#         if type(level) == LevelType:
#             continue
#         if level.Elevation > wall_level.Elevation:
#             if not level.Name.startswith('Floor'):
#                 continue
#             upper_level = level
#             break
#     if upper_level:
#         elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
#         current_window_bbox = current_window.get_BoundingBox(None)
#         current_window_x, current_window_y = current_window_bbox.Min.X, current_window_bbox.Min.Y
#         upperWindow = None
#         distance = float('inf')
#         for elem in elements:
#             if (elem.Category is not None and
#                     elem.Category.Id == current_window.Category.Id and
#                     elem.LevelId == upper_level.Id):
#                 if upperWindow is None:
#                     upperWindow = elem
#                     continue
#                 elem_bbox = elem.get_BoundingBox(None)
#                 elem_x, elem_y = elem_bbox.Min.X, elem_bbox.Min.Y
#                 new_distance = math.sqrt((elem_x - current_window_x) ** 2 + (elem_y - current_window_y) ** 2)
#                 if new_distance < distance:
#                     distance = new_distance
#                     upperWindow = elem
#         if upperWindow is not None:
#             # print('Upper window:', upperWindow.Name)
#             return UnitUtils.ConvertToInternalUnits(
#                 (upperWindow.get_BoundingBox(None).Min.Z - current_window.get_BoundingBox(None).Max.Z) * 30.48,
#                 UnitTypeId.Centimeters)
#     return None




# Old code
# def find_floors_offsets(wall_lvl):
#     floor_collector = FilteredElementCollector(doc).OfClass(Floor)
#     upper_floor, lower_floor = None, None
#     upper_floor_level_best_height, lower_floor_level_best_height = float('inf'), float('-inf')
#     for floor in floor_collector:
#         floor_level_id = floor.LevelId
#         floor_level = None
#         if floor_level_id:
#             floor_level = doc.GetElement(floor_level_id)
#         if floor_level and wall_lvl:
#             if floor_level.Elevation > wall_lvl.Elevation:
#                 if floor_level.Elevation < upper_floor_level_best_height:
#                     upper_floor_level_best_height = floor_level.Elevation
#                     upper_floor = floor
#             elif floor_level.Elevation <= wall_lvl.Elevation:
#                 if floor_level.Elevation > lower_floor_level_best_height:
#                     lower_floor_level_best_height = floor_level.Elevation
#                     lower_floor = floor
#     if upper_floor:
#         # print("Upper Floor:", upper_floor.Name)
#         upper_floor_params = upper_floor.Parameters
#         upper_floor_height = None
#         for param in upper_floor_params:
#             if param.Definition.Name == 'Core Thickness':
#                 upper_floor_height = param.AsDouble()
#                 break
#         if upper_floor_height is not None:
#             upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(upper_floor_height * 30.48,
#                                                                           UnitTypeId.Centimeters)
#             # print('Top floor height in cm', upper_floor_height * 30.48)
#         else:
#             upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#         window_bbox = windowFamilyObject.get_BoundingBox(None)  # False for local coordinate system
#         floor_bbox = upper_floor.get_BoundingBox(None)
#         window_top_z = window_bbox.Max.Z
#         floor_bottom_z = floor_bbox.Min.Z
#         distance_z = floor_bottom_z - window_top_z
#         upper_floor_distance_cm = UnitUtils.ConvertToInternalUnits(distance_z * 30.48, UnitTypeId.Centimeters)
#         # print('Distance to upper floor in cm', distance_z * 30.48)
#     else:
#         upper_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#         upper_floor_distance_cm = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#         # print("No upper floor found for the wall")
#     if lower_floor:
#         # print("Lower Floor:", lower_floor.Name)
#         lower_floor_params = lower_floor.Parameters
#         lower_floor_height = None
#         for param in lower_floor_params:
#             if param.Definition.Name == 'Core Thickness':
#                 lower_floor_height = param.AsDouble()
#                 break
#         if lower_floor_height is not None:
#             lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(lower_floor_height * 30.48, UnitTypeId.Centimeters)
#             # print('Height of lower floor in cm', lower_floor_height * 30.48)
#         else:
#             lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#         window_bbox = windowFamilyObject.get_BoundingBox(None)  # False for local coordinate system
#         floor_bbox = lower_floor.get_BoundingBox(None)
#         window_top_z = window_bbox.Min.Z
#         floor_bottom_z = floor_bbox.Max.Z
#         distance_z = window_top_z - floor_bottom_z
#         lower_floor_distance_cm = UnitUtils.ConvertToInternalUnits(distance_z * 30.48, UnitTypeId.Centimeters)
#         # print('Distance to lower floor in cm', distance_z * 30.48)
#     else:
#         lower_floor_in_centimeters = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#         lower_floor_distance_cm = UnitUtils.ConvertToInternalUnits(0, UnitTypeId.Centimeters)
#         # print("No lower floor found for the wall")
#     return lower_floor_distance_cm, lower_floor_in_centimeters, upper_floor_distance_cm, upper_floor_in_centimeters


def geographical_finding_algorythm(start_point, end_point, object_to_find_name=None, object_to_find_categoty=None, ignore_id=None):
    if object_to_find_name is None and object_to_find_categoty is None:
        return None

    min_x, max_x = min(start_point.X, end_point.X), max(start_point.X, end_point.X)
    min_y = min(start_point.Y, end_point.Y)
    min_z = min(start_point.Z, end_point.Z)
    max_y = max(start_point.Y, end_point.Y)
    max_z = max(start_point.Z, end_point.Z)

    # Create the bounding box
    outline = Outline(XYZ(min_x, min_y, min_z), XYZ(max_x, max_y, max_z))


    # if start_point < end_point:
    #     outline = Outline(end_point, start_point)
    # else:
    #     outline = Outline(start_point, end_point)
    bboxFilter = BoundingBoxIntersectsFilter(outline)
    allIntersections = FilteredElementCollector(doc).WherePasses(bboxFilter)
    intersections = []
    if object_to_find_categoty is not None:
        for intersec in allIntersections:
            # print(intersec.Category.Name)
            if hasattr(intersec.Category, 'Name') and intersec.Category.Name == object_to_find_categoty.Name:
                if ignore_id == intersec.Id:
                    continue
                intersections.append(intersec)
                # print(intersec, intersec.Id)
    else:
        for intersec in allIntersections:
            if intersec.Name == object_to_find_name:
                if ignore_id == intersec.Id:
                    continue
                intersections.append(intersec)
                # print(intersec, intersec.Id)
    return intersections


def find_upper_window(current_window):
    window = geographical_finding_algorythm(
        current_window.Location.Point - win_width * vector + XYZ(0, 0, win_height),
        current_window.Location.Point + win_width * vector + XYZ(0, 0, win_height + wall_height + wallDepth),
        object_to_find_name=current_window.Name,
        ignore_id=current_window.Id
    )
    print('found windows', window)
    if len(window):
        return window[0]
    else:
        return None


def get_wall_direction_vector(wall):
    # curve = wall.Location.Curve
    # direction_vector = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    # print('AAAAAAAAAA')
    # print(direction_vector)
    # print(curve.Direction)
    # print('AAAAAAAAAA')
    return wall.Location.Curve.Direction


def calculate_distance_in_direction(element1, element2, normalized_direction_vector):
    # Get the location points of the elements
    try:
        location1 = element1.Location.Curve.GetEndPoint(0)
    except:
        location1 = element1.Location.Point
    try:
        location2 = element2.Location.Curve.GetEndPoint(0)
    except:
        location2 = element2.Location.Point

    # Compute the vector from element1 to element2
    vector_between_elements = location2 - location1

    # Project the vector between elements onto the direction vector
    distance = vector_between_elements.DotProduct(normalized_direction_vector)

    return abs(distance)


def find_floors_offsets(current_window):
    host_wall = current_window.Host
    host_wall_bbox = host_wall.get_BoundingBox(None)
    current_window_bbox = current_window.get_BoundingBox(None)

    categories = doc.Settings.Categories
    floor_category = None
    for c in categories:
        if c.Name == 'Floors':
            floor_category = c
    if floor_category is None:
        raise Exception('There is no floor\'s category named "Floors"')

    top_floor = geographical_finding_algorythm(
        current_window_bbox.Min,
        XYZ(current_window_bbox.Max.X, current_window_bbox.Max.Y, host_wall.get_BoundingBox(None).Max.Z + 3),
        object_to_find_categoty=floor_category
    )
    if len(top_floor):
        top_floor_bbox = top_floor[0].get_BoundingBox(None)
        top_floor_height = top_floor_bbox.Max.Z - top_floor_bbox.Min.Z
    else:
        top_floor_height = 0

    top_window = geographical_finding_algorythm(
        current_window_bbox.Min,
        XYZ(current_window_bbox.Max.X, current_window_bbox.Max.Y, host_wall.get_BoundingBox(None).Max.Z + abs(host_wall_bbox.Max.Z - host_wall_bbox.Min.Z)),
        object_to_find_name=windowFamilyObject.Name,
        ignore_id=windowFamilyObject.Id
    )
    if len(top_window):
        top_window = top_window[0]
        top_floor_offset = top_window.get_BoundingBox(None).Min.Z - current_window_bbox.Max.Z
    else:
        top_floor_offset = host_wall_bbox.Max.Z - current_window_bbox.Max.Z + top_floor_height
    top_offset_in_cm = UnitUtils.ConvertToInternalUnits(top_floor_offset * 30.48 + 30, UnitTypeId.Centimeters)

    bottom_floor = geographical_finding_algorythm(
        current_window_bbox.Max,
        XYZ(current_window_bbox.Min.X, current_window_bbox.Min.Y, host_wall.get_BoundingBox(None).Min.Z - 3),
        object_to_find_categoty=floor_category
    )
    if len(bottom_floor):
        bottom_floor_bbox = bottom_floor[0].get_BoundingBox(None)
        bottom_floor_height = bottom_floor_bbox.Max.Z - bottom_floor_bbox.Min.Z
    else:
        bottom_floor_height = 0
    bottom_offset_in_cm = UnitUtils.ConvertToInternalUnits((current_window_bbox.Min.Z - host_wall_bbox.Min.Z + bottom_floor_height) * 30.48 + 30, UnitTypeId.Centimeters)

    return top_offset_in_cm, bottom_offset_in_cm

# def get_position_vector(window_object):
#     print('haha')
#     fam = window_object.Symbol.Family
#     print(fam)
#     famDoc = doc.EditFamily(fam)
#     coll = FilteredElementCollector(famDoc).WhereElementIsNotElementType().ToElements()
#     counter = 1
#     for elem in coll:
#         print('try', counter, len(coll))
#         counter += 1
#         try:
#             if elem.Name == 'Extrusion':
#                 print(elem)
#                 break
#         except AttributeError:
#             pass
    # print(window_object)
    # print(window_object.GetSubComponentIds())
    # print(window_object.GetCopingIds())
    # print(window_object.GetFamilyPointPlacementReferences())
    # for child_id in window_object.GetSubComponentIds():
    #     print(child_id)
    # for child_id in window_object.GetCopingIds():
    #     print(child_id)
    # for child_id in window_object.GetFamilyPointPlacementReferences():
    #     print(child_id)


# ------------------------------ MAIN ------------------------------

selection = uidoc.Selection.GetElementIds()
if len(selection) != 1:
    TaskDialog.Show("Selection Error", "Please select only one element. Plugin will now stop.")
    sys.exit()
windowFamilyObject = doc.GetElement(selection[0])  # type: FamilyInstance
print("Selected Element:", windowFamilyObject.Symbol.Family.Name)


window_origin = windowFamilyObject.Location.Point   # type: XYZ
host_object = windowFamilyObject.Host               # type: Wall
# curve = host_object.Location.Curve                  # type: Curve
# pt_start = curve.GetEndPoint(0)                     # type: XYZ
# pt_end = curve.GetEndPoint(1)                       # type: XYZ
# vector = pt_end - pt_start                          # type: XYZ
# vector = vector.Normalize()
win_height = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.GENERIC_HEIGHT).AsDouble()
win_width = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble()
win_depth = UnitUtils.ConvertToInternalUnits(40, UnitTypeId.Centimeters)
vector = get_wall_direction_vector(host_object)
# print('old window_origin', window_origin)
# window_origin -= vector * 10 / 30.48
# print('new window_origin', window_origin)
# print(vector)
wall_bbox = host_object.get_BoundingBox(None)
wall_height = wall_bbox.Max.Z - wall_bbox.Min.Z
wallDepth = UnitUtils.ConvertToInternalUnits(host_object.Width, UnitTypeId.Feet)
wall_level_id = host_object.LevelId

if wall_level_id:
    wall_level = doc.GetElement(wall_level_id)  # Get the Level element
else:
    print("Wall does not have an associated level. Plugin will now stop.")
    sys.exit()

# print('----------')
# win = find_upper_window(windowFamilyObject)
# if win:
#     print(win, win.Id)
# find_floors_offsets(windowFamilyObject)
# print('----------')


# print('-' for _ in range(50))


# geographical_finding_algorythm(window_origin + XYZ(0, 0, win_height), window_origin + XYZ(0, 0, win_height + 300 / 30.48), object_to_find_name=windowFamilyObject.Name, object_to_find_categoty=None)
# get_position_vector(windowFamilyObject)

# print(windowFamilyObject.FacingOrientation)


# -------------------------------------------------------------------------------- Front view
def get_front_view():
    top_offset, bottom_offset = find_floors_offsets(windowFamilyObject)


    end_point_left = window_origin + (win_width + 120 / 30.48) * get_wall_direction_vector(host_object)
    end_point_right = window_origin - (win_width + 120 / 30.48) * get_wall_direction_vector(host_object)


    # print('points', end_point_left, window_origin, end_point_right)
    # print(get_wall_direction_vector(host_object))
    # print(window_origin)
    # print(end_point_left)
    # print(window_origin - end_point_right)
    print('left:')
    windows = geographical_finding_algorythm(
        window_origin,
        end_point_left,
        object_to_find_name=windowFamilyObject.Name,
        ignore_id=windowFamilyObject.Id)
    left_offset = UnitUtils.ConvertToInternalUnits(120, UnitTypeId.Centimeters)
    if len(windows):
        # print('---', 'windows')
        # for w in windows:
        #     print(w, w.Id)
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
        if len(windows):
            left_offset = UnitUtils.ConvertToInternalUnits((best_distance - win_width) * 30.48, UnitTypeId.Centimeters)
            print('NEW LEFT OFFSET', best_distance * 30.48)
    else:
        walls = geographical_finding_algorythm(
        window_origin,
        end_point_left,
        object_to_find_categoty=host_object.Category,
        ignore_id=host_object.Id)
        print('---', 'walls')
        for w in walls:
            print(w, w.Id)
        print(walls)
        print(win_depth)
        if len(walls):
            left_distance = walls[0].Location.Curve.Distance(window_origin)
            left_offset = UnitUtils.ConvertToInternalUnits((left_distance - win_width - 10/30.48) * 30.48, UnitTypeId.Centimeters)
            # left_offset = UnitUtils.ConvertToInternalUnits(calculate_distance_in_direction(windowFamilyObject, walls[0], get_wall_direction_vector(host_object)) * 30.48, UnitTypeId.Centimeters)
            print('new wall distance:', left_distance)

    print('left offset')
    print(left_offset)

    print('right:')
    windows = geographical_finding_algorythm(window_origin, end_point_right, object_to_find_name=windowFamilyObject.Name, ignore_id=windowFamilyObject.Id)
    right_offset = UnitUtils.ConvertToInternalUnits(120, UnitTypeId.Centimeters)
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
        right_offset = UnitUtils.ConvertToInternalUnits((best_distance - win_width) * 30.48, UnitTypeId.Centimeters)
        print('NEW Right OFFSET', best_distance * 30.48)
    else:
        walls = geographical_finding_algorythm(
            window_origin,
            end_point_left,
            object_to_find_categoty=host_object.Category,
            ignore_id=host_object.Id)
        print('---', 'walls')
        for w in walls:
            print(w, w.Id)
        print(walls)
        print(win_depth)
        if len(walls):
            right_distance = walls[0].Location.Curve.Distance(window_origin)
            left_offset = UnitUtils.ConvertToInternalUnits((right_distance - win_width - 10 / 30.48) * 30.48,
                                                           UnitTypeId.Centimeters)
            # left_offset = UnitUtils.ConvertToInternalUnits(calculate_distance_in_direction(windowFamilyObject, walls[0], get_wall_direction_vector(host_object)) * 30.48, UnitTypeId.Centimeters)
            print('new wall distance:', right_distance)

    print('right')
    print(right_offset)
    print('done')





    transform = Transform.Identity
    transform.Origin = window_origin
    transform.BasisX = vector
    transform.BasisY = XYZ.BasisZ
    transform.BasisZ = vector.CrossProduct(XYZ.BasisZ)
    section_box = BoundingBoxXYZ()
    offset = UnitUtils.ConvertToInternalUnits(1, UnitTypeId.Centimeters)
    # section_box.Min = XYZ(-win_width - (offset * 50), 0 - (offset * 50), -win_depth)
    # section_box.Max = XYZ(win_width + (offset * 50), win_height + (offset * 50), win_depth)
    offset_30cm = UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)
    offset_120cm = UnitUtils.ConvertToInternalUnits(120, UnitTypeId.Centimeters)


    section_box.Min = XYZ(
        -win_width - right_offset, 0 - bottom_offset, -win_depth
    )
    section_box.Max = XYZ(
        win_width + left_offset, win_height + top_offset,
        win_depth
    )




    section_box.Transform = transform
    section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
    views = FilteredElementCollector(doc).OfClass(View).ToElements()
    viewTemplates = [v for v in views if v.IsTemplate and "Window_View" in v.Name.ToString()]
    win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
    win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
    new_name = 'MAMAD_Window_Front_View'
    for i in range(10):
        try:
            win_elevation.Name = new_name
            print('Created a section: {}'.format(new_name))
            break
        except:
            new_name += '*'

    # # Getting upper offset if upper floor exists
    # upper_distance = find_upper_window(wall_level_id, windowFamilyObject)
    # if upper_distance is None:
    #     top_offset = None
    # else:
    #     top_offset = upper_distance + UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)
    #
    # (lower_floor_distance_cm,
    #  lower_floor_in_centimeters,
    #  upper_floor_distance_cm,
    #  upper_floor_in_centimeters) = find_floors_offsets(wall_level)
    #
    #
    # transform = Transform.Identity
    # transform.Origin = window_origin
    # transform.BasisX = vector
    # transform.BasisY = XYZ.BasisZ
    # transform.BasisZ = vector.CrossProduct(XYZ.BasisZ)
    # section_box = BoundingBoxXYZ()
    # offset = UnitUtils.ConvertToInternalUnits(1, UnitTypeId.Centimeters)
    # # section_box.Min = XYZ(-win_width - (offset * 50), 0 - (offset * 50), -win_depth)
    # # section_box.Max = XYZ(win_width + (offset * 50), win_height + (offset * 50), win_depth)
    # offset_30cm = UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Centimeters)
    # offset_120cm = UnitUtils.ConvertToInternalUnits(120, UnitTypeId.Centimeters)
    #
    #
    # section_box.Min = XYZ(
    #     -win_width - offset_120cm, 0 - lower_floor_distance_cm - lower_floor_in_centimeters - offset_30cm, -win_depth
    # )
    #
    # if top_offset is None:
    #     section_box.Max = XYZ(
    #         win_width + offset_120cm, win_height + upper_floor_distance_cm + upper_floor_in_centimeters + offset_30cm,
    #         win_depth
    #     )
    # else:
    #     section_box.Max = XYZ(
    #         win_width + offset_120cm, win_height + top_offset,
    #         win_depth
    #     )
    #
    #
    # # section_box.Transform = transform
    # # section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
    # # win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
    # # new_name = 'py_{}'.format(windowFamilyObject.Symbol.Family.Name)
    #
    #
    # section_box.Transform = transform
    # section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
    # views = FilteredElementCollector(doc).OfClass(View).ToElements()
    # viewTemplates = [v for v in views if v.IsTemplate and "Window_View" in v.Name.ToString()]
    # win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
    # win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
    # new_name = 'MAMAD_Window_Front_View'
    #
    # # win_elevation.Name = new_name + '123123'
    # # print('Created a section: {}'.format(new_name))
    #
    # for i in range(10):
    #     try:
    #         win_elevation.Name = new_name
    #         print('Created a section: {}'.format(new_name))
    #         break
    #     except:
    #         new_name += '*'


# -------------------------------------------------------------------------------- From down to up --- SHOULD BE CALLOUT
def get_callout():
    pass

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


# -------------------------------------------------------------------------------- Perpendicular window
def get_perpendicular_window_section():

    top_offset, bottom_offset = find_floors_offsets(windowFamilyObject)

    print('Vector', vector, 'Position', windowFamilyObject.FacingOrientation)

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

    if (windowFamilyObject.FacingOrientation.X == -vector.Y and vector.Y != 0) or (windowFamilyObject.FacingOrientation.Y == vector.X and vector.X != 0):
        exterior_offset = offset_70cm
        interior_offset = offset_50cm
    else:
        exterior_offset = offset_50cm
        interior_offset = offset_70cm

    if windowFamilyObject.FacingOrientation.X + vector.Y < 0.00001:
        print('One side!')
    else:
        print('Not one side!')

    section_box.Min = XYZ(
        -wallDepth / 2 - exterior_offset, 0 - bottom_offset, -win_width / 2
    )
    section_box.Max = XYZ(
        wallDepth / 2 + interior_offset, win_height + top_offset,
        -win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
    )

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


# -------------------------------------------------------------------------------- Perpendicular shelter
def get_perpendicular_shelter_section():
    top_offset, bottom_offset = find_floors_offsets(windowFamilyObject)

    # print('WFO', windowFamilyObject.FacingOrientation, 'vec', vector)
    # print(windowFamilyObject.FacingOrientation.X, vector.Y, '=', windowFamilyObject.FacingOrientation.X + vector.Y)
    # print(windowFamilyObject.FacingOrientation.Y, vector.X, '=', windowFamilyObject.FacingOrientation.Y + vector.X)

    if (windowFamilyObject.FacingOrientation.X == -vector.Y and vector.Y != 0) or (windowFamilyObject.FacingOrientation.Y == vector.X and vector.X != 0):
        print('One side!')
    else:
        print('Not one side!')

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
    if (windowFamilyObject.FacingOrientation.X == -vector.Y and vector.Y != 0) or (windowFamilyObject.FacingOrientation.Y == vector.X and vector.X != 0):
        exterior_offset = offset_70cm
        interior_offset = offset_50cm
    else:
        exterior_offset = offset_50cm
        interior_offset = offset_70cm

    section_box.Min = XYZ(
        -wallDepth / 2 - exterior_offset, 0 - bottom_offset, win_width / 2
    )
    section_box.Max = XYZ(
        wallDepth / 2 + interior_offset, win_height + top_offset,
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


transaction = Transaction(doc, 'Generate Window Sections')
transaction.Start()

# get_front_view()
# get_perpendicular_window_section()
# get_perpendicular_shelter_section()
get_callout()

transaction.Commit()
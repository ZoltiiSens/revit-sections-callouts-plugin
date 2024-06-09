# ------------------------------ IMPORTS ------------------------------
import math
import sys

from Autodesk.Revit.ApplicationServices import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *

from Autodesk.Revit.DB.Structure import *
from System.Collections.Generic import List

# from pyrevit import forms


# ------------------------------ VARIABLES ------------------------------
__title__ = "Sections automation"
__doc__ = """
Plugin which creates 3 sections depending on 
"""
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = __revit__.ActiveUIDocument.Document  # type: Document
app = __revit__.Application  # type: Application
createdoc = doc.Create  # type: Autodesk.Revit.Creation.Document
print('HI!')


# ------------------------------ Functions ------------------------------
def geographical_finding_algorythm(start_point, end_point, object_to_find_name=None, object_to_find_categoty=None,
                                   object_to_find_builtin_categoty=None, ignore_id=None):
    if object_to_find_name is None and object_to_find_categoty is None and object_to_find_builtin_categoty is None:
        return None

    min_x, max_x = min(start_point.X, end_point.X), max(start_point.X, end_point.X)
    min_y = min(start_point.Y, end_point.Y)
    min_z = min(start_point.Z, end_point.Z)
    max_y = max(start_point.Y, end_point.Y)
    max_z = max(start_point.Z, end_point.Z)
    outline = Outline(XYZ(min_x, min_y, min_z), XYZ(max_x, max_y, max_z))

    bboxFilter = BoundingBoxIntersectsFilter(outline)
    allIntersections = FilteredElementCollector(doc).WherePasses(bboxFilter)
    intersections = []
    if object_to_find_categoty is not None:
        for intersec in allIntersections:
            if hasattr(intersec.Category, 'Name') and intersec.Category.Name == object_to_find_categoty.Name:
                if ignore_id == intersec.Id:
                    continue
                intersections.append(intersec)
    elif object_to_find_name is not None:
        for intersec in allIntersections:
            if intersec.Name == object_to_find_name:
                if ignore_id == intersec.Id:
                    continue
                intersections.append(intersec)
    elif object_to_find_builtin_categoty is not None:
        allIntersections = allIntersections.OfCategory(object_to_find_builtin_categoty)
        for intersec in allIntersections:
            if ignore_id == intersec.Id:
                continue
            intersections.append(intersec)
    return intersections


# Now unnecessary
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


# Now unnecessary
def calculate_distance_in_direction(element1, element2, normalized_direction_vector):
    try:
        location1 = element1.Location.Curve.GetEndPoint(0)
    except:
        location1 = element1.Location.Point
    try:
        location2 = element2.Location.Curve.GetEndPoint(0)
    except:
        location2 = element2.Location.Point

    vector_between_elements = location2 - location1
    distance = vector_between_elements.DotProduct(normalized_direction_vector)

    return abs(distance)


def get_wall_direction_vector(wall):
    return wall.Location.Curve.Direction


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
        XYZ(current_window_bbox.Max.X, current_window_bbox.Max.Y,
            host_wall.get_BoundingBox(None).Max.Z + abs(host_wall_bbox.Max.Z - host_wall_bbox.Min.Z)),
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
    bottom_offset_in_cm = UnitUtils.ConvertToInternalUnits(
        (current_window_bbox.Min.Z - host_wall_bbox.Min.Z + bottom_floor_height) * 30.48 + 30, UnitTypeId.Centimeters)

    return top_offset_in_cm, bottom_offset_in_cm


def find_rebars_by_quantity_and_spacing(view, start_point, end_point, quantity, spacing, which_to_show):
    rebars = geographical_finding_algorythm(
        start_point,
        end_point,
        object_to_find_builtin_categoty=BuiltInCategory.OST_Rebar)
    target_rebar = None
    for reb in rebars:
        if spacing - 0.5 < reb.MaxSpacing * 30.48 < spacing + 0.5 and reb.Quantity == quantity:
            target_rebar = reb
            break
    if target_rebar is None:
        print('Can\'t find rebar set(front view)')
    else:
        target_rebar.SetUnobscuredInView(doc.ActiveView, True)
        target_rebar.SetPresentationMode(view, RebarPresentationMode.Select)
        for i in range(target_rebar.NumberOfBarPositions):
            target_rebar.SetBarHiddenStatus(view, i, True)
        target_rebar.SetBarHiddenStatus(view, which_to_show, False)


def find_rebars_on_view(view):
    return FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()


# def create_rebar_tag(rebar, view, tag_type=None, orientation=0):
#     '''
#
#     :param rebar:
#     :param view:
#     :param tag_type:
#     :param orientation: 0 == Horizontal, 1 == Vertical
#     :return:
#     '''
#
#     rebar_set = doc.GetElement(rebar.Id)
#     if not rebar_set:
#         raise Exception("Rebar set not found")
#
#
#     # Create a tag for the rebar set
#     tag_mode = TagMode.TM_ADDBY_CATEGORY
#     tag_orientation = TagOrientation.Horizontal
#
#     # Find a valid location for the tag
#     # For simplicity, place it at the centroid of the bounding box
#     rebar_bbox = rebar.get_BoundingBox(None)
#     rebar_point = XYZ(
#         (rebar_bbox.Min.X + rebar_bbox.Max.X) / 2,
#         (rebar_bbox.Min.Y + rebar_bbox.Max.Y) / 2,
#         (rebar_bbox.Min.Z + rebar_bbox.Max.Z) / 2,
#     )
#
#     # Create the tag
#     rebar_tag = IndependentTag.Create(doc, view.Id, Reference(rebar_set), True, tag_mode, tag_orientation,
#                                       rebar_point)
#
#     if not rebar_tag:
#         raise Exception("Failed to create rebar tag")


def get_tag_types():
    familySymbols = FilteredElementCollector(doc).OfClass(FamilySymbol)
    bendingDetailTypes = FilteredElementCollector(doc).OfClass(RebarBendingDetailType)
    result = {}
    for famS in familySymbols:
        name = famS.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
        if 'Horizontal_Bars' in name:
            result['Horizontal_Bars'] = famS
        if 'Column_Vertical' in name:
            result['Column_Vertical'] = famS
        if 'Wall&Col_Vertical+Length' in name:
            result['Wall&Col_Vertical+Length'] = famS
        if 'Link&U-Shape+Length' in name:
            result['Link&U-Shape+Length'] = famS

    # for bdt in bendingDetailTypes:
    #     name = bdt.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
    #     print(name, bdt.Id, bdt.FamilyName, name_param)
    #     if 'Bending Detail 2 (No hooks)' in name:
    #         print('found')
    #         result['Bending Detail 2 (No hooks)'] = bdt
    return result


def create_rebar_tag(view, all_rebars, tag_mode, tag_orientation, tag_type_name, tag_position, partitionName, leader_end_condition=LeaderEndCondition.Attached, create_only_for_one = False):
    filtered_rebars = []
    for rebar in all_rebars:
        if rebar.LookupParameter("Partition").AsString() == partitionName:
            filtered_rebars.append(rebar)
    tag = None
    for i, rebar in enumerate(filtered_rebars):
        for subelement in rebar.GetSubelements():
            if tag is None:
                tag = IndependentTag.Create(
                    doc,
                    view.Id,
                    subelement.GetReference(),
                    True,
                    tag_mode,
                    tag_orientation,
                    XYZ(0, 0, 3))
                tag.ChangeTypeId(tagTypes[tag_type_name].Id)
                tag.TagHeadPosition = tag_position
                tag.LeaderEndCondition = leader_end_condition
                if create_only_for_one:
                    break
                continue
            lll = List[Reference]()
            lll.Add(subelement.GetReference())
            tag.AddReferences(lll)
    return tag


def create_rebar_tag_depending_on_rebar(view, all_rebars, tag_mode, tag_orientation, tag_type_name, tag_position, partitionName, leader_end_condition=LeaderEndCondition.Attached, create_only_for_one = False):
    filtered_rebars = []
    for rebar in all_rebars:
        if rebar.LookupParameter("Partition").AsString() == partitionName:
            filtered_rebars.append(rebar)
    tag = None
    for i, rebar in enumerate(filtered_rebars):
        for subelement in rebar.GetSubelements():
            if tag is None:
                tag = IndependentTag.Create(
                    doc,
                    view.Id,
                    subelement.GetReference(),
                    True,
                    tag_mode,
                    tag_orientation,
                    XYZ(0, 0, 0))
                tag.ChangeTypeId(tagTypes[tag_type_name].Id)

                rebar_bbox = rebar.get_BoundingBox(None)
                print(rebar_bbox)
                tag_coordinates = XYZ(
                    (rebar_bbox.Min.X + rebar_bbox.Max.X) / 2,
                    (rebar_bbox.Min.Y + rebar_bbox.Max.Y) / 2,
                    (rebar_bbox.Min.Z + rebar_bbox.Max.Z) / 2
                )
                tag.TagHeadPosition = tag_position + tag_coordinates

                tag.LeaderEndCondition = leader_end_condition
                if create_only_for_one:
                    break
                continue
            lll = List[Reference]()
            lll.Add(subelement.GetReference())
            tag.AddReferences(lll)
    return tag


def create_bending_detail(view, all_rebars, tag_mode, tag_orientation, tag_type_name, tag_position, partitionName, leader_end_condition=LeaderEndCondition.Attached, create_only_for_one = False):
    filtered_rebars = []
    for rebar in all_rebars:
        if rebar.LookupParameter("Partition").AsString() == partitionName:
            filtered_rebars.append(rebar)
    tag = None
    for i, rebar in enumerate(filtered_rebars):
        for subelement in rebar.GetSubelements():
            if tag is None:
                tag = RebarBendingDetail.Create(
                    doc,
                    view.Id,
                    subelement.GetReference(),
                    0,
                    tag_mode,
                    tag_orientation,
                    XYZ(0, 0, 3))
                tag.ChangeTypeId(tagTypes[tag_type_name].Id)
                tag.TagHeadPosition = tag_position
                tag.LeaderEndCondition = leader_end_condition
                if create_only_for_one:
                    break
                continue
            lll = List[Reference]()
            lll.Add(subelement.GetReference())
            tag.AddReferences(lll)
    return tag


# ------------------------------ MAIN ------------------------------
selection = uidoc.Selection.GetElementIds()
if len(selection) != 1:
    TaskDialog.Show("Selection Error", "Please select only one element. Plugin will now stop.")
    sys.exit()
windowFamilyObject = doc.GetElement(selection[0])
print("Selected Element:", windowFamilyObject.Symbol.Family.Name)

tagTypes = get_tag_types()

window_origin = windowFamilyObject.Location.Point
window_bbox = windowFamilyObject.get_BoundingBox(None)
host_object = windowFamilyObject.Host
win_height = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.GENERIC_HEIGHT).AsDouble()
win_width = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble()
win_depth = UnitUtils.ConvertToInternalUnits(40, UnitTypeId.Centimeters)
vector = get_wall_direction_vector(host_object)
wall_bbox = host_object.get_BoundingBox(None)
wall_height = wall_bbox.Max.Z - wall_bbox.Min.Z
wallDepth = UnitUtils.ConvertToInternalUnits(host_object.Width, UnitTypeId.Feet)
wall_level_id = host_object.LevelId

if wall_level_id:
    wall_level = doc.GetElement(wall_level_id)  # Get the Level element
else:
    print("Wall does not have an associated level. Plugin will now stop.")
    sys.exit()


# -------------------------------------------------------------------------------- Front view
def get_front_view():
    top_offset, bottom_offset = find_floors_offsets(windowFamilyObject)

    end_point_left = window_origin + (win_width + 200 / 30.48) * get_wall_direction_vector(host_object)
    end_point_right = window_origin - (win_width + 200 / 30.48) * get_wall_direction_vector(host_object)
    windows = geographical_finding_algorythm(
        window_origin,
        end_point_left,
        object_to_find_name=windowFamilyObject.Name,
        ignore_id=windowFamilyObject.Id)
    left_offset = UnitUtils.ConvertToInternalUnits(120, UnitTypeId.Centimeters)
    print("windows", windows)
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
        if len(windows):
            left_offset = UnitUtils.ConvertToInternalUnits((best_distance - win_width) * 30.48, UnitTypeId.Centimeters)
    else:
        walls = geographical_finding_algorythm(
            window_origin,
            end_point_left,
            object_to_find_categoty=host_object.Category,
            ignore_id=host_object.Id)
        if len(walls):
            left_distance = walls[0].Location.Curve.Distance(window_origin)
            left_offset = UnitUtils.ConvertToInternalUnits((left_distance - win_width - 10 / 30.48) * 30.48,
                                                           UnitTypeId.Centimeters)
    windows = geographical_finding_algorythm(window_origin, end_point_right,
                                             object_to_find_name=windowFamilyObject.Name,
                                             ignore_id=windowFamilyObject.Id)
    right_offset = UnitUtils.ConvertToInternalUnits(120, UnitTypeId.Centimeters)
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
        right_offset = UnitUtils.ConvertToInternalUnits((best_distance - win_width) * 30.48, UnitTypeId.Centimeters)
    else:
        walls = geographical_finding_algorythm(window_origin, end_point_left,
                                               object_to_find_categoty=host_object.Category, ignore_id=host_object.Id)
        if len(walls):
            right_distance = walls[0].Location.Curve.Distance(window_origin)
            right_offset = UnitUtils.ConvertToInternalUnits((right_distance - win_width - 10 / 30.48) * 30.48,
                                                           UnitTypeId.Centimeters)

    transform = Transform.Identity
    transform.Origin = window_origin
    transform.BasisX = vector
    transform.BasisY = XYZ.BasisZ
    transform.BasisZ = vector.CrossProduct(XYZ.BasisZ)
    section_box = BoundingBoxXYZ()
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

    left_rebar_start_point = XYZ(
        window_bbox.Min.X + left_offset * get_wall_direction_vector(host_object).X,
        window_bbox.Min.Y + left_offset * get_wall_direction_vector(host_object).Y,
        window_bbox.Min.Z
    )
    left_rebar_end_point = XYZ(
        window_bbox.Max.X + left_offset * get_wall_direction_vector(host_object).X,
        window_bbox.Max.Y + left_offset * get_wall_direction_vector(host_object).Y,
        window_bbox.Max.Z
    )
    find_rebars_by_quantity_and_spacing(win_elevation, window_bbox.Min, window_bbox.Max, 5, 20, 3)
    find_rebars_by_quantity_and_spacing(win_elevation, window_bbox.Min, window_bbox.Max, 8, 10, 2)
    find_rebars_by_quantity_and_spacing(win_elevation, left_rebar_start_point, left_rebar_end_point, 5, 20, 3)
    find_rebars_by_quantity_and_spacing(win_elevation, left_rebar_start_point, left_rebar_end_point, 8, 10, 2)

    all_rebars = find_rebars_on_view(win_elevation)
    rebar_ids_to_hide = List[ElementId]()
    for reb in all_rebars:
        if reb.GetHostId() != host_object.Id:
            rebar_ids_to_hide.Add(reb.Id)
    win_elevation.HideElements(rebar_ids_to_hide)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(1, 1, win_height + 0.85),
        'Window_Detail_1')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Column_Vertical',
        window_origin + XYZ((win_width + 1.2) * vector.X, (win_width + 1.2) * vector.Y, win_height / 2),
        'Window_Detail_2')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(1, 1, -0.85),
        'Window_Detail_3')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Column_Vertical',
        window_origin + XYZ(-(win_width + 1) * vector.X, -(win_width + 1) * vector.Y, win_height / 2),
        'Window_Detail_4')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Column_Vertical',
        window_origin + XYZ(-(win_width / 2 + 0.5) * vector.X, -(win_width / 2 + 0.5) * vector.Y, win_height / 2),
        'Window_Detail_5')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Column_Vertical',
        window_origin + XYZ(-(win_width / 2 + 1.5) * vector.X, -(win_width / 2 + 1.5) * vector.Y, win_height / 2),
        'Window_Detail_6')

    win_elevation.Scale = 25
    new_name = 'MAMAD_Window_Front_View'
    for i in range(10):
        try:
            win_elevation.Name = new_name
            print('Created a section: {}'.format(new_name))
            break
        except:
            new_name += '*'


# -------------------------------------------------------------------------------- From down to up --- SHOULD BE CALLOUT
def get_callout():
    end_point_left = window_origin + (win_width + 200 / 30.48) * get_wall_direction_vector(host_object)
    end_point_right = window_origin - (win_width + 200 / 30.48) * get_wall_direction_vector(host_object)

    windows = geographical_finding_algorythm(
        window_origin,
        end_point_left,
        object_to_find_name=windowFamilyObject.Name,
        ignore_id=windowFamilyObject.Id)
    left_point = window_origin + (win_width + 120 / 30.48) * get_wall_direction_vector(host_object)
    print("left point 120")
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt((window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
                left_point = window.Location.Point + XYZ(
                    win_width / 2 * get_wall_direction_vector(host_object).X,
                    win_width / 2 * get_wall_direction_vector(host_object).Y,
                    0)
                print("left point window", left_point)
    else:
        walls = geographical_finding_algorythm(
            window_origin,
            end_point_left,
            object_to_find_categoty=host_object.Category,
            ignore_id=host_object.Id)
        if len(walls):
            left_distance = walls[0].Location.Curve.Distance(window_origin)
            left_offset = (left_distance - win_width - 3 / 30.48)
            left_point = window_origin + XYZ(
                (win_width + left_offset) * get_wall_direction_vector(host_object).X,
                (win_width + left_offset) * get_wall_direction_vector(host_object).Y,
                0)
            print("left point wall", left_point)

    windows = geographical_finding_algorythm(
        window_origin,
        end_point_right,
        object_to_find_name=windowFamilyObject.Name,
        ignore_id=windowFamilyObject.Id)
    right_point = window_origin - (win_width + 120 / 30.48) * get_wall_direction_vector(host_object)
    print("right point 120")
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
                right_point = window.Location.Point - XYZ(
                        win_width / 2 * get_wall_direction_vector(host_object).X,
                        win_width / 2 * get_wall_direction_vector(host_object).Y,
                        0)
                print("right point window", right_point)
    else:
        walls = geographical_finding_algorythm(
            window_origin,
            end_point_right,
            object_to_find_categoty=host_object.Category,
            ignore_id=host_object.Id)
        if len(walls):
            right_distance = walls[0].Location.Curve.Distance(window_origin)
            right_offset = (right_distance - win_width - 22 / 30.48)
            right_point = window_origin - XYZ(
                (win_width + right_offset) * get_wall_direction_vector(host_object).X,
                (win_width + right_offset) * get_wall_direction_vector(host_object).Y,
                0)
            print("right point wall", right_point)

    print("points", left_point, right_point, window_origin)

    # Plan creation
    current_window_level_id = windowFamilyObject.LevelId
    fec = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    for x in fec:
        if ViewFamily.StructuralPlan == x.ViewFamily:
            fec = x
            break
    print(fec)
    structuralPlan = ViewPlan.Create(doc, fec.Id, current_window_level_id)

    perpendicular_vector = XYZ(-vector.Y, vector.X, vector.Z)

    start_point = XYZ(
        right_point.X + (perpendicular_vector.X * (host_object.Width / 2 + 10 / 30.48)),
        right_point.Y + (perpendicular_vector.Y * (host_object.Width / 2 + 10 / 30.48)),
        window_origin.Z
    )
    end_point = XYZ(
        left_point.X - (perpendicular_vector.X * (host_object.Width / 2 + 100 / 30.48)),
        left_point.Y - (perpendicular_vector.Y * (host_object.Width / 2 + 100 / 30.48)),
        window_origin.Z + 20 / 30.48
    )

    callout = ViewSection.CreateCallout(
        doc,
        structuralPlan.Id,
        fec.Id,
        start_point,
        end_point
    )

    # Modifying viewRange
    sill_height_param = windowFamilyObject.LookupParameter("Sill Height")
    if sill_height_param:
         sill_height_param = sill_height_param.AsDouble()  # Sill height is typically stored in feet
    else:
        raise Exception("Sill Height parameter not found in the window family instance.")

    view_range = callout.GetViewRange()
    view_range.SetOffset(PlanViewPlane.TopClipPlane, 70 / 30.48 + sill_height_param)
    view_range.SetOffset(PlanViewPlane.CutPlane, 70 / 30.48 + sill_height_param)
    view_range.SetOffset(PlanViewPlane.BottomClipPlane, 30 / 30.48 + sill_height_param)
    view_range.SetOffset(PlanViewPlane.ViewDepthPlane, 30 / 30.48 + sill_height_param)
    callout.SetViewRange(view_range)
    callout.Scale = 25

    # Set template
    views = FilteredElementCollector(doc).OfClass(View).ToElements()
    viewTemplates = [v for v in views if v.IsTemplate and "MAMAD_Window_Callout" in v.Name.ToString()]
    try:
        callout.ApplyViewTemplateParameters(viewTemplates[0])
    except IndexError:
        print('!!! There is no view template "MAMAD_Window_Callout"')

    all_rebars = find_rebars_on_view(callout)
    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-win_width * vector.X + 2 * perpendicular_vector.X, -win_width * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Detail_4')

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-(win_width - 1) * vector.X + 2 * perpendicular_vector.X,
                            -(win_width - 1) * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Detail_5',
        create_only_for_one=True)

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-(win_width - 1.5) * vector.X + 2 * perpendicular_vector.X,
                            -(win_width - 1.5) * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Detail_6',
        create_only_for_one=True)

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(2 * perpendicular_vector.X,
                            2 * perpendicular_vector.Y, 0),
        'Window_Front_Detail_11')

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(0.5 * vector.X + 2 * perpendicular_vector.X,
                            0.5 * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Front_Detail_10')

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ((win_width - 0.5) * vector.X + 2 * perpendicular_vector.X,
                            (win_width - 0.5) * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Front_Detail_9')

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ((win_width) * vector.X + 2 * perpendicular_vector.X,
                            (win_width) * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Detail_2')

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ((win_width + 1.5) * vector.X + 2 * perpendicular_vector.X,
                            (win_width + 1.5) * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Front_Detail_4',
        create_only_for_one=True)

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ((win_width + 1) * vector.X + 2 * perpendicular_vector.X,
                            (win_width + 1) * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Front_Detail_3',
        create_only_for_one=True)

    create_rebar_tag(
        callout,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ((win_width * 2) * vector.X + 2 * perpendicular_vector.X,
                            (win_width * 2) * vector.Y + 2 * perpendicular_vector.Y, 0),
        'Window_Detail_20')


    new_name = 'MAMAD_Window_Detail'
    for i in range(10):
        try:
            callout.Name = new_name
            print('Created a callout: {}'.format(new_name))
            break
        except:
            new_name += '*'

    parent_view_param = callout.get_Parameter(BuiltInParameter.SECTION_PARENT_VIEW_NAME)
    if parent_view_param and parent_view_param.IsReadOnly is False:
        parent_view_param.Set(ElementId.InvalidElementId)
    doc.Delete(structuralPlan.Id)


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
    offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
    offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)

    if (windowFamilyObject.FacingOrientation.X == -vector.Y and vector.Y != 0) or (
            windowFamilyObject.FacingOrientation.Y == vector.X and vector.X != 0):
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
    win_elevation.Scale = 25
    # new_name = 'py_{}'.format(windowFamilyObject.Symbol.Family.Name)

    all_rebars = find_rebars_on_view(win_elevation)
    # rebar_ids_to_hide = List[ElementId]()
    # for reb in all_rebars:
    #     if reb.GetHostId() != host_object.Id:
    #         rebar_ids_to_hide.Add(reb.Id)
    # win_elevation.HideElements(rebar_ids_to_hide)

    perpendicular_vector = XYZ(-vector.Y, vector.X, vector.Z)

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-1.5 * perpendicular_vector.X, -1.5 * perpendicular_vector.Y, 0.8),
        # XYZ(0, 0, 0),
        'Window_Detail_7')

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1.5 * perpendicular_vector.X, -1.5 * perpendicular_vector.Y, -(win_height + 1)),
    #     'Window_Detail_7')

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-2.4 * perpendicular_vector.X, -2.4 * perpendicular_vector.Y, 1),
        'Window_Front_Detail_6',
        leader_end_condition=LeaderEndCondition.Free,
        create_only_for_one=True)

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1.8 * perpendicular_vector.X, -1.8 * perpendicular_vector.Y, -(win_height - 0.2)),
    #     'Window_Front_Detail_6',
    #     leader_end_condition=LeaderEndCondition.Free,
    #     create_only_for_one=True)

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-1.29 * perpendicular_vector.X, -1.29 * perpendicular_vector.Y, 1.5),
        'Window_Front_Detail_5',
        create_only_for_one=True)

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1.8 * perpendicular_vector.X, -1.8 * perpendicular_vector.Y, -(win_height - 1)),
    #     'Window_Front_Detail_5',
    #     create_only_for_one=True)

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-1.3 * perpendicular_vector.X, -1.3 * perpendicular_vector.Y, -0.1),
        'Window_Detail_3')

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1.5 * perpendicular_vector.X, -1.5 * perpendicular_vector.Y, -(win_height - 2.9)),
    #     'Window_Detail_3')

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-1 * perpendicular_vector.X, -1 * perpendicular_vector.Y, 1),
        'Window_Front_Detail_8')

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1 * perpendicular_vector.X, -1 * perpendicular_vector.Y, 1),
    #     'Window_Front_Detail_8')

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-1 * perpendicular_vector.X, -1 * perpendicular_vector.Y, - 1),
        'Window_Front_Detail_7')

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1 * perpendicular_vector.X, -1 * perpendicular_vector.Y, win_height - 1),
    #     'Window_Front_Detail_7')

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-1.3 * perpendicular_vector.X, -1.3 * perpendicular_vector.Y, 0.1),
        'Window_Detail_1')

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1.5 * perpendicular_vector.X, -1.5 * perpendicular_vector.Y, win_height + 0.3),
    #     'Window_Detail_1')

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-2.4 * perpendicular_vector.X, -2.4 * perpendicular_vector.Y, 0),
        'Window_Front_Detail_2',
        create_only_for_one=True)

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1.8 * perpendicular_vector.X, -1.8 * perpendicular_vector.Y, win_height + 0.9),
    #     'Window_Front_Detail_2',
    #     create_only_for_one=True)

    create_rebar_tag_depending_on_rebar(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        XYZ(-1.29 * perpendicular_vector.X, -1.29 * perpendicular_vector.Y, -0.47),
        'Window_Front_Detail_1',
        create_only_for_one=True)

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1.8 * perpendicular_vector.X, -1.8 * perpendicular_vector.Y, win_height + 1.4),
    #     'Window_Front_Detail_1',
    #     create_only_for_one=True)

    # RebarBendingDetail.Create(
    #     doc,
    #     win_elevation.Id,
    #
    # )

    # create_bending_detail(
    #     doc,
    #     win_elevation.Id
    # )


    # tag = IndependentTag.Create(
    #     doc,
    #     view.Id,
    #     subelement.GetReference(),
    #     True,
    #     tag_mode,
    #     tag_orientation,
    #     XYZ(0, 0, 3))



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

    if (windowFamilyObject.FacingOrientation.X == -vector.Y and vector.Y != 0) or (
            windowFamilyObject.FacingOrientation.Y == vector.X and vector.X != 0):
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
    offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
    offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)
    if (windowFamilyObject.FacingOrientation.X == -vector.Y and vector.Y != 0) or (
            windowFamilyObject.FacingOrientation.Y == vector.X and vector.X != 0):
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
    win_elevation.Scale = 25


    all_rebars = find_rebars_on_view(win_elevation)
    perpendicular_vector = XYZ(-vector.Y, vector.X, vector.Z)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-1.65 * perpendicular_vector.X, -1.65 * perpendicular_vector.Y, -(win_height + 0.9)),
        'Window_Detail_7',
        leader_end_condition=LeaderEndCondition.Free)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-1.9 * perpendicular_vector.X, -1.9 * perpendicular_vector.Y, -(win_height - 0.8)),
        'Window_Front_Detail_6',
        leader_end_condition=LeaderEndCondition.Free,
        create_only_for_one= True)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-1.9 * perpendicular_vector.X, -1.9 * perpendicular_vector.Y, -(win_height - 1.2)),
        'Window_Front_Detail_5',
        leader_end_condition=LeaderEndCondition.Free,
        create_only_for_one=True)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-1.65 * perpendicular_vector.X, -1.65 * perpendicular_vector.Y, -0.6),
        'Window_Detail_3',
        leader_end_condition=LeaderEndCondition.Free)

    # create_rebar_tag(
    #     win_elevation,
    #     all_rebars,
    #     TagMode.TM_ADDBY_CATEGORY,
    #     TagOrientation.Horizontal,
    #     'Horizontal_Bars',
    #     window_origin + XYZ(-1.9 * perpendicular_vector.X, -1.9 * perpendicular_vector.Y, win_height / 2),
    #     'Window_Detail_14',
    #     leader_end_condition=LeaderEndCondition.Free,
    #     create_only_for_one= True)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-1.65 * perpendicular_vector.X, -1.65 * perpendicular_vector.Y, win_height),
        'Window_Detail_1',
        leader_end_condition=LeaderEndCondition.Free,
        create_only_for_one=True)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-1.9 * perpendicular_vector.X, -1.9 * perpendicular_vector.Y, win_height + 2),
        'Window_Front_Detail_2',
        leader_end_condition=LeaderEndCondition.Free,
        create_only_for_one=True)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-1.9 * perpendicular_vector.X, -1.9 * perpendicular_vector.Y, win_height + 1.4),
        'Window_Front_Detail_1',
        leader_end_condition=LeaderEndCondition.Free,
        create_only_for_one=True)

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-1.65 * perpendicular_vector.X, -1.65 * perpendicular_vector.Y, win_height + 2.4),
        'Window_Detail_14',
        leader_end_condition=LeaderEndCondition.Free)


    for i in range(10):
        try:
            win_elevation.Name = new_name
            print('Created a section: {}'.format(new_name))
            break
        except:
            new_name += '*'


transaction = Transaction(doc, 'Generate Window Sections')
transaction.Start()
try:
    # get_front_view()
    get_perpendicular_window_section()
    # get_perpendicular_shelter_section()
    # get_callout()
except Exception as err:
    print('ERROR!', err)
    print(vars(err))

transaction.Commit()

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
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = __revit__.ActiveUIDocument.Document  # type: Document
app = __revit__.Application  # type: Application
createdoc = doc.Create  # type: Autodesk.Revit.Creation.Document
print('HI!')


# ------------------------------ Functions ------------------------------
def geographical_finding_algorythm(start_point, end_point, object_to_find_name=None, object_to_find_categoty=None,
                                   ignore_id=None):
    if object_to_find_name is None and object_to_find_categoty is None:
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
    else:
        for intersec in allIntersections:
            if intersec.Name == object_to_find_name:
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


# ------------------------------ MAIN ------------------------------
selection = uidoc.Selection.GetElementIds()
if len(selection) != 1:
    TaskDialog.Show("Selection Error", "Please select only one element. Plugin will now stop.")
    sys.exit()
windowFamilyObject = doc.GetElement(selection[0])
print("Selected Element:", windowFamilyObject.Symbol.Family.Name)

window_origin = windowFamilyObject.Location.Point
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
    current_window_level_id = windowFamilyObject.LevelId
    current_window_level = doc.GetElement(current_window_level_id)
    # newViewFamilyType = createdoc.NewViewPlan('callout_view', current_window_level, ViewPlanType.FloorPlan)
    # newViewFamilyType = ViewPlan.Create('callout_view', current_window_level, ViewPlanType.FloorPlan)

    view_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()

    # TODO:
    #  Structural plan (not floor)

    # print('123123123')
    # for vt in view_types:
    #     print(vt)
    # print('123123123')

    fec = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    for x in fec:
        if ViewFamily.StructuralPlan == x.ViewFamily:
            fec = x
            break

    print(fec)

    structuralPlan = ViewPlan.Create(doc, fec.Id, current_window_level_id)


    # view_types_plans = [vt for vt in view_types if vt.ViewFamily == ViewFamily.FloorPlan]
    # floor_plan_type = view_types_plans[0]
    #
    # newViewFamilyType = ViewPlan.Create(
    #     doc,
    #     floor_plan_type.Id,
    #     current_window_level_id
    # )

    end_point_left = window_origin + (win_width + 200 / 30.48) * get_wall_direction_vector(host_object)
    end_point_right = window_origin - (win_width + 200 / 30.48) * get_wall_direction_vector(host_object)

    windows = geographical_finding_algorythm(
        window_origin,
        end_point_left,
        object_to_find_name=windowFamilyObject.Name,
        ignore_id=windowFamilyObject.Id)
    left_offset = 120 / 30.48
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
        if len(windows):
            left_offset = (best_distance - win_width) * 30.48
    else:
        walls = geographical_finding_algorythm(
            window_origin,
            end_point_left,
            object_to_find_categoty=host_object.Category,
            ignore_id=host_object.Id)
        if len(walls):
            left_distance = walls[0].Location.Curve.Distance(window_origin)
            left_offset = (left_distance - win_width - 10 / 30.48) * 30.48
    windows = geographical_finding_algorythm(window_origin, end_point_right,
                                             object_to_find_name=windowFamilyObject.Name,
                                             ignore_id=windowFamilyObject.Id)
    right_offset = 120
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
        right_offset = (best_distance - win_width) * 30.48
    else:
        walls = geographical_finding_algorythm(window_origin, end_point_left,
                                               object_to_find_categoty=host_object.Category, ignore_id=host_object.Id)
        if len(walls):
            right_distance = walls[0].Location.Curve.Distance(window_origin)
            right_offset = (right_distance - win_width - 10 / 30.48) * 30.48


    # print(vector)
    window_bbox = windowFamilyObject.get_BoundingBox(None)
    perpendicular_vector = XYZ(-vector.Y, vector.X, vector.Z)
    # print(perpendicular_vector)
    # print('st_x', ((right_offset / 30.48 + win_width) * vector.X + perpendicular_vector.X * 10 / 30.48) * 30.48, window_bbox.Min.X, window_bbox.Max.X)
    # print('st_y', ((right_offset / 30.48 + win_width) * vector.Y + perpendicular_vector.Y * 10 / 30.48) * 30.48, window_bbox.Min.Y, window_bbox.Max.Y)
    # print('left offset', left_offset)
    # print('right offset', right_offset)

    start_point = XYZ(
        # window_bbox.Min.X + ((right_offset / 30.48 + win_width) * vector.X + perpendicular_vector.X * 0 / 30.48),
        # window_bbox.Min.Y + ((right_offset / 30.48 + win_width) * vector.Y + perpendicular_vector.Y * 0 / 30.48),
        window_bbox.Min.X + (right_offset / 30.48 + 2 * win_width) * vector.X,
        window_bbox.Min.Y + (right_offset / 30.48 + 2 * win_width) * vector.Y,
        window_origin.Z
    )

    end_point = XYZ(
        # window_bbox.Max.X - ((left_offset / 30.48 + win_width) * vector.X - perpendicular_vector.X * 0 / 30.48),
        # window_bbox.Max.Y - ((left_offset / 30.48 + win_width) * vector.Y - perpendicular_vector.Y * 0 / 30.48),
        window_bbox.Max.X - (left_offset / 30.48 + 2 * win_width) * vector.X,
        window_bbox.Max.Y - (left_offset / 30.48 + 2 * win_width) * vector.Y,
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

    # Set template
    views = FilteredElementCollector(doc).OfClass(View).ToElements()
    viewTemplates = [v for v in views if v.IsTemplate and "MAMAD_Window_Callout" in v.Name.ToString()]
    try:
        callout.ApplyViewTemplateParameters(viewTemplates[0])
    except IndexError:
        print('!!! There is no view template "MAMAD_Window_Callout"')

    # print(view_range)
    # print(type(callout))

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

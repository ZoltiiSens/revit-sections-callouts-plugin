# ------------------------------ IMPORTS ------------------------------
import math
import sys
import Autodesk

from Autodesk.Revit.ApplicationServices import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *

from Autodesk.Revit.DB.Structure import *
from System.Collections.Generic import List

# ------------------------------ VARIABLES ------------------------------
__title__ = "Start tagging"
__doc__ = """
Creates 3 sections and one callout and tags rebars on it 
"""
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = __revit__.ActiveUIDocument.Document  # type: Document
app = __revit__.Application  # type: Application
print('HI!')


# ------------------------------ VERIFYING ------------------------------
# import requests
# from subprocess import check_output
# from tempfile import gettempdir
# import os
#
#
# def check_jot_state(plugin_id):
#     try:
#         serials = check_output('wmic diskdrive get Name, SerialNumber').decode().split('\n')[1:]
#         drive_serial = None
#         for serial in serials:
#             if 'DRIVE0' in serial:
#                 drive_serial = serial.split('DRIVE0')[-1].strip()
#         with open(os.path.sep.join(gettempdir().split(os.path.sep)[:-1]) + os.path.sep + 'jot.tmp', 'r') as file:
#             jot_token = file.read()
#         params = {
#             'plugin_id': plugin_id,
#             'machine_ids': drive_serial + '---' + 'smth',
#             'jot_token': jot_token
#         }
#         resp = requests.get('http://localhost:7878/auth-plugin', params=params)
#         if resp.content == 'true':
#             print('Accessed')
#         else:
#             print('You have no access to this plugin. Details:')
#             print(resp.content)
#             sys.exit()
#     except Exception as err:
#         print(err)
#         sys.exit()
#
#
# check_jot_state(5)


# ------------------------------ Functions ------------------------------
def geographical_finding_algorythm(start_point, end_point, object_to_find_name=None, object_to_find_categoty=None,
                                   object_to_find_builtin_category=None, ignore_id=None):
    if object_to_find_name is None and object_to_find_categoty is None and object_to_find_builtin_category is None:
        return None

    min_x, max_x = min(start_point.X, end_point.X), max(start_point.X, end_point.X)
    min_y = min(start_point.Y, end_point.Y)
    min_z = min(start_point.Z, end_point.Z)
    max_y = max(start_point.Y, end_point.Y)
    max_z = max(start_point.Z, end_point.Z)
    outline = Outline(XYZ(min_x, min_y, min_z), XYZ(max_x, max_y, max_z))

    bbox_filter = BoundingBoxIntersectsFilter(outline)
    all_intersections = FilteredElementCollector(doc).WherePasses(bbox_filter)
    intersections = []
    if object_to_find_categoty is not None:
        for intersection in all_intersections:
            if hasattr(intersection.Category, 'Name') and intersection.Category.Name == object_to_find_categoty.Name:
                if ignore_id == intersection.Id:
                    continue
                intersections.append(intersection)
    elif object_to_find_name is not None:
        for intersection in all_intersections:
            if intersection.Name == object_to_find_name:
                if ignore_id == intersection.Id:
                    continue
                intersections.append(intersection)
    elif object_to_find_builtin_category is not None:
        all_intersections = all_intersections.OfCategory(object_to_find_builtin_category)
        for intersection in all_intersections:
            if ignore_id == intersection.Id:
                continue
            intersections.append(intersection)
    return intersections


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
        print(':cross_mark: ERROR! There is no necessary floor category "Flooes". Plugin will stop now...')
        raise Exception('Can\'t find!')

    top_floor = geographical_finding_algorythm(
        current_window_bbox.Min,
        XYZ(current_window_bbox.Max.X, current_window_bbox.Max.Y, host_wall_bbox.Max.Z + 3),
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
    top_offset_in_cm = UnitUtils.ConvertToInternalUnits(abs(top_floor_offset) * 30.48 + 30, UnitTypeId.Centimeters)
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
        object_to_find_builtin_category=BuiltInCategory.OST_Rebar)
    target_rebar = None
    if rebars is None:
        return
    for reb in rebars:
        if spacing - 0.5 < reb.MaxSpacing * 30.48 < spacing + 0.5 and reb.Quantity == quantity:
            target_rebar = reb
            break
    if target_rebar is None:
        print(':warning_sign: WARNING! Can\'t find rebar set(front view)')
        return
    else:
        target_rebar.SetUnobscuredInView(doc.ActiveView, True)
        target_rebar.SetPresentationMode(view, RebarPresentationMode.Select)
        for i in range(target_rebar.NumberOfBarPositions):
            target_rebar.SetBarHiddenStatus(view, i, True)
        target_rebar.SetBarHiddenStatus(view, which_to_show, False)
        return target_rebar


def find_rebars_on_view(view):
    return FilteredElementCollector(doc, view.Id).OfCategory(
        BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()


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
    if 'Horizontal_Bars' not in result.keys():
        print(':cross_mark: ERROR! There is no necessary tag type "Horizontal_Bars". Plugin will stop now...')
        raise Exception('Can\'t find!')
    if 'Column_Vertical' not in result.keys():
        print(':cross_mark: ERROR! There is no necessary tag type "Column_Vertical". Plugin will stop now...')
        raise Exception('Can\'t find!')
    if 'Wall&Col_Vertical+Length' not in result.keys():
        print(':cross_mark: ERROR! There is no necessary tag type "Wall&Col_Vertical+Length". Plugin will stop now...')
        raise Exception('Can\'t find!')
    if 'Link&U-Shape+Length' not in result.keys():
        print(':cross_mark: ERROR! There is no necessary tag type "Link&U-Shape+Length". Plugin will stop now...')
        raise Exception('Can\'t find!')
    for bdt in bendingDetailTypes:
        name = bdt.LookupParameter("Type Name").AsString()
        if 'Bending Detail 2 (No hooks)' in name:
            result['Bending Detail 2 (No hooks)'] = bdt
    if 'Bending Detail 2 (No hooks)' not in result.keys():
        print(':cross_mark: ERROR! There is no necessary bending detail type "Bending Detail 2 (No hooks)". Plugin will stop now...')
        raise Exception('Can\'t find!')
    return result


def get_shapes_ids():
    rebars = FilteredElementCollector(doc).OfClass(RebarShape)
    U_5_shape = None
    Link_4 = None
    for rebar in rebars:
        for param in rebar.Parameters:
            if param.AsString() == '5_U-Shape':
                U_5_shape = rebar
                break
            if param.AsString() == '4_Link':
                Link_4 = rebar
                break
        if U_5_shape is not None and Link_4 is not None:
            break
    if U_5_shape is None:
        print(':cross_mark: ERROR! There is no necessary shape "5_U-Shape". Plugin will stop now...')
        raise Exception('Can\'t find!')
    if Link_4 is None:
        print(':cross_mark:ERROR! There is no necessary shape "4_Link". Plugin will stop now...')
        raise Exception('Can\'t find!')
    return {
        '5_U-Shape': U_5_shape,
        '4_Link': Link_4
    }


def create_rebar_tag(view, all_rebars, tag_mode, tag_orientation, tag_type_name, tag_position, partitionName,
                     leader_end_condition=LeaderEndCondition.Attached, create_only_for_one=False):
    filtered_rebars = []
    for rebar in all_rebars:
        if rebar.LookupParameter("Partition").AsString() == partitionName:
            filtered_rebars.append(rebar)
    tag = None
    for i, rebar in enumerate(filtered_rebars):
        for subelement in rebar.GetSubelements():
            if tag is None:
                print(rebar)
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


def create_rebar_tag_depending_on_rebar(view, all_rebars, tag_mode, tag_orientation, tag_type_name, tag_position,
                                        partitionName, leader_end_condition=LeaderEndCondition.Attached,
                                        create_only_for_one=False, has_leader=True):
    filtered_rebars = []
    for rebar in all_rebars:
        if rebar.LookupParameter("Partition").AsString() == partitionName:
            filtered_rebars.append(rebar)
    tag = None
    for i, rebar in enumerate(filtered_rebars):
        for subelement in rebar.GetSubelements():
            if tag is None:
                print(rebar)
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
                tag_coordinates = XYZ(
                    (rebar_bbox.Min.X + rebar_bbox.Max.X) / 2,
                    (rebar_bbox.Min.Y + rebar_bbox.Max.Y) / 2,
                    (rebar_bbox.Min.Z + rebar_bbox.Max.Z) / 2
                )
                tag.TagHeadPosition = tag_position + tag_coordinates

                tag.LeaderEndCondition = leader_end_condition
                tag.HasLeader = has_leader
                if create_only_for_one:
                    break
                continue
            lll = List[Reference]()
            lll.Add(subelement.GetReference())
            tag.AddReferences(lll)
    return tag


def create_rebar_tag_depending_on_window(view, all_rebars, tag_mode, tag_orientation, tag_type_name, tag_position,
                                        partitionName, leader_end_condition=LeaderEndCondition.Attached,
                                        create_only_for_one=False, has_leader=True):
    filtered_rebars = []
    for rebar in all_rebars:
        if rebar.LookupParameter("Partition").AsString() == partitionName:
            filtered_rebars.append(rebar)
    tag = None
    for i, rebar in enumerate(filtered_rebars):
        for subelement in rebar.GetSubelements():
            if tag is None:
                print(rebar)
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

                x_coord = rebar_bbox.Min.X if abs(window_origin.X - rebar_bbox.Min.X) < abs(window_origin.X - rebar_bbox.Max.X) else rebar_bbox.Max.X
                y_coord = rebar_bbox.Min.Y if abs(window_origin.Y - rebar_bbox.Min.Y) < abs(window_origin.Y - rebar_bbox.Max.Y) else rebar_bbox.Max.Y

                tag_coordinates = XYZ(
                    x_coord,
                    y_coord,
                    (rebar_bbox.Min.Z + rebar_bbox.Max.Z) / 2
                )
                tag.TagHeadPosition = tag_position + tag_coordinates

                tag.LeaderEndCondition = leader_end_condition
                tag.HasLeader = has_leader
                if create_only_for_one:
                    break
                continue
            lll = List[Reference]()
            lll.Add(subelement.GetReference())
            tag.AddReferences(lll)
    return tag


def check_type_of_ulink_hor_rebar(view, all_rebars, rebarShapes):
    filtered_rebars = []
    filtered_rebars_hor_in = []
    for rebar in all_rebars:
        if rebar.LookupParameter("Partition").AsString() == 'U/Link_Hor':
            filtered_rebars.append(rebar)
    for rebar in all_rebars:
        if rebar.LookupParameter("Partition").AsString() == 'Hor_In':
            filtered_rebars_hor_in.append(rebar)
    if len(filtered_rebars) == 0:
        # print('0')
        pass
    elif len(filtered_rebars) == 1 and filtered_rebars[0].GetShapeId() == rebarShapes['5_U-Shape'].Id and len(filtered_rebars_hor_in) > 0:
        # print('1')
        # 3
        create_bending_detail(
            view,
            filtered_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(0 * vector.X + -3.9 * perpendicular_vector.X, 0 * vector.Y + -3.9 * perpendicular_vector.Y, 0),
            'U/Link_Hor'
        )

        create_rebar_tag_depending_on_rebar(
            view,
            filtered_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(0 * vector.X + -3.15 * perpendicular_vector.X, 0 * vector.Y + -3.15 * perpendicular_vector.Y, 0),
            'U/Link_Hor',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            view,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(0 * vector.X + -1.75 * perpendicular_vector.X, 0 * vector.Y + -1.75 * perpendicular_vector.Y, 0),
            'Hor_In'
        )

        create_rebar_tag_depending_on_rebar(
            view,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(0 * vector.X + -1.55 * perpendicular_vector.X, 0 * vector.Y + -1.55 * perpendicular_vector.Y, 0),
            'Hor_In',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            view,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(0 * vector.X + -2.05 * perpendicular_vector.X, 0 * vector.Y + -2.05 * perpendicular_vector.Y, 0),
            'Hor_Out'
        )

        create_rebar_tag_depending_on_rebar(
            view,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(0 * vector.X + -2.25 * perpendicular_vector.X, 0 * vector.Y + -2.25 * perpendicular_vector.Y, 0),
            'Hor_Out',
            create_only_for_one=True,
            has_leader=False)
    elif len(filtered_rebars) == 2 and filtered_rebars[0].GetShapeId() == rebarShapes['5_U-Shape'].Id and len(filtered_rebars_hor_in) > 0:
        # print('2')
        for reb in filtered_rebars:
            create_bending_detail(
                view,
                [reb],
                'Bending Detail 2 (No hooks)',
                XYZ(0 * vector.X + -3.9 * perpendicular_vector.X, 0 * vector.Y + -3.9 * perpendicular_vector.Y, 0),
                'U/Link_Hor'
            )

            create_rebar_tag_depending_on_rebar(
                view,
                [reb],
                TagMode.TM_ADDBY_CATEGORY,
                TagOrientation.Vertical,
                'Link&U-Shape+Length',
                XYZ(0 * vector.X + -3.15 * perpendicular_vector.X, 0 * vector.Y + -3.15 * perpendicular_vector.Y, 0),
                'U/Link_Hor',
                create_only_for_one=True,
                has_leader=False)
        # 3
        # create_bending_detail(
        #     view,
        #     filtered_rebars[:0],
        #     'Bending Detail 2 (No hooks)',
        #     XYZ(0 * vector.X + -3.9 * perpendicular_vector.X, 0 * vector.Y + -3.9 * perpendicular_vector.Y, 0),
        #     'U/Link_Hor'
        # )
        #
        # create_rebar_tag_depending_on_rebar(
        #     view,
        #     filtered_rebars[:0],
        #     TagMode.TM_ADDBY_CATEGORY,
        #     TagOrientation.Vertical,
        #     'Link&U-Shape+Length',
        #     XYZ(0 * vector.X + -3.15 * perpendicular_vector.X, 0 * vector.Y + -3.15 * perpendicular_vector.Y, 0),
        #     'U/Link_Hor',
        #     create_only_for_one=True,
        #     has_leader=False)
        #
        # create_bending_detail(
        #     view,
        #     filtered_rebars[1:],
        #     'Bending Detail 2 (No hooks)',
        #     XYZ(0 * vector.X + -3.9 * perpendicular_vector.X, 0 * vector.Y + -3.9 * perpendicular_vector.Y, 0),
        #     'U/Link_Hor'
        # )
        #
        # create_rebar_tag_depending_on_rebar(
        #     view,
        #     filtered_rebars[1:],
        #     TagMode.TM_ADDBY_CATEGORY,
        #     TagOrientation.Vertical,
        #     'Link&U-Shape+Length',
        #     XYZ(0 * vector.X + -3.15 * perpendicular_vector.X, 0 * vector.Y + -3.15 * perpendicular_vector.Y, 0),
        #     'U/Link_Hor',
        #     create_only_for_one=True,
        #     has_leader=False)

        create_bending_detail(
            view,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(0 * vector.X + -3.15 * perpendicular_vector.X, 0 * vector.Y + -3.15 * perpendicular_vector.Y, 0),
            'Hor_In'
        )

        create_rebar_tag_depending_on_rebar(
            view,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(0 * vector.X + -2.95 * perpendicular_vector.X, 0 * vector.Y + -2.95 * perpendicular_vector.Y, 0),
            'Hor_In',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            view,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(0 * vector.X + -4.65 * perpendicular_vector.X, 0 * vector.Y + -4.65 * perpendicular_vector.Y, 0),
            'Hor_Out'
        )

        create_rebar_tag_depending_on_rebar(
            view,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(0 * vector.X + -4.45 * perpendicular_vector.X, 0 * vector.Y + -4.45 * perpendicular_vector.Y, 0),
            'Hor_Out',
            create_only_for_one=True,
            has_leader=False)
    elif len(filtered_rebars) == 2 and filtered_rebars[0].GetShapeId() == rebarShapes['5_U-Shape'].Id and len(filtered_rebars_hor_in) == 0:
        # print('3')
        # 2
        create_bending_detail(
            view,
            filtered_rebars[:0],
            'Bending Detail 2 (No hooks)',
            XYZ(0 * vector.X + -3.9 * perpendicular_vector.X, 0 * vector.Y + -3.9 * perpendicular_vector.Y, 0),
            'U/Link_Hor'
        )

        create_rebar_tag_depending_on_rebar(
            view,
            filtered_rebars[:0],
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(0 * vector.X + -3.15 * perpendicular_vector.X, 0 * vector.Y + -3.15 * perpendicular_vector.Y, 0),
            'U/Link_Hor',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            view,
            filtered_rebars[1:],
            'Bending Detail 2 (No hooks)',
            XYZ(0 * vector.X + -2.1 * perpendicular_vector.X, 0 * vector.Y + -2.1 * perpendicular_vector.Y, 0),
            'U/Link_Hor'
        )

        create_rebar_tag_depending_on_rebar(
            view,
            filtered_rebars[1:],
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(0 * vector.X + -2.15 * perpendicular_vector.X, 0 * vector.Y + -15 * perpendicular_vector.Y, 0),
            'U/Link_Hor',
            create_only_for_one=True,
            has_leader=False)
    elif len(filtered_rebars) == 1 and filtered_rebars[0].GetShapeId() == rebarShapes['4_Link'].Id:
        print('4 hahaha')


def create_bending_detail(view, all_rebars, tag_type_name, tag_position, partitionName, create_only_for_one=False):
    try:
        filtered_rebars = []
        for rebar in all_rebars:
            if rebar.LookupParameter("Partition").AsString() == partitionName:
                filtered_rebars.append(rebar)
        bdetail = None
        for i, rebar in enumerate(filtered_rebars):
            if bdetail is None:
                rebar_bbox = rebar.get_BoundingBox(None)
                bdetail_coordinates = XYZ(
                    (rebar_bbox.Min.X + rebar_bbox.Max.X) / 2,
                    (rebar_bbox.Min.Y + rebar_bbox.Max.Y) / 2,
                    (rebar_bbox.Min.Z + rebar_bbox.Max.Z) / 2
                )
                bdetailPosition = tag_position + bdetail_coordinates

                bdetail = RebarBendingDetail.Create(
                    document=doc,
                    viewId=view.Id,
                    reinforcementElementId=rebar.Id,
                    reinforcementElementSubelementKey=0,
                    bendingDetailType=tagTypes[tag_type_name],
                    position=bdetailPosition,
                    rotation=0)
                if create_only_for_one:
                    break
                break
        return bdetail
    except:
        print('-- Error happened while bending detail was creating! It wasn\'t been created: ' + partitionName)


def create_text_note(view, text, position):
    text_note_types = FilteredElementCollector(doc).OfClass(TextNoteType)
    text_note_type = None
    for tnt in text_note_types:
        name = tnt.LookupParameter("Type Name").AsString()
        if '3.5mm Ariall with border' in name:
            text_note_type = tnt
            break
    if text_note_type is None:
        print(':cross_mark: ERROR! There is no "3.5mm Ariall with border" text note start. Plugin will stop now...')
        return
        # TODO: change return to raise
        # raise Exception('Can\'t find!')
    text_note_options = TextNoteOptions(text_note_type.Id)
    text_note = TextNote.Create(doc, view.Id, position, text, text_note_options)
    return text_note


def create_detail_component(view, location):
    collector = FilteredElementCollector(doc)
    family_symbols = collector.OfCategory(BuiltInCategory.OST_DetailComponents).OfClass(FamilySymbol)
    break_line_type = None
    for fs in family_symbols:
        fs_name = fs.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
        if fs_name == 'Break Line':
            break_line_type = fs
            break

    if break_line_type is None:
        print(':cross_mark: ERROR! There is no necessary detail type "Break Line". Plugin will stop now...')
        raise Exception('Can\'t find!')

    detail_component = doc.Create.NewFamilyInstance(location, break_line_type, view)
    return detail_component


def create_spot_elevation(view, element, point, ):
    reference = Reference(element)
    offset_point = point + XYZ(perpendicular_vector.X * 5, perpendicular_vector.Y * 5, 0)

    # TODO: get getting type out of function
    family_symbols = FilteredElementCollector(doc).OfClass(SpotDimensionType)
    spot_dimension_type = None
    for fs in family_symbols:
        fs_name = fs.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        try:
            if fs_name == 'Arrow (Project)':
                spot_dimension_type = fs
                break
        except:
            print(':cross_mark: ERROR! There is no necessary spot elevation type: "Arrow (Project)". Plugin will stop now...')
            raise Exception('Cant\'t find!')

    spot_dimension = doc.Create.NewSpotElevation(
        view,
        reference,
        point,
        offset_point,
        offset_point,
        offset_point,
        True
    )
    spot_dimension.SpotDimensionType = spot_dimension_type


# ------------------------------ MAIN ------------------------------
selection = uidoc.Selection.GetElementIds()
if len(selection) != 1:
    TaskDialog.Show("Selection Error", "Please select only one element. Plugin will now stop.")
    sys.exit()
windowFamilyObject = doc.GetElement(selection[0])
print("Selected Element:", windowFamilyObject.Symbol.Family.Name)

tagTypes = get_tag_types()
rebarShapes = get_shapes_ids()

window_origin = windowFamilyObject.Location.Point
window_bbox = windowFamilyObject.get_BoundingBox(None)
host_object = windowFamilyObject.Host
win_height = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.GENERIC_HEIGHT).AsDouble()
win_width = windowFamilyObject.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble()
win_depth = UnitUtils.ConvertToInternalUnits(40, UnitTypeId.Centimeters)
win_center_point = XYZ((window_bbox.Min.X + window_bbox.Max.X) / 2, (window_bbox.Min.Y + window_bbox.Max.Y) / 2, (window_bbox.Min.Z + window_bbox.Max.Z) / 2)

perpendicular_vector = windowFamilyObject.FacingOrientation
perpendicular_vector = XYZ(-perpendicular_vector.X, -perpendicular_vector.Y, -perpendicular_vector.Z)
vector = XYZ(-perpendicular_vector.Y, perpendicular_vector.X, perpendicular_vector.Z)

wall_bbox = host_object.get_BoundingBox(None)
wall_height = wall_bbox.Max.Z - wall_bbox.Min.Z
wallDepth = UnitUtils.ConvertToInternalUnits(host_object.Width, UnitTypeId.Feet)
# wall_level_id = host_object.LevelId
#
# if wall_level_id:
#     wall_level = doc.GetElement(wall_level_id)  # Get the Level element
# else:
#     print("Wall does not have an associated level. Plugin will now stop.")
#     sys.exit()


# -------------------------------------------------------------------------------- Front view
def get_front_view():
    top_offset, bottom_offset = find_floors_offsets(windowFamilyObject)

    end_point_left = window_origin + (win_width + 200 / 30.48) * vector
    end_point_right = window_origin - (win_width + 200 / 30.48) * vector
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
    win_elevation.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(0)
    win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
    left_rebar_start_point = XYZ(
        window_bbox.Min.X + left_offset * vector.X,
        window_bbox.Min.Y + left_offset * vector.Y,
        window_bbox.Min.Z
    )
    left_rebar_end_point = XYZ(
        window_bbox.Max.X + left_offset * vector.X,
        window_bbox.Max.Y + left_offset * vector.Y,
        window_bbox.Max.Z
    )
    left_reb_5 = find_rebars_by_quantity_and_spacing(win_elevation, window_bbox.Min, window_bbox.Max, 5, 20, 3)
    left_reb_8 = find_rebars_by_quantity_and_spacing(win_elevation, window_bbox.Min, window_bbox.Max, 8, 10, 2)
    right_reb_5 = find_rebars_by_quantity_and_spacing(win_elevation, left_rebar_start_point, left_rebar_end_point, 5,
                                                      20, 3)
    right_reb_8 = find_rebars_by_quantity_and_spacing(win_elevation, left_rebar_start_point, left_rebar_end_point, 8,
                                                      10, 2)

    all_rebars = find_rebars_on_view(win_elevation)
    rebar_ids_to_hide = List[ElementId]()
    for reb in all_rebars:
        if reb.GetHostId() != host_object.Id:
            rebar_ids_to_hide.Add(reb.Id)
    try:
        win_elevation.HideElements(rebar_ids_to_hide)
    except:
        print('There are no rebars to hide')
    # perpendicular_vector = XYZ(-vector.Y, vector.X, vector.Z)

    # all_rebars = geographical_finding_algorythm(
    #     XYZ(
    #         window_origin.X - perpendicular_vector.X * wallDepth / 2 - vector.X * (win_width + 50 / 30.48),
    #         window_origin.Y - perpendicular_vector.X * wallDepth / 2 - vector.Y * (win_width + 50 / 30.48),
    #         window_origin.Z - win_height
    #     ),
    #     XYZ(
    #         window_origin.X + perpendicular_vector.X * wallDepth / 2 + vector.X * (win_width + 50 / 30.48),
    #         window_origin.Y + perpendicular_vector.Y * wallDepth / 2 + vector.Y * (win_width + 50 / 30.48),
    #         window_origin.Z + win_height * 2
    #     ),
    #     object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    # )

    # all_rebars = geographical_finding_algorythm(
    #     XYZ(
    #         window_origin.X - perpendicular_vector.X * wallDepth - vector.X * (win_width + 60 / 30.48),
    #         window_origin.Y - perpendicular_vector.X * wallDepth - vector.Y * (win_width + 60 / 30.48),
    #         window_origin.Z - win_height
    #     ),
    #     XYZ(
    #         window_origin.X + perpendicular_vector.X * wallDepth + vector.X * (win_width + 60 / 30.48),
    #         window_origin.Y + perpendicular_vector.Y * wallDepth + vector.Y * (win_width + 60 / 30.48),
    #         window_origin.Z + win_height * 2
    #     ),
    #     object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    # )

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-vector.X * 1, -vector.Y * 1, win_height + 0.85),
        'WD_Hor_14_up')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Column_Vertical',
        window_origin + XYZ((win_width + 1.2) * -vector.X, (win_width + 1.2) * -vector.Y, win_height / 2),
        'WD_Vert_14')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Horizontal,
        'Horizontal_Bars',
        window_origin + XYZ(-vector.X * 1, -vector.Y * 1, -0.85),
        'WD_Hor_14_down')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Column_Vertical',
        window_origin + XYZ((win_width / 2 - 0.5) * vector.X, (win_width / 2 - 0.5) * vector.Y, win_height / 2),
        # XYZ(-0.2 * vector.X + 2 * perpendicular_vector.X, -0.2 * vector.Y + 2 * perpendicular_vector.Y, 0),
        'WD_Vert_14_sh_in')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Wall&Col_Vertical+Length',
        window_origin + XYZ((win_width / 2 + 0.5) * vector.X, (win_width / 2 + 0.5) * vector.Y, win_height / 2 - 0.2),
        'WD_Vert_14_In')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Wall&Col_Vertical+Length',
        window_origin + XYZ((win_width / 2 + 1.5) * vector.X, (win_width / 2 + 1.5) * vector.Y, win_height / 2 - 0.2),
        'WD_Vert_14_Out')

    create_rebar_tag(
        win_elevation,
        all_rebars,
        TagMode.TM_ADDBY_CATEGORY,
        TagOrientation.Vertical,
        'Column_Vertical',
        window_origin + XYZ((win_width + 1) * vector.X, (win_width + 1) * vector.Y, win_height / 2),
        'WD_Vert_14_sh_out')



    create_detail_component(win_elevation,
                            XYZ(window_origin.X, window_origin.Y, window_origin.Z + win_height / 2) + XYZ(
                                (left_offset + win_width) * vector.X, (left_offset + win_width) * vector.Y, 0))
    create_detail_component(win_elevation,
                            XYZ(window_origin.X, window_origin.Y, window_origin.Z + win_height / 2) + XYZ(
                                (right_offset + win_width) * -vector.X, (right_offset + win_width) * -vector.Y, 0))

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
    end_point_left = window_origin + (win_width + 200 / 30.48) * vector
    end_point_right = window_origin - (win_width + 200 / 30.48) * vector

    windows = geographical_finding_algorythm(
        window_origin,
        end_point_left,
        object_to_find_name=windowFamilyObject.Name,
        ignore_id=windowFamilyObject.Id)
    left_point = window_origin + (win_width + 120 / 30.48) * vector
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
                left_point = window.Location.Point + XYZ(
                    win_width / 2 * vector.X,
                    win_width / 2 * vector.Y,
                    0)
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
                (win_width + left_offset) * vector.X,
                (win_width + left_offset) * vector.Y,
                0)

    windows = geographical_finding_algorythm(
        window_origin,
        end_point_right,
        object_to_find_name=windowFamilyObject.Name,
        ignore_id=windowFamilyObject.Id)
    right_point = window_origin - (win_width + 120 / 30.48) * vector
    if len(windows):
        best_distance = float('inf')
        for window in windows:
            distance = abs(math.sqrt(
                (window_origin.Y - window.Location.Point.Y) ** 2 + (window_origin.X - window.Location.Point.X) ** 2))
            if distance < best_distance:
                best_distance = distance
                right_point = window.Location.Point - XYZ(
                    win_width / 2 * vector.X,
                    win_width / 2 * vector.Y,
                    0)
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
                (win_width + right_offset) * vector.X,
                (win_width + right_offset) * vector.Y,
                0)

    # Plan creation
    current_window_level_id = windowFamilyObject.LevelId
    fec = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    for x in fec:
        if ViewFamily.StructuralPlan == x.ViewFamily:
            fec = x
            break
    structuralPlan = ViewPlan.Create(doc, fec.Id, current_window_level_id)

    # perpendicular_vector = XYZ(-vector.Y, vector.X, vector.Z)

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

    callout.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(0)

    # Modifying viewRange
    sill_height_param = windowFamilyObject.LookupParameter("Sill Height")
    if sill_height_param:
        sill_height_param = sill_height_param.AsDouble()  # Sill height is typically stored in feet
    else:
        print(':cross_mark: ERROR! Sill Height parameter not found in the window family instance. Plugin will stop now...')
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
        print(':cross_mark: ERROR! There is no necessary view template "MAMAD_Window_Callout". Plugin will stop now...')
        # TODO: get raise back
        # raise Exception('Can\'t find!')

    # all_rebars = find_rebars_on_view(callout)

    rebars_sets = []
    rebars_sets.append(geographical_finding_algorythm(
        XYZ(
            window_origin.X + perpendicular_vector.X * wallDepth / 2,
            window_origin.Y + perpendicular_vector.Y * wallDepth / 2,
            window_origin.Z
        ),
        XYZ(
            window_origin.X - perpendicular_vector.X * wallDepth / 2 - vector.X * 320 / 30.48,
            window_origin.Y - perpendicular_vector.Y * wallDepth / 2 - vector.Y * 320 / 30.48,
            window_origin.Z + 20 / 30.48
        ),
        object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    ))
    rebars_sets.append(geographical_finding_algorythm(
        XYZ(
            window_origin.X + perpendicular_vector.X * wallDepth / 2,
            window_origin.Y + perpendicular_vector.Y * wallDepth / 2,
            window_origin.Z
        ),
        XYZ(
            window_origin.X - perpendicular_vector.X * wallDepth / 2 + vector.X * 320 / 30.48,
            window_origin.Y - perpendicular_vector.Y * wallDepth / 2 + vector.Y * 320 / 30.48,
            window_origin.Z + 20 / 30.48
        ),
        object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    ))
    for all_rebars in rebars_sets:
        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(0.5 * vector.X + 1.9 * perpendicular_vector.X, 0.5 * vector.Y + 1.9 * perpendicular_vector.Y, 0),
            'Vert_Corner')

        create_rebar_tag_depending_on_window(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.3 * perpendicular_vector.X + 0.8 * vector.X, 2.3 * perpendicular_vector.Y + 0.8 * vector.Y, 0),
            'Vert_Out',
            create_only_for_one=True)

        create_rebar_tag_depending_on_window(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.8 * perpendicular_vector.X - 0 * vector.X, 1.8 * perpendicular_vector.Y - 0 * vector.Y, 0),
            'Vert_In',
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(-0.4 * vector.X + 1.9 * perpendicular_vector.X, -0.4 * vector.Y + 1.9 * perpendicular_vector.Y, 0),
            'WD_Vert_14_sh_out')

        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(-1 * vector.X + 1.5 * perpendicular_vector.X, -1 * vector.Y + 1.5 * perpendicular_vector.Y, 0),
            'WD_Vert_14_In',
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(0 * vector.X + 2 * perpendicular_vector.X, 0 * vector.Y + 2 * perpendicular_vector.Y, 0),
            'WD_Vert_14_Out',
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(0 * vector.X + 2.6 * perpendicular_vector.X, 0 * vector.Y + 2.6 * perpendicular_vector.Y, 0),
            'WD_Vert_14_sh_in')

        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(-0.7 * vector.X + 2.3 * perpendicular_vector.X, -0.7 * vector.Y + 2.3 * perpendicular_vector.Y, 0),
            'Vert_8_sh')

        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.0 * vector.X + 2 * perpendicular_vector.X, 1.0 * vector.Y + 2 * perpendicular_vector.Y, 0),
            'Vert_8')

        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(0.6 * vector.X + 2.5 * perpendicular_vector.X, 0.6 * vector.Y + 2.5 * perpendicular_vector.Y, 0),
            'WD_Vert_14')


        create_bending_detail(
            callout,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-1.5 * vector.X + -1.2 * perpendicular_vector.X, -1.5 * vector.Y + -1.2 * perpendicular_vector.Y, 0),
            'U_Hor_Small_sh'
        )
        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-1.5 * vector.X + -0.85 * perpendicular_vector.X, -1.5 * vector.Y + -0.85 * perpendicular_vector.Y, 0),
            'U_Hor_Small_sh',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            callout,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(1.7 * vector.X + -2.35 * perpendicular_vector.X, 1.7 * vector.Y + -2.35 * perpendicular_vector.Y, 0),
            'U_Hor_Small'
        )
        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(1.9 * vector.X + -2.0 * perpendicular_vector.X, 1.9 * vector.Y + -2.0 * perpendicular_vector.Y, 0),
            'U_Hor_Small',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            callout,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(0 * vector.X + -1.7 * perpendicular_vector.X, 0 * vector.Y + -1.7 * perpendicular_vector.Y, 0),
            'U_Hor'
        )
        create_rebar_tag_depending_on_rebar(
            callout,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(0 * vector.X + -1.15 * perpendicular_vector.X, 0 * vector.Y + -1.15 * perpendicular_vector.Y, 0),
            'U_Hor',
            create_only_for_one=True,
            has_leader=False)

        check_type_of_ulink_hor_rebar(callout, all_rebars, rebarShapes)


    create_text_note(callout, b'\xd7\x97\xd7\x95\xd7\xa5'.decode('UTF-8'), XYZ(
        window_origin.X - 4 * perpendicular_vector.X - 1 * vector.X,
        window_origin.Y - 4 * perpendicular_vector.Y - 1 * vector.Y,
        window_origin.Z))

    create_text_note(callout, b'\xd7\xa4\xd7\xa0\xd7\x99\xd7\x9d'.decode('UTF-8'), XYZ(
        window_origin.X + 4 * perpendicular_vector.X - 1 * vector.X,
        window_origin.Y + 4 * perpendicular_vector.Y - 1 * vector.Y,
        window_origin.Z))

    create_detail_component(callout, XYZ(
        left_point.X + -vector.X * 10 / 30.48,
        left_point.Y + -vector.Y * 10 / 30.48,
        window_origin.Z
    ))
    create_detail_component(callout, XYZ(
        right_point.X + vector.X * 10 / 30.48,
        right_point.Y + vector.Y * 10 / 30.48,
        window_origin.Z
    ))

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

    transform = Transform.Identity
    transform.Origin = window_origin
    vector_perp = get_wall_direction_vector(host_object).CrossProduct(XYZ.BasisZ)
    transform.BasisX = vector_perp
    transform.BasisY = XYZ.BasisZ
    transform.BasisZ = vector_perp.CrossProduct(XYZ.BasisZ)
    section_box = BoundingBoxXYZ()
    offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
    offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)

    if (windowFamilyObject.FacingOrientation.X == -get_wall_direction_vector(host_object).Y and get_wall_direction_vector(host_object).Y != 0) or (
            windowFamilyObject.FacingOrientation.Y == get_wall_direction_vector(host_object).X and get_wall_direction_vector(host_object).X != 0):
        exterior_offset = offset_70cm
        interior_offset = offset_50cm
    else:
        exterior_offset = offset_50cm
        interior_offset = offset_70cm

    section_box.Min = XYZ(
        -wallDepth / 2 - exterior_offset, 0 - bottom_offset, -win_width / 2
    )
    section_box.Max = XYZ(
        wallDepth / 2 + interior_offset,
        win_height + top_offset,
        -win_width / 2 + UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters)
    )

    section_box.Transform = transform
    section_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
    views = FilteredElementCollector(doc).OfClass(View).ToElements()
    viewTemplates = [v for v in views if v.IsTemplate and "Section_Reinforcement" in v.Name.ToString()]

    win_elevation = ViewSection.CreateSection(doc, section_type_id, section_box)
    win_elevation.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(0)

    try:
        win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
    except:
        print(':cross_mark: ERROR! There is no necessary view template "Section_Reinforcement". Plugin will stop now...')
        raise Exception('Can\'t find!')
    win_elevation.Scale = 25


    # rebars_sets = []
    # rebars_sets.append(geographical_finding_algorythm(
    #     XYZ(window_bbox.Max.X, window_bbox.Max.Y, window_origin.Z),
    #     XYZ(window_bbox.Min.X, window_bbox.Min.Y, window_origin.Z + 10),
    #     object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    # ))
    # rebars_sets.append(geographical_finding_algorythm(
    #     XYZ(window_bbox.Max.X, window_bbox.Max.Y, window_origin.Z),
    #     XYZ(window_bbox.Min.X, window_bbox.Min.Y, window_origin.Z - 5.6),
    #     object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    # ))

    rebars_sets = []
    rebars_sets.append(geographical_finding_algorythm(
        XYZ(window_bbox.Max.X, window_bbox.Max.Y, (window_bbox.Min.Z + window_bbox.Max.Z) / 2),
        XYZ(window_bbox.Min.X, window_bbox.Min.Y, (window_bbox.Min.Z + window_bbox.Max.Z) / 2 + 8),
        object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    ))
    rebars_sets.append(geographical_finding_algorythm(
        XYZ(window_bbox.Max.X, window_bbox.Max.Y, (window_bbox.Min.Z + window_bbox.Max.Z) / 2),
        XYZ(window_bbox.Min.X, window_bbox.Min.Y, (window_bbox.Min.Z + window_bbox.Max.Z) / 2 - 8),
        object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    ))

    for all_rebars in rebars_sets:
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.7 * perpendicular_vector.X, 2.7 * perpendicular_vector.Y, -0.8),
            'Hor_T_Corner')

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.9 * perpendicular_vector.X, 1.9 * perpendicular_vector.Y, -0.8),
            'Hor_B_Corner')

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(3.4 * perpendicular_vector.X, 3.4 * perpendicular_vector.Y, -0.8),
            'Hor_Out_Above',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.29 * perpendicular_vector.X, 2.29 * perpendicular_vector.Y, -1.4),
            'Hor_In_Above',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.3 * perpendicular_vector.X, 2.3 * perpendicular_vector.Y, 0.1),
            'WD_Hor_14_up')

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.4 * perpendicular_vector.X, 2.4 * perpendicular_vector.Y, -0.6),
            'Hor_8_up')

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.9 * perpendicular_vector.X, 1.9 * perpendicular_vector.Y, 0),
            'U_Hor',
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.4 * perpendicular_vector.X, 2.4 * perpendicular_vector.Y, 1),
            'Hor_8_down')

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.3 * perpendicular_vector.X, 2.3 * perpendicular_vector.Y, -0.1),
            'WD_Hor_14_down')

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(3.4 * perpendicular_vector.X, 3.4 * perpendicular_vector.Y, 0.8),
            'Hor_Out_Below',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.29 * perpendicular_vector.X, 2.29 * perpendicular_vector.Y, 1.4),
            'Hor_In_Below',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-1.3 * perpendicular_vector.X, -1.3 * perpendicular_vector.Y, 1.5),
            'U_Vert_Small_down'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-0.85 * perpendicular_vector.X, -0.85 * perpendicular_vector.Y, 1.5),
            'U_Vert_Small_down',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            view=win_elevation,
            all_rebars=all_rebars,
            tag_type_name='Bending Detail 2 (No hooks)',
            tag_position=XYZ(-2.4 * perpendicular_vector.X, -2.4 * perpendicular_vector.Y, -1.9),
            partitionName='U_Vert_Small_up'
        )

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-1.95 * perpendicular_vector.X, -1.95 * perpendicular_vector.Y, -1.9),
            'U_Vert_Small_up',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-3.2 * perpendicular_vector.X, -3.2 * perpendicular_vector.Y, 0),
            'U_Above_Out'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-3.55 * perpendicular_vector.X, -3.55 * perpendicular_vector.Y, 0),
            'U_Above_Out',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-1.8 * perpendicular_vector.X, -1.8 * perpendicular_vector.Y, 0),
            'U_Above_In'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-1.45 * perpendicular_vector.X, -1.45 * perpendicular_vector.Y, 0),
            'U_Above_In',
            create_only_for_one=True,
            has_leader=False)

        # create_bending_detail(
        #     win_elevation,
        #     all_rebars,
        #     'Bending Detail 2 (No hooks)',
        #     XYZ(4.35 * perpendicular_vector.X, 4.35 * perpendicular_vector.Y, 0),
        #     'Bending_5'
        # )
        # create_rebar_tag_depending_on_rebar(
        #     win_elevation,
        #     all_rebars,
        #     TagMode.TM_ADDBY_CATEGORY,
        #     TagOrientation.Vertical,
        #     'Link&U-Shape+Length',
        #     XYZ(5.35 * perpendicular_vector.X, 5.35 * perpendicular_vector.Y, 0),
        #     'Bending_5',
        #     create_only_for_one=True,
        #     has_leader=False)

        # create_bending_detail(
        #     win_elevation,
        #     all_rebars,
        #     'Bending Detail 2 (No hooks)',
        #     XYZ(1.8 * perpendicular_vector.X, 1.8 * perpendicular_vector.Y, 0),
        #     'Bending_6'
        # )
        # create_rebar_tag_depending_on_rebar(
        #     win_elevation,
        #     all_rebars,
        #     TagMode.TM_ADDBY_CATEGORY,
        #     TagOrientation.Vertical,
        #     'Link&U-Shape+Length',
        #     XYZ(1.45 * perpendicular_vector.X, 1.45 * perpendicular_vector.Y, 0),
        #     'Bending_6',
        #     create_only_for_one=True,
        #     has_leader=False)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-3.2 * perpendicular_vector.X, -3.2 * perpendicular_vector.Y, 0),
            'L_Out'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-3.55 * perpendicular_vector.X, -3.55 * perpendicular_vector.Y, 0),
            'L_Out',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-1.8 * perpendicular_vector.X, -1.8 * perpendicular_vector.Y, 0),
            'L_In'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-1.45 * perpendicular_vector.X, -1.45 * perpendicular_vector.Y, 0),
            'L_In',
            create_only_for_one=True,
            has_leader=False)
        #
        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-2.5 * perpendicular_vector.X, -2.5 * perpendicular_vector.Y, 0),
            'U_Vert_Starter'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-2.15 * perpendicular_vector.X, -2.15 * perpendicular_vector.Y, 0),
            'U_Vert_Starter',
            create_only_for_one=True,
            has_leader=False)

    create_text_note(win_elevation, b'\xd7\x97\xd7\x95\xd7\xa5'.decode('UTF-8'), window_origin + XYZ(
        -4 * perpendicular_vector.X,
        -4 * perpendicular_vector.Y,
        win_height / 2))
    create_text_note(win_elevation, b'\xd7\xa4\xd7\xa0\xd7\x99\xd7\x9d'.decode('UTF-8'), window_origin + XYZ(
        4 * perpendicular_vector.X,
        4 * perpendicular_vector.Y,
        win_height / 2))

    categories = doc.Settings.Categories
    floor_category = None
    for c in categories:
        if c.Name == 'Floors':
            floor_category = c
    if floor_category is None:
        print(':cross_mark: ERROR! There is no necessary floor category "Flooes". Plugin will stop now...')
        raise Exception('There is no floor\'s category named "Floors"')

    create_spot_elevation(win_elevation, host_object, window_origin)
    create_spot_elevation(win_elevation, host_object, window_origin + XYZ(0, 0, win_height))
    try:
        floors = geographical_finding_algorythm(window_origin, window_origin + XYZ(0.1, 0.1, top_offset + win_height),
                                                object_to_find_categoty=floor_category)
        top_floor = floors[0]
        top_win = geographical_finding_algorythm(window_origin, window_origin + XYZ(0.1, 0.1, top_offset + win_height),
                                                 object_to_find_name=windowFamilyObject.Name)
        if len(top_win):
            create_spot_elevation(win_elevation, host_object, XYZ(window_origin.X, window_origin.Y, top_win[0].get_BoundingBox(None).Min.Z + 1.5 / 30.48))
        create_spot_elevation(win_elevation, host_object,
                              XYZ(window_origin.X, window_origin.Y, top_floor.get_BoundingBox(None).Max.Z))
        create_detail_component(win_elevation, XYZ(
            window_origin.X + (interior_offset + wallDepth / 2) * perpendicular_vector.X,
            window_origin.Y + (interior_offset + wallDepth / 2) * perpendicular_vector.Y,
            top_floor.get_BoundingBox(None).Max.Z - (
                    top_floor.get_BoundingBox(None).Max.Z - top_floor.get_BoundingBox(None).Min.Z) / 4 * 3
        ))
        top_wall = geographical_finding_algorythm(window_origin, window_origin + XYZ(0, 0, top_offset + win_height),
                                                  object_to_find_categoty=host_object.Category,
                                                  ignore_id=host_object.Id)
        if len(top_wall) and not len(top_win):
            create_detail_component(win_elevation, XYZ(
                window_origin.X,
                window_origin.Y,
                window_origin.Z + win_height + top_offset - 30 / 30.48
            ))
    except:
        print('-- Can\'t find top floor. Spot elevation will not be created(Section 1)')

    try:
        floors = geographical_finding_algorythm(window_origin, window_origin + XYZ(0.1, 0.1, -bottom_offset),
                                                object_to_find_categoty=floor_category)
        bottom_floor = floors[0]
        create_spot_elevation(win_elevation, host_object,
                              XYZ(window_origin.X, window_origin.Y, bottom_floor.get_BoundingBox(None).Max.Z))
        create_detail_component(win_elevation, XYZ(
            window_origin.X + (interior_offset + wallDepth / 2) * perpendicular_vector.X,
            window_origin.Y + (interior_offset + wallDepth / 2) * perpendicular_vector.Y,
            bottom_floor.get_BoundingBox(None).Max.Z - (
                    top_floor.get_BoundingBox(None).Max.Z - top_floor.get_BoundingBox(None).Min.Z) / 2
        ))
        bottom_wall = geographical_finding_algorythm(window_origin, window_origin - XYZ(0, 0, bottom_offset),
                                                     object_to_find_categoty=host_object.Category,
                                                     ignore_id=host_object.Id)
        if len(bottom_wall):
            create_detail_component(win_elevation, XYZ(
                window_origin.X,
                window_origin.Y,
                window_origin.Z - bottom_offset
            ))
    except:
        print('-- Can\'t find bottom floor. Spot elevation will not be created(Section 1)')

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


    transform = Transform.Identity
    transform.Origin = window_origin
    vector_perp = get_wall_direction_vector(host_object).CrossProduct(XYZ.BasisZ)
    transform.BasisX = vector_perp
    transform.BasisY = XYZ.BasisZ
    transform.BasisZ = vector_perp.CrossProduct(XYZ.BasisZ)
    section_box = BoundingBoxXYZ()
    offset_50cm = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Centimeters)
    offset_70cm = UnitUtils.ConvertToInternalUnits(70, UnitTypeId.Centimeters)
    if (windowFamilyObject.FacingOrientation.X == -get_wall_direction_vector(host_object).Y and get_wall_direction_vector(host_object).Y != 0) or (
            windowFamilyObject.FacingOrientation.Y == get_wall_direction_vector(host_object).X and get_wall_direction_vector(host_object).X != 0):
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
    win_elevation.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(0)
    win_elevation.ApplyViewTemplateParameters(viewTemplates[0])
    new_name = 'MAMAD_Window_Section_2'
    win_elevation.Scale = 25

    rebars_sets = []
    rebars_sets.append(geographical_finding_algorythm(
        XYZ(window_bbox.Max.X - 0.1, window_bbox.Max.Y - 0.1, window_origin.Z),
        XYZ(window_bbox.Min.X + 0.1 , window_bbox.Min.Y + 0.1, window_origin.Z + 10),
        object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    ))
    rebars_sets.append(geographical_finding_algorythm(
        XYZ(window_bbox.Max.X - 0.1, window_bbox.Max.Y - 0.1, window_origin.Z),
        XYZ(window_bbox.Min.X + 0.1, window_bbox.Min.Y + 0.1, window_origin.Z - 5.6),
        object_to_find_builtin_category=BuiltInCategory.OST_Rebar
    ))
    for i, all_rebars in enumerate(rebars_sets):
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.9 * perpendicular_vector.X, 1.9 * perpendicular_vector.Y, 0.8),
            'Hor_T_Corner',
            leader_end_condition=LeaderEndCondition.Free)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.5 * perpendicular_vector.X, 1.5 * perpendicular_vector.Y, 0.8),
            'Hor_B_Corner',
            leader_end_condition=LeaderEndCondition.Free)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.4 * perpendicular_vector.X, 2.4 * perpendicular_vector.Y, 1.5),
            'Hor_Out',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.29 * perpendicular_vector.X, 1.29 * perpendicular_vector.Y, 1),
            'Hor_In',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.65 * perpendicular_vector.X, 1.3 * perpendicular_vector.Y, -0.1),
            'WD_Hor_14_down',
            leader_end_condition=LeaderEndCondition.Free)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.9 * perpendicular_vector.X, 1.9 * perpendicular_vector.Y, 0),
            'U_Hor',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(1.65 * perpendicular_vector.X, 1.65 * perpendicular_vector.Y, 0.1),
            'WD_Hor_14_up',
            leader_end_condition=LeaderEndCondition.Free
        )

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.29 * perpendicular_vector.X, 2.29 * perpendicular_vector.Y, -1.4),
            'Hor_In_Above',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(2.29 * perpendicular_vector.X, 2.29 * perpendicular_vector.Y, 1.4),
            'Hor_In_Below',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(3.4 * perpendicular_vector.X, 3.4 * perpendicular_vector.Y, 0.8),
            'Hor_Out_Below',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)

        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            'Horizontal_Bars',
            XYZ(3.4 * perpendicular_vector.X, 3.4 * perpendicular_vector.Y, -0.8),
            'Hor_Out_Above',
            leader_end_condition=LeaderEndCondition.Free,
            create_only_for_one=True)


        # create_rebar_tag_depending_on_rebar(
        #     win_elevation,
        #     all_rebars,
        #     TagMode.TM_ADDBY_CATEGORY,
        #     TagOrientation.Horizontal,
        #     'Horizontal_Bars',
        #     XYZ(-2.4 * perpendicular_vector.X, -2.4 * perpendicular_vector.Y, 0.5),
        #     'Hor_Out',
        #     leader_end_condition=LeaderEndCondition.Free,
        #     create_only_for_one=True)

        # create_rebar_tag_depending_on_rebar(
        #     win_elevation,
        #     all_rebars,
        #     TagMode.TM_ADDBY_CATEGORY,
        #     TagOrientation.Horizontal,
        #     'Horizontal_Bars',
        #     XYZ(-1.29 * perpendicular_vector.X, -1.29 * perpendicular_vector.Y, 0),
        #     'Hor_In',
        #     leader_end_condition=LeaderEndCondition.Free,
        #     create_only_for_one=True)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-2.4 * perpendicular_vector.X, -2.4 * perpendicular_vector.Y, 0),
            'WD_Vert_14_Out'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-2.2 * perpendicular_vector.X, -2.2 * perpendicular_vector.Y, 0),
            'WD_Vert_14_Out',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-3.4 * perpendicular_vector.X, -3.4 * perpendicular_vector.Y, 0),
            'U_Above_Out'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-3.7 * perpendicular_vector.X, -3.7 * perpendicular_vector.Y, 0),
            'U_Above_Out',
            create_only_for_one=True,
            has_leader=False)

        # create_bending_detail(
        #     win_elevation,
        #     all_rebars,
        #     'Bending Detail 2 (No hooks)',
        #     XYZ(-3.4 * perpendicular_vector.X, -3.4 * perpendicular_vector.Y, 0),
        #     'Bending_5'
        # )
        # create_rebar_tag_depending_on_rebar(
        #     win_elevation,
        #     all_rebars,
        #     TagMode.TM_ADDBY_CATEGORY,
        #     TagOrientation.Vertical,
        #     'Link&U-Shape+Length',
        #     XYZ(-4.4 * perpendicular_vector.X, -4.4 * perpendicular_vector.Y, 0),
        #     'Bending_5',
        #     create_only_for_one=True,
        #     has_leader=False)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-3.4 * perpendicular_vector.X, -3.4 * perpendicular_vector.Y, 0),
            'L_Out'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-4.15 * perpendicular_vector.X, -4.15 * perpendicular_vector.Y, 0),
            'L_Out',
            create_only_for_one=True,
            has_leader=False)

        create_bending_detail(
            win_elevation,
            all_rebars,
            'Bending Detail 2 (No hooks)',
            XYZ(-2.5 * perpendicular_vector.X, -2.5 * perpendicular_vector.Y, 0),
            'U_Vert_Starter'
        )
        create_rebar_tag_depending_on_rebar(
            win_elevation,
            all_rebars,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Vertical,
            'Link&U-Shape+Length',
            XYZ(-2.2 * perpendicular_vector.X, -2.2 * perpendicular_vector.Y, 0),
            'U_Vert_Starter',
            create_only_for_one=True,
            has_leader=False)

        if not i:
            create_bending_detail(
                win_elevation,
                all_rebars,
                'Bending Detail 2 (No hooks)',
                XYZ(-1.9 * perpendicular_vector.X, -1.9 * perpendicular_vector.Y, 0),
                'WD_Vert_14_In'
            )
            create_rebar_tag_depending_on_rebar(
                win_elevation,
                all_rebars,
                TagMode.TM_ADDBY_CATEGORY,
                TagOrientation.Vertical,
                'Link&U-Shape+Length',
                XYZ(-2.1 * perpendicular_vector.X, -2.1 * perpendicular_vector.Y, 0),
                'WD_Vert_14_In',
                create_only_for_one=True,
                has_leader=False)

    create_text_note(win_elevation, b'\xd7\x97\xd7\x95\xd7\xa5'.decode('UTF-8'), window_origin + XYZ(
        -4 * perpendicular_vector.X,
        -4 * perpendicular_vector.Y,
        win_height / 2))
    create_text_note(win_elevation, b'\xd7\xa4\xd7\xa0\xd7\x99\xd7\x9d'.decode('UTF-8'), window_origin + XYZ(
        4 * perpendicular_vector.X,
        4 * perpendicular_vector.Y,
        win_height / 2))

    categories = doc.Settings.Categories
    floor_category = None
    for c in categories:
        if c.Name == 'Floors':
            floor_category = c
    if floor_category is None:
        print(':cross_mark: ERROR! There is no necessary floor category "Flooes". Plugin will stop now...')
        raise Exception('There is no floor\'s category named "Floors"')

    create_spot_elevation(win_elevation, host_object, window_origin + XYZ(0, 0, win_height))
    try:
        floors = geographical_finding_algorythm(window_origin, window_origin + XYZ(0, 0, top_offset + win_height),
                                                object_to_find_categoty=floor_category)
        top_floor = floors[0]
        create_spot_elevation(win_elevation, host_object,
                              XYZ(window_origin.X, window_origin.Y, top_floor.get_BoundingBox(None).Max.Z))
        create_detail_component(win_elevation, XYZ(
            window_origin.X + (interior_offset + wallDepth / 2) * perpendicular_vector.X,
            window_origin.Y + (interior_offset + wallDepth / 2) * perpendicular_vector.Y,
            top_floor.get_BoundingBox(None).Max.Z - (
                    top_floor.get_BoundingBox(None).Max.Z - top_floor.get_BoundingBox(None).Min.Z) / 4 * 3
        ))
    except:
        print('-- Can\'t find top floor. Spot elevation will not be created(Section 2)')

    try:
        floors = geographical_finding_algorythm(window_origin, window_origin + XYZ(0, 0, -bottom_offset),
                                                object_to_find_categoty=floor_category)
        bottom_floor = floors[0]
        create_spot_elevation(win_elevation, host_object,
                              XYZ(window_origin.X, window_origin.Y, bottom_floor.get_BoundingBox(None).Max.Z))
        create_detail_component(win_elevation, XYZ(
            window_origin.X + (interior_offset + wallDepth / 2) * perpendicular_vector.X,
            window_origin.Y + (interior_offset + wallDepth / 2) * perpendicular_vector.Y,
            bottom_floor.get_BoundingBox(None).Max.Z - (
                    top_floor.get_BoundingBox(None).Max.Z - top_floor.get_BoundingBox(None).Min.Z) / 2
        ))
    except:
        print('-- Can\'t find top floor. Spot elevation will not be created(Section 2)')

    top_wall = geographical_finding_algorythm(window_origin, window_origin + XYZ(0, 0, top_offset + win_height),
                                              object_to_find_categoty=host_object.Category, ignore_id=host_object.Id)
    if len(top_wall):
        create_detail_component(win_elevation, XYZ(
            window_origin.X,
            window_origin.Y,
            window_origin.Z + win_height + top_offset - 10 / 30.48
        ))

    bottom_wall = geographical_finding_algorythm(window_origin, window_origin - XYZ(0, 0, bottom_offset),
                                                 object_to_find_categoty=host_object.Category, ignore_id=host_object.Id)
    if len(bottom_wall):
        create_detail_component(win_elevation, XYZ(
            window_origin.X,
            window_origin.Y,
            window_origin.Z - bottom_offset
        ))

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
    get_front_view()
    get_perpendicular_window_section()
    get_perpendicular_shelter_section()
    get_callout()
except Exception as err:
    print('Error happened! Contact developer to fix it.', err)
finally:
    transaction.Commit()

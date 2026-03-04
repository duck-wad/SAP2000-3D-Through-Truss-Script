import comtypes.client
import math


def sap_open():
    # create API helper object
    helper = comtypes.client.CreateObject("SAP2000v1.Helper")
    helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
    sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
    if sap_object is None:
        sap_object = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
        sap_object.ApplicationStart()

    return sap_object


def sap_close(sap_object):
    ret = sap_object.ApplicationExit(False)
    sap_object = None


def sap_initialize_model(base_file_path, sap_object):

    sap_model = sap_object.SapModel
    ret = sap_model.File.OpenFile(base_file_path)
    return sap_model


def sap_set_sections(
    sap_model,
    bottom_chord_name,
    top_chord_name,
    diag_web_name,
    vert_web_name,
    lateral_name,
    bottom_chord_sec,
    top_chord_sec,
    diag_web_sec,
    vert_web_sec,
    lateral_sec,
):

    sap_model.FrameObj.SetSection(
        Name=bottom_chord_name, PropName=bottom_chord_sec, ItemType=1
    )
    sap_model.FrameObj.SetSection(
        Name=top_chord_name, PropName=top_chord_sec, ItemType=1
    )
    sap_model.FrameObj.SetSection(Name=diag_web_name, PropName=diag_web_sec, ItemType=1)
    sap_model.FrameObj.SetSection(Name=vert_web_name, PropName=vert_web_sec, ItemType=1)
    sap_model.FrameObj.SetSection(Name=lateral_name, PropName=lateral_sec, ItemType=1)

    sap_model.DesignSteel.SetDesignSection(
        Name=bottom_chord_name,
        PropName=bottom_chord_sec,
        LastAnalysis=False,
        ItemType=1,
    )
    sap_model.DesignSteel.SetDesignSection(
        Name=top_chord_name, PropName=top_chord_sec, LastAnalysis=False, ItemType=1
    )
    sap_model.DesignSteel.SetDesignSection(
        Name=diag_web_name, PropName=diag_web_sec, LastAnalysis=False, ItemType=1
    )
    sap_model.DesignSteel.SetDesignSection(
        Name=vert_web_name, PropName=vert_web_sec, LastAnalysis=False, ItemType=1
    )
    sap_model.DesignSteel.SetDesignSection(
        Name=lateral_name, PropName=lateral_sec, LastAnalysis=False, ItemType=1
    )

    names = [
        bottom_chord_name,
        top_chord_name,
        diag_web_name,
        vert_web_name,
        lateral_name,
    ]
    frame_ids = []

    for name in names:
        _, types, ids, ret = sap_model.GroupDef.GetAssignments(name)
        types = list(types)
        ids = list(ids)
        # only need to the ids corresponding with a type of 2 (2=frame)
        ids = [id for t, id in zip(types, ids) if t == 2]
        frame_ids.append(ids)

    return frame_ids


def sap_central_node(sap_model, bottom_chord_frames):
    # just picked this manually from model
    return "9"


def sap_run_analysis(sap_model, file_path):
    ret = sap_model.File.Save(file_path)
    ret = sap_model.Analyze.RunAnalysis()


def sap_deflection(sap_model, bottom_chord_frames, span_length):

    deflection_limit = span_length / 360.0

    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput("SLS 1")

    # get the central node vertical displacement
    center_node = sap_central_node(sap_model, bottom_chord_frames)
    _, _, _, _, _, _, _, _, temp, _, _, _, ret = sap_model.Results.JointDispl(
        center_node, 0, 0, [], [], [], [], [], [], [], [], [], [], []
    )
    # return absolute value
    deflection = (
        abs(temp[0]) * 0.0254
    )  # convert from inches to m idk why sap model is in inches and dont know how to change it
    percentage = deflection / deflection_limit * 100
    return deflection, percentage


def sap_module_mass(sap_model, num_modules):

    # get the results from the 'DEAD' load case
    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput("DEAD")

    _, _, _, _, _, _, reaction, _, _, _, _, _, _, ret = sap_model.Results.BaseReact(
        0, [], [], [], [], [], [], [], [], [], 0, 0, 0
    )
    print(reaction[0])
    # convert kip to kN then kN to kg
    total_mass = reaction[0] * 4.4482216
    total_mass = total_mass / 9.81 * 1000
    module_mass = total_mass / num_modules

    return module_mass


def sap_member_design(sap_model):

    cases = [
        "ULS 1_s",
        "ULS 3_ecc-1_s",
        "ULS 3_ecc-2_s",
        "ULS 3_ecc+1_s",
        "ULS 3_ecc+2_s",
        "ULS 3-_s",
        "ULS 3+_s",
        "ULS 4_ecc-1_s",
        "ULS 4_ecc-2_s",
        "ULS 4_ecc+1_s",
        "ULS 4_ecc+2_s",
        "ULS 4-_s",
        "ULS 4+_s",
        "ULS 5_s",
        "ULS 6_s",
    ]

    for case in cases:

        ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret = sap_model.Results.Setup.SetCaseSelectedForOutput(case)

        num_failed = 0
        names = []

        ret = sap_model.DesignSteel.StartDesign()
        _, num_failed, _, names, ret = sap_model.DesignSteel.VerifyPassed(
            0, num_failed, 0, names
        )

        if num_failed != 0:
            passed = False
        else:
            passed = True

    return passed

import comtypes.client
import math
import sys


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

    names = [
        bottom_chord_name,
        top_chord_name,
        diag_web_name,
        vert_web_name,
        lateral_name,
    ]
    sections = [
        bottom_chord_sec,
        top_chord_sec,
        diag_web_sec,
        vert_web_sec,
        lateral_sec,
    ]

    for i in range(len(names)):
        sap_model.FrameObj.SetSection(Name=names[i], PropName=sections[i], ItemType=1)
        sap_model.DesignSteel.SetDesignSection(
            Name=names[i],
            PropName=sections[i],
            LastAnalysis=False,
            ItemType=1,
        )


def sap_central_node():
    # just picked this manually from model
    return "9"


def sap_run_analysis(sap_model, file_path):
    ret = sap_model.File.Save(file_path)
    ret = sap_model.Analyze.RunAnalysis()


def sap_deflection(sap_model, span_length):

    deflection_limit = span_length / 360.0

    ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = sap_model.Results.Setup.SetCaseSelectedForOutput("SLS 1")

    # get the central node vertical displacement
    center_node = sap_central_node()
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

    passed = True
    failed_cases = []

    for case in cases:

        ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret = sap_model.Results.Setup.SetCaseSelectedForOutput(case)
        num_failed = 0
        names = []

        ret = sap_model.DesignSteel.StartDesign()
        _, num_failed, _, names, ret = sap_model.DesignSteel.VerifyPassed(
            0, num_failed, 0, names
        )

        if num_failed > 0:
            passed = False
            failed_cases.append(case)

    return passed, failed_cases


def sap_vibration_analysis(sap_model):
    # get the vertical and lateral accelerations from model
    # check if they are in comfort class 2 (below 1m/s2 for vertical, 0.3m/s2 for lateral)

    v_cases = ["P_1V_t", "P_2V_t", "P_3V_t", "P_4V_t", "P_5V_t"]
    l_cases = ["P_1L_t", "P_2L_t", "P_3L_t", "P_4L_t", "P_5L_t"]
    modes = [1, 2, 3, 4, 5]
    v_accel = []
    l_accel = []

    is_class2 = True

    # check the vertical acceleration for modes 1-5
    for case in v_cases:
        ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret = sap_model.Results.Setup.SetCaseSelectedForOutput(case)

        _, _, _, _, _, _, _, _, U3, _, _, _, ret = sap_model.Results.JointAccAbs(
            sap_central_node(), 0, 0, [], [], [], [], [], [], [], [], [], [], []
        )
        U3_max = U3[0]
        # convert in/s2 to m/s2
        U3_max = U3_max * 0.0254
        v_accel.append(U3_max)

        if U3_max > 1.0:
            is_class2 = False

    # now check the lateral acceleration for modes 1-5
    for case in l_cases:
        ret = sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret = sap_model.Results.Setup.SetCaseSelectedForOutput(case)

        _, _, _, _, _, _, _, U2, _, _, _, _, ret = sap_model.Results.JointAccAbs(
            sap_central_node(), 0, 0, [], [], [], [], [], [], [], [], [], [], []
        )
        U2_max = U2[0]
        # convert in/s2 to m/s2
        U2_max = U2_max * 0.0254
        l_accel.append(U2_max)

        if U2_max > 0.3:
            is_class2 = False

    return is_class2, v_accel, l_accel

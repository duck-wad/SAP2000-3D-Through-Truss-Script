import os
from tqdm import tqdm
import sys

from sap_interface import *
from define_geometry import *
from define_sections import *

if __name__ == "__main__":

    """------------------------ DEFINE MODEL PARAMETERS ------------------------"""

    height = 2.5
    module_length = 15.0
    module_divisions = 3  # applies to the bottom chord
    segment_length = module_length / module_divisions
    # each span contains 2 modules, so there will be num_spans x 2 modules
    # num_spans should be an odd number
    # ONLY CONSIDERING 1 SPAN FOR 3D
    num_spans = 3
    assert num_spans % 2 != 0
    num_modules = 2 * num_spans
    span_length = 2 * module_length
    total_length = span_length * num_spans

    # default analysis is for steel. for aluminum set is_alu to true
    is_alu = False

    """ ------------------------ DEFINE LOADS ------------------------ """

    deck_width = 3.0  # meters
    deck_pressure = 1.85  # kPa (Simon calculation)

    pedestrian_pressure = 3.86  # kPa (Simon calculation)
    pedestrian_density = 1.5  # p/m2

    trib_area = deck_width / 2.0

    # compute UDLs
    live_UDL = trib_area * pedestrian_pressure
    concrete_deck_UDL = deck_pressure * trib_area

    """ ------------------------ INITIALIZE MODEL ------------------------ """

    # import the section combinations
    # will be a list of lists, each sublist has the combinations for different section types
    # ex. round, box, round-box combo
    # each combination has 4 sections: top chord, bottom chord, web, lateral
    # but lateral is matched with web for simplicity
    section_combinations = create_section_combinations_steel()

    # set file paths
    root_path = os.getcwd()
    base_file_path = root_path + "/BASE.sdb"
    os.makedirs("./models", exist_ok=True)
    model_path = root_path + "/models/MODEL.sdb"

    # open SAP application
    sap_object = sap_open()

    # delete old output file if it exists
    results_file = "output.xlsx"
    results_path = root_path + os.sep + results_file
    if os.path.exists(results_path):
        os.remove(results_path)
    sheet_names = ["Box Box Box", "Box Box Round"]

    first_write = True

    # initialize the model from BASE in root folder
    # this model will stay open until end of process
    # just changing sections in this model and extracting results

    for index, combination_type in enumerate(section_combinations):
        results = []
        sheet_name = sheet_names[index]

        for combo_index, combination in enumerate(tqdm(combination_type)):

            sap_model = sap_initialize_model(base_file_path, sap_object)

            top_chord_section = combination[0]
            bottom_chord_section = combination[1]
            web_section = combination[2]
            lateral_section = combination[3]

            """ ------------------------ CREATE SAP MODEL ------------------------ """

            # collect the frame objects from the model based on their assigned group
            # and also change the frame section to be the one in the combination
            (
                bottom_chord_frames,
                top_chord_frames,
                diagonal_web_frames,
                vertical_web_frames,
                lateral_frames,
            ) = sap_set_sections(
                sap_model,
                "BOTTOM_CHORD",
                "TOP_CHORD",
                "DIAG_WEB",
                "VERT_WEB",
                "LATERAL",
                bottom_chord_section,
                top_chord_section,
                web_section,
                web_section,
                lateral_section,
            )

            """ ------------------------ RUN MODEL AND COLLECT RESULTS ------------------------ """
            sap_run_analysis(sap_model, model_path)

            deflection, deflection_percentage = sap_deflection(
                sap_model, bottom_chord_frames, span_length
            )

            # get the reaction output from dead case and divide by num_modules
            module_mass = sap_module_mass(sap_model, num_modules)

            # verify frames pass steel design check, and get list of sections that fail if ULS does not pass
            passed = sap_member_design(sap_model)

            results.append(
                {
                    "Top chord": top_chord_section,
                    "Bottom chord": bottom_chord_section,
                    "Web members": web_section,
                    "Max vertical deflection for SLS (m)": deflection,
                    "Percentage of deflection limit for SLS (%)": deflection_percentage,
                    "Module mass (kg)": module_mass,
                    "Passed member design check for ULS": passed,
                }
            )

            # log results to console
            tqdm.write(
                f"Top chord section: {top_chord_section}, Bottom chord section: {bottom_chord_section}, Web member section: {web_section}, Lateral member section: {lateral_section}"
            )
            tqdm.write(
                f"Deflection of central node for SLS (mm): {round(deflection * 1000, 4)}"
            )
            tqdm.write(
                f"Percentage of deflection limit for SLS (%): {round(deflection_percentage)}"
            )
            tqdm.write(f"Mass of single module (kg): {round(module_mass, 4)}")
            tqdm.write(f"Passed member design check for ULS: {passed}")

            # write result to excel
            if combo_index % 10 == 0:
                write_to_excel(results, results_path, sheet_name, first_write)
                tqdm.write("Successfully updated output file.")
                first_write = False

            sap_model = None

    sap_close(sap_object)

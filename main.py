import os
from tqdm import tqdm
import sys

from sap_interface import *
from define_sections import *

if __name__ == "__main__":

    """------------------------ DEFINE MODEL PARAMETERS ------------------------"""

    num_spans = 3
    assert num_spans % 2 != 0
    num_modules = 2 * num_spans
    span_length = 30.0

    """ ------------------------ INITIALIZE MODEL ------------------------ """

    # import the section combinations
    # will be a list of lists, each sublist has the combinations for different section types
    # ex. round, box, round-box combo
    # each combination has 4 sections: top chord, bottom chord, web, lateral
    # but lateral is matched with web for simplicity
    section_combinations = create_section_combinations_steel()
    print(section_combinations[0][0])

    # set file paths. put run model in models folder to
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
    # sheet_names = ["Box Box Box", "Box Box Round"]
    sheet_names = ["Box Box Box"]

    first_write = True

    for index, combination_type in enumerate(section_combinations):
        results = []
        sheet_name = sheet_names[index]

        for combo_index, combination in enumerate(tqdm(combination_type)):

            sap_model = sap_initialize_model(base_file_path, sap_object)

            """ ------------------------ SET MEMBER SECTIONS ------------------------ """

            top_chord_section = combination[0]
            bottom_chord_section = combination[1]
            diag_web_section = combination[2]
            vert_web_section = combination[3]
            lateral_section = combination[4]

            print(combination)

            # change the section of the member groups
            sap_set_sections(
                sap_model,
                "BOTTOM_CHORD",
                "TOP_CHORD",
                "DIAG_WEB",
                "VERT_WEB",
                "LATERAL",
                bottom_chord_section,
                top_chord_section,
                diag_web_section,
                vert_web_section,
                lateral_section,
            )

            """ ------------------------ RUN MODEL AND COLLECT RESULTS ------------------------ """
            sap_run_analysis(sap_model, model_path)

            deflection, deflection_percentage = sap_vert_deflection(
                sap_model, span_length
            )

            lat_deflection, lat_deflection_percentage = sap_lat_deflection(sap_model)

            # get the reaction output from dead case and divide by num_modules
            module_mass = sap_module_mass(sap_model, num_modules)

            # verify frames pass steel design check, and get list of sections that fail if ULS does not pass
            passed, failed_cases, num_failed = sap_member_design(sap_model)

            is_class2, v_accel, l_accel = sap_vibration_analysis(sap_model)

            results.append(
                {
                    "Top chord": top_chord_section,
                    "Bottom chord": bottom_chord_section,
                    "Diagonal Web members": diag_web_section,
                    "Vertical Web members": vert_web_section,
                    "Laterals": lateral_section,
                    "Max vertical deflection for SLS (m)": deflection,
                    "Percentage of deflection limit (L/360) for SLS (%)": deflection_percentage,
                    "Max lateral deflection for SLS (m)": lat_deflection,
                    "Percentage of deflection limit (100mm) for SLS (%)": lat_deflection_percentage,
                    "Module mass (kg)": module_mass,
                    "Passed member design check for ULS": passed,
                    "Failed ULS cases": failed_cases,
                    "Number of failed members": num_failed,
                    "Class 2": is_class2,
                    "Vertical acceleration 1 (m/s2)": v_accel[0],
                    "Vertical acceleration 2 (m/s2)": v_accel[1],
                    "Vertical acceleration 3 (m/s2)": v_accel[2],
                    "Vertical acceleration 4 (m/s2)": v_accel[3],
                    "Vertical acceleration 5 (m/s2)": v_accel[4],
                    "Lateral acceleration 1 (m/s2)": l_accel[0],
                    "Lateral acceleration 2 (m/s2)": l_accel[1],
                    "Lateral acceleration 3 (m/s2)": l_accel[2],
                    "Lateral acceleration 4 (m/s2)": l_accel[3],
                    "Lateral acceleration 5 (m/s2)": l_accel[4],
                }
            )

            # log results to console
            tqdm.write(
                f"Top chord section: {top_chord_section}, Bottom chord section: {bottom_chord_section}, Diagonal web member section: {diag_web_section}, Vertical web member section: {vert_web_section} Lateral member section: {lateral_section}"
            )
            tqdm.write(
                f"Vertical deflection of central node for SLS (mm): {round(deflection * 1000, 3)}"
            )
            tqdm.write(
                f"Percentage of vertical deflection limit for SLS (%): {round(deflection_percentage)}"
            )
            tqdm.write(
                f"Lateral deflection of central node for SLS (mm): {round(lat_deflection * 1000, 3)}"
            )
            tqdm.write(
                f"Percentage of lateral deflection limit for SLS (%): {round(lat_deflection_percentage)}"
            )
            tqdm.write(f"Mass of single module (kg): {round(module_mass, 3)}")
            tqdm.write(f"Passed member design check for ULS: {passed}")
            tqdm.write(f"Failed ULS cases: {failed_cases}")
            tqdm.write(f"Number of failed members: {num_failed}")
            tqdm.write(f"Is comfort class 2?: {is_class2}")
            tqdm.write(
                f"Acceleration in mode 1 (m/s2): vertical {v_accel[0]}, lateral {l_accel[0]}"
            )
            tqdm.write(
                f"Acceleration in mode 2 (m/s2): vertical {v_accel[1]}, lateral {l_accel[1]}"
            )
            tqdm.write(
                f"Acceleration in mode 3 (m/s2): vertical {v_accel[2]}, lateral {l_accel[2]}"
            )
            tqdm.write(
                f"Acceleration in mode 4 (m/s2): vertical {v_accel[3]}, lateral {l_accel[3]}"
            )
            tqdm.write(
                f"Acceleration in mode 5 (m/s2): vertical {v_accel[4]}, lateral {l_accel[4]}"
            )

            # write result to excel
            if combo_index % 10 == 0:
                write_to_excel(results, results_path, sheet_name, first_write)
                tqdm.write("Successfully updated output file.")
                first_write = False

            sap_model = None

    sap_close(sap_object)

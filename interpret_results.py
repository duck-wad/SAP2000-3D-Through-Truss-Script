import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import sys


def plot_save(plt, out_path, section, name):

    section_formatted = (section.replace(" ", "")).lower()
    name_formatted = name.replace(" ", "")
    full_name = "_".join([section_formatted, name_formatted])
    full_name = full_name + ".pdf"
    full_path = out_path + os.sep + full_name
    plt.savefig(full_path, format="pdf")


def get_uls_indices(df):
    uls_result = df["Passed member design check for ULS"].to_numpy()
    index_true = np.where(uls_result)[0]
    index_false = np.where(~uls_result)[0]
    return index_true, index_false


def determine_optimal_section(mass, comfort):

    valid_indices = np.where(comfort)[0]  # indices where comfort == True
    min_index = valid_indices[np.argmin(mass[valid_indices])]

    return min_index


def plot_mass_vs_deflection(
    mass_uls_passed,
    mass_uls_failed,
    deflection_uls_passed,
    deflection_uls_failed,
    optimal_mass,
    optimal_deflection,
    optimal_sections,
    sheet_name,
    out_path,
):

    plt.figure(figsize=(8, 6))
    plt.scatter(
        mass_uls_failed,
        deflection_uls_failed * 1000,
        label="ULS Failed",
        color="red",
        s=10,
    )
    plt.scatter(
        mass_uls_passed,
        deflection_uls_passed * 1000,
        label="ULS Passed",
        color="green",
        s=10,
    )
    plt.scatter(
        optimal_mass,
        optimal_deflection * 1000,
        label="Optimal Section Combination",
        color="cyan",
        s=20,
    )
    plt.ylabel("SLS Deflection [mm]")
    plt.xlabel("Mass of module [kg]")
    plt.grid(True)
    plt.minorticks_on()
    plt.legend()
    plt.annotate(
        optimal_sections,
        xy=(optimal_mass, optimal_deflection * 1000),
        xytext=(80, 80),  # offset in pixels
        textcoords="offset points",
        arrowprops=dict(arrowstyle="->", color="cyan", lw=1.5),
        fontsize=9,
        bbox=dict(
            boxstyle="round,pad=0.3", edgecolor="cyan", facecolor="white", alpha=0.9
        ),
    )

    plot_save(plt, out_path, sheet_name, "mass vs deflection")


def plot_mass_vs_acceleration(
    mass,
    acceleration_v,
    acceleration_l,
    sheet_name,
    out_path,
):

    mode_colors = ["blue", "orange", "green", "red", "purple"]

    # -------- Vertical acceleration plot --------
    plt.figure(figsize=(8, 6))

    for i in range(acceleration_v.shape[1]):  # 5 modes
        plt.scatter(
            mass, acceleration_v[:, i], label=f"Mode {i+1}", color=mode_colors[i], s=10
        )

    plt.ylabel("Vertical acceleration [m/s²]")
    plt.xlabel("Mass of module [kg]")
    plt.title("Mass vs Vertical Acceleration")
    plt.grid(True)
    plt.minorticks_on()
    plt.legend()

    plot_save(plt, out_path, sheet_name, "mass vs vertical acceleration")

    # -------- Lateral acceleration plot --------
    plt.figure(figsize=(8, 6))

    for i in range(acceleration_l.shape[1]):  # 5 modes
        plt.scatter(
            mass, acceleration_l[:, i], label=f"Mode {i+1}", color=mode_colors[i], s=10
        )

    plt.ylabel("Lateral acceleration [m/s²]")
    plt.xlabel("Mass of module [kg]")
    plt.title("Mass vs Lateral Acceleration")
    plt.grid(True)
    plt.minorticks_on()
    plt.legend()

    plot_save(plt, out_path, sheet_name, "mass vs lateral acceleration")


# interpret_results can be run from main.py, or just from the run() function in this file
def interpret_results(file_path, sheets, folderpath):

    os.makedirs("./plots", exist_ok=True)
    out_path = folderpath + "/plots"

    dfs = []
    for sheet in sheets:
        dfs.append(pd.read_excel(file_path, sheet_name=sheet))

    # track the optimal deflections, harmonics, and masses for each section combination type

    for index, df in enumerate(dfs):

        top_chord = df["Top chord"].to_numpy()
        bottom_chord = df["Bottom chord"].to_numpy()
        diag_web = df["Diagonal Web members"].to_numpy()
        vert_web = df["Vertical Web members"].to_numpy()
        lateral = df["Laterals"].to_numpy()
        mass = df["Module mass (kg)"].to_numpy()
        comfort_class = df["Class 2"].to_numpy()
        deflection = df["Max vertical deflection for SLS (m)"].to_numpy()
        acceleration_v = np.column_stack(
            [
                df["Vertical acceleration 1 (m/s2)"],
                df["Vertical acceleration 2 (m/s2)"],
                df["Vertical acceleration 3 (m/s2)"],
                df["Vertical acceleration 4 (m/s2)"],
                df["Vertical acceleration 5 (m/s2)"],
            ]
        )

        acceleration_l = np.column_stack(
            [
                df["Lateral acceleration 1 (m/s2)"],
                df["Lateral acceleration 2 (m/s2)"],
                df["Lateral acceleration 3 (m/s2)"],
                df["Lateral acceleration 4 (m/s2)"],
                df["Lateral acceleration 5 (m/s2)"],
            ]
        )

        # get the indices of section combos that pass ULS and split into separate lists
        index_true, index_false = get_uls_indices(df)

        mass_uls_passed = mass[index_true]
        mass_uls_failed = mass[index_false]
        comfort_uls_passed = comfort_class[index_true]
        comfort_uls_failed = comfort_class[index_false]
        deflection_uls_passed = deflection[index_true]
        deflection_uls_failed = deflection[index_false]
        top_chord_uls_passed = top_chord[index_true]
        bottom_chord_uls_passed = bottom_chord[index_true]
        diag_web_uls_passed = diag_web[index_true]
        vert_web_uls_passed = vert_web[index_true]
        lateral_uls_passed = lateral[index_true]

        high_score_index = determine_optimal_section(
            mass_uls_passed, comfort_uls_passed
        )

        optimal_mass = mass_uls_passed[high_score_index]
        optimal_deflection = deflection_uls_passed[high_score_index]
        optimal_top_chord = top_chord_uls_passed[high_score_index]
        optimal_bottom_chord = bottom_chord_uls_passed[high_score_index]
        optimal_diag_web = diag_web_uls_passed[high_score_index]
        optimal_vert_web = vert_web_uls_passed[high_score_index]
        optimal_lateral = lateral_uls_passed[high_score_index]

        # with line breaks for plotting
        optimal_sections_spaced = (
            "Top: "
            + optimal_top_chord
            + "\nBottom: "
            + optimal_bottom_chord
            + "\nDiagonal Web: "
            + optimal_diag_web
            + "\nVertical Web: "
            + optimal_vert_web
            + "\nLateral: "
            + optimal_lateral
        )

        plot_mass_vs_deflection(
            mass_uls_passed,
            mass_uls_failed,
            deflection_uls_passed,
            deflection_uls_failed,
            optimal_mass,
            optimal_deflection,
            optimal_sections_spaced,
            sheets[index],
            out_path,
        )

        plot_mass_vs_acceleration(
            mass,
            acceleration_v,
            acceleration_l,
            sheets[index],
            out_path,
        )


def run():
    root_path = os.getcwd()
    file_path = root_path + "/output.xlsx"
    # sheet_names = ["Aluminum"]
    sheet_names = ["Steel"]
    interpret_results(file_path, sheet_names, root_path)


if __name__ == "__main__":
    run()

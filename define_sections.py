import xml.etree.ElementTree as ET
import pandas as pd
from fractions import Fraction
import sys


def load_xml_steel():
    # load xml file from SAP2000 installation folder
    """ tree = ET.parse(
        r"C:\Program Files\Computers and Structures\SAP2000 26\Property Libraries\Sections\CISC10.xml"
    )
    root = tree.getroot()

    ns = {"csi": "http://www.csiberkeley.com"}

    # find all steel pipe (round)
    round_pipe = root.findall(".//csi:STEEL_PIPE", ns)

    # filter the labels
    hss_round = []
    for pipe in round_pipe:
        label = pipe.find("csi:LABEL", ns)
        if label is not None and label.text.startswith("HS"):
            hss_round.append(label.text)

    # find all HSS sections
    hss = root.findall(".//csi:STEEL_BOX", ns)

    hss_box = []
    for pipe in hss:
        label = pipe.find("csi:LABEL", ns)
        if label is not None and label.text.startswith("HS"):
            hss_box.append(label.text)

    # write sections to excel file for easier reference
    max_len = max(len(hss_round), len(hss_box))
    hss_round_excel = hss_round.copy()
    hss_box_excel = hss_box.copy()
    hss_round_excel += [None] * (max_len - len(hss_round))
    hss_box_excel += [None] * (max_len - len(hss_box))
    """
    output_path = "./steel_sections.xlsx"
    df = pd.read_excel(output_path)
    # Create DataFrame
    hss_round = (df['HSS Round']).dropna().tolist()
    hss_box = df['HSS Box'].dropna().tolist()

    return hss_round, hss_box


def filter_HSS_sections_steel(
    sections, min_depth, min_thick, max_depth, max_thick, asym=False
):

    filtered_sections = []
    for section in sections:
        # name is in format HS###X## (round) and HS###X###X## (box)
        # split section name into before and after 'X'
        # depth is always the first dim and thickness always the last dim
        parts = section.split("X")
        # get diameter as everything after 'HS' and before 'X', and convert to int
        depth = int(parts[0][2:])
        # get numbers after 'X' and convert to float
        thick = float(parts[-1])
        # get width if applicable
        width = depth
        if len(parts) == 3:
            width = int(parts[1])

        if (
            depth >= min_depth
            and thick >= min_thick
            and depth <= max_depth
            and thick <= max_thick
        ):
            if asym == True and width != depth:
                pass
            else:
                filtered_sections.append(section)

    return filtered_sections


def get_depth(section):
    parts = section.replace("HS", "").split("X")
    return int(float(parts[0]))


def get_width(section):
    parts = section.replace("HS", "").split("X")

    # for round hss
    if len(parts) == 2:
        return int(float(parts[0]))

    # rectangular/square HSS
    return int(float(parts[1]))


def valid_combinations_steel(
    top_sections, bottom_sections, web_sections, lateral_sections
):

    combinations = []
    min_d_diff = 50

    for top in top_sections:

        top_d = get_depth(top)
        top_w = get_width(top)

        for bottom in bottom_sections:

            bottom_d = get_depth(bottom)
            bottom_w = get_width(bottom)

            # width top chord = width bottom chord
            if top_w == bottom_w and top_d == bottom_d:

                for diag_web in web_sections:

                    diag_d = get_depth(diag_web)
                    diag_w = get_width(diag_web)

                    # width diag_web <= width bottom chord
                    if diag_w <= bottom_w:

                        for vert_web in web_sections:

                            vert_d = get_depth(vert_web)
                            vert_w = get_width(vert_web)

                            # width vert_web = depth vert_web (square only)
                            if vert_w == vert_d:

                                # width vert_web <= width bottom chord
                                if vert_w <= bottom_w:

                                    for lateral in lateral_sections:

                                        lat_d = get_depth(lateral)
                                        lat_w = get_width(lateral)

                                        # depth lateral <= depth top chord
                                        if lat_d <= top_d:

                                            # width lateral = depth lateral (square only)
                                            if lat_w == lat_d:

                                                combinations.append(
                                                    [
                                                        top,
                                                        bottom,
                                                        diag_web,
                                                        vert_web,
                                                        lateral,
                                                    ]
                                                )

    return combinations


# we are limiting bottom and top chord to be box only for connection purposes
# top chord > bottom chord > web
# for box limit to square section no rectangle (for now to limit combinations)
def create_section_combinations_steel():

    round, box = load_xml_steel()

    """ ------------------------ FILTER ROUND SECTIONS ------------------------ """
    # web is smaller than bottom chord. don't specify a min or max diameter, that will be taken care of in
    # the valid_combinations_steel function. limit thickness 7.9-9.5
    web_round = filter_HSS_sections_steel(round, 152, 7.9, 500, 9.5)

    """ ------------------------ FILTER BOX SECTIONS ------------------------ """

    # sort box sections based on size (depth, width, thickness)
    # in xml they are ordered by square first then rectangle
    box = sorted(
        box,
        key=lambda x: (
            int(x[2 : x.index("X")]),
            int(x[x.index("X") + 1 : x.index("X", x.index("X") + 1)]),
            float(x[x.index("X", x.index("X") + 1) + 1 :]),
        ),
    )
    # reverse list
    box.reverse()

    # limit top chord to be depth 200+ and thickness 7.9-9.5
    top_chord_box = filter_HSS_sections_steel(box, 152, 6.4, 203, 8)
    # limit bottom chord depth bottom limit 152, 305 top, 7.9-9.5
    bottom_chord_box = filter_HSS_sections_steel(box, 152, 6.4, 203, 8)
    # limit web depth bottom limit 152, top limit 254
    # limit the web to be only square sections
    web_box = filter_HSS_sections_steel(box, 152, 6.4, 203, 8, asym=True)
    lateral_box = filter_HSS_sections_steel(box, 127, 6.4, 178, 8, asym=True)

    """ ------------------------ CREATE COMBINATIONS ------------------------ """

    # [top, bottom, web, lateral]
    box_box_box = valid_combinations_steel(
        top_chord_box, bottom_chord_box, web_box, lateral_box
    )
    """ box_box_round = valid_combinations_steel(
        top_chord_box, bottom_chord_box, web_round, web_round
    ) """

    # only do box_box_box for now
    return [box_box_box]


def parse_fraction(s: str) -> float:
    s = s.strip()
    if " " in s:  # mixed number like "1 1/2"
        whole, frac = s.split()
        return float(whole) + float(Fraction(frac))
    return float(Fraction(s))  # simple float or fraction


def write_to_excel(results, path, sheet, first_write=False):

    df = pd.DataFrame(results)

    if first_write:
        with pd.ExcelWriter(path, mode="w") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

    else:
        with pd.ExcelWriter(
            path,
            mode="a",
            if_sheet_exists="replace",
        ) as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

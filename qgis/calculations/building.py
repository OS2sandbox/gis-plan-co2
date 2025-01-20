from .const import constants as c
import json

#from tabulate import tabulate
import math
from .utils.utility import (
    Average,
    find_key_index,
    find_key_index_double,
    average_of_range,
    range_average,
)


# Define the path to the JSON file
import os.path
PATH = os.path.join(os.path.dirname(__file__), 'const')
json_file_path = PATH + '/build/opbyg_dict.json'

# Read and parse the JSON file
with open(json_file_path, "r") as json_file:
    opbygDict = json.load(json_file)

json_file_path = PATH + '/structure/structure_number_of_floors.json'
with open(json_file_path) as f:
    structure_number_of_floors = json.load(f)

json_file_path = PATH + '/structure/structure_floor_height.json'
with open(json_file_path) as f:
    structure_floor_height = json.load(f)

json_file_path = PATH + '/structure/structure_roof_angle.json'
with open(json_file_path) as f:
    structure_roof_angle = json.load(f)

json_file_path = PATH + '/structure/structure_number_of_basement_floors.json'
with open(json_file_path) as f:
    structure_number_of_basement_floors = json.load(f)

json_file_path = PATH + '/structure/structure_basement_floor_height.json'
with open(json_file_path) as f:
    structure_basement_floor_height = json.load(f)


class Building:

    def __init__(
        self,
        study_period,
        construction_year,
        emission_factor,
        phases,
        calculate_people,
        include_demolition,
        num_residents,
        num_workplaces,
        area,
        width,
        perimeter,
        height,
        num_floors,
        num_base_floors,
        basement_depth,
        parking_basement,
        condition,
        typology,
        ground_typology,
        ground_ftf,
        structure,
        wwr,
        prim_facade,
        sec_facade,
        prim_facade_proportion,
        roof_type,
        roof_material,
        roof_angle,
        heating,
    ):

        self.building_typology = typology
        self.ERROR_MESSAGES = []
        (
            roof_angle,
            ROOF_THICKNESS,
            ROOF_HEIGHT,
            ground_ftf,
            ftf,
            ground_typology,
            heating,
            structure,
            slab_thickness,
            ftc,
            ftc_ground,
            wwr,
            prim_facade_proportion,
            sec_facade_proportion,
        ) = self.initialize_dimensions(
            roof_type,
            roof_material,
            roof_angle,
            typology,
            structure,
            width,
            height,
            num_floors,
            ground_typology,
            ground_ftf,
            heating,
            wwr,
            prim_facade,
            sec_facade,
            prim_facade_proportion,
        )

        num_people, num_residents, num_workplaces = self.initialize_people(
            ground_typology,
            typology,
            num_floors,
            area,
            condition,
            num_residents,
            num_workplaces,
        )

        # CHANGED
        index_basement = c.STRUCTURE_OVERVIEW["Assembly version"].index("basement")
        index_hollow_core_slab = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            "hollow-core slab"
        )
        index_concrete_finish = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            "concrete finish"
        )

        if num_base_floors == 0:  # If no basement
            TYPOLOGY_BASEMENT = c.NO_BASEMENT_TYPOLOGY
        else:
            if parking_basement == c.PARKING_BASEMENT[0]:
                TYPOLOGY_BASEMENT = "parking"
            else:
                TYPOLOGY_BASEMENT = typology

        if basement_depth is None:
            if num_base_floors == 0:
                basement_depth = 0
            else:
                basement_depth = (
                    c.MIN_FTC_BASEMENT[TYPOLOGY_BASEMENT]
                    + c.STRUCTURE_OVERVIEW["Thickness [m]"][index_hollow_core_slab]
                    + c.STRUCTURE_OVERVIEW["Thickness [m]"][index_concrete_finish]
                ) * num_base_floors + c.STRUCTURE_OVERVIEW["Thickness [m]"][
                    index_basement
                ]

        (volume_below_ground, volume_above_ground, volume_total, volume_top_floor) = (
            self.quantity_calculation(
                area, basement_depth, height, roof_angle, ROOF_HEIGHT, ftf
            )
        )

        # region Direct quantities
        FLOOR_AREA_TOTAL = area * num_floors
        FLOOR_AREA_ABOVE_GROUND = area * (num_floors - 1)
        FLOOR_AREA_GROUND = area
        BASEMENT_AREA = num_base_floors * area
        TOTAL_AREA = FLOOR_AREA_TOTAL + BASEMENT_AREA

        if roof_type == c.ROOF_TYPE[0]:
            ROOF_FINISH_AREA = area
        else:
            ROOF_FINISH_AREA = (
                (area / width) * (width / 2) / math.cos(math.radians(roof_angle)) * 2
            )

        roof_var = 0
        if roof_type == "steep":
            roof_var = ROOF_HEIGHT / 2
        VOL_ABOVE = area * (height - roof_var)

        if num_base_floors > 0:
            BASE_WALL_AREA = perimeter * basement_depth
        else:
            BASE_WALL_AREA = 0
        # endregion

        # region Calculated quantities

        # In Excel it repeats the wwr calculation here, no need to repeat here
        EXTERNAL_WALL_AREA = perimeter * (height - ROOF_HEIGHT)
        FACADE_AREA = EXTERNAL_WALL_AREA * (1 - wwr)
        PRIM_FACADE_AREA = FACADE_AREA * prim_facade_proportion
        SEC_FACADE_AREA = FACADE_AREA * sec_facade_proportion
        WINDOW_AREA = EXTERNAL_WALL_AREA * wwr

        if (num_floors + num_base_floors) >= c.MIN_FLOORS_STAIRS_DICT[typology]:
            NUM_STAIRCASES = math.ceil(area / c.MAX_AREA_CORE_DICT[typology])
        else:
            NUM_STAIRCASES = 0

        if typology == "detached house" or typology == "terraced house":
            NUM_CORES = 0
        else:
            if (num_floors + num_base_floors) >= c.MIN_FLOORS_CORE_DICT[typology]:
                NUM_CORES = math.ceil(area / c.MAX_AREA_CORE_DICT[typology])
            else:
                NUM_CORES = 0

        NUM_ELEVATORS = NUM_CORES

        index_thickness_facade = self.thickness_facade(
            c.STRUCTURE_OVERVIEW, structure, typology, prim_facade
        )

        facade_object = next(
            (item for item in opbygDict if item["item"] == index_thickness_facade), None
        )

        USE_AREA_ABOVE = (
            (area - (NUM_CORES * math.pow(c.LENGTH_CORE_DICT[typology], 2)))
            - (
                perimeter * facade_object["thickness [m]"]
            )  # ((114.05 - ( 1 * 0.5 ) ) - (44.35 * 0.3293)) * (8-1)
        ) * (num_floors - 1)

        USE_AREA_GROUND = (
            area
            - (
                NUM_CORES * math.pow(c.LENGTH_CORE_DICT[typology], 2)
            )  # * c.LENGTH_CORE_DICT[typology])
        ) - (perimeter * facade_object["thickness [m]"])

        if num_base_floors == 0:  # If no basement
            USE_AREA_BELOW = 0
            FTF_BASEMENT = 0
            FTC_BASEMENT = 0
            # Not sure if this is needed
            LENGTH_INTERNAL_AREA_BASE = 0
            LENGTH_INTERNAL_WALLS_BASE = 0
            AREA_INTERNAL_WALLS_BASE = 0
        else:
            # this is correct with excel
            FTF_BASEMENT = (
                basement_depth - c.STRUCTURE_OVERVIEW["Thickness [m]"][index_basement]
            ) / num_base_floors

            USE_AREA_BELOW = (
                area
                - (
                    NUM_CORES
                    * c.LENGTH_CORE_DICT[TYPOLOGY_BASEMENT]
                    * c.LENGTH_CORE_DICT[typology]
                )
            ) * num_base_floors
            #JCN/2024-10-24
            if FTF_BASEMENT < 0:
                self.ERROR_MESSAGES.append("Too low floor-to-floor basement height")
                return

            # FTC_BASEMENT = FTF_BASEMENT - (c.STRUCTURE_OVERVIEW['Thickness [m]'][index_hollow_core_slab] +
            #                                c.STRUCTURE_OVERVIEW['Thickness [m]'][index_concrete_finish])
            # wrong in excel, it references FTC above ground instead of basement
            # FIX EXCEL AND THEN USE THE ABOVE
            FTC_BASEMENT = ftc

            if FTC_BASEMENT < c.MIN_FTC_BASEMENT[TYPOLOGY_BASEMENT]:
                self.ERROR_MESSAGES.append("Too low floor-to-ceiling height")
            LENGTH_INTERNAL_AREA_BASE = c.INTERNAL_WALLS_LENGTH_DICT[TYPOLOGY_BASEMENT]
            LENGTH_INTERNAL_WALLS_BASE = LENGTH_INTERNAL_AREA_BASE * USE_AREA_BELOW
            AREA_INTERNAL_WALLS_BASE = LENGTH_INTERNAL_WALLS_BASE * FTC_BASEMENT

        VOL_BELOW = area * basement_depth
        TOTAL_VOL = VOL_BELOW + VOL_ABOVE

        TOTAL_BUILD_HEIGHT = height + basement_depth

        STAIRCASE_HEIGHT = (TOTAL_BUILD_HEIGHT - (ftf + ROOF_HEIGHT)) * NUM_STAIRCASES

        TOTAL_CORE_HEIGHT = (TOTAL_BUILD_HEIGHT - ROOF_HEIGHT) * NUM_CORES

        LENGTH_STAIRCASE = TOTAL_CORE_HEIGHT
        if TOTAL_CORE_HEIGHT == 0:
            LENGTH_STAIRCASE = STAIRCASE_HEIGHT

        # Building height [m] * 4 (since it’s a rectangular core with 4 sides) * core side size [m] (refers to constant value in the tab building constants) = core wall area [m2]
        CORE_AREA = TOTAL_CORE_HEIGHT * 4 * c.CORE_LENGTH

        LENGTH_INTERNAL_AREA = c.INTERNAL_WALLS_LENGTH_DICT[typology]  # detach 0.2

        LENGTH_INTERNAL_WALLS = (
            LENGTH_INTERNAL_AREA * USE_AREA_ABOVE
            + c.INTERNAL_WALLS_LENGTH_DICT_M[typology]  # 0
        )

        AREA_INTERNAL_WALLS = LENGTH_INTERNAL_WALLS * ftc

        # new
        if typology == "detached house" or typology == "terraced house":
            AREA_LOW_SETTLEMENT = FLOOR_AREA_ABOVE_GROUND
            AREA_LOW_SETTLEMENT_BELOW = BASEMENT_AREA
        else:
            AREA_LOW_SETTLEMENT = 0
            AREA_LOW_SETTLEMENT_BELOW = 0

        LENGTH_INTERNAL_AREA_GROUND = c.INTERNAL_WALLS_LENGTH_DICT[ground_typology]
        LENGTH_INTERNAL_WALLS_GROUND = (
            LENGTH_INTERNAL_AREA_GROUND * USE_AREA_GROUND
            + c.INTERNAL_WALLS_LENGTH_DICT_M[ground_typology]
        )

        AREA_INTERNAL_WALLS_GROUND = LENGTH_INTERNAL_WALLS_GROUND * ftc_ground
        AREA_DEMOLITION = TOTAL_AREA
        # endregion

        # region structural quantities

        (
            item_key_structure_number_of_floors,
            item_structure_floor_height,
            item_structure_roof_angle,
            item_key_structure_number_of_basement_floors,
            item_structure_basement_floor_height,
        ) = self.structure_calculation(typology, structure, num_floors, num_base_floors)

        SLAB_PER_VOLUME = item_key_structure_number_of_floors["slab (above ground)"] * (
            item_structure_floor_height["slab (above ground)"]["a"] * math.pow(ftf, -1)
        ) + (
            (volume_top_floor / volume_above_ground)
            * (
                item_structure_roof_angle["slab (above ground)"]["a"] * roof_angle
                + item_structure_roof_angle["slab (above ground)"]["b"]
            )
        )

        BEAM_PER_VOLUME = item_key_structure_number_of_floors["beam (above ground)"] * (
            item_structure_floor_height["beam (above ground)"]["a"] * math.pow(ftf, -1)
        ) + (
            (volume_top_floor / volume_above_ground)
            * (
                item_structure_roof_angle["beam (above ground)"]["a"]
                * math.pow(roof_angle, 2)
                + item_structure_roof_angle["beam (above ground)"]["b"] * roof_angle
                + item_structure_roof_angle["beam (above ground)"]["c"]
            )
        )

        VERTICAL_PER_VOLUME = item_key_structure_number_of_floors[
            "vertical (above ground)"
        ] * (
            item_structure_floor_height["vertical (above ground)"]["a"]
            * math.pow(ftf, 2)
            + item_structure_floor_height["vertical (above ground)"]["b"] * ftf
            + item_structure_floor_height["vertical (above ground)"]["c"]
        ) + (
            (volume_top_floor / volume_above_ground)
            * (
                item_structure_roof_angle["vertical (above ground)"]["a"]
                * math.pow(roof_angle, 2)
                + item_structure_roof_angle["vertical (above ground)"]["b"] * roof_angle
                + item_structure_roof_angle["vertical (above ground)"]["c"]
            )
        )
        if item_key_structure_number_of_basement_floors is not None:
            SLAB_BELOW_PER_VOLUME = item_key_structure_number_of_basement_floors[
                "slab (below ground)"
            ] * (
                item_structure_basement_floor_height["slab (below ground)"]["a"]
                * math.pow(FTF_BASEMENT, -1)
            )

            BEAM_BELOW_PER_VOLUME = item_key_structure_number_of_basement_floors[
                "beam (below ground)"
            ] * (
                item_structure_basement_floor_height["beam (below ground)"]["a"]
                * math.pow(FTF_BASEMENT, -1)
            )

            VERTICAL_BELOW_PER_VOLUME = item_key_structure_number_of_basement_floors[
                "vertical (below ground)"
            ] * (
                item_structure_basement_floor_height["vertical (below ground)"]["a"]
                * FTF_BASEMENT
                + item_structure_basement_floor_height["vertical (below ground)"]["b"]
            )
        else:
            SLAB_BELOW_PER_VOLUME = 0
            BEAM_BELOW_PER_VOLUME = 0
            VERTICAL_BELOW_PER_VOLUME = 0

        foundation_below = 0
        if num_base_floors != 0:
            foundation_below = (
                item_key_structure_number_of_basement_floors["foundation"]
                * (
                    item_structure_basement_floor_height["foundation"]["a"]
                    * math.pow(
                        FTF_BASEMENT,
                        item_structure_basement_floor_height["foundation"]["b"],
                    )
                )
            ) * (volume_below_ground / volume_total)

        FOUNDATION_PER_VOLUME = (
            (
                item_key_structure_number_of_floors["foundation"]
                * (
                    item_structure_floor_height["foundation"]["a"] * math.pow(ftf, 2)
                    + item_structure_floor_height["foundation"]["b"] * ftf
                    + item_structure_floor_height["foundation"]["c"]
                )
                + (
                    (volume_top_floor / volume_above_ground)
                    * (
                        item_structure_roof_angle["foundation"]["a"]
                        * math.pow(roof_angle, 2)
                        + item_structure_roof_angle["foundation"]["b"] * roof_angle
                        + item_structure_roof_angle["foundation"]["c"]
                    )
                )
            )
            * (volume_above_ground / volume_total)
        ) + foundation_below

        SLAB = SLAB_PER_VOLUME * volume_above_ground
        BEAM = BEAM_PER_VOLUME * volume_above_ground
        VERTICAL = VERTICAL_PER_VOLUME * volume_above_ground
        SLAB_BELOW = SLAB_BELOW_PER_VOLUME * volume_below_ground
        BEAM_BELOW = BEAM_BELOW_PER_VOLUME * volume_below_ground
        VERTICAL_BELOW = VERTICAL_BELOW_PER_VOLUME * volume_below_ground
        FOUNDATION = FOUNDATION_PER_VOLUME * volume_total

        # region Modified final values
        # List to define how much material we will keep
        TRANSFORMATION_POTENTIAL = c.BUILDING_CONDITION[condition]
        # Correct quantities - All quantities and Modified quantities section in 1.1_BygningBeregning
        FOUNDATION *= TRANSFORMATION_POTENTIAL[0]
        LENGTH_LINE = TRANSFORMATION_POTENTIAL[0] * perimeter
        AREA_GROUND_SLAB = area * TRANSFORMATION_POTENTIAL[0]
        SLAB *= TRANSFORMATION_POTENTIAL[1]
        BEAM *= TRANSFORMATION_POTENTIAL[1]
        VERTICAL *= TRANSFORMATION_POTENTIAL[1]
        SLAB_BELOW *= TRANSFORMATION_POTENTIAL[0]
        BEAM_BELOW *= TRANSFORMATION_POTENTIAL[0]
        VERTICAL_BELOW *= TRANSFORMATION_POTENTIAL[0]
        BASE_WALL_AREA *= TRANSFORMATION_POTENTIAL[0]
        PRIM_FACADE_AREA *= TRANSFORMATION_POTENTIAL[2]
        SEC_FACADE_AREA *= TRANSFORMATION_POTENTIAL[2]
        WINDOW_AREA *= TRANSFORMATION_POTENTIAL[3]
        CORE_AREA *= TRANSFORMATION_POTENTIAL[1]
        AREA_INTERNAL_WALLS *= TRANSFORMATION_POTENTIAL[4]
        AREA_INTERNAL_WALLS_GROUND *= TRANSFORMATION_POTENTIAL[4]
        AREA_INTERNAL_WALLS_BASE *= TRANSFORMATION_POTENTIAL[4]
        AREA_ROOF_FINISH = ROOF_FINISH_AREA * TRANSFORMATION_POTENTIAL[2]
        AREA_FLOOR_FINISH = USE_AREA_ABOVE * TRANSFORMATION_POTENTIAL[4]
        AREA_FLOOR_FINISH_GROUND = USE_AREA_GROUND * TRANSFORMATION_POTENTIAL[4]
        AREA_FLOOR_FINISH_BELOW = USE_AREA_BELOW * TRANSFORMATION_POTENTIAL[4]
        AREA_CEILING_FINISH = USE_AREA_ABOVE * TRANSFORMATION_POTENTIAL[4]
        AREA_CEILING_FINISH_GROUND = USE_AREA_GROUND * TRANSFORMATION_POTENTIAL[4]
        AREA_CEILING_FINISH_BELOW = USE_AREA_BELOW * TRANSFORMATION_POTENTIAL[4]
        LENGTH_STAIRCASE *= TRANSFORMATION_POTENTIAL[4]
        AREA_TECHNICAL = FLOOR_AREA_ABOVE_GROUND * TRANSFORMATION_POTENTIAL[5]
        AREA_TECHNICAL_GROUND = FLOOR_AREA_GROUND * TRANSFORMATION_POTENTIAL[5]
        AREA_TECHNICAL_BELOW = BASEMENT_AREA * TRANSFORMATION_POTENTIAL[5]
        NUM_ELEVATORS = NUM_ELEVATORS * TRANSFORMATION_POTENTIAL[5]
        AREA_OPERATIONAL = FLOOR_AREA_ABOVE_GROUND * TRANSFORMATION_POTENTIAL[7]
        # AREA_OPERATIONAL
        AREA_OPERATIONAL_GROUND = FLOOR_AREA_GROUND * TRANSFORMATION_POTENTIAL[7]
        AREA_OPERATIONAL_BELOW = BASEMENT_AREA * TRANSFORMATION_POTENTIAL[7]
        AREA_CONSTRUCTION = TOTAL_AREA * TRANSFORMATION_POTENTIAL[6]
        AREA_LOW_SETTLEMENT *= TRANSFORMATION_POTENTIAL[1]
        AREA_LOW_SETTLEMENT_BELOW *= TRANSFORMATION_POTENTIAL[0]
        # endregion

        # region Collect indices of assembly IDs
        index_study_period = c.STUDY_PERIOD.index(study_period)

        item_pad = next(
            (
                item
                for item in opbygDict
                if item["assembly_version"] == c.PAD_FOUNDATION[structure]
            ),
            None,
        )
        index_pad = item_pad["item"]

        if typology == "detached house" or typology == "terraced house":

            index_line_foundation_item = next(
                (
                    item
                    for item in opbygDict
                    if item["assembly_version"] == "low settlement"
                ),
                None,
            )
            index_line_foundation = index_line_foundation_item["item"]

        else:
            index_line_foundation = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                c.LINE_FOUNDATION[structure]
            )

        if num_base_floors > 0:
            index_ground_slab = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "basement"
            )
        else:
            if typology == "detached house" or typology == "terraced house":
                index_ground_slab = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                    "low settlement no basement"
                )
            else:
                index_ground_slab = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                    "high settlement no basement"
                )

        index_slab_above = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.SLAB_STRUCTURE[structure]
        )
        index_vertical_above = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.VERTICAL_STRUCTURE[structure]
        )
        index_beam_above = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.BEAM_STRUCTURE[structure]
        )

        index_slab_below = index_hollow_core_slab
        index_vertical_below = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            "reinforced concrete column"
        )
        index_beam_below = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            "reinforced concrete beam"
        )
        if num_base_floors > 2:
            index_basement_wall = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "3 floors"
            )
        elif num_base_floors == 2:
            index_basement_wall = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "2 floors"
            )
        else:
            index_basement_wall = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "1 floor"
            )

        if typology == "detached house" or typology == "terraced house":
            if "concrete" in structure:
                index_primary_facade = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                    "low settlement concrete building " + prim_facade
                )
                if SEC_FACADE_AREA > 0:
                    index_sec_facade = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                        "low settlement concrete building " + sec_facade
                    )
            else:
                index_primary_facade = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                    "low settlement timber building " + prim_facade
                )
                if SEC_FACADE_AREA > 0:
                    index_sec_facade = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                        "low settlement timber building " + sec_facade
                    )
        else:
            index_primary_facade = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                prim_facade
            )
            if SEC_FACADE_AREA > 0:
                index_sec_facade = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                    sec_facade
                )

        index_window = c.STRUCTURE_OVERVIEW["Assembly version"].index("window")

        index_core_wall = find_key_index_double(
            c.STRUCTURE_OVERVIEW,
            "Assembly",
            "Assembly version",
            "core wall",
            c.CORE_WALL[structure],
        )
        index_internal_wall = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.INTERNAL_WALLS[structure]
        )
        index_internal_basement_wall = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            "steel stud (light)"
        )

        if typology == "detached house" or typology == "terraced house":
            if "concrete" in structure:
                index_roof_finish = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                    "low settlement concrete building "
                    + roof_type
                    + " "
                    + roof_material
                )
            else:
                index_roof_finish = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                    "low settlement timber building " + roof_type + " " + roof_material
                )
        else:
            index_roof_finish = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                roof_type + " " + roof_material
            )

        index_floor_finish_above = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.FLOOR_FINISH_DICT[typology]
        )
        index_floor_finish_ground = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.FLOOR_FINISH_DICT[ground_typology]
        )
        index_floor_finish_below = index_concrete_finish

        index_ceiling_finish_above = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.CEILING_FINISH_DICT[typology]
        )
        index_ceiling_finish_ground = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.CEILING_FINISH_DICT[ground_typology]
        )

        try:
            index_ceiling_finish_below = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                c.CEILING_FINISH_DICT[TYPOLOGY_BASEMENT]
            )
        except:
            index_ceiling_finish_below = index_ceiling_finish_above

        total_num_floors = num_floors + num_base_floors
        if total_num_floors >= c.MIN_FLOORS_CORE_DICT[typology] and (
            typology != "detached house" and typology != "terraced house"
        ):
            index_staircase = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "staircase and elevator"
            )
        else:
            index_staircase = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "staircase"
            )

        index_elevator = c.STRUCTURE_OVERVIEW["Assembly version"].index("elevator")

        index_technical_above = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.TECHNICAL_SYSTEMS_DICT[typology]
        )
        index_technical_ground = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            c.TECHNICAL_SYSTEMS_DICT[ground_typology]
        )
        if num_base_floors != 0 and num_base_floors is not None:
            index_technical_below = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                c.TECHNICAL_SYSTEMS_DICT[TYPOLOGY_BASEMENT]
            )

        if "concrete" in structure:
            index_low_slab_above = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "low settlement hollow-core slab"
            )
        else:
            index_low_slab_above = c.STRUCTURE_OVERVIEW["Assembly version"].index(
                "rib panels"
            )

        index_low_slab_below = c.STRUCTURE_OVERVIEW["Assembly version"].index(
            "low settlement hollow-core slab"
        )
        # endregion

        # region Emissions

        # region Emissions of selected assemblies
        PAD_GWP_A1_A3 = c.GWP_A1_A3[index_pad][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            PAD_GWP_A4 = c.GWP_A4[index_pad][index_study_period]
            PAD_GWP_A5 = c.GWP_A5[index_pad][index_study_period]
        else:
            PAD_GWP_A4 = 0
            PAD_GWP_A5 = 0
        PAD_GWP_B4 = c.GWP_B4[index_pad][index_study_period]
        PAD_GWP_C3 = c.GWP_C3[index_pad][index_study_period]
        PAD_GWP_C4 = c.GWP_C4[index_pad][index_study_period]
        PAD_GWP_D = c.GWP_D[index_pad][index_study_period]
        PAD_GWP = (
            PAD_GWP_A1_A3
            + PAD_GWP_A4
            + PAD_GWP_A5
            + PAD_GWP_B4
            + PAD_GWP_C3
            + PAD_GWP_C4
        )

        LINE_GWP_A1_A3 = c.GWP_A1_A3[index_line_foundation][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            LINE_GWP_A4 = c.GWP_A4[index_line_foundation][index_study_period]
            LINE_GWP_A5 = c.GWP_A5[index_line_foundation][index_study_period]
        else:
            LINE_GWP_A4 = 0
            LINE_GWP_A5 = 0
        LINE_GWP_B4 = c.GWP_B4[index_line_foundation][index_study_period]
        LINE_GWP_C3 = c.GWP_C3[index_line_foundation][index_study_period]
        LINE_GWP_C4 = c.GWP_C4[index_line_foundation][index_study_period]
        LINE_GWP_D = c.GWP_D[index_line_foundation][index_study_period]
        LINE_GWP = (
            LINE_GWP_A1_A3
            + LINE_GWP_A4
            + LINE_GWP_A5
            + LINE_GWP_B4
            + LINE_GWP_C3
            + LINE_GWP_C4
        )

        GROUND_SLAB_GWP_A1_A3 = c.GWP_A1_A3[index_ground_slab][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            GROUND_SLAB_GWP_A4 = c.GWP_A4[index_ground_slab][index_study_period]
            GROUND_SLAB_GWP_A5 = c.GWP_A5[index_ground_slab][index_study_period]
        else:
            GROUND_SLAB_GWP_A4 = 0
            GROUND_SLAB_GWP_A5 = 0
        GROUND_SLAB_GWP_B4 = c.GWP_B4[index_ground_slab][index_study_period]
        GROUND_SLAB_GWP_C3 = c.GWP_C3[index_ground_slab][index_study_period]
        GROUND_SLAB_GWP_C4 = c.GWP_C4[index_ground_slab][index_study_period]
        GROUND_SLAB_GWP_D = c.GWP_D[index_ground_slab][index_study_period]
        GROUND_SLAB_GWP = (
            GROUND_SLAB_GWP_A1_A3
            + GROUND_SLAB_GWP_A4
            + GROUND_SLAB_GWP_A5
            + GROUND_SLAB_GWP_B4
            + GROUND_SLAB_GWP_C3
            + GROUND_SLAB_GWP_C4
        )

        LOW_SLAB_A_GWP_A1_A3 = c.GWP_A1_A3[index_low_slab_above][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            LOW_SLAB_A_GWP_A4 = c.GWP_A4[index_low_slab_above][index_study_period]
            LOW_SLAB_A_GWP_A5 = c.GWP_A5[index_low_slab_above][index_study_period]
        else:
            LOW_SLAB_A_GWP_A4 = 0
            LOW_SLAB_A_GWP_A5 = 0
        LOW_SLAB_A_GWP_B4 = c.GWP_B4[index_low_slab_above][index_study_period]
        LOW_SLAB_A_GWP_C3 = c.GWP_C3[index_low_slab_above][index_study_period]
        LOW_SLAB_A_GWP_C4 = c.GWP_C4[index_low_slab_above][index_study_period]
        LOW_SLAB_A_GWP_D = c.GWP_D[index_low_slab_above][index_study_period]
        LOW_SLAB_A_GWP = (
            LOW_SLAB_A_GWP_A1_A3
            + LOW_SLAB_A_GWP_A4
            + LOW_SLAB_A_GWP_A5
            + LOW_SLAB_A_GWP_B4
            + LOW_SLAB_A_GWP_C3
            + LOW_SLAB_A_GWP_C4
        )

        LOW_SLAB_B_GWP_A1_A3 = c.GWP_A1_A3[index_low_slab_below][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            LOW_SLAB_B_GWP_A4 = c.GWP_A4[index_low_slab_below][index_study_period]
            LOW_SLAB_B_GWP_A5 = c.GWP_A5[index_low_slab_below][index_study_period]
        else:
            LOW_SLAB_B_GWP_A4 = 0
            LOW_SLAB_B_GWP_A5 = 0
        LOW_SLAB_B_GWP_B4 = c.GWP_B4[index_low_slab_below][index_study_period]
        LOW_SLAB_B_GWP_C3 = c.GWP_C3[index_low_slab_below][index_study_period]
        LOW_SLAB_B_GWP_C4 = c.GWP_C4[index_low_slab_below][index_study_period]
        LOW_SLAB_B_GWP_D = c.GWP_D[index_low_slab_below][index_study_period]
        LOW_SLAB_B_GWP = (
            LOW_SLAB_B_GWP_A1_A3
            + LOW_SLAB_B_GWP_A4
            + LOW_SLAB_B_GWP_A5
            + LOW_SLAB_B_GWP_B4
            + LOW_SLAB_B_GWP_C3
            + LOW_SLAB_B_GWP_C4
        )

        SLAB_ABOVE_GWP_A1_A3 = c.GWP_A1_A3[index_slab_above][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            SLAB_ABOVE_GWP_A4 = c.GWP_A4[index_slab_above][index_study_period]
            SLAB_ABOVE_GWP_A5 = c.GWP_A5[index_slab_above][index_study_period]
        else:
            SLAB_ABOVE_GWP_A4 = 0
            SLAB_ABOVE_GWP_A5 = 0
        SLAB_ABOVE_GWP_B4 = c.GWP_B4[index_slab_above][index_study_period]
        SLAB_ABOVE_GWP_C3 = c.GWP_C3[index_slab_above][index_study_period]
        SLAB_ABOVE_GWP_C4 = c.GWP_C4[index_slab_above][index_study_period]
        SLAB_ABOVE_GWP_D = c.GWP_D[index_slab_above][index_study_period]
        SLAB_ABOVE_GWP = (
            SLAB_ABOVE_GWP_A1_A3
            + SLAB_ABOVE_GWP_A4
            + SLAB_ABOVE_GWP_A5
            + SLAB_ABOVE_GWP_B4
            + SLAB_ABOVE_GWP_C3
            + SLAB_ABOVE_GWP_C4
        )

        VERTICAL_ABOVE_GWP_A1_A3 = c.GWP_A1_A3[index_vertical_above][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            VERTICAL_ABOVE_GWP_A4 = c.GWP_A4[index_vertical_above][index_study_period]
            VERTICAL_ABOVE_GWP_A5 = c.GWP_A5[index_vertical_above][index_study_period]
        else:
            VERTICAL_ABOVE_GWP_A4 = 0
            VERTICAL_ABOVE_GWP_A5 = 0
        VERTICAL_ABOVE_GWP_B4 = c.GWP_B4[index_vertical_above][index_study_period]
        VERTICAL_ABOVE_GWP_C3 = c.GWP_C3[index_vertical_above][index_study_period]
        VERTICAL_ABOVE_GWP_C4 = c.GWP_C4[index_vertical_above][index_study_period]
        VERTICAL_ABOVE_GWP_D = c.GWP_D[index_vertical_above][index_study_period]
        VERTICAL_ABOVE_GWP = (
            VERTICAL_ABOVE_GWP_A1_A3
            + VERTICAL_ABOVE_GWP_A4
            + VERTICAL_ABOVE_GWP_A5
            + VERTICAL_ABOVE_GWP_B4
            + VERTICAL_ABOVE_GWP_C3
            + VERTICAL_ABOVE_GWP_C4
        )

        BEAM_ABOVE_GWP_A1_A3 = c.GWP_A1_A3[index_beam_above][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            BEAM_ABOVE_GWP_A4 = c.GWP_A4[index_beam_above][index_study_period]
            BEAM_ABOVE_GWP_A5 = c.GWP_A5[index_beam_above][index_study_period]
        else:
            BEAM_ABOVE_GWP_A4 = 0
            BEAM_ABOVE_GWP_A5 = 0
        BEAM_ABOVE_GWP_B4 = c.GWP_B4[index_beam_above][index_study_period]
        BEAM_ABOVE_GWP_C3 = c.GWP_C3[index_beam_above][index_study_period]
        BEAM_ABOVE_GWP_C4 = c.GWP_C4[index_beam_above][index_study_period]
        BEAM_ABOVE_GWP_D = c.GWP_D[index_beam_above][index_study_period]
        BEAM_ABOVE_GWP = (
            BEAM_ABOVE_GWP_A1_A3
            + BEAM_ABOVE_GWP_A4
            + BEAM_ABOVE_GWP_A5
            + BEAM_ABOVE_GWP_B4
            + BEAM_ABOVE_GWP_C3
            + BEAM_ABOVE_GWP_C4
        )

        SLAB_BELOW_GWP_A1_A3 = c.GWP_A1_A3[index_slab_below][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            SLAB_BELOW_GWP_A4 = c.GWP_A4[index_slab_below][index_study_period]
            SLAB_BELOW_GWP_A5 = c.GWP_A5[index_slab_below][index_study_period]
        else:
            SLAB_BELOW_GWP_A4 = 0
            SLAB_BELOW_GWP_A5 = 0
        SLAB_BELOW_GWP_B4 = c.GWP_B4[index_slab_below][index_study_period]
        SLAB_BELOW_GWP_C3 = c.GWP_C3[index_slab_below][index_study_period]
        SLAB_BELOW_GWP_C4 = c.GWP_C4[index_slab_below][index_study_period]
        SLAB_BELOW_GWP_D = c.GWP_D[index_slab_below][index_study_period]
        SLAB_BELOW_GWP = (
            SLAB_BELOW_GWP_A1_A3
            + SLAB_BELOW_GWP_A4
            + SLAB_BELOW_GWP_A5
            + SLAB_BELOW_GWP_B4
            + SLAB_BELOW_GWP_C3
            + SLAB_BELOW_GWP_C4
        )

        VERTICAL_BELOW_GWP_A1_A3 = c.GWP_A1_A3[index_vertical_below][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            VERTICAL_BELOW_GWP_A4 = c.GWP_A4[index_vertical_below][index_study_period]
            VERTICAL_BELOW_GWP_A5 = c.GWP_A5[index_vertical_below][index_study_period]
        else:
            VERTICAL_BELOW_GWP_A4 = 0
            VERTICAL_BELOW_GWP_A5 = 0
        VERTICAL_BELOW_GWP_B4 = c.GWP_B4[index_vertical_below][index_study_period]
        VERTICAL_BELOW_GWP_C3 = c.GWP_C3[index_vertical_below][index_study_period]
        VERTICAL_BELOW_GWP_C4 = c.GWP_C4[index_vertical_below][index_study_period]
        VERTICAL_BELOW_GWP_D = c.GWP_D[index_vertical_below][index_study_period]
        VERTICAL_BELOW_GWP = (
            VERTICAL_BELOW_GWP_A1_A3
            + VERTICAL_BELOW_GWP_A4
            + VERTICAL_BELOW_GWP_A5
            + VERTICAL_BELOW_GWP_B4
            + VERTICAL_BELOW_GWP_C3
            + VERTICAL_BELOW_GWP_C4
        )

        BEAM_BELOW_GWP_A1_A3 = c.GWP_A1_A3[index_beam_below][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            BEAM_BELOW_GWP_A4 = c.GWP_A4[index_beam_below][index_study_period]
            BEAM_BELOW_GWP_A5 = c.GWP_A5[index_beam_below][index_study_period]
        else:
            BEAM_BELOW_GWP_A4 = 0
            BEAM_BELOW_GWP_A5 = 0
        BEAM_BELOW_GWP_B4 = c.GWP_B4[index_beam_below][index_study_period]
        BEAM_BELOW_GWP_C3 = c.GWP_C3[index_beam_below][index_study_period]
        BEAM_BELOW_GWP_C4 = c.GWP_C4[index_beam_below][index_study_period]
        BEAM_BELOW_GWP_D = c.GWP_D[index_beam_below][index_study_period]
        BEAM_BELOW_GWP = (
            BEAM_BELOW_GWP_A1_A3
            + BEAM_BELOW_GWP_A4
            + BEAM_BELOW_GWP_A5
            + BEAM_BELOW_GWP_B4
            + BEAM_BELOW_GWP_C3
            + BEAM_BELOW_GWP_C4
        )

        BASEMENT_WALL_GWP_A1_A3 = c.GWP_A1_A3[index_basement_wall][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            BASEMENT_WALL_GWP_A4 = c.GWP_A4[index_basement_wall][index_study_period]
            BASEMENT_WALL_GWP_A5 = c.GWP_A5[index_basement_wall][index_study_period]
        else:
            BASEMENT_WALL_GWP_A4 = 0
            BASEMENT_WALL_GWP_A5 = 0
        BASEMENT_WALL_GWP_B4 = c.GWP_B4[index_basement_wall][index_study_period]
        BASEMENT_WALL_GWP_C3 = c.GWP_C3[index_basement_wall][index_study_period]
        BASEMENT_WALL_GWP_C4 = c.GWP_C4[index_basement_wall][index_study_period]
        BASEMENT_WALL_GWP_D = c.GWP_D[index_basement_wall][index_study_period]
        BASEMENT_WALL_GWP = (
            BASEMENT_WALL_GWP_A1_A3
            + BASEMENT_WALL_GWP_A4
            + BASEMENT_WALL_GWP_A5
            + BASEMENT_WALL_GWP_B4
            + BASEMENT_WALL_GWP_C3
            + BASEMENT_WALL_GWP_C4
        )

        PRIMARY_FACADE_GWP_A1_A3 = c.GWP_A1_A3[index_primary_facade][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            PRIMARY_FACADE_GWP_A4 = c.GWP_A4[index_primary_facade][index_study_period]
            PRIMARY_FACADE_GWP_A5 = c.GWP_A5[index_primary_facade][index_study_period]
        else:
            PRIMARY_FACADE_GWP_A4 = 0
            PRIMARY_FACADE_GWP_A5 = 0
        PRIMARY_FACADE_GWP_B4 = c.GWP_B4[index_primary_facade][index_study_period]
        PRIMARY_FACADE_GWP_C3 = c.GWP_C3[index_primary_facade][index_study_period]
        PRIMARY_FACADE_GWP_C4 = c.GWP_C4[index_primary_facade][index_study_period]
        PRIMARY_FACADE_GWP_D = c.GWP_D[index_primary_facade][index_study_period]
        PRIMARY_FACADE_GWP = (
            PRIMARY_FACADE_GWP_A1_A3
            + PRIMARY_FACADE_GWP_A4
            + PRIMARY_FACADE_GWP_A5
            + PRIMARY_FACADE_GWP_B4
            + PRIMARY_FACADE_GWP_C3
            + PRIMARY_FACADE_GWP_C4
        )

        # If there is a secondary facade
        if SEC_FACADE_AREA > 0:
            SECONDARY_FACADE_GWP_A1_A3 = c.GWP_A1_A3[index_sec_facade][
                index_study_period
            ]
            if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
                SECONDARY_FACADE_GWP_A4 = c.GWP_A4[index_sec_facade][index_study_period]
                SECONDARY_FACADE_GWP_A5 = c.GWP_A5[index_sec_facade][index_study_period]
            else:
                SECONDARY_FACADE_GWP_A4 = 0
                SECONDARY_FACADE_GWP_A5 = 0
            SECONDARY_FACADE_GWP_B4 = c.GWP_B4[index_sec_facade][index_study_period]
            SECONDARY_FACADE_GWP_C3 = c.GWP_C3[index_sec_facade][index_study_period]
            SECONDARY_FACADE_GWP_C4 = c.GWP_C4[index_sec_facade][index_study_period]
            SECONDARY_FACADE_GWP_D = c.GWP_D[index_sec_facade][index_study_period]
            SECONDARY_FACADE_GWP = (
                SECONDARY_FACADE_GWP_A1_A3
                + SECONDARY_FACADE_GWP_A4
                + SECONDARY_FACADE_GWP_A5
                + SECONDARY_FACADE_GWP_B4
                + SECONDARY_FACADE_GWP_C3
                + SECONDARY_FACADE_GWP_C4
            )
        else:
            SECONDARY_FACADE_GWP_A1_A3 = 0
            SECONDARY_FACADE_GWP_A4 = 0
            SECONDARY_FACADE_GWP_A5 = 0
            SECONDARY_FACADE_GWP_B4 = 0
            SECONDARY_FACADE_GWP_C3 = 0
            SECONDARY_FACADE_GWP_C4 = 0
            SECONDARY_FACADE_GWP_D = 0
            SECONDARY_FACADE_GWP = 0

        WINDOW_GWP_A1_A3 = c.GWP_A1_A3[index_window][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            WINDOW_GWP_A4 = c.GWP_A4[index_window][index_study_period]
            WINDOW_GWP_A5 = c.GWP_A5[index_window][index_study_period]
        else:
            WINDOW_GWP_A4 = 0
            WINDOW_GWP_A5 = 0
        WINDOW_GWP_B4 = c.GWP_B4[index_window][index_study_period]
        WINDOW_GWP_C3 = c.GWP_C3[index_window][index_study_period]
        WINDOW_GWP_C4 = c.GWP_C4[index_window][index_study_period]
        WINDOW_GWP_D = c.GWP_D[index_window][index_study_period]
        WINDOW_GWP = (
            WINDOW_GWP_A1_A3
            + WINDOW_GWP_A4
            + WINDOW_GWP_A5
            + WINDOW_GWP_B4
            + WINDOW_GWP_C3
            + WINDOW_GWP_C4
        )

        CORE_WALL_GWP_A1_A3 = c.GWP_A1_A3[index_core_wall][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            CORE_WALL_GWP_A4 = c.GWP_A4[index_core_wall][index_study_period]
            CORE_WALL_GWP_A5 = c.GWP_A5[index_core_wall][index_study_period]
        else:
            CORE_WALL_GWP_A4 = 0
            CORE_WALL_GWP_A5 = 0
        CORE_WALL_GWP_B4 = c.GWP_B4[index_core_wall][index_study_period]
        CORE_WALL_GWP_C3 = c.GWP_C3[index_core_wall][index_study_period]
        CORE_WALL_GWP_C4 = c.GWP_C4[index_core_wall][index_study_period]
        CORE_WALL_GWP_D = c.GWP_D[index_core_wall][index_study_period]
        CORE_WALL_GWP = (
            CORE_WALL_GWP_A1_A3
            + CORE_WALL_GWP_A4
            + CORE_WALL_GWP_A5
            + CORE_WALL_GWP_B4
            + CORE_WALL_GWP_C3
            + CORE_WALL_GWP_C4
        )

        INTERNAL_WALL_GWP_A1_A3 = c.GWP_A1_A3[index_internal_wall][index_study_period]

        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            INTERNAL_WALL_GWP_A4 = c.GWP_A4[index_internal_wall][index_study_period]
            INTERNAL_WALL_GWP_A5 = c.GWP_A5[index_internal_wall][index_study_period]
        else:
            INTERNAL_WALL_GWP_A4 = 0
            INTERNAL_WALL_GWP_A5 = 0
        INTERNAL_WALL_GWP_B4 = c.GWP_B4[index_internal_wall][index_study_period]
        INTERNAL_WALL_GWP_C3 = c.GWP_C3[index_internal_wall][index_study_period]
        INTERNAL_WALL_GWP_C4 = c.GWP_C4[index_internal_wall][index_study_period]
        INTERNAL_WALL_GWP_D = c.GWP_D[index_internal_wall][index_study_period]
        INTERNAL_WALL_GWP = (
            INTERNAL_WALL_GWP_A1_A3
            + INTERNAL_WALL_GWP_A4
            + INTERNAL_WALL_GWP_A5
            + INTERNAL_WALL_GWP_B4
            + INTERNAL_WALL_GWP_C3
            + INTERNAL_WALL_GWP_C4
        )

        INTERNAL_BASEMENT_WALL_GWP_A1_A3 = c.GWP_A1_A3[index_internal_basement_wall][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            INTERNAL_BASEMENT_WALL_GWP_A4 = c.GWP_A4[index_internal_basement_wall][
                index_study_period
            ]
            INTERNAL_BASEMENT_WALL_GWP_A5 = c.GWP_A5[index_internal_basement_wall][
                index_study_period
            ]
        else:
            INTERNAL_BASEMENT_WALL_GWP_A4 = 0
            INTERNAL_BASEMENT_WALL_GWP_A5 = 0
        INTERNAL_BASEMENT_WALL_GWP_B4 = c.GWP_B4[index_internal_basement_wall][
            index_study_period
        ]
        INTERNAL_BASEMENT_WALL_GWP_C3 = c.GWP_C3[index_internal_basement_wall][
            index_study_period
        ]
        INTERNAL_BASEMENT_WALL_GWP_C4 = c.GWP_C4[index_internal_basement_wall][
            index_study_period
        ]
        INTERNAL_BASEMENT_WALL_GWP_D = c.GWP_D[index_internal_basement_wall][
            index_study_period
        ]
        INTERNAL_BASEMENT_WALL_GWP = (
            INTERNAL_BASEMENT_WALL_GWP_A1_A3
            + INTERNAL_BASEMENT_WALL_GWP_A4
            + INTERNAL_BASEMENT_WALL_GWP_A5
            + INTERNAL_BASEMENT_WALL_GWP_B4
            + INTERNAL_BASEMENT_WALL_GWP_C3
            + INTERNAL_BASEMENT_WALL_GWP_C4
        )

        ROOF_FINISH_GWP_A1_A3 = c.GWP_A1_A3[index_roof_finish][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            ROOF_FINISH_GWP_A4 = c.GWP_A4[index_roof_finish][index_study_period]
            ROOF_FINISH_GWP_A5 = c.GWP_A5[index_roof_finish][index_study_period]
        else:
            ROOF_FINISH_GWP_A4 = 0
            ROOF_FINISH_GWP_A5 = 0
        ROOF_FINISH_GWP_B4 = c.GWP_B4[index_roof_finish][index_study_period]
        ROOF_FINISH_GWP_C3 = c.GWP_C3[index_roof_finish][index_study_period]
        ROOF_FINISH_GWP_C4 = c.GWP_C4[index_roof_finish][index_study_period]
        ROOF_FINISH_GWP_D = c.GWP_D[index_roof_finish][index_study_period]
        ROOF_FINISH_GWP = (
            ROOF_FINISH_GWP_A1_A3
            + ROOF_FINISH_GWP_A4
            + ROOF_FINISH_GWP_A5
            + ROOF_FINISH_GWP_B4
            + ROOF_FINISH_GWP_C3
            + ROOF_FINISH_GWP_C4
        )

        # region Floor
        FLOOR_FINISH_ABOVE_GWP_A1_A3 = c.GWP_A1_A3[index_floor_finish_above][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            FLOOR_FINISH_ABOVE_GWP_A4 = c.GWP_A4[index_floor_finish_above][
                index_study_period
            ]
            FLOOR_FINISH_ABOVE_GWP_A5 = c.GWP_A5[index_floor_finish_above][
                index_study_period
            ]
        else:
            FLOOR_FINISH_ABOVE_GWP_A4 = 0
            FLOOR_FINISH_ABOVE_GWP_A5 = 0
        FLOOR_FINISH_ABOVE_GWP_B4 = c.GWP_B4[index_floor_finish_above][
            index_study_period
        ]
        FLOOR_FINISH_ABOVE_GWP_C3 = c.GWP_C3[index_floor_finish_above][
            index_study_period
        ]
        FLOOR_FINISH_ABOVE_GWP_C4 = c.GWP_C4[index_floor_finish_above][
            index_study_period
        ]
        FLOOR_FINISH_ABOVE_GWP_D = c.GWP_D[index_floor_finish_above][index_study_period]
        FLOOR_FINISH_GWP = (
            FLOOR_FINISH_ABOVE_GWP_A1_A3
            + FLOOR_FINISH_ABOVE_GWP_A4
            + FLOOR_FINISH_ABOVE_GWP_A5
            + FLOOR_FINISH_ABOVE_GWP_B4
            + FLOOR_FINISH_ABOVE_GWP_C3
            + FLOOR_FINISH_ABOVE_GWP_C4
        )

        FLOOR_FINISH_GROUND_GWP_A1_A3 = c.GWP_A1_A3[index_floor_finish_ground][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            FLOOR_FINISH_GROUND_GWP_A4 = c.GWP_A4[index_floor_finish_ground][
                index_study_period
            ]
            FLOOR_FINISH_GROUND_GWP_A5 = c.GWP_A5[index_floor_finish_ground][
                index_study_period
            ]
        else:
            FLOOR_FINISH_GROUND_GWP_A4 = 0
            FLOOR_FINISH_GROUND_GWP_A5 = 0
        FLOOR_FINISH_GROUND_GWP_B4 = c.GWP_B4[index_floor_finish_ground][
            index_study_period
        ]
        FLOOR_FINISH_GROUND_GWP_C3 = c.GWP_C3[index_floor_finish_ground][
            index_study_period
        ]
        FLOOR_FINISH_GROUND_GWP_C4 = c.GWP_C4[index_floor_finish_ground][
            index_study_period
        ]
        FLOOR_FINISH_GROUND_GWP_D = c.GWP_D[index_floor_finish_ground][
            index_study_period
        ]
        FLOOR_FINISH_GROUND_GWP = (
            FLOOR_FINISH_GROUND_GWP_A1_A3
            + FLOOR_FINISH_GROUND_GWP_A4
            + FLOOR_FINISH_GROUND_GWP_A5
            + FLOOR_FINISH_GROUND_GWP_B4
            + FLOOR_FINISH_GROUND_GWP_C3
            + FLOOR_FINISH_GROUND_GWP_C4
        )

        FLOOR_FINISH_BELOW_GWP_A1_A3 = c.GWP_A1_A3[index_floor_finish_below][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            FLOOR_FINISH_BELOW_GWP_A4 = c.GWP_A4[index_floor_finish_below][
                index_study_period
            ]
            FLOOR_FINISH_BELOW_GWP_A5 = c.GWP_A5[index_floor_finish_below][
                index_study_period
            ]
        else:
            FLOOR_FINISH_BELOW_GWP_A4 = 0
            FLOOR_FINISH_BELOW_GWP_A5 = 0
        FLOOR_FINISH_BELOW_GWP_B4 = c.GWP_B4[index_floor_finish_below][
            index_study_period
        ]
        FLOOR_FINISH_BELOW_GWP_C3 = c.GWP_C3[index_floor_finish_below][
            index_study_period
        ]
        FLOOR_FINISH_BELOW_GWP_C4 = c.GWP_C4[index_floor_finish_below][
            index_study_period
        ]
        FLOOR_FINISH_BELOW_GWP_D = c.GWP_D[index_floor_finish_below][index_study_period]
        FLOOR_FINISH_BELOW_GWP = (
            FLOOR_FINISH_BELOW_GWP_A1_A3
            + FLOOR_FINISH_BELOW_GWP_A4
            + FLOOR_FINISH_BELOW_GWP_A5
            + FLOOR_FINISH_BELOW_GWP_B4
            + FLOOR_FINISH_BELOW_GWP_C3
            + FLOOR_FINISH_BELOW_GWP_C4
        )
        # endregion

        # region Ceiling
        CEILING_FINISH_ABOVE_GWP_A1_A3 = c.GWP_A1_A3[index_ceiling_finish_above][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            CEILING_FINISH_ABOVE_GWP_A4 = c.GWP_A4[index_ceiling_finish_above][
                index_study_period
            ]
            CEILING_FINISH_ABOVE_GWP_A5 = c.GWP_A5[index_ceiling_finish_above][
                index_study_period
            ]
        else:
            CEILING_FINISH_ABOVE_GWP_A4 = 0
            CEILING_FINISH_ABOVE_GWP_A5 = 0
        CEILING_FINISH_ABOVE_GWP_B4 = c.GWP_B4[index_ceiling_finish_above][
            index_study_period
        ]
        CEILING_FINISH_ABOVE_GWP_C3 = c.GWP_C3[index_ceiling_finish_above][
            index_study_period
        ]
        CEILING_FINISH_ABOVE_GWP_C4 = c.GWP_C4[index_ceiling_finish_above][
            index_study_period
        ]
        CEILING_FINISH_ABOVE_GWP_D = c.GWP_D[index_ceiling_finish_above][
            index_study_period
        ]
        CEILING_FINISH_ABOVE_GWP = (
            CEILING_FINISH_ABOVE_GWP_A1_A3
            + CEILING_FINISH_ABOVE_GWP_A4
            + CEILING_FINISH_ABOVE_GWP_A5
            + CEILING_FINISH_ABOVE_GWP_B4
            + CEILING_FINISH_ABOVE_GWP_C3
            + CEILING_FINISH_ABOVE_GWP_C4
        )

        CEILING_FINISH_GROUND_GWP_A1_A3 = c.GWP_A1_A3[index_ceiling_finish_ground][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            CEILING_FINISH_GROUND_GWP_A4 = c.GWP_A4[index_ceiling_finish_ground][
                index_study_period
            ]
            CEILING_FINISH_GROUND_GWP_A5 = c.GWP_A5[index_ceiling_finish_ground][
                index_study_period
            ]
        else:
            CEILING_FINISH_GROUND_GWP_A4 = 0
            CEILING_FINISH_GROUND_GWP_A5 = 0
        CEILING_FINISH_GROUND_GWP_B4 = c.GWP_B4[index_ceiling_finish_ground][
            index_study_period
        ]
        CEILING_FINISH_GROUND_GWP_C3 = c.GWP_C3[index_ceiling_finish_ground][
            index_study_period
        ]
        CEILING_FINISH_GROUND_GWP_C4 = c.GWP_C4[index_ceiling_finish_ground][
            index_study_period
        ]
        CEILING_FINISH_GROUND_GWP_D = c.GWP_D[index_ceiling_finish_ground][
            index_study_period
        ]
        CEILING_FINISH_GROUND_GWP = (
            CEILING_FINISH_GROUND_GWP_A1_A3
            + CEILING_FINISH_GROUND_GWP_A4
            + CEILING_FINISH_GROUND_GWP_A5
            + CEILING_FINISH_GROUND_GWP_B4
            + CEILING_FINISH_GROUND_GWP_C3
            + CEILING_FINISH_GROUND_GWP_C4
        )

        CEILING_FINISH_BELOW_GWP_A1_A3 = c.GWP_A1_A3[index_ceiling_finish_below][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            CEILING_FINISH_BELOW_GWP_A4 = c.GWP_A4[index_ceiling_finish_below][
                index_study_period
            ]
            CEILING_FINISH_BELOW_GWP_A5 = c.GWP_A5[index_ceiling_finish_below][
                index_study_period
            ]
        else:
            CEILING_FINISH_BELOW_GWP_A4 = 0
            CEILING_FINISH_BELOW_GWP_A5 = 0
        CEILING_FINISH_BELOW_GWP_B4 = c.GWP_B4[index_ceiling_finish_below][
            index_study_period
        ]
        CEILING_FINISH_BELOW_GWP_C3 = c.GWP_C3[index_ceiling_finish_below][
            index_study_period
        ]
        CEILING_FINISH_BELOW_GWP_C4 = c.GWP_C4[index_ceiling_finish_below][
            index_study_period
        ]
        CEILING_FINISH_BELOW_GWP_D = c.GWP_D[index_ceiling_finish_below][
            index_study_period
        ]
        CEILING_FINISH_BELOW_GWP = (
            CEILING_FINISH_BELOW_GWP_A1_A3
            + CEILING_FINISH_BELOW_GWP_A4
            + CEILING_FINISH_BELOW_GWP_A5
            + CEILING_FINISH_BELOW_GWP_B4
            + CEILING_FINISH_BELOW_GWP_C3
            + CEILING_FINISH_BELOW_GWP_C4
        )
        # endregion

        STAIRCASE_GWP_A1_A3 = c.GWP_A1_A3[index_staircase][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            STAIRCASE_GWP_A4 = c.GWP_A4[index_staircase][index_study_period]
            STAIRCASE_GWP_A5 = c.GWP_A5[index_staircase][index_study_period]
        else:
            STAIRCASE_GWP_A4 = 0
            STAIRCASE_GWP_A5 = 0
        STAIRCASE_GWP_B4 = c.GWP_B4[index_staircase][index_study_period]
        STAIRCASE_GWP_C3 = c.GWP_C3[index_staircase][index_study_period]
        STAIRCASE_GWP_C4 = c.GWP_C4[index_staircase][index_study_period]
        STAIRCASE_GWP_D = c.GWP_D[index_staircase][index_study_period]
        STAIRCASE_GWP = (
            STAIRCASE_GWP_A1_A3
            + STAIRCASE_GWP_A4
            + STAIRCASE_GWP_A5
            + STAIRCASE_GWP_B4
            + STAIRCASE_GWP_C3
            + STAIRCASE_GWP_C4
        )

        ELEVATOR_GWP_A1_A3 = c.GWP_A1_A3[index_elevator][index_study_period]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            ELEVATOR_GWP_A4 = c.GWP_A4[index_elevator][index_study_period]
            ELEVATOR_GWP_A5 = c.GWP_A5[index_elevator][index_study_period]
        else:
            ELEVATOR_GWP_A4 = 0
            ELEVATOR_GWP_A5 = 0
        ELEVATOR_GWP_B4 = c.GWP_B4[index_elevator][index_study_period]
        ELEVATOR_GWP_C3 = c.GWP_C3[index_elevator][index_study_period]
        ELEVATOR_GWP_C4 = c.GWP_C4[index_elevator][index_study_period]
        ELEVATOR_GWP_D = c.GWP_D[index_elevator][index_study_period]
        ELEVATOR_GWP = (
            ELEVATOR_GWP_A1_A3
            + ELEVATOR_GWP_A4
            + ELEVATOR_GWP_A5
            + ELEVATOR_GWP_B4
            + ELEVATOR_GWP_C3
            + ELEVATOR_GWP_C4
        )

        # region Technical
        TECHNICAL_ABOVE_GWP_A1_A3 = c.GWP_A1_A3[index_technical_above][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            TECHNICAL_ABOVE_GWP_A4 = c.GWP_A4[index_technical_above][index_study_period]
            TECHNICAL_ABOVE_GWP_A5 = c.GWP_A5[index_technical_above][index_study_period]
        else:
            TECHNICAL_ABOVE_GWP_A4 = 0
            TECHNICAL_ABOVE_GWP_A5 = 0
        TECHNICAL_ABOVE_GWP_B4 = c.GWP_B4[index_technical_above][index_study_period]
        TECHNICAL_ABOVE_GWP_C3 = c.GWP_C3[index_technical_above][index_study_period]
        TECHNICAL_ABOVE_GWP_C4 = c.GWP_C4[index_technical_above][index_study_period]
        TECHNICAL_ABOVE_GWP_D = c.GWP_D[index_technical_above][index_study_period]
        TECHNICAL_ABOVE_GWP = (
            TECHNICAL_ABOVE_GWP_A1_A3
            + TECHNICAL_ABOVE_GWP_A4
            + TECHNICAL_ABOVE_GWP_A5
            + TECHNICAL_ABOVE_GWP_B4
            + TECHNICAL_ABOVE_GWP_C3
            + TECHNICAL_ABOVE_GWP_C4
        )

        TECHNICAL_GROUND_GWP_A1_A3 = c.GWP_A1_A3[index_technical_ground][
            index_study_period
        ]
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            TECHNICAL_GROUND_GWP_A4 = c.GWP_A4[index_technical_ground][
                index_study_period
            ]
            TECHNICAL_GROUND_GWP_A5 = c.GWP_A5[index_technical_ground][
                index_study_period
            ]
        else:
            TECHNICAL_GROUND_GWP_A4 = 0
            TECHNICAL_GROUND_GWP_A5 = 0
        TECHNICAL_GROUND_GWP_B4 = c.GWP_B4[index_technical_ground][index_study_period]
        TECHNICAL_GROUND_GWP_C3 = c.GWP_C3[index_technical_ground][index_study_period]
        TECHNICAL_GROUND_GWP_C4 = c.GWP_C4[index_technical_ground][index_study_period]
        TECHNICAL_GROUND_GWP_D = c.GWP_D[index_technical_ground][index_study_period]
        TECHNICAL_GROUND_GWP = (
            TECHNICAL_GROUND_GWP_A1_A3
            + TECHNICAL_GROUND_GWP_A4
            + TECHNICAL_GROUND_GWP_A5
            + TECHNICAL_GROUND_GWP_B4
            + TECHNICAL_GROUND_GWP_C3
            + TECHNICAL_GROUND_GWP_C4
        )

        if num_base_floors == 0 or num_base_floors is None:
            TECHNICAL_BELOW_GWP_A1_A3 = TECHNICAL_BELOW_GWP_A4 = (
                TECHNICAL_BELOW_GWP_A5
            ) = TECHNICAL_BELOW_GWP_B4 = TECHNICAL_BELOW_GWP_C3 = (
                TECHNICAL_BELOW_GWP_C4
            ) = TECHNICAL_BELOW_GWP_D = TECHNICAL_BELOW_GWP = 0
        else:
            TECHNICAL_BELOW_GWP_A1_A3 = c.GWP_A1_A3[index_technical_below][
                index_study_period
            ]
            if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
                TECHNICAL_BELOW_GWP_A4 = c.GWP_A4[index_technical_below][
                    index_study_period
                ]
                TECHNICAL_BELOW_GWP_A5 = c.GWP_A5[index_technical_below][
                    index_study_period
                ]
            else:
                TECHNICAL_BELOW_GWP_A4 = 0
                TECHNICAL_BELOW_GWP_A5 = 0
            TECHNICAL_BELOW_GWP_B4 = c.GWP_B4[index_technical_below][index_study_period]
            TECHNICAL_BELOW_GWP_C3 = c.GWP_C3[index_technical_below][index_study_period]
            TECHNICAL_BELOW_GWP_C4 = c.GWP_C4[index_technical_below][index_study_period]
            TECHNICAL_BELOW_GWP_D = c.GWP_D[index_technical_below][index_study_period]
            TECHNICAL_BELOW_GWP = (
                TECHNICAL_BELOW_GWP_A1_A3
                + TECHNICAL_BELOW_GWP_A4
                + TECHNICAL_BELOW_GWP_A5
                + TECHNICAL_BELOW_GWP_B4
                + TECHNICAL_BELOW_GWP_C3
                + TECHNICAL_BELOW_GWP_C4
            )
            # endregion
        # endregion

        # region Emissions of assemblies
        # Foundation
        PAD_GWP_A1_A3_TOTAL = PAD_GWP_A1_A3 * FOUNDATION
        PAD_GWP_A4_TOTAL = PAD_GWP_A4 * FOUNDATION
        PAD_GWP_A5_TOTAL = PAD_GWP_A5 * FOUNDATION
        PAD_GWP_B4_TOTAL = PAD_GWP_B4 * FOUNDATION
        PAD_GWP_C3_TOTAL = PAD_GWP_C3 * FOUNDATION
        PAD_GWP_C4_TOTAL = PAD_GWP_C4 * FOUNDATION
        PAD_GWP_D_TOTAL = PAD_GWP_D * FOUNDATION
        PAD_GWP_TOTAL = (
            PAD_GWP_A1_A3_TOTAL
            + PAD_GWP_A4_TOTAL
            + PAD_GWP_A5_TOTAL
            + PAD_GWP_B4_TOTAL
            + PAD_GWP_C3_TOTAL
            + PAD_GWP_C4_TOTAL
        )

        LINE_GWP_A1_A3_TOTAL = LINE_GWP_A1_A3 * LENGTH_LINE
        LINE_GWP_A4_TOTAL = LINE_GWP_A4 * LENGTH_LINE
        LINE_GWP_A5_TOTAL = LINE_GWP_A5 * LENGTH_LINE
        LINE_GWP_B4_TOTAL = LINE_GWP_B4 * LENGTH_LINE
        LINE_GWP_C3_TOTAL = LINE_GWP_C3 * LENGTH_LINE
        LINE_GWP_C4_TOTAL = LINE_GWP_C4 * LENGTH_LINE
        LINE_GWP_D_TOTAL = LINE_GWP_D * LENGTH_LINE
        LINE_GWP_TOTAL = (
            LINE_GWP_A1_A3_TOTAL
            + LINE_GWP_A4_TOTAL
            + LINE_GWP_A5_TOTAL
            + LINE_GWP_B4_TOTAL
            + LINE_GWP_C3_TOTAL
            + LINE_GWP_C4_TOTAL
        )

        GROUND_SLAB_GWP_A1_A3_TOTAL = GROUND_SLAB_GWP_A1_A3 * AREA_GROUND_SLAB
        GROUND_SLAB_GWP_A4_TOTAL = GROUND_SLAB_GWP_A4 * AREA_GROUND_SLAB
        GROUND_SLAB_GWP_A5_TOTAL = GROUND_SLAB_GWP_A5 * AREA_GROUND_SLAB
        GROUND_SLAB_GWP_B4_TOTAL = GROUND_SLAB_GWP_B4 * AREA_GROUND_SLAB
        GROUND_SLAB_GWP_C3_TOTAL = GROUND_SLAB_GWP_C3 * AREA_GROUND_SLAB
        GROUND_SLAB_GWP_C4_TOTAL = GROUND_SLAB_GWP_C4 * AREA_GROUND_SLAB
        GROUND_SLAB_GWP_D_TOTAL = GROUND_SLAB_GWP_D * AREA_GROUND_SLAB
        GROUND_SLAB_GWP_TOTAL = (
            GROUND_SLAB_GWP_A1_A3_TOTAL
            + GROUND_SLAB_GWP_A4_TOTAL
            + GROUND_SLAB_GWP_A5_TOTAL
            + GROUND_SLAB_GWP_B4_TOTAL
            + GROUND_SLAB_GWP_C3_TOTAL
            + GROUND_SLAB_GWP_C4_TOTAL
        )

        LOW_SLAB_A_GWP_A1_A3_TOTAL = LOW_SLAB_A_GWP_A1_A3 * AREA_LOW_SETTLEMENT
        LOW_SLAB_A_GWP_A4_TOTAL = LOW_SLAB_A_GWP_A4 * AREA_LOW_SETTLEMENT
        LOW_SLAB_A_GWP_A5_TOTAL = LOW_SLAB_A_GWP_A5 * AREA_LOW_SETTLEMENT
        LOW_SLAB_A_GWP_B4_TOTAL = LOW_SLAB_A_GWP_B4 * AREA_LOW_SETTLEMENT
        LOW_SLAB_A_GWP_C3_TOTAL = LOW_SLAB_A_GWP_C3 * AREA_LOW_SETTLEMENT
        LOW_SLAB_A_GWP_C4_TOTAL = LOW_SLAB_A_GWP_C4 * AREA_LOW_SETTLEMENT
        LOW_SLAB_A_GWP_D_TOTAL = LOW_SLAB_A_GWP_D * AREA_LOW_SETTLEMENT
        LOW_SLAB_A_GWP_TOTAL = (
            LOW_SLAB_A_GWP_A1_A3_TOTAL
            + LOW_SLAB_A_GWP_A4_TOTAL
            + LOW_SLAB_A_GWP_A5_TOTAL
            + LOW_SLAB_A_GWP_B4_TOTAL
            + LOW_SLAB_A_GWP_C3_TOTAL
            + LOW_SLAB_A_GWP_C4_TOTAL
        )

        LOW_SLAB_B_GWP_A1_A3_TOTAL = LOW_SLAB_B_GWP_A1_A3 * AREA_LOW_SETTLEMENT_BELOW
        LOW_SLAB_B_GWP_A4_TOTAL = LOW_SLAB_B_GWP_A4 * AREA_LOW_SETTLEMENT_BELOW
        LOW_SLAB_B_GWP_A5_TOTAL = LOW_SLAB_B_GWP_A5 * AREA_LOW_SETTLEMENT_BELOW
        LOW_SLAB_B_GWP_B4_TOTAL = LOW_SLAB_B_GWP_B4 * AREA_LOW_SETTLEMENT_BELOW
        LOW_SLAB_B_GWP_C3_TOTAL = LOW_SLAB_B_GWP_C3 * AREA_LOW_SETTLEMENT_BELOW
        LOW_SLAB_B_GWP_C4_TOTAL = LOW_SLAB_B_GWP_C4 * AREA_LOW_SETTLEMENT_BELOW
        LOW_SLAB_B_GWP_D_TOTAL = LOW_SLAB_B_GWP_D * AREA_LOW_SETTLEMENT_BELOW
        LOW_SLAB_B_GWP_TOTAL = (
            LOW_SLAB_B_GWP_A1_A3_TOTAL
            + LOW_SLAB_B_GWP_A4_TOTAL
            + LOW_SLAB_B_GWP_A5_TOTAL
            + LOW_SLAB_B_GWP_B4_TOTAL
            + LOW_SLAB_B_GWP_C3_TOTAL
            + LOW_SLAB_B_GWP_C4_TOTAL
        )

        SLAB_ABOVE_GWP_A1_A3_TOTAL = SLAB_ABOVE_GWP_A1_A3 * SLAB
        SLAB_ABOVE_GWP_A4_TOTAL = SLAB_ABOVE_GWP_A4 * SLAB
        SLAB_ABOVE_GWP_A5_TOTAL = SLAB_ABOVE_GWP_A5 * SLAB
        SLAB_ABOVE_GWP_B4_TOTAL = SLAB_ABOVE_GWP_B4 * SLAB
        SLAB_ABOVE_GWP_C3_TOTAL = SLAB_ABOVE_GWP_C3 * SLAB
        SLAB_ABOVE_GWP_C4_TOTAL = SLAB_ABOVE_GWP_C4 * SLAB
        SLAB_ABOVE_GWP_D_TOTAL = SLAB_ABOVE_GWP_D * SLAB
        SLAB_ABOVE_GWP_TOTAL = (
            SLAB_ABOVE_GWP_A1_A3_TOTAL
            + SLAB_ABOVE_GWP_A4_TOTAL
            + SLAB_ABOVE_GWP_A5_TOTAL
            + SLAB_ABOVE_GWP_B4_TOTAL
            + SLAB_ABOVE_GWP_C3_TOTAL
            + SLAB_ABOVE_GWP_C4_TOTAL
        )

        VERTICAL_ABOVE_GWP_A1_A3_TOTAL = VERTICAL_ABOVE_GWP_A1_A3 * VERTICAL
        VERTICAL_ABOVE_GWP_A4_TOTAL = VERTICAL_ABOVE_GWP_A4 * VERTICAL
        VERTICAL_ABOVE_GWP_A5_TOTAL = VERTICAL_ABOVE_GWP_A5 * VERTICAL
        VERTICAL_ABOVE_GWP_B4_TOTAL = VERTICAL_ABOVE_GWP_B4 * VERTICAL
        VERTICAL_ABOVE_GWP_C3_TOTAL = VERTICAL_ABOVE_GWP_C3 * VERTICAL
        VERTICAL_ABOVE_GWP_C4_TOTAL = VERTICAL_ABOVE_GWP_C4 * VERTICAL
        VERTICAL_ABOVE_GWP_D_TOTAL = VERTICAL_ABOVE_GWP_D * VERTICAL
        VERTICAL_ABOVE_GWP_TOTAL = (
            VERTICAL_ABOVE_GWP_A1_A3_TOTAL
            + VERTICAL_ABOVE_GWP_A4_TOTAL
            + VERTICAL_ABOVE_GWP_A5_TOTAL
            + VERTICAL_ABOVE_GWP_B4_TOTAL
            + VERTICAL_ABOVE_GWP_C3_TOTAL
            + VERTICAL_ABOVE_GWP_C4_TOTAL
        )

        BEAM_ABOVE_GWP_A1_A3_TOTAL = BEAM_ABOVE_GWP_A1_A3 * BEAM
        BEAM_ABOVE_GWP_A4_TOTAL = BEAM_ABOVE_GWP_A4 * BEAM
        BEAM_ABOVE_GWP_A5_TOTAL = BEAM_ABOVE_GWP_A5 * BEAM
        BEAM_ABOVE_GWP_B4_TOTAL = BEAM_ABOVE_GWP_B4 * BEAM
        BEAM_ABOVE_GWP_C3_TOTAL = BEAM_ABOVE_GWP_C3 * BEAM
        BEAM_ABOVE_GWP_C4_TOTAL = BEAM_ABOVE_GWP_C4 * BEAM
        BEAM_ABOVE_GWP_D_TOTAL = BEAM_ABOVE_GWP_D * BEAM
        BEAM_ABOVE_GWP_TOTAL = (
            BEAM_ABOVE_GWP_A1_A3_TOTAL
            + BEAM_ABOVE_GWP_A4_TOTAL
            + BEAM_ABOVE_GWP_A5_TOTAL
            + BEAM_ABOVE_GWP_B4_TOTAL
            + BEAM_ABOVE_GWP_C3_TOTAL
            + BEAM_ABOVE_GWP_C4_TOTAL
        )

        SLAB_BELOW_GWP_A1_A3_TOTAL = SLAB_BELOW_GWP_A1_A3 * SLAB_BELOW
        SLAB_BELOW_GWP_A4_TOTAL = SLAB_BELOW_GWP_A4 * SLAB_BELOW
        SLAB_BELOW_GWP_A5_TOTAL = SLAB_BELOW_GWP_A5 * SLAB_BELOW
        SLAB_BELOW_GWP_B4_TOTAL = SLAB_BELOW_GWP_B4 * SLAB_BELOW
        SLAB_BELOW_GWP_C3_TOTAL = SLAB_BELOW_GWP_C3 * SLAB_BELOW
        SLAB_BELOW_GWP_C4_TOTAL = SLAB_BELOW_GWP_C4 * SLAB_BELOW
        SLAB_BELOW_GWP_D_TOTAL = SLAB_BELOW_GWP_D * SLAB_BELOW
        SLAB_BELOW_GWP_TOTAL = (
            SLAB_BELOW_GWP_A1_A3_TOTAL
            + SLAB_BELOW_GWP_A4_TOTAL
            + SLAB_BELOW_GWP_A5_TOTAL
            + SLAB_BELOW_GWP_B4_TOTAL
            + SLAB_BELOW_GWP_C3_TOTAL
            + SLAB_BELOW_GWP_C4_TOTAL
        )

        BEAM_BELOW_GWP_A1_A3_TOTAL = BEAM_BELOW_GWP_A1_A3 * BEAM_BELOW
        BEAM_BELOW_GWP_A4_TOTAL = BEAM_BELOW_GWP_A4 * BEAM_BELOW
        BEAM_BELOW_GWP_A5_TOTAL = BEAM_BELOW_GWP_A5 * BEAM_BELOW
        BEAM_BELOW_GWP_B4_TOTAL = BEAM_BELOW_GWP_B4 * BEAM_BELOW
        BEAM_BELOW_GWP_C3_TOTAL = BEAM_BELOW_GWP_C3 * BEAM_BELOW
        BEAM_BELOW_GWP_C4_TOTAL = BEAM_BELOW_GWP_C4 * BEAM_BELOW
        BEAM_BELOW_GWP_D_TOTAL = BEAM_BELOW_GWP_D * BEAM_BELOW
        BEAM_BELOW_GWP_TOTAL = (
            BEAM_BELOW_GWP_A1_A3_TOTAL
            + BEAM_BELOW_GWP_A4_TOTAL
            + BEAM_BELOW_GWP_A5_TOTAL
            + BEAM_BELOW_GWP_B4_TOTAL
            + BEAM_BELOW_GWP_C3_TOTAL
            + BEAM_BELOW_GWP_C4_TOTAL
        )

        VERTICAL_BELOW_GWP_A1_A3_TOTAL = VERTICAL_BELOW_GWP_A1_A3 * VERTICAL_BELOW
        VERTICAL_BELOW_GWP_A4_TOTAL = VERTICAL_BELOW_GWP_A4 * VERTICAL_BELOW
        VERTICAL_BELOW_GWP_A5_TOTAL = VERTICAL_BELOW_GWP_A5 * VERTICAL_BELOW
        VERTICAL_BELOW_GWP_B4_TOTAL = VERTICAL_BELOW_GWP_B4 * VERTICAL_BELOW
        VERTICAL_BELOW_GWP_C3_TOTAL = VERTICAL_BELOW_GWP_C3 * VERTICAL_BELOW
        VERTICAL_BELOW_GWP_C4_TOTAL = VERTICAL_BELOW_GWP_C4 * VERTICAL_BELOW
        VERTICAL_BELOW_GWP_D_TOTAL = VERTICAL_BELOW_GWP_D * VERTICAL_BELOW
        VERTICAL_BELOW_GWP_TOTAL = (
            VERTICAL_BELOW_GWP_A1_A3_TOTAL
            + VERTICAL_BELOW_GWP_A4_TOTAL
            + VERTICAL_BELOW_GWP_A5_TOTAL
            + VERTICAL_BELOW_GWP_B4_TOTAL
            + VERTICAL_BELOW_GWP_C3_TOTAL
            + VERTICAL_BELOW_GWP_C4_TOTAL
        )

        BASE_WALL_GWP_A1_A3_TOTAL = BASEMENT_WALL_GWP_A1_A3 * BASE_WALL_AREA
        BASE_WALL_GWP_A4_TOTAL = BASEMENT_WALL_GWP_A4 * BASE_WALL_AREA
        BASE_WALL_GWP_A5_TOTAL = BASEMENT_WALL_GWP_A5 * BASE_WALL_AREA
        BASE_WALL_GWP_B4_TOTAL = BASEMENT_WALL_GWP_B4 * BASE_WALL_AREA
        BASE_WALL_GWP_C3_TOTAL = BASEMENT_WALL_GWP_C3 * BASE_WALL_AREA
        BASE_WALL_GWP_C4_TOTAL = BASEMENT_WALL_GWP_C4 * BASE_WALL_AREA
        BASE_WALL_GWP_D_TOTAL = BASEMENT_WALL_GWP_D * BASE_WALL_AREA
        BASE_WALL_GWP_TOTAL = (
            BASE_WALL_GWP_A1_A3_TOTAL
            + BASE_WALL_GWP_A4_TOTAL
            + BASE_WALL_GWP_A5_TOTAL
            + BASE_WALL_GWP_B4_TOTAL
            + BASE_WALL_GWP_C3_TOTAL
            + BASE_WALL_GWP_C4_TOTAL
        )

        PRIM_FACADE_GWP_A1_A3_TOTAL = PRIMARY_FACADE_GWP_A1_A3 * PRIM_FACADE_AREA

        PRIM_FACADE_GWP_A4_TOTAL = PRIMARY_FACADE_GWP_A4 * PRIM_FACADE_AREA
        PRIM_FACADE_GWP_A5_TOTAL = PRIMARY_FACADE_GWP_A5 * PRIM_FACADE_AREA
        PRIM_FACADE_GWP_B4_TOTAL = PRIMARY_FACADE_GWP_B4 * PRIM_FACADE_AREA
        PRIM_FACADE_GWP_C3_TOTAL = PRIMARY_FACADE_GWP_C3 * PRIM_FACADE_AREA
        PRIM_FACADE_GWP_C4_TOTAL = PRIMARY_FACADE_GWP_C4 * PRIM_FACADE_AREA
        PRIM_FACADE_GWP_D_TOTAL = PRIMARY_FACADE_GWP_D * PRIM_FACADE_AREA
        PRIM_FACADE_GWP_TOTAL = (
            PRIM_FACADE_GWP_A1_A3_TOTAL
            + PRIM_FACADE_GWP_A4_TOTAL
            + PRIM_FACADE_GWP_A5_TOTAL
            + PRIM_FACADE_GWP_B4_TOTAL
            + PRIM_FACADE_GWP_C3_TOTAL
            + PRIM_FACADE_GWP_C4_TOTAL
        )

        SEC_FACADE_GWP_A1_A3_TOTAL = SECONDARY_FACADE_GWP_A1_A3 * SEC_FACADE_AREA
        SEC_FACADE_GWP_A4_TOTAL = SECONDARY_FACADE_GWP_A4 * SEC_FACADE_AREA
        SEC_FACADE_GWP_A5_TOTAL = SECONDARY_FACADE_GWP_A5 * SEC_FACADE_AREA
        SEC_FACADE_GWP_B4_TOTAL = SECONDARY_FACADE_GWP_B4 * SEC_FACADE_AREA
        SEC_FACADE_GWP_C3_TOTAL = SECONDARY_FACADE_GWP_C3 * SEC_FACADE_AREA
        SEC_FACADE_GWP_C4_TOTAL = SECONDARY_FACADE_GWP_C4 * SEC_FACADE_AREA
        SEC_FACADE_GWP_D_TOTAL = SECONDARY_FACADE_GWP_D * SEC_FACADE_AREA
        SEC_FACADE_GWP_TOTAL = (
            SEC_FACADE_GWP_A1_A3_TOTAL
            + SEC_FACADE_GWP_A4_TOTAL
            + SEC_FACADE_GWP_A5_TOTAL
            + SEC_FACADE_GWP_B4_TOTAL
            + SEC_FACADE_GWP_C3_TOTAL
            + SEC_FACADE_GWP_C4_TOTAL
        )

        WINDOW_GWP_A1_A3_TOTAL = WINDOW_GWP_A1_A3 * WINDOW_AREA
        WINDOW_GWP_A4_TOTAL = WINDOW_GWP_A4 * WINDOW_AREA
        WINDOW_GWP_A5_TOTAL = WINDOW_GWP_A5 * WINDOW_AREA
        WINDOW_GWP_B4_TOTAL = WINDOW_GWP_B4 * WINDOW_AREA
        WINDOW_GWP_C3_TOTAL = WINDOW_GWP_C3 * WINDOW_AREA
        WINDOW_GWP_C4_TOTAL = WINDOW_GWP_C4 * WINDOW_AREA
        WINDOW_GWP_D_TOTAL = WINDOW_GWP_D * WINDOW_AREA
        WINDOW_GWP_TOTAL = (
            WINDOW_GWP_A1_A3_TOTAL
            + WINDOW_GWP_A4_TOTAL
            + WINDOW_GWP_A5_TOTAL
            + WINDOW_GWP_B4_TOTAL
            + WINDOW_GWP_C3_TOTAL
            + WINDOW_GWP_C4_TOTAL
        )
        # wrong
        CORE_GWP_A1_A3_TOTAL = CORE_WALL_GWP_A1_A3 * CORE_AREA
        CORE_GWP_A4_TOTAL = CORE_WALL_GWP_A4 * CORE_AREA
        CORE_GWP_A5_TOTAL = CORE_WALL_GWP_A5 * CORE_AREA
        CORE_GWP_B4_TOTAL = CORE_WALL_GWP_B4 * CORE_AREA
        CORE_GWP_C3_TOTAL = CORE_WALL_GWP_C3 * CORE_AREA
        CORE_GWP_C4_TOTAL = CORE_WALL_GWP_C4 * CORE_AREA
        CORE_GWP_D_TOTAL = CORE_WALL_GWP_D * CORE_AREA
        CORE_GWP_TOTAL = (
            CORE_GWP_A1_A3_TOTAL
            + CORE_GWP_A4_TOTAL
            + CORE_GWP_A5_TOTAL
            + CORE_GWP_B4_TOTAL
            + CORE_GWP_C3_TOTAL
            + CORE_GWP_C4_TOTAL
        )

        INTERNAL_WALLS_GWP_A1_A3_TOTAL = INTERNAL_WALL_GWP_A1_A3 * AREA_INTERNAL_WALLS
        INTERNAL_WALLS_GWP_A4_TOTAL = INTERNAL_WALL_GWP_A4 * AREA_INTERNAL_WALLS
        INTERNAL_WALLS_GWP_A5_TOTAL = INTERNAL_WALL_GWP_A5 * AREA_INTERNAL_WALLS
        INTERNAL_WALLS_GWP_B4_TOTAL = INTERNAL_WALL_GWP_B4 * AREA_INTERNAL_WALLS
        INTERNAL_WALLS_GWP_C3_TOTAL = INTERNAL_WALL_GWP_C3 * AREA_INTERNAL_WALLS
        INTERNAL_WALLS_GWP_C4_TOTAL = INTERNAL_WALL_GWP_C4 * AREA_INTERNAL_WALLS

        INTERNAL_WALLS_GWP_D_TOTAL = INTERNAL_WALL_GWP_D * AREA_INTERNAL_WALLS
        INTERNAL_WALLS_GWP_TOTAL = (
            INTERNAL_WALLS_GWP_A1_A3_TOTAL
            + INTERNAL_WALLS_GWP_A4_TOTAL
            + INTERNAL_WALLS_GWP_A5_TOTAL
            + INTERNAL_WALLS_GWP_B4_TOTAL
            + INTERNAL_WALLS_GWP_C3_TOTAL
            + INTERNAL_WALLS_GWP_C4_TOTAL
        )

        INTERNAL_WALLS_GROUND_GWP_A1_A3_TOTAL = (
            INTERNAL_WALL_GWP_A1_A3 * AREA_INTERNAL_WALLS_GROUND
        )
        INTERNAL_WALLS_GROUND_GWP_A4_TOTAL = (
            INTERNAL_WALL_GWP_A4 * AREA_INTERNAL_WALLS_GROUND
        )
        INTERNAL_WALLS_GROUND_GWP_A5_TOTAL = (
            INTERNAL_WALL_GWP_A5 * AREA_INTERNAL_WALLS_GROUND
        )
        INTERNAL_WALLS_GROUND_GWP_B4_TOTAL = (
            INTERNAL_WALL_GWP_B4 * AREA_INTERNAL_WALLS_GROUND
        )
        INTERNAL_WALLS_GROUND_GWP_C3_TOTAL = (
            INTERNAL_WALL_GWP_C3 * AREA_INTERNAL_WALLS_GROUND
        )
        INTERNAL_WALLS_GROUND_GWP_C4_TOTAL = (
            INTERNAL_WALL_GWP_C4 * AREA_INTERNAL_WALLS_GROUND
        )
        INTERNAL_WALLS_GROUND_GWP_D_TOTAL = (
            INTERNAL_WALL_GWP_D * AREA_INTERNAL_WALLS_GROUND
        )
        INTERNAL_WALLS_GROUND_GWP_TOTAL = (
            INTERNAL_WALLS_GROUND_GWP_A1_A3_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_A4_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_A5_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_B4_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_C3_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_C4_TOTAL
        )

        INTERNAL_WALLS_BASE_GWP_A1_A3_TOTAL = (
            INTERNAL_BASEMENT_WALL_GWP_A1_A3 * AREA_INTERNAL_WALLS_BASE
        )

        INTERNAL_WALLS_BASE_GWP_A4_TOTAL = (
            INTERNAL_BASEMENT_WALL_GWP_A4 * AREA_INTERNAL_WALLS_BASE
        )
        INTERNAL_WALLS_BASE_GWP_A5_TOTAL = (
            INTERNAL_BASEMENT_WALL_GWP_A5 * AREA_INTERNAL_WALLS_BASE
        )
        INTERNAL_WALLS_BASE_GWP_B4_TOTAL = (
            INTERNAL_BASEMENT_WALL_GWP_B4 * AREA_INTERNAL_WALLS_BASE
        )
        INTERNAL_WALLS_BASE_GWP_C3_TOTAL = (
            INTERNAL_BASEMENT_WALL_GWP_C3 * AREA_INTERNAL_WALLS_BASE
        )
        INTERNAL_WALLS_BASE_GWP_C4_TOTAL = (
            INTERNAL_BASEMENT_WALL_GWP_C4 * AREA_INTERNAL_WALLS_BASE
        )
        INTERNAL_WALLS_BASE_GWP_D_TOTAL = (
            INTERNAL_BASEMENT_WALL_GWP_D * AREA_INTERNAL_WALLS_BASE
        )
        INTERNAL_WALLS_BASE_GWP_TOTAL = (
            INTERNAL_WALLS_BASE_GWP_A1_A3_TOTAL
            + INTERNAL_WALLS_BASE_GWP_A4_TOTAL
            + INTERNAL_WALLS_BASE_GWP_A5_TOTAL
            + INTERNAL_WALLS_BASE_GWP_B4_TOTAL
            + INTERNAL_WALLS_BASE_GWP_C3_TOTAL
            + INTERNAL_WALLS_BASE_GWP_C4_TOTAL
        )

        ROOF_FINISH_GWP_A1_A3_TOTAL = ROOF_FINISH_GWP_A1_A3 * AREA_ROOF_FINISH
        ROOF_FINISH_GWP_A4_TOTAL = ROOF_FINISH_GWP_A4 * AREA_ROOF_FINISH
        ROOF_FINISH_GWP_A5_TOTAL = ROOF_FINISH_GWP_A5 * AREA_ROOF_FINISH
        ROOF_FINISH_GWP_B4_TOTAL = ROOF_FINISH_GWP_B4 * AREA_ROOF_FINISH
        ROOF_FINISH_GWP_C3_TOTAL = ROOF_FINISH_GWP_C3 * AREA_ROOF_FINISH
        ROOF_FINISH_GWP_C4_TOTAL = ROOF_FINISH_GWP_C4 * AREA_ROOF_FINISH
        ROOF_FINISH_GWP_D_TOTAL = ROOF_FINISH_GWP_D * AREA_ROOF_FINISH
        ROOF_FINISH_GWP_TOTAL = (
            ROOF_FINISH_GWP_A1_A3_TOTAL
            + ROOF_FINISH_GWP_A4_TOTAL
            + ROOF_FINISH_GWP_A5_TOTAL
            + ROOF_FINISH_GWP_B4_TOTAL
            + ROOF_FINISH_GWP_C3_TOTAL
            + ROOF_FINISH_GWP_C4_TOTAL
        )

        FLOOR_FINISH_ABOVE_GWP_A1_A3_TOTAL = (
            FLOOR_FINISH_ABOVE_GWP_A1_A3 * AREA_FLOOR_FINISH
        )
        FLOOR_FINISH_ABOVE_GWP_A4_TOTAL = FLOOR_FINISH_ABOVE_GWP_A4 * AREA_FLOOR_FINISH
        FLOOR_FINISH_ABOVE_GWP_A5_TOTAL = FLOOR_FINISH_ABOVE_GWP_A5 * AREA_FLOOR_FINISH
        FLOOR_FINISH_ABOVE_GWP_B4_TOTAL = FLOOR_FINISH_ABOVE_GWP_B4 * AREA_FLOOR_FINISH
        FLOOR_FINISH_ABOVE_GWP_C3_TOTAL = FLOOR_FINISH_ABOVE_GWP_C3 * AREA_FLOOR_FINISH
        FLOOR_FINISH_ABOVE_GWP_C4_TOTAL = FLOOR_FINISH_ABOVE_GWP_C4 * AREA_FLOOR_FINISH
        FLOOR_FINISH_ABOVE_GWP_D_TOTAL = FLOOR_FINISH_ABOVE_GWP_D * AREA_FLOOR_FINISH
        FLOOR_FINISH_ABOVE_GWP_TOTAL = (
            FLOOR_FINISH_ABOVE_GWP_A1_A3_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_A4_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_A5_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_B4_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_C3_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_C4_TOTAL
        )

        FLOOR_FINISH_GROUND_GWP_A1_A3_TOTAL = (
            FLOOR_FINISH_GROUND_GWP_A1_A3 * AREA_FLOOR_FINISH_GROUND
        )
        FLOOR_FINISH_GROUND_GWP_A4_TOTAL = (
            FLOOR_FINISH_GROUND_GWP_A4 * AREA_FLOOR_FINISH_GROUND
        )
        FLOOR_FINISH_GROUND_GWP_A5_TOTAL = (
            FLOOR_FINISH_GROUND_GWP_A5 * AREA_FLOOR_FINISH_GROUND
        )
        FLOOR_FINISH_GROUND_GWP_B4_TOTAL = (
            FLOOR_FINISH_GROUND_GWP_B4 * AREA_FLOOR_FINISH_GROUND
        )
        FLOOR_FINISH_GROUND_GWP_C3_TOTAL = (
            FLOOR_FINISH_GROUND_GWP_C3 * AREA_FLOOR_FINISH_GROUND
        )
        FLOOR_FINISH_GROUND_GWP_C4_TOTAL = (
            FLOOR_FINISH_GROUND_GWP_C4 * AREA_FLOOR_FINISH_GROUND
        )
        FLOOR_FINISH_GROUND_GWP_D_TOTAL = (
            FLOOR_FINISH_GROUND_GWP_D * AREA_FLOOR_FINISH_GROUND
        )
        FLOOR_FINISH_GROUND_GWP_TOTAL = (
            FLOOR_FINISH_GROUND_GWP_A1_A3_TOTAL
            + FLOOR_FINISH_GROUND_GWP_A4_TOTAL
            + FLOOR_FINISH_GROUND_GWP_A5_TOTAL
            + FLOOR_FINISH_GROUND_GWP_B4_TOTAL
            + FLOOR_FINISH_GROUND_GWP_C3_TOTAL
            + FLOOR_FINISH_GROUND_GWP_C4_TOTAL
        )

        FLOOR_FINISH_BELOW_GWP_A1_A3_TOTAL = (
            FLOOR_FINISH_BELOW_GWP_A1_A3 * AREA_FLOOR_FINISH_BELOW
        )
        FLOOR_FINISH_BELOW_GWP_A4_TOTAL = (
            FLOOR_FINISH_BELOW_GWP_A4 * AREA_FLOOR_FINISH_BELOW
        )
        FLOOR_FINISH_BELOW_GWP_A5_TOTAL = (
            FLOOR_FINISH_BELOW_GWP_A5 * AREA_FLOOR_FINISH_BELOW
        )
        FLOOR_FINISH_BELOW_GWP_B4_TOTAL = (
            FLOOR_FINISH_BELOW_GWP_B4 * AREA_FLOOR_FINISH_BELOW
        )
        FLOOR_FINISH_BELOW_GWP_C3_TOTAL = (
            FLOOR_FINISH_BELOW_GWP_C3 * AREA_FLOOR_FINISH_BELOW
        )
        FLOOR_FINISH_BELOW_GWP_C4_TOTAL = (
            FLOOR_FINISH_BELOW_GWP_C4 * AREA_FLOOR_FINISH_BELOW
        )
        FLOOR_FINISH_BELOW_GWP_D_TOTAL = (
            FLOOR_FINISH_BELOW_GWP_D * AREA_FLOOR_FINISH_BELOW
        )
        FLOOR_FINISH_BELOW_GWP_TOTAL = (
            FLOOR_FINISH_BELOW_GWP_A1_A3_TOTAL
            + FLOOR_FINISH_BELOW_GWP_A4_TOTAL
            + FLOOR_FINISH_BELOW_GWP_A5_TOTAL
            + FLOOR_FINISH_BELOW_GWP_B4_TOTAL
            + FLOOR_FINISH_BELOW_GWP_C3_TOTAL
            + FLOOR_FINISH_BELOW_GWP_C4_TOTAL
        )

        CEILING_FINISH_ABOVE_GWP_A1_A3_TOTAL = (
            CEILING_FINISH_ABOVE_GWP_A1_A3 * AREA_CEILING_FINISH
        )
        CEILING_FINISH_ABOVE_GWP_A4_TOTAL = (
            CEILING_FINISH_ABOVE_GWP_A4 * AREA_CEILING_FINISH
        )
        CEILING_FINISH_ABOVE_GWP_A5_TOTAL = (
            CEILING_FINISH_ABOVE_GWP_A5 * AREA_CEILING_FINISH
        )
        CEILING_FINISH_ABOVE_GWP_B4_TOTAL = (
            CEILING_FINISH_ABOVE_GWP_B4 * AREA_CEILING_FINISH
        )
        CEILING_FINISH_ABOVE_GWP_C3_TOTAL = (
            CEILING_FINISH_ABOVE_GWP_C3 * AREA_CEILING_FINISH
        )
        CEILING_FINISH_ABOVE_GWP_C4_TOTAL = (
            CEILING_FINISH_ABOVE_GWP_C4 * AREA_CEILING_FINISH
        )
        CEILING_FINISH_ABOVE_GWP_D_TOTAL = (
            CEILING_FINISH_ABOVE_GWP_D * AREA_CEILING_FINISH
        )
        CEILING_FINISH_ABOVE_GWP_TOTAL = (
            CEILING_FINISH_ABOVE_GWP_A1_A3_TOTAL
            + CEILING_FINISH_ABOVE_GWP_A4_TOTAL
            + CEILING_FINISH_ABOVE_GWP_A5_TOTAL
            + CEILING_FINISH_ABOVE_GWP_B4_TOTAL
            + CEILING_FINISH_ABOVE_GWP_C3_TOTAL
            + CEILING_FINISH_ABOVE_GWP_C4_TOTAL
        )

        CEILING_FINISH_GROUND_GWP_A1_A3_TOTAL = (
            CEILING_FINISH_GROUND_GWP_A1_A3 * AREA_CEILING_FINISH_GROUND
        )
        CEILING_FINISH_GROUND_GWP_A4_TOTAL = (
            CEILING_FINISH_GROUND_GWP_A4 * AREA_CEILING_FINISH_GROUND
        )
        CEILING_FINISH_GROUND_GWP_A5_TOTAL = (
            CEILING_FINISH_GROUND_GWP_A5 * AREA_CEILING_FINISH_GROUND
        )
        CEILING_FINISH_GROUND_GWP_B4_TOTAL = (
            CEILING_FINISH_GROUND_GWP_B4 * AREA_CEILING_FINISH_GROUND
        )
        CEILING_FINISH_GROUND_GWP_C3_TOTAL = (
            CEILING_FINISH_GROUND_GWP_C3 * AREA_CEILING_FINISH_GROUND
        )
        CEILING_FINISH_GROUND_GWP_C4_TOTAL = (
            CEILING_FINISH_GROUND_GWP_C4 * AREA_CEILING_FINISH_GROUND
        )
        CEILING_FINISH_GROUND_GWP_D_TOTAL = (
            CEILING_FINISH_GROUND_GWP_D * AREA_CEILING_FINISH_GROUND
        )
        CEILING_FINISH_GROUND_GWP_TOTAL = (
            CEILING_FINISH_GROUND_GWP_A1_A3_TOTAL
            + CEILING_FINISH_GROUND_GWP_A4_TOTAL
            + CEILING_FINISH_GROUND_GWP_A5_TOTAL
            + CEILING_FINISH_GROUND_GWP_B4_TOTAL
            + CEILING_FINISH_GROUND_GWP_C3_TOTAL
            + CEILING_FINISH_GROUND_GWP_C4_TOTAL
        )

        CEILING_FINISH_BELOW_GWP_A1_A3_TOTAL = (
            CEILING_FINISH_BELOW_GWP_A1_A3 * AREA_CEILING_FINISH_BELOW
        )
        CEILING_FINISH_BELOW_GWP_A4_TOTAL = (
            CEILING_FINISH_BELOW_GWP_A4 * AREA_CEILING_FINISH_BELOW
        )
        CEILING_FINISH_BELOW_GWP_A5_TOTAL = (
            CEILING_FINISH_BELOW_GWP_A5 * AREA_CEILING_FINISH_BELOW
        )
        CEILING_FINISH_BELOW_GWP_B4_TOTAL = (
            CEILING_FINISH_BELOW_GWP_B4 * AREA_CEILING_FINISH_BELOW
        )
        CEILING_FINISH_BELOW_GWP_C3_TOTAL = (
            CEILING_FINISH_BELOW_GWP_C3 * AREA_CEILING_FINISH_BELOW
        )
        CEILING_FINISH_BELOW_GWP_C4_TOTAL = (
            CEILING_FINISH_BELOW_GWP_C4 * AREA_CEILING_FINISH_BELOW
        )
        CEILING_FINISH_BELOW_GWP_D_TOTAL = (
            CEILING_FINISH_BELOW_GWP_D * AREA_CEILING_FINISH_BELOW
        )
        CEILING_FINISH_BELOW_GWP_TOTAL = (
            CEILING_FINISH_BELOW_GWP_A1_A3_TOTAL
            + CEILING_FINISH_BELOW_GWP_A4_TOTAL
            + CEILING_FINISH_BELOW_GWP_A5_TOTAL
            + CEILING_FINISH_BELOW_GWP_B4_TOTAL
            + CEILING_FINISH_BELOW_GWP_C3_TOTAL
            + CEILING_FINISH_BELOW_GWP_C4_TOTAL
        )

        STAIRCASE_GWP_A1_A3_TOTAL = STAIRCASE_GWP_A1_A3 * LENGTH_STAIRCASE
        STAIRCASE_GWP_A4_TOTAL = STAIRCASE_GWP_A4 * LENGTH_STAIRCASE
        STAIRCASE_GWP_A5_TOTAL = STAIRCASE_GWP_A5 * LENGTH_STAIRCASE
        STAIRCASE_GWP_B4_TOTAL = STAIRCASE_GWP_B4 * LENGTH_STAIRCASE
        STAIRCASE_GWP_C3_TOTAL = STAIRCASE_GWP_C3 * LENGTH_STAIRCASE
        STAIRCASE_GWP_C4_TOTAL = STAIRCASE_GWP_C4 * LENGTH_STAIRCASE
        STAIRCASE_GWP_D_TOTAL = STAIRCASE_GWP_D * LENGTH_STAIRCASE
        STAIRCASE_GWP_TOTAL = (
            STAIRCASE_GWP_A1_A3_TOTAL
            + STAIRCASE_GWP_A4_TOTAL
            + STAIRCASE_GWP_A5_TOTAL
            + STAIRCASE_GWP_B4_TOTAL
            + STAIRCASE_GWP_C3_TOTAL
            + STAIRCASE_GWP_C4_TOTAL
        )

        ELEVATOR_GWP_A1_A3_TOTAL = ELEVATOR_GWP_A1_A3 * NUM_ELEVATORS
        ELEVATOR_GWP_A4_TOTAL = ELEVATOR_GWP_A4 * NUM_ELEVATORS
        ELEVATOR_GWP_A5_TOTAL = ELEVATOR_GWP_A5 * NUM_ELEVATORS
        ELEVATOR_GWP_B4_TOTAL = ELEVATOR_GWP_B4 * NUM_ELEVATORS
        ELEVATOR_GWP_C3_TOTAL = ELEVATOR_GWP_C3 * NUM_ELEVATORS
        ELEVATOR_GWP_C4_TOTAL = ELEVATOR_GWP_C4 * NUM_ELEVATORS
        ELEVATOR_GWP_D_TOTAL = ELEVATOR_GWP_D * NUM_ELEVATORS
        ELEVATOR_GWP_TOTAL = (
            ELEVATOR_GWP_A1_A3_TOTAL
            + ELEVATOR_GWP_A4_TOTAL
            + ELEVATOR_GWP_A5_TOTAL
            + ELEVATOR_GWP_B4_TOTAL
            + ELEVATOR_GWP_C3_TOTAL
            + ELEVATOR_GWP_C4_TOTAL
        )

        TECHNICAL_ABOVE_GWP_A1_A3_TOTAL = TECHNICAL_ABOVE_GWP_A1_A3 * AREA_TECHNICAL
        TECHNICAL_ABOVE_GWP_A4_TOTAL = TECHNICAL_ABOVE_GWP_A4 * AREA_TECHNICAL
        TECHNICAL_ABOVE_GWP_A5_TOTAL = TECHNICAL_ABOVE_GWP_A5 * AREA_TECHNICAL
        TECHNICAL_ABOVE_GWP_B4_TOTAL = TECHNICAL_ABOVE_GWP_B4 * AREA_TECHNICAL
        TECHNICAL_ABOVE_GWP_C3_TOTAL = TECHNICAL_ABOVE_GWP_C3 * AREA_TECHNICAL
        TECHNICAL_ABOVE_GWP_C4_TOTAL = TECHNICAL_ABOVE_GWP_C4 * AREA_TECHNICAL
        TECHNICAL_ABOVE_GWP_D_TOTAL = TECHNICAL_ABOVE_GWP_D * AREA_TECHNICAL
        TECHNICAL_ABOVE_GWP_TOTAL = (
            TECHNICAL_ABOVE_GWP_A1_A3_TOTAL
            + TECHNICAL_ABOVE_GWP_A4_TOTAL
            + TECHNICAL_ABOVE_GWP_A5_TOTAL
            + TECHNICAL_ABOVE_GWP_B4_TOTAL
            + TECHNICAL_ABOVE_GWP_C3_TOTAL
            + TECHNICAL_ABOVE_GWP_C4_TOTAL
        )

        TECHNICAL_GROUND_GWP_A1_A3_TOTAL = (
            TECHNICAL_GROUND_GWP_A1_A3 * AREA_TECHNICAL_GROUND
        )
        TECHNICAL_GROUND_GWP_A4_TOTAL = TECHNICAL_GROUND_GWP_A4 * AREA_TECHNICAL_GROUND
        TECHNICAL_GROUND_GWP_A5_TOTAL = TECHNICAL_GROUND_GWP_A5 * AREA_TECHNICAL_GROUND
        TECHNICAL_GROUND_GWP_B4_TOTAL = TECHNICAL_GROUND_GWP_B4 * AREA_TECHNICAL_GROUND
        TECHNICAL_GROUND_GWP_C3_TOTAL = TECHNICAL_GROUND_GWP_C3 * AREA_TECHNICAL_GROUND
        TECHNICAL_GROUND_GWP_C4_TOTAL = TECHNICAL_GROUND_GWP_C4 * AREA_TECHNICAL_GROUND
        TECHNICAL_GROUND_GWP_D_TOTAL = TECHNICAL_GROUND_GWP_D * AREA_TECHNICAL_GROUND
        TECHNICAL_GROUND_GWP_TOTAL = (
            TECHNICAL_GROUND_GWP_A1_A3_TOTAL
            + TECHNICAL_GROUND_GWP_A4_TOTAL
            + TECHNICAL_GROUND_GWP_A5_TOTAL
            + TECHNICAL_GROUND_GWP_B4_TOTAL
            + TECHNICAL_GROUND_GWP_C3_TOTAL
            + TECHNICAL_GROUND_GWP_C4_TOTAL
        )

        TECHNICAL_BELOW_GWP_A1_A3_TOTAL = (
            TECHNICAL_BELOW_GWP_A1_A3 * AREA_TECHNICAL_BELOW
        )
        TECHNICAL_BELOW_GWP_A4_TOTAL = TECHNICAL_BELOW_GWP_A4 * AREA_TECHNICAL_BELOW
        TECHNICAL_BELOW_GWP_A5_TOTAL = TECHNICAL_BELOW_GWP_A5 * AREA_TECHNICAL_BELOW
        TECHNICAL_BELOW_GWP_B4_TOTAL = TECHNICAL_BELOW_GWP_B4 * AREA_TECHNICAL_BELOW
        TECHNICAL_BELOW_GWP_C3_TOTAL = TECHNICAL_BELOW_GWP_C3 * AREA_TECHNICAL_BELOW
        TECHNICAL_BELOW_GWP_C4_TOTAL = TECHNICAL_BELOW_GWP_C4 * AREA_TECHNICAL_BELOW
        TECHNICAL_BELOW_GWP_D_TOTAL = TECHNICAL_BELOW_GWP_D * AREA_TECHNICAL_BELOW
        TECHNICAL_BELOW_GWP_TOTAL = (
            TECHNICAL_BELOW_GWP_A1_A3_TOTAL
            + TECHNICAL_BELOW_GWP_A4_TOTAL
            + TECHNICAL_BELOW_GWP_A5_TOTAL
            + TECHNICAL_BELOW_GWP_B4_TOTAL
            + TECHNICAL_BELOW_GWP_C3_TOTAL
            + TECHNICAL_BELOW_GWP_C4_TOTAL
        )
        # endregion

        # region Emissions of construction and demolition
        DIESEL_USE = c.EMISSION_FACTOR_DIESEL * c.DIESEL_USE
        # removed if statement and added lookup
        ELECTRICITY_USE = (
            c.emission_factor_lookup_table[emission_factor][construction_year][0]
            * c.ELECTRICITY_USE
        )

        CONSTRUCTION_GWP_TOTAL = AREA_CONSTRUCTION * (DIESEL_USE + ELECTRICITY_USE)

        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            DEMOLITION_GWP_C1 = c.DEMOLITION_DICT[structure] * AREA_DEMOLITION
        else:
            DEMOLITION_GWP_C1 = 0

        # endregion

        # region Emissions Operational energy // DriftEnergi

        ENERGY_USE_ABOVE = c.ENERGY_USE_DICT[typology][condition] * study_period
        HEATING_USE_ABOVE = c.HEATING_DICT[typology][condition] * study_period

        try:
            ENERGY_USE_BELOW = (
                c.ENERGY_USE_DICT[TYPOLOGY_BASEMENT][condition] * study_period
            )
        except:
            ENERGY_USE_BELOW = 0

        try:
            HEAT_USE_BELOW = c.HEATING_DICT[TYPOLOGY_BASEMENT][condition] * study_period
        except:
            HEAT_USE_BELOW = 0

        ENERGY_USE_GROUND = c.ENERGY_USE_DICT[ground_typology][condition] * study_period
        HEATING_USE_GROUND = c.HEATING_DICT[ground_typology][condition] * study_period

        AV_FACTOR_ELECTRICITY = range_average(
            c.emission_factor_lookup_table,
            emission_factor,
            construction_year,
            study_period,
            "electricity",
        )

        AV_FACTOR_DISTRICT = range_average(
            c.emission_factor_lookup_table,
            emission_factor,
            construction_year,
            study_period,
            "district heating",
        )

        # NEW FACTOR 3/3.5 added

        if heating == "district heating":
            HEATING_USE_ABOVE = HEATING_USE_ABOVE / 3
            HEATING_USE_GROUND = HEATING_USE_GROUND / 3
            HEAT_USE_BELOW = HEAT_USE_BELOW / 3
        else:
            HEATING_USE_ABOVE = HEATING_USE_ABOVE / 3.5
            HEATING_USE_GROUND = HEATING_USE_GROUND / 3.5
            HEAT_USE_BELOW = HEAT_USE_BELOW / 3.5

        EMISSIONS_ENERGY_USE_ABOVE = ENERGY_USE_ABOVE * AV_FACTOR_ELECTRICITY
        EMISSIONS_ENERGY_USE_GROUND = ENERGY_USE_GROUND * AV_FACTOR_ELECTRICITY
        EMISSIONS_ENERGY_USE_BELOW = ENERGY_USE_BELOW * AV_FACTOR_ELECTRICITY

        if heating == "district heating":
            EMISSIONS_HEATING_ABOVE = HEATING_USE_ABOVE * AV_FACTOR_DISTRICT
            EMISSIONS_HEATING_GROUND = HEATING_USE_GROUND * AV_FACTOR_DISTRICT
            EMISSIONS_HEATING_BELOW = HEAT_USE_BELOW * AV_FACTOR_DISTRICT
        else:
            EMISSIONS_HEATING_ABOVE = HEATING_USE_ABOVE * AV_FACTOR_ELECTRICITY
            EMISSIONS_HEATING_GROUND = HEATING_USE_GROUND * AV_FACTOR_ELECTRICITY
            EMISSIONS_HEATING_BELOW = HEAT_USE_BELOW * AV_FACTOR_ELECTRICITY

        EMISSIONS_OPERATIONAL_ABOVE = (
            EMISSIONS_ENERGY_USE_ABOVE + EMISSIONS_HEATING_ABOVE
        ) * AREA_OPERATIONAL
        EMISSIONS_OPERATIONAL_GROUND = (
            EMISSIONS_ENERGY_USE_GROUND + EMISSIONS_HEATING_GROUND
        ) * AREA_OPERATIONAL_GROUND
        EMISSIONS_OPERATIONAL_BELOW = (
            EMISSIONS_ENERGY_USE_BELOW + EMISSIONS_HEATING_BELOW
        ) * AREA_OPERATIONAL_BELOW
        EMISSIONS_OPERATIONAL = (
            EMISSIONS_OPERATIONAL_ABOVE
            + EMISSIONS_OPERATIONAL_GROUND
            + EMISSIONS_OPERATIONAL_BELOW
        )

        self.EMISSIONS_OPERATIONAL = EMISSIONS_OPERATIONAL

        # endregion

        # endregion

        # region Results
        EMISSIONS_FOUNDATION_GWP_A1_A3 = (
            PAD_GWP_A1_A3_TOTAL
            + LINE_GWP_A1_A3_TOTAL
            + GROUND_SLAB_GWP_A1_A3_TOTAL
            + SLAB_BELOW_GWP_A1_A3_TOTAL
            + LOW_SLAB_B_GWP_A1_A3_TOTAL
            + VERTICAL_BELOW_GWP_A1_A3_TOTAL
            + BEAM_BELOW_GWP_A1_A3_TOTAL
            + BASE_WALL_GWP_A1_A3_TOTAL
        )

        EMISSIONS_FOUNDATION_GWP_A4 = (
            PAD_GWP_A4_TOTAL
            + LINE_GWP_A4_TOTAL
            + GROUND_SLAB_GWP_A4_TOTAL
            + SLAB_BELOW_GWP_A4_TOTAL
            + LOW_SLAB_B_GWP_A4_TOTAL
            + VERTICAL_BELOW_GWP_A4_TOTAL
            + BEAM_BELOW_GWP_A4_TOTAL
            + BASE_WALL_GWP_A4_TOTAL
        )
        EMISSIONS_FOUNDATION_GWP_A5 = (
            PAD_GWP_A5_TOTAL
            + LINE_GWP_A5_TOTAL
            + GROUND_SLAB_GWP_A5_TOTAL
            + SLAB_BELOW_GWP_A5_TOTAL
            + LOW_SLAB_B_GWP_A5_TOTAL
            + VERTICAL_BELOW_GWP_A5_TOTAL
            + BEAM_BELOW_GWP_A5_TOTAL
            + BASE_WALL_GWP_A5_TOTAL
        )
        EMISSIONS_FOUNDATION_GWP_B4 = (
            PAD_GWP_B4_TOTAL
            + LINE_GWP_B4_TOTAL
            + GROUND_SLAB_GWP_B4_TOTAL
            + SLAB_BELOW_GWP_B4_TOTAL
            + LOW_SLAB_B_GWP_B4_TOTAL
            + VERTICAL_BELOW_GWP_B4_TOTAL
            + BEAM_BELOW_GWP_B4_TOTAL
            + BASE_WALL_GWP_B4_TOTAL
        )
        EMISSIONS_FOUNDATION_GWP_C3 = (
            PAD_GWP_C3_TOTAL
            + LINE_GWP_C3_TOTAL
            + GROUND_SLAB_GWP_C3_TOTAL
            + SLAB_BELOW_GWP_C3_TOTAL
            + LOW_SLAB_B_GWP_C3_TOTAL
            + VERTICAL_BELOW_GWP_C3_TOTAL
            + BEAM_BELOW_GWP_C3_TOTAL
            + BASE_WALL_GWP_C3_TOTAL
        )
        EMISSIONS_FOUNDATION_GWP_C4 = (
            PAD_GWP_C4_TOTAL
            + LINE_GWP_C4_TOTAL
            + GROUND_SLAB_GWP_C4_TOTAL
            + SLAB_BELOW_GWP_C4_TOTAL
            + LOW_SLAB_B_GWP_C4_TOTAL
            + VERTICAL_BELOW_GWP_C4_TOTAL
            + BEAM_BELOW_GWP_C4_TOTAL
            + BASE_WALL_GWP_C4_TOTAL
        )
        if condition == "demolition":
            (
                EMISSIONS_FOUNDATION_GWP_A1_A3,
                EMISSIONS_FOUNDATION_GWP_A4,
                EMISSIONS_FOUNDATION_GWP_A5,
                EMISSIONS_FOUNDATION_GWP_B4,
                EMISSIONS_FOUNDATION_GWP_C3,
                EMISSIONS_FOUNDATION_GWP_C4,
            ) = (0, 0, 0, 0, 0, 0)
        self.EMISSIONS_FOUNDATION_GWP_TOTAL = (
            EMISSIONS_FOUNDATION_GWP_A1_A3
            + EMISSIONS_FOUNDATION_GWP_A4
            + EMISSIONS_FOUNDATION_GWP_A5
            + EMISSIONS_FOUNDATION_GWP_B4
            + EMISSIONS_FOUNDATION_GWP_C3
            + EMISSIONS_FOUNDATION_GWP_C4
        )
        EMISSIONS_STRUCTURE_GWP_A1_A3 = (
            SLAB_ABOVE_GWP_A1_A3_TOTAL
            + LOW_SLAB_A_GWP_A1_A3_TOTAL
            + VERTICAL_ABOVE_GWP_A1_A3_TOTAL
            + BEAM_ABOVE_GWP_A1_A3_TOTAL
            + CORE_GWP_A1_A3_TOTAL
        )

        EMISSIONS_STRUCTURE_GWP_A4 = (
            SLAB_ABOVE_GWP_A4_TOTAL
            + LOW_SLAB_A_GWP_A4_TOTAL
            + VERTICAL_ABOVE_GWP_A4_TOTAL
            + BEAM_ABOVE_GWP_A4_TOTAL
            + CORE_GWP_A4_TOTAL
        )
        EMISSIONS_STRUCTURE_GWP_A5 = (
            SLAB_ABOVE_GWP_A5_TOTAL
            + LOW_SLAB_A_GWP_A5_TOTAL
            + VERTICAL_ABOVE_GWP_A5_TOTAL
            + BEAM_ABOVE_GWP_A5_TOTAL
            + CORE_GWP_A5_TOTAL
        )
        EMISSIONS_STRUCTURE_GWP_B4 = (
            SLAB_ABOVE_GWP_B4_TOTAL
            + LOW_SLAB_A_GWP_B4_TOTAL
            + VERTICAL_ABOVE_GWP_B4_TOTAL
            + BEAM_ABOVE_GWP_B4_TOTAL
            + CORE_GWP_B4_TOTAL
        )
        EMISSIONS_STRUCTURE_GWP_C3 = (
            SLAB_ABOVE_GWP_C3_TOTAL
            + LOW_SLAB_A_GWP_C3_TOTAL
            + VERTICAL_ABOVE_GWP_C3_TOTAL
            + BEAM_ABOVE_GWP_C3_TOTAL
            + CORE_GWP_C3_TOTAL
        )
        EMISSIONS_STRUCTURE_GWP_C4 = (
            SLAB_ABOVE_GWP_C4_TOTAL
            + LOW_SLAB_A_GWP_C4_TOTAL
            + VERTICAL_ABOVE_GWP_C4_TOTAL
            + BEAM_ABOVE_GWP_C4_TOTAL
            + CORE_GWP_C4_TOTAL
        )
        if condition == "demolition":
            (
                EMISSIONS_STRUCTURE_GWP_A1_A3,
                EMISSIONS_STRUCTURE_GWP_A4,
                EMISSIONS_STRUCTURE_GWP_A5,
                EMISSIONS_STRUCTURE_GWP_B4,
                EMISSIONS_STRUCTURE_GWP_C3,
                EMISSIONS_STRUCTURE_GWP_C4,
            ) = (0, 0, 0, 0, 0, 0)
        self.EMISSIONS_STRUCTURE_GWP_TOTAL = (
            EMISSIONS_STRUCTURE_GWP_A1_A3
            + EMISSIONS_STRUCTURE_GWP_A4
            + EMISSIONS_STRUCTURE_GWP_A5
            + EMISSIONS_STRUCTURE_GWP_B4
            + EMISSIONS_STRUCTURE_GWP_C3
            + EMISSIONS_STRUCTURE_GWP_C4
        )

        # Print this
        EMISSIONS_ENVELOPE_GWP_A1_A3 = (
            PRIM_FACADE_GWP_A1_A3_TOTAL
            + SEC_FACADE_GWP_A1_A3_TOTAL
            + ROOF_FINISH_GWP_A1_A3_TOTAL
        )

        EMISSIONS_ENVELOPE_GWP_A4 = (
            PRIM_FACADE_GWP_A4_TOTAL
            + SEC_FACADE_GWP_A4_TOTAL
            + ROOF_FINISH_GWP_A4_TOTAL
        )
        EMISSIONS_ENVELOPE_GWP_A5 = (
            PRIM_FACADE_GWP_A5_TOTAL
            + SEC_FACADE_GWP_A5_TOTAL
            + ROOF_FINISH_GWP_A5_TOTAL
        )
        EMISSIONS_ENVELOPE_GWP_B4 = (
            PRIM_FACADE_GWP_B4_TOTAL
            + SEC_FACADE_GWP_B4_TOTAL
            + ROOF_FINISH_GWP_B4_TOTAL
        )
        EMISSIONS_ENVELOPE_GWP_C3 = (
            PRIM_FACADE_GWP_C3_TOTAL
            + SEC_FACADE_GWP_C3_TOTAL
            + ROOF_FINISH_GWP_C3_TOTAL
        )
        EMISSIONS_ENVELOPE_GWP_C4 = (
            PRIM_FACADE_GWP_C4_TOTAL
            + SEC_FACADE_GWP_C4_TOTAL
            + ROOF_FINISH_GWP_C4_TOTAL
        )
        if condition == "demolition":
            (
                EMISSIONS_ENVELOPE_GWP_A1_A3,
                EMISSIONS_ENVELOPE_GWP_A4,
                EMISSIONS_ENVELOPE_GWP_A5,
                EMISSIONS_ENVELOPE_GWP_B4,
                EMISSIONS_ENVELOPE_GWP_C3,
                EMISSIONS_ENVELOPE_GWP_C4,
            ) = (0, 0, 0, 0, 0, 0)
        EMISSIONS_ENVELOPE_GWP_TOTAL = (
            EMISSIONS_ENVELOPE_GWP_A1_A3
            + EMISSIONS_ENVELOPE_GWP_A4
            + EMISSIONS_ENVELOPE_GWP_A5
            + EMISSIONS_ENVELOPE_GWP_B4
            + EMISSIONS_ENVELOPE_GWP_C3
            + EMISSIONS_ENVELOPE_GWP_C4
        )

        EMISSIONS_WINDOW_GWP_A1_A3 = WINDOW_GWP_A1_A3_TOTAL
        EMISSIONS_WINDOW_GWP_A4 = WINDOW_GWP_A4_TOTAL
        EMISSIONS_WINDOW_GWP_A5 = WINDOW_GWP_A5_TOTAL
        EMISSIONS_WINDOW_GWP_B4 = WINDOW_GWP_B4_TOTAL
        EMISSIONS_WINDOW_GWP_C3 = WINDOW_GWP_C3_TOTAL
        EMISSIONS_WINDOW_GWP_C4 = WINDOW_GWP_C4_TOTAL
        if condition == "demolition":
            (
                EMISSIONS_WINDOW_GWP_A1_A3,
                EMISSIONS_WINDOW_GWP_A4,
                EMISSIONS_WINDOW_GWP_A5,
                EMISSIONS_WINDOW_GWP_B4,
                EMISSIONS_WINDOW_GWP_C3,
                EMISSIONS_WINDOW_GWP_C4,
            ) = (0, 0, 0, 0, 0, 0)
        EMISSIONS_WINDOW_GWP_TOTAL = (
            EMISSIONS_WINDOW_GWP_A1_A3
            + EMISSIONS_WINDOW_GWP_A4
            + EMISSIONS_WINDOW_GWP_A5
            + EMISSIONS_WINDOW_GWP_B4
            + EMISSIONS_WINDOW_GWP_C3
            + EMISSIONS_WINDOW_GWP_C4
        )

        EMISSIONS_INTERIOR_GWP_A1_A3 = (
            INTERNAL_WALLS_GWP_A1_A3_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_A1_A3_TOTAL
            + INTERNAL_WALLS_BASE_GWP_A1_A3_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_A1_A3_TOTAL
            + FLOOR_FINISH_GROUND_GWP_A1_A3_TOTAL
            + FLOOR_FINISH_BELOW_GWP_A1_A3_TOTAL
            + CEILING_FINISH_ABOVE_GWP_A1_A3_TOTAL
            + CEILING_FINISH_GROUND_GWP_A1_A3_TOTAL
            + CEILING_FINISH_BELOW_GWP_A1_A3_TOTAL
            + STAIRCASE_GWP_A1_A3_TOTAL
        )

        EMISSIONS_INTERIOR_GWP_A4 = (
            INTERNAL_WALLS_GWP_A4_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_A4_TOTAL
            + INTERNAL_WALLS_BASE_GWP_A4_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_A4_TOTAL
            + FLOOR_FINISH_GROUND_GWP_A4_TOTAL
            + FLOOR_FINISH_BELOW_GWP_A4_TOTAL
            + CEILING_FINISH_ABOVE_GWP_A4_TOTAL
            + CEILING_FINISH_GROUND_GWP_A4_TOTAL
            + CEILING_FINISH_BELOW_GWP_A4_TOTAL
            + STAIRCASE_GWP_A4_TOTAL
        )
        EMISSIONS_INTERIOR_GWP_A5 = (
            INTERNAL_WALLS_GWP_A5_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_A5_TOTAL
            + INTERNAL_WALLS_BASE_GWP_A5_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_A5_TOTAL
            + FLOOR_FINISH_GROUND_GWP_A5_TOTAL
            + FLOOR_FINISH_BELOW_GWP_A5_TOTAL
            + CEILING_FINISH_ABOVE_GWP_A5_TOTAL
            + CEILING_FINISH_GROUND_GWP_A5_TOTAL
            + CEILING_FINISH_BELOW_GWP_A5_TOTAL
            + STAIRCASE_GWP_A5_TOTAL
        )
        EMISSIONS_INTERIOR_GWP_B4 = (
            INTERNAL_WALLS_GWP_B4_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_B4_TOTAL
            + INTERNAL_WALLS_BASE_GWP_B4_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_B4_TOTAL
            + FLOOR_FINISH_GROUND_GWP_B4_TOTAL
            + FLOOR_FINISH_BELOW_GWP_B4_TOTAL
            + CEILING_FINISH_ABOVE_GWP_B4_TOTAL
            + CEILING_FINISH_GROUND_GWP_B4_TOTAL
            + CEILING_FINISH_BELOW_GWP_B4_TOTAL
            + STAIRCASE_GWP_B4_TOTAL
        )
        EMISSIONS_INTERIOR_GWP_C3 = (
            INTERNAL_WALLS_GWP_C3_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_C3_TOTAL
            + INTERNAL_WALLS_BASE_GWP_C3_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_C3_TOTAL
            + FLOOR_FINISH_GROUND_GWP_C3_TOTAL
            + FLOOR_FINISH_BELOW_GWP_C3_TOTAL
            + CEILING_FINISH_ABOVE_GWP_C3_TOTAL
            + CEILING_FINISH_GROUND_GWP_C3_TOTAL
            + CEILING_FINISH_BELOW_GWP_C3_TOTAL
            + STAIRCASE_GWP_C3_TOTAL
        )
        EMISSIONS_INTERIOR_GWP_C4 = (
            INTERNAL_WALLS_GWP_C4_TOTAL
            + INTERNAL_WALLS_GROUND_GWP_C4_TOTAL
            + INTERNAL_WALLS_BASE_GWP_C4_TOTAL
            + FLOOR_FINISH_ABOVE_GWP_C4_TOTAL
            + FLOOR_FINISH_GROUND_GWP_C4_TOTAL
            + FLOOR_FINISH_BELOW_GWP_C4_TOTAL
            + CEILING_FINISH_ABOVE_GWP_C4_TOTAL
            + CEILING_FINISH_GROUND_GWP_C4_TOTAL
            + CEILING_FINISH_BELOW_GWP_C4_TOTAL
            + STAIRCASE_GWP_C4_TOTAL
        )
        if condition == "demolition":
            (
                EMISSIONS_INTERIOR_GWP_A1_A3,
                EMISSIONS_INTERIOR_GWP_A4,
                EMISSIONS_INTERIOR_GWP_A5,
                EMISSIONS_INTERIOR_GWP_B4,
                EMISSIONS_INTERIOR_GWP_C3,
                EMISSIONS_INTERIOR_GWP_C4,
            ) = (0, 0, 0, 0, 0, 0)
        EMISSIONS_INTERIOR_GWP_TOTAL = (
            EMISSIONS_INTERIOR_GWP_A1_A3
            + EMISSIONS_INTERIOR_GWP_A4
            + EMISSIONS_INTERIOR_GWP_A5
            + EMISSIONS_INTERIOR_GWP_B4
            + EMISSIONS_INTERIOR_GWP_C3
            + EMISSIONS_INTERIOR_GWP_C4
        )

        self.EMISSIONS_INTERIOR_GWP_TOTAL = EMISSIONS_INTERIOR_GWP_TOTAL

        EMISSIONS_TECHNICAL_GWP_A1_A3 = (
            TECHNICAL_ABOVE_GWP_A1_A3_TOTAL
            + TECHNICAL_GROUND_GWP_A1_A3_TOTAL
            + TECHNICAL_BELOW_GWP_A1_A3_TOTAL
            + ELEVATOR_GWP_A1_A3_TOTAL
        )
        EMISSIONS_TECHNICAL_GWP_A4 = (
            TECHNICAL_ABOVE_GWP_A4_TOTAL
            + TECHNICAL_GROUND_GWP_A4_TOTAL
            + TECHNICAL_BELOW_GWP_A4_TOTAL
            + ELEVATOR_GWP_A4_TOTAL
        )
        EMISSIONS_TECHNICAL_GWP_A5 = (
            TECHNICAL_ABOVE_GWP_A5_TOTAL
            + TECHNICAL_GROUND_GWP_A5_TOTAL
            + TECHNICAL_BELOW_GWP_A5_TOTAL
            + ELEVATOR_GWP_A5_TOTAL
        )
        EMISSIONS_TECHNICAL_GWP_B4 = (
            TECHNICAL_ABOVE_GWP_B4_TOTAL
            + TECHNICAL_GROUND_GWP_B4_TOTAL
            + TECHNICAL_BELOW_GWP_B4_TOTAL
            + ELEVATOR_GWP_B4_TOTAL
        )
        EMISSIONS_TECHNICAL_GWP_C3 = (
            TECHNICAL_ABOVE_GWP_C3_TOTAL
            + TECHNICAL_GROUND_GWP_C3_TOTAL
            + TECHNICAL_BELOW_GWP_C3_TOTAL
            + ELEVATOR_GWP_C3_TOTAL
        )
        EMISSIONS_TECHNICAL_GWP_C4 = (
            TECHNICAL_ABOVE_GWP_C4_TOTAL
            + TECHNICAL_GROUND_GWP_C4_TOTAL
            + TECHNICAL_BELOW_GWP_C4_TOTAL
            + ELEVATOR_GWP_C4_TOTAL
        )
        if condition == "demolition":
            (
                EMISSIONS_TECHNICAL_GWP_A1_A3,
                EMISSIONS_TECHNICAL_GWP_A4,
                EMISSIONS_TECHNICAL_GWP_A5,
                EMISSIONS_TECHNICAL_GWP_B4,
                EMISSIONS_TECHNICAL_GWP_C3,
                EMISSIONS_TECHNICAL_GWP_C4,
            ) = (0, 0, 0, 0, 0, 0)
        EMISSIONS_TECHNICAL_GWP_TOTAL = (
            EMISSIONS_TECHNICAL_GWP_A1_A3
            + EMISSIONS_TECHNICAL_GWP_A4
            + EMISSIONS_TECHNICAL_GWP_A5
            + EMISSIONS_TECHNICAL_GWP_B4
            + EMISSIONS_TECHNICAL_GWP_C3
            + EMISSIONS_TECHNICAL_GWP_C4
        )

        EMISSIONS_CONSTRUCTION_GWP_A5 = CONSTRUCTION_GWP_TOTAL
        if condition == "demolition":
            EMISSIONS_CONSTRUCTION_GWP_TOTAL = 0
        else:
            if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
                EMISSIONS_CONSTRUCTION_GWP_TOTAL = EMISSIONS_CONSTRUCTION_GWP_A5
            else:
                EMISSIONS_CONSTRUCTION_GWP_TOTAL = 0

        # CONFIRM 4 row in excel is empty for Emissions of operational energy
        if condition == "demolition":
            EMISSIONS_OPERATIONAL_GWP_B6 = 0
            EMISSIONS_OPERATIONAL_GWP_TOTAL = EMISSIONS_OPERATIONAL_GWP_B6
        else:
            EMISSIONS_OPERATIONAL_GWP_B6 = EMISSIONS_OPERATIONAL
            EMISSIONS_OPERATIONAL_GWP_TOTAL = EMISSIONS_OPERATIONAL_GWP_B6

        # DEMOLITION

        DEMOLITION_GWP_C3 = 0
        DEMOLITION_GWP_C4 = 0
        if include_demolition and condition == "demolition":
            DEMOLITION_GWP_C3 = (
                PAD_GWP_C3_TOTAL
                + LINE_GWP_C3_TOTAL
                + GROUND_SLAB_GWP_C3_TOTAL
                + LOW_SLAB_A_GWP_C3_TOTAL
                + SLAB_ABOVE_GWP_C3_TOTAL
                + VERTICAL_ABOVE_GWP_C3_TOTAL
                + BEAM_ABOVE_GWP_C3_TOTAL
                + LOW_SLAB_B_GWP_C3_TOTAL
                + SLAB_BELOW_GWP_C3_TOTAL
                + VERTICAL_BELOW_GWP_C3_TOTAL
                + BEAM_BELOW_GWP_C3_TOTAL
                + BASE_WALL_GWP_C3_TOTAL
                + PRIM_FACADE_GWP_C3_TOTAL
                + SEC_FACADE_GWP_C3_TOTAL
                + WINDOW_GWP_C3_TOTAL
                + CORE_GWP_C3_TOTAL
                + INTERNAL_WALLS_GWP_C3_TOTAL
                + INTERNAL_WALLS_GROUND_GWP_C3_TOTAL
                + INTERNAL_WALLS_BASE_GWP_C3_TOTAL
                + ROOF_FINISH_GWP_C3_TOTAL
                + FLOOR_FINISH_GROUND_GWP_C3_TOTAL
                + FLOOR_FINISH_ABOVE_GWP_C3_TOTAL
                + FLOOR_FINISH_BELOW_GWP_C3_TOTAL
                + CEILING_FINISH_ABOVE_GWP_C3_TOTAL
                + CEILING_FINISH_GROUND_GWP_C3_TOTAL
                + CEILING_FINISH_BELOW_GWP_C3_TOTAL
                + STAIRCASE_GWP_C3_TOTAL
                + ELEVATOR_GWP_C3_TOTAL
                + TECHNICAL_ABOVE_GWP_C3_TOTAL
                + TECHNICAL_BELOW_GWP_C3_TOTAL
                + TECHNICAL_GROUND_GWP_C3_TOTAL
            )

            DEMOLITION_GWP_C4 = (
                PAD_GWP_C4_TOTAL
                + LINE_GWP_C4_TOTAL
                + GROUND_SLAB_GWP_C4_TOTAL
                + LOW_SLAB_A_GWP_C4_TOTAL
                + SLAB_ABOVE_GWP_C4_TOTAL
                + VERTICAL_ABOVE_GWP_C4_TOTAL
                + BEAM_ABOVE_GWP_C4_TOTAL
                + LOW_SLAB_B_GWP_C4_TOTAL
                + SLAB_BELOW_GWP_C4_TOTAL
                + VERTICAL_BELOW_GWP_C4_TOTAL
                + BEAM_BELOW_GWP_C4_TOTAL
                + BASE_WALL_GWP_C4_TOTAL
                + PRIM_FACADE_GWP_C4_TOTAL
                + SEC_FACADE_GWP_C4_TOTAL
                + WINDOW_GWP_C4_TOTAL
                + CORE_GWP_C4_TOTAL
                + INTERNAL_WALLS_GWP_C4_TOTAL
                + INTERNAL_WALLS_GROUND_GWP_C4_TOTAL
                + INTERNAL_WALLS_BASE_GWP_C4_TOTAL
                + ROOF_FINISH_GWP_C4_TOTAL
                + FLOOR_FINISH_GROUND_GWP_C4_TOTAL
                + FLOOR_FINISH_ABOVE_GWP_C4_TOTAL
                + FLOOR_FINISH_BELOW_GWP_C4_TOTAL
                + CEILING_FINISH_ABOVE_GWP_C4_TOTAL
                + CEILING_FINISH_GROUND_GWP_C4_TOTAL
                + CEILING_FINISH_BELOW_GWP_C4_TOTAL
                + STAIRCASE_GWP_C4_TOTAL
                + ELEVATOR_GWP_C4_TOTAL
                + TECHNICAL_ABOVE_GWP_C4_TOTAL
                + TECHNICAL_BELOW_GWP_C4_TOTAL
                + TECHNICAL_GROUND_GWP_C4_TOTAL
            )
        elif include_demolition is False and condition == "demolition":
            DEMOLITION_GWP_C1 = 0

        self.EMISSIONS_DEMOLITION_GWP_TOTAL = (
            DEMOLITION_GWP_C1 + DEMOLITION_GWP_C3 + DEMOLITION_GWP_C4
        )

        # endregion

        # region Results for categories
        EMISSIONS_FOUNDATION_M2 = self.EMISSIONS_FOUNDATION_GWP_TOTAL / TOTAL_AREA
        EMISSIONS_FOUNDATION_M2_YEAR = EMISSIONS_FOUNDATION_M2 / study_period
        try:
            EMISSIONS_FOUNDATION_PERCENT = self.EMISSIONS_FOUNDATION_GWP_TOTAL / (
                self.EMISSIONS_FOUNDATION_GWP_TOTAL
                + self.EMISSIONS_STRUCTURE_GWP_TOTAL
                + EMISSIONS_ENVELOPE_GWP_TOTAL
                + EMISSIONS_WINDOW_GWP_TOTAL
                + EMISSIONS_INTERIOR_GWP_TOTAL
                + EMISSIONS_TECHNICAL_GWP_TOTAL
                + EMISSIONS_CONSTRUCTION_GWP_TOTAL
                + EMISSIONS_OPERATIONAL_GWP_TOTAL
            )
        except ZeroDivisionError:
            EMISSIONS_FOUNDATION_PERCENT = 0
        # For structure
        EMISSIONS_STRUCTURE_M2 = self.EMISSIONS_STRUCTURE_GWP_TOTAL / TOTAL_AREA
        EMISSIONS_STRUCTURE_M2_YEAR = EMISSIONS_STRUCTURE_M2 / study_period
        try:
            EMISSIONS_STRUCTURE_PERCENT = self.EMISSIONS_STRUCTURE_GWP_TOTAL / (
                self.EMISSIONS_FOUNDATION_GWP_TOTAL
                + self.EMISSIONS_STRUCTURE_GWP_TOTAL
                + EMISSIONS_ENVELOPE_GWP_TOTAL
                + EMISSIONS_WINDOW_GWP_TOTAL
                + EMISSIONS_INTERIOR_GWP_TOTAL
                + EMISSIONS_TECHNICAL_GWP_TOTAL
                + EMISSIONS_CONSTRUCTION_GWP_TOTAL
                + EMISSIONS_OPERATIONAL_GWP_TOTAL
            )
        except ZeroDivisionError:
            EMISSIONS_STRUCTURE_PERCENT = 0
        # For ENVELOPE
        EMISSIONS_ENVELOPE_M2 = EMISSIONS_ENVELOPE_GWP_TOTAL / TOTAL_AREA
        EMISSIONS_ENVELOPE_M2_YEAR = EMISSIONS_ENVELOPE_M2 / study_period

        # For WINDOW
        EMISSIONS_WINDOW_M2 = EMISSIONS_WINDOW_GWP_TOTAL / TOTAL_AREA
        EMISSIONS_WINDOW_M2_YEAR = EMISSIONS_WINDOW_M2 / study_period

        # For INTERIOR
        EMISSIONS_INTERIOR_M2 = EMISSIONS_INTERIOR_GWP_TOTAL / TOTAL_AREA
        EMISSIONS_INTERIOR_M2_YEAR = EMISSIONS_INTERIOR_M2 / study_period

        # For TECHNICAL
        EMISSIONS_TECHNICAL_M2 = EMISSIONS_TECHNICAL_GWP_TOTAL / TOTAL_AREA
        EMISSIONS_TECHNICAL_M2_YEAR = EMISSIONS_TECHNICAL_M2 / study_period

        # For CONSTRUCTION
        EMISSIONS_CONSTRUCTION_M2 = EMISSIONS_CONSTRUCTION_GWP_TOTAL / TOTAL_AREA
        EMISSIONS_CONSTRUCTION_M2_YEAR = EMISSIONS_CONSTRUCTION_M2 / study_period

        # For OPERATIONAL
        EMISSIONS_OPERATIONAL_M2 = EMISSIONS_OPERATIONAL_GWP_TOTAL / TOTAL_AREA
        EMISSIONS_OPERATIONAL_M2_YEAR = EMISSIONS_OPERATIONAL_M2 / study_period

        try:
            EMISSIONS_ENVELOPE_PERCENT = EMISSIONS_ENVELOPE_GWP_TOTAL / (
                self.EMISSIONS_FOUNDATION_GWP_TOTAL
                + self.EMISSIONS_STRUCTURE_GWP_TOTAL
                + EMISSIONS_ENVELOPE_GWP_TOTAL
                + EMISSIONS_WINDOW_GWP_TOTAL
                + EMISSIONS_INTERIOR_GWP_TOTAL
                + EMISSIONS_TECHNICAL_GWP_TOTAL
                + EMISSIONS_CONSTRUCTION_GWP_TOTAL
                + EMISSIONS_OPERATIONAL_GWP_TOTAL
            )

            EMISSIONS_WINDOW_PERCENT = EMISSIONS_WINDOW_GWP_TOTAL / (
                self.EMISSIONS_FOUNDATION_GWP_TOTAL
                + self.EMISSIONS_STRUCTURE_GWP_TOTAL
                + EMISSIONS_ENVELOPE_GWP_TOTAL
                + EMISSIONS_WINDOW_GWP_TOTAL
                + EMISSIONS_INTERIOR_GWP_TOTAL
                + EMISSIONS_TECHNICAL_GWP_TOTAL
                + EMISSIONS_CONSTRUCTION_GWP_TOTAL
                + EMISSIONS_OPERATIONAL_GWP_TOTAL
            )

            EMISSIONS_INTERIOR_PERCENT = EMISSIONS_INTERIOR_GWP_TOTAL / (
                self.EMISSIONS_FOUNDATION_GWP_TOTAL
                + self.EMISSIONS_STRUCTURE_GWP_TOTAL
                + EMISSIONS_ENVELOPE_GWP_TOTAL
                + EMISSIONS_WINDOW_GWP_TOTAL
                + EMISSIONS_INTERIOR_GWP_TOTAL
                + EMISSIONS_TECHNICAL_GWP_TOTAL
                + EMISSIONS_CONSTRUCTION_GWP_TOTAL
                + EMISSIONS_OPERATIONAL_GWP_TOTAL
            )

            EMISSIONS_TECHNICAL_PERCENT = EMISSIONS_TECHNICAL_GWP_TOTAL / (
                self.EMISSIONS_FOUNDATION_GWP_TOTAL
                + self.EMISSIONS_STRUCTURE_GWP_TOTAL
                + EMISSIONS_ENVELOPE_GWP_TOTAL
                + EMISSIONS_WINDOW_GWP_TOTAL
                + EMISSIONS_INTERIOR_GWP_TOTAL
                + EMISSIONS_TECHNICAL_GWP_TOTAL
                + EMISSIONS_CONSTRUCTION_GWP_TOTAL
                + EMISSIONS_OPERATIONAL_GWP_TOTAL
            )

            EMISSIONS_CONSTRUCTION_PERCENT = EMISSIONS_CONSTRUCTION_GWP_TOTAL / (
                self.EMISSIONS_FOUNDATION_GWP_TOTAL
                + self.EMISSIONS_STRUCTURE_GWP_TOTAL
                + EMISSIONS_ENVELOPE_GWP_TOTAL
                + EMISSIONS_WINDOW_GWP_TOTAL
                + EMISSIONS_INTERIOR_GWP_TOTAL
                + EMISSIONS_TECHNICAL_GWP_TOTAL
                + EMISSIONS_CONSTRUCTION_GWP_TOTAL
                + EMISSIONS_OPERATIONAL_GWP_TOTAL
            )

            EMISSIONS_OPERATIONAL_PERCENT = EMISSIONS_OPERATIONAL_GWP_TOTAL / (
                self.EMISSIONS_FOUNDATION_GWP_TOTAL
                + self.EMISSIONS_STRUCTURE_GWP_TOTAL
                + EMISSIONS_ENVELOPE_GWP_TOTAL
                + EMISSIONS_WINDOW_GWP_TOTAL
                + EMISSIONS_INTERIOR_GWP_TOTAL
                + EMISSIONS_TECHNICAL_GWP_TOTAL
                + EMISSIONS_CONSTRUCTION_GWP_TOTAL
                + EMISSIONS_OPERATIONAL_GWP_TOTAL
            )

        except ZeroDivisionError:
            EMISSIONS_ENVELOPE_PERCENT = 0
            EMISSIONS_WINDOW_PERCENT = 0
            EMISSIONS_INTERIOR_PERCENT = 0
            EMISSIONS_TECHNICAL_PERCENT = 0
            EMISSIONS_CONSTRUCTION_PERCENT = 0
            EMISSIONS_OPERATIONAL_PERCENT = 0

        # endregion

        # region Results for life-cycle phases
        EMISSIONS_GWP_A1_A3 = (
            EMISSIONS_FOUNDATION_GWP_A1_A3
            + EMISSIONS_STRUCTURE_GWP_A1_A3
            + EMISSIONS_ENVELOPE_GWP_A1_A3
            + EMISSIONS_WINDOW_GWP_A1_A3
            + EMISSIONS_INTERIOR_GWP_A1_A3
            + EMISSIONS_TECHNICAL_GWP_A1_A3
        )
        EMISSIONS_GWP_A1_A3_M2 = EMISSIONS_GWP_A1_A3 / TOTAL_AREA
        EMISSIONS_GWP_A1_A3_M2_YEAR = EMISSIONS_GWP_A1_A3_M2 / study_period

        EMISSIONS_GWP_A4 = (
            EMISSIONS_FOUNDATION_GWP_A4
            + EMISSIONS_STRUCTURE_GWP_A4
            + EMISSIONS_ENVELOPE_GWP_A4
            + EMISSIONS_WINDOW_GWP_A4
            + EMISSIONS_INTERIOR_GWP_A4
            + EMISSIONS_TECHNICAL_GWP_A4
        )
        EMISSIONS_GWP_A4_M2 = EMISSIONS_GWP_A4 / TOTAL_AREA
        EMISSIONS_GWP_A4_M2_YEAR = EMISSIONS_GWP_A4_M2 / study_period

        EMISSIONS_GWP_A5 = (
            EMISSIONS_FOUNDATION_GWP_A5
            + EMISSIONS_STRUCTURE_GWP_A5
            + EMISSIONS_ENVELOPE_GWP_A5
            + EMISSIONS_WINDOW_GWP_A5
            + EMISSIONS_INTERIOR_GWP_A5
            + EMISSIONS_TECHNICAL_GWP_A5
            + EMISSIONS_CONSTRUCTION_GWP_A5
        )
        EMISSIONS_GWP_A5_M2 = EMISSIONS_GWP_A5 / TOTAL_AREA
        EMISSIONS_GWP_A5_M2_YEAR = EMISSIONS_GWP_A5_M2 / study_period

        EMISSIONS_GWP_B4 = (
            EMISSIONS_FOUNDATION_GWP_B4
            + EMISSIONS_STRUCTURE_GWP_B4
            + EMISSIONS_ENVELOPE_GWP_B4
            + EMISSIONS_WINDOW_GWP_B4
            + EMISSIONS_INTERIOR_GWP_B4
            + EMISSIONS_TECHNICAL_GWP_B4
        )
        EMISSIONS_GWP_B4_M2 = EMISSIONS_GWP_B4 / TOTAL_AREA
        EMISSIONS_GWP_B4_M2_YEAR = EMISSIONS_GWP_B4_M2 / study_period

        EMISSIONS_GWP_B6 = EMISSIONS_OPERATIONAL_GWP_B6
        EMISSIONS_GWP_B6_M2 = EMISSIONS_GWP_B6 / TOTAL_AREA
        EMISSIONS_GWP_B4_M2_YEAR = EMISSIONS_GWP_B6_M2 / study_period

        EMISSIONS_GWP_C1 = DEMOLITION_GWP_C1
        EMISSIONS_GWP_C1_M2 = EMISSIONS_GWP_C1 / TOTAL_AREA
        EMISSIONS_GWP_C1_M2_YEAR = EMISSIONS_GWP_C1_M2 / study_period

        EMISSIONS_GWP_C3 = (
            EMISSIONS_FOUNDATION_GWP_C3
            + EMISSIONS_STRUCTURE_GWP_C3
            + EMISSIONS_ENVELOPE_GWP_C3
            + EMISSIONS_WINDOW_GWP_C3
            + EMISSIONS_INTERIOR_GWP_C3
            + EMISSIONS_TECHNICAL_GWP_C3
            + DEMOLITION_GWP_C3
        )
        EMISSIONS_GWP_C3_M2 = EMISSIONS_GWP_C3 / TOTAL_AREA
        EMISSIONS_GWP_C3_M2_YEAR = EMISSIONS_GWP_C3_M2 / study_period

        EMISSIONS_GWP_C4 = (
            EMISSIONS_FOUNDATION_GWP_C4
            + EMISSIONS_STRUCTURE_GWP_C4
            + EMISSIONS_ENVELOPE_GWP_C4
            + EMISSIONS_WINDOW_GWP_C4
            + EMISSIONS_INTERIOR_GWP_C4
            + EMISSIONS_TECHNICAL_GWP_C4
            + DEMOLITION_GWP_C4
        )
        EMISSIONS_GWP_C4_M2 = EMISSIONS_GWP_C4 / TOTAL_AREA
        EMISSIONS_GWP_C4_M2_YEAR = EMISSIONS_GWP_C4_M2 / study_period

        try:
            EMISSIONS_A1_A3_PERCENT = EMISSIONS_GWP_A1_A3 / (
                EMISSIONS_GWP_A1_A3
                + EMISSIONS_GWP_A4
                + EMISSIONS_GWP_B4
                + EMISSIONS_GWP_B6
                + EMISSIONS_GWP_C1
                + EMISSIONS_GWP_C3
                + EMISSIONS_GWP_C4
            )
        except:
            EMISSIONS_A1_A3_PERCENT = 0

        try:
            EMISSIONS_A4_PERCENT = EMISSIONS_GWP_A4 / (
                EMISSIONS_GWP_A1_A3
                + EMISSIONS_GWP_A4
                + EMISSIONS_GWP_A5
                + EMISSIONS_GWP_B4
                + EMISSIONS_GWP_B6
                + EMISSIONS_GWP_C1
                + EMISSIONS_GWP_C3
                + EMISSIONS_GWP_C4
            )
        except:
            EMISSIONS_A4_PERCENT = 0

        try:
            EMISSIONS_A5_PERCENT = EMISSIONS_GWP_A5 / (
                EMISSIONS_GWP_A1_A3
                + EMISSIONS_GWP_A4
                + EMISSIONS_GWP_A5
                + EMISSIONS_GWP_B4
                + EMISSIONS_GWP_B6
                + EMISSIONS_GWP_C1
                + EMISSIONS_GWP_C3
                + EMISSIONS_GWP_C4
            )
        except:
            EMISSIONS_A5_PERCENT = 0

        try:
            EMISSIONS_B4_PERCENT = EMISSIONS_GWP_B4 / (
                EMISSIONS_GWP_A1_A3
                + EMISSIONS_GWP_A4
                + EMISSIONS_GWP_A5
                + EMISSIONS_GWP_B4
                + EMISSIONS_GWP_B6
                + EMISSIONS_GWP_C1
                + EMISSIONS_GWP_C3
                + EMISSIONS_GWP_C4
            )
        except:
            EMISSIONS_B4_PERCENT = 0

        try:
            EMISSIONS_B6_PERCENT = EMISSIONS_GWP_B6 / (
                EMISSIONS_GWP_A1_A3
                + EMISSIONS_GWP_A4
                + EMISSIONS_GWP_A5
                + EMISSIONS_GWP_B4
                + EMISSIONS_GWP_B6
                + EMISSIONS_GWP_C1
                + EMISSIONS_GWP_C3
                + EMISSIONS_GWP_C4
            )
        except:
            EMISSIONS_B4_PERCENT = 0

        try:
            EMISSIONS_C1_PERCENT = EMISSIONS_GWP_C1 / (
                EMISSIONS_GWP_A1_A3
                + EMISSIONS_GWP_A4
                + EMISSIONS_GWP_A5
                + EMISSIONS_GWP_B4
                + EMISSIONS_GWP_B6
                + EMISSIONS_GWP_C1
                + EMISSIONS_GWP_C3
                + EMISSIONS_GWP_C4
            )
        except:
            EMISSIONS_C1_PERCENT = 0

        try:
            EMISSIONS_C3_PERCENT = EMISSIONS_GWP_C3 / (
                EMISSIONS_GWP_A1_A3
                + EMISSIONS_GWP_A4
                + EMISSIONS_GWP_A5
                + EMISSIONS_GWP_B4
                + EMISSIONS_GWP_B6
                + EMISSIONS_GWP_C1
                + EMISSIONS_GWP_C3
                + EMISSIONS_GWP_C4
            )
        except:
            EMISSIONS_C3_PERCENT = 0

        try:
            EMISSIONS_C4_PERCENT = EMISSIONS_GWP_C4 / (
                EMISSIONS_GWP_A1_A3
                + EMISSIONS_GWP_A4
                + EMISSIONS_GWP_A5
                + EMISSIONS_GWP_B4
                + EMISSIONS_GWP_B6
                + EMISSIONS_GWP_C1
                + EMISSIONS_GWP_C3
                + EMISSIONS_GWP_C4
            )
        except:
            EMISSIONS_C4_PERCENT = 0
        # endregion

        # region Total results
        self.TOTAL_EMISSIONS = (
            self.EMISSIONS_FOUNDATION_GWP_TOTAL
            + self.EMISSIONS_STRUCTURE_GWP_TOTAL
            + EMISSIONS_ENVELOPE_GWP_TOTAL
            + EMISSIONS_WINDOW_GWP_TOTAL
            + EMISSIONS_INTERIOR_GWP_TOTAL
            + EMISSIONS_TECHNICAL_GWP_TOTAL
            + EMISSIONS_CONSTRUCTION_GWP_TOTAL
            + self.EMISSIONS_DEMOLITION_GWP_TOTAL
            + EMISSIONS_OPERATIONAL_GWP_TOTAL
        )

        self.EMISSIONS_CONSTRUCTION_GWP_TOTAL = EMISSIONS_CONSTRUCTION_GWP_TOTAL
        self.EMISSIONS_TECHNICAL_GWP_TOTAL = EMISSIONS_TECHNICAL_GWP_TOTAL
        self.EMISSIONS_WINDOW_GWP_TOTAL = EMISSIONS_WINDOW_GWP_TOTAL
        self.EMISSIONS_ENVELOPE_GWP_TOTAL = EMISSIONS_ENVELOPE_GWP_TOTAL
        self.TOTAL_EMISSIONS_M2 = self.TOTAL_EMISSIONS / TOTAL_AREA

        self.TOTAL_EMISSIONS_M2_YEAR = self.TOTAL_EMISSIONS_M2 / study_period
        if num_people != 0:
            self.TOTAL_EMISSIONS_PERSON = self.TOTAL_EMISSIONS / num_people
        else:
            self.TOTAL_EMISSIONS_PERSON = 0

        self.TOTAL_EMISSIONS_PERSON_YEAR = self.TOTAL_EMISSIONS_PERSON / study_period
        # endregion
        #JCN/2024-10-24: Værdi for alle kategorier skal udstilles
        self.EMISSIONS_ENVELOPE_GWP_TOTAL = EMISSIONS_ENVELOPE_GWP_TOTAL
        self.EMISSIONS_WINDOW_GWP_TOTAL = EMISSIONS_WINDOW_GWP_TOTAL
        self.EMISSIONS_INTERIOR_GWP_TOTAL = EMISSIONS_INTERIOR_GWP_TOTAL
        self.EMISSIONS_TECHNICAL_GWP_TOTAL = EMISSIONS_TECHNICAL_GWP_TOTAL
        self.EMISSIONS_CONSTRUCTION_GWP_TOTAL = EMISSIONS_CONSTRUCTION_GWP_TOTAL
        #self.EMISSIONS_DEMOLITION_GWP_TOTAL = EMISSIONS_DEMOLITION_GWP_TOTAL
        self.EMISSIONS_OPERATIONAL_GWP_TOTAL = EMISSIONS_OPERATIONAL_GWP_TOTAL

    # region CLASS FUNCTIONS

    def initialize_dimensions(
        self,
        roof_type,
        roof_material,
        roof_angle,
        typology,
        structure,
        width,
        height,
        num_floors,
        ground_typology,
        ground_ftf,
        heating,
        wwr,
        prim_facade,
        sec_facade,
        prim_facade_proportion,
    ):
        # region lookup to calculate roof thickness
        struct = structure.split()[0]
        if struct in ["concrete", "not"] and typology in [
            "detached house",
            "terraced house",
        ]:
            item_roof_thickness = next(
                (
                    item
                    for item in opbygDict
                    if item["assembly_version"]
                    == "low settlement concrete building "
                    + roof_type
                    + " "
                    + roof_material
                ),
                None,
            )
        elif struct in ["timber", "hybrid"] and typology in [
            "detached house",
            "terraced house",
        ]:
            item_roof_thickness = next(
                (
                    item
                    for item in opbygDict
                    if item["assembly_version"]
                    == "low settlement timber building "
                    + roof_type
                    + " "
                    + roof_material
                ),
                None,
            )
        else:
            item_roof_thickness = next(
                (
                    item
                    for item in opbygDict
                    if item["assembly_version"] == roof_type + " " + roof_material
                ),
                None,
            )
        # endregion

        # region roof calculation
        roof_angle = 0 if roof_type == "flat" else roof_angle or c.ROOF_ANGLE[typology]
        ROOF_THICKNESS = item_roof_thickness["thickness [m]"]
        ROOF_HEIGHT = math.tan(math.radians(roof_angle)) * (width / 2) + ROOF_THICKNESS

        # endregion

        # region ground ftf calculation
        if ground_typology is None:
            ground_typology = typology

        if ground_ftf is not None:
            try:
                ftf = ((height - ROOF_HEIGHT) - ground_ftf) / (num_floors - 1)
            except ZeroDivisionError:
                self.ERROR_MESSAGES.append(
                    "The values from floor-to-floor height for ground floor and basement do not match the number of floors."
                )
                return None, None, None, None, None
        else:
            ground_ftf = (height - ROOF_HEIGHT) / num_floors
            ftf = ground_ftf

        # endregion

        # region handle heating systems
        if heating == None:
            heating = "electricity"
        # endregion

        # region choose structure from lookup table
        if structure == "not decided":
            structure = c.STRUCTURE_DICT[typology]
        # endregion

        # region calculate slab thickness
        slab_thickness = 0

        if typology in [
            "detached house",
            "terraced house",
        ]:
            struct = structure.split()[0]
            if struct in ["concrete", "not"]:
                slab_item = next(
                    (item for item in opbygDict if item["assembly_id"] == "#o0411"),
                    None,
                )
                slab_thickness = slab_item["thickness [m]"]
            else:
                slab_item = next(
                    (item for item in opbygDict if item["assembly_id"] == "#o0412"),
                    None,
                )
                slab_thickness = slab_item["thickness [m]"]
        else:
            slab_thickness = c.SLAB_THICKNESS[structure]

        # endregion

        # region calculate floor to ceilings
        ftc = ftf - slab_thickness
        ftc_ground = ground_ftf - slab_thickness
        if ftc < c.MIN_FTC[typology] or ftc_ground < c.MIN_FTC[ground_typology]:
            self.ERROR_MESSAGES.append("Too low floor-to-ceiling-height.")

        # endregion

        # region window to wall ratio
        if wwr is None:
            wwr = c.WWR_DICT[typology]
        else:
            # Convert WWR from %
            wwr /= 100
        # endregion
        # region calculate facade percentages
        prim_facade_proportion /= 100 if sec_facade is not None else 1
        sec_facade_proportion = (
            1 - prim_facade_proportion if sec_facade is not None else 0
        )

        # endregion

        # region calculae volumes

        # endregion
        return (
            roof_angle,
            ROOF_THICKNESS,
            ROOF_HEIGHT,
            ground_ftf,
            ftf,
            ground_typology,
            heating,
            structure,
            slab_thickness,
            ftc,
            ftc_ground,
            wwr,
            prim_facade_proportion,
            sec_facade_proportion,
        )

    def quantity_calculation(
        self, area, basement_depth, height, roof_angle, ROOF_HEIGHT, ftf
    ):
        # region calculate volumes
        volume_below_ground = basement_depth * area
        volume_above_ground = area * height
        volume_top_floor = area * (ftf + ROOF_HEIGHT)
        if roof_angle != 0:
            volume_above_ground = area * (height - (ROOF_HEIGHT / 2))
            volume_top_floor = area * (ftf + (ROOF_HEIGHT / 2))
        volume_total = volume_below_ground + volume_above_ground

        # endregion

        return (
            volume_below_ground,
            volume_above_ground,
            volume_total,
            volume_top_floor,
        )

    def structure_calculation(self, typology, structure, num_floors, num_base_floors):
        # Construct the key using the provided parameters
        key_above_three_index = f"{typology}-{structure}-{num_floors}"
        key_above_two_index = f"{typology}-{structure}"

        key_below_three_index = f"{typology}-{structure}-{num_base_floors}"
        key_below_two_index = f"{typology}-{structure}"

        # Lookup the item in the structure_number_of_floors dictionary
        item_key_structure_number_of_floors = structure_number_of_floors.get(
            key_above_three_index, "key_structure_number_of_floors not found"
        )
        item_structure_floor_height = structure_floor_height.get(
            key_above_two_index, "key_structure_floor_height not found"
        )
        item_structure_roof_angle = structure_roof_angle.get(
            key_above_two_index, "structure_roof_angle not found"
        )

        if num_base_floors != 0:
            item_key_structure_number_of_basement_floors = (
                structure_number_of_basement_floors.get(
                    key_below_three_index,
                    "structure_number_of_basement_floors not found",
                )
            )

            item_structure_basement_floor_height = structure_basement_floor_height.get(
                key_below_two_index, "structure_basement_floor_height not found"
            )
        else:
            item_key_structure_number_of_basement_floors = None
            item_structure_basement_floor_height = None
        # Return the found item or a message if not found
        return (
            item_key_structure_number_of_floors,
            item_structure_floor_height,
            item_structure_roof_angle,
            item_key_structure_number_of_basement_floors,
            item_structure_basement_floor_height,
        )

    def initialize_people(
        self,
        ground_typology,
        typology,
        num_floors,
        area,
        condition,
        num_residents,
        num_workplaces,
    ):

        # residents, workplaces, people
        floor_ground = 1
        if ground_typology is None:
            floor_ground = 0

        if num_residents is None:
            if condition == "demolition":
                num_residents = 0
                num_workplaces = 0
            elif (
                typology == "detached house"
                or typology == "terraced house"
                or typology == "apartment building"
            ):
                num_residents = round(
                    ((num_floors - floor_ground) * area) / c.OCCUPANCY_DICT[typology]
                )
            else:
                num_residents = 0

        if num_workplaces is None:
            if condition == "demolition":
                num_workplaces = 0
            elif typology == "office":
                num_workplaces = round(
                    ((num_floors - floor_ground) * area) / c.OCCUPANCY_DICT[typology]
                )
                if ground_typology == "office":
                    num_workplaces += round(area / c.OCCUPANCY_DICT[ground_typology])
            else:
                num_workplaces = 0

        num_people = num_workplaces + num_residents
        return num_people, num_residents, num_workplaces

    def thickness_facade(self, structureOverView, structure, typology, prim_facade):
        factor = 0  # default value
        struct = structure.split()[0]

        if struct == "concrete" and typology in ["detached house", "terraced house"]:
            factor = structureOverView["Assembly version"].index(
                "low settlement concrete building " + prim_facade
            )
        elif struct in ["timber", "hybrid"] and typology in [
            "detached house",
            "terraced house",
        ]:
            factor = structureOverView["Assembly version"].index(
                "low settlement timber building " + prim_facade
            )
        else:
            factor = structureOverView["Assembly version"].index(prim_facade)
        return factor

    # endregion

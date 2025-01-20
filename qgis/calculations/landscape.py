from .const import constants as c
import math


class Landscape:

    def __init__(self, category, type, condition, length, width, area,
                 num_trees, type_trees, size_trees,
                 num_shrubs, type_shrubs, size_shrubs,
                 study_period, phases):

        if category == 'roads and paths' and width is None:
            width = c.ROAD_WIDTH_DICT[type]

        if area is None:
            area = width * length

        index_study_period = c.STUDY_PERIOD.index(study_period)

        if condition == 'new':
            TYPE_ID = c.ASSEMBLY_SELECTION_LANDSCAPE[type][0]
        elif condition == 'refurbishment':
            TYPE_ID = c.ASSEMBLY_SELECTION_LANDSCAPE[type][1]
        else:
            TYPE_ID = c.ASSEMBLY_SELECTION_LANDSCAPE[type][2]

        if num_trees is not None:
            if type_trees == 'deciduous':
                TREES_ID = c.ASSEMBLY_SELECTION_TREES[size_trees][0]
            else:
                TREES_ID = c.ASSEMBLY_SELECTION_TREES[size_trees][1]

            if type_shrubs == 'deciduous':
                SHRUBS_ID = c.ASSEMBLY_SELECTION_SHRUBS[size_shrubs][0]
            else:
                SHRUBS_ID = c.ASSEMBLY_SELECTION_SHRUBS[size_shrubs][1]

            TREE_GWP_A1_A3 = c.GWP_A1_A3_LANDSCAPE[TREES_ID][index_study_period] * num_trees
            if phases == "A1-A3, A4, A5, B4, B6, C3, C4":
                TREE_GWP_A4 = c.GWP_A4_LANDSCAPE[TREES_ID][index_study_period] * num_trees
                TREE_GWP_A5 = c.GWP_A5_LANDSCAPE[TREES_ID][index_study_period] * num_trees
            else:
                TREE_GWP_A4 = 0
                TREE_GWP_A5 = 0
            TREE_GWP_B4 = c.GWP_B4_LANDSCAPE[TREES_ID][index_study_period] * num_trees
            TREE_GWP_B6 = c.GWP_B6_LANDSCAPE[TREES_ID][index_study_period] * num_trees
            TREE_GWP_C3 = c.GWP_C3_LANDSCAPE[TREES_ID][index_study_period] * num_trees
            TREE_GWP_C4 = c.GWP_C4_LANDSCAPE[TREES_ID][index_study_period] * num_trees
            TREE_GWP_D = c.GWP_D_LANDSCAPE[TREES_ID][index_study_period] * num_trees

            self.EMISSIONS_TREE = TREE_GWP_A1_A3 + TREE_GWP_A4 + TREE_GWP_A5 + TREE_GWP_B4 + TREE_GWP_B6 + TREE_GWP_C3 + TREE_GWP_C4

            SHRUB_GWP_A1_A3 = c.GWP_A1_A3_LANDSCAPE[SHRUBS_ID][index_study_period] * num_shrubs
            if phases == "A1-A3, A4, A5, B4, B6, C3, C4":
                SHRUB_GWP_A4 = c.GWP_A4_LANDSCAPE[SHRUBS_ID][index_study_period] * num_shrubs
                SHRUB_GWP_A5 = c.GWP_A5_LANDSCAPE[SHRUBS_ID][index_study_period] * num_shrubs
            else:
                SHRUB_GWP_A4 = 0
                SHRUB_GWP_A5 = 0
            SHRUB_GWP_B4 = c.GWP_B4_LANDSCAPE[SHRUBS_ID][index_study_period] * num_shrubs
            SHRUB_GWP_B6 = c.GWP_B6_LANDSCAPE[SHRUBS_ID][index_study_period] * num_shrubs
            SHRUB_GWP_C3 = c.GWP_C3_LANDSCAPE[SHRUBS_ID][index_study_period] * num_shrubs
            SHRUB_GWP_C4 = c.GWP_C4_LANDSCAPE[SHRUBS_ID][index_study_period] * num_shrubs
            SHRUB_GWP_D = c.GWP_D_LANDSCAPE[SHRUBS_ID][index_study_period] * num_shrubs

            self.EMISSIONS_SHRUB = SHRUB_GWP_A1_A3 + SHRUB_GWP_A4 + SHRUB_GWP_A5 + SHRUB_GWP_B4 + SHRUB_GWP_B6 + SHRUB_GWP_C3 + SHRUB_GWP_C4
        else:
            self.EMISSIONS_TREE = 0
            self.EMISSIONS_SHRUB = 0

        TYPE_GWP_A1_A3 = c.GWP_A1_A3_LANDSCAPE[TYPE_ID][index_study_period] * area
        if phases == "A1-A3, A4, A5, B4, B6, C1, C3, C4":
            TYPE_GWP_A4 = c.GWP_A4_LANDSCAPE[TYPE_ID][index_study_period] * area
            TYPE_GWP_A5 = c.GWP_A5_LANDSCAPE[TYPE_ID][index_study_period] * area
        else:
            TYPE_GWP_A4 = 0
            TYPE_GWP_A5 = 0
        TYPE_GWP_B4 = c.GWP_B4_LANDSCAPE[TYPE_ID][index_study_period] * area
        TYPE_GWP_B6 = c.GWP_B6_LANDSCAPE[TYPE_ID][index_study_period] * area
        TYPE_GWP_C3 = c.GWP_C3_LANDSCAPE[TYPE_ID][index_study_period] * area
        TYPE_GWP_C4 = c.GWP_C4_LANDSCAPE[TYPE_ID][index_study_period] * area
        TYPE_GWP_D = c.GWP_D_LANDSCAPE[TYPE_ID][index_study_period] * area

        self.EMISSIONS_TYPE = TYPE_GWP_A1_A3 + TYPE_GWP_A4 + TYPE_GWP_A5 + TYPE_GWP_B4 + TYPE_GWP_B6 + TYPE_GWP_C3 + TYPE_GWP_C4

        # region Results

        self.EMISSIONS_TYPE_M2 = self.EMISSIONS_TYPE / area
        self.EMISSIONS_TYPE_M2_YEAR = self.EMISSIONS_TYPE_M2 / study_period

        self.EMISSIONS_TREES_SHRUBS = self.EMISSIONS_SHRUB + self.EMISSIONS_TREE
        self.EMISSIONS_TREES_SHRUBS_M2 = self.EMISSIONS_TREES_SHRUBS / area
        self.EMISSIONS_TREES_SHRUBS_M2_YEAR = self.EMISSIONS_TREES_SHRUBS_M2 / study_period

        # TO DO create more results depending on what we will visualize
        # endregion

# region Dictionaries to assign values based on structure (Concrete (load bearing), Concrete (framed), Hybrid timber (framed), Timber (load bearing), Timber (framed))
VERTICAL_STRUCTURE = {
    "concrete (load bearing)": "reinforced concrete wall",
    "concrete (framed)": "reinforced concrete column",
    "hybrid timber (framed)": "glulam column",
    "timber (load bearing)": "clt wall",
    "timber (framed)": "glulam column"
}

BEAM_STRUCTURE = {
    "concrete (load bearing)": "no beam",
    "concrete (framed)": "reinforced concrete beam",
    "hybrid timber (framed)": "glulam beam",
    "timber (load bearing)": "no beam",
    "timber (framed)": "glulam beam"
}

SLAB_STRUCTURE = {
    "concrete (load bearing)": "hollow-core slab",
    "concrete (framed)": "hollow-core slab",
    "hybrid timber (framed)": "hybrid concrete/clt slab",
    "timber (load bearing)": "clt slab",
    "timber (framed)": "lvl slab"
}

SLAB_THICKNESS = {
    "concrete (load bearing)": 0.32,
    "concrete (framed)": 0.32,
    "hybrid timber (framed)": 0.28,
    "timber (load bearing)": 0.20,
    "timber (framed)": 0.453
}

PAD_FOUNDATION = {
    "concrete (load bearing)": "pad foundation",
    "concrete (framed)": "pad foundation",
    "hybrid timber (framed)": "pad foundation",
    "timber (load bearing)": "pad foundation",
    "timber (framed)": "pad foundation"
}

LINE_FOUNDATION = {
    "concrete (load bearing)": "concrete building",
    "concrete (framed)": "concrete building",
    "hybrid timber (framed)": "timber building",
    "timber (load bearing)": "timber building",
    "timber (framed)": "timber building"
}

CORE_WALL = {
    "concrete (load bearing)": "reinforced concrete wall",
    "concrete (framed)": "reinforced concrete wall",
    "hybrid timber (framed)": "reinforced concrete wall",
    "timber (load bearing)": "clt wall",
    "timber (framed)": "clt wall"
}

INTERNAL_WALLS = {
    "concrete (load bearing)": "concrete wall",
    "concrete (framed)": "steel stud (light)",
    "hybrid timber (framed)": "timber stud (light)",
    "timber (load bearing)": "timber stud (light)",
    "timber (framed)": "timber stud (light)"
}

DEMOLITION_DICT = {
    "concrete (load bearing)": 5.47,
    "concrete (framed)": 5.47,
    "hybrid timber (framed)": 5.43,
    "timber (load bearing)": 5.37,
    "timber (framed)": 5.37
}
# endregion
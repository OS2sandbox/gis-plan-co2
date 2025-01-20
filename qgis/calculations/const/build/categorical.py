# region categorical (dropdowns)
PARKING_BASEMENT = ["yes", "no"]
CONDITION_LIST = [
    "new",
    "transformation",
    "renovation",
    "refurbishment",
    "keep",
    "demolition",
]  # CHANGED added demolition
TYPOLOGY = [
    "detached house",
    "terraced house",
    "apartment building",
    "office",
    "school",
    "institution",
    "retail",
    "parking",
    "industry",
    "transport",
    "hotel",
    "hospital",
]
NO_BASEMENT_TYPOLOGY = "no basement"

STRUCTURE_LIST = [
    "concrete (load bearing)",
    "concrete (framed)",
    "hybrid timber (framed)",
    "timber (load bearing)",
    "timber (framed)",
    "not decided",
]

# Facade
FACADE = [
    "brick (heavy facade)",
    "aluminium (light facade)",
    "zinc (light facade)",
    "panel brick (light facade)",
    "fibercement (light facade)",
    "timber (light facade)",
]

# Roof type
# If this list changes, the ROOF_FINISH_AREA calculation should be modified
ROOF_TYPE = ["flat", "steep"]

# Roof material
ROOF_MATERIAL = ["bitumen", "tiles", "steel", "fibercement", "slate", "zinc"]

# Heating source
HEATING_SOURCE = ["district heating", "electricity", "not decided"]
HEATING_DEFAULT = "district heating"
# endregion

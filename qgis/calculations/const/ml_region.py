# region for ML model

# Dict to find which typology will be passed to the model
# (several typologies are sharing the same Live load and span, so the model was trained with individual group typologies
MODEL_TYPOS = {
    "detached house": 0,
    "terraced house": 0,
    "apartment building": 0,
    "office": 2,
    "school": 2,
    "institution": 2,
    "retail": 4,
    "parking": 2,
    "industry": 1,
    "transport": 5,
    "hotel": 0,
    "hospital": 2
}
# endregion
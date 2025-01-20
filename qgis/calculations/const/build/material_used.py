# Dictionary to define amount of material used based on transformation potential
# Every list is: Foundation, Structure, Envelope, Window, Interior, Technical Systems, Construction, Operational Energy
# Refer to 1.3_Appendiks_BygningTilstand
BUILDING_CONDITION = {
    "new": [1, 1, 1, 1, 1, 1, 1, 1],
    "transformation": [0, .25, .75, 1, 1, 1, .50, 1],
    "renovation": [0, 0, .25, 1, 1, 1, .25, 1],
    "refurbishment": [0, 0, 0, 0, 1, 0, 0, 1],
    "keep": [0, 0, 0, 0, 0, 0, 0, 1],
    "demolition": [1, 1, 1, 1, 1, 1, 0, 0]
}

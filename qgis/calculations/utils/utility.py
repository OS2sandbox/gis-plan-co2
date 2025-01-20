def Average(lst):
    return sum(lst) / len(lst)


def find_key_index(dictionary, key):
    keys = list(dictionary.keys())
    if key in keys:
        return keys.index(key)
    else:
        return -1  # Key not found


def find_key_index_double(dictionary, key1, key2, type_assembly, assembly_version):
    indices = []
    if key1 in dictionary:
        for i, item in enumerate(dictionary[key1]):
            if item == type_assembly:
                indices.append(i)
    key2_index = -1
    if key2 in dictionary:
        for i, item in enumerate(dictionary[key2]):
            if i in indices and item == assembly_version:
                key2_index = i
    return key2_index


def average_of_range(dictionary, start_index, end_index, item_index):
    values = list(dictionary.values())
    if start_index < 0 or end_index >= len(values):
        raise ValueError("Indices are out of range")

    values_in_range = values[start_index:end_index]
    first_items = [item[item_index] for item in values_in_range]
    return sum(first_items) / len(first_items)


class EmissionFactorError(Exception):
    pass


class InvalidYearError(EmissionFactorError):
    def __init__(self, year):
        self.message = f"Invalid construction year: {year}. Maximum allowed is 2075."
        super().__init__(self.message)


class InvalidStudyPeriodError(EmissionFactorError):
    def __init__(self, period):
        self.message = f"Invalid study period: {period}. Maximum allowed is 80 years."
        super().__init__(self.message)


class EmissionFactorNotFoundError(EmissionFactorError):
    def __init__(self, emission_factor):
        self.message = f"Emission factor '{emission_factor}' not found."
        super().__init__(self.message)


class KeyNotFoundError(EmissionFactorError):
    def __init__(self, key):
        self.message = f"Key '{key}' not found in the dataset."
        super().__init__(self.message)


def range_average(dictionary, emission_factor, construction_year, study_period, key):

    if construction_year > 2040:
        raise InvalidYearError(key)
    if study_period > 80:
        raise InvalidStudyPeriodError(key)
    # Normalize the key to handle case-insensitive inputs
    key_lower = key.lower()
    emission_factor_index = None
    if key_lower == "electricity":
        emission_factor_index = 0
    elif key_lower == "district heating":
        emission_factor_index = 1
    elif key_lower == "gas":
        emission_factor_index = 2
    else:
        raise KeyNotFoundError(key)

    # Fetch the emission factors for the selected category
    category_data = dictionary[emission_factor]

    # Calculate the start and end years
    start_year = construction_year
    end_year = construction_year + study_period

    # Collect the emission factors for the given range
    factors = []
    for year in range(start_year, end_year):

        if year in category_data:
            factors.append(category_data[year][emission_factor_index])
        else:
            raise KeyError(
                f"Year '{year}' not found in the dataset for '{emission_factor}'."
            )

    # Return the average of the collected factors
    return sum(factors) / len(factors) if factors else 0

import xlrd
from statistics import stdev, mean

OECD_sheet = xlrd.open_workbook(r'C:\Users\natha\Downloads\ind study project.xlsx')
sheet = OECD_sheet.sheet_by_index(0)

countries = {
    "USA": 43,
    "Japan": 25,
    "NZ": 32,
    "Norway": 33,
    "Iceland": 21
}

data_categories = {
    "Housing": 4,
    "Unemployment": 10,
    "Employment": 9,
    "Earnings": 11,
    "Support": 12,
    "Skills": 14,
    "Years": 15,
    "Pollution": 16,
    "Quality": 17,
    "Stakeholder": 18,
    "Voters": 19,
    "Expectancy": 20,
    "Satisfaction": 22,
    "Health": 21,
    "Hours": 25,
}


def take_stdev_column(column: int):
    """Find the standard deviation of all the data values in a column"""
    stdev_list = []
    for i in range(0, 48):
        if sheet.cell_type(i, column) == 2:  # ignores non-mumbers
            stdev_list.append(sheet.cell_value(i, column))
        else:
            continue
    return stdev(stdev_list)


def take_average_column(column: int):
    """Find the average of all the data values in a column"""
    average_list = []
    for i in range(0, 48):
        if sheet.cell_type(i, column) == 2:  # ignores non-numbers
            average_list.append(sheet.cell_value(i, column))
        else:
            continue
    return mean(average_list)


def country_value_delta_pos_good(country: str, data_value: str):
    """Find the delta between a particular countries value for a
    category and the average, in terms of the number of standard
    deviations from the mean, for data types where a bigger std dev is
    better.
    """
    std_dev_value = take_stdev_column(data_categories[data_value])
    average_value = take_average_column(data_categories[data_value])
    country_value = sheet.cell_value(countries[country],
                                     data_categories[data_value])
    pos_difference = (country_value - average_value) / std_dev_value
    return pos_difference


def country_value_delta_neg_good(country: str, data_value: str):
    std_dev_value = take_stdev_column(data_categories[data_value])
    average_value = take_average_column(data_categories[data_value])
    country_value = sheet.cell_value(countries[country],
                                     data_categories[data_value])
    neg_difference = (average_value - country_value) / std_dev_value
    return neg_difference


stu_1_pos = ["Employment", "Earnings", "Skills"]
stu_2_pos = ["Quality", "Voters", "Health"]
stu_3_pos = ["Earnings", "Support", "Stakeholder", "Satisfaction"]
stu_4_pos = ["Expectancy", "Voters", "Earnings"]
stu_5_pos = ["Satisfaction", "Health"]
stu_1_neg = ["Housing"]
stu_2_neg = ["Pollution"]
stu_3_neg = []
stu_4_neg = ["Unemployment"]
stu_5_neg = ["Pollution", "Housing"]
norway_max_pos = ["Satisfaction", "Quality", "Support"]
norway_max_neg = ["Hours"]

stu_1_dict = {
    "USA": 0,
    "Japan": 0,
    "NZ": 0,
    "Norway": 0,
    "Iceland": 0
}

stu_2_dict = {
    "USA": 0,
    "Japan": 0,
    "NZ": 0,
    "Norway": 0,
    "Iceland": 0
}

stu_3_dict = {
    "USA": 0,
    "Japan": 0,
    "NZ": 0,
    "Norway": 0,
    "Iceland": 0
}

stu_4_dict = {
    "USA": 0,
    "Japan": 0,
    "NZ": 0,
    "Norway": 0,
    "Iceland": 0
}

stu_5_dict = {
    "USA": 0,
    "Japan": 0,
    "NZ": 0,
    "Norway": 0,
    "Iceland": 0
}

norway_dict = {
    "USA": 0,
    "Japan": 0,
    "NZ": 0,
    "Norway": 0,
    "Iceland": 0
}


def stu_list_evaluator(pos: list, neg: list, dictionary: dict):
    for country in countries:
        for data in neg:
            value = country_value_delta_neg_good(country, data)
            dictionary[country] += value
        for data in pos:
            value = country_value_delta_pos_good(country, data)
            dictionary[country] += value


def main():
    stu_list_evaluator(stu_1_pos, stu_1_neg, stu_1_dict)
    stu_list_evaluator(stu_2_pos, stu_2_neg, stu_2_dict)
    stu_list_evaluator(stu_3_pos, stu_3_neg, stu_3_dict)
    stu_list_evaluator(stu_4_pos, stu_4_neg, stu_4_dict)
    stu_list_evaluator(stu_5_pos, stu_5_neg, stu_5_dict)
    print({k: v for k, v in sorted(stu_1_dict.items(), key=lambda item: item[1])})
    print({k: v for k, v in sorted(stu_2_dict.items(), key=lambda item: item[1])})
    print({k: v for k, v in sorted(stu_3_dict.items(), key=lambda item: item[1])})
    print({k: v for k, v in sorted(stu_4_dict.items(), key=lambda item: item[1])})
    print({k: v for k, v in sorted(stu_5_dict.items(), key=lambda item: item[1])})
    # dict sorter is not mine; comes from stack overflow user
    # Devin Jeanpierre


main()

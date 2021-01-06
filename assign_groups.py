#!/bin/python3

import math
import argparse
import os.path
from openpyxl import Workbook, load_workbook

parser = argparse.ArgumentParser(
    description="Group assignment tool created for TDT4140 at NTNU, authored by Ådne Karstad"
)
parser.add_argument(
    "--group_size",
    type=int,
    default=8,
    help="Size of groups",
)
parser.add_argument(
    "-O",
    type=str,
    default="assigned_groups",
    help="Filename of output, e.g. '-O result', will output a file result.xlsx",
)
parser.add_argument("-f", type=str, help="Path to file to process")
args = parser.parse_args()

ATTRIBUTES_MAP = {
    "A": "name",
    "B": "email",
    "C": "provided_email",
    "D": "prog_skill",
    "E": "taken_it1901",
    "F": "grant_anonym",
    "G": "availability",
    "H": "ref_group",
    "I": "eng_group",
    "J": "program",
}


class Student(object):
    """ Object to represent the required fields of students in TDT4140 """

    def __init__(
        self,
        email,
        provided_email,
        name,
        taken_it1901,
        prog_skill,
        availability,
        program,
        grant_anonym,
        ref_group,
        eng_group,
    ):
        """
        email:              provided by logged in user
        provided_email:     written by user
        name:               name
        taken_it1901        student has completed the couse 1901
        prog_skill          self described skill ranging from 1-5
        availability:       for given dates that demos are run
        program:            study program that students are enrolled in
        grant_anonym:       we are allowed to use their deliveries to train assistants
        ref_group:          student wants to join the reference group
        eng_group:          student wants to attend an engelish speaking group

        Student objects should provide attributes to sort on so that we can assign the student to a
        prefered group.

        Should be sorted on if the student have positive grant_anonym, taken_it1901, prog_skill
        """

        self.email = email
        self.provided_email = provided_email
        self.name = name
        self.taken_it1901 = taken_it1901
        self.prog_skill = prog_skill
        self.availability = availability
        self.program = program
        self.grant_anonym = grant_anonym
        self.ref_group = ref_group
        self.eng_group = eng_group

    def get_prog_skill(self):
        return self.prog_skill

    def __getattribute__(self, name: str):
        return super().__getattribute__(name)

    def __repr__(self) -> str:
        return f"{self.provided_email} | skill={self.prog_skill} | grant_anonym={self.grant_anonym}"


def write_column_names(sheet):
    """ Add column name to row 1 """
    for key, attribute in zip(ATTRIBUTES_MAP.keys(), ATTRIBUTES_MAP.values()):
        sheet[f"{key}1"] = f"{attribute}"


def store_students_in_list(workbook, students=None):
    """ Temporal storage for students to be sorted """
    if not students:
        students = []

    for i, student in enumerate(workbook.active.values):
        if i == 0:
            continue
        students.append(
            Student(
                email=student[3],
                name=student[4],
                provided_email=student[5],
                taken_it1901=student[6],
                prog_skill=student[7],
                availability=student[8],
                program=student[9],
                grant_anonym=student[10],
                ref_group=student[11],
                eng_group=student[12],
            )
        )

    return students


def convert_answer_to_bool(answer):
    if answer.lower() == "ja":
        return True
    else:
        return False


def sort_students(students):
    """
    student: list of student objects

    sorted on attributes:
        - stud.grant_anonym
        - stud.taken_it1901
        - stud.prog_skill

    Results in the topological sorting with the highest priority of grant_anonym,
    and least significant of programming skill.
    """

    students = sorted(students, key=lambda stud: stud.prog_skill)
    students = sorted(
        students,
        key=lambda stud: convert_answer_to_bool(stud.taken_it1901),
        reverse=True,
    )
    students = sorted(
        students,
        key=lambda stud: convert_answer_to_bool(stud.grant_anonym),
        reverse=True,
    )

    return students


def write_student_to_sheet(sheet, row, student=None):
    for key, attribute in zip(ATTRIBUTES_MAP.keys(), ATTRIBUTES_MAP.values()):
        sheet[f"{key}{row}"] = f"{getattr(student, attribute)}"


def assign_students_to_groups(
    sheet, number_of_groups, students, group_size=8, offset_value=2
):
    student_index = 0  # Represent the index of student to be appeneded to the group
    group_index = 0  # Represent the group to append a student to

    while len(students) > 0:
        if group_index % number_of_groups == 0:

            group_index = 0
            student_index += 1

        write_student_to_sheet(
            sheet,
            row=calculate_row(group_index, group_size, offset_value, student_index),
            student=students.pop(),
        )

        group_index += 1

    return sheet


def write_group_number_headers(number_of_groups, sheet, group_size=8, offset_value=2):
    """ Adds the group numbers to where rows students should be appended to """
    for group_number in range(number_of_groups):
        sheet[
            f"A{calculate_row(group_number, group_size=group_size, offset_value=offset_value)}"
        ] = f"Group: {group_number + 1}"

    return sheet


def calculate_row(group_index, group_size, offset_value=2, student_index=0):
    """
    Calculate the relevant row
    Arguments:
        - group_index -> what group is currently referenced
        - group_size -> what is the maximum size of groups
        - offset_value -> what row does group 1 start at
        - student_index -> what row in a given group is referenced
    """
    return offset_value + group_index * (group_size + 1) + student_index


def main():

    if not os.path.isfile(args.f):
        raise FileNotFoundError(
            f"Cannot locate the file that you have provided {args.f}"
        )

    loaded_workbook = load_workbook(filename=f"{args.f}")

    group_size = args.group_size
    students = store_students_in_list(workbook=loaded_workbook)
    students = sort_students(students)
    number_of_students = len(students)
    number_of_groups = math.ceil(number_of_students / group_size)

    new_workbook = Workbook()
    sheet = new_workbook.active

    write_column_names(sheet)
    write_group_number_headers(sheet=sheet, number_of_groups=number_of_groups)
    assign_students_to_groups(
        sheet=sheet, number_of_groups=number_of_groups, students=students
    )

    new_workbook.save(filename=f"{args.O}.xlsx")


if __name__ == "__main__":
    main()
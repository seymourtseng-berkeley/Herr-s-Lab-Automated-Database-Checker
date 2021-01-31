import unittest
from Verifier import *
import os, sys
import openpyxl as pyxl

test_loc = os.path.join(sys.path[0], 'test_sample_dataset.xlsx')
sheets = pyxl.load_workbook(test_loc)


class TestMainFunctions(unittest.TestCase):

    def test_make_person(self):
        bio_people = make_person(sheets.worksheets[0])
        self.assertEqual("Allison Dennis", bio_people[0].name)
        self.assertEqual(None, bio_people[3].institution)

    def test_make_department(self):

        departments = [department.name for department in make_department(sheets)]
        self.assertEqual(['Bioengineering',
                          'Biochemistry',
                          'Chemical Engineering',
                          'Mechanical Engineering',
                          'Electrical Engineering',
                          'Computer Science',
                          'Materials Science'], departments)


if __name__ == '__main__':
    unittest.main()
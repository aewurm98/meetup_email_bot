import unittest
import pandas as pd
from io import StringIO
from email_bot import read_students, generate_group, update_selected_students, reset_selections

class TestEmailBot(unittest.TestCase):

    def setUp(self):
        # Sample CSV data
        self.csv_data = StringIO("""Student Email,Number of Selections,Section
student1@example.com,0,A
student2@example.com,1,B
student3@example.com,2,A
student4@example.com,0,B
student5@example.com,1,A
student6@example.com,0,B
""")
        self.students_df = pd.read_csv(self.csv_data)

    def test_read_students(self):
        self.csv_data.seek(0)
        df = read_students(self.csv_data)
        self.assertEqual(len(df), 6)
        self.assertListEqual(list(df.columns), ['Student Email', 'Number of Selections', 'Section'])

    def _test_generate_group(self, use_sections):
        group = generate_group(self.students_df, group_size=3, use_sections=use_sections)
        self.assertEqual(len(group), 3)
        self.assertTrue(all(email in self.students_df['Student Email'].values for email in group))

    def test_generate_group_without_sections(self):
        self._test_generate_group(use_sections=False)

    def test_generate_group_with_sections(self):
        self._test_generate_group(use_sections=True)

    def test_update_selected_students(self):
        selected_students = ['student1@example.com', 'student4@example.com']
        self.csv_data.seek(0)
        updated_df = update_selected_students(self.csv_data, selected_students)
        self.assertEqual(updated_df.loc[updated_df['Student Email'] == 'student1@example.com', 'Number of Selections'].values[0], 1)
        self.assertEqual(updated_df.loc[updated_df['Student Email'] == 'student4@example.com', 'Number of Selections'].values[0], 1)

    def test_reset_selections(self):
        self.csv_data.seek(0)
        reset_df = reset_selections(self.csv_data)
        self.assertTrue((reset_df['Number of Selections'] == 0).all())

if __name__ == '__main__':
    unittest.main()
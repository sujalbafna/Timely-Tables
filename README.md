Overview
This project automates the generation of academic timetables for educational institutions. It schedules lectures and labs for multiple divisions while ensuring conflict-free assignments of subjects and faculties. The program offers user-friendly inputs, displays the timetables in the console, and exports them to a color-coded Excel file.

Features
Custom Inputs: Configure divisions, time slots, subjects, faculties, and labs.
Randomized Scheduling: Ensures balanced subject and faculty distribution.
Conflict-Free Assignments: Prevents faculty overlap across slots.
Export to Excel: Saves timetables in a formatted, color-coded Excel file.
Yellow: Lunch breaks.
Green: Lab sessions.
Blue: Regular subjects.
Requirements
Install the required Python libraries using pip:

bash
Copy code
pip install pandas openpyxl
How to Use
Clone the repository:
bash
Copy code
git clone https://github.com/your-username/timetable-generator.git
Navigate to the project directory:
bash
Copy code
cd timetable-generator
Run the script:
bash
Copy code
python timetable_generator.py
Follow the prompts to input divisions, slots, subjects, and faculties.
Outputs
Console Display: View the generated timetable for each division.
Excel File: The program saves the timetable as a formatted Timetables.xlsx file.
Example Workflow
Input division names, time slots, subjects, and faculties.
View the generated timetable.
Open the saved Timetables.xlsx file to see a neatly formatted timetable with color coding.
Contribution
Feel free to fork this repository, submit issues, or create pull requests for improvements or bug fixes.

License
This project is licensed under the MIT License.

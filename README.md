# automated generation of college feast seating plans

This repository contains Python code and CSV datasets for generating seating plans for College Feast and Subject Dinner events.

The seating plans are created using mixed-integer linear programming models. The models assign diners to seats while considering seat capacity, High Table requirements, gender balance, guest mixing, seating closeness, and other event-specific rules.

## repository structure

code/
- college_feast_xpress.py
- subject_dinner_single_stage.py
- subject_dinner_two_stage.py
- plot.py

csv data/
- all_layouts_seats.csv
- all_layouts_seat_pairs.csv
- layout_summary.csv
- README.txt

data provided/
- original workbook and layout files

outputs csv/
- generated seating plan outputs

## main scripts

- college_feast_xpress.py: generates the College Feast seating plan using a three-stage optimisation model.
- subject_dinner_single_stage.py: generates the Subject Dinner seating plan using a compact single-stage model.
- subject_dinner_two_stage.py: generates the Subject Dinner seating plan using a two-stage model.
- plot.py: creates seating layout figures from the CSV layout data.

## csv data

The CSV files describe the seating layouts used by the optimisation models.

- all_layouts_seats.csv contains one row per seat.
- all_layouts_seat_pairs.csv contains close-seat pairs and their weights.
- layout_summary.csv contains validation counts for each layout.
- README.txt explains how the layout CSV files were generated.

## requirements

The scripts use Python with openpyxl, xpress, and matplotlib.

FICO Xpress must be installed locally with a valid licence to run the optimisation models.

## how to run

Before running the scripts, check the file paths at the top of each Python file and update them if needed.

Example commands:

- python "code/subject_dinner_single_stage.py"
- python "code/subject_dinner_two_stage.py"
- python "code/college_feast_xpress.py"
- python "code/plot.py"

For the College Feast model, these optional commands can also be used:

- python "code/college_feast_xpress.py" --validate-data
- python "code/college_feast_xpress.py" --stage1-only

## outputs

The generated seating plans are saved as CSV files in the outputs csv/ folder.

Expected output files include:

- college_feast_stage1_assignments.csv
- college_feast_main_course_solution.csv
- college_feast_dessert_solution.csv
- subject_dinner_solution_single_stage.csv
- subject_dinner_solution_two_stage.csv

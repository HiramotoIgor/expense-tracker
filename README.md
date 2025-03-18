# Expense Tracker

A simple script that processes and visualizes budget and expense data from an Excel file. It reads budget and expense data from an input Excel file, groups expenses by month and year, and generates a detailed analysis.

## How to Use It

1. **Create an Excel file** named `Expense_Tracker_Template.xlsx` (if using the template file, skip to step 3).
2. **Add Two Sheets**:
   - **Budget Sheet**: Name the sheet `Budget`. It should have the following columns:
     - `Category`: The name of the expense category (e.g., Food, Rent, Transportation).
     - `Monthly Budget`: The budgeted amount for each category.

     Example:
     ```
     |  Category      |  Monthly Budget  |
     |----------------|------------------|
     |  Food          |        300       |
     |  Rent          |       1000       |
     |  Other         |        200       |
     ```

   - **Expenses Sheet**: Name the sheet `Expenses`. It should have the following columns:
     - `Date`: The date of the expense (in `YYYY-MM-DD` format).
     - `Category`: The category of the expense (must match categories in the Budget sheet).
     - `Amount`: The amount spent.

     Example:
     ```
     |    Date       |    Category    |    Amount    |
     |---------------|----------------|--------------|
     | 2023-10-01    |      Food      |      50      |
     | 2023-10-05    |      Rent      |    1000      |
     | 2023-10-10    |      Other     |      30      |
     ```

3. **Save the File**: Save the Excel file in the same directory as the Python script.
4. **Open a Terminal or Command Prompt** and navigate to the directory where the script and Excel file are located.
5. **Run the script** using:
   ```bash
   python expense_tracker.py
6. **Output** The script should generate a new Excel file named `output.xlsx` in the same directory.
   
# Output File (output.xlsx)
The `output.xlsx` file contains:
   - A sheet for each month and year (e.g., 2023, Oct).
   - Key metrics such as total monthly spending, average daily spending, and budget status.
   - A breakdown of spending by category, including percentages and budget comparisons.
   - A pie chart visualizing the distribution of expenses across categories.
     
# Requirements
   - Python 3.x
   - Libraries: pandas, numpy, openpyxl, matplotlib

Install the required libraries using:
	
	pip install pandas numpy openpyxl matplotlib
# To Do
   - Turn it into a GUI app with `Tkinter` or `PyQt` for a friendlier interface.
   - Add machine learning to predict future expenses (using libraries like `scikit-learn`).

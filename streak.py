# TODO:
# 1. ADD A GET LATEST STREAK VALUES FUNCTION TO CALL BEFORE VALIDATE TODAYS COMPLETION / EXPORT / VIEW DETAILS -- done
# 2. ADD VIEW DETAILS FUNCTION -- done
# 3. ADD EXPORT FUNCTION -- done
# 4. CHANGE THE CODE TO CREATE DATA FOLDER FOR SOURCE CSV -- done
# 5. RENAME EVERYTHING (STREAK -> HABIT)
# 6. CHANGE USER PROMPTS TO MATCH THE THEME
# 7. ADD TRY EXCEPT EVERYWHERE FOR MISSING VALUE (INCASE SOMEONE CHANGES CSV AFTER CLICKING ON A CHOICE)

from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook
import pandas as pd
import os
from dateutil.relativedelta import relativedelta
from InquirerPy import inquirer
from pathlib import Path

class Habit:
      
    def __init__(self) -> None:
        # Always Fixed
        self.source_folder = Path("data")
        self.source_csv = "streak_list.csv"
        self.dest_folder = Path.home() / "Documents"
        self.dest_workbook = f"streaks_{datetime.now().date().strftime("%Y%m%d")}_{datetime.now().strftime("%H%M")}.xlsx"
        
        # Loaded Everytime
        self.name = "New Streak"
        self.start_date = datetime.now().date()
        self.duration = 0
        self.end_date = self.start_date
        self.done_today = False
        self.curr_streak = 0
        self.longest_streak = 0
        self.total_successful_days = 0
        self.num_resets = 0
        self.last_updated = datetime.now().date()
        self.headers = ["Streak Name", "Start Date", "Commited Days", "End Date", "Completed Today", "Current Streak", "Longest Streak", "Total Successful Days", "Number of Resets", "Last Updated On"]
    

# """
# =========================================================================================
#     To check if the csv contains a valid dataframe
# =========================================================================================
# """
    def _verify_dataframe(self) -> pd.DataFrame:
        source_file = self.source_folder / self.source_csv
        try:
            self.source_folder.mkdir(parents=True, exist_ok=True)
            df = pd.read_csv(source_file)
            df['Start Date'] = pd.to_datetime(df['Start Date']).dt.date
            df['End Date'] = pd.to_datetime(df['End Date']).dt.date
            df['Last Updated On'] = pd.to_datetime(df['Last Updated On']).dt.date
            if list(df.columns) == self.headers:
                return df
        except Exception as e:
            print(str(e))
            pass
        
        df = pd.DataFrame(columns=self.headers)
        df.to_csv(source_file, index=False)
        return df


# """
# =========================================================================================
#     Loading from csv. Input will be name and the streak loaded will be the one with
#     with start and end dates that encapsulate today's date
# =========================================================================================
# """
    def _load_from_csv(self, name, start_date) -> None:
        df = self._verify_dataframe()
        streak_data_dict = {}
        start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
        
        row = df[(df['Streak Name'] == name) & 
                (df['Start Date'] == start_date)
            ]

        if not row.empty:
            streak_data_dict = row.iloc[0].to_dict()

        df.drop(row.index, inplace=True)
        df.to_csv(self.source_csv, index=False)

        self.name = streak_data_dict['Streak Name']
        self.start_date = streak_data_dict['Start Date']
        self.duration = streak_data_dict['Commited Days']
        self.end_date = streak_data_dict['End Date']
        self.done_today = streak_data_dict['Completed Today']
        self.curr_streak = streak_data_dict['Current Streak']
        self.longest_streak = streak_data_dict['Longest Streak']
        self.total_successful_days = streak_data_dict['Total Successful Days']
        self.num_resets = streak_data_dict['Number of Resets']
        self.last_updated = streak_data_dict['Last Updated On']
        return


# """
# =========================================================================================
#     Getting latest streak values before any operation is to be done on them
# =========================================================================================
# """
    def _get_latest_streak(self, name, start_date):
        self._load_from_csv(name, start_date)
        date_today = datetime.now().date()

        if (date_today - self.last_updated) > timedelta(days=1):
            if self.curr_streak != 0:
                self.num_resets += 1
            self.curr_streak = 0
            self.done_today = False
            self.last_updated = date_today
        return


# """
# =========================================================================================
#     Accept input for custom destination path
# =========================================================================================
# """
    def _add_custom_dest_path(self):
        while True:
            user_input = input("Enter folder path for saving export file. Default is Documents folder (Use '/'): ")
            try:
                if len(user_input) == 0:
                    self.dest_folder.mkdir(parents=True, exist_ok=True)
                    return
                path = Path(user_input).expanduser()
                if path.is_dir():
                    self.dest_folder = path
                    return
                else:
                    print("Wrong path provided. Please provide a valid folder")
            except Exception as e:
                print("Exception: Wrong path provided. Please provide a valid folder")

# """
# =========================================================================================
#     Accepting input for streak name.
#     Name can't be empty, and has to be a string nothing sneaky
# =========================================================================================
# """
    def _add_name(self) -> None:
        while True:
            user_input = input("Enter streak name: ")
            try:

                # Names can't be empty
                if len(user_input) == 0:
                    print("Error: Streak name can't be empty!")
                else:
                    self.name = user_input
                    return
                
            except Exception:

                # Just for safety sake
                print("Invalid format. Please enter a valid name.")


# """
# =========================================================================================
#     Accepting input for streak start date.
# =========================================================================================
# """
    def _add_start_date(self) -> None:
        while True:

            # Start date should only be inputted in the format YYYY-MM-DD.
            user_input = input("Enter streak start date (YYYY-MM-DD): ")
            try:

                # Start date can't be empty
                if len(user_input) == 0:
                    print("Error: Streak start date can't be empty!")
                else:
                    selected_date = datetime.strptime(user_input, "%Y-%m-%d").date()
                    today = datetime.now().date()

                    # Start date can't be before the day streak is being created
                    if selected_date < today:
                        print("Error: Start date can't be before today!")

                    # Start date can't be more than 2 years after the day streak is being created. Just a personal choice.
                    elif relativedelta(selected_date, today).years >=2:
                        print("Date too far in the future! Think about the person you are, not the one you will be.")
                    else:
                        self.start_date = selected_date
                        return
            except ValueError:
                print("Invalid format. Please use YYYY-MM-DD")


# """
# =========================================================================================
#     Accept input for streak duration
# =========================================================================================
# """
    def _add_duration(self) -> None:
        curr_year = self.start_date.year
        is_leap_year = (curr_year % 4 == 0 and curr_year % 100 != 0) or (curr_year % 400 == 0)

        while True:
            user_input = input("How many days are you commiting for?: ")
            try:

                # Duration can't be empty
                if len(user_input) == 0:
                    print("Come on now, commit for atleast 1 day!")
                else:
                    user_input = int(user_input)

                    # Duration can't be negative or 0. What's even the point then
                    if user_input <= 0:
                        print("Come on now, commit for atleast 1 day!")
                    
                    # Duration can't be more than a year. This is to manage the number of worksheets getting created.
                    # Duration can't be greater than 366 at all
                    # Duration can only be greater than 365 (i.e. 366 because greater than that isn't allowed at all) 
                    # if the streak is set on 1 Jan of a leap year
                    if user_input > 366 or user_input > 365 and (not is_leap_year or not (self.start_date.day == 1 and self.start_date.month == 1)):
                        print("Make commitments for one year at a time pls")
                    else:
                        self.duration = user_input
                        self.end_date = self.start_date + timedelta(days = self.duration) - timedelta(days = 1)
                        return
            except ValueError:
                print("Enter valid number of days")


# """
# =========================================================================================
#     Function to confirm is streak has been completed today
# =========================================================================================
# """
    def _validate_todays_completion(self, first_time) -> None:
        if not first_time:
            self._get_latest_streak()
            option_choices = ["Yes, ofcourse :)", "No, I couldn't :(", "Go Back"]
        else:
            option_choices = ["Yes, ofcourse :)", "No, I couldn't :("]
        
        if self.start_date <= datetime.now().date() <= self.end_date:
            done_today = inquirer.select(
                message="Did you complete this task today, btw?",
                choices=option_choices
            ).execute()

            if done_today == "Yes, ofcourse :)":
                self.done_today = True
                self.curr_streak += 1
                if self.curr_streak > self.longest_streak:
                    self.longest_streak = self.curr_streak
                self.total_successful_days += 1
                self.last_updated = datetime.now().date()
            elif done_today == "No, I Couldn't :(":
                self.done_today = False
                if self.curr_streak != 0:
                    self.num_resets += 1
                self.curr_streak = 0
                self.last_updated = datetime.now().date()
        return


# """
# =========================================================================================
#     Function to check if a streak exists with the same name with clashing time period
# =========================================================================================
# """
    def _check_overlapping_time_period(self, df) -> bool:
        df_same_name = df[df['Streak Name'] == self.name]
        if df_same_name.empty:
            return False
        overlapping = df_same_name[(df_same_name['Start Date'] <= self.end_date) & (df_same_name['End Date'] >= self.start_date)]
        return not overlapping.empty


# """
# =========================================================================================
#     Function to add streak to the csv file
# =========================================================================================
# """
    def _add_streak_to_csv(self) -> bool:
        df = self._verify_dataframe()

        overlaps = self._check_overlapping_time_period(df)
        if not overlaps:
            curr_df = pd.DataFrame([[self.name, self.start_date, self.duration, self.end_date, self.done_today, self.curr_streak, self.longest_streak, self.total_successful_days, self.num_resets, datetime.now().date()]], columns=self.headers)
            df = pd.concat([df, curr_df], ignore_index=True)
            df.to_csv(self.source_csv, index=False)
            return True
        print(f"Another streak of the same name exists in between the time period of this one. Choose another name.")
        return False


# """
# =========================================================================================
#     Function that triggers when new streak is to be created.
#     This will call all other functions in order from name input to final addition to excel file
# =========================================================================================
# """
    def create_streak(self) -> None:
        while True:
            self._add_name()
            self._add_start_date()
            self._add_duration()
            self._validate_todays_completion(first_time = True)
            added = self._add_streak_to_csv()
            if added:
                return
    

# """
# =========================================================================================
#     Function to view latest details about a habit
# =========================================================================================
# """
    def view_latest_details(self, name, start_date) -> dict:
        try:
            self._get_latest_streak(name, start_date)
            self._add_streak_to_csv()

            df = pd.read_csv(self.source_csv)
            df['Start Date'] = pd.to_datetime(df['Start Date']).dt.date
            df['End Date'] = pd.to_datetime(df['End Date']).dt.date
            df['Last Updated On'] = pd.to_datetime(df['Last Updated On']).dt.date
            row = df[(df['Streak Name'] == self.name) & 
                    (df['Start Date'] <= datetime.now().date()) & 
                    (df['End Date'] >= datetime.now().date())
                ]
            if not row.empty:
                return row.iloc[0].to_dict()
        except Exception as e:
            return {'Error': str(e)}




# """
# =========================================================================================
#     To verify if the worksheet for the required month exists and create it if not
# =========================================================================================
# """
def verify_worksheet(workbook, month_name, headers):
    
        
    if month_name not in workbook.sheetnames:
        worksheet = workbook.create_sheet(month_name)
    else:
        worksheet = workbook[month_name]

    verify_headers(worksheet=worksheet, headers=headers)


# """
# =========================================================================================
#     To check if the headers in the worksheet are what are required.
#     Basically filtering out anything else if it is put in the excel file manually
# =========================================================================================
# """
def verify_headers(worksheet, headers):

    # Get all the cells in the first row of the worksheet
    actual = [
        worksheet.cell(row=1, column=i).value
        for i in range(1, len(headers) + 1)
    ]

    if actual != headers:

        # Delete all rows where any data may be present
        worksheet.delete_rows(1, worksheet.max_row)

        # Add the actual headers required
        worksheet.append(headers)


# """
# =========================================================================================
#     Function to add streak to the excel file
# =========================================================================================
# """
def add_streaks_to_excel(source_df, dest_path, headers) -> bool:
    workbook = Workbook()

    for _, row in source_df.iterrows():
        start_month = row['Start Date'].replace(day=1)
        end_month = row['End Date'].replace(day=1)

        while (start_month <= end_month):
            month_name = start_month.strftime("%b %Y")
            verify_worksheet(workbook=workbook, month_name=month_name, headers=headers)
            row['Start Date'] = str(row['Start Date'])
            row['End Date'] = str(row['End Date'])
            row['Last Updated On'] = str(row['Last Updated On'])

            workbook[month_name].append(list(row))
            start_month += relativedelta(months=1)

    del workbook['Sheet']
    workbook._sheets.sort(
        key = lambda ws: datetime.strptime(ws.title, "%b %Y")
    )
    workbook.save(dest_path)


# """
# =========================================================================================
#     Function to export the entire list to excel
# =========================================================================================
# """
def export_current_list():
    habit = Habit()
    source_df = habit._verify_dataframe()
    habit._add_custom_dest_path()
    dest_path = habit.dest_folder / habit.dest_workbook
    headers = habit.headers
    add_streaks_to_excel(source_df, dest_path, headers)

# streak = Habit()
# streak.create_streak()
# streak.create_streak()
# streak.create_streak()
# streak_name = input("Streak name?: ")
# streak_date = input("Start Date?: ")
# print(streak.view_latest_details(streak_name, streak_date))
# streak._add_streaks_to_excel()
export_current_list()
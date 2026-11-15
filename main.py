# TODO:
# 1. ADD OPTION TO CREATE, DELETE, UPDATE, EXPORT, VIEW DETAILS FOR A HABIT
# 2. MAKE AN EXE

from terminaltexteffects.effects.effect_blackhole import Blackhole
from terminaltexteffects.effects.effect_wipe import Wipe
from terminaltexteffects.effects.effect_middleout import MiddleOut
from openpyxl import load_workbook, Workbook
from pyfiglet import figlet_format
import os
from datetime import datetime
import pandas as pd
from InquirerPy import inquirer

def terminal_effect_print(effect):
    with effect.terminal_output() as terminal:
        for frame in effect:
            terminal.print(frame)

def introduction():
    title = Blackhole(figlet_format("STREAK - IT🫦", font='ansi_shadow'))
    terminal_effect_print(effect=title)

    subtitle = "Streak it for me baby. Put in as much as you want, I can take it all"
    terminal_effect_print(Wipe(subtitle))

    separator = "============================================================================="
    terminal_effect_print(MiddleOut(separator))


def find_workbook(workbook_path, sheet_name):
    if not os.path.exists(workbook_path):
        workbook = Workbook()
        workbook['Sheet'].title = sheet_name
    else:
        workbook = load_workbook(workbook_path)
    
    if not sheet_name in workbook.sheetnames:
        workbook.create_sheet(sheet_name)
    
    workbook.save(workbook_path)

def check_columns(goals_df):
    required_columns = ["Streak Name", "Start Date", "Commited Days", "End Date", "Completed", "Number of Resets"]
    if list(goals_df.columns) != required_columns:
        goals_df = pd.DataFrame()
    
    for col in required_columns:
        goals_df[col] = None
    return goals_df

def new_streak():
    name = input("Enter streak name: ")
    start_date = input("Enter start date (yyyy/mm/dd): ")
    period = input("How many days are you willing to pledge?: ")

def main():
    try:
        introduction()

        expense_workbook_path = "expenses.xlsx"
        sheet_name = f"{datetime.now().strftime("%B")} {datetime.now().year}"

        find_workbook(expense_workbook_path, sheet_name)

        goals_df = pd.read_excel(expense_workbook_path, sheet_name=sheet_name)
        
        goals_df = check_columns(goals_df)

        curr_action = ""
        while True:
            curr_action = inquirer.select(
                message="Record current streaks, or create a new one?",
                choices=["Current Streaks", "Create New", "Exit"]
            ).execute()
            if curr_action == "Exit":
                break
            new_streak()
            
    except KeyboardInterrupt:
        print("\n\ngoodbye")


if __name__ == "__main__":
    main()

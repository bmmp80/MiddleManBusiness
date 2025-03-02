# # main.py
import CreateDatabase
import CreateFullReportWithPivotTables
import DataInsertGUI


def main():
    CreateDatabase.create_database()
    CreateDatabase.add_algriculture_data()
    DataInsertGUI.setup_gui()
    CreateFullReportWithPivotTables.main()

if __name__ == "__main__":
    main()


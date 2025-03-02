# # main.py
import CreateDatabase
import CreateFullReportWithPivotTables
def main():
    CreateDatabase.create_database()

    import DataInsertGUI

    # Ensure refresh_product_list is only called before mainloop
    dataaddagain.run()

    ExportAndCalculateData.run_all()

if __name__ == "__main__":
    main()


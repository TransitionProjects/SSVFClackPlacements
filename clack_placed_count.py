"""
Create a report showing placements by SSVF into Clackamas County locations.

This script is for automating the processing of the data for the SSVF teams
regular reports to Clackamas County showing the number of vets placed into
said county.
"""
__version__ = "1.0"
__author__ = "David Marienburg"
__maintainer__ = "David Marienburg"

import pandas as pd

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class CreatePlacementReport:
    def __init__(self):
        """
        Initialize the class creating two data frames from the raw report.
        """
        self.raw_report = askopenfilename(
            title="Open the Placement Report v.5c + CoC Location",
            initialdir="//tproserver/Reports/Monthly Reports/"
        )
        self.placements = pd.read_excel(
            self.raw_report,
            sheet_name="Placement Data"
        )
        self.entries = pd.read_excel(
            self.raw_report,
            sheet_name="Entries"
        )


    def process_dataframes(self):
        """
        Create the two processed dataframes.

        Merge the self.entries and self.placements datrames on the client uid
        columns.  Then slice the resulting dataframe so that only participants
        who's Client Location(7690) value contains OR-507.  Remove any rows
        where the entry exit entry date is less than the placement date and
        finally, deduplicate based on client uid.  Return the processed
        dataframe.

        return: a pandas dataframe
        """
        merged = self.placements.merge(
            self.entries,
            on="Client Uid",
            how="inner"
        )

        definite = merged[
            merged["Client Location(7690)"].str.contains("OR-507") &
            merged["Placement Date(3072)"].notna() &
            (merged["Placement Date(3072)"].dt.date >= merged["Entry Exit Entry Date"].dt.date)
        ].drop_duplicates(
            subset="Client Uid"
        )[[
            "Household Uid",
            "Client Uid",
            "Client First Name",
            "Client Last Name",
            "Client Location(7690)",
            "Placement Date(3072)",
            "Reporting Program (TPI)(8748)"
        ]]

        possible = merged[
            merged["Client Location(7690)"].str.contains("OR-507") &
            ~(merged["Client Uid"].isin(definite["Client Uid"]))
        ].drop_duplicates(
            subset="Client Uid"
        )[[
            "Household Uid",
            "Client Uid",
            "Client First Name",
            "Client Last Name",
            "Client Location(7690)",
            "Placement Date(3072)",
            "Entry Exit Entry Date",
            "Reporting Program (TPI)(8748)"
        ]]

        return (definite, possible)

    def save_to_excel(self):
        """
        Save the output of the process_dataframes to an excel spreadsheet.

        return: Boolean
        """
        definite, possible = self.process_dataframes()
        writer = pd.ExcelWriter(
            asksaveasfilename(
                title="Save the Processed SSVF Placement Report",
                initialdir="//tproserver/Reports/Monthly Reports",
                defaultextension=".xlsx",
                initialfile="SSVFClackPlacements(Processed).xlsx"
            ),
            engine="xlsxwriter"
        )

        definite.to_excel(writer, sheet_name="Pts Placed in Clackamas", index=False)
        possible.to_excel(writer, sheet_name="Pts Possibly Placed in Clack", index=False)
        self.placements.to_excel(writer, sheet_name="Raw Placement Data", index=False)
        self.entries.to_excel(writer, sheet_name="Raw Entry Data", index=False)

        writer.save()

if __name__ == "__main__":
    run = CreatePlacementReport()
    run.save_to_excel()

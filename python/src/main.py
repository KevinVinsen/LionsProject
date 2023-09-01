import logging

from pylightxl.pylightxl import utility_columnletter2num
from ruamel.yaml import YAML
from os.path import exists

import pylightxl as xl
import argparse

LOG = logging.getLogger(__name__)


def copy_data(db, sport, sport_columns, sheet_row, row):
    LOG.info(f"Copying data for {row[2]} {row[3]} to {sport['tab']}")
    sport_tab_name = sport["tab"]
    destination_sheet = db.ws(sport_tab_name)
    for column in sport_columns:
        if "from" in column:
            from_column = column["from"]
            to_column = column["to"]
            if from_column == "tab name":
                value = sport_tab_name
            else:
                value = row[utility_columnletter2num(from_column) - 1]

            destination_sheet.update_address(
                address=f"{to_column}{sheet_row[sport_tab_name]}", val=value
            )

    sheet_row[sport_tab_name] += 1


def copy_row(db, yaml_config, sheet_row, row) -> bool:
    if row[2] == "":
        # No first name, so skip
        return False

    # Find the preferred sport
    for sport in yaml_config["sports"]:
        column_number = utility_columnletter2num(sport["school column"])
        if row[column_number - 1] == 1:
            # Found the sport
            copy_data(db, sport, yaml_config["by sport columns"], sheet_row, row)
            return True

    # If we are here we didn't find the sport
    LOG.error(f"Could not find sport for {row[0]} {row[1]} {row[2]} {row[3]}")
    return False


def trim_row_data(row_in):
    row_out = []
    for cell in row_in:
        if cell is None:
            row_out.append("")
        elif isinstance(cell, str):
            row_out.append(cell.strip())
        else:
            row_out.append(cell)

    return row_out


def process_school(db, yaml_config, data_frame, school, sheet_row):
    LOG.info(f"Processing school {school['name']}")
    work_sheet = data_frame.ws(school["tab"])
    reading_heading = True
    count = 0
    for row in work_sheet.rows:
        row = trim_row_data(row)
        if reading_heading:
            if (
                row[0] == "No"
                and row[1] == "School"
                and row[2] == "First name"
                and row[3] == "Surname"
            ):
                reading_heading = False
                LOG.info(f"Found heading row for {school['name']}")

        elif copy_row(db, yaml_config, sheet_row, row):
            count += 1

    LOG.info(f"Processed {count} rows for {school['name']}")


def fill_sheets(db, yaml_config, data_frame):
    sheet_row = {sport["tab"]: 2 for sport in yaml_config["sports"]}
    for school in yaml_config["schools"]:
        # Process school
        process_school(db, yaml_config, data_frame, school, sheet_row)


def build_database(yaml_config, data_frame) -> xl.Database:
    db = xl.Database()

    add_sheets(db, yaml_config)
    fill_sheets(db, yaml_config, data_frame)

    return db


def add_sheets(db, yaml_config):
    # Added the sheets
    for sheet in yaml_config["sports"]:
        LOG.info(f"Create sheet for {sheet['tab']}")
        tab_name = sheet["tab"]
        db.add_ws(tab_name)

        # Add the columns
        for column in yaml_config["by sport columns"]:
            db.ws(tab_name).update_address(
                address=f"{column['to']}1", val=column["name"]
            )


def do_copy(args):
    # Load the YAML
    with open(args.yaml_file, "r") as yaml_file:
        yaml = YAML()
        yaml_config = yaml.load(yaml_file)

    # Load the Excel file
    data_frame = xl.readxl(args.input_file, None)

    # Start writing the output file
    db = build_database(yaml_config, data_frame)

    # Write out the Excel file
    xl.writexl(db=db, fn=f"{args.output_file}")


def user_confirm(question):
    reply = str(input(f"{question} (y/n): ")).lower().strip()
    return reply[0] in ["y", "yes"]


def main(args):
    if not exists(args.input_file):
        LOG.error(f"File {args.input_file} does not exist")
        return

    if exists(args.output_file):
        LOG.warning(f"File {args.output_file} already exists")

        # Confirm overwrite
        if user_confirm("Overwrite file?"):
            LOG.info("Overwriting file")

        else:
            LOG.info("Not overwriting file")
            return

    if not exists(args.yaml_file):
        LOG.error(f"File {args.yaml_file} does not exists")
        return

    do_copy(args)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    parser = argparse.ArgumentParser("Copy the data across")

    parser.add_argument("input_file", help="The input Excel to parse")
    parser.add_argument("output_file", help="The output Excel file")
    parser.add_argument("yaml_file", help="The yaml configuration file")
    main(parser.parse_args())

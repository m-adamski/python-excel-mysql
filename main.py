import os
import sys
import argparse
import openpyxl


def main():
    parser = argparse.ArgumentParser(description="Generate MySQL queries based on input data")
    parser.add_argument("xlsx", type=str, help="Path to the XLSX file containing the input data")
    parser.add_argument("table", type=str, help="The name of the table in the database")
    parser.add_argument("--statement", type=str, default="insert", help="Statement (insert or update)")
    parser.add_argument("--postfix", type=str, help="The content of the postfix that will be added to each generated query")

    # Parse provided arguments
    try:
        arguments = parser.parse_args()

        if arguments.xlsx is not None and arguments.table is not None:
            if os.path.isfile(arguments.xlsx):
                response = parse_xlsx(arguments.xlsx)

                if len(response["headers"]) > 0 and len(response["objects"]) > 0 and len(response["values"]) > 0:
                    if arguments.statement.casefold() == "insert":
                        for curr_value in response["values"]:
                            print(generate_insert(arguments.table, response["headers"], curr_value, arguments.postfix))
                    else:
                        for curr_object in response["objects"]:
                            print(generate_update(arguments.table, response["headers"], curr_object, arguments.postfix))

    except argparse.ArgumentError:
        parser.print_help()
        sys.exit(1)


def parse_xlsx(filename) -> dict:
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active

    response = {
        "headers": [],
        "objects": [],
        "values": []
    }

    for row_num, row in enumerate(sheet.rows):
        objects = []
        values = []

        for cell_num, cell in enumerate(row):
            if row_num == 0:
                response["headers"].append(cell.value)
            else:
                objects.append({response["headers"][cell_num]: cell.value})
                values.append(cell.value)

        if len(objects) > 0:
            response["objects"].append(objects)
            response["values"].append(values)

    return response


def generate_insert(table, headers, values, postfix) -> str:
    headers_content = ", ".join(headers)
    values_list = []

    # Generate values list
    for curr_value in values:
        if isinstance(curr_value, str) and curr_value != "NULL":
            tmp = curr_value.replace("'", "\"")
            values_list.append(f"'{str(tmp)}'")
        else:
            values_list.append(str(curr_value))

    # Join generated items
    values_content = ", ".join(values_list)

    # Check provided postfix
    if postfix is not None:
        postfix = f" {postfix}"
    else:
        postfix = ""

    # Generate INSERT statement
    return f"INSERT INTO {table} ({headers_content}) VALUES ({values_content}){postfix};"


def generate_update(table, headers, objects, postfix) -> str:
    values_list = []

    for curr_object in objects:
        for key, value in curr_object.items():
            if isinstance(value, str) and value != "NULL":
                tmp = value.replace("'", "\"")
                values_list.append(f"{key}='{str(tmp)}'")
            else:
                values_list.append(f"{key}={str(value)}")

    # Join generated items
    values_content = ", ".join(values_list)

    # Check provided postfix
    if postfix is not None:
        postfix = f" {postfix}"
    else:
        postfix = ""

    # Generate UPDATE statement
    return f"UPDATE {table} SET {values_content}{postfix};"


if __name__ == "__main__":
    main()

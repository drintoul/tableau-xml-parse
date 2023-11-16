def get_datasources(workbook):
    """Build XML Tree and find Datasources in Workbook
    """

    import xml.etree.ElementTree as ET

    try:

        tree = ET.parse(workbook)
        datasources = tree.find("datasources").findall("datasource")
        return datasources
    except Exception as e:
        print(e)
        return False


def extract_fields(datasource):
    """Extract fields from datasource
    """

    try:

        sheet = datasource.attrib["caption"][:25]

    except Exception:

        sheet = datasource.attrib["name"][:25]

    fields = []

    for column in datasource.findall("column"):

        try:
            name = column.attrib["name"]
            role = column.attrib["role"]
            datatype = column.attrib["datatype"]

            try:

                calc = column.find("calculation").attrib["formula"]

            except Exception:

                calc = None

            fields.append(name, role, datatype, calc)

        except Exception:

            continue

    return sheet, fields


def process_workbooks(twb):
    """Extract fields from Tableau Workbook
    """

    from files import to_df, to_excel

    for file in twb:

        datasources = get_datasources(file)

        for i, datasource in enumerate(datasources):

            try:

                sheet, fields = extract_fields(datasource)
                df = to_df(fields)
                df["file"] = file.split("\\")[-1:]
                print(f"{sheet} has {df.shape[0]:,} fields")
                to_excel(df, "fields{i:0>3}.xlsx", sheet)

            except Exception as e:
                print(e)
                pass

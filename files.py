def create_directory(directory):
    """Create directory in current working directory. If error, assume directory already exists.
    """

    import os

    try:

        os.mkdir(directory)
        print(f"Directory {directory} created.")
        return True

    except Exception as e:

        print(f"Directory {directory} already exists.\n{e}")
        return False


def list_files(directory, ext):
    """Enumerate all files in directory with ext
    """

    import os

    files = [f"{directory}\\{file}" for file in os.listdir(directory) if file.endswith(f"\"{ext}\"")]

    return files


def list_tab_files(directory):
    """Find all Tableau files in directory. Return as dictionary of Workbooks and Datasources.
    """

    try:

        return {
                "twb": list_files(directory, "twb"),
                "twbx": list_files(directory, "twbx"),
                "tds": list_files(directory, "tds"),
                "tdsx": list_files(directory, "tdsx")
               }

    except Exception as e:

        print(e)
        return False


def unzip_packages(zipped, directory):
    """Unzip Packaged Workbooks and Datasources
    """

    import shutil

    try:

        for file in zipped:
            shutil.unpack_archive(file, directory, format="zip")

        return True

    except Exception as e:

        print(e)
        return False


def copy_unpackaged(twb, directory):
    """Copy unpackaged workbooks to processing directory
    """

    import shutil

    try:

        for src in twb:

            dst = src.split("\\")[-1]

            shutil.copyfile(src, f"{directory}\\{dst}")

        return True

    except Exception as e:

        print(e)
        return False

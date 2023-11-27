# tableau
<p>Unzip all packaged workbooks, copy to a single directory, then systematically parse each Tableau Workbook XML file looking for fields and calculations. Output as Excel.

<p>In addition to main.py, there are three files with specific functions:

<h2>file</h2>
<ul>
  <li>create_directory
  <li>list_files
  <li>list_tab_files = calls list_files
  <li>unzip_packages = unzip packaged Tableau workbooks
  <li>copy_unpackaged
</ul>

<h2>excel</h2>
<ul>
  <li>to_df = creates Pandas DataFrame from fields
  <li>to_excel = creates Microsoft Excel workbook from CSV
  <li>summarize_excel = merge all Excel workbooks into single file
  <li>colorize_and_format = sets sheet tab color, column widths and word wraps formula column
</ul>

<h2>tableau</h2>
<ul>
  <li>get_datasources = enumerate datasources found in workbook; note that parameters are also a 'datasource'
  <li>extract_fields = get attributes for each field
  <li>process_workbooks = generate Excel file for each Tableau workbook with sheet tab for each datasource and row for each field
</ul>

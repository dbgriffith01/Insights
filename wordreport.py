import win32com.client as client
from insights import (create_macro_template, get_files_in_dir,
                      write_macro_to_file, close_mtb)
from os.path import join, basename, splitext


path_to_graphs = r'C:\Insights\Graphs'
path_to_data = r'C:\Insights\Data'
path_to_reports = r'C:\Insights\Reports'
excel_files = get_files_in_dir(path_to_data, suffix='xlsx')

def create_report(word_app, var_name, path_to_report, path_to_graphs):
    doc = word_app.Documents.Open(join(path_to_report,'template','quality_report_template.docx'))
    varname = doc.Bookmarks("variable_name").Range.Text = var_name
    cchart = doc.Bookmarks('cc_chart').Range.InlineShapes.AddPicture(join(path_to_graphs, 'ICHAR_' + var_name) + '.jpg')
    tchart = doc.Bookmarks('tol_int').Range.InlineShapes.AddPicture(join(path_to_graphs, 'TOLIN_' + var_name) + '.jpg')
    doc.SaveAs(join(path_to_report, var_name + '.docx'))
    doc.Close()

with open(r'C:\Insights\Macros\spc.txt', 'r') as f:
    macro = f.read()
# launch Minitab
mtb = client.Dispatch('Mtb.Application')
ui = mtb.UserInterface
project = mtb.ActiveProject
commands = project.Commands
execute = project.ExecuteCommand

ui.DisplayAlerts = False

for file in excel_files:
    commands.Delete()
    # run Minitab commands on that Excel File
    execute(macro.format(filename=file))
    # save the graph
    for i in range(1,commands.Count+1):
        for output in commands.Item(i).Outputs:
            if output.OutputType==0: # 0 Output type is a graph
                variable_name = splitext(basename(file))[0]
                fname = join(path_to_graphs, commands.Item(i).CommandLanguage[:5] + '_' + variable_name)
                output.Graph.SaveAs(fname, True, 1)

word = client.Dispatch('Word.Application')

for file in excel_files:
    variable_name = splitext(basename(file))[0]
    create_report(word, variable_name, path_to_reports, path_to_graphs)

word.Quit
del word
close_mtb

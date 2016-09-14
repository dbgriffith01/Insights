from os import listdir, system
from os.path import join
import psutil
# This is code to be used for the Minitab Insights Conference
# Must of these functions are specific
def get_command_language(commands):
    mtb_commands = []
    for i in range(1, commands.Count + 1):
        mtb_commands.append(commands.Item(i).CommandLanguage)
    return mtb_commands

def create_macro_template(mtb_coms):
    commands = get_command_language(mtb_coms)
    mtb_commands = []
    for i in range(len(commands)):
        command = commands[i]
        if i==0:
            command = command.split(';')
            command[0]="WOPEN '{filename}'"
            del command[-1]
            command = ';\n'.join(command) + '.'
        else:
            command = command.split(';')
            command = ';\n'.join(command)
        mtb_commands.append(command)
    return '\n'.join(mtb_commands)

def write_macro_to_file(macro_file, macro):
    with open(macro_file, 'w') as f:
        f.write(macro)
    print('File Saved!')

# from stackoverflow
# http://stackoverflow.com/questions/9234560/find-all-csv-files-in-a-directory-using-python
def get_files_in_dir( path_to_dir, suffix=".xlsx" ):
    filenames = listdir(path_to_dir)
    return [join(path_to_dir,filename) for filename in filenames if filename.endswith( suffix )]


def get_mtb_processes():
    mtb_pids=[]
    pids = psutil.pids()
    for pid in pids:
        if psutil.Process(pid).name() == 'Mtb.exe':
            mtb_pids.append(pid)
    return set(mtb_pids)

def close_mtb():
    result = system('taskkill /im mtb.exe /F')
def launch():
    before_procs = get_mtb_processes()
    while True:
        mtb = client.Dispatch('Mtb.Application')
        after_procs = get_mtb_processes()
        return mtb, list(after_procs)[0]

from colorama import init, Fore, Back, Style
import threading, sys, os, argparse

init() # color in terminal
join = os.path.join
walk = os.walk
splitext = os.path.splitext
system = os.system
root_dir = "."
writing_to_console = False
calling_executable = False
consoles = []
threads = []


parser = argparse.ArgumentParser(description="convert all documents into pdf's recursively from the directory this script is being runned.")
parser.add_argument("-v", "--verbose", action="store_true",help="increase output verbosity")
parser.add_argument("-o", "--overwrite", action="store_true",help="overwrite old files")
args = parser.parse_args()

def generatePDF(cwd,directory, file):
    global writing_to_console
    if(directory != "" and file != ""):
        global calling_executable
        current_instance = os.popen("OfficeToPDF \"" + os.getcwd() + join(directory, file)[1:] + "\"")
        consoles.append(current_instance)
        reply = current_instance.read()
        
        if "Input file can not be handled. Must be Word, PowerPoint, Excel, Outlook, Publisher, XPS or Visio" in reply:
            # while someone is writing_to_consolem wait
            while (writing_to_console):
                 pass # wait
            writing_to_console = True
            sys.stdout.flush()
            print(Fore.RED + format("[ERROR]:","10") + Style.RESET_ALL + reply)
            sys.stdout.flush()
            writing_to_console = False
        elif "":
            # no reply -> success
            if args.verbose:
                writing_to_console = True
                sys.stdout.flush()
                # print(Fore.GREEN + format("[generated]:\t","20") + Style.RESET_ALL + splitext(file)[0] + ".pdf")
                sys.stdout.flush()
                writing_to_console = False    
        else:
            # while someone is writing_to_consolem wait
            if args.verbose:
                while (writing_to_console):
                    pass # wait
                writing_to_console = True
                sys.stdout.flush()
                print(Fore.GREEN + format("[DONE]: ","10") + Style.RESET_ALL + reply + format(splitext(file)[0] + ".pdf","22") + format("converted successfully","22"))
                sys.stdout.flush()
                writing_to_console = False
            

for directory, subdirectories, files in walk(root_dir):
    for file in files:
        if "~" in file or ".txt" in file or ".py" in file:
            if args.verbose:
                while writing_to_console:
                    pass #wait
                writing_to_console = True
                sys.stdout.flush()
                print(Fore.YELLOW + format("[SKIP]:","10") + Style.RESET_ALL + format(file,"22") + format("not a target extention","22"))
                sys.stdout.flush()
                writing_to_console = False
                continue

        if ".docx" in file or ".rtf" in file:
            if not (args.overwrite):
                if(len(files) > 1):
                    # multiple files in directory
                    filename1 = splitext(join(directory, files[0]))[0]
                    fileext1 = splitext(join(directory, files[0]))[1]
                    filename2 = splitext(join(directory, files[1]))[0]
                    fileext2 = splitext(join(directory, files[1]))[1]
                    
                    if filename1 == filename2 and fileext1 == ".pdf" or fileext2 == ".pdf":
                        if args.verbose:
                            while writing_to_console:
                                pass
                            writing_to_console = True
                            sys.stdout.flush()
                            print(Fore.YELLOW + format("[SKIP]: ","10") + Style.RESET_ALL + format(file,"22") + format("target file already exsist","22"))
                            sys.stdout.flush()
                            writing_to_console = False
                            continue
            else:
                # It's target extention, there is no pdf found with the same name. Generate the file 
                cwd = os.getcwd()               
                new_tread = threading.Thread(target=generatePDF,args=[cwd, directory,file])
                threads.append(new_tread)
                try:
                    threads[-1].start()
                except:
                    pass
                while writing_to_console:
                    pass
                writing_to_console = True
                # print(Fore.GREEN + format("[generating]:\t","22") + Style.RESET_ALL + splitext(file)[0] + ".pdf")
                writing_to_console = False

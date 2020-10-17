import time
import pandas as pd
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
pd.options.mode.chained_assignment = None
file_folder="\\data_files\\"
output_files="\\Output_files\\"

 

#function to combine files
def combiner():
    
    #create data_files and output files folders
    try:
        os.mkdir(os.getcwd()+output_files)
        os.mkdir(os.getcwd()+file_folder)
    except:
        pass

 

    major_df=pd.DataFrame([])
    desti_name=os.getcwd()+output_files+"\major.csv"
    no_closed_faults=[]

 

    if os.path.isfile(os.getcwd()+file_folder+"\major.csv"):
        #remove current major csv
        os.remove(os.getcwd()+file_folder+"\major.csv")

 

    for file in os.listdir(os.getcwd()+file_folder):
        if file is not "major.csv":
            file_content=pd.read_excel(os.getcwd()+file_folder+file)
            #drop unnamed columns...removes unclean portions of excel
            file_content.drop(file_content.columns[file_content.columns.str.contains('unnamed',case = False)],axis = 1, inplace = True)
            #criteria for finding open faults and closed faults header rows and use them to pick open and closed faults
            fault_type=file_content[file_content.isnull().sum(axis=1)>(len(file_content.columns)-2)].index#to pick indices of open,closed and changes headers
                
            for x in list(fault_type):
                change_loc=0;
                if "ope" in str(file_content.loc[[x]].dropna(axis=1)).lower():
                    open_loc=x
                elif "clo" in str(file_content.loc[[x]].dropna(axis=1)).lower():
                    close_loc=x
                elif "cha" in str(file_content.loc[[x]].dropna(axis=1)).lower():
                    change_loc=x
            try:
                #pick each category
                if(close_loc>open_loc):
                    opened_df=file_content.loc[open_loc+1:close_loc-1]
                if(change_loc>close_loc):
                    closed_df=file_content.loc[close_loc+1:change_loc-1]
                else:
                    closed_df=file_content.loc[close_loc+1:]
                #dd status column to indentify state of faults    
                opened_df["Status"]="Open"
                closed_df["Status"]="Closed"
                #concat open and closed faults
                #combined_faults=pd.concat([opened_df,closed_df])
                #concat to main data frame
                major_df=pd.concat([major_df,closed_df])
                #reset index variables so that next excle can be read with new index values
                del open_loc,change_loc,close_loc
            except:
                no_closed_faults.append(file)
                closed_df=pd.DataFrame(no_closed_faults)
                closed_df.to_csv(os.getcwd()+output_files+"No_closed_faults.csv",index=False)
                
    #reset index and save to major.csv
    major_df.reset_index().drop(columns=["index"]).to_csv(desti_name)
    
    
    
class  file_handler(FileSystemEventHandler):
    

 

    def dispatch(event):
        if(event.event_type=="created"):
            file_type=event.src_path.split(".")[-1]
            if "xl" in file_type:
                print(event.src_path+"  added into folder")
                combiner()

 

        
file_observer=Observer()
file_observer.schedule(file_handler,os.getcwd()+file_folder,recursive=True)

 

file_observer.start()
try:
    print("Script watching for new files into folder...")
    while True:
            time.sleep(5)
except KeyboardInterrupt:
        file_observer.stop()
        file_observer.join()

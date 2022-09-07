
#!/usr/bin/env python3
import sys, glob, pandas as pd,csv,os

def main():
    #get all files with extension .xlsb
    if len(sys.argv) > 0:
        try:
            input_dir = sys.argv[1]
        except Exception as e:
            print("input dir not provided \n usage: ./",os.path.basename(__file__), ' <INPUT_DIR_FULL_PATH/>')
            return
    xlsb_files = glob.glob(sys.argv[1] + '/*.xlsb')
    header,out_data = [] , []
    #iterate on all files
    for file_path in xlsb_files:
        print("Processing :",file_path)
        #read file
        df = pd.read_excel(file_path, engine='pyxlsb',  index_col=None, header=None,na_filter=0)
        #trnapose
        for i, row in df.iterrows():
        #consider the row till here as columns
            if i  < 12:
                header.append(row[0])
            else:
                out_row = []
                for j in range (15):
                    #append the data
                    out_row.append(row[j])
                out_data.append(out_row)
        #open output file
        with open(file_path + '.csv', 'w', newline='\n') as file:
            writer = csv.writer(file)
            writer.writerows([header])
            writer.writerows(out_data)
            print("Writitng output to file ",file_path + '.csv')
        #rename the file for now
        os.rename(file_path, file_path + '_done')
    return
#run the program
if __name__ == "__main__":
    main()
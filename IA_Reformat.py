import os
import csv
import openpyxl
import argparse

def main(fname, snum):
    id_dict = dict()
    main_titles = ["LUI", "GivenName", "FamilyName", "StudentID"]
    repeated_titles = ["Subject Name","IA1 Result","IA2 Result","IA3 Result"]

    orig_path = "./" + fname
    
    assert os.path.isfile(orig_path) == True, "Error, the file name given does not exist in this folder, please check the spelling and try again"
    
    wb = openpyxl.load_workbook(orig_path)
    ws = wb.active

    data = list(ws.iter_rows(values_only=True))
    data = [list(lines) for lines in data][1:]

    full_title = main_titles + (repeated_titles)*snum

    with open("formattedATARPredictions.csv", "w", newline='', encoding='utf-8') as new_file:
        writer = csv.writer(new_file)
        writer.writerow(full_title)
        for lines in data:
            if lines[1] not in id_dict:
                id_dict[lines[1]] = lines[1:5] + [lines[0]] + lines[5:]
                # for dat in lines[5:]:
                #     id_dict[lines[1]] += [f"=\"{dat}\""]
            else:
                id_dict[lines[1]] += [lines[0]] + lines[5:]
                # for dat in lines[5:]:
                #     id_dict[lines[1]] += [f"=\"{dat}\""]
        
        for key, value in id_dict.items():
            writer.writerow(value)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Script to reformat ATAR results document",
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)

    parser.add_argument("-fname", type=str, nargs='?', const="", default="", 
                        help="Filename of Excel file to reformat")
    parser.add_argument('-smax', type=int, nargs='?', const=7, default=7, 
                        help='an integer for the accumulator')

    args = parser.parse_args()
    assert args.fname != "", "Please provide a filename argument by running the following command: python script.py -fname 'your_filename'"
    main(args.fname, args.smax)

    print("Finished Successfully")
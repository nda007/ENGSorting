import os
import csv
import openpyxl
import argparse

def main(fname, snum):
    id_dict = dict()
    formatted_text = []
    main_titles = ["SId", "PolyRank", "TEA", "TEAPercent", "ATAR"]
    repeated_titles = ["Id","Type","LetterGrade","RawResult","ScaledResult", "Used"]

    orig_path = "./" + fname
    
    assert os.path.isfile(orig_path) == True, "Error, the file name given does not exist in this folder, please check the spelling and try again"
    
    wb = openpyxl.load_workbook(orig_path)
    ws = wb.active

    data = list(ws.iter_rows(values_only=True))
    data = [list(lines)[1:] for lines in data][1:]

    labelled_subjects = []
    for i in range(1, snum):
        labelled_subjects += ["Sub" + str(i) + title for title in repeated_titles]
    full_title = [main_titles[0]] + labelled_subjects + main_titles[1:]
    formatted_text.append(full_title)

    for lines in data:
        if lines[0] not in id_dict:
            id_dict[lines[0]] = lines[1:]
        else:
            del id_dict[lines[0]][-4:]
            id_dict[lines[0]] += lines[1:]

    for key, value in id_dict.items():
        full_line = [key] + value[:-4]
        while len(full_line) != len(full_title)-4:
            full_line.append("")
        full_line += value[-4:]

        formatted_text.append(full_line)

    with open("formattedATARPredictions.csv", "w", newline='', encoding='utf-8') as new_file:
        writer = csv.writer(new_file)
        for row in formatted_text:
            writer.writerow(row)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Script to reformat ATAR results document",
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)

    parser.add_argument("-fname", type=str, nargs='?', const="", default="", 
                        help="Filename of Excel file to reformat")
    parser.add_argument('-smax', type=int, nargs='?', const=7, default=7, 
                        help='an integer for the accumulator')

    args = parser.parse_args()
    assert args.fname != "", "Please provide a filename argument by running the following command: python script.py -fname 'your_filename'"
    main(args.fname, args.smax+1)

    print("Finished Successfully")

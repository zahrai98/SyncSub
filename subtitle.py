import webvtt
from datetime import datetime, timedelta
import pandas as pd
import argparse
from pathlib import Path
import openpyxl

# datetime format
dt_format = "%H:%M:%S.%f"

def parse_argument():
    """
    Returns the tranclation path and subtitle path
    """
    tranclation, subtitle = None, None
    parser = argparse.ArgumentParser()
    parser.add_argument("SubtitlePath", help="path for vtt subtitle file")
    parser.add_argument("TranclationPath",help="path for vtt  tranclation file")
    args = parser.parse_args()
    if Path(args.SubtitlePath).exists() and Path(args.TranclationPath).exists():
        subtitle, tranclation = args.SubtitlePath, args.TranclationPath
        return subtitle, tranclation
    else:
        print("Index or others folder does not exist or")
        print("You have not entered the quotation mark for paths")
        print("command:separator.py 'SubtitlePath' 'TranclationPath'")       
    return None
    

def merge_similar_times(sub_vtt, counter=1):
    """
    Rrturns the subtitle with non-repetitive time
    Args:
        sub_vtt(class webvtt.webvtt.WebVTT)
        counter(int) 
    Returns:
        subtitle(list) : list of [start time , end time , sentence] of each line
    """
    subtitle = []
    while(counter < len(sub_vtt)):
        sentence = sub_vtt[counter].text
        if counter + 1 >= len(sub_vtt):
            subtitle.append([sub_vtt[counter].start, sub_vtt[counter].end, sub_vtt[counter].text])
            counter += 1
            continue
        while(sub_vtt[counter].start == sub_vtt[counter+1].start and
              sub_vtt[counter].end == sub_vtt[counter+1].end):
            counter += 1
            sentence += " " + sub_vtt[counter].text
            if counter + 1 >= len(sub_vtt):
                break
        subtitle.append([sub_vtt[counter].start, sub_vtt[counter].end, sub_vtt[counter].text])
        counter += 1
    return subtitle


def search_analogous_sentence(tranclation_sub, subtitle, counter=0,start_search=3, search_number=10, diff_second=2.5):
    """
    Rrturns a list with analogous sentence of subtitle and tranclation
    Args:
        tranclation_sub(list) : list of [start time , end time , sentence] of tranclation
        subtitle(list) : list of [start time , end time , sentence] of subtitle
    Returns:
        result_subtitle(list) : list [subtitle index, tranclation index, start time , subtitle sentence, tranclation sentence] of analogous line
    """
    result_subtitle = []
    while(counter < len(subtitle)):
        sub_start_time = datetime.strptime(subtitle[counter][0], dt_format)
        sub_sentence = subtitle[counter][2]
        tranc_index = -1 
        if start_search + search_number < len(tranclation_sub):
            end_search = start_search + search_number
        else:
            end_search = len(tranclation_sub)
        for i in range(start_search-3,end_search):
            tranc_start_time = datetime.strptime(tranclation_sub[i][0], dt_format)
            if abs(tranc_start_time - sub_start_time) < timedelta(seconds=diff_second) :
                if tranc_index == -1 :
                    tranc_index = i
                elif (tranc_start_time - sub_start_time) < abs(datetime.strptime(tranclation_sub[tranc_index][0], dt_format) - sub_start_time):
                    tranc_index = i          
        if tranc_index == -1 :
            start_search +=1
        else:
            start_search = tranc_index + 1
            if  len(result_subtitle) != 0 and tranc_index == result_subtitle[-1][1]:
                result_subtitle[-1][3] += sub_sentence
            else:
                result_subtitle.append([counter,tranc_index,sub_start_time,sub_sentence,tranclation_sub[tranc_index][2]])        
        counter += 1
    return result_subtitle


def merge_short_sentence(short_subtitle, result_subtitle, short_sub_index):
    """
    Rrturns list that it's short sentence merged
    Args:
        short_subtitle(list) : list of [start time , end time , sentence]
        result_subtitle(list) : result_subtitle(list) : list [subtitle index, tranclation index, start time , subtitle sentence, tranclation sentence]
        short_sub_index(int) : the index of subtitle that we want merge in the result_sutitle (subtitle:0  tranclation:1)
    Returns:
        result_subtitle(list) : modified list [subtitle index, tranclation index, start time , subtitle sentence, tranclation sentence]
    """
    for i in range(1,len(result_subtitle)):
        index_previous_row = result_subtitle[i-1][short_sub_index]
        index_row = result_subtitle[i][short_sub_index]
        if index_row - index_previous_row -1 >=1 :
            for j in range(index_previous_row+1, index_row):
                result_subtitle[i-1][4] += short_subtitle[j][2]          
    return result_subtitle


def make_exel_form(subtitle_result):
    """
    Rrturns result_subtitle in Excel form
    Args:
        result_subtitle(list) : list [subtitle index, tranclation index, start time , subtitle sentence, tranclation sentence]
    Returns:
        result_subtitle(list) : modified list [start time(change format) , subtitle sentence, tranclation sentence]
    """
    for i in range(len(subtitle_result)):
        subtitle_result[i].pop(0)
        subtitle_result[i].pop(0)
        start_time = subtitle_result[i][0]
        start_second = start_time.second
        start_minute = start_time.minute
        if start_minute != 0 :
            time = "{}:{}".format(start_minute,start_second)
        else:
            time = "{} s".format(start_second)
        subtitle_result[i][0] = time
    return subtitle_result


def create_exel_file(exel_name, sheet_name, column, result_subtitle):
    """
    create exel file
    Args:
        exel_name(str) : name of exel
        sheet_name(str) : name of sheet in exel
        column(list) : names of columns 
        result_subtitle(list) : list [start time , subtitle sentence, tranclation sentence]
    """
    exel_name += ".xlsx"
    df = pd.DataFrame(result_subtitle, columns=column)
    writer = pd.ExcelWriter(exel_name, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=sheet_name,index=False)
    writer.save()
    

def run_program(subtitle_path, tranclation_path):
    tranclation_text = webvtt.read(tranclation_path)
    subtitle_text = webvtt.read(subtitle_path)          
    tranclation = merge_similar_times(tranclation_text)
    subtitle = merge_similar_times(subtitle_text) 
    result_subtitle = search_analogous_sentence(tranclation, subtitle)
    result_subtitle = merge_short_sentence(subtitle, result_subtitle, 0)   
    result_subtitle = merge_short_sentence(tranclation, result_subtitle, 1)    
    result_exel_form_subtitle = make_exel_form(result_subtitle)   
    create_exel_file("test", "test_sheet", ['Time','Translation', 'Subtitle'], result_exel_form_subtitle)  


if __name__ == "__main__":
    subtitle_path, tranclation_path = parse_argument()
    if subtitle_path and tranclation_path:
        run_program(Path(subtitle_path),Path(tranclation_path))      
        

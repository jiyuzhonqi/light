import pandas as pd
import json
import argparse
import os
import shutil
from pathlib import Path

current_directory = os.path.dirname(os.path.abspath(__file__))
print(current_directory)
zip_folder_path = current_directory+"/output"
zip_out_path = current_directory
song_list_json_path = current_directory+"/output/songlist.json"
project_output_dir = current_directory+"/output"
project_res_dir = current_directory+"/res"
project_dir = current_directory

def excel_to_json(excel_path,json_path='output.json'):
    try:
        color_dic = handle_color(excel_path)
        duanluo_dic = handle_duanluo(excel_path)
        dict_new1 = {}
        if not color_dic:
            dict_new1 = {"sectionData":duanluo_dic}
        else:
            dict_new1 = {"colorData":color_dic,"sectionData":duanluo_dic}
        # dict_new1 = {"sectionData":duanluo_dic}
        json_data = json.dumps(dict_new1)
        # json_data = aes_encode(json_data)
        with open(json_path,'w') as f:
            f.write(json_data)
        print(f"path:{excel_path}")
    except Exception as e:
        print(f"failed {str(e)} path:{excel_path}")


def zip():
    folder_path = zip_folder_path
    output_path = zip_out_path
    shutil.make_archive(output_path, 'zip', folder_path)

def write_song_list_text():
    current_dir = project_output_dir
    file_names = all_input_files(current_dir)
    song_list_text_path = song_list_json_path
    song_list = ""
    song_array = []
    for file_name in file_names:
         if file_name == '.DS_Store' or file_name == 'songlist.json' :
             continue
         file_name_without_extension = os.path.splitext(file_name)[0]
         song_list += file_name_without_extension + '\r\n'
         song_array.append(file_name_without_extension) 
    song_array_json = json.dumps(song_array)
    print(song_array_json)
    with open(song_list_text_path,'w') as f:
            f.write(song_array_json)
            f.close()

# 颜色
def handle_color(excel_path):
    df = pd.read_excel(
            excel_path,
            sheet_name='颜色',
            engine='openpyxl'
        )
    if df.isnull().values.all():
        return {}
    df_cleaned = df.dropna()
    df_cleaned = df_cleaned.dropna(axis=1)
    timestamp_change(df_cleaned)
    dict = df_cleaned.to_dict(orient='records')    
    light_color_dict = {"lightColors":dict}
    return light_color_dict

# 段落
def handle_duanluo(excel_path):
    df = pd.read_excel(
            excel_path,
            sheet_name='段落',
            dtype={'time':str},
            engine='openpyxl'
        )
    df_cleaned = df.dropna()
    df_cleaned = df_cleaned.dropna(axis=1)
    timestamp_change(df_cleaned)
    dict = df_cleaned.to_dict(orient='records')  
    section_dict = {"sectionList":dict}  
    return section_dict
    
#转换时间格式
def timestamp_change(excel): 
     excel['time'] = "1970-01-01 "+excel['time'] 
     for index, row in excel.iterrows():
         excel.at[index, 'time'] = pd.to_datetime(excel.at[index, 'time'],format="%Y-%m-%d %H:%M:%S:%f").timestamp() * 1000
     excel['time'] = excel['time'].astype("int")

def file_exists(file_path):
    return os.path.exists(file_path)

def all_input_files(path_name):
    folder_path = Path(path_name)
    file_names = [f.name for f in folder_path.iterdir() if f.is_file()]
    
    return file_names

def traverse_folder(folder_path):
    path = Path(folder_path)
    for file_path in path.glob('**/*'):
        if file_path.is_file():
            print(file_path)

def transform_res_file():
    current_dir = project_dir
    path = current_dir+"/"+"res"
    file_names = all_input_files(path)
    output_path = current_dir+"/"+"output"
    for file_name in file_names:
         if file_name == '.DS_Store':
             continue
         file_name_without_extension = os.path.splitext(file_name)[0]
         json_output_file_name = file_name_without_extension + ".kgo"
         xlsx_input_file_path = path + "/" + file_name
         json_output_file_path = output_path + "/" + json_output_file_name
         excel_to_json(xlsx_input_file_path,json_output_file_path)
    try:
        os.remove(output_path+'/'+".DS_Store")
    except OSError:
        print(" rm -rf dstore fialed")
    write_song_list_text()
    zip()
    
def test():
    input = "/Users/yanghuafeng/Downloads/color_table.xlsx"
    output = "/Users/yanghuafeng/Downloads/color_table.kgo"
    excel_to_json(input,output)

def initPath():
    current_dir = os.getcwd()
    print(current_dir)
    global project_output_dir
    global project_dir
    global song_list_json_path
    global zip_out_path
    global zip_folder_path
    global project_res_dir
    project_dir = current_dir
    project_output_dir = current_dir + "/output"
    shutil.rmtree(project_output_dir)
    os.mkdir(project_output_dir)
    project_res_dir = current_dir +"/res"
    song_list_json_path = project_output_dir + "/songlist.json"
    zip_out_path = current_dir + "/lightfeature/"
    zip_folder_path = project_output_dir

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i","--input",help="input name",type=str,default="/Users/yanghuafeng/Downloads/color_table.xlsx")
    parser.add_argument(
        "-o",
        "--output",
        help="Output directory",
        type=str,
        default="",
    )
    parser.add_argument(
        "-t",
        "--type",
        help="type",
        type=str,
        default="",
    )
    args = parser.parse_args()
     # print(f"type: {args.type}")

    # initPath()
    transform_res_file()


if __name__ == "__main__":
    main()
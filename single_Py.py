import tkinter as tk
from tkinter import filedialog, messagebox
from loguru import logger
import pandas as pd
import configparser
import time, os


# 将原有的函数保留在这里，以便在GUI中使用

def get_max_source(data: dict, c):
    return int(data.get(c))


def get_user_config(data: dict, c):
    return int(data.get(c)) / 100

def compute_excel_function(file, config_file, template_option):
    result_df = pd.DataFrame(
        columns=['学校', '年级', '班级', '科目', '学生总数', '计算人数', '过差', '及格', '优秀', '平均分'])
    ret_file_name = str(time.time()) + "RET.xlsx"
    logger.info("正在处理上传文件{}，模板选项为{}".format(file, template_option))
    data = pd.read_excel(file)
    columns = data.columns.tolist()
    logger.info("文件{}列名为{}".format(file, columns))
    base_elements = ["学校", "年级", "班级", "排名", "学号", "姓名", "考号"]
    class_name = [item for item in columns if item not in base_elements]

    config = configparser.ConfigParser()
    config.read(config_file, encoding='utf-8')
    logger.info("读取配置文件{}".format(config_file))

    max_source_dict = {}
    if 'max_source' in config.sections():
        for option in config.options('max_source'):
            max_source_dict[option] = config.get('max_source', option)
    logger.info("读取满分配置{}".format(max_source_dict))

    user_dict = {}
    if 'config' in config.sections():
        for option in config.options('config'):
            user_dict[option] = config.get('config', option)
    logger.info("读取用户配置{}".format(user_dict))

    schools = data["学校"].unique()
    logger.info("学校列表{}".format(schools))

    for school in schools:
        school_data = data[data["学校"] == school]
        grades = school_data["年级"].unique()
        logger.info("学校{}年级列表{}".format(school, grades))
        for grade in grades:
            grade_data = school_data[school_data["年级"] == grade]
            classes = grade_data["班级"].unique()
            logger.info("学校{}年级{}班级列表{}".format(school, grade, classes))
            for _class in classes:
                class_data = grade_data[grade_data["班级"] == _class]
                logger.info("学校{}年级{}班级{}数据{}".format(school, grade, _class, class_data))

                student_count = class_data.shape[0]
                student_count1 = student_count * get_user_config(user_dict, "top_percent")
                if student_count1 % 1 != 0:
                    student_count1 = student_count1 // 1 + 1

                for c in class_name:
                    sorted_data = class_data.sort_values(by=c, ascending=False)
                    sorted_data = sorted_data.head(int(student_count1))

                    if c != "总分":
                        guocha = len(sorted_data[sorted_data[c] < get_max_source(max_source_dict, c) * get_user_config(
                            user_dict, "guocha_percent")]) / student_count1
                        jige = len(sorted_data[
                                       sorted_data[c] >= get_max_source(max_source_dict, c) * get_user_config(user_dict,
                                                                                                              "jige_percent")]) / student_count1
                        youxiu = len(sorted_data[sorted_data[c] >= get_max_source(max_source_dict, c) * get_user_config(
                            user_dict, "youxiu_percent")]) / student_count1
                        avg = sorted_data[c].mean()
                    else:
                        non_empty_columns = [col for col in sorted_data.columns if sorted_data[col].notna().any()]
                        class_name1 = [item for item in non_empty_columns if item not in base_elements]
                        all_source = 0
                        for _c in class_name1:
                            if _c != "总分":
                                all_source += int(max_source_dict.get(_c))
                        guocha = len(sorted_data[sorted_data[c] < all_source * get_user_config(user_dict,
                                                                                               "guocha_percent")]) / student_count1
                        jige = len(sorted_data[sorted_data[c] >= all_source * get_user_config(user_dict,
                                                                                              "jige_percent")]) / student_count1
                        youxiu = len(sorted_data[sorted_data[c] >= all_source * get_user_config(user_dict,
                                                                                                "youxiu_percent")]) / student_count1
                        avg = sorted_data[c].mean()
                    _result = pd.Series(
                        [school, grade, _class, c, student_count, student_count1, guocha, jige, youxiu, avg],
                        index=result_df.columns)
                    result_df = result_df._append(_result, ignore_index=True)

    input_file_directory = os.path.dirname(file)
    result_file = os.path.join(input_file_directory, ret_file_name)
    result_df.to_excel(result_file, index=False)

    return result_file


def browse_file():
    file_path = filedialog.askopenfilename()
    app_config['file'] = file_path
    logger.info("选择了文件：{}".format(file_path))


def browse_config_file():
    config_file_path = filedialog.askopenfilename()
    app_config['config_file'] = config_file_path
    logger.info("选择了配置文件：{}".format(config_file_path))


def compute_excel():
    file = app_config['file']
    config_file = app_config['config_file']

    if not file or not config_file:
        messagebox.showerror("错误", "请选择文件和配置文件")
        return

    result_file = compute_excel_function(file, config_file, None)
    messagebox.showinfo("成功", "计算完成，结果文件名：{}".format(result_file))


app = tk.Tk()
app.title("Excel计算器")
app_config = {"file": None, "config_file": None}

frame = tk.Frame(app)
frame.pack(padx=10, pady=10)

browse_file_button = tk.Button(frame, text="选择文件", command=browse_file)
browse_file_button.grid(row=0, column=0, padx=5, pady=5)

browse_config_file_button = tk.Button(frame, text="选择配置文件", command=browse_config_file)
browse_config_file_button.grid(row=0, column=1, padx=5, pady=5)

compute_button = tk.Button(frame, text="计算", command=compute_excel)
compute_button.grid(row=1, columnspan=2, pady=10)

app.mainloop()
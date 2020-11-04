from docx import Document
import datetime
import json


def load_config(filename="./config.json"):
    with open(filename) as f:
        config = json.load(f)
    if config["student_name"] == "张三":
        print("警告：似乎你忘记修改了 config.json 文件。")
    return config


def date_generator(week_num=8, working_day_num=5):
    date = datetime.datetime.strptime("08-31", "%m-%d")
    for week in range(week_num):
        for working_day in range(working_day_num):
            yield date.strftime("%m 月 %d 日")
            if working_day == working_day_num - 1:
                delta = datetime.timedelta(days=3)
            else:
                delta = datetime.timedelta(days=1)
            date += delta


def content_generator(config, num):
    paper_list = config["paper_list"]
    experiment_list = config["experiment_list"]
    task_list = config["task_list"]
    course_list = []

    content = {"student_name": config["student_name"],
               "id": config["id"],
               "lab_name": config["lab_name"],
               "major": config["major"],
               "tutor_name": config["tutor_name"],
               "project_name": config["project_name"],
               "todo_task_1": "",
               "todo_task_2": "",
               "progress_1": "",
               "progress_2": "",
               "tomorrow_plan_1": "",
               "tomorrow_plan_2": "",
               "others_1": "",
               "others_2": ""
               }
    while True:
        yield content


def generate_log(content):
    date = content["date"]
    print(f"Generating log for day: {date}...", end=" ")
    source_filename = f"./template.docx"
    target_filename = f"./generated/{date}.docx"
    doc = Document(source_filename)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:  # Keep the style
                        for key in content.keys():
                            if key in run.text:
                                run.text = run.text.replace(key, content[key])
                                break
    doc.save(target_filename)
    print("Done.")


def main():
    cfg = load_config()
    week_num = 8
    working_day_num = 5
    date = date_generator(week_num, working_day_num)
    content = content_generator(cfg, week_num * working_day_num)
    for current_date in date:
        current_content = next(content)
        current_content["date"] = current_date
        generate_log(current_content)


if __name__ == '__main__':
    main()

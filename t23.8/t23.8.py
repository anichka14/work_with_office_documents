import docx
import os
import re


"""T23.8 Скласти програму, яка здійснює заміну новим рядком рядка, що
відповідає заданому регулярному виразу, у знайдених у заданому каталозі та
його підкаталогах усіх документах MS Word."""


def find_text(folder, regex_1, regex_2):
    regex_1 = re.compile(regex_1)

    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith("docx"):
                filepath = os.path.join(root, file)
                change_text_in_document(filename=file, filepath=filepath, regex_1=regex_1, regex_2=regex_2)


def change_text_in_document(filename, filepath, regex_1, regex_2):
    doc = docx.Document(filepath)

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            match = re.search(regex_1, run.text)
            if match is not None:
                run.text = change(run.text, regex_1, regex_2)
    doc.save(filename)


def change(string, regex_1, regex_2):
    return re.sub(regex_1, regex_2, string)


if __name__ == "__main__":
    need_text = r"\d"
    change_text = r"***"
    find_text("phil", need_text, change_text)






    # для того, щоб замінити всі header i footer треба відкривати документ у xml форматі і за допомогою xpath
    # робити заміну

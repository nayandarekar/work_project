"""
This script is to create the word report file from excle file
"""
import os
import shutil


def create_new_word_file(base_path):
    """
    from current base path create new working word file 
    by copying and renameing template word file
    """
    template_file = 'template_HFRF_Report.docx'
    report_file_name = 'HFRF_Report.docx'
    template_file_path = os.path.join(base_path, template_file)
    new_file_path = os.path.join(base_path, report_file_name)
    try:
        path = shutil.copyfile(template_file_path,new_file_path)
        return path
    except FileNotFoundError:
        print (f"Report word template file not found at path: {template_file_path}")


def main():
    """This is main function"""
    base_path = os.getcwd()
    new_file_path = create_new_word_file(base_path)
    print (new_file_path)


if __name__ == "__main__":
    main()

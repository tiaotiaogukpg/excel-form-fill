def main():
    print("excel-form-fill：从 Excel 读取平时成绩与考试成绩，分别同步进系统。")
    print("用法：")
    print("  生成示例：python create_sample_excel.py")
    print("  读取并填表：python fill_form.py -e sample_data_usual_exam.xlsx -u <录入页URL> [--dry-run]")
    print("  Excel 支持「平时成绩」「考试成绩」两列，或单列「成绩」（同时填两列）。")


if __name__ == "__main__":
    main()

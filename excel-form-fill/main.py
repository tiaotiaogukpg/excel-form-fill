def main():
    print("excel-form-fill：Excel → 强校验 → 任务生成（browser-use 自动成绩录入）")
    print("用法：")
    print("  1. 看说明：python main.py")
    print("  2. 填表（先 dry-run 看任务）：python fill_form.py -e <Excel路径> -u <录入页URL> --dry-run")
    print("  Excel 必须同时有「平时成绩」「考试成绩」两列，否则强校验禁止填表。详见 USAGE.md。")


if __name__ == "__main__":
    main()

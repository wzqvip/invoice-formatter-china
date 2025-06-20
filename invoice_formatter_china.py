# invoice_formatter_china.py

import os
import pdfplumber
import json
import csv
import re
import shutil
import argparse
import pandas as pd

def extract_invoice_texts(folder_path, output_json):
    results = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            filepath = os.path.join(folder_path, filename)
            with pdfplumber.open(filepath) as pdf:
                text = "\n".join([page.extract_text() or "" for page in pdf.pages])
            results.append({"文件名": filename, "内容": text})
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

def ask_gpt_to_format(json_path, openai_key, output_csv):
    from openai import OpenAI
    client = OpenAI(api_key=openai_key)

    with open(json_path, "r", encoding="utf-8") as f:
        invoice_data = json.load(f)

    messages = [
        {
            "role": "system",
            "content": (
                "你是一个中文发票内容整理助手。请根据用户提供的发票文本数组提取关键信息，"
                "每行生成一个CSV行，字段包括：文件名、发票号码、开票日期、购买方、销售方、价税合计（元）。"
                "直接返回CSV纯文本，无需解释说明。"
            )
        },
        {
            "role": "user",
            "content": json.dumps(invoice_data, ensure_ascii=False)
        }
    ]

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        temperature=0.0,
        max_tokens=4096
    )

    csv_text = response.choices[0].message.content.strip()
    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        for line in csv_text.splitlines():
            f.write(line + "\n")

def copy_and_rename_files(csv_file, source_folder, target_folder):
    os.makedirs(target_folder, exist_ok=True)

    def sanitize_filename(name):
        return re.sub(r'[\\/:"*?<>|]', '_', name)

    with open(csv_file, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            old_filename = row["文件名"]
            sale_name = row["销售方"]
            date = row["开票日期"]
            total = row["价税合计（元）"]

            new_filename = sanitize_filename(f"{sale_name} - {total} - {date}.pdf")
            old_path = os.path.join(source_folder, old_filename)
            new_path = os.path.join(target_folder, new_filename)

            if os.path.exists(old_path):
                shutil.copy2(old_path, new_path)
                print(f"✔ 已复制并重命名为: {new_filename}")
            else:
                print(f"⚠ 未找到文件: {old_filename}")

def main():
    parser = argparse.ArgumentParser(description="中国发票 PDF 内容提取与 GPT 结构化输出工具")
    parser.add_argument("--pdf_dir", type=str, required=True, help="PDF 发票文件夹路径")
    parser.add_argument("--json_out", type=str, default="invoice_texts.json", help="提取输出的 JSON 文件名")
    parser.add_argument("--csv_out", type=str, default="invoices.csv", help="GPT 输出的 CSV 文件名")
    parser.add_argument("--api_key", type=str, help="你的 OpenAI API Key（web 模式可省略）")
    parser.add_argument("--rename", type=int, default=0, help="是否复制并重命名发票文件（0/1）")
    parser.add_argument("--excel", type=int, default=0, help="是否将CSV转为Excel（0/1）")
    parser.add_argument("--web", type=int, default=0, help="是否使用 ChatGPT 网页版手动处理（0/1）")

    args = parser.parse_args()

    print("[1] 提取发票文本...")
    extract_invoice_texts(args.pdf_dir, args.json_out)
    print(f"✅ 已保存发票文本到 {args.json_out}")

    if args.web:
        print("[2] 网页版处理模式：")
        print("请将以下 JSON 文本复制粘贴到 ChatGPT 中（提示词见系统消息）：\n")
        with open(args.json_out, "r", encoding="utf-8") as f:
            print(f.read())
        
        print(f"\n🔁 粘贴 GPT 返回的 CSV 内容后，将其保存为: {args.csv_out}")
        input(f"📎 完成后请按 Enter 键继续...")  # 等待用户确认
        if not os.path.exists(args.csv_out):
            print(f"❌ 找不到 {args.csv_out}，请确认你已经粘贴并保存了 GPT 返回的 CSV。")
            return

    else:
        if not args.api_key:
            print("❌ 未提供 API Key，无法使用 GPT 模式")
            return
        print("[2] 调用 GPT-4o 提取关键信息...")
        ask_gpt_to_format(args.json_out, args.api_key, args.csv_out)
        print(f"✅ GPT 输出已保存为 {args.csv_out}")

    if args.rename:
        print("[3] 正在复制并重命名发票文件...")
        copy_and_rename_files(args.csv_out, args.pdf_dir, "./整理后发票")
        print("✅ 发票已复制并整理到 ./整理后发票")

    if args.excel:
        print("[4] 正在将 CSV 转为 Excel...")
        df = pd.read_csv(args.csv_out)
        df.to_excel(args.csv_out.replace(".csv", ".xlsx"), index=False)
        print("✅ 已生成 Excel 文件")

if __name__ == "__main__":
    main()

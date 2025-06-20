# invoice_formatter_china.py
# Purpose: Extract and format Chinese invoice (PDF) content using NLP for expense reporting

import os
import pdfplumber
import json
from openai import OpenAI
import csv
import argparse

# === Step 1: Extract all text from PDF invoices ===
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

# === Step 2: Call OpenAI GPT-4o API to format content ===
def ask_gpt_to_format(json_path, openai_key, output_csv):
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

    csv_text = response.choices[0].message.content

    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        for line in csv_text.strip().splitlines():
            f.write(line + "\n")

# === Example Usage ===
# extract_invoice_texts("./invoices", "invoice_texts.json")
# ask_gpt_to_format("invoice_texts.json", openai_key="sk-...", output_csv="invoices.csv")

def main():
    parser = argparse.ArgumentParser(description="中国发票 PDF 内容提取与 GPT 结构化输出工具")
    parser.add_argument("--pdf_dir", type=str, required=True, help="PDF 发票文件夹路径")
    parser.add_argument("--json_out", type=str, default="invoice_texts.json", help="提取输出的 JSON 文件名")
    parser.add_argument("--csv_out", type=str, default="invoices.csv", help="GPT 输出的 CSV 文件名")
    parser.add_argument("--api_key", type=str, required=True, help="你的 OpenAI API Key")
    args = parser.parse_args()

    print("[1] 提取发票文本...")
    extract_invoice_texts(args.pdf_dir, args.json_out)
    print(f"✅ 已保存发票文本到 {args.json_out}")

    print("[2] 通过 GPT-4o 提取关键信息并输出 CSV...")
    ask_gpt_to_format(args.json_out, args.api_key, args.csv_out)
    print(f"✅ GPT 输出已保存为 {args.csv_out}")

if __name__ == "__main__":
    main()
    import pandas as pd
    df = pd.read_csv("invoices.csv")
    df.to_excel("invoices.xlsx", index=False)

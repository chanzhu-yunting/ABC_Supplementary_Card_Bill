# -*- coding:utf-8 -*-
"""
作者: Yunting Zhao
时间: 2025/1/13 10:36
"""
import pandas as pd
import email
from email.policy import default
from bs4 import BeautifulSoup
import re
import csv


def extract_clean_email_body(eml_file_path):
    """
    提取并清理EML文件中的邮件正文。
    """
    try:
        with open(eml_file_path, 'rb') as eml_file:
            msg = email.message_from_binary_file(eml_file, policy=default)

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    body += part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8")
                elif content_type == "text/html" and "attachment" not in content_disposition:
                    html_content = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8")
                    body += clean_html(html_content)
        else:
            content_type = msg.get_content_type()
            if content_type == "text/plain":
                body = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8")
            elif content_type == "text/html":
                html_content = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8")
                body = clean_html(html_content)
        return body

    except Exception as e:
        print(f"提取邮件正文时出错：{e}")
        return ""


def clean_html(html_content):
    """
    清理HTML内容，提取纯文本。
    """
    soup = BeautifulSoup(html_content, "html.parser")
    for img in soup.find_all("img"):
        img.decompose()
    for tag in soup(["script", "style"]):
        tag.decompose()
    text = soup.get_text(separator="\n").strip().replace("\xa0", " ")
    return text


def remove_unwanted_lines(content, exclude_english=True):
    """
    删除指定内容中的空行和以英文字母开头的行。
    """
    lines = content.splitlines()
    filtered_lines = [
        line for line in lines if line.strip() and (not exclude_english or not re.match(r'^[a-zA-Z]', line))
    ]
    return "\n".join(filtered_lines)


def process_email_to_csv(content, output_file):
    """
    提取有效交易信息并保存为CSV文件。
    """
    try:
        lines = content.split('\n')
        start_index, end_index = -1, -1
        for i, line in enumerate(lines):
            if "交易明细" in line:
                start_index = i + 1
            if "温馨提示" in line:
                end_index = i
                break

        if start_index == -1 or end_index == -1 or start_index >= end_index:
            print("未找到有效交易内容，无法处理！")
            return

        # 处理指定范围内的行
        filtered_lines = []
        for i, line in enumerate(lines):
            if start_index <= i < end_index:
                # 删除包含指定内容的行
                if any(keyword in line for keyword in [
                    "TDate", "PDate", "Card No.", "Type",
                    "City/Merchant Name/Branches", "Tran Amt/Curr", "Sett Amt/Curr"
                ]):
                    # if "Sett Amt/Curr" in line:
                    #     # 一旦命中并删除 "Sett Amt/Curr"，终止处理
                    #     break
                    continue  # 其他匹配内容的行直接跳过
                filtered_lines.append(line)

        # valid_lines = [line.strip() for line in lines[start_index:end_index] if line.strip()]
        valid_lines = [line.strip() for line in filtered_lines if line.strip()]
        transactions, temp_transaction = [], []

        for line in valid_lines:
            temp_transaction.append(line)
            # 使用正则表达式检查最后两个元素是否匹配 [数字]/CNY 格式
            if len(temp_transaction) >= 3:
                last_two_elements = temp_transaction[-2:]
                pattern = r'^-?\d+(\.\d+)?/CNY$'  # 匹配正负数字，带可选小数点，且以 /CNY 结尾
                if all(re.match(pattern, element) for element in last_two_elements):
                    transactions.append(temp_transaction)
                    temp_transaction = []
                # 保证7行交易的处理逻辑不变
                if len(temp_transaction) == 7:
                    transactions.append(temp_transaction)
                    temp_transaction = []

        if temp_transaction:
            transactions.append(temp_transaction)

        # with open(output_file, 'w', encoding='utf-8', newline='') as csvfile:
        #     csv_writer = csv.writer(csvfile)
        #     csv_writer.writerows(transactions)

        # print(f"交易信息已保存为CSV文件：{output_file}")
        return pd.DataFrame(transactions[1:], columns=transactions[0])
    except Exception as e:
        print(f"处理交易信息时出错：{e}")


def calculate_account_balance(csv_file, card_number):
    """
    计算指定卡号的累计入账金额。
    """
    try:
        with open(csv_file, 'r', encoding='utf-8') as f:
            csv_reader = csv.reader(f)
            header = next(csv_reader)
            if "卡号后四位" not in header or "入账金额/币种(支出为-)" not in header:
                print("CSV文件格式不正确！")
                return None

            card_index = header.index("卡号后四位")
            amount_index = header.index("入账金额/币种(支出为-)")
            return sum(
                float(row[amount_index].split('/')[0])
                for row in csv_reader if len(row) > max(card_index, amount_index) and row[card_index] == card_number
            )
    except FileNotFoundError:
        print(f"文件未找到：{csv_file}")

    except Exception as e:
        print(f"计算账户余额时出错：{e}")
    return None


def calculate_account_balance_df(df, card_number):
    """
    计算指定卡号的累计入账金额。

    参数:
        df: pandas.DataFrame - 包含交易记录的 DataFrame。
        card_number: str - 卡号的后四位。

    返回:
        float - 累计入账金额，若出错则返回 None。
    """
    try:
        # 检查 DataFrame 是否包含必要的列
        required_columns = ["卡号后四位", "入账金额/币种(支出为-)"]
        if not all(col in df.columns for col in required_columns):
            print("DataFrame 格式不正确！")
            return None

        # 筛选出目标卡号的交易记录
        filtered_df = df[df["卡号后四位"] == card_number]

        # 提取金额部分并计算累计入账金额
        total_amount = filtered_df["入账金额/币种(支出为-)"] \
            .str.split('/', expand=True)[0] \
            .astype(float) \
            .sum()

        return total_amount

    except Exception as e:
        print(f"计算账户余额时出错：{e}")
        return None


def get_transaction_amounts(csv_file, card_number):
    """
    获取指定卡号的每笔交易金额列表。
    """
    try:
        with open(csv_file, 'r', encoding='utf-8') as f:
            csv_reader = csv.reader(f)
            header = next(csv_reader)
            if "卡号后四位" not in header or "入账金额/币种(支出为-)" not in header:
                print("CSV文件格式不正确！")
                return []

            card_index = header.index("卡号后四位")
            amount_index = header.index("入账金额/币种(支出为-)")
            return [
                float(row[amount_index].split('/')[0])
                for row in csv_reader if len(row) > max(card_index, amount_index) and row[card_index] == card_number
            ]
    except FileNotFoundError:
        print(f"文件未找到：{csv_file}")
    except Exception as e:
        print(f"获取交易金额列表时出错：{e}")
    return []


def get_transaction_amounts_df(df, card_number):
    """
    获取指定卡号的每笔交易金额列表。

    参数:
        df: pandas.DataFrame - 包含交易记录的 DataFrame。
        card_number: str - 卡号的后四位。

    返回:
        list[float] - 每笔交易金额列表，若出错返回空列表。
    """
    try:
        # 检查 DataFrame 是否包含必要的列
        required_columns = ["卡号后四位", "入账金额/币种(支出为-)"]
        if not all(col in df.columns for col in required_columns):
            print("DataFrame 格式不正确！")
            return []

        # 筛选指定卡号的交易记录
        filtered_df = df[df["卡号后四位"] == card_number]

        # 提取交易金额部分并转换为浮点数
        transaction_amounts = (
            filtered_df["入账金额/币种(支出为-)"]
            .str.split("/", expand=True)[0]
            .astype(float)
            .tolist()
        )

        return transaction_amounts

    except Exception as e:
        print(f"获取交易金额列表时出错：{e}")
        return []


def analyse_deal_mess(email_path, card_number, output_csv="transactions.txt"):
    """
    分析邮件并输出交易信息和统计结果。
    """
    content = []
    clean_body = extract_clean_email_body(email_path)
    # filtered_body = remove_unwanted_lines(clean_body)
    data = process_email_to_csv(clean_body, output_csv)
    total_balance = calculate_account_balance_df(data, card_number)
    if total_balance is not None:
        content.append(total_balance)
        # print(f"卡号尾号 {card_number} 的账户累计入账金额为：{total_balance:.2f}")
    else:
        return None
    amounts = get_transaction_amounts_df(data, card_number)
    content.append(amounts)
    # print(f"卡号尾号 {card_number} 的账户上月总计消费 {len(amounts)} 笔")
    # print(f"每笔交易金额：{amounts}")
    return content


if __name__ == "__main__":

    card_number = "1234"
    acc_mess = analyse_deal_mess("中国农业银行金穗信用卡电子对账单.eml", card_number)
    if len(acc_mess) == 2:
        total_balance = acc_mess[0]
        amounts = acc_mess[1]
        print(f"卡号尾号 {card_number} 的账户累计入账金额为：{total_balance:.2f}")
        print(f"卡号尾号 {card_number} 的账户上月总计消费 {len(amounts)} 笔")
        print(f"每笔交易金额：{amounts}")



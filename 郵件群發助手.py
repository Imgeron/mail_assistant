import pandas as pd
import win32com.client as win32
from datetime import datetime
import os

# 读取mail_loop表格
mail_list = pd.read_excel('mail_loop.xlsx', sheet_name='mail list')
mail_content = pd.read_excel('mail_loop.xlsx', sheet_name='郵件設定')

# 读取data表格并用"0"替换空值
data = pd.read_excel('data.xlsx').fillna(0)

# 获取当前登录用户的Outlook邮箱地址
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
current_user = namespace.CurrentUser
current_user_email = current_user.Address

# 删除空行
data.dropna(how='all', inplace=True)

# 获取当前文件夹路径
current_folder = os.getcwd()

# 遍历mail_list表格的每一行
for index, row in mail_list.iterrows():
    vendor = row['vendor']
    to_emails = row['to'].replace(';', ',').split(',')
    cc = row['cc']
    # 检查cc是否为空
    if pd.isna(cc) or cc is None:
        cc_emails = []
    else:
        cc_emails = cc.replace('，', ',').split(',')

    # 根据vendor筛选data表格的数据
    vendor_data = data[data['vendor'] == vendor].copy()

    # 检查vendor是否在data中存在
    if vendor_data.empty:
        print(f'No data found for vendor: {vendor}')
        continue

    # 确保整行都是非空值，保留表头
    vendor_data.dropna(how='all', inplace=True)

    # 创建Outlook应用程序对象
    outlook = win32.Dispatch('Outlook.Application')

    # 创建邮件对象
    mail = outlook.CreateItem(0)
    html_body = '<html><body><p>' + mail_content.iloc[0]["郵件內容"].replace("\n", "<br>") + '</p><br><br>' + vendor_data.to_html(index=False) + '</body></html>'
    mail.HTMLBody = html_body
    subject_date = datetime.now().strftime("%Y%m%d")
    mail.Subject = f'{mail_content.iloc[0]["郵件主旨"]} {vendor} ({subject_date})'

    # 添加收件人和抄送
    mail.To = ';'.join(to_emails)
    mail.CC = ';'.join(cc_emails)

    # 添加附件
    attachment_name = f'{vendor}_data_{subject_date}.xlsx'
    attachment_path = os.path.join(current_folder, attachment_name)
    vendor_data.to_excel(attachment_path, index=False)
    mail.Attachments.Add(attachment_path)

    # 发送邮件
    mail.Send()
    print(f'Sent email for vendor: {vendor}')

    # 删除临时附件文件
    os.remove(attachment_path)

# 检查data中的vendor是否在mail_loop中存在
missing_vendors = set(data['vendor']) - set(mail_list['vendor'])
if missing_vendors:
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = '邮件发送错误'
    mail.Body = f'以下厂商在mail_loop中不存在，请维护mail_loop表格：\n\n{", ".join(missing_vendors)}'
    mail.To = current_user_email
    mail.Send()
    print('Sent email for missing vendors')
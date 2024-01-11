"""
Класс для работы с почтовым ящиком при обновлении прайс-листов

Использование класса:
- создаем объект, передаем в него настройки подключения к почте
- далее можем обращаться к методам

Методы:
- проверка статуса подключения к своей почте (подключен/не подключен)
- проверить письма от определенного отправителя
- скачать файлы из определенного письма
"""


#Библиотеки для работы с почтой
import imaplib
import email
import os
# --------------------------------------------------------------
# --------------------------------------------------------------

class docpart_imap:
    
    # --------------------------------------------------------------
    
    #В конструкторе иницализируем настройки подключения к почте
    def __init__(self, email_server, email_encryption, email_port, email_username, email_password):
        self.email_server = email_server
        self.email_encryption = email_encryption
        self.email_port = email_port
        self.email_username = email_username
        self.email_password = email_password
    
    # --------------------------------------------------------------
    # Метод пометки всех писем как прочитанных
    def mark_all_messages_as_seen(self):
        with imaplib.IMAP4_SSL(self.email_server, self.email_port) as M:
            M.login(self.email_username, self.email_password)
            M.select('INBOX')
            typ, data = M.search(None, 'ALL')
            for num in data[0].split():
                M.store(num, '+FLAGS', '\\Seen')
    
    #Метод получения статуса подключения к почте
    def get_imap_status(self):
        try:
            with imaplib.IMAP4_SSL(self.email_server, self.email_port) as M:
                M.login(self.email_username, self.email_password)
                M.select('INBOX')
                return True
        except:
            return False
    
    # --------------------------------------------------------------
    
    #Получения писем от определенного отправителя. Вернет список объектов с описанием писем.
    def get_messages_by_sender(self, sender, need_mark_seen):
        messages = []  # List to return

        with imaplib.IMAP4_SSL(self.email_server, self.email_port) as M:
            M.login(self.email_username, self.email_password)
            M.select('INBOX')
            typ, data = M.search(None, 'UNSEEN', 'FROM', sender)
            for num in data[0].split():
                typ, data = M.fetch(num, '(RFC822)')
                message = {}
                message['uid'] = num  # ID письма
                raw_message_data = data[0][1]
                msg = email.message_from_bytes(raw_message_data)
                
                # Decode and handle encoded headers
                message['date'] = self.decode_header(msg['Date'])
                message['from_'] = self.decode_header(msg['From'])
                message['subject'] = self.decode_header(msg['Subject'])
                
                message['attachments'] = []  # List of attachments

                # Process attachments
                for part in msg.walk():
                    if part.get_filename():
                        filename = self.decode_header(part.get_filename())
                        message['attachments'].append(filename)

                messages.append(message)

        return messages

    def decode_header(self, header):
        try:
            # Try to decode header using UTF-8, and if it fails, replace errors
            return email.header.decode_header(header)[0][0].decode('utf-8', 'replace')
        except Exception:
            return header
        
    # --------------------------------------------------------------
    
    # Метод скачивания файла с именем filename из сообщения с message_uid в директорию dest_dir. Здесь нет учета случая, при котором в одном письме могут быть файлы с одинаковыми именами. Скачивание произвойдет для первого файла. Если на практике встретится такой случай с файлами с одинаковыми именами в одном письме, то, можно будет доработать - искать файл под нескольким критериям
    def download_file_from_message(self, message_uid, filename, dest_dir, need_mark_seen):
        try:
            with imaplib.IMAP4_SSL(self.email_server, self.email_port) as M:
                M.login(self.email_username, self.email_password)
                M.select('INBOX')
                typ, data = M.fetch(message_uid, '(RFC822)')

                for response_part in data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        for part in msg.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            if part.get('Content-Disposition') is None:
                                continue
                            if filename.endswith('.zip' or '.rar' or '.7z' or '.gz' or '.tar'):
                                if part.get_filename() == filename:
                                    file_path = os.path.join(dest_dir, filename)
                                    with open(file_path, 'wb') as f:
                                        f.write(part.get_payload(decode=True))
                                    return True
                            else:
                                file_path = f"{dest_dir}/{filename}"
                                with open(file_path, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                return True
        except Exception as e:
            print(f"Error downloading file: {e}")
        return False

    
    
    # --------------------------------------------------------------

# --------------------------------------------------------------

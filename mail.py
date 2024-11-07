import imaplib
import email
from bs4 import BeautifulSoup
from email.utils import getaddresses
import re
import csv

# Подключение к почтовому серверу Mail.ru
IMAP_SERVER = "imap.mail.ru"
EMAIL_ACCOUNT = "k_kuhta@mail.ru"  # Замените на ваш email
PASSWORD = "0fh6Rj1cFybEhJuwak6Y"  # Замените на ваш пароль

# Подключение и аутентификация
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL_ACCOUNT, PASSWORD)

# Переход в папку "Входящие" и получение всех писем
mail.select("inbox")

# Поиск всех писем
result, data = mail.search(None, "ALL")

# Собираем адреса электронной почты
email_addresses = set()


def decode_body(body, charset):
    """Функция для декодирования тела письма с несколькими кодировками."""
    try:
        return body.decode(charset or 'utf-8')
    except UnicodeDecodeError:
        try:
            return body.decode('latin-1')
        except UnicodeDecodeError:
            return body.decode('ISO-8859-1', errors='ignore')


for num in data[0].split():
    result, msg_data = mail.fetch(num, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    # Получение адресов из заголовков
    from_address = msg.get("From")
    to_addresses = msg.get("To")
    cc_addresses = msg.get("Cc")

    all_addresses = getaddresses([from_address, to_addresses, cc_addresses])
    for name, address in all_addresses:
        if address:
            email_addresses.add(address)

    # Если письмо содержит текст/HTML, парсим его
    for part in msg.walk():
        if part.get_content_type() in ["text/plain", "text/html"]:
            charset = part.get_content_charset()
            body = part.get_payload(decode=True)
            decoded_body = decode_body(body, charset)

            # Проверка на наличие HTML-тегов перед парсингом
            if "<html" in decoded_body.lower() or "<body" in decoded_body.lower():
                soup = BeautifulSoup(decoded_body, "lxml")
                found_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", soup.get_text())
            else:
                found_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", decoded_body)
            for email_address in found_emails:
                email_addresses.add(email_address)

# Переход в папку "Отправленные" и повторение поиска
mail.select('"Отправленные"')
result, data = mail.search(None, "ALL")

for num in data[0].split():
    result, msg_data = mail.fetch(num, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    from_address = msg.get("From")
    to_addresses = msg.get("To")
    cc_addresses = msg.get("Cc")

    all_addresses = getaddresses([from_address, to_addresses, cc_addresses])
    for name, address in all_addresses:
        if address:
            email_addresses.add(address)

    for part in msg.walk():
        if part.get_content_type() in ["text/plain", "text/html"]:
            charset = part.get_content_charset()
            body = part.get_payload(decode=True)
            decoded_body = decode_body(body, charset)

            if "<html" in decoded_body.lower() or "<body" in decoded_body.lower():
                soup = BeautifulSoup(decoded_body, "lxml")
                found_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", soup.get_text())
            else:
                found_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", decoded_body)
            for email_address in found_emails:
                email_addresses.add(email_address)

# Закрытие соединения
mail.logout()

# Запись уникальных email-адресов в файл CSV
with open("email_addresses.csv", "w", newline="") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Email"])  # Записываем заголовок
    for address in sorted(email_addresses):
        writer.writerow([address])

print("Найденные адреса электронной почты сохранены в 'email_addresses.csv'")

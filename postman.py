import pandas as pd
import smtplib
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import time
import re

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

banner = '''

██████╗  ██████╗ ███████╗████████╗███╗   ███╗ █████╗ ███╗   ██╗
██╔══██╗██╔═══██╗██╔════╝╚══██╔══╝████╗ ████║██╔══██╗████╗  ██║
██████╔╝██║   ██║███████╗   ██║   ██╔████╔██║███████║██╔██╗ ██║
██╔═══╝ ██║   ██║╚════██║   ██║   ██║╚██╔╝██║██╔══██║██║╚██╗██║
██║     ╚██████╔╝███████║   ██║   ██║ ╚═╝ ██║██║  ██║██║ ╚████║
╚═╝      ╚═════╝ ╚══════╝   ╚═╝   ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═══╝

            Creator: @gitem
'''

def read_accounts_from_excel(excel_file):
    """
    Читает данные аккаунтов из Excel файла
    
    Аргументы:
    excel_file - путь к Excel файлу
    
    Возвращает:
    list of dict - список словарей с данными аккаунтов
    """
    try:
        # Читаем Excel файл
        df = pd.read_excel(excel_file)
        
        # Проверяем наличие необходимых колонок
        required_columns = ['name_mail', 'passwd_mail']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error(f"Отсутствуют колонки: {missing_columns}")
            return []
        
        accounts = []
        
        for index, row in df.iterrows():
            account_data = {}
            
            # Читаем данные из колонок
            account_data['login'] = str(row['name_mail']).strip()
            account_data['password'] = str(row['passwd_mail']).strip()
            
            # Пропускаем пустые строки
            if not account_data['login'] or not account_data['password']:
                continue
            
            # Автоматически определяем SMTP сервер из email
            email_domain = account_data['login'].split('@')[-1].lower()
            
            # Сопоставление доменов с SMTP серверами
            smtp_servers = {
                'gmail.com': ('smtp.gmail.com', 587),
                'googlemail.com': ('smtp.gmail.com', 587),
                'yahoo.com': ('smtp.mail.yahoo.com', 587),
                'outlook.com': ('smtp-mail.outlook.com', 587),
                'hotmail.com': ('smtp-mail.outlook.com', 587),
                'live.com': ('smtp-mail.outlook.com', 587),
                'mail.ru': ('smtp.mail.ru', 587),
                'bk.ru': ('smtp.mail.ru', 587),
                'list.ru': ('smtp.mail.ru', 587),
                'inbox.ru': ('smtp.mail.ru', 587),
                'yandex.ru': ('smtp.yandex.ru', 587),
                'ya.ru': ('smtp.yandex.ru', 587),
                'rambler.ru': ('smtp.rambler.ru', 587),
                'lenta.ru': ('smtp.rambler.ru', 587),
                'autorambler.ru': ('smtp.rambler.ru', 587),
                'myrambler.ru': ('smtp.rambler.ru', 587),
                'ro.ru': ('smtp.rambler.ru', 587),
                'icloud.com': ('smtp.mail.me.com', 587),
                'me.com': ('smtp.mail.me.com', 587),
                'mac.com': ('smtp.mail.me.com', 587)
            }
            
            if email_domain in smtp_servers:
                account_data['smtp_server'], account_data['smtp_port'] = smtp_servers[email_domain]
            else:
                # Для неизвестных доменов используем стандартный SMTP
                account_data['smtp_server'] = f'smtp.{email_domain}'
                account_data['smtp_port'] = 587
            
            accounts.append(account_data)
            logger.info(f"Загружен аккаунт: {account_data['login']} -> {account_data['smtp_server']}")
        
        logger.info(f"Всего загружено аккаунтов: {len(accounts)}")
        return accounts
        
    except Exception as e:
        logger.error(f"Ошибка чтения Excel файла: {e}")
        print(f"\n Ошибка при чтении файла {excel_file}: {e}")
        return []

def send_bulk_emails(accounts, recipients, subject, body, delay=1):
    """
    Отправка писем с множества аккаунтов
    
    Аргументы:
    accounts - список аккаунтов из read_accounts_from_excel()
    recipients - список email получателей
    subject - тема письма
    body - текст письма
    delay - задержка между отправками в секундах
    """
    successful_sends = 0
    failed_sends = 0
    
    print(f"\n Начинаем массовую рассылку:")
    print(f" Аккаунтов: {len(accounts)}")
    print(f" Получателей: {len(recipients)}")
    print(f" Всего писем: {len(accounts) * len(recipients)}")
    print(f" Задержка: {delay} сек\n")
    
    for i, account in enumerate(accounts, 1):
        print(f"🔧 Работа с аккаунтом {i}/{len(accounts)}: {account['login']}")
        
        for j, recipient in enumerate(recipients, 1):
            try:
                print(f"📨 Отправка {j}/{len(recipients)}: {account['login']} -> {recipient}")
                
                # Используем универсальную функцию отправки
                success = send_email(
                    login=account['login'],
                    app_password=account['password'],
                    to_address=recipient,
                    subject=subject,
                    body=body,
                    smtp_server=account['smtp_server'],
                    smtp_port=account['smtp_port']
                )
                
                if success:
                    successful_sends += 1
                    logger.info(f"✓ Успешно отправлено: {account['login']} -> {recipient}")
                    print(f"✅ Успешно: {account['login']} -> {recipient}")
                else:
                    failed_sends += 1
                    logger.error(f"✗ Ошибка отправки: {account['login']} -> {recipient}")
                    print(f"❌ Ошибка: {account['login']} -> {recipient}")
                
                # Задержка между отправками
                if j < len(recipients):
                    time.sleep(delay)
                    
            except Exception as e:
                failed_sends += 1
                logger.error(f"Критическая ошибка при отправке: {e}")
                print(f" Критическая ошибка: {e}")
                continue
        
        # Задержка между сменой аккаунтов
        if i < len(accounts):
            print(f"⏳ Задержка {delay * 2} сек перед сменой аккаунта...")
            time.sleep(delay * 2)
    
    print(f"\n Итоги рассылки:")
    print(f"✅ Успешно отправлено: {successful_sends}")
    print(f"❌ Не удалось отправить: {failed_sends}")
    
    logger.info(f"Итоги рассылки: Успешно - {successful_sends}, Неудачно - {failed_sends}")
    return successful_sends, failed_sends

def send_email(login, app_password, to_address, subject, body, smtp_server, smtp_port=587):
    """
    Универсальная функция отправки email через любой SMTP сервер
    """
    try:
        # Создаем сообщение
        msg = MIMEMultipart()
        msg['From'] = login
        msg['To'] = to_address
        msg['Subject'] = Header(subject, 'utf-8')
        
        # Добавляем тело письма
        text_part = MIMEText(body, 'plain', 'utf-8')
        msg.attach(text_part)
        
        # Устанавливаем соединение и отправляем
        logger.info(f"Подключаемся к {smtp_server}:{smtp_port}")
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Включаем шифрование
            logger.info("Шифрование TLS включено")
            
            server.login(login, app_password)
            logger.info("Аутентификация успешна")
            
            server.send_message(msg)
            logger.info(f"Письмо успешно отправлено с {login} на {to_address}")
            
        return True
        
    except smtplib.SMTPAuthenticationError as e:
        logger.error(f"Ошибка аутентификации для {login}: {e}")
        return False
        
    except Exception as e:
        logger.error(f"Общая ошибка при отправке с {login}: {e}")
        return False

def read_recipients_from_excel(excel_file, sheet_name=0, column_name='recipients'):
    """
    Читает список получателей из Excel файла
    
    Аргументы:
    excel_file - путь к Excel файлу
    sheet_name - имя или индекс листа
    column_name - название колонки с email
    """
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            logger.error(f"Колонка '{column_name}' не найдена в файле")
            return []
        
        recipients = []
        for email in df[column_name].dropna():
            email_str = str(email).strip()
            if re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email_str):
                recipients.append(email_str)
        
        logger.info(f"Загружено получателей: {len(recipients)}")
        return recipients
        
    except Exception as e:
        logger.error(f"Ошибка чтения получателей из Excel: {e}")
        return []

def main():
    print(banner)

    # Чтение аккаунтов из Excel
    excel_file = input("Введите путь к Excel файлу с аккаунтами (например: accounts.xlsx): ").strip()
    accounts = read_accounts_from_excel(excel_file)
    
    if not accounts:
        print("❌ Не удалось загрузить аккаунты из Excel файла!")
        return

    # Чтение получателей
    recipients_file = input("Введите путь к Excel файлу с получателями (или нажмите Enter для ручного ввода): ").strip()
    recipients = []
    
    if recipients_file:
        recipients = read_recipients_from_excel(recipients_file)
    
    if not recipients:
        recipients_input = input("Введите email получателей (через запятую): ")
        recipients = [email.strip() for email in recipients_input.split(',') if email.strip()]

    subject = input("Введите тему письма: ").strip()
    body = input("Введите текст письма: ").strip()

    # Настройка задержки
    try:
        delay = float(input("Введите задержку между отправками в секундах (по умолчанию 1): ") or "1")
    except ValueError:
        delay = 1

    print(f"\n{'='*50}")
    print("ПОДТВЕРЖДЕНИЕ РАССЫЛКИ")
    print(f"{'='*50}")
    print(f"Аккаунтов: {len(accounts)}")
    print(f"Получателей: {len(recipients)}")
    print(f"Всего писем: {len(accounts) * len(recipients)}")
    print(f"Задержка: {delay} сек")
    print(f"Тема: {subject}")
    print(f"{'='*50}")
    
    confirm = input("\nНачать рассылку? (y/n): ")
    if confirm.lower() != 'y':
        print("❌ Рассылка отменена!")
        return

    # Запуск массовой отправки
    successful, failed = send_bulk_emails(
        accounts=accounts,
        recipients=recipients,
        subject=subject,
        body=body,
        delay=delay
    )

    print(f"\n🎉 Рассылка завершена!")
    print(f"✅ Успешно отправлено: {successful}")
    print(f"❌ Не удалось отправить: {failed}")

if __name__ == "__main__":
    main()
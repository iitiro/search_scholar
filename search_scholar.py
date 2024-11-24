import os
import requests
from scholarly import scholarly
import pandas as pd
from datetime import datetime

# Файл з ключовими словами
keywords_file = 'keywords.txt'

# Назва папки для збереження результатів і файлів
output_folder = '!search_scholar'

# Створення основної папки, якщо її немає
os.makedirs(output_folder, exist_ok=True)

# Поточна дата і час для позначення запиту
current_time = datetime.now().strftime('%Y-%m-%d %H-%M')

# Зчитування ключових слів з файлу
with open(keywords_file, 'r', encoding='utf-8') as file:
    keywords = [line.strip() for line in file if line.strip()]

# Кількість потрібних результатів для кожного ключового слова
desired_results = 80

# Список для збереження всіх даних
all_data = []

# Цикл для кожного запиту
for keyword in keywords:
    # Пропускаємо ключові слова, які починаються з символу #
    if keyword.startswith("#"):
        print(f"Пропуск ключового слова: {keyword}")
        continue

    print(f"Обробка ключового слова: {keyword}")
    
    # Створення підпапки для поточного ключового слова з додаванням дати і часу
    keyword_folder = os.path.join(output_folder, f"{current_time} {keyword}")
    os.makedirs(keyword_folder, exist_ok=True)

    # Виконання пошукового запиту в Google Scholar
    search_query = scholarly.search_pubs(keyword)

    # Збір даних з результатів пошуку для поточного ключового слова
    data = []
    for i, result in enumerate(search_query):
        # Отримання поточного часу для фіксації дати та часу запиту
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Отримання інформації про публікацію
        title = result.get('bib', {}).get('title', 'No title')
        authors = result.get('bib', {}).get('author', 'No authors')
        abstract = result.get('bib', {}).get('abstract', 'No abstract')
        year = result.get('bib', {}).get('pub_year', 'No year')
        journal = result.get('bib', {}).get('venue', 'No journal')
        url = result.get('pub_url', 'No URL')

        # Якщо посилання закінчується на ".pdf", завантажуємо цей файл
        status = 'Not Attempted'
        if url.endswith('.pdf'):
            try:
                pdf_response = requests.get(url, timeout=10)
                pdf_response.raise_for_status()  # Перевірка на помилки у відповіді
                
                # Збереження PDF-файлу у підпапку ключевого слова
                pdf_name = os.path.join(keyword_folder, f"file_{i + 1}.pdf")
                with open(pdf_name, 'wb') as pdf_file:
                    pdf_file.write(pdf_response.content)
                    
                status = 'Успішно завантажено'
                print(f"Файл {pdf_name} завантажено успішно.")
            except requests.exceptions.RequestException as e:
                status = f"Помилка при завантаженні: {e}"
                print(f"Помилка при завантаженні файлу {url}: {e}")

        # Додаємо у список
        data.append([i + 1, title, authors, year, journal, abstract, url, status, timestamp, keyword])

        # Перевірка на досягнення бажаної кількості результатів
        if len(data) >= desired_results:
            break

    # Перевірка, якщо бажаних результатів менше бажачаної кількості
    if len(data) < desired_results:
        print(f"Для ключевого слова '{keyword}' знайдено лише {len(data)} результатів, хоча хотілось {desired_results}.")

    # Додаємо дані до загального списку
    all_data.extend(data)

# Трансформація даних в DataFrame
df = pd.DataFrame(all_data, columns=['Index', 'Title', 'Authors', 'Year', 'Journal', 'Abstract', 'URL', 'Status', 'Timestamp', 'Keyword'])

# Збереження в Excel з додаванням дати і часу до назви файлу
excel_file_path = os.path.join(output_folder, f'scholar_results_{current_time}.xlsx')
df.to_excel(excel_file_path, index=False)

print(f"Готово! Результати пошуку і файли збережені в папці '{output_folder}'.")
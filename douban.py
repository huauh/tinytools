import requests
import re
import pandas as pd
from bs4 import BeautifulSoup


def main():
    url = r'https://movie.douban.com/chart'
    headers = {
        'User-Agent':
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36'
    }
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'lxml')
    title = []
    release_dates = []
    actor = []
    rating = []
    reviewers_count = []

    for i in soup.find_all(name='div', attrs={'class': 'pl2'}):
        title.extend([
            re.sub(r'\s+', ' ', j.text).strip() for j in i.find_all(name='a')
        ])
        for j in i.find_all(name='p', attrs={'class': 'pl'}):
            temp_date = ''.join(
                re.findall(r'\d{4}-\d{2}-\d{2}.*\)\s*\/', j.text)).strip()
            actor.append(j.text.replace(temp_date, '').strip())
            release_dates.append(temp_date[:-2])
        rating.extend([
            float(j.text.strip())
            for j in i.find_all(name='span', attrs={'class': 'rating_nums'})
        ])
        reviewers_count.extend([
            int(j.text[1:-5])
            for j in i.find_all(name='span', attrs={'class': 'pl'})
        ])

    data = pd.DataFrame({
        '影名': title,
        '映日': release_dates,
        '演员': actor,
        '评分': rating,
        '人数': reviewers_count
    })
    print(data)


if __name__ == "__main__":
    main()

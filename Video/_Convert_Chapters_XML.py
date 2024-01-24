from bs4 import BeautifulSoup # pip install bs4
import re


with open(r"C:\Media Output\Handbrake Output\_Process/Chapters.xml", 'r') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')


chapters = soup.select('editionentry > chapteratom')
chapters_array = []


for chapter in chapters:
    time = re.search(r'(\d{2}):(\d{2}):(\d{2})', str(chapter))
    hrs = int(time.group(1))
    mins = int(time.group(2))
    secs = int(time.group(3))

    minutes = (hrs * 60) + mins
    seconds = secs + (minutes * 60)
    timestamp = (seconds * 1000)
    chap = {
        "title": re.sub(r'(=|;|#|\n)', r'\\\1', chapter.find('chapterstring').text),
        "startTime": timestamp
    }
    chapters_array.append(chap)

text = ";FFMETADATA1\n\n"

for i in range(len(chapters_array)-1):
    chap = chapters_array[i]
    title = chap['title']
    start = chap['startTime']
    end = chapters_array[i+1]['startTime']-1
    text += f"\n[CHAPTER]\nTIMEBASE=1/1000\nSTART={start}\nEND={end}\ntitle={title}\n"

text = text.strip()

print(text)
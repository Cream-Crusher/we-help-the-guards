import os
import argparse
from random import randint, randrange
from datetime import timedelta, datetime
from docxtpl import DocxTemplate


def save_file(schedule_territories, current_date):
    doc = DocxTemplate("pattern.docx")
    context = {
        'now_day_from': current_date,
        'now_day_before': current_date + timedelta(days=1),
        'schedule_territories': schedule_territories
    }
    doc.render(context)
    doc.save("security_guard_bypass_schedule.docx")
    

def get_schedule(set_time, objects, base=5):
    schedule_territories = {}

    for num in range(0, args.start):
        territories = {}
        
        for territory in objects:
            schedule = []
            crawl_time = set_time
            key = 0

            while crawl_time.hour != 7 and crawl_time.hour != 6:
                
                if key == 0:
                    crawl_time = set_time - timedelta(minutes=randrange(-5, 15, 5))
                    schedule.append(crawl_time)
                    key = 1

                time_between_crawl = randint(90, 120)
                time_between_crawl = base * round(time_between_crawl/base)
                after_the_crawl = timedelta(minutes=time_between_crawl)
                crawl_time = crawl_time+after_the_crawl
                schedule.append(crawl_time)

                territories.update({territory: schedule})

        schedule_territories.update({num: territories})
    return schedule_territories


def get_datetime():
    now = datetime.now()
    current_date = datetime(now.year, now.month, now.day)
    set_time = datetime(now.year, now.month, now.day, hour=8, minute=0)
    return current_date, set_time


def get_args():
    parser = argparse.ArgumentParser(description='Получение расписания')
    parser.add_argument('start', default='2', help='кол дней', type=int)
    args = parser.parse_args()
    return args
    

if __name__ == '__main__':
    args = get_args()
    objects = ['ГСМ', 'Cтоянка', 'Ангар']
    current_date, set_time = get_datetime()
    schedule_territories = get_schedule(set_time, objects)
    save_file(schedule_territories, current_date)

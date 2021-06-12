import os
from random import randint, randrange
from datetime import timedelta, datetime
from docxtpl import DocxTemplate


def world(now, schedule_territories):
    now_day = datetime(now.year, now.month, now.day)
    doc = DocxTemplate("pattern.docx")
    context = {
    'now_day_from': now_day,
    'now_day_before': now_day + timedelta(days=1),
    'schedule_territories': schedule_territories
    }
    doc.render(context)
    doc.save("security_guard_bypass_schedule.docx")


def get_a_schedule_rounded_to_5(set_time, base=5):
    territories = ['ГСМ', 'Cтоянка', 'Ангар']
    schedule_territories = {}

    for territory in territories:
        schedule = []
        crawl_time = set_time
        num = 0

        while crawl_time.hour != 7 and crawl_time.hour != 6:
            
            if num == 0:
                crawl_time = set_time - timedelta(minutes=randrange(-5, 15, 5))
                schedule.append(crawl_time)

            num = 1
            time_between_crawl = randint(90, 120)
            time_between_crawl = base * round(time_between_crawl/base)
            after_the_crawl = timedelta(minutes=time_between_crawl)
            crawl_time = crawl_time+after_the_crawl
            schedule.append(crawl_time)

        schedule_territories.update({territory: schedule})

    return schedule_territories


def get_date_and_time():
    now = datetime.now()
    set_time = datetime(now.year, now.month, now.day, hour=8, minute=0, second=0, microsecond=0)
    return now, set_time


if __name__ == '__main__':
    now, set_time = get_date_and_time()
    schedule_territories = get_a_schedule_rounded_to_5(set_time)
    world(now, schedule_territories)

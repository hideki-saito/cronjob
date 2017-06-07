import os

from crontab import CronTab

from config import *

user_cron   = CronTab(user=cronJob_user)
current_directory = os.getcwd()

job  = user_cron.new(command='python %s' %os.path.join(current_directory, 'fill_excel.py'))
job.minute.every(interval_minutes)

user_cron.write()
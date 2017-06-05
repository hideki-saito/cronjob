from crontab import CronTab
from config import *

my_cron = CronTab(user=cronJob_user)
for job in my_cron:
    print job
    print job.frequency_per_hour()
    # my_cron.remove(job)
    # my_cron.write()
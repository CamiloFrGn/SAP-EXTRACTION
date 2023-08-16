from datetime import datetime, timedelta
from datetime import datetime, timedelta
import time
import timeit
import warnings
warnings.filterwarnings('ignore')
import sys
import schedule

def schedule_job():
    print("working")
    time.sleep(10)
schedule.every(5).seconds.do(schedule_job)
while True:
    schedule.run_pending()
    time.sleep(5)
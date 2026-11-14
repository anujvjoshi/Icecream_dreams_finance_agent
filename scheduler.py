import os
import logging
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from apscheduler.schedulers.blocking import BlockingScheduler
from dotenv import load_dotenv
from agent import execute_daily_refresh

load_dotenv()
BASE_DIR = Path(__file__).resolve().parent
(BASE_DIR / "logs").mkdir(exist_ok=True)
logging.basicConfig(
    filename=BASE_DIR / "logs" / "scheduler.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

def main():
    tz = os.getenv("TIMEZONE", "Asia/Kolkata")
    run_time = os.getenv("SCHEDULE_TIME", "09:00")
    hour, minute = map(int, run_time.split(":"))
    scheduler = BlockingScheduler(timezone=ZoneInfo(tz))
    scheduler.add_job(
        execute_daily_refresh,
        trigger="cron",
        hour=hour,
        minute=minute,
        id="daily_expense_refresh",
        replace_existing=True,
        max_instances=1,
        coalesce=True,
    )
    print(f"Scheduler running. Daily refresh at {run_time} ({tz}).")
    print("Keep this terminal/process running. Press Ctrl+C to stop.")
    scheduler.start()

if __name__ == "__main__":
    main()

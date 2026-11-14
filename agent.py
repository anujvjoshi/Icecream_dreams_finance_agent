import logging
from pathlib import Path
from processor import run_refresh, save_failure

BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    filename=LOG_DIR / "agent_execution.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

def execute_daily_refresh():
    try:
        result = run_refresh()
        logging.info(
            "Refresh successful | sheet=%s | transactions=%s | total=%.2f | issues=%s",
            result["ledger_sheet"],
            result["metrics"]["transaction_count"],
            result["metrics"]["total_lifetime"],
            len(result["issues"]),
        )
        print(
            f"SUCCESS: {result['title']} | {result['ledger_sheet']} | "
            f"{result['metrics']['transaction_count']} transactions | "
            f"₹{result['metrics']['total_lifetime']:,.2f} total | "
            f"{len(result['issues'])} issues"
        )
        return result
    except Exception as exc:
        logging.exception("Daily refresh failed")
        try:
            save_failure(exc)
        except Exception:
            logging.exception("Could not save failure status")
        print(f"FAILED: {exc}")
        raise

if __name__ == "__main__":
    execute_daily_refresh()

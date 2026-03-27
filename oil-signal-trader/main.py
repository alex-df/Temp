import os
from datetime import datetime
from dotenv import load_dotenv
from apscheduler.schedulers.blocking import BlockingScheduler
from scraper.twitter_monitor import TwitterMonitor
from analyzer.sentiment import analyze_tweet
from notifier.email_sender import send_alert

load_dotenv()

monitor = TwitterMonitor()
scheduler = BlockingScheduler()


def check_and_alert():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Checking for new tweets...")
    new_tweets = monitor.check_new_tweets()

    if not new_tweets:
        print("No new tweets.")
        return

    for tweet in new_tweets:
        print(f"New tweet from @{tweet['username']}: {tweet['text'][:80]}...")

        analysis = analyze_tweet(tweet)
        print(f"Signal: {analysis['signal']} ({analysis['confidence']}%)")

        # Envoyer alerte si signal fort ou confiance élevée
        if analysis["signal"] != "NEUTRAL" or analysis["confidence"] >= 70:
            send_alert(tweet, analysis)
            print(f"Alert sent to {os.getenv('ALERT_EMAIL')}")
        else:
            print("Signal NEUTRAL faible — pas d'email envoyé.")


@scheduler.scheduled_job("interval", minutes=1)
def scheduled_check():
    try:
        check_and_alert()
    except Exception as e:
        print(f"[ERROR] {e}")


if __name__ == "__main__":
    print("=" * 50)
    print("Oil Signal Trader — Démarrage")
    print(f"Comptes surveillés : {list(monitor.user_ids.keys())}")
    print("Polling toutes les 60 secondes")
    print("=" * 50)

    # Premier check immédiat au démarrage
    check_and_alert()

    # Puis toutes les minutes
    scheduler.start()

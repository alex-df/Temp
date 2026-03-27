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

        # Envoyer alerte uniquement si signal fort ET confiance >= 90%
        if analysis["signal"] != "NEUTRAL" and analysis["confidence"] >= 90:
            send_alert(tweet, analysis)
            print(f"Alert sent to {os.getenv('ALERT_EMAIL')}")
        else:
            print(f"Signal {analysis['signal']} {analysis['confidence']}% — sous le seuil, pas d'email.")


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
    print("Polling toutes les 60 secondes | Seuil email : 90%")
    print("=" * 50)

    # Initialiser les derniers IDs sans traiter les tweets existants
    monitor.initialize_last_ids()
    print("Initialisation terminée — en attente de nouveaux tweets...")

    # Scheduler uniquement, pas de check immédiat
    scheduler.start()

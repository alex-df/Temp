import os
from dotenv import load_dotenv
load_dotenv()

from datetime import datetime, timezone
from apscheduler.schedulers.blocking import BlockingScheduler
from scraper.twitter_monitor import TwitterMonitor
from analyzer.sentiment import analyze_tweet
from notifier.email_sender import send_alert

monitor = TwitterMonitor()
scheduler = BlockingScheduler()

# Cooldown : mémorise le dernier envoi par compte (évite le spam)
# Format : {username: datetime}
last_alert_sent = {}
COOLDOWN_HOURS = 4
CONFIDENCE_THRESHOLD = 90


def check_and_alert():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Checking...")
    new_tweets = monitor.check_new_tweets()

    if not new_tweets:
        print("Aucun nouveau tweet pertinent.")
        return

    for tweet in new_tweets:
        username = tweet['username']
        print(f"Tweet pertinent @{username}: {tweet['text'][:80]}...")

        # Vérifier cooldown
        last_sent = last_alert_sent.get(username)
        if last_sent:
            hours_since = (datetime.now(timezone.utc) - last_sent).total_seconds() / 3600
            if hours_since < COOLDOWN_HOURS:
                print(f"Cooldown actif @{username} — {COOLDOWN_HOURS - hours_since:.1f}h restantes")
                continue

        analysis = analyze_tweet(tweet)
        signal = analysis['signal']
        confidence = analysis['confidence']
        print(f"Signal: {signal} ({confidence}%)")

        if signal != "NEUTRAL" and confidence >= CONFIDENCE_THRESHOLD:
            send_alert(tweet, analysis)
            last_alert_sent[username] = datetime.now(timezone.utc)
            print(f"ALERTE envoyée — {signal} {confidence}% @{username}")
        else:
            print(f"Sous le seuil ({confidence}% < {CONFIDENCE_THRESHOLD}%) — pas d'email.")


@scheduler.scheduled_job("interval", minutes=1)
def scheduled_check():
    try:
        check_and_alert()
    except Exception as e:
        print(f"[ERROR] {e}")


if __name__ == "__main__":
    print("=" * 50)
    print("Oil Signal Trader — Démarrage")
    print(f"Comptes : {list(monitor.user_ids.keys())}")
    print(f"Seuil   : {CONFIDENCE_THRESHOLD}% | Cooldown : {COOLDOWN_HOURS}h")
    print(f"Filtre  : mots-clés pétrole actif")
    print("=" * 50)

    monitor.initialize_last_ids()
    print("Initialisation terminée — en attente de nouveaux tweets...")

    scheduler.start()

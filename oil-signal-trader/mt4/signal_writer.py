import os
from datetime import datetime, timezone

SIGNAL_FILE = os.getenv("MT4_SIGNAL_FILE", "./mt4/signal.txt")


def write_signal(analysis: dict, tweet_data: dict):
    """
    Phase 2 — Écrit le signal dans un fichier lu par l'EA MT4.
    Format : SIGNAL,confidence,timestamp,username
    Exemple : BUY,85,2026-03-27T14:32:00Z,realDonaldTrump
    """
    signal = analysis["signal"]
    confidence = analysis["confidence"]
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    username = tweet_data["username"]

    os.makedirs(os.path.dirname(SIGNAL_FILE), exist_ok=True)

    with open(SIGNAL_FILE, "w") as f:
        f.write(f"{signal},{confidence},{timestamp},{username}\n")

    print(f"MT4 signal écrit : {signal} ({confidence}%) — @{username}")

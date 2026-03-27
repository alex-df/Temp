import os
import tweepy

# Comptes ciblés — uniquement ceux à fort impact direct sur le Brent
TRACKED_USERNAMES = [
    "realDonaldTrump",   # Trump — sanctions Iran, politique énergie US
    "IrnaEnglish",      # Agence presse officielle iranienne
    "OPECSecretariat",  # OPEC — décisions production
]

# Mots-clés : un tweet doit contenir au moins un de ces mots pour être analysé
OIL_KEYWORDS = [
    "oil", "petroleum", "brent", "opec", "barrel", "crude",
    "sanctions", "iran", "energy", "supply", "pipeline",
    "refinery", "production cut", "embargo", "nuclear deal",
    "strait of hormuz", "middle east", "saudi", "venezuela"
]


def is_oil_relevant(text: str) -> bool:
    """Filtre rapide — n'appelle Claude que si le tweet parle vraiment de pétrole."""
    text_lower = text.lower()
    return any(kw in text_lower for kw in OIL_KEYWORDS)


class TwitterMonitor:
    def __init__(self):
        self.client = tweepy.Client(
            bearer_token=os.getenv("X_BEARER_TOKEN"),
            consumer_key=os.getenv("X_API_KEY"),
            consumer_secret=os.getenv("X_API_SECRET"),
            access_token=os.getenv("X_ACCESS_TOKEN"),
            access_token_secret=os.getenv("X_ACCESS_TOKEN_SECRET"),
            wait_on_rate_limit=True,
        )
        self.user_ids = {}
        self.last_tweet_ids = {}
        self._resolve_user_ids()

    def _resolve_user_ids(self):
        for username in TRACKED_USERNAMES:
            try:
                resp = self.client.get_user(username=username)
                if resp.data:
                    self.user_ids[username] = resp.data.id
                    print(f"Résolu : @{username} -> {resp.data.id}")
            except Exception as e:
                print(f"[WARN] Impossible de résoudre @{username} : {e}")

    def initialize_last_ids(self):
        """Mémorise les tweets actuels sans les traiter — évite le rattrapage."""
        for username, user_id in self.user_ids.items():
            try:
                resp = self.client.get_users_tweets(
                    id=user_id,
                    max_results=5,
                    exclude=["retweets", "replies"],
                )
                if resp.data:
                    self.last_tweet_ids[user_id] = resp.data[0].id
                    print(f"Init @{username} — dernier tweet ID mémorisé")
            except Exception as e:
                print(f"[WARN] Init @{username} : {e}")

    def check_new_tweets(self):
        """Vérifie les nouveaux tweets et filtre par mots-clés pétrole."""
        new_tweets = []

        for username, user_id in self.user_ids.items():
            try:
                since_id = self.last_tweet_ids.get(user_id)
                resp = self.client.get_users_tweets(
                    id=user_id,
                    since_id=since_id,
                    max_results=5,
                    tweet_fields=["created_at", "text", "public_metrics"],
                    exclude=["retweets", "replies"],
                )

                if resp.data:
                    for tweet in resp.data:
                        if not is_oil_relevant(tweet.text):
                            print(f"@{username} — tweet ignoré (hors sujet pétrole)")
                            continue
                        new_tweets.append({
                            "username": username,
                            "tweet_id": str(tweet.id),
                            "text": tweet.text,
                            "created_at": tweet.created_at.strftime("%Y-%m-%d %H:%M UTC"),
                            "likes": tweet.public_metrics.get("like_count", 0),
                            "retweets": tweet.public_metrics.get("retweet_count", 0),
                        })
                    self.last_tweet_ids[user_id] = resp.data[0].id

            except Exception as e:
                print(f"[ERROR] Fetch @{username} : {e}")

        return new_tweets

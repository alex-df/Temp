import os
import tweepy

# Comptes à surveiller : username -> user_id (résolu au démarrage)
TRACKED_USERNAMES = [
    "realDonaldTrump",   # Trump — sanctions Iran, politique énergie
    "POTUS",            # Compte officiel présidence US
    "IrnaEnglish",      # Agence presse officielle iranienne
    "OPECSecretariat",  # OPEC — décisions production
    "EnergyIntel",      # Intelligence énergie
]


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
        self.user_ids = {}         # {username: user_id}
        self.last_tweet_ids = {}   # {user_id: last_tweet_id} — déduplication
        self._resolve_user_ids()

    def _resolve_user_ids(self):
        """Convertit les usernames en user IDs au démarrage."""
        for username in TRACKED_USERNAMES:
            try:
                resp = self.client.get_user(username=username)
                if resp.data:
                    self.user_ids[username] = resp.data.id
                    print(f"Résolu : @{username} -> {resp.data.id}")
            except Exception as e:
                print(f"[WARN] Impossible de résoudre @{username} : {e}")

    def initialize_last_ids(self):
        """Mémorise les tweets actuels sans les traiter — évite le rattrapage au démarrage."""
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
        """Vérifie les nouveaux tweets de tous les comptes surveillés."""
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
                        new_tweets.append({
                            "username": username,
                            "tweet_id": str(tweet.id),
                            "text": tweet.text,
                            "created_at": tweet.created_at.strftime("%Y-%m-%d %H:%M UTC"),
                            "likes": tweet.public_metrics.get("like_count", 0),
                            "retweets": tweet.public_metrics.get("retweet_count", 0),
                        })
                    # Mémoriser le dernier tweet ID pour la prochaine vérification
                    self.last_tweet_ids[user_id] = resp.data[0].id

            except Exception as e:
                print(f"[ERROR] Fetch @{username} : {e}")

        return new_tweets

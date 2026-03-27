import os
import json
import anthropic

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))


def analyze_tweet(tweet_data: dict) -> dict:
    """
    Analyse l'impact géopolitique d'un tweet sur le prix du Brent.
    Prompt conservateur — 90%+ uniquement pour événements majeurs confirmés.
    """
    prompt = f"""Tu es un trader senior en matières premières, spécialisé Brent crude oil.
Ton rôle est d'identifier UNIQUEMENT les événements majeurs qui vont faire bouger le Brent de +2% minimum.

Compte    : @{tweet_data['username']}
Tweet     : "{tweet_data['text']}"
Publié le : {tweet_data['created_at']}
Likes     : {tweet_data.get('likes', 0)} | Retweets : {tweet_data.get('retweets', 0)}

RÈGLES STRICTES pour la confiance :
- 90-100% : Événement MAJEUR CONFIRMÉ (sanctions officielles, guerre déclarée, coupe OPEC annoncée, blocage Hormuz)
- 70-89%  : Événement significatif mais non confirmé (menace crédible, négociations rompues)
- 50-69%  : Signal faible ou ambigu
- 0-49%   : Pas d'impact direct sur le Brent

NE PAS donner >70% pour :
- Tweets généraux sur la politique sans mention pétrole/énergie/Iran
- Rumeurs sans source officielle
- Commentaires économiques généraux
- Répétition d'une news déjà connue

Signaux BUY (hausse Brent) :
- Sanctions Iran/Venezuela officielles ou durcissement
- Conflit militaire Moyen-Orient affectant production/transport
- Réduction production OPEC+ annoncée officiellement
- Blocage Détroit Hormuz

Signaux SELL (baisse Brent) :
- Accord nucléaire Iran signé → retour production
- Augmentation production OPEC+
- Libération réserves stratégiques US massive
- Cessez-le-feu confirmé zone pétrolière

Réponds UNIQUEMENT avec ce JSON :
{{
  "signal": "BUY" | "SELL" | "NEUTRAL",
  "confidence": <entier 0-100>,
  "impact_timeframe": "immédiat" | "court_terme" | "moyen_terme",
  "summary": "<résumé exécutif 2-3 phrases>",
  "reasoning": "<raisonnement détaillé>",
  "price_impact_estimate": "<ex: +2% à +4%>",
  "key_factors": ["facteur1", "facteur2"]
}}"""

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]

    return json.loads(raw)

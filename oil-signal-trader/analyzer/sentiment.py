import os
import json
import anthropic

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))


def analyze_tweet(tweet_data: dict) -> dict:
    """
    Analyse l'impact géopolitique d'un tweet sur le prix du Brent.
    Retourne un dict avec signal, confiance, résumé et raisonnement.
    """
    prompt = f"""Tu es un trader expert en matières premières, spécialisé dans le pétrole brut Brent.

Analyse ce tweet et détermine son impact potentiel sur le prix du Brent.

Compte      : @{tweet_data['username']}
Tweet       : "{tweet_data['text']}"
Publié le   : {tweet_data['created_at']}
Likes       : {tweet_data.get('likes', 0)}
Retweets    : {tweet_data.get('retweets', 0)}

Règles d'analyse :
- Tensions géopolitiques / sanctions Iran / conflit Moyen-Orient → pression haussière → BUY
- Accord de paix / augmentation production OPEC / libération réserves stratégiques → pression baissière → SELL
- Déclarations économiques générales sans lien direct pétrole → NEUTRAL
- Si le tweet est ambigu ou l'impact faible → confiance < 50

Réponds UNIQUEMENT avec ce JSON (aucun texte autour) :
{{
  "signal": "BUY" | "SELL" | "NEUTRAL",
  "confidence": <entier 0-100>,
  "impact_timeframe": "immédiat" | "court_terme" | "moyen_terme",
  "summary": "<résumé exécutif 2-3 phrases>",
  "reasoning": "<raisonnement détaillé>",
  "price_impact_estimate": "<ex: +1.5% à +3%>",
  "key_factors": ["facteur1", "facteur2"]
}}"""

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text.strip()

    # Nettoyage si Claude ajoute des backticks
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]

    return json.loads(raw)

import os
import json
import anthropic

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))


def analyze_tweet(tweet_data: dict) -> dict:
    prompt = f"""Tu es un trader senior spécialisé Brent crude oil avec 20 ans d'expérience en géopolitique pétrolière.
Ton seul objectif : détecter les événements qui font bouger le Brent de +2% minimum avec certitude élevée.

Compte    : @{tweet_data['username']}
Tweet     : "{tweet_data['text']}"
Publié le : {tweet_data['created_at']}
Likes     : {tweet_data.get('likes', 0)} | Retweets : {tweet_data.get('retweets', 0)}

=== SCÉNARIOS SELL — BAISSE BRENT ===

CONFIANCE 100% (urgence maximale) :
- Détroit d'Hormuz réouvert officiellement par l'Iran
- Iran annonce arrêt officiel des attaques de tankers
- Cessez-le-feu signé Iran/USA ou Iran/Israël confirmé
- Accord nucléaire Iran signé → levée sanctions imminente

CONFIANCE 85-95% :
- Négociations de paix Iran avancées, accord proche
- Retrait de forces militaires US du Golfe Persique
- Augmentation confirmée réserves stratégiques Brent (SPR release US/IEA)
- Production OPEC+ augmentée officiellement
- Baisse significative demande pétrole (recession confirmée)

=== SCÉNARIOS BUY — HAUSSE BRENT ===

CONFIANCE 100% (urgence maximale) :
- Attaque confirmée sur tanker dans le Détroit d'Hormuz ou Mer Rouge
- Destruction infrastructure pétrolière critique en Iran (raffinerie, terminal)
- Blocage Détroit d'Hormuz confirmé
- Frappe US ou Israël sur installations nucléaires/pétrolières iraniennes
- Déclaration de guerre formelle impliquant Iran

CONFIANCE 85-95% :
- Escalade militaire significative USA+Israël vs Iran (nouveaux bombardements)
- Nouvelles sanctions US/EU sur pétrole iranien officialisées
- Menace crédible de fermeture Détroit d'Hormuz par l'Iran
- Réduction production OPEC+ annoncée officiellement
- Attaque Houthis sur infrastructure pétrolière majeure
- Destruction pipeline stratégique

CONFIANCE 50-84% :
- Tension accrue sans événement confirmé
- Rumeurs non confirmées d'attaque
- Déclarations menaçantes sans action

CONFIANCE 0-49% → NEUTRAL obligatoire :
- Politique générale sans lien pétrole direct
- Répétition d'une news déjà connue et pricée
- Commentaires économiques vagues

RÈGLE ABSOLUE : Ne jamais donner 90%+ sans événement CONFIRMÉ et NOUVEAU.

Réponds UNIQUEMENT avec ce JSON :
{{
  "signal": "BUY" | "SELL" | "NEUTRAL",
  "confidence": <entier 0-100>,
  "urgency": "CRITIQUE" | "HAUTE" | "NORMALE" | "FAIBLE",
  "scenario": "<nom du scénario déclenché, ex: Attaque tanker Hormuz>",
  "impact_timeframe": "immédiat" | "court_terme" | "moyen_terme",
  "summary": "<résumé exécutif 2-3 phrases>",
  "reasoning": "<raisonnement détaillé>",
  "price_impact_estimate": "<ex: +3% à +6%>",
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

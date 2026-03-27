import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

SIGNAL_COLORS = {
    "BUY":     {"bg": "#d4edda", "border": "#28a745", "emoji": "🟢"},
    "SELL":    {"bg": "#f8d7da", "border": "#dc3545", "emoji": "🔴"},
    "NEUTRAL": {"bg": "#fff3cd", "border": "#ffc107", "emoji": "🟡"},
}

URGENCY_STYLE = {
    "CRITIQUE": {"bg": "#dc3545", "text": "white",  "label": "🚨 URGENCE CRITIQUE"},
    "HAUTE":    {"bg": "#fd7e14", "text": "white",  "label": "⚠️ URGENCE HAUTE"},
    "NORMALE":  {"bg": "#ffc107", "text": "black",  "label": "🟡 NORMALE"},
    "FAIBLE":   {"bg": "#6c757d", "text": "white",  "label": "ℹ️ FAIBLE"},
}


def send_alert(tweet_data: dict, analysis: dict):
    signal = analysis["signal"]
    confidence = analysis["confidence"]
    urgency = analysis.get("urgency", "NORMALE")
    scenario = analysis.get("scenario", "")
    sig_style = SIGNAL_COLORS.get(signal, SIGNAL_COLORS["NEUTRAL"])
    urg_style = URGENCY_STYLE.get(urgency, URGENCY_STYLE["NORMALE"])

    subject = (
        f"{urg_style['label']} — 🛢️ BRENT {sig_style['emoji']} {signal} [{confidence}%] "
        f"— @{tweet_data['username']}"
    )

    html = f"""
    <html><body style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;">

      <!-- Urgence banner -->
      <div style="background:{urg_style['bg']};color:{urg_style['text']};
                  padding:12px 20px;border-radius:8px 8px 0 0;text-align:center;
                  font-size:16px;font-weight:bold;">
        {urg_style['label']} &nbsp;—&nbsp; {scenario}
      </div>

      <!-- Header -->
      <div style="background:#0d1b2a;color:white;padding:16px 20px;">
        <h2 style="margin:0;">🛢️ Oil Signal Trader — Brent Crude</h2>
      </div>

      <!-- Tweet -->
      <div style="background:#f8f9fa;padding:16px 20px;border-left:4px solid #1da1f2;">
        <p style="margin:0;font-size:12px;color:#555;">@{tweet_data['username']} · {tweet_data['created_at']}</p>
        <p style="margin:8px 0 0;font-size:15px;font-style:italic;">"{tweet_data['text']}"</p>
        <p style="margin:6px 0 0;font-size:12px;color:#888;">
          ❤️ {tweet_data.get('likes', 0)} &nbsp; 🔁 {tweet_data.get('retweets', 0)}
        </p>
      </div>

      <!-- Signal -->
      <div style="background:{sig_style['bg']};border-left:4px solid {sig_style['border']};
                  padding:16px 20px;">
        <h3 style="margin:0 0 8px;">
          {sig_style['emoji']} Signal : {signal} &nbsp;|&nbsp; Confiance : {confidence}%
        </h3>
        <p style="margin:0;">{analysis['summary']}</p>
      </div>

      <!-- Détails -->
      <div style="background:white;padding:16px 20px;border:1px solid #e0e0e0;">
        <h4 style="margin:0 0 10px;">Analyse détaillée</h4>
        <p style="margin:0 0 10px;">{analysis['reasoning']}</p>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          <tr>
            <td style="padding:6px 8px;background:#f5f5f5;width:40%;"><strong>Scénario</strong></td>
            <td style="padding:6px 8px;">{scenario}</td>
          </tr>
          <tr>
            <td style="padding:6px 8px;background:#f5f5f5;"><strong>Impact estimé</strong></td>
            <td style="padding:6px 8px;">{analysis.get('price_impact_estimate', 'N/A')}</td>
          </tr>
          <tr>
            <td style="padding:6px 8px;background:#f5f5f5;"><strong>Horizon</strong></td>
            <td style="padding:6px 8px;">{analysis.get('impact_timeframe', 'N/A')}</td>
          </tr>
          <tr>
            <td style="padding:6px 8px;background:#f5f5f5;"><strong>Facteurs clés</strong></td>
            <td style="padding:6px 8px;">{', '.join(analysis.get('key_factors', []))}</td>
          </tr>
        </table>
      </div>

      <!-- Footer -->
      <div style="background:#f5f5f5;padding:10px 20px;border-radius:0 0 8px 8px;
                  font-size:11px;color:#888;">
        ⚠️ Analyse automatisée — pas de conseil financier.
        <br>Oil Signal Trader · alexandre@defranca.ch
      </div>

    </body></html>
    """

    message = Mail(
        from_email=os.getenv("FROM_EMAIL"),
        to_emails=os.getenv("ALERT_EMAIL"),
        subject=subject,
        html_content=html,
    )

    sg = SendGridAPIClient(os.getenv("SENDGRID_API_KEY"))
    response = sg.send(message)
    print(f"Email [{response.status_code}] : {subject}")

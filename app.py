from flask import Flask, render_template, request, jsonify, session
from groq import Groq
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'sahara-dev-secret-2026')

client = Groq(api_key=os.environ.get('GROQ_API_KEY'))

CRISIS_KEYWORDS = [
    'suicide', 'suicidal', 'kill myself', 'end my life', 'want to die',
    'self harm', 'self-harm', 'hurt myself', 'cutting myself', 'no reason to live',
    "can't go on", 'cannot go on', 'worthless', 'better off dead', 'end it all',
    'take my life', 'not worth living', 'ending it', 'want to end it',
    'harming myself', 'harm myself', 'no point in living', 'nothing to live for',
    "don't want to be here", 'dont want to be here', 'rather be dead',
    'wish i was dead', 'life is not worth', 'give up on life', 'no hope left',
    'disappear forever', 'world without me'
]

SYSTEM_PROMPT = """You are Sahara, a warm and supportive peer listener chatbot for college students at IIT Jodhpur. "Sahara" (सहारा) means support and refuge in Hindi.

Your role:
- Listen empathetically to students experiencing stress, anxiety, loneliness, exam pressure, or burnout
- Acknowledge the specific feeling the student expressed BEFORE offering any advice
- Ask ONE thoughtful follow-up question to understand the student's situation more deeply
- Offer genuine emotional support, practical coping strategies, and mental wellness tips
- Suggest healthy habits, breathing exercises, or study break techniques when relevant
- Vary your responses — never repeat the same suggestion in consecutive messages
- Keep responses warm, conversational, and concise (2-4 sentences)

Strict boundaries:
- You are NOT a therapist or doctor. NEVER diagnose any mental health condition.
- NEVER recommend or mention any medications.
- You are a first point of contact — a supportive peer, not a clinician.
- If the student expresses persistent or serious distress, gently suggest speaking with the IIT Jodhpur counselor at counselor@iitj.ac.in or calling iCall: 9152987821.
- Never dismiss or minimise what the student is feeling, even if it seems minor.

Tone: Like a caring, understanding batchmate — warm, non-judgmental, real. Simple English. Avoid clinical or overly formal language."""


def detect_crisis(message: str) -> bool:
    return any(kw in message.lower() for kw in CRISIS_KEYWORDS)


@app.route('/')
def index():
    if 'messages' not in session:
        session['messages'] = []
    if 'mood_log' not in session:
        session['mood_log'] = []
    return render_template('index.html')


@app.route('/chat', methods=['POST'])
def chat():
    data = request.get_json()
    user_message = (data.get('message') or '').strip()
    mood_context = data.get('mood', '')

    if not user_message:
        return jsonify({'error': 'Empty message'}), 400

    # Layer 3 — Safety Filter
    if detect_crisis(user_message):
        return jsonify({
            'response': (
                "I'm really glad you reached out — you are not alone in this. "
                "What you're sharing sounds very serious, and I want to make sure "
                "you get the right support right now. Please reach out to one of the resources below."
            ),
            'is_crisis': True,
            'resources': [
                {'label': 'iCall (National Helpline)', 'value': '9152987821'},
                {'label': 'NIMHANS Helpline', 'value': '080-46110007'},
                {'label': 'IIT Jodhpur Student Wellness Cell', 'value': 'counselor@iitj.ac.in'},
                {'label': 'Vandrevala Foundation (24x7)', 'value': '1860-2662-345'},
            ]
        })

    # Layer 2 — AI Brain
    if 'messages' not in session:
        session['messages'] = []

    messages = list(session['messages'])
    messages.append({'role': 'user', 'content': user_message})
    messages = messages[-12:]

    # Inject mood context into system prompt if available
    system = SYSTEM_PROMPT
    if mood_context:
        system += f"\n\nNote: The student has indicated their current mood is: {mood_context}. Keep this in mind."

    try:
        response = client.chat.completions.create(
            model='llama-3.3-70b-versatile',
            max_tokens=350,
            messages=[{'role': 'system', 'content': system}] + messages,
        )
        reply = response.choices[0].message.content
    except Exception as e:
        return jsonify({'error': f'AI error: {str(e)}'}), 500

    messages.append({'role': 'assistant', 'content': reply})
    session['messages'] = messages
    session.modified = True

    return jsonify({'response': reply, 'is_crisis': False})


@app.route('/mood', methods=['POST'])
def log_mood():
    """Layer 4 — Mood Tracker: log a mood entry for the session."""
    data = request.get_json()
    mood = data.get('mood', '')
    score = data.get('score', 3)

    if 'mood_log' not in session:
        session['mood_log'] = []

    entry = {
        'mood': mood,
        'score': score,
        'time': datetime.now().strftime('%H:%M'),
        'date': datetime.now().strftime('%b %d')
    }
    log = list(session['mood_log'])
    log.append(entry)
    session['mood_log'] = log[-10:]  # keep last 10
    session.modified = True

    return jsonify({'status': 'ok', 'log': session['mood_log']})


@app.route('/mood', methods=['GET'])
def get_mood():
    return jsonify({'log': session.get('mood_log', [])})


@app.route('/report', methods=['POST'])
def report():
    data = request.get_json()
    flagged = {
        'message': (data.get('message') or '')[:500],
        'time': datetime.now().isoformat()
    }
    if 'flagged_responses' not in session:
        session['flagged_responses'] = []
    flags = list(session.get('flagged_responses', []))
    flags.append(flagged)
    session['flagged_responses'] = flags[-10:]
    session.modified = True
    return jsonify({'status': 'ok'})


@app.route('/reset', methods=['POST'])
def reset():
    session.clear()
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

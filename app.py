from flask import Flask, render_template, request, jsonify, session
import anthropic
import os

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'manas-dev-secret-2026')

client = anthropic.Anthropic(api_key=os.environ.get('ANTHROPIC_API_KEY'))

CRISIS_KEYWORDS = [
    'suicide', 'suicidal', 'kill myself', 'end my life', 'want to die',
    'self harm', 'self-harm', 'hurt myself', 'cutting myself', 'no reason to live',
    'can\'t go on', 'cannot go on', 'worthless', 'better off dead', 'end it all',
    'take my life', 'not worth living'
]

SYSTEM_PROMPT = """You are Sahara, a warm and supportive peer listener chatbot designed for college students at IIT Jodhpur. "Sahara" (सहारा) means support and refuge in Hindi.

Your role:
- Listen empathetically to students experiencing stress, anxiety, loneliness, exam pressure, or burnout
- Offer genuine emotional support, practical coping strategies, and mental wellness tips
- Ask thoughtful follow-up questions to understand the student better
- Suggest healthy habits, breathing exercises, or study break techniques when appropriate
- Keep responses warm, conversational, and concise (2-4 sentences)

Strict boundaries:
- You are NOT a therapist or doctor. NEVER diagnose any mental health condition.
- NEVER recommend or mention any medications.
- You are a first point of contact — a supportive peer, not a clinician.
- If the student needs more support than you can offer, gently suggest speaking with the IIT Jodhpur counselor.

Tone: Like a caring, understanding batchmate — warm, non-judgmental, and real. Avoid clinical or overly formal language. Use simple English."""


def detect_crisis(message: str) -> bool:
    msg = message.lower()
    return any(keyword in msg for keyword in CRISIS_KEYWORDS)


@app.route('/')
def index():
    if 'messages' not in session:
        session['messages'] = []
    return render_template('index.html')


@app.route('/chat', methods=['POST'])
def chat():
    data = request.get_json()
    user_message = (data.get('message') or '').strip()

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
            ]
        })

    # Layer 2 — AI Brain
    if 'messages' not in session:
        session['messages'] = []

    messages = list(session['messages'])
    messages.append({'role': 'user', 'content': user_message})

    # Keep last 12 turns to avoid token overflow
    messages = messages[-12:]

    try:
        response = client.messages.create(
            model='claude-haiku-4-5-20251001',
            max_tokens=350,
            system=SYSTEM_PROMPT,
            messages=messages,
        )
        reply = response.content[0].text
    except Exception as e:
        return jsonify({'error': f'AI error: {str(e)}'}), 500

    messages.append({'role': 'assistant', 'content': reply})
    session['messages'] = messages
    session.modified = True

    return jsonify({'response': reply, 'is_crisis': False})


@app.route('/reset', methods=['POST'])
def reset():
    session.clear()
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

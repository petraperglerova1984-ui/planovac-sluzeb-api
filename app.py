"""
Flask API pro plánování služeb - čte a píše do Google Sheets
"""
from flask import Flask, request, jsonify
from flask_cors import CORS
import os
from planner_sheets import plan_shifts_from_sheets

app = Flask(__name__)
CORS(app)

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        "status": "ok",
        "message": "Plánovač služeb API běží",
        "version": "1.0"
    })

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "healthy"})

@app.route('/plan', methods=['POST'])
def plan():
    """
    Endpoint pro spuštění plánování
    """
    try:
        # Získej parametry z requestu (pokud nějaké jsou)
        data = request.get_json() or {}
        
        # Spusť plánování
        result = plan_shifts_from_sheets()
        
        return jsonify({
            "status": "success",
            "message": "Plánování dokončeno úspěšně",
            "details": result
        })
    
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

"""
Flask API V2 - Jednodušší a čistší
"""
from flask import Flask, request, jsonify
from flask_cors import CORS
import os
from planner_sheets_v2 import plan_shifts_v2

app = Flask(__name__)
CORS(app)

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        "status": "ok",
        "message": "Plánovač služeb V2 API",
        "version": "2.0"
    })

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "healthy"})

@app.route('/plan', methods=['POST'])
def plan():
    """
    Endpoint pro plánování
    Očekává: { "sheet_name": "CERVEN" }
    """
    try:
        data = request.get_json() or {}
        sheet_name = data.get('sheet_name', 'CERVEN')
        
        print(f"Přijat request pro plánování: {sheet_name}")
        
        result = plan_shifts_v2(sheet_name)
        
        return jsonify({
            "status": "success",
            "message": f"Plánování dokončeno pro {sheet_name}",
            "details": result
        })
    
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"CHYBA: {error_details}")
        
        return jsonify({
            "status": "error",
            "message": str(e),
            "details": error_details
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

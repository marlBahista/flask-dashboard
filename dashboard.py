from flask import Flask, render_template, jsonify, request, send_from_directory
from flask_socketio import SocketIO
import threading
import time
import os
import subprocess

app = Flask(__name__)
socketio = SocketIO(app)

# Global Variables
total_coins = 0
remaining_paper = 18  # Default matches PAPER_FEED_CAPACITY from printUI16.py
ink_status = "Checking..."
transactions = []  # Store printing history

UPLOAD_FOLDER = os.path.join(os.getcwd(), "static")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/upload_screenshot", methods=["POST"])
def upload_screenshot():
    file = request.files["file"]
    if file:
        file.save(os.path.join(UPLOAD_FOLDER, "capture.png"))
        return {"message": "Screenshot uploaded"}, 200
    return {"message": "No file received"}, 400

@app.route("/get_screenshot")
def get_screenshot():
    return send_from_directory(UPLOAD_FOLDER, "capture.png")

def check_ink_levels():
    global ink_status
    time.sleep(5)  # Simulate checking process
    ink_status = "✅ Ink Levels OK"
    socketio.emit("update_ink", {"ink_status": ink_status})

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/get_status")
def get_status():
    """Retrieve the latest status, including remaining paper."""
    return jsonify({
        "total_coins": total_coins,
        "remaining_paper": remaining_paper,
        "ink_status": ink_status,
        "transactions": transactions
    })

@app.route("/update_paper", methods=["POST"])
def update_paper():
    """Update paper count from printUI16.py and sync with the dashboard."""
    global remaining_paper
    data = request.get_json()
    remaining_paper = data["remaining_paper"]
    
    # Emit real-time update to dashboard
    socketio.emit("update_paper", {"remaining_paper": remaining_paper})
    
    return jsonify({"message": "Paper count updated"}), 200

@app.route("/add_transaction", methods=["POST"])
def add_transaction():
    """Log new transactions and update total earnings."""
    global total_coins
    data = request.get_json()
    transactions.append(data)
    total_coins += data["cost"]  # Update total earnings

    socketio.emit("update_transactions", {"transactions": transactions})
    socketio.emit("update_coins", {"total_coins": total_coins})
    return jsonify({"message": "Transaction logged"}), 200

@app.route("/log_transaction", methods=["POST"])
def log_transaction():
    """Log messages from printUI16.py"""
    data = request.get_json()
    print(f"📋 Log: {data['message']}")
    return jsonify({"message": "Log received"}), 200

if __name__ == "__main__":
    threading.Thread(target=check_ink_levels, daemon=True).start()
    socketio.run(app, host="0.0.0.0", port=5000)

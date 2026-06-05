from flask import Flask, request, send_file, jsonify
import json, os, tempfile
from generate_kickoff import generate

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate_pptx():
    config = request.get_json()
    if not config:
        return jsonify({"error": "No JSON body"}), 400

    with tempfile.TemporaryDirectory() as tmp:
        output_path = os.path.join(tmp, f"Kick-Off_{config.get('brand_name','output')}.pptx")
        try:
            generate(config, output_path)
        except Exception as e:
            return jsonify({"error": str(e)}), 500

        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=f"Kick-Off_{config.get('brand_name','output')}.pptx"
        )

@app.route('/', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))

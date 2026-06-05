from flask import Flask, request, send_file, jsonify
import json, os, tempfile, traceback
from generate_kickoff import generate

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate_pptx():
    config = request.get_json()
    if not config:
        return jsonify({"error": "No JSON body"}), 400

    print(f"CONFIG RECEIVED: {json.dumps(config, ensure_ascii=False)}", flush=True)

    with tempfile.TemporaryDirectory() as tmp:
        output_path = os.path.join(tmp, f"Kick-Off_{config.get('brand_name','output')}.pptx")
        try:
            generate(config, output_path)
        except Exception as e:
            tb = traceback.format_exc()
            print(f"ERROR:\n{tb}", flush=True)
            return jsonify({"error": str(e), "traceback": tb}), 500

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

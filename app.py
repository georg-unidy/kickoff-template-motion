from flask import Flask, request, send_file
import json, os, subprocess, tempfile, shutil

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    config = request.get_json()
    
    with tempfile.TemporaryDirectory() as tmp:
        config_path = os.path.join(tmp, 'config.json')
        output_path = os.path.join(tmp, 'kickoff.pptx')
        
        with open(config_path, 'w') as f:
            json.dump(config, f)
        
        subprocess.run([
            'python', 'generate_kickoff.py',
            '--config', config_path,
            '--output', output_path
        ], check=True)
        
        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=f"Kick-Off_{config.get('brand_name','')}.pptx"
        )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
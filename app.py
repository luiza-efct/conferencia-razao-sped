"""
HUB EFCT — Conferência Razão × SPED
Backend Flask que recebe os 4 arquivos via multipart e devolve o xlsx cruzado.
"""
import io
import os
import logging
import traceback

from flask import Flask, request, send_file, jsonify, Response

from logica import processar_cruzamento

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB total upload

logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s %(message)s')


@app.route('/')
def index():
    """Serve o index.html da raiz."""
    path = os.path.join(BASE_DIR, 'index.html')
    with open(path, 'r', encoding='utf-8') as f:
        return Response(f.read(), mimetype='text/html; charset=utf-8')


@app.route('/health')
def health():
    """Health check para o Render."""
    return jsonify({'status': 'ok'})


@app.route('/api/processar', methods=['POST'])
def processar():
    """
    Recebe 4 arquivos (razao, c100efd, a100, f100) via multipart/form-data
    e devolve o xlsx cruzado com as 2 abas (RAZÃO CONTABIL + PENDÊNCIAS DE CRUZAMENTO).
    """
    try:
        razao = request.files.get('razao')
        c100efd = request.files.get('c100efd')
        a100 = request.files.get('a100')
        f100 = request.files.get('f100')
        skip_cnae = request.form.get('skip_cnae', 'false').lower() in ('true', '1', 'on')

        if not all([razao, c100efd, a100, f100]):
            return jsonify({
                'erro': 'Os 4 arquivos são obrigatórios: razao, c100efd, a100, f100',
                'recebidos': {
                    'razao': bool(razao), 'c100efd': bool(c100efd),
                    'a100': bool(a100), 'f100': bool(f100)
                }
            }), 400

        app.logger.info(
            'Iniciando cruzamento: razao=%s c100efd=%s a100=%s f100=%s skip_cnae=%s',
            razao.filename, c100efd.filename, a100.filename, f100.filename, skip_cnae
        )

        out_bytes, stats = processar_cruzamento(
            razao_stream=razao.stream,
            c100efd_stream=c100efd.stream,
            a100_stream=a100.stream,
            f100_stream=f100.stream,
            skip_cnae=skip_cnae,
        )

        # Nome do arquivo de saída derivado do nome da Razão
        original = razao.filename or 'razao'
        base = original.rsplit('.', 1)[0]
        out_filename = f'{base} - CRUZADO.xlsx'

        app.logger.info(
            'Cruzamento OK: %s linhas · matches=%s · pendencias_cnpjs=%s',
            stats.get('linhas_razao'), stats.get('match_any'), stats.get('pendencias')
        )

        response = send_file(
            io.BytesIO(out_bytes),
            as_attachment=True,
            download_name=out_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        # Expõe stats no header para a UI ler
        for k, v in stats.items():
            response.headers[f'X-Stat-{k}'] = str(v)
        response.headers['Access-Control-Expose-Headers'] = ', '.join(
            f'X-Stat-{k}' for k in stats.keys()
        )
        return response

    except ValueError as e:
        # Erros de validação dos arquivos
        app.logger.warning('Validação falhou: %s', e)
        return jsonify({'erro': str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        app.logger.error('Erro no processamento: %s', e)
        return jsonify({
            'erro': f'Erro interno no processamento: {str(e)}'
        }), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=False)

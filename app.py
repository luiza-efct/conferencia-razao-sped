"""
HUB EFCT — Conferência Razão × SPED
Backend Flask que recebe os 4 arquivos via multipart e devolve o xlsx cruzado.
"""
import io
import os
import logging
import traceback

from flask import Flask, request, send_file, jsonify, Response

from logica import processar_cruzamento, preview_padrao

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


def _real(fs):
    """Returns FileStorage só se um arquivo real foi enviado (filename não vazio)."""
    if fs is None or not getattr(fs, 'filename', None) or not fs.filename.strip():
        return None
    return fs


@app.route('/api/preview', methods=['POST'])
def preview():
    """
    Etapa de CONFERÊNCIA do padrão: recebe só a Razão + um índice de estratégia e
    devolve uma amostra (Complemento Histórico → NF + Razão Social) pra a usuária
    aprovar ("OK") ou pedir outra leitura ("Repensar" → próxima estratégia).
    Não processa os blocos — é rápido.
    """
    try:
        razao = _real(request.files.get('razao'))
        if not razao:
            return jsonify({'erro': 'Envie a Razão Contábil para conferir o padrão.'}), 400
        try:
            estrategia = int(request.form.get('estrategia', 0))
        except (TypeError, ValueError):
            estrategia = 0
        ex_hist = (request.form.get('exemplo_historico') or '').strip() or None
        ex_razao = (request.form.get('exemplo_razao') or '').strip() or None
        ex_nf = (request.form.get('exemplo_nf') or '').strip() or None
        resultado = preview_padrao(razao.stream, estrategia=estrategia,
                                   exemplo_historico=ex_hist, exemplo_razao=ex_razao,
                                   exemplo_nf=ex_nf)
        return jsonify(resultado)
    except ValueError as e:
        app.logger.warning('Preview falhou: %s', e)
        return jsonify({'erro': str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        app.logger.error('Erro no preview: %s', e)
        return jsonify({'erro': f'Erro ao gerar a prévia: {str(e)}'}), 500


@app.route('/api/processar', methods=['POST'])
def processar():
    """
    Recebe 4 arquivos (razao, c100efd, a100, f100) via multipart/form-data
    e devolve o xlsx cruzado com as 2 abas (RAZÃO CONTABIL + PENDÊNCIAS DE CRUZAMENTO).
    """
    try:
        razao = _real(request.files.get('razao'))
        c100efd = _real(request.files.get('c100efd'))
        a100 = _real(request.files.get('a100'))  # opcional
        f100 = _real(request.files.get('f100'))  # opcional
        # Consulta CNPJ→Razão Social/CNAE é SEMPRE feita quando há CNPJ identificado

        # Apenas razao + c100efd são obrigatórios. A100/F100 são opcionais —
        # empresas que não escrituram esses blocos podem processar mesmo assim.
        if not razao or not c100efd:
            return jsonify({
                'erro': 'Os arquivos Razão Contábil e C100-EFD FISCAL são obrigatórios. '
                        'A100-CONTRI e F100-CONTRI são opcionais.',
                'recebidos': {
                    'razao': bool(razao), 'c100efd': bool(c100efd),
                    'a100': bool(a100), 'f100': bool(f100)
                }
            }), 400

        app.logger.info(
            'Iniciando cruzamento: razao=%s c100efd=%s a100=%s f100=%s',
            razao.filename, c100efd.filename,
            a100.filename if a100 else '(não enviado)',
            f100.filename if f100 else '(não enviado)',
        )

        try:
            estrategia = int(request.form.get('estrategia', 0))
        except (TypeError, ValueError):
            estrategia = 0
        ex_hist = (request.form.get('exemplo_historico') or '').strip() or None
        ex_razao = (request.form.get('exemplo_razao') or '').strip() or None
        ex_nf = (request.form.get('exemplo_nf') or '').strip() or None

        out_bytes, stats = processar_cruzamento(
            razao_stream=razao.stream,
            c100efd_stream=c100efd.stream,
            a100_stream=a100.stream if a100 else None,
            f100_stream=f100.stream if f100 else None,
            estrategia=estrategia,
            exemplo_historico=ex_hist,
            exemplo_razao=ex_razao,
            exemplo_nf=ex_nf,
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

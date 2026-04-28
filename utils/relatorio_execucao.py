import csv
import time
from datetime import datetime
from pathlib import Path

from utils.paths import appdata_dir

try:
    from version import VERSAO_ATUAL
except Exception:
    VERSAO_ATUAL = "desconhecida"


CSV_CAMPOS = [
    "id_execucao", "versao", "status", "inicio", "fim", "data_corte",
    "pasta_origem", "pasta_destino", "duracao_segundos", "duracao_formatada",
    "pdfs_encontrados", "pdfs_fora_da_data", "pdfs_ignorados", "pdfs_sem_dados",
    "pdfs_processados", "arquivos_organizados", "arquivos_ja_existentes",
    "pdfs_movidos_para_erros", "falhas_ao_mover_para_erros", "falhas_organizacao",
    "raes_lancados", "duplicidades_barradas", "cnpjs_nao_encontrados",
    "falhas_definitivas", "erros_pendencias", "tempo_medio_segundos_por_rae",
    "tempo_medio_minutos_por_rae", "tempo_medio_segundos_por_pdf_processado",
    "tempo_medio_minutos_por_pdf_processado",
]


def _agora_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def novo_resumo_execucao(pasta_origem: str, pasta_destino: str, data_corte: str) -> dict:
    return {
        "id_execucao": _agora_id(),
        "versao": VERSAO_ATUAL,
        "status": "em_execucao",
        "inicio": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "fim": "",
        "duracao_segundos": 0,
        "_inicio_perf": time.perf_counter(),
        "pasta_origem": pasta_origem,
        "pasta_destino": pasta_destino,
        "data_corte": data_corte,
        "pdfs_encontrados": 0,
        "pdfs_fora_da_data": 0,
        "pdfs_ignorados": 0,
        "pdfs_sem_dados": 0,
        "pdfs_processados": 0,
        "arquivos_organizados": 0,
        "arquivos_ja_existentes": 0,
        "pdfs_movidos_para_erros": 0,
        "falhas_ao_mover_para_erros": 0,
        "falhas_organizacao": 0,
        "raes_lancados": 0,
        "duplicidades_barradas": 0,
        "cnpjs_nao_encontrados": 0,
        "falhas_definitivas": 0,
        "erros": [],
    }


def finalizar_resumo_execucao(resumo: dict, status: str) -> dict:
    resumo["status"] = status
    resumo["fim"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    inicio_perf = resumo.pop("_inicio_perf", None)
    if inicio_perf is not None:
        resumo["duracao_segundos"] = round(time.perf_counter() - inicio_perf, 2)

    resumo.update(calcular_metricas_tempo(resumo))
    salvar_resumo_txt(resumo)
    salvar_relatorio_individual_txt(resumo)
    salvar_historico_csv(resumo)
    return resumo


def calcular_metricas_tempo(resumo: dict) -> dict:
    duracao = float(resumo.get("duracao_segundos", 0) or 0)
    raes_lancados = int(resumo.get("raes_lancados", 0) or 0)
    pdfs_processados = int(resumo.get("pdfs_processados", 0) or 0)

    tempo_medio_rae = round(duracao / raes_lancados, 2) if raes_lancados > 0 else 0
    tempo_medio_pdf = round(duracao / pdfs_processados, 2) if pdfs_processados > 0 else 0

    return {
        "tempo_medio_segundos_por_rae": tempo_medio_rae,
        "tempo_medio_minutos_por_rae": round(tempo_medio_rae / 60, 2) if tempo_medio_rae else 0,
        "tempo_medio_segundos_por_pdf_processado": tempo_medio_pdf,
        "tempo_medio_minutos_por_pdf_processado": round(tempo_medio_pdf / 60, 2) if tempo_medio_pdf else 0,
    }


def formatar_duracao(segundos: float | int) -> str:
    segundos = int(segundos or 0)
    horas, resto = divmod(segundos, 3600)
    minutos, seg = divmod(resto, 60)
    if horas:
        return f"{horas}h {minutos}min {seg}s"
    if minutos:
        return f"{minutos}min {seg}s"
    return f"{seg}s"


def caminho_relatorio_execucao() -> Path:
    return appdata_dir() / "ultimo_relatorio_execucao.txt"


def caminho_pasta_relatorios() -> Path:
    pasta = appdata_dir() / "relatorios"
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def caminho_historico_execucoes() -> Path:
    return caminho_pasta_relatorios() / "historico_execucoes.csv"


def caminho_relatorio_individual(resumo: dict) -> Path:
    id_execucao = resumo.get("id_execucao") or _agora_id()
    return caminho_pasta_relatorios() / f"relatorio_{id_execucao}.txt"


def _fmt_numero_planilha(valor) -> str:
    if isinstance(valor, float):
        return f"{valor:.2f}".replace(".", ",")
    return str(valor)


def _linha_csv(resumo: dict) -> dict:
    erros = resumo.get("erros", []) or []
    linha = {}
    for campo in CSV_CAMPOS:
        if campo == "duracao_formatada":
            valor = formatar_duracao(resumo.get("duracao_segundos", 0))
        elif campo == "erros_pendencias":
            valor = len(erros)
        else:
            valor = resumo.get(campo, 0)
        linha[campo] = _fmt_numero_planilha(valor)
    return linha


def formatar_resumo_para_usuario(resumo: dict) -> str:
    if not resumo:
        return "Processo concluído, mas nenhum resumo detalhado foi gerado."

    tempo_medio_rae = resumo.get("tempo_medio_segundos_por_rae", 0)
    tempo_medio_pdf = resumo.get("tempo_medio_segundos_por_pdf_processado", 0)
    texto_tempo_medio_rae = formatar_duracao(tempo_medio_rae) if tempo_medio_rae else "Não calculado"
    texto_tempo_medio_pdf = formatar_duracao(tempo_medio_pdf) if tempo_medio_pdf else "Não calculado"

    linhas = [
        "Resumo da execução", "",
        f"Status: {resumo.get('status', '-')}",
        f"Versão: {resumo.get('versao', VERSAO_ATUAL)}",
        f"Início: {resumo.get('inicio', '-')}",
        f"Fim: {resumo.get('fim', '-')}",
        f"Tempo total: {formatar_duracao(resumo.get('duracao_segundos', 0))}",
        f"Tempo médio por RAE lançado: {texto_tempo_medio_rae}",
        f"Tempo médio por PDF processado: {texto_tempo_medio_pdf}",
        "",
        f"PDFs encontrados: {resumo.get('pdfs_encontrados', 0)}",
        f"PDFs fora da data: {resumo.get('pdfs_fora_da_data', 0)}",
        f"PDFs ignorados: {resumo.get('pdfs_ignorados', 0)}",
        f"PDFs sem dados suficientes: {resumo.get('pdfs_sem_dados', 0)}",
        f"PDFs processados: {resumo.get('pdfs_processados', 0)}",
        "",
        f"PDFs movidos para destino: {resumo.get('arquivos_organizados', 0)}",
        f"Arquivos já existentes no destino: {resumo.get('arquivos_ja_existentes', 0)}",
        f"PDFs enviados para pasta de erros: {resumo.get('pdfs_movidos_para_erros', 0)}",
        f"Falhas ao mover para pasta de erros: {resumo.get('falhas_ao_mover_para_erros', 0)}",
        f"Falhas ao organizar após RAE lançado: {resumo.get('falhas_organizacao', 0)}",
        "",
        f"RAEs lançados com sucesso: {resumo.get('raes_lancados', 0)}",
        f"Duplicidades barradas: {resumo.get('duplicidades_barradas', 0)}",
        f"CNPJs não encontrados: {resumo.get('cnpjs_nao_encontrados', 0)}",
        f"Falhas definitivas: {resumo.get('falhas_definitivas', 0)}",
        f"Erros/pendências: {len(resumo.get('erros', []))}",
    ]

    erros = resumo.get("erros", [])
    if erros:
        linhas.extend(["", "Pendências:"])
        for erro in erros[:10]:
            if isinstance(erro, dict):
                linhas.append(f"- {erro.get('cnpj', '-')}: {erro.get('motivo', '-')}")
            else:
                linhas.append(f"- {erro}")
        if len(erros) > 10:
            linhas.append(f"- ... e mais {len(erros) - 10} pendência(s).")

    linhas.extend(["", f"Último relatório salvo em: {caminho_relatorio_execucao()}", f"Histórico CSV salvo em: {caminho_historico_execucoes()}"])
    return "\n".join(linhas)


def salvar_resumo_txt(resumo: dict) -> Path | None:
    try:
        caminho = caminho_relatorio_execucao()
        caminho.write_text(formatar_resumo_para_usuario(resumo), encoding="utf-8")
        return caminho
    except Exception:
        return None


def salvar_relatorio_individual_txt(resumo: dict) -> Path | None:
    try:
        caminho = caminho_relatorio_individual(resumo)
        caminho.write_text(formatar_resumo_para_usuario(resumo), encoding="utf-8")
        return caminho
    except Exception:
        return None


def salvar_historico_csv(resumo: dict) -> Path | None:
    try:
        caminho = caminho_historico_execucoes()
        novo_arquivo = not caminho.exists()
        with caminho.open("a", newline="", encoding="utf-8-sig") as arquivo:
            writer = csv.DictWriter(arquivo, fieldnames=CSV_CAMPOS, delimiter=";")
            if novo_arquivo:
                writer.writeheader()
            writer.writerow(_linha_csv(resumo))
        return caminho
    except Exception:
        return None

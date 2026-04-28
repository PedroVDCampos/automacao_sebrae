import time
from datetime import datetime
from pathlib import Path

from utils.paths import appdata_dir


def novo_resumo_execucao(pasta_origem: str, pasta_destino: str, data_corte: str) -> dict:
    """Cria o objeto de resumo da execução.

    A chave _inicio_perf é interna e é removida ao finalizar o resumo.
    """
    return {
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
    salvar_resumo_txt(resumo)
    return resumo


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


def formatar_resumo_para_usuario(resumo: dict) -> str:
    if not resumo:
        return "Processo concluído, mas nenhum resumo detalhado foi gerado."

    linhas = [
        "Resumo da execução",
        "",
        f"Status: {resumo.get('status', '-')}",
        f"Início: {resumo.get('inicio', '-')}",
        f"Fim: {resumo.get('fim', '-')}",
        f"Tempo total: {formatar_duracao(resumo.get('duracao_segundos', 0))}",
        "",
        f"PDFs encontrados: {resumo.get('pdfs_encontrados', 0)}",
        f"PDFs fora da data: {resumo.get('pdfs_fora_da_data', 0)}",
        f"PDFs ignorados: {resumo.get('pdfs_ignorados', 0)}",
        f"PDFs sem dados suficientes: {resumo.get('pdfs_sem_dados', 0)}",
        f"PDFs processados: {resumo.get('pdfs_processados', 0)}",
        "",
        f"Arquivos organizados: {resumo.get('arquivos_organizados', 0)}",
        f"Arquivos já existentes no destino: {resumo.get('arquivos_ja_existentes', 0)}",
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

    linhas.extend(["", f"Relatório salvo em: {caminho_relatorio_execucao()}"])
    return "\n".join(linhas)


def salvar_resumo_txt(resumo: dict) -> Path | None:
    try:
        caminho = caminho_relatorio_execucao()
        caminho.write_text(formatar_resumo_para_usuario(resumo), encoding="utf-8")
        return caminho
    except Exception:
        return None

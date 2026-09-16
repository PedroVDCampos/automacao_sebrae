import csv, os
from datetime import datetime
from utils.logger import configurar_logger
from utils.paths import caminho_historico
logger = configurar_logger()
CAMPOS = ["id_execucao","versao","inicio","fim","duracao_segundos","status","pasta_origem","pasta_destino","data_corte","pdfs_encontrados","pdfs_fora_da_data","pdfs_processados","pdfs_ignorados","pdfs_sem_dados","arquivos_organizados","arquivos_ja_existentes","arquivos_com_erro","formalizacoes","alteracoes","declaracoes","boletos_das","parcelamentos","baixas","erros"]
def novo_resumo_execucao(origem,destino,data_corte,versao):
    agora=datetime.now()
    return {"id_execucao":agora.strftime("%Y%m%d_%H%M%S_%f"),"versao":versao,"inicio":agora.isoformat(timespec="seconds"),"fim":"","duracao_segundos":0.0,"status":"em_andamento","pasta_origem":origem,"pasta_destino":destino,"data_corte":data_corte,"pdfs_encontrados":0,"pdfs_fora_da_data":0,"pdfs_processados":0,"pdfs_ignorados":0,"pdfs_sem_dados":0,"arquivos_organizados":0,"arquivos_ja_existentes":0,"arquivos_com_erro":0,"formalizacoes":0,"alteracoes":0,"declaracoes":0,"boletos_das":0,"parcelamentos":0,"baixas":0,"erros":[]}
def finalizar_resumo_execucao(resumo,status):
    fim=datetime.now(); resumo["fim"]=fim.isoformat(timespec="seconds")
    try: resumo["duracao_segundos"]=round((fim-datetime.fromisoformat(resumo["inicio"])).total_seconds(),3)
    except Exception: pass
    resumo["status"]=status; salvar_historico(resumo); return resumo
def salvar_historico(resumo):
    caminho=caminho_historico(); novo=not os.path.exists(caminho)
    linha={c:(" | ".join(map(str,resumo.get(c,[]))) if c=="erros" else resumo.get(c,"")) for c in CAMPOS}
    try:
        with open(caminho,"a",newline="",encoding="utf-8-sig") as f:
            w=csv.DictWriter(f,fieldnames=CAMPOS)
            if novo:w.writeheader()
            w.writerow(linha)
    except Exception as e: logger.error("Não foi possível salvar histórico: %s",e)

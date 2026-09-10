#!/usr/bin/env python3
# Requer: pandas, openpyxl
# Uso: python cruzamento_planilhas.py
import pandas as pd
import numpy as np
from collections import Counter
import datetime
import re

# Arquivos esperados (mude se necessário)
FILE_PROF = "Professores e códigos.xlsx"
FILE_PASTA = "Pasta.xlsx"
FILE_TURMAS = "Turmas.xlsx"

# Palavras a excluir (confirmadas)
EXCLUDE_KEYWORDS = ["projeto de extensão","tcc","tópicos","topicos","estágio","estagio"]

def contains_exclude(s):
    if pd.isna(s):
        return False
    s = str(s).lower()
    return any(k in s for k in EXCLUDE_KEYWORDS)

def norm(s):
    if pd.isna(s):
        return ""
    return re.sub(r'\s+',' ', str(s).strip().lower())

def pick(col_list, candidates):
    cols = {c.lower():c for c in col_list}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    for cand in candidates:
        for c in col_list:
            if cand.lower() in c.lower():
                return c
    return None

def infer_turno_from_time(val):
    if pd.isna(val) or val=="":
        return None
    # pandas Timestamp/time
    if isinstance(val, (pd.Timestamp, datetime.time, datetime.datetime)):
        h = val.hour if hasattr(val, 'hour') else None
    else:
        s = str(val).strip()
        # tenta HH:MM ou HH:MM:SS
        m = re.match(r'(\d{1,2}):(\d{2})', s)
        if m:
            h = int(m.group(1))
        else:
            # tenta converter com pandas
            try:
                t = pd.to_datetime(s, errors='coerce')
                h = t.hour if not pd.isna(t) else None
            except:
                h = None
    if h is None:
        return None
    if h < 12:
        return "manhã"
    if h >= 18:
        return "noite"
    return "tarde"

def parse_date_safe(d):
    try:
        dt = pd.to_datetime(d, dayfirst=True, errors='coerce')
        return dt
    except:
        return pd.NaT

def main():
    # Lê planilhas (primeira aba de cada)
    prof = pd.read_excel(FILE_PROF, sheet_name=0, dtype=object)
    pasta = pd.read_excel(FILE_PASTA, sheet_name=0, dtype=object)
    turmas = pd.read_excel(FILE_TURMAS, sheet_name=0, dtype=object)

    prof_cols = prof.columns.tolist()
    tur_cols = turmas.columns.tolist()
    pasta_cols = pasta.columns.tolist()

    # Detecta colunas heurísticamente
    prof_disc = pick(prof_cols, ["disciplina","materia","matéria"])
    prof_prof = pick(prof_cols, ["professor","docente","responsável","responsavel","prof"])
    prof_n1 = pick(prof_cols, ["N1","dia","data","dia da aplicação","data n1","dia_n1"])

    tur_disc = pick(tur_cols, ["disciplina","materia","matéria"])
    tur_prof = pick(tur_cols, ["professor","docente","responsável","responsavel","prof"])
    tur_inicio = pick(tur_cols, ["inicio","hora_inicio","horario inicio","hora de inicio","início"])
    tur_fim = pick(tur_cols, ["fim","termino","horario fim","hora fim"])
    tur_pres = pick(tur_cols, ["presencial","modalidade","tipo"])
    tur_turno = pick(tur_cols, ["turno","periodo"])

    pasta_nome = pick(pasta_cols, ["aluna","nome","nome da aluna","alunas","nome_aluna"])
    pasta_disc = pick(pasta_cols, ["disciplina","materia","matéria"])
    pasta_prof = pick(pasta_cols, ["professor","docente","responsavel"])

    # Normalizações e exclusões
    if prof_disc:
        prof['__disc_norm'] = prof[prof_disc].apply(norm)
        prof = prof[~prof[prof_disc].apply(contains_exclude)]
    else:
        prof['__disc_norm'] = prof.index.astype(str)

    if tur_disc:
        turmas['__disc_norm'] = turmas[tur_disc].apply(norm)
    else:
        turamas['__disc_norm'] = turamas.index.astype(str)  # fallback (não esperado)

    if prof_prof:
        prof['__prof_norm'] = prof[prof_prof].apply(norm)
    else:
        prof['__prof_norm'] = ""

    if tur_prof:
        turamas = turamas.rename(columns={tur_prof:"__turma_prof"})
        turamas['__turma_prof_norm'] = turamas['__turma_prof'].apply(norm)
    else:
        turamas['__turma_prof_norm'] = ""

    tur = turamas.copy()
    tur['_start_raw'] = tur[tur_inicio] if tur_inicio in tur.columns else None
    tur['_end_raw'] = tur[tur_fim] if tur_fim in tur.columns else None
    # turno preferindo coluna, senão inferir por hora
    if tur_turno and tur_turno in tur.columns:
        tur['_turno'] = tur[tur_turno].apply(lambda x: norm(x) if not pd.isna(x) else "")
    else:
        tur['_turno'] = tur['_start_raw'].apply(infer_turno_from_time)

    def find_turmas_for(disc_name, prof_name=None):
        discn = norm(disc_name)
        profn = norm(prof_name) if prof_name else ""
        cand = tur[tur.get('__disc_norm', pd.Series()).fillna("").astype(str) == discn]
        if profn:
            if '__turma_prof_norm' in cand.columns:
                cand2 = cand[cand['__turma_prof_norm'].str.contains(profn, na=False)]
                if not cand2.empty:
                    return cand2
        return cand

    rows = []
    for idx, r in pasta.iterrows():
        nome = r[pasta_nome] if pasta_nome else f"aluna_{idx}"
        if isinstance(nome, float) and pd.isna(nome):
            nome = f"aluna_{idx}"
        if pasta_disc:
            disc = r[pasta_disc]
            prof_from_pasta = r[pasta_prof] if pasta_prof in pasta.columns else None
            if pd.isna(disc) or str(disc).strip()=="":
                # sem disciplina na linha
                rows.append({"aluna":nome, "disciplina":"", "professor": prof_from_pasta or "", "dia_N1":"N1 não cadastrado", "inicio":"","fim":"","presencial":"","turno_detectado":"","observacoes":"linha sem disciplina"})
                continue
            tur_found = find_turmas_for(disc, prof_from_pasta)
            if tur_found is None or tur_found.empty:
                # tentar procurar no prof sheet por disciplina
                prof_matches = prof[prof['__disc_norm'] == norm(disc)]
                if not prof_matches.empty:
                    for _, pm in prof_matches.iterrows():
                        professor_final = pm[prof_prof] if prof_prof in pm.columns else ""
                        n1 = pm[prof_n1] if prof_n1 in prof.columns else ""
                        n1v = n1 if not pd.isna(n1) and str(n1).strip()!="" else "N1 não cadastrado"
                        t2 = find_turmas_for(disc, professor_final)
                        if t2 is None or t2.empty:
                            rows.append({
                                "aluna": nome, "disciplina": disc, "professor": professor_final or "",
                                "dia_N1": n1v, "inicio": "", "fim": "", "presencial": "", "turno_detectado": "", "observacoes":"turma não encontrada"
                            })
                        else:
                            for _, tt in t2.iterrows():
                                rows.append({
                                    "aluna": nome, "disciplina": disc, "professor": professor_final or "",
                                    "dia_N1": n1v,
                                    "inicio": tt.get(tur_inicio,"") if tur_inicio in tt.index else "",
                                    "fim": tt.get(tur_fim,"") if tur_fim in tt.index else "",
                                    "presencial": tt.get(tur_pres,"") if tur_pres in tt.index else "",
                                    "turno_detectado": tt.get('_turno',""),
                                    "observacoes":"via_prof_sheet"
                                })
                else:
                    rows.append({
                        "aluna": nome, "disciplina": disc, "professor": prof_from_pasta or "",
                        "dia_N1": "N1 não cadastrado",
                        "inicio":"", "fim":"", "presencial":"", "turno_detectado":"", "observacoes":"nenhuma correspondência"
                    })
            else:
                for _, tt in tur_found.iterrows():
                    # localizar N1 pela combinação disciplina+professor em prof sheet (melhor esforço)
                    dia_n1 = ""
                    if prof_disc and prof_prof and prof_n1:
                        possible = prof[(prof['__disc_norm'] == norm(disc))]
                        if not possible.empty and '__prof_norm' in possible.columns:
                            # se professor conhecido na turma, tenta casar
                            profn_t = tt.get('__turma_prof',"")
                            pm = possible[possible['__prof_norm'].str.contains(norm(profn_t), na=False)]
                            if pm.empty:
                                pm = possible
                            dia_n1 = pm.iloc[0].get(prof_n1,"") if prof_n1 in pm.columns else ""
                    dia_n1v = dia_n1 if dia_n1 and not pd.isna(dia_n1) and str(dia_n1).strip()!="" else "N1 não cadastrado"
                    rows.append({
                        "aluna": nome, "disciplina": disc, "professor": tt.get('__turma_prof',"") or (prof_from_pasta or ""),
                        "dia_N1": dia_n1v,
                        "inicio": tt.get(tur_inicio,"") if tur_inicio in tt.index else "",
                        "fim": tt.get(tur_fim,"") if tur_fim in tt.index else "",
                        "presencial": tt.get(tur_pres,"") if tur_pres in tt.index else "",
                        "turno_detectado": tt.get('_turno',""),
                        "observacoes":"via_pasta_disc"
                    })
        else:
            # pasta não tem disciplina: procurar por professor/disciplinas associados
            prof_from_pasta = r[pasta_prof] if pasta_prof in pasta.columns else None
            if not prof_from_pasta or pd.isna(prof_from_pasta):
                rows.append({"aluna":nome,"disciplina":"","professor":"","dia_N1":"N1 não cadastrado","inicio":"","fim":"","presencial":"","turno_detectado":"","observacoes":"sem dados"})
                continue
            prof_matches = prof[prof['__prof_norm'].str.contains(norm(prof_from_pasta), na=False)]
            if prof_matches.empty:
                t_matches = tur[tur.get('__turma_prof_norm', pd.Series()).str.contains(norm(prof_from_pasta), na=False)]
                if t_matches.empty:
                    rows.append({"aluna":nome,"disciplina":"","professor":prof_from_pasta,"dia_N1":"N1 não cadastrado","inicio":"","fim":"","presencial":"","turno_detectado":"","observacoes":"professor não encontrado"})
                else:
                    for _, tt in t_matches.iterrows():
                        rows.append({
                            "aluna":nome,"disciplina":tt.get(tur_disc,""),"professor":tt.get('__turma_prof',''),
                            "dia_N1":"N1 não cadastrado", "inicio":tt.get(tur_inicio,""), "fim":tt.get(tur_fim,""),
                            "presencial":tt.get(tur_pres,""), "turno_detectado":tt.get('_turno',""), "observacoes":"via_turmas"
                        })
            else:
                for _, pm in prof_matches.iterrows():
                    disc = pm.get(prof_disc,"")
                    dia_n1 = pm.get(prof_n1,"") if prof_n1 in pm.columns else ""
                    dia_n1v = dia_n1 if dia_n1 and not pd.isna(dia_n1) and str(dia_n1).strip()!="" else "N1 não cadastrado"
                    t_matches = find_turmas_for(disc, prof_from_pasta)
                    if t_matches is None or t_matches.empty:
                        rows.append({
                            "aluna":nome,"disciplina":disc,"professor":prof_from_pasta,"dia_N1":dia_n1v,"inicio":"","fim":"","presencial":"","turno_detectado":"","observacoes":"sem turma"
                        })
                    else:
                        for _, tt in t_matches.iterrows():
                            rows.append({
                                "aluna":nome,"disciplina":disc,"professor":prof_from_pasta,"dia_N1":dia_n1v,
                                "inicio":tt.get(tur_inicio,""), "fim":tt.get(tur_fim,""), "presencial":tt.get(tur_pres,""),
                                "turno_detectado":tt.get('_turno',""), "observacoes":"via_prof_match"
                            })

    out = pd.DataFrame(rows, columns=["aluna","disciplina","professor","dia_N1","inicio","fim","presencial","turno_detectado","observacoes"])

    # normalizar turnos
    out['turno_detectado'] = out['turno_detectado'].astype(str).fillna("").replace({"nan":""}).str.strip().str.lower()
    out['turno_detectado'] = out['turno_detectado'].replace({"manha":"manhã"})

    # Inferir turno por aluna (maioria manhã/noite). Em caso de empate -> "Indeterminado (empate)" (regra A).
    resumo = []
    for nome, g in out.groupby('aluna'):
        turnos = [t for t in g['turno_detectado'] if t and t in ['manhã','noite','tarde']]
        # considerar apenas manhã e noite para decisão
        turnos_norm = ['manhã' if t.startswith('man') else t for t in turnos]
        c = Counter(turnos_norm)
        man = c.get('manhã',0)
        noite = c.get('noite',0)
        if man>noite:
            turno_inferido = "manhã"
        elif noite>man:
            turno_inferido = "noite"
        elif man==noite and man==0:
            turno_inferido = "indeterminado"
        else:
            # empate
            turno_inferido = "Indeterminado (empate)"

        # Ordena detalhes por dia_N1 (datas) e por inicio (horário)
        g2 = g.copy()
        g2['__dia_parsed'] = g2['dia_N1'].apply(lambda d: parse_date_safe(d) if d and d!="N1 não cadastrado" else pd.NaT)
        def parse_time_safe(t):
            if pd.isna(t) or t=="":
                return pd.NaT
            try:
                tt = pd.to_datetime(t, errors='coerce')
                return tt.time() if not pd.isna(tt) else pd.NaT
            except:
                return pd.NaT
        g2['__time_parsed'] = g2['inicio'].apply(parse_time_safe)
        g2 = g2.sort_values(by=['__dia_parsed','__time_parsed'], ascending=[True, True], na_position='last')
        detalhes = g2[["disciplina","professor","dia_N1","inicio","fim","presencial","turno_detectado","observacoes"]].to_dict(orient='records')
        resumo.append({"aluna":nome, "turno_inferido":turno_inferido, "detalhes":detalhes})

    # salvar resultados
    out.to_csv("resultado_cruzamento.csv", index=False, encoding='utf-8-sig')
    out.to_excel("resultado_cruzamento.xlsx", index=False)

    resumo_df = pd.DataFrame([{"aluna":r["aluna"], "turno_inferido":r["turno_inferido"], "detalhes":str(r["detalhes"])} for r in resumo])
    resumo_df.to_excel("resumo_por_aluna.xlsx", index=False)

    print("Arquivos gerados:\n - resultado_cruzamento.xlsx\n - resultado_cruzamento.csv\n - resumo_por_aluna.xlsx")
    print("Se quiser, envie aqui os arquivos gerados e eu verifico/ajusto e preparo o relatório final.")

if __name__ == "__main__":
    main()

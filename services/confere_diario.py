"""Leitura e conferência dos relatórios de notas do Diário Online do Padin."""

import hashlib
import io
import re
from collections import defaultdict

import pdfplumber

from services.conferir_listas import (
    SED_TRANSFER_STATUSES,
    clean_display,
    format_turma_key,
    normalize_name_for_match,
    normalize_text,
    normalize_turma_lista,
)


STATUS_LABELS = {
    "ATIVO": "Ativo",
    "TE": "Transferido",
    "REM": "Remanejado",
    "NCOM": "Não compareceu",
    "AUSENTE": "Não localizado",
    "DESCONHECIDO": "Não reconhecido",
    "AMBIGUO": "Revisar identidade",
}


def _movement(value):
    raw = clean_display(value)
    normalized = normalize_text(raw)
    if not normalized:
        return "ATIVO", ""
    if normalized.startswith("T E") or normalized.startswith("TE ") or normalized == "TE":
        return "TE", ""
    if normalized.startswith("NCOM") or normalized.startswith("N COM"):
        return "NCOM", ""
    match = re.match(r"REM\.?\s*(DE|P\s*/|PARA)\s+(.+?)(?:\s+-\s+\d{1,2}/\d{1,2})?$", raw, re.I)
    if match:
        other_class = normalize_turma_lista(match.group(2))
        if other_class:
            return ("REM_DE" if match.group(1).upper() == "DE" else "REM_P"), other_class
    return "DESCONHECIDO", ""


def extract_diario_pdf_records(pdf_bytes, filename):
    records = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            header = (page.extract_text() or "").splitlines()[:5]
            header_text = " ".join(header)
            class_match = re.search(
                r"Turma:\s*(\d{1,2}\s*[º°]?\s*Ano\s+[A-Z](?:\([IP]\))?)",
                header_text,
                re.I,
            )
            # Alguns relatórios terminam com uma página de assinaturas, sem turma nem alunos.
            if not class_match:
                continue
            if "PADIN MOUTA" not in normalize_text(header_text):
                raise ValueError("O relatório não foi identificado como Diário Online do José Padin Mouta.")
            turma = clean_display(class_match.group(1)) if class_match else ""
            turma_key = normalize_turma_lista(turma)
            if not turma_key:
                raise ValueError("Não foi possível identificar a turma no cabeçalho do Diário Online.")

            found_table = False
            for table in page.extract_tables():
                for index, row in enumerate(table):
                    if len(row) < 3 or normalize_text(row[1]) != "NOME DO ALUNO":
                        continue
                    if "MOVIMENTACAO" not in normalize_text(row[2]):
                        continue
                    found_table = True
                    for student in table[index + 1:]:
                        if len(student) < 3 or not re.fullmatch(r"\d+", clean_display(student[0])):
                            continue
                        nome = clean_display(student[1])
                        if not nome:
                            continue
                        situacao, outra_turma_key = _movement(student[2])
                        records.append({
                            "source": "diario",
                            "pdf_origem": filename,
                            "turma": turma,
                            "turma_key": turma_key,
                            "numero_chamada": str(int(clean_display(student[0]))),
                            "nome": nome,
                            "nome_match_norm": normalize_name_for_match(nome),
                            "situacao": situacao,
                            "movimentacao": clean_display(student[2]),
                            "outra_turma_key": outra_turma_key,
                        })
            if not found_table:
                raise ValueError("A tabela de alunos não foi localizada em uma página do relatório.")
            if hasattr(page, "flush_cache"):
                page.flush_cache()
    if not records:
        raise ValueError("Nenhum aluno foi extraído do relatório do Diário Online.")
    return records


def extract_diario_pdfs(file_items):
    info = {"records": [], "errors": [], "successful_files": [], "duplicate_files": []}
    hashes = set()
    class_files = {}
    for item in file_items:
        filename = item["filename"]
        content = item["content"]
        digest = hashlib.sha256(content).hexdigest()
        if digest in hashes:
            info["duplicate_files"].append(filename)
            continue
        hashes.add(digest)
        try:
            records = extract_diario_pdf_records(content, filename)
        except Exception as exc:
            info["errors"].append({"arquivo": filename, "erro": str(exc)})
            continue
        repeated = {record["turma_key"] for record in records if record["turma_key"] in class_files}
        if repeated:
            repeated_labels = ", ".join(format_turma_key(key) for key in sorted(repeated))
            raise ValueError(
                f"Há mais de um PDF do Diário Online para a turma {repeated_labels}. "
                "Envie somente o relatório mais atual de cada turma."
            )
        for key in {record["turma_key"] for record in records}:
            class_files[key] = filename
        info["records"].extend(records)
        info["successful_files"].append(filename)
    return info


def _effective_status(records, source):
    if source == "lista":
        active = [record for record in records if record.get("situacao") in {"MA", "TR", "MN", "NF"}]
    elif source == "sed":
        active = [record for record in records if record.get("situacao") == "ATIVO"]
    else:
        active = [record for record in records if record.get("situacao") in {"ATIVO", "REM_DE"}]
    if len(active) == 1:
        return "ATIVO", active[0], False
    if len(active) > 1:
        return "AMBIGUO", None, True
    statuses = {record.get("situacao") for record in records}
    if source == "sed":
        if statuses & SED_TRANSFER_STATUSES:
            return "TE", None, False
        if "REMA" in statuses:
            return "REM", None, False
    elif source == "diario":
        if "TE" in statuses:
            return "TE", None, False
        if "REM_P" in statuses:
            return "REM", None, False
    else:
        if "TE" in statuses:
            return "TE", None, False
        if "REM" in statuses:
            return "REM", None, False
    if "NCOM" in statuses:
        return "NCOM", None, False
    return "DESCONHECIDO", None, False


def _display_source(records, source):
    status, current, ambiguous = _effective_status(records, source)
    status_records = {
        "lista": {"TE": {"TE"}, "REM": {"REM"}, "NCOM": {"NCOM"}},
        "sed": {"TE": SED_TRANSFER_STATUSES, "REM": {"REMA"}, "NCOM": {"NCOM"}},
        "diario": {"TE": {"TE"}, "REM": {"REM_P"}, "NCOM": {"NCOM"}},
    }
    selected = current or next(
        (record for record in records if record.get("situacao") in status_records[source].get(status, set())),
        records[0] if records else None,
    )
    return {
        "status": status if records else "AUSENTE",
        "nome": clean_display(selected.get("nome")) if selected else "",
        "turma": format_turma_key(current.get("turma_key")) if current else (format_turma_key(selected.get("turma_key")) if selected else ""),
        "turma_key": current.get("turma_key", "") if current else "",
        "numero": clean_display(current.get("numero_chamada")) if current else "",
        "ra": clean_display(selected.get("ra")) if selected else "",
        "nascimento": clean_display(selected.get("data_nascimento")) if selected else "",
        "pdf": selected.get("pdf_origem", "") if selected else "",
        "ambiguous": ambiguous,
    }


def _one_character_edit(first, second):
    if abs(len(first) - len(second)) > 1:
        return False
    if len(first) == len(second):
        return sum(left != right for left, right in zip(first, second)) == 1
    shorter, longer = sorted((first, second), key=len)
    return any(shorter == longer[:index] + longer[index + 1:] for index in range(len(longer)))


def _names_differ_by_one_letter(first, second):
    first_tokens = normalize_text(first).split()
    second_tokens = normalize_text(second).split()
    if len(first_tokens) < 3 or len(first_tokens) != len(second_tokens):
        return False
    differences = [(left, right) for left, right in zip(first_tokens, second_tokens) if left != right]
    return (
        len(differences) == 1
        and min(map(len, differences[0])) >= 4
        and _one_character_edit(*differences[0])
    )


def _unique_active_record(records, source):
    status, current, ambiguous = _effective_status(records, source)
    return current if status == "ATIVO" and not ambiguous else None


def _reconcile_diario_name_typos(groups, scope):
    """Une apenas grafias quase iguais com turma, chamada e RA concordantes."""
    anchors = {}
    diario_only = {}
    for key, group in groups.items():
        if group["lista"] and group["sed"] and not group["diario"]:
            lista = _unique_active_record(group["lista"], "lista")
            sed = _unique_active_record(group["sed"], "sed")
            if not lista or not sed or lista.get("turma_key") != sed.get("turma_key"):
                continue
            lista_ra = set(lista.get("ra_keys") or ())
            sed_ra = set(sed.get("ra_keys") or ())
            if (lista.get("turma_key") not in scope or not lista.get("numero_chamada")
                    or not lista_ra or not sed_ra or lista_ra.isdisjoint(sed_ra)):
                continue
            anchors[key] = lista
        elif group["diario"] and not group["lista"] and not group["sed"]:
            diario = _unique_active_record(group["diario"], "diario")
            if diario and diario.get("turma_key") in scope and diario.get("numero_chamada"):
                diario_only[key] = diario

    candidates = []
    for diario_key, diario in diario_only.items():
        for anchor_key, lista in anchors.items():
            if (diario["turma_key"] == lista["turma_key"]
                    and diario["numero_chamada"] == lista["numero_chamada"]
                    and _names_differ_by_one_letter(lista["nome"], diario["nome"])):
                candidates.append((diario_key, anchor_key))

    diario_counts = {key: sum(diario_key == key for diario_key, _ in candidates) for key in diario_only}
    anchor_counts = {key: sum(anchor_key == key for _, anchor_key in candidates) for key in anchors}
    for diario_key, anchor_key in candidates:
        if diario_counts[diario_key] == 1 and anchor_counts[anchor_key] == 1:
            groups[anchor_key]["diario"].extend(groups[diario_key]["diario"])
            del groups[diario_key]


def compare_three_sources(lista_records, sed_records, diario_info):
    """Compara por nome; pequenas grafias só são vinculadas com evidências extras."""
    diario_records = diario_info["records"]
    sed_classes = {r.get("turma_key") for r in sed_records if r.get("turma_key")}
    diario_classes = {r.get("turma_key") for r in diario_records if r.get("turma_key")}
    scope = sed_classes & diario_classes
    if not scope:
        raise ValueError("SED e Diário Online não têm turmas em comum nos PDFs enviados.")
    groups = defaultdict(lambda: {"lista": [], "sed": [], "diario": []})
    for source, records in (("lista", lista_records), ("sed", sed_records), ("diario", diario_records)):
        for record in records:
            key = record.get("nome_match_norm") or normalize_name_for_match(record.get("nome"))
            if key:
                groups[key][source].append(record)
    _reconcile_diario_name_typos(groups, scope)

    rows = []
    for name_groups in groups.values():
        if not any(record.get("turma_key") in scope for records in name_groups.values() for record in records):
            continue
        views = {source: _display_source(records, source) for source, records in name_groups.items()}
        nome = next((r["nome"] for source in ("lista", "sed", "diario") for r in name_groups[source]), "")
        issues = []
        issue_types = []

        def add_issue(kind, message):
            issues.append(message)
            issue_types.append(kind)

        if (views["lista"]["nome"] and views["sed"]["nome"] and views["diario"]["nome"]
                and normalize_name_for_match(views["diario"]["nome"])
                != normalize_name_for_match(views["lista"]["nome"])):
            add_issue("nome", "Nome cadastrado de forma diferente no Diário Online; confira a grafia.")

        ra_sets = [
            {tuple(r.get("ra_keys") or ()) for r in name_groups["lista"] if r.get("ra_keys")},
            {tuple(r.get("ra_keys") or ()) for r in name_groups["sed"] if r.get("ra_keys")},
        ]
        repeated_identity = any(len(values) > 1 for values in ra_sets) or any(view["ambiguous"] for view in views.values())
        lista_ra = {key for values in ra_sets[0] for key in values}
        sed_ra = {key for values in ra_sets[1] for key in values}
        divergent_ra = bool(lista_ra and sed_ra and not lista_ra.intersection(sed_ra))
        if repeated_identity or divergent_ra:
            if repeated_identity:
                add_issue("identidade", "Nome repetido ou mais de uma turma ativa: conferir a identidade manualmente.")
            if divergent_ra:
                if views["lista"]["ra"] in {"", "0"}:
                    add_issue("identidade", "RA não informado na Lista Piloto; confira o cadastro antes de vincular os registros.")
                else:
                    add_issue("identidade", "RA divergente entre Lista Piloto e SED; confirme a identidade antes de vincular o Diário.")
        else:
            present = {source: view for source, view in views.items() if view["status"] != "AUSENTE"}
            if any(view["status"] == "ATIVO" for view in present.values()):
                for source, label in (("lista", "na Lista Piloto"), ("sed", "no SED"), ("diario", "no Diário Online")):
                    if views[source]["status"] == "AUSENTE":
                        add_issue("matricula", f"Aluno ativo não localizado {label}.")
            statuses = {view["status"] for view in present.values()}
            if len(statuses) > 1 and "DESCONHECIDO" not in statuses:
                add_issue("situacao", "Situação diferente entre as fontes (ativo, transferido, remanejado ou não compareceu).")
            if "DESCONHECIDO" in statuses:
                add_issue("situacao", "Situação não reconhecida em uma das fontes: conferir manualmente.")
            active_classes = {view["turma_key"] for view in present.values() if view["status"] == "ATIVO"}
            if len(active_classes) > 1:
                add_issue("turma", "Aluno ativo em turmas diferentes entre as fontes.")
            if (views["lista"]["status"] == "ATIVO" and views["diario"]["status"] == "ATIVO"
                    and views["lista"]["turma_key"] == views["diario"]["turma_key"]
                    and views["lista"]["numero"] and views["diario"]["numero"]
                    and views["lista"]["numero"] != views["diario"]["numero"]):
                add_issue("chamada", "Número da chamada diferente entre Lista Piloto e Diário Online.")

            for record in name_groups["diario"]:
                if record["situacao"] not in {"REM_DE", "REM_P"}:
                    continue
                origin = record["outra_turma_key"] if record["situacao"] == "REM_DE" else record["turma_key"]
                if (record["situacao"] == "REM_P" and record["outra_turma_key"] in scope
                        and not any(r["turma_key"] == record["outra_turma_key"] and r["situacao"] in {"ATIVO", "REM_DE"}
                                    for r in name_groups["diario"])):
                    add_issue("remanejamento", "Remanejado para outra turma, mas não localizado no relatório de destino do Diário.")
                if origin not in scope:
                    continue
                if not any(r.get("turma_key") == origin and r.get("situacao") == "REMA" for r in name_groups["sed"]):
                    add_issue("remanejamento", "Remanejamento do Diário sem registro correspondente na turma de origem do SED.")
                    break
            for record in name_groups["sed"]:
                if record.get("situacao") != "REMA" or record.get("turma_key") not in scope:
                    continue
                if not any(r["situacao"] in {"REM_DE", "REM_P"} and
                           (r["outra_turma_key"] if r["situacao"] == "REM_DE" else r["turma_key"]) == record["turma_key"]
                           for r in name_groups["diario"]):
                    add_issue("remanejamento", "Remanejamento do SED sem registro correspondente no Diário Online.")
                    break

        movement = "; ".join(sorted({r["movimentacao"] for r in name_groups["diario"] if r.get("movimentacao")}))
        unique_types = list(dict.fromkeys(issue_types))
        priority = next(
            (kind for kind in ("identidade", "nome", "matricula", "situacao", "turma", "remanejamento", "chamada")
             if kind in unique_types),
            "ok",
        )
        turma_key = next(
            (views[source]["turma_key"] for source in ("diario", "lista", "sed") if views[source]["turma_key"]),
            normalize_turma_lista(views["diario"]["turma"] or views["lista"]["turma"] or views["sed"]["turma"]),
        )
        rows.append({
            "nome": nome,
            "nome_lista": views["lista"]["nome"],
            "nome_sed": views["sed"]["nome"],
            "nome_diario": views["diario"]["nome"],
            "turma_key": turma_key,
            "turma_lista": views["lista"]["turma"],
            "turma_sed": views["sed"]["turma"],
            "turma_diario": views["diario"]["turma"],
            "situacao_lista": views["lista"]["status"],
            "situacao_sed": views["sed"]["status"],
            "situacao_diario": views["diario"]["status"],
            "situacao_lista_display": STATUS_LABELS.get(views["lista"]["status"], views["lista"]["status"]),
            "situacao_sed_display": STATUS_LABELS.get(views["sed"]["status"], views["sed"]["status"]),
            "situacao_diario_display": STATUS_LABELS.get(views["diario"]["status"], views["diario"]["status"]),
            "numero_lista": views["lista"]["numero"],
            "numero_diario": views["diario"]["numero"],
            "ra_lista": views["lista"]["ra"],
            "ra_sed": views["sed"]["ra"],
            "nascimento_lista": views["lista"]["nascimento"],
            "nascimento_sed": views["sed"]["nascimento"],
            "movimentacao_diario": movement,
            "pdf_diario": views["diario"]["pdf"],
            "pendencias": list(dict.fromkeys(issues)),
            "pendencias_tipos": unique_types,
            "tipo_principal": priority,
            "categoria": "pendente" if issues else "ok",
        })
    rows.sort(key=lambda r: (r["categoria"] != "pendente", normalize_turma_lista(r["turma_lista"] or r["turma_sed"] or r["turma_diario"]), normalize_text(r["nome"])))
    type_totals = {
        kind: sum(kind in row["pendencias_tipos"] for row in rows)
        for kind in ("identidade", "nome", "matricula", "situacao", "turma", "remanejamento", "chamada")
    }
    return {
        "summary": {
            "total_alunos": len(rows),
            "total_pendencias": sum(bool(row["pendencias"]) for row in rows),
            "total_ok": sum(not row["pendencias"] for row in rows),
            "total_pdfs_lidos": len(diario_info["successful_files"]),
            "total_turmas_conferidas": len(scope),
            "total_somente_chamada": sum(row["pendencias_tipos"] == ["chamada"] for row in rows),
            "total_outros_alertas": sum(bool(set(row["pendencias_tipos"]) - {"chamada"}) for row in rows),
            "totais_tipos": type_totals,
        },
        "turmas_conferidas": [format_turma_key(key) for key in sorted(scope)],
        "turmas_diario_sem_sed": [format_turma_key(key) for key in sorted(diario_classes - sed_classes)],
        "turmas_sed_sem_diario": [format_turma_key(key) for key in sorted(sed_classes - diario_classes)],
        "rows": rows,
        "pdf_errors": diario_info["errors"],
        "duplicate_files": diario_info["duplicate_files"],
    }

from dataclasses import dataclass, field, replace


DEFAULT_LISTA_COLUMNS = {
    "turma": ["SERIE", "SÉRIE", "TURMA"],
    "nome": ["NOME", "NOME DO ALUNO", "ALUNO"],
    "data_nascimento": ["DATA NASC", "DATA NASC.", "DATA DE NASCIMENTO", "NASCIMENTO"],
    "ra": ["RA", "R.A."],
    "situacao": ["COD", "CÓD", "SITUACAO", "SITUAÇÃO", "STATUS"],
    "observacoes": ["OBS", "OBSERVACAO", "OBSERVAÇÃO"],
}


DEFAULT_STATUS_MAP = {
    "MA": "MA",
    "MATRICULA ATIVA": "MA",
    "MATRÍCULA ATIVA": "MA",
    "ATIVO": "MA",
    "TE": "TE",
    "TRANSFERENCIA EXPEDIDA": "TE",
    "TRANSFERÊNCIA EXPEDIDA": "TE",
    "TRANSFERIDO": "TE",
    "REM": "REM",
    "REMANEJADO": "REM",
}


@dataclass(frozen=True)
class ConfereSchoolConfig:
    id: str
    nome: str
    sheet_name: str | int | None = "LISTA CORRIDA"
    column_mode: str = "headers"
    header_row: int = 1
    data_start_row: int | None = None
    columns: dict = field(default_factory=lambda: dict(DEFAULT_LISTA_COLUMNS))
    status_map: dict = field(default_factory=lambda: dict(DEFAULT_STATUS_MAP))
    description: str = ""
    ra_digits_only: bool = False
    strip_ra_uf_suffix: bool = False
    allow_class_sheets: bool = False

    def to_template_dict(self):
        return {
            "id": self.id,
            "nome": self.nome,
            "description": self.description,
        }


CONFERE_SCHOOL_CONFIGS = {
    "padin": ConfereSchoolConfig(
        id="padin",
        nome="E.M José Padin Mouta",
        sheet_name="LISTA CORRIDA",
        column_mode="headers",
        header_row=1,
        columns=DEFAULT_LISTA_COLUMNS,
        description="Modelo atual da Lista Piloto com aba LISTA CORRIDA.",
    ),
    "mahatma_gandhi": ConfereSchoolConfig(
        id="mahatma_gandhi",
        nome="E.M Mahatma Gandhi",
        sheet_name="Verifica\u00e7\u00e3o SED",
        column_mode="headers",
        header_row=1,
        data_start_row=2,
        columns={
            "turma": ["TURMA"],
            "nome": ["NOME DO ALUNO", "NOME", "ALUNO"],
            "data_nascimento": ["NASC.", "NASC", "DATA NASC", "DATA NASC.", "DATA DE NASCIMENTO"],
            "ra": ["R.A.", "RA", "R A"],
            "situacao": ["C\u00d3D.", "COD.", "C\u00d3D", "COD", "SITUACAO", "SITUA\u00c7\u00c3O"],
            "observacoes": ["OBSERVA\u00c7\u00c3O", "OBSERVACAO", "OBS"],
        },
        status_map={
            **DEFAULT_STATUS_MAP,
            "": "MA",
            "0": "MA",
            "PNEE": "MA",
            "TR": "TE",
            "T R": "TE",
            "TE": "TE",
            "T E": "TE",
            "MN": "MA",
            "M N": "MA",
            "MATRICULA NORMAL": "MA",
            "REM": "REM",
            "NF": "NF",
            "N F": "NF",
            "NAO FREQUENTE": "NF",
            "NCOM": "NCOM",
            "N COM": "NCOM",
            "NAO COMPARECEU": "NCOM",
        },
        description="Aba Verifica\u00e7\u00e3o SED ou abas vis\u00edveis por turma da Lista Piloto.",
        ra_digits_only=True,
        strip_ra_uf_suffix=True,
        allow_class_sheets=True,
    ),
}

CONFERE_SCHOOL_CONFIGS["maria_nilza"] = replace(
    CONFERE_SCHOOL_CONFIGS["mahatma_gandhi"],
    id="maria_nilza",
    nome="E.M. Maria Nilza da Silva Rom\u00e3o, Prof.\u00aa",
    description="Mesmo modelo de leitura do Mahatma Gandhi, incluindo abas vis\u00edveis por turma.",
)


def default_confere_school_id():
    return "padin"


def default_confere_school_config():
    return CONFERE_SCHOOL_CONFIGS[default_confere_school_id()]


def get_confere_school_config(school_id):
    if not school_id:
        return None
    return CONFERE_SCHOOL_CONFIGS.get(str(school_id).strip())


def list_confere_schools():
    return [config.to_template_dict() for config in CONFERE_SCHOOL_CONFIGS.values()]

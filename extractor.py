"""
Extrai metadados de arquivos .pbix sem precisar abrir o Power BI Desktop.
Um .pbix é um ZIP contendo DataModel (formato ABF) e outros recursos.
"""

import zipfile
import json
import re
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Column:
    name: str
    data_type: str = "Text"
    description: str = ""
    is_hidden: bool = False
    is_calculated: bool = False
    expression: str = ""


@dataclass
class Measure:
    name: str
    expression: str = ""
    description: str = ""
    table: str = ""
    format_string: str = ""


@dataclass
class Table:
    name: str
    description: str = ""
    columns: list = field(default_factory=list)
    measures: list = field(default_factory=list)
    is_hidden: bool = False
    source_type: str = ""


@dataclass
class Relationship:
    from_table: str
    from_column: str
    to_table: str
    to_column: str
    cardinality: str = "ManyToOne"
    cross_filter: str = "Single"
    is_active: bool = True


@dataclass
class DataSource:
    name: str
    kind: str = "Unknown"
    connection_string: str = ""


@dataclass
class PBIXModel:
    file_name: str
    tables: list = field(default_factory=list)
    relationships: list = field(default_factory=list)
    data_sources: list = field(default_factory=list)
    report_pages: list = field(default_factory=list)
    raw_metadata: dict = field(default_factory=dict)
    extraction_warnings: list = field(default_factory=list)

    @property
    def all_measures(self):
        measures = []
        for t in self.tables:
            for m in t.measures:
                m.table = t.name
                measures.append(m)
        return measures

    @property
    def summary(self):
        return {
            "tables": len(self.tables),
            "columns": sum(len(t.columns) for t in self.tables),
            "measures": sum(len(t.measures) for t in self.tables),
            "relationships": len(self.relationships),
            "data_sources": len(self.data_sources),
            "report_pages": len(self.report_pages),
        }


def _infer_source_type(table_raw: dict) -> str:
    partitions = table_raw.get("partitions", [])
    for p in partitions:
        source = p.get("source", {})
        kind = source.get("type", source.get("kind", ""))
        if kind:
            return kind
        expr = source.get("expression", "")
        if isinstance(expr, list):
            expr = "\n".join(expr)
        if "Value.NativeQuery" in expr or "Sql.Database" in expr:
            return "SQL"
        if "SharePoint" in expr:
            return "SharePoint"
        if "Web.Contents" in expr or "Web.Page" in expr:
            return "Web"
        if "Excel.Workbook" in expr or "Csv.Document" in expr:
            return "File"
        if expr:
            return "Power Query (M)"
    return "Import"


def _parse_cardinality(card: str) -> str:
    mapping = {
        "manytoone": "N:1",
        "onetomany": "1:N",
        "onetoone": "1:1",
        "manytomany": "N:N",
    }
    return mapping.get(card.lower(), card)


def _extract_from_bim(bim_text: str, model: PBIXModel):
    """Parse JSON do arquivo model.bim / DataModel JSON."""
    try:
        data = json.loads(bim_text)
    except Exception:
        model.extraction_warnings.append("Não foi possível parsear o BIM JSON.")
        return

    model.raw_metadata = data

    db = data.get("model", data.get("database", data))
    tables_raw = db.get("tables", [])

    for t_raw in tables_raw:
        name = t_raw.get("name", "")
        if not name:
            continue

        table = Table(
            name=name,
            description=t_raw.get("description", ""),
            is_hidden=t_raw.get("isHidden", False),
            source_type=_infer_source_type(t_raw),
        )

        for c_raw in t_raw.get("columns", []):
            col_name = c_raw.get("name", "")
            if not col_name:
                continue
            col = Column(
                name=col_name,
                data_type=c_raw.get("dataType", "Text"),
                description=c_raw.get("description", ""),
                is_hidden=c_raw.get("isHidden", False),
                is_calculated=c_raw.get("type", "") == "calculated",
                expression=c_raw.get("expression", ""),
            )
            if isinstance(col.expression, list):
                col.expression = "\n".join(col.expression)
            table.columns.append(col)

        for m_raw in t_raw.get("measures", []):
            m_name = m_raw.get("name", "")
            if not m_name:
                continue
            expr = m_raw.get("expression", "")
            if isinstance(expr, list):
                expr = "\n".join(expr)
            meas = Measure(
                name=m_name,
                expression=expr,
                description=m_raw.get("description", ""),
                format_string=m_raw.get("formatString", ""),
                table=name,
            )
            table.measures.append(meas)

        model.tables.append(table)

    # Relacionamentos
    for r_raw in db.get("relationships", []):
        rel = Relationship(
            from_table=r_raw.get("fromTable", ""),
            from_column=r_raw.get("fromColumn", ""),
            to_table=r_raw.get("toTable", ""),
            to_column=r_raw.get("toColumn", ""),
            cardinality=_parse_cardinality(r_raw.get("cardinality", "ManyToOne")),
            cross_filter=r_raw.get("crossFilteringBehavior", "Single"),
            is_active=r_raw.get("isActive", True),
        )
        if rel.from_table and rel.to_table:
            model.relationships.append(rel)

    # Data sources / expressões
    for ds_raw in db.get("dataSources", []):
        ds = DataSource(
            name=ds_raw.get("name", ds_raw.get("connectionDetails", {}).get("protocol", "Source")),
            kind=ds_raw.get("type", ds_raw.get("kind", "Unknown")),
        )
        model.data_sources.append(ds)

    # Se não achou data sources, infere pelas tabelas
    if not model.data_sources:
        seen = set()
        for t in model.tables:
            if t.source_type and t.source_type not in seen:
                seen.add(t.source_type)
                model.data_sources.append(DataSource(name=t.source_type, kind=t.source_type))


def _extract_report_pages(layout_bytes: bytes, model: PBIXModel):
    """Extrai páginas do report layout."""
    try:
        text = layout_bytes.decode("utf-8-sig", errors="replace")
        data = json.loads(text)
        sections = data.get("sections", [])
        for s in sections:
            name = s.get("displayName", s.get("name", ""))
            if name:
                model.report_pages.append(name)
    except Exception:
        pass


def extract_pbix(file_path: str | Path) -> PBIXModel:
    """Ponto de entrada principal: extrai modelo de um .pbix."""
    path = Path(file_path)
    model = PBIXModel(file_name=path.name)

    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    if not zipfile.is_zipfile(path):
        raise ValueError(f"Arquivo não é um ZIP válido: {path.name}")

    with zipfile.ZipFile(path, "r") as z:
        names = z.namelist()

        # --- Tenta DataModelSchema (export do Power BI Service) ---
        bim_candidates = [
            n for n in names
            if n.endswith("model.bim") or n.endswith("DataModelSchema")
               or ("DataModel" in n and n.endswith(".json"))
        ]

        # --- Tenta Report/Layout ---
        layout_candidates = [n for n in names if n.endswith("Layout")]

        if bim_candidates:
            with z.open(bim_candidates[0]) as f:
                content = f.read().decode("utf-8-sig", errors="replace")
                _extract_from_bim(content, model)
        else:
            # Tenta extrair diretamente do DataModel (binário ABF via JSON interno)
            data_model_files = [n for n in names if "DataModel" in n]
            found = False
            for dm in data_model_files:
                try:
                    with z.open(dm) as f:
                        raw = f.read()
                    # Tenta achar JSON embutido
                    text = raw.decode("utf-8", errors="replace")
                    json_matches = re.findall(r'\{["\']model["\']:\s*\{', text)
                    if json_matches:
                        start = text.find('{"model":')
                        if start == -1:
                            start = text.find("{'model':")
                        if start != -1:
                            _extract_from_bim(text[start:], model)
                            found = True
                            break
                except Exception:
                    continue
            if not found:
                model.extraction_warnings.append(
                    "DataModel binário não pôde ser lido diretamente. "
                    "Para extração completa, exporte o .pbix como .pbit ou salve como 'Power BI Template'."
                )

        if layout_candidates:
            with z.open(layout_candidates[0]) as f:
                _extract_report_pages(f.read(), model)

    return model

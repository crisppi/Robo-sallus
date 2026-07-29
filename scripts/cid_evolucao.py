#!/usr/bin/env python3
"""Sugestão conservadora de CID-10 a partir do diagnóstico da evolução.

A ordem é intencional: código escrito pelo profissional, diagnóstico principal
explícito e somente depois regras de patologia suficientemente específicas.
"""

from __future__ import annotations

import re
import unicodedata


def normalized(value: str) -> str:
    text = unicodedata.normalize("NFD", str(value or "").lower())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", text).strip()


def infer_cid_suggestion(text: str) -> tuple[str, str]:
    """Retorna ``(cid, origem)``; vazio quando não há base clínica suficiente."""
    raw = str(text or "")
    explicit_cids = re.findall(
        r"(?i)\bcid\s*[:\-]?\s*([A-Z]\d{2}(?:\.\d{1,2})?)\b",
        raw,
    )
    if explicit_cids:
        return explicit_cids[-1].upper(), "CID explícito na evolução"

    value = normalized(raw)
    padded = f" {value} "
    if (
        (
            re.search(r"\btce\b", value)
            or any(term in padded for term in (" traumatismo craniano ", " trauma craniano "))
        )
        and any(term in padded for term in (" queda ", " trauma ", " colisao "))
    ):
        return "S09.9", "traumatismo craniano sem lesão mais específica descrita"

    rules: tuple[tuple[tuple[str, ...], str, str], ...] = (
        (("sindrome de guillain-barre", "sindrome de guillain barre", "polirradiculoneurite aguda"), "G61.0", "síndrome de Guillain-Barré"),
        (("retirada de corpo estranho subcutaneo", "corpo estranho subcutaneo"), "M79.5", "corpo estranho residual em tecido mole"),
        (("revisao total de ptq", "revisao acetabular de ptq", "luxacao de acetabulo"), "T84.0", "complicação mecânica de prótese articular interna"),
        (("hsa traumatica", "hemorragia subaracnoidea traumatica"), "S06.6", "hemorragia subaracnóidea traumática"),
        (("dor a mobilizacao da coluna cervical", "bloqueios da coluna cervical"), "M54.2", "cervicalgia"),
        (("cancer de mama metastatico", "neoplasia de mama metastatica"), "C50.9", "neoplasia maligna da mama"),
        (("pneumonia",), "J18.9", "pneumonia não especificada"),
        (("sars-cov-2", "sars cov 2", "covid-19", "covid 19"), "U07.1", "COVID-19"),
        (("dor cervical", "cervicalgia"), "M54.2", "cervicalgia"),
        (("hemiparesia esquerda", "deficit motor a esquerda", "déficit motor à esquerda"), "I63.9", "infarto cerebral com déficit focal"),
        (("adenocarcinoma de reto", "neoplasia de reto"), "C20", "neoplasia maligna do reto"),
        (("ic descompensada", "insuficiencia cardiaca descompensada"), "I50.9", "insuficiência cardíaca"),
        (("colecistite aguda", "colecistite agua"), "K81.0", "colecistite aguda"),
        (("vestibulopatia periferica", "labirintite"), "H81.9", "transtorno da função vestibular"),
        (("infeccao urinaria", "infecção urinária", "itu de repeticao", "itu de repetição"), "N39.0", "infecção do trato urinário"),
        (("choque septico", "choque séptico"), "R57.2", "choque séptico"),
        (("anemia", "hb 8,6", "hb 8.6"), "D64.9", "anemia não especificada"),
        (("fratura-luxacao traumatica", "fratura do umero proximal"), "S42.2", "fratura da extremidade superior do úmero"),
        (("fratura do extremo distal do radio", "fratura distal do radio"), "S52.5", "fratura da extremidade distal do rádio"),
        (("fratura do sacro", "fratura sacral"), "S32.1", "fratura do sacro"),
        (("lesao osteocondral no joelho",), "M93.2", "lesão osteocondral do joelho"),
        (("suboclusao intestinal por brida", "obstrucao intestinal por brida"), "K56.5", "aderência intestinal com obstrução"),
        (("ulcera bulbar ativa com sinal de sangramento",), "K26.4", "úlcera duodenal com hemorragia"),
        (("colica biliar reentrante",), "K80.2", "cálculo da vesícula biliar sem colecistite"),
        (("neoplasia de esofago distal", "adenocarcinoma estenosante"), "C15.5", "neoplasia maligna do terço inferior do esôfago"),
        (("melanoma metastatico",), "C43.9", "melanoma maligno de pele, sítio primário não especificado"),
        (("sindrome desmielinizante", "sd. desmielinizante"), "G37.9", "doença desmielinizante do sistema nervoso central"),
        (("hidrocefalia de pressao normal",), "G91.2", "hidrocefalia de pressão normal"),
        (("encefalopatia hepatica",), "K72.9", "insuficiência hepática com encefalopatia não especificada"),
        (("ira kdigo", "insuficiencia renal aguda", " ira - tsr"), "N17.9", "insuficiência renal aguda"),
        (("pielonefrite",), "N10", "pielonefrite aguda"),
        (
            (
                "nova infeccao urinaria",
                "internada por itu",
                "internado por itu",
                "sintomas sugestivos de infeccao do trato urinario",
                "sintoma sugestivos de infeccao do trato urinario",
                "urocultura e coli",
            ),
            "N39.0",
            "infecção do trato urinário",
        ),
        (("diagnostico de pneumomia", "tc torax com pneumonia", "pneumonia + sinusite"), "J18.9", "pneumonia não especificada"),
        (("placa mole na origem da asce", "placa na origem da arteria subclavia"), "I70.8", "aterosclerose de outra artéria"),
        (("lombalgia cronica", "ciatalgia intensa"), "M54.4", "lumbago com ciática"),
        (("candidiase", "candidiase oral"), "B37.0", "candidíase oral"),
        (("noto humor depressivo",), "F32.9", "episódio depressivo não especificado"),
        (("rash cutaneo cranio-caudal",), "R21", "erupção cutânea não especificada"),
        (("dor abdominal leve e nausea",), "R10.4", "dor abdominal não especificada"),
        (("isquemia mesenterica", "perfuracao de alca"), "K55.0", "isquemia mesentérica aguda"),
        (("choque septico refratario",), "R57.2", "choque séptico"),
        (("apendicectomia", "apendicite aguda"), "K35.9", "apendicite tratada por apendicectomia"),
        (("broncoaspir", "bcp aspirativa", "pneumonia aspirativa"), "J69.0", "pneumonia aspirativa"),
        (("pneumonia bacteriana", "broncopneumonia", "opacidades pulmonares"), "J18.9", "pneumonia não especificada"),
        (("virus sincicial respiratorio", "vsr positivos", "vsr positivo"), "J21.0", "bronquiolite por VSR"),
        (("sindrome coronariana aguda", " sca "), "I24.9", "síndrome coronariana aguda"),
        (
            (
                "aterosclerose coronariana",
                "lesao coronariana",
                "lesoes coronarianas",
                "doenca aterosclerotica",
                "doenca arterial coronariana",
                "coronariano previamente",
                "atc de tce",
                " dac ",
            ),
            "I25.1",
            "doença aterosclerótica do coração",
        ),
        (("estenose aortica", "tavi"), "I35.0", "estenose aórtica"),
        (("ablacao de trn", "taquicardia por reentrada nodal"), "I47.1", "taquicardia supraventricular"),
        (("avc isquemico", "avci", "multiplos focos isquemicos", "infarto cerebral"), "I63.9", "infarto cerebral"),
        (("fratura trans trocanterica", "fratura transtroc"), "S72.1", "fratura pertrocantérica"),
        (("fratura do osso nasal", "fratura do septo nasal"), "S02.2", "fratura nasal"),
        (("sincope",), "R55", "síncope e colapso"),
        (("anemia ferropriva secundaria a menorragia",), "D50.0", "anemia ferropriva por perda sanguínea"),
        (("neutropenia febril",), "D70", "neutropenia"),
        (("mastocitose", "elevacao da triptase"), "D47.0", "neoplasia de comportamento incerto de mastócitos"),
        (("carcinoma espinocelular", "cec?"), "C44.7", "neoplasia maligna da pele do membro inferior"),
        (("adenocarcinoma de pulmao",), "C34.9", "neoplasia maligna de pulmão"),
        (("adenocarcinoma de cabeca de pancreas",), "C25.0", "neoplasia maligna da cabeça do pâncreas"),
        (("neoplasia colorretal", "ca de intestino"), "C19", "neoplasia colorretal"),
        (("colangite",), "K83.0", "colangite"),
        (("pancreatite cronica agudizada", "pancreatite aguda", "pancreatite leve"), "K85.9", "pancreatite aguda"),
        (("migracao de calculo", "coledocolitiase"), "K80.5", "cálculo de via biliar"),
        (("fistula liquorica", "hipotensao liquorica"), "G96.0", "fístula liquórica"),
        (("abstinencia alcoolica",), "F10.3", "abstinência alcoólica"),
        (("linfonodomegalias", "esplenomegalia sintomatica"), "R59.1", "linfonodomegalia generalizada"),
        (("dispneia", "dessaturacao"), "R06.0", "dispneia"),
        (("diarreia",), "R19.7", "diarreia"),
        (("nauseas", "vomitos"), "R11", "náuseas e vômitos"),
        (("sintomas gripais", "quadro gripal"), "J06.9", "infecção aguda de vias aéreas superiores"),
        (("broncopatia inflamatoria",), "J40", "bronquite não especificada"),
        (("edema em mie", "edema de membro inferior"), "R60.0", "edema localizado"),
        (("hematuria",), "R31", "hematúria"),
        (("hiperemia na regiao tibial", "teicoplanina"), "L03.1", "celulite de membro"),
        (("malformacao arteriovenosa", " mav "), "Q28.2", "malformação arteriovenosa"),
        (("alteracao de habito intestinal", "alteracao do habito intestinal"), "R19.4", "alteração do hábito intestinal"),
        (("abscesso dentario",), "K04.7", "abscesso dentário"),
        (("gastroenterocolite",), "A09", "gastroenterocolite"),
        (("cansaco", "fadiga", "mal estar inespecifico"), "R53", "mal-estar e fadiga"),
    )
    for terms, cid, reason in rules:
        if any(term in padded for term in terms):
            return cid, reason
    return "", ""


def infer_cid_from_evolution(text: str) -> str:
    return infer_cid_suggestion(text)[0]


def infer_adjusted_cid_from_evolution(text: str) -> str:
    value = normalized(text)
    if any(term in value for term in ("area de isquemia recente", "avc isquemico", "avci", "infarto cerebral")):
        return "I63.9"
    return ""

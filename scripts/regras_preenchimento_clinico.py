#!/usr/bin/env python3
"""Regras auxiliares para preencher a base clinica antes do lancamento."""

from __future__ import annotations


def regra_cid_internacao() -> str:
    return (
        "Preencher CID de internacao com o diagnostico inicial que motivou "
        "a internacao. Se houver mudanca/evolucao diagnostica posterior, "
        "preencher CID ajustado com o diagnostico atualizado."
    )


def regra_comorbidades() -> str:
    return (
        "Preencher comorbidades a partir dos antecedentes/comorbidades "
        "descritos no relatorio. Se o relatorio nao informar comorbidades, "
        "usar 'Sem comorbidades'."
    )


def regra_sinais_vitais() -> str:
    return (
        "Nao inventar PA, FC, FR, SpO2, temperatura ou outros sinais vitais. "
        "Preencher apenas quando o relatorio informar o valor ou quando o "
        "Salus oferecer opcao explicita de nao mensurado."
    )

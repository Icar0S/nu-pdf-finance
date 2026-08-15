"""Erros do pipeline. Todos herdam de ErroFatura para a CLI capturar num lugar so."""


class ErroFatura(Exception):
    """Base de tudo que pode dar errado ao processar uma fatura."""


class ErroExtracao(ErroFatura):
    """O PDF nao tem a estrutura esperada de uma fatura do Nubank."""


class ErroReconciliacao(ErroFatura):
    """As contas da fatura nao fecham. O import e rejeitado inteiro."""


class ErroExport(ErroFatura):
    """A planilha nao esta no formato que o export espera."""

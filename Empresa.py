class Empresa:
    def __init__(self, link=None, reclamacoes_respondidas=None, voltariam_negocio=None, indice_solucao=None, nota_consumidor=None):
        self.link = link
        self.reclamacoes_respondidas = reclamacoes_respondidas
        self.voltariam_negocio = voltariam_negocio
        self.indice_solucao = indice_solucao
        self.nota_consumidor = nota_consumidor


    def to_dict(self):
        return {
            "Link": self.link,
            "Reclamações Respondidas": self.reclamacoes_respondidas,
            "Voltariam ao Negócio": self.voltariam_negocio,
            "Índice de Solução": self.indice_solucao,
            "Nota do Consumidor": self.nota_consumidor,
        }
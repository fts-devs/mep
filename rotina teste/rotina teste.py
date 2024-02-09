import json
import openpyxl

def cadastrar_usuario():
    nome = input("Digite seu nome: ")
    email = input("Digite seu email: ")
    
    # Criar um dicionário com as informações do usuário
    usuario = {"nome": nome, "email": email}
    
    # Converter o dicionário para formato JSON e salvar em um arquivo
    with open("usuario.json", "w") as arquivo:
        json.dump(usuario, arquivo)

def escolher_vestibular():
    vestibulares = ["Vestibular A", "Vestibular B", "Vestibular C", "Vestibular D", "Vestibular E", "Vestibular F"]
    
    print("Escolha um vestibular:")
    for i, vestibular in enumerate(vestibulares, start=1):
        print(f"{i}. {vestibular}")

    escolha = int(input("Digite o número do vestibular desejado: "))
    return vestibulares[escolha - 1]

def definir_horas_estudo():
    horas_estudo = input("Digite o número de horas que pretende estudar diariamente (mínimo 2, máximo 3): ")
    
    # Substituir vírgula por ponto
    horas_estudo = horas_estudo.replace(',', '.')

    # Garantir que o número de horas esteja entre 2 e 3
    horas_estudo = max(2, min(3, float(horas_estudo)))
    
    return horas_estudo

def excluir_domingos():
    opcao = input("Deseja excluir os domingos dos estudos? (S/N): ")
    return opcao.lower() == "s"

# ...

def gerar_cronograma(vestibular, horas_estudo, excluir_domingo):
    # Abrir a planilha Excel
    planilha = openpyxl.load_workbook("cfis.xlsx")  
    # Selecione a primeira planilha
    sheet = planilha.active
    
    # Mapeamento dos dias da semana para as células da tabela
    cronograma = {
        "Segunda": "A1:A6",
        "Terça": "A8:A10",
        "Quarta": "A5:A9",
        "Quinta": "A50:A52",
        "Sexta": "A11:A15",
        "Sábado": "A29:A30",
        "Domingo":"A35:A40"
    }
    
    # Excluir domingo, se necessário
    if excluir_domingo:
        del cronograma["Domingo"]
    else:
        cronograma["Domingo"] = "A80:A83"
    # Criar lista de estudos com base no cronograma
    lista_estudos = {dia: [cell.value if isinstance(cell, openpyxl.cell.cell.Cell) else cell[0].value for cell in sheet[cronograma[dia]]] for dia in cronograma}

    # Salvar lista de estudos em um arquivo txt com codificação UTF-8
    with open(f"{vestibular}_cronograma.txt", "w", encoding="utf-8") as arquivo:
        json.dump(lista_estudos, arquivo, ensure_ascii=False, indent=2)

    print("Cronograma gerado e salvo com sucesso!")

# ...


# Função principal
def main():
    cadastrar_usuario()
    
    vestibular_escolhido = escolher_vestibular()
    print(f"Você escolheu o vestibular: {vestibular_escolhido}")
    
    horas_estudo = definir_horas_estudo()
    print(f"Você pretende estudar {horas_estudo} horas por dia.")
    
    excluir_domingo = excluir_domingos()
    if excluir_domingo:
        print("Você optou por excluir os domingos dos estudos.")
    else:
        print("Você não escolheu excluir os domingos dos estudos.")

    gerar_cronograma(vestibular_escolhido, horas_estudo, excluir_domingo)

if __name__ == "__main__":
    main()

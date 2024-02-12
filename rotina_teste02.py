import openpyxl
from datetime import datetime, timedelta

#Caso para três dias de estudo. A lógica precisa ser adaptada para mais dias (ou menos '_' )

def obter_aulas_do_dia(sheet, n_s, tempo_diario, aulas_por_dia):
    aulas_do_dia = []
    
    for row in sheet.iter_rows(min_row=2, max_col=6, max_row=514):
        id_aula = int(row[0].value)
        if id_aula is not None and isinstance(id_aula, (int, float)):
            id_aula = int(id_aula)
            frente = row[1].value
            grande_topico = row[2].value
            relevancia = row[5].value
            topico = f'{grande_topico} - {row[3].value}'
            aula = row[4].value


        if id_aula <= aulas_por_dia and relevancia in ['BAI','MED', 'ALT']:
            aulas_do_dia.append({
                'ID': id_aula,
                'Frente': frente,
                'Tópico': topico,
                'Aula': aula
            })

    return aulas_do_dia

def gerar_cronograma():
    # Receber informações do usuário
    nome_aluno = input("Digite o nome do aluno: ")
    data_inicio = datetime.strptime(input("Digite a data de início (dd/mm/aaaa): "), "%d/%m/%Y")
    data_fim = datetime.strptime(input("Digite a data de término (dd/mm/aaaa): "), "%d/%m/%Y")

    dias_semana = input("Digite os dias da semana que o aluno gostaria de estudar (separados por vírgula): ").split(',')
    dias_semana = [dia.strip().capitalize() for dia in dias_semana]
    tempo_diario = float(input("Digite a quantidade diária de horas de estudo (entre 1,5 e 4): "))

    # Calcular o número de semanas com base nas datas fornecidas
    n_s = (data_fim - data_inicio).days // 7 + 1

    # Carregar a planilha Excel
    planilha = openpyxl.load_workbook("Conteúdos_de_Física_CR.xlsx")
    sheet = planilha.active

    # Montar o cronograma
    cronograma = {}
    
    for i in range(n_s * len(dias_semana)):
        dia_semana = dias_semana[i % len(dias_semana)]
        data_atual = data_inicio + timedelta(days=i * 7)

        if data_atual.day == 11 and dia_semana == 'Sexta':
            continue  # Se o primeiro dia de estudo for sexta, pule para a próxima semana

        # Verificar a quantidade de aulas por dia com base nas condições fornecidas
        aulas_por_dia = 2
        if 20 < n_s <= 30:
            if 1.5 <= tempo_diario <= 2:
                aulas_por_dia = 3
            elif 2.5 <= tempo_diario <= 3:
                aulas_por_dia = 4
            elif 3.5 <= tempo_diario <= 4:
                aulas_por_dia = 5
        elif 15 <= n_s <= 20:
            if 1.5 <= tempo_diario <= 2:
                aulas_por_dia = 3
            elif 2.5 <= tempo_diario <= 3:
                aulas_por_dia = 4
            elif 3.5 <= tempo_diario <= 4:
                aulas_por_dia = 5
        elif 10 <= n_s < 15:
            if 1.5 <= tempo_diario <= 2:
                aulas_por_dia = 4
            elif 2.5 <= tempo_diario <= 3:
                aulas_por_dia = 5
            elif 3.5 <= tempo_diario <= 4:
                aulas_por_dia = 6
        else:
            print("Intervalo de semanas não suportado.")
            return None

        aulas_do_dia = obter_aulas_do_dia(sheet, n_s, tempo_diario, aulas_por_dia)

        cronograma[data_atual.strftime("%d/%m/%Y")] = {
            'Frente do dia': aulas_do_dia[0]['Frente'],
            'Aulas do dia': aulas_do_dia,
            'Exemplos resolvidos': 'Refaça os exemplos resolvidos',
            'Questões da lista': f'Comece a lista do {aulas_do_dia[0]["Tópico"]}',
            'Tempo de estudo': f'{tempo_diario} horas'
        }

    # Salvar cronograma em um arquivo txt
    with open(f'cronograma_{nome_aluno}.txt', 'w') as file:
        for data, dia in cronograma.items():
            file.write(f'Data: {data}\n')
            file.write(f'Frente do dia: {dia["Frente do dia"]}\n')
            file.write('Aulas do dia:\n')
            for aula in dia['Aulas do dia']:
                file.write(f'  - Aula {aula["ID"]}: {aula["Aula"]}\n')
            file.write(f'- Exemplos resolvidos: {dia["Exemplos resolvidos"]}\n')
            file.write(f'- Questões da lista: {dia["Questões da lista"]}\n')
            file.write(f'- Tempo de estudo: {dia["Tempo de estudo"]}\n\n')


gerar_cronograma()
